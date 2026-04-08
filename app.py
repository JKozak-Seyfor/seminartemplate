import streamlit as st
import zipfile
import io
import re
import json
import anthropic
import urllib.request
import urllib.error
from html.parser import HTMLParser

st.set_page_config(page_title="Šablona školení", page_icon="📄", layout="centered")
st.title("📄 Automatické vyplnění šablony školení")
st.caption("Nahrajte šablonu a zadejte URL stránky školení — zbývající vyplní AI.")


# ── Stažení stránky ───────────────────────────────────────────────────────────

class TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self._parts = []
        self._skip = False

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "nav", "footer"):
            self._skip = True

    def handle_endtag(self, tag):
        if tag in ("script", "style", "nav", "footer"):
            self._skip = False
        if tag in ("p", "li", "h1", "h2", "h3", "h4", "br", "div"):
            self._parts.append("\n")

    def handle_data(self, data):
        if not self._skip and data.strip():
            self._parts.append(data.strip())

    def get_text(self):
        return " ".join(t for t in self._parts if t.strip())


def fetch_page_text(url: str) -> tuple[str, str | None]:
    try:
        req = urllib.request.Request(
            url, headers={"User-Agent": "Mozilla/5.0 (compatible; TemplateBot/1.0)"}
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
        parser = TextExtractor()
        parser.feed(raw)
        text = parser.get_text()
        text = re.sub(r" {2,}", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text[:12000], None
    except Exception as e:
        return "", str(e)


# ── Word XML helpers ──────────────────────────────────────────────────────────

def find_green_groups(xml_str: str) -> list[dict]:
    groups = []
    for p_match in re.finditer(r"<w:p[ >][\s\S]*?</w:p>", xml_str):
        p_xml, p_start, runs, pos = p_match.group(), p_match.start(), [], 0
        while pos < len(p_xml):
            s1 = p_xml.find("<w:r>", pos)
            s2 = p_xml.find("<w:r ", pos)
            if s1 == -1 and s2 == -1:
                break
            r_start = s2 if s1 == -1 else s1 if s2 == -1 else min(s1, s2)
            r_end = p_xml.find("</w:r>", r_start)
            if r_end == -1:
                break
            run_xml = p_xml[r_start: r_end + 6]
            is_green = "w:highlight" in run_xml and (
                'w:val="green"' in run_xml or "w:val='green'" in run_xml
            )
            t_m = re.search(r"<w:t[^>]*>([\s\S]*?)</w:t>", run_xml)
            runs.append({
                "xml": run_xml, "text": t_m.group(1) if t_m else "",
                "is_green": is_green,
                "abs_start": p_start + r_start, "abs_end": p_start + r_end + 6,
            })
            pos = r_end + 6
        i = 0
        while i < len(runs):
            if runs[i]["is_green"]:
                j, text = i, ""
                while j < len(runs) and runs[j]["is_green"]:
                    text += runs[j]["text"]
                    j += 1
                if text.strip():
                    groups.append({
                        "text": text.strip(), "runs": runs[i:j],
                        "start_pos": runs[i]["abs_start"],
                        "end_pos": runs[j - 1]["abs_end"],
                    })
                i = j
            else:
                i += 1
    return groups


def apply_replacements(xml_str: str, groups: list[dict], replacements: list[str]) -> str:
    items = sorted(
        [{"group": g, "new_text": replacements[i]} for i, g in enumerate(groups)],
        key=lambda x: x["group"]["start_pos"], reverse=True,
    )
    result = xml_str
    for item in items:
        g = item["group"]
        fr = g["runs"][0]["xml"]
        fr = fr.replace('<w:highlight w:val="green"/>', "")
        fr = fr.replace("<w:highlight w:val='green'/>", "")
        escaped = (item["new_text"] or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        fr = re.sub(r"<w:t[^>]*>[\s\S]*?</w:t>",
                    f'<w:t xml:space="preserve">{escaped}</w:t>', fr)
        result = result[: g["start_pos"]] + fr + result[g["end_pos"]:]
    return result


# ── AI extrakce ───────────────────────────────────────────────────────────────

DYNAMIC_KEYWORDS = {"termín", "datum", "cena", "kč", "price", "date"}


def extract_from_page(page_text: str, url: str, groups: list[dict], api_key: str) -> list[dict]:
    client = anthropic.Anthropic(api_key=api_key)

    field_list = "\n".join([
        f'Pole {i+1}: "{g["text"][:200]}{"..." if len(g["text"]) > 200 else ""}"'
        for i, g in enumerate(groups)
    ])

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": f"""Zde je obsah webové stránky školení (URL: {url}):

---
{page_text}
---

V šabloně Word dokumentu jsou tato zeleně zvýrazněná pole k aktualizaci:
{field_list}

Na základě obsahu stránky navrhni nový text pro každé pole. Pravidla:
- Název školení: přesný název ze stránky
- Termín/datum: pokud na stránce NENÍ, vrať přesně "__CHYBÍ__"
- Cena: pokud na stránce NENÍ, vrať přesně "__CHYBÍ__"
- Jméno lektora: přesné jméno ze stránky
- Profil lektora: zkopíruj celý profil/bio lektora
- Obsah/program: zkopíruj celý program ze stránky (i s odrážkami)
- Pro koho je školení: pokud na stránce, zkopíruj; pokud ne, odvoď z tématu
- Zachovej původní délku a styl

DŮLEŽITÉ: Vrať POUZE JSON pole o {len(groups)} stringách. Žádný jiný text, žádné backticky."""}],
    )

    text_content = "".join(b.text for b in response.content if hasattr(b, "text"))
    json_match = re.search(r"\[[\s\S]*\]", text_content)

    try:
        values = json.loads(json_match.group()) if json_match else []
    except Exception:
        values = []

    while len(values) < len(groups):
        values.append(groups[len(values)]["text"])

    results = []
    for i, g in enumerate(groups):
        val = values[i] if i < len(values) else g["text"]
        warning = None
        if val == "__CHYBÍ__":
            val = g["text"]
            warning = "⚠️ Toto pole nebylo na stránce nalezeno (pravděpodobně termín nebo cena, které se načítají dynamicky). Doplňte ručně."
        elif val.strip() == g["text"].strip():
            lower = g["text"].lower()
            if any(kw in lower for kw in DYNAMIC_KEYWORDS):
                warning = "⚠️ Hodnota se nezměnila — pravděpodobně dynamické pole (termín/cena). Zkontrolujte ručně."
        results.append({"value": val, "warning": warning})

    return results


# ── UI ────────────────────────────────────────────────────────────────────────

api_key = st.secrets.get("ANTHROPIC_API_KEY", "") if hasattr(st, "secrets") else ""
if not api_key:
    api_key = st.text_input(
        "Anthropic API klíč", type="password",
        help="Váš API klíč z console.anthropic.com. Nebo uložte do Streamlit secrets jako ANTHROPIC_API_KEY.",
        placeholder="sk-ant-...",
    )

st.divider()

col1, col2 = st.columns([3, 2])
with col1:
    url = st.text_input("URL stránky školení", placeholder="https://studiow.cz/...")
with col2:
    template_file = st.file_uploader("Word šablona (.docx)", type=["docx"])

if st.button("🔍 Zpracovat stránku", type="primary",
             disabled=not (url and template_file and api_key), use_container_width=True):

    with st.spinner("Čtu šablonu..."):
        raw = template_file.read()
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            xml_str = zf.read("word/document.xml").decode("utf-8")
        groups = find_green_groups(xml_str)

    if not groups:
        st.error("V šabloně nebyly nalezeny žádné **zeleně zvýrazněné** oblasti. "
                 "Zkontrolujte, zda je text označen zeleným highlightem (ne barvou písma).")
        st.stop()

    st.success(f"Nalezeno **{len(groups)} zelených polí** v šabloně.")

    with st.spinner("Načítám stránku školení..."):
        page_text, fetch_error = fetch_page_text(url)

    if fetch_error or not page_text.strip():
        st.error(f"Nepodařilo se načíst stránku: {fetch_error or 'prázdná odpověď'}")
        st.stop()

    with st.spinner("Claude analyzuje obsah stránky..."):
        try:
            results = extract_from_page(page_text, url, groups, api_key)
        except Exception as e:
            st.error(f"Chyba při volání API: {e}")
            st.stop()

    # Smaž staré hodnoty text_area — jinak Streamlit ignoruje nový value=
    old_count = len(st.session_state.get("groups") or [])
    for i in range(old_count + 5):
        st.session_state.pop(f"field_{i}", None)

    st.session_state.update({"groups": groups, "results": results,
                              "xml_str": xml_str, "raw_zip": raw})

# ── Kontrola a stažení ────────────────────────────────────────────────────────

if "groups" in st.session_state:
    st.divider()
    st.subheader("✏️ Zkontrolujte a upravte navržená pole")

    groups = st.session_state["groups"]
    results = st.session_state["results"]
    edited = []

    warn_count = sum(1 for r in results if r["warning"])
    if warn_count:
        st.warning(
            f"**{warn_count} {'pole vyžaduje' if warn_count == 1 else 'pole vyžadují'} ruční doplnění** "
            "— termín a cena se na stránce Studiow načítají dynamicky a nelze je automaticky přečíst."
        )

    for i, g in enumerate(groups):
        res = results[i] if i < len(results) else {"value": g["text"], "warning": None}
        with st.expander(
            f"Pole {i+1} — {g['text'][:55]}{'...' if len(g['text']) > 55 else ''}",
            expanded=bool(res["warning"])
        ):
            st.markdown(f"**Původní text v šabloně:**\n\n> {g['text'][:300]}{'...' if len(g['text']) > 300 else ''}")
            if res["warning"]:
                st.warning(res["warning"])
            new_val = st.text_area("Nový text", value=res["value"], key=f"field_{i}", height=120)
            edited.append(new_val)

    st.divider()

    if st.button("⬇️ Stáhnout upravený dokument", type="primary", use_container_width=True):
        modified_xml = apply_replacements(st.session_state["xml_str"], groups, edited)
        out_buf = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(st.session_state["raw_zip"])) as src:
            with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as dst:
                for item in src.infolist():
                    dst.writestr(item,
                        modified_xml.encode("utf-8") if item.filename == "word/document.xml"
                        else src.read(item.filename))
        st.download_button(
            "📥 Klikněte zde pro stažení", data=out_buf.getvalue(),
            file_name="registracni_stranka.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        st.success("Dokument je připraven ke stažení!")
