import streamlit as st
import zipfile
import io
import re
import json
import urllib.request

st.set_page_config(page_title="Šablona školení", page_icon="📄", layout="centered")
st.title("📄 Automatické vyplnění šablony školení")
st.caption("Nahrajte šablonu a zadejte URL stránky školení — zbývající vyplní AI.")


# ── Word XML helpers ──────────────────────────────────────────────────────────

def extract_paragraph_text(p_xml: str) -> str:
    """Extrahuje čistý text z XML odstavce."""
    texts = re.findall(r"<w:t[^>]*>([\s\S]*?)</w:t>", p_xml)
    return "".join(texts).strip()


def find_green_groups(xml_str: str) -> list[dict]:
    # Nejdřív sesbírej všechny odstavce s jejich textem a pozicí
    paragraphs = []
    for p_match in re.finditer(r"<w:p[ >][\s\S]*?</w:p>", xml_str):
        paragraphs.append({
            "xml": p_match.group(),
            "start": p_match.start(),
            "text": extract_paragraph_text(p_match.group()),
            "is_heading": bool(re.search(r'<w:pStyle w:val="[^"]*[Hh]eading', p_match.group())
                               or re.search(r"<w:b/>|<w:b\s", p_match.group())),
        })

    groups = []
    for p_idx, p_data in enumerate(paragraphs):
        p_xml, p_start = p_data["xml"], p_data["start"]
        runs, pos = [], 0
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
                    # Najdi nejbližší nadpis/label před tímto polem (max 5 odstavců zpět)
                    context_label = ""
                    for look_back in range(p_idx - 1, max(p_idx - 6, -1), -1):
                        candidate = paragraphs[look_back]["text"]
                        if candidate and len(candidate) < 120:
                            context_label = candidate
                            break

                    groups.append({
                        "text": text.strip(),
                        "context": context_label,
                        "runs": runs[i:j],
                        "start_pos": runs[i]["abs_start"],
                        "end_pos": runs[j - 1]["abs_end"],
                    })
                i = j
            else:
                i += 1
    return groups


def text_to_word_runs(new_text: str, base_run_xml: str) -> str:
    """
    Převede text (potenciálně s \\n) na sekvenci Word runů s <w:br/> místo zalomení.
    Zachová formátování (tučné, kurzíva atd.) z base_run_xml.
    """
    # Odstraň zelené zvýraznění z base runu
    clean_run = base_run_xml
    clean_run = clean_run.replace('<w:highlight w:val="green"/>', "")
    clean_run = clean_run.replace("<w:highlight w:val='green'/>", "")

    # Vyextrahuj rPr (formátování) z base runu
    rpr_match = re.search(r"<w:rPr>([\s\S]*?)</w:rPr>", clean_run)
    rpr = f"<w:rPr>{rpr_match.group(1)}</w:rPr>" if rpr_match else ""

    lines = (new_text or "").split("\n")
    runs = []

    for idx, line in enumerate(lines):
        escaped = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if idx == 0:
            # První řádek: použij původní run (zachová veškeré jeho atributy)
            run = re.sub(
                r"<w:t[^>]*>[\s\S]*?</w:t>",
                f'<w:t xml:space="preserve">{escaped}</w:t>',
                clean_run,
            )
            runs.append(run)
        else:
            # Každý další řádek: nejdřív run s <w:br/>, pak run s textem
            runs.append(f"<w:r>{rpr}<w:br/></w:r>")
            if escaped:
                runs.append(f'<w:r>{rpr}<w:t xml:space="preserve">{escaped}</w:t></w:r>')

    return "".join(runs)


def apply_replacements(xml_str: str, groups: list[dict], replacements: list[str]) -> str:
    items = sorted(
        [{"group": g, "new_text": replacements[i]} for i, g in enumerate(groups)],
        key=lambda x: x["group"]["start_pos"], reverse=True,
    )
    result = xml_str
    for item in items:
        g = item["group"]
        replacement_xml = text_to_word_runs(item["new_text"], g["runs"][0]["xml"])
        result = result[: g["start_pos"]] + replacement_xml + result[g["end_pos"]:]
    return result


# ── Make webhook volání ───────────────────────────────────────────────────────

def call_make_webhook(webhook_url: str, page_url: str, groups: list[dict]) -> list[dict]:
    """
    Pošle POST na Make webhook s URL stránky a seznamem polí obohacených o kontext.
    Make vrátí JSON pole: [{"value": "...", "warning": "..."}, ...]
    """
    # Každé pole pošleme jako objekt s textem a kontextem ze šablony
    fields_payload = [
        {
            "text": g["text"],
            "context": g.get("context", ""),
            "label": f"[{g.get('context', '')}] {g['text'][:80]}" if g.get("context") else g["text"][:80],
        }
        for g in groups
    ]

    payload = json.dumps({
        "url": page_url,
        "fields": fields_payload,
    }).encode("utf-8")

    req = urllib.request.Request(
        webhook_url,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    with urllib.request.urlopen(req, timeout=60) as resp:
        body = resp.read().decode("utf-8")

    data = json.loads(body)

    # Make může vrátit buď pole objektů [{"value":...}] nebo pole stringů ["..."]
    results = []
    for i, item in enumerate(data):
        original_text = groups[i]["text"] if i < len(groups) else ""
        if isinstance(item, dict):
            val = item.get("value", original_text)
            warning = item.get("warning", None)
        else:
            val = str(item)
            warning = None

        if val == "__CHYBÍ__":
            val = original_text
            warning = "⚠️ Hodnota nenalezena na stránce — doplňte ručně."
        elif val.strip() == original_text.strip():
            lower = original_text.lower()
            if any(k in lower for k in ("termín", "datum", "cena", "kč")):
                warning = "⚠️ Hodnota se nezměnila — zkontrolujte ručně."

        results.append({"value": val, "warning": warning})

    while len(results) < len(groups):
        results.append({"value": groups[len(results)]["text"], "warning": None})

    return results


# ── UI ────────────────────────────────────────────────────────────────────────

webhook_url = "https://hook.eu2.make.com/0osnyixa6ivhkm2rvpjapqdtqqxwlsip"

st.divider()

col1, col2 = st.columns([3, 2])
with col1:
    url = st.text_input("URL stránky školení", placeholder="https://studiow.cz/...")
with col2:
    template_file = st.file_uploader("Word šablona (.docx)", type=["docx"])

if st.button("🔍 Zpracovat stránku", type="primary",
             disabled=not (url and template_file and webhook_url), use_container_width=True):

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

    with st.spinner("Make zpracovává stránku a extrahuje data... (může trvat 15–30 s)"):
        try:
            results = call_make_webhook(webhook_url, url, groups)
        except Exception as e:
            st.error(f"Chyba při volání Make webhooku: {e}")
            st.stop()

    # Přepiš session state před renderem widgetů
    for i, res in enumerate(results):
        st.session_state[f"field_{i}"] = res["value"]

    st.session_state.update({"groups": groups, "results": results,
                              "xml_str": xml_str, "raw_zip": raw})


# ── Kontrola a stažení ────────────────────────────────────────────────────────

if "groups" in st.session_state:
    st.divider()
    st.subheader("✏️ Zkontrolujte a upravte navržená pole")

    groups = st.session_state["groups"]
    results = st.session_state["results"]
    edited = []

    warn_count = sum(1 for r in results if r.get("warning"))
    if warn_count:
        st.warning(
            f"**{warn_count} {'pole vyžaduje' if warn_count == 1 else 'pole vyžadují'} ruční doplnění** "
            "— termín a cena se na Studiow načítají dynamicky a nelze je automaticky přečíst."
        )

    for i, g in enumerate(groups):
        res = results[i] if i < len(results) else {"value": g["text"], "warning": None}
        with st.expander(
            f"Pole {i+1} — {g['text'][:55]}{'...' if len(g['text']) > 55 else ''}",
            expanded=bool(res.get("warning"))
        ):
            st.markdown(f"**Původní text v šabloně:**\n\n> {g['text'][:300]}{'...' if len(g['text']) > 300 else ''}")
            if res.get("warning"):
                st.warning(res["warning"])
            new_val = st.text_area("Nový text", key=f"field_{i}", height=120)
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
