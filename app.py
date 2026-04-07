import streamlit as st
import zipfile
import io
import re
import json
import anthropic

st.set_page_config(
    page_title="Šablona školení",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Automatické vyplnění šablony školení")
st.caption("Nahrajte šablonu a zadejte URL stránky školení — zbývající vyplní AI.")

# ── Helpers ──────────────────────────────────────────────────────────────────

def find_green_groups(xml_str: str) -> list[dict]:
    """Najde skupiny zeleně zvýrazněných běhů v XML dokumentu."""
    groups = []
    p_pattern = re.compile(r"<w:p[ >][\s\S]*?</w:p>")

    for p_match in p_pattern.finditer(xml_str):
        p_xml = p_match.group()
        p_start = p_match.start()
        runs = []
        pos = 0

        while pos < len(p_xml):
            s1 = p_xml.find("<w:r>", pos)
            s2 = p_xml.find("<w:r ", pos)
            if s1 == -1 and s2 == -1:
                break
            if s1 == -1:
                r_start = s2
            elif s2 == -1:
                r_start = s1
            else:
                r_start = min(s1, s2)

            r_end = p_xml.find("</w:r>", r_start)
            if r_end == -1:
                break

            run_xml = p_xml[r_start : r_end + 6]
            is_green = "w:highlight" in run_xml and (
                'w:val="green"' in run_xml or "w:val='green'" in run_xml
            )
            t_match = re.search(r"<w:t[^>]*>([\s\S]*?)</w:t>", run_xml)
            text = t_match.group(1) if t_match else ""

            runs.append(
                {
                    "xml": run_xml,
                    "text": text,
                    "is_green": is_green,
                    "abs_start": p_start + r_start,
                    "abs_end": p_start + r_end + 6,
                }
            )
            pos = r_end + 6

        i = 0
        while i < len(runs):
            if runs[i]["is_green"]:
                j, text = i, ""
                while j < len(runs) and runs[j]["is_green"]:
                    text += runs[j]["text"]
                    j += 1
                if text.strip():
                    groups.append(
                        {
                            "text": text.strip(),
                            "runs": runs[i:j],
                            "start_pos": runs[i]["abs_start"],
                            "end_pos": runs[j - 1]["abs_end"],
                        }
                    )
                i = j
            else:
                i += 1

    return groups


def apply_replacements(xml_str: str, groups: list[dict], replacements: list[str]) -> str:
    """Nahradí zeleně zvýrazněný text novými hodnotami."""
    items = sorted(
        [
            {"group": g, "new_text": replacements[i]}
            for i, g in enumerate(groups)
        ],
        key=lambda x: x["group"]["start_pos"],
        reverse=True,
    )

    result = xml_str
    for item in items:
        g = item["group"]
        new_text = item["new_text"] or ""
        first_run = g["runs"][0]["xml"]

        first_run = first_run.replace('<w:highlight w:val="green"/>', "")
        first_run = first_run.replace("<w:highlight w:val='green'/>", "")

        escaped = (
            new_text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        first_run = re.sub(
            r"<w:t[^>]*>[\s\S]*?</w:t>",
            f'<w:t xml:space="preserve">{escaped}</w:t>',
            first_run,
        )

        result = result[: g["start_pos"]] + first_run + result[g["end_pos"] :]

    return result


def extract_from_url(url: str, groups: list[dict], api_key: str) -> list[str]:
    """Zavolá Claude API s web search a vrátí navržené náhrady."""
    client = anthropic.Anthropic(api_key=api_key)

    field_list = "\n".join(
        [
            f'Pole {i+1}: "{g["text"][:150]}{"..." if len(g["text"]) > 150 else ""}"'
            for i, g in enumerate(groups)
        ]
    )

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=3000,
        tools=[{"type": "web_search_20250305", "name": "web_search"}],
        messages=[
            {
                "role": "user",
                "content": f"""Vyhledej a přečti tuto webovou stránku o školení: {url}

V šabloně Word dokumentu jsou tato zeleně zvýrazněná místa k aktualizaci:
{field_list}

Na základě informací z webové stránky navrhni nový obsah pro každé pole.
Pokyny:
- Název školení: použij přesný název ze stránky
- Datum/termín: použij termín ze stránky včetně roku
- Cena: použij cenu ze stránky včetně DPH informace
- Jméno lektora: přesné jméno ze stránky
- Profil/bio lektora: zkopíruj nebo shrň profil lektora
- Obsah školení: zkopíruj obsah/program ze stránky
- Pro koho je školení: pokud je na stránce, zkopíruj; pokud ne, odvoď z tématu
- Zachovej původní délku a styl textu

Odpověz POUZE JSON polem o {len(groups)} prvcích (stringy v pořadí polí výše).
Žádný jiný text, žádné markdown backticky.""",
            }
        ],
    )

    text_content = "".join(
        block.text for block in response.content if hasattr(block, "text")
    )

    json_match = re.search(r"\[[\s\S]*\]", text_content)
    if json_match:
        try:
            result = json.loads(json_match.group())
            while len(result) < len(groups):
                result.append(groups[len(result)]["text"])
            return result
        except json.JSONDecodeError:
            pass

    return [g["text"] for g in groups]


# ── API klíč ─────────────────────────────────────────────────────────────────

api_key = st.secrets.get("ANTHROPIC_API_KEY", "") if hasattr(st, "secrets") else ""
if not api_key:
    api_key = st.text_input(
        "Anthropic API klíč",
        type="password",
        help="Váš API klíč z console.anthropic.com. Nebo ho uložte do Streamlit secrets jako ANTHROPIC_API_KEY.",
        placeholder="sk-ant-...",
    )

st.divider()

# ── Krok 1: Vstup ────────────────────────────────────────────────────────────

col1, col2 = st.columns([3, 2])
with col1:
    url = st.text_input(
        "URL stránky školení",
        placeholder="https://studiow.cz/...",
        help="Stránka, ze které se mají vytáhnout informace o školení.",
    )
with col2:
    template_file = st.file_uploader(
        "Word šablona (.docx)",
        type=["docx"],
        help="Soubor se zeleně zvýrazněnými poli.",
    )

process_btn = st.button(
    "🔍 Zpracovat stránku",
    type="primary",
    disabled=not (url and template_file and api_key),
    use_container_width=True,
)

# ── Krok 2: Zpracování ───────────────────────────────────────────────────────

if process_btn:
    with st.spinner("Čtu šablonu..."):
        raw = template_file.read()
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            xml_bytes = zf.read("word/document.xml")
        xml_str = xml_bytes.decode("utf-8")
        groups = find_green_groups(xml_str)

    if not groups:
        st.error(
            "V šabloně nebyly nalezeny žádné **zeleně zvýrazněné** oblasti. "
            "Zkontrolujte, zda je text označen zeleným highlightem (ne barvou písma)."
        )
        st.stop()

    st.success(f"Nalezeno **{len(groups)} zelených polí** v šabloně.")

    with st.spinner(f"Načítám stránku školení a extrahuji data..."):
        try:
            replacements = extract_from_url(url, groups, api_key)
        except Exception as e:
            st.error(f"Chyba při volání API: {e}")
            st.stop()

    st.session_state["groups"] = groups
    st.session_state["replacements"] = replacements
    st.session_state["xml_str"] = xml_str
    st.session_state["raw_zip"] = raw

# ── Krok 3: Kontrola a úprava ────────────────────────────────────────────────

if "groups" in st.session_state:
    st.divider()
    st.subheader("✏️ Zkontrolujte a upravte navržená pole")
    st.caption("Claude navrhl následující náhrady. Upravte text dle potřeby, pak klikněte na Stáhnout.")

    groups = st.session_state["groups"]
    replacements = list(st.session_state["replacements"])
    edited = []

    for i, g in enumerate(groups):
        with st.expander(f"Pole {i+1} — {g['text'][:60]}{'...' if len(g['text']) > 60 else ''}", expanded=True):
            st.markdown(
                f"**Původní text v šabloně:**\n\n> {g['text'][:300]}{'...' if len(g['text']) > 300 else ''}",
            )
            new_val = st.text_area(
                "Nový text",
                value=replacements[i] if i < len(replacements) else g["text"],
                key=f"field_{i}",
                height=120,
            )
            edited.append(new_val)

    st.divider()

    if st.button("⬇️ Stáhnout upravený dokument", type="primary", use_container_width=True):
        xml_str = st.session_state["xml_str"]
        raw_zip = st.session_state["raw_zip"]

        modified_xml = apply_replacements(xml_str, groups, edited)

        out_buf = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(raw_zip)) as src_zip:
            with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as dst_zip:
                for item in src_zip.infolist():
                    if item.filename == "word/document.xml":
                        dst_zip.writestr(item, modified_xml.encode("utf-8"))
                    else:
                        dst_zip.writestr(item, src_zip.read(item.filename))

        st.download_button(
            label="📥 Klikněte zde pro stažení",
            data=out_buf.getvalue(),
            file_name="registracni_stranka.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        st.success("Dokument je připraven ke stažení!")
