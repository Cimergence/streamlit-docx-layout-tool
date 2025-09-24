
# 🧩 DOCX Layout Refitter (Streamlit, no AI)

Batch‑apply a **new Word layout** to hundreds or thousands of legacy `.docx` files — **without any generative AI**.  
Images, tables, and lists are preserved by composing each legacy document **into your template**.

## Why this tool?
- **Confidential by design:** No AI, no external calls. Run locally or on your own infra.
- **Deterministic:** Strict rules (style mappings, find/replace) instead of opaque AI transformations.
- **Fast at scale:** Drop in a `.zip` with thousands of docs; get a `.zip` out with the refitted versions.

---

## Quick start

```bash
# 1) Create & activate a virtual env (recommended)
python -m venv .venv
source .venv/bin/activate    # Windows: .venv\Scripts\activate

# 2) Install dependencies
pip install -r requirements.txt

# 3) Run the app
streamlit run app.py
```

Open the local URL shown in the terminal.

> **Tip:** For best confidentiality, run on an offline machine if needed.

---

## How it works

1. **Template** – Provide your **new-layout** `.docx` (or use the bundled default).  
   This controls margins, header/footer, orientation, and styles.
2. **Compose** – Each legacy `.docx` is **appended into the template** (via `docxcompose`), preserving **images, tables, and lists**.
3. **Normalize** – Optional **style mappings** and **regex find/replace** rules are applied.
4. **Export** – Download a `.zip` containing all refitted documents.

---

## Configuration (YAML)

Example (`sample_layout_config.yaml`):

```yaml
page_setup:
  orientation: portrait          # portrait | landscape
  margins_mm: { top: 20, right: 15, bottom: 20, left: 25 }

header_footer:
  header_text: "Confidential — New Layout"
  footer_text: "© Your Company"
  include_page_numbers: true

style_map:
  "Heading 1": "Title"
  "Heading 2": "Heading 1"
  "Heading 3": "Heading 2"

find_replace:
  - pattern: "\\bACME Corp\\b"
    replace: "Your Company"
  - pattern: "\\s{2,}"
    replace: " "
```

- **`style_map`** remaps paragraph styles by name. Styles must exist in your template.
- **`find_replace`** is regex-based and operates on paragraph text (best‑effort across runs).

> ⚠️ **Note on formatting:** Regex replacements may simplify run‑level formatting within a paragraph (e.g., mixed bold/italic inside a line). If this matters for specific patterns, prefer adjusting the **template styles** over text‑level replacements.

---

## Inputs & outputs

- **Inputs**
  - One or more `.docx` files **or** a `.zip` of `.docx`.
  - Optional: a `.docx` **template**.
  - Optional: a **YAML config** file (or paste directly in the app).

- **Outputs**
  - A `.zip` with `_refit.docx` files, one per input.

---

## Limitations & tips

- This tool focuses on **layout application**, not semantic rewriting.
- Complex per‑character formatting inside a single paragraph may not be perfectly preserved if regex replacements are applied.
- If you need strictly identical styles, ensure the **target styles** exist in your template and map the source styles accordingly.
- For best results, keep a **clean, style‑driven template**. Avoid direct formatting in the template itself; rely on styles.

---

## Running in Docker (optional)

```dockerfile
# See Dockerfile in this repo
docker build -t docx-refitter .
docker run -p 8501:8501 -v "$PWD":/app docx-refitter
```

Then open http://localhost:8501

---

## Security & privacy

- No network or AI calls.
- All processing happens in-memory per upload.
- You control the runtime (local machine, on‑prem, or private cloud).

---

## Tech stack

- [Streamlit](https://streamlit.io/) for UI
- [python-docx](https://python-docx.readthedocs.io/) for DOCX manipulation
- [docxcompose](https://github.com/python-openxml/python-docx/issues/165#issuecomment-438304665) to append documents safely
- [PyYAML](https://pyyaml.org/) for config

---

## FAQ

**Q: Does it preserve images, tables, lists?**  
**A:** Yes — these are preserved when composing into the template.

**Q: Can I apply my corporate header/footer and page numbers?**  
**A:** Yes — configure via YAML `header_footer` and `page_setup` or by using your own template.

**Q: Can I map styles automatically?**  
**A:** Yes — define `style_map`. Ensure the target styles exist in your template.

**Q: Is AI used anywhere?**  
**A:** No. This is 100% deterministic and local.

---

## License

MIT — do whatever you want, no warranty.
