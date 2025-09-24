
import io
import os
import re
import zipfile
import tempfile
from pathlib import Path
from typing import Dict, Any, List, Optional

import streamlit as st
import yaml

from processor import (
    build_default_template,
    compose_into_template,
    apply_style_mappings,
    apply_find_replace,
    set_header_footer,
    apply_page_setup,
)

st.set_page_config(page_title="DOCX Layout Refitter", page_icon="🧩", layout="wide")

st.title("🧩 DOCX Layout Refitter (no AI)")
st.caption("Batch‑apply a new layout to legacy Word documents—securely and deterministically, no generative AI involved.")

with st.expander("How this works", expanded=False):
    st.markdown("""
    - **Input**: Your legacy `.docx` files (individual files or a `.zip`).
    - **Template**: A **new‑layout** Word template (`.docx`) that defines page size, margins, header/footer, and styles.
    - **Process**: For each legacy document, the app **composes** its full content **into the template** (preserving images, tables, lists), 
      then applies optional **style mappings** and **find/replace** rules.
    - **Output**: A `.zip` containing all refitted `.docx` files with the new layout.
    """)

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1) New layout template (.docx)")
    uploaded_template = st.file_uploader("Upload your new-layout template (optional)", type=["docx"], key="tpl")
    use_default = st.checkbox("Use bundled default template", value=uploaded_template is None, help="If checked, a clean default template is used.")

    st.subheader("2) Config (optional)")
    cfg_file = st.file_uploader("Upload YAML config for mappings/replacements", type=["yml", "yaml"], key="cfg")
    if st.button("Load sample config into editor"):
        st.session_state["cfg_text"] = (
            "# Sample config\n"
            "page_setup:\n"
            "  orientation: portrait   # portrait|landscape\n"
            "  margins_mm: {top: 20, right: 15, bottom: 20, left: 25}\n"
            "header_footer:\n"
            "  header_text: \"Confidential — New Layout\"\n"
            "  footer_text: \"© Your Company\"\n"
            "  include_page_numbers: true\n"
            "style_map:\n"
            "  \"Heading 1\": \"Title\"\n"
            "  \"Heading 2\": \"Heading 1\"\n"
            "  \"Heading 3\": \"Heading 2\"\n"
            "find_replace:\n"
            "  - pattern: \"\\bACME Corp\\b\"\n"
            "    replace: \"Your Company\"\n"
            "  - pattern: \"\\s{2,}\"\n"
            "    replace: \" \"\n"
        )
    cfg_text = st.text_area("Edit config (YAML)", value=st.session_state.get("cfg_text", ""), height=240, help="Optional. If provided, this overrides the uploaded YAML file.")

with col2:
    st.subheader("3) Legacy documents")
    files = st.file_uploader(
        "Upload one or more .docx files, or a .zip containing .docx files",
        type=["docx", "zip"], accept_multiple_files=True, key="docs"
    )
    st.markdown("You can **drag & drop** a large set of files or a single `.zip` with up to thousands of documents.")

    st.subheader("4) Actions")
    preview_btn = st.button("🔍 Preview first document")
    run_btn = st.button("▶️ Process all")

# Parse / load config
def load_config() -> Dict[str, Any]:
    if cfg_text.strip():
        try:
            return yaml.safe_load(cfg_text) or {}
        except Exception as e:
            st.error(f"Invalid YAML in editor: {e}")
            return {}
    if cfg_file is not None:
        try:
            return yaml.safe_load(cfg_file.getvalue()) or {}
        except Exception as e:
            st.error(f"Invalid YAML file: {e}")
            return {}
    return {}

cfg: Dict[str, Any] = load_config()

# Prepare template
def get_template_bytes() -> bytes:
    if not use_default and uploaded_template is not None:
        return uploaded_template.getvalue()
    # Build a clean default template
    return build_default_template()

def iter_input_docs(uploaded_items) -> List[tuple[str, bytes]]:
    """Return list of (filename, bytes) for .docx to process, expanding any uploaded zip files."""
    results = []
    for item in uploaded_items or []:
        name = item.name.lower()
        if name.endswith(".docx"):
            results.append((item.name, item.getvalue()))
        elif name.endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(item.getvalue())) as zf:
                for zi in zf.infolist():
                    if not zi.is_dir() and zi.filename.lower().endswith(".docx"):
                        results.append((Path(zi.filename).name, zf.read(zi)))
    return results

def process_one(name: str, data: bytes, tpl_bytes: bytes, cfg: Dict[str, Any]) -> bytes:
    """Return processed .docx bytes for a single input document."""
    # Step 1: compose legacy content into template (preserves images, tables, lists)
    composed = compose_into_template(data, tpl_bytes)

    # Step 2: apply page setup/header/footer on the composed doc
    if "page_setup" in cfg and isinstance(cfg["page_setup"], dict):
        composed = apply_page_setup(composed, cfg["page_setup"])
    if "header_footer" in cfg and isinstance(cfg["header_footer"], dict):
        composed = set_header_footer(composed, cfg["header_footer"])

    # Step 3: style mappings and find/replace
    if "style_map" in cfg and isinstance(cfg["style_map"], dict):
        composed = apply_style_mappings(composed, cfg["style_map"])
    if "find_replace" in cfg and isinstance(cfg["find_replace"], list):
        composed = apply_find_replace(composed, cfg["find_replace"])

    return composed

# Buttons behavior
docs = iter_input_docs(files)

if preview_btn:
    if not docs:
        st.warning("Please upload at least one .docx or a zip of .docx files.")
    else:
        tpl_bytes = get_template_bytes()
        name, data = docs[0]
        with st.spinner(f"Previewing {name}…"):
            out_bytes = process_one(name, data, tpl_bytes, cfg)
        st.success("Preview ready.")
        st.download_button("Download preview .docx", data=out_bytes, file_name=f"PREVIEW_{Path(name).stem}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if run_btn:
    if not docs:
        st.warning("Please upload at least one .docx or a zip of .docx files.")
    else:
        progress = st.progress(0.0, text="Processing…")
        tpl_bytes = get_template_bytes()

        out_zip_bytes = io.BytesIO()
        with zipfile.ZipFile(out_zip_bytes, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for i, (name, data) in enumerate(docs, start=1):
                try:
                    out_bytes = process_one(name, data, tpl_bytes, cfg)
                    out_name = f"{Path(name).stem}_refit.docx"
                    zf.writestr(out_name, out_bytes)
                except Exception as e:
                    # Include an error marker file for this item
                    zf.writestr(f"{Path(name).stem}__ERROR.txt", f"Failed to process {name}: {e}")
                progress.progress(i / len(docs), text=f"Processed {i} / {len(docs)}")

        progress.progress(1.0, text="Done.")
        st.success(f"Processed {len(docs)} file(s).")
        st.download_button(
            "⬇️ Download all as .zip",
            data=out_zip_bytes.getvalue(),
            file_name="refitted_documents.zip",
            mime="application/zip",
        )

st.markdown("---")
st.caption("Tip: For the best confidentiality, run this app **locally** or in your own infrastructure. No external AI or network calls are used.")
