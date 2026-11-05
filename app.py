import json
from pathlib import Path

import pandas as pd
import streamlit as st

from extractors import TESSERACT_OK, extract_document

SAMPLES = Path(__file__).parent / "samples"
FILE_TYPES = ["pdf", "docx", "txt", "html", "htm", "png", "jpg", "jpeg", "tif", "tiff", "bmp"]

st.set_page_config(page_title="Clinical Document Ingestion", page_icon="", layout="wide")

st.markdown(
    """
    <style>
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #2bb5da 0%, #4a9ad9 100%);
    }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { font-weight: 800; margin-top: 1.2rem; }
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p { color: #10222e; }
    .main-title { text-align: center; font-size: 2.6rem; font-weight: 800; margin: 2.5rem 0 0.2rem 0; }
    .sub-title { text-align: center; color: #3b4a56; font-size: 1.1rem; margin-bottom: 2.2rem; }
    .file-line { color: #5a6670; margin-bottom: 0.6rem; }
    button[data-baseweb="tab"] { font-size: 1rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner="Reading files...")
def run(name, data, ocr, dpi, lang):
    return extract_document(name, data, ocr=ocr, dpi=dpi, lang=lang).to_dict()


def show_table(rows):
    st.dataframe(pd.DataFrame(rows), hide_index=True)


def key_value_rows(values):
    return [{"Field": k, "Value": json.dumps(v) if isinstance(v, (list, dict)) else str(v)} for k, v in values.items()]


with st.sidebar:
    st.header("Settings")
    mode = st.radio("Where from?", ["Sample files", "Upload"])
    files = []
    if mode == "Sample files":
        names = sorted(p.name for p in SAMPLES.iterdir()) if SAMPLES.exists() else []
        picked = st.multiselect("Files", names, default=names)
        files = [(n, (SAMPLES / n).read_bytes()) for n in picked]
    else:
        uploads = st.file_uploader("Files", type=FILE_TYPES, accept_multiple_files=True)
        files = [(u.name, u.getvalue()) for u in uploads]

    st.header("OCR")
    use_ocr = st.checkbox("OCR scanned PDF pages", value=True)
    dpi = st.select_slider("Scan quality (DPI)", options=[150, 200, 300, 400], value=200)
    lang = st.text_input("OCR language", value="eng")
    if not TESSERACT_OK:
        st.warning("Tesseract not found. Scanned pages and images can't be read.")

    results = []
    if files:
        results = [run(n, d, use_ocr, dpi, lang.strip() or "eng") for n, d in files]
        st.header("View")
        labels = [r["source"]["file_name"] for r in results]
        current = results[labels.index(st.selectbox("Open file", labels))]

st.markdown('<div class="main-title"> Clinical Document Ingestion</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Drop in any clinical file and get clean text with its source details.</div>',
    unsafe_allow_html=True,
)

if not files:
    st.info("Choose a sample file or upload your own from the sidebar.")
    st.stop()

stats = current["stats"]
page_note = f"{stats['pages']} pages" if stats["pages"] else "no page count"
st.markdown(
    f'<div class="file-line">{current["source"]["file_name"]} - {current["type"]}, {page_note}</div>',
    unsafe_allow_html=True,
)

tab_summary, tab_text, tab_structure, tab_meta, tab_export = st.tabs(
    ["Summary", "Text", "Structure", "Metadata", "Export"]
)

with tab_summary:
    st.subheader("Counts for each file")
    show_table(
        [
            {
                "File": r["source"]["file_name"],
                "Type": r["type"],
                "Method": r["method"],
                "Pages": r["stats"]["pages"],
                "Blocks": r["stats"]["blocks"],
                "Words": r["stats"]["words"],
                "OCR pages": r["stats"]["ocr_pages"],
                "OCR confidence": r["stats"]["avg_ocr_confidence"],
                "Seconds": r["stats"]["seconds"],
                "Warnings": len(r["warnings"]),
            }
            for r in results
        ]
    )
    st.caption("OCR confidence runs from 0 to 1. Check the text by eye when it is low.")

    st.subheader("Problems found")
    if current["warnings"]:
        for w in current["warnings"]:
            st.warning(w)
    else:
        st.success("No problems found in this file.")

with tab_text:
    st.subheader("Extracted text")
    text = current["text"]
    if stats["pages"] and stats["pages"] > 1:
        choice = st.selectbox("Page", ["All pages"] + list(range(1, stats["pages"] + 1)))
        if choice != "All pages":
            text = "\n\n".join(b["text"] for b in current["blocks"] if b.get("page") == choice)
    st.text_area("Text", text, height=420, label_visibility="collapsed")
    st.download_button("Download as .txt", text, file_name=f"{Path(current['source']['file_name']).stem}.txt")

with tab_structure:
    st.subheader("Blocks")
    kinds = sorted({b["kind"] for b in current["blocks"]})
    chosen = st.multiselect("Show", kinds, default=kinds)
    rows = [
        {
            "Page": b.get("page"),
            "Kind": b["kind"],
            "Level": b.get("level"),
            "Source": b["source"],
            "Confidence": b.get("confidence"),
            "Text": b["text"],
        }
        for b in current["blocks"]
        if b["kind"] in chosen
    ]
    show_table(rows)
    tables = [b for b in current["blocks"] if b["kind"] == "table"]
    for i, t in enumerate(tables, 1):
        with st.expander(f"Table {i}"):
            st.dataframe(pd.DataFrame(t["rows"]), hide_index=True)

with tab_meta:
    st.subheader("Source")
    show_table(key_value_rows(current["source"]))
    st.subheader("Document properties")
    if current["metadata"]:
        show_table(key_value_rows(current["metadata"]))
    else:
        st.caption("This file has no embedded properties.")

with tab_export:
    st.subheader("Standard output")
    st.caption("Every format ends up in the same shape: source, type, metadata, stats, warnings and blocks.")
    preview = dict(current)
    preview["blocks"] = current["blocks"][:5]
    preview["text"] = current["text"][:300]
    st.json(preview, expanded=1)
    name = Path(current["source"]["file_name"]).stem
    left, right = st.columns(2)
    left.download_button(
        "Download this file as JSON",
        json.dumps(current, indent=2, ensure_ascii=False),
        file_name=f"{name}.json",
        mime="application/json",
    )
    right.download_button(
        "Download all files as JSON",
        json.dumps(results, indent=2, ensure_ascii=False),
        file_name="all_documents.json",
        mime="application/json",
    )
