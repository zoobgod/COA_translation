"""
Pharmaceutical COA Translator — Streamlit Application

Upload a pharmaceutical Certificate of Analysis (COA) in PDF format,
translate it to Russian using OpenAI with a pharmaceutical glossary,
and download the result as a formatted Word document.
"""

import streamlit as st

from modules.pdf_extractor import extract_text_from_pdf
from modules.translator import translate_text
from modules.doc_generator import generate_doc_from_template

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="COA Translator — EN → RU",
    page_icon="💊",
    layout="centered",
)

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .main-header {
        text-align: center;
        padding: 1rem 0;
    }
    .main-header h1 {
        color: #1E88E5;
        font-size: 2rem;
    }
    .status-box {
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .info-metric {
        background-color: #f0f2f6;
        padding: 0.75rem;
        border-radius: 0.5rem;
        text-align: center;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="main-header">'
    "<h1>Pharmaceutical COA Translator</h1>"
    "<p>English → Russian | Certificate of Analysis</p>"
    "</div>",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Sidebar — Settings
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Settings")

    api_key = st.text_input(
        "OpenAI API Key",
        type="password",
        help="Enter your OpenAI API key. It is used only for the current session and never stored.",
    )

    model_choice = st.selectbox(
        "Translation Model",
        options=["gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-3.5-turbo"],
        index=0,
        help="gpt-4o recommended for best quality pharmaceutical translations.",
    )

    st.divider()
    st.markdown("**About**")
    st.markdown(
        "This app translates pharmaceutical Certificate of Analysis (COA) "
        "documents from English to Russian using AI with a specialized "
        "pharmaceutical glossary."
    )
    st.markdown(
        "**Features:**\n"
        "- Multi-method PDF extraction\n"
        "- OCR fallback for scanned docs\n"
        "- 200+ pharma term glossary\n"
        "- Formatted Word output"
    )

# ---------------------------------------------------------------------------
# Main content
# ---------------------------------------------------------------------------
st.subheader("1. Upload COA PDF")

uploaded_file = st.file_uploader(
    "Upload a Certificate of Analysis in PDF format",
    type=["pdf"],
    help="Supports text-based and scanned (OCR) PDFs up to 50 MB.",
)

if uploaded_file is not None:
    pdf_bytes = uploaded_file.read()
    file_size_mb = len(pdf_bytes) / (1024 * 1024)

    col1, col2 = st.columns(2)
    with col1:
        st.metric("File", uploaded_file.name)
    with col2:
        st.metric("Size", f"{file_size_mb:.2f} MB")

    # ------------------------------------------------------------------
    # Step 2: Extract text
    # ------------------------------------------------------------------
    st.subheader("2. Extract Text")

    if "extraction_result" not in st.session_state or st.session_state.get(
        "last_file"
    ) != uploaded_file.name:
        with st.spinner("Extracting text from PDF..."):
            extraction = extract_text_from_pdf(pdf_bytes)
            st.session_state["extraction_result"] = extraction
            st.session_state["last_file"] = uploaded_file.name
    else:
        extraction = st.session_state["extraction_result"]

    if extraction["success"]:
        st.success(
            f"Text extracted successfully using **{extraction['method']}** "
            f"({extraction['page_count']} page(s))"
        )

        with st.expander("Preview extracted text", expanded=False):
            preview = extraction["text"][:3000]
            if len(extraction["text"]) > 3000:
                preview += "\n\n... [truncated for preview]"
            st.text(preview)

        # ------------------------------------------------------------------
        # Step 3: Translate
        # ------------------------------------------------------------------
        st.subheader("3. Translate to Russian")

        if not api_key:
            st.warning("Please enter your OpenAI API key in the sidebar to proceed.")
        else:
            translate_btn = st.button(
                "Translate to Russian",
                type="primary",
                use_container_width=True,
            )

            if translate_btn or st.session_state.get("translation_result"):
                if translate_btn:
                    progress_bar = st.progress(0, text="Translating...")
                    status_text = st.empty()

                    def update_progress(current, total):
                        pct = int(current / total * 100)
                        progress_bar.progress(
                            pct,
                            text=f"Translating chunk {current}/{total}...",
                        )
                        status_text.text(
                            f"Processing chunk {current} of {total}"
                        )

                    with st.spinner("Translating document..."):
                        result = translate_text(
                            text=extraction["text"],
                            api_key=api_key,
                            model=model_choice,
                            progress_callback=update_progress,
                        )

                    progress_bar.empty()
                    status_text.empty()

                    st.session_state["translation_result"] = result
                else:
                    result = st.session_state["translation_result"]

                if result["success"]:
                    st.success(
                        f"Translation complete! Model: **{result['model_used']}**, "
                        f"Chunks: **{result['chunks_translated']}**"
                    )

                    with st.expander("Preview translated text", expanded=False):
                        preview_ru = result["translated_text"][:3000]
                        if len(result["translated_text"]) > 3000:
                            preview_ru += "\n\n... [сокращено для предварительного просмотра]"
                        st.text(preview_ru)

                    # ----------------------------------------------------------
                    # Step 4: Generate & Download Word doc
                    # ----------------------------------------------------------
                    st.subheader("4. Download Word Document")

                    if "doc_bytes" not in st.session_state or translate_btn:
                        with st.spinner("Generating Word document..."):
                            doc_bytes = generate_doc_from_template(
                                translated_text=result["translated_text"],
                                original_filename=uploaded_file.name,
                                extraction_method=extraction["method"],
                                model_used=result["model_used"],
                            )
                            st.session_state["doc_bytes"] = doc_bytes
                    else:
                        doc_bytes = st.session_state["doc_bytes"]

                    # Output filename
                    base_name = uploaded_file.name.rsplit(".", 1)[0]
                    output_filename = f"{base_name}_RU.docx"

                    st.download_button(
                        label="Download Translated COA (.docx)",
                        data=doc_bytes,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True,
                    )

                    st.info(
                        "The translated document is ready for download. "
                        "We recommend having a qualified pharmaceutical specialist "
                        "review the translation before official use."
                    )

                else:
                    st.error(f"Translation failed: {result['error']}")

                    if "api_key" in str(result["error"]).lower() or "auth" in str(
                        result["error"]
                    ).lower():
                        st.warning(
                            "This looks like an authentication error. "
                            "Please check your OpenAI API key."
                        )

    else:
        st.error(
            "Could not extract text from the PDF. "
            "The file may be corrupted or empty."
        )
        st.info(
            "Tips:\n"
            "- Ensure the PDF is not password-protected\n"
            "- Try a different PDF if the file seems corrupted\n"
            "- For scanned documents, ensure pytesseract is installed"
        )

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.divider()
st.markdown(
    "<div style='text-align: center; color: #888; font-size: 0.8rem;'>"
    "COA Translator v1.0 | Pharmaceutical glossary with 200+ terms | "
    "Powered by OpenAI"
    "</div>",
    unsafe_allow_html=True,
)
