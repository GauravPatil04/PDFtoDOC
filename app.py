import streamlit as st
import tempfile
import os
from io import BytesIO

# Editable-text conversion
from pdf2docx import Converter

# Exact-layout fallback (pages as images)
import fitz  # PyMuPDF
from docx import Document


def convert_pdf_to_docx_text(pdf_path: str, output_path: str, start: int | None = None, end: int | None = None) -> None:
    """Convert PDF to DOCX using pdf2docx preserving editable text/layout."""
    cv = Converter(pdf_path)
    try:
        cv.convert(output_path, start=start, end=end)
    finally:
        cv.close()


def convert_pdf_to_docx_images(pdf_path: str) -> bytes:
    """Convert PDF pages to images and place them into DOCX for exact visual fidelity.
    Returns DOCX bytes.
    """
    docx = Document()

    # Use default section to compute available content width
    section = docx.sections[0]
    available_width = section.page_width - section.left_margin - section.right_margin

    pdf_doc = fitz.open(pdf_path)
    try:
        for i in range(pdf_doc.page_count):
            page = pdf_doc.load_page(i)
            # Render at 2x scale for sharper text
            zoom = 2.0
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)

            # Save image to memory
            img_bytes = pix.tobytes("png")
            bio = BytesIO(img_bytes)

            # Add image to DOCX and fit within page width
            docx.add_picture(bio, width=available_width)

            # Page break except after last page
            if i < pdf_doc.page_count - 1:
                docx.add_page_break()
    finally:
        pdf_doc.close()

    out = BytesIO()
    docx.save(out)
    out.seek(0)
    return out.read()


def main():
    st.set_page_config(page_title="PDF → DOCX Converter", page_icon="📄", layout="centered")
    st.title("PDF → DOCX Converter")
    st.caption("Convert PDF to DOCX preserving layout. Choose text or exact-image mode.")

    uploaded = st.file_uploader("Upload a PDF", type=["pdf"], accept_multiple_files=False)

    mode = st.radio(
        "Conversion mode",
        options=["Preserve editable text (recommended)", "Exact layout as images (fallback)"],
        index=0,
        help=(
            "Text mode attempts to keep text editable and structure preserved. "
            "Fallback image mode guarantees the same look by embedding page images."
        ),
    )

    # Optional: page range for text mode
    start_page = None
    end_page = None
    with st.expander("Optional settings"):
        if uploaded is not None:
            # Determine total pages for guidance
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in:
                tmp_in.write(uploaded.getvalue())
                tmp_in_path = tmp_in.name
            try:
                total_pages = fitz.open(tmp_in_path).page_count
            except Exception:
                total_pages = None
            finally:
                try:
                    os.unlink(tmp_in_path)
                except Exception:
                    pass

            st.write(f"Pages in PDF: {total_pages if total_pages is not None else 'Unknown'}")
        col1, col2 = st.columns(2)
        with col1:
            start_page = st.number_input("Start page (1-based)", min_value=1, value=1, step=1)
        with col2:
            end_page = st.number_input("End page (leave 0 for all)", min_value=0, value=0, step=1)
        if end_page == 0:
            end_page = None

    convert_btn = st.button("Convert", type="primary", disabled=(uploaded is None))

    if convert_btn and uploaded is not None:
        with st.spinner("Converting… This may take a moment."):
            # Write uploaded PDF to a temp file for library compatibility
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf.write(uploaded.getvalue())
                tmp_pdf_path = tmp_pdf.name

            file_stem = os.path.splitext(uploaded.name)[0]

            try:
                if mode.startswith("Preserve editable text"):
                    # pdf2docx requires a filesystem path for output
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                        out_path = tmp_docx.name

                    # Sanitize page range
                    s = int(start_page) if start_page and start_page >= 1 else None
                    e = int(end_page) if end_page and (end_page == 0 or end_page >= 1) else None

                    convert_pdf_to_docx_text(tmp_pdf_path, out_path, start=s, end=e)

                    # Read the DOCX content back into memory
                    with open(out_path, "rb") as f:
                        docx_bytes = f.read()

                    # Offer download
                    st.success("Conversion complete.")
                    st.download_button(
                        label="Download DOCX",
                        data=docx_bytes,
                        file_name=f"{file_stem}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                else:
                    # Exact layout via images embedded into DOCX
                    docx_bytes = convert_pdf_to_docx_images(tmp_pdf_path)
                    st.success("Conversion complete (image-based).")
                    st.download_button(
                        label="Download DOCX",
                        data=docx_bytes,
                        file_name=f"{file_stem}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
            except Exception as exc:
                st.error(f"Conversion failed: {exc}")
            finally:
                # Clean up temp files
                try:
                    os.unlink(tmp_pdf_path)
                except Exception:
                    pass
                try:
                    if mode.startswith("Preserve editable text"):
                        os.unlink(out_path)
                except Exception:
                    pass

    st.divider()
    st.markdown(
        """
        - Text mode uses `pdf2docx` to preserve editable text and layout.
        - Fallback mode embeds page images for exact visual fidelity when text mode struggles.
        - Large or graphically complex PDFs may take longer to process.
        """
    )


if __name__ == "__main__":
    main()
