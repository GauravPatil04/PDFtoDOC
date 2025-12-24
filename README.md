# PDF → DOCX Converter (Streamlit)

A simple Streamlit app to convert PDF files to DOCX either by preserving editable text (via `pdf2docx`) or by embedding page images for exact visual fidelity.

## Quick Start

1. Install dependencies (Windows, PowerShell):

```bash
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

2. Run the app:

```bash
streamlit run app.py
```

3. In the browser UI:
- Upload a PDF.
- Choose a conversion mode:
  - Preserve editable text (recommended)
  - Exact layout as images (fallback)
- Click Convert, then download the DOCX.

## Notes
- Text mode keeps text editable and attempts to preserve structure/layout. Some highly formatted PDFs may convert imperfectly; use the fallback mode for an exact visual match.
- Fallback mode places each page as an image inside the DOCX, ensuring identical appearance but making the content non-editable.
- Page range selection is available for text mode.
- Large PDFs or those with many high-resolution pages may take longer and use more memory.
