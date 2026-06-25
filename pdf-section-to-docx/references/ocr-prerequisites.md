# OCR Prerequisites

## Runtime requirement

This skill's script relies on:

- Python packages already imported by the script: `PyMuPDF` (`fitz`) and `python-docx`
- `tesseract` on the system `PATH` for scanned or image-only PDF pages

## Install Tesseract on Windows

If `tesseract` is missing, install Tesseract OCR and ensure the `tesseract.exe` directory is on `PATH`.

After installation, verify with:

```powershell
tesseract --version
```

## Why the script fails closed

The user requirement is "no omission." For pages that appear image-only, the script does not silently skip OCR.
It stops with a clear error if `tesseract` is unavailable.

## OCR language packs

Default OCR language is `chi_sim+eng`.

Change it when needed:

```powershell
python .\scripts\pdf_section_to_docx.py .\sample.pdf --ocr-lang eng
python .\scripts\pdf_section_to_docx.py .\sample.pdf --ocr-lang chi_sim+eng
```
