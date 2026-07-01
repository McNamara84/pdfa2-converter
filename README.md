# pdfa2-converter

Batch conversion from mixed source formats in `input_files/` to PDF/A-2
documents in `output_pdfs/`.

## Quick start

```bash
python3 convert_to_pdfa2.py
```

Useful options:

```bash
python3 convert_to_pdfa2.py --overwrite
python3 convert_to_pdfa2.py --workers 4
python3 convert_to_pdfa2.py --dry-run
python3 convert_to_pdfa2.py --validate required
```

The tool writes:

- converted PDFs to `output_pdfs/`
- a per-file report to `output_pdfs/manifest.csv`
- detailed command output and tracebacks to `output_pdfs/conversion.log`

## Supported input classes

- existing PDFs
- common image formats such as JPG, PNG, TIFF, GIF, BMP and WEBP
- Office/OpenDocument files supported by LibreOffice, for example DOCX, XLSX,
  PPTX, ODT, ODS and ODP
- text-like files such as XML, TXT, CSV, JSON and Markdown

Files with unsupported extensions are not silently ignored. They are recorded
as `unsupported` in the manifest.

## Requirements

Required:

- Python 3
- Pillow for image and text rendering
- LibreOffice / `soffice` for Office documents
- Ghostscript / `gs` for PDF/A-2 normalization
- an sRGB ICC profile; common Linux paths are detected automatically

Optional but strongly recommended:

- veraPDF / `verapdf` for real PDF/A validation

The script first looks for `verapdf` on `PATH`. If it is not installed
system-wide, it also uses a local installation at `.tools/verapdf/verapdf`.
Without veraPDF the script can create PDF/A-2 candidates via Ghostscript, but
it cannot independently certify conformance. Use `--validate required` in a
production workflow.

## Important limits

This is not impossible, but it is also not a pure-Python problem. Reliable
conversion of 1.8 million heterogeneous files needs:

- format-specific converters
- PDF/A validation
- retry/error handling
- logging and manifest data
- enough CPU, disk space and process isolation for batch execution

This repository provides the local conversion pipeline and error reporting.
For a production run at 1.8 million files, run it on a controlled worker setup,
keep the manifest, and review all `failed` and `unsupported` rows.
 
