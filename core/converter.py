from pathlib import Path

import pikepdf

from core.hashing import sha256_file

METADATA_KEY = "/DocToPDFSourceHash"


def convert(docx_path: str) -> str:
    """Convert a .docx file to PDF, embedding a SHA-256 hash of the source."""
    docx_path = Path(docx_path)
    pdf_path = docx_path.with_suffix(".pdf")

    source_hash = sha256_file(str(docx_path))

    # Use win32com directly for more control than docx2pdf wrapper
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "Microsoft Word is required for conversion.\n"
            "Please install the desktop version of Microsoft Word and try again."
        )

    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except Exception:
            raise RuntimeError(
                "Could not connect to Microsoft Word.\n\n"
                "Please ensure the desktop version of Microsoft Word is installed.\n"
                "(The Microsoft Store version is not supported.)"
            )
        word.Visible = False
        word.DisplayAlerts = False
        try:
            doc = word.Documents.Open(
                str(docx_path.resolve()),
                ReadOnly=True,
                AddToRecentFiles=False,
            )
            doc.SaveAs2(
                str(pdf_path.resolve()),
                FileFormat=17,  # wdFormatPDF
            )
            doc.Close(SaveChanges=False)
        finally:
            word.Quit()
    finally:
        pythoncom.CoUninitialize()

    # Embed the source hash into PDF metadata
    with pikepdf.open(str(pdf_path), allow_overwriting_input=True) as pdf:
        # Write to Document Info dictionary
        pdf.docinfo[METADATA_KEY] = source_hash
        # Also write to XMP metadata
        with pdf.open_metadata() as meta:
            meta["pdfx:DocToPDFSourceHash"] = source_hash
        pdf.save(str(pdf_path))

    return str(pdf_path)
