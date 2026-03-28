from pathlib import Path

import pikepdf

from core.hashing import sha256_file

METADATA_KEY = "/DocToPDFSourceHash"


def convert(docx_path: str) -> str:
    """Convert a .docx file to PDF, embedding a SHA-256 hash of the source."""
    docx_path = Path(docx_path)
    pdf_path = docx_path.with_suffix(".pdf")

    # Check if the output PDF exists and is writable before starting Word
    if pdf_path.exists():
        try:
            with open(pdf_path, "a"):
                pass
        except PermissionError:
            raise PermissionError(
                f"Cannot overwrite '{pdf_path.name}'.\n"
                "The file may be open in another program."
            )

    source_hash = sha256_file(str(docx_path))

    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "Microsoft Word is required for conversion.\n"
            "Please install the desktop version of Microsoft Word and try again."
        )

    pythoncom.CoInitialize()
    word = None
    doc = None
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

        doc = word.Documents.Open(
            str(docx_path.resolve()),
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        doc.SaveAs2(
            str(pdf_path.resolve()),
            FileFormat=17,  # wdFormatPDF
        )
    finally:
        # Close document and Word in reverse order, guarding each step
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    # Embed the source hash into PDF metadata
    try:
        with pikepdf.open(str(pdf_path), allow_overwriting_input=True) as pdf:
            pdf.docinfo[METADATA_KEY] = source_hash
            with pdf.open_metadata() as meta:
                meta["pdfx:DocToPDFSourceHash"] = source_hash
            pdf.save(str(pdf_path))
    except Exception:
        # PDF was created but metadata embedding failed — still usable
        pass

    return str(pdf_path)
