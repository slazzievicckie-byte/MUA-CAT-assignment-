from pathlib import Path

try:
    from pdf2docx import Converter
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "Missing dependency: pdf2docx. Install it with: pip install pdf2docx"
    ) from exc


def convert_pdf_to_docx(pdf_path: Path) -> None:
    docx_path = pdf_path.with_suffix(".docx")

    if docx_path.exists():
        print(f"Skipping: {pdf_path.name} (already converted to {docx_path.name})")
        return

    print(f"Converting: {pdf_path.name} -> {docx_path.name}")
    cv = Converter(str(pdf_path))
    cv.convert(str(docx_path))
    cv.close()


def main() -> None:
    repo_root = Path(__file__).resolve().parent
    pdf_files = sorted(repo_root.glob("*.pdf"))

    if not pdf_files:
        print("No PDF files found in this directory.")
        return

    for pdf_file in pdf_files:
        convert_pdf_to_docx(pdf_file)

    print(f"Done. Converted {len(pdf_files)} PDF file(s) to DOCX.")


if __name__ == "__main__":
    main()
