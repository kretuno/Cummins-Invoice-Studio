from __future__ import annotations

import argparse
from pathlib import Path

from exporter import export_report_to_excel
from parser import CumminsInvoiceParser
from utils import collect_pdf_files, ensure_xlsx_path, human_total_usd, human_total_weight


def main() -> None:
    args = parse_args()
    source_path = Path(args.source).expanduser().resolve()

    if not source_path.exists():
        raise SystemExit(f"Source does not exist: {source_path}")

    source_type = detect_source_type(source_path)
    pdf_files, temp_dir = collect_pdf_files(source_type, source_path)

    try:
        report = CumminsInvoiceParser().parse_files(pdf_files)
        output_path = ensure_xlsx_path(args.output) if args.output else default_output_path(source_path)
        export_report_to_excel(report, output_path)

        print("Analysis completed")
        print(f"Source: {source_path}")
        print(f"PDF files: {len(pdf_files)}")
        print(f"Invoices: {len(report.invoices)}")
        print(f"Total net weight: {human_total_weight(report.total_weight)}")
        print(f"Total gross weight: {human_total_weight(report.total_gross_weight)}")
        print(f"Total places: {report.total_packages}")
        print(f"Total USD: {human_total_usd(report.total_usd)}")
        print(f"Duplicates excluded: {len(report.duplicates)}")
        print(f"Notes/issues: {len(report.issues)}")
        print(f"Excel saved: {output_path}")

        if report.issues:
            print("\nNotes:")
            for issue in report.issues[:20]:
                print(f"- [{issue.level}] {issue.source_file}: {issue.message}")
            if len(report.issues) > 20:
                print(f"- ... and {len(report.issues) - 20} more note(s)")
    finally:
        if temp_dir is not None:
            temp_dir.cleanup()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Analyze Cummins PDF invoices and export totals to Excel.",
    )
    parser.add_argument(
        "source",
        help="Path to one PDF file, a folder with PDFs, or a ZIP archive.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output .xlsx path. Defaults to cummins_invoice_summary.xlsx near the source.",
    )
    return parser.parse_args()


def detect_source_type(path: Path) -> str:
    if path.is_dir():
        return "folder"
    if path.suffix.lower() == ".pdf":
        return "file"
    if path.suffix.lower() == ".zip":
        return "zip"
    raise SystemExit("Source must be a PDF file, a folder, or a ZIP archive.")


def default_output_path(source_path: Path) -> Path:
    base_dir = source_path if source_path.is_dir() else source_path.parent
    return base_dir / "cummins_invoice_summary.xlsx"


if __name__ == "__main__":
    main()
