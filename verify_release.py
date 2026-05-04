from __future__ import annotations

from pathlib import Path

from invoice_parser import CumminsInvoiceParser


ROOT = Path(__file__).resolve().parent
CASES = [
    ("INV", 18, 612.615),
    ("20-04-26", 56, 7345.615),
]


def main() -> None:
    parser = CumminsInvoiceParser()
    failures: list[str] = []

    for folder_name, expected_count, expected_weight in CASES:
        folder = ROOT / folder_name
        pdfs = sorted(path for path in folder.iterdir() if path.is_file() and path.suffix.lower() == ".pdf")
        report = parser.parse_files(pdfs)
        actual_weight = report.total_weight
        print(
            f"{folder_name}: pdf={len(pdfs)} invoices={len(report.invoices)} "
            f"errors={len(report.error_files)} weight={actual_weight:.3f}"
        )

        if len(pdfs) != expected_count:
            failures.append(f"{folder_name}: expected {expected_count} PDFs, got {len(pdfs)}")
        if len(report.invoices) != expected_count:
            failures.append(f"{folder_name}: expected {expected_count} invoices, got {len(report.invoices)}")
        if abs(actual_weight - expected_weight) > 0.0005:
            failures.append(f"{folder_name}: expected {expected_weight:.3f} kg, got {actual_weight:.3f} kg")
        if report.error_files:
            failures.append(f"{folder_name}: unexpected parse errors in {', '.join(report.error_files)}")

    if failures:
        print("\nFAIL")
        for failure in failures:
            print(f"- {failure}")
        raise SystemExit(1)

    print("\nPASS")


if __name__ == "__main__":
    main()
