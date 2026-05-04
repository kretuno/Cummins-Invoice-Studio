from __future__ import annotations

import re
import subprocess
import tempfile
import os
import shutil
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable

try:
    from PyPDF2 import PdfReader
except ImportError:  # pragma: no cover - handled at runtime for missing dependency
    PdfReader = None  # type: ignore[assignment]


INVOICE_RE = re.compile(r"Invoice No\s*:?\s*(?P<value>\d+)", re.IGNORECASE)
ORDER_TOTAL_RE = re.compile(
    r"Order Total:\s*(?:\n|\r\n?|\s)*(?:TO PAY\s*(?:\n|\r\n?|\s)*)?(?P<amount>[\d,]+\.\d{2})\s*USD",
    re.IGNORECASE | re.DOTALL,
)
TO_PAY_RE = re.compile(r"(?P<amount>[\d,]+\.\d{2})\s+USD\s+TO PAY", re.IGNORECASE)
FINAL_TOTAL_RE = re.compile(r"(?:Final Total|Net Invoice Amount)\s+(?P<amount>[\d,]+\.\d{2})", re.IGNORECASE)
TOTAL_PIECES_RE = re.compile(
    r"TOTAL PIECES\s*\([^)]*\)\s*(?P<count>\d+)",
    re.IGNORECASE,
)
GROSS_WEIGHT_RE = re.compile(r"GROSS WEIGHT\s*:\s*(?P<weight>[\d,]+\.\d{3})\s*KG", re.IGNORECASE)
PACKING_ROW_RE = re.compile(
    r"\bPU\d+\s+BG-\d+\s+\d+\s+ROAD\s+CUSTOMER PICK UP\s+"
    r"(?P<gross>[\d,]+\.\d{3})\s+(?P<net>[\d,]+\.\d{3})",
    re.IGNORECASE,
)
PAGE_SPLIT_RE = re.compile(r"\nINVOICE VAT\s*\nPage:\s*\n", re.IGNORECASE)
SELLER_SPLIT_RE = re.compile(r"\nSELLER:\s*\n", re.IGNORECASE)
ITEM_NO_MARKER = "Item No."
CUSTOM_STAT_MARKER = "Custom Stat No."
ORDER_TOTAL_MARKER = "Order Total:"
WEIGHT_LINE_RE = re.compile(r"^\d+(?:,\d{3})*\.\d{3}$")
CUSTOMS_CODE_RE = re.compile(r"^\d{8}$")
ITEM_CONTEXT_RE = re.compile(r"[A-Z0-9-]{8,15}(?:\s+[A-Z])?\s+\[[A-Z]{2}\]\s+\d")
COMPACT_WEIGHT_RE = re.compile(
    r"(?P<customs>\d{8})(?P<weight>\d+\.\d{3})0\s+[\d,]+\.\d{2}\s+\d{9}\s+\[[A-Z]{2}\]\s+0"
)
ORIGIN_COUNTRY_RE = re.compile(r"\[(?P<country>[A-Z]{2})\]")
ATMUS_ORIGIN_COUNTRY_RE = re.compile(r"_DS(?P<country>[A-Z]{2})\b")


@dataclass
class InvoiceRecord:
    invoice_no: str
    source_file: str
    positions: int
    line_weights: list[float]
    net_weight: float
    total_usd: float
    gross_weight: float | None = None
    package_count: int | None = None
    origin_countries: list[str] = field(default_factory=list)
    comments: list[str] = field(default_factory=list)


@dataclass
class DuplicateRecord:
    invoice_no: str
    kept_file: str
    excluded_file: str
    reason: str


@dataclass
class ParseIssue:
    source_file: str
    level: str
    message: str


@dataclass
class ParseReport:
    invoices: list[InvoiceRecord] = field(default_factory=list)
    duplicates: list[DuplicateRecord] = field(default_factory=list)
    issues: list[ParseIssue] = field(default_factory=list)

    @property
    def total_weight(self) -> float:
        return _sum_decimal((item.net_weight for item in self.invoices), places=3)

    @property
    def total_usd(self) -> float:
        return _sum_decimal((item.total_usd for item in self.invoices), places=2)

    @property
    def total_gross_weight(self) -> float:
        return _sum_decimal((item.gross_weight or 0 for item in self.invoices), places=3)

    @property
    def total_packages(self) -> int:
        return sum(item.package_count or 0 for item in self.invoices)

    @property
    def error_files(self) -> list[str]:
        seen: set[str] = set()
        files: list[str] = []
        for issue in self.issues:
            if issue.level != "ERROR" or issue.source_file in seen:
                continue
            seen.add(issue.source_file)
            files.append(issue.source_file)
        return files


class CumminsInvoiceParser:
    """Parser tailored to the Cummins invoice layout from the provided samples."""

    def parse_files(self, pdf_paths: Iterable[Path]) -> ParseReport:
        report = ParseReport()
        seen_invoices: dict[str, InvoiceRecord] = {}

        for path in pdf_paths:
            try:
                record = self.parse_file(path)
            except Exception as exc:
                report.issues.append(ParseIssue(path.name, "ERROR", f"Failed to parse PDF: {exc}"))
                continue

            if record is None:
                continue

            if record.invoice_no in seen_invoices:
                kept = seen_invoices[record.invoice_no]
                report.duplicates.append(
                    DuplicateRecord(
                        invoice_no=record.invoice_no,
                        kept_file=kept.source_file,
                        excluded_file=record.source_file,
                        reason="Duplicate Invoice No detected; latest file excluded.",
                    )
                )
                report.issues.append(
                    ParseIssue(
                        record.source_file,
                        "WARNING",
                        f"Duplicate invoice {record.invoice_no} excluded.",
                    )
                )
                continue

            seen_invoices[record.invoice_no] = record
            report.invoices.append(record)
            for comment in record.comments:
                level = "WARNING" if comment.startswith("WARNING:") else "INFO"
                message = comment.removeprefix("WARNING: ").strip() if level == "WARNING" else comment
                report.issues.append(ParseIssue(record.source_file, level, message))

        return report

    def parse_file(self, pdf_path: Path) -> InvoiceRecord | None:
        attempts: list[tuple[str, str]] = [("PyPDF2", self._extract_with_pypdf2(pdf_path))]
        if shutil.which("swift"):
            try:
                attempts.append(("PDFKit", self._extract_with_pdfkit(pdf_path)))
            except Exception as exc:
                attempts.append(("PDFKit", ""))
                errors = [f"PDFKit: {exc}"]
        else:
            errors = ["PDFKit: unavailable on this system."]
        errors = locals().get("errors", [])
        candidates: list[InvoiceRecord] = []

        for source_name, text in attempts:
            if not text or not text.strip():
                errors.append(f"{source_name}: PDF text is empty.")
                continue

            record, parse_error = self._parse_record_from_text(pdf_path, source_name, text)
            if record is not None:
                candidates.append(record)
                continue
            if parse_error:
                errors.append(parse_error)

        if candidates:
            candidates.sort(
                key=lambda record: (
                    record.positions,
                    record.net_weight,
                    len(record.origin_countries),
                    1 if any("PDFKit" in comment for comment in record.comments) else 0,
                ),
                reverse=True,
            )
            chosen = candidates[0]
            if len(candidates) > 1:
                comparison = {(record.positions, record.net_weight, tuple(record.line_weights)) for record in candidates}
                if len(comparison) > 1:
                    engine_summary = "; ".join(
                        f"{self._comment_source(record.comments)}: positions={record.positions}, weight={record.net_weight:.3f}"
                        for record in candidates
                    )
                    chosen.comments.append(
                        f"WARNING: Extraction engines disagreed. Selected {self._comment_source(chosen.comments)}. {engine_summary}"
                    )
            return chosen

        raise ValueError("; ".join(errors) if errors else "Unable to parse PDF.")

    def _parse_record_from_text(self, pdf_path: Path, source_name: str, text: str) -> tuple[InvoiceRecord | None, str | None]:
        filter_record = self._parse_filter_record_from_text(pdf_path, source_name, text)
        if filter_record is not None:
            return filter_record, None

        invoice_no = self.extract_invoice_no(text)
        total_usd = self.extract_total_usd(text)
        line_weights = self.extract_line_weights(text)
        origin_countries = self.extract_origin_countries(text)

        if not invoice_no:
            return None, f"{source_name}: Invoice No not found."
        if total_usd is None:
            return None, f"{source_name}: Total USD not found."
        if not line_weights:
            return None, f"{source_name}: Net Weight values not found in item rows."

        comments: list[str] = [f"Parsed using {source_name} text extraction."]
        if len(set(line_weights)) == 1 and len(line_weights) > 1:
            comments.append("Multiple positions share the same extracted weight value.")

        return (
            InvoiceRecord(
                invoice_no=invoice_no,
                source_file=pdf_path.name,
                positions=len(line_weights),
                line_weights=line_weights,
                net_weight=_sum_decimal(line_weights, places=3),
                total_usd=total_usd,
                origin_countries=origin_countries,
                comments=comments,
            ),
            None,
        )

    def _parse_filter_record_from_text(self, pdf_path: Path, source_name: str, text: str) -> InvoiceRecord | None:
        if "Atmus Filtration" not in text and "TOTAL PIECES" not in text and "GROSS WEIGHT" not in text:
            return None

        invoice_no = self.extract_invoice_no(text)
        total_usd = self.extract_total_usd(text)
        package_count = self.extract_package_count(text)
        gross_weight = self.extract_gross_weight(text)
        packing_weights = self.extract_packing_weights(text)
        net_weights = [net for _gross, net in packing_weights]
        gross_weights = [gross for gross, _net in packing_weights]

        if not invoice_no or total_usd is None or not net_weights:
            return None

        comments: list[str] = [f"Parsed using {source_name} text extraction.", "Detected Atmus filter invoice layout."]
        net_weight = _sum_decimal(net_weights, places=3)
        if gross_weight is None and gross_weights:
            gross_weight = _sum_decimal(gross_weights, places=3)
        elif gross_weights:
            calculated_gross = _sum_decimal(gross_weights, places=3)
            if gross_weight != calculated_gross:
                comments.append(
                    f"WARNING: Header gross weight {gross_weight:.3f} differs from packing rows {calculated_gross:.3f}."
                )

        if package_count is not None and package_count != len(net_weights):
            comments.append(
                f"WARNING: TOTAL PIECES {package_count} differs from packing rows {len(net_weights)}."
            )

        return InvoiceRecord(
            invoice_no=invoice_no,
            source_file=pdf_path.name,
            positions=package_count or len(net_weights),
            line_weights=net_weights,
            net_weight=net_weight,
            gross_weight=gross_weight,
            package_count=package_count or len(net_weights),
            total_usd=total_usd,
            origin_countries=self.extract_origin_countries(text),
            comments=comments,
        )

    def extract_text(self, pdf_path: Path) -> str:
        text = self._extract_with_pypdf2(pdf_path)
        if text.strip():
            return text

        return self._extract_with_pdfkit(pdf_path)

    def _extract_with_pypdf2(self, pdf_path: Path) -> str:
        if PdfReader is None:
            return ""

        reader = PdfReader(str(pdf_path))
        parts = [(page.extract_text() or "") for page in reader.pages]
        return "\n".join(parts)

    def _extract_with_pdfkit(self, pdf_path: Path) -> str:
        swift_script = """
import Foundation
import PDFKit

let path = CommandLine.arguments[1]
guard let document = PDFDocument(url: URL(fileURLWithPath: path)) else {
    fputs("OPEN FAILED\\n", stderr)
    exit(2)
}

var output = ""
for index in 0..<document.pageCount {
    if let page = document.page(at: index), let text = page.string {
        output += text + "\\n\\n"
    }
}

print(output)
"""

        with tempfile.NamedTemporaryFile("w", suffix=".swift", delete=False) as handle:
            handle.write(swift_script)
            script_path = Path(handle.name)

        try:
            cmd = [
                "swift",
                "-module-cache-path",
                "/tmp/clang-module-cache",
                str(script_path),
                str(pdf_path),
            ]
            env = {"CLANG_MODULE_CACHE_PATH": "/tmp/clang-module-cache"}
            completed = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                check=False,
                env={**os.environ, **env},
            )
            if completed.returncode != 0:
                raise ValueError(completed.stderr.strip() or "PDFKit fallback failed.")
            return completed.stdout
        finally:
            script_path.unlink(missing_ok=True)

    def extract_invoice_no(self, text: str) -> str | None:
        match = INVOICE_RE.search(text)
        if match:
            return match.group("value")

        reverse_match = re.search(r"(?P<value>\d{6,10})\s+Invoice No:", text, re.IGNORECASE)
        if reverse_match:
            return reverse_match.group("value")

        lines = [line.strip() for line in text.splitlines()]
        for index, line in enumerate(lines):
            if line != "Invoice No:":
                continue

            for candidate in lines[index + 1 : index + 20]:
                if re.fullmatch(r"\d{6,10}", candidate):
                    return candidate

        return None

    def extract_total_usd(self, text: str) -> float | None:
        match = ORDER_TOTAL_RE.search(text) or TO_PAY_RE.search(text) or FINAL_TOTAL_RE.search(text)
        if not match:
            return None
        return self._to_float(match.group("amount"))

    def extract_package_count(self, text: str) -> int | None:
        match = TOTAL_PIECES_RE.search(text)
        if not match:
            return None
        return int(match.group("count"))

    def extract_gross_weight(self, text: str) -> float | None:
        match = GROSS_WEIGHT_RE.search(text)
        if not match:
            return None
        return self._to_float(match.group("weight"))

    def extract_packing_weights(self, text: str) -> list[tuple[float, float]]:
        weights: list[tuple[float, float]] = []
        for match in PACKING_ROW_RE.finditer(text):
            gross = self._to_float(match.group("gross"))
            net = self._to_float(match.group("net"))
            if gross is None or net is None:
                continue
            weights.append((gross, net))
        return weights

    def extract_line_weights(self, text: str) -> list[float]:
        weights: list[float] = []
        for section in self._iter_item_sections(text):
            weights.extend(self._extract_weights_from_section(section))
        if not weights:
            weights = self._extract_compact_line_weights(text)
        return weights

    def extract_origin_countries(self, text: str) -> list[str]:
        countries: list[str] = []
        seen: set[str] = set()
        for match in ORIGIN_COUNTRY_RE.finditer(text):
            country = f"[{match.group('country')}]"
            if country in seen:
                continue
            seen.add(country)
            countries.append(country)
        for match in ATMUS_ORIGIN_COUNTRY_RE.finditer(text):
            country = f"[{match.group('country')}]"
            if country in seen:
                continue
            seen.add(country)
            countries.append(country)
        return countries

    def _iter_item_sections(self, text: str) -> Iterable[str]:
        pages = self._split_pages(text)
        for page in pages:
            if ITEM_NO_MARKER not in page:
                continue

            start = page.find(ITEM_NO_MARKER)
            end_candidates = [
                idx for idx in (page.find(ORDER_TOTAL_MARKER, start), page.find("If the invoice is not disputed", start)) if idx != -1
            ]
            end = min(end_candidates) if end_candidates else len(page)
            section = page[start:end]
            if SELLER_SPLIT_RE.search(section):
                section = SELLER_SPLIT_RE.split(section, maxsplit=1)[0]
            yield section

    def _split_pages(self, text: str) -> list[str]:
        parts = PAGE_SPLIT_RE.split(text)
        if len(parts) <= 1:
            return [text]
        return [part for part in parts if part.strip()]

    def _extract_weights_from_section(self, section: str) -> list[float]:
        if CUSTOM_STAT_MARKER not in section:
            return []

        body = section.split(CUSTOM_STAT_MARKER, maxsplit=1)[1]
        lines = [line.strip() for line in body.splitlines() if line.strip()]

        weights: list[float] = []
        for index, line in enumerate(lines):
            if not ITEM_CONTEXT_RE.search(line):
                continue

            customs_index = self._find_customs_code_index(lines, index)
            if customs_index is None:
                continue

            candidate_lines = lines[index + 1 : customs_index]
            candidate_weights = [candidate for candidate in candidate_lines if WEIGHT_LINE_RE.fullmatch(candidate)]
            if not candidate_weights:
                continue

            value = self._to_float(candidate_weights[-1])
            if value is not None:
                weights.append(value)

        return weights

    def _find_customs_code_index(self, lines: list[str], item_index: int) -> int | None:
        search_end = min(len(lines), item_index + 10)
        for index in range(item_index + 1, search_end):
            if CUSTOMS_CODE_RE.fullmatch(lines[index]):
                return index
        return None

    @staticmethod
    def _comment_source(comments: list[str]) -> str:
        for comment in comments:
            if "Parsed using " in comment:
                return comment.removeprefix("Parsed using ").removesuffix(" text extraction.")
        return "unknown engine"

    def _extract_compact_line_weights(self, text: str) -> list[float]:
        weights: list[float] = []
        for match in COMPACT_WEIGHT_RE.finditer(text):
            value = self._to_float(match.group("weight"))
            if value is not None:
                weights.append(value)
        return weights

    @staticmethod
    def _to_float(value: str) -> float | None:
        try:
            normalized = value.replace(",", "")
            return float(Decimal(normalized))
        except (InvalidOperation, ValueError):
            return None


def _sum_decimal(values: Iterable[float], places: int) -> float:
    total = sum((Decimal(str(value)) for value in values), Decimal("0"))
    quantum = Decimal("1").scaleb(-places)
    return float(total.quantize(quantum))
