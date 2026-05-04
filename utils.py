from __future__ import annotations

import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Iterable


SUPPORTED_SUFFIXES = {".pdf"}


def collect_pdf_files(source_type: str, source_path: Path) -> tuple[list[Path], tempfile.TemporaryDirectory[str] | None]:
    if source_type == "file":
        return [source_path], None

    if source_type == "folder":
        files = sorted(
            path for path in source_path.rglob("*") if path.is_file() and path.suffix.lower() in SUPPORTED_SUFFIXES
        )
        return files, None

    if source_type == "zip":
        temp_dir = tempfile.TemporaryDirectory(prefix="cummins_pdf_")
        extract_dir = Path(temp_dir.name)
        with zipfile.ZipFile(source_path) as archive:
            archive.extractall(extract_dir)
        files = sorted(
            path for path in extract_dir.rglob("*") if path.is_file() and path.suffix.lower() in SUPPORTED_SUFFIXES
        )
        return files, temp_dir

    raise ValueError(f"Unsupported source type: {source_type}")


def ensure_xlsx_path(path: str) -> Path:
    candidate = Path(path)
    if candidate.suffix.lower() != ".xlsx":
        return candidate.with_suffix(".xlsx")
    return candidate


def format_weights(weights: Iterable[float]) -> str:
    return ", ".join(f"{weight:.3f}" for weight in weights)


def format_countries(countries: Iterable[str]) -> str:
    unique_countries: list[str] = []
    seen: set[str] = set()
    for country in countries:
        if country in seen:
            continue
        seen.add(country)
        unique_countries.append(country)
    return ", ".join(unique_countries) if unique_countries else "-"


def human_total_weight(value: float) -> str:
    return f"{value:.3f} kg"


def human_total_usd(value: float) -> str:
    return f"{value:,.2f} USD"


from tkinter import filedialog


def choose_file(title: str, allowed_extensions: list[str]) -> str:
    file_types = [("Files", [f".{ext}" for ext in allowed_extensions])]
    path = filedialog.askopenfilename(title=title, filetypes=file_types)
    return path if path else ""


def choose_folder(title: str) -> str:
    path = filedialog.askdirectory(title=title)
    return path if path else ""


def choose_save_file(title: str, default_name: str) -> str:
    path = filedialog.asksaveasfilename(title=title, initialfile=default_name, defaultextension=".xlsx")
    return path if path else ""
