from __future__ import annotations

import queue
import threading
import tkinter as tk
import os
import sys
from dataclasses import dataclass, field
from pathlib import Path
from tempfile import TemporaryDirectory
from tkinter import messagebox

from exporter import export_report_to_excel
from parser import CumminsInvoiceParser, ParseReport
from utils import (
    choose_file,
    choose_folder,
    choose_save_file,
    collect_pdf_files,
    ensure_xlsx_path,
    format_countries,
    format_weights,
    human_total_usd,
    human_total_weight,
)


APP_TITLE = "Cummins PDF Invoice Processor"
UI_BUILD = "CANVAS BUILD 2026-04-28-01"

BG = "#eef2f6"
SURFACE = "#ffffff"
SURFACE_ALT = "#f8fafc"
INK = "#172033"
MUTED = "#667085"
BORDER = "#d7dde6"
ACCENT = "#d99621"
BLUE = "#235c9f"
GREEN = "#1f9d66"
RED = "#b42318"


def _format_optional_weight(value: float | None) -> str:
    return f"{value:.3f} kg" if value is not None else "-"


@dataclass
class AppState:
    source_type: str | None = None
    source_path: Path | None = None
    report: ParseReport = field(default_factory=ParseReport)
    temp_dir: TemporaryDirectory[str] | None = None
    pdf_count: int = 0
    is_analyzing: bool = False
    status: str = "Готово к работе"
    result_text: str = "Выберите PDF, папку или ZIP. После анализа результат останется здесь."
    log_lines: list[str] = field(default_factory=list)


@dataclass
class AnalysisResult:
    request_id: int
    pdf_count: int
    report: ParseReport
    temp_dir: TemporaryDirectory[str] | None = None
    error: str | None = None


class CumminsInvoiceApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"{UI_BUILD} | {APP_TITLE}")
        self.root.geometry("1180x780")
        self.root.minsize(980, 640)
        self.root.configure(bg=BG)

        self.state = AppState()
        self.parser = CumminsInvoiceParser()
        self.analysis_queue: queue.Queue[AnalysisResult] = queue.Queue()
        self.analysis_request_id = 0

        self._build_ui()
        self._log("INFO", "Приложение запущено.")
        self._draw()

    def _build_ui(self) -> None:
        self.toolbar = tk.Frame(self.root, bg=BG)
        self.toolbar.pack(fill="x", padx=18, pady=(16, 8))

        buttons = [
            ("Выбрать PDF", self.on_choose_pdf, "#ffffff", INK),
            ("Выбрать папку", self.on_choose_folder, "#ffffff", INK),
            ("Выбрать ZIP", self.on_choose_zip, "#ffffff", INK),
            ("Анализировать", self.on_analyze, "#ffe6ad", "#5d3a00"),
            ("Сохранить Excel", self.on_save_excel, "#dbeafe", "#173f75"),
            ("Очистить", self.on_clear, "#fee2e2", "#8f1d1d"),
        ]
        for title, command, bg, fg in buttons:
            tk.Button(
                self.toolbar,
                text=title,
                command=command,
                bg=bg,
                fg=fg,
                activebackground=bg,
                activeforeground=fg,
                relief="solid",
                bd=1,
                padx=13,
                pady=7,
                font=("Helvetica", 11, "bold"),
            ).pack(side="left", padx=(0, 8))

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=18, pady=(0, 16))
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(body, bg=BG, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        self.canvas.bind("<Configure>", lambda _event: self._draw())
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def on_choose_pdf(self) -> None:
        self._set_status("Открываю выбор PDF...")
        path = choose_file("Select a Cummins PDF invoice", ["pdf"])
        if not path:
            self._set_status("Выбор PDF отменен")
            self._log("INFO", "Выбор PDF отменен.")
            self._draw()
            return
        self._set_source("file", Path(path))
        self.on_analyze()

    def on_choose_folder(self) -> None:
        self._set_status("Открываю выбор папки...")
        path = choose_folder("Select a folder with Cummins PDF invoices")
        if not path:
            self._set_status("Выбор папки отменен")
            self._log("INFO", "Выбор папки отменен.")
            self._draw()
            return
        self._set_source("folder", Path(path))
        self.on_analyze()

    def on_choose_zip(self) -> None:
        self._set_status("Открываю выбор ZIP...")
        path = choose_file("Select a ZIP archive with PDFs", ["zip"])
        if not path:
            self._set_status("Выбор ZIP отменен")
            self._log("INFO", "Выбор ZIP отменен.")
            self._draw()
            return
        self._set_source("zip", Path(path))
        self.on_analyze()

    def on_analyze(self) -> None:
        if self.state.is_analyzing:
            return
        if self.state.source_path is None or self.state.source_type is None:
            messagebox.showwarning("Источник не выбран", "Сначала выберите PDF, папку или ZIP.", parent=self.root)
            return

        self._cleanup_temp_dir()
        self.analysis_request_id += 1
        request_id = self.analysis_request_id
        self.state.is_analyzing = True
        self.state.report = ParseReport()
        self.state.pdf_count = 0
        self.state.result_text = "Идет анализ выбранного источника. Подождите несколько секунд..."
        self._set_status("Анализ...")
        self._log("INFO", f"Старт анализа: {self.state.source_path}")
        self._draw()

        worker = threading.Thread(
            target=self._analysis_worker,
            args=(request_id, self.state.source_type, self.state.source_path),
            daemon=True,
        )
        worker.start()
        self.root.after(120, lambda: self._poll_analysis(request_id))

    def on_save_excel(self) -> None:
        if not (self.state.report.invoices or self.state.report.issues or self.state.report.duplicates):
            messagebox.showwarning("Нет данных", "Сначала выполните анализ.", parent=self.root)
            return

        path = choose_save_file("Save Excel report as", "cummins_invoice_summary.xlsx")
        if not path:
            self._set_status("Сохранение Excel отменено")
            self._draw()
            return

        destination = ensure_xlsx_path(path)
        try:
            export_report_to_excel(self.state.report, destination)
        except Exception as exc:
            self._set_status("Ошибка сохранения Excel")
            self._log("ERROR", f"Ошибка сохранения Excel: {exc}")
            self._draw()
            messagebox.showerror("Ошибка сохранения", str(exc), parent=self.root)
            return

        self._set_status(f"Excel сохранен: {destination.name}")
        self._log("INFO", f"Excel сохранен: {destination}")
        self._draw()

    def on_clear(self) -> None:
        self.analysis_request_id += 1
        self._cleanup_temp_dir()
        self.state = AppState()
        self._log("INFO", "Очищено.")
        self._draw()

    def _set_source(self, source_type: str, source_path: Path) -> None:
        self.state.source_type = source_type
        self.state.source_path = source_path
        self.state.report = ParseReport()
        self.state.pdf_count = 0
        self.state.result_text = "Источник выбран. Анализ запускается..."
        self._set_status(f"Источник выбран: {source_path.name}")
        self._log("INFO", f"Источник выбран: {source_path}")
        self._draw()

    def _analysis_worker(self, request_id: int, source_type: str, source_path: Path) -> None:
        try:
            pdf_files, temp_dir = collect_pdf_files(source_type, source_path)
            report = self.parser.parse_files(pdf_files) if pdf_files else ParseReport()
            result = AnalysisResult(request_id, len(pdf_files), report, temp_dir=temp_dir)
        except Exception as exc:
            result = AnalysisResult(request_id, 0, ParseReport(), error=str(exc))
        self.analysis_queue.put(result)

    def _poll_analysis(self, request_id: int) -> None:
        if request_id != self.analysis_request_id:
            return

        try:
            result = self.analysis_queue.get_nowait()
        except queue.Empty:
            if self.state.is_analyzing:
                self.root.after(120, lambda: self._poll_analysis(request_id))
            return

        self.state.is_analyzing = False
        if result.error:
            self.state.result_text = f"Ошибка анализа:\n{result.error}"
            self._set_status("Ошибка анализа")
            self._log("ERROR", result.error)
            self._draw()
            return

        self.state.pdf_count = result.pdf_count
        self.state.report = result.report
        self.state.temp_dir = result.temp_dir

        if result.pdf_count == 0:
            self.state.result_text = "В выбранном источнике не найдено PDF-файлов."
            self._set_status("PDF-файлы не найдены")
            self._log("WARNING", "PDF-файлы не найдены.")
        elif result.report.invoices:
            self._set_status(f"Анализ завершен: найдено {len(result.report.invoices)} инвойсов")
            self._log(
                "INFO",
                f"Готово. PDF: {result.pdf_count}; инвойсы: {len(result.report.invoices)}; "
                f"дубликаты: {len(result.report.duplicates)}; заметки: {len(result.report.issues)}.",
            )
        else:
            self.state.result_text = "PDF обработаны, но инвойсы Cummins не найдены. Смотрите лог ниже."
            self._set_status("Инвойсы не найдены")
            self._log("WARNING", "Инвойсы не найдены.")

        self._draw()

    def _draw(self) -> None:
        self.canvas.delete("all")
        width = max(self.canvas.winfo_width(), 900)
        x = 18
        y = 12
        usable = width - 36

        y = self._draw_header(x, y, usable)
        y = self._draw_source(x, y + 10, usable)
        y = self._draw_metrics(x, y + 10, usable)
        y = self._draw_results(x, y + 10, usable)
        y = self._draw_table(x, y + 10, usable)
        y = self._draw_log(x, y + 10, usable)
        y = self._draw_status(x, y + 10, usable)

        self.canvas.configure(scrollregion=(0, 0, width, y + 24))
        self.root.title(f"{UI_BUILD} | {APP_TITLE} | {self.state.status}")

    def _draw_header(self, x: int, y: int, w: int) -> int:
        h = 92
        self._card(x, y, w, h)
        self._text(x + 18, y + 18, APP_TITLE, 22, "bold", INK)
        self._text(x + 18, y + 50, "Анализ PDF-инвойсов Cummins, контроль дублей и экспорт в Excel.", 12, "normal", MUTED)
        self._pill(x + w - 260, y + 22, 230, 30, UI_BUILD, "#fff1c7", "#5d3a00")
        return y + h

    def _draw_source(self, x: int, y: int, w: int) -> int:
        h = 88
        self._card(x, y, w, h)
        self._section_title(x + 16, y + 14, "Источник")
        self._text(x + 16, y + 44, self._source_text(), 11, "normal", INK, width=w - 32)
        return y + h

    def _draw_metrics(self, x: int, y: int, w: int) -> int:
        gap = 10
        card_w = (w - gap * 3) / 4
        values = [
            ("Инвойсов", str(len(self.state.report.invoices))),
            ("Вес нетто", human_total_weight(self.state.report.total_weight)),
            ("Вес брутто", human_total_weight(self.state.report.total_gross_weight)),
            ("Сумма USD", human_total_usd(self.state.report.total_usd)),
        ]
        for idx, (title, value) in enumerate(values):
            cx = x + idx * (card_w + gap)
            self._card(cx, y, card_w, 86)
            self._text(cx + 14, y + 14, title, 10, "bold", MUTED)
            self._text(cx + 14, y + 38, value, 20, "bold", INK)
        return y + 86

    def _draw_results(self, x: int, y: int, w: int) -> int:
        text = self._result_text()
        line_count = max(5, min(12, len(text.splitlines())))
        h = 78 + line_count * 22
        self._card(x, y, w, h)
        self._section_title(x + 16, y + 14, "Результат анализа")
        self.canvas.create_rectangle(x + 16, y + 44, x + w - 16, y + h - 16, fill=SURFACE_ALT, outline=BORDER)
        self._text(x + 30, y + 58, text, 12, "normal", INK, width=w - 60)
        return y + h

    def _draw_table(self, x: int, y: int, w: int) -> int:
        invoices = self.state.report.invoices
        row_h = 30
        visible_rows = max(5, min(14, len(invoices) if invoices else 5))
        h = 76 + row_h * visible_rows
        self._card(x, y, w, h)
        self._section_title(x + 16, y + 14, "Таблица инвойсов")

        table_x = x + 16
        table_y = y + 44
        table_w = w - 32
        cols = [
            ("Invoice No", 0.10),
            ("Source File", 0.18),
            ("Positions", 0.07),
            ("Countries", 0.10),
            ("Places", 0.06),
            ("Line Weights (kg)", 0.22),
            ("Net Weight", 0.10),
            ("Gross Weight", 0.10),
            ("Total USD", 0.07),
        ]
        self.canvas.create_rectangle(table_x, table_y, table_x + table_w, table_y + row_h, fill="#dbe4ee", outline=BORDER)
        cx = table_x
        for title, ratio in cols:
            cw = table_w * ratio
            self._text(cx + 6, table_y + 8, title, 10, "bold", INK, width=cw - 10)
            cx += cw

        if not invoices:
            self.canvas.create_rectangle(table_x, table_y + row_h, table_x + table_w, table_y + row_h * 2, fill="#ffffff", outline=BORDER)
            self._text(table_x + 8, table_y + row_h + 8, "После анализа строки появятся здесь.", 11, "normal", MUTED)
            return y + h

        for index, record in enumerate(invoices[:visible_rows]):
            ry = table_y + row_h * (index + 1)
            fill = "#ffffff" if index % 2 == 0 else "#f8fafc"
            self.canvas.create_rectangle(table_x, ry, table_x + table_w, ry + row_h, fill=fill, outline=BORDER)
            values = [
                record.invoice_no,
                record.source_file,
                str(record.positions),
                format_countries(record.origin_countries),
                str(record.package_count) if record.package_count is not None else "",
                format_weights(record.line_weights),
                f"{record.net_weight:.3f}",
                f"{record.gross_weight:.3f}" if record.gross_weight is not None else "",
                f"{record.total_usd:,.2f}",
            ]
            cx = table_x
            for value, (_title, ratio) in zip(values, cols):
                cw = table_w * ratio
                clipped = value if len(value) <= 42 else value[:39] + "..."
                self._text(cx + 6, ry + 8, clipped, 9, "normal", INK, width=cw - 10)
                cx += cw

        if len(invoices) > visible_rows:
            self._text(table_x, y + h - 24, f"Показано {visible_rows} из {len(invoices)}. Все строки будут в Excel.", 10, "normal", MUTED)
        return y + h

    def _draw_log(self, x: int, y: int, w: int) -> int:
        lines = self._log_lines()
        h = 70 + min(8, max(4, len(lines))) * 20
        self._card(x, y, w, h)
        self._section_title(x + 16, y + 14, "Лог")
        log_x = x + 16
        log_y = y + 44
        self.canvas.create_rectangle(log_x, log_y, x + w - 16, y + h - 14, fill="#111827", outline="#111827")
        self._text(log_x + 12, log_y + 10, "\n".join(lines[-8:]), 10, "normal", "#e5e7eb", width=w - 56, font_family="Courier")
        return y + h

    def _draw_status(self, x: int, y: int, w: int) -> int:
        h = 46
        self._card(x, y, w, h)
        self._text(x + 16, y + 14, self.state.status, 11, "bold", BLUE, width=w - 32)
        return y + h

    def _result_text(self) -> str:
        report = self.state.report
        if not report.invoices:
            return self.state.result_text
        lines = [
            f"PDF-файлов найдено: {self.state.pdf_count}",
            f"Распознано инвойсов: {len(report.invoices)}",
            f"Не распознано PDF: {len(report.error_files)}",
            f"Исключено дубликатов: {len(report.duplicates)}",
            f"Общий вес нетто: {human_total_weight(report.total_weight)}",
            f"Общий вес брутто: {human_total_weight(report.total_gross_weight)}",
            f"Количество мест: {report.total_packages}",
            f"Общая сумма USD: {human_total_usd(report.total_usd)}",
            "",
            "Инвойсы:",
        ]
        for record in report.invoices:
            lines.append(
                f"- {record.invoice_no}: {record.net_weight:.3f} kg, "
                f"gross: {_format_optional_weight(record.gross_weight)}, "
                f"{record.total_usd:,.2f} USD, мест: {record.package_count or '-'}, "
                f"позиций: {record.positions}, файл: {record.source_file}"
            )
        if report.duplicates:
            lines.append("")
            lines.append("Дубликаты:")
            for duplicate in report.duplicates:
                lines.append(
                    f"- {duplicate.invoice_no}: оставлен {duplicate.kept_file}, исключен {duplicate.excluded_file}"
                )
        if report.error_files:
            lines.append("")
            lines.append("Не распознаны PDF:")
            for source_file in report.error_files:
                lines.append(f"- {source_file}")
        if report.issues:
            lines.append(f"Заметки/предупреждения: {len(report.issues)}. Подробности в логе.")
        return "\n".join(lines)

    def _log_lines(self) -> list[str]:
        lines = list(self.state.log_lines)
        for duplicate in self.state.report.duplicates:
            lines.append(
                f"[WARNING] Дубликат {duplicate.invoice_no}: оставлен {duplicate.kept_file}, исключен {duplicate.excluded_file}"
            )
        for issue in self.state.report.issues:
            lines.append(f"[{issue.level}] {issue.source_file}: {issue.message}")
        return lines or ["[INFO] Лог пока пуст."]

    def _source_text(self) -> str:
        if self.state.source_path is None:
            return "Источник не выбран. Нажмите 'Выбрать PDF', 'Выбрать папку' или 'Выбрать ZIP'."
        label = {"file": "PDF", "folder": "Папка", "zip": "ZIP"}.get(self.state.source_type, "Источник")
        return f"{label}: {self.state.source_path}"

    def _card(self, x: float, y: float, w: float, h: float) -> None:
        self.canvas.create_rectangle(x, y, x + w, y + h, fill=SURFACE, outline=BORDER)

    def _section_title(self, x: float, y: float, text: str) -> None:
        self._text(x, y, text, 12, "bold", INK)

    def _pill(self, x: float, y: float, w: float, h: float, text: str, fill: str, fg: str) -> None:
        self.canvas.create_rectangle(x, y, x + w, y + h, fill=fill, outline=fill)
        self._text(x + 10, y + 8, text, 10, "bold", fg, width=w - 20)

    def _text(
        self,
        x: float,
        y: float,
        text: str,
        size: int,
        weight: str,
        fill: str,
        width: float | None = None,
        font_family: str = "Helvetica",
    ) -> None:
        self.canvas.create_text(
            x,
            y,
            text=text,
            anchor="nw",
            fill=fill,
            font=(font_family, size, weight),
            width=width,
        )

    def _on_mousewheel(self, event: tk.Event) -> None:
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _set_status(self, status: str) -> None:
        self.state.status = status
        self.root.title(f"{UI_BUILD} | {APP_TITLE} | {status}")
        self.root.update_idletasks()

    def _log(self, level: str, message: str) -> None:
        self.state.log_lines.append(f"[{level}] {message}")

    def _cleanup_temp_dir(self) -> None:
        if self.state.temp_dir is not None:
            self.state.temp_dir.cleanup()
            self.state.temp_dir = None


def legacy_main() -> None:
    root = tk.Tk()
    app = CumminsInvoiceApp(root)
    root.protocol("WM_DELETE_WINDOW", lambda: _on_close(app))
    root.mainloop()


def _on_close(app: CumminsInvoiceApp) -> None:
    app._cleanup_temp_dir()
    app.root.destroy()


if __name__ == "__main__":
    venv_python = Path(__file__).resolve().parent / ".venv" / "bin" / "python"
    if venv_python.exists() and Path(sys.executable).resolve() != venv_python.resolve():
        os.execv(str(venv_python), [str(venv_python), str(Path(__file__).resolve())])

    try:
        from qt_app import main as qt_main

        qt_main()
    except Exception:
        legacy_main()
