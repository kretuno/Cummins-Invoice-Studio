from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

try:
    from PySide6.QtCore import QThread, Signal, Qt
    from PySide6.QtGui import QIcon
    from PySide6.QtWidgets import (
        QApplication,
        QFileDialog,
        QFrame,
        QGridLayout,
        QHBoxLayout,
        QLabel,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QTableWidget,
        QTableWidgetItem,
        QTextEdit,
        QVBoxLayout,
        QWidget,
    )
except ImportError:  # pragma: no cover
    print("PySide6 is not installed. Run: pip install PySide6")
    raise

import parser as parser_module

from exporter import export_report_to_excel
from parser import CumminsInvoiceParser, ParseReport
from utils import collect_pdf_files, ensure_xlsx_path, format_countries, format_weights, human_total_usd, human_total_weight


APP_TITLE = "Cummins Invoice Studio"
APP_VERSION = "1.0.8"
APP_SUBTITLE = "Парсер PDF-инвойсов Cummins: общий вес, сумма USD, страны происхождения и Excel-экспорт."
DEVELOPER_NAME = "Eduard Osipov"
DEVELOPER_EMAIL = "edosipov@gmail.com"
DEVELOPER_PHONE = "+380675694704"
ROOT = Path(__file__).resolve().parent
ICON_PATH = ROOT / "assets" / "app_icon_512.png"


def _format_optional_weight(value: float | None) -> str:
    return f"{value:.3f} kg" if value is not None else "-"


def _detect_parser_build() -> str:
    parser_path = getattr(parser_module, "__file__", None)
    if not parser_path:
        return "embedded"

    try:
        return datetime.fromtimestamp(Path(parser_path).stat().st_mtime).strftime("%Y-%m-%d %H:%M")
    except OSError:
        return "embedded"


PARSER_BUILD = _detect_parser_build()


class AnalysisWorker(QThread):
    finished = Signal(object, int)
    failed = Signal(str)

    def __init__(self, source_type: str, source_path: Path) -> None:
        super().__init__()
        self.source_type = source_type
        self.source_path = source_path

    def run(self) -> None:
        temp_dir = None
        try:
            pdf_files, temp_dir = collect_pdf_files(self.source_type, self.source_path)
            report = CumminsInvoiceParser().parse_files(pdf_files) if pdf_files else ParseReport()
            self.finished.emit(report, len(pdf_files))
        except Exception as exc:
            self.failed.emit(str(exc))
        finally:
            if temp_dir is not None:
                temp_dir.cleanup()


class MetricCard(QFrame):
    def __init__(self, title: str) -> None:
        super().__init__()
        self.setObjectName("MetricCard")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(4)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("MetricTitle")
        self.value_label = QLabel("0")
        self.value_label.setObjectName("MetricValue")

        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)

    def set_value(self, value: str) -> None:
        self.value_label.setText(value)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_TITLE} {APP_VERSION} | parser {PARSER_BUILD}")
        self.resize(1160, 720)
        self.setMinimumSize(860, 560)
        if ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(ICON_PATH)))

        self.source_type: str | None = None
        self.source_path: Path | None = None
        self.report = ParseReport()
        self.worker: AnalysisWorker | None = None
        self.pdf_count = 0

        self._build_ui()
        self._apply_style()
        self._render_empty()

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        header = QFrame()
        header.setObjectName("Panel")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(18, 14, 18, 14)
        title = QLabel(APP_TITLE)
        title.setObjectName("Title")
        subtitle = QLabel(APP_SUBTITLE)
        subtitle.setObjectName("Subtitle")
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        layout.addWidget(header)

        toolbar = QFrame()
        toolbar.setObjectName("Panel")
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 12, 12, 12)
        toolbar_layout.setSpacing(8)

        self.pdf_button = QPushButton("Выбрать PDF")
        self.folder_button = QPushButton("Выбрать папку")
        self.zip_button = QPushButton("Выбрать ZIP")
        self.analyze_button = QPushButton("Анализировать")
        self.save_button = QPushButton("Сохранить Excel")
        self.clear_button = QPushButton("Очистить")
        self.about_button = QPushButton("О программе")

        self.analyze_button.setObjectName("AccentButton")
        self.save_button.setObjectName("BlueButton")
        self.clear_button.setObjectName("DangerButton")
        self.about_button.setObjectName("GhostButton")

        for button in (self.pdf_button, self.folder_button, self.zip_button):
            toolbar_layout.addWidget(button)

        toolbar_layout.addStretch(1)
        for button in (self.analyze_button, self.save_button, self.clear_button, self.about_button):
            toolbar_layout.addWidget(button)

        layout.addWidget(toolbar)

        self.pdf_button.clicked.connect(self.choose_pdf)
        self.folder_button.clicked.connect(self.choose_folder)
        self.zip_button.clicked.connect(self.choose_zip)
        self.analyze_button.clicked.connect(self.analyze)
        self.save_button.clicked.connect(self.save_excel)
        self.clear_button.clicked.connect(self.clear)
        self.about_button.clicked.connect(self.show_about)

        source_panel = QFrame()
        source_panel.setObjectName("Panel")
        source_layout = QVBoxLayout(source_panel)
        source_layout.setContentsMargins(16, 12, 16, 12)
        source_title = QLabel("Источник")
        source_title.setObjectName("SectionTitle")
        self.source_label = QLabel("Источник не выбран")
        self.source_label.setWordWrap(True)
        source_layout.addWidget(source_title)
        source_layout.addWidget(self.source_label)
        layout.addWidget(source_panel)

        countries_panel = QFrame()
        countries_panel.setObjectName("Panel")
        countries_layout = QVBoxLayout(countries_panel)
        countries_layout.setContentsMargins(16, 12, 16, 12)
        countries_title = QLabel("Страны происхождения")
        countries_title.setObjectName("SectionTitle")
        self.countries_label = QLabel("Страны появятся после анализа")
        self.countries_label.setWordWrap(True)
        countries_layout.addWidget(countries_title)
        countries_layout.addWidget(self.countries_label)
        layout.addWidget(countries_panel)

        metrics = QFrame()
        metrics_layout = QGridLayout(metrics)
        metrics_layout.setContentsMargins(0, 0, 0, 0)
        metrics_layout.setHorizontalSpacing(12)
        self.invoice_card = MetricCard("Инвойсов")
        self.weight_card = MetricCard("Вес нетто")
        self.gross_weight_card = MetricCard("Вес брутто")
        self.usd_card = MetricCard("Сумма USD")
        metrics_layout.addWidget(self.invoice_card, 0, 0)
        metrics_layout.addWidget(self.weight_card, 0, 1)
        metrics_layout.addWidget(self.gross_weight_card, 0, 2)
        metrics_layout.addWidget(self.usd_card, 0, 3)
        layout.addWidget(metrics)

        content = QFrame()
        content.setObjectName("Panel")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(14, 10, 14, 14)
        content_layout.setSpacing(8)

        result_title = QLabel("Результат анализа")
        result_title.setObjectName("SectionTitle")
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setMinimumHeight(80)
        self.result_text.setMaximumHeight(135)
        self.result_text.setObjectName("ResultText")

        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels(
            [
                "Invoice No",
                "Source File",
                "Positions",
                "Origin Countries",
                "Places",
                "Line Weights (kg)",
                "Net Weight (kg)",
                "Gross Weight (kg)",
                "Total USD",
            ]
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setMinimumHeight(180)

        log_title = QLabel("Лог")
        log_title.setObjectName("SectionTitle")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(70)
        self.log_text.setMaximumHeight(120)
        self.log_text.setObjectName("LogText")

        content_layout.addWidget(result_title)
        content_layout.addWidget(self.result_text)
        content_layout.addWidget(self.table, stretch=1)
        content_layout.addWidget(log_title)
        content_layout.addWidget(self.log_text)
        layout.addWidget(content, stretch=1)

        self.status_label = QLabel("Готово к работе")
        self.status_label.setObjectName("Status")
        layout.addWidget(self.status_label)

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background: #eef2f6; }
            QFrame#Panel, QFrame#MetricCard {
                background: #ffffff;
                border: 1px solid #d7dde6;
                border-radius: 12px;
            }
            QLabel#Title {
                color: #172033;
                font-size: 24px;
                font-weight: 800;
            }
            QLabel#Subtitle, QLabel {
                color: #667085;
                font-size: 13px;
            }
            QLabel#SectionTitle {
                color: #172033;
                font-size: 15px;
                font-weight: 700;
            }
            QLabel#MetricTitle {
                color: #667085;
                font-size: 12px;
                font-weight: 700;
            }
            QLabel#MetricValue {
                color: #172033;
                font-size: 24px;
                font-weight: 800;
            }
            QPushButton {
                background: #f8fafc;
                border: 1px solid #cfd7e3;
                border-radius: 10px;
                padding: 8px 14px;
                color: #172033;
                font-weight: 700;
            }
            QPushButton:hover { background: #eef4fb; }
            QPushButton#AccentButton { background: #ffe7b4; color: #5d3a00; }
            QPushButton#BlueButton { background: #dbeafe; color: #173f75; }
            QPushButton#DangerButton { background: #fee2e2; color: #8f1d1d; }
            QPushButton#GhostButton { background: #ffffff; color: #667085; }
            QTextEdit#ResultText {
                background: #f8fafc;
                border: 1px solid #d7dde6;
                border-radius: 10px;
                color: #172033;
                font-size: 14px;
            }
            QTextEdit#LogText {
                background: #111827;
                border-radius: 10px;
                color: #e5e7eb;
                font-family: Menlo;
                font-size: 12px;
            }
            QTableWidget {
                background: #ffffff;
                alternate-background-color: #f8fafc;
                border: 1px solid #d7dde6;
                border-radius: 10px;
                gridline-color: #e5e7eb;
                color: #172033;
            }
            QHeaderView::section {
                background: #dbe4ee;
                color: #172033;
                font-weight: 700;
                padding: 6px;
                border: 0;
            }
            QLabel#Status {
                color: #235c9f;
                font-weight: 700;
                padding: 4px;
            }
            """
        )

    def choose_pdf(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Выберите PDF", "", "PDF Files (*.pdf)")
        if path:
            self.set_source("file", Path(path))
            self.analyze()

    def choose_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Выберите папку с PDF")
        if path:
            self.set_source("folder", Path(path))
            self.analyze()

    def choose_zip(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Выберите ZIP", "", "ZIP Files (*.zip)")
        if path:
            self.set_source("zip", Path(path))
            self.analyze()

    def set_source(self, source_type: str, source_path: Path) -> None:
        self.source_type = source_type
        self.source_path = source_path
        self.report = ParseReport()
        self.pdf_count = 0
        self.source_label.setText(f"{source_type.upper()}: {source_path}")
        self.countries_label.setText("Страны появятся после анализа")
        self.set_status(f"Источник выбран: {source_path.name}")
        self.result_text.setPlainText("Источник выбран. Запускаю анализ...")
        self.log(f"[INFO] Источник выбран: {source_path}")
        self.render_report()

    def analyze(self) -> None:
        if self.source_type is None or self.source_path is None:
            QMessageBox.warning(self, "Источник не выбран", "Сначала выберите PDF, папку или ZIP.")
            return
        if self.worker is not None and self.worker.isRunning():
            return

        self.set_status("Анализ...")
        self.result_text.setPlainText("Идет анализ. Подождите несколько секунд...")
        self.table.setRowCount(0)
        self.worker = AnalysisWorker(self.source_type, self.source_path)
        self.worker.finished.connect(self.on_analysis_finished)
        self.worker.failed.connect(self.on_analysis_failed)
        self.worker.start()

    def on_analysis_finished(self, report: ParseReport, pdf_count: int) -> None:
        self.report = report
        self.pdf_count = pdf_count
        if pdf_count == 0:
            self.set_status("PDF-файлы не найдены")
            self.result_text.setPlainText("В выбранном источнике не найдено PDF-файлов.")
        elif report.invoices:
            self.set_status(
                f"Анализ завершен: PDF {pdf_count}; инвойсы {len(report.invoices)}; "
                f"ошибки {len(report.error_files)}"
            )
            self.log(
                f"[INFO] Готово. PDF: {pdf_count}; инвойсы: {len(report.invoices)}; "
                f"дубликаты: {len(report.duplicates)}; заметки: {len(report.issues)}."
            )
        else:
            self.set_status("Инвойсы не найдены")
            self.result_text.setPlainText("PDF обработаны, но инвойсы Cummins не найдены. Смотрите лог.")
        self.render_report()

    def on_analysis_failed(self, error: str) -> None:
        self.set_status("Ошибка анализа")
        self.result_text.setPlainText(f"Ошибка анализа:\n{error}")
        self.log(f"[ERROR] {error}")

    def render_report(self) -> None:
        self.invoice_card.set_value(str(len(self.report.invoices)))
        self.weight_card.set_value(human_total_weight(self.report.total_weight))
        self.gross_weight_card.set_value(human_total_weight(self.report.total_gross_weight))
        self.usd_card.set_value(human_total_usd(self.report.total_usd))
        self.countries_label.setText(self.format_report_countries())

        self.table.setRowCount(len(self.report.invoices))
        for row, record in enumerate(self.report.invoices):
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
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column in (2, 4, 6, 7, 8):
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row, column, item)
        self.table.resizeColumnsToContents()

        if not self.report.invoices:
            if not self.result_text.toPlainText().strip():
                self.result_text.setPlainText("Выберите PDF, папку или ZIP. После анализа результат останется здесь.")
            return

        lines = [
            f"PDF-файлов найдено: {self.pdf_count}",
            f"Распознано инвойсов: {len(self.report.invoices)}",
            f"Не распознано PDF: {len(self.report.error_files)}",
            f"Исключено дубликатов: {len(self.report.duplicates)}",
            f"Общий вес нетто: {human_total_weight(self.report.total_weight)}",
            f"Общий вес брутто: {human_total_weight(self.report.total_gross_weight)}",
            f"Количество мест: {self.report.total_packages}",
            f"Общая сумма USD: {human_total_usd(self.report.total_usd)}",
            f"Страны происхождения: {self.format_report_countries()}",
            "",
            "Инвойсы:",
        ]
        for record in self.report.invoices:
            lines.append(
                f"- {record.invoice_no}: {record.net_weight:.3f} kg, "
                f"gross: {_format_optional_weight(record.gross_weight)}, "
                f"{record.total_usd:,.2f} USD, мест: {record.package_count or '-'}, "
                f"позиций: {record.positions}, "
                f"страны: {format_countries(record.origin_countries)}, файл: {record.source_file}"
            )
        if self.report.duplicates:
            lines.append("")
            lines.append("Дубликаты:")
            for duplicate in self.report.duplicates:
                lines.append(
                    f"- {duplicate.invoice_no}: оставлен {duplicate.kept_file}, исключен {duplicate.excluded_file}"
                )
        if self.report.error_files:
            lines.append("")
            lines.append("Не распознаны PDF:")
            for source_file in self.report.error_files:
                lines.append(f"- {source_file}")
        if self.report.issues:
            lines.append(f"Заметки/предупреждения: {len(self.report.issues)}. Подробности в логе.")
            for issue in self.report.issues:
                self.log(f"[{issue.level}] {issue.source_file}: {issue.message}")
        self.result_text.setPlainText("\n".join(lines))

    def save_excel(self) -> None:
        if not (self.report.invoices or self.report.issues or self.report.duplicates):
            QMessageBox.warning(self, "Нет данных", "Сначала выполните анализ.")
            return

        path, _ = QFileDialog.getSaveFileName(self, "Сохранить Excel", "cummins_invoice_summary.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        destination = ensure_xlsx_path(path)
        try:
            export_report_to_excel(self.report, destination)
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка сохранения", str(exc))
            self.log(f"[ERROR] Ошибка сохранения Excel: {exc}")
            return
        self.set_status(f"Excel сохранен: {destination.name}")
        self.log(f"[INFO] Excel сохранен: {destination}")

    def show_about(self) -> None:
        QMessageBox.about(
            self,
            "О программе",
            f"<b>{APP_TITLE}</b><br>"
            f"Версия: {APP_VERSION}<br>"
            f"Parser build: {PARSER_BUILD}<br><br>"
            f"Назначение: {APP_SUBTITLE}<br><br>"
            f"Контроль качества: сравнение движков PDF, предупреждения о расхождениях, "
            f"регрессионная проверка известных наборов инвойсов.<br><br>"
            f"Разработчик: {DEVELOPER_NAME}<br>"
            f"Email: {DEVELOPER_EMAIL}<br>"
            f"Телефон: {DEVELOPER_PHONE}",
        )

    def clear(self) -> None:
        self.source_type = None
        self.source_path = None
        self.report = ParseReport()
        self.pdf_count = 0
        self.source_label.setText("Источник не выбран")
        self.countries_label.setText("Страны появятся после анализа")
        self.result_text.setPlainText("Выберите PDF, папку или ZIP. После анализа результат останется здесь.")
        self.log_text.clear()
        self.set_status("Очищено")
        self.render_report()

    def log(self, text: str) -> None:
        self.log_text.append(text)

    def set_status(self, status: str) -> None:
        self.status_label.setText(status)
        self.setWindowTitle(f"{APP_TITLE} {APP_VERSION} | parser {PARSER_BUILD} | {status}")

    def _render_empty(self) -> None:
        self.source_label.setText("Источник не выбран")
        self.countries_label.setText("Страны появятся после анализа")
        self.result_text.setPlainText("Выберите PDF, папку или ZIP. После анализа результат останется здесь.")
        self.invoice_card.set_value("0")
        self.weight_card.set_value("0.000 kg")
        self.gross_weight_card.set_value("0.000 kg")
        self.usd_card.set_value("0.00 USD")
        self.log(f"[INFO] Версия: {APP_VERSION}; parser build: {PARSER_BUILD}")
        self.log("[INFO] Приложение запущено.")

    def format_report_countries(self) -> str:
        countries: list[str] = []
        seen: set[str] = set()
        for record in self.report.invoices:
            for country in record.origin_countries:
                if country in seen:
                    continue
                seen.add(country)
                countries.append(country)
        return format_countries(countries)


def main() -> None:
    app = QApplication(sys.argv)
    if ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(ICON_PATH)))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
