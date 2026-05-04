from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QRectF
from PySide6.QtGui import QColor, QFont, QGuiApplication, QImage, QPainter, QPen


ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
ASSETS.mkdir(exist_ok=True)


def draw_icon(size: int) -> QImage:
    image = QImage(size, size, QImage.Format_ARGB32)
    image.fill(Qt.GlobalColor.transparent)

    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    outer = QRectF(size * 0.06, size * 0.06, size * 0.88, size * 0.88)
    inner = QRectF(size * 0.12, size * 0.12, size * 0.76, size * 0.76)

    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor("#102033"))
    painter.drawRoundedRect(outer, size * 0.18, size * 0.18)

    painter.setBrush(QColor("#f4b63c"))
    painter.drawRoundedRect(inner, size * 0.16, size * 0.16)

    painter.setBrush(QColor("#ffffff"))
    paper = QRectF(size * 0.28, size * 0.18, size * 0.44, size * 0.54)
    painter.drawRoundedRect(paper, size * 0.05, size * 0.05)

    painter.setPen(QPen(QColor("#102033"), max(2, size // 60)))
    for row in range(3):
        y = size * (0.29 + row * 0.09)
        painter.drawLine(int(size * 0.34), int(y), int(size * 0.66), int(y))

    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor("#102033"))
    painter.drawRoundedRect(QRectF(size * 0.33, size * 0.61, size * 0.34, size * 0.11), size * 0.03, size * 0.03)

    painter.setPen(QColor("#ffffff"))
    font = QFont("Helvetica", max(10, size // 8), QFont.Weight.Bold)
    painter.setFont(font)
    painter.drawText(QRectF(size * 0.33, size * 0.61, size * 0.34, size * 0.11), Qt.AlignmentFlag.AlignCenter, "CI")

    painter.end()
    return image


def main() -> None:
    app = QGuiApplication([])
    png_path = ASSETS / "app_icon_512.png"
    ico_path = ASSETS / "app_icon.ico"
    icns_path = ASSETS / "app_icon.icns"

    image_512 = draw_icon(512)
    image_512.save(str(png_path), "PNG")
    image_512.save(str(ico_path), "ICO")
    image_512.save(str(icns_path), "ICNS")

    print(f"Created: {png_path}")
    print(f"Created: {ico_path}")
    print(f"Created: {icns_path}")
    app.quit()


if __name__ == "__main__":
    main()
