from PySide6.QtWidgets import QWidget, QVBoxLayout, QScrollArea
from PySide6.QtCore import Qt


def create_scroll_layout(parent):
    """
    يعمل Layout قابل للـ Scroll لأي شاشة.
    استخدمه بدل QVBoxLayout(self).
    """
    outer_layout = QVBoxLayout(parent)

    scroll_area = QScrollArea()
    scroll_area.setWidgetResizable(True)
    scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
    scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)

    content_widget = QWidget()
    content_layout = QVBoxLayout(content_widget)

    scroll_area.setWidget(content_widget)
    outer_layout.addWidget(scroll_area)

    return content_layout