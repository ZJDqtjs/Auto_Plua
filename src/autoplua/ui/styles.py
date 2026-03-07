from __future__ import annotations

SIDEBAR_EXPANDED_STYLE = """
QPushButton {
    border: none;
    padding: 14px 16px;
    font-size: 16px;
    text-align: left;
    background-color: transparent;
    qproperty-iconSize: 22px 22px;
}
QPushButton:checked {
    background-color: #e5f1ff;
    color: #0b6ecf;
    font-weight: 700;
    border-left: 4px solid #0b6ecf;
}
QPushButton:hover:!checked {
    background-color: #f3f3f3;
}
"""


SIDEBAR_COLLAPSED_STYLE = """
QPushButton {
    border: none;
    padding: 0;
    font-size: 16px;
    text-align: center;
    background-color: transparent;
    qproperty-iconSize: 22px 22px;
}
QPushButton:checked {
    background-color: #e5f1ff;
    color: #0b6ecf;
    font-weight: 700;
    border-left: 4px solid #0b6ecf;
}
QPushButton:hover:!checked {
    background-color: #f3f3f3;
}
"""
