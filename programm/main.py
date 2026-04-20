"""台股程式交易回測系統 — 進入點"""
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QFont, QFontDatabase
from ui_main import MainWindow


def _pick_font() -> QFont:
    """選擇系統中可用的中文字型（正黑 → 雅黑 → 宋體 → 系統預設）"""
    available = QFontDatabase().families()
    for name in ('Microsoft JhengHei', 'Microsoft YaHei',
                 'Microsoft YaHei UI', 'SimHei', 'NSimSun'):
        if name in available:
            return QFont(name, 11)
    return QFont()   # 系統預設


def main():
    app = QApplication(sys.argv)
    app.setFont(_pick_font())
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
