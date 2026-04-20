"""台股程式交易回測系統 — 主視窗 (PyQt5)"""
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Qt5Agg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QComboBox, QPushButton, QGroupBox,
    QRadioButton, QTableWidget, QTableWidgetItem, QSplitter,
    QFileDialog, QMessageBox, QStatusBar, QHeaderView,
    QButtonGroup, QSizePolicy,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor

from utils      import fetch_data, save_to_excel, save_chart_png
from indicators import calculate_all
from strategy   import apply_strategy
from backtest   import run_backtest

plt.rcParams['font.sans-serif'] = [
    'Microsoft JhengHei', 'Microsoft YaHei', 'Microsoft YaHei UI',
    'SimHei', 'Arial Unicode MS',
]
plt.rcParams['axes.unicode_minus'] = False

STRATEGIES = ['均線糾結策略', 'RSI策略', 'MACD策略', '布林通道策略']

# ─────────────────────────────────────────────────────────────────
#  淺色 / 深色樣式
# ─────────────────────────────────────────────────────────────────

LIGHT_STYLE = """
QMainWindow          { background-color: #f0f2f5; }
QWidget              { font-size: 13px; color: #1a1a2e; }
QGroupBox {
    font-weight: bold; font-size: 13px;
    border: 1px solid #c8cdd6; border-radius: 6px;
    margin-top: 14px; padding: 10px 8px 8px 8px;
    background-color: #ffffff;
}
QGroupBox::title     { subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #1e3a5f; }
QLineEdit, QComboBox {
    border: 1px solid #b8bec9; border-radius: 4px;
    padding: 5px 8px; background: #ffffff; min-height: 30px; font-size: 13px;
}
QLineEdit:focus, QComboBox:focus { border: 1.5px solid #2563eb; }
QPushButton#runBtn {
    background-color: #2563eb; color: #ffffff; border: none; border-radius: 5px;
    padding: 10px 28px; font-weight: bold; font-size: 15px; min-height: 42px;
}
QPushButton#runBtn:hover    { background-color: #1d4ed8; }
QPushButton#runBtn:pressed  { background-color: #1e40af; }
QPushButton#runBtn:disabled { background-color: #94a3b8; }
QPushButton#saveBtn {
    background-color: #475569; color: #ffffff; border: none; border-radius: 4px;
    padding: 8px 18px; font-size: 13px; min-height: 38px;
}
QPushButton#saveBtn:hover   { background-color: #334155; }
QPushButton#themeBtn {
    background-color: #e2e8f0; color: #334155; border: 1px solid #cbd5e1;
    border-radius: 4px; padding: 6px 14px; font-size: 12px; min-height: 34px;
}
QPushButton#themeBtn:hover  { background-color: #cbd5e1; }
QRadioButton            { spacing: 6px; font-size: 13px; }
QRadioButton::indicator { width: 15px; height: 15px; }
QTableWidget {
    border: 1px solid #c8cdd6; border-radius: 3px;
    gridline-color: #e2e8f0;
    selection-background-color: #bfdbfe; selection-color: #000;
    alternate-background-color: #f1f5f9;
}
QTableWidget QHeaderView::section {
    background-color: #1e3a5f; color: #ffffff;
    font-weight: bold; padding: 4px; border: none; font-size: 12px;
}
QScrollBar:vertical   { width: 9px; }
QScrollBar:horizontal { height: 9px; }
QStatusBar  { background: #dbeafe; color: #1e3a5f; font-size: 12px; padding-left: 6px; }
QSplitter::handle { background: #cbd5e1; }
"""

DARK_STYLE = """
QMainWindow          { background-color: #0f172a; }
QWidget              { font-size: 13px; color: #e2e8f0; }
QGroupBox {
    font-weight: bold; font-size: 13px;
    border: 1px solid #334155; border-radius: 6px;
    margin-top: 14px; padding: 10px 8px 8px 8px;
    background-color: #1e293b;
}
QGroupBox::title     { subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #60a5fa; }
QLineEdit, QComboBox {
    border: 1px solid #334155; border-radius: 4px;
    padding: 5px 8px; background: #1e293b; color: #e2e8f0;
    min-height: 30px; font-size: 13px;
}
QLineEdit:focus, QComboBox:focus { border: 1.5px solid #3b82f6; }
QComboBox QAbstractItemView {
    background: #1e293b; color: #e2e8f0;
    selection-background-color: #3b82f6; selection-color: #fff;
}
QPushButton#runBtn {
    background-color: #3b82f6; color: #ffffff; border: none; border-radius: 5px;
    padding: 10px 28px; font-weight: bold; font-size: 15px; min-height: 42px;
}
QPushButton#runBtn:hover    { background-color: #2563eb; }
QPushButton#runBtn:pressed  { background-color: #1d4ed8; }
QPushButton#runBtn:disabled { background-color: #475569; }
QPushButton#saveBtn {
    background-color: #334155; color: #e2e8f0; border: none; border-radius: 4px;
    padding: 8px 18px; font-size: 13px; min-height: 38px;
}
QPushButton#saveBtn:hover   { background-color: #475569; }
QPushButton#themeBtn {
    background-color: #1e293b; color: #94a3b8; border: 1px solid #334155;
    border-radius: 4px; padding: 6px 14px; font-size: 12px; min-height: 34px;
}
QPushButton#themeBtn:hover  { background-color: #334155; color: #e2e8f0; }
QRadioButton            { spacing: 6px; font-size: 13px; color: #e2e8f0; }
QRadioButton::indicator { width: 15px; height: 15px; }
QTableWidget {
    border: 1px solid #334155; border-radius: 3px;
    background-color: #1e293b; color: #e2e8f0;
    gridline-color: #334155;
    selection-background-color: #3b82f6; selection-color: #fff;
    alternate-background-color: #172033;
}
QTableWidget QHeaderView::section {
    background-color: #0f172a; color: #60a5fa;
    font-weight: bold; padding: 4px; border: none; font-size: 12px;
    border-bottom: 1px solid #334155;
}
QScrollBar:vertical   { width: 9px; background: #1e293b; }
QScrollBar:horizontal { height: 9px; background: #1e293b; }
QStatusBar  { background: #0f172a; color: #60a5fa; font-size: 12px; padding-left: 6px; }
QSplitter::handle { background: #334155; }
"""

# ─────────────────────────────────────────────────────────────────
#  圖表顏色配置
# ─────────────────────────────────────────────────────────────────

CHART_LIGHT = dict(
    fig_bg   = '#f8f8f8',
    ax_bg    = '#f8f8f8',
    grid     = '#e2e8f0',
    text     = '#1a1a2e',
    up       = '#2a9d8f',
    down     = '#e76f51',
    ma5      = '#3b82f6',
    ma10     = '#f59e0b',
    ma20     = '#8b5cf6',
    bb       = '#64748b',
    rsi_line = '#6d28d9',
    rsi_ob   = '#dc2626',
    rsi_os   = '#16a34a',
    dif      = '#3b82f6',
    macd     = '#f59e0b',
    hist_pos = '#2a9d8f',
    hist_neg = '#e76f51',
    signal_b = '#16a34a',
    signal_s = '#dc2626',
    legend_fc= '#ffffff',
    legend_ec= '#c8cdd6',
)

CHART_DARK = dict(
    fig_bg   = '#0f172a',
    ax_bg    = '#1e293b',
    grid     = '#2d3748',
    text     = '#e2e8f0',
    up       = '#00c9a7',
    down     = '#ff6b6b',
    ma5      = '#60a5fa',
    ma10     = '#fbbf24',
    ma20     = '#a78bfa',
    bb       = '#94a3b8',
    rsi_line = '#c084fc',
    rsi_ob   = '#f87171',
    rsi_os   = '#4ade80',
    dif      = '#60a5fa',
    macd     = '#fbbf24',
    hist_pos = '#00c9a7',
    hist_neg = '#ff6b6b',
    signal_b = '#4ade80',
    signal_s = '#f87171',
    legend_fc= '#1e293b',
    legend_ec= '#334155',
)

# ─────────────────────────────────────────────────────────────────
#  背景工作執行緒
# ─────────────────────────────────────────────────────────────────

class BacktestWorker(QThread):
    finished = pyqtSignal(object, object, str)   # df, kpi, symbol
    error    = pyqtSignal(str)

    def __init__(self, ticker, interval, strategy):
        super().__init__()
        self.ticker   = ticker
        self.interval = interval
        self.strategy = strategy

    def run(self):
        try:
            df, symbol = fetch_data(self.ticker, self.interval)
            df         = calculate_all(df)
            df         = apply_strategy(df, self.strategy)
            kpi = run_backtest(df)[2]
            self.finished.emit(df, kpi, symbol)
        except Exception as exc:
            self.error.emit(str(exc))


# ─────────────────────────────────────────────────────────────────
#  Matplotlib 嵌入畫布
# ─────────────────────────────────────────────────────────────────

class ChartCanvas(FigureCanvas):
    def __init__(self):
        self.fig = Figure(figsize=(10, 6), dpi=100)
        super().__init__(self.fig)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._c = CHART_LIGHT
        self._apply_bg()
        self._placeholder()

    def _apply_bg(self):
        self.fig.set_facecolor(self._c['fig_bg'])

    def set_dark(self, dark: bool):
        self._c = CHART_DARK if dark else CHART_LIGHT
        self._apply_bg()
        self._placeholder()

    def _placeholder(self):
        self.fig.clear()
        self._apply_bg()
        ax = self.fig.add_subplot(111)
        ax.set_facecolor(self._c['ax_bg'])
        ax.text(0.5, 0.5, '請輸入股票代號並按下「開始回測」',
                ha='center', va='center', fontsize=13,
                color=self._c['text'], alpha=0.45, transform=ax.transAxes)
        ax.axis('off')
        self.draw()

    # ── 輔助：蠟燭 ─────────────────────────────────────────────
    def _draw_candles(self, ax, df_plot):
        for i, (_, row) in enumerate(df_plot.iterrows()):
            up    = float(row['Close']) >= float(row['Open'])
            color = self._c['up'] if up else self._c['down']
            bbot  = min(float(row['Open']), float(row['Close']))
            bh    = abs(float(row['Close']) - float(row['Open'])) or float(row['Close']) * 0.001
            ax.bar(i, bh, bottom=bbot, color=color, width=0.55, linewidth=0)
            ax.plot([i, i], [float(row['Low']), float(row['High'])],
                    color=color, linewidth=0.9)

    # ── 輔助：買賣標記 ──────────────────────────────────────────
    def _draw_signals(self, ax, df_plot):
        if 'Signal' not in df_plot.columns:
            return
        for i, (_, row) in enumerate(df_plot.iterrows()):
            v = row['Signal']
            sig = 0 if pd.isna(v) else int(v)
            if sig == 1:
                ax.scatter(i, float(row['Low'])  * 0.997, marker='^',
                           color=self._c['signal_b'], s=90, zorder=5)
            elif sig == -1:
                ax.scatter(i, float(row['High']) * 1.003, marker='v',
                           color=self._c['signal_s'], s=90, zorder=5)

    # ── 輔助：坐標軸共同設定 ────────────────────────────────────
    def _style_ax(self, ax, ticks, labels, ylabel=''):
        ax.set_facecolor(self._c['ax_bg'])
        ax.tick_params(colors=self._c['text'], labelsize=7)
        ax.yaxis.label.set_color(self._c['text'])
        for sp in ax.spines.values():
            sp.set_color(self._c['grid'])
        ax.grid(True, color=self._c['grid'], linewidth=0.6)
        ax.set_xticks(ticks)
        ax.set_xticklabels(labels, rotation=25, ha='right',
                           fontsize=7, color=self._c['text'])
        if ylabel:
            ax.set_ylabel(ylabel, fontsize=8, color=self._c['text'])

    # ── 主繪圖 ──────────────────────────────────────────────────
    def plot(self, df: pd.DataFrame, symbol: str, strategy: str) -> Figure:
        self.fig.clear()
        self._apply_bg()

        c        = self._c
        df_plot  = df.tail(50).copy()
        n        = len(df_plot)
        xs       = list(range(n))
        step     = max(1, n // 8)
        ticks    = list(range(0, n, step))
        labels   = [str(df_plot.index[t])[:16] for t in ticks]

        show_rsi  = (strategy == 'RSI策略')
        show_macd = (strategy == 'MACD策略')
        has_sub   = show_rsi or show_macd

        if has_sub:
            ax1 = self.fig.add_axes([0.07, 0.38, 0.91, 0.56])
            ax2 = self.fig.add_axes([0.07, 0.07, 0.91, 0.26])
        else:
            ax1 = self.fig.add_axes([0.07, 0.12, 0.91, 0.80])

        # ── K 線 ────────────────────────────────────────────────
        self._draw_candles(ax1, df_plot)

        # ── 均線（全策略顯示）────────────────────────────────────
        for col, clr, lbl in [('MA5', c['ma5'], 'MA5'),
                               ('MA10', c['ma10'], 'MA10'),
                               ('MA20', c['ma20'], 'MA20')]:
            ax1.plot(xs, df_plot[col].values, color=clr, linewidth=1.2, label=lbl)

        # ── 布林通道 ────────────────────────────────────────────
        if strategy == '布林通道策略':
            upper = df_plot['BB_UPPER'].values
            lower = df_plot['BB_LOWER'].values
            mid   = df_plot['BB_MID'].values
            ax1.plot(xs, upper, color=c['bb'], linewidth=1.0,
                     linestyle='--', label='BB上軌')
            ax1.plot(xs, mid,   color=c['bb'], linewidth=0.8,
                     linestyle=':',  label='BB中軌')
            ax1.plot(xs, lower, color=c['bb'], linewidth=1.0,
                     linestyle='--', label='BB下軌')
            ax1.fill_between(xs, lower, upper, alpha=0.07, color=c['bb'])

        # ── 買賣訊號 ────────────────────────────────────────────
        self._draw_signals(ax1, df_plot)

        # ── 主圖修飾 ────────────────────────────────────────────
        self._style_ax(ax1, ticks, labels)
        ax1.set_title(f'{symbol}  近50筆K線走勢圖  [{strategy}]',
                      fontsize=11, fontweight='bold', pad=6, color=c['text'])
        ax1.legend(loc='upper left', fontsize=7, framealpha=0.8, ncol=2,
                   facecolor=c['legend_fc'], edgecolor=c['legend_ec'],
                   labelcolor=c['text'])

        # ── RSI 子圖 ────────────────────────────────────────────
        if show_rsi:
            rsi    = np.array(df_plot['RSI'].values, dtype=float)
            rsi_nn = np.nan_to_num(rsi, nan=50.0)
            ax2.plot(xs, rsi, color=c['rsi_line'], linewidth=1.2, label='RSI(14)')
            ax2.axhline(70, color=c['rsi_ob'], linestyle='--', linewidth=0.85, alpha=0.7)
            ax2.axhline(30, color=c['rsi_os'], linestyle='--', linewidth=0.85, alpha=0.7)
            ax2.fill_between(xs, 30, rsi_nn,
                             where=(rsi_nn < 30), alpha=0.15, color=c['rsi_os'])
            ax2.fill_between(xs, 70, rsi_nn,
                             where=(rsi_nn > 70), alpha=0.15, color=c['rsi_ob'])
            ax2.set_ylim(0, 100)
            self._style_ax(ax2, ticks, labels, ylabel='RSI')
            ax2.legend(loc='upper left', fontsize=7,
                       facecolor=c['legend_fc'], edgecolor=c['legend_ec'],
                       labelcolor=c['text'])

        # ── MACD 子圖 ───────────────────────────────────────────
        if show_macd:
            dif  = np.array(df_plot['DIF'].values,       dtype=float)
            sig  = np.array(df_plot['MACD_SIG'].values,  dtype=float)
            hist = np.array(df_plot['MACD_HIST'].values, dtype=float)
            bar_colors = [c['hist_pos'] if h >= 0 else c['hist_neg']
                          for h in np.nan_to_num(hist)]
            ax2.bar(xs, hist, color=bar_colors, width=0.55, alpha=0.7, label='Histogram')
            ax2.plot(xs, dif, color=c['dif'],  linewidth=1.2, label='DIF')
            ax2.plot(xs, sig, color=c['macd'], linewidth=1.2, label='MACD')
            ax2.axhline(0, color=c['grid'], linewidth=0.7)
            self._style_ax(ax2, ticks, labels, ylabel='MACD')
            ax2.legend(loc='upper left', fontsize=7, ncol=3,
                       facecolor=c['legend_fc'], edgecolor=c['legend_ec'],
                       labelcolor=c['text'])

        self.draw()
        return self.fig


# ─────────────────────────────────────────────────────────────────
#  主視窗
# ─────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('台股程式交易回測系統')
        self.setMinimumSize(1200, 800)
        self.resize(1440, 900)

        self._df     = None
        self._symbol = ''
        self._fig    = None
        self._worker = None
        self._dark   = False          # 預設淺色

        self._build_ui()
        self.setStyleSheet(LIGHT_STYLE)
        self._set_status('就緒  |  30分K = 近59天  /  60分K = 近1年  |  輸入股票代號後按下「開始回測」')

    # ── UI 建構 ──────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 8, 10, 6)
        root.setSpacing(6)

        root.addWidget(self._build_input_panel())

        mid = QSplitter(Qt.Horizontal)
        mid.addWidget(self._build_chart_area())
        mid.addWidget(self._build_kpi_panel())
        mid.setSizes([900, 320])
        root.addWidget(mid, stretch=5)

        root.addWidget(self._build_data_table(), stretch=3)

        self._status_bar = QStatusBar()
        self.setStatusBar(self._status_bar)

    # ── 輸入區 ───────────────────────────────────────────────────

    def _build_input_panel(self) -> QGroupBox:
        grp = QGroupBox('輸入設定')
        lay = QHBoxLayout(grp)
        lay.setSpacing(12)

        lay.addWidget(QLabel('股票代號：'))
        self.ticker_input = QLineEdit()
        self.ticker_input.setPlaceholderText('例：2330')
        self.ticker_input.setFixedWidth(95)
        self.ticker_input.returnPressed.connect(self._on_run)
        lay.addWidget(self.ticker_input)

        lay.addWidget(QLabel('K線週期：'))
        self.interval_combo = QComboBox()
        self.interval_combo.addItems(['30分鐘K  (近59日)', '60分鐘K  (近1年)'])
        self.interval_combo.setFixedWidth(165)
        lay.addWidget(self.interval_combo)

        # 策略選擇（2×2 格線）
        strat_grp  = QGroupBox('策略選擇')
        strat_grid = QGridLayout(strat_grp)
        strat_grid.setHorizontalSpacing(20)
        strat_grid.setVerticalSpacing(4)

        self._btn_grp = QButtonGroup(self)
        labels = ['均線糾結策略', 'RSI 超買超賣', 'MACD 交叉', '布林通道']
        self._radios = []
        for idx, lbl in enumerate(labels):
            row, col = idx // 2, idx % 2
            rb = QRadioButton(lbl)
            self._btn_grp.addButton(rb, idx)
            strat_grid.addWidget(rb, row, col)
            self._radios.append(rb)
        self._radios[0].setChecked(True)
        lay.addWidget(strat_grp)

        lay.addStretch()

        # 深色切換
        self.theme_btn = QPushButton('深色模式')
        self.theme_btn.setObjectName('themeBtn')
        self.theme_btn.setCheckable(True)
        self.theme_btn.clicked.connect(self._toggle_theme)
        lay.addWidget(self.theme_btn)

        self.run_btn = QPushButton('開始回測')
        self.run_btn.setObjectName('runBtn')
        self.run_btn.clicked.connect(self._on_run)
        lay.addWidget(self.run_btn)

        excel_btn = QPushButton('儲存 Excel')
        excel_btn.setObjectName('saveBtn')
        excel_btn.clicked.connect(self._save_excel)
        lay.addWidget(excel_btn)

        png_btn = QPushButton('儲存圖表')
        png_btn.setObjectName('saveBtn')
        png_btn.clicked.connect(self._save_png)
        lay.addWidget(png_btn)

        return grp

    # ── K 線圖區 ─────────────────────────────────────────────────

    def _build_chart_area(self) -> QGroupBox:
        grp = QGroupBox('K線走勢圖')
        lay = QVBoxLayout(grp)
        lay.setContentsMargins(4, 4, 4, 4)
        self.canvas = ChartCanvas()
        lay.addWidget(self.canvas)
        return grp

    # ── KPI 表格 ─────────────────────────────────────────────────

    def _build_kpi_panel(self) -> QGroupBox:
        grp = QGroupBox('投資績效 KPI')
        lay = QVBoxLayout(grp)
        self.kpi_table = QTableWidget(0, 2)
        self.kpi_table.setHorizontalHeaderLabels(['指標', '數值'])
        self.kpi_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.kpi_table.verticalHeader().setVisible(False)
        self.kpi_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.kpi_table.setAlternatingRowColors(True)
        lay.addWidget(self.kpi_table)
        return grp

    # ── K 線資料表格 ─────────────────────────────────────────────

    def _build_data_table(self) -> QGroupBox:
        grp  = QGroupBox('最後 50 筆 K 線資料')
        lay  = QVBoxLayout(grp)
        cols = ['日期時間', '開盤 Open', '最高 High', '最低 Low',
                '收盤 Close', '成交量 Volume']
        self.data_table = QTableWidget(0, len(cols))
        self.data_table.setHorizontalHeaderLabels(cols)
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.data_table.verticalHeader().setVisible(False)
        self.data_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setMaximumHeight(250)
        lay.addWidget(self.data_table)
        return grp

    # ── 主題切換 ─────────────────────────────────────────────────

    def _toggle_theme(self, checked: bool):
        self._dark = checked
        self.setStyleSheet(DARK_STYLE if checked else LIGHT_STYLE)
        self.theme_btn.setText('淺色模式' if checked else '深色模式')
        self.canvas.set_dark(checked)
        # 若已有圖表，重新繪製以套用新配色
        if self._df is not None:
            self._fig = self.canvas.plot(
                self._df, self._symbol, self._strategy_name())

    # ── 事件處理 ─────────────────────────────────────────────────

    def _strategy_name(self) -> str:
        return STRATEGIES[self._btn_grp.checkedId()]

    def _interval_code(self) -> str:
        return '30m' if self.interval_combo.currentIndex() == 0 else '60m'

    def _on_run(self):
        ticker = self.ticker_input.text().strip()
        if not ticker:
            QMessageBox.warning(self, '輸入錯誤', '請輸入股票代號。')
            return

        self.run_btn.setEnabled(False)
        self.run_btn.setText('回測中...')
        interval = self._interval_code()
        desc     = '59日' if interval == '30m' else '1年'
        self._set_status(f'正在抓取 {ticker}（{desc}，{interval}）...')

        self._worker = BacktestWorker(ticker, interval, self._strategy_name())
        self._worker.finished.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.start()

    def _on_done(self, df, kpi, symbol):
        self._df     = df
        self._symbol = symbol

        self._fill_data_table(df)
        self._fill_kpi_table(kpi)
        self.canvas.set_dark(self._dark)
        self._fig = self.canvas.plot(df, symbol, self._strategy_name())

        total = kpi.get('總交易筆數', '0')
        roi   = kpi.get('投資報酬率',  'N/A')
        bars  = len(df)
        start = str(df.index[0])[:16]
        end   = str(df.index[-1])[:16]
        self._set_status(
            f'回測完成  |  {symbol}  |  資料：{bars}筆（{start} ~ {end}）'
            f'  |  交易：{total}筆  |  報酬率：{roi}')

        self.run_btn.setEnabled(True)
        self.run_btn.setText('開始回測')

    def _on_error(self, msg):
        QMessageBox.critical(self, '執行錯誤', msg)
        self._set_status(f'錯誤：{msg}')
        self.run_btn.setEnabled(True)
        self.run_btn.setText('開始回測')

    # ── 填充表格 ─────────────────────────────────────────────────

    def _fill_data_table(self, df: pd.DataFrame):
        df50 = df.tail(50)
        self.data_table.setRowCount(len(df50))
        for r, (ts, row) in enumerate(df50.iterrows()):
            vals = [str(ts)[:19],
                    f'{float(row["Open"]):.2f}',
                    f'{float(row["High"]):.2f}',
                    f'{float(row["Low"]):.2f}',
                    f'{float(row["Close"]):.2f}',
                    f'{int(row["Volume"]):,}']
            for c, v in enumerate(vals):
                item = QTableWidgetItem(v)
                item.setTextAlignment(Qt.AlignCenter)
                self.data_table.setItem(r, c, item)
        self.data_table.scrollToBottom()

    def _fill_kpi_table(self, kpi: dict):
        self.kpi_table.setRowCount(len(kpi))
        for r, (name, val) in enumerate(kpi.items()):
            n_item = QTableWidgetItem(name)
            v_item = QTableWidgetItem(str(val))
            n_item.setTextAlignment(Qt.AlignLeft  | Qt.AlignVCenter)
            v_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

            if '報酬率' in name and '%' in str(val):
                try:
                    pct = float(str(val).replace('%', '').replace('+', ''))
                    if self._dark:
                        v_item.setBackground(
                            QColor('#14532d') if pct >= 0 else QColor('#450a0a'))
                        v_item.setForeground(
                            QColor('#4ade80') if pct >= 0 else QColor('#f87171'))
                    else:
                        v_item.setBackground(
                            QColor('#dcfce7') if pct >= 0 else QColor('#fee2e2'))
                        v_item.setForeground(
                            QColor('#14532d') if pct >= 0 else QColor('#7f1d1d'))
                except ValueError:
                    pass

            self.kpi_table.setItem(r, 0, n_item)
            self.kpi_table.setItem(r, 1, v_item)

    # ── 儲存功能 ─────────────────────────────────────────────────

    def _save_excel(self):
        if self._df is None:
            QMessageBox.information(self, '提示', '請先執行回測。')
            return
        path, _ = QFileDialog.getSaveFileName(
            self, '儲存 Excel', f'{self._symbol}_data.xlsx', 'Excel (*.xlsx)')
        if path:
            save_to_excel(self._df, path)
            self._set_status(f'Excel 已儲存：{path}')

    def _save_png(self):
        if self._fig is None:
            QMessageBox.information(self, '提示', '請先執行回測。')
            return
        path, _ = QFileDialog.getSaveFileName(
            self, '儲存圖表', f'{self._symbol}_chart.png', 'PNG (*.png)')
        if path:
            save_chart_png(self._fig, path)
            self._set_status(f'圖表已儲存：{path}')

    def _set_status(self, msg: str):
        self._status_bar.showMessage(msg)
