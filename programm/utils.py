"""資料抓取、儲存工具函式"""
import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta

# 設定中文字型（Windows）
matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial Unicode MS']
matplotlib.rcParams['axes.unicode_minus'] = False


def fetch_data(ticker: str, interval: str = '30m') -> tuple[pd.DataFrame, str]:
    """
    從 yfinance 抓取台股資料
    ticker : 台灣股票代號（如 '2330'，自動附加 .TW）
    interval: '30m' 或 '60m'
    回傳 (df, symbol_string)
    """
    symbol = ticker.strip().upper()
    if not symbol.endswith('.TW'):
        symbol = symbol + '.TW'

    # yfinance 免費 API 上限：30m 底層為 15m 聚合，最多 59 天；60m 最多 365 天
    days  = 59 if interval == '30m' else 365
    end   = datetime.now()
    start = end - timedelta(days=days)

    df = yf.download(
        symbol, start=start, end=end,
        interval=interval, progress=False, auto_adjust=True,
    )

    if df is None or df.empty:
        raise ValueError(f'無法取得 {symbol} 的資料，請確認股票代號是否正確或網路連線。')

    # 展平多層欄位（yfinance 多版本相容）
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)

    # 移除 timezone（避免 mplfinance / matplotlib 顯示問題）
    df.index = pd.to_datetime(df.index)
    if df.index.tz is not None:
        df.index = df.index.tz_localize(None)

    needed = [c for c in ['Open', 'High', 'Low', 'Close', 'Volume'] if c in df.columns]
    df = df[needed].dropna()

    if len(df) < 20:
        raise ValueError(f'{symbol} 資料筆數不足（{len(df)} 筆），無法計算技術指標，請換其他週期或代號。')

    return df, symbol


def save_to_excel(df: pd.DataFrame, path: str = 'stock_data.xlsx') -> str:
    df.to_excel(path, index=True)
    return path


def save_chart_png(fig, path: str = 'chart.png') -> str:
    fig.savefig(path, dpi=150, bbox_inches='tight')
    return path
