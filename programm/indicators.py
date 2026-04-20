"""技術指標計算模組"""
import numpy as np
import pandas as pd


def calculate_ma(df: pd.DataFrame) -> pd.DataFrame:
    for w in [5, 10, 20]:
        df[f'MA{w}'] = df['Close'].rolling(w).mean()
    return df


def calculate_rsi(df: pd.DataFrame, period: int = 14) -> pd.DataFrame:
    delta = df['Close'].diff()
    gain  = delta.clip(lower=0).rolling(period).mean()
    loss  = (-delta.clip(upper=0)).rolling(period).mean()
    rs    = gain / loss.replace(0, np.nan)
    df['RSI'] = 100 - (100 / (1 + rs))
    return df


def calculate_macd(df: pd.DataFrame, fast=12, slow=26, signal=9) -> pd.DataFrame:
    ema_fast        = df['Close'].ewm(span=fast,   adjust=False).mean()
    ema_slow        = df['Close'].ewm(span=slow,   adjust=False).mean()
    df['DIF']       = ema_fast - ema_slow
    df['MACD_SIG']  = df['DIF'].ewm(span=signal,  adjust=False).mean()
    df['MACD_HIST'] = df['DIF'] - df['MACD_SIG']
    return df


def calculate_bollinger(df: pd.DataFrame, period: int = 20, mult: float = 2.0) -> pd.DataFrame:
    mid              = df['Close'].rolling(period).mean()
    std              = df['Close'].rolling(period).std()
    df['BB_MID']     = mid
    df['BB_UPPER']   = mid + mult * std
    df['BB_LOWER']   = mid - mult * std
    return df


def calculate_ma_convergence(df: pd.DataFrame) -> pd.DataFrame:
    """均線糾結：三均線最大最小差距 / 收盤價 <= 0.5%"""
    ma_cols = ['MA5', 'MA10', 'MA20']
    valid   = df[ma_cols].notna().all(axis=1)
    ma_max  = df.loc[valid, ma_cols].max(axis=1)
    ma_min  = df.loc[valid, ma_cols].min(axis=1)
    df['MA_Conv'] = False
    df.loc[valid, 'MA_Conv'] = (ma_max - ma_min) / df.loc[valid, 'Close'] <= 0.005
    return df


def calculate_all(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = calculate_ma(df)
    df = calculate_rsi(df)
    df = calculate_macd(df)
    df = calculate_bollinger(df)
    df = calculate_ma_convergence(df)
    return df
