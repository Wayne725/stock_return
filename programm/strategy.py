"""交易策略產生訊號模組"""
import pandas as pd


def _safe_signal(val) -> int:
    try:
        return 0 if pd.isna(val) else int(val)
    except (TypeError, ValueError):
        return 0


def strategy_ma_convergence(df: pd.DataFrame) -> pd.DataFrame:
    """
    策略一：均線糾結突破
    買進：前一根均線糾結 + 當根 Close > MA5 > MA10 > MA20
    賣出：Close < MA20（其餘停損/停利由回測引擎處理）
    """
    signals = pd.Series(0, index=df.index, dtype=int, name='Signal')

    for i in range(1, len(df)):
        prev = df.iloc[i - 1]
        curr = df.iloc[i]

        if pd.isna([curr['MA5'], curr['MA10'], curr['MA20']]).any():
            continue

        if (prev.get('MA_Conv', False) and
                curr['Close'] > curr['MA5'] > curr['MA10'] > curr['MA20']):
            signals.iloc[i] = 1
        elif curr['Close'] < curr['MA20']:
            signals.iloc[i] = -1

    df = df.copy()
    df['Signal'] = signals
    return df


def strategy_rsi(df: pd.DataFrame, buy_level: int = 30, sell_level: int = 70) -> pd.DataFrame:
    """
    策略二：RSI 超買超賣
    買進：RSI < 30（超賣）
    賣出：RSI > 70（超買）
    """
    signals = pd.Series(0, index=df.index, dtype=int, name='Signal')

    for i in range(len(df)):
        rsi = df.iloc[i].get('RSI', float('nan'))
        if pd.isna(rsi):
            continue
        if rsi < buy_level:
            signals.iloc[i] = 1
        elif rsi > sell_level:
            signals.iloc[i] = -1

    df = df.copy()
    df['Signal'] = signals
    return df


def strategy_macd(df: pd.DataFrame) -> pd.DataFrame:
    """
    策略三：MACD 黃金 / 死亡交叉
    買進：DIF 由下往上穿越 MACD_SIG（黃金交叉）
    賣出：DIF 由上往下穿越 MACD_SIG（死亡交叉）
    """
    signals = pd.Series(0, index=df.index, dtype=int, name='Signal')

    for i in range(1, len(df)):
        prev = df.iloc[i - 1]
        curr = df.iloc[i]

        if pd.isna([curr['DIF'], curr['MACD_SIG'],
                    prev['DIF'], prev['MACD_SIG']]).any():
            continue

        if prev['DIF'] < prev['MACD_SIG'] and curr['DIF'] > curr['MACD_SIG']:
            signals.iloc[i] = 1                    # 黃金交叉
        elif prev['DIF'] > prev['MACD_SIG'] and curr['DIF'] < curr['MACD_SIG']:
            signals.iloc[i] = -1                   # 死亡交叉

    df = df.copy()
    df['Signal'] = signals
    return df


def strategy_bollinger(df: pd.DataFrame) -> pd.DataFrame:
    """
    策略四：布林通道突破
    買進：Close 跌破下軌（超賣）
    賣出：Close 突破上軌（超買）
    """
    signals = pd.Series(0, index=df.index, dtype=int, name='Signal')

    for i in range(len(df)):
        curr = df.iloc[i]

        if pd.isna([curr.get('BB_UPPER', float('nan')),
                    curr.get('BB_LOWER', float('nan'))]).any():
            continue

        if curr['Close'] < curr['BB_LOWER']:
            signals.iloc[i] = 1
        elif curr['Close'] > curr['BB_UPPER']:
            signals.iloc[i] = -1

    df = df.copy()
    df['Signal'] = signals
    return df


def apply_strategy(df: pd.DataFrame, strategy_name: str) -> pd.DataFrame:
    if strategy_name == '均線糾結策略':
        return strategy_ma_convergence(df)
    elif strategy_name == 'RSI策略':
        return strategy_rsi(df)
    elif strategy_name == 'MACD策略':
        return strategy_macd(df)
    elif strategy_name == '布林通道策略':
        return strategy_bollinger(df)
    return df
