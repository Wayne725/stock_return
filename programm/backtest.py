"""回測引擎：模擬交易、扣除手續費、計算 KPI"""
import numpy as np
import pandas as pd

FEE_RATE    = 0.001425   # 雙向手續費
STOP_LOSS   = -0.02      # 停損 -2%
TAKE_PROFIT =  0.05      # 停利 +5%


def _get_sig(row) -> int:
    val = row.get('Signal', 0)
    try:
        return 0 if pd.isna(val) else int(val)
    except (TypeError, ValueError):
        return 0


def run_backtest(df: pd.DataFrame, initial_capital: float = 1_000_000.0):
    """
    執行回測並回傳 (trades_df, equity_df, kpi_dict, final_capital)

    買進策略：All in（全部資金買進）
    手續費計算：交易金額 × 0.001425（買進賣出皆收）
    """
    capital     = initial_capital
    shares      = 0.0
    entry_price = 0.0
    cost_basis  = 0.0   # 買進時的實際花費（含手續費）
    entry_time  = None
    trades      = []
    equity_hist = []

    for i in range(len(df)):
        row   = df.iloc[i]
        price = float(row['Close'])
        ts    = df.index[i]
        sig   = _get_sig(row)

        equity_hist.append({'time': ts, 'equity': capital + shares * price})

        # ── 持倉中：檢查出場條件 ────────────────────────────────
        if shares > 0:
            ret         = (price - entry_price) / entry_price
            exit_reason = None
            if ret <= STOP_LOSS:
                exit_reason = '停損'
            elif ret >= TAKE_PROFIT:
                exit_reason = '停利'
            elif sig == -1:
                exit_reason = '訊號賣出'

            if exit_reason:
                proceeds = shares * price * (1 - FEE_RATE)
                pnl      = proceeds - cost_basis
                trades.append({
                    'entry_time' : entry_time,
                    'exit_time'  : ts,
                    'entry_price': round(entry_price, 2),
                    'exit_price' : round(price, 2),
                    'pnl_pct'   : round(ret * 100, 2),
                    'pnl'       : round(pnl, 0),
                    'reason'    : exit_reason,
                })
                capital     = proceeds
                shares      = 0.0
                entry_price = 0.0
                cost_basis  = 0.0
                entry_time  = None

        # ── 空手：檢查買進條件 ──────────────────────────────────
        elif shares == 0 and sig == 1:
            # All in：capital / (price × (1 + fee)) = 持有股數
            shares      = capital / (price * (1 + FEE_RATE))
            cost_basis  = capital                          # 買進花費的總金額
            entry_price = price
            entry_time  = ts
            capital     = 0.0

    # ── 強制平倉（回測結束仍持倉） ──────────────────────────────
    if shares > 0:
        final_price = float(df.iloc[-1]['Close'])
        proceeds    = shares * final_price * (1 - FEE_RATE)
        ret         = (final_price - entry_price) / entry_price
        trades.append({
            'entry_time' : entry_time,
            'exit_time'  : df.index[-1],
            'entry_price': round(entry_price, 2),
            'exit_price' : round(final_price, 2),
            'pnl_pct'   : round(ret * 100, 2),
            'pnl'       : round(proceeds - cost_basis, 0),
            'reason'    : '強制平倉',
        })
        capital = proceeds

    cols = ['entry_time', 'exit_time', 'entry_price', 'exit_price',
            'pnl_pct', 'pnl', 'reason']
    trades_df = pd.DataFrame(trades, columns=cols) if trades else pd.DataFrame(columns=cols)
    equity_df = pd.DataFrame(equity_hist).set_index('time')
    kpi       = _calc_kpi(initial_capital, capital, trades_df, equity_df)

    return trades_df, equity_df, kpi, capital


def _calc_kpi(initial_capital, final_capital, trades_df, equity_df) -> dict:
    roi = (final_capital - initial_capital) / initial_capital * 100

    if len(trades_df) > 0:
        n_win    = int((trades_df['pnl'] > 0).sum())
        n_total  = len(trades_df)
        win_rate = n_win / n_total * 100
        risk     = float(trades_df['pnl_pct'].std())
    else:
        n_win = n_total = 0
        win_rate = risk = 0.0

    eq_arr = equity_df['equity'].values.astype(float)
    peak   = np.maximum.accumulate(eq_arr)
    dd     = (eq_arr - peak) / peak * 100
    max_dd = float(dd.min()) if len(dd) > 0 else 0.0

    return {
        '原始投資金額'       : f'${initial_capital:>13,.0f}',
        '最後結算金額'       : f'${final_capital:>13,.2f}',
        '投資報酬率'        : f'{roi:+.2f}%',
        '風險(報酬率標準差)' : f'{risk:.2f}%',
        '獲利比率'          : f'{win_rate:.1f}%',
        '最大回落'          : f'{max_dd:.2f}%',
        '總交易筆數'        : str(n_total),
        '獲利交易筆數'       : str(n_win),
    }
