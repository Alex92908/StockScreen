#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
stock_data.py — 纯数据分析，结果输出到 stdout
OpenClaw 读取输出后通过飞书推送
不含任何推送逻辑
"""

import os, json, sys
import warnings; warnings.filterwarnings('ignore')
from datetime import datetime, timedelta
import requests
from requests.adapters import HTTPAdapter

_orig_get = requests.Session.get

def _patched_get(self, url, **kwargs):
    kwargs.setdefault('headers', {})
    kwargs['headers'].setdefault('User-Agent',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 '
        '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    kwargs['headers'].setdefault('Referer', 'https://quote.eastmoney.com/')
    kwargs['headers'].setdefault('Accept', 'application/json, text/plain, */*')
    return _orig_get(self, url, **kwargs)

requests.Session.get = _patched_get

STATE_PATH = os.path.expanduser("~/stock/state.json")

def load_state():
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"x_codes": [], "x_names": {}, "date": ""}

def save_state(state):
    os.makedirs(os.path.dirname(STATE_PATH), exist_ok=True)
    with open(STATE_PATH, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def _ak():
    import akshare as ak
    return ak

def fetch_a_spot():
    return _ak().stock_zh_a_spot_em()

def fetch_hk_spot():
    try:
        return _ak().stock_hk_spot_em()
    except:
        return None

def fetch_hist(code, days=65):
    try:
        ak = _ak()
        start = (datetime.now() - timedelta(days=days)).strftime('%Y%m%d')
        end   = datetime.now().strftime('%Y%m%d')
        df = ak.stock_zh_a_hist(symbol=code, period="daily",
                                 start_date=start, end_date=end, adjust="qfq")
        return df if df is not None and not df.empty else None
    except:
        return None

def fetch_index(symbol, days=90):
    try:
        df = _ak().stock_zh_index_daily_em(symbol=symbol)
        col = {'close':'收盘','open':'开盘','high':'最高','low':'最低','volume':'成交量'}
        df = df.rename(columns={k:v for k,v in col.items() if k in df.columns})
        return df.tail(days)
    except:
        return None

def fetch_money_flow():
    try:
        return _ak().stock_individual_fund_flow_rank(indicator="今日")
    except:
        return None

def fetch_sector_flow():
    try:
        return _ak().stock_sector_fund_flow_rank()
    except:
        return None

def fetch_news():
    try:
        df = _ak().stock_news_em()
        return df.head(10) if df is not None and not df.empty else None
    except:
        return None

# 技术指标
def ma(s, n): return s.rolling(n).mean()

def rsi(hist, n=14):
    try:
        c = hist['收盘'].astype(float)
        d = c.diff()
        g = d.where(d>0,0).rolling(n).mean()
        l = (-d.where(d<0,0)).rolling(n).mean()
        return round((100-100/(1+g/l)).iloc[-1], 1)
    except: return None

def is_ma_bull(hist):
    try:
        c = hist['收盘'].astype(float)
        return ma(c,5).iloc[-1] > ma(c,10).iloc[-1] > ma(c,20).iloc[-1]
    except: return False

def is_macd_golden(hist):
    try:
        c = hist['收盘'].astype(float)
        e1 = c.ewm(span=12,adjust=False).mean()
        e2 = c.ewm(span=26,adjust=False).mean()
        m = e1-e2; s = m.ewm(span=9,adjust=False).mean()
        return m.iloc[-2]<s.iloc[-2] and m.iloc[-1]>s.iloc[-1]
    except: return False

def is_kdj_golden(hist):
    try:
        c  = hist['收盘'].astype(float)
        lo = hist['最低'].astype(float).rolling(9,min_periods=9).min()
        hi = hist['最高'].astype(float).rolling(9,min_periods=9).max()
        rsv= (c-lo)/(hi-lo)*100
        k  = rsv.ewm(com=2).mean(); d = k.ewm(com=2).mean()
        return k.iloc[-2]<=d.iloc[-2] and k.iloc[-1]>d.iloc[-1]
    except: return False

def consecutive_limit_up(hist):
    n = 0
    for i in range(len(hist)-1,-1,-1):
        try:
            if float(hist.iloc[i]['涨跌幅'])>=9.5: n+=1
            else: break
        except: break
    return n

def parse_flow(v):
    try:
        if isinstance(v,(int,float)): return float(v)
        v = str(v).strip().replace(',','')
        neg = v.startswith('-'); v = v.lstrip('-')
        if '亿' in v: r = float(v.replace('亿',''))*1e8
        elif '万' in v: r = float(v.replace('万',''))*1e4
        else: r = float(v)
        return -r if neg else r
    except: return 0.0

def _c(v): return f"+{v:.2f}%" if v>=0 else f"{v:.2f}%"

# ──────────────────────────────────────────────────────────
# 各分析模块输出纯文本
# ──────────────────────────────────────────────────────────

def report_market():
    lines = ["【大盘趋势】"]
    for symbol, label in [("sh000001","上证"),("sz399001","深证"),("sz399006","创业板")]:
        df = fetch_index(symbol)
        if df is None or len(df)<20: continue
        try:
            c = df['收盘'].astype(float)
            m5,m10,m20,m60 = ma(c,5).iloc[-1],ma(c,10).iloc[-1],ma(c,20).iloc[-1],ma(c,60).iloc[-1]
            e1=c.ewm(span=12,adjust=False).mean(); e2=c.ewm(span=26,adjust=False).mean()
            mc=e1-e2; sg=mc.ewm(span=9,adjust=False).mean()
            d=c.diff(); g=d.where(d>0,0).rolling(14).mean(); l=(-d.where(d<0,0)).rolling(14).mean()
            r=round((100-100/(1+g/l)).iloc[-1],1)
            bull=m5>m10>m20>m60; bear=m5<m10<m20<m60
            mgold=mc.iloc[-1]>sg.iloc[-1] and mc.iloc[-2]<=sg.iloc[-2]
            if bull and mgold:     trend="强势上涨🚀"
            elif bull:             trend="震荡上涨📈"
            elif bear and mgold:   trend="强势下跌🔻"
            elif bear:             trend="震荡下跌📉"
            else:                  trend="震荡整理↔️"
            chg=round((c.iloc[-1]-c.iloc[-2])/c.iloc[-2]*100,2)
            bull_str="多头🟢" if bull else "空头🔴"
            lines.append(f"{label} {c.iloc[-1]} ({_c(chg)}) | {trend} | {bull_str} | RSI {r}")
        except: pass

    # 强势板块
    try:
        sf = fetch_sector_flow()
        if sf is not None:
            strong = []
            for _,row in sf.iterrows():
                chg=0
                for col in ['今日涨跌幅','涨跌幅','涨幅']:
                    if col in row.index:
                        try: chg=float(row[col]); break
                        except: pass
                flow=0
                for col in row.index:
                    if '净流入' in col or '净额' in col:
                        flow=parse_flow(row[col]); break
                if chg>=2 and flow>0:
                    strong.append(f"{row.get('名称','')} {_c(chg)} 净入{flow/1e8:.1f}亿")
            if strong:
                lines.append("\n强势板块：" + " | ".join(strong[:5]))
    except: pass

    return "\n".join(lines)


def report_limit_up(df_spot, top_n=10):
    lines = ["【涨停票分析】"]
    limit = df_spot[
        (df_spot['涨跌幅']>=9.5) &
        (~df_spot['代码'].str.startswith('688')) &
        (~df_spot['代码'].str.startswith('430'))
    ].copy()
    lines.append(f"今日涨停共 {len(limit)} 只，重点如下：")

    results = []
    for _,row in limit.iterrows():
        code = str(row['代码'])
        try:
            hist = fetch_hist(code, days=30)
            cons = consecutive_limit_up(hist) if hist is not None else 1
            vr   = float(row.get('量比',0))
            to   = float(row.get('换手率',0))
            if cons>=3:          risk="高风险⚠️"
            elif cons==2 or vr>5:risk="中风险⚡"
            else:                risk="低风险✅"
            tags=[]
            if cons>1: tags.append(f"连续{cons}板🔥")
            else: tags.append("首板")
            if vr>=3: tags.append(f"量比{vr:.1f}")
            if hist is not None and is_ma_bull(hist): tags.append("均线多头")
            results.append((cons, vr, f"{row.get('名称','')}({code}) {row.get('最新价',0)}元 {risk} {'·'.join(tags)}"))
        except: continue

    results.sort(key=lambda x:(x[0],x[1]), reverse=True)
    for _,_,desc in results[:top_n]:
        lines.append(f"· {desc}")
    return "\n".join(lines)


def report_money_flow(top_n=10):
    lines = ["【主力资金净流入 Top10】"]
    df = fetch_money_flow()
    if df is None or df.empty:
        return "主力资金数据获取失败"
    results = []
    for _,row in df.iterrows():
        try:
            code=str(row.get('代码','')); name=str(row.get('名称',''))
            chg=0
            for col in ['今日涨跌幅','涨跌幅']:
                if col in row.index:
                    try: chg=float(row[col]); break
                    except: pass
            flow=0
            for col in row.index:
                if '主力净流入' in col:
                    flow=parse_flow(row[col]); break
            if flow>0:
                results.append((flow, f"{name}({code}) {_c(chg)} 净入{flow/1e8:.2f}亿"))
        except: continue
    results.sort(reverse=True)
    for i,(_,desc) in enumerate(results[:top_n],1):
        lines.append(f"{i}. {desc}")
    return "\n".join(lines)


def report_x_set(df_spot, cfg_path=None):
    lines = ["【选股X集合】（主板均线多头+量能）"]

    # 默认参数
    f = {
        "change_min":2.0,"change_max":9.4,
        "turnover_min":2.0,"turnover_max":20.0,
        "vol_ratio_min":1.5,"cap_min_yi":30.0,"cap_max_yi":800.0,
        "ma_bullish":True,"top_n":8
    }
    if cfg_path and os.path.exists(cfg_path):
        with open(cfg_path) as f2:
            cfg = json.load(f2)
            f.update(cfg.get('x_filter', {}))

    df = df_spot.copy()
    df = df[
        (df['涨跌幅']>=f['change_min']) & (df['涨跌幅']<=f['change_max']) &
        (df['换手率']>=f['turnover_min']) & (df['换手率']<=f['turnover_max']) &
        (df['量比']>=f['vol_ratio_min'])
    ]
    if '总市值' in df.columns:
        df = df[(df['总市值']>=f['cap_min_yi']*1e8)&(df['总市值']<=f['cap_max_yi']*1e8)]
    df = df[df['代码'].str.startswith('6')|df['代码'].str.startswith('0')]
    df = df[~df['代码'].str.startswith('688')]

    results = []
    for _,row in df.iterrows():
        code = str(row['代码'])
        try:
            hist = fetch_hist(code, days=65)
            if hist is None or len(hist)<20: continue
            if not is_ma_bull(hist): continue
            score=3; tags=["均线多头✅"]
            if is_macd_golden(hist): score+=2; tags.append("MACD金叉📈")
            if is_kdj_golden(hist):  score+=1; tags.append("KDJ金叉")
            vr=float(row.get('量比',0))
            if vr>=3: score+=2; tags.append(f"量比{vr:.1f}🔥")
            elif vr>=2: score+=1; tags.append(f"量比{vr:.1f}")
            r=rsi(hist)
            cap=float(row.get('总市值',0))/1e8
            results.append((score,
                f"{row.get('名称','')}({code}) "
                f"💰{row.get('最新价',0)} {_c(float(row.get('涨跌幅',0)))} "
                f"换手{float(row.get('换手率',0)):.1f}% 市值{cap:.0f}亿 RSI {r}\n"
                f"   信号: {'·'.join(tags)}"
            ))
            if len(results)>=f['top_n']*4: break
        except: continue

    results.sort(reverse=True)
    final = results[:f['top_n']]

    # 保存X集合代码
    state = load_state()
    state['x_codes'] = []
    state['x_names'] = {}
    for _,df_row in df.head(f['top_n']).iterrows():
        code = str(df_row['代码'])
        state['x_codes'].append(code)
        state['x_names'][code] = str(df_row.get('名称',''))
    state['date'] = datetime.now().strftime('%Y%m%d')
    save_state(state)

    if not final:
        lines.append("今日暂无满足条件的股票")
    else:
        for _,desc in final:
            lines.append(f"· {desc}")
    return "\n".join(lines)


def report_watchlist(df_spot, codes_names):
    lines = ["【重点关注】"]
    spot = {str(r['代码']): r for _,r in df_spot.iterrows()}
    for code, name in codes_names:
        row = spot.get(str(code))
        if row is None:
            lines.append(f"{name}({code}) 数据获取失败")
            continue
        price=float(row.get('最新价',0))
        change=float(row.get('涨跌幅',0))
        vr=float(row.get('量比',0))
        to=float(row.get('换手率',0))
        hist=fetch_hist(code,days=65)
        tags=[]
        r_val=None
        if hist is not None and len(hist)>=20:
            r_val=rsi(hist)
            if is_ma_bull(hist):    tags.append("均线多头✅")
            if is_macd_golden(hist):tags.append("MACD金叉📈")
        if change>=9.5:  tags.append("🚀涨停！")
        elif change>=5:  tags.append("强势↑")
        elif change<=-5: tags.append("⚠️大跌")
        tag_str = " ".join(tags) if tags else "正常"
        lines.append(
            f"{name}({code}) 💰{price} {_c(change)} "
            f"量比{vr:.1f} 换手{to:.1f}% RSI {r_val} | {tag_str}"
        )
    return "\n".join(lines)


def report_hk(threshold=20.0):
    df = fetch_hk_spot()
    if df is None or df.empty:
        return f"港股数据获取失败"
    results = []
    for _,row in df.iterrows():
        try:
            chg=float(row.get('涨跌幅',0))
            if chg>=threshold:
                results.append(f"{row.get('名称','')}({row.get('代码','')}) {row.get('最新价',0)} {_c(chg)}")
        except: continue
    if not results:
        return f"港股：暂无涨幅>{threshold:.0f}%个股"
    return f"【港股涨幅>{threshold:.0f}%预警】\n" + "\n".join(results)


def report_news():
    df = fetch_news()
    if df is None or df.empty: return ""
    lines = ["【今日热点新闻】"]
    for i,(_,row) in enumerate(df.iterrows(),1):
        t = str(row.get('新闻标题', row.get('标题','')))
        if t and len(t)>5: lines.append(f"{i}. {t[:70]}")
        if i>=8: break
    return "\n".join(lines)


def report_x_intraday(df_spot):
    """盘中X集合实时状态"""
    state = load_state()
    if state.get('date') != datetime.now().strftime('%Y%m%d'):
        return ""
    codes = state.get('x_codes',[])
    names = state.get('x_names',{})
    if not codes: return ""
    spot = {str(r['代码']): r for _,r in df_spot.iterrows()}
    lines = ["【X集合实时跟踪】"]
    for code in codes:
        row = spot.get(str(code))
        if row:
            chg=float(row.get('涨跌幅',0))
            lines.append(f"{names.get(code,'')}({code}) 💰{row.get('最新价',0)} {_c(chg)}")
    return "\n".join(lines)


# ──────────────────────────────────────────────────────────
# 三个模式的输出
# ──────────────────────────────────────────────────────────
WATCHLIST = [("002195","岩山科技"), ("600879","航天电子")]
CFG_PATH  = os.path.expanduser("~/stock/config.json")
HK_PCT    = 20.0

def run_morning():
    now = datetime.now().strftime('%m/%d %H:%M')
    print(f"=== 股票晨报 {now} ===\n")
    df = fetch_a_spot()
    # 全市场概况
    up=len(df[df['涨跌幅']>0]); dn=len(df[df['涨跌幅']<0]); lu=len(df[df['涨跌幅']>=9.5])
    print(f"全市场：涨{up}家 跌{dn}家 涨停{lu}只\n")
    print(report_market()); print()
    print(report_news()); print()
    print(report_limit_up(df)); print()
    print(report_money_flow()); print()
    print(report_x_set(df, CFG_PATH)); print()
    print(report_watchlist(df, WATCHLIST)); print()
    print(report_hk(HK_PCT))
    print("\n⚠️ 以上为自动分析仅供参考，股市有风险，投资需谨慎")

def run_intraday():
    now = datetime.now().strftime('%H:%M')
    print(f"=== 盘中监控 {now} ===\n")
    df = fetch_a_spot()
    print(report_watchlist(df, WATCHLIST)); print()
    print(report_x_intraday(df)); print()
    print(report_hk(HK_PCT))

def run_close():
    now = datetime.now().strftime('%m/%d %H:%M')
    print(f"=== 收盘日报 {now} ===\n")
    df = fetch_a_spot()
    up=len(df[df['涨跌幅']>0]); dn=len(df[df['涨跌幅']<0])
    fl=len(df[df['涨跌幅']==0]); lu=len(df[df['涨跌幅']>=9.5]); ld=len(df[df['涨跌幅']<=-9.5])
    print(f"全市场收盘：涨{up}家 跌{dn}家 平{fl}家 涨停{lu}只 跌停{ld}只\n")
    print(report_market()); print()
    print(report_limit_up(df)); print()
    print(report_money_flow(top_n=8)); print()
    print(report_watchlist(df, WATCHLIST)); print()
    # X集合收盘表现
    state = load_state()
    if state.get('date')==datetime.now().strftime('%Y%m%d'):
        spot={str(r['代码']):r for _,r in df.iterrows()}
        codes=state.get('x_codes',[]); names=state.get('x_names',{})
        if codes:
            print("【X集合今日表现】")
            for code in codes:
                row=spot.get(str(code))
                if row:
                    chg=float(row.get('涨跌幅',0))
                    icon="📈" if chg>0 else ("📉" if chg<0 else "➡️")
                    print(f"{icon} {names.get(code,'')}({code}) 收盘{row.get('最新价',0)} {_c(chg)}")
            print()
    print(report_hk(HK_PCT))
    print("\n⚠️ 以上为自动分析仅供参考，股市有风险，投资需谨慎")


if __name__ == "__main__":
    run_morning()
    # mode = sys.argv[1] if len(sys.argv)>1 else "morning"
    # if mode == "morning":  run_morning()
    # elif mode == "intraday": run_intraday()
    # elif mode == "close":  run_close()
