"""
FINREPORT INTELLIGENCE - Bước 5 (Phát hiện tín hiệu) + Bước 6 (Risk score / mức ưu tiên)
Kiến trúc: Rule-engine + Threshold Tuning động.
  - Ngưỡng KHÔNG cố định: mỗi chỉ tiêu được đối chiếu (cross-join) với p25/p75/IQR/fence của ĐỒNG NGÀNH CÙNG NĂM (file Bước 4)
    và với mức thay đổi so với năm trước, chuẩn hóa theo IQR của ngành.
  - Các tham số điều chỉnh (tau, k0, t0, wp, ngưỡng severity) được tinh chỉnh bằng lưới + leave-one-company-out trên nhãn mẫu BTC.
Điểm sáng Silver áp dụng ở bước này: phát hiện đồng thời nhiều tín hiệu; baseline để so sánh; KPI thực nghiệm (leave-one-company-out); health score.
CHƯA thuộc bước này (làm ở Bước 7-9, nháp trong step7_9_draft.py): giải thích lý do, truy vết nguồn, verify, confidence, what-if.
INPUT : master1_1_with_metrics.xlsx (Bước 1-2), B4_benchmark_nganh_theonam.xlsx (đầu ra Bước 4), financial_scenarios_sample.csv (nhãn mẫu, để tinh chỉnh/đo KPI; tùy chọn)
OUTPUT (out/): B5_tin_hieu.xlsx, B6_risk_score.xlsx, B5B6_KPI.xlsx, B5B6_doi_chieu_nhan_mau.xlsx
Chạy: python3 step5_6_signal_risk.py
"""
import os, time, itertools, json
import numpy as np, pandas as pd

MASTER = "master1_1_with_metrics.xlsx"; BENCH = "B4_benchmark_nganh_theonam.xlsx"; LABELS = "financial_scenarios_sample.csv"
OUT = "out"; MODEL = "FinReport-DynamicRule-v1"
K = ["company_id", "fiscal_year"]
FAMILY = {  # 6 họ tín hiệu (tag family) và trọng số trong risk score
 "MARGIN": ({"GROSS_MARGIN_DECLINE", "PROFITABILITY_PRESSURE", "LOW_MARGIN"}, .60),
 "RECEIVABLES": ({"RECEIVABLES_RISK"}, .45),
 "INVENTORY": ({"INVENTORY_RISK"}, .45),
 "CASH_QUALITY": ({"CFO_WEAKNESS", "CFO_PRESSURE", "PROFIT_CASH_DIVERGENCE"}, .50),
 "LEVERAGE": ({"LEVERAGE_RISK", "HIGH_LEVERAGE", "INTEREST_BURDEN"}, .60),
 "LIQUIDITY_WC": ({"LIQUIDITY_STRESS", "LIQUIDITY_PRESSURE", "WORKING_CAPITAL_PRESSURE"}, .60)}
TAG2FAM = {t: f for f, (ts, _) in FAMILY.items() for t in ts}
DEFAULT = dict(tau=.35, k0=0.0, t0=0.5, wp=.5, cuts=(.10, .35, .75), tau_press=.7)
SRC = {}  # tên trường -> file nguồn Data Pack
for f, cols in {
 "income_statement.csv": ["revenue_vnd_mn", "cost_of_goods_sold_vnd_mn", "operating_profit_vnd_mn", "net_income_vnd_mn", "interest_expense_vnd_mn"],
 "balance_sheet.csv": ["accounts_receivable_vnd_mn", "inventory_vnd_mn", "total_current_assets_vnd_mn", "total_current_liabilities_vnd_mn",
                       "short_term_debt_vnd_mn", "long_term_debt_vnd_mn", "total_assets_vnd_mn", "total_equity_vnd_mn"],
 "cashflow.csv": ["net_cash_from_operating_activities_vnd_mn", "capital_expenditure_vnd_mn"]}.items():
    for c in cols: SRC[c] = f
# chỉ tiêu -> các trường gốc dùng để tính (phục vụ evidence_fields / truy vết)
EVID = {"gross_margin_pct": ["revenue_vnd_mn", "cost_of_goods_sold_vnd_mn"], "operating_margin_pct": ["revenue_vnd_mn", "operating_profit_vnd_mn"],
 "net_margin_pct": ["revenue_vnd_mn", "net_income_vnd_mn"], "dso_days": ["accounts_receivable_vnd_mn", "revenue_vnd_mn"],
 "inventory_days": ["inventory_vnd_mn", "cost_of_goods_sold_vnd_mn"], "cfo_to_net_income": ["net_cash_from_operating_activities_vnd_mn", "net_income_vnd_mn"],
 "debt_to_assets_pct": ["short_term_debt_vnd_mn", "long_term_debt_vnd_mn", "total_assets_vnd_mn"], "debt_to_equity": ["short_term_debt_vnd_mn", "long_term_debt_vnd_mn", "total_equity_vnd_mn"],
 "interest_coverage": ["operating_profit_vnd_mn", "interest_expense_vnd_mn"], "free_cash_flow_vnd_mn": ["net_cash_from_operating_activities_vnd_mn", "capital_expenditure_vnd_mn"],
 "current_ratio": ["total_current_assets_vnd_mn", "total_current_liabilities_vnd_mn"], "cfo_margin_pct": ["net_cash_from_operating_activities_vnd_mn", "revenue_vnd_mn"],
 "revenue_growth_pct": ["revenue_vnd_mn"]}
BMETRICS = ["gross_margin_pct", "operating_margin_pct", "net_margin_pct", "dso_days", "inventory_days", "cfo_to_net_income", "debt_to_assets_pct",
            "debt_to_equity", "interest_coverage", "current_ratio", "free_cash_flow_vnd_mn", "revenue_growth_pct", "cfo_margin_pct"]

# ======================================================================= nạp & chuẩn bị
def load(master=MASTER, bench=BENCH):
    raw = pd.read_excel(master, header=None)
    h = next(i for i in range(min(15, len(raw))) if {"company_id", "fiscal_year"} <= set(raw.iloc[i].astype(str)))
    df = pd.read_excel(master, header=h).dropna(subset=K).copy(); df.fiscal_year = df.fiscal_year.astype(int)
    df = df.sort_values(K).reset_index(drop=True)
    df["cfo_vnd_mn"] = df.net_cash_from_operating_activities_vnd_mn
    df["cfo_margin_pct"] = df.cfo_vnd_mn / df.revenue_vnd_mn * 100
    df["total_debt_vnd_mn"] = df.short_term_debt_vnd_mn + df.long_term_debt_vnd_mn
    g = df.groupby("company_id")
    df["ar_growth_pct"] = g.accounts_receivable_vnd_mn.pct_change() * 100
    df["inv_growth_pct"] = g.inventory_vnd_mn.pct_change() * 100
    df["debt_growth_pct"] = g.total_debt_vnd_mn.pct_change() * 100
    df["ar_gap"] = df.ar_growth_pct - df.revenue_growth_pct          # phải thu tăng nhanh hơn doanh thu (điểm %)
    df["inv_gap"] = df.inv_growth_pct - df.revenue_growth_pct
    for m in BMETRICS:
        df[f"{m}__chg1"] = g[m].diff(); df[f"{m}__chg2"] = g[m].diff(2)
    cm = df.cfo_to_net_income.where(df.net_income_vnd_mn > 0)       # CFO/LNR chỉ có nghĩa khi lợi nhuận ròng > 0
    df["cfo_to_net_income__chg1"] = cm.groupby(df.company_id).diff(); df["cfo_to_net_income__chg2"] = cm.groupby(df.company_id).diff(2)
    # benchmark: nạp file Bước 4 và ghép (cross-join) theo năm tài chính
    b = pd.read_excel(bench)
    b = b[b.n_doanh_nghiep >= 5]
    ext = []                                                        # benchmark cho chỉ tiêu dẫn xuất (cfo_margin) cùng phương pháp Bước 4
    for y, gg in df.groupby("fiscal_year"):
        x = gg.cfo_margin_pct.dropna(); q1, q3 = x.quantile(.25), x.quantile(.75); i = q3 - q1
        ext.append(dict(metric="cfo_margin_pct", fiscal_year=y, n_doanh_nghiep=len(x), median=x.median(), p25=q1, p75=q3, iqr=i, lower_fence=q1 - 1.5 * i, upper_fence=q3 + 1.5 * i))
    b = pd.concat([b, pd.DataFrame(ext)], ignore_index=True)
    piv = b.pivot(index="fiscal_year", columns="metric", values=["median", "p25", "p75", "iqr", "lower_fence", "upper_fence"])
    piv.columns = [f"{m}__{s}" for s, m in piv.columns]
    df = df.merge(piv.reset_index(), on="fiscal_year", how="left")
    return df, b

# ======================================================================= lõi: điểm tín hiệu động
def ramp(x, a, b): return np.clip((x - a) / (b - a), 0, 1)
def peer_part(df, m, d, P):
    """d=+1: càng cao càng xấu (so với p75); d=-1: càng thấp càng xấu (so với p25). Đơn vị: số IQR vượt quá quartile; 1,5 IQR = fence."""
    ref = df[f"{m}__p75"] if d > 0 else df[f"{m}__p25"]
    z = d * (df[m] - ref) / df[f"{m}__iqr"].replace(0, np.nan)
    return ramp(z, P["k0"], 1.5)
def trend_part(df, m, d, P):
    ch = pd.concat([df[f"{m}__chg1"], df[f"{m}__chg2"] / 2], axis=1).apply(lambda r: r.max() if d < 0 else r.min(), axis=1)  # lấy thay đổi theo hướng xấu nhất  (d>0: xấu = tăng)
    ch = pd.concat([df[f"{m}__chg1"], df[f"{m}__chg2"] / 2], axis=1)
    worst = (ch * d).max(axis=1, skipna=True)                       # theo hướng xấu
    z = worst / df[f"{m}__iqr"].replace(0, np.nan)
    return ramp(z, P["t0"], P["t0"] + 1.0), z
def S(df, m, d, P, wp=None):
    wp = P["wp"] if wp is None else wp
    pp = peer_part(df, m, d, P); tp, tz = trend_part(df, m, d, P)
    s = np.where(tp.notna(), wp * pp + (1 - wp) * tp, pp)           # năm đầu chưa có xu hướng: chỉ dựa vị trí đồng ngành
    if P.get("and_gate"):                                            # (thí nghiệm giảm false positive) cần cả vị trí lẫn xu hướng cùng xấu
        s = np.where(tp.notna() & ~((pp > 0) & (tp > 0)), 0, s)
    return pd.Series(s, index=df.index).fillna(0), pp.fillna(0), tp.fillna(0)

def detect(df, P, tracer=None):
    """tracer (tùy chọn): đối tượng ghi lại tự động các cột dữ liệu mà luật đã đọc cho TỪNG tag (dùng cho Auto Audit Log ở lớp Gold)."""
    G = (lambda c: c) if P.get("gates", True) else (lambda c: True)   # gates = điều kiện lọc giảm cảnh báo sai
    """Trả về (T, parts): T = điểm S (0..1) của từng tag; parts = (peer, trend) để giải thích."""
    T, parts = {}, {}
    if tracer: tracer.take_cols(); T = tracer.tdict()
    def put(tag, s, pp=None, tp=None, metric=None, d=None, cols=None):
        T[tag] = s.fillna(0).clip(0, 1); parts[tag] = (pp, tp, metric, d)
        if tracer: tracer.commit(tag, cols)
    # --- Phải thu / Tồn kho: vị trí DSO, Inventory Days so với ngành + thay đổi, kèm điều kiện tăng nhanh hơn doanh thu
    s, pp, tp = S(df, "dso_days", +1, P); put("RECEIVABLES_RISK", s * G(df.ar_gap > 0), pp, tp, "dso_days", +1)
    s, pp, tp = S(df, "inventory_days", +1, P); put("INVENTORY_RISK", s * G(df.inv_gap > 0), pp, tp, "inventory_days", +1)
    # --- Chất lượng dòng tiền
    ni_ok = df.net_income_vnd_mn > 0
    d2 = df.copy(); d2["cfo_to_net_income"] = df.cfo_to_net_income.where(ni_ok)
    s, pp, tp = S(d2, "cfo_to_net_income", -1, P)
    s = s.where(ni_ok & G(df.cfo_to_net_income < df["cfo_to_net_income__median"]), 0); s = np.where(df.cfo_vnd_mn < 0, 1.0, s); s = pd.Series(s, index=df.index)
    div = (df.net_margin_pct >= df["net_margin_pct__median"]) & (df.cfo_vnd_mn > 0) & (s >= P["tau"])
    grp = tracer.take_cols() if tracer else None                     # ba tag dòng tiền dùng chung các cột đã đọc ở trên
    put("PROFIT_CASH_DIVERGENCE", s.where(div, 0), pp, tp, "cfo_to_net_income", -1, cols=grp)
    put("CFO_PRESSURE", s.where(~div & (s >= P["tau_press"]), 0), pp, tp, "cfo_to_net_income", -1, cols=grp)
    put("CFO_WEAKNESS", s.where(~div & (s < P["tau_press"]), 0), pp, tp, "cfo_to_net_income", -1, cols=grp)
    # --- Đòn bẩy / lãi vay
    s1, pp1, tp1 = S(df, "debt_to_assets_pct", +1, P); s2, _, _ = S(df, "debt_to_equity", +1, P)
    put("LEVERAGE_RISK", np.maximum(s1, s2), pp1, tp1, "debt_to_assets_pct", +1)
    s, pp, tp = S(df, "interest_coverage", -1, P); put("INTEREST_BURDEN", s, pp, tp, "interest_coverage", -1)
    # --- Biên lợi nhuận: nghiêng về xu hướng (giảm tích lũy 2 năm) cho gross margin
    s, pp, tp = S(df, "gross_margin_pct", -1, P, wp=.25); put("GROSS_MARGIN_DECLINE", s * G(df.gross_margin_pct__chg2 < 0), pp, tp, "gross_margin_pct", -1)
    s, pp, tp = S(df, "operating_margin_pct", -1, P)
    dec = (df.operating_margin_pct__chg1.fillna(0) < 0) | (df.operating_margin_pct__chg2.fillna(0) < 0)
    put("PROFITABILITY_PRESSURE", s * G(dec), pp, tp, "operating_margin_pct", -1)
    # --- Thanh khoản / dòng tiền
    cm, _, _ = S(df, "cfo_margin_pct", -1, P)
    liq = np.where((df.cfo_vnd_mn < 0) & (df.free_cash_flow_vnd_mn < 0), 1.0, np.where(df.free_cash_flow_vnd_mn < 0, cm, 0))
    put("LIQUIDITY_STRESS", pd.Series(liq, index=df.index), None, None, "free_cash_flow_vnd_mn", -1)
    dz = pd.concat([df.dso_days__chg1 / df.dso_days__iqr, df.inventory_days__chg1 / df.inventory_days__iqr], axis=1).max(axis=1)
    wc = ramp(dz, P["t0"], P["t0"] + 1.0).fillna(0) * ((T["CFO_PRESSURE"] > 0) | (T["PROFIT_CASH_DIVERGENCE"] > 0) | (T["LIQUIDITY_STRESS"] > 0) | (T["CFO_WEAKNESS"] > 0))
    put("WORKING_CAPITAL_PRESSURE", wc, None, None, "dso_days", +1)
    # --- Nhóm "nền yếu" (distress baseline): vị trí thấp/cao so với ngành, không cần xu hướng
    s, pp, tp = S(df, "operating_margin_pct", -1, P, wp=1.0); put("LOW_MARGIN", s.where(T["PROFITABILITY_PRESSURE"] < P["tau"], 0), pp, tp, "operating_margin_pct", -1)
    s, pp, tp = S(df, "debt_to_assets_pct", +1, P, wp=1.0)
    put("HIGH_LEVERAGE", s.where((df.operating_margin_pct < df["operating_margin_pct__median"]) & (T["LEVERAGE_RISK"] < P["tau"]), 0), pp, tp, "debt_to_assets_pct", +1)
    s, pp, tp = S(df, "current_ratio", -1, P, wp=1.0); put("LIQUIDITY_PRESSURE", s.where(df.cfo_margin_pct < df["cfo_margin_pct__median"], 0), pp, tp, "current_ratio", -1)
    if P.get("drop_family"):                                        # (thí nghiệm feature importance) bỏ một họ tín hiệu
        for t in FAMILY[P["drop_family"]][0]: T[t] = T[t] * 0
    return dict(T), parts

def positives(df, P):
    z = lambda m, d, c="chg1": d * df[f"{m}__{c}"] / df[f"{m}__iqr"].replace(0, np.nan)
    good = .25
    mi = (z("gross_margin_pct", 1) >= good) & (z("operating_margin_pct", 1) >= good)
    pos = {"MARGIN_IMPROVEMENT": mi & (df.operating_margin_pct >= df.operating_margin_pct__median),
           "MARGIN_RECOVERY": mi & (df.operating_margin_pct < df.operating_margin_pct__median),
           "DELEVERAGING": (df.debt_growth_pct < 0) & (z("debt_to_assets_pct", -1) >= good),
           "WORKING_CAPITAL_IMPROVEMENT": (z("dso_days", -1) >= .3) & (z("inventory_days", -1) >= .3),
           "STRONG_GROWTH": (df.revenue_growth_pct > df.revenue_growth_pct__median) & (df.operating_margin_pct > df.operating_margin_pct__p75),
           "STRONG_LIQUIDITY": df.current_ratio > df.current_ratio__upper_fence,
           "LOW_LEVERAGE": df.debt_to_assets_pct < df.debt_to_assets_pct__p25}
    return {k: v.fillna(False) for k, v in pos.items()}

# ======================================================================= Bước 6: risk score, severity, health, scenario
SEV = ["LOW", "MEDIUM", "HIGH", "CRITICAL"]
def score_rows(df, T, P, drop=None):
    fam = {f: pd.concat([T[t] * (T[t] >= P["tau"]) for t in ts], axis=1).max(axis=1) for f, (ts, w) in FAMILY.items()}
    prod = np.ones(len(df))
    for f, (ts, w) in FAMILY.items():
        if f != drop: prod *= 1 - w * fam[f].values        # noisy-OR: nhiều tín hiệu cùng lúc cộng dồn, không vượt 1
    return pd.Series(1 - prod, index=df.index), fam

def scenario(r, tags, pos):
    T = set(tags); fam = {f: bool(T & ts) for f, (ts, _) in FAMILY.items()}; nf = sum(fam.values())
    if r.operating_margin_pct <= 0 and r.debt_to_assets_pct > r.debt_to_assets_pct__p75: return "DISTRESS_BASELINE"      # lỗ hoạt động + đòn bẩy cao hơn ngành
    if "MARGIN_RECOVERY" in pos and "DELEVERAGING" in pos: return "FINANCIAL_RECOVERY"
    if nf >= 4 and fam["MARGIN"] and (fam["RECEIVABLES"] or fam["INVENTORY"]): return "MULTI_FACTOR_STRESS"
    if "LIQUIDITY_STRESS" in T: return "LIQUIDITY_STRESS"
    if "STRONG_GROWTH" in pos and "LOW_LEVERAGE" in pos: return "STRONG_PERFORMANCE"
    if len(pos) >= 2 and not T: return "OPERATIONAL_IMPROVEMENT"
    if fam["MARGIN"] and r.gross_margin_pct__chg2 < 0: return "MARGIN_DETERIORATION"
    if fam["LEVERAGE"] and (r.debt_growth_pct > 25 or r.debt_to_assets_pct > r.debt_to_assets_pct__p75): return "DEBT_EXPANSION"
    if "RECEIVABLES_RISK" in T and r.revenue_growth_pct < r.revenue_growth_pct__median + 3: return "RECEIVABLES_RISK"
    if "INVENTORY_RISK" in T: return "INVENTORY_BUILDUP"
    if "PROFIT_CASH_DIVERGENCE" in T: return "PROFIT_CASH_DIVERGENCE"
    if fam["CASH_QUALITY"] and r.revenue_growth_pct >= r.revenue_growth_pct__median: return "REVENUE_GROWTH_WEAK_CASH"
    if "RECEIVABLES_RISK" in T: return "RECEIVABLES_RISK"
    if fam["LEVERAGE"]: return "DEBT_EXPANSION"
    if fam["CASH_QUALITY"]: return "REVENUE_GROWTH_WEAK_CASH"
    return "NORMAL_STABLE"

def health(r, sev, pos, prev, score):
    weak = (r.operating_margin_pct < r.operating_margin_pct__p25) or (r.interest_coverage < r.interest_coverage__p25)
    if sev in ("HIGH", "CRITICAL"): return "HIGH_RISK" if (sev == "CRITICAL" or weak) else "WATCH"
    if sev == "MEDIUM": return "WATCH"
    if prev is not None and prev - score >= .25 and len(pos) >= 2: return "RECOVERING"
    return "HEALTHY"

# ======================================================================= chạy một bộ tham số
def run(df, P, with_text=False):
    T, parts = detect(df, P); score, fam = score_rows(df, T, P, drop=P.get("drop")); pos = positives(df, P)
    cuts = P["cuts"]; sev = pd.Series(np.array(SEV)[np.searchsorted(cuts, score.values, side="right")], index=df.index)
    tags, scen, hl, prevs = [], [], [], {}
    for i, r in df.iterrows():
        fired = [t for t in T if T[t].iloc[i] >= P["tau"] and TAG2FAM[t] != P.get("drop")]
        ps = [k for k, v in pos.items() if v.iloc[i]]
        tags.append(fired + ([] if fired else ps)); scen.append(scenario(r, fired, ps))
        hl.append(health(r, sev.iloc[i], ps, prevs.get(r.company_id), score.iloc[i])); prevs[r.company_id] = score.iloc[i]
    O = df[K].copy(); O["risk_score"] = score.round(4); O["predicted_primary_scenario"] = scen; O["predicted_severity"] = sev
    O["predicted_overall_health"] = hl; O["alert_tags"] = [";".join(t) for t in tags]
    return O, T, parts, pos

# ======================================================================= đánh giá KPI
def fams(s): t = {x for x in str(s).split(";") if x and x != "nan"}; return {f for f, (ts, _) in FAMILY.items() if t & ts}
def evaluate(O, L):
    J = L.merge(O, on=K, suffixes=("_t", "_p")); r = {}
    r["severity_acc"] = (J.severity == J.predicted_severity).mean(); r["health_acc"] = (J.overall_health == J.predicted_overall_health).mean()
    r["scenario_acc"] = (J.primary_scenario == J.predicted_primary_scenario).mean()
    hc = J.severity.isin(["HIGH", "CRITICAL"]); ph = J.predicted_severity.isin(["HIGH", "CRITICAL"])
    r["highcritical_recall"] = (hc & ph).sum() / max(hc.sum(), 1); r["highcritical_precision"] = (hc & ph).sum() / max(ph.sum(), 1)
    tp = fp = fn = 0
    for _, x in J.iterrows():
        a, b = fams(x.alert_tags_t), fams(x.alert_tags_p); tp += len(a & b); fp += len(b - a); fn += len(a - b)
    P_, R_ = tp / max(tp + fp, 1), tp / max(tp + fn, 1); r["tagfam_precision"], r["tagfam_recall"] = P_, R_; r["tagfam_f1"] = 2 * P_ * R_ / max(P_ + R_, 1e-9)
    return r, J
def objective(r): return r["severity_acc"] + r["health_acc"] + r["tagfam_f1"] + r["scenario_acc"]

def baseline_old_rules(df):
    """Phương án cơ bản để so sánh: bộ luật ban đầu của nhóm (ngưỡng quartile/fence cố định theo Bước 4, điểm cộng cố định)."""
    sc = np.zeros(len(df)); tags = [[] for _ in range(len(df))]
    def add(mask, w, name):
        nonlocal sc
        for i in np.where(mask.fillna(False).values)[0]: tags[i].append(name)
        sc += w * mask.fillna(False).values
    add(df.net_margin_pct < df.net_margin_pct__p25, .25, "LOW_MARGIN_VS_PEERS"); add(df.inventory_days > df.inventory_days__upper_fence, .20, "CRITICAL_INVENTORY_BUILDUP")
    add(df.dso_days > df.dso_days__upper_fence, .15, "CRITICAL_DSO_DELAY"); add(df.debt_to_equity > df.debt_to_equity__p75, .20, "HIGH_LEVERAGE")
    prev = df.groupby("company_id").free_cash_flow_vnd_mn.shift(1)
    add((df.free_cash_flow_vnd_mn < 0) & (prev < 0), .35, "SUSTAINED_NEGATIVE_FCF"); add((df.free_cash_flow_vnd_mn < 0) & ~(prev < 0), .15, "NEGATIVE_FCF")
    sc = np.minimum(sc, 1)
    sev = np.where(sc == 0, "LOW", np.where(sc <= .3, "LOW", np.where(sc <= .6, "MEDIUM", "CRITICAL")))   # NONE/LOW -> LOW cho hợp lệ BTC
    hl = np.where(sc == 0, "HEALTHY", np.where(sc <= .3, "WATCH", np.where(sc <= .6, "WATCH", "HIGH_RISK")))  # WARNING không thuộc 4 lớp BTC -> WATCH
    O = df[K].copy(); O["risk_score"] = sc; O["predicted_severity"] = sev; O["predicted_overall_health"] = hl
    O["alert_tags"] = [";".join(t) for t in tags]; O["predicted_primary_scenario"] = "NA"
    return O

# ======================================================================= tinh chỉnh ngưỡng (grid + leave-one-company-out)
GRID = dict(tau=[.25, .35, .45], k0=[0, .5], t0=[0, .5, 1.0], wp=[.4, .6],
            cuts=[(.05, .25, .60), (.10, .35, .75), (.15, .45, .85)])
def tune(df, L):
    combos = [dict(zip(GRID, v), tau_press=.7) for v in itertools.product(*GRID.values())]
    cache = []
    for P in combos:
        O, *_ = run(df, P); r, J = evaluate(O, L); cache.append((P, r, J, objective(r)))
    best = max(cache, key=lambda x: x[3])
    # leave-one-company-out: chọn tham số tốt nhất trên các công ty còn lại, chấm trên công ty bị loại
    comps = sorted(L.company_id.unique()); rows = []
    for c in comps:
        def sc(x):
            J = x[2]; tr = J[J.company_id != c]
            return (tr.severity == tr.predicted_severity).mean() + (tr.overall_health == tr.predicted_overall_health).mean() + (tr.primary_scenario == tr.predicted_primary_scenario).mean()
        bp = max(cache, key=sc); J = bp[2]; te = J[J.company_id == c]
        rows.append(te.assign(_sev=te.severity == te.predicted_severity, _hl=te.overall_health == te.predicted_overall_health, _sc=te.primary_scenario == te.predicted_primary_scenario))
    cv = pd.concat(rows)
    return best, cv, cache

# ======================================================================= main
def main():
    os.makedirs(OUT, exist_ok=True); t0 = time.time()
    for f in (MASTER, BENCH):
        if not os.path.exists(f): raise SystemExit(f"Không tìm thấy {f}. Đặt cùng thư mục với script.")
    df, bench = load()
    L = pd.read_csv(LABELS) if os.path.exists(LABELS) else None
    P = dict(DEFAULT)
    if L is not None:
        best, cv, cache = tune(df, L); P = best[0]
        print("Tham số sau tinh chỉnh:", {k: P[k] for k in GRID})
    O, T, parts, pos = run(df, P)

    # ---- B5: danh sách tín hiệu phát hiện được (chỉ dữ liệu phát hiện, chưa có giải thích bằng lời, chưa truy vết nguồn)
    rows = []
    for i, r in df.iterrows():
        for t in T:
            s_ = T[t].iloc[i]
            if s_ >= P["tau"]:
                m = parts[t][2]
                rows.append(dict(company_id=r.company_id, fiscal_year=int(r.fiscal_year), tag=t, family=TAG2FAM[t], metric=m,
                                 value=r[m] if m in r.index else np.nan, peer_p25=r.get(f"{m}__p25"), peer_median=r.get(f"{m}__median"),
                                 peer_p75=r.get(f"{m}__p75"), upper_fence=r.get(f"{m}__upper_fence"), lower_fence=r.get(f"{m}__lower_fence"),
                                 signal_score=s_, threshold_tau=P["tau"]))
    A = pd.DataFrame(rows).round(3); A.to_excel(f"{OUT}/B5_tin_hieu.xlsx", index=False)

    # ---- B6: risk score, severity, health, scenario, mức ưu tiên
    neg = pd.DataFrame({t: T[t] >= P["tau"] for t in T})                   # tín hiệu rủi ro đã kích hoạt
    O["n_risk_signals"] = neg.sum(axis=1).values
    O["n_signal_families"] = [len({TAG2FAM[t] for t in T if neg[t].iloc[i]}) for i in range(len(O))]
    O["health_score"] = (100 * (1 - O.risk_score)).round(1)                 # điểm sức khỏe 0-100 (100 = khỏe nhất)
    O["priority_rank"] = O.sort_values(["fiscal_year", "risk_score"], ascending=[True, False]).groupby("fiscal_year").cumcount().add(1).reindex(O.index)
    cols = K + ["risk_score", "health_score", "n_risk_signals", "n_signal_families", "predicted_severity", "predicted_overall_health", "predicted_primary_scenario", "alert_tags", "priority_rank"]
    O[cols].round(4).to_excel(f"{OUT}/B6_risk_score.xlsx", index=False)

    # ---- KPI so với baseline
    if L is not None:
        rb, _ = evaluate(baseline_old_rules(df), L); rn, J = evaluate(O, L)
        cvr = dict(severity_acc=cv._sev.mean(), health_acc=cv._hl.mean(), scenario_acc=cv._sc.mean())
        K_ = pd.DataFrame([dict(phuong_an="Baseline (luật cũ, điểm cộng cố định)", **{k: rb[k] for k in ["severity_acc", "health_acc", "highcritical_recall"]}),
                           dict(phuong_an="Ngưỡng động, tinh chỉnh trên toàn bộ nhãn (tối ưu trong mẫu)", **rn),
                           dict(phuong_an="Ngưỡng động, leave-one-company-out (ước lượng trung thực)", **cvr)])
        K_.round(3).to_excel(f"{OUT}/B5B6_KPI.xlsx", index=False); J.to_excel(f"{OUT}/B5B6_doi_chieu_nhan_mau.xlsx", index=False)
        print("\nKPI trên 18 nhãn mẫu BTC:"); print(K_.round(2).to_string(index=False))
    last = int(df.fiscal_year.max())
    print(f"\nB5: {len(A)} tín hiệu | B6: {len(O)} dòng | {time.time()-t0:.1f}s\n")
    print(O[O.fiscal_year == last].sort_values("priority_rank")[["company_id", "risk_score", "predicted_severity", "predicted_overall_health", "predicted_primary_scenario", "priority_rank"]].to_string(index=False))

if __name__ == "__main__":
    main()
