"""
FINREPORT INTELLIGENCE - GOLD LAYER (một file, ba nâng cấp)
 A. Auto Audit Log  : rule-engine (Bước 5) chạy trên một lõi tính toán tự ghi vết. Mọi lần đọc dữ liệu (ô nào của master, dòng mấy, giá trị bao nhiêu,
                      công thức nào, mốc benchmark nào) được ghi lại tại khoảnh khắc luật được đánh giá. evidence_fields KHÔNG gõ tay.
 B. Dynamic Template: câu giải thích (alert_reason) sinh từ các con số đã ghi vết (giá trị thật, rào chặn ngành cùng năm). Không dùng LLM.
 C. Natural-language QA: bộ phân tích câu hỏi tiếng Việt (Bước 9) chạy trên cùng lõi; bằng chứng của đáp án cũng tự ghi vết.
INPUT : master1_1_with_metrics.xlsx, B4_benchmark_nganh_theonam.xlsx, data_dictionary.csv (tùy chọn: tự ánh xạ cột -> file nguồn),
        questions_sample.csv (+ answer_key_sample.csv, B6_risk_score.xlsx tùy chọn để đối chiếu)
PHỤ THUỘC: step5_6_signal_risk.py và step9_qa_basic.py cùng thư mục (bộ luật + bộ phân tích câu hỏi của nhóm)
OUTPUT (out_gold/): G_audit_log.json, G_audit_log_phang.xlsx, G_giai_thich_truy_vet.xlsx, G_qa_submission.csv, G_kiem_tra.json
Chạy: python3 gold_layer.py
"""
import os, re, json, time, hashlib, datetime, unicodedata
import numpy as np, pandas as pd
import step5_6_signal_risk as S5
import step9_qa_basic as Q9

MASTER, BENCH, DICT = "master1_1_with_metrics.xlsx", "B4_benchmark_nganh_theonam.xlsx", "data_dictionary.csv"
QUESTIONS, KEYFILE, TEMPLATE, B6, B9 = "questions_sample.csv", "answer_key_sample.csv", "qa_submission_template_Case3.csv", "out/B6_risk_score.xlsx", "out/B9_qa_submission.csv"
OUT = "out_gold"; MODEL = "FinReport-Gold-AutoAudit"
P = dict(tau=.25, k0=.5, t0=0, wp=.4, cuts=(.1, .35, .75), tau_press=.7)          # tham số đã tinh chỉnh ở Bước 5-6
FILES = ["company_info.csv", "income_statement.csv", "balance_sheet.csv", "cashflow.csv", "financial_notes.csv"]

def column_file_map():
    """Ánh xạ tên cột -> file nguồn Data Pack (đọc từ data_dictionary.csv; nếu thiếu thì suy từ tiền tố)."""
    m = {}
    if os.path.exists(DICT):
        d = pd.read_csv(DICT, encoding="utf-8-sig")
        for _, r in d[d.file_name.isin(FILES)].iterrows(): m.setdefault(r.field_name, r.file_name)
    return m
COLFILE = column_file_map()

# ============================================================================ LÕI TÍNH TOÁN TỰ GHI VẾT
class Core:
    """Mọi truy cập dữ liệu đi qua raw()/metric()/bench(); mỗi lần đọc được ghi vào ngữ cảnh đang mở (begin/end)."""
    def __init__(self, master=MASTER, bench=BENCH):
        raw = pd.read_excel(master, header=None)
        h = next(i for i in range(min(15, len(raw))) if {"company_id", "fiscal_year"} <= set(raw.iloc[i].astype(str)))
        df = pd.read_excel(master, header=h).dropna(subset=["company_id", "fiscal_year"]).copy()
        df["fiscal_year"] = df.fiscal_year.astype(int); df["excel_row"] = df.index + h + 2
        self.M = df.set_index(["company_id", "fiscal_year"]); self.master_path = master
        _, self.B = S5.load(master, bench)                                     # bảng benchmark (file Bước 4 + benchmark dẫn xuất biên CFO)
        self.Bi = self.B.set_index(["metric", "fiscal_year"]); self.bench_path = bench
        self._cur = None; self.cache = {}
    # ---- ngữ cảnh ghi vết
    def begin(self): self._cur = {}
    def end(self): out = list(self._cur.values()); self._cur = None; return out
    def _log(self, key, rec):
        if self._cur is not None and key not in self._cur: self._cur[key] = rec
    # ---- đọc dữ liệu gốc
    def raw(self, c, y, col):
        if (c, y) not in self.M.index or col not in self.M.columns: return np.nan
        v = self.M.at[(c, y), col]; v = float(v) if isinstance(v, (int, float, np.number)) and not pd.isna(v) else v
        self._log(("raw", c, y, col), dict(type="raw", file=COLFILE.get(col, "?"), field=col, company_id=c, fiscal_year=y, row=int(self.M.at[(c, y), "excel_row"]), value=v))
        return v
    # ---- chỉ tiêu: tính từ công thức, đọc dữ liệu gốc qua raw() nên tự có lineage
    def metric(self, c, y, name):
        key = (c, y, name); n0 = None
        if self._cur is not None: n0 = len(self._cur)
        f = FORMULAS[name][0]; v = f(self, c, y)
        v = np.nan if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))) else float(v)
        if not np.isnan(v): self._log(("metric", c, y, name), dict(type="derived", metric=name, company_id=c, fiscal_year=y, value=v, formula=FORMULAS[name][1]))
        return v
    def bench(self, metric, y, stat):
        if (metric, y) not in self.Bi.index: return np.nan
        v = self.Bi.at[(metric, y), stat]
        if pd.isna(v): return np.nan
        src = "derived_from_master(same IQR method)" if metric == "cfo_margin_pct" else os.path.basename(self.bench_path)
        self._log(("bench", metric, y, stat), dict(type="benchmark", file=src, metric=metric, fiscal_year=y, stat=stat, value=float(v)))
        return float(v)

def _avg(core, c, y, col):
    a, b = core.raw(c, y, col), core.raw(c, y - 1, col); return np.nan if pd.isna(b) else (a + b) / 2
def _div(a, b): return np.nan if (pd.isna(a) or pd.isna(b) or b == 0) else a / b
R = lambda col: (lambda k, c, y: k.raw(c, y, col))
FORMULAS = {   # tên: (hàm, mô tả công thức)
 "revenue_growth_pct": (lambda k, c, y: _div(k.raw(c, y, "revenue_vnd_mn"), k.raw(c, y - 1, "revenue_vnd_mn")) * 100 - 100 if not pd.isna(k.raw(c, y - 1, "revenue_vnd_mn")) else np.nan, "(Revenue_t / Revenue_t-1 - 1) x 100"),
 "gross_margin_pct": (lambda k, c, y: _div(k.raw(c, y, "gross_profit_vnd_mn"), k.raw(c, y, "revenue_vnd_mn")) * 100, "Gross Profit / Revenue x 100"),
 "operating_margin_pct": (lambda k, c, y: _div(k.raw(c, y, "operating_profit_vnd_mn"), k.raw(c, y, "revenue_vnd_mn")) * 100, "Operating Profit / Revenue x 100"),
 "net_margin_pct": (lambda k, c, y: _div(k.raw(c, y, "net_income_vnd_mn"), k.raw(c, y, "revenue_vnd_mn")) * 100, "Net Income / Revenue x 100"),
 "current_ratio": (lambda k, c, y: _div(k.raw(c, y, "total_current_assets_vnd_mn"), k.raw(c, y, "total_current_liabilities_vnd_mn")), "Current Assets / Current Liabilities"),
 "quick_ratio": (lambda k, c, y: _div(k.raw(c, y, "cash_and_cash_equivalents_vnd_mn") + k.raw(c, y, "accounts_receivable_vnd_mn"), k.raw(c, y, "total_current_liabilities_vnd_mn")), "(Cash + AR) / Current Liabilities"),
 "debt_to_assets_pct": (lambda k, c, y: _div(k.raw(c, y, "short_term_debt_vnd_mn") + k.raw(c, y, "long_term_debt_vnd_mn"), k.raw(c, y, "total_assets_vnd_mn")) * 100, "(ShortDebt + LongDebt) / Total Assets x 100"),
 "debt_to_equity": (lambda k, c, y: _div(k.raw(c, y, "short_term_debt_vnd_mn") + k.raw(c, y, "long_term_debt_vnd_mn"), k.raw(c, y, "total_equity_vnd_mn")), "(ShortDebt + LongDebt) / Total Equity"),
 "roa_pct": (lambda k, c, y: _div(k.raw(c, y, "net_income_vnd_mn"), _avg(k, c, y, "total_assets_vnd_mn")) * 100, "Net Income / Average Total Assets x 100"),
 "roe_pct": (lambda k, c, y: _div(k.raw(c, y, "net_income_vnd_mn"), _avg(k, c, y, "total_equity_vnd_mn")) * 100, "Net Income / Average Equity x 100"),
 "cfo_to_net_income": (lambda k, c, y: _div(k.raw(c, y, "net_cash_from_operating_activities_vnd_mn"), k.raw(c, y, "net_income_vnd_mn")), "CFO / Net Income"),
 "cfo_to_net_income_valid": (lambda k, c, y: _div(k.raw(c, y, "net_cash_from_operating_activities_vnd_mn"), k.raw(c, y, "net_income_vnd_mn")) if k.raw(c, y, "net_income_vnd_mn") > 0 else np.nan, "CFO / Net Income (chỉ khi Net Income > 0)"),
 "free_cash_flow_vnd_mn": (lambda k, c, y: k.raw(c, y, "net_cash_from_operating_activities_vnd_mn") + k.raw(c, y, "capital_expenditure_vnd_mn"), "CFO + CAPEX (CAPEX mang dấu âm)"),
 "dso_days": (lambda k, c, y: _div(_avg(k, c, y, "accounts_receivable_vnd_mn"), k.raw(c, y, "revenue_vnd_mn")) * 365, "Average AR / Revenue x 365"),
 "inventory_days": (lambda k, c, y: _div(_avg(k, c, y, "inventory_vnd_mn"), k.raw(c, y, "cost_of_goods_sold_vnd_mn")) * 365, "Average Inventory / COGS x 365"),
 "dpo_days": (lambda k, c, y: _div(_avg(k, c, y, "accounts_payable_vnd_mn"), k.raw(c, y, "cost_of_goods_sold_vnd_mn")) * 365, "Average AP / COGS x 365"),
 "interest_coverage": (lambda k, c, y: _div(k.raw(c, y, "operating_profit_vnd_mn"), k.raw(c, y, "interest_expense_vnd_mn")), "Operating Profit / Interest Expense"),
 "working_capital_vnd_mn": (lambda k, c, y: k.raw(c, y, "total_current_assets_vnd_mn") - k.raw(c, y, "total_current_liabilities_vnd_mn"), "Current Assets - Current Liabilities"),
 # chỉ tiêu phụ cho rule-engine
 "cfo_margin_pct": (lambda k, c, y: _div(k.raw(c, y, "net_cash_from_operating_activities_vnd_mn"), k.raw(c, y, "revenue_vnd_mn")) * 100, "CFO / Revenue x 100"),
 "total_debt_vnd_mn": (lambda k, c, y: k.raw(c, y, "short_term_debt_vnd_mn") + k.raw(c, y, "long_term_debt_vnd_mn"), "ShortDebt + LongDebt"),
 "ar_growth_pct": (lambda k, c, y: (_div(k.raw(c, y, "accounts_receivable_vnd_mn"), k.raw(c, y - 1, "accounts_receivable_vnd_mn")) - 1) * 100, "(AR_t / AR_t-1 - 1) x 100"),
 "inv_growth_pct": (lambda k, c, y: (_div(k.raw(c, y, "inventory_vnd_mn"), k.raw(c, y - 1, "inventory_vnd_mn")) - 1) * 100, "(Inventory_t / Inventory_t-1 - 1) x 100"),
 "debt_growth_pct": (lambda k, c, y: (_div(k.raw(c, y, "short_term_debt_vnd_mn") + k.raw(c, y, "long_term_debt_vnd_mn"), k.raw(c, y - 1, "short_term_debt_vnd_mn") + k.raw(c, y - 1, "long_term_debt_vnd_mn")) - 1) * 100, "(Debt_t / Debt_t-1 - 1) x 100"),
}
MASTER_METRICS = ["revenue_growth_pct", "gross_margin_pct", "operating_margin_pct", "net_margin_pct", "current_ratio", "quick_ratio", "debt_to_assets_pct", "debt_to_equity",
                  "roa_pct", "roe_pct", "cfo_to_net_income", "free_cash_flow_vnd_mn", "dso_days", "inventory_days", "dpo_days", "interest_coverage", "working_capital_vnd_mn"]

# ============================================================================ A. RULE-ENGINE TỰ GHI VẾT (cùng logic Bước 5-6)
ramp = lambda x, a, b: np.nan if pd.isna(x) else float(np.clip((x - a) / (b - a), 0, 1))
def S_tr(k, c, y, m, d, wp=None, vm=None):
    """Điểm tín hiệu 0..1 = wp * (vị trí so với quartile ngành cùng năm) + (1-wp) * (xu hướng chuẩn hóa theo IQR). Mọi lần đọc được ghi vết."""
    wp = P["wp"] if wp is None else wp; vm = vm or m
    v = k.metric(c, y, vm); ref = k.bench(m, y, "p75" if d > 0 else "p25"); iqr = k.bench(m, y, "iqr")
    for st in ("p25", "median", "p75", "upper_fence", "lower_fence"): k.bench(m, y, st)      # toàn bộ mốc ngành dùng làm bối cảnh cho lời giải thích
    pp = ramp(d * (v - ref) / iqr, P["k0"], 1.5) if not (pd.isna(v) or pd.isna(ref) or pd.isna(iqr) or iqr == 0) else np.nan
    ch = [d * (v - k.metric(c, y - 1, vm)), d * (v - k.metric(c, y - 2, vm)) / 2]; ch = [x for x in ch if not pd.isna(x)]
    tp = ramp(max(ch) / iqr, P["t0"], P["t0"] + 1) if ch and not (pd.isna(iqr) or iqr == 0) else np.nan
    s = (wp * pp + (1 - wp) * tp) if not pd.isna(tp) else pp
    return (0.0 if pd.isna(s) else float(s)), pp, tp
def gt0(x): return (not pd.isna(x)) and x > 0

def detect_traced(k, c, y):
    """Trả về (scores, logs, positives, pos_logs) cho một công ty-năm. logs[tag] = danh sách bản ghi truy vết."""
    T, LG = {}, {}
    def tag(name, fn):
        k.begin(); s = fn(); LG[name] = k.end(); T[name] = float(np.clip(0 if pd.isna(s) else s, 0, 1))
    tau = P["tau"]
    arg, rg = k.metric(c, y, "ar_growth_pct"), k.metric(c, y, "revenue_growth_pct")   # (ghi ra ngoài ngữ cảnh; thêm lại vào log của tag cần)
    tag("RECEIVABLES_RISK", lambda: S_tr(k, c, y, "dso_days", +1)[0] * (gt0(k.metric(c, y, "ar_growth_pct") - k.metric(c, y, "revenue_growth_pct"))))
    tag("INVENTORY_RISK", lambda: S_tr(k, c, y, "inventory_days", +1)[0] * (gt0(k.metric(c, y, "inv_growth_pct") - k.metric(c, y, "revenue_growth_pct"))))
    # chất lượng dòng tiền
    ni, cfo = k.raw(c, y, "net_income_vnd_mn"), k.raw(c, y, "net_cash_from_operating_activities_vnd_mn")
    k.begin(); s0, _, _ = S_tr(k, c, y, "cfo_to_net_income", -1, vm="cfo_to_net_income_valid")
    for st in ("p25", "p75"): k.bench("cfo_to_net_income", y, st)
    med = k.bench("cfo_to_net_income", y, "median"); x = k.metric(c, y, "cfo_to_net_income"); nm_ = k.metric(c, y, "net_margin_pct"); nmmed = k.bench("net_margin_pct", y, "median")
    ni = k.raw(c, y, "net_income_vnd_mn"); cfo = k.raw(c, y, "net_cash_from_operating_activities_vnd_mn"); cq = k.end()
    s = s0 if (ni > 0 and not pd.isna(med) and x < med) else 0.0
    if cfo < 0: s = 1.0
    div = (not pd.isna(nmmed)) and nm_ >= nmmed and cfo > 0 and s >= tau
    for nme, val in [("PROFIT_CASH_DIVERGENCE", s if div else 0.0), ("CFO_PRESSURE", s if (not div and s >= P["tau_press"]) else 0.0), ("CFO_WEAKNESS", s if (not div and s < P["tau_press"]) else 0.0)]:
        T[nme] = float(np.clip(val, 0, 1)); LG[nme] = cq
    tag("LEVERAGE_RISK", lambda: max(S_tr(k, c, y, "debt_to_assets_pct", +1)[0], S_tr(k, c, y, "debt_to_equity", +1)[0]) if (k.metric(c, y, "debt_growth_pct") or True) else 0)
    tag("INTEREST_BURDEN", lambda: S_tr(k, c, y, "interest_coverage", -1)[0])
    def gm():
        s = S_tr(k, c, y, "gross_margin_pct", -1, wp=.25)[0]; a, b = k.metric(c, y, "gross_margin_pct"), k.metric(c, y - 2, "gross_margin_pct")
        return s * (gt0(b - a) if not pd.isna(b) else 0)
    tag("GROSS_MARGIN_DECLINE", gm)
    def prof():
        s = S_tr(k, c, y, "operating_margin_pct", -1)[0]; a = k.metric(c, y, "operating_margin_pct")
        d1, d2 = a - k.metric(c, y - 1, "operating_margin_pct"), a - k.metric(c, y - 2, "operating_margin_pct")
        return s * ((0 if pd.isna(d1) else d1) < 0 or (0 if pd.isna(d2) else d2) < 0)
    tag("PROFITABILITY_PRESSURE", prof)
    def liq():
        f, cf = k.metric(c, y, "free_cash_flow_vnd_mn"), k.raw(c, y, "net_cash_from_operating_activities_vnd_mn")
        cm = S_tr(k, c, y, "cfo_margin_pct", -1)[0]                       # luôn đọc biên CFO và mốc ngành (bối cảnh cho lời giải thích)
        if cf < 0 and f < 0: return 1.0
        return cm if f < 0 else 0.0
    tag("LIQUIDITY_STRESS", liq)
    def wc():
        z = []
        for m in ("dso_days", "inventory_days"):
            a, b, iq = k.metric(c, y, m), k.metric(c, y - 1, m), k.bench(m, y, "iqr")
            if not (pd.isna(a) or pd.isna(b) or pd.isna(iq) or iq == 0): z.append((a - b) / iq)
        gate = any(T[t] > 0 for t in ("CFO_PRESSURE", "PROFIT_CASH_DIVERGENCE", "LIQUIDITY_STRESS", "CFO_WEAKNESS"))
        r = ramp(max(z), P["t0"], P["t0"] + 1) if z else np.nan
        return (0.0 if pd.isna(r) else r) * gate
    tag("WORKING_CAPITAL_PRESSURE", wc)
    for t in ("CFO_PRESSURE", "PROFIT_CASH_DIVERGENCE", "LIQUIDITY_STRESS", "CFO_WEAKNESS"):   # tín hiệu điều kiện hóa WC: gộp vết của chúng khi điều kiện thỏa
        if T[t] > 0 and T["WORKING_CAPITAL_PRESSURE"] > 0: LG["WORKING_CAPITAL_PRESSURE"] += [e for e in LG[t] if e not in LG["WORKING_CAPITAL_PRESSURE"]]
    def low_margin():
        s = S_tr(k, c, y, "operating_margin_pct", -1, wp=1.0)[0]; return s if T["PROFITABILITY_PRESSURE"] < tau else 0.0
    tag("LOW_MARGIN", low_margin)
    def high_lev():
        s = S_tr(k, c, y, "debt_to_assets_pct", +1, wp=1.0)[0]; om, omm = k.metric(c, y, "operating_margin_pct"), k.bench("operating_margin_pct", y, "median")
        return s if (not pd.isna(omm) and om < omm and T["LEVERAGE_RISK"] < tau) else 0.0
    tag("HIGH_LEVERAGE", high_lev)
    def liq_pr():
        s = S_tr(k, c, y, "current_ratio", -1, wp=1.0)[0]; cm, cmm = k.metric(c, y, "cfo_margin_pct"), k.bench("cfo_margin_pct", y, "median")
        return s if (not pd.isna(cmm) and cm < cmm) else 0.0
    tag("LIQUIDITY_PRESSURE", liq_pr)
    # điểm tích cực
    POS, PL = {}, {}
    def pos(name, fn):
        k.begin(); r = bool(fn()); PL[name] = k.end(); POS[name] = r
    def z(m, d):
        a, b, iq = k.metric(c, y, m), k.metric(c, y - 1, m), k.bench(m, y, "iqr")
        return np.nan if (pd.isna(a) or pd.isna(b) or pd.isna(iq) or iq == 0) else d * (a - b) / iq
    def mi_calc():
        gz, oz = z("gross_margin_pct", 1), z("operating_margin_pct", 1)
        om, omm = k.metric(c, y, "operating_margin_pct"), k.bench("operating_margin_pct", y, "median")
        return (gz >= .25) and (oz >= .25), (om >= omm)
    pos("MARGIN_IMPROVEMENT", lambda: mi_calc()[0] and mi_calc()[1]); pos("MARGIN_RECOVERY", lambda: mi_calc()[0] and not mi_calc()[1])
    pos("DELEVERAGING", lambda: (k.metric(c, y, "debt_growth_pct") < 0) and (z("debt_to_assets_pct", -1) >= .25))
    pos("WORKING_CAPITAL_IMPROVEMENT", lambda: (z("dso_days", -1) >= .3) and (z("inventory_days", -1) >= .3))
    pos("STRONG_GROWTH", lambda: (k.metric(c, y, "revenue_growth_pct") > k.bench("revenue_growth_pct", y, "median")) and (k.metric(c, y, "operating_margin_pct") > k.bench("operating_margin_pct", y, "p75")))
    pos("STRONG_LIQUIDITY", lambda: k.metric(c, y, "current_ratio") > k.bench("current_ratio", y, "upper_fence"))
    pos("LOW_LEVERAGE", lambda: k.metric(c, y, "debt_to_assets_pct") < k.bench("debt_to_assets_pct", y, "p25"))
    return T, LG, POS, PL

# ============================================================================ B. DYNAMIC TEMPLATE (số liệu lấy từ bản ghi vết)
class View:
    """Truy cập giá trị đã ghi vết theo tên; thiếu bản ghi -> KeyError (đảm bảo mọi con số trong câu đều có nguồn)."""
    def __init__(self, log, y): self.l = log; self.y = y
    def m(self, name, dy=0): return next(e["value"] for e in self.l if e["type"] == "derived" and e["metric"] == name and e["fiscal_year"] == self.y + dy)
    def r(self, col, dy=0): return next(e["value"] for e in self.l if e["type"] == "raw" and e["field"] == col and e["fiscal_year"] == self.y + dy)
    def b(self, metric, stat): return next(e["value"] for e in self.l if e["type"] == "benchmark" and e["metric"] == metric and e["stat"] == stat and e["fiscal_year"] == self.y)
    def has(self, kind, **kw): return any(e["type"] == kind and all(e.get(a) == v for a, v in kw.items()) for e in self.l)
ud = lambda v: "tăng" if v >= 0 else "giảm"
def beyond(v, ref, d): return (v > ref) if d > 0 else (v < ref)

def sentence(tag, V, y):
    def cmp(m, d, stat, unit="", dec=0):
        v, ref = V.m(m), V.b(m, stat); side = "cao hơn" if d > 0 else "thấp hơn"
        w = f"{side} {stat} ngành năm {y}" if beyond(v, ref, d) else f"chưa {'vượt' if d > 0 else 'chạm'} {stat} ngành năm {y}"
        return f"{w} ({ref:,.{dec}f}{unit}; trung vị {V.b(m,'median'):,.{dec}f}{unit})"
    ch = lambda m, dy=-1, dec=0: (V.m(m) - V.m(m, dy))
    if tag == "RECEIVABLES_RISK":
        return f"Số ngày thu tiền (DSO) đạt {V.m('dso_days'):.0f} ngày, {cmp('dso_days', 1, 'p75', ' ngày')}; khoản phải thu {ud(V.m('ar_growth_pct'))} {abs(V.m('ar_growth_pct')):.1f}% trong khi doanh thu {ud(V.m('revenue_growth_pct'))} {abs(V.m('revenue_growth_pct')):.1f}%."
    if tag == "INVENTORY_RISK":
        return f"Số ngày tồn kho đạt {V.m('inventory_days'):.0f} ngày, {cmp('inventory_days', 1, 'p75', ' ngày')}; tồn kho {ud(V.m('inv_growth_pct'))} {abs(V.m('inv_growth_pct')):.1f}% trong khi doanh thu {ud(V.m('revenue_growth_pct'))} {abs(V.m('revenue_growth_pct')):.1f}%."
    if tag in ("CFO_PRESSURE", "CFO_WEAKNESS", "PROFIT_CASH_DIVERGENCE"):
        cfo, ni = V.r("net_cash_from_operating_activities_vnd_mn"), V.r("net_income_vnd_mn")
        base = f"Dòng tiền kinh doanh (CFO) {'âm ' if cfo < 0 else 'đạt '}{cfo:,.0f} triệu VND so với lợi nhuận ròng {ni:,.0f} triệu"
        if ni > 0: base += f" (CFO/Lợi nhuận ròng = {cfo/ni:.2f}; p25 ngành {V.b('cfo_to_net_income','p25'):.2f}, trung vị {V.b('cfo_to_net_income','median'):.2f})"
        if tag == "PROFIT_CASH_DIVERGENCE": base += f"; biên lợi nhuận ròng {V.m('net_margin_pct'):.1f}% cao hơn trung vị ngành ({V.b('net_margin_pct','median'):.1f}%) nhưng tiền thu về không theo kịp"
        return base + "."
    if tag in ("LEVERAGE_RISK", "HIGH_LEVERAGE"):
        s = f"Nợ vay/Tổng tài sản đạt {V.m('debt_to_assets_pct'):.1f}%, {cmp('debt_to_assets_pct', 1, 'p75', '%', 1)}"
        if V.has("derived", metric="debt_growth_pct"): s += f"; nợ vay {ud(V.m('debt_growth_pct'))} {abs(V.m('debt_growth_pct')):.1f}% so với năm trước"
        if V.has("derived", metric="debt_to_equity"): s += f"; Debt/Equity = {V.m('debt_to_equity'):.2f}"
        return s + "."
    if tag == "INTEREST_BURDEN":
        return f"Hệ số trả lãi (Interest Coverage) chỉ {V.m('interest_coverage'):.1f} lần, {cmp('interest_coverage', -1, 'p25', ' lần', 1)}."
    if tag == "GROSS_MARGIN_DECLINE":
        return f"Biên lợi nhuận gộp {V.m('gross_margin_pct'):.1f}%, giảm {V.m('gross_margin_pct', -2) - V.m('gross_margin_pct'):.1f} điểm % so với hai năm trước ({V.m('gross_margin_pct', -2):.1f}%); trung vị ngành {V.b('gross_margin_pct','median'):.1f}%."
    if tag in ("PROFITABILITY_PRESSURE", "LOW_MARGIN"):
        s = f"Biên lợi nhuận hoạt động đạt {V.m('operating_margin_pct'):.1f}%, {cmp('operating_margin_pct', -1, 'p25', '%', 1)}"
        if V.has("derived", metric="operating_margin_pct", fiscal_year=y - 1): s += f"; thay đổi {V.m('operating_margin_pct') - V.m('operating_margin_pct', -1):+.1f} điểm % so với năm trước"
        return s + "."
    if tag == "LIQUIDITY_STRESS":
        return f"Dòng tiền kinh doanh {V.r('net_cash_from_operating_activities_vnd_mn'):,.0f} triệu và dòng tiền tự do {V.m('free_cash_flow_vnd_mn'):,.0f} triệu VND (âm); biên CFO/doanh thu {V.m('cfo_margin_pct'):.1f}% so với p25 ngành {V.b('cfo_margin_pct','p25'):.1f}%."
    if tag == "LIQUIDITY_PRESSURE":
        return f"Hệ số thanh toán hiện hành {V.m('current_ratio'):.2f}, {cmp('current_ratio', -1, 'p25', '', 2)}; biên CFO/doanh thu {V.m('cfo_margin_pct'):.1f}% thấp hơn trung vị ngành ({V.b('cfo_margin_pct','median'):.1f}%)."
    if tag == "WORKING_CAPITAL_PRESSURE":
        return f"Số ngày thu tiền {V.m('dso_days') - V.m('dso_days', -1):+.0f} ngày và số ngày tồn kho {V.m('inventory_days') - V.m('inventory_days', -1):+.0f} ngày so với năm trước, đi cùng dòng tiền yếu."
    # điểm tích cực
    if tag in ("MARGIN_IMPROVEMENT", "MARGIN_RECOVERY"):
        return f"Biên lợi nhuận gộp {V.m('gross_margin_pct') - V.m('gross_margin_pct', -1):+.1f} điểm % và biên hoạt động {V.m('operating_margin_pct') - V.m('operating_margin_pct', -1):+.1f} điểm % so với năm trước (biên hoạt động {V.m('operating_margin_pct'):.1f}%; trung vị ngành {V.b('operating_margin_pct','median'):.1f}%)."
    if tag == "DELEVERAGING": return f"Nợ vay giảm {abs(V.m('debt_growth_pct')):.1f}% so với năm trước; Nợ/Tổng tài sản còn {V.m('debt_to_assets_pct'):.1f}%."
    if tag == "WORKING_CAPITAL_IMPROVEMENT": return f"Số ngày thu tiền {V.m('dso_days') - V.m('dso_days', -1):+.0f} ngày và số ngày tồn kho {V.m('inventory_days') - V.m('inventory_days', -1):+.0f} ngày so với năm trước."
    if tag == "STRONG_GROWTH": return f"Doanh thu tăng {V.m('revenue_growth_pct'):.1f}% (trung vị ngành {V.b('revenue_growth_pct','median'):.1f}%), biên hoạt động {V.m('operating_margin_pct'):.1f}% vượt p75 ngành ({V.b('operating_margin_pct','p75'):.1f}%)."
    if tag == "STRONG_LIQUIDITY": return f"Hệ số thanh toán hiện hành {V.m('current_ratio'):.2f}, vượt rào chặn trên của ngành ({V.b('current_ratio','upper_fence'):.2f})."
    if tag == "LOW_LEVERAGE": return f"Nợ vay/Tổng tài sản chỉ {V.m('debt_to_assets_pct'):.1f}%, thấp hơn p25 ngành ({V.b('debt_to_assets_pct','p25'):.1f}%)."
    raise KeyError(tag)

def evidence_fields(log, c):
    g = {}
    for e in log:
        if e["type"] == "raw": g.setdefault(e["file"], {"y": set(), "f": []}); g[e["file"]]["y"].add(e["fiscal_year"]); g[e["file"]]["f"].append(e["field"]) if e["field"] not in g[e["file"]]["f"] else None
    lab = lambda ys: (f"{min(ys)}-{max(ys)}" if len(ys) > 1 and sorted(ys) == list(range(min(ys), max(ys) + 1)) else ";".join(map(str, sorted(ys))))
    return ";".join(f"{f}:{c}:{lab(v['y'])}:{'+'.join(v['f'])}" for f, v in g.items())

# ============================================================================ C. QA TỰ GHI VẾT
SYN = [(r"hệ số thanh toán hiện hành", "current ratio"), (r"hệ số thanh toán nhanh", "quick ratio"), (r"dòng tiền tự do", "free cash flow"),
       (r"khả năng (thanh toán|trả) lãi( vay)?", "interest coverage"), (r"số ngày thu tiền bình quân", "dso"), (r"số ngày (phải trả|trả tiền)", "dpo"),
       (r"nợ (vay )?(trên|/) ?tổng tài sản", "debt/assets"), (r"nợ (vay )?(trên|/) ?vốn chủ( sở hữu)?", "debt/equity"), (r"(?<!biên )lợi nhuận ròng", "lợi nhuận sau thuế")]
def normalize_question(t):
    t = unicodedata.normalize("NFC", t); t = re.sub(r"\s+", " ", t).strip()
    for rx, rep in SYN: t = re.sub(rx, rep, t, flags=re.I)
    return t

def qa_run(core, questions, to_ev):
    """Chạy bộ phân tích câu hỏi của Bước 9 trên lõi tự ghi vết: đáp án và bằng chứng đều đi ra từ cùng các lần đọc dữ liệu."""
    def gval(c, y, key):
        if key in FORMULAS: return core.metric(c, y, key)
        return core.raw(c, y, key)
    Q9.val = gval; rows, audit = [], []
    for _, q in questions.iterrows():
        t0 = time.perf_counter(); core.begin()
        try: a, _, expl = Q9.answer(normalize_question(q.question_text_vi), q.get("target_company_id"), q.get("target_fiscal_year")); err = ""
        except Exception as e: a, expl, err = "", f"Chưa trả lời được: {e}", str(e)
        log = core.end(); ms = (time.perf_counter() - t0) * 1000
        raws = [e for e in log if e["type"] == "raw"]; ncomp = len({e["company_id"] for e in raws})
        ev = []
        for e in raws: ev.append((e["file"], "all_companies" if ncomp >= len(Q9.NAME) else e["company_id"], e["fiscal_year"], e["field"]))
        rows.append(dict(question_id=q.question_id, answer=a, evidence_source_files=";".join(sorted({x[0] for x in ev})), evidence_source_fields=Q9.fmt_evidence(ev), explanation=expl,
                         confidence=0.95 if a != "" else 0.0, model_name=MODEL, processing_time_ms=round(ms, 3)))
        audit.append(dict(question_id=q.question_id, answer=a, inputs=[e for e in log if e["type"] in ("raw", "derived")]))
    return pd.DataFrame(rows), audit

# ============================================================================ CHẠY TOÀN BỘ
def sha(path): return hashlib.sha256(open(path, "rb").read()).hexdigest()[:16]
def main():
    os.makedirs(OUT, exist_ok=True); t_all = time.time(); k = Core(); checks = {}
    # ---- 0. kiểm tra lõi: chỉ tiêu tính lại bằng công thức == chỉ tiêu trong master (Bước 2)
    bad = 0; tot = 0
    for (c, y) in k.M.index:
        for m in MASTER_METRICS:
            mv = k.M.at[(c, y), m]; v = k.metric(c, y, m); tot += 1
            if (pd.isna(mv) != pd.isna(v)) or (not pd.isna(mv) and abs(mv - v) > 1e-6 * max(1, abs(mv))): bad += 1
    checks["metric_recompute"] = dict(cells=tot, mismatches=bad)
    # ---- A. rule-engine tự ghi vết
    t0 = time.time(); recs, T_rows, POS_rows = [], [], []
    keys = sorted(k.M.index)
    for c, y in keys:
        T, LG, POS, PL = detect_traced(k, c, y); T_rows.append(T); POS_rows.append(POS); recs.append((c, y, T, LG, POS, PL))
    t_detect = time.time() - t0
    # đối chiếu với Bước 5-6 (vectorized)
    df5, _ = S5.load(); O5, T5, _, pos5 = S5.run(df5, P)
    Tdf = pd.DataFrame(T_rows, index=pd.MultiIndex.from_tuples(keys, names=["company_id", "fiscal_year"]))
    T5df = pd.DataFrame({t: T5[t].values for t in T5}, index=pd.MultiIndex.from_frame(df5[["company_id", "fiscal_year"]]))
    maxdiff = float((Tdf.sort_index() - T5df.sort_index()[Tdf.columns]).abs().max().max())
    # kết quả cuối (risk score, severity, health, scenario) dùng đúng hàm của Bước 5-6
    Tser = {t: pd.Series(Tdf[t].reindex(pd.MultiIndex.from_frame(df5[["company_id", "fiscal_year"]])).values, index=df5.index) for t in Tdf.columns}
    score, _ = S5.score_rows(df5, Tser, P); sev = pd.Series(np.array(S5.SEV)[np.searchsorted(P["cuts"], score.values, side="right")], index=df5.index)
    posd = {k_: pd.Series([dict(zip([(c, y) for c, y, *_ in recs], [r[4] for r in recs]))[(c_, y_)][k_] for c_, y_ in zip(df5.company_id, df5.fiscal_year)], index=df5.index) for k_ in recs[0][4]}
    out_rows, prevs = [], {}
    for i, r in df5.iterrows():
        fired = [t for t in Tdf.columns if Tser[t].iloc[i] >= P["tau"]]; ps = [kk for kk, v in posd.items() if v.iloc[i]]
        scen = S5.scenario(r, fired, ps); hl = S5.health(r, sev.iloc[i], ps, prevs.get(r.company_id), score.iloc[i]); prevs[r.company_id] = score.iloc[i]
        out_rows.append(dict(company_id=r.company_id, fiscal_year=int(r.fiscal_year), risk_score=round(float(score.iloc[i]), 4), predicted_severity=sev.iloc[i], predicted_overall_health=hl,
                             predicted_primary_scenario=scen, tags=fired if fired else ps, fired=fired, pos=ps))
    G = pd.DataFrame(out_rows)
    checks["engine_equals_step5_6"] = dict(max_abs_score_diff=maxdiff, risk_score_equal=bool(np.allclose(G.risk_score.values, O5.risk_score.round(4).values, atol=1e-4)),
                                           severity_equal=bool((G.predicted_severity.values == O5.predicted_severity.values).all()), health_equal=bool((G.predicted_overall_health.values == O5.predicted_overall_health.values).all()),
                                           scenario_equal=bool((G.predicted_primary_scenario.values == O5.predicted_primary_scenario.values).all()), tags_equal=int(sum(set(a) == set(str(b).split(";")) - {""} if b == b else set(a) == set() for a, b in zip(G.tags, O5.alert_tags))))
    # ---- B + A: giải thích động và audit log
    reasons, evid, audit_recs, flat = [], [], [], []
    lookup = {(c, y): (T, LG, POS, PL) for c, y, T, LG, POS, PL in recs}; unmatched = 0; no_raw = 0
    for r in out_rows:
        c, y = r["company_id"], r["fiscal_year"]; T, LG, POS, PL = lookup[(c, y)]; parts, logs_all = [], []
        if r["fired"]:
            for t in sorted(r["fired"], key=lambda t: -T[t]):
                V = View(LG[t], y); parts.append(f"{t}: " + sentence(t, V, y)); logs_all += LG[t]
                if not any(e["type"] == "raw" for e in LG[t]): no_raw += 1
                inputs = [dict(e) for e in LG[t]]; audit_recs.append(dict(company_id=c, fiscal_year=y, tag=t, family=S5.TAG2FAM[t], signal_score=round(T[t], 4), inputs=inputs))
                for e in inputs: flat.append(dict(company_id=c, fiscal_year=y, tag=t, signal_score=round(T[t], 4), **{kk: vv for kk, vv in e.items() if kk not in ("company_id", "fiscal_year")}, input_company=e.get("company_id"), input_year=e["fiscal_year"]))
        elif r["pos"]:
            for t in r["pos"]:
                V = View(PL[t], y); parts.append(f"{t}: " + sentence(t, V, y)); logs_all += PL[t]
            parts = ["Không có tín hiệu rủi ro vượt ngưỡng động của ngành cùng năm. Điểm tích cực: " + " | ".join(parts)] if parts else parts
        if not parts: parts = ["Không có tín hiệu rủi ro vượt ngưỡng động của ngành cùng năm."]
        reasons.append(" | ".join(parts)); evid.append(evidence_fields(logs_all, c) or "income_statement.csv;balance_sheet.csv;cashflow.csv")
    G["alert_tags"] = [";".join(t) for t in G.tags]; G["alert_reason"] = reasons; G["evidence_fields"] = evid
    G["priority_rank"] = G.sort_values(["fiscal_year", "risk_score"], ascending=[True, False]).groupby("fiscal_year").cumcount().add(1).reindex(G.index)
    # kiểm tra toàn vẹn bằng chứng: giá trị ghi vết == ô tương ứng trong master (đúng dòng Excel đã ghi)
    chk = mism = 0
    for rec in audit_recs:
        for e in rec["inputs"]:
            if e["type"] == "raw":
                chk += 1
                if abs(k.M.at[(e["company_id"], e["fiscal_year"]), e["field"]] - e["value"]) > 1e-9 or int(k.M.at[(e["company_id"], e["fiscal_year"]), "excel_row"]) != e["row"]: mism += 1
    n_fired = len(audit_recs)
    checks["evidence_integrity"] = dict(fired_signals=n_fired, signals_with_raw_evidence=n_fired - no_raw, traceability_pct=round(100 * (n_fired - no_raw) / max(n_fired, 1), 1), raw_cells_checked=chk, mismatches=mism)
    checks["explanations"] = dict(rows=len(G), contains_nan=int(G.alert_reason.str.contains("nan", case=False).sum()))
    cols = ["company_id", "fiscal_year", "risk_score", "predicted_severity", "predicted_overall_health", "predicted_primary_scenario", "alert_tags", "alert_reason", "evidence_fields", "priority_rank"]
    G[cols].to_excel(f"{OUT}/G_giai_thich_truy_vet.xlsx", index=False)
    pd.DataFrame(flat).to_excel(f"{OUT}/G_audit_log_phang.xlsx", index=False)
    # ---- C. QA
    Qs = pd.read_csv(QUESTIONS, encoding="utf-8-sig"); QS, qa_audit = qa_run(k, Qs, None)
    if os.path.exists(TEMPLATE): QS = QS[list(pd.read_csv(TEMPLATE, encoding="utf-8-sig").columns)]
    QS.to_csv(f"{OUT}/G_qa_submission.csv", index=False, encoding="utf-8-sig")
    qa = dict(questions=len(QS), answered=int((QS.answer != "").sum()), mean_latency_ms=round(float(QS.processing_time_ms.mean()), 3))
    if os.path.exists(B9):
        b9 = pd.read_csv(B9, encoding="utf-8-sig").set_index("question_id"); qa["answers_equal_B9"] = int(sum(str(QS.set_index("question_id").answer[i]) == str(b9.answer[i]) for i in b9.index))
    if os.path.exists(KEYFILE):
        K_ = pd.read_csv(KEYFILE, encoding="utf-8-sig").set_index("question_id"); ok = tot_ = 0; evm = 0
        def pairs(s): return {(seg.split(":")[0], f) for seg in str(s).split("|") for f in (seg.split(":")[-1].split(";") if seg.count(":") >= 3 else [])}
        for _, r in QS.iterrows():
            kk = K_.loc[r.question_id]; ref, a = str(kk.reference_answer), str(r.answer)
            if kk.answer_unit_or_type == "short_text": continue
            try: good = abs(float(a) - float(ref)) <= float(kk.numeric_tolerance) + 1e-9
            except Exception: good = a.strip() == ref.strip() or (str(kk.normalized_answer).upper() == "TRUE" and a == "Có")
            ok += good; tot_ += 1
            evm += pairs(r.evidence_source_fields) >= pairs(kk.evidence_reference)      # bằng chứng tự ghi vết bao phủ đủ (file, trường) của đáp án mẫu
        qa.update(auto_graded=tot_, correct=int(ok), evidence_covers_reference=int(evm))
    # thử diễn đạt khác (không có trong bộ mẫu)
    para = [("Hệ số thanh toán hiện hành của C007 năm 2025 là bao nhiêu?", None), ("Khả năng trả lãi vay của C005 năm 2025 bằng bao nhiêu lần?", None),
            ("Dòng tiền tự do của C004 năm 2025 là bao nhiêu triệu VND?", None), ("Lợi nhuận ròng của C009 năm 2025 là bao nhiêu triệu VND?", None)]
    expected = [Q9.val("C007", 2025, "current_ratio"), Q9.val("C005", 2025, "interest_coverage"), Q9.val("C004", 2025, "free_cash_flow_vnd_mn"), Q9.val("C009", 2025, "net_income_vnd_mn")]
    pq = []
    for (q, _), e in zip(para, expected):
        core_a = qa_run(k, pd.DataFrame([dict(question_id="P", question_text_vi=q, target_company_id=None, target_fiscal_year=None)]), None)[0].answer.iloc[0]
        try: pq.append(abs(float(core_a) - float(e)) < 0.01)
        except Exception: pq.append(False)
    qa["paraphrase_ok"] = f"{sum(pq)}/{len(pq)}"
    checks["qa"] = qa
    # ---- audit log JSON (run log + từng tín hiệu + từng câu QA)
    log = dict(run=dict(run_id=datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S"), model=MODEL, layer="Gold", parameters={kk: (list(v) if isinstance(v, tuple) else v) for kk, v in P.items()},
                        input_hash={os.path.basename(p): sha(p) for p in (MASTER, BENCH) if os.path.exists(p)}, rows=len(G), signals=n_fired, detect_time_s=round(t_detect, 3), total_time_s=round(time.time() - t_all, 2)),
               checks=checks, signals=audit_recs, qa=qa_audit)
    json.dump(log, open(f"{OUT}/G_audit_log.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=lambda o: None if (isinstance(o, float) and np.isnan(o)) else (float(o) if isinstance(o, (np.floating,)) else int(o) if isinstance(o, np.integer) else str(o)))
    json.dump(checks, open(f"{OUT}/G_kiem_tra.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    pd.set_option("display.max_colwidth", 400, "display.width", 250)
    print(json.dumps(checks, ensure_ascii=False, indent=1))
    ex = G[(G.company_id == "C004") & (G.fiscal_year == 2025)].iloc[0]
    print("\nVí dụ C004 2025:\n", ex.alert_reason, "\n", ex.evidence_fields)

if __name__ == "__main__":
    main()
