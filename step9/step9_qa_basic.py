"""
BƯỚC 9 (mức tối thiểu): NUMERICAL QA - trả lời câu hỏi số liệu từ Data Pack bằng rule-engine
INPUT : master1_1_with_metrics.xlsx (số liệu gốc + 17 chỉ tiêu, Bước 1-2), questions_sample.csv (câu hỏi), qa_submission_template_Case3.csv (khung đầu ra)
        answer_key_sample.csv (tùy chọn: chỉ để TỰ CHẤM, hệ thống KHÔNG dùng khi trả lời)
OUTPUT: out/B9_qa_submission.csv  (question_id, answer, evidence_source_files, evidence_source_fields, explanation, confidence, model_name, processing_time_ms)
        out/B9_qa_tu_cham.csv     (đối chiếu với đáp án mẫu, nếu có answer_key)
Cách xử lý: (1) nhận diện chỉ tiêu, doanh nghiệp, năm, phép toán từ câu hỏi tiếng Việt bằng từ khóa/regex;
            (2) tra hoặc tính từ master theo công thức của Data Pack; (3) trả đáp án + file/trường/kỳ dữ liệu nguồn.
Không tra theo question_id, không dùng đáp án mẫu. Chạy: python3 step9_qa_basic.py [questions.csv]
"""
import os, re, sys, time
import numpy as np, pandas as pd

MASTER = "master1_1_with_metrics.xlsx"
QUESTIONS = sys.argv[1] if len(sys.argv) > 1 else "questions_sample.csv"
TEMPLATE = "qa_submission_template_Case3.csv"; KEYFILE = "answer_key_sample.csv"
OUT = "out"; MODEL = "FinReport-QA-RuleEngine-basic"

def load_master(path=MASTER):
    raw = pd.read_excel(path, header=None)
    h = next(i for i in range(min(15, len(raw))) if {"company_id", "fiscal_year"} <= set(raw.iloc[i].astype(str)))
    df = pd.read_excel(path, header=h).dropna(subset=["company_id", "fiscal_year"]).copy(); df.fiscal_year = df.fiscal_year.astype(int)
    df["total_debt_vnd_mn"] = df.short_term_debt_vnd_mn + df.long_term_debt_vnd_mn
    return df
D = load_master(); A = D.set_index(["company_id", "fiscal_year"])
NAME = D.drop_duplicates("company_id").set_index("company_id").company_name.to_dict()
def val(c, y, k): return float(A.loc[(c, y), k])
def nm(c): return f"{c} - {NAME[c]}"
def fmt(v, d): return str(int(round(v))) if d == 0 else str(round(v, d))

IS, BS, CF, FN = "income_statement.csv", "balance_sheet.csv", "cashflow.csv", "financial_notes.csv"
F = lambda f, *c: [(f, x) for x in c]
# (regex, cột trong master, số chữ số thập phân, [(file, trường nguồn)])
RATIO = [
 (r"free cash flow|fcf", "free_cash_flow_vnd_mn", 0, F(CF, "net_cash_from_operating_activities_vnd_mn", "capital_expenditure_vnd_mn")),
 (r"cfo\s*/\s*net income|cfo\s*/\s*ni", "cfo_to_net_income", 3, F(CF, "net_cash_from_operating_activities_vnd_mn") + F(IS, "net_income_vnd_mn")),
 (r"interest coverage", "interest_coverage", 3, F(IS, "operating_profit_vnd_mn", "interest_expense_vnd_mn")),
 (r"quick ratio", "quick_ratio", 3, F(BS, "cash_and_cash_equivalents_vnd_mn", "accounts_receivable_vnd_mn", "total_current_liabilities_vnd_mn")),
 (r"current ratio", "current_ratio", 3, F(BS, "total_current_assets_vnd_mn", "total_current_liabilities_vnd_mn")),
 (r"debt\s*/\s*assets", "debt_to_assets_pct", 2, F(BS, "short_term_debt_vnd_mn", "long_term_debt_vnd_mn", "total_assets_vnd_mn")),
 (r"debt\s*/\s*equity", "debt_to_equity", 3, F(BS, "short_term_debt_vnd_mn", "long_term_debt_vnd_mn", "total_equity_vnd_mn")),
 (r"gross margin|biên lợi nhuận gộp", "gross_margin_pct", 2, F(IS, "gross_profit_vnd_mn", "revenue_vnd_mn")),
 (r"operating margin|biên lợi nhuận hoạt động", "operating_margin_pct", 2, F(IS, "operating_profit_vnd_mn", "revenue_vnd_mn")),
 (r"net margin|biên lợi nhuận ròng", "net_margin_pct", 2, F(IS, "net_income_vnd_mn", "revenue_vnd_mn")),
 (r"\bdso\b", "dso_days", 1, F(BS, "accounts_receivable_vnd_mn") + F(IS, "revenue_vnd_mn")),
 (r"inventory days|số ngày tồn kho", "inventory_days", 1, F(BS, "inventory_vnd_mn") + F(IS, "cost_of_goods_sold_vnd_mn")),
 (r"\bdpo\b", "dpo_days", 1, F(BS, "accounts_payable_vnd_mn") + F(IS, "cost_of_goods_sold_vnd_mn")),
 (r"working capital|vốn lưu động", "working_capital_vnd_mn", 0, F(BS, "total_current_assets_vnd_mn", "total_current_liabilities_vnd_mn")),
 (r"\broa\b", "roa_pct", 2, F(IS, "net_income_vnd_mn") + F(BS, "total_assets_vnd_mn")),
 (r"\broe\b", "roe_pct", 2, F(IS, "net_income_vnd_mn") + F(BS, "total_equity_vnd_mn")),
 (r"capex intensity|cường độ capex", "capex_intensity_pct", 2, F(FN, "capex_intensity_pct")),
 (r"nợ đáo hạn", "debt_maturing_within_12m_pct", 2, F(FN, "debt_maturing_within_12m_pct")),
 (r"thời hạn tín dụng|credit terms", "average_customer_credit_terms_days", 0, F(FN, "average_customer_credit_terms_days")),
 (r"lãi suất vay", "average_borrowing_rate_pct", 2, F(FN, "average_borrowing_rate_pct")),
 (r"tập trung khách hàng", "major_customer_concentration_pct", 0, F(FN, "major_customer_concentration_pct")),
 (r"xuất khẩu", "export_revenue_pct", 0, F(FN, "export_revenue_pct"))]
GROWTH = [(r"doanh thu", "revenue_vnd_mn", IS, [("revenue_vnd_mn",)]), (r"tổng nợ vay|nợ vay", "total_debt_vnd_mn", BS, [("short_term_debt_vnd_mn",), ("long_term_debt_vnd_mn",)]),
          (r"phải thu", "accounts_receivable_vnd_mn", BS, [("accounts_receivable_vnd_mn",)]), (r"tồn kho", "inventory_vnd_mn", BS, [("inventory_vnd_mn",)])]
DIRECT = [(r"cfo|lưu chuyển tiền thuần từ hoạt động kinh doanh", "net_cash_from_operating_activities_vnd_mn", CF), (r"tổng nợ phải trả", "total_liabilities_vnd_mn", BS),
 (r"tổng tài sản", "total_assets_vnd_mn", BS), (r"vốn chủ sở hữu", "total_equity_vnd_mn", BS), (r"giá vốn", "cost_of_goods_sold_vnd_mn", IS),
 (r"lợi nhuận gộp", "gross_profit_vnd_mn", IS), (r"lợi nhuận hoạt động", "operating_profit_vnd_mn", IS), (r"lợi nhuận sau thuế|net income", "net_income_vnd_mn", IS),
 (r"chi phí lãi vay", "interest_expense_vnd_mn", IS), (r"doanh thu", "revenue_vnd_mn", IS), (r"tiền và tương đương|cash and cash", "cash_and_cash_equivalents_vnd_mn", BS),
 (r"phải thu", "accounts_receivable_vnd_mn", BS), (r"tồn kho", "inventory_vnd_mn", BS), (r"capex", "capital_expenditure_vnd_mn", CF)]
SRC = {"revenue_vnd_mn": IS, "accounts_receivable_vnd_mn": BS, "inventory_vnd_mn": BS}

def find_metric(t):
    for rx, key, dec, src in RATIO:
        if re.search(rx, t): return key, dec, src
    return None
def growth_pct(c, y0, y1, col): return (val(c, y1, col) / val(c, y0, col) - 1) * 100
def metric_order(t):   # thứ tự xuất hiện của phải thu / tồn kho / doanh thu trong câu (cho câu boolean)
    cand = sorted([(k, t.find(w)) for k, w in [("accounts_receivable_vnd_mn", "phải thu"), ("inventory_vnd_mn", "tồn kho"), ("revenue_vnd_mn", "doanh thu")] if t.find(w) >= 0], key=lambda x: x[1])
    return cand[0][0], cand[1][0]

def years_label(ys):
    ys = sorted(set(ys)); return f"{ys[0]}-{ys[-1]}" if len(ys) > 1 and ys == list(range(ys[0], ys[-1] + 1)) else ";".join(map(str, ys))
def fmt_evidence(ev):
    """ev: list (file, company|'all_companies', year, field) -> 'file:công ty:năm:trường1;trường2|file:...' giống đáp án mẫu của BTC"""
    g = {}
    for f, c, y, fld in ev:
        k = (f, c); g.setdefault(k, {"y": [], "f": []}); g[k]["y"].append(y); g[k]["f"].append(fld) if fld not in g[k]["f"] else None
    return "|".join(f"{f}:{c}:{years_label(v['y'])}:{';'.join(v['f'])}" for (f, c), v in g.items())

def answer(text, tc=None, ty=None):
    t = text.lower(); comps = [c.upper() for c in re.findall(r"c\d{3}", t)] or ([tc] if isinstance(tc, str) else [])
    yrs = [int(y) for y in re.findall(r"20\d\d", t)] or ([int(str(ty)[:4])] if ty == ty and ty not in (None, "") else [])
    ev = []
    def E(files, c, ys):
        for f, fld in files:
            for y in ys: ev.append((f, c, y, fld))
    # 1) lọc danh sách: "những doanh nghiệp nào có ... âm"
    if re.search(r"những doanh nghiệp nào", t) and re.search(r"\bâm\b", t):
        y = yrs[0]; m = find_metric(t); key, files = (m[0], m[2]) if m else ("net_cash_from_operating_activities_vnd_mn", F(CF, "net_cash_from_operating_activities_vnd_mn"))
        hits = sorted(c for c in NAME if val(c, y, key) < 0); E(files, "all_companies", [y])
        return ";".join(nm(c) for c in hits), ev, f"Lọc các doanh nghiệp có {key} < 0 năm {y}"
    # 2) đối chiếu ending cash
    if "ending cash" in t and "cash and cash equivalents" in t:
        c, y = comps[0], yrs[0]; a, b = val(c, y, "ending_cash_vnd_mn"), val(c, y, "cash_and_cash_equivalents_vnd_mn")
        ev += [(CF, c, y, "ending_cash_vnd_mn"), (BS, c, y, "cash_and_cash_equivalents_vnd_mn")]
        return ("Có" if a == b else "Không"), ev, f"Ending Cash = {a:,.0f}; tiền trên bảng cân đối = {b:,.0f}"
    # 3) mô tả xu hướng
    if "xu hướng" in t and comps and len(yrs) >= 2 and "nào" not in t.replace("nào đó", ""):
        c = comps[0]; y0, y1 = min(yrs), max(yrs); Y = list(range(y0, y1 + 1))
        s = lambda k: [val(c, y, k) for y in Y]; rev, cfo, cn, ar = s("revenue_vnd_mn"), s("net_cash_from_operating_activities_vnd_mn"), s("cfo_to_net_income"), s("accounts_receivable_vnd_mn")
        pc = lambda x: (x[-1] / x[0] - 1) * 100
        parts = [f"Doanh thu {'tăng' if rev[-1] > rev[0] else 'giảm'} {abs(pc(rev)):.1f}% ({y0}-{y1})", f"CFO {'tăng' if cfo[-1] > cfo[0] else 'giảm'} {abs(pc(cfo)):.1f}% ({cfo[0]:,.0f} → {cfo[-1]:,.0f} triệu)",
                 f"CFO/Lợi nhuận ròng từ {cn[0]:.2f} xuống {cn[-1]:.2f}", f"phải thu {'tăng' if ar[-1] > ar[0] else 'giảm'} {abs(pc(ar)):.1f}%"]
        imp = []
        if rev[-1] > rev[0] and cn[-1] < cn[0]: imp.append("khả năng chuyển lợi nhuận thành dòng tiền suy yếu (chất lượng dòng tiền giảm)")
        if pc(ar) > pc(rev): imp.append("phải thu tăng nhanh hơn doanh thu, tạo áp lực vốn lưu động")
        E(F(IS, "revenue_vnd_mn", "net_income_vnd_mn") + F(CF, "net_cash_from_operating_activities_vnd_mn") + F(BS, "accounts_receivable_vnd_mn"), c, Y)
        return "; ".join(parts) + ". Gợi ý: " + ("; ".join(imp) if imp else "chưa thấy tín hiệu bất thường") + ".", ev, "Mô tả xu hướng từ số liệu nguồn"
    # 4) xếp hạng toàn ngành
    if re.search(r"doanh nghiệp nào|liệt kê", t) and re.search(r"cao nhất|thấp nhất|mạnh nhất|lớn nhất|nhỏ nhất", t):
        n = 3 if re.search(r"\b3\b|ba doanh nghiệp", t) else 1; m = find_metric(t)
        if not m: raise ValueError("không nhận diện chỉ tiêu xếp hạng")
        key, dec, files = m
        if re.search(r"suy giảm|giảm.*mạnh nhất", t) and len(yrs) >= 2:
            y0, y1 = min(yrs), max(yrs); vals = {c: val(c, y1, key) - val(c, y0, key) for c in NAME}; desc = False; E(files, "all_companies", [y0, y1]); expl = f"Thay đổi {key} {y0}→{y1}; nhỏ nhất = suy giảm mạnh nhất"
        else:
            y = yrs[0]; vals = {c: val(c, y, key) for c in NAME}; desc = not re.search(r"thấp nhất|nhỏ nhất", t); E(files, "all_companies", [y]); expl = f"Xếp hạng {key} năm {y}"
        top = sorted(vals, key=lambda c: vals[c], reverse=desc)[:n]
        return ";".join(nm(c) for c in top), ev, expl + ": " + ", ".join(f"{c}={vals[c]:.3f}" for c in top)
    # 5) Có/Không: A tăng nhanh hơn B
    if re.search(r"có .*hay không", t) and re.search(r"nhanh hơn|cao hơn", t):
        c = comps[0]; a, b = metric_order(t); y0, y1 = (min(yrs), max(yrs)) if len(yrs) >= 2 else (yrs[0] - 1, yrs[0])
        ga, gb = growth_pct(c, y0, y1, a), growth_pct(c, y0, y1, b)
        for k in (a, b): ev += [(SRC[k], c, y, k) for y in (y0, y1)]
        return ("Có" if ga > gb else "Không"), ev, f"{a} tăng {ga:.1f}% so với {b} tăng {gb:.1f}% ({y0}→{y1})"
    # 6) thay đổi điểm phần trăm
    if re.search(r"(giảm|tăng) bao nhiêu điểm phần trăm", t) and len(yrs) >= 2:
        key, dec, files = find_metric(t); c = comps[0]; y0, y1 = min(yrs), max(yrs)
        d = val(c, y0, key) - val(c, y1, key) if "giảm" in t else val(c, y1, key) - val(c, y0, key); E(files, c, [y0, y1])
        return fmt(d, 2), ev, f"{key} {y0} = {val(c, y0, key):.2f}, {y1} = {val(c, y1, key):.2f}"
    # 7) tốc độ tăng trưởng
    if re.search(r"tốc độ tăng|tăng trưởng", t) and comps and yrs:
        for rx, col, fl, raws in GROWTH:
            if re.search(rx, t):
                c, y = comps[0], yrs[0]
                for r_ in raws:
                    for yy in (y - 1, y): ev.append((fl, c, yy, r_[0]))
                g_ = growth_pct(c, y - 1, y, col); return fmt(g_, 2), ev, f"{col}: {val(c, y - 1, col):,.0f} → {val(c, y, col):,.0f}, tăng {g_:.2f}%"
    # 8) tỷ số, chỉ tiêu, thuyết minh
    m = find_metric(t)
    if m and comps and yrs:
        key, dec, files = m; c, y = comps[0], yrs[0]
        E(files, c, [y - 1, y] if key in ("dso_days", "inventory_days", "dpo_days", "roa_pct", "roe_pct") else [y])
        return fmt(val(c, y, key), dec), ev, f"{key} của {c} năm {y} = {val(c, y, key):.4f}"
    # 9) giá trị trực tiếp
    for rx, key, fl in DIRECT:
        if re.search(rx, t) and comps and yrs:
            c, y = comps[0], yrs[0]; ev.append((fl, c, y, key)); return fmt(val(c, y, key), 0), ev, f"Giá trị trực tiếp {key}"
    raise ValueError("không nhận diện được câu hỏi")

def main():
    os.makedirs(OUT, exist_ok=True)
    Q = pd.read_csv(QUESTIONS, encoding="utf-8-sig"); rows = []
    for _, q in Q.iterrows():
        t0 = time.perf_counter()
        try: a, ev, expl = answer(q.question_text_vi, q.get("target_company_id"), q.get("target_fiscal_year")); err = ""
        except Exception as e: a, ev, expl, err = "", [], f"Chưa trả lời được: {e}", str(e)
        ms = (time.perf_counter() - t0) * 1000
        rows.append(dict(question_id=q.question_id, answer=a, evidence_source_files=";".join(sorted({e[0] for e in ev})), evidence_source_fields=fmt_evidence(ev),
                         explanation=expl, confidence=0.95 if a != "" else 0.0, model_name=MODEL, processing_time_ms=round(ms, 3)))
    S = pd.DataFrame(rows)
    if os.path.exists(TEMPLATE): S = S[list(pd.read_csv(TEMPLATE, encoding="utf-8-sig").columns)]   # đúng thứ tự cột mẫu
    S.to_csv(f"{OUT}/B9_qa_submission.csv", index=False, encoding="utf-8-sig")
    print(f"Đã trả lời {int((S.answer != '').sum())}/{len(S)} câu -> {OUT}/B9_qa_submission.csv | thời gian trung bình {S.processing_time_ms.mean():.2f} ms/câu")
    if os.path.exists(KEYFILE):   # tự chấm (chỉ để đo KPI; không ảnh hưởng đáp án)
        K = pd.read_csv(KEYFILE, encoding="utf-8-sig").set_index("question_id"); g = []
        for _, r in S.iterrows():
            k = K.loc[r.question_id]; ref, a = str(k.reference_answer), str(r.answer); tol = k.numeric_tolerance
            if k.question_type == "TREND_ANALYSIS" and k.answer_unit_or_type == "short_text": ok = None          # văn bản: chấm tay
            else:
                try: ok = abs(float(a) - float(ref)) <= float(tol) + 1e-9
                except Exception: ok = a.strip() == ref.strip() or (str(k.normalized_answer).upper() == "TRUE" and a == "Có")
            g.append(dict(question_id=r.question_id, question_type=k.question_type, difficulty=k.difficulty, answer=a, reference=ref, dung=ok))
        G = pd.DataFrame(g); G.to_csv(f"{OUT}/B9_qa_tu_cham.csv", index=False, encoding="utf-8-sig"); auto = G[G.dung.notna()]
        print(f"Tự chấm: {int(auto.dung.sum())}/{len(auto)} câu có đáp án số/lựa chọn đúng ({100*auto.dung.mean():.0f}%); {int(G.dung.isna().sum())} câu văn bản cần chấm tay")
        if (auto.dung == False).any(): print(auto[auto.dung == False].to_string(index=False))
        print(G.groupby("question_type").dung.agg(["count", "sum"]).rename(columns={"count": "so_cau", "sum": "dung"}).to_string())

if __name__ == "__main__":
    main()
