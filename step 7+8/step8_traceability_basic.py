"""
BƯỚC 8 (mức tối thiểu): TRUY VẾT NHẬN XÉT (TRACEABILITY) - gắn cứng tên trường dữ liệu
INPUT : B7_giai_thich.xlsx (cột alert_tags; có thể dùng B6_risk_score.xlsx)
OUTPUT: B8_truy_vet.xlsx (company_id, fiscal_year, alert_tags, evidence_fields)
Cách xử lý: mỗi tag được lập trình viên gắn cứng với danh sách (file nguồn, trường dữ liệu) trong bảng TAG_FIELDS.
evidence_fields = các trường đó, kèm công ty và năm của dòng, ví dụ "balance_sheet.csv:C004:2025:inventory_vnd_mn".
Chưa dùng trích xuất tự động từ công thức (thuộc mức nâng cao).
Chạy: python3 step8_traceability_basic.py [đường_dẫn_B7]
"""
import sys, os
import pandas as pd

INPUT = sys.argv[1] if len(sys.argv) > 1 else "out/B7_giai_thich.xlsx"
MASTER = "master1_1_with_metrics.xlsx"          # chỉ để kiểm tra tên trường có tồn tại (không bắt buộc)
OUTPUT = "out/B8_truy_vet.xlsx"

IS, BS, CF = "income_statement.csv", "balance_sheet.csv", "cashflow.csv"
DEBT = [(BS, "short_term_debt_vnd_mn"), (BS, "long_term_debt_vnd_mn"), (BS, "total_assets_vnd_mn")]
TAG_FIELDS = {   # tag -> các (file, trường) gắn cứng
    "RECEIVABLES_RISK": [(BS, "accounts_receivable_vnd_mn"), (IS, "revenue_vnd_mn")],
    "INVENTORY_RISK": [(BS, "inventory_vnd_mn"), (IS, "cost_of_goods_sold_vnd_mn"), (IS, "revenue_vnd_mn")],
    "CFO_PRESSURE": [(CF, "net_cash_from_operating_activities_vnd_mn"), (IS, "net_income_vnd_mn")],
    "CFO_WEAKNESS": [(CF, "net_cash_from_operating_activities_vnd_mn"), (IS, "net_income_vnd_mn")],
    "PROFIT_CASH_DIVERGENCE": [(CF, "net_cash_from_operating_activities_vnd_mn"), (IS, "net_income_vnd_mn")],
    "LEVERAGE_RISK": DEBT + [(BS, "total_equity_vnd_mn")],
    "HIGH_LEVERAGE": DEBT + [(BS, "total_equity_vnd_mn")],
    "INTEREST_BURDEN": [(IS, "operating_profit_vnd_mn"), (IS, "interest_expense_vnd_mn")],
    "GROSS_MARGIN_DECLINE": [(IS, "revenue_vnd_mn"), (IS, "cost_of_goods_sold_vnd_mn")],
    "PROFITABILITY_PRESSURE": [(IS, "revenue_vnd_mn"), (IS, "operating_profit_vnd_mn")],
    "LOW_MARGIN": [(IS, "revenue_vnd_mn"), (IS, "operating_profit_vnd_mn")],
    "LIQUIDITY_STRESS": [(CF, "net_cash_from_operating_activities_vnd_mn"), (CF, "capital_expenditure_vnd_mn"), (IS, "revenue_vnd_mn")],
    "LIQUIDITY_PRESSURE": [(BS, "total_current_assets_vnd_mn"), (BS, "total_current_liabilities_vnd_mn"), (CF, "net_cash_from_operating_activities_vnd_mn")],
    "WORKING_CAPITAL_PRESSURE": [(BS, "accounts_receivable_vnd_mn"), (BS, "inventory_vnd_mn"), (IS, "revenue_vnd_mn"), (IS, "cost_of_goods_sold_vnd_mn")],
    "MARGIN_IMPROVEMENT": [(IS, "revenue_vnd_mn"), (IS, "cost_of_goods_sold_vnd_mn"), (IS, "operating_profit_vnd_mn")],
    "MARGIN_RECOVERY": [(IS, "revenue_vnd_mn"), (IS, "cost_of_goods_sold_vnd_mn"), (IS, "operating_profit_vnd_mn")],
    "DELEVERAGING": DEBT,
    "WORKING_CAPITAL_IMPROVEMENT": [(BS, "accounts_receivable_vnd_mn"), (BS, "inventory_vnd_mn"), (IS, "revenue_vnd_mn"), (IS, "cost_of_goods_sold_vnd_mn")],
    "STRONG_GROWTH": [(IS, "revenue_vnd_mn"), (IS, "operating_profit_vnd_mn")],
    "STRONG_LIQUIDITY": [(BS, "total_current_assets_vnd_mn"), (BS, "total_current_liabilities_vnd_mn")],
    "LOW_LEVERAGE": DEBT,
}

def evidence(company, year, tags):
    tags = [t for t in str(tags).split(";") if t and t != "nan"]
    if not tags:   # không có tín hiệu: ghi nguồn ở mức báo cáo
        return ";".join(f"{f}:{company}:{year}" for f in (IS, BS, CF))
    seen, out = set(), []
    for t in tags:
        for f, c in TAG_FIELDS.get(t, []):
            k = (f, c)
            if k not in seen: seen.add(k); out.append(k)
    by_file = {}
    for f, c in out: by_file.setdefault(f, []).append(c)           # gom các trường cùng file: file:công ty:năm:trường1+trường2
    return ";".join(f"{f}:{company}:{year}:{'+'.join(cols)}" for f, cols in by_file.items())

def main():
    if not os.path.exists(INPUT): sys.exit(f"Không tìm thấy {INPUT}. Chạy Bước 7 trước hoặc truyền đường dẫn.")
    d = pd.read_excel(INPUT)
    if not {"company_id", "fiscal_year", "alert_tags"} <= set(d.columns): sys.exit("Thiếu cột company_id, fiscal_year hoặc alert_tags")
    out = d[["company_id", "fiscal_year", "alert_tags"]].copy()
    out["evidence_fields"] = [evidence(r.company_id, int(r.fiscal_year), r.alert_tags) for r in out.itertuples()]
    unknown = {t for s in out.alert_tags.dropna() for t in str(s).split(";")} - set(TAG_FIELDS)
    if os.path.exists(MASTER):   # kiểm tra tên trường có thật trong Data Pack
        raw = pd.read_excel(MASTER, header=None)
        h = next(i for i in range(min(15, len(raw))) if "company_id" in set(raw.iloc[i].astype(str)))
        cols = set(pd.read_excel(MASTER, header=h).columns)
        missing = {c for v in TAG_FIELDS.values() for _, c in v} - cols
        print("Kiểm tra tên trường:", "tất cả tồn tại trong master" if not missing else f"THIẾU {sorted(missing)}")
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True); out.to_excel(OUTPUT, index=False)
    print(f"Đã xuất {len(out)} dòng -> {OUTPUT}" + (f" | tag chưa gắn trường: {sorted(unknown)}" if unknown else ""))
    pd.set_option("display.max_colwidth", 200, "display.width", 250)
    print(out[(out.fiscal_year == out.fiscal_year.max())].head(4).to_string(index=False))

if __name__ == "__main__":
    main()
