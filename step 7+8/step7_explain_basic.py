"""
BƯỚC 7 (mức tối thiểu): GIẢI THÍCH LÝ DO CẢNH BÁO - Rule-based String Mapping
INPUT : B6_risk_score.xlsx (cột alert_tags, predicted_primary_scenario từ Bước 5-6)
OUTPUT: B7_giai_thich.xlsx  (company_id, fiscal_year, alert_tags, alert_reason, predicted_primary_scenario)
Cách xử lý: mỗi tag được map cứng sang một câu tiếng Việt cố định; các câu của cùng một dòng được nối bằng " | ".
Chưa có số liệu cụ thể, chưa truy vết nguồn (thuộc Bước 8).
Chạy: python3 step7_explain_basic.py [đường_dẫn_B6]
"""

import sys, os
import pandas as pd

INPUT = (
    sys.argv[1]
    if len(sys.argv) > 1
    else "file output/buoc_6_tao_risk_score/B6_risk_score.xlsx"
)
OUTPUT = "file output/buoc_7_giai_thich_ly_do_canh_bao/B7_giai_thich.xlsx"

TAG_TEXT = {
    # --- tín hiệu rủi ro
    "RECEIVABLES_RISK": "Khoản phải thu tăng nhanh hơn doanh thu, số ngày thu tiền cao hơn mặt bằng ngành.",
    "INVENTORY_RISK": "Hàng tồn kho tăng nhanh hơn doanh thu, số ngày tồn kho cao hơn mặt bằng ngành.",
    "CFO_PRESSURE": "Dòng tiền từ hoạt động kinh doanh yếu, không đủ tương xứng với lợi nhuận.",
    "CFO_WEAKNESS": "Dòng tiền từ hoạt động kinh doanh thấp hơn lợi nhuận ròng.",
    "PROFIT_CASH_DIVERGENCE": "Lợi nhuận ghi nhận tốt nhưng dòng tiền kinh doanh không theo kịp.",
    "LEVERAGE_RISK": "Đòn bẩy tài chính cao hơn mặt bằng ngành.",
    "HIGH_LEVERAGE": "Tỷ lệ nợ cao trong khi hiệu quả hoạt động dưới mức trung vị ngành.",
    "INTEREST_BURDEN": "Khả năng trả lãi vay thấp hơn mặt bằng ngành.",
    "GROSS_MARGIN_DECLINE": "Biên lợi nhuận gộp giảm so với các năm trước.",
    "PROFITABILITY_PRESSURE": "Biên lợi nhuận hoạt động thấp hơn ngành và đang suy giảm.",
    "LOW_MARGIN": "Biên lợi nhuận hoạt động thấp hơn mặt bằng ngành.",
    "LIQUIDITY_STRESS": "Dòng tiền kinh doanh và dòng tiền tự do âm hoặc rất thấp, áp lực thanh khoản.",
    "LIQUIDITY_PRESSURE": "Khả năng thanh toán ngắn hạn thấp hơn ngành, dòng tiền kinh doanh dưới trung vị.",
    "WORKING_CAPITAL_PRESSURE": "Chu kỳ thu tiền và tồn kho kéo dài thêm, gây áp lực lên vốn lưu động.",
    # --- điểm tích cực
    "MARGIN_IMPROVEMENT": "Biên lợi nhuận cải thiện so với năm trước.",
    "MARGIN_RECOVERY": "Biên lợi nhuận phục hồi từ mức thấp.",
    "DELEVERAGING": "Doanh nghiệp đang giảm nợ.",
    "WORKING_CAPITAL_IMPROVEMENT": "Chu kỳ thu tiền và tồn kho được rút ngắn.",
    "STRONG_GROWTH": "Tăng trưởng doanh thu tốt, biên lợi nhuận hoạt động cao hơn ngành.",
    "STRONG_LIQUIDITY": "Khả năng thanh toán ngắn hạn mạnh.",
    "LOW_LEVERAGE": "Mức nợ thấp hơn mặt bằng ngành.",
}
NO_TAG_TEXT = (
    "Không phát hiện tín hiệu tài chính đáng chú ý so với đồng ngành cùng năm."
)


def explain(tags):
    tags = [t for t in str(tags).split(";") if t and t != "nan"]
    if not tags:
        return NO_TAG_TEXT
    return " | ".join(
        TAG_TEXT.get(t, f"Tín hiệu {t} (chưa có câu mô tả).") for t in tags
    )


def main():
    if not os.path.exists(INPUT):
        sys.exit(f"Không tìm thấy {INPUT}. Chạy Bước 5-6 trước hoặc truyền đường dẫn.")
    b6 = pd.read_excel(INPUT)
    need = {"company_id", "fiscal_year", "alert_tags", "predicted_primary_scenario"}
    if need - set(b6.columns):
        sys.exit(f"Thiếu cột: {need - set(b6.columns)}")
    out = b6[
        ["company_id", "fiscal_year", "alert_tags", "predicted_primary_scenario"]
    ].copy()
    out.insert(3, "alert_reason", out.alert_tags.map(explain))
    out = out[
        [
            "company_id",
            "fiscal_year",
            "alert_tags",
            "alert_reason",
            "predicted_primary_scenario",
        ]
    ]
    unknown = {t for s in out.alert_tags.dropna() for t in str(s).split(";")} - set(
        TAG_TEXT
    )
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    out.to_excel(OUTPUT, index=False)
    print(
        f"Đã xuất {len(out)} dòng -> {OUTPUT}"
        + (
            f" | tag chưa có câu mô tả: {sorted(unknown)}"
            if unknown
            else " | mọi tag đều có câu mô tả"
        )
    )
    print(out[out.fiscal_year == out.fiscal_year.max()].head(4).to_string(index=False))


if __name__ == "__main__":
    main()
