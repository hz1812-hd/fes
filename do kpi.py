import pandas as pd
from sklearn.metrics import precision_score, recall_score, f1_score


def calculate_kpis(df_predict, df_truth):
    # Nối dữ liệu dự đoán và đáp án
    merged = pd.merge(df_predict, df_truth, on=["company_id", "fiscal_year"])

    # 1. Đo lường Alert Tổng quát (Binary: Có rủi ro vs Khỏe mạnh)
    y_pred_alert = (merged["predicted_overall_health"] != "HEALTHY").astype(int)
    y_true_alert = (merged["true_overall_health"] != "HEALTHY").astype(int)

    precision = precision_score(y_true_alert, y_pred_alert)
    recall = recall_score(y_true_alert, y_pred_alert)
    f1 = f1_score(y_true_alert, y_pred_alert)

    print(f"--- KPI CẢNH BÁO TỔNG QUÁT ---")
    print(f"Precision: {precision:.2%} (Hạn chế báo động giả)")
    print(f"Recall:    {recall:.2%} (Hạn chế bỏ sót rủi ro)")
    print(f"F1-Score:  {f1:.2%}")

    # 2. Đo lường High/Critical Recall (Không bỏ lọt bạo bệnh)
    # Lọc ra những ca nguy kịch thực sự
    critical_cases = merged[merged["true_overall_health"] == "HIGH_RISK"]

    if len(critical_cases) > 0:
        # Trong số các ca nguy kịch thực sự, model bắt được bao nhiêu ca?
        correctly_caught = critical_cases[
            critical_cases["predicted_overall_health"] == "HIGH_RISK"
        ]
        critical_recall = len(correctly_caught) / len(critical_cases)

        print(f"\n--- KPI BẮT BẠO BỆNH (High/Critical) ---")
        print(
            f"Bắt trúng: {len(correctly_caught)} / {len(critical_cases)} ca nguy kịch."
        )
        print(f"Critical Recall: {critical_recall:.2%}")
    else:
        print("\nKhông có ca HIGH_RISK nào trong tập đáp án mẫu.")


# Cách gọi hàm (Ví dụ)
# df_predict = pd.read_excel('step5_anomaly_detection_results.xlsx')
# df_truth = pd.read_csv('financial_scenarios_sample.csv')
# calculate_kpis(df_predict, df_truth)
