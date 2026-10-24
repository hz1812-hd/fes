"""Thí nghiệm Silver cho Bước 5-6 (chạy SAU step5_6_signal_risk.py; không thay đổi kết quả Bước 5-6).
 A. Feature importance: bỏ lần lượt từng họ tín hiệu, đo độ giảm chất lượng.
 B. Giảm false positive: điều kiện 'vị trí và xu hướng cùng xấu' (and_gate) so với bản gốc.
 C. So sánh nhiều phương pháp: risk_score động vs luật cũ vs Isolation Forest vs 1 chỉ tiêu đơn lẻ (AUC phát hiện HIGH/CRITICAL).
Kết quả: out/Silver_thuc_nghiem.xlsx. Lưu ý: chỉ có 18 nhãn mẫu nên mọi số đều mang tính tham khảo."""
import numpy as np, pandas as pd
from sklearn.ensemble import IsolationForest
from sklearn.metrics import roc_auc_score
import step5_6_signal_risk as S

df, bench = S.load(); L = pd.read_csv(S.LABELS)
best, cv, cache = S.tune(df, L); P = best[0]
def ev(P_):
    O, *_ = S.run(df, P_); r, J = S.evaluate(O, L); return O, r
O0, r0 = ev(P)

# ---- A. feature importance (ablation theo họ tín hiệu)
rows = [dict(ho_bi_bo="(đầy đủ)", **{k: r0[k] for k in ["severity_acc", "health_acc", "scenario_acc", "highcritical_recall", "tagfam_f1"]})]
for f in S.FAMILY:
    _, r = ev(dict(P, drop_family=f))
    rows.append(dict(ho_bi_bo=f, **{k: r[k] for k in ["severity_acc", "health_acc", "scenario_acc", "highcritical_recall", "tagfam_f1"]}))
A = pd.DataFrame(rows)
for k in ["severity_acc", "health_acc", "scenario_acc", "highcritical_recall", "tagfam_f1"]: A[f"giam_{k}"] = A[k].iloc[0] - A[k]
A["tam_quan_trong"] = A[["giam_severity_acc", "giam_health_acc", "giam_scenario_acc"]].sum(axis=1)

# ---- B. giảm false positive
rows = []
for name, P_ in [("Bản gốc (tổng có trọng số)", P), ("and_gate: vị trí VÀ xu hướng cùng xấu", dict(P, and_gate=True))]:
    O, r = ev(P_); T, _ = S.detect(df, P_)
    n_sig = int(sum((T[t] >= P_["tau"]).sum() for t in T))
    rows.append(dict(phuong_an=name, so_tin_hieu_tren_60_dong=n_sig, tag_precision=r["tagfam_precision"], tag_recall=r["tagfam_recall"], tag_f1=r["tagfam_f1"],
                     severity_acc=r["severity_acc"], health_acc=r["health_acc"], highcritical_recall=r["highcritical_recall"]))
B = pd.DataFrame(rows)

# ---- C. so sánh phương pháp (AUC phát hiện HIGH/CRITICAL trên 18 dòng có nhãn)
J = L.merge(O0, on=S.K, suffixes=("_t", "_p")); y = J.severity.isin(["HIGH", "CRITICAL"]).astype(int)
def auc_of(score_df): return roc_auc_score(y, L[S.K].merge(score_df, on=S.K).iloc[:, -1])
feat = ["gross_margin_pct", "operating_margin_pct", "net_margin_pct", "current_ratio", "debt_to_assets_pct", "debt_to_equity", "cfo_to_net_income",
        "dso_days", "inventory_days", "interest_coverage", "revenue_growth_pct"]
X = df[S.K + feat].copy()
for f in feat: X[f] = X.groupby("fiscal_year")[f].transform(lambda v: (v - v.median()) / (v.quantile(.75) - v.quantile(.25) + 1e-9))   # chuẩn hóa theo ngành cùng năm
X[feat] = X[feat].fillna(0).clip(-5, 5)
iso = IsolationForest(n_estimators=300, random_state=0).fit(X[feat]); X["score"] = -iso.score_samples(X[feat])
base = S.baseline_old_rules(df)[S.K + ["risk_score"]]
one = df[S.K].assign(score=-df.net_margin_pct)
C = pd.DataFrame([
    dict(phuong_phap="Risk score ngưỡng động (đã tinh chỉnh)", AUC_high_critical=auc_of(O0[S.K + ["risk_score"]])),
    dict(phuong_phap="Luật cũ (điểm cộng cố định)", AUC_high_critical=auc_of(base)),
    dict(phuong_phap="Isolation Forest (17 chỉ tiêu, chuẩn hóa theo ngành)", AUC_high_critical=auc_of(X[S.K + ["score"]])),
    dict(phuong_phap="Chỉ một chỉ tiêu: biên ròng thấp", AUC_high_critical=auc_of(one))])
C["so_dong_co_nhan"] = len(y); C["so_ca_high_critical"] = int(y.sum())
with pd.ExcelWriter(f"{S.OUT}/Silver_thuc_nghiem.xlsx") as w:
    A.round(3).to_excel(w, sheet_name="A_feature_importance", index=False); B.round(3).to_excel(w, sheet_name="B_giam_false_positive", index=False)
    C.round(3).to_excel(w, sheet_name="C_so_sanh_phuong_phap", index=False)
pd.set_option("display.width", 220)
print("A.\n", A[["ho_bi_bo", "severity_acc", "health_acc", "scenario_acc", "tagfam_f1", "tam_quan_trong"]].round(2).to_string(index=False))
print("\nB.\n", B.round(2).to_string(index=False)); print("\nC.\n", C.round(2).to_string(index=False))
