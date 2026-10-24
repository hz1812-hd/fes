"""
FINREPORT INTELLIGENCE - Bước 3 (Xu hướng) và Bước 4 (So sánh đồng ngành)

INPUT : master1_1_with_metrics.xlsx  (đầu ra Bước 1-2: 60 dòng công ty-năm, 17 chỉ tiêu)
OUTPUT (thư mục out/):
  B3_xu_huong.xlsx                    xu hướng từng công ty x chỉ tiêu
  B4_benchmark_nganh_theonam.xlsx     benchmark ngành theo năm (tham chiếu)
  B4_so_sanh_dong_nganh.xlsx          so sánh từng công ty-năm với đồng ngành (đầu vào Bước 5)
  charts/xu_huong_<company>.png       12 biểu đồ xu hướng (mỗi công ty)
  charts/dong_nganh_<metric>_2025.png biểu đồ so sánh đồng ngành theo chỉ tiêu
  charts/scatter_DE_vs_NM_2025.png    vị thế đồng ngành và quỹ đạo 2021->2025
Chạy: python3 step3_4_trend_peer.py [đường_dẫn_master.xlsx]
"""
import sys, os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

INPUT = sys.argv[1] if len(sys.argv) > 1 else "master1_1_with_metrics.xlsx"
OUT = "out"; CH = os.path.join(OUT, "charts")
MIN_N = 5            # năm nào có < 5 doanh nghiệp thì không tính benchmark
BASE_YEAR = 2021     # năm gốc cho xu hướng (chỉ tiêu trống 2021 dùng năm đầu tiên có số)
STABLE_PCT = 0.05    # thay đổi < 5% so với năm gốc coi là ổn định
LATEST = None        # None = năm lớn nhất trong dữ liệu

# tên, đơn vị, hướng tốt (+1 càng cao càng tốt, -1 càng thấp càng tốt, 0 không kết luận hướng)
M = {
 "revenue_growth_pct": ("Tăng trưởng doanh thu", "%", +1),
 "gross_margin_pct": ("Biên lợi nhuận gộp", "%", +1),
 "operating_margin_pct": ("Biên lợi nhuận hoạt động", "%", +1),
 "net_margin_pct": ("Biên lợi nhuận ròng", "%", +1),
 "current_ratio": ("Hệ số thanh toán hiện hành", "lần", +1),
 "quick_ratio": ("Hệ số thanh toán nhanh", "lần", +1),
 "debt_to_assets_pct": ("Nợ/Tổng tài sản", "%", -1),
 "debt_to_equity": ("Nợ/Vốn chủ sở hữu", "lần", -1),
 "roa_pct": ("ROA", "%", +1),
 "roe_pct": ("ROE", "%", +1),
 "cfo_to_net_income": ("CFO/Lợi nhuận ròng", "lần", +1),
 "free_cash_flow_vnd_mn": ("Dòng tiền tự do", "triệu VND", +1),
 "dso_days": ("Số ngày thu tiền (DSO)", "ngày", -1),
 "inventory_days": ("Số ngày tồn kho", "ngày", -1),
 "dpo_days": ("Số ngày trả tiền (DPO)", "ngày", 0),
 "interest_coverage": ("Khả năng trả lãi", "lần", +1),
 "working_capital_vnd_mn": ("Vốn lưu động", "triệu VND", +1),
}
KEY = ["net_margin_pct", "debt_to_equity", "cfo_to_net_income", "dso_days", "inventory_days", "free_cash_flow_vnd_mn"]

# ---------------------------------------------------------------- đọc & kiểm tra
def load(path):
    if not os.path.exists(path):
        sys.exit(f"Không tìm thấy '{path}'. Đặt file cùng thư mục với script hoặc truyền đường dẫn.")
    raw = pd.read_excel(path, header=None)
    # tự tìm dòng tiêu đề (dòng chứa company_id và fiscal_year), bỏ qua phần trang trí
    hdr = next(i for i in range(min(15, len(raw))) if {"company_id", "fiscal_year"} <= set(raw.iloc[i].astype(str)))
    df = pd.read_excel(path, header=hdr)
    miss = [m for m in M if m not in df.columns]
    if miss: sys.exit(f"Thiếu cột chỉ tiêu: {miss}")
    df = df.dropna(subset=["company_id", "fiscal_year"]).copy()
    df["fiscal_year"] = df["fiscal_year"].astype(int)
    assert not df.duplicated(["company_id", "fiscal_year"]).any(), "Trùng company_id-fiscal_year"
    return df.sort_values(["company_id", "fiscal_year"]).reset_index(drop=True)

# ---------------------------------------------------------------- Bước 4a: benchmark
def benchmark(df):
    rows = []
    for m, (name, unit, _) in M.items():
        for y, g in df.groupby("fiscal_year"):
            x = g[m].dropna(); n = len(x)
            if n < MIN_N:
                rows.append(dict(metric=m, ten_chi_tieu=name, don_vi=unit, fiscal_year=y, n_doanh_nghiep=n,
                                 ghi_chu=f"Không đủ dữ liệu (n={n} < {MIN_N})")); continue
            q1, q3 = x.quantile(.25), x.quantile(.75); iqr = q3 - q1
            lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
            rows.append(dict(metric=m, ten_chi_tieu=name, don_vi=unit, fiscal_year=y, n_doanh_nghiep=n,
                             mean=x.mean(), median=x.median(), std=x.std(), min=x.min(), p25=q1, p75=q3, max=x.max(),
                             iqr=iqr, lower_fence=lo, upper_fence=hi, n_outlier=int(((x < lo) | (x > hi)).sum()), ghi_chu=""))
    return pd.DataFrame(rows)

# ---------------------------------------------------------------- Bước 3: xu hướng
def trend(df, latest):
    rows = []
    for c, g in df.groupby("company_id"):
        g = g.set_index("fiscal_year")
        for m, (name, unit, d) in M.items():
            s = g[m].dropna()
            if len(s) < 2: continue
            y0 = BASE_YEAR if BASE_YEAR in s.index else s.index.min()
            if latest not in s.index: continue
            v0, v1 = s[y0], s[latest]
            chg = v1 - v0
            pct = chg / abs(v0) if v0 not in (0, np.nan) and abs(v0) > 1e-9 else np.nan
            slope = np.polyfit(s.index.values, s.values, 1)[0] if len(s) >= 3 else np.nan
            # số năm xấu đi liên tiếp tính đến năm mới nhất (chỉ khi có hướng tốt/xấu)
            streak = 0
            if d != 0:
                vals = s.sort_index()
                for i in range(len(vals) - 1, 0, -1):
                    if (vals.iloc[i] - vals.iloc[i - 1]) * d < 0: streak += 1
                    else: break
            if d == 0: label = "không xác định hướng"
            elif not np.isnan(pct) and abs(pct) < STABLE_PCT: label = "ổn định"
            else: label = "cải thiện" if chg * d > 0 else "xấu đi"
            yoy = s.pct_change() if (s.abs() > 1e-9).all() else s.diff()
            rows.append(dict(company_id=c, metric=m, ten_chi_tieu=name, don_vi=unit, nam_goc=y0, gia_tri_goc=v0,
                             nam_moi_nhat=latest, gia_tri_moi_nhat=v1, thay_doi_tuyet_doi=chg, thay_doi_pct=pct,
                             do_doc_theo_nam=slope, so_nam_xau_di_lien_tiep=streak, xu_huong=label,
                             chuoi_gia_tri=" | ".join(f"{k}:{v:.2f}" for k, v in s.items())))
    return pd.DataFrame(rows)

# ---------------------------------------------------------------- Bước 4b: so sánh từng công ty với đồng ngành
def compare(df, bm):
    b = bm.set_index(["metric", "fiscal_year"])
    rows = []
    for m, (name, unit, d) in M.items():
        for y, g in df.groupby("fiscal_year"):
            if (m, y) not in b.index or b.loc[(m, y), "n_doanh_nghiep"] < MIN_N: continue
            ref = b.loc[(m, y)]
            gg = g[["company_id", m, "net_income_vnd_mn"]].dropna(subset=[m]) if "net_income_vnd_mn" in g else g[["company_id", m]].dropna()
            vals = gg.set_index("company_id")[m]
            for c, v in vals.items():
                peers = vals.drop(c)                       # đồng ngành KHÔNG gồm chính công ty
                med = peers.median()
                if d == 0: rank = np.nan
                else:
                    order = (vals * d).rank(ascending=False, method="min")   # 1 = tốt nhất
                    rank = order[c]
                pct_rank = (vals < v).sum() / (len(vals) - 1) if len(vals) > 1 else np.nan
                out = "none"
                if v < ref["lower_fence"]: out = "dưới ngưỡng"
                elif v > ref["upper_fence"]: out = "trên ngưỡng"
                if d == 0 or out == "none": side = ""
                else: side = "thuận lợi" if ((out == "trên ngưỡng") == (d > 0)) else "bất lợi"
                if d == 0: vt = "không xác định hướng"
                elif v < ref["p25"]: vt = "dưới p25 " + ("(kém)" if d > 0 else "(tốt)")
                elif v > ref["p75"]: vt = "trên p75 " + ("(tốt)" if d > 0 else "(kém)")
                else: vt = "trong khoảng p25–p75"
                note = ""
                if m == "cfo_to_net_income" and "net_income_vnd_mn" in gg.columns and gg.loc[gg.company_id == c, "net_income_vnd_mn"].iloc[0] <= 0:
                    note = "Lợi nhuận ròng ≤ 0: tỷ số không có ý nghĩa, không diễn giải là tốt"
                if note:                                   # tỷ số vô nghĩa: không xếp hạng, không gắn tốt/kém
                    vt, side, rank = "không có ý nghĩa (LN ròng ≤ 0)", "", np.nan
                rows.append(dict(company_id=c, fiscal_year=y, metric=m, ten_chi_tieu=name, gia_tri=v,
                                 trung_vi_dong_nganh_loai_chinh_no=med, chenh_lech_vs_trung_vi=v - med,
                                 chenh_lech_pct=(v - med) / abs(med) if abs(med) > 1e-9 else np.nan,
                                 hang_tot_nhat_la_1=rank, so_dn=len(vals), phan_vi=pct_rank, vi_tri=vt,
                                 ngoai_nguong_IQR=out, huong_ngoai_nguong=side, ghi_chu=note))
    return pd.DataFrame(rows)

# ---------------------------------------------------------------- biểu đồ
INK = "#1d1d1b"; PEER = "#898781"; C1 = "#2a78d6"; C2 = "#eb6834"
def chart_company(df, bm, c):
    g = df[df.company_id == c].set_index("fiscal_year")
    fig, ax = plt.subplots(2, 3, figsize=(13, 6.6)); ax = ax.ravel()
    for a, m in zip(ax, KEY):
        name, unit, _ = M[m]
        b = bm[(bm.metric == m) & (bm.n_doanh_nghiep >= MIN_N)].set_index("fiscal_year")
        a.fill_between(b.index, b.p25, b.p75, color=PEER, alpha=.18, label="p25–p75 đồng ngành")
        a.plot(b.index, b["median"], "--", color=PEER, label="Trung vị đồng ngành")
        s = g[m].where(g.net_income_vnd_mn > 0) if (m == 'cfo_to_net_income' and 'net_income_vnd_mn' in g) else g[m]
        s = s.dropna(); a.plot(s.index, s.values, "-o", color=C1, lw=2.2, ms=5, label=c)
        if len(s): a.annotate(f"{s.iloc[-1]:,.1f}", (s.index[-1], s.iloc[-1]), textcoords="offset points", xytext=(4, 6), fontsize=8, color=C1)
        a.set_title(f"{name} ({unit})", fontsize=10, color=INK, loc="left"); a.set_xticks(range(2021, 2026)); a.grid(alpha=.25)
        for sp in ("top", "right"): a.spines[sp].set_visible(False)
    h, l = ax[0].get_legend_handles_labels(); fig.legend(h, l, loc="lower center", ncol=3, frameon=False)
    fig.suptitle(f"{c}: xu hướng 2021–2025 so với đồng ngành", x=.01, ha="left", fontsize=13, fontweight="bold")
    fig.tight_layout(rect=(0, .05, 1, .95)); fig.savefig(f"{CH}/xu_huong_{c}.png", dpi=150); plt.close(fig)

def chart_peer(df, bm, m, y):
    name, unit, d = M[m]
    s = df[df.fiscal_year == y].set_index("company_id")[m].dropna().sort_values(ascending=(d <= 0))
    r = bm[(bm.metric == m) & (bm.fiscal_year == y)].iloc[0]
    fig, a = plt.subplots(figsize=(8, 4.2))
    a.barh(s.index, s.values, color=[C2 if (v < r.lower_fence or v > r.upper_fence) else C1 for v in s.values])
    a.axvline(r["median"], color=PEER, ls="--", label=f"Trung vị = {r['median']:.2f}")
    a.axvspan(r.p25, r.p75, color=PEER, alpha=.15, label="p25–p75")
    a.set_title(f"{name} ({unit}) năm {y}; cam = ngoài ngưỡng IQR", loc="left", fontsize=11, color=INK)
    a.legend(frameon=False, fontsize=8); a.grid(axis="x", alpha=.25)
    for sp in ("top", "right"): a.spines[sp].set_visible(False)
    fig.tight_layout(); fig.savefig(f"{CH}/dong_nganh_{m}_{y}.png", dpi=150); plt.close(fig)

def chart_scatter(df, y):
    s = df[df.fiscal_year == y].set_index("company_id"); s0 = df[df.fiscal_year == BASE_YEAR].set_index("company_id")
    fig, a = plt.subplots(figsize=(8, 5.6))
    for c in s.index:
        a.annotate("", xy=(s.loc[c, "debt_to_equity"], s.loc[c, "net_margin_pct"]), xytext=(s0.loc[c, "debt_to_equity"], s0.loc[c, "net_margin_pct"]),
                   arrowprops=dict(arrowstyle="->", color=PEER, lw=.9))
        a.scatter(s.loc[c, "debt_to_equity"], s.loc[c, "net_margin_pct"], color=C1, zorder=3); a.text(s.loc[c, "debt_to_equity"] + .01, s.loc[c, "net_margin_pct"] + .3, c, fontsize=8)
    a.axvline(s.debt_to_equity.median(), ls="--", color=PEER); a.axhline(s.net_margin_pct.median(), ls="--", color=PEER)
    a.set_xlabel(f"Debt/Equity {y} (lần)"); a.set_ylabel(f"Biên lợi nhuận ròng {y} (%)"); a.grid(alpha=.25)
    a.set_title(f"Vị thế đồng ngành {y} và quỹ đạo từ {BASE_YEAR} (mũi tên); đường đứt = trung vị", loc="left", fontsize=11)
    for sp in ("top", "right"): a.spines[sp].set_visible(False)
    fig.tight_layout(); fig.savefig(f"{CH}/scatter_DE_vs_NM_{y}.png", dpi=150); plt.close(fig)

# ---------------------------------------------------------------- chạy
def r2(d):
    """Làm tròn 2 chữ số thập phân khi xuất file (tính toán bên trong vẫn giữ nguyên độ chính xác)."""
    return d.round(2)

def main():
    os.makedirs(CH, exist_ok=True)
    df = load(INPUT); latest = LATEST or int(df.fiscal_year.max())
    print(f"Đọc {INPUT}: {len(df)} dòng, {df.company_id.nunique()} công ty, năm {df.fiscal_year.min()}–{latest}")

    bm = benchmark(df); tr = trend(df, latest); cp = compare(df, bm)
    r2(bm).to_excel(f"{OUT}/B4_benchmark_nganh_theonam.xlsx", index=False)
    r2(tr).to_excel(f"{OUT}/B3_xu_huong.xlsx", index=False)
    # tóm tắt nhanh năm mới nhất: số chỉ tiêu bất lợi / thuận lợi so với đồng ngành
    cp_l = cp[cp.fiscal_year == latest]
    summ = cp_l.assign(kem=cp_l.vi_tri.str.contains(r"\(kém\)"), tot=cp_l.vi_tri.str.contains(r"\(tốt\)"),
                       bat_loi=cp_l.huong_ngoai_nguong.eq("bất lợi")).groupby("company_id")[["kem", "tot", "bat_loi"]].sum() \
        .rename(columns={"kem": "so_chi_tieu_kem_hon_dong_nganh", "tot": "so_chi_tieu_tot_hon_dong_nganh", "bat_loi": "so_chi_tieu_ngoai_nguong_bat_loi"})
    xu = tr.assign(xau=tr.xu_huong.eq("xấu đi"), tot=tr.xu_huong.eq("cải thiện")).groupby("company_id")[["xau", "tot"]].sum() \
        .rename(columns={"xau": "so_chi_tieu_xau_di_so_voi_nam_goc", "tot": "so_chi_tieu_cai_thien_so_voi_nam_goc"})
    with pd.ExcelWriter(f"{OUT}/B4_so_sanh_dong_nganh.xlsx") as w:
        r2(cp).to_excel(w, sheet_name="so_sanh_chi_tiet", index=False)
        summ.join(xu).reset_index().to_excel(w, sheet_name=f"tom_tat_{latest}", index=False)

    for c in sorted(df.company_id.unique()): chart_company(df, bm, c)
    for m in KEY: chart_peer(df, bm, m, latest)
    chart_scatter(df, latest)

    # ---- tự kiểm tra
    assert bm.groupby("metric").size().eq(df.fiscal_year.nunique()).all()
    assert cp.company_id.nunique() == df.company_id.nunique()
    n_chart = len(os.listdir(CH))
    print(f"B4 benchmark: {len(bm)} dòng | B3 xu hướng: {len(tr)} dòng | B4 so sánh: {len(cp)} dòng | biểu đồ: {n_chart}")
    print(f"\nTóm tắt {latest}:"); print(summ.join(xu).to_string())
    print("\nNgoài ngưỡng IQR (năm mới nhất):")
    print(cp_l[cp_l.ngoai_nguong_IQR != "none"][["company_id", "metric", "gia_tri", "ngoai_nguong_IQR", "huong_ngoai_nguong", "ghi_chu"]].to_string(index=False))
    print(f"\nXong. Kết quả trong thư mục '{OUT}/'.")

if __name__ == "__main__":
    main()
