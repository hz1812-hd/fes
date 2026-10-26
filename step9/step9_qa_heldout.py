"""Kiểm tra tổng quát hoá Bước 9: câu hỏi mới (không có trong bộ mẫu); đáp án được tính ĐỘC LẬP bằng pandas từ các cột số liệu gốc của master
(không dùng 17 cột chỉ tiêu đã tính, không dùng engine). Chạy: python3 step9_qa_heldout.py"""
import pandas as pd, numpy as np
from step9_qa_basic import answer, NAME, D
i = b = c = n = D
def g(df, co, y, col): return float(df[(df.company_id == co) & (df.fiscal_year == y)][col].iloc[0])
def nm(co): return f"{co} - {NAME[co]}"
T=[]  # (câu hỏi, đáp án độc lập, dung sai)
T.append(("Doanh thu của C011 trong năm 2023 là bao nhiêu triệu VND?",g(i,'C011',2023,'revenue_vnd_mn'),0))
T.append(("Tổng tài sản của C002 năm 2024 là bao nhiêu triệu VND?",g(b,'C002',2024,'total_assets_vnd_mn'),0))
T.append(("Lợi nhuận sau thuế của C009 năm 2025 là bao nhiêu triệu VND?",g(i,'C009',2025,'net_income_vnd_mn'),0))
T.append(("Biên lợi nhuận hoạt động của C012 năm 2024 là bao nhiêu phần trăm?",g(i,'C012',2024,'operating_profit_vnd_mn')/g(i,'C012',2024,'revenue_vnd_mn')*100,0.05))
T.append(("Current Ratio của C003 năm 2022 bằng bao nhiêu lần?",g(b,'C003',2022,'total_current_assets_vnd_mn')/g(b,'C003',2022,'total_current_liabilities_vnd_mn'),0.01))
T.append(("Debt/Equity của C010 năm 2023 bằng bao nhiêu lần?",(g(b,'C010',2023,'short_term_debt_vnd_mn')+g(b,'C010',2023,'long_term_debt_vnd_mn'))/g(b,'C010',2023,'total_equity_vnd_mn'),0.01))
T.append(("Interest Coverage của C006 năm 2024 bằng bao nhiêu lần?",g(i,'C006',2024,'operating_profit_vnd_mn')/g(i,'C006',2024,'interest_expense_vnd_mn'),0.02))
T.append(("CFO/Net Income của C001 năm 2022 bằng bao nhiêu lần?",g(c,'C001',2022,'net_cash_from_operating_activities_vnd_mn')/g(i,'C001',2022,'net_income_vnd_mn'),0.01))
T.append(("Free Cash Flow của C008 năm 2023 là bao nhiêu triệu VND, với FCF = CFO + CAPEX?",g(c,'C008',2023,'net_cash_from_operating_activities_vnd_mn')+g(c,'C008',2023,'capital_expenditure_vnd_mn'),0))
T.append(("Tốc độ tăng doanh thu của C005 năm 2023 so với 2022 là bao nhiêu phần trăm?",(g(i,'C005',2023,'revenue_vnd_mn')/g(i,'C005',2022,'revenue_vnd_mn')-1)*100,0.05))
T.append(("Tốc độ tăng tổng nợ vay của C003 năm 2024 so với 2023 là bao nhiêu phần trăm?",((g(b,'C003',2024,'short_term_debt_vnd_mn')+g(b,'C003',2024,'long_term_debt_vnd_mn'))/(g(b,'C003',2023,'short_term_debt_vnd_mn')+g(b,'C003',2023,'long_term_debt_vnd_mn'))-1)*100,0.05))
T.append(("Biên lợi nhuận gộp của C012 đã giảm bao nhiêu điểm phần trăm từ 2021 đến 2025?",(g(i,'C012',2021,'gross_profit_vnd_mn')/g(i,'C012',2021,'revenue_vnd_mn')-g(i,'C012',2025,'gross_profit_vnd_mn')/g(i,'C012',2025,'revenue_vnd_mn'))*100,0.05))
T.append(("DSO của C007 năm 2023 là bao nhiêu ngày nếu dùng Accounts Receivable bình quân đầu-cuối năm?",(g(b,'C007',2022,'accounts_receivable_vnd_mn')+g(b,'C007',2023,'accounts_receivable_vnd_mn'))/2/g(i,'C007',2023,'revenue_vnd_mn')*365,0.2))
T.append(("Inventory Days của C008 năm 2025 là bao nhiêu ngày nếu dùng tồn kho bình quân đầu-cuối năm?",(g(b,'C008',2024,'inventory_vnd_mn')+g(b,'C008',2025,'inventory_vnd_mn'))/2/g(i,'C008',2025,'cost_of_goods_sold_vnd_mn')*365,0.2))
T.append(("DPO của C011 năm 2025 là bao nhiêu ngày nếu dùng Accounts Payable bình quân đầu-cuối năm?",(g(b,'C011',2024,'accounts_payable_vnd_mn')+g(b,'C011',2025,'accounts_payable_vnd_mn'))/2/g(i,'C011',2025,'cost_of_goods_sold_vnd_mn')*365,0.2))
T.append(("Working Capital của C005 năm 2022 là bao nhiêu triệu VND?",g(b,'C005',2022,'total_current_assets_vnd_mn')-g(b,'C005',2022,'total_current_liabilities_vnd_mn'),0))
T.append(("Tỷ lệ nợ đáo hạn trong vòng 12 tháng của C012 năm 2024 là bao nhiêu phần trăm?",g(n,'C012',2024,'debt_maturing_within_12m_pct'),0.01))
# xếp hạng / lọc / boolean — đáp án độc lập
yr=2024
gm={co:g(i,co,yr,'gross_profit_vnd_mn')/g(i,co,yr,'revenue_vnd_mn') for co in NAME}
T.append(("Doanh nghiệp nào có Gross Margin thấp nhất trong năm 2024?",nm(min(gm,key=gm.get)),None))
da={co:(g(b,co,2023,'short_term_debt_vnd_mn')+g(b,co,2023,'long_term_debt_vnd_mn'))/g(b,co,2023,'total_assets_vnd_mn') for co in NAME}
T.append(("Doanh nghiệp nào có tỷ lệ Debt/Assets thấp nhất trong năm 2023?",nm(min(da,key=da.get)),None))
cr={co:g(b,co,2022,'total_current_assets_vnd_mn')/g(b,co,2022,'total_current_liabilities_vnd_mn') for co in NAME}
T.append(("Hãy liệt kê 3 doanh nghiệp có Current Ratio cao nhất năm 2022 theo thứ tự giảm dần.",';'.join(nm(x) for x in sorted(cr,key=cr.get,reverse=True)[:3]),None))
neg=[co for co in sorted(NAME) if g(c,co,2023,'net_cash_from_operating_activities_vnd_mn')<0]
T.append(("Những doanh nghiệp nào có CFO âm trong năm 2024?",';'.join(nm(x) for x in sorted([co for co in NAME if g(c,co,2024,'net_cash_from_operating_activities_vnd_mn')<0])),None))
a=(g(b,'C002',2025,'accounts_receivable_vnd_mn')/g(b,'C002',2023,'accounts_receivable_vnd_mn'))>(g(i,'C002',2025,'revenue_vnd_mn')/g(i,'C002',2023,'revenue_vnd_mn'))
T.append(("Đối với C002, trong giai đoạn 2023-2025, khoản phải thu có tăng nhanh hơn doanh thu hay không?",'Có' if a else 'Không',None))
a=(g(b,'C010',2025,'inventory_vnd_mn')/g(b,'C010',2024,'inventory_vnd_mn'))>(g(i,'C010',2025,'revenue_vnd_mn')/g(i,'C010',2024,'revenue_vnd_mn'))
T.append(("Trong năm 2025, tốc độ tăng tồn kho của C010 so với 2024 có cao hơn tốc độ tăng doanh thu hay không?",'Có' if a else 'Không',None))
def run():
    rows=[]
    for q,exp,tol in T:
        try: a,ev,ex=answer(q)
        except Exception as e: a,ev,ex=f'ERR {e}',[],''
        try: good=abs(float(a)-float(exp))<=tol+1e-9
        except Exception: good=(str(a).strip()==str(exp).strip())
        rows.append(dict(question_text_vi=q,answer_system=a,answer_independent=exp,tolerance=tol,ok=good))
    return pd.DataFrame(rows)
if __name__=='__main__':
    R=run();print(f'held-out: {R.ok.sum()}/{len(R)}')
