import streamlit as st

# ── Page config ────────────────────────────────────────────────────────────
st.set_page_config(page_title="房贷计算器", page_icon="🏠", layout="wide")

# ── Helper ─────────────────────────────────────────────────────────────────
def monthly_payment(principal: float, annual_rate: float, years: int):
    """等额本息：返回 (月供, 总利息, 还款总额)，单位均为元"""
    if principal <= 0 or years <= 0:
        return 0.0, 0.0, 0.0
    r = annual_rate / 100 / 12          # 月利率
    n = years * 12                       # 总期数
    if r == 0:
        m = principal / n
    else:
        m = principal * r * (1 + r) ** n / ((1 + r) ** n - 1)
    total = m * n
    interest = total - principal
    return m, interest, total


# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 基本设置")

    property_type = st.radio("房产类型", ["首套房", "二套房"], horizontal=True)
    loan_type = st.radio(
        "贷款方式",
        ["纯商业贷款", "纯公积金贷款", "组合贷款"],
    )
    years = st.selectbox(
        "贷款年限", [5, 10, 15, 20, 25, 30],
        index=5,  # 默认 30 年
    )

# ── 利率 & 规则 ────────────────────────────────────────────────────────────
gjj_rate = 2.6                                 # 公积金利率不变
sd_rate = 3.0 if property_type == "首套房" else 3.3   # 二套房商贷 3.3%

uses_gjj = loan_type in ("纯公积金贷款", "组合贷款")
min_down_pay = 20 if uses_gjj else 15           # 公积金≥20%，纯商贷≥15%

# ── 利率提示 ───────────────────────────────────────────────────────────────
st.markdown(
    f"**公积金利率** `{gjj_rate}%` ｜ "
    f"**商贷利率** `{sd_rate}%` ｜ "
    f"**最低首付** `{min_down_pay}%`"
)

# ── 输入区 ─────────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    total_price_wan = st.number_input(
        "房屋总价（万元）", min_value=1.0, value=200.0, step=1.0, format="%.0f",
    )
with col2:
    down_pay_pct = st.slider(
        "首付比例（%）",
        min_value=min_down_pay, max_value=100, value=min_down_pay, step=1,
        help=f"最低 {min_down_pay}%",
    )

max_gjj_wan: float = 0.0
if uses_gjj:
    max_gjj_wan = st.number_input(
        "公积金最高贷款额度（万元）",
        min_value=0.0, value=80.0, step=1.0, format="%.0f",
        help="各地公积金中心规定的最高可贷额度，可自行修改",
    )

# ── 计算 ───────────────────────────────────────────────────────────────────
total_price   = total_price_wan * 10000          # → 元
down_payment  = total_price * down_pay_pct / 100
total_loan    = total_price - down_payment
max_gjj       = max_gjj_wan * 10000

# 分配贷款
if loan_type == "纯公积金贷款":
    gjj_loan = min(total_loan, max_gjj)
    sd_loan  = 0.0
    if total_loan > max_gjj and max_gjj > 0:
        st.warning(
            f"所需贷款 **{total_loan/10000:.2f} 万元** 超出公积金最高额度 "
            f"**{max_gjj/10000:.0f} 万元**，请增加首付或切换为组合贷"
        )
elif loan_type == "纯商业贷款":
    gjj_loan = 0.0
    sd_loan  = total_loan
else:  # 组合贷
    gjj_loan = min(total_loan, max_gjj)
    sd_loan  = total_loan - gjj_loan

# 分别计算月供
gjj_m, gjj_interest, gjj_total = monthly_payment(gjj_loan, gjj_rate, years)
sd_m,  sd_interest,  sd_total  = monthly_payment(sd_loan,  sd_rate,  years)

# ── 结果展示 ───────────────────────────────────────────────────────────────
st.divider()
st.header("📊 计算结果")

# 首付
st.metric("首付金额", f"{down_payment/10000:,.2f} 万元")

st.divider()

# 公积金部分
show_gjj = loan_type in ("纯公积金贷款", "组合贷款")
show_sd  = loan_type in ("纯商业贷款", "组合贷款")

if show_gjj:
    st.subheader("🔵 公积金贷款")
    gc1, gc2, gc3, gc4 = st.columns(4)
    gc1.metric("贷款总额", f"{gjj_loan/10000:,.2f} 万元")
    gc2.metric("每月还款", f"{gjj_m:,.0f} 元")
    gc3.metric("总利息", f"{gjj_interest/10000:,.2f} 万元")
    gc4.metric("还款总额", f"{gjj_total/10000:,.2f} 万元")

if show_sd:
    st.subheader("🟠 商业贷款")
    sc1, sc2, sc3, sc4 = st.columns(4)
    sc1.metric("贷款总额", f"{sd_loan/10000:,.2f} 万元")
    sc2.metric("每月还款", f"{sd_m:,.0f} 元")
    sc3.metric("总利息", f"{sd_interest/10000:,.2f} 万元")
    sc4.metric("还款总额", f"{sd_total/10000:,.2f} 万元")

# 组合贷合计
if loan_type == "组合贷款":
    st.divider()
    st.subheader("📋 合计汇总")
    tc1, tc2, tc3, tc4 = st.columns(4)
    tc1.metric("总贷款额", f"{(gjj_loan + sd_loan)/10000:,.2f} 万元")
    tc2.metric("合计月供", f"{gjj_m + sd_m:,.0f} 元")
    tc3.metric("合计总利息", f"{(gjj_interest + sd_interest)/10000:,.2f} 万元")
    tc4.metric("合计还款总额", f"{(gjj_total + sd_total)/10000:,.2f} 万元")

# 脚注
st.divider()
st.caption("💡 采用**等额本息**还款方式计算，结果仅供参考，实际以银行审批为准。")
