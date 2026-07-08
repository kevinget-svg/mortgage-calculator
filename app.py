import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ────────────────────────────────────────────────────────────
st.set_page_config(page_title="房贷计算器", page_icon="🏠", layout="wide")

# ── Constants ──────────────────────────────────────────────────────────────
PRICE_OPTIONS = [50, 80, 100, 120, 150, 180, 200, 220, 250, 280, 300, 350, 400, 450, 500]
PRICE_DEFAULTS = [150, 200, 250]
YEAR_OPTIONS = [5, 10, 15, 20, 25, 30]

# ── Helper ─────────────────────────────────────────────────────────────────
def monthly_payment(principal: float, annual_rate: float, years: int):
    """等额本息：返回 (月供, 总利息, 还款总额)，单位均为元"""
    if principal <= 0 or years <= 0:
        return 0.0, 0.0, 0.0
    r = annual_rate / 100 / 12
    n = years * 12
    if r == 0:
        m = principal / n
    else:
        m = principal * r * (1 + r) ** n / ((1 + r) ** n - 1)
    total = m * n
    interest = total - principal
    return m, interest, total


def make_scenario_label(price: float, dp: int):
    """生成场景标签，如 '200万/20%'"""
    return f"{price:.0f}万/{dp}%"


def build_excel(df: pd.DataFrame, loan_type: str, years: int) -> bytes:
    """用 openpyxl 生成带格式的 Excel 文件，返回 bytes"""
    wb = Workbook()
    ws = wb.active
    ws.title = "贷款对比结果"

    # 标题行
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
    title_cell = ws.cell(row=1, column=1, value=f"房贷计算结果 — {loan_type}，{years}年")
    title_cell.font = Font(name="微软雅黑", size=14, bold=True)
    title_cell.alignment = Alignment(horizontal="center")

    # 空行
    ws.append([])

    # 表头
    header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    header_font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=3, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # 数据行
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    for row_idx, row in df.iterrows():
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row=row_idx + 4, column=col_idx, value=val)
            cell.font = Font(name="微软雅黑", size=10)
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border

    # 自适应列宽
    for col_idx in range(1, len(df.columns) + 1):
        max_len = max(
            len(str(df.columns[col_idx - 1])),
            df.iloc[:, col_idx - 1].astype(str).str.len().max() if len(df) > 0 else 0,
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = max(max_len + 4, 12)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 基本设置")

    property_type = st.radio("房产类型", ["首套房", "二套房"], horizontal=True)
    loan_type = st.radio(
        "贷款方式",
        ["纯商业贷款", "纯公积金贷款", "组合贷款"],
    )
    years = st.selectbox("贷款年限", YEAR_OPTIONS, index=5)

# ── 利率 & 规则 ────────────────────────────────────────────────────────────
gjj_rate = 2.6
sd_rate = 3.0 if property_type == "首套房" else 3.3

uses_gjj = loan_type in ("纯公积金贷款", "组合贷款")
min_down_pay = 20 if uses_gjj else 15

st.markdown(
    f"**公积金利率** `{gjj_rate}%` ｜ "
    f"**商贷利率** `{sd_rate}%` ｜ "
    f"**最低首付** `{min_down_pay}%`"
)

# ── 多选输入区 ─────────────────────────────────────────────────────────────
st.subheader("📝 选择对比场景")
col1, col2, col3 = st.columns(3)

with col1:
    selected_prices = st.multiselect(
        "房屋总价（万元）",
        options=PRICE_OPTIONS,
        default=[p for p in PRICE_DEFAULTS if p in PRICE_OPTIONS],
        help="可多选，同时对比不同总价",
    )

with col2:
    dp_pcts = list(range(min_down_pay, 55, 5))
    default_dp = [min_down_pay, min_down_pay + 5, min_down_pay + 10]
    default_dp = [d for d in default_dp if d in dp_pcts]
    selected_dps = st.multiselect(
        "首付比例（%）",
        options=dp_pcts,
        default=default_dp[:3],
        help="可多选，同时对比不同首付比例",
    )

with col3:
    max_gjj_wan = 80.0
    if uses_gjj:
        max_gjj_wan = st.number_input(
            "公积金最高贷款额度（万元）",
            min_value=0.0, value=80.0, step=1.0, format="%.0f",
            help="各地公积金中心规定的最高可贷额度",
        )

if not selected_prices or not selected_dps:
    st.info("👆 请在上述多选框中至少各选一项，即可查看计算结果。")
    st.stop()

# ── 计算所有组合 ───────────────────────────────────────────────────────────
rows = []
max_gjj = max_gjj_wan * 10000

for price_wan in selected_prices:
    total_price = price_wan * 10000
    for dp_pct in selected_dps:
        down_payment = total_price * dp_pct / 100
        total_loan = total_price - down_payment

        if loan_type == "纯公积金贷款":
            gjj_loan = min(total_loan, max_gjj) if max_gjj > 0 else total_loan
            sd_loan = 0.0
            gjj_note = "⚠已达上限" if (total_loan > max_gjj and max_gjj > 0) else ""
        elif loan_type == "纯商业贷款":
            gjj_loan = 0.0
            sd_loan = total_loan
            gjj_note = ""
        else:  # 组合贷
            gjj_loan = min(total_loan, max_gjj) if max_gjj > 0 else 0
            sd_loan = total_loan - gjj_loan
            gjj_note = "⚠已达上限" if (max_gjj > 0 and gjj_loan >= max_gjj) else ""

        gjj_m, gjj_interest, gjj_total = monthly_payment(gjj_loan, gjj_rate, years)
        sd_m, sd_interest, sd_total = monthly_payment(sd_loan, sd_rate, years)

        row = {
            "场景": make_scenario_label(price_wan, dp_pct),
            "房屋总价(万)": price_wan,
            "首付比例": f"{dp_pct}%",
            "首付(万)": round(down_payment / 10000, 2),
            "公积金贷款(万)": f"{gjj_loan/10000:.2f}{gjj_note}" if gjj_loan > 0 else 0,
            "公积金月供(元)": round(gjj_m, 0) if gjj_loan > 0 else 0,
            "公积金总利息(万)": round(gjj_interest / 10000, 2) if gjj_loan > 0 else 0,
            "商贷(万)": round(sd_loan / 10000, 2) if sd_loan > 0 else 0,
            "商贷月供(元)": round(sd_m, 0) if sd_loan > 0 else 0,
            "商贷总利息(万)": round(sd_interest / 10000, 2) if sd_loan > 0 else 0,
            "合计月供(元)": round(gjj_m + sd_m, 0),
            "合计总利息(万)": round((gjj_interest + sd_interest) / 10000, 2),
        }
        rows.append(row)

df = pd.DataFrame(rows)

# ── 动态筛选展示列 ─────────────────────────────────────────────────────────
if loan_type == "纯商业贷款":
    display_cols = ["场景", "房屋总价(万)", "首付比例", "首付(万)",
                    "商贷(万)", "商贷月供(元)", "商贷总利息(万)", "合计月供(元)", "合计总利息(万)"]
elif loan_type == "纯公积金贷款":
    display_cols = ["场景", "房屋总价(万)", "首付比例", "首付(万)",
                    "公积金贷款(万)", "公积金月供(元)", "公积金总利息(万)", "合计月供(元)", "合计总利息(万)"]
else:
    display_cols = ["场景", "房屋总价(万)", "首付比例", "首付(万)",
                    "公积金贷款(万)", "公积金月供(元)", "公积金总利息(万)",
                    "商贷(万)", "商贷月供(元)", "商贷总利息(万)",
                    "合计月供(元)", "合计总利息(万)"]

df_display = df[display_cols].copy()

# ── 结果表格 ───────────────────────────────────────────────────────────────
st.divider()
st.header("📊 对比结果")

# 高亮最低月供行
def highlight_min(s):
    if s.name == "合计月供(元)" or s.name == "商贷月供(元)" or s.name == "公积金月供(元)":
        return ['background-color: #d4edda; font-weight: bold' if v == s.min() else '' for v in s]
    return ['' for _ in s]

styled = df_display.style.apply(highlight_min, axis=0)
st.dataframe(styled, use_container_width=True, hide_index=True)

# ── 可视化 ─────────────────────────────────────────────────────────────────
st.subheader("📈 图表对比")

chart_df = df.copy()
chart_df["场景"] = pd.Categorical(chart_df["场景"], categories=chart_df["场景"].tolist(), ordered=True)

col_a, col_b = st.columns(2)

with col_a:
    # 月供对比
    if loan_type == "组合贷款":
        fig1 = go.Figure()
        fig1.add_bar(name="公积金月供", x=chart_df["场景"], y=chart_df["公积金月供(元)"],
                     text=chart_df["公积金月供(元)"].apply(lambda x: f"{x:,.0f}"),
                     textposition="inside", marker_color="#E8B88A")
        fig1.add_bar(name="商贷月供", x=chart_df["场景"], y=chart_df["商贷月供(元)"],
                     text=chart_df["商贷月供(元)"].apply(lambda x: f"{x:,.0f}"),
                     textposition="inside", marker_color="#C5602D")
        fig1.update_layout(barmode="stack", title="月供构成对比（元）", height=400,
                           legend=dict(orientation="h", y=1.15))
    else:
        y_col = "商贷月供(元)" if loan_type == "纯商业贷款" else "公积金月供(元)"
        name = "商贷月供" if loan_type == "纯商业贷款" else "公积金月供"
        fig1 = px.bar(chart_df, x="场景", y=y_col, title=f"{name}对比（元）",
                       text=chart_df[y_col].apply(lambda x: f"{x:,.0f}"),
                       color="首付比例",
                       color_discrete_sequence=px.colors.qualitative.Set2)
        fig1.update_traces(textposition="outside")
        fig1.update_layout(height=400, legend=dict(orientation="h", y=1.15))
    st.plotly_chart(fig1, use_container_width=True)

with col_b:
    # 总利息对比
    if loan_type == "组合贷款":
        fig2 = go.Figure()
        fig2.add_bar(name="公积金利息", x=chart_df["场景"], y=chart_df["公积金总利息(万)"],
                     text=chart_df["公积金总利息(万)"].apply(lambda x: f"{x:.1f}万"),
                     textposition="inside", marker_color="#E8B88A")
        fig2.add_bar(name="商贷利息", x=chart_df["场景"], y=chart_df["商贷总利息(万)"],
                     text=chart_df["商贷总利息(万)"].apply(lambda x: f"{x:.1f}万"),
                     textposition="inside", marker_color="#C5602D")
        fig2.update_layout(barmode="stack", title="总利息构成对比（万元）", height=400,
                           legend=dict(orientation="h", y=1.15))
    else:
        y_col = "商贷总利息(万)" if loan_type == "纯商业贷款" else "公积金总利息(万)"
        name = "商贷总利息" if loan_type == "纯商业贷款" else "公积金总利息"
        fig2 = px.bar(chart_df, x="场景", y=y_col, title=f"{name}对比（万元）",
                       text=chart_df[y_col].apply(lambda x: f"{x:.1f}万"),
                       color="首付比例",
                       color_discrete_sequence=px.colors.qualitative.Set2)
        fig2.update_traces(textposition="outside")
        fig2.update_layout(height=400, legend=dict(orientation="h", y=1.15))
    st.plotly_chart(fig2, use_container_width=True)

# 月供趋势线（房子总价 vs 月供，按首付比例分组）
st.subheader("📉 月供趋势")
trend_df = df.copy()
trend_fig = px.line(
    trend_df, x="房屋总价(万)", y="合计月供(元)", color="首付比例",
    markers=True, title="不同首付比例下，月供随房价变化趋势",
    color_discrete_sequence=px.colors.qualitative.Set2,
)
trend_fig.update_traces(texttemplate=trend_df["合计月供(元)"].apply(lambda x: f"{x:,.0f}"),
                        textposition="top center")
trend_fig.update_layout(height=420)
st.plotly_chart(trend_fig, use_container_width=True)

# ── Excel 导出 ─────────────────────────────────────────────────────────────
st.divider()
st.subheader("📥 导出结果")

excel_bytes = build_excel(df_display, loan_type, years)
st.download_button(
    label="📥 导出 Excel 文件",
    data=excel_bytes,
    file_name=f"房贷计算结果_{loan_type}_{years}年.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ── 脚注 ───────────────────────────────────────────────────────────────────
st.divider()
st.caption("💡 采用**等额本息**还款方式计算，结果仅供参考，实际以银行审批为准。")
