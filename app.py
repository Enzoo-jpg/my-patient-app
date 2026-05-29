import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 1. 网页基础配置与全局命名
st.set_page_config(page_title="脱落率计算工具", layout="wide")
st.title("📊 脱落率计算工具（医院/药房全链路下钻与流失预警系统）")
st.markdown("---")

# 2. 侧边栏：数据上传、标准模板下载与多维下钻筛选
with st.sidebar:
    st.header("📂 数据中心")
    
    # 提供标准 Excel 模板下载，避免用户格式传错
    st.markdown("##### ⬇️ 第一步：下载/核对标准格式")
    template_buffer = io.BytesIO()
    temp_wb = Workbook()
    temp_ws = temp_wb.active
    temp_ws.title = "标准格式示例"
    temp_ws.views.sheetView[0].showGridLines = True
    temp_ws.append(["患者姓名", "所属药房", "1月购药量", "2月购药量", "3月购药量"])
    temp_ws.append(["张三", "北京第一药房", 1, 0, 1])
    temp_ws.append(["李四", "上海中心药房", 2, 1, 0])
    temp_wb.save(template_buffer)
    
    st.download_button(
        label="📥 下载标准 Excel 表头模板",
        data=template_buffer.getvalue(),
        file_name="脱落率工具_标准表头模板.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.markdown(" ")
    
    # 上传文件
    uploaded_file = st.file_uploader("第二步：上传您填好的数据报表", type=["xlsx", "xls"])
    st.markdown("---")
    st.header("🎯 多维下钻漏斗")
    
if uploaded_file:
    try:
        # 读取数据
        df_raw = pd.read_excel(uploaded_file)
        
        # 自动识别月份购药量列并排序
        buy_cols = [col for col in df_raw.columns if "购药量" in col and col != "总购药量"]
        buy_cols.sort(key=lambda x: int(''.join(filter(str.isdigit, x))))
        months_nums = [int(''.join(filter(str.isdigit, c))) for c in buy_cols]
        
        if "患者姓名" not in df_raw.columns or "所属药房" not in df_raw.columns:
            st.error("❌ 格式错误：表格中必须包含 '患者姓名' 和 '所属药房' 这两列！")
        elif len(buy_cols) < 2:
            st.error("❌ 格式错误：未检测到足够的购药量列（至少需要连续两个月的数据）。")
        else:
            df = df_raw.copy()
            
            # --- 核心计算引擎 ---
            # 1. 滚动可用库存（按标准业务规则：患者每月消耗 1 盒）
            for i, m_num in enumerate(months_nums):
                stock_col = f"{m_num}月可用库存"
                buy_col = f"{m_num}月购药量"
                if i == 0:
                    df[stock_col] = df[buy_col].fillna(0)
                else:
                    prev_stock_col = f"{months_nums[i-1]}月可用库存"
                    df[stock_col] = df.apply(lambda r: max(0, r[prev_stock_col] - 1) + r[buy_col], axis=1)
            
            # 2. 每月新增脱落判定（上月本店有药可用 且 本月无药可用）
            drop_cols_map = {}
            for i in range(1, len(months_nums)):
                curr_m = months_nums[i]
                prev_m = months_nums[i-1]
                d_col = f"{curr_m}月新增脱落"
                df[d_col] = df.apply(
                    lambda r: 1 if r[f"{prev_m}月可用库存"] > 0 and r[f"{curr_m}月可用库存"] == 0 else 0, axis=1
                )
                drop_cols_map[f"{curr_m}月"] = d_col
            
            # 3. 期末结算
            last_m = months_nums[-1]
            last_stock_col = f"{last_m}月可用库存"
            df["总购药量"] = df[buy_cols].sum(axis=1)
            final_drop_col = f"截止{last_m}月最终脱落"
            df[final_drop_col] = df.apply(lambda r: 1 if r["总购药量"] > 0 and r[last_stock_col] == 0 else 0, axis=1)
            
            # --- 侧边栏交互：动态下钻控制逻辑 ---
            all_pharmacies_list = list(df["所属药房"].dropna().unique())
            
            with st.sidebar:
                selected_pharmacies = st.multiselect(
                    "1️⃣ 按药房下钻 (可多选)", 
                    options=all_pharmacies_list,
                    default=all_pharmacies_list
                )
                
                status_filter = st.selectbox("2️⃣ 按患者状态过滤", ["显示所有人", "仅看已脱落患者", "仅看在治(活跃)患者"])
                
                specific_month = "全部月份"
                if status_filter == "仅看已脱落患者":
                    specific_month = st.selectbox("3️⃣ 查看哪个月新增的脱落？", ["全部月份"] + list(drop_cols_map.keys()) + [f"截止{last_m}月最终脱落"])

            # 开始执行数据切片（下钻）
            view_df = df.copy()
            
            # 药房多选过滤逻辑
            if selected_pharmacies:
                view_df = view_df[view_df["所属药房"].isin(selected_pharmacies)]
                base_df = df[df["所属药房"].isin(selected_pharmacies)].copy()
                pharmacy_label = ", ".join(selected_pharmacies)
            else:
                base_df = df.copy()
                pharmacy_label = "全部药房 (未选择时默认全选)"
            
            # 状态过滤
            all_drop_fields = list(drop_cols_map.values()) + [final_drop_col]
            if status_filter == "仅看已脱落患者":
                if specific_month == "全部月份":
                    view_df = view_df[view_df[all_drop_fields].sum(axis=1) > 0]
                elif specific_month in drop_cols_map:
                    view_df = view_df[view_df[drop_cols_map[specific_month]] == 1]
                else:
                    view_df = view_df[view_df[final_drop_col] == 1]
            elif status_filter == "仅看在治(活跃)患者":
                view_df = view_df[view_df[last_stock_col] > 0]
                
            # --- 前端大屏渲染 ---
            st.subheader(f"💯 核心流失指标总览 ({pharmacy_label})")
            
            total_patients = len(base_df)
            active_end = len(base_df[base_df[last_stock_col] > 0])
            total_dropped = base_df[all_drop_fields].sum(axis=1).apply(lambda x: 1 if x > 0 else 0).sum()
            cum_drop_rate = total_dropped / total_patients if total_patients > 0 else 0
            
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("总监测分析患者数", f"{total_patients} 人")
            kpi2.metric(f"{last_m}月底维持治疗人数", f"{active_end} 人")
            kpi3.metric("累计真实脱落人数", f"{total_dropped} 人")
            kpi4.metric("整体累计脱落率", f"{cum_drop_rate:.1%}")
            
            st.markdown("---")
            
            # --- 月度趋势图表 ---
            st.subheader("📈 月度脱落率与实际购药表现趋势分析")
            st.caption("💡 脱落率精准定义：[上月本店有药可用且本月无药可用用户数] / [上月本店有药可用用户数]")
            
            trend_months, trend_rates, beg_counts, drop_counts, buy_counts = [], [], [], [] ,[]
            
            first_m = months_nums[0]
            first_m_buyers = len(base_df[base_df[f"{first_m}月购药量"] > 0])
            
            for i in range(1, len(months_nums)):
                curr_m = months_nums[i]
                prev_m = months_nums[i-1]
                
                beg_in_treatment = len(base_df[base_df[f"{prev_m}月可用库存"] > 0])
                new_dropped = base_df[f"{curr_m}月新增脱落"].sum()
                m_rate = new_dropped / beg_in_treatment if beg_in_treatment > 0 else 0
                current_buyers = len(base_df[base_df[f"{curr_m}月购药量"] > 0])
                
                trend_months.append(f"{curr_m}月")
                trend_rates.append(m_rate * 100)
                beg_counts.append(beg_in_treatment)
                drop_counts.append(new_dropped)
                buy_counts.append(current_buyers)
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=trend_months, y=trend_rates, mode='lines+markers+text', 
                text=[f"{r:.1%}" for r in [x/100 for x in trend_rates]], textposition="top center",
                name='月度脱落率', line=dict(color='#E67E22', width=3)
            ))
            fig.update_layout(title=f"所选药房 - 月度脱落率走势图", height=350, yaxis=dict(ticksuffix="%"))
            
            chart_col, table_col = st.columns([3, 2])
            with chart_col:
                st.plotly_chart(fig, use_container_width=True)
            with table_col:
                table_months = [f"{first_m}月"] + trend_months
                table_beg = ["-"] + beg_counts
                table_drop = ["-"] + drop_counts
                table_rates = ["-"] + [f"{x/100:.1%}" for x in trend_rates]
                table_buyers = [first_m_buyers] + buy_counts
                
                summary_df = pd.DataFrame({
                    "月份": table_months, 
                    "当月购药患者数": table_buyers,
                    "上月有药可用用户数": table_beg, 
                    "上月有药且本月无药用户数": table_drop,
                    "月度脱落率": table_rates
                })
                st.dataframe(summary_df, use_container_width=True, hide_index=True)
                
            st.markdown("---")
            
            # --- 下钻明细数据与双
