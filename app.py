import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 1. 网页基础配置
st.set_page_config(page_title="高级全链路脱落分析系统", layout="wide")
st.title("📊 医院/药房患者全链路下钻分析与流失预警系统")
st.markdown("---")

# 2. 侧边栏：数据上传与多维下钻筛选
with st.sidebar:
    st.header("📂 数据中心")
    uploaded_file = st.file_uploader("第一步：上传原始报表", type=["xlsx", "xls"])
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
            # 1. 滚动可用库存
            for i, m_num in enumerate(months_nums):
                stock_col = f"{m_num}月可用库存"
                buy_col = f"{m_num}月购药量"
                if i == 0:
                    df[stock_col] = df[buy_col].fillna(0)
                else:
                    prev_stock_col = f"{months_nums[i-1]}月可用库存"
                    df[stock_col] = df.apply(lambda r: max(0, r[prev_stock_col] - 1) + r[buy_col], axis=1)
            
            # 2. 每月新增脱落
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
            final_drop_col = f"截止{last_m}月真实脱落"
            df[final_drop_col] = df.apply(lambda r: 1 if r["总购药量"] > 0 and r[last_stock_col] == 0 else 0, axis=1)
            
            # --- 侧边栏交互：动态下钻控制逻辑 ---
            pharmacies = ["全部药房"] + list(df["所属药房"].dropna().unique())
            
            with st.sidebar:
                selected_pharmacy = st.selectbox("1️⃣ 按药房下钻", pharmacies)
                status_filter = st.selectbox("2️⃣ 按患者状态过滤", ["显示所有人", "仅看已脱落患者", "仅看在治(活跃)患者"])
                
                # 如果选择看脱落患者，可以进一步下钻到具体月份
                specific_month = "全部月份"
                if status_filter == "仅看已脱落患者":
                    specific_month = st.selectbox("3️⃣ 查看哪个月新增的脱落？", ["全部月份"] + list(drop_cols_map.keys()) + [f"截止{last_m}月最终脱落"])

            # 开始执行数据切片（下钻）
            view_df = df.copy()
            
            # 药房过滤
            if selected_pharmacy != "全部药房":
                view_df = view_df[view_df["所属药房"] == selected_pharmacy]
            
            # 状态过滤
            if status_filter == "仅看已脱落患者":
                if specific_month == "全部月份":
                    # 只要任何一个月脱落过，或者最终脱落了
                    all_drop_fields = list(drop_cols_map.values()) + [final_drop_col]
                    view_df = view_df[view_df[all_drop_fields].sum(axis=1) > 0]
                elif specific_month in drop_cols_map:
                    view_df = view_df[view_df[drop_cols_map[specific_month]] == 1]
                else:
                    view_df = view_df[view_df[final_drop_col] == 1]
            elif status_filter == "仅看在治(活跃)患者":
                view_df = view_df[view_df[last_stock_col] > 0]
                
            # --- 前端大屏渲染 ---
            st.subheader(f"💯 核心流失指标总览 ({selected_pharmacy})")
            
            # 大盘数字基于当前选择的药房（不受状态过滤影响，保证全局观）
            base_df = df if selected_pharmacy == "全部药房" else df[df["所属药房"] == selected_pharmacy]
            total_patients = len(base_df)
            active_end = len(base_df[base_df[last_stock_col] > 0])
            total_dropped = base_df[final_drop_col].sum()
            cum_drop_rate = total_dropped / total_patients if total_patients > 0 else 0
            
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("总监测分析患者数", f"{total_patients} 人")
            kpi2.metric(f"{last_m}月底维持治疗人数", f"{active_end} 人")
            kpi3.metric("累计真实脱落人数", f"{total_dropped} 人")
            kpi4.metric("整体累计脱落率", f"{cum_drop_rate:.1%}")
            
            st.markdown("---")
            
            # --- 月度动态趋势图表 ---
            st.subheader("📈 月度动态新增脱落趋势分析")
            trend_months, trend_rates, beg_counts, drop_counts = [], [], [], []
            for i in range(1, len(months_nums)):
                curr_m = months_nums[i]
                prev_m = months_nums[i-1]
                beg_in_treatment = len(base_df[base_df[f"{prev_m}月可用库存"] > 0])
                new_dropped = base_df[f"{curr_m}月新增脱落"].sum()
                m_rate = new_dropped / beg_in_treatment if beg_in_treatment > 0 else 0
                
                trend_months.append(f"{curr_m}月")
                trend_rates.append(m_rate * 100)
                beg_counts.append(beg_in_treatment)
                drop_counts.append(new_dropped)
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=trend_months, y=trend_rates, mode='lines+markers+text', 
                text=[f"{r:.1%}" for r in [x/100 for x in trend_rates]], textposition="top center",
                name='新增脱落率', line=dict(color='#E67E22', width=3)
            ))
            fig.update_layout(title=f"{selected_pharmacy} - 各月动态新增脱落率走势", height=350, yaxis=dict(ticksuffix="%"))
            
            chart_col, table_col = st.columns([3, 2])
            with chart_col:
                st.plotly_chart(fig, use_container_width=True)
            with table_col:
                summary_df = pd.DataFrame({
                    "月份": trend_months, "期初在治人数": beg_counts, "当月新增脱落": drop_counts,
                    "当月动态脱落率": [f"{x/100:.1%}" for x in trend_rates]
                })
                st.dataframe(summary_df, use_container_width=True, hide_index=True)
                
            st.markdown("---")
            
            # --- 下钻明细数据与双表格展现 ---
            st.subheader("📋 深度下钻明细数据中心")
            st.info(f"💡 当前视图已联动过滤。药房：【{selected_pharmacy}】 | 状态