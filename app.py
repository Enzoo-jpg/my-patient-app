import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# 1. 网页基础配置
st.set_page_config(page_title="动态患者脱落分析系统", layout="wide")
st.title("📊 医院/药房患者用药生命周期与动态脱落分析系统")
st.markdown("---")

# 2. 侧边栏：上传文件与筛选条件
with st.sidebar:
    st.header("📂 数据中心")
    uploaded_file = st.file_uploader("第一步：上传原始报表", type=["xlsx", "xls"])
    st.markdown("""
    **💡 Excel 格式要求：**
    1. 必须包含 **'患者姓名'** 和 **'所属药房'** 列。
    2. 购药量列名必须按规则命名，例如：`1月购药量`, `2月购药量` ... `12月购药量`。
    """)

if uploaded_file:
    try:
        # 读取数据
        df_raw = pd.read_excel(uploaded_file)
        
        # 自动识别表格中到底包含哪些月份的购药量
        buy_cols = [col for col in df_raw.columns if "购药量" in col and col != "总购药量"]
        # 按月份数字大小排序，确保 2月在10月前面
        buy_cols.sort(key=lambda x: int(''.join(filter(str.isdigit, x))))
        
        # 提取纯月份数字列表 (如 [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])
        months_nums = [int(''.join(filter(str.isdigit, c))) for c in buy_cols]
        
        # 基础列校验
        if "患者姓名" not in df_raw.columns or "所属药房" not in df_raw.columns:
            st.error("❌ 格式错误：表格中必须包含 '患者姓名' 和 '所属药房' 这两列！")
        elif len(buy_cols) < 2:
            st.error("❌ 格式错误：未检测到足够的购药量列（至少需要包含连续两个月的数据）。")
        else:
            
            # --- 核心计算引擎 (利用动态循环，自动支持 1-12 月) ---
            df = df_raw.copy()
            
            # 1. 动态滚动计算每月可用库存 (不回补逻辑)
            for i, m_num in enumerate(months_nums):
                stock_col = f"{m_num}月可用库存"
                buy_col = f"{m_num}月购药量"
                
                if i == 0:
                    # 首月库存 = 首月购药量
                    df[stock_col] = df[buy_col].fillna(0)
                else:
                    # 本月库存 = MAX(0, 上月库存 - 1) + 本月购药量
                    prev_stock_col = f"{months_nums[i-1]}月可用库存"
                    df[stock_col] = df.apply(lambda r: max(0, r[prev_stock_col] - 1) + r[buy_col], axis=1)
            
            # 2. 动态计算每月新增脱落 (选项B：断药不重复计入)
            for i in range(1, len(months_nums)):
                curr_m = months_nums[i]
                prev_m = months_nums[i-1]
                drop_col = f"{curr_m}月新增脱落"
                
                curr_stock = f"{curr_m}月可用库存"
                prev_stock = f"{prev_m}月可用库存"
                
                # 判定：上月手里有药(>0)且本月手里没药(==0)，记为 1，否则为 0
                df[drop_col] = df.apply(lambda r: 1 if r[prev_stock] > 0 and r[curr_stock] == 0 else 0, axis=1)
            
            # 3. 动态结算期末留存与真实脱落
            last_m = months_nums[-1]  # 报表里的最后一个月 (可能是5月，也可能是12月)
            last_stock_col = f"{last_m}月可用库存"
            
            df["总购药量"] = df[buy_cols].sum(axis=1)
            final_drop_col = f"截止{last_m}月真实脱落"
            
            # 期末动态看留存：今年买过药，但最后一个月手里库存为 0 的人
            df[final_drop_col] = df.apply(lambda r: 1 if r["总购药量"] > 0 and r[last_stock_col] == 0 else 0, axis=1)
            
            # --- 药房筛选联动功能 ---
            pharmacies = ["全部药房"] + list(df["所属药房"].dropna().unique())
            with st.sidebar:
                st.markdown("---")
                st.header("🎯 维度筛选")
                selected_pharmacy = st.selectbox("选择要分析的药房", pharmacies)
            
            # 根据选择切分数据
            if selected_pharmacy != "全部药房":
                view_df = df[df["所属药房"] == selected_pharmacy].copy()
            else:
                view_df = df.copy()
                
            # --- 前端大屏渲染 ---
            st.subheader(f"💯 核心流失指标总览 ({selected_pharmacy})")
            
            total_patients = len(view_df)
            active_end = len(view_df[view_df[last_stock_col] > 0])
            total_dropped = view_df[final_drop_col].sum()
            cum_drop_rate = total_dropped / total_patients if total_patients > 0 else 0
            
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("总监测分析患者数", f"{total_patients} 人")
            kpi2.metric(f"{last_m}月底维持治疗人数", f"{active_end} 人")
            kpi3.metric("累计真实脱落人数", f"{total_dropped} 人")
            kpi4.metric("整体累计脱落率", f"{cum_drop_rate:.1%}")
            
            st.markdown("---")
            
            # --- 月度动态趋势 (全自动适应月份数量) ---
            st.subheader("📈 月度动态新增脱落趋势分析")
            
            trend_months = []
            trend_rates = []
            beg_counts = []
            drop_counts = []
            
            for i in range(1, len(months_nums)):
                curr_m = months_nums[i]
                prev_m = months_nums[i-1]
                
                prev_stock = f"{prev_m}月可用库存"
                curr_drop = f"{curr_m}月新增脱落"
                
                beg_in_treatment = len(view_df[view_df[prev_stock] > 0])
                new_dropped = view_df[curr_drop].sum()
                m_rate = new_dropped / beg_in_treatment if beg_in_treatment > 0 else 0
                
                trend_months.append(f"{curr_m}月")
                trend_rates.append(m_rate * 100)
                beg_counts.append(beg_in_treatment)
                drop_counts.append(new_dropped)
            
            # 绘制 Plotly 动态图表
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=trend_months, y=trend_rates, 
                mode='lines+markers+text', 
                text=[f"{r:.1%}" for r in [x/100 for x in trend_rates]],
                textposition="top center",
                name='新增脱落率', 
                line=dict(color='#2980B9', width=3)
            ))
            fig.update_layout(
                title=f"{selected_pharmacy} - 各月动态新增脱落率趋势 (%)", 
                xaxis_title="分析月份", 
                yaxis_title="脱落率 (%)", 
                height=400,
                yaxis=dict(ticksuffix="%")
            )
            
            chart_col, table_col = st.columns([3, 2])
            with chart_col:
                st.plotly_chart(fig, use_container_width=True)
            with table_col:
                summary_df = pd.DataFrame({
                    "月份": trend_months,
                    "期初在治人数 (分母)": beg_counts,
                    "当月新增脱落 (分子)": drop_counts,
                    "当月动态脱落率": [f"{x/100:.1%}" for x in trend_rates]
                })
                st.dataframe(summary_df, use_container_width=True, hide_index=True)
                
            st.markdown("---")
            
            # --- 智能明细大表与变色高亮 ---
            st.subheader("📋 智能计算结果明细表 (可在右上角搜索/过滤)")
            
            # 自动组合高亮列
            drop_cols = [f"{m}月新增脱落" for m in months_nums[1:]] + [final_drop_col]
            stock_cols = [f"{m}月可用库存" for m in months_nums]
            display_cols = ["患者姓名", "所属药房"] + buy_cols + stock_cols + drop_cols
            
            # 变色样式函数
            def highlight_dropped(val):
                return 'background-color: #FADBD8; color: #922B21; font-weight: bold;' if val == 1 else ''
            
            styled_df = view_df[display_cols].style.applymap(highlight_dropped, subset=drop_cols)
            st.dataframe(styled_df, use_container_width=True)
            
            # --- 导出结果下载 ---
            st.markdown("### 📥 下载清洗后报表")
            import io
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                view_df[display_cols].to_excel(writer, index=False)
            
            st.download_button(
                label=f"点击下载【{selected_pharmacy}】的计算结果 Excel",
                data=buffer.getvalue(),
                file_name=f"{selected_pharmacy}_患者脱落分析报表.xlsx",
                mime="application/vnd.ms-excel"
            )
            
    except Exception as e:
        st.error(f"🚨 运行出错！请检查 Excel 格式是否规范。错误详情: {e}")
else:
    st.info("💡 期待您的数据：请在左侧侧边栏上传包含 '所属药房' 和购药数据的 Excel 文件。")