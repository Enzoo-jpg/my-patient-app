import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# 1. 网页标题与样式设置
st.set_page_config(page_title="患者脱落分析工具", layout="wide")
st.title("📊 患者用药生命周期与脱落动态分析工具")
st.markdown("---")

# 2. 上传文件组件
uploaded_file = st.file_uploader("👉 第一步：请上传您的原始购药 Excel 报表（需包含：患者姓名、1月-5月购药量）", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # 读取用户上传的 Excel
        df = pd.read_excel(uploaded_file)
        
        # 检查必要的列是否存在
        required_cols = ["患者姓名", "1月购药量", "2月购药量", "3月购药量", "4月购药量", "5月购药量"]
        if not all(col in df.columns for col in required_cols):
            st.error("❌ 上传的表格列名不规范！请确保包含：'患者姓名', '1月购药量', '2月购药量', '3月购药量', '4月购药量', '5月购药量'")
        else:
            # --- 3. 核心算法逻辑：滚动计算库存（不回补） ---
            df["1月可用库存"] = df["1月购药量"]
            df["2月可用库存"] = df.apply(lambda r: max(0, r["1月可用库存"] - 1) + r["2月购药量"], axis=1)
            df["3月可用库存"] = df.apply(lambda r: max(0, r["2月可用库存"] - 1) + r["3月购药量"], axis=1)
            df["4月可用库存"] = df.apply(lambda r: max(0, r["3月可用库存"] - 1) + r["4月购药量"], axis=1)
            df["5月可用库存"] = df.apply(lambda r: max(0, r["4月可用库存"] - 1) + r["5月购药量"], axis=1)

            # --- 4. 核心算法逻辑：新增脱落判定（选项B） ---
            df["2月新增脱落"] = df.apply(lambda r: 1 if r["1月可用库存"] > 0 and r["2月可用库存"] == 0 else 0, axis=1)
            df["3月新增脱落"] = df.apply(lambda r: 1 if r["2月可用库存"] > 0 and r["3月可用库存"] == 0 else 0, axis=1)
            df["4月新增脱落"] = df.apply(lambda r: 1 if r["3月可用库存"] > 0 and r["4月可用库存"] == 0 else 0, axis=1)
            df["5月新增脱落"] = df.apply(lambda r: 1 if r["4月可用库存"] > 0 and r["5月可用库存"] == 0 else 0, axis=1)

            # --- 5. 核心算法逻辑：期末累计真实脱落 ---
            df["总购药量"] = df[["1月购药量", "2月购药量", "3月购药量", "4月购药量", "5月购药量"]].sum(axis=1)
            df["截止5月真实脱落"] = df.apply(lambda r: 1 if r["总购药量"] > 0 and r["5月可用库存"] == 0 else 0, axis=1)

            # --- 6. 看板指标计算 ---
            total_patients = len(df)
            active_end = len(df[df["5月可用库存"] > 0])
            total_dropped = df["截止5月真实脱落"].sum()
            cum_drop_rate = total_dropped / total_patients if total_patients > 0 else 0

            # --- 7. 前端页面渲染：核心 KPI 卡片 ---
            st.subheader("二、核心流失指标总览")
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("总监测分析患者数", f"{total_patients} 人")
            kpi2.metric("5月底维持治疗人数", f"{active_end} 人")
            kpi3.metric("累计真实脱落人数", f"{total_dropped} 人")
            kpi4.metric("整体累计脱落率", f"{cum_drop_rate:.1%}")
            
            st.markdown("---")

            # --- 8. 前端页面渲染：月度动态脱落趋势 ---
            st.subheader("三、月度动态新增脱落分析")
            
            # 计算各月期初在治与新增脱落
            beg_2 = len(df[df["1月可用库存"] > 0])
            drop_2 = df["2月新增脱落"].sum()
            rate_2 = drop_2 / beg_2 if beg_2 > 0 else 0

            beg_3 = len(df[df["2月可用库存"] > 0])
            drop_3 = df["3月新增脱落"].sum()
            rate_3 = drop_3 / beg_3 if beg_3 > 0 else 0

            beg_4 = len(df[df["3月可用库存"] > 0])
            drop_4 = df["4月新增脱落"].sum()
            rate_4 = drop_4 / beg_4 if beg_4 > 0 else 0

            beg_5 = len(df[df["4月可用库存"] > 0])
            drop_5 = df["5月新增脱落"].sum()
            rate_5 = drop_5 / beg_5 if beg_5 > 0 else 0

            # 准备图表数据
            months = ["2月", "3月", "4月", "5月"]
            rates = [rate_2 * 100, rate_3 * 100, rate_4 * 100, rate_5 * 100]

            # 绘制折线图
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=months, y=rates, mode='lines+markers', name='新增脱落率', line=dict(color='#E67E22', width=3)))
            fig.update_layout(title="各月动态新增脱落率趋势 (%)", xaxis_title="月份", yaxis_title="脱落率 (%)", height=400)
            
            chart_col, table_col = st.columns([3, 2])
            with chart_col:
                st.plotly_chart(fig, use_container_width=True)
            with table_col:
                summary_df = pd.DataFrame({
                    "月份": months,
                    "期初在治人数": [beg_2, beg_3, beg_4, beg_5],
                    "当月新增脱落": [drop_2, drop_3, drop_4, drop_5],
                    "当月动态脱落率": [f"{rate_2:.1%}", f"{rate_3:.1%}", f"{rate_4:.1%}", f"{rate_5:.1%}"]
                })
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

            st.markdown("---")

            # --- 9. 前端页面渲染：明细数据展示与下载 ---
            st.subheader("四、智能计算结果明细 (脱落患者已红色高亮)")
            
            # 格式化表格，让 1（脱落）看起来更显眼
            def highlight_dropped(val):
                if val == 1:
                    return 'background-color: #FADBD8; color: #922B21; font-weight: bold;'
                return ''
            
            display_cols = ["患者姓名", "1月购药量", "2月购药量", "3月购药量", "4月购药量", "5月购药量", 
                            "1月可用库存", "2月可用库存", "3月可用库存", "4月可用库存", "5月可用库存",
                            "2月新增脱落", "3月新增脱落", "4月新增脱落", "5月新增脱落", "截止5月真实脱落"]
            
            styled_df = df[display_cols].style.applymap(highlight_dropped, subset=["2月新增脱落", "3月新增脱落", "4月新增脱落", "5月新增脱落", "截止5月真实脱落"])
            st.dataframe(styled_df, use_container_width=True)

            # 下载按钮
            st.markdown("### 📥 下载分析报表")
            @st.cache_data
            def convert_df(input_df):
                return input_df.to_excel(index=False, engine='openpyxl')
            
            # 生成临时excel供下载
            import io
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df[display_cols].to_excel(writer, index=False)
            
            st.download_button(
                label="点击下载带有计算结果的 Excel 表格",
                data=buffer.getvalue(),
                file_name="患者脱落分析结果报表.xlsx",
                mime="application/vnd.ms-excel"
            )
    except Exception as e:
        st.error(f"处理文件时出错，请确保Excel格式正确。错误信息: {e}")