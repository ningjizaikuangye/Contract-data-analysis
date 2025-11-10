import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import plotly.graph_objects as go
import matplotlib as mpl
from matplotlib import font_manager
import plotly.io as pio
import datetime
import os
import base64
import tempfile

# ===== 终极字体解决方案 =====
def setup_chinese_font():
    """确保中文显示的终极方案"""
    try:
        # 1. 尝试使用系统字体
        system_fonts = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS', 
                       'WenQuanYi Micro Hei', 'STHeiti', 'PingFang SC']
        
        # 查找可用字体
        available_font = None
        for font in system_fonts:
            try:
                font_path = font_manager.findfont(font)
                if font_path:
                    available_font = font
                    break
            except:
                continue
        
        # 2. 如果找到系统字体则使用
        if available_font:
            plt.rcParams['font.family'] = available_font
            plt.rcParams['axes.unicode_minus'] = False
            pio.templates.default = "plotly_white"
            pio.templates["plotly_white"].layout.font.family = available_font
            st.success(f"使用系统字体: {available_font}")
            return True
        
        # 3. 系统字体不可用时，使用Web安全字体回退
        plt.rcParams['font.family'] = ['sans-serif']
        plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'Microsoft YaHei', 'SimSun']
        plt.rcParams['axes.unicode_minus'] = False
        
        # 4. 强制设置Plotly使用相同字体
        pio.templates.default = "plotly_white"
        pio.templates["plotly_white"].layout.font.family = "Arial Unicode MS, Microsoft YaHei, sans-serif"
        
        return True
        
    except Exception as e:
        st.error(f"字体设置错误: {str(e)}")
        return False

# 初始化字体
setup_chinese_font()

# ===== 主应用代码 =====
def main():
    st.set_page_config(page_title="分包合同数据分析系统", layout="wide")
    st.title("📊 分包合同数据分析系统")
    
    # 1. 数据加载
    @st.cache_data
    def load_data():
        try:
            df = pd.read_excel("合同2.0系统数据.xlsm", sheet_name="Items", engine='openpyxl')
            
            # 数据清洗
            date_cols = [c for c in ['签订时间', '履行期限(起)', '履行期限(止)'] if c in df.columns]
            for col in date_cols:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            
            if '标的金额' in df.columns:
                df['标的金额'] = pd.to_numeric(df['标的金额'], errors='coerce')
                df['标的金额(万元)'] = df['标的金额'] / 10000
            
            return df
        except Exception as e:
            st.error(f"数据加载失败: {str(e)}")
            return None
    
    df = load_data()
    if df is None:
        st.stop()
    
    # 2. 筛选界面
    with st.sidebar:
        st.header("🔍 数据筛选")
        
        # 时间范围
        min_date = df['签订时间'].min().to_pydatetime()
        max_date = df['签订时间'].max().to_pydatetime()
        start_date, end_date = st.date_input(
            "合同签订时间范围", 
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )
        
        # 金额范围
        min_amount = st.number_input("最小金额(万元)", 
                                   value=float(df['标的金额(万元)'].min()), 
                                   min_value=0.0)
        max_amount = st.number_input("最大金额(万元)", 
                                   value=float(df['标的金额(万元)'].max()), 
                                   min_value=0.0)
        
        # 部门和采购类型
        departments = st.multiselect("承办部门", options=df['承办部门'].unique().tolist())
        procurement_types = st.multiselect("采购类型", options=df['选商方式'].unique().tolist())
    
    # 3. 应用筛选
    filtered_df = df[
        (df['签订时间'] >= pd.to_datetime(start_date)) & 
        (df['签订时间'] <= pd.to_datetime(end_date)) &
        (df['标的金额(万元)'] >= min_amount) &
        (df['标的金额(万元)'] <= max_amount)
    ]
    
    if departments:
        filtered_df = filtered_df[filtered_df['承办部门'].isin(departments)]
    if procurement_types:
        filtered_df = filtered_df[filtered_df['选商方式'].isin(procurement_types)]
    
    st.success(f"🔍 筛选到 {len(filtered_df)} 条记录")
    
    # 4. 数据分析展示
    tab1, tab2, tab3 = st.tabs(["数据表格", "统计分析", "图表展示"])
    
    with tab1:
        st.dataframe(filtered_df, height=400)
    
    with tab2:
        st.subheader("📈 基本统计信息")
        st.write(filtered_df.describe())
        
        st.subheader("🏢 部门合同统计")
        dept_stats = filtered_df.groupby('承办部门').agg(
            合同数量=('标的金额', 'count'),
            总金额_万元=('标的金额(万元)', 'sum'),
            平均金额_万元=('标的金额(万元)', 'mean')
        ).sort_values('总金额_万元', ascending=False)
        st.dataframe(dept_stats.style.format("{:.2f}"))
    
    with tab3:
        st.subheader("📊 采购类型分析")
        
        # 确保使用正确字体创建图表
        plt.rcParams['font.family'] = plt.rcParams['font.family'][0] if isinstance(plt.rcParams['font.family'], list) else plt.rcParams['font.family']
        
        # 采购类型分布
        fig1, ax1 = plt.subplots(figsize=(10, 6))
        counts = filtered_df['选商方式'].value_counts()
        counts.plot(kind='bar', ax=ax1, color='#1f77b4')
        
        ax1.set_title("各采购类型合同数量", fontsize=14, fontproperties=FontProperties(family=plt.rcParams['font.family']))
        ax1.set_xlabel("采购类型", fontsize=12, fontproperties=FontProperties(family=plt.rcParams['font.family']))
        ax1.set_ylabel("合同数量", fontsize=12, fontproperties=FontProperties(family=plt.rcParams['font.family']))
        
        # 旋转标签避免重叠
        plt.xticks(rotation=45, ha='right')
        
        # 添加数值标签
        for i, v in enumerate(counts):
            ax1.text(i, v, str(v), ha='center', va='bottom', fontsize=10)
        
        st.pyplot(fig1)
        
        # 金额分布
        st.subheader("💰 合同金额分布")
        fig2, ax2 = plt.subplots(figsize=(10, 6))
        amounts = filtered_df.groupby('选商方式')['标的金额(万元)'].sum().sort_values(ascending=False)
        amounts.plot(kind='bar', ax=ax2, color='#2ca02c')
        
        ax2.set_title("各采购类型合同金额(万元)", fontsize=14, fontproperties=FontProperties(family=plt.rcParams['font.family']))
        ax2.set_xlabel("采购类型", fontsize=12, fontproperties=FontProperties(family=plt.rcParams['font.family']))
        ax2.set_ylabel("金额(万元)", fontsize=12, fontproperties=FontProperties(family=plt.rcParams['font.family']))
        
        plt.xticks(rotation=45, ha='right')
        
        for i, v in enumerate(amounts):
            ax2.text(i, v, f"{v:.2f}", ha='center', va='bottom', fontsize=10)
        
        st.pyplot(fig2)
        
        # 交互式图表
        st.subheader("📈 交互式分析")
        fig3 = go.Figure()
        
        # 添加数据
        for dept in filtered_df['承办部门'].unique()[:5]:  # 限制显示前5个部门
            dept_data = filtered_df[filtered_df['承办部门'] == dept]
            fig3.add_trace(go.Box(
                y=dept_data['标的金额(万元)'],
                name=dept,
                boxpoints='all',
                jitter=0.3,
                pointpos=-1.8,
                marker=dict(size=4),
                line=dict(width=1)
            ))
        
        # 更新布局
        fig3.update_layout(
            title='各部门合同金额分布(万元)',
            yaxis_title='金额(万元)',
            font=dict(
                family=plt.rcParams['font.family'],
                size=12
            ),
            showlegend=True,
            height=500
        )
        
        st.plotly_chart(fig3, use_container_width=True)
    
    # 5. 数据导出
    st.subheader("💾 数据导出")
    csv = filtered_df.to_csv(index=False).encode('utf-8-sig')  # 使用utf-8-sig确保Excel能正确识别
    
    st.download_button(
        label="下载筛选结果(CSV)",
        data=csv,
        file_name=f"合同数据_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime='text/csv'
    )

if __name__ == "__main__":
    main()
