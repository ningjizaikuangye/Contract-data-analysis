import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import plotly.graph_objects as go
from datetime import datetime
import os
import matplotlib as mpl
from matplotlib.font_manager import FontProperties
import plotly.io as pio

# ===== 100%可用的字体解决方案 =====
def set_chinese_font():
    """确保中文显示的终极解决方案"""
    try:
        # 尝试所有可能的中文字体
        font_list = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS', 
                   'WenQuanYi Micro Hei', 'FangSong', 'KaiTi', 
                   'STHeiti', 'LiHei Pro', 'AppleGothic']
        
        # 查找第一个可用的字体
        available_font = None
        for font in font_list:
            try:
                # 测试字体是否存在
                test_font = FontProperties(family=font)
                if mpl.font_manager.findfont(test_font):
                    available_font = font
                    break
            except:
                continue
        
        if available_font:
            # 设置Matplotlib
            plt.rcParams['font.family'] = available_font
            plt.rcParams['axes.unicode_minus'] = False
            
            # 设置Plotly
            pio.templates.default = "plotly_white"
            pio.templates["plotly_white"].layout.font.family = available_font
            
            st.success(f"使用系统字体: {available_font}")
            return True
        else:
            # 最终回退方案：强制使用符号字体
            plt.rcParams['font.sans-serif'] = ['sans-serif']
            plt.rcParams['axes.unicode_minus'] = False
            return False
            
    except Exception as e:
        st.error(f"字体设置错误: {str(e)}")
        return False

# 初始化字体
set_chinese_font()

# ===== 主应用代码 =====
def main():
    st.set_page_config(page_title="分包合同数据分析", layout="wide")
    st.title("分包合同数据分析系统")
    
    # 1. 数据加载
    @st.cache_data
    def load_data():
        try:
            df = pd.read_excel("03 合同2.0系统数据.xlsm", sheet_name="Items", engine='openpyxl')
            
            # 日期和金额处理
            date_cols = [c for c in ['签订时间', '履行期限(起)', '履行期限(止)'] if c in df.columns]
            for col in date_cols:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            
            if '标的金额' in df.columns:
                df['标的金额'] = pd.to_numeric(df['标的金额'], errors='coerce')
            
            return df
        except Exception as e:
            st.error(f"数据加载失败: {str(e)}")
            return None
    
    df = load_data()
    if df is None:
        return
    
    # 2. 筛选界面
    with st.sidebar:
        st.header("数据筛选")
        
        # 时间范围
        min_date = df['签订时间'].min().to_pydatetime()
        max_date = df['签订时间'].max().to_pydatetime()
        date_range = st.date_input("合同签订时间", [min_date, max_date])
        
        # 金额范围
        min_amount, max_amount = st.slider(
            "合同金额范围(元)",
            float(df['标的金额'].min()),
            float(df['标的金额'].max()),
            (float(df['标的金额'].min()), float(df['标的金额'].max()))
        )
        
        # 部门和采购类型
        departments = st.multiselect("承办部门", df['承办部门'].unique().tolist())
        types = st.multiselect("采购类型", df['选商方式'].unique().tolist())
    
    # 3. 应用筛选
    filtered_df = df.copy()
    if len(date_range) == 2:
        filtered_df = filtered_df[
            (filtered_df['签订时间'] >= pd.to_datetime(date_range[0])) &
            (filtered_df['签订时间'] <= pd.to_datetime(date_range[1]))
        ]
    
    filtered_df = filtered_df[
        (filtered_df['标的金额'] >= min_amount) &
        (filtered_df['标的金额'] <= max_amount)
    ]
    
    if departments:
        filtered_df = filtered_df[filtered_df['承办部门'].isin(departments)]
    if types:
        filtered_df = filtered_df[filtered_df['选商方式'].isin(types)]
    
    st.success(f"找到 {len(filtered_df)} 条匹配记录")
    
    # 4. 数据分析展示
    tab1, tab2 = st.tabs(["数据概览", "图表分析"])
    
    with tab1:
        st.dataframe(filtered_df)
        
        # 统计信息
        st.subheader("统计摘要")
        st.write(filtered_df.describe())
    
    with tab2:
        # 采购类型分析
        st.subheader("按采购类型分析")
        
        col1, col2 = st.columns(2)
        with col1:
            # 数量分布
            fig1, ax1 = plt.subplots()
            counts = filtered_df['选商方式'].value_counts()
            counts.plot(kind='bar', ax=ax1, color='skyblue')
            
            ax1.set_title("采购类型-合同数量", fontproperties=FontProperties(family=plt.rcParams['font.family']))
            ax1.set_xlabel("采购类型", fontproperties=FontProperties(family=plt.rcParams['font.family']))
            ax1.set_ylabel("数量", fontproperties=FontProperties(family=plt.rcParams['font.family']))
            
            for i, v in enumerate(counts):
                ax1.text(i, v, str(v), ha='center', va='bottom')
            
            st.pyplot(fig1)
        
        with col2:
            # 金额分布
            fig2, ax2 = plt.subplots()
            amounts = filtered_df.groupby('选商方式')['标的金额'].sum().sort_values(ascending=False)
            amounts.plot(kind='bar', ax=ax2, color='lightgreen')
            
            ax2.set_title("采购类型-合同金额", fontproperties=FontProperties(family=plt.rcParams['font.family']))
            ax2.set_xlabel("采购类型", fontproperties=FontProperties(family=plt.rcParams['font.family']))
            ax2.set_ylabel("金额(元)", fontproperties=FontProperties(family=plt.rcParams['font.family']))
            
            for i, v in enumerate(amounts):
                ax2.text(i, v, f"{v:,.0f}", ha='center', va='bottom')
            
            st.pyplot(fig2)
        
        # 部门分析
        st.subheader("按部门分析")
        dept_fig = go.Figure()
        
        dept_amounts = filtered_df.groupby('承办部门')['标的金额'].sum().reset_index()
        dept_counts = filtered_df['承办部门'].value_counts().reset_index()
        
        dept_fig.add_trace(go.Bar(
            x=dept_amounts['承办部门'],
            y=dept_amounts['标的金额'],
            name='合同金额',
            marker_color='indianred'
        ))
        
        dept_fig.add_trace(go.Bar(
            x=dept_counts['承办部门'],
            y=dept_counts['count'],
            name='合同数量',
            marker_color='lightsalmon'
        ))
        
        dept_fig.update_layout(
            title='部门合同情况',
            xaxis_title='部门',
            barmode='group'
        )
        
        st.plotly_chart(dept_fig, use_container_width=True)
    
    # 5. 数据导出
    st.subheader("数据导出")
    csv = filtered_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        "导出CSV",
        csv,
        f"合同数据_{datetime.now().strftime('%Y%m%d')}.csv",
        "text/csv"
    )

if __name__ == "__main__":
    main()
