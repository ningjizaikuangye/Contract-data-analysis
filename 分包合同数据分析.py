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
# ==================== 应用主代码 ====================
# 设置页面布局
st.set_page_config(page_title="分包合同数据分析系统", layout="wide")
st.title("分包合同数据分析系统")
# 定义文件路径
file_path = r"03 合同2.0系统数据.xlsm"  # 确保文件在同一个目录下
# 检查文件是否存在
if not os.path.exists(file_path):
    st.error(f"文件未找到: {file_path}")
    st.stop()
# 读取Excel数据
@st.cache_data
def load_data():
    try:
        df = pd.read_excel(file_path, sheet_name="Items", engine='openpyxl')
        date_cols = ['签订时间', '履行期限(起)', '履行期限(止)']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        if '标的金额' in df.columns:
            df['标的金额'] = pd.to_numeric(df['标的金额'], errors='coerce')
        if '承办部门' in df.columns:
            df['承办部门'] = df['承办部门'].fillna('未知部门')
        return df
    except Exception as e:
        st.error(f"读取数据时出错: {str(e)}")
        return None
df = load_data()
if df is None:
    st.stop()
current_time = datetime.now()
# ==================== 侧边栏筛选 ====================
with st.sidebar:
    st.header("筛选条件")
    
    # 时间范围
    min_date = df['签订时间'].min().to_pydatetime()
    max_date = df['签订时间'].max().to_pydatetime()
    start_date = st.date_input("最早签订时间", min_date, min_value=min_date, max_value=max_date)
    end_date = st.date_input("最晚签订时间", max_date, min_value=min_date, max_value=max_date)
    
    # 金额范围
    min_amount = float(df['标的金额'].min())
    max_amount = float(df['标的金额'].max())
    col1, col2 = st.columns(2)
    with col1:
        min_amount_input = st.number_input("最低合同金额 (元)", min_value=min_amount, max_value=max_amount, 
                                         value=min_amount, step=1.0, format="%.0f")
    with col2:
        max_amount_input = st.number_input("最高合同金额 (元)", min_value=min_amount, max_value=max_amount, 
                                         value=max_amount, step=1.0, format="%.0f")
    
    # 部门筛选
    departments = df['承办部门'].unique().tolist()
    selected_departments = st.multiselect("选择承办部门", departments, default=departments)
    
    # 采购类别(动态更新)
    if selected_departments:
        procurement_types = df[df['承办部门'].isin(selected_departments)]['选商方式'].unique().tolist()
    else:
        procurement_types = df['选商方式'].unique().tolist()
    selected_types = st.multiselect("选择采购类别", procurement_types, default=procurement_types)
    
    # 图表类型选择
    chart_type = st.radio("选择图表类型", ["2D图表", "3D交互图表"])
    
    apply_filter = st.button("执行筛选条件")
# ==================== 主页面内容 ====================
if apply_filter:
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    
    filtered_df = df[
        (df['签订时间'] >= start_date) & 
        (df['签订时间'] <= end_date) & 
        (df['标的金额'] >= min_amount_input) & 
        (df['标的金额'] <= max_amount_input) & 
        (df['选商方式'].isin(selected_types)) &
        (df['承办部门'].isin(selected_departments))
    ].copy()
    
    st.success(f"筛选到 {len(filtered_df)} 条记录")
    st.dataframe(filtered_df.head())
    
    # 采购类别分析
    st.subheader("采购类别分析")
    
    if chart_type == "2D图表":
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("采购类别合同数量")
            if not filtered_df.empty:
                fig, ax = plt.subplots(figsize=(10, 6))
                counts = filtered_df['选商方式'].value_counts()
                counts.plot(kind='bar', ax=ax, color='skyblue')
                ax.set_ylabel("合同数量", fontsize=12)
                ax.set_xlabel("采购类别", fontsize=12)
                ax.set_title("采购类别合同数量分布", fontsize=14)
                
                plt.xticks(rotation=45, ha='right')
                
                for i, v in enumerate(counts):
                    ax.text(i, v + 0.5, str(v), ha='center', va='bottom')
                
                plt.tight_layout()
                st.pyplot(fig)
            else:
                st.warning("没有符合条件的数据")
                
        with col2:
            st.subheader("采购类别合同金额")
            if not filtered_df.empty:
                fig, ax = plt.subplots(figsize=(10, 6))
                amount_by_type = filtered_df.groupby('选商方式')['标的金额'].sum().sort_values(ascending=False)
                amount_by_type.plot(kind='bar', ax=ax, color='lightgreen')
                ax.set_ylabel("合同金额 (元)", fontsize=12)
                ax.set_xlabel("采购类别", fontsize=12)
                ax.set_title("采购类别合同金额分布", fontsize=14)
                
                plt.xticks(rotation=45, ha='right')
                
                for i, v in enumerate(amount_by_type):
                    ax.text(i, v + max(amount_by_type)*0.01, f"{v:,.0f}", 
                           ha='center', va='bottom')
                
                plt.tight_layout()
                st.pyplot(fig)
            else:
                st.warning("没有符合条件的数据")
    
    else:  # 3D交互图表
        if not filtered_df.empty:
            counts = filtered_df['选商方式'].value_counts().reset_index()
            counts.columns = ['采购类别', '合同数量']
            amounts = filtered_df.groupby('选商方式')['标的金额'].sum().reset_index()
            amounts.columns = ['采购类别', '合同金额']
            
            st.subheader("采购类别3D分析(数量与金额)")
            fig = go.Figure()
            
            for i, row in counts.iterrows():
                fig.add_trace(go.Scatter3d(
                    x=[row['采购类别'], row['采购类别']],
                    y=['数量', '数量'],
                    z=[0, row['合同数量']],
                    mode='lines',
                    line=dict(color='skyblue', width=10),
                    name=f"{row['采购类别']} 数量",
                    showlegend=False,
                    hoverinfo='text',
                    hovertext=f"采购类别: {row['采购类别']}<br>数量: {row['合同数量']}"
                ))
            
            max_count = counts['合同数量'].max()
            max_amount = amounts['合同金额'].max()
            
            for i, row in amounts.iterrows():
                scaled_amount = row['合同金额'] / max_amount * max_count
                fig.add_trace(go.Scatter3d(
                    x=[row['采购类别'], row[' procurement_category']],
                    y=['金额', '金额'],
                    z=[0, scaled_amount],
                    mode='lines',
                    line=dict(color='lightgreen', width=10),
                    name=f"{row['采购类别']} 金额",
                    showlegend=False,
                    hoverinfo='text',
                    hovertext=f"采购类别: {row['采购类别']}<br>金额: {row['合同金额']:,.0f}元"
                ))
            
            fig.update_layout(
                scene=dict(
                    xaxis_title='采购类别',
                    yaxis_title='指标类型',
                    zaxis_title='值',
                    camera=dict(
                        up=dict(x=0, y=0, z=1),
                        center=dict(x=0, y=0, z=0),
                        eye=dict(x=1.5, y=1.5, z=0.8)
                    ),
                    aspectratio=dict(x=1.5, y=1, z=0.8)
                ),
                title="采购类别3D分析(数量与金额)",
                font=dict(family=plt.rcParams['font.family']),
                width=1000,
                height=600,
                margin=dict(l=50, r=50, b=50, t=50),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5)
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            with st.expander("3D图表操作指南"):
                st.markdown("""
                **交互操作**:
                - 鼠标左键拖动: 旋转视角
                - 鼠标右键拖动: 平移视图
                - 鼠标滚轮: 缩放视图
                - 悬停在柱子上: 查看详细数据
                
                **图表说明**:
                - 蓝色柱子: 合同实际数量
                - 绿色柱子: 合同金额(按比例缩放)
                """)
        else:
            st.warning("没有符合条件的数据用于生成3D图表")
    
    # 在建项目分析
    st.subheader("在建项目分析")
    
    ongoing_projects = df[
        (df['履行期限(止)'] > current_time) &
        (df['承办部门'].isin(selected_departments)) &
        (df['选商方式'].isin(selected_types))
    ].copy()
    
    if not ongoing_projects.empty:
        ongoing_projects['年份'] = ongoing_projects['履行期限(起)'].dt.year
        
        yearly_stats = ongoing_projects.groupby('年份').agg(
            项目数量=('标的金额', 'count'),
            合同金额=('标的金额', 'sum')
        ).reset_index()
        
        col3, col4 = st.columns(2)
        
        with col3:
            st.subheader("在建项目数量按年份分布")
            fig, ax = plt.subplots(figsize=(10, 6))
            yearly_stats.plot(x='年份', y='项目数量', kind='bar', ax=ax, color='teal')
            ax.set_ylabel("项目数量", fontsize=12)
            ax.set_xlabel("年份", fontsize=12)
            ax.set_title("在建项目数量按年份分布", fontsize=14)
            
            for i, v in enumerate(yearly_stats['项目数量']):
                ax.text(i, v + 0.5, str(v), ha='center', va='bottom')
            
            plt.xticks(rotation=0)
            plt.tight_layout()
            st.pyplot(fig)
        
        with col4:
            st.subheader("在建项目金额按年份分布")
            fig, ax = plt.subplots(figsize=(10, 6))
            yearly_stats.plot(x='年份', y='合同金额', kind='bar', ax=ax, color='purple')
            ax.set_ylabel("合同金额 (元)", fontsize=12)
            ax.set_xlabel("年份", fontsize=12)
            ax.set_title("在建项目金额按年份分布", fontsize=14)
            
            for i, v in enumerate(yearly_stats['合同金额']):
                ax.text(i, v + max(yearly_stats['合同金额'])*0.01, f"{v:,.0f}", 
                       ha='center', va='bottom')
            
            plt.xticks(rotation=0)
            plt.tight_layout()
            st.pyplot(fig)
        
        with st.expander("在建项目详情"):
            st.dataframe(ongoing_projects.sort_values('履行期限(止)', ascending=True))
    else:
        st.warning("没有符合条件的在建项目")
    
    st.subheader("数据导出")
    csv = filtered_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="下载筛选结果 (CSV)",
        data=csv,
        file_name=f"分包合同数据_{datetime.now().strftime('%Y%m%d')}.csv",
        mime='text_csv'
    )
else:
    st.info("请在左侧边栏设置筛选条件，然后点击'执行筛选条件'按钮")
with st.expander("原始数据统计信息"):
    st.subheader("数据概览")
    st.write(f"总记录数: {len(df)}")
    
    st.subheader("各字段统计")
    st.write(df.describe(include='all'))
    
    st.subheader("前5条记录")
    st.dataframe(df.head())
