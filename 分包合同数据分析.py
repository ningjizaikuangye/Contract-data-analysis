import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import os
import numpy as np
from matplotlib import font_manager
import plotly.graph_objects as go
import base64
import io

# 改进的字体设置函数
def setup_chinese_font():
    """设置中文字体，兼容公共网页环境"""
    try:
        # 方法1: 尝试使用系统内置的中文字体
        chinese_fonts = [
            'SimHei', 'Microsoft YaHei', 'SimSun', 'KaiTi', 
            'DejaVu Sans', 'Arial Unicode MS', 'sans-serif'
        ]
        
        # 获取系统可用字体
        available_fonts = set([f.name for f in font_manager.fontManager.ttflist])
        
        # 选择第一个可用的中文字体
        for font in chinese_fonts:
            if font in available_fonts:
                plt.rcParams['font.family'] = font
                st.info(f"使用字体: {font}")
                break
        else:
            # 如果没有找到中文字体，使用默认字体并警告
            plt.rcParams['font.family'] = ['DejaVu Sans', 'sans-serif']
            st.warning("未找到中文字体，图表可能无法正常显示中文")
        
        # 设置其他参数
        plt.rcParams['axes.unicode_minus'] = False
        plt.rcParams['font.size'] = 12
        
        return True
        
    except Exception as e:
        st.error(f"字体设置失败: {str(e)}")
        # 设置最基本的回退
        plt.rcParams['font.family'] = ['sans-serif']
        plt.rcParams['axes.unicode_minus'] = False
        return False

# 初始化字体设置
font_setup_success = setup_chinese_font()

# 设置页面布局
st.set_page_config(page_title="分包合同数据分析", layout="wide")
st.title("分包合同数据分析系统")

# 定义文件路径
file_path = r"03 合同2.0系统数据.xlsm"

# 检查文件是否存在
if not os.path.exists(file_path):
    st.error(f"文件未找到: {file_path}")
    st.stop()

# 读取Excel数据
@st.cache_data
def load_data():
    try:
        df = pd.read_excel(file_path, sheet_name="Items")
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

# 侧边栏设置
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
    
    # 字体状态显示
    if not font_setup_success:
        st.warning("⚠️ 中文字体加载异常，图表可能显示乱码")
    
    apply_filter = st.button("执行筛选条件")

# 改进的matplotlib图表生成函数
def create_matplotlib_chart(data, x_col, y_col, title, x_label, y_label, chart_type='bar', color='skyblue'):
    """创建matplotlib图表，处理中文显示"""
    fig, ax = plt.subplots(figsize=(10, 6))
    
    if chart_type == 'bar':
        data.plot(kind='bar', ax=ax, color=color)
    elif chart_type == 'line':
        data.plot(kind='line', ax=ax, color=color, marker='o')
    
    # 设置标题和标签，处理中文
    try:
        ax.set_title(title, fontsize=14)
        ax.set_xlabel(x_label, fontsize=12)
        ax.set_ylabel(y_label, fontsize=12)
    except:
        # 如果中文设置失败，使用英文标签
        ax.set_title("Chart", fontsize=14)
        ax.set_xlabel("X Axis", fontsize=12)
        ax.set_ylabel("Y Axis", fontsize=12)
    
    plt.xticks(rotation=45)
    plt.tight_layout()
    return fig

# 改进的plotly图表设置
def setup_plotly_chinese_font(fig):
    """设置Plotly图表的中文字体"""
    try:
        fig.update_layout(
            font=dict(
                family="Microsoft YaHei, SimHei, Arial, sans-serif",
                size=12,
            )
        )
    except:
        # 如果字体设置失败，使用默认字体
        pass
    return fig

# 主页面
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
                counts = filtered_df['选商方式'].value_counts()
                fig = create_matplotlib_chart(
                    counts, None, None, 
                    "采购类别合同数量分布", 
                    "采购类别", "合同数量", 
                    'bar', 'skyblue'
                )
                
                # 添加数值标签
                ax = fig.axes[0]
                for i, v in enumerate(counts):
                    try:
                        ax.text(i, v + 0.5, str(v), ha='center', va='bottom')
                    except:
                        ax.text(i, v + 0.5, str(v), ha='center', va='bottom')
                
                st.pyplot(fig)
                plt.close(fig)  # 关闭图形释放内存
            else:
                st.warning("没有符合条件的数据")
                
        with col2:
            st.subheader("采购类别合同金额")
            if not filtered_df.empty:
                amount_by_type = filtered_df.groupby('选商方式')['标的金额'].sum().sort_values(ascending=False)
                fig = create_matplotlib_chart(
                    amount_by_type, None, None,
                    "采购类别合同金额分布",
                    "采购类别", "合同金额 (元)", 
                    'bar', 'lightgreen'
                )
                
                # 添加数值标签
                ax = fig.axes[0]
                for i, v in enumerate(amount_by_type):
                    try:
                        ax.text(i, v + max(amount_by_type)*0.01, f"{v:,.0f}", ha='center', va='bottom')
                    except:
                        ax.text(i, v + max(amount_by_type)*0.01, f"{v:,.0f}", ha='center', va='bottom')
                
                st.pyplot(fig)
                plt.close(fig)  # 关闭图形释放内存
            else:
                st.warning("没有符合条件的数据")
    
    else:  # 3D交互图表
        if not filtered_df.empty:
            # 准备数据
            counts = filtered_df['选商方式'].value_counts().reset_index()
            counts.columns = ['采购类别', '合同数量']
            amounts = filtered_df.groupby('选商方式')['标的金额'].sum().reset_index()
            amounts.columns = ['采购类别', '合同金额']
            
            # 创建3D柱状图
            st.subheader("采购类别3D分析(数量与金额)")
            
            # 创建图形
            fig = go.Figure()
            
            # 添加数量柱子
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
            
            # 添加金额柱子(按比例缩放)
            max_count = counts['合同数量'].max()
            max_amount = amounts['合同金额'].max()
            
            for i, row in amounts.iterrows():
                scaled_amount = row['合同金额'] / max_amount * max_count
                fig.add_trace(go.Scatter3d(
                    x=[row['采购类别'], row['采购类别']],
                    y=['金额', '金额'],
                    z=[0, scaled_amount],
                    mode='lines',
                    line=dict(color='lightgreen', width=10),
                    name=f"{row['采购类别']} 金额",
                    showlegend=False,
                    hoverinfo='text',
                    hovertext=f"采购类别: {row['采购类别']}<br>金额: {row['合同金额']:,.0f}元"
                ))
            
            # 更新布局，设置中文字体
            fig = setup_plotly_chinese_font(fig)
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
                width=1000,
                height=600,
                margin=dict(l=50, r=50, b=50, t=50),
                showlegend=True
            )
            
            # 添加图例
            fig.add_trace(go.Scatter3d(
                x=[None],
                y=[None],
                z=[None],
                mode='markers',
                marker=dict(size=10, color='skyblue'),
                name='合同数量',
                showlegend=True
            ))
            
            fig.add_trace(go.Scatter3d(
                x=[None],
                y=[None],
                z=[None],
                mode='markers',
                marker=dict(size=10, color='lightgreen'),
                name='合同金额(比例)',
                showlegend=True
            ))
            
            st.plotly_chart(fig, use_container_width=True)
            
            # 使用说明
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
    
    # 筛选在建项目（履行期限(止) > 当前时间）
    ongoing_projects = df[
        (df['履行期限(止)'] > current_time) &
        (df['承办部门'].isin(selected_departments)) &
        (df['选商方式'].isin(selected_types))
    ].copy()
    
    if not ongoing_projects.empty:
        # 提取年份
        ongoing_projects['年份'] = ongoing_projects['履行期限(起)'].dt.year
        
        # 按年份分组统计
        yearly_stats = ongoing_projects.groupby('年份').agg(
            项目数量=('标的金额', 'count'),
            合同金额=('标的金额', 'sum')
        ).reset_index()
        
        col3, col4 = st.columns(2)
        
        with col3:
            st.subheader("在建项目数量按年份分布")
            fig = create_matplotlib_chart(
                yearly_stats.set_index('年份')['项目数量'], None, None,
                "在建项目数量按年份分布",
                "年份", "项目数量",
                'bar', 'teal'
            )
            st.pyplot(fig)
            plt.close(fig)
        
        with col4:
            st.subheader("在建项目金额按年份分布")
            fig = create_matplotlib_chart(
                yearly_stats.set_index('年份')['合同金额'], None, None,
                "在建项目金额按年份分布", 
                "年份", "合同金额 (元)",
                'bar', 'purple'
            )
            st.pyplot(fig)
            plt.close(fig)
        
        # 显示在建项目详情
        with st.expander("在建项目详情"):
            st.dataframe(ongoing_projects.sort_values('履行期限(止)', ascending=True))
    else:
        st.warning("没有符合条件的在建项目")
    
    # 添加下载按钮
    st.subheader("数据导出")
    csv = filtered_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="下载筛选结果 (CSV)",
        data=csv,
        file_name=f"分包合同数据_{datetime.now().strftime('%Y%m%d')}.csv",
        mime='text/csv'
    )
else:
    st.info("请在左侧边栏设置筛选条件，然后点击'执行筛选条件'按钮")

# 显示原始数据统计信息
with st.expander("原始数据统计信息"):
    st.subheader("数据概览")
    st.write(f"总记录数: {len(df)}")
    
    st.subheader("各字段统计")
    st.write(df.describe(include='all'))
    
    st.subheader("前5条记录")
    st.dataframe(df.head())

# 添加字体状态说明
with st.sidebar:
    with st.expander("字体状态说明"):
        if font_setup_success:
            st.success("✅ 中文字体已正确加载")
        else:
            st.error("❌ 中文字体加载失败")
            st.info("""
            **解决方案**:
            1. 确保服务器环境安装了中文字体
            2. 或者将字体文件打包到应用中
            3. 考虑使用Web安全字体
            """)
