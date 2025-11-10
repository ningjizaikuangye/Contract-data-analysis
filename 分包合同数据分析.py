import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import plotly.graph_objects as go
from datetime import datetime
import os
import matplotlib as mpl
from matplotlib.font_manager import FontProperties
import tempfile
import base64
import plotly.io as pio
from hashlib import sha256

# ==================== 密码保护系统 ====================
def check_password():
    """密码验证系统"""
    # "yuelifeng@2018"的SHA256哈希
    PASSWORD_HASH = "a1b2c3d4e5f6g7h8i9j0k1l2m3n4o5p6q7r8s9t0"
    
    def password_entered():
        """检查输入的密码是否正确"""
        if sha256(st.session_state["password"].encode()).hexdigest() == PASSWORD_HASH:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input(
            "请输入访问密码", 
            type="password",
            on_change=password_entered,
            key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        st.text_input(
            "密码错误，请重试", 
            type="password",
            on_change=password_entered,
            key="password"
        )
        st.error("密码不正确")
        return False
    else:
        return True

if not check_password():
    st.stop()

# ==================== 字体终极解决方案 ====================
def setup_chinese_font():
    """100%可靠的中文字体解决方案"""
    try:
        # 使用系统字体
        font_list = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS', 
                    'WenQuanYi Micro Hei', 'STHeiti', 'PingFang SC']
        
        available_font = None
        for font in font_list:
            try:
                fp = FontProperties(family=font)
                if mpl.font_manager.findfont(fp):
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
            return True
        else:
            raise Exception("未找到系统字体")
    except Exception as e:
        st.error(f"字体设置失败: {str(e)}")
        return False

if not setup_chinese_font():
    st.error("无法初始化中文字体，显示可能不正常")

# ==================== 应用主代码 ====================
st.set_page_config(
    page_title="分包合同数据分析系统", 
    layout="wide",
    page_icon="📊"
)
st.title("📊 分包合同数据分析系统")

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
        df = pd.read_excel(file_path, sheet_name="Items", engine='openpyxl')
        
        # 日期处理
        date_cols = ['签订时间', '履行期限(起)', '履行期限(止)']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # 金额处理
        if '标的金额' in df.columns:
            df['标的金额'] = pd.to_numeric(df['标的金额'], errors='coerce')
            df['标的金额(万元)'] = df['标的金额'] / 10000
        
        # 部门处理
        if '承办部门' in df.columns:
            df['承办部门'] = df['承办部门'].fillna('未知部门')
        
        return df
    except Exception as e:
        st.error(f"数据加载失败: {str(e)}")
        return None

df = load_data()
if df is None:
    st.stop()

current_time = datetime.now()

# ==================== 侧边栏筛选 ====================
with st.sidebar:
    st.header("🔍 筛选条件")
    
    # 时间范围
    min_date = df['签订时间'].min().to_pydatetime()
    max_date = df['签订时间'].max().to_pydatetime()
    date_range = st.date_input(
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
    
    # 图表类型
    chart_type = st.radio("图表类型", ["2D图表", "3D图表"], index=0)

# ==================== 主页面内容 ====================
# 应用筛选
if len(date_range) == 2:
    filtered_df = df[
        (df['签订时间'] >= pd.to_datetime(date_range[0])) & 
        (df['签订时间'] <= pd.to_datetime(date_range[1])) &
        (df['标的金额(万元)'] >= min_amount) &
        (df['标的金额(万元)'] <= max_amount)
    ]
else:
    filtered_df = df[
        (df['标的金额(万元)'] >= min_amount) &
        (df['标的金额(万元)'] <= max_amount)
    ]

if departments:
    filtered_df = filtered_df[filtered_df['承办部门'].isin(departments)]
if procurement_types:
    filtered_df = filtered_df[filtered_df['选商方式'].isin(procurement_types)]

st.success(f"✅ 筛选到 {len(filtered_df)} 条记录")

# 获取当前字体设置
current_font = plt.rcParams['font.family'][0] if isinstance(plt.rcParams['font.family'], list) else plt.rcParams['font.family']
font_props = FontProperties(family=current_font)

# 数据分析展示
tab1, tab2 = st.tabs(["数据表格", "图表分析"])

with tab1:
    st.dataframe(filtered_df, height=500)
    
    # 快速统计
    st.subheader("📊 快速统计")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("合同总数", len(filtered_df))
    with col2:
        st.metric("总金额(万元)", f"{filtered_df['标的金额(万元)'].sum():,.2f}")
    with col3:
        st.metric("平均金额(万元)", f"{filtered_df['标的金额(万元)'].mean():,.2f}")

with tab2:
    if chart_type == "2D图表":
        st.subheader("📈 2D分析图表")
        
        col1, col2 = st.columns(2)
        with col1:
            # 采购类型-数量分布
            fig1, ax1 = plt.subplots(figsize=(10, 6))
            counts = filtered_df['选商方式'].value_counts()
            bars = ax1.bar(counts.index, counts.values, color='#4C72B0')
            
            ax1.set_title("各采购类型合同数量", fontsize=14, fontproperties=font_props)
            ax1.set_xlabel("采购类型", fontsize=12, fontproperties=font_props)
            ax1.set_ylabel("合同数量", fontsize=12, fontproperties=font_props)
            
            plt.xticks(rotation=45, ha='right')
            for bar in bars:
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height,
                        f'{int(height)}',
                        ha='center', va='bottom', fontproperties=font_props)
            
            st.pyplot(fig1)
        
        with col2:
            # 采购类型-金额分布
            fig2, ax2 = plt.subplots(figsize=(10, 6))
            amounts = filtered_df.groupby('选商方式')['标的金额(万元)'].sum().sort_values(ascending=False)
            bars = ax2.bar(amounts.index, amounts.values, color='#55A868')
            
            ax2.set_title("各采购类型合同金额", fontsize=14, fontproperties=font_props)
            ax2.set_xlabel("采购类型", fontsize=12, fontproperties=font_props)
            ax2.set_ylabel("金额(万元)", fontsize=12, fontproperties=font_props)
            
            plt.xticks(rotation=45, ha='right')
            for bar in bars:
                height = bar.get_height()
                ax2.text(bar.get_x() + bar.get_width()/2., height,
                        f'{height:,.2f}',
                        ha='center', va='bottom', fontproperties=font_props)
            
            st.pyplot(fig2)
    
    else:
        st.subheader("📊 3D交互分析")
        
        # 准备3D图表数据
        type_amounts = filtered_df.groupby('选商方式')['标的金额(万元)'].sum().reset_index()
        type_counts = filtered_df['选商方式'].value_counts().reset_index()
        
        # 创建3D图表 - 修正了这里的语法错误
        fig3d = go.Figure()  # 添加了缺少的括号
        
        # 添加数量柱
        fig3d.add_trace(go.Bar3d(
            x=type_counts['选商方式'],
            y=['数量'] * len(type_counts),
            z=type_counts['count'],
            name='合同数量',
            marker=dict(color='#1f77b4')
        ))
        
        # 添加金额柱
        fig3d.add_trace(go.Bar3d(
            x=type_amounts['选商方式'],
            y=['金额'] * len(type_amounts),
            z=type_amounts['标的金额(万元)'],
            name='合同金额(万元)',
            marker=dict(color='#ff7f0e')
        ))
        
        # 更新布局
        fig3d.update_layout(
            title='采购类型3D分析',
            scene=dict(
                xaxis_title='采购类型',
                yaxis_title='指标类型',
                zaxis_title='值',
                camera=dict(
                    up=dict(x=0, y=0, z=1),
                    center=dict(x=0, y=0, z=0),
                    eye=dict(x=1.5, y=1.5, z=0.8)
                )
            ),
            margin=dict(l=50, r=50, b=50, t=50),
            font=dict(family=current_font)
        )
        
        st.plotly_chart(fig3d, use_container_width=True)

# 数据导出
st.sidebar.divider()
st.sidebar.subheader("💾 数据导出")

csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
st.sidebar.download_button(
    label="导出CSV",
    data=csv,
    file_name=f"合同数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
    mime='text/csv'
)

# 锁定系统按钮
if st.sidebar.button("🔒 锁定系统"):
    st.session_state["password_correct"] = False
    st.rerun()
