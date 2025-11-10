import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import os
import numpy as np
from matplotlib import font_manager
import plotly.graph_objects as go
import matplotlib
import base64
import tempfile
import io
from PIL import Image, ImageFont, ImageDraw

# 设置页面布局
st.set_page_config(page_title="分包合同数据分析", layout="wide")
st.title("分包合同数据分析系统")

# 字体解决方案选择
st.sidebar.header("字体解决方案")
font_solution = st.sidebar.radio(
    "选择字体解决方案",
    ["自动检测", "强制系统字体", "Base64嵌入字体", "图像渲染方案"],
    help="如果2D图表中文显示异常，请尝试不同的解决方案"
)

# 方案1: 自动检测字体
def setup_font_auto():
    """自动检测并设置中文字体"""
    try:
        # 方法1: 直接设置常见中文字体
        plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'SimSun', 'KaiTi', 
                                         'DejaVu Sans', 'Arial Unicode MS', 'sans-serif']
        plt.rcParams['axes.unicode_minus'] = False
        
        # 方法2: 检查系统可用字体
        available_fonts = [f.name for f in font_manager.fontManager.ttflist]
        chinese_fonts = ['SimHei', 'Microsoft YaHei', 'SimSun', 'KaiTi', 
                        'WenQuanYi Micro Hei', 'DejaVu Sans']
        
        for font in chinese_fonts:
            if any(font in f for f in available_fonts):
                plt.rcParams['font.family'] = font
                return f"自动选择: {font}"
        
        # 方法3: 重建字体缓存
        font_manager._rebuild()
        plt.rcParams['font.family'] = 'sans-serif'
        return "使用默认sans-serif字体"
        
    except Exception as e:
        return f"自动检测失败: {str(e)}"

# 方案2: 强制使用系统字体
def setup_font_force():
    """强制使用系统字体"""
    try:
        # 清除所有设置
        matplotlib.rcParams.update(matplotlib.rcParamsDefault)
        
        # 强制设置字体
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Helvetica', 'sans-serif']
        plt.rcParams['axes.unicode_minus'] = False
        
        # 尝试找到任何可用的中文字体
        font_paths = [
            '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
            '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
            'C:/Windows/Fonts/simsun.ttc',
            'C:/Windows/Fonts/msyh.ttc',
        ]
        
        for path in font_paths:
            if os.path.exists(path):
                font_prop = font_manager.FontProperties(fname=path)
                plt.rcParams['font.family'] = font_prop.get_name()
                return f"强制使用: {os.path.basename(path)}"
        
        return "使用系统默认字体"
        
    except Exception as e:
        return f"强制设置失败: {str(e)}"

# 方案3: Base64嵌入字体
def setup_font_base64():
    """使用Base64嵌入字体"""
    try:
        # 这里可以添加一个Base64编码的字体文件
        # 由于字体文件较大，这里只做演示
        # 实际使用时，可以将字体文件转换为base64并嵌入
        
        # 使用默认字体，但设置更兼容的参数
        plt.rcParams['font.family'] = 'DejaVu Sans'
        plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'sans-serif']
        plt.rcParams['axes.unicode_minus'] = False
        
        return "Base64字体方案已启用(需添加字体文件)"
        
    except Exception as e:
        return f"Base64方案失败: {str(e)}"

# 方案4: 图像渲染方案
def create_chart_with_pil(data, title, xlabel, ylabel, chart_type='bar'):
    """使用PIL创建图表，避免字体问题"""
    # 创建matplotlib图表但不显示中文
    fig, ax = plt.subplots(figsize=(10, 6))
    
    if chart_type == 'bar':
        if hasattr(data, 'values'):
            values = data.values
            labels = [str(i) for i in data.index]
        else:
            values = data
            labels = [str(i) for i in range(len(data))]
        
        bars = ax.bar(range(len(values)), values, color='skyblue', alpha=0.7)
        
        # 设置英文标签避免中文问题
        ax.set_title("Chart", fontsize=14)
        ax.set_xlabel("Category", fontsize=12)
        ax.set_ylabel("Value", fontsize=12)
        
        # 设置x轴标签
        ax.set_xticks(range(len(labels)))
        ax.set_xticklabels([f"Cat{i}" for i in range(len(labels))], rotation=45, ha='right')
        
        # 添加数值标签
        for i, bar in enumerate(bars):
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + max(values)*0.01,
                    f'{height:,.0f}', ha='center', va='bottom', fontsize=10)
    
    plt.tight_layout()
    
    # 将图表保存为图像
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    plt.close(fig)
    
    # 使用PIL添加中文文本
    img = Image.open(buf)
    draw = ImageDraw.Draw(img)
    
    # 尝试使用PIL的字体（如果可用）
    try:
        # 尝试加载字体
        font = ImageFont.truetype("arial.ttf", 24)
        title_font = ImageFont.truetype("arial.ttf", 28)
    except:
        # 使用默认字体
        font = ImageFont.load_default()
        title_font = ImageFont.load_default()
    
    # 添加中文标题和标签（作为图像文本）
    try:
        # 标题
        draw.text((50, 20), title, fill='black', font=title_font)
        # x轴标签
        draw.text((400, 350), xlabel, fill='black', font=font)
        # y轴标签
        draw.text((20, 180), ylabel, fill='black', font=font, angle=90)
        
        # 添加分类标签
        for i, label in enumerate(labels):
            try:
                draw.text((100 + i*80, 330), label, fill='black', font=font)
            except:
                draw.text((100 + i*80, 330), f"类别{i}", fill='black', font=font)
                
    except Exception as e:
        st.sidebar.warning(f"PIL文本渲染失败: {str(e)}")
    
    # 将图像转换回BytesIO
    output_buf = io.BytesIO()
    img.save(output_buf, format='PNG')
    output_buf.seek(0)
    
    return output_buf

# 根据选择的方案设置字体
font_status = ""
if font_solution == "自动检测":
    font_status = setup_font_auto()
elif font_solution == "强制系统字体":
    font_status = setup_font_force()
elif font_solution == "Base64嵌入字体":
    font_status = setup_font_base64()
# 图像渲染方案在图表生成时处理

st.sidebar.info(f"字体状态: {font_status}")

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
    
    apply_filter = st.button("执行筛选条件")

# 通用图表生成函数
def create_matplotlib_chart(data, title, xlabel, ylabel, chart_type='bar', color='skyblue'):
    """创建matplotlib图表，处理中文显示"""
    # 如果使用图像渲染方案，直接返回PIL图像
    if font_solution == "图像渲染方案":
        return create_chart_with_pil(data, title, xlabel, ylabel, chart_type)
    
    # 否则使用常规matplotlib方法
    fig, ax = plt.subplots(figsize=(10, 6))
    
    if chart_type == 'bar':
        if hasattr(data, 'values'):
            values = data.values
            labels = data.index
        else:
            values = data
            labels = range(len(data))
        
        bars = ax.bar(range(len(values)), values, color=color, alpha=0.7)
        
        # 设置标题和标签
        try:
            ax.set_title(title, fontsize=14, fontweight='bold')
            ax.set_xlabel(xlabel, fontsize=12)
            ax.set_ylabel(ylabel, fontsize=12)
        except:
            # 如果中文设置失败，使用英文
            ax.set_title("Chart", fontsize=14, fontweight='bold')
            ax.set_xlabel("X Axis", fontsize=12)
            ax.set_ylabel("Y Axis", fontsize=12)
        
        # 设置x轴标签
        if hasattr(data, 'index'):
            try:
                ax.set_xticks(range(len(labels)))
                ax.set_xticklabels(labels, rotation=45, ha='right')
            except:
                # 如果设置中文标签失败，使用索引
                ax.set_xticks(range(len(labels)))
                ax.set_xticklabels([f"Cat{i}" for i in range(len(labels))], rotation=45, ha='right')
        
        # 在柱子上方显示数值
        for i, bar in enumerate(bars):
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + max(values)*0.01,
                    f'{height:,.0f}' if '金额' in ylabel else f'{height:.0f}',
                    ha='center', va='bottom', fontsize=10)
    
    plt.tight_layout()
    return fig

# 设置Plotly中文字体
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
    
    # 采购类别分析
    st.subheader("采购类别分析")
    
    if chart_type == "2D图表":
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("采购类别合同数量")
            if not filtered_df.empty:
                counts = filtered_df['选商方式'].value_counts()
                
                if font_solution == "图像渲染方案":
                    # 使用PIL图像
                    img_buf = create_matplotlib_chart(
                        counts, 
                        "采购类别合同数量分布", 
                        "采购类别", 
                        "合同数量", 
                        'bar', 
                        'skyblue'
                    )
                    st.image(img_buf, use_column_width=True)
                else:
                    # 使用常规matplotlib
                    fig = create_matplotlib_chart(
                        counts, 
                        "采购类别合同数量分布", 
                        "采购类别", 
                        "合同数量", 
                        'bar', 
                        'skyblue'
                    )
                    st.pyplot(fig)
                    plt.close(fig)
            else:
                st.warning("没有符合条件的数据")
                
        with col2:
            st.subheader("采购类别合同金额")
            if not filtered_df.empty:
                amount_by_type = filtered_df.groupby('选商方式')['标的金额'].sum().sort_values(ascending=False)
                
                if font_solution == "图像渲染方案":
                    # 使用PIL图像
                    img_buf = create_matplotlib_chart(
                        amount_by_type,
                        "采购类别合同金额分布",
                        "采购类别", 
                        "合同金额 (元)", 
                        'bar', 
                        'lightgreen'
                    )
                    st.image(img_buf, use_column_width=True)
                else:
                    # 使用常规matplotlib
                    fig = create_matplotlib_chart(
                        amount_by_type,
                        "采购类别合同金额分布",
                        "采购类别", 
                        "合同金额 (元)", 
                        'bar', 
                        'lightgreen'
                    )
                    st.pyplot(fig)
                    plt.close(fig)
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
            if font_solution == "图像渲染方案":
                img_buf = create_matplotlib_chart(
                    yearly_stats.set_index('年份')['项目数量'],
                    "在建项目数量按年份分布",
                    "年份", 
                    "项目数量", 
                    'bar', 
                    'teal'
                )
                st.image(img_buf, use_column_width=True)
            else:
                fig = create_matplotlib_chart(
                    yearly_stats.set_index('年份')['项目数量'],
                    "在建项目数量按年份分布",
                    "年份", 
                    "项目数量", 
                    'bar', 
                    'teal'
                )
                st.pyplot(fig)
                plt.close(fig)
        
        with col4:
            st.subheader("在建项目金额按年份分布")
            if font_solution == "图像渲染方案":
                img_buf = create_matplotlib_chart(
                    yearly_stats.set_index('年份')['合同金额'],
                    "在建项目金额按年份分布", 
                    "年份", 
                    "合同金额 (元)",
                    'bar', 
                    'purple'
                )
                st.image(img_buf, use_column_width=True)
            else:
                fig = create_matplotlib_chart(
                    yearly_stats.set_index('年份')['合同金额'],
                    "在建项目金额按年份分布", 
                    "年份", 
                    "合同金额 (元)",
                    'bar', 
                    'purple'
                )
                st.pyplot(fig)
                plt.close(fig)
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

# 字体解决方案说明
with st.sidebar:
    with st.expander("字体解决方案说明"):
        st.markdown("""
        **解决方案说明**:
        
        1. **自动检测** - 尝试自动检测系统中可用的中文字体
        2. **强制系统字体** - 强制使用系统默认字体，避免字体缓存问题
        3. **Base64嵌入字体** - 将字体文件嵌入代码中(需要添加字体文件)
        4. **图像渲染方案** - 使用PIL库渲染中文，完全避免matplotlib字体问题
        
        **如果2D图表仍显示方框**:
        - 优先尝试"图像渲染方案"
        - 或使用3D交互图表(Plotly兼容性更好)
        """)

# 显示原始数据统计信息
with st.expander("原始数据统计信息"):
    st.subheader("数据概览")
    st.write(f"总记录数: {len(df)}")
    
    st.subheader("各字段统计")
    st.write(df.describe(include='all'))
    
    st.subheader("前5条记录")
    st.dataframe(df.head())
