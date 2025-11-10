import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import plotly.graph_objects as go
import matplotlib as mpl
from matplotlib import font_manager
from datetime import datetime
import plotly.io as pio
import os
from matplotlib.font_manager import FontProperties

# ===== 字体终极解决方案 =====
def setup_chinese_font():
    """确保中文显示的100%可靠方案"""
    try:
        # 系统字体优先级列表（跨平台兼容）
        font_preference = [
            'Microsoft YaHei',    # Windows
            'SimHei',             # Windows
            'Arial Unicode MS',   # Mac
            'PingFang SC',        # Mac
            'WenQuanYi Micro Hei',# Linux
            'Noto Sans CJK SC',   # Linux
            'sans-serif'          # 最终回退
        ]
        
        # 测试并选择第一个可用的字体
        available_font = None
        for font in font_preference:
            try:
                test_font = FontProperties(family=font)
                font_path = font_manager.findfont(test_font)
                if font_path:
                    available_font = font
                    break
            except:
                continue
        
        if available_font:
            # 配置Matplotlib
            plt.rcParams['font.family'] = available_font
            plt.rcParams['axes.unicode_minus'] = False
            
            # 配置Plotly
            pio.templates.default = "plotly_white"
            pio.templates["plotly_white"].layout.font.family = available_font
            
            st.success(f"字体设置成功: 使用 {available_font}")
            return True
        else:
            raise RuntimeError("未找到任何可用字体")
            
    except Exception as e:
        st.warning(f"字体设置警告: {str(e)}，使用备用方案")
        # 强制回退到基本字体
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['axes.unicode_minus'] = False
        pio.templates.default = "plotly_white"
        return False

# 初始化字体
setup_chinese_font()

# ===== 主应用程序 =====
def main():
    # 页面配置
    st.set_page_config(
        page_title="分包合同数据分析系统",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    st.title("📈 分包合同数据分析系统")
    
    # 1. 数据加载函数
    @st.cache_data
    def load_data():
        try:
            # 读取Excel文件
            df = pd.read_excel("03 合同2.0系统数据.xlsm", sheet_name="Items", engine='openpyxl')
            
            # 日期列处理
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
            st.error(f"⚠️ 数据加载失败: {str(e)}")
            return None
    
    # 加载数据
    df = load_data()
    if df is None:
        st.stop()
    
    # 2. 获取当前时间（带错误处理）
    try:
        current_time = datetime.now()
    except Exception as e:
        st.warning(f"时间获取警告: {str(e)}，使用替代方案")
        current_time = pd.Timestamp.now()  # 使用pandas的备用方案
    
    # 3. 侧边栏筛选器
    with st.sidebar:
        st.header("🔍 数据筛选条件")
        
        # 时间范围选择
        min_date = df['签订时间'].min().to_pydatetime()
        max_date = df['签订时间'].max().to_pydatetime()
        date_range = st.date_input(
            "合同签订时间范围",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )
        
        # 金额范围选择
        col1, col2 = st.columns(2)
        with col1:
            min_amount = st.number_input(
                "最小金额(万元)",
                value=float(df['标的金额(万元)'].min()),
                min_value=0.0,
                step=0.01
            )
        with col2:
            max_amount = st.number_input(
                "最大金额(万元)",
                value=float(df['标的金额(万元)'].max()),
                min_value=0.0,
                step=0.01
            )
        
        # 部门和采购类型多选
        departments = st.multiselect(
            "选择承办部门",
            options=df['承办部门'].unique().tolist(),
            default=df['承办部门'].unique().tolist()
        )
        
        procurement_types = st.multiselect(
            "选择采购类型",
            options=df['选商方式'].unique().tolist(),
            default=df['选商方式'].unique().tolist()
        )
        
        # 图表类型选择
        chart_type = st.radio("图表类型", ["2D图表", "3D图表"], index=0)
    
    # 4. 应用筛选条件
    if len(date_range) == 2:
        filtered_df = df[
            (df['签订时间'] >= pd.to_datetime(date_range[0])) &
            (df['签订时间'] <= pd.to_datetime(date_range[1])) &
            (df['标的金额(万元)'] >= min_amount) &
            (df['标的金额(万元)'] <= max_amount) &
            (df['承办部门'].isin(departments)) &
            (df['选商方式'].isin(procurement_types))
        ]
    else:
        filtered_df = df[
            (df['标的金额(万元)'] >= min_amount) &
            (df['标的金额(万元)'] <=max_amount) &
            (df['承办部门'].isin(departments)) &
            (df['选商方式'].isin(procurement_types))
        ]
    
    st.success(f"✅ 筛选到 {len(filtered_df)} 条记录")
    
    # 5. 主显示区域
    tab1, tab2, tab3 = st.tabs(["数据浏览", "统计分析", "图表展示"])
    
    with tab1:
        st.dataframe(filtered_df, height=500, use_container_width=True)
        
        # 快速统计
        st.subheader("📊 快速统计")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("合同总数", len(filtered_df))
        with col2:
            st.metric("总金额(万元)", f"{filtered_df['标的金额(万元)'].sum():,.2f}")
        with col3:
            st.metric("平均金额(万元)", f"{filtered_df['标的金额'].mean()/10000:,.2f}")
    
    with tab2:
        # 部门分析
        st.subheader("🏢 按部门分析")
        dept_stats = filtered_df.groupby('承办部门').agg(
            合同数量=('标的金额', 'count'),
            总金额_万元=('标的金额(万元)', 'sum'),
            平均金额_万元=('标的金额(万元)', 'mean')
        ).sort_values('总金额_万元', ascending=False)
        
        st.dataframe(
            dept_stats.style.format({
                '总金额_万元': '{:,.2f}',
                '平均金额_万元': '{:,.2f}'
            }),
            height=400
        )
        
        # 采购类型分析
        st.subheader("🛒 按采购类型分析")
        type_stats = filtered_df.groupby('选商方式').agg(
            合同数量=('标的金额', 'count'),
            总金额_万元=('标的金额(万元)', 'sum')
        ).sort_values('总金额_万元', ascending=False)
        
        st.dataframe(
            type_stats.style.format({'总金额_万元': '{:,.2f}'}),
            height=400
        )
    
    with tab3:
        # 获取当前字体设置
        current_font = plt.rcParams['font.family'][0] if isinstance(plt.rcParams['font.family'], list) else plt.rcParams['font.family']
        font_props = FontProperties(family=current_font)
        
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
                
                # 旋转标签并添加数值
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
            
            # 创建3D图表
            fig3d = go.Figure()
            
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
    
    # 6. 数据导出功能
    st.sidebar.divider()
    st.sidebar.subheader("💾 数据导出")
    
    csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
    st.sidebar.download_button(
        label="导出CSV",
        data=csv,
        file_name=f"合同数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime='text/csv'
    )

if __name__ == "__main__":
    main()
