import streamlit as st
import pandas as pd
import numpy as np
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from datetime import datetime
import plotly.express as px
import openpyxl

# ==================== 页面配置 ====================
st.set_page_config(
    page_title="跨境电商加货计算系统",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 自定义样式 ====================
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 初始化 Session State ====================
if 'history' not in st.session_state:
    st.session_state.history = []
if 'last_result' not in st.session_state:
    st.session_state.last_result = None

# ==================== 侧边栏 ====================
with st.sidebar:
    st.markdown("## ⚙️ 参数设置")
    
    weeks_threshold = st.number_input(
        "可售周数阈值",
        min_value=1,
        max_value=52,
        value=24,
        step=1
    )
    
    st.markdown("---")
    st.markdown("### 🚚 物流时效（天）")
    
    lead_times = {
        '美西': st.number_input("美西", min_value=1, max_value=60, value=22, step=1),
        '美东': st.number_input("美东", min_value=1, max_value=60, value=24, step=1),
        '芝加哥': st.number_input("芝加哥", min_value=1, max_value=60, value=22, step=1),
        '美南': st.number_input("美南", min_value=1, max_value=60, value=23, step=1),
    }
    
    st.markdown("---")
    st.markdown("### ✉️ 邮件发送")
    email_to = st.text_input("收件邮箱", placeholder="example@email.com")
    send_email_btn = st.button("📧 发送邮件", type="primary", use_container_width=True)

# ==================== 主区域 ====================
st.markdown('<p class="main-header">📦 跨境电商加货计算系统</p>', unsafe_allow_html=True)
st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📤 上传发货数据")
    file_main = st.file_uploader(
        "选择发货系统数据文件",
        type=['xlsx', 'xls', 'csv'],
        key="file_main"
    )

with col2:
    st.markdown("### 📤 上传周动销数据")
    file_sales = st.file_uploader(
        "选择周动销数据文件",
        type=['xlsx', 'xls', 'csv'],
        key="file_sales"
    )

# ==================== 数据处理函数 ====================
def load_main_file(file):
    """加载发货数据文件 - 完全按照原始 jiaohuo.py 的逻辑"""
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        sheet = wb['汽摩配汇总']
        
        # 列映射 - 与原始代码完全一致
        columns_map = {
            '包裹编码': 'A',
            '版本SKU': 'C',
            '理论-美西': 'M',
            '理论-美东': 'N',
            '理论-芝加哥': 'O',
            '理论-美南': 'P',
            '实际-美西': 'Q',
            '实际-美东': 'R',
            '实际-芝加哥': 'S',
            '实际-美南': 'T',
            '美国FBM库存': 'DH',
            '国内可发库存': 'HU'
        }
        
        # 从第4行开始提取
        data_rows = []
        for row in range(4, sheet.max_row + 1):
            row_data = {}
            for name, col in columns_map.items():
                val = sheet[f'{col}{row}'].value
                if val is None or val == '':
                    row_data[name] = None
                else:
                    row_data[name] = val
            if row_data.get('包裹编码') and str(row_data['包裹编码']).strip():
                data_rows.append(row_data)
        
        wb.close()
        
        if len(data_rows) == 0:
            st.error("❌ 未提取到任何数据")
            return None
        
        df_extracted = pd.DataFrame(data_rows)
        
        # 转换数值列
        numeric_cols = ['理论-美西', '理论-美东', '理论-芝加哥', '理论-美南',
                        '实际-美西', '实际-美东', '实际-芝加哥', '实际-美南',
                        '美国FBM库存', '国内可发库存']
        for col in numeric_cols:
            if col in df_extracted.columns:
                df_extracted[col] = pd.to_numeric(df_extracted[col], errors='coerce').fillna(0)
        
        return df_extracted
        
    except Exception as e:
        st.error(f"读取发货文件失败: {e}")
        return None

def load_sales_file(file):
    """加载周动销数据文件"""
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        
        # 查找"包裹编码"列
        for col in df.columns:
            if '包裹' in str(col) or '编码' in str(col):
                df.rename(columns={col: '包裹编码'}, inplace=True)
                break
        
        # 查找"周动销"列
        for col in df.columns:
            if '动销' in str(col) or '周' in str(col):
                df.rename(columns={col: '周动销'}, inplace=True)
                break
        
        return df
    except Exception as e:
        st.error(f"读取周动销文件失败: {e}")
        return None

def calculate_restock(df_main, df_sales, weeks_threshold, lead_times):
    """执行加货计算 - 完全按照原始 jiaohuo.py 的逻辑"""
    
    if df_main is None or len(df_main) == 0:
        return None, "❌ 主数据为空"
    
    if df_sales is None or len(df_sales) == 0:
        return None, "❌ 周动销数据为空"
    
    # 合并数据
    df = df_main.merge(
        df_sales[['包裹编码', '周动销']],
        on='包裹编码',
        how='left'
    )
    df['周动销'] = pd.to_numeric(df['周动销'], errors='coerce').fillna(0)
    
    # 计算可售周数
    def calc_weeks(row):
        if row['周动销'] > 0 and row['美国FBM库存'] > 0:
            return row['美国FBM库存'] / row['周动销']
        elif row['周动销'] <= 0 and row['美国FBM库存'] > 0:
            return 999
        else:
            return 999
    
    df['可售周数'] = df.apply(calc_weeks, axis=1)
    
    # 筛选可售周数 < 阈值
    df_filtered = df[df['可售周数'] < weeks_threshold].copy()
    
    if len(df_filtered) == 0:
        return None, f"⚠️ 没有可售周数低于 {weeks_threshold} 的包裹"
    
    # 计算需求量
    df_filtered['需求量'] = (weeks_threshold - df_filtered['可售周数']) * df_filtered['周动销']
    
    # 计算发货总量
    df_filtered['发货总量'] = df_filtered.apply(
        lambda row: min(row['国内可发库存'], row['需求量'])
        if row['国内可发库存'] > 0 and row['需求量'] > 0 else 0,
        axis=1
    )
    
    # 计算四仓应发
    warehouses = ['美西', '美东', '芝加哥', '美南']
    theory_cols = ['理论-美西', '理论-美东', '理论-芝加哥', '理论-美南']
    actual_cols = ['实际-美西', '实际-美东', '实际-芝加哥', '实际-美南']
    
    for wh, theory_col, actual_col in zip(warehouses, theory_cols, actual_cols):
        lt = lead_times[wh]
        df_filtered[f'应发-{wh}'] = df_filtered.apply(
            lambda row, wh=wh, theory_col=theory_col, actual_col=actual_col, lt=lt:
            row['周动销'] * row[theory_col] * lt - row['美国FBM库存'] * row[actual_col],
            axis=1
        )
        df_filtered[f'应发-{wh}'] = df_filtered[f'应发-{wh}'].clip(lower=0)
    
    # 计算最终四仓发货
    df_filtered['应发总和'] = df_filtered[[f'应发-{wh}' for wh in warehouses]].sum(axis=1)
    
    for wh in warehouses:
        df_filtered[f'比例-{wh}'] = np.where(
            df_filtered['应发总和'] > 0,
            df_filtered[f'应发-{wh}'] / df_filtered['应发总和'],
            0
        )
        df_filtered[f'最终发货-{wh}'] = df_filtered['发货总量'] * df_filtered[f'比例-{wh}']
    
    # 整理输出列
    original_cols = ['包裹编码', '版本SKU', '理论-美西', '理论-美东', '理论-芝加哥', '理论-美南',
                     '实际-美西', '实际-美东', '实际-芝加哥', '实际-美南', '美国FBM库存', '国内可发库存']
    calc_cols = ['周动销', '可售周数', '需求量', '发货总量']
    should_cols = [f'应发-{wh}' for wh in warehouses]
    ratio_cols = [f'比例-{wh}' for wh in warehouses]
    final_cols = [f'最终发货-{wh}' for wh in warehouses]
    
    all_output_cols = original_cols + calc_cols + should_cols + ratio_cols + final_cols
    df_output = df_filtered[all_output_cols].copy()
    
    # 四舍五入
    int_cols = ['需求量', '发货总量'] + should_cols + final_cols
    for col in int_cols:
        if col in df_output.columns:
            df_output[col] = df_output[col].round(0).astype(int)
    
    for col in ratio_cols:
        if col in df_output.columns:
            df_output[col] = df_output[col].round(4)
    
    return df_output, f"✅ 计算完成！共 {len(df_output)} 个包裹需要加货"

def send_email_with_attachment(to_email, df_result):
    """发送邮件带附件"""
    try:
        smtp_server = os.getenv('SMTP_SERVER', 'smtp.qq.com')
        smtp_port = int(os.getenv('SMTP_PORT', '587'))
        sender_email = os.getenv('SENDER_EMAIL', '')
        sender_password = os.getenv('SENDER_PASSWORD', '')
        
        if not sender_email or not sender_password:
            return False, "❌ 请配置发件邮箱和授权码"
        
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Subject'] = f"加货计算结果 - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        
        body = f"""
您好！

这是您的加货计算结果，生成时间为 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

结果摘要：
- 总包裹数：{len(df_result)}
- 总发货量：{df_result['发货总量'].sum()}
- 美西发货：{df_result['最终发货-美西'].sum()}
- 美东发货：{df_result['最终发货-美东'].sum()}
- 芝加哥发货：{df_result['最终发货-芝加哥'].sum()}
- 美南发货：{df_result['最终发货-美南'].sum()}

此邮件由 跨境电商加货计算系统 自动发送
"""
        
        msg.attach(MIMEText(body, 'plain'))
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='加货结果')
        buffer.seek(0)
        
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(buffer.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename=加货结果_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
        )
        msg.attach(part)
        
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        
        return True, f"✅ 邮件已发送至 {to_email}"
        
    except Exception as e:
        return False, f"❌ 发送失败: {str(e)}"

# ==================== 主要逻辑 ====================
if file_main is not None and file_sales is not None:
    
    with st.spinner("正在加载数据..."):
        df_main = load_main_file(file_main)
        df_sales = load_sales_file(file_sales)
    
    if df_main is not None and df_sales is not None:
        
        st.markdown("---")
        st.markdown("### 📋 数据预览")
        
        tab1, tab2 = st.tabs(["📄 发货数据", "📄 周动销数据"])
        
        with tab1:
            st.dataframe(df_main.head(10), use_container_width=True)
            st.caption(f"共 {len(df_main)} 行")
        
        with tab2:
            st.dataframe(df_sales.head(10), use_container_width=True)
            st.caption(f"共 {len(df_sales)} 行")
        
        st.markdown("---")
        if st.button("🚀 执行加货计算", type="primary", use_container_width=True):
            
            with st.spinner("正在执行计算..."):
                result, message = calculate_restock(
                    df_main,
                    df_sales,
                    weeks_threshold,
                    lead_times
                )
                
                if result is None:
                    st.warning(message)
                else:
                    st.success(message)
                    
                    st.session_state.last_result = result
                    
                    st.session_state.history.append({
                        'time': datetime.now().strftime('%Y-%m-%d %H:%M'),
                        'rows': len(result),
                        'total': result['发货总量'].sum()
                    })
                    
                    # ===== 汇总统计 =====
                    st.markdown("---")
                    st.markdown("### 📊 汇总统计")
                    
                    col1, col2, col3, col4, col5 = st.columns(5)
                    with col1:
                        st.metric("📦 总包裹数", len(result))
                    with col2:
                        st.metric("📦 总发货量", f"{result['发货总量'].sum():.0f}")
                    with col3:
                        st.metric("🏷️ 美西", f"{result['最终发货-美西'].sum():.0f}")
                    with col4:
                        st.metric("🏷️ 美东", f"{result['最终发货-美东'].sum():.0f}")
                    with col5:
                        st.metric("🏷️ 芝加哥", f"{result['最终发货-芝加哥'].sum():.0f}")
                    
                    st.metric("🏷️ 美南", f"{result['最终发货-美南'].sum():.0f}")
                    
                    # ===== 可视化 =====
                    st.markdown("---")
                    st.markdown("### 📈 可视化分析")
                    
                    warehouse_totals = {
                        '仓库': ['美西', '美东', '芝加哥', '美南'],
                        '发货量': [
                            result['最终发货-美西'].sum(),
                            result['最终发货-美东'].sum(),
                            result['最终发货-芝加哥'].sum(),
                            result['最终发货-美南'].sum()
                        ]
                    }
                    fig = px.bar(
                        pd.DataFrame(warehouse_totals),
                        x='仓库',
                        y='发货量',
                        title='各仓库发货量对比',
                        color='仓库',
                        text='发货量'
                    )
                    fig.update_traces(texttemplate='%{text:.0f}', textposition='outside')
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # ===== 显示结果表格 =====
                    st.markdown("---")
                    st.markdown("### 📋 计算结果详情")
                    
                    display_cols = ['包裹编码', '版本SKU', '周动销', '美国FBM库存', '可售周数',
                                   '需求量', '国内可发库存', '发货总量',
                                   '最终发货-美西', '最终发货-美东', '最终发货-芝加哥', '最终发货-美南']
                    display_cols = [col for col in display_cols if col in result.columns]
                    
                    st.dataframe(
                        result[display_cols],
                        use_container_width=True,
                        height=400
                    )
                    
                    # ===== 导出结果 =====
                    st.markdown("---")
                    st.markdown("### 💾 导出结果")
                    
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        result.to_excel(writer, index=False, sheet_name='加货结果')
                    excel_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 下载 Excel 文件",
                        data=excel_buffer.getvalue(),
                        file_name=f"发货计算结果_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )
                    
                    # ===== 发送邮件 =====
                    if send_email_btn and email_to:
                        with st.spinner("正在发送邮件..."):
                            success, msg = send_email_with_attachment(email_to, result)
                            if success:
                                st.success(msg)
                            else:
                                st.error(msg)
                    elif send_email_btn and not email_to:
                        st.warning("⚠️ 请先输入收件邮箱地址")
                    
                    # ===== 历史记录 =====
                    if len(st.session_state.history) > 0:
                        st.markdown("---")
                        st.markdown("### 📜 历史记录")
                        history_df = pd.DataFrame(st.session_state.history)
                        st.dataframe(history_df, use_container_width=True)
else:
    st.info("👆 请在上方上传发货数据和周动销数据文件")

st.markdown("---")
st.caption("📦 跨境电商加货计算系统 v1.0")
