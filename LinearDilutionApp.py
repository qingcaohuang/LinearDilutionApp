import streamlit as st
import pandas as pd
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import os
from datetime import datetime
import io

# --- 0. 版本号 ---
VERSION = "v1.3.1"

# --- 1. 基础工具函数 (完全不动) ---
def get_densities(temp):
    """根据温度输出纯水和生理盐水(0.9% NaCl)的密度 (g/cm3)"""
    rho_water = 1000 * (1 - (temp + 288.9414) / (508929.2 * (temp + 68.12963)) * (temp - 3.9863)**2)
    rho_water_g = round(rho_water / 1000, 5)
    rho_saline_g = round(rho_water_g * 1.0064, 4) 
    return rho_water_g, rho_saline_g

def calc_theoretical_masses(tc, tm, c_h, rho_h, c_l, rho_l):
    """计算理论质量，确保非负 [0, tm]"""
    if tc >= c_h: return tm, 0.0
    if tc <= c_l: return 0.0, tm
    k1 = (c_h - tc) / rho_h
    k2 = (tc - c_l) / rho_l
    if (k1 + k2) == 0: return 0.0, tm
    m_h = (tm * k2) / (k1 + k2)
    m_h = max(0.0, min(float(m_h), float(tm)))
    return m_h, tm - m_h

def calc_actual_volume_conc(m_h, m_l, c_h, rho_h, c_l, rho_l):
    """回算实际体积浓度"""
    v_h = m_h / rho_h
    v_l = m_l / rho_l
    if (v_h + v_l) == 0: return 0.0
    return (v_h * c_h + v_l * c_l) / (v_h + v_l)

# --- 2. PDF 生成类 (带页脚) ---
class PDFWithFooter(FPDF):
    def __init__(self, version, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.version = version

    def footer(self):
        self.set_y(-15)
        # --- 修改1：移除重复的 add_font，直接使用 set_font ---
        # 字体已经在 create_pdf 中被添加，这里只需检查文件存在性来决定是否使用
        current_dir = os.path.dirname(os.path.abspath(__file__))
        font_path = os.path.join(current_dir, "font.ttf")
        
        if os.path.exists(font_path): 
            # 假设外部已注册名为 'Font'
            self.set_font('Font', '', 8)
        else: 
            self.set_font('Arial', 'I', 8)
            
        self.set_text_color(150, 150, 150)
        
        # 版本信息
        if os.path.exists(font_path):
            version_text = f"版本: {self.version} | 程序创建者：Rong"
        else:
            version_text = f"Version: {self.version} | Creator: Rong"
            
        self.cell(0, 10, text=version_text, align='R')
        self.set_y(-15)
        
        # 页码信息
        if os.path.exists(font_path):
            page_num_text = f"第 {self.page_no()} 页"
        else:
            page_num_text = f"Page {self.page_no()}"
            
        self.cell(0, 10, text=page_num_text, align='L')

def create_pdf(df_main, df_mid, title, meta_info):
    version = meta_info.pop("程序版本", "N/A")
    pdf = PDFWithFooter(version=version)
    
    # --- 修改2：设置页边距 (左=25mm, 上=20mm, 右=20mm) ---
    # 默认是10mm。左边距增大2.5倍 -> 25mm，右边距增大1倍(通常指翻倍) -> 20mm
    pdf.set_margins(left=25, top=20, right=20)
    
    pdf.add_page()
    
    # 字体加载逻辑
    current_dir = os.path.dirname(os.path.abspath(__file__))
    font_path = os.path.join(current_dir, "font.ttf")
    font_ok = False
    
    if os.path.exists(font_path):
        pdf.add_font('Font', '', font_path) # 在这里注册一次即可
        pdf.set_font('Font', size=16)
        font_ok = True
    else: 
        pdf.set_font('Arial', size=16)
        pdf.set_text_color(255, 0, 0)
        pdf.cell(0, 10, text="Warning: font.ttf not found.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_text_color(0, 0, 0)

    # 1. 标题 (宽度设为0表示利用剩余宽度，align='C'会自动在margin之间居中)
    display_title = title if font_ok else "Linear Dilution Report"
    pdf.cell(0, 10, text=display_title, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
    pdf.ln(5)
    
    # 2. 元数据
    if font_ok: pdf.set_font('Font', size=10)
    else: pdf.set_font('Arial', size=10)
    
    # 计算有效宽度：A4宽(210) - 左边距(15) - 右边距(20) = 175
    effective_page_width = pdf.w - pdf.l_margin - pdf.r_margin
    col_width_meta = effective_page_width / 2
    
    items = list(meta_info.items())
    for i in range(0, len(items), 2):
        k1, v1 = items[i]
        if not font_ok: k1 = "Item"; v1 = str(v1).encode('ascii', 'ignore').decode('ascii')
        pdf.cell(col_width_meta, 8, text=f"{k1}: {v1}", new_x=XPos.RIGHT, new_y=YPos.TOP)
        
        if i + 1 < len(items):
            k2, v2 = items[i+1]
            if not font_ok: k2 = "Item"; v2 = str(v2).encode('ascii', 'ignore').decode('ascii')
            pdf.cell(col_width_meta, 8, text=f"{k2}: {v2}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        else: pdf.ln(8)
        
    pdf.ln(4)
    # 画线：从左边距开始，到 (页宽-右边距) 结束
    pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
    pdf.ln(5)
    
    # 3. 中间浓度表
    if font_ok: 
        pdf.set_font('Font', size=11)
        pdf.cell(0, 10, text="一、中间浓度配置详情", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font('Font', size=10)
    else: 
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(0, 10, text="1. Intermediate Prep Details", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font('Arial', size=10)
        
    col_width_mid = effective_page_width / len(df_mid.columns)
    pdf.set_fill_color(245, 245, 245)
    for col in df_mid.columns: 
        txt = str(col) if font_ok else "Col"
        pdf.cell(col_width_mid, 8, text=txt, border=1, align='C', fill=True)
    pdf.ln()
    for _, row in df_mid.iterrows():
        for i, item in enumerate(row):
            if isinstance(item, float): val = f"{item:.2f}" if "浓度" in df_mid.columns[i] else f"{item:.1f}"
            else: val = str(item)
            if not font_ok: val = str(val).encode('ascii', 'ignore').decode('ascii')
            pdf.cell(col_width_mid, 8, text=val, border=1, align='C')
        pdf.ln()
    pdf.ln(10)
    
    # 4. 梯度表
    if font_ok: 
        pdf.set_font('Font', size=11)
        pdf.cell(0, 10, text="二、分段梯度稀释方案", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font('Font', size=10)
    else: 
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(0, 10, text="2. Gradient Dilution Plan", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font('Arial', size=10)
        
    cols = df_main.columns.tolist()
    col_width = effective_page_width / len(cols)
    pdf.set_fill_color(235, 235, 235)
    for col in cols: 
        txt = str(col) if font_ok else "Col"
        pdf.cell(col_width, 10, text=txt, border=1, align='C', fill=True)
    pdf.ln()
    for _, row in df_main.iterrows():
        for i, item in enumerate(row):
            if i == 0: val = str(int(item))
            elif isinstance(item, (int, float)):
                if "浓度" in cols[i]: val = f"{item:.2f}"
                else: val = f"{item:.1f}"
            else: val = str(item)
            if not font_ok: val = str(val).encode('ascii', 'ignore').decode('ascii')
            pdf.cell(col_width, 10, text=val, border=1, align='C')
        pdf.ln()
    return pdf.output()

# --- 3. 界面初始化 ---
st.set_page_config(page_title="线性评价样本制备程序", layout="wide")

# CSS: 1. 增加左侧栏宽度; 2. 移除主界面顶部的巨大留空
st.markdown("""
    <style>
        [data-testid="stSidebar"] { min-width: 500px; max-width: 500px; }
        .block-container { padding-top: 1.5rem; }
        h4 { margin-top: 0rem !important; margin-bottom: 0.2rem !important; }
    </style>
    """, unsafe_allow_html=True)

# 将主标题和小标题放在同一行，并控制字号
st.markdown("""
    <h2>🧪 体外诊断线性材料配制程序 
    <span style="font-size: 0.65em; font-weight: normal; color: #666;">— 适用称重稀释法</span>
    </h4>
    """, unsafe_allow_html=True)
st.caption(f"版本: {VERSION}")

# XLSX 导入逻辑
import_data = {}
with st.sidebar.expander("📂 导入 XLSX 存档", expanded=False):
    uploaded_file = st.file_uploader("导入 XLSX 存档", type="xlsx", label_visibility="collapsed")
    if uploaded_file:
        try:
            df_settings = pd.read_excel(uploaded_file, sheet_name="配置参数")
            import_data = dict(zip(df_settings["参数"], df_settings["数值"]))
            df_grad_import = pd.read_excel(uploaded_file, sheet_name="梯度方案")
            st.success("存档导入成功！")
        except Exception as e:
            st.error(f"导入失败: {e}")

with st.sidebar:
    st.subheader("⚙️ 基础设置") 
    current_date_str = datetime.now().strftime("%Y%m%d")
    final_exp_name = import_data.get("实验内容")
    if not final_exp_name and uploaded_file: final_exp_name = os.path.splitext(uploaded_file.name)[0]
    if not final_exp_name: final_exp_name = f"线性稀释实验-{current_date_str}"
    exp_name = st.text_input("实验内容名称", value=str(final_exp_name))
    
    c_u1, c_u2 = st.columns(2)
    unit_conc = c_u1.text_input("浓度单位", value=import_data.get("浓度单位", "mg/L"))
    unit_mass = c_u2.text_input("质量单位", value=import_data.get("质量单位", "mg"))
    
    input_temp = st.number_input("环境温度 (°C)", value=float(import_data.get("环境温度", 22.0)), step=0.5)
    rho_w, rho_s = get_densities(input_temp)
    c_d1, c_d2 = st.columns(2)
    c_d1.write(f"💧 **纯水密度**: `{rho_w}`"); c_d2.write(f"🏥 **生理盐水**: `{rho_s}`")
    
    st.markdown("### 材料参数")
    c_p1, c_p2 = st.columns(2)
    c_h_orig = c_p1.number_input(f"高浓度材料浓度", value=float(import_data.get("原液浓度A", 100.0)))
    rho_h_orig = c_p2.number_input("高浓度材料密度 (g/cm3)", value=float(import_data.get("原液密度A", 1.0500)), format="%.4f")
    c_p3, c_p4 = st.columns(2)
    c_l_orig = c_p3.number_input(f"低浓度材料浓度", value=float(import_data.get("原液浓度B", 0.0)))
    rho_l_orig = c_p4.number_input("低浓度材料密度 (g/cm3)", value=float(import_data.get("原液密度B", 1.0000)), format="%.4f")

    c_p5, c_p6 = st.columns(2)
    num_points = c_p5.number_input("稀释点数量 (含端点)", min_value=3, max_value=20, value=int(import_data.get("样本点数", 8)), step=1)
    target_tm_each = c_p6.number_input(f"各梯度点配置量 ({unit_mass})", value=float(import_data.get("单点计划总量",350)), step=5.0)

# --- 4. 预计算中间浓度 (完全不动) ---
target_c_mid_guess = round((c_h_orig + c_l_orig)/2, 2)
mid_idx_guess = num_points // 2
pts_low_temp = [c_l_orig + i * ((target_c_mid_guess - c_l_orig) / mid_idx_guess) for i in range(mid_idx_guess)]
pts_high_temp = [target_c_mid_guess + i * ((c_h_orig - target_c_mid_guess) / (num_points - mid_idx_guess - 1)) for i in range(num_points - mid_idx_guess)]
all_targets_temp = pts_low_temp + pts_high_temp
total_mid_usage_theo = 0.0
for t_conc in all_targets_temp:
    if t_conc > target_c_mid_guess + 0.0001: _, m_mid_needed = calc_theoretical_masses(t_conc, target_tm_each, c_h_orig, rho_h_orig, target_c_mid_guess, 1.0)
    else: m_mid_needed, _ = calc_theoretical_masses(t_conc, target_tm_each, target_c_mid_guess, 1.0, c_l_orig, rho_l_orig)
    total_mid_usage_theo += m_mid_needed
suggested_prep_m = round(total_mid_usage_theo * 1.1, 1)

# --- 5. 步骤一：中间配置 (降档标题 + 并列布局) ---
st.markdown("#### 1️⃣ 中间浓度配置") # 标题缩小一档
with st.container(border=True):
    col_left, col_right = st.columns(2)
    with col_left:
        target_c_mid = st.number_input(f"中间目标浓度 ({unit_conc})", value=float(import_data.get("中间目标浓度", target_c_mid_guess)), step=0.1)
        prep_m_mid = st.number_input(f"配置总质量 (建议: 总需求×1.1)", value=float(import_data.get("中间计划总量", max(suggested_prep_m, 100.0))), step=10.0)
        m_h_theo, m_l_theo = calc_theoretical_masses(target_c_mid, prep_m_mid, c_h_orig, rho_h_orig, c_l_orig, rho_l_orig)
        st.info(f"💡 建议：高浓度材料 {m_h_theo:.1f} + 低浓度材料 {m_l_theo:.1f} (理论用量和: {total_mid_usage_theo:.1f})")
    with col_right:
        m_h_mid_act = st.number_input("加入高浓度材料 (实测质量)", value=float(import_data.get("中间实测A", round(m_h_theo, 1))), min_value=0.0, step=0.1, format="%.1f", key="mid_h_val")
        m_l_mid_act = st.number_input("加入低浓度材料 (实测质量)", value=float(import_data.get("中间实测B", round(m_l_theo, 1))), min_value=0.0, step=0.1, format="%.1f", key="mid_l_val")
        actual_c_mid = calc_actual_volume_conc(m_h_mid_act, m_l_mid_act, c_h_orig, rho_h_orig, c_l_orig, rho_l_orig)
        denom = (m_h_mid_act/rho_h_orig) + (m_l_mid_act/rho_l_orig)
        actual_rho_mid = (m_h_mid_act + m_l_mid_act) / denom if denom > 0 else 1.0
        st.warning(f"🧪 **中间浓度实际参数**：浓度 **{actual_c_mid:.2f}**，密度 **{actual_rho_mid:.4f}**")

# --- 6. 步骤二：分段梯度稀释方案 (降档标题 + 字体大小一致) ---
st.markdown("#### 2️⃣ 分段梯度稀释方案") # 标题缩小一档
mid_idx = num_points // 2
pts_low = [c_l_orig + i * ((actual_c_mid - c_l_orig) / mid_idx) for i in range(mid_idx)]
pts_high = [actual_c_mid + i * ((c_h_orig - actual_c_mid) / (num_points - mid_idx - 1)) for i in range(num_points - mid_idx)]
all_targets = pts_low + pts_high
h_cols = st.columns([0.5, 1.2, 1.2, 1.2, 1.2, 1.2, 1.2])
headers = ["序号", "目标浓度", "材料A", "材料B", "加入A质量", "加入B质量", "实际浓度"]
for col, lab in zip(h_cols, headers): col.write(f"**{lab}**")

results_data = []
total_high_used = m_h_mid_act
total_low_used = m_l_mid_act
for i, t_conc in enumerate(all_targets):
    idx = i + 1
    if t_conc > actual_c_mid + 0.0001: m_a_name, m_b_name, ca, ra, cb, rb = "高浓度", "中间浓度", c_h_orig, rho_h_orig, actual_c_mid, actual_rho_mid
    else: m_a_name, m_b_name, ca, ra, cb, rb = "中间浓度", "低浓度", actual_c_mid, actual_rho_mid, c_l_orig, rho_l_orig
    imp_tc, imp_ma, imp_mb = t_conc, None, None
    if uploaded_file and 'df_grad_import' in locals():
        if i < len(df_grad_import): imp_tc, imp_ma, imp_mb = df_grad_import.iloc[i]["目标浓度"], df_grad_import.iloc[i]["加入A质量"], df_grad_import.iloc[i]["加入B质量"]
    r_cols = st.columns([0.5, 1.2, 1.2, 1.2, 1.2, 1.2, 1.2])
    r_cols[0].write(f"{idx}")
    row_tc = r_cols[1].number_input(f"tc_{i}", value=float(imp_tc), format="%.2f", key=f"row_tc_{i}", label_visibility="collapsed", step=0.1)
    # 使用 st.write 确保字体大小与主界面一致
    r_cols[2].write(m_a_name); r_cols[3].write(m_b_name)
    m_a_theo_row, m_b_theo_row = calc_theoretical_masses(row_tc, target_tm_each, ca, ra, cb, rb)
    row_ma = r_cols[4].number_input(f"ma_{i}", value=float(imp_ma if imp_ma is not None else round(m_a_theo_row, 1)), min_value=0.0, step=0.1, format="%.1f", key=f"row_ma_{i}", label_visibility="collapsed")
    row_mb = r_cols[5].number_input(f"mb_{i}", value=float(imp_mb if imp_mb is not None else round(m_b_theo_row, 1)), min_value=0.0, step=0.1, format="%.1f", key=f"row_mb_{i}", label_visibility="collapsed")
    row_act_c = calc_actual_volume_conc(row_ma, row_mb, ca, ra, cb, rb)
    r_cols[6].write(f"**{row_act_c:.2f}**")
    results_data.append({"序号": idx, "目标浓度": row_tc, "材料A": m_a_name, "材料B": m_b_name, "加入A质量": row_ma, "加入B质量": row_mb, "最终实际浓度": row_act_c})
    if m_a_name == "高浓度": total_high_used += row_ma
    if m_b_name == "低浓度": total_low_used += row_mb

# --- 7. 数据导出区域 (完全不动) ---
c_x1, c_x2 = st.columns(2)
with c_x1:
    if st.button("💾 导出 XLSX 存档"):
        settings_dict = {
            "程序版本": VERSION, "实验内容": exp_name, "浓度单位": unit_conc, "质量单位": unit_mass, "环境温度": input_temp,
            "原液浓度A": c_h_orig, "原液密度A": rho_h_orig, "原液浓度B": c_l_orig, "原液密度B": rho_l_orig,
            "样本点数": num_points, "单点计划总量": target_tm_each, "中间目标浓度": target_c_mid, "中间计划总量": prep_m_mid,
            "中间实测A": m_h_mid_act, "中间实测B": m_l_mid_act
        }
        df_settings = pd.DataFrame(list(settings_dict.items()), columns=["参数", "数值"])
        df_grad = pd.DataFrame(results_data)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_settings.to_excel(writer, sheet_name="配置参数", index=False)
            df_grad.to_excel(writer, sheet_name="梯度方案", index=False)
        st.download_button("📥 点击下载 XLSX", data=output.getvalue(), file_name=f"{exp_name}_{datetime.now().strftime('%H%M')}.xlsx")
with c_x2:
    if st.button("📑 生成实验 PDF 报告"):
        mid_prep_df = pd.DataFrame([
            {"组分": "高浓度材料", "理论质量": m_h_theo, "加入质量": m_h_mid_act, "目标浓度": "-", "实际配置浓度": "-"},
            {"组分": "低浓度材料", "理论质量": m_l_theo, "加入质量": m_l_mid_act, "目标浓度": "-", "实际配置浓度": "-"},
            {"组分": "合计(中间浓度材料)", "理论质量": m_h_theo + m_l_theo, "加入质量": m_h_mid_act + m_l_mid_act, "目标浓度": target_c_mid, "实际配置浓度": actual_c_mid}
        ])
        mid_prep_df.columns = ["组分", f"理论质量({unit_mass})", f"加入质量({unit_mass})", f"目标浓度({unit_conc})", f"实际配置浓度({unit_conc})"]
        pdf_meta = {
            "程序版本": VERSION, "实验内容": exp_name, "环境温度": f"{input_temp} degC", "水密度": f"{rho_w} g/cm3",
            "生理盐水密度": f"{rho_s} g/cm3", "高浓度材料": f"{c_h_orig} {unit_conc} (密度:{rho_h_orig:.4f})",
            "低浓度材料": f"{c_l_orig} {unit_conc} (密度:{rho_l_orig:.4f})", "中间浓度材料": f"{actual_c_mid:.2f} {unit_conc} (密度:{actual_rho_mid:.4f})",
            "高浓度材料合计量": f"{total_high_used:.1f} {unit_mass}", "低浓度材料合计量": f"{total_low_used:.1f} {unit_mass}",
            "导出时间": datetime.now().strftime("%Y-%m-%d %H:%M")
        }
        pdf_out = create_pdf(pd.DataFrame(results_data), mid_prep_df, "线性评价样本制备记录", pdf_meta)
        st.download_button("📥 点击下载 PDF", data=bytes(pdf_out), file_name=f"Report_{exp_name}_{datetime.now().strftime('%H%M')}.pdf", mime="application/pdf")