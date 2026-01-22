import streamlit as st
import pandas as pd
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import os
from datetime import datetime
import io

# --- 0. 版本号 ---
VERSION = "v1.3.2"

# --- 1. 基础工具函数 ---
def get_densities(temp):
    """根据温度输出纯水和生理盐水(0.9% NaCl)的密度 (g/cm3)"""
    rho_water = 1000 * (1 - (temp + 288.9414) / (508929.2 * (temp + 68.12963)) * (temp - 3.9863)**2)
    rho_water_g = round(rho_water / 1000, 5)
    # 生理盐水密度换算 (0.9% NaCl)
    rho_saline_g = round(rho_water_g * 1.0064, 5) 
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

# --- 2. PDF 生成类 ---
class PDFWithFooter(FPDF):
    def __init__(self, version, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.version = version

    def footer(self):
        self.set_y(-15)
        current_dir = os.path.dirname(os.path.abspath(__file__))
        font_path = os.path.join(current_dir, "font.ttf")
        if os.path.exists(font_path):
            self.set_font('Font', '', 8)
            version_text = f"版本: {self.version} | 程序创建者：Rong | 第 {self.page_no()} 页"
        else:
            self.set_font('Arial', 'I', 8)
            version_text = f"Version: {self.version} | Creator: Rong | Page {self.page_no()}"
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, text=version_text, align='R')

def create_pdf(df_main, df_mid, title, meta_info):
    version = meta_info.get("程序版本", "N/A")
    pdf = PDFWithFooter(version=version)

    # 设置页边距：左25mm, 上20mm, 右20mm
    pdf.set_margins(left=25, top=20, right=20)
    pdf.add_page()
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    font_path = os.path.join(current_dir, "font.ttf")
    font_ok = False
    if os.path.exists(font_path):
        pdf.add_font('Font', '', font_path)
        pdf.set_font('Font', size=16)
        font_ok = True
    else:
        pdf.set_font('Arial', size=16)

    # 1. 标题
    pdf.cell(0, 10, text=title if font_ok else "Linear Dilution Report", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
    pdf.ln(5)
    
    # 2. 元数据
    pdf.set_font('Font' if font_ok else 'Arial', size=10)
    effective_width = pdf.w - pdf.l_margin - pdf.r_margin
    items = list(meta_info.items())
    for i in range(0, len(items), 2):
        k1, v1 = items[i]
        pdf.cell(effective_width/2, 8, text=f"{k1}: {v1}", new_x=XPos.RIGHT, new_y=YPos.TOP)
        if i + 1 < len(items):
            k2, v2 = items[i+1]
            pdf.cell(effective_width/2, 8, text=f"{k2}: {v2}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        else: pdf.ln(8)
            
    pdf.ln(4)
    pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
    pdf.ln(5)

    # 3. 中间浓度配置详情
    pdf.set_font('Font' if font_ok else 'Arial', size=11)
    pdf.cell(0, 10, text="一、中间浓度配置详情" if font_ok else "1. Intermediate Prep", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.set_font('Font' if font_ok else 'Arial', size=9)
    col_width_mid = effective_width / len(df_mid.columns)
    pdf.set_fill_color(245, 245, 245)
    for col in df_mid.columns:
        pdf.cell(col_width_mid, 8, text=str(col), border=1, align='C', fill=True)
    pdf.ln()
    for _, row in df_mid.iterrows():
        for i, item in enumerate(row):
            # 修复点：先检查是否为数字，再根据列名判断保留位数
            if isinstance(item, (int, float)):
                val = f"{item:.2f}" if "浓度" in df_mid.columns[i] else f"{item:.1f}"
            else:
                val = str(item)
            pdf.cell(col_width_mid, 8, text=val, border=1, align='C')
        pdf.ln()
    pdf.ln(10)

    # 4. 分段梯度稀释方案
    pdf.set_font('Font' if font_ok else 'Arial', size=11)
    pdf.cell(0, 10, text="二、分段梯度稀释方案" if font_ok else "2. Gradient Plan", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font('Font' if font_ok else 'Arial', size=9)
    cols = df_main.columns.tolist()
    col_width = effective_width / len(cols)
    pdf.set_fill_color(235, 235, 235)
    for col in cols:
        pdf.cell(col_width, 10, text=str(col), border=1, align='C', fill=True)
    pdf.ln()
    for _, row in df_main.iterrows():
        for i, item in enumerate(row):
            if i == 0: val = str(int(item))
            elif isinstance(item, (int, float)):
                val = f"{item:.2f}" if "浓度" in cols[i] else f"{item:.1f}"
            else: val = str(item)
            pdf.cell(col_width, 10, text=val, border=1, align='C')
        pdf.ln()
    return pdf.output()

# --- 3. 界面初始化 ---
st.set_page_config(page_title="线性评价样本制备程序", layout="wide")

st.markdown("""
    <style>
        [data-testid="stSidebar"] { min-width: 450px; max-width: 450px; }
        .block-container { padding-top: 1.5rem; }
    </style>
    """, unsafe_allow_html=True)

st.markdown("""
    <h2>🧪 体外诊断线性材料配制程序 
    <span style="font-size: 0.65em; font-weight: normal; color: #666;">— 适用称重稀释法</span>
    </h4>
    """, unsafe_allow_html=True)
st.caption(f"版本: {VERSION}")

# XLSX 导入
import_data = {}
with st.sidebar.expander("📂 导入 XLSX 存档", expanded=False):
    uploaded_file = st.file_uploader("选择存档文件", type="xlsx", label_visibility="collapsed")
    if uploaded_file:
        try:
            df_settings = pd.read_excel(uploaded_file, sheet_name="配置参数")
            import_data = dict(zip(df_settings["参数"], df_settings["数值"]))
            df_grad_import = pd.read_excel(uploaded_file, sheet_name="梯度方案")
            st.success("导入成功")
        except: st.error("导入失败")

with st.sidebar:
    st.subheader("⚙️ 基础设置")
    cur_date = datetime.now().strftime("%Y%m%d")
    default_name = import_data.get("实验内容", f"线性稀释实验-{cur_date}")
    exp_name = st.text_input("实验内容名称", value=str(default_name))
    
    c_u1, c_u2 = st.columns(2)
    unit_conc = c_u1.text_input("浓度单位", value=import_data.get("浓度单位", "mg/L"))
    unit_mass = c_u2.text_input("质量单位", value=import_data.get("质量单位", "mg"))
    
    input_temp = st.number_input("环境温度 (°C)", value=float(import_data.get("环境温度", 22.0)), step=0.5)
    rho_w, rho_s = get_densities(input_temp)
    st.write(f"💧 **纯水密度**: `{rho_w}` g/cm3  |  🏥 **生理盐水**: `{rho_s}` g/cm3")
    
    st.markdown("---")
    c_p1, c_p2 = st.columns(2)
    c_h_orig = c_p1.number_input("高浓度材料浓度", value=float(import_data.get("原液浓度A", 100.0)), step=1.0)
    rho_h_orig = c_p2.number_input("高浓度材料密度 (g/cm3)", value=float(import_data.get("原液密度A", 1.0500)), format="%.4f", step=0.001)
    
    c_p3, c_p4 = st.columns(2)
    c_l_orig = c_p3.number_input("低浓度材料浓度", value=float(import_data.get("原液浓度B", 0.0)), step=1.0)
    rho_l_orig = c_p4.number_input("低浓度材料密度 (g/cm3)", value=float(import_data.get("原液密度B", rho_w)), format="%.4f", step=0.001)

    c_p5, c_p6 = st.columns(2)
    num_points = c_p5.number_input("样本数量", min_value=3, max_value=20, value=int(import_data.get("样本点数", 8)), step=1)
    target_tm_each = c_p6.number_input(f"单点配置量 ({unit_mass})", value=float(import_data.get("单点计划总量", 350.0)), step=10.0)

# --- 4. 预计算与中间配置 ---
target_c_mid_guess = round((c_h_orig + c_l_orig)/2, 2)
mid_idx_guess = num_points // 2
pts_low_temp = [c_l_orig + i * ((target_c_mid_guess - c_l_orig) / mid_idx_guess) for i in range(mid_idx_guess)]
pts_high_temp = [target_c_mid_guess + i * ((c_h_orig - target_c_mid_guess) / (num_points - mid_idx_guess - 1)) for i in range(num_points - mid_idx_guess)]
all_targets_temp = pts_low_temp + pts_high_temp
total_mid_usage = 0.0
for t_c in all_targets_temp:
    if t_c > target_c_mid_guess + 0.0001: _, m_mid = calc_theoretical_masses(t_c, target_tm_each, c_h_orig, rho_h_orig, target_c_mid_guess, 1.0)
    else: m_mid, _ = calc_theoretical_masses(t_c, target_tm_each, target_c_mid_guess, 1.0, c_l_orig, rho_l_orig)
    total_mid_usage += m_mid
suggested_m = round(total_mid_usage * 1.1, 1)

st.markdown("#### 1️⃣ 中间浓度配置")
with st.container(border=True):
    col_l, col_r = st.columns(2)
    with col_l:
        # 动态 Key：当原液参数变化时，强制重置输入框
        target_c_mid = st.number_input(f"中间目标浓度 ({unit_conc})", value=float(import_data.get("中间目标浓度", target_c_mid_guess)), step=0.1, key=f"tcm_{c_h_orig}_{c_l_orig}")
        prep_m_mid = st.number_input("中间配置总量 (mg)", value=float(import_data.get("中间计划总量", max(suggested_m, 100.0))), step=10.0, key=f"pmm_{suggested_m}")
        m_h_theo, m_l_theo = calc_theoretical_masses(target_c_mid, prep_m_mid, c_h_orig, rho_h_orig, c_l_orig, rho_l_orig)
        st.info(f"💡 建议：高浓度材料 {m_h_theo:.1f} + 低浓度材料 {m_l_theo:.1f} (理论用量和: {total_mid_usage:.1f})")
    with col_r:
        # 实际加入质量使用动态 Key，随理论建议值变化自动同步
        m_h_mid_act = st.number_input("加入高浓度实测", value=float(import_data.get("中间实测A", round(m_h_theo, 1))), min_value=0.0, step=0.1, format="%.1f", key=f"mha_{m_h_theo}")
        m_l_mid_act = st.number_input("加入低浓度实测", value=float(import_data.get("中间实测B", round(m_l_theo, 1))), min_value=0.0, step=0.1, format="%.1f", key=f"mla_{m_l_theo}")
        actual_c_mid = calc_actual_volume_conc(m_h_mid_act, m_l_mid_act, c_h_orig, rho_h_orig, c_l_orig, rho_l_orig)
        denom = (m_h_mid_act/rho_h_orig) + (m_l_mid_act/rho_l_orig)
        actual_rho_mid = (m_h_mid_act + m_l_mid_act) / denom if denom > 0 else 1.0
        st.warning(f"🧪 **中间实际参数**：浓度 **{actual_c_mid:.2f}** | 密度 **{actual_rho_mid:.4f}**")

# --- 5. 梯度方案 ---
st.markdown("#### 2️⃣ 分段梯度稀释方案")
mid_idx = num_points // 2
pts_low = [c_l_orig + i * ((actual_c_mid - c_l_orig) / mid_idx) for i in range(mid_idx)]
pts_high = [actual_c_mid + i * ((c_h_orig - actual_c_mid) / (num_points - mid_idx - 1)) for i in range(num_points - mid_idx)]
all_targets = pts_low + pts_high

h_cols = st.columns([0.5, 1.2, 1.2, 1.2, 1.2, 1.2, 1.2])
headers = ["序号", "目标浓度", "材料A", "材料B", "加入A质量", "加入B质量", "实际浓度"]
for col, lab in zip(h_cols, headers): col.write(f"**{lab}**")

results_data = []
total_h = m_h_mid_act
total_l = m_l_mid_act

for i, t_c in enumerate(all_targets):
    idx = i + 1
    if t_c > actual_c_mid + 0.0001: m_a, m_b, ca, ra, cb, rb = "高浓度", "中间浓度", c_h_orig, rho_h_orig, actual_c_mid, actual_rho_mid
    else: m_a, m_b, ca, ra, cb, rb = "中间浓度", "低浓度", actual_c_mid, actual_rho_mid, c_l_orig, rho_l_orig
    
    imp_tc, imp_ma, imp_mb = t_c, None, None
    if 'df_grad_import' in locals() and i < len(df_grad_import):
        imp_tc, imp_ma, imp_mb = df_grad_import.iloc[i]["目标浓度"], df_grad_import.iloc[i]["加入A质量"], df_grad_import.iloc[i]["加入B质量"]
    
    r_cols = st.columns([0.5, 1.2, 1.2, 1.2, 1.2, 1.2, 1.2])
    r_cols[0].write(f"{idx}")
    # 动态 Key 确保梯度目标随中间浓度变化刷新
    row_tc = r_cols[1].number_input(f"tc_{i}", value=float(imp_tc), format="%.2f", step=0.1, key=f"rtc_{i}_{actual_c_mid}", label_visibility="collapsed")
    r_cols[2].write(m_a); r_cols[3].write(m_b)
    
    m_a_t, m_b_t = calc_theoretical_masses(row_tc, target_tm_each, ca, ra, cb, rb)
    # 梯度实测框使用动态 Key，确保理论配比变化时强制更新输入框
    row_ma = r_cols[4].number_input(f"ma_{i}", value=float(imp_ma if imp_ma is not None else round(m_a_t, 1)), min_value=0.0, step=0.1, format="%.1f", key=f"rma_{i}_{actual_c_mid}_{row_tc}", label_visibility="collapsed")
    row_mb = r_cols[5].number_input(f"mb_{i}", value=float(imp_mb if imp_mb is not None else round(m_b_t, 1)), min_value=0.0, step=0.1, format="%.1f", key=f"rmb_{i}_{actual_c_mid}_{row_tc}", label_visibility="collapsed")
    
    act_c = calc_actual_volume_conc(row_ma, row_mb, ca, ra, cb, rb)
    r_cols[6].write(f"**{act_c:.2f}**")
    results_data.append({"序号": idx, "目标浓度": row_tc, "材料A": m_a, "材料B": m_b, "加入A质量": row_ma, "加入B质量": row_mb, "最终实际浓度": act_c})
    if m_a == "高浓度": total_h += row_ma
    if m_b == "低浓度": total_l += row_mb

# --- 6. 导出 ---
st.divider()
ex_l, ex_r = st.columns(2)
with ex_l:
    if st.button("💾 导出 XLSX 存档", use_container_width=True):
        s_dict = {
            "实验内容": exp_name, "浓度单位": unit_conc, "质量单位": unit_mass, "环境温度": input_temp,
            "原液浓度A": c_h_orig, "原液密度A": rho_h_orig, "原液浓度B": c_l_orig, "原液密度B": rho_l_orig,
            "样本点数": num_points, "单点计划总量": target_tm_each, "中间目标浓度": target_c_mid, "中间计划总量": prep_m_mid,
            "中间实测A": m_h_mid_act, "中间实测B": m_l_mid_act
        }
        df_s = pd.DataFrame(list(s_dict.items()), columns=["参数", "数值"])
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_s.to_excel(writer, sheet_name="配置参数", index=False)
            pd.DataFrame(results_data).to_excel(writer, sheet_name="梯度方案", index=False)
        st.download_button("📥 下载 XLSX", data=output.getvalue(), file_name=f"{exp_name}.xlsx", use_container_width=True)

with ex_r:
    if st.button("📑 生成 PDF 报告", use_container_width=True):
        df_mid_pdf = pd.DataFrame([
            {"组分": "高浓度材料", "理论质量": m_h_theo, "加入质量": m_h_mid_act, "目标浓度": "-", "实际配置浓度": "-"},
            {"组分": "低浓度材料", "理论质量": m_l_theo, "加入质量": m_l_mid_act, "目标浓度": "-", "实际配置浓度": "-"},
            {"组分": "合计(中间浓度)", "理论质量": m_h_theo+m_l_theo, "加入质量": m_h_mid_act+m_l_mid_act, "目标浓度": target_c_mid, "实际配置浓度": actual_c_mid}
        ])
        df_mid_pdf.columns = ["组分", f"理论质量({unit_mass})", f"加入质量({unit_mass})", f"目标浓度({unit_conc})", f"实际配置浓度({unit_conc})"]
        meta = {
            "实验内容": exp_name, "环境温度": f"{input_temp} degC", "水密度": f"{rho_w} g/cm3", "生理盐水密度": f"{rho_s} g/cm3",
            "高浓度材料": f"{c_h_orig} (D:{rho_h_orig})", "低浓度材料": f"{c_l_orig} (D:{rho_l_orig})", "中间材料": f"{actual_c_mid:.2f} (D:{actual_rho_mid:.4f})",
            "高浓度材料合计": f"{total_h:.1f}", "低浓度材料合计": f"{total_l:.1f}", "生成时间": datetime.now().strftime("%Y-%m-%d %H:%M"), "程序版本": VERSION
        }
        pdf_out = create_pdf(pd.DataFrame(results_data), df_mid_pdf, "线性评价样本制备记录", meta)
        st.download_button("📥 下载 PDF", data=bytes(pdf_out), file_name=f"Report_{exp_name}.pdf", use_container_width=True)