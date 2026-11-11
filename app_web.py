"""
Northern Solar - Web Sistem Pengekstrak DWG & Penjana Laporan PVsyst PDF
=======================================================================
Aliran Kerja 2 Langkah (Diasingkan):
- Langkah 1: Muat naik DWG PDF ➡️ Jana & Muat Turun Fail Excel Template ([NamaPelanggan]_Template.xlsx).
  Drafter buka fail Excel untuk menyemak & mengesahkan data.
- Langkah 2: Muat naik Fail Excel yang telah disemak ➡️ Jana & Muat Turun Laporan PVsyst PDF Bersih.
  Auto-pilih template (1Battery vs NoBattery), 100% bebas highlight kuning, graf & carta dikemaskini.
"""

import os
import io
import re
import math
import openpyxl
import pandas as pd
import streamlit as st
import pymupdf
import dwg_excel_extractor as extractor

# ==============================================================================
# KONFIGURASI HALAMAN STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="Northern Solar - PVSyst & Excel Automation System",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Theme Adaptive)
st.markdown("""
<style>
    .main-title {
        font-size: 26px;
        font-weight: 800;
        color: #0284C7;
        margin-bottom: 2px;
    }
    .sub-title {
        font-size: 14px;
        color: #64748B;
        margin-bottom: 20px;
    }
    .step-card {
        border-radius: 10px;
        padding: 18px;
        border: 1px solid #BAE6FD;
        background-color: #F0F9FF;
        margin-bottom: 20px;
    }
    .download-box {
        border-radius: 10px;
        padding: 20px;
        border: 2px solid #10B981;
        background-color: #ECFDF5;
        margin-top: 15px;
        margin-bottom: 20px;
    }
    .info-badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
        margin-right: 6px;
    }
    .stDownloadButton button {
        font-size: 16px !important;
        font-weight: 700 !important;
        padding: 12px 28px !important;
        border-radius: 8px !important;
        width: 100% !important;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# BANTUAN VISUAL PVSYST FIELD PARAMETERS (SVG)
# ==============================================================================

def render_pvsyst_orientation_svg(tilt, azimuth):
    """
    Menjana visual PVsyst persis Gambar 1:
    - Kiri: Tilt Diagram (Side View)
    - Kanan: Azimuth Diagram (Top View - paksi West, East, South, pivot dot, modul PV biru, red dot)
    """
    tilt = float(tilt)
    azimuth = float(azimuth)

    t_rad = math.radians(min(89.0, max(0.0, tilt)))
    L_tilt = 62.0
    x0_t, y0_t = 85.0, 78.0
    x1_t = x0_t + L_tilt * math.cos(t_rad)
    y1_t = y0_t - L_tilt * math.sin(t_rad)

    cx, cy = 295.0, 70.0
    w, h = 60.0, 26.0
    south_len = 30.0

    svg_parts = [
        f'<svg viewBox="0 0 420 135" xmlns="http://www.w3.org/2000/svg" style="width:100%; max-width:440px; height:auto; background:#ffffff; border:1px solid #94a3b8; border-radius:6px; box-shadow:0 1px 2px rgba(0,0,0,0.05); font-family: Segoe UI, Arial, sans-serif;">',
        f'<text x="95" y="22" font-size="13" font-weight="bold" text-anchor="middle" fill="#000000">Tilt {tilt:.1f}°</text>',
        f'<line x1="20" y1="{y0_t}" x2="175" y2="{y0_t}" stroke="#000000" stroke-width="1.8" />',
        f'<circle cx="{x0_t}" cy="{y0_t}" r="3" fill="#000000" />',
        f'<line x1="{x0_t}" y1="{y0_t}" x2="{x1_t:.1f}" y2="{y1_t:.1f}" stroke="#0000ff" stroke-width="2.6" stroke-linecap="round" />',
        f'<circle cx="{x1_t:.1f}" cy="{y1_t:.1f}" r="3.5" fill="#ff0000" />',
        f'<line x1="195" y1="10" x2="195" y2="125" stroke="#cbd5e1" stroke-width="1" stroke-dasharray="3,3" />',
        f'<text x="{cx}" y="22" font-size="13" font-weight="bold" text-anchor="middle" fill="#000000">Azimuth {azimuth:.0f}°</text>',
        f'<g transform="rotate({azimuth:.1f}, {cx}, {cy})">',
        f'<rect x="{cx - w/2:.1f}" y="{cy - h:.1f}" width="{w}" height="{h}" fill="#c6d9f1" stroke="#0000ff" stroke-width="2" />',
        f'<circle cx="{cx + w/2:.1f}" cy="{cy - h:.1f}" r="3.5" fill="#ff0000" />',
        f'</g>',
        f'<line x1="210" y1="{cy}" x2="380" y2="{cy}" stroke="#000000" stroke-width="1.8" />',
        f'<text x="210" y="{cy - 5}" font-size="11" font-weight="bold" text-anchor="start" fill="#000000">West</text>',
        f'<text x="380" y="{cy - 5}" font-size="11" font-weight="bold" text-anchor="end" fill="#000000">East</text>',
        f'<line x1="{cx}" y1="{cy}" x2="{cx}" y2="{cy + south_len}" stroke="#000000" stroke-width="1.8" />',
        f'<circle cx="{cx}" cy="{cy}" r="3" fill="#000000" />',
        f'<text x="{cx}" y="{cy + south_len + 16}" font-size="11" font-weight="bold" text-anchor="middle" fill="#000000">South</text>',
        f'</svg>'
    ]
    return "".join(svg_parts)

def render_pvsyst_orientation_image(tilt, azimuth):
    """Menjana imej visual PVsyst PNG yang tajam dan 100% bebas ralat perenggan/DOM."""
    svg_str = render_pvsyst_orientation_svg(tilt, azimuth)
    try:
        doc = pymupdf.open(stream=svg_str.encode("utf-8"), filetype="svg")
        pix = doc[0].get_pixmap(dpi=180)
        return pix.tobytes("png")
    except Exception:
        return svg_str

def parse_tilt_azimuth(str_val):
    """Mengekstrak nilai (tilt, azimuth) daripada format '20/180°' atau '20/0'."""
    try:
        clean = str_val.replace("°", "").strip()
        parts = clean.split("/")
        if len(parts) == 2:
            return float(parts[0]), float(parts[1])
    except Exception:
        pass
    return 20.0, 0.0

def main():
    current_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() and __file__ else os.getcwd()
    default_excel_template = os.path.join(current_dir, "Template.xlsx")
    
    templates_status = {
        "🔋 Battery (2 Ori)": os.path.join(current_dir, "BATTERY (2 Orientation).pdf"),
        "🔋 Battery (3 Ori)": os.path.join(current_dir, "BATTERY (3 Orientation).pdf"),
        "🔋 Battery (4 Ori)": os.path.join(current_dir, "BATTERY (4 Orientation).pdf"),
        "⚡ No-Battery (2 Ori)": os.path.join(current_dir, "NO BATTERY (2 Orientation).pdf"),
        "⚡ No-Battery (3 Ori)": os.path.join(current_dir, "NO BATTERY (3 Orientation).pdf"),
        "⚡ No-Battery (4 Ori)": os.path.join(current_dir, "NO BATTERY (4 Orientation).pdf"),
    }

    # SIDEBAR
    with st.sidebar:
        st.image("https://northernsolar.com.my/wp-content/uploads/2023/06/northern-solar-logo.png", width=220)
        st.markdown("### ☀️ Modul Automasi Drafter")
        st.info("""
        **Aliran Kerja Profesional:**
        
        **Langkah 1 (Tab 1):**
        1. Muat naik fail lukisan DWG PDF.
        2. Sistem isi semua sel kuning dalam fail Excel.
        3. Muat turun & semak fail Excel.
        
        **Langkah 2 (Tab 2):**
        1. Muat naik fail Excel yang telah disahkan.
        2. Sistem auto-pilih daripada **6 Template PVsyst PDF** (2/3/4 Orientasi x Battery/No-Battery).
        3. Laporan PVsyst PDF dijana 100% bersih tanpa highlight kuning dan sifar pertindihan teks.
        """)
        
        st.markdown("---")
        st.markdown("#### ⚙️ Status 6 Template Sistem")
        st.write(f"📊 **Excel Template:** `Template.xlsx` {'✅' if os.path.exists(default_excel_template) else '❌'}")
        for t_label, t_path in templates_status.items():
            t_fname = os.path.basename(t_path)
            st.write(f"{t_label}: `{t_fname}` {'✅' if os.path.exists(t_path) else '❌'}")
        
        st.markdown("---")
        st.caption("🔒 Formula, PVsyst Graph Data & GlobHor_DiffHor 100% dilindungi.")

    # HEADER
    st.markdown('<div class="main-title">☀️ NORTHERN SOLAR - SISTEM PENJANAAN PVSYST & EXCEL</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Automasi pemprosesan lukisan teknikal DWG AutoCAD kepada fail Excel Template dan penjanaan semula Laporan Simulasi PVsyst PDF (Matriks 6 Template).</div>', unsafe_allow_html=True)

    if not os.path.exists(default_excel_template):
        st.error("Ralat: Fail Template.xlsx tidak dijumpai dalam folder projek.")
        return

    # TABS UTAMA (LANGKAH 1 VS LANGKAH 2)
    tab1, tab2 = st.tabs([
        "📄 LANGKAH 1: Muat Naik DWG PDF ➡️ Jana Excel Template",
        "📊 LANGKAH 2: Muat Naik Excel Disemak ➡️ Jana Laporan PVsyst PDF"
    ])

    # ==========================================================================
    # TAB 1: DWG PDF ➡️ EXCEL TEMPLATE
    # ==========================================================================
    with tab1:
        st.markdown("""
        <div class="step-card">
            <h4>📄 Langkah 1: Pengekstrakan Lukisan DWG PDF ke Template Excel</h4>
            Muat naik fail lukisan <b>DWG PDF AutoCAD</b> untuk mengisi semua sel kuning dalam <code>Template.xlsx</code> secara automatik (Nama Pelanggan, Alamat Ringkas, Orientasi Roof Fall 1-4, Inverter, Modul PV, dan Bateri).
        </div>
        """, unsafe_allow_html=True)

        uploaded_dwg = st.file_uploader(
            "Pilih atau Tarik (Drag & Drop) Fail Lukisan DWG PDF (.pdf)",
            type=["pdf"],
            key="tab1_dwg_uploader"
        )

        if uploaded_dwg:
            with st.spinner("🔍 Sedang mengekstrak maklumat lukisan DWG PDF..."):
                try:
                    dwg_bytes = uploaded_dwg.getvalue()
                    dwg_data = extractor.extract_dwg_info(dwg_bytes)
                    project_data = extractor.combine_project_data(dwg_data=dwg_data)
                except Exception as e:
                    st.error(f"Ralat mengekstrak fail DWG: {str(e)}")
                    return

            client_name = project_data.get("client_name") or "Pelanggan"
            safe_name = re.sub(r'[\\/*?:"<>|]', '', client_name).strip() or "Projek"
            kwp_val = project_data.get("system_size_kwp", 0.0)

            st.success(f"✅ Berjaya mengekstrak maklumat projek: **{client_name}** ({kwp_val} kWp)")

            # Ringkasan Metrik
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Nama Pelanggan", client_name)
            m2.metric("Kapasiti Sistem", f"{kwp_val} kWp", f"{project_data.get('system_size_kwac', 0.0)} kWac")
            m3.metric("Jumlah Panel", f"{project_data['pv_module'].get('total_panels', 0)} PCS", f"{project_data['pv_module'].get('unit_power_w', 630)}W")
            m4.metric("Model Inverter", project_data['inverter'].get('model') or "N/A", f"{project_data['inverter'].get('unit_power_kw', 0)} kW")
            m5.metric("Bateri", f"{project_data['battery'].get('units', 0)} Unit", f"{project_data['battery'].get('nominal_energy_kwh', 0)} kWh")

            # Reset session state jika fail DWG baharu dimuat naik
            if st.session_state.get("last_uploaded_dwg") != uploaded_dwg.name:
                st.session_state["last_uploaded_dwg"] = uploaded_dwg.name
                for k in list(st.session_state.keys()):
                    if k.startswith("plane_tilt_") or k.startswith("azimuth_") or k == "num_roof_orientations" or k in ["custom_annual_load", "custom_weather_data"] or k.startswith("custom_pv_") or k.startswith("pdf_generated_"):
                        del st.session_state[k]

            # ------------------------------------------------------------------
            # PARAMETER SISTEM: MODUL PV, ANNUAL LOAD & WEATHER DATA
            # ------------------------------------------------------------------
            st.markdown("---")
            st.markdown("### ⚙️ Parameter Sistem (Modul PV, Annual Load & Weather Data)")
            st.caption("Nilai dicadangkan secara automatik daripada pengekstrakan DWG dan negeri. Anda boleh memilih modul PV atau mengubah nilai ini secara terus di bawah:")

            col_pv, col_p1, col_p2 = st.columns([1.3, 1, 1])
            with col_pv:
                pv_choices = [
                    "JA Solar JAM72D42-650/LB (650 Wp)",
                    "Astronergy ASTRO-N7 CHSM66RN-630W (630 Wp)",
                    "JA Solar JAM72D42-640/LB (640 Wp)",
                    "JA Solar JAM72D42-635/LB (635 Wp)",
                    "JA Solar JAM66D50-650/GB (650 Wp)"
                ]
                cur_mod = str(project_data["pv_module"].get("model", "")).upper()
                cur_w = project_data["pv_module"].get("unit_power_w", 630)

                default_idx = 0
                if "650" in cur_mod or cur_w == 650:
                    if "JAM66D50" in cur_mod:
                        default_idx = 4
                    else:
                        default_idx = 0
                elif "640" in cur_mod or cur_w == 640:
                    default_idx = 2
                elif "635" in cur_mod or cur_w == 635:
                    default_idx = 3
                elif "CHSM" in cur_mod or "ASTRO" in cur_mod or cur_w == 630:
                    default_idx = 1

                sel_pv = st.selectbox(
                    "☀️ Pilihan Modul PV [Sheet PAGE 3]:",
                    options=pv_choices,
                    index=default_idx,
                    key=f"custom_pv_{uploaded_dwg.name}",
                    help="Pilih modul PV yang digunakan. Pengeluar, model, fail .PAN dan kuasa Wp akan dikemas kini secara automatik."
                )

                if "JAM72D42-650" in sel_pv:
                    project_data["pv_module"]["manufacturer"] = "JA Solar"
                    project_data["pv_module"]["model"] = "JAM72D42-650/LB"
                    project_data["pv_module"]["pan_file"] = "JAM72D42-650/LB.PAN"
                    project_data["pv_module"]["unit_power_w"] = 650
                elif "JAM66D50-650" in sel_pv:
                    project_data["pv_module"]["manufacturer"] = "JA Solar"
                    project_data["pv_module"]["model"] = "JAM66D50-650/GB"
                    project_data["pv_module"]["pan_file"] = "JAM66D50-650/GB.PAN"
                    project_data["pv_module"]["unit_power_w"] = 650
                elif "JAM72D42-640" in sel_pv:
                    project_data["pv_module"]["manufacturer"] = "JA Solar"
                    project_data["pv_module"]["model"] = "JAM72D42-640/LB"
                    project_data["pv_module"]["pan_file"] = "JAM72D42-640/LB.PAN"
                    project_data["pv_module"]["unit_power_w"] = 640
                elif "JAM72D42-635" in sel_pv:
                    project_data["pv_module"]["manufacturer"] = "JA Solar"
                    project_data["pv_module"]["model"] = "JAM72D42-635/LB"
                    project_data["pv_module"]["pan_file"] = "JAM72D42-635/LB.PAN"
                    project_data["pv_module"]["unit_power_w"] = 635
                else:
                    project_data["pv_module"]["manufacturer"] = "Astronergy"
                    project_data["pv_module"]["model"] = "ASTRO-N7 CHSM66RN-630W"
                    project_data["pv_module"]["pan_file"] = "CHSM66RN(DG)F-BH-630.PAN"
                    project_data["pv_module"]["unit_power_w"] = 630

                if project_data["pv_module"].get("total_panels", 0) > 0:
                    project_data["system_size_kwp"] = round((project_data["pv_module"]["total_panels"] * project_data["pv_module"]["unit_power_w"]) / 1000, 2)

            col_p1, col_p2 = st.columns(2)
            with col_p1:
                init_load = int(project_data.get("used_energy_kwh") or 11400)
                user_load = st.number_input(
                    "⚡ Annual Load (kWh) — [Sheet PAGE 1 sel B6]:",
                    min_value=0,
                    max_value=10000000,
                    value=init_load,
                    step=100,
                    key="custom_annual_load",
                    help="Nilai beban tahunan (Annual Load) yang akan dimasukkan ke sel kuning B6 pada sheet PAGE 1."
                )
                project_data["used_energy_kwh"] = int(user_load)

            with col_p2:
                detected_state = project_data.get("state") or "Tidak Dikesan"
                init_weather = project_data.get("weather_data_variance")
                if init_weather is None:
                    init_weather = extractor.get_weather_assumption_for_state(detected_state)
                user_weather = st.number_input(
                    f"⛅ Weather Data Variance (%) — [Sheet PAGE 7 sel B2] (Negeri: {detected_state}):",
                    min_value=0.0,
                    max_value=20.0,
                    value=float(init_weather),
                    step=0.05,
                    format="%.2f",
                    key="custom_weather_data",
                    help=f"Nilai varians data cuaca untuk sel kuning B2 pada sheet PAGE 7. Lalai bagi negeri {detected_state} ialah {init_weather:.2f}%."
                )
                project_data["weather_data_variance"] = float(user_weather)

            # ------------------------------------------------------------------
            # WIDGET TETAPAN ORIENTASI PVSYST (TILT & AZIMUTH)
            # ------------------------------------------------------------------
            st.markdown("---")
            st.markdown("### 🧭 Tetapan Orientasi Bumbung (PVsyst Field Parameters)")
            st.caption("Laraskan sudut kecondongan (**Plane tilt**) dan arah (**Azimuth**) mengikut lukisan teknikal PVsyst. Nilai ini akan dimasukkan terus ke dalam sel kuning sheet `PAGE2` (`B9`, `D9`, `F9`, `I9`) fail Excel.")

            oris_detected = project_data.get("orientations", [])
            active_detected_count = sum(1 for o in oris_detected if o.get("str_val") not in ["0/0", "0", ""])
            if active_detected_count < 1:
                active_detected_count = 1

            num_active = st.radio(
                "Bilangan Cerun Bumbung (Orientasi Aktif):",
                options=[1, 2, 3, 4],
                index=min(max(active_detected_count, 1), 4) - 1,
                horizontal=True,
                key="num_roof_orientations",
                help="Pilih bilangan orientasi cerun bumbung aktif bagi projek ini (1 hingga 4)."
            )

            custom_orientations = []

            for i in range(1, 5):
                if i <= num_active:
                    # Ambil nilai default daripada pengekstrakan DWG jika ada
                    def_tilt, def_az = 20.0, 0.0
                    if i <= len(oris_detected) and oris_detected[i-1].get("str_val") not in ["0/0", "0", ""]:
                        def_tilt, def_az = parse_tilt_azimuth(oris_detected[i-1]["str_val"])
                    elif i == 1:
                        def_tilt, def_az = 20.0, 180.0

                    with st.container():
                        st.markdown(f"##### 📐 Orientasi #{i} (Roof Fall #{i})")
                        c_left, c_right = st.columns([1, 1.3])
                        with c_left:
                            st.markdown("""
                            <div style="font-size:12px; font-weight:600; color:#475569; margin-bottom:4px;">
                                Field parameters
                            </div>
                            """, unsafe_allow_html=True)

                            t_val = st.number_input(
                                f"Plane tilt (°) [#{i}]",
                                min_value=0.0,
                                max_value=89.0,
                                value=float(def_tilt),
                                step=1.0,
                                format="%.1f",
                                key=f"plane_tilt_{i}"
                            )
                            az_val = st.number_input(
                                f"Azimuth (°) [#{i}]",
                                min_value=-180.0,
                                max_value=180.0,
                                value=float(def_az),
                                step=1.0,
                                format="%.1f",
                                key=f"azimuth_{i}",
                                help="Sudut Azimuth PVsyst: 0° = South, -90° = East, 90° = West, 180°/-180° = North"
                            )
                            st.text_input(
                                f"Base tilt angle (°) [#{i}]",
                                value="0.0°",
                                disabled=True,
                                key=f"base_tilt_{i}"
                            )
                        with c_right:
                            img_data = render_pvsyst_orientation_image(t_val, az_val)
                            st.image(img_data, use_container_width=True)
                        st.markdown("<div style='margin-bottom:12px;'></div>", unsafe_allow_html=True)

                    t_str = f"{int(t_val) if t_val.is_integer() else t_val}"
                    az_str = f"{int(az_val) if az_val.is_integer() else az_val}"
                    custom_orientations.append({"id": i, "str_val": f"{t_str}/{az_str}°"})
                else:
                    custom_orientations.append({"id": i, "str_val": "0/0"})

            # Kemas kini orientations dalam project_data sebelum pengisian template Excel
            project_data["orientations"] = custom_orientations

            # ------------------------------------------------------------------
            # SEMAKAN BUTIRAN DATA EXCEL (DOUBLE CHECK)
            # ------------------------------------------------------------------
            # ------------------------------------------------------------------
            # SEMAKAN & EDIT BUTIRAN DATA EXCEL (EDITABLE TABLE)
            # ------------------------------------------------------------------
            st.markdown("---")
            with st.expander("🔍 Semak & Edit Butiran Data Yang Diisi Ke Dalam Template Excel", expanded=False):
                st.caption("💡 **Boleh Diedit**: Anda boleh terus mengubah nilai pada lajur **'Nilai Dimasukkan'** di bawah. Sebarang perubahan akan mengemas kini fail Excel dan Laporan PVsyst PDF secara automatik!")
                mapped_dict = extractor.get_mapped_values_for_template(project_data)
                rows = []
                for (s_name, coord), val in mapped_dict.items():
                    param_label = extractor.get_template_cell_label(s_name, coord)
                    rows.append({
                        "_sheet": s_name,
                        "_coord": coord,
                        "Sheet": s_name,
                        "Butiran (Ikut Template Excel)": param_label,
                        "Nilai Dimasukkan": str(val) if val is not None else ""
                    })
                df_map = pd.DataFrame(rows)
                edited_df = st.data_editor(
                    df_map,
                    column_config={
                        "_sheet": None,
                        "_coord": None,
                        "Sheet": st.column_config.TextColumn("Sheet", disabled=True),
                        "Butiran (Ikut Template Excel)": st.column_config.TextColumn("Butiran (Ikut Template Excel)", disabled=True),
                        "Nilai Dimasukkan": st.column_config.TextColumn("Nilai Dimasukkan", disabled=False)
                    },
                    use_container_width=True,
                    hide_index=True,
                    key=f"editor_excel_map_{uploaded_dwg.name}"
                )

            # Proses sebarang overrides/perubahan daripada edited_df
            custom_overrides = {}
            if edited_df is not None:
                for _, r in edited_df.iterrows():
                    s_name = r["_sheet"]
                    coord = r["_coord"]
                    orig_val = mapped_dict.get((s_name, coord))
                    new_val_str = str(r["Nilai Dimasukkan"]).strip()
                    if new_val_str != str(orig_val if orig_val is not None else ""):
                        try:
                            if "." in new_val_str:
                                val_to_use = float(new_val_str)
                            else:
                                val_to_use = int(new_val_str)
                        except ValueError:
                            val_to_use = new_val_str
                        custom_overrides[(s_name, coord)] = val_to_use

                        # Kemas kini project_data jika parameter teras diubah
                        if s_name == "PAGE 1" and coord == "B1":
                            project_data["client_name"] = str(val_to_use)
                        elif s_name == "PAGE 1" and coord == "B3":
                            project_data["short_address"] = str(val_to_use)
                        elif s_name == "PAGE2" and coord == "F13":
                            try:
                                n_u = int(val_to_use)
                                project_data["battery"]["units"] = n_u
                                project_data["battery"]["capacity_ah"] = n_u * 314.0
                            except Exception:
                                pass
                        elif s_name == "PAGE2" and coord == "F14":
                            try:
                                project_data["battery"]["voltage_v"] = float(val_to_use)
                            except Exception:
                                pass
                        elif s_name == "PAGE2" and coord == "F15":
                            try:
                                project_data["battery"]["capacity_ah"] = float(val_to_use)
                            except Exception:
                                pass
                        elif s_name == "PAGE2" and coord == "D14":
                            try:
                                project_data["inverter"]["unit_power_kw"] = float(val_to_use)
                                project_data["system_size_kwac"] = float(val_to_use)
                            except Exception:
                                pass
                        elif s_name == "PAGE2" and coord == "B13":
                            try:
                                project_data["pv_module"]["total_panels"] = int(val_to_use)
                                if project_data["pv_module"].get("unit_power_w", 0) > 0:
                                    project_data["system_size_kwp"] = round((int(val_to_use) * project_data["pv_module"]["unit_power_w"]) / 1000, 2)
                            except Exception:
                                pass
                        elif s_name == "PAGE 3" and coord == "B14":
                            project_data["pv_module"]["manufacturer"] = str(val_to_use)
                        elif s_name == "PAGE 3" and coord == "B15":
                            project_data["pv_module"]["model"] = str(val_to_use)
                        elif s_name == "PAGE 3" and coord == "B16":
                            project_data["pv_module"]["pan_file"] = str(val_to_use)
                        elif s_name == "PAGE 3" and coord == "B17":
                            try:
                                project_data["pv_module"]["unit_power_w"] = int(val_to_use)
                                if project_data["pv_module"].get("total_panels", 0) > 0:
                                    project_data["system_size_kwp"] = round((project_data["pv_module"]["total_panels"] * int(val_to_use)) / 1000, 2)
                            except Exception:
                                pass

            # Isi Sel Kuning Excel dalam memori
            with open(default_excel_template, "rb") as f:
                excel_tmpl_bytes = f.read()

            tmpl_wb = openpyxl.load_workbook(io.BytesIO(excel_tmpl_bytes))
            extractor.fill_yellow_cells(tmpl_wb, project_data, custom_overrides=custom_overrides)
            out_stream = io.BytesIO()
            tmpl_wb.save(out_stream)
            filled_excel_bytes = out_stream.getvalue()

            excel_filename = f"{safe_name}.xlsx"
            pdf_filename = f"{safe_name}.pdf"

            # ------------------------------------------------------------------
            # PENJANAAN LAPORAN PVSYST PDF (TERUS) & MUAT TURUN
            # ------------------------------------------------------------------
            st.markdown("### 📑 Penjanaan Laporan PVsyst PDF & Muat Turun")
            st.write("Selepas mengesahkan data di atas, anda boleh **terus menjana Laporan PVsyst PDF** dengan satu klik (tanpa perlu muat turun fail Excel terlebih dahulu):")

            col_act1, col_act2 = st.columns([1.5, 1])
            with col_act1:
                btn_gen_pdf = st.button("🚀 JANA LAPORAN PVSYST PDF SEKARANG", type="primary", use_container_width=True, key="btn_gen_pdf_tab1")
            with col_act2:
                st.download_button(
                    label=f"📥 Muat Turun Excel ({excel_filename})",
                    data=filled_excel_bytes,
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_download_excel_tab1",
                    use_container_width=True,
                    help="Pilihan: Muat turun salinan fail Excel untuk semakan atau simpanan arkib."
                )

            state_key = f"pdf_generated_{uploaded_dwg.name}"
            if btn_gen_pdf or st.session_state.get(state_key):
                st.session_state[state_key] = True

                with st.spinner("📑 Sedang menjana semula Laporan PVsyst PDF bersih..."):
                    try:
                        ex_data = extractor.extract_excel_info(filled_excel_bytes)
                        full_proj_data = extractor.combine_project_data(dwg_data=project_data, excel_data=ex_data)
                        target_tmpl_path = extractor.select_pvsyst_template(full_proj_data)
                        target_tmpl_name = os.path.basename(target_tmpl_path)
                        clean_pdf_bytes = extractor.generate_pvsyst_pdf(target_tmpl_path, full_proj_data)
                    except Exception as e:
                        st.error(f"Ralat semasa menjana PDF PVsyst: {str(e)}")
                        clean_pdf_bytes = None

                if clean_pdf_bytes:
                    st.markdown('<div class="download-box">', unsafe_allow_html=True)
                    st.markdown(f"### 🎉 Laporan PVsyst PDF Berjaya Dijana: `{pdf_filename}`")
                    st.write(f"Template dipilih automatik: **`{target_tmpl_name}`** | 100% Bersih & Sifar Highlight Kuning.")

                    st.download_button(
                        label=f"⬇️ MUAT TURUN LAPORAN PVSYST PDF ({pdf_filename})",
                        data=clean_pdf_bytes,
                        file_name=pdf_filename,
                        mime="application/pdf",
                        key="btn_download_pdf_tab1",
                        type="primary"
                    )
                    st.markdown('</div>', unsafe_allow_html=True)

                    # Pratonton Halaman PDF
                    st.markdown("#### 👁️ Pratonton Halaman Laporan PVsyst Yang Dijana")
                    doc_prev = pymupdf.open(stream=clean_pdf_bytes, filetype="pdf")
                    p_cols = st.columns(3)
                    with p_cols[0]:
                        st.caption("Muka Surat 1: Cover Projek")
                        st.image(doc_prev[0].get_pixmap(dpi=120).tobytes("png"), use_container_width=True)
                    with p_cols[1]:
                        st.caption("Muka Surat 2: System Summary")
                        if len(doc_prev) >= 2:
                            st.image(doc_prev[1].get_pixmap(dpi=120).tobytes("png"), use_container_width=True)
                    with p_cols[2]:
                        st.caption(f"Muka Surat {len(doc_prev)-1}: P50-P90 Evaluation")
                        st.image(doc_prev[-2 if len(doc_prev)>=10 else -1].get_pixmap(dpi=120).tobytes("png"), use_container_width=True)

        else:
            st.info("👈 Sila muat naik fail lukisan DWG PDF di atas untuk memulakan.")


    # ==========================================================================
    # TAB 2: EXCEL DISEDIAKAN ➡️ PVSYST PDF REPORT
    # ==========================================================================
    with tab2:
        st.markdown("""
        <div class="step-card">
            <h4>📊 Langkah 2: Penjanaan Laporan PVsyst PDF daripada Fail Excel Disemak</h4>
            Muat naik fail <b>Excel yang telah disemak oleh drafter</b>. Sistem akan mengesan konfigurasi sistem secara automatik mengikut matriks 6 template:
            <ul>
                <li>🔋 <b>Auto-Pilih Template:</b>
                    <ul>
                        <li><b>2 Orientasi:</b> <code>BATTERY (2 Orientation).pdf</code> (jika ada bateri) / <code>NO BATTERY (2 Orientation).pdf</code> (jika tiada bateri)</li>
                        <li><b>3 Orientasi:</b> <code>BATTERY (3 Orientation).pdf</code> (jika ada bateri) / <code>NO BATTERY (3 Orientation).pdf</code> (jika tiada bateri)</li>
                        <li><b>4 Orientasi:</b> <code>BATTERY (4 Orientation).pdf</code> (jika ada bateri) / <code>NO BATTERY (4 Orientation).pdf</code> (jika tiada bateri)</li>
                    </ul>
                </li>
                <li>✨ <b>100% Bebas Highlight Kuning:</b> Semua kotak highlight kuning dan anotasi dipadam sepenuhnya.</li>
                <li>📈 <b>Format Kemas & Sifar Pertindihan:</b> Teks dan formula diganti dengan tepat tanpa sebarang overlap.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

        uploaded_checked_excel = st.file_uploader(
            "Pilih atau Tarik (Drag & Drop) Fail Excel Yang Telah Disemak (.xlsx)",
            type=["xlsx", "xls"],
            key="tab2_excel_uploader"
        )

        if uploaded_checked_excel:
            with st.spinner("🔍 Sedang membaca maklumat daripada fail Excel..."):
                try:
                    excel_bytes = uploaded_checked_excel.getvalue()
                    checked_excel_data = extractor.extract_excel_info(excel_bytes)
                    checked_project_data = extractor.combine_project_data(excel_data=checked_excel_data)
                except Exception as e:
                    st.error(f"Ralat membaca fail Excel: {str(e)}")
                    return

            client_name_ex = checked_project_data.get("client_name") or "Pelanggan"
            safe_name_ex = re.sub(r'[\\/*?:"<>|]', '', client_name_ex).strip() or "Projek"
            kwp_ex = checked_project_data.get("system_size_kwp", 0.0)
            bat_units_ex = checked_project_data.get("battery", {}).get("units", 0)
            has_bat = bat_units_ex > 0

            # Pemilihan Template Automatik (Matriks 6 Template)
            target_tmpl_path = extractor.select_pvsyst_template(checked_project_data)
            target_tmpl_name = os.path.basename(target_tmpl_path)

            st.success(f"✅ Data projek dimuatkan: **{client_name_ex}** ({kwp_ex} kWp)")

            # Ringkasan Maklumat Disemak
            col_inf1, col_inf2, col_inf3, col_inf4 = st.columns(4)
            col_inf1.metric("Pelanggan", client_name_ex)
            col_inf1.caption(f"📍 {checked_project_data.get('short_address', '-')}")
            
            col_inf2.metric("Kapasiti PV / AC", f"{kwp_ex} kWp", f"{checked_project_data.get('system_size_kwac', 0.0)} kWac")
            col_inf2.caption(f"☀️ Orientasi: {len(checked_project_data.get('orientations', []))} falls")
            
            pr_val = checked_project_data.get('perf_ratio_pr', 0.0)
            col_inf3.metric("Performance Ratio (PR)", f"{pr_val:.2f} %", f"Ratio: {round(pr_val/100.0, 3):.3f}")
            col_inf3.caption(f"⚡ Tenaga: {checked_project_data.get('produced_energy_kwh', 0):,.0f} kWh/thn")
            
            col_inf4.metric("Konfigurasi Bateri", f"{bat_units_ex} Unit" if has_bat else "Tiada Bateri", f"Template: {target_tmpl_name}")
            col_inf4.caption(f"🔋 {f'{bat_units_ex}-Battery System' if has_bat else 'No-Battery System'}")

            # Penjanaan PDF Bersih
            with st.spinner("📑 Sedang menjana semula Laporan PVsyst PDF bersih..."):
                try:
                    clean_pdf_bytes = extractor.generate_pvsyst_pdf(target_tmpl_path, checked_project_data)
                except Exception as e:
                    st.error(f"Ralat semasa menjana PDF PVsyst: {str(e)}")
                    return

            pdf_out_name = f"{safe_name_ex}.pdf"

            # Download Box
            st.markdown('<div class="download-box">', unsafe_allow_html=True)
            st.markdown(f"### 📑 Muat Turun Laporan PVsyst PDF: `{pdf_out_name}`")
            st.write(f"Laporan PDF rasmi telah dijana berdasarkan template **{target_tmpl_name}**. Semua highlight kuning telah dipadam dan maklumat projek baharu telah dimasukkan secara kemas:")

            st.download_button(
                label=f"⬇️ MUAT TURUN LAPORAN PVSYST PDF ({pdf_out_name})",
                data=clean_pdf_bytes,
                file_name=pdf_out_name,
                mime="application/pdf",
                key="btn_download_pdf_tab2"
            )
            st.markdown('</div>', unsafe_allow_html=True)

            # Preview Pages
            st.markdown("#### 👁️ Praterap Halaman Laporan PVsyst Yang Dijana")
            doc_preview = pymupdf.open(stream=clean_pdf_bytes, filetype="pdf")
            p_cols = st.columns(3)
            
            # Cover Page
            with p_cols[0]:
                st.caption("Muka Surat 1: Cover Projek")
                img1 = doc_preview[0].get_pixmap(dpi=120).tobytes("png")
                st.image(img1, use_container_width=True)

            # Summary Page
            with p_cols[1]:
                st.caption("Muka Surat 2: System Summary")
                if len(doc_preview) >= 2:
                    img2 = doc_preview[1].get_pixmap(dpi=120).tobytes("png")
                    st.image(img2, use_container_width=True)

            # P50 Page
            with p_cols[2]:
                st.caption(f"Muka Surat {len(doc_preview)-1}: P50-P90 Evaluation")
                img_p50 = doc_preview[-2 if len(doc_preview)>=10 else -1].get_pixmap(dpi=120).tobytes("png")
                st.image(img_p50, use_container_width=True)

        else:
            st.info("👈 Sila muat naik fail Excel projek yang telah disemak di atas untuk menjana Laporan PVsyst PDF.")


if __name__ == "__main__":
    main()
