"""
FinCorp Analytics — Suite de Análisis Financiero Integral de Empresas
Construido con Streamlit, Pandas y Plotly.
Incluye:
- Módulo de Autenticación y Cuentas de Usuario (SQLite y hashing SHA-256)
- Módulo de Exportación Multiformato (Excel .xlsx, Word .docx y PDF)
- Análisis Estructural (Vertical y Horizontal - NC 3 / NC 6)
- Ratios Financieros y Descomposición DuPont (3 y 5 pasos con promedios)
- Flujo de Efectivo Libre (FELF & FELP con Cascada Waterfall)
- Diagnóstico Ejecutivo Automatizado con Semáforos
- Editor de Datos Interactivo
"""

import streamlit as st
import pandas as pd
import numpy as np
import io

from modules.data_loader import (
    load_default_data,
    parse_uploaded_file,
    get_years_from_df,
    validate_accounting_equation,
    generate_template_csv,
    generate_template_excel
)
from modules.financial_math import (
    get_balance_totals,
    get_balance_averages,
    get_income_statement_totals,
    calculate_vertical_analysis,
    calculate_horizontal_analysis,
    calculate_financial_ratios,
    calculate_dupont_3_step,
    calculate_dupont_5_step,
    calculate_dupont_attribution,
    calculate_felf_felp,
    calculate_fcf_and_valuation,
    DEFAULT_UFV_RATES
)
from modules.charts import (
    create_felf_waterfall,
    create_felp_waterfall,
    create_fcf_waterfall,
    create_dupont_attribution_chart,
    create_roe_roa_trend_chart,
    create_margins_trend_chart,
    create_liquidity_chart,
    create_cash_conversion_cycle_chart,
    create_dupont_breakdown_chart,
    create_financial_radar
)
from modules.interpretations import (
    evaluate_financial_health,
    generate_executive_summary
)
from modules.auth import (
    authenticate_user,
    register_user
)
from modules.exporter import (
    generate_excel_report,
    generate_word_report,
    generate_pdf_report
)

# ---------------------------------------------------------
# Configuración General de la Página
# ---------------------------------------------------------
st.set_page_config(
    page_title="FinCorp Analytics | Suite Financiera",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# Estilos CSS Personalizados
# ---------------------------------------------------------
st.markdown("""
<meta name="google" content="notranslate">
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    .header-box {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #0284C7 100%);
        padding: 22px 28px;
        border-radius: 14px;
        color: white;
        margin-bottom: 22px;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
    }
    .header-title {
        font-size: 26px;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .header-sub {
        font-size: 13.5px;
        color: #93C5FD;
        margin-top: 5px;
        margin-bottom: 0;
    }

    .kpi-card {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 16px 18px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .kpi-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08);
    }
    .kpi-title {
        font-size: 11.5px;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 22px;
        font-weight: 700;
        color: #0F172A;
        margin: 5px 0;
    }
    .kpi-delta-pos {
        font-size: 11.5px;
        font-weight: 600;
        color: #10B981;
    }
    .kpi-delta-neg {
        font-size: 11.5px;
        font-weight: 600;
        color: #EF4444;
    }

    .badge-card {
        padding: 14px 18px;
        border-radius: 10px;
        margin-bottom: 12px;
        display: flex;
        align-items: flex-start;
        gap: 14px;
    }
    .badge-green {
        background-color: #ECFDF5;
        border-left: 5px solid #10B981;
    }
    .badge-orange {
        background-color: #FFFBEB;
        border-left: 5px solid #F59E0B;
    }
    .badge-red {
        background-color: #FEF2F2;
        border-left: 5px solid #EF4444;
    }
    
    .auth-container {
        max-width: 480px;
        margin: 40px auto;
        background: #FFFFFF;
        padding: 35px;
        border-radius: 16px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 15px 30px rgba(0, 0, 0, 0.08);
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# Gestión del Estado de Sesión y Autenticación
# ---------------------------------------------------------
if "user_authenticated" not in st.session_state:
    st.session_state.user_authenticated = False
    st.session_state.username = ""
    st.session_state.user_email = ""
    
if "financial_data" not in st.session_state:
    st.session_state.financial_data = load_default_data()

# ---------------------------------------------------------
# PANTALLA DE AUTENTICACIÓN (LOGIN / REGISTRO)
# ---------------------------------------------------------
if not st.session_state.user_authenticated:
    st.markdown("""
    <div style="text-align: center; margin-top: 25px; margin-bottom: 10px;">
        <h1 style="color: #1E3A8A; font-weight: 800; font-size: 34px; margin-bottom: 4px;">FinCorp Analytics</h1>
        <p style="color: #64748B; font-size: 16px;">Plataforma Corporativa de Análisis Financiero Integral y Valoración</p>
    </div>
    """, unsafe_allow_html=True)
    
    col_l, col_center, col_r = st.columns([1, 1.8, 1])
    with col_center:
        auth_tab_login, auth_tab_reg = st.tabs(["🔑 Iniciar Sesión", "📝 Crear Cuenta Nueva"])
        
        with auth_tab_login:
            st.markdown("#### Ingreso a la Plataforma")
            login_user = st.text_input("Usuario o Correo Electrónico:", key="input_login_user", placeholder="ej. demo o juan_perez")
            login_pass = st.text_input("Contraseña:", type="password", key="input_login_pass", placeholder="••••••••")
            
            if st.button("Iniciar Sesión", type="primary", use_container_width=True):
                user_info = authenticate_user(login_user, login_pass)
                if user_info:
                    st.session_state.user_authenticated = True
                    st.session_state.username = user_info["username"]
                    st.session_state.user_email = user_info["email"]
                    st.success(f"¡Bienvenido, {user_info['username'].capitalize()}! Ingresando a la plataforma...")
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas. Verifica tu usuario y contraseña.")
                    
            st.markdown("---")
            st.caption("⚡ **Acceso Rápido de Demostración:**")
            if st.button("🚀 Ingresar con Cuenta Demo", type="secondary", use_container_width=True):
                st.session_state.user_authenticated = True
                st.session_state.username = "demo"
                st.session_state.user_email = "demo@fincorp.com"
                st.rerun()
                    
        with auth_tab_reg:
            st.markdown("#### Registro de Nueva Cuenta")
            reg_user = st.text_input("Nombre de Usuario:", key="input_reg_user", placeholder="ej. juan_perez")
            reg_email = st.text_input("Correo Electrónico:", key="input_reg_email", placeholder="ej. juan@empresa.com")
            reg_pass = st.text_input("Contraseña (mínimo 6 caracteres):", type="password", key="input_reg_pass", placeholder="••••••••")
            
            if st.button("Crear Cuenta", type="primary", use_container_width=True):
                ok, msg = register_user(reg_user, reg_email, reg_pass)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)
                    
    # Detener la ejecución del resto de la app hasta que el usuario se autentique
    st.stop()

# ---------------------------------------------------------
# BARRA LATERAL (SIDEBAR): USUARIO Y ARCHIVOS
# ---------------------------------------------------------
with st.sidebar:
    # Perfil del Usuario Autenticado
    st.markdown(f"""
    <div style="background: #F1F5F9; border-radius: 10px; padding: 14px; margin-bottom: 15px; border: 1px solid #CBD5E1;">
        <div style="font-weight: 700; color: #0F172A; font-size: 15px;">👤 {st.session_state.username.capitalize()}</div>
        <div style="color: #64748B; font-size: 12px;">{st.session_state.user_email}</div>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.user_authenticated = False
        st.rerun()
            
    st.markdown("---")
    st.subheader("📂 Datos Financieros")
    
    modo_fuente = st.radio(
        "Fuente de los Estados Financieros:",
        ["📊 Datos Modelo (Demo 3 Años)", "📁 Cargar Archivo (Excel / CSV)"],
        index=0
    )
    
    if modo_fuente == "📁 Cargar Archivo (Excel / CSV)":
        uploaded_file = st.file_uploader(
            "Selecciona un archivo .csv o .xlsx",
            type=["csv", "xlsx", "xls"]
        )
        if uploaded_file is not None:
            file_sig = f"{uploaded_file.name}_{uploaded_file.size}"
            if st.session_state.get("last_uploaded_sig") != file_sig:
                try:
                    loaded_df = parse_uploaded_file(uploaded_file)
                    st.session_state.financial_data = loaded_df
                    st.session_state.last_uploaded_sig = file_sig
                    new_years = [str(y).strip() for y in get_years_from_df(loaded_df)]
                    if new_years:
                        st.session_state.selected_focal_year = str(new_years[-1])
                        st.session_state.focal_year_widget = str(new_years[-1])
                    st.success("✅ Archivo cargado exitosamente.")
                except Exception as e:
                    st.error(f"Error al procesar el archivo: {e}")
    else:
        if st.button("🔄 Restaurar Datos Predeterminados"):
            st.session_state.financial_data = load_default_data()
            default_years = [str(y).strip() for y in get_years_from_df(st.session_state.financial_data)]
            if default_years:
                st.session_state.selected_focal_year = str(default_years[-1])
                st.session_state.focal_year_widget = str(default_years[-1])
            st.success("Datos modelo restaurados.")

    df_active = st.session_state.financial_data
    years = [str(y).strip() for y in get_years_from_df(df_active)]
    
    if not years:
        st.error("No se detectaron columnas de periodos válidas.")
        st.stop()
        
    st.markdown("---")
    # Sincronización robusta y bidireccional del Periodo Focal seleccionado
    if "selected_focal_year" not in st.session_state or str(st.session_state.selected_focal_year) not in years:
        st.session_state.selected_focal_year = str(years[-1])

    def _on_focal_year_change():
        if "focal_year_widget" in st.session_state:
            st.session_state.selected_focal_year = str(st.session_state.focal_year_widget)

    try:
        focal_index = years.index(str(st.session_state.selected_focal_year))
    except (ValueError, IndexError):
        focal_index = len(years) - 1
        st.session_state.selected_focal_year = str(years[focal_index])

    st.selectbox(
        "Periodo Focal:",
        options=years,
        index=focal_index,
        key="focal_year_widget",
        on_change=_on_focal_year_change
    )
    selected_year = str(st.session_state.selected_focal_year)
    company_name = st.text_input("Razón Social de la Empresa:", value="Empresa Industrial Andina S.A.")
    nit_val = st.text_input("NIT / Identificación Fiscal:", value="1028475021")
    
    # Validación contable
    val_status = validate_accounting_equation(df_active)
    if val_status["is_balanced"]:
        st.success("⚖️ Balance Cuadrado (Activos = Pasivo + Pat.)")
    else:
        st.warning("⚠️ Descuadre Contable detectado.")
        
    st.markdown("---")
    st.subheader("📥 Descargar Plantillas")
    col_csv, col_xls = st.columns(2)
    with col_csv:
        st.download_button(
            label="📄 CSV",
            data=generate_template_csv(),
            file_name="plantilla_financiera.csv",
            mime="text/csv",
            use_container_width=True
        )
    with col_xls:
        try:
            excel_bytes = generate_template_excel()
            st.download_button(
                label="📊 Excel",
                data=excel_bytes,
                file_name="plantilla_financiera.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception:
            pass

# ---------------------------------------------------------
# Procesamiento de Cálculos Globales
# ---------------------------------------------------------
bal_totals = get_balance_totals(df_active)
bal_averages = get_balance_averages(df_active)
inc_totals = get_income_statement_totals(df_active)
ratios_df = calculate_financial_ratios(df_active)
dupont_3_df = calculate_dupont_3_step(df_active)
dupont_5_df = calculate_dupont_5_step(df_active)
felf_df, felp_df, comp_df, fcf_components = calculate_felf_felp(df_active)
fcf_df = felf_df
attr_result = calculate_dupont_attribution(df_active, years[0], selected_year) if len(years) > 1 else None

# ---------------------------------------------------------
# Encabezado Principal y Pasarela de Pago
# ---------------------------------------------------------
st.markdown("""
<div class="header-box">
    <h1 class="header-title">FinCorp Analytics — Suite de Análisis Financiero Integral</h1>
    <p class="header-sub">Análisis Estructural (NC 3 / NC 6) • DuPont Jerárquico con Atribución • FELF & FELP • Diagnóstico Automatizado • Exportaciones</p>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# Pestañas de Navegación
# ---------------------------------------------------------
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Resumen & Exportación",
    "📑 Análisis Estructural (NC 3/6)",
    "📐 Ratios & DuPont Jerárquico",
    "💵 Flujos Libres (FELF & FELP)",
    "🧠 Diagnóstico & Semáforo",
    "✏️ Editor Interactivo"
])

# =========================================================
# TAB 1: RESUMEN EJECUTIVO & EXPORTACIÓN MULTIFORMATO
# =========================================================
with tab1:
    st.subheader(f"Resumen de Desempeño Corporativo — Ejercicio {selected_year}")
    
    prev_idx = years.index(selected_year) - 1
    has_prev = prev_idx >= 0
    prev_year = years[prev_idx] if has_prev else None
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    # Ventas
    ventas_val = inc_totals["Ventas"].get(selected_year, 0)
    ventas_prev = inc_totals["Ventas"].get(prev_year, 0) if has_prev else None
    ventas_delta = ((ventas_val - ventas_prev) / abs(ventas_prev) * 100) if (has_prev and ventas_prev) else None
    with col1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">Ingresos por Ventas</div>
            <div class="kpi-value">${ventas_val:,.0f}</div>
            <div class="{ 'kpi-delta-pos' if (ventas_delta or 0) >= 0 else 'kpi-delta-neg' }">
                { f'{ventas_delta:+.1f}% vs {prev_year}' if ventas_delta is not None else 'Año Base' }
            </div>
        </div>
        """, unsafe_allow_html=True)

    # EBITDA
    ebitda_val = inc_totals["EBITDA"].get(selected_year, 0)
    ebitda_prev = inc_totals["EBITDA"].get(prev_year, 0) if has_prev else None
    ebitda_delta = ((ebitda_val - ebitda_prev) / abs(ebitda_prev) * 100) if (has_prev and ebitda_prev) else None
    with col2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">EBITDA Generado</div>
            <div class="kpi-value">${ebitda_val:,.0f}</div>
            <div class="{ 'kpi-delta-pos' if (ebitda_delta or 0) >= 0 else 'kpi-delta-neg' }">
                { f'{ebitda_delta:+.1f}% vs {prev_year}' if ebitda_delta is not None else 'Año Base' }
            </div>
        </div>
        """, unsafe_allow_html=True)

    # Margen Operativo
    mo_series = ratios_df[ratios_df["Ratio"] == "Margen Operativo"].iloc[0]
    mo_val = mo_series[selected_year]
    mo_prev = mo_series[prev_year] if has_prev else None
    mo_delta = (mo_val - mo_prev) if has_prev else None
    with col3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">Margen Operativo</div>
            <div class="kpi-value">{mo_val:.1f}%</div>
            <div class="{ 'kpi-delta-pos' if (mo_delta or 0) >= 0 else 'kpi-delta-neg' }">
                { f'{mo_delta:+.1f} p.p. vs {prev_year}' if mo_delta is not None else 'Año Base' }
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ROE
    roe_series = ratios_df[ratios_df["Ratio"] == "ROE"].iloc[0]
    roe_val = roe_series[selected_year]
    roe_prev = roe_series[prev_year] if has_prev else None
    roe_delta = (roe_val - roe_prev) if has_prev else None
    with col4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">ROE (Retorno Patrimonio)</div>
            <div class="kpi-value">{roe_val:.1f}%</div>
            <div class="{ 'kpi-delta-pos' if (roe_delta or 0) >= 0 else 'kpi-delta-neg' }">
                { f'{roe_delta:+.1f} p.p. vs {prev_year}' if roe_delta is not None else 'Año Base' }
            </div>
        </div>
        """, unsafe_allow_html=True)

    # Flujo Libre Firma (FELF)
    felf_val = fcf_components["felf_op1"].get(selected_year, 0)
    with col5:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">Flujo Libre Firma (FELF)</div>
            <div class="kpi-value">${felf_val:,.0f}</div>
            <div class="{ 'kpi-delta-pos' if felf_val >= 0 else 'kpi-delta-neg' }">
                { 'Generación Libre Positiva' if felf_val >= 0 else 'Déficit de Flujo Libre' }
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    col_chart_left, col_chart_right = st.columns([1, 1])
    with col_chart_left:
        fig_radar = create_financial_radar(ratios_df, selected_year)
        st.plotly_chart(fig_radar, use_container_width=True)
    with col_chart_right:
        fig_roe_roa = create_roe_roa_trend_chart(ratios_df, years)
        st.plotly_chart(fig_roe_roa, use_container_width=True)

    st.markdown("---")
    st.subheader("📥 Módulo de Exportación y Descarga de Informes")
    st.caption("Descarga el análisis financiero completo en los 3 formatos corporativos con gráficos, fórmulas nativas y diagnóstico DuPont:")
    
    # Preparar figuras analíticas para incrustación en reportes
    fig_mar = create_margins_trend_chart(ratios_df, years)
    fig_liq = create_liquidity_chart(ratios_df, years)
    fig_felf = create_felf_waterfall(fcf_components, selected_year, "Opción 1")
    if attr_result:
        summary_attr, attr_df_data = attr_result
        fig_attr = create_dupont_attribution_chart(attr_df_data, selected_year, years[0])
    else:
        summary_attr, attr_df_data = None, None
        fig_attr = None
    
    figures_dict = {
        "radar": fig_radar,
        "roe_roa": fig_roe_roa,
        "margins": fig_mar,
        "liquidity": fig_liq,
        "felf": fig_felf,
        "dupont_attr": fig_attr
    }
    
    col_exp_xl, col_exp_doc, col_exp_pdf = st.columns(3)
    
    # 1. Exportación Excel (Modelo Dinámico con 5 Pestañas)
    with col_exp_xl:
        try:
            excel_data = generate_excel_report(
                df_active, ratios_df, dupont_3_df, dupont_5_df, felf_df, years,
                felp_df=felp_df, comp_df=comp_df,
                attr_df=attr_df_data,
                company_name=company_name, nit=nit_val, base_year=years[0]
            )
            st.download_button(
                label="📊 Descargar Excel (.xlsx)",
                data=excel_data,
                file_name=f"analisis_financiero_{selected_year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.caption("✅ 5 Pestañas: Dashboard, EEFF V/H, 40 Ratios, DuPont y Flujos Libres.")
        except Exception as e:
            st.error(f"Error generando Excel: {e}")
            
    # 2. Exportación Word (Informe Narrativo con Gráficos)
    with col_exp_doc:
        try:
            summary_text = generate_executive_summary(ratios_df, fcf_components, dupont_3_df, years)
            word_data = generate_word_report(
                df_active, ratios_df, dupont_3_df, dupont_5_df, fcf_components, years,
                summary_text, company_name=company_name, nit=nit_val,
                attr_result=attr_result, felf_df=felf_df, felp_df=felp_df, comp_df=comp_df,
                figures_dict=figures_dict, base_year=years[0]
            )
            st.download_button(
                label="📝 Descargar Word (.docx)",
                data=word_data,
                file_name=f"informe_ejecutivo_{selected_year}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            st.caption("✅ Portada, índice, gráficos PNG, diagnóstico DuPont y recomendaciones.")
        except Exception as e:
            st.error(f"Error generando Word: {e}")

    # 3. Exportación PDF (Reporte Formal Imprimible)
    with col_exp_pdf:
        try:
            summary_text = generate_executive_summary(ratios_df, fcf_components, dupont_3_df, years)
            pdf_data = generate_pdf_report(
                df_active, ratios_df, dupont_3_df, dupont_5_df, fcf_components, years,
                summary_text, company_name=company_name, nit=nit_val,
                attr_result=attr_result, felf_df=felf_df, felp_df=felp_df,
                figures_dict=figures_dict, base_year=years[0]
            )
            st.download_button(
                label="📑 Descargar PDF (.pdf)",
                data=pdf_data,
                file_name=f"reporte_financiero_{selected_year}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            st.caption("✅ Formato Carta, paginación dinámica 'Página X de Y' y gráficos.")
        except Exception as e:
            st.error(f"Error generando PDF: {e}")

# =========================================================
# TAB 2: ANÁLISIS ESTRUCTURAL (VERTICAL & HORIZONTAL)
# =========================================================
with tab2:
    st.subheader("Análisis Estructural de Estados Financieros")
    
    tipo_analisis = st.radio(
        "Seleccionar Enfoque Analítico:",
        ["📐 Análisis Vertical (% de Participación)", "📈 Análisis Horizontal (Variación $ y %)"],
        horizontal=True
    )
    
    col_sel_est, col_empty = st.columns([1, 2])
    with col_sel_est:
        estado_sel = st.selectbox("Estado Financiero a Explorar:", ["Balance General", "Estado de Resultados"])

    if "Vertical" in tipo_analisis:
        st.markdown(f"#### Análisis Vertical — {estado_sel}")
        st.caption("📌 Expresa cada cuenta del Balance como % del Activo Total y cada cuenta del Estado de Resultados como % de los Ingresos Operativos.")
        df_vert = calculate_vertical_analysis(df_active, estado_sel)
        format_dict = {}
        for c in df_vert.columns:
            if "($)" in c:
                format_dict[c] = "${:,.0f}"
            elif "(%)" in c:
                format_dict[c] = "{:.2f}%"
        st.dataframe(df_vert.style.format(format_dict), use_container_width=True, height=450)
        
    else:
        st.markdown(f"#### Análisis Horizontal — {estado_sel} (Normas Contables Bolivianas NC 3 / NC 6)")
        
        col_h1, col_h2, col_h3 = st.columns([1.2, 1.2, 1.8])
        with col_h1:
            modo_comp = st.selectbox(
                "Modo de Comparación:",
                ["Año Anterior Inmediato (t vs t-1)", "Año Base Fijo (vs t0)"],
                key="sel_modo_comp"
            )
        with col_h2:
            base_yr_sel = years[0]
            if "Año Base" in modo_comp:
                base_options = years[:-1] if len(years) > 1 else years
                b_idx = 0
                if "sel_base_yr" in st.session_state and st.session_state.sel_base_yr in base_options:
                    b_idx = base_options.index(st.session_state.sel_base_yr)
                base_yr_sel = st.selectbox("Año Base Fijo de Referencia:", base_options, index=b_idx, key="sel_base_yr")
            else:
                st.caption("Comparación secuencial entre gestiones continuas.")
        with col_h3:
            ajustar_ufv = st.checkbox(
                "Ajuste por Inflación (NC 3 / NC 6 - UFV)",
                value=False,
                help="Reexpresa los montos del periodo de referencia a poder adquisitivo de cierre según la razón de UFVs oficiales del BCB (Norma Contable N° 3 de Estados Financieros a Moneda Constante)."
            )
            if ajustar_ufv:
                st.caption("🔹 Moneda Constante activada. Factor de ajuste: UFV_cierre / UFV_base.")
                
        comp_param = "base" if "Año Base" in modo_comp else "previous"
        df_horiz = calculate_horizontal_analysis(
            df_active, estado_sel, compare_mode=comp_param, base_year=base_yr_sel, apply_ufv=ajustar_ufv
        )
        
        format_dict = {}
        for c in df_horiz.columns:
            if "($)" in c:
                format_dict[c] = "${:,.0f}"
            elif "(%)" in c:
                format_dict[c] = "{:+.2f}%"
                
        def color_variations(val):
            if isinstance(val, (int, float)):
                if val > 15:
                    return 'background-color: #D1FAE5; color: #065F46; font-weight: bold;'
                elif val < -15:
                    return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'
            return ''

        styler = df_horiz.style.format(format_dict)
        if hasattr(styler, "map"):
            styled_table = styler.map(color_variations, subset=[c for c in df_horiz.columns if "(%)" in c])
        else:
            styled_table = styler.applymap(color_variations, subset=[c for c in df_horiz.columns if "(%)" in c])

        st.dataframe(styled_table, use_container_width=True, height=450)

# =========================================================
# TAB 3: RATIOS FINANCIEROS & DESCOMPOSICIÓN DUPONT
# =========================================================
with tab3:
    st.subheader("Indicadores Financieros & Descomposición DuPont")
    
    subtab_ratios, subtab_dupont = st.tabs(["📊 Razones Financieras (40 Ratios con Promedios)", "🧬 Descomposición DuPont & Atribución Causa Raíz"])
    
    with subtab_ratios:
        st.info("💡 **Criterio de Medición Contable**: Los ratios que vinculan cuentas de balance (Situación Financiera) utilizan el **PROMEDIO** entre el inicio y cierre de gestión ((Inicio + Cierre) / 2) según la técnica financiera estándar.")
        
        cat_filtro = st.multiselect(
            "Filtrar por Categoría de Ratios:",
            options=["Actividad", "Liquidez", "Rentabilidad", "Solvencia"],
            default=["Actividad", "Liquidez", "Rentabilidad", "Solvencia"]
        )
        
        df_ratios_filtered = ratios_df[ratios_df["Categoría"].isin(cat_filtro)]
        format_ratios = {y: "{:.2f}" for y in years}
        st.dataframe(df_ratios_filtered.style.format(format_ratios), use_container_width=True, height=450)
        
        st.markdown("---")
        st.subheader("Visualizaciones de Ratios Operativos y Financieros")
        col_g1, col_g2, col_g3 = st.columns(3)
        with col_g1:
            fig_liq = create_liquidity_chart(ratios_df, years)
            st.plotly_chart(fig_liq, use_container_width=True)
        with col_g2:
            fig_mar = create_margins_trend_chart(ratios_df, years)
            st.plotly_chart(fig_mar, use_container_width=True)
        with col_g3:
            fig_ccc = create_cash_conversion_cycle_chart(ratios_df, years)
            st.plotly_chart(fig_ccc, use_container_width=True)

    with subtab_dupont:
        dup_view = st.radio(
            "Seleccionar Módulo DuPont:",
            ["🧬 DuPont 3 y 5 Pasos (con Promedios)", "🌳 Desagregación Jerárquica", "🎯 Atribución Causa Raíz (Sensibilidad de Cuentas)"],
            horizontal=True
        )
        
        if "3 y 5 Pasos" in dup_view:
            st.markdown("### Sistema DuPont de 3 Pasos (con Promedios de Balance)")
            st.latex(r"\text{ROE (Rentabilidad sobre Patrimonio)} = \text{Margen Neto} \times \text{Rotación de Activos} \times \text{Multiplicador de Capital}")
            format_d3 = {y: "{:.2f}" for y in years}
            st.dataframe(dupont_3_df.style.format(format_d3), use_container_width=True)
            
            fig_dup_break = create_dupont_breakdown_chart(dupont_3_df, years)
            st.plotly_chart(fig_dup_break, use_container_width=True)
            
            st.markdown("---")
            st.markdown("### Sistema DuPont Extendido de 5 Pasos")
            st.latex(r"\text{ROE (Rentabilidad sobre Patrimonio)} = \left(\frac{\text{Resultado Neto}}{\text{EBT}}\right) \times \left(\frac{\text{EBT}}{\text{EBIT}}\right) \times \left(\frac{\text{EBIT}}{\text{Ingresos}}\right) \times \left(\frac{\text{Ingresos}}{\text{Prom. Activos}}\right) \times \left(\frac{\text{Prom. Activos}}{\text{Prom. Patrimonio}}\right)")
            format_d5 = {y: "{:.3f}" for y in years}
            st.dataframe(dupont_5_df.style.format(format_d5), use_container_width=True)
            
        elif "Desagregación Jerárquica" in dup_view:
            st.markdown("### Estructura Jerárquica de Descomposición DuPont")
            st.caption("Desglose piramidal de rentabilidad conectado con los componentes del Estado de Resultados y Balance General:")
            
            col_h_roe, col_h_roa = st.columns(2)
            with col_h_roe:
                st.markdown(f"""
                <div style="background: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 12px; padding: 18px; margin-bottom: 15px;">
                    <h4 style="color: #1E3A8A; margin-top: 0;">1. Rentabilidad sobre Patrimonio (ROE)</h4>
                    <p style="font-size: 14px; margin-bottom: 8px;"><b>ROE = ROA × Apalancamiento Financiero</b></p>
                    <ul style="font-size: 13px; color: #334155; line-height: 1.6;">
                        <li><b>ROA:</b> Retorno sobre activos totales (Margen Neto × Rotación).</li>
                        <li><b>Apalancamiento:</b> Promedio de Activos Totales / Promedio de Patrimonio Neto.</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
            with col_h_roa:
                st.markdown(f"""
                <div style="background: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 12px; padding: 18px; margin-bottom: 15px;">
                    <h4 style="color: #0D9488; margin-top: 0;">2. Rentabilidad sobre Activos (ROA)</h4>
                    <p style="font-size: 14px; margin-bottom: 8px;"><b>ROA = Margen Neto × Rotación de Activos</b></p>
                    <ul style="font-size: 13px; color: #334155; line-height: 1.6;">
                        <li><b>Margen Neto:</b> Resultado Neto / Ingresos Operativos.</li>
                        <li><b>Rotación de Activos:</b> Ingresos Operativos / Promedio de Activos Totales.</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
                
            col_d_res, col_d_act, col_d_pas = st.columns(3)
            with col_d_res:
                st.markdown(f"""
                <div style="background: #EFF6FF; border: 1px solid #93C5FD; border-radius: 10px; padding: 14px;">
                    <h5 style="color: #1E3A8A; margin-top: 0;">3. Desagregación del Resultado Neto</h5>
                    <p style="font-size: 12px; color: #475569;">(Ingresos y Ganancias) − (Gastos y Pérdidas)</p>
                    <ul style="font-size: 12.5px; color: #1E293B;">
                        <li>(+) Ingresos por Ventas</li>
                        <li>(-) Costo de lo Vendido</li>
                        <li>(-) Gastos Operativos (Adm. y Ventas)</li>
                        <li>(-) Depreciaciones y Amortizaciones</li>
                        <li>(-) Gastos Financieros (Intereses)</li>
                        <li>(-) Impuesto sobre la Renta (IUE)</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
            with col_d_act:
                st.markdown(f"""
                <div style="background: #F0FDF4; border: 1px solid #86EFAC; border-radius: 10px; padding: 14px;">
                    <h5 style="color: #166534; margin-top: 0;">4. Desagregación de Activos Totales</h5>
                    <p style="font-size: 12px; color: #475569;">Activos Corrientes vs. No Corrientes</p>
                    <ul style="font-size: 12.5px; color: #1E293B;">
                        <li><b>Activos Corrientes:</b>
                            <br>• Disponibilidades (Efectivo / ICP)
                            <br>• Cuentas por Cobrar
                            <br>• Inventarios
                            <br>• Otros Activos Corrientes
                        </li>
                        <li><b>Activos No Corrientes:</b>
                            <br>• Inversiones de Capital (PPE Neto)
                            <br>• Intangibles y Otros LP
                        </li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
            with col_d_pas:
                st.markdown(f"""
                <div style="background: #FEF2F2; border: 1px solid #FCA5A5; border-radius: 10px; padding: 14px;">
                    <h5 style="color: #991B1B; margin-top: 0;">5. Desagregación de Pasivos Totales</h5>
                    <p style="font-size: 12px; color: #475569;">Pasivos Corrientes vs. No Corrientes</p>
                    <ul style="font-size: 12.5px; color: #1E293B;">
                        <li><b>Pasivos Corrientes:</b>
                            <br>• Cuentas por Pagar Comerciales
                            <br>• Deuda Financiera Corto Plazo
                            <br>• Pasivos Laborales e Impuestos
                        </li>
                        <li><b>Pasivos No Corrientes:</b>
                            <br>• Deuda Financiera Largo Plazo
                            <br>• Otros Pasivos Largo Plazo
                        </li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
                
        else:
            st.markdown("### Análisis de Sensibilidad y Atribución Causa Raíz de Variaciones")
            st.caption("Identifica con precisión matemática qué cuentas contables específicas impulsaron o deterioraron el ROE y ROA entre dos periodos:")
            
            col_sel_p1, col_sel_p2 = st.columns(2)
            with col_sel_p1:
                base_options = years[:-1] if len(years) > 1 else years
                base_idx = 0
                if "attr_base_yr" in st.session_state and str(st.session_state.attr_base_yr) in base_options:
                    base_idx = base_options.index(str(st.session_state.attr_base_yr))
                base_attr_yr = st.selectbox("Periodo Base:", base_options, index=base_idx, key="attr_base_yr")
            with col_sel_p2:
                target_options = years[1:] if len(years) > 1 else years
                target_idx = len(target_options) - 1
                if "attr_target_yr" in st.session_state and str(st.session_state.attr_target_yr) in target_options:
                    target_idx = target_options.index(str(st.session_state.attr_target_yr))
                elif selected_year in target_options:
                    target_idx = target_options.index(selected_year)
                target_attr_yr = st.selectbox("Periodo Comparativo:", target_options, index=target_idx, key="attr_target_yr")
                
            summary_attr, attr_df = calculate_dupont_attribution(df_active, base_attr_yr, target_attr_yr)
            
            kcol1, kcol2, kcol3, kcol4 = st.columns(4)
            with kcol1:
                d_roe = summary_attr["delta_roe"]
                st.metric(
                    f"Variación Neta ROE ({target_attr_yr} vs {base_attr_yr})",
                    f"{summary_attr['roe_target']:.2f}%",
                    f"{d_roe:+.2f} p.p."
                )
            with kcol2:
                st.metric(
                    "Efecto Margen Neto",
                    f"{summary_attr['margen_neto_target']:.2f}%",
                    f"{summary_attr['efecto_margen']:+.2f} p.p."
                )
            with kcol3:
                st.metric(
                    "Efecto Rotación de Activos",
                    f"{summary_attr['rotacion_activos_target']:.2f}x",
                    f"{summary_attr['efecto_rotacion']:+.2f} p.p."
                )
            with kcol4:
                st.metric(
                    "Efecto Apalancamiento",
                    f"{summary_attr['multiplicador_capital_target']:.2f}x",
                    f"{summary_attr['efecto_apalancamiento']:+.2f} p.p."
                )
                
            st.markdown("<br>", unsafe_allow_html=True)
            fig_attr = create_dupont_attribution_chart(attr_df, target_attr_yr, base_attr_yr)
            st.plotly_chart(fig_attr, use_container_width=True)
            
            st.markdown("#### Matriz Detallada de Cuentas de Impacto (Ranking de Drivers y Detractores)")
            
            def color_impact(val):
                if isinstance(val, (int, float)):
                    if val > 0.05:
                        return 'background-color: #D1FAE5; color: #065F46; font-weight: bold;'
                    elif val < -0.05:
                        return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'
                return ''
                
            format_attr = {
                "Variación ($)": "${:,.0f}",
                "Impacto Estimado ROE (p.p.)": "{:+.2f} p.p.",
                "Impacto Estimado ROA (p.p.)": "{:+.2f} p.p."
            }
            
            styler_attr = attr_df.style.format(format_attr)
            if hasattr(styler_attr, "map"):
                styled_attr_table = styler_attr.map(color_impact, subset=["Impacto Estimado ROE (p.p.)", "Impacto Estimado ROA (p.p.)"])
            else:
                styled_attr_table = styler_attr.applymap(color_impact, subset=["Impacto Estimado ROE (p.p.)", "Impacto Estimado ROA (p.p.)"])
                
            st.dataframe(styled_attr_table, use_container_width=True)

# =========================================================
# TAB 4: FLUJO DE CAJA LIBRE & EBITDA
# =========================================================
# =========================================================
# TAB 4: FLUJOS DE EFECTIVO LIBRES (FELF & FELP)
# =========================================================
with tab4:
    st.subheader("Flujos de Efectivo Libres Corporativos (FELF & FELP)")
    
    tipo_flujo = st.radio(
        "Seleccionar Perspectiva de Flujo de Efectivo Libre:",
        ["🏢 Flujo Libre de la Firma (FELF / FCFF)", "👥 Flujo Libre del Patrimonio (FELP / FCFE)", "⚖️ Conciliación Comparativa (Opciones 1 vs 2)"],
        horizontal=True
    )
    
    format_curr = {y: "${:,.0f}" for y in years}
    
    if "FELF" in tipo_flujo:
        col_met1, col_met2 = st.columns([1.5, 2.5])
        with col_met1:
            metodo_felf = st.selectbox(
                "Metodología de Cálculo FELF:",
                ["Opción 1 (Base RAII / EBIT)", "Opción 2 (Base FEAO / Operativo)"],
                key="sel_metodo_felf"
            )
        with col_met2:
            if "Opción 1" in metodo_felf:
                st.latex(r"\text{FELF (Flujo Libre de la Firma)} = \text{RAII} - \text{IR} + \text{PNGME} - \text{CAPEX} \pm \text{MNAPO}")
            else:
                st.latex(r"\text{FELF (Flujo Libre de la Firma)} = \text{FEAO} + \text{GF} \times (1 - \tau) - \text{CAPEX}")
                
        fig_felf = create_felf_waterfall(fcf_components, selected_year, method=metodo_felf)
        st.plotly_chart(fig_felf, use_container_width=True)
        
        st.markdown("---")
        st.markdown(f"#### Estado de Determinación del FELF — Ejercicio {selected_year}")
        st.dataframe(felf_df.style.format(format_curr), use_container_width=True)
        
        st.markdown("""
        **Glosario Metodológico FELF:**
        - **RAII:** Resultado Antes de Intereses e Impuestos (EBIT operativo).
        - **IR:** Impuesto Reexpresado / Calculado sobre la operación ($\text{RAII} \times \tau$).
        - **PNGME:** Provisiones y Gastos No Monetarios del Ejercicio (Depreciaciones y Amortizaciones).
        - **CAPEX:** Inversiones netas de capital en activo fijo.
        - **MNAPO:** Variación del Capital de Trabajo Neto Operativo ($\Delta\text{KTNO}$).
        - **FEAO:** Flujo de Efectivo generado por Actividades de Operación.
        - **GF $\times (1 - \tau)$:** Escudo fiscal derivado de los gastos financieros por intereses.
        """)
        
    elif "FELP" in tipo_flujo:
        col_met1, col_met2 = st.columns([1.5, 2.5])
        with col_met1:
            metodo_felp = st.selectbox(
                "Metodología de Cálculo FELP:",
                ["Opción 1 (Base Resultado Neto)", "Opción 2 (Base FEAO / Operativo)"],
                key="sel_metodo_felp"
            )
        with col_met2:
            if "Opción 1" in metodo_felp:
                st.latex(r"\text{FELP (Flujo Libre del Patrimonio)} = \text{RN} + \text{PNGME} - \text{CAPEX} \pm \text{MNAPO} \pm \Delta\text{D}")
            else:
                st.latex(r"\text{FELP (Flujo Libre del Patrimonio)} = \text{FEAO} - \text{CAPEX} \pm \Delta\text{D}")
                
        fig_felp = create_felp_waterfall(fcf_components, selected_year, method=metodo_felp)
        st.plotly_chart(fig_felp, use_container_width=True)
        
        st.markdown("---")
        st.markdown(f"#### Estado de Determinación del FELP — Ejercicio {selected_year}")
        st.dataframe(felp_df.style.format(format_curr), use_container_width=True)
        
        st.markdown("""
        **Glosario Metodológico FELP:**
        - **RN:** Resultado Neto del ejercicio atribuible a los accionistas.
        - **PNGME:** Provisiones y Gastos No Monetarios (D&A).
        - **CAPEX:** Inversiones en Activos Fijos del periodo.
        - **MNAPO:** Variación del Capital de Trabajo Neto Operativo ($\Delta\text{KTNO}$).
        - **$\Delta\text{D}$:** Variación o Servicio Neto de Deuda Financiera ($\text{Deuda}_t - \text{Deuda}_{t-1}$).
        - **FEAO:** Flujo de Efectivo de Actividades de Operación.
        """)
        
    else:
        st.markdown("#### Conciliación Comparativa: Opciones 1 vs. 2 de FELF y FELP")
        st.caption("Verificación de la congruencia matemática entre la estimación analítica y la basada en flujos de efectivo:")
        st.dataframe(comp_df.style.format(format_curr), use_container_width=True)
        
        st.markdown("---")
        col_fc1, col_fc2 = st.columns(2)
        with col_fc1:
            felf1_now = fcf_components["felf_op1"].get(selected_year, 0)
            felf2_now = fcf_components["felf_op2"].get(selected_year, 0)
            st.markdown(f"""
            <div style="background: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 10px; padding: 16px;">
                <h5 style="color: #1E3A8A; margin-top: 0;">🏢 FELF (Firma) en {selected_year}</h5>
                <p style="font-size: 15px; margin: 4px 0;"><b>Opción 1:</b> US$ {felf1_now:,.0f}</p>
                <p style="font-size: 15px; margin: 4px 0;"><b>Opción 2:</b> US$ {felf2_now:,.0f}</p>
                <p style="font-size: 13px; color: #64748B; margin-top: 6px;">Excedente disponible para remunerar a todos los proveedores de capital (bancos y accionistas).</p>
            </div>
            """, unsafe_allow_html=True)
        with col_fc2:
            felp1_now = fcf_components["felp_op1"].get(selected_year, 0)
            felp2_now = fcf_components["felp_op2"].get(selected_year, 0)
            st.markdown(f"""
            <div style="background: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 10px; padding: 16px;">
                <h5 style="color: #7C3AED; margin-top: 0;">👥 FELP (Patrimonio) en {selected_year}</h5>
                <p style="font-size: 15px; margin: 4px 0;"><b>Opción 1:</b> US$ {felp1_now:,.0f}</p>
                <p style="font-size: 15px; margin: 4px 0;"><b>Opción 2:</b> US$ {felp2_now:,.0f}</p>
                <p style="font-size: 13px; color: #64748B; margin-top: 6px;">Caja libre remanente tras atender necesidades operativas, reinversión y deuda neta.</p>
            </div>
            """, unsafe_allow_html=True)

# =========================================================
# TAB 5: DIAGNÓSTICO FINANCIERO & ALERTAS AUTOMÁTICAS
# =========================================================
with tab5:
    st.subheader(f"Semáforo de Salud Financiera & Alertas Tempranas — Periodo {selected_year}")
    
    diagnosticos = evaluate_financial_health(ratios_df, fcf_components, selected_year)
    for d in diagnosticos:
        color = d.get("color", d.get("Color", "green"))
        color_class = f"badge-{color}"
        estado_badge = d.get("Estado", "🟢 Normal")
        icono = estado_badge.split()[0] if estado_badge else "📌"
        pilar = d.get("Pilar", "General")
        indicador = d.get("Indicador", "")
        valor = d.get("Valor", "")
        interpretacion = d.get("Interpretación", d.get("interpretacion", ""))
        
        st.markdown(f"""
        <div class="badge-card {color_class}">
            <div style="font-size: 24px; min-width: 35px;">{icono}</div>
            <div>
                <div style="font-weight: 700; font-size: 15px; color: #0F172A;">
                    {pilar} — {indicador}: <span style="font-size: 16px; color: #1E3A8A;">{valor}</span> 
                    <span style="font-weight: 500; font-size: 13px; color: #64748B;">({estado_badge})</span>
                </div>
                <div style="font-size: 13.5px; color: #334155; margin-top: 4px; line-height: 1.4;">
                    {interpretacion}
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("---")
    st.subheader("🎯 Matriz de Recomendaciones Estratégicas")
    col_rec1, col_rec2 = st.columns(2)
    with col_rec1:
        st.markdown("""
        **⚡ Plan de Acción Operativo Inmediato:**
        - **Política de Crédito y Cobranzas**: Monitorear DSO para evitar atrasos que deterioren el flujo operativo.
        - **Gestión de Stock**: Optimizar inventarios de menor rotación para liberar efectivo atrapado en almacén.
        - **Negociación con Proveedores**: Mantener plazos de crédito que financien holgadamente el ciclo de inventarios (DPO > DIO).
        """)
    with col_rec2:
        st.markdown("""
        **🏛️ Directrices Financieras y de Estructura de Capital:**
        - **Capacidad de Endeudamiento**: Mantener la cobertura de intereses por encima de 3.0x para preservar grado de inversión.
        - **Disciplina de Inversión (CAPEX)**: Asegurar que todo proyecto de capital rinda una TIR superior al WACC corporativo.
        - **Calidad del ROE**: Fomentar que el crecimiento de la rentabilidad provenga de eficiencia en costos y rotación, y no exclusivamente de deuda.
        """)

# =========================================================
# TAB 6: EDITOR DE DATOS EN VIVO
# =========================================================
with tab6:
    st.subheader("Editor Interactivo de Estados Financieros")
    
    st.info("Modifica cualquier celda numérica de la tabla a continuación para simular escenarios y recalcular automáticamente todas las métricas, gráficos e interpretaciones en tiempo real.")
    
    edited_df = st.data_editor(
        df_active,
        use_container_width=True,
        num_rows="dynamic",
        key="data_editor_grid"
    )
    
    col_save, col_reset = st.columns([1, 4])
    with col_save:
        if st.button("💾 Aplicar Cambios Editados", type="primary"):
            st.session_state.financial_data = edited_df
            st.success("Cambios aplicados exitosamente. Las métricas y gráficos han sido actualizados.")
            st.rerun()
