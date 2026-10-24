"""
Módulo de Visualizaciones Interactivas con Plotly
FinCorp Analytics — Gráficos Ejecutivos Corporativos

Incluye:
- Cascadas Waterfall para FELF y FELP (Opciones 1 y 2)
- Gráfico de Atribución y Causa Raíz DuPont (Drivers vs Detractores de ROE/ROA)
- Evolución histórica de Rentabilidad (ROE vs ROA con spread)
- Márgenes de Utilidad (Bruto, EBITDA, Operativo, Neto)
- Liquidez y Prueba Ácida con líneas de referencia estándar
- Ciclo de Conversión de Efectivo (PPC, EPI, PPP, CCE)
- Descomposición DuPont de 3 Pasos
- Radar de Desempeño Financiero Integral
"""

import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import numpy as np
from typing import Dict, Any, List, Optional

# Paleta cromática corporativa moderna (Executive Navy & Slate)
COLORS = {
    "primary": "#1E3A8A",       # Deep Navy
    "secondary": "#3B82F6",     # Corporate Blue
    "accent": "#0D9488",        # Teal
    "success": "#10B981",       # Emerald Green
    "warning": "#F59E0B",       # Amber
    "danger": "#EF4444",        # Crimson Red
    "neutral_dark": "#1F2937",  # Dark Charcoal
    "neutral_light": "#F3F4F6", # Light Gray
    "purple": "#8B5CF6",
    "pink": "#EC4899"
}

THEME_LAYOUT = dict(
    font=dict(family="Inter, Roboto, sans-serif", size=13, color="#1F2937"),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(248,250,252,0.6)",
    hoverlabel=dict(bgcolor="#1E293B", font_size=12, font_color="#FFFFFF"),
    margin=dict(l=40, r=40, t=50, b=40),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
)


def create_felf_waterfall(felf_components: Dict[str, Any], selected_year: str, method: str = "Opción 1") -> go.Figure:
    """
    Gráfico de Cascada Waterfall para el Flujo de Efectivo Libre de la Firma (FELF).
    - Opción 1: NOPAT -> (+) D&A -> (-) CAPEX -> (-/+) MNAPO -> (=) FELF
    - Opción 2: FEAO -> (+) GF * (1 - t) -> (-) CAPEX -> (=) FELF
    """
    y = selected_year
    
    if "Opción 2" in method or "2" in method:
        feao = float(felf_components["feao"].get(y, 0))
        tasa = float(felf_components["tasa_tax"].get(y, 0.25))
        escudo = float(felf_components["gf"].get(y, 0)) * (1.0 - tasa)
        capex = -abs(float(felf_components["capex"].get(y, 0)))
        felf_total = float(felf_components["felf_op2"].get(y, 0))
        
        x_labels = ["FEAO", "(+) GF × (1 - τ)", "(-) CAPEX", "(=) FELF (Opción 2)"]
        y_values = [feao, escudo, capex, felf_total]
        measures = ["relative", "relative", "relative", "total"]
        title_text = f"<b>FELF (Firma) — Opción 2 (Base FEAO): Periodo {selected_year}</b>"
    else:
        nopat = float(felf_components["nopat"].get(y, 0))
        depr = float(felf_components["pngme"].get(y, 0))
        capex = -abs(float(felf_components["capex"].get(y, 0)))
        mnapo = -float(felf_components["mnapo"].get(y, 0))
        felf_total = float(felf_components["felf_op1"].get(y, 0))
        
        x_labels = ["NOPAT (RAII - IR)", "(+) PNGME (D&A)", "(-) CAPEX", "(-/+) MNAPO (ΔKTNO)", "(=) FELF (Opción 1)"]
        y_values = [nopat, depr, capex, mnapo, felf_total]
        measures = ["relative", "relative", "relative", "relative", "total"]
        title_text = f"<b>FELF (Firma) — Opción 1 (Base RAII): Periodo {selected_year}</b>"
        
    fig = go.Figure(go.Waterfall(
        name="FELF",
        orientation="v",
        measure=measures,
        x=x_labels,
        textposition="outside",
        text=[f"${val:,.0f}" for val in y_values],
        y=y_values,
        connector={"line": {"color": "#94A3B8", "width": 1.5, "dash": "dot"}},
        decreasing={"marker": {"color": COLORS["danger"]}},
        increasing={"marker": {"color": COLORS["success"]}},
        totals={"marker": {"color": COLORS["primary"]}}
    ))
    
    fig.update_layout(
        title=title_text,
        yaxis_title="Monto ($)",
        **THEME_LAYOUT
    )
    fig.update_yaxes(gridcolor="#E2E8F0")
    return fig


def create_felp_waterfall(felf_components: Dict[str, Any], selected_year: str, method: str = "Opción 1") -> go.Figure:
    """
    Gráfico de Cascada Waterfall para el Flujo de Efectivo Libre del Patrimonio (FELP).
    - Opción 1: RN -> (+) PNGME -> (-) CAPEX -> (-/+) MNAPO -> (+/-) ΔD -> (=) FELP
    - Opción 2: FEAO -> (-) CAPEX -> (+/-) ΔD -> (=) FELP
    """
    y = selected_year
    delta_deuda = float(felf_components["delta_deuda"].get(y, 0))
    capex = -abs(float(felf_components["capex"].get(y, 0)))
    
    if "Opción 2" in method or "2" in method:
        feao = float(felf_components["feao"].get(y, 0))
        felp_total = float(felf_components["felp_op2"].get(y, 0))
        
        x_labels = ["FEAO", "(-) CAPEX", "(+/-) Δ Deuda", "(=) FELP (Opción 2)"]
        y_values = [feao, capex, delta_deuda, felp_total]
        measures = ["relative", "relative", "relative", "total"]
        title_text = f"<b>FELP (Patrimonio) — Opción 2 (Base FEAO): Periodo {selected_year}</b>"
    else:
        rn = float(felf_components["rn"].get(y, 0))
        pngme = float(felf_components["pngme"].get(y, 0))
        mnapo = -float(felf_components["mnapo"].get(y, 0))
        felp_total = float(felf_components["felp_op1"].get(y, 0))
        
        x_labels = ["Resultado Neto", "(+) PNGME (D&A)", "(-) CAPEX", "(-/+) MNAPO (ΔKTNO)", "(+/-) Δ Deuda", "(=) FELP (Opción 1)"]
        y_values = [rn, pngme, capex, mnapo, delta_deuda, felp_total]
        measures = ["relative", "relative", "relative", "relative", "relative", "total"]
        title_text = f"<b>FELP (Patrimonio) — Opción 1 (Base RN): Periodo {selected_year}</b>"
        
    fig = go.Figure(go.Waterfall(
        name="FELP",
        orientation="v",
        measure=measures,
        x=x_labels,
        textposition="outside",
        text=[f"${val:,.0f}" for val in y_values],
        y=y_values,
        connector={"line": {"color": "#94A3B8", "width": 1.5, "dash": "dot"}},
        decreasing={"marker": {"color": COLORS["danger"]}},
        increasing={"marker": {"color": COLORS["success"]}},
        totals={"marker": {"color": COLORS["purple"]}}
    ))
    
    fig.update_layout(
        title=title_text,
        yaxis_title="Monto ($)",
        **THEME_LAYOUT
    )
    fig.update_yaxes(gridcolor="#E2E8F0")
    return fig


# Compatibilidad hacia atrás
def create_fcf_waterfall(fcf_components: Dict[str, Any], selected_year: str) -> go.Figure:
    return create_felf_waterfall(fcf_components, selected_year, method="Opción 1")


def create_dupont_attribution_chart(attribution_df: pd.DataFrame, target_year: str, base_year: str) -> go.Figure:
    """
    Gráfico de barras horizontales divergentes para el análisis de causa raíz DuPont.
    Muestra la contribución de cada cuenta contable en puntos porcentuales (p.p.) al cambio en el ROE.
    """
    # Ordenar por impacto
    df_sorted = attribution_df.sort_values(by="Impacto Estimado ROE (p.p.)", ascending=True)
    
    colors = [COLORS["success"] if x >= 0 else COLORS["danger"] for x in df_sorted["Impacto Estimado ROE (p.p.)"]]
    
    fig = go.Figure(go.Bar(
        x=df_sorted["Impacto Estimado ROE (p.p.)"],
        y=df_sorted["Cuenta"],
        orientation="h",
        marker_color=colors,
        text=[f"{v:+.2f} p.p." for v in df_sorted["Impacto Estimado ROE (p.p.)"]],
        textposition="outside"
    ))
    
    fig.add_vline(x=0, line_width=1.5, line_color="#475569")
    
    fig.update_layout(
        title=f"<b>Atribución Causa Raíz al ROE: Periodo {target_year} vs. {base_year}</b>",
        xaxis_title="Impacto en el ROE (puntos porcentuales)",
        yaxis=dict(autorange="reversed"),
        **THEME_LAYOUT
    )
    fig.update_xaxes(gridcolor="#E2E8F0")
    return fig


def create_roe_roa_trend_chart(ratios_df: pd.DataFrame, years: List[str]) -> go.Figure:
    """
    Gráfico de barras agrupadas y líneas superpuestas para la evolución de ROE vs ROA.
    """
    roe_row = ratios_df[ratios_df["Ratio"] == "ROE"].iloc[0]
    roa_row = ratios_df[ratios_df["Ratio"] == "ROA"].iloc[0]
    
    roe_vals = [roe_row[y] for y in years]
    roa_vals = [roa_row[y] for y in years]
    
    fig = go.Figure()
    
    # Barra ROA
    fig.add_trace(go.Bar(
        x=years,
        y=roa_vals,
        name="ROA (Rentabilidad Activos)",
        marker_color=COLORS["secondary"],
        text=[f"{v:.1f}%" for v in roa_vals],
        textposition="outside"
    ))
    
    # Barra ROE
    fig.add_trace(go.Bar(
        x=years,
        y=roe_vals,
        name="ROE (Rentabilidad Patrimonio)",
        marker_color=COLORS["primary"],
        text=[f"{v:.1f}%" for v in roe_vals],
        textposition="outside"
    ))
    
    # Spread ROE - ROA (Línea de Efecto Apalancamiento)
    spread = [roe - roa for roe, roa in zip(roe_vals, roa_vals)]
    fig.add_trace(go.Scatter(
        x=years,
        y=spread,
        name="Spread Financiero (ROE - ROA)",
        mode="lines+markers",
        line=dict(color=COLORS["accent"], width=3, dash="dash"),
        marker=dict(size=8)
    ))
    
    fig.update_layout(
        title="<b>Evolución Comparativa de Rentabilidad: ROE vs. ROA</b>",
        yaxis_title="Porcentaje (%)",
        barmode="group",
        **THEME_LAYOUT
    )
    fig.update_yaxes(gridcolor="#E2E8F0")
    return fig


def create_margins_trend_chart(ratios_df: pd.DataFrame, years: List[str]) -> go.Figure:
    """
    Evolución temporal de los 4 márgenes clave: Bruto, EBITDA, Operativo y Neto.
    """
    mb = ratios_df[ratios_df["Ratio"] == "Margen Bruto"].iloc[0][years].values
    me = ratios_df[ratios_df["Ratio"] == "Margen EBITDA"].iloc[0][years].values
    mo = ratios_df[ratios_df["Ratio"] == "Margen Operativo"].iloc[0][years].values
    mn = ratios_df[ratios_df["Ratio"] == "Margen Neto"].iloc[0][years].values
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=years, y=mb, mode="lines+markers", name="Margen Bruto",
        line=dict(color=COLORS["secondary"], width=3),
        marker=dict(size=7)
    ))
    fig.add_trace(go.Scatter(
        x=years, y=me, mode="lines+markers", name="Margen EBITDA",
        line=dict(color=COLORS["purple"], width=2.5, dash="dash"),
        marker=dict(size=6)
    ))
    fig.add_trace(go.Scatter(
        x=years, y=mo, mode="lines+markers", name="Margen Operativo",
        line=dict(color=COLORS["accent"], width=3),
        marker=dict(size=7)
    ))
    fig.add_trace(go.Scatter(
        x=years, y=mn, mode="lines+markers", name="Margen Neto",
        line=dict(color=COLORS["primary"], width=3),
        marker=dict(size=7)
    ))
    
    fig.update_layout(
        title="<b>Estructura y Evolución de Márgenes de Utilidad</b>",
        yaxis_title="Margen (%)",
        **THEME_LAYOUT
    )
    fig.update_yaxes(gridcolor="#E2E8F0")
    return fig


def create_liquidity_chart(ratios_df: pd.DataFrame, years: List[str]) -> go.Figure:
    """
    Gráfico de barras para Liquidez Corriente, Prueba Ácida y Razón Rápida con líneas de referencia.
    """
    rc = ratios_df[ratios_df["Ratio"] == "Liquidez Corriente"].iloc[0][years].values
    pa = ratios_df[ratios_df["Ratio"] == "Prueba Ácida"].iloc[0][years].values
    rr = ratios_df[ratios_df["Ratio"] == "Razón Rápida"].iloc[0][years].values
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=years, y=rc, name="Liquidez Corriente",
        marker_color=COLORS["secondary"],
        text=[f"{v:.2f}x" for v in rc], textposition="outside"
    ))
    fig.add_trace(go.Bar(
        x=years, y=pa, name="Prueba Ácida",
        marker_color=COLORS["accent"],
        text=[f"{v:.2f}x" for v in pa], textposition="outside"
    ))
    fig.add_trace(go.Bar(
        x=years, y=rr, name="Razón Rápida",
        marker_color="#0284C7",
        text=[f"{v:.2f}x" for v in rr], textposition="outside"
    ))
    
    # Línea umbral prudencial para Prueba Ácida (1.0x)
    fig.add_hline(y=1.0, line_dash="dot", line_color=COLORS["danger"],
                  annotation_text="Umbral Mínimo Prueba Ácida (1.0x)", annotation_position="top left")
    # Línea umbral prudencial para Razón Corriente (1.5x)
    fig.add_hline(y=1.5, line_dash="dash", line_color=COLORS["warning"],
                  annotation_text="Umbral Óptimo Corriente (1.5x)", annotation_position="top right")
    
    fig.update_layout(
        title="<b>Indicadores de Liquidez Inmediata vs. Umbrales Financieros</b>",
        yaxis_title="Ratio (veces)",
        barmode="group",
        **THEME_LAYOUT
    )
    fig.update_yaxes(gridcolor="#E2E8F0")
    return fig


def create_cash_conversion_cycle_chart(ratios_df: pd.DataFrame, years: List[str]) -> go.Figure:
    """
    Visualización del Ciclo Operativo y Ciclo de Conversión de Efectivo:
    CCE = PPC (Cobro) + EPI (Inventarios) - PPP (Pago).
    """
    def find_val(pattern):
        sub = ratios_df[ratios_df["Ratio"].str.contains(pattern, case=False, regex=True)]
        if not sub.empty:
            return sub.iloc[0][years].values
        return np.zeros(len(years))

    ppc = find_val("PPC|Cobro")
    epi = find_val("EPI|Inventario")
    ppp = find_val("PPP|Pago")
    cce = find_val("Conversión del Efectivo|CCC")
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(x=years, y=epi, name="(+) Edad Promedio Inventario (EPI)", marker_color="#64748B"))
    fig.add_trace(go.Bar(x=years, y=ppc, name="(+) Periodo Promedio Cobro (PPC)", marker_color="#0EA5E9"))
    fig.add_trace(go.Bar(x=years, y=-ppp, name="(-) Periodo Promedio Pago (PPP)", marker_color="#F97316"))
    
    fig.add_trace(go.Scatter(
        x=years, y=cce, name="Ciclo de Conversión de Efectivo (CCE)",
        mode="lines+markers+text",
        text=[f"{v:.0f} d" for v in cce],
        textposition="top center",
        line=dict(color=COLORS["primary"], width=3.5),
        marker=dict(size=9, color=COLORS["primary"])
    ))
    
    fig.update_layout(
        title="<b>Ciclo de Conversión de Efectivo (CCE = PPC + EPI - PPP)</b>",
        yaxis_title="Días",
        barmode="relative",
        **THEME_LAYOUT
    )
    fig.update_yaxes(gridcolor="#E2E8F0")
    return fig


def create_dupont_breakdown_chart(dupont_3_df: pd.DataFrame, years: List[str]) -> go.Figure:
    """
    Visualiza la contribución de los 3 pilares de DuPont al ROE a lo largo del tiempo.
    """
    m_neto = dupont_3_df[dupont_3_df["Paso"].str.contains("Margen Neto")].iloc[0][years].values
    rot_act = dupont_3_df[dupont_3_df["Paso"].str.contains("Rotación de Activos")].iloc[0][years].values
    mult_cap = dupont_3_df[dupont_3_df["Paso"].str.contains("Multiplicador")].iloc[0][years].values
    roe = dupont_3_df[dupont_3_df["Paso"].str.contains("ROE DuPont")].iloc[0][years].values
    
    base_m = m_neto[0] if m_neto[0] != 0 else 1
    base_r = rot_act[0] if rot_act[0] != 0 else 1
    base_c = mult_cap[0] if mult_cap[0] != 0 else 1
    base_roe = roe[0] if roe[0] != 0 else 1
    
    idx_m = (m_neto / base_m) * 100
    idx_r = (rot_act / base_r) * 100
    idx_c = (mult_cap / base_c) * 100
    idx_roe = (roe / base_roe) * 100
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(x=years, y=idx_m, mode="lines+markers", name="1. Margen Neto (Base 100)", line=dict(color=COLORS["success"], width=2.5)))
    fig.add_trace(go.Scatter(x=years, y=idx_r, mode="lines+markers", name="2. Rotación de Activos (Base 100)", line=dict(color=COLORS["secondary"], width=2.5)))
    fig.add_trace(go.Scatter(x=years, y=idx_c, mode="lines+markers", name="3. Apalancamiento Financiero (Base 100)", line=dict(color=COLORS["warning"], width=2.5)))
    fig.add_trace(go.Scatter(x=years, y=idx_roe, mode="lines+markers", name="ROE DuPont (Base 100)", line=dict(color=COLORS["primary"], width=4, dash="solid")))
    
    fig.update_layout(
        title="<b>Dinámica de Impulsores DuPont 3 Pasos (Evolución Indexada Base 100)</b>",
        yaxis_title="Índice Base 100",
        **THEME_LAYOUT
    )
    fig.update_yaxes(gridcolor="#E2E8F0")
    return fig


def create_financial_radar(ratios_df: pd.DataFrame, selected_year: str) -> go.Figure:
    """
    Gráfico Radar de diagnóstico general en escala relativa (0 a 100 pts)
    evaluando: Liquidez, Solvencia, Rotación, Margen Operativo y Retorno Patrimonial.
    """
    y = selected_year
    
    def get_single_ratio(name):
        sub = ratios_df[ratios_df["Ratio"].str.contains(name, case=False)]
        if not sub.empty:
            val = sub.iloc[0][y]
            return float(val) if not np.isnan(val) else 0.0
        return 0.0
        
    rc = get_single_ratio("Liquidez Corriente|Razón Corriente")
    cob_int = get_single_ratio("Cobertura de Intereses")
    rot_at = get_single_ratio("Rotación de Activos Totales")
    margen_op = get_single_ratio("Margen Operativo")
    roe = get_single_ratio("^ROE$|ROE \\(Rentabilidad")
    
    # Normalización a escala 0 - 100
    score_liq = min(max((rc / 2.0) * 100, 10), 100)
    score_solv = min(max((cob_int / 6.0) * 100, 10), 100)
    score_eficiencia = min(max((rot_at / 1.5) * 100, 10), 100)
    score_margen = min(max((margen_op / 20.0) * 100, 10), 100)
    score_roe = min(max((roe / 25.0) * 100, 10), 100)
    
    categories = [
        "Liquidez (Cte)",
        "Solvencia (Cob. Int)",
        "Eficiencia Activos",
        "Margen Operativo",
        "Rentabilidad (ROE)"
    ]
    scores = [score_liq, score_solv, score_eficiencia, score_margen, score_roe]
    categories.append(categories[0])
    scores.append(scores[0])
    
    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=scores,
        theta=categories,
        fill="toself",
        name=f"Año {selected_year}",
        fillcolor="rgba(30, 58, 138, 0.25)",
        line=dict(color=COLORS["primary"], width=2.5)
    ))
    
    fig.update_layout(
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 100], gridcolor="#E2E8F0")
        ),
        title=f"<b>Radar de Salud Financiera — Periodo {selected_year}</b>",
        **THEME_LAYOUT
    )
    return fig
