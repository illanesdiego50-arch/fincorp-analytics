"""
Módulo de Interpretación Ejecutiva y Diagnóstico Automático
FinCorp Analytics — Evaluaciones y Dictámenes Corporativos

Genera:
1. Semáforo de Salud Financiera automatizado (Verde, Amarillo, Rojo)
2. Diagnósticos de Liquidez, Solvencia, Actividad, Rentabilidad y Flujo Libre
3. Dictamen narrativo para Comités y Directorios
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional


def evaluate_financial_health(ratios_df: pd.DataFrame, fcf_components: Dict[str, Any], year: str) -> List[Dict[str, Any]]:
    """
    Evalúa cada indicador clave para el año seleccionado y devuelve un semáforo
    (🟢 Verde / Óptimo, 🟡 Amarillo / Precaución, 🔴 Rojo / Crítico) con diagnóstico.
    """
    y = str(year)
    diagnostics = []

    def get_ratio_val(patterns: List[str]) -> Optional[float]:
        for pat in patterns:
            sub = ratios_df[ratios_df["Ratio"].str.contains(pat, case=False, regex=True)]
            if not sub.empty:
                val = sub.iloc[0].get(y, np.nan)
                if not pd.isna(val):
                    return float(val)
        return None

    # 1. Liquidez Corriente
    rc = get_ratio_val(["Liquidez Corriente", "Razón Corriente"])
    if rc is not None:
        if rc >= 1.5:
            status = "🟢 Óptimo"
            color = "green"
            msg = f"La liquidez corriente de {rc:.2f}x indica una adecuada cobertura de obligaciones a corto plazo sin tensiones operativas."
        elif rc >= 1.1:
            status = "🟡 Aceptable / Monitoreo"
            color = "orange"
            msg = f"Liquidez corriente de {rc:.2f}x. Cobertura justa de pasivos corrientes; se recomienda supervisar el calendario de vencimientos."
        else:
            status = "🔴 Riesgo de Liquidez"
            color = "red"
            msg = f"Liquidez corriente de {rc:.2f}x por debajo de la paridad (1.0x). Existe un déficit de capital de trabajo que puede comprometer pagos inmediatos."
            
        diagnostics.append({
            "Pilar": "Liquidez",
            "Indicador": "Liquidez Corriente",
            "Valor": f"{rc:.2f}x",
            "Estado": status,
            "color": color,
            "Color": color,
            "Interpretación": msg
        })

    # 2. Prueba Ácida
    pa = get_ratio_val(["Prueba Ácida", "Quick Ratio"])
    if pa is not None:
        if pa >= 1.0:
            status = "🟢 Saludable"
            color = "green"
            msg = f"Prueba ácida de {pa:.2f}x (≥ 1.0x). La empresa puede saldar la totalidad de sus pasivos de corto plazo sin depender de la venta de inventarios."
        elif pa >= 0.8:
            status = "🟡 Precaución"
            color = "orange"
            msg = f"Prueba ácida de {pa:.2f}x (< 1.0x). Existe una moderada dependencia de la rotación y liquidación del inventario para cubrir obligaciones inmediatas."
        else:
            status = "🔴 Alerta Crítica"
            color = "red"
            msg = f"Prueba ácida severa de {pa:.2f}x. Alta vulnerabilidad ante posibles retrasos en cobranzas o desaceleración en ventas de inventarios."
            
        diagnostics.append({
            "Pilar": "Liquidez Inmediata",
            "Indicador": "Prueba Ácida",
            "Valor": f"{pa:.2f}x",
            "Estado": status,
            "color": color,
            "Color": color,
            "Interpretación": msg
        })

    # 3. Nivel de Endeudamiento / Pasivo Total
    end = get_ratio_val(["Pasivo Total", "Razón de Endeudamiento"])
    if end is not None:
        if end <= 50.0:
            status = "🟢 Conservador"
            color = "green"
            msg = f"Nivel de pasivo total del {end:.1f}% sobre activos. Estructura de capital sólida con amplio margen para financiamiento adicional si se requiere."
        elif end <= 70.0:
            status = "🟡 Moderado"
            color = "orange"
            msg = f"Pasivo total del {end:.1f}%. Nivel habitual en industrias de capital intensivo, pero exige seguimiento a la carga financiera."
        else:
            status = "🔴 Alto Apalancamiento"
            color = "red"
            msg = f"Endeudamiento elevado del {end:.1f}%. La mayor parte de los activos está financiada por acreedores, elevando el riesgo financiero."
            
        diagnostics.append({
            "Pilar": "Solvencia",
            "Indicador": "Pasivo Total / Endeudamiento",
            "Valor": f"{end:.1f}%",
            "Estado": status,
            "color": color,
            "Color": color,
            "Interpretación": msg
        })

    # 4. Cobertura de Intereses
    cob = get_ratio_val(["Cobertura de Intereses"])
    if cob is not None:
        if cob >= 3.5:
            status = "🟢 Sólida"
            color = "green"
            msg = f"Cobertura de intereses de {cob:.2f}x. El EBIT cubre ampliamente el costo de la deuda, garantizando tranquilidad crediticia."
        elif cob >= 2.0:
            status = "🟡 Adecuada"
            color = "orange"
            msg = f"Cobertura de {cob:.2f}x. Capacidad suficiente de pago, pero vulnerable ante contracciones repentinas en ventas o alzas de tasas."
        else:
            status = "🔴 Tensión Financiera"
            color = "red"
            msg = f"Cobertura crítica de {cob:.2f}x (< 2.0x). El servicio de intereses absorbe una proporción riesgosa de la utilidad operativa."
            
        diagnostics.append({
            "Pilar": "Solvencia",
            "Indicador": "Cobertura de Intereses",
            "Valor": f"{cob:.2f}x",
            "Estado": status,
            "color": color,
            "Color": color,
            "Interpretación": msg
        })

    # 5. Ciclo de Conversión de Efectivo (CCE / CCC)
    cce = get_ratio_val(["Conversión del Efectivo", "CCC"])
    if cce is not None:
        if cce <= 60:
            status = "🟢 Ágil"
            color = "green"
            msg = f"Ciclo de conversión de efectivo ágil de {cce:.0f} días. La compañía recupera rápidamente la liquidez invertida en el ciclo comercial."
        elif cce <= 100:
            status = "🟡 Convencional"
            color = "orange"
            msg = f"Ciclo de efectivo de {cce:.0f} días. Rango estándar comercial; evaluar si es posible extender plazos con proveedores u optimizar inventario."
        else:
            status = "🔴 Ciclo Prolongado"
            color = "red"
            msg = f"Ciclo dilatado de {cce:.0f} días. Los fondos permanecen atrapados en la cadena de capital de trabajo por tiempo excesivo."
            
        diagnostics.append({
            "Pilar": "Actividad Operativa",
            "Indicador": "Ciclo de Conversión de Efectivo",
            "Valor": f"{cce:.0f} días",
            "Estado": status,
            "color": color,
            "Color": color,
            "Interpretación": msg
        })

    # 6. Rentabilidad Patrimonial (ROE)
    roe = get_ratio_val(["^ROE$", "ROE \\(Rentabilidad"])
    if roe is not None:
        if roe >= 15.0:
            status = "🟢 Excelente"
            color = "green"
            msg = f"Rentabilidad sobre el patrimonio del {roe:.1f}%. Generación de valor atractiva y superior al costo de oportunidad del capital contable."
        elif roe >= 8.0:
            status = "🟡 Moderada"
            color = "orange"
            msg = f"ROE del {roe:.1f}%. Rendimiento positivo aunque con potencial de mejora en margen operativo o rotación de activos."
        else:
            status = "🔴 Débil"
            color = "red"
            msg = f"ROE bajo del {roe:.1f}%. Retorno insuficiente para los accionistas frente a los riesgos asumidos."
            
        diagnostics.append({
            "Pilar": "Rentabilidad",
            "Indicador": "ROE (Retorno al Patrimonio)",
            "Valor": f"{roe:.1f}%",
            "Estado": status,
            "color": color,
            "Color": color,
            "Interpretación": msg
        })

    # 7. Flujo Libre FELF / FCF
    felf_val = float(fcf_components.get("felf_op1", fcf_components.get("fcf", {})).get(y, 0))
    if felf_val > 0:
        status = "🟢 Positivo"
        color = "green"
        msg = f"Flujo de Efectivo Libre positivo de ${felf_val:,.0f}. La operación autofinancia su reposición de activos (CAPEX) y genera excedentes."
    else:
        status = "🔴 Negativo / En Inversión"
        color = "red"
        msg = f"Flujo de Efectivo Libre deficitario (${felf_val:,.0f}). Los requerimientos de CAPEX y capital de trabajo absorbieron la totalidad del flujo operativo."
        
    diagnostics.append({
        "Pilar": "Flujo de Efectivo Libre",
        "Indicador": "FELF (Flujo Libre de la Firma)",
        "Valor": f"US$ {felf_val:,.0f}",
        "Estado": status,
        "color": color,
        "Color": color,
        "Interpretación": msg
    })

    return diagnostics


def generate_executive_summary(ratios_df: pd.DataFrame, fcf_components: Dict[str, Any], dupont_3_df: pd.DataFrame, years: List[str]) -> str:
    """
    Genera un informe narrativo ejecutivo consolidado analizando la trayectoria entre el año inicial y final.
    """
    y_ini = years[0]
    y_fin = years[-1]
    
    def get_val(pattern: str, yr: str, default: float = 0.0) -> float:
        sub = ratios_df[ratios_df["Ratio"].str.contains(pattern, case=False, regex=True)]
        if not sub.empty:
            v = sub.iloc[0].get(yr, default)
            return float(v) if not pd.isna(v) else default
        return default
        
    roe_fin = get_val("^ROE$|ROE \\(Rentabilidad", y_fin)
    roe_ini = get_val("^ROE$|ROE \\(Rentabilidad", y_ini)
    delta_roe = roe_fin - roe_ini
    
    roa_fin = get_val("^ROA$|ROA \\(Rentabilidad", y_fin)
    roa_ini = get_val("^ROA$|ROA \\(Rentabilidad", y_ini)
    
    rc_fin = get_val("Liquidez Corriente|Razón Corriente", y_fin, default=1.0)
    pa_fin = get_val("Prueba Ácida", y_fin, default=1.0)
    
    end_fin = get_val("Pasivo Total|Razón de Endeudamiento", y_fin, default=50.0)
    cob_fin = get_val("Cobertura de Intereses", y_fin, default=2.0)
    
    fcf_fin = float(fcf_components.get("felf_op1", fcf_components.get("fcf", {})).get(y_fin, 0))
    ebitda_fin = float(fcf_components["ebitda"].get(y_fin, 0))
    
    # Análisis de los inductores de DuPont
    m_neto_fin = float(dupont_3_df[dupont_3_df["Paso"].str.contains("Margen Neto")].iloc[0][y_fin])
    m_neto_ini = float(dupont_3_df[dupont_3_df["Paso"].str.contains("Margen Neto")].iloc[0][y_ini])
    
    rot_fin = float(dupont_3_df[dupont_3_df["Paso"].str.contains("Rotación de Activos")].iloc[0][y_fin])
    rot_ini = float(dupont_3_df[dupont_3_df["Paso"].str.contains("Rotación de Activos")].iloc[0][y_ini])
    
    lev_fin = float(dupont_3_df[dupont_3_df["Paso"].str.contains("Multiplicador|Apalancamiento")].iloc[0][y_fin])
    lev_ini = float(dupont_3_df[dupont_3_df["Paso"].str.contains("Multiplicador|Apalancamiento")].iloc[0][y_ini])
    
    drivers = []
    if m_neto_fin > m_neto_ini:
        drivers.append("expansión del Margen Neto")
    if rot_fin > rot_ini:
        drivers.append("mayor rotación y productividad de activos")
    if lev_fin > lev_ini:
        drivers.append("incremento del apalancamiento financiero")
    driver_text = ", ".join(drivers) if drivers else "estabilidad en sus parámetros operativos y patrimoniales"
    
    summary = f"""
### 📌 Dictamen Financiero Ejecutivo ({y_ini} - {y_fin})

1. **Rentabilidad y Retorno al Accionista (ROE/ROA con Promedios)**:
   - El ROE cerró en **{roe_fin:.2f}%** en {y_fin} (frente a {roe_ini:.2f}% en {y_ini}, variación neta de **{delta_roe:+.2f} p.p.**).
   - El ROA se posicionó en **{roa_fin:.2f}%**, registrando un spread financiero de **{(roe_fin - roa_fin):.2f} p.p.** a favor del capital accionario.
   - De acuerdo con la descomposición DuPont, los inductores clave fueron: **{driver_text}**.

2. **Posición de Liquidez y Capital de Trabajo**:
   - Liquidez Corriente: **{rc_fin:.2f}x** y Prueba Ácida: **{pa_fin:.2f}x**.
   - {"🟢 La empresa cuenta con holgura suficiente para hacer frente a obligaciones corrientes sin requerir la realización forzada de inventarios." if pa_fin >= 1.0 else "🟡 Se observa cierta tensión en liquidez inmediata (Prueba Ácida < 1.0x), requiriéndose vigilar plazos de cobranzas a clientes (PPC)."}

3. **Estructura de Solvencia y Cobertura**:
   - Pasivo Total sobre Activos: **{end_fin:.1f}%**.
   - Cobertura de Intereses: **{cob_fin:.2f}x**, {"indicando una capacidad robusta de pago frente al costo del endeudamiento." if cob_fin >= 3.0 else "requiriendo prudencia para evitar sobrecargas de servicio de deuda."}

4. **Flujos Libres (FELF) y EBITDA**:
   - Generación de EBITDA corporativo: **US$ {ebitda_fin:,.0f}** en {y_fin}.
   - Flujo de Efectivo Libre de la Firma (FELF): **US$ {fcf_fin:,.0f}**, {"confirmando una operación autosuficiente capaz de remunerar deuda y accionistas tras cubrir CAPEX y capital de trabajo." if fcf_fin > 0 else "evidenciando absorción transitoria de caja por proyectos de inversión de capital o acumulación de existencias."}
"""
    return summary
