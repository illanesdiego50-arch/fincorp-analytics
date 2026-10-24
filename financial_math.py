"""
Módulo de Cálculo Matemático Financiero y Normativa Contable Boliviana (NC 3 / NC 6)
FinCorp Analytics — Suite Integral de Análisis Financiero

Implementa:
1. Totales y Subtotales de Balance y Estado de Resultados
2. Promedios entre inicio y fin de gestión para cuentas de Situación Financiera
3. Catálogo completo de Razones Financieras (Actividad, Liquidez, Rentabilidad, Solvencia)
4. Descomposición Jerárquica DuPont y Motor de Atribución de Cuentas de Impacto (Root Cause Analysis)
5. Flujos de Efectivo Libres FELF y FELP (Opciones 1 y 2 con conciliación)
6. Análisis Estructural Vertical y Horizontal con ajuste opcional por inflación / UFV (NC 3 / NC 6)
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Any, Tuple, Optional
from modules.data_loader import get_years_from_df, get_account_series

# Valores UFV de cierre de gestión referenciales de Bolivia (Banco Central de Bolivia)
DEFAULT_UFV_RATES = {
    "2021": 2.37376,
    "2022": 2.40858,
    "2023": 2.47895,
    "2024": 2.54520,
    "2025": 2.61000
}


def get_balance_totals(df: pd.DataFrame) -> Dict[str, pd.Series]:
    """
    Calcula los totales y subtotales clave del Balance General para todos los años disponibles.
    """
    years = get_years_from_df(df)
    
    # Activo Corriente
    ac_mask = (df["Estado"] == "Balance General") & (df["Categoria"] == "Activo Corriente")
    ac_series = df[ac_mask][years].sum(axis=0).astype(float)
    
    # Activo No Corriente
    anc_mask = (df["Estado"] == "Balance General") & (df["Categoria"] == "Activo No Corriente")
    anc_series = df[anc_mask][years].sum(axis=0).astype(float)
    
    # Total Activos
    total_activos = ac_series + anc_series
    
    # Pasivo Corriente
    pc_mask = (df["Estado"] == "Balance General") & (df["Categoria"] == "Pasivo Corriente")
    pc_series = df[pc_mask][years].sum(axis=0).astype(float)
    
    # Pasivo No Corriente
    pnc_mask = (df["Estado"] == "Balance General") & (df["Categoria"] == "Pasivo No Corriente")
    pnc_series = df[pnc_mask][years].sum(axis=0).astype(float)
    
    # Total Pasivos
    total_pasivos = pc_series + pnc_series
    
    # Patrimonio
    pat_mask = (df["Estado"] == "Balance General") & (df["Categoria"] == "Patrimonio")
    pat_series = df[pat_mask][years].sum(axis=0).astype(float)
    
    total_pasivo_patrimonio = total_pasivos + pat_series
    
    # Deuda Financiera (Corto Plazo + Largo Plazo)
    deuda_cp = get_account_series(df, "Deuda Financiera Corto Plazo")
    deuda_lp = get_account_series(df, "Deuda Financiera Largo Plazo")
    deuda_total = deuda_cp + deuda_lp
    
    # Activos Fijos Netos (PPE)
    activos_fijos = get_account_series(df, "Propiedad Planta y Equipo")
    
    return {
        "Activo Corriente": ac_series,
        "Activo No Corriente": anc_series,
        "Activos": total_activos,
        "Pasivo Corriente": pc_series,
        "Pasivo No Corriente": pnc_series,
        "Pasivos": total_pasivos,
        "Patrimonio": pat_series,
        "Pasivo y Patrimonio": total_pasivo_patrimonio,
        "Deuda Corto Plazo": deuda_cp,
        "Deuda Largo Plazo": deuda_lp,
        "Deuda Total": deuda_total,
        "Activos Fijos": activos_fijos
    }


def get_series_average(series: pd.Series, years: List[str]) -> pd.Series:
    """
    Calcula el promedio entre el inicio y final de gestión para una serie de balance.
    Para el primer año disponible (año 0), se utiliza el valor de cierre del año base.
    Para los años subsecuentes (año i >= 1), se calcula: (Valor_i + Valor_{i-1}) / 2.
    """
    avg_series = pd.Series(index=years, dtype=float)
    for i, y in enumerate(years):
        if i == 0:
            avg_series[y] = float(series[y])
        else:
            y_prev = years[i - 1]
            avg_series[y] = (float(series[y]) + float(series[y_prev])) / 2.0
    return avg_series


def get_balance_averages(df: pd.DataFrame) -> Dict[str, pd.Series]:
    """
    Genera series de promedios para todas las cuentas y agregados del Balance General.
    """
    years = get_years_from_df(df)
    bal = get_balance_totals(df)
    
    cxc = get_account_series(df, "Cuentas por Cobrar")
    inv = get_account_series(df, "Inventarios")
    cxp = get_account_series(df, "Cuentas por Pagar")
    ctn = bal["Activo Corriente"] - bal["Pasivo Corriente"]
    
    return {
        "Activos": get_series_average(bal["Activos"], years),
        "Patrimonio": get_series_average(bal["Patrimonio"], years),
        "Activo Corriente": get_series_average(bal["Activo Corriente"], years),
        "Pasivo Corriente": get_series_average(bal["Pasivo Corriente"], years),
        "Activos Fijos": get_series_average(bal["Activos Fijos"], years),
        "Cuentas por Cobrar": get_series_average(cxc, years),
        "Inventarios": get_series_average(inv, years),
        "Cuentas por Pagar": get_series_average(cxp, years),
        "Capital de Trabajo Neto": get_series_average(ctn, years),
        "Deuda Total": get_series_average(bal["Deuda Total"], years)
    }


def get_income_statement_totals(df: pd.DataFrame) -> Dict[str, pd.Series]:
    """
    Calcula los subtotales clave del Estado de Resultados:
    Ingresos Operativos, Costo de lo Vendido, Resultado Bruto, Gastos Operativos,
    EBITDA, Resultado Operativo (EBIT/RAII), EBT, Impuestos, Resultado Neto y NOPAT.
    """
    years = get_years_from_df(df)
    
    ventas = get_account_series(df, "Ingresos por Ventas")
    costo_ventas = get_account_series(df, "Costo de Ventas")
    utilidad_bruta = ventas - costo_ventas
    
    sga = get_account_series(df, "Gastos de Administracion y Ventas")
    depr_amort = get_account_series(df, "Depreciacion y Amortizacion")
    gastos_operativos_totales = sga + depr_amort
    
    ebit = utilidad_bruta - gastos_operativos_totales
    ebitda = ebit + depr_amort
    
    gastos_financieros = get_account_series(df, "Gastos Financieros")
    otros_ingresos = get_account_series(df, "Otros Ingresos")
    
    ebt = ebit - gastos_financieros + otros_ingresos
    impuestos = get_account_series(df, "Impuesto sobre la Renta")
    
    utilidad_neta = ebt - impuestos
    
    # Tasa efectiva calculada o configurada
    tasa_param = get_account_series(df, "Tasa Impositiva")
    tasa_efectiva = pd.Series(index=years, dtype=float)
    for y in years:
        if tasa_param.get(y, 0) > 0:
            tasa_efectiva[y] = float(tasa_param[y]) / 100.0 if float(tasa_param[y]) > 1.0 else float(tasa_param[y])
        elif ebt.get(y, 0) > 0:
            tasa_efectiva[y] = min(max(float(impuestos.get(y, 0)) / float(ebt.get(y, 1)), 0.0), 0.40)
        else:
            tasa_efectiva[y] = 0.25  # 25% IUE estándar boliviano
            
    nopat = ebit * (1.0 - tasa_efectiva)
    
    return {
        "Ventas": ventas,
        "Ingresos Operativos": ventas,
        "Costo de Ventas": costo_ventas,
        "Costo de lo Vendido": costo_ventas,
        "Utilidad Bruta": utilidad_bruta,
        "Resultado Bruto": utilidad_bruta,
        "Gastos Operativos": sga,
        "Gastos Operativos Totales": gastos_operativos_totales,
        "Depreciacion y Amortizacion": depr_amort,
        "EBITDA": ebitda,
        "EBIT": ebit,
        "Resultado Operativo": ebit,
        "RAII": ebit,
        "Gastos Financieros": gastos_financieros,
        "Otros Ingresos": otros_ingresos,
        "EBT": ebt,
        "Impuestos": impuestos,
        "Utilidad Neta": utilidad_neta,
        "Resultado Neto": utilidad_neta,
        "Tasa Efectiva": tasa_efectiva,
        "NOPAT": nopat
    }


def safe_div(num: Any, denom: Any, multiplier: float = 1.0) -> pd.Series:
    """División segura entre Series o escalares que retorna pd.Series evitando ZeroDivisionError."""
    if isinstance(num, pd.Series):
        idx = num.index
    elif isinstance(denom, pd.Series):
        idx = denom.index
    else:
        idx = None

    n = num.values if isinstance(num, pd.Series) else num
    d = denom.values if isinstance(denom, pd.Series) else denom
    
    with np.errstate(divide='ignore', invalid='ignore'):
        denom_safe = np.where(d != 0, d, np.nan)
        res = (n / denom_safe) * multiplier
    return pd.Series(res, index=idx, dtype=float)


def calculate_financial_ratios(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula el catálogo completo de razones financieras requeridas, asegurando
    el uso de PROMEDIOS (inicio y fin de gestión) en las cuentas de balance aplicables:
    - 12 Razones de Actividad
    - 8 Razones de Liquidez
    - 8 Razones de Rentabilidad
    - 12 Razones de Solvencia
    """
    years = get_years_from_df(df)
    bal = get_balance_totals(df)
    bal_avg = get_balance_averages(df)
    inc = get_income_statement_totals(df)
    
    # Cuentas de balance a fecha de corte (cierre)
    efectivo = get_account_series(df, "Efectivo y Equivalentes")
    icp = get_account_series(df, "Inversiones a Corto Plazo")
    disponibilidades = efectivo + icp
    cxc = get_account_series(df, "Cuentas por Cobrar")
    inv = get_account_series(df, "Inventarios")
    cxp = get_account_series(df, "Cuentas por Pagar")
    capex = get_account_series(df, "CAPEX")
    if capex.sum() == 0:
        capex = get_account_series(df, "Inversiones en Activo Fijo")
        
    dividendos = get_account_series(df, "Dividendos")
    feao_data = get_account_series(df, "FEAO")
    if feao_data.sum() == 0:
        feao_data = get_account_series(df, "Flujo de Efectivo Actividades Operativas")
        
    # Compras del periodo: si existe en datos se toma, sino se deduce: Costo + Inv_fin - Inv_ini
    compras_data = get_account_series(df, "Compras")
    compras = pd.Series(index=years, dtype=float)
    for i, y in enumerate(years):
        if compras_data.get(y, 0) > 0:
            compras[y] = float(compras_data[y])
        elif i == 0:
            compras[y] = float(inc["Costo de lo Vendido"][y])
        else:
            y_prev = years[i - 1]
            compras[y] = float(inc["Costo de lo Vendido"][y]) + float(inv[y]) - float(inv[y_prev])
            
    # Flujo de Efectivo de Actividades de Operación (FEAO)
    feao = pd.Series(index=years, dtype=float)
    for i, y in enumerate(years):
        if feao_data.get(y, 0) != 0:
            feao[y] = float(feao_data[y])
        else:
            ktno_y = (cxc[y] + inv[y]) - cxp[y]
            ktno_prev = (cxc[years[i-1]] + inv[years[i-1]]) - cxp[years[i-1]] if i > 0 else ktno_y
            feao[y] = float(inc["Resultado Neto"][y]) + float(inc["Depreciacion y Amortizacion"][y]) - (ktno_y - ktno_prev)

    ratios = []

    # =========================================================
    # 1. RAZONES DE ACTIVIDAD (12 Ratios con Promedios de Balance)
    # =========================================================
    # 1.1 Rotación de Cuentas por Cobrar = Ingresos Operativos / Promedio Cuentas por Cobrar
    rot_cxc = safe_div(inc["Ingresos Operativos"], bal_avg["Cuentas por Cobrar"])
    ratios.append({"Categoría": "Actividad", "Ratio": "Rotación de Cuentas por Cobrar", "Fórmula": "Ingresos Operativos / Promedio CxC", "Unidad": "veces", **rot_cxc.to_dict()})
    
    # 1.2 Periodo Promedio de Cobro (PPC) = 365 / Rotación Cuentas por Cobrar
    ppc = safe_div(365.0, rot_cxc)
    ratios.append({"Categoría": "Actividad", "Ratio": "Periodo Promedio de Cobro (PPC)", "Fórmula": "365 / Rotación CxC", "Unidad": "días", **ppc.to_dict()})
    
    # 1.3 Rotación de Inventarios = Costo de lo Vendido / Promedio Inventarios
    rot_inv = safe_div(inc["Costo de lo Vendido"], bal_avg["Inventarios"])
    ratios.append({"Categoría": "Actividad", "Ratio": "Rotación de Inventarios", "Fórmula": "Costo de lo Vendido / Promedio Inventarios", "Unidad": "veces", **rot_inv.to_dict()})
    
    # 1.4 Edad Promedio de Inventario (EPI) = 365 / Rotación Inventarios
    epi = safe_div(365.0, rot_inv)
    ratios.append({"Categoría": "Actividad", "Ratio": "Edad Promedio de Inventario (EPI)", "Fórmula": "365 / Rotación Inventarios", "Unidad": "días", **epi.to_dict()})
    
    # 1.5 Rotación Cuentas por Pagar = Compras / Promedio Cuentas por Pagar
    rot_cxp = safe_div(compras, bal_avg["Cuentas por Pagar"])
    ratios.append({"Categoría": "Actividad", "Ratio": "Rotación Cuentas por Pagar", "Fórmula": "Compras / Promedio CxP", "Unidad": "veces", **rot_cxp.to_dict()})
    
    # 1.6 Periodo Promedio de Pago (PPP) = 365 / Rotación Cuentas por Pagar
    ppp = safe_div(365.0, rot_cxp)
    ratios.append({"Categoría": "Actividad", "Ratio": "Periodo Promedio de Pago (PPP)", "Fórmula": "365 / Rotación CxP", "Unidad": "días", **ppp.to_dict()})
    
    # 1.7 Rotación del Capital de Trabajo = Ingresos Operativos / Promedio Capital de Trabajo Neto
    rot_ctn = safe_div(inc["Ingresos Operativos"], bal_avg["Capital de Trabajo Neto"])
    ratios.append({"Categoría": "Actividad", "Ratio": "Rotación del Capital de Trabajo", "Fórmula": "Ingresos Operativos / Promedio CTN", "Unidad": "veces", **rot_ctn.to_dict()})
    
    # 1.8 Rotación de Activos Fijos = Ingresos Operativos / Promedio Activos Fijos
    rot_af = safe_div(inc["Ingresos Operativos"], bal_avg["Activos Fijos"])
    ratios.append({"Categoría": "Actividad", "Ratio": "Rotación de Activos Fijos", "Fórmula": "Ingresos Operativos / Promedio Activos Fijos", "Unidad": "veces", **rot_af.to_dict()})
    
    # 1.9 Rotación de Activos Totales = Ingresos Operativos / Promedio Activos Totales
    rot_at = safe_div(inc["Ingresos Operativos"], bal_avg["Activos"])
    ratios.append({"Categoría": "Actividad", "Ratio": "Rotación de Activos Totales", "Fórmula": "Ingresos Operativos / Promedio Activos Totales", "Unidad": "veces", **rot_at.to_dict()})
    
    # 1.10 Eficiencia Operativa = (Costo de lo Vendido + Gastos Operativos) / Promedio Activos Totales
    eficiencia_op = safe_div(inc["Costo de lo Vendido"] + inc["Gastos Operativos Totales"], bal_avg["Activos"])
    ratios.append({"Categoría": "Actividad", "Ratio": "Eficiencia Operativa", "Fórmula": "(Costo + Gastos Op.) / Promedio Activos", "Unidad": "veces", **eficiencia_op.to_dict()})
    
    # 1.11 Eficiencia Administrativa = Gastos Operativos / Promedio Activos Totales
    eficiencia_adm = safe_div(inc["Gastos Operativos"], bal_avg["Activos"])
    ratios.append({"Categoría": "Actividad", "Ratio": "Eficiencia Administrativa", "Fórmula": "Gastos Operativos / Promedio Activos", "Unidad": "veces", **eficiencia_adm.to_dict()})
    
    # 1.12 Cobertura del FEAO = FEAO / (Dividendos + Deuda Corto Plazo + CAPEX)
    denom_feao = dividendos + bal["Deuda Corto Plazo"] + capex
    cob_feao = safe_div(feao, denom_feao)
    ratios.append({"Categoría": "Actividad", "Ratio": "Cobertura del FEAO", "Fórmula": "FEAO / (Dividendos + Deuda CP + CAPEX)", "Unidad": "veces", **cob_feao.to_dict()})

    # =========================================================
    # 2. RAZONES DE LIQUIDEZ (8 Ratios)
    # =========================================================
    # 2.1 Liquidez Corriente = Activo Corriente / Pasivo Corriente
    liq_cte = safe_div(bal["Activo Corriente"], bal["Pasivo Corriente"])
    ratios.append({"Categoría": "Liquidez", "Ratio": "Liquidez Corriente", "Fórmula": "Activo Corriente / Pasivo Corriente", "Unidad": "veces", **liq_cte.to_dict()})
    
    # 2.2 Prueba Ácida = (Activo Corriente - Inventarios) / Pasivo Corriente
    prueba_acida = safe_div(bal["Activo Corriente"] - inv, bal["Pasivo Corriente"])
    ratios.append({"Categoría": "Liquidez", "Ratio": "Prueba Ácida", "Fórmula": "(Activo Corriente - Inventarios) / Pasivo Corriente", "Unidad": "veces", **prueba_acida.to_dict()})
    
    # 2.3 Razón Rápida = (Disponibilidades + ICP + Cuentas por Cobrar) / Pasivo Corriente
    razon_rapida = safe_div(disponibilidades + cxc, bal["Pasivo Corriente"])
    ratios.append({"Categoría": "Liquidez", "Ratio": "Razón Rápida", "Fórmula": "(Disponibilidades + ICP + CxC) / Pasivo Corriente", "Unidad": "veces", **razon_rapida.to_dict()})
    
    # 2.4 Razón de Efectivo = (Disponibilidades + ICP) / Pasivo Corriente
    razon_efectivo = safe_div(disponibilidades, bal["Pasivo Corriente"])
    ratios.append({"Categoría": "Liquidez", "Ratio": "Razón de Efectivo", "Fórmula": "(Disponibilidades + ICP) / Pasivo Corriente", "Unidad": "veces", **razon_efectivo.to_dict()})
    
    # 2.5 Capital de Trabajo Neto = Activo Corriente - Pasivo Corriente
    ctn = bal["Activo Corriente"] - bal["Pasivo Corriente"]
    ratios.append({"Categoría": "Liquidez", "Ratio": "Capital de Trabajo Neto", "Fórmula": "Activo Corriente - Pasivo Corriente", "Unidad": "$", **ctn.to_dict()})
    
    # 2.6 Prueba Defensiva = (Disponibilidades + ICP + Cuentas por Cobrar) / Gastos Diarios en Efectivo
    gastos_diarios = (inc["Costo de lo Vendido"] + inc["Gastos Operativos"]) / 365.0
    prueba_defensiva = safe_div(disponibilidades + cxc, gastos_diarios)
    ratios.append({"Categoría": "Liquidez", "Ratio": "Prueba Defensiva", "Fórmula": "(Disp. + ICP + CxC) / Gastos Diarios en Efectivo", "Unidad": "días", **prueba_defensiva.to_dict()})
    
    # 2.7 Ciclo Operativo = PPC + EPI
    ciclo_operativo = ppc + epi
    ratios.append({"Categoría": "Liquidez", "Ratio": "Ciclo Operativo", "Fórmula": "PPC + EPI", "Unidad": "días", **ciclo_operativo.to_dict()})
    
    # 2.8 Ciclo de Conversión del Efectivo = PPC + EPI - PPP
    cce = ppc + epi - ppp
    ratios.append({"Categoría": "Liquidez", "Ratio": "Ciclo de Conversión del Efectivo", "Fórmula": "PPC + EPI - PPP", "Unidad": "días", **cce.to_dict()})

    # =========================================================
    # 3. RAZONES DE RENTABILIDAD (8 Ratios con Promedios de Balance)
    # =========================================================
    # 3.1 Margen Bruto = Resultado Bruto / Ingresos Operativos
    mb = safe_div(inc["Resultado Bruto"], inc["Ingresos Operativos"], multiplier=100.0)
    ratios.append({"Categoría": "Rentabilidad", "Ratio": "Margen Bruto", "Fórmula": "Resultado Bruto / Ingresos Operativos", "Unidad": "%", **mb.to_dict()})
    
    # 3.2 Margen EBITDA = EBITDA / Ingresos Operativos
    mebitda = safe_div(inc["EBITDA"], inc["Ingresos Operativos"], multiplier=100.0)
    ratios.append({"Categoría": "Rentabilidad", "Ratio": "Margen EBITDA", "Fórmula": "EBITDA / Ingresos Operativos", "Unidad": "%", **mebitda.to_dict()})
    
    # 3.3 Margen Operativo = Resultado Operativo / Ingresos Operativos
    mo = safe_div(inc["Resultado Operativo"], inc["Ingresos Operativos"], multiplier=100.0)
    ratios.append({"Categoría": "Rentabilidad", "Ratio": "Margen Operativo", "Fórmula": "Resultado Operativo / Ingresos Operativos", "Unidad": "%", **mo.to_dict()})
    
    # 3.4 Margen Neto = Resultado Neto / Ingresos Operativos
    mn = safe_div(inc["Resultado Neto"], inc["Ingresos Operativos"], multiplier=100.0)
    ratios.append({"Categoría": "Rentabilidad", "Ratio": "Margen Neto", "Fórmula": "Resultado Neto / Ingresos Operativos", "Unidad": "%", **mn.to_dict()})
    
    # 3.5 ROA = Resultado Neto / Promedio de Activos
    roa = safe_div(inc["Resultado Neto"], bal_avg["Activos"], multiplier=100.0)
    ratios.append({"Categoría": "Rentabilidad", "Ratio": "ROA", "Fórmula": "Resultado Neto / Promedio Activos", "Unidad": "%", **roa.to_dict()})
    
    # 3.6 ROA Operativo = Resultado Operativo / Promedio de Activos
    roa_op = safe_div(inc["Resultado Operativo"], bal_avg["Activos"], multiplier=100.0)
    ratios.append({"Categoría": "Rentabilidad", "Ratio": "ROA Operativo", "Fórmula": "Resultado Operativo / Promedio Activos", "Unidad": "%", **roa_op.to_dict()})
    
    # 3.7 ROE = Resultado Neto / Promedio de Patrimonio Neto
    roe = safe_div(inc["Resultado Neto"], bal_avg["Patrimonio"], multiplier=100.0)
    ratios.append({"Categoría": "Rentabilidad", "Ratio": "ROE", "Fórmula": "Resultado Neto / Promedio Patrimonio Neto", "Unidad": "%", **roe.to_dict()})
    
    # 3.8 ROIC = NOPAT / Promedio de Capital Invertido
    cap_inv = bal["Deuda Total"] + bal["Patrimonio"] - disponibilidades
    cap_inv_avg = get_series_average(cap_inv, years)
    roic = safe_div(inc["NOPAT"], cap_inv_avg, multiplier=100.0)
    ratios.append({"Categoría": "Rentabilidad", "Ratio": "ROIC", "Fórmula": "NOPAT / Promedio Capital Invertido", "Unidad": "%", **roic.to_dict()})

    # =========================================================
    # 4. RAZONES DE SOLVENCIA (12 Ratios con Promedios donde Aplica)
    # =========================================================
    # 4.1 Pasivo Total = Pasivo / Activo
    pasivo_total = safe_div(bal["Pasivos"], bal["Activos"], multiplier=100.0)
    ratios.append({"Categoría": "Solvencia", "Ratio": "Pasivo Total", "Fórmula": "Pasivo / Activo", "Unidad": "%", **pasivo_total.to_dict()})
    
    # 4.2 Pasivo a Patrimonio = Pasivo / Patrimonio Neto
    pasivo_patrimonio = safe_div(bal["Pasivos"], bal["Patrimonio"])
    ratios.append({"Categoría": "Solvencia", "Ratio": "Pasivo a Patrimonio", "Fórmula": "Pasivo / Patrimonio Neto", "Unidad": "veces", **pasivo_patrimonio.to_dict()})
    
    # 4.3 Deuda Total = Deuda / Activo
    deuda_activos = safe_div(bal["Deuda Total"], bal["Activos"], multiplier=100.0)
    ratios.append({"Categoría": "Solvencia", "Ratio": "Deuda Total", "Fórmula": "Deuda Financiera / Activo", "Unidad": "%", **deuda_activos.to_dict()})
    
    # 4.4 Deuda a Capital = Deuda / Capital (Deuda + Patrimonio)
    capital_total = bal["Deuda Total"] + bal["Patrimonio"]
    deuda_capital = safe_div(bal["Deuda Total"], capital_total)
    ratios.append({"Categoría": "Solvencia", "Ratio": "Deuda a Capital", "Fórmula": "Deuda / (Deuda + Patrimonio)", "Unidad": "veces", **deuda_capital.to_dict()})
    
    # 4.5 Deuda a Patrimonio = Deuda / Patrimonio Neto
    deuda_patrimonio = safe_div(bal["Deuda Total"], bal["Patrimonio"])
    ratios.append({"Categoría": "Solvencia", "Ratio": "Deuda a Patrimonio", "Fórmula": "Deuda / Patrimonio Neto", "Unidad": "veces", **deuda_patrimonio.to_dict()})
    
    # 4.6 Costo Implícito de la Deuda = Gastos Financieros / Promedio de Deuda
    costo_deuda = safe_div(inc["Gastos Financieros"], bal_avg["Deuda Total"], multiplier=100.0)
    ratios.append({"Categoría": "Solvencia", "Ratio": "Costo Implícito de la Deuda", "Fórmula": "Gastos Financieros / Promedio Deuda", "Unidad": "%", **costo_deuda.to_dict()})
    
    # 4.7 Plazo Implícito de la Deuda = Deuda Total / Deuda a CP
    plazo_deuda = safe_div(bal["Deuda Total"], bal["Deuda Corto Plazo"])
    ratios.append({"Categoría": "Solvencia", "Ratio": "Plazo Implícito de la Deuda", "Fórmula": "Deuda Total / Deuda a CP", "Unidad": "veces", **plazo_deuda.to_dict()})
    
    # 4.8 Deuda a EBITDA = Deuda / EBITDA
    deuda_ebitda = safe_div(bal["Deuda Total"], inc["EBITDA"])
    ratios.append({"Categoría": "Solvencia", "Ratio": "Deuda a EBITDA", "Fórmula": "Deuda Financiera / EBITDA", "Unidad": "veces", **deuda_ebitda.to_dict()})
    
    # 4.9 Multiplicador del Apalancamiento Financiero = Promedio Activos / Promedio Patrimonio Neto
    mult_apalancamiento = safe_div(bal_avg["Activos"], bal_avg["Patrimonio"])
    ratios.append({"Categoría": "Solvencia", "Ratio": "Multiplicador del Apalancamiento Financiero", "Fórmula": "Promedio Activos / Promedio Patrimonio", "Unidad": "veces", **mult_apalancamiento.to_dict()})
    
    # 4.10 Cobertura de Intereses = Resultado Operativo / Gastos Financieros
    cob_intereses = safe_div(inc["Resultado Operativo"], inc["Gastos Financieros"])
    ratios.append({"Categoría": "Solvencia", "Ratio": "Cobertura de Intereses", "Fórmula": "Resultado Operativo / Gastos Financieros", "Unidad": "veces", **cob_intereses.to_dict()})
    
    # 4.11 EBITDA a Servicio de la Deuda = EBITDA / (Gastos Financieros + Deuda a CP)
    servicio_deuda = inc["Gastos Financieros"] + bal["Deuda Corto Plazo"]
    ebitda_servicio = safe_div(inc["EBITDA"], servicio_deuda)
    ratios.append({"Categoría": "Solvencia", "Ratio": "EBITDA a Servicio de la Deuda", "Fórmula": "EBITDA / (Gastos Fin. + Deuda CP)", "Unidad": "veces", **ebitda_servicio.to_dict()})
    
    # 4.12 Cobertura del Servicio de la Deuda = (EBITDA + Activo Corriente) / (Gastos Financieros + Deuda a CP)
    cob_servicio = safe_div(inc["EBITDA"] + bal["Activo Corriente"], servicio_deuda)
    ratios.append({"Categoría": "Solvencia", "Ratio": "Cobertura del Servicio de la Deuda", "Fórmula": "(EBITDA + Activo Corriente) / (Gastos Fin. + Deuda CP)", "Unidad": "veces", **cob_servicio.to_dict()})

    return pd.DataFrame(ratios)


# =========================================================
# 2. DESAGREGACIÓN DUPONT Y ANÁLISIS DE ATRIBUCIÓN (ROOT CAUSE)
# =========================================================
def calculate_dupont_3_step(df: pd.DataFrame) -> pd.DataFrame:
    """
    Descomposición DuPont de 3 Pasos con promedios de balance:
    ROE = Margen Neto × Rotación de Activos × Apalancamiento Financiero
    Donde:
    - Margen Neto = Resultado Neto / Ingresos Operativos
    - Rotación de Activos = Ingresos Operativos / Promedio Activos Totales
    - Apalancamiento Financiero = Promedio Activos / Promedio Patrimonio Neto
    """
    years = get_years_from_df(df)
    bal_avg = get_balance_averages(df)
    inc = get_income_statement_totals(df)
    
    margen_neto = inc["Resultado Neto"] / np.where(inc["Ingresos Operativos"] != 0, inc["Ingresos Operativos"], np.nan)
    rotacion_activos = inc["Ingresos Operativos"] / np.where(bal_avg["Activos"] != 0, bal_avg["Activos"], np.nan)
    multiplicador_capital = bal_avg["Activos"] / np.where(bal_avg["Patrimonio"] != 0, bal_avg["Patrimonio"], np.nan)
    
    roa_d = (margen_neto * rotacion_activos) * 100.0
    roe_d = (margen_neto * rotacion_activos * multiplicador_capital) * 100.0
    
    records = [
        {"Paso": "1. Margen Neto (Resultado Neto / Ingresos)", "Tipo": "Eficiencia Operativa", "Unidad": "%", **(margen_neto * 100.0).to_dict()},
        {"Paso": "2. Rotación de Activos (Ingresos / Promedio Activos)", "Tipo": "Eficiencia en Uso de Activos", "Unidad": "veces", **rotacion_activos.to_dict()},
        {"Paso": "ROA DuPont (Paso 1 × Paso 2)", "Tipo": "Retorno sobre Activos", "Unidad": "%", **roa_d.to_dict()},
        {"Paso": "3. Multiplicador de Capital (Prom. Activos / Prom. Patrimonio)", "Tipo": "Apalancamiento Financiero", "Unidad": "veces", **multiplicador_capital.to_dict()},
        {"Paso": "ROE DuPont (Paso 1 × Paso 2 × Paso 3)", "Tipo": "Rentabilidad del Accionista", "Unidad": "%", **roe_d.to_dict()}
    ]
    return pd.DataFrame(records)


def calculate_dupont_5_step(df: pd.DataFrame) -> pd.DataFrame:
    """
    Descomposición DuPont de 5 Pasos con promedios de balance:
    ROE = Carga Fiscal × Carga Financiera × Margen Operativo × Rotación Activos × Apalancamiento
    """
    years = get_years_from_df(df)
    bal_avg = get_balance_averages(df)
    inc = get_income_statement_totals(df)
    
    carga_fiscal = inc["Resultado Neto"] / np.where(inc["EBT"] != 0, inc["EBT"], np.nan)
    carga_financiera = inc["EBT"] / np.where(inc["Resultado Operativo"] != 0, inc["Resultado Operativo"], np.nan)
    margen_operativo = inc["Resultado Operativo"] / np.where(inc["Ingresos Operativos"] != 0, inc["Ingresos Operativos"], np.nan)
    rotacion_activos = inc["Ingresos Operativos"] / np.where(bal_avg["Activos"] != 0, bal_avg["Activos"], np.nan)
    multiplicador_capital = bal_avg["Activos"] / np.where(bal_avg["Patrimonio"] != 0, bal_avg["Patrimonio"], np.nan)
    
    roe_calculado = (carga_fiscal * carga_financiera * margen_operativo * rotacion_activos * multiplicador_capital) * 100.0
    
    records = [
        {"Paso": "1. Carga Fiscal (Resultado Neto / EBT)", "Interpretación": "Impacto tributario (1 - Tasa Efectiva)", "Unidad": "factor", **carga_fiscal.to_dict()},
        {"Paso": "2. Carga Financiera (EBT / Resultado Operativo)", "Interpretación": "Impacto de la deuda/intereses", "Unidad": "factor", **carga_financiera.to_dict()},
        {"Paso": "3. Margen Operativo (Resultado Op. / Ingresos)", "Interpretación": "Eficacia operativa del negocio", "Unidad": "%", **(margen_operativo * 100.0).to_dict()},
        {"Paso": "4. Rotación de Activos (Ingresos / Prom. Activos)", "Interpretación": "Rotación y productividad de activos", "Unidad": "veces", **rotacion_activos.to_dict()},
        {"Paso": "5. Multiplicador de Capital (Prom. Activos / Prom. Patrimonio)", "Interpretación": "Apalancamiento patrimonial", "Unidad": "veces", **multiplicador_capital.to_dict()},
        {"Paso": "ROE DuPont 5 Pasos (Producto de 1 a 5)", "Interpretación": "Rentabilidad sobre Capital Contable", "Unidad": "%", **roe_calculado.to_dict()}
    ]
    return pd.DataFrame(records)


def calculate_dupont_attribution(df: pd.DataFrame, base_year: str, target_year: str) -> Tuple[Dict[str, Any], pd.DataFrame]:
    """
    Análisis de Sensibilidad y Atribución Causa Raíz de DuPont para ROE y ROA entre dos periodos.
    Descompone la variación en:
    1. Efecto Margen Neto, Efecto Rotación de Activos, Efecto Apalancamiento Financiero
    2. Atribución específica por cuenta contable:
       - Estado de Resultados (Ingresos, Costos, Gastos Operativos, Gastos Financieros, Impuestos)
       - Activos Corrientes y No Corrientes (Disponibilidades, CxC, Inventarios, PPE)
       - Pasivos Corrientes y No Corrientes (CxP, Deudas Financieras, etc.)
    """
    bal = get_balance_totals(df)
    bal_avg = get_balance_averages(df)
    inc = get_income_statement_totals(df)
    
    y0, y1 = base_year, target_year
    
    # 1. Factores DuPont para y0 e y1
    mn0 = float(inc["Resultado Neto"][y0]) / float(inc["Ingresos Operativos"][y0]) if inc["Ingresos Operativos"][y0] else 0
    mn1 = float(inc["Resultado Neto"][y1]) / float(inc["Ingresos Operativos"][y1]) if inc["Ingresos Operativos"][y1] else 0
    
    rot0 = float(inc["Ingresos Operativos"][y0]) / float(bal_avg["Activos"][y0]) if bal_avg["Activos"][y0] else 0
    rot1 = float(inc["Ingresos Operativos"][y1]) / float(bal_avg["Activos"][y1]) if bal_avg["Activos"][y1] else 0
    
    mult0 = float(bal_avg["Activos"][y0]) / float(bal_avg["Patrimonio"][y0]) if bal_avg["Patrimonio"][y0] else 0
    mult1 = float(bal_avg["Activos"][y1]) / float(bal_avg["Patrimonio"][y1]) if bal_avg["Patrimonio"][y1] else 0
    
    roe0 = mn0 * rot0 * mult0 * 100.0
    roe1 = mn1 * rot1 * mult1 * 100.0
    delta_roe = roe1 - roe0
    
    roa0 = mn0 * rot0 * 100.0
    roa1 = mn1 * rot1 * 100.0
    delta_roa = roa1 - roa0
    
    # Descomposición logarítmica aditiva de primer orden (Log-Mean / Taylor)
    # ΔROE = Efecto Margen + Efecto Rotación + Efecto Multiplicador
    avg_mult = (mult0 + mult1) / 2.0
    avg_rot = (rot0 + rot1) / 2.0
    avg_mn = (mn0 + mn1) / 2.0
    
    efecto_margen = (mn1 - mn0) * avg_rot * avg_mult * 100.0
    efecto_rotacion = avg_mn * (rot1 - rot0) * avg_mult * 100.0
    efecto_apalancamiento = avg_mn * avg_rot * (mult1 - mult0) * 100.0
    
    # Atribución por cuenta específica
    attribution_records = []
    
    # Subcuentas de Resultados (Impacto en Margen)
    ventas0, ventas1 = float(inc["Ingresos Operativos"][y0]), float(inc["Ingresos Operativos"][y1])
    costo0, costo1 = float(inc["Costo de lo Vendido"][y0]), float(inc["Costo de lo Vendido"][y1])
    sga0, sga1 = float(inc["Gastos Operativos"][y0]), float(inc["Gastos Operativos"][y1])
    depr0, depr1 = float(inc["Depreciacion y Amortizacion"][y0]), float(inc["Depreciacion y Amortizacion"][y1])
    gf0, gf1 = float(inc["Gastos Financieros"][y0]), float(inc["Gastos Financieros"][y1])
    imp0, imp1 = float(inc["Impuestos"][y0]), float(inc["Impuestos"][y1])
    
    def_factor_mn = (avg_rot * avg_mult * 100.0) / (ventas1 if ventas1 != 0 else 1.0)
    
    delta_ventas = ventas1 - ventas0
    delta_costos = -(costo1 - costo0)
    delta_sga = -(sga1 - sga0)
    delta_depr = -(depr1 - depr0)
    delta_gf = -(gf1 - gf0)
    delta_imp = -(imp1 - imp0)
    
    accounts_er = [
        ("Ingresos por Ventas", "Estado de Resultados", delta_ventas, delta_ventas * def_factor_mn * (mn1 / (abs(mn1) if mn1 != 0 else 1))),
        ("Costo de lo Vendido", "Estado de Resultados", -(costo1 - costo0), delta_costos * def_factor_mn),
        ("Gastos de Administración y Ventas", "Estado de Resultados", -(sga1 - sga0), delta_sga * def_factor_mn),
        ("Depreciación y Amortización", "Estado de Resultados", -(depr1 - depr0), delta_depr * def_factor_mn),
        ("Gastos Financieros (Intereses)", "Estado de Resultados", -(gf1 - gf0), delta_gf * def_factor_mn),
        ("Impuesto sobre la Renta", "Estado de Resultados", -(imp1 - imp0), delta_imp * def_factor_mn),
    ]
    
    for acct, cat, d_val, imp_roe in accounts_er:
        attribution_records.append({
            "Cuenta": acct,
            "Área / Categoría": cat,
            "Variación ($)": d_val,
            "Impacto Estimado ROE (p.p.)": imp_roe,
            "Impacto Estimado ROA (p.p.)": imp_roe / (avg_mult if avg_mult != 0 else 1.0),
            "Efecto": "Driver de Crecimiento (+)" if imp_roe > 0.05 else ("Factor Detractor (-)" if imp_roe < -0.05 else "Neutral")
        })
        
    # Subcuentas de Balance (Impacto en Rotación y Apalancamiento)
    cxc0, cxc1 = float(get_account_series(df, "Cuentas por Cobrar")[y0]), float(get_account_series(df, "Cuentas por Cobrar")[y1])
    inv0, inv1 = float(get_account_series(df, "Inventarios")[y0]), float(get_account_series(df, "Inventarios")[y1])
    ppe0, ppe1 = float(get_account_series(df, "Propiedad Planta y Equipo")[y0]), float(get_account_series(df, "Propiedad Planta y Equipo")[y1])
    cxp0, cxp1 = float(get_account_series(df, "Cuentas por Pagar")[y0]), float(get_account_series(df, "Cuentas por Pagar")[y1])
    deuda0, deuda1 = float(bal["Deuda Total"][y0]), float(bal["Deuda Total"][y1])
    
    # Mayor activo reduce rotación (ceteris paribus)
    act_total_avg = float(bal_avg["Activos"][y1]) if float(bal_avg["Activos"][y1]) != 0 else 1.0
    factor_rot_acct = - (avg_mn * avg_mult * 100.0 * (rot1 / act_total_avg))
    
    accounts_bs = [
        ("Cuentas por Cobrar", "Activo Corriente", cxc1 - cxc0, (cxc1 - cxc0) * factor_rot_acct),
        ("Inventarios", "Activo Corriente", inv1 - inv0, (inv1 - inv0) * factor_rot_acct),
        ("Propiedad, Planta y Equipo (CAPEX)", "Activo No Corriente", ppe1 - ppe0, (ppe1 - ppe0) * factor_rot_acct),
        ("Cuentas por Pagar Proveedores", "Pasivo Corriente", cxp1 - cxp0, (cxp1 - cxp0) * (efecto_apalancamiento / (bal["Pasivos"][y1] if bal["Pasivos"][y1] else 1.0))),
        ("Deuda Financiera Total", "Pasivo / Apalancamiento", deuda1 - deuda0, (deuda1 - deuda0) * (efecto_apalancamiento / (bal["Deuda Total"][y1] if bal["Deuda Total"][y1] else 1.0)))
    ]
    
    for acct, cat, d_val, imp_roe in accounts_bs:
        attribution_records.append({
            "Cuenta": acct,
            "Área / Categoría": cat,
            "Variación ($)": d_val,
            "Impacto Estimado ROE (p.p.)": imp_roe,
            "Impacto Estimado ROA (p.p.)": imp_roe / (avg_mult if avg_mult != 0 else 1.0),
            "Efecto": "Driver de Crecimiento (+)" if imp_roe > 0.05 else ("Factor Detractor (-)" if imp_roe < -0.05 else "Neutral")
        })
        
    attr_df = pd.DataFrame(attribution_records)
    # Ordenar por magnitud del impacto
    attr_df = attr_df.sort_values(by="Impacto Estimado ROE (p.p.)", ascending=False).reset_index(drop=True)
    
    summary = {
        "base_year": base_year,
        "target_year": target_year,
        "roe_base": roe0,
        "roe_target": roe1,
        "delta_roe": delta_roe,
        "roa_base": roa0,
        "roa_target": roa1,
        "delta_roa": delta_roa,
        "efecto_margen": efecto_margen,
        "efecto_rotacion": efecto_rotacion,
        "efecto_apalancamiento": efecto_apalancamiento,
        "margen_neto_base": mn0 * 100.0,
        "margen_neto_target": mn1 * 100.0,
        "rotacion_activos_base": rot0,
        "rotacion_activos_target": rot1,
        "multiplicador_capital_base": mult0,
        "multiplicador_capital_target": mult1
    }
    
    return summary, attr_df


# =========================================================
# 3. FLUJOS DE EFECTIVO LIBRES (FELF Y FELP) — OPCIÓN 1 Y OPCIÓN 2
# =========================================================
def calculate_felf_felp(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    """
    Calcula:
    1. Flujo de Efectivo Libre para la Firma (FELF / FCFF):
       - Opción 1: FELF = RAII - IR + PNGME - CAPEX - MNAPO
       - Opción 2: FELF = FEAO + GF * (1 - τ) - CAPEX
    2. Flujo de Efectivo Libre para el Patrimonio (FELP / FCFE):
       - Opción 1: FELP = RN + PNGME - CAPEX - MNAPO + ΔD
       - Opción 2: FELP = FEAO - CAPEX + ΔD
       
    Retorna:
    - Tabla FELF detallada (Opciones 1 y 2)
    - Tabla FELP detallada (Opciones 1 y 2)
    - Tabla Comparativa de Conciliación
    - Diccionario de Componentes para Gráficos
    """
    years = get_years_from_df(df)
    bal = get_balance_totals(df)
    inc = get_income_statement_totals(df)
    
    raii = inc["Resultado Operativo"]
    tasa_tax = inc["Tasa Efectiva"]
    ir = raii * tasa_tax
    pngme = inc["Depreciacion y Amortizacion"]
    capex = get_account_series(df, "CAPEX")
    if capex.sum() == 0:
        capex = get_account_series(df, "Inversiones en Activo Fijo")
        
    cxc = get_account_series(df, "Cuentas por Cobrar")
    inv = get_account_series(df, "Inventarios")
    cxp = get_account_series(df, "Cuentas por Pagar")
    
    ktno = (cxc + inv) - cxp
    mnapo = pd.Series(0.0, index=years)
    for i in range(1, len(years)):
        mnapo[years[i]] = ktno[years[i]] - ktno[years[i - 1]]
        
    # FEAO (Flujo de Efectivo de Actividades de Operación)
    feao_data = get_account_series(df, "FEAO")
    if feao_data.sum() == 0:
        feao_data = get_account_series(df, "Flujo de Efectivo Actividades Operativas")
        
    feao = pd.Series(index=years, dtype=float)
    for i, y in enumerate(years):
        if feao_data.get(y, 0) != 0:
            feao[y] = float(feao_data[y])
        else:
            feao[y] = float(inc["Resultado Neto"][y]) + float(pngme[y]) - float(mnapo[y])
            
    gf = inc["Gastos Financieros"]
    rn = inc["Resultado Neto"]
    
    # Variación neta de deuda (ΔD = Deuda_t - Deuda_{t-1})
    delta_deuda = pd.Series(0.0, index=years)
    for i in range(1, len(years)):
        delta_deuda[years[i]] = bal["Deuda Total"][years[i]] - bal["Deuda Total"][years[i - 1]]
        
    # FELF Opción 1: RAII - IR + PNGME - CAPEX - MNAPO
    felf_op1 = raii - ir + pngme - capex - mnapo
    
    # FELF Opción 2: FEAO + GF * (1 - τ) - CAPEX
    felf_op2 = feao + (gf * (1.0 - tasa_tax)) - capex
    
    # FELP Opción 1: RN + PNGME - CAPEX - MNAPO + ΔD
    felp_op1 = rn + pngme - capex - mnapo + delta_deuda
    
    # FELP Opción 2: FEAO - CAPEX + ΔD
    felp_op2 = feao - capex + delta_deuda

    # Construir tablas estructuradas
    felf_rows = [
        {"Opción": "Opción 1", "Concepto": "(+) Resultado Operativo (RAII / EBIT)", "Tipo": "Base Operativa", **raii.to_dict()},
        {"Opción": "Opción 1", "Concepto": "(-) Impuesto Calculado (IR = RAII × τ)", "Tipo": "Fiscal", **(-ir).to_dict()},
        {"Opción": "Opción 1", "Concepto": "(=) NOPAT (RAII - IR)", "Tipo": "Subtotal Operativo", **(raii - ir).to_dict()},
        {"Opción": "Opción 1", "Concepto": "(+) Provisiones y Gastos No Monetarios (PNGME / D&A)", "Tipo": "No desembolsable", **pngme.to_dict()},
        {"Opción": "Opción 1", "Concepto": "(-) Inversiones de Capital (CAPEX)", "Tipo": "Inversión", **(-capex).to_dict()},
        {"Opción": "Opción 1", "Concepto": "(-/+) Variación Capital de Trabajo Neto Operativo (MNAPO)", "Tipo": "Operativo Circulante", **(-mnapo).to_dict()},
        {"Opción": "Opción 1", "Concepto": "(=) FELF — Flujo de Efectivo Libre Firma (Opción 1)", "Tipo": "Resultado", **felf_op1.to_dict()},
        {"Opción": "---", "Concepto": "--------------------------------------------------------", "Tipo": "Separador", **{y: 0.0 for y in years}},
        {"Opción": "Opción 2", "Concepto": "(+) Flujo Efectivo Actividades Operativas (FEAO)", "Tipo": "Flujo Efectivo", **feao.to_dict()},
        {"Opción": "Opción 2", "Concepto": "(+) Escudo Fiscal de Intereses (GF × (1 - τ))", "Tipo": "Financiero Fiscal", **(gf * (1.0 - tasa_tax)).to_dict()},
        {"Opción": "Opción 2", "Concepto": "(-) Inversiones de Capital (CAPEX)", "Tipo": "Inversión", **(-capex).to_dict()},
        {"Opción": "Opción 2", "Concepto": "(=) FELF — Flujo de Efectivo Libre Firma (Opción 2)", "Tipo": "Resultado", **felf_op2.to_dict()}
    ]
    felf_df = pd.DataFrame(felf_rows)

    felp_rows = [
        {"Opción": "Opción 1", "Concepto": "(+) Resultado Neto (RN)", "Tipo": "Base Patrimonial", **rn.to_dict()},
        {"Opción": "Opción 1", "Concepto": "(+) Provisiones y Gastos No Monetarios (PNGME / D&A)", "Tipo": "No desembolsable", **pngme.to_dict()},
        {"Opción": "Opción 1", "Concepto": "(-) Inversiones de Capital (CAPEX)", "Tipo": "Inversión", **(-capex).to_dict()},
        {"Opción": "Opción 1", "Concepto": "(-/+) Variación Capital Trabajo Neto Operativo (MNAPO)", "Tipo": "Operativo Circulante", **(-mnapo).to_dict()},
        {"Opción": "Opción 1", "Concepto": "(+/-) Variación Neta de Deuda Financiera (ΔD)", "Tipo": "Financiamiento", **delta_deuda.to_dict()},
        {"Opción": "Opción 1", "Concepto": "(=) FELP — Flujo Libre del Patrimonio (Opción 1)", "Tipo": "Resultado", **felp_op1.to_dict()},
        {"Opción": "---", "Concepto": "--------------------------------------------------------", "Tipo": "Separador", **{y: 0.0 for y in years}},
        {"Opción": "Opción 2", "Concepto": "(+) Flujo Efectivo Actividades Operativas (FEAO)", "Tipo": "Flujo Efectivo", **feao.to_dict()},
        {"Opción": "Opción 2", "Concepto": "(-) Inversiones de Capital (CAPEX)", "Tipo": "Inversión", **(-capex).to_dict()},
        {"Opción": "Opción 2", "Concepto": "(+/-) Variación Neta de Deuda Financiera (ΔD)", "Tipo": "Financiamiento", **delta_deuda.to_dict()},
        {"Opción": "Opción 2", "Concepto": "(=) FELP — Flujo Libre del Patrimonio (Opción 2)", "Tipo": "Resultado", **felp_op2.to_dict()}
    ]
    felp_df = pd.DataFrame(felp_rows)

    # Tabla comparativa y de discrepancia
    diff_felf = felf_op1 - felf_op2
    diff_felp = felp_op1 - felp_op2
    comp_rows = [
        {"Flujo": "FELF (Firma)", "Metodología": "Opción 1 (RAII - IR + D&A - CAPEX - MNAPO)", **felf_op1.to_dict()},
        {"Flujo": "FELF (Firma)", "Metodología": "Opción 2 (FEAO + GF(1-t) - CAPEX)", **felf_op2.to_dict()},
        {"Flujo": "FELF (Firma)", "Metodología": "Diferencia de Conciliación", **diff_felf.to_dict()},
        {"Flujo": "FELP (Patrimonio)", "Metodología": "Opción 1 (RN + D&A - CAPEX - MNAPO + ΔD)", **felp_op1.to_dict()},
        {"Flujo": "FELP (Patrimonio)", "Metodología": "Opción 2 (FEAO - CAPEX + ΔD)", **felp_op2.to_dict()},
        {"Flujo": "FELP (Patrimonio)", "Metodología": "Diferencia de Conciliación", **diff_felp.to_dict()}
    ]
    comp_df = pd.DataFrame(comp_rows)

    components = {
        "years": years,
        "raii": raii,
        "ir": ir,
        "nopat": raii - ir,
        "pngme": pngme,
        "depr_amort": pngme,
        "capex": capex,
        "mnapo": mnapo,
        "delta_ktno": mnapo,
        "feao": feao,
        "gf": gf,
        "rn": rn,
        "delta_deuda": delta_deuda,
        "tasa_tax": tasa_tax,
        "felf_op1": felf_op1,
        "felf_op2": felf_op2,
        "felp_op1": felp_op1,
        "felp_op2": felp_op2,
        "ebitda": inc["EBITDA"],
        "ktno": ktno,
        "fcf": felf_op1  # Compatibilidad con módulos existentes
    }

    return felf_df, felp_df, comp_df, components


# =========================================================
# 4. ANÁLISIS VERTICAL Y HORIZONTAL (NORMATIVA BOLIVIA NC 3 / NC 6)
# =========================================================
def calculate_vertical_analysis(df: pd.DataFrame, estado: str) -> pd.DataFrame:
    """
    Calcula el Análisis Vertical según la práctica contable boliviana:
    - Balance General: porcentaje relativo al Activo Total (o Pasivo + Patrimonio).
    - Estado de Resultados: porcentaje relativo a los Ingresos Operativos (Ventas Netas).
    """
    years = get_years_from_df(df)
    sub_df = df[df["Estado"] == estado].copy()
    
    if estado == "Balance General":
        totals = get_balance_totals(df)
        base = totals["Activos"]
    else:
        inc_totals = get_income_statement_totals(df)
        base = inc_totals["Ingresos Operativos"]
        
    result_df = sub_df[["Categoria", "Cuenta"]].copy()
    
    for y in years:
        denom = base[y] if base.get(y, 0) != 0 else 1.0
        result_df[f"{y} ($)"] = sub_df[y].astype(float)
        result_df[f"{y} (%)"] = (sub_df[y].astype(float) / denom) * 100.0
        
    return result_df


def calculate_horizontal_analysis(
    df: pd.DataFrame,
    estado: str,
    compare_mode: str = "previous",
    base_year: Optional[str] = None,
    apply_ufv: bool = False,
    ufv_dict: Optional[Dict[str, float]] = None
) -> pd.DataFrame:
    """
    Calcula el Análisis Horizontal con opciones avanzadas de normativa contable (NC 3 / NC 6):
    - compare_mode: 'previous' (compara respecto al año anterior inmediato t vs t-1)
                    'base' (compara todos los años respecto a un año base fijo)
    - apply_ufv: reexpresa los valores a moneda constante / homogénea mediante la razón de UFVs.
      Factor = UFV_t / UFV_referencia.
    - Genera para cada cuenta:
      - Monto nominal o reexpresado
      - Variación Absoluta ($)
      - Variación Porcentual (%)
    """
    years = get_years_from_df(df)
    sub_df = df[df["Estado"] == estado].copy()
    result_df = sub_df[["Categoria", "Cuenta"]].copy()
    
    # Resolver diccionario de UFVs
    ufvs = DEFAULT_UFV_RATES.copy()
    if ufv_dict:
        ufvs.update(ufv_dict)
    else:
        # Verificar si hay fila UFV en Datos Complementarios
        ufv_series = get_account_series(df, "UFV")
        for y in years:
            if ufv_series.get(y, 0) > 0:
                ufvs[y] = float(ufv_series[y])
                
    # Montos por año
    amounts = {}
    for y in years:
        amounts[y] = sub_df[y].astype(float).values
        result_df[f"{y} ($)"] = amounts[y]
        
    if compare_mode == "base":
        ref_year = base_year if base_year in years else years[0]
        ref_vals = amounts[ref_year]
        ref_ufv = ufvs.get(ref_year, 1.0)
        
        for y in years:
            if y == ref_year:
                continue
            curr_vals = amounts[y]
            curr_ufv = ufvs.get(y, 1.0)
            
            if apply_ufv:
                # Reexpresión NC 3: el monto base se ajusta a poder adquisitivo del año y
                factor_ufv = curr_ufv / ref_ufv if ref_ufv != 0 else 1.0
                adjusted_ref = ref_vals * factor_ufv
                diff_abs = curr_vals - adjusted_ref
                denom = np.where(np.abs(adjusted_ref) > 1e-6, np.abs(adjusted_ref), np.nan)
                diff_pct = np.where(~np.isnan(denom), (diff_abs / denom) * 100.0, 0.0)
                
                col_name = f"Var {y} vs Base {ref_year} (NC3 UFV)"
            else:
                diff_abs = curr_vals - ref_vals
                denom = np.where(np.abs(ref_vals) > 1e-6, np.abs(ref_vals), np.nan)
                diff_pct = np.where(~np.isnan(denom), (diff_abs / denom) * 100.0, 0.0)
                col_name = f"Var {y} vs Base {ref_year}"
                
            result_df[f"{col_name} ($)"] = diff_abs
            result_df[f"{col_name} (%)"] = diff_pct
            
    else:
        # Comparación respecto al año anterior inmediato (t vs t-1)
        for i in range(1, len(years)):
            y_prev = years[i - 1]
            y_curr = years[i]
            
            curr_vals = amounts[y_curr]
            prev_vals = amounts[y_prev]
            
            if apply_ufv:
                ufv_prev = ufvs.get(y_prev, 1.0)
                ufv_curr = ufvs.get(y_curr, 1.0)
                factor_ufv = ufv_curr / ufv_prev if ufv_prev != 0 else 1.0
                adjusted_prev = prev_vals * factor_ufv
                
                diff_abs = curr_vals - adjusted_prev
                denom = np.where(np.abs(adjusted_prev) > 1e-6, np.abs(adjusted_prev), np.nan)
                diff_pct = np.where(~np.isnan(denom), (diff_abs / denom) * 100.0, 0.0)
                col_name = f"Var {y_curr}/{y_prev} (NC3 UFV)"
            else:
                diff_abs = curr_vals - prev_vals
                denom = np.where(np.abs(prev_vals) > 1e-6, np.abs(prev_vals), np.nan)
                diff_pct = np.where(~np.isnan(denom), (diff_abs / denom) * 100.0, 0.0)
                col_name = f"Var {y_curr}/{y_prev}"
                
            result_df[f"{col_name} ($)"] = diff_abs
            result_df[f"{col_name} (%)"] = diff_pct

    return result_df


# Función de compatibilidad con módulos anteriores
def calculate_fcf_and_valuation(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    felf_df, felp_df, comp_df, components = calculate_felf_felp(df)
    return felf_df, components
