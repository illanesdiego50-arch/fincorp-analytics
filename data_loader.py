"""
Módulo de Carga, Validación y Procesamiento de Datos Financieros para Bolivia
FinCorp Analytics — Suite Integral de Finanzas Corporativas (TIC III - UMSA)

Características Clave:
1. Limpieza y normalización de formatos numéricos bolivianos:
   - Separadores de miles con punto (ej. 25.015.392) o coma (25,015,392).
   - Signos negativos con espacios (ej. '- 288.510.850') o entre paréntesis (ej. '(288.510.850)').
   - Valores nulos, guiones ('-', '—') y vacíos convertidos a 0.0.
   - Cifras UFV con 5 decimales (ej. 2.37376, 2.40898).
2. Detección automática de fila UFV en cabecera o cuerpo e integración en Datos Complementarios (NC 3 / NC 6).
3. Omisión de subtotales y totales para evitar doble cómputo al consolidar categorías contables.
4. Motor de mapeo dinámico de sinónimos contables estándar en Bolivia (empresas comerciales e industriales).
5. Soporte para archivos con Balance General y Estado de Resultados en la misma hoja (uno debajo del otro) o en hojas separadas.
"""

import io
import re
import unicodedata
from typing import Tuple, List, Dict, Any, Optional
import pandas as pd
import numpy as np


# =========================================================
# 1. MOTOR DE LIMPIEZA NUMÉRICA (FORMATO BOLIVIANO)
# =========================================================
def clean_numeric_value(val: Any) -> float:
    """
    Convierte cualquier representación numérica (texto, float, int) en float estándar de Python.
    Maneja:
    - Puntos como separador de miles: '25.015.392' -> 25015392.0
    - Negativos con espacio: '- 288.510.850' -> -288510850.0
    - Negativos entre paréntesis: '(288.510.850)' -> -288510850.0
    - Guiones y nulos: '-', '—', 'nan', '' -> 0.0
    - Valores decimales con coma o punto: '2.37376' o '2,37376' -> 2.37376
    - Símbolos de moneda: 'Bs. 120.000', '$ 45.000' -> 120000.0, 45000.0
    """
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val) if not np.isnan(val) else 0.0

    s = str(val).strip()
    if s in ["-", "—", "–", "", "None", "nan", "null", "N/A", "n/a", ".", ",", "0"]:
        return 0.0

    # Detectar signo negativo
    is_neg = False
    if s.startswith("(") and s.endswith(")"):
        is_neg = True
        s = s[1:-1].strip()
    elif s.startswith("-"):
        is_neg = True
        s = s[1:].strip()
    elif s.endswith("-"):
        is_neg = True
        s = s[:-1].strip()

    # Eliminar símbolos de moneda y texto no numérico
    s = s.replace("Bs.", "").replace("Bs", "").replace("BS", "").replace("bs", "")
    s = s.replace("$", "").replace("USD", "").replace("BOB", "").strip()
    
    # Conservar solo dígitos, puntos y comas
    s = re.sub(r"[^\d.,]", "", s)
    if not s:
        return 0.0

    # Determinar rol de '.' y ','
    if "." in s and "," in s:
        # Ejemplo: 25.015.392,50 (Punto miles, coma decimal)
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            # Ejemplo: 25,015,392.50 (Coma miles, punto decimal)
            s = s.replace(",", "")
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 2:
            # Múltiples puntos: 25.015.392 -> separador de miles
            s = "".join(parts)
        elif len(parts) == 2:
            # Un solo punto. Distinguir entre separador de miles (ej. 150.000) o decimal UFV (ej. 2.37376)
            # En Bolivia, los montos monetarios usan puntos de miles (3 dígitos a la derecha).
            # La UFV tiene valor entre 1 y 10 y 4-5 dígitos decimales (ej. 2.37376)
            int_part, dec_part = parts[0], parts[1]
            if len(dec_part) == 3 and not (len(int_part) == 1 and int_part in ["1", "2", "3"]):
                s = int_part + dec_part
            else:
                # Decimal genuino (ej. 2.37376 o 15.5)
                s = int_part + "." + dec_part
    elif "," in s:
        parts = s.split(",")
        if len(parts) > 2:
            s = "".join(parts)
        elif len(parts) == 2:
            int_part, dec_part = parts[0], parts[1]
            if len(dec_part) == 3 and not (len(int_part) == 1 and int_part in ["1", "2", "3"]):
                s = int_part + dec_part
            else:
                s = int_part + "." + dec_part

    try:
        num = float(s)
        return -num if is_neg else num
    except ValueError:
        return 0.0


# =========================================================
# 2. FILTRADO DE SUBTOTALES Y TOTALES (PREVENCIÓN DUPLICADOS)
# =========================================================
def normalize_text(text: str) -> str:
    """Normaliza texto eliminando tildes, signos de puntuación y espacios redundantes."""
    if not text or pd.isna(text):
        return ""
    text = unicodedata.normalize('NFKD', str(text)).encode('ASCII', 'ignore').decode('utf-8')
    text = text.lower().strip()
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


SUBTOTAL_KEYWORDS = [
    # Activos
    "total activo", "total activos", "activo total", "activos totales", "total del activo",
    "total activo corriente", "total activos corrientes", "activo corriente total",
    "total activo no corriente", "total activos no corrientes", "activo no corriente total",
    "total bienes de uso", "total activo fijo", "total existencias", "total disponibilidades",
    
    # Pasivos
    "total pasivo", "total pasivos", "pasivo total", "pasivos totales", "total del pasivo",
    "total pasivo corriente", "total pasivos corrientes", "pasivo corriente total",
    "total pasivo no corriente", "total pasivos no corrientes", "pasivo no corriente total",
    "total deudas comerciales", "total obligaciones financieras",
    
    # Patrimonio
    "total patrimonio", "total patrimonio neto", "patrimonio total", "patrimonio neto total",
    "total del patrimonio", "total capital contable",
    
    # Ecuación Contable Balance
    "total pasivo y patrimonio", "total pasivo y patrimonio neto", "pasivo y patrimonio total",
    "total pasivos y patrimonio",
    
    # Resultados (Subtotales intermedios)
    "utilidad bruta", "resultado bruto", "total ingresos brutos", "margen bruto",
    "utilidad operativa", "resultado operativo", "utilidad de operacion",
    "utilidad antes de impuestos", "resultado antes de impuestos", "utilidad antes de iue",
    "utilidad neta", "resultado neto", "utilidad del ejercicio", "resultado del ejercicio",
    "utilidad neta del ejercicio", "resultado neto del ejercicio", "ganancia neta", "perdida neta",
    "ebitda", "ebit", "ebt",
    
    # Totales generales de ingresos y costos
    "total costos", "total de costos", "total costo de ventas", "total gastos operativos",
    "total gastos", "total de gastos", "total ingresos", "total de ingresos", "total egresos"
]


def is_subtotal_or_total_row(text: str) -> bool:
    """Identifica si una fila corresponde a un total o subtotal que deba omitirse de la consolidación."""
    norm = normalize_text(text)
    if not norm:
        return True

    # 1. Filas que empiezan con 'total ' o contienen ' total' o son 'total'
    if norm.startswith("total ") or " total" in norm or norm == "total":
        return True

    # 2. Subtotales calculados en el Estado de Resultados
    subtotal_phrases = [
        "utilidad bruta", "resultado bruto", "margen bruto", "ganancia bruta",
        "utilidad de operacion", "utilidad operativa", "resultado operativo", "ganancia operativa", "ganancia de operacion",
        "utilidad antes de", "resultado antes de", "ganancia antes",
        "utilidad neta", "resultado neto", "utilidad del ejercicio", "resultado del ejercicio",
        "ganancia neta", "perdida neta", "resultado retenido", "ebitda", "ebit", "ebt"
    ]
    for sp in subtotal_phrases:
        if sp in norm:
            return True

    # 3. Coincidencia con palabras clave de subtotales
    for kw in SUBTOTAL_KEYWORDS:
        if norm == kw or norm == f"total {kw}":
            return True

    # 4. Casos donde la línea es exactamente un encabezado de categoría o sección
    exact_headers = [
        "activo", "activos", "activo corriente", "activos corrientes",
        "activo no corriente", "activos no corrientes",
        "pasivo", "pasivos", "pasivo corriente", "pasivos corrientes",
        "pasivo no corriente", "pasivos no corrientes",
        "patrimonio", "patrimonio neto", "pasivo y patrimonio",
        "estado de resultados", "estado de ganancias y perdidas",
        "ingresos", "costos", "gastos", "gastos de operacion",
        "datos complementarios", "datos complementarios y flujos", "flujos"
    ]
    if norm in exact_headers:
        return True

    return False


# =========================================================
# 3. DICCIONARIO DE MAPEO DINÁMICO (SINÓNIMOS EN BOLIVIA)
# =========================================================
# Mapeo estructurado: termino_normalizado -> (Cuenta_Estandar, Estado_Estandar, Categoria_Estandar)
SYNONYM_MAPPING: Dict[str, Tuple[str, str, str]] = {
    # Disponibilidades
    "disponibilidades": ("Efectivo y Equivalentes", "Balance General", "Activo Corriente"),
    "caja y bancos": ("Efectivo y Equivalentes", "Balance General", "Activo Corriente"),
    "caja": ("Efectivo y Equivalentes", "Balance General", "Activo Corriente"),
    "bancos": ("Efectivo y Equivalentes", "Balance General", "Activo Corriente"),
    "efectivo": ("Efectivo y Equivalentes", "Balance General", "Activo Corriente"),
    "efectivo y equivalentes": ("Efectivo y Equivalentes", "Balance General", "Activo Corriente"),
    "efectivo en caja y bancos": ("Efectivo y Equivalentes", "Balance General", "Activo Corriente"),
    "bancos moneda nacional": ("Efectivo y Equivalentes", "Balance General", "Activo Corriente"),
    "bancos moneda extranjera": ("Efectivo y Equivalentes", "Balance General", "Activo Corriente"),
    "caja chica": ("Efectivo y Equivalentes", "Balance General", "Activo Corriente"),

    # Inversiones temporales
    "inversiones temporales": ("Inversiones a Corto Plazo (ICP)", "Balance General", "Activo Corriente"),
    "inversiones a corto plazo": ("Inversiones a Corto Plazo (ICP)", "Balance General", "Activo Corriente"),
    "inversiones a corto plazo icp": ("Inversiones a Corto Plazo (ICP)", "Balance General", "Activo Corriente"),
    "valores negociables": ("Inversiones a Corto Plazo (ICP)", "Balance General", "Activo Corriente"),
    "depositos a plazo fijo cp": ("Inversiones a Corto Plazo (ICP)", "Balance General", "Activo Corriente"),
    "dpf corto plazo": ("Inversiones a Corto Plazo (ICP)", "Balance General", "Activo Corriente"),

    # Cuentas por cobrar comerciales
    "cuentas por cobrar": ("Cuentas por Cobrar Comerciales", "Balance General", "Activo Corriente"),
    "cuentas por cobrar comerciales": ("Cuentas por Cobrar Comerciales", "Balance General", "Activo Corriente"),
    "clientes": ("Cuentas por Cobrar Comerciales", "Balance General", "Activo Corriente"),
    "deudores por ventas": ("Cuentas por Cobrar Comerciales", "Balance General", "Activo Corriente"),
    "documentos por cobrar": ("Cuentas por Cobrar Comerciales", "Balance General", "Activo Corriente"),
    "documentos por cobrar comerciales": ("Cuentas por Cobrar Comerciales", "Balance General", "Activo Corriente"),
    "cartera de clientes": ("Cuentas por Cobrar Comerciales", "Balance General", "Activo Corriente"),
    "creditos comerciales": ("Cuentas por Cobrar Comerciales", "Balance General", "Activo Corriente"),

    # Otras cuentas por cobrar y diferidos CP
    "otras cuentas por cobrar": ("Otros Activos Corrientes", "Balance General", "Activo Corriente"),
    "gastos pagados por adelantado": ("Otros Activos Corrientes", "Balance General", "Activo Corriente"),
    "anticipos a proveedores": ("Otros Activos Corrientes", "Balance General", "Activo Corriente"),
    "credito fiscal": ("Otros Activos Corrientes", "Balance General", "Activo Corriente"),
    "credito fiscal iva": ("Otros Activos Corrientes", "Balance General", "Activo Corriente"),
    "impuestos por recuperar": ("Otros Activos Corrientes", "Balance General", "Activo Corriente"),
    "exigible cp": ("Otros Activos Corrientes", "Balance General", "Activo Corriente"),
    "otros deudores": ("Otros Activos Corrientes", "Balance General", "Activo Corriente"),
    "otros activos corrientes": ("Otros Activos Corrientes", "Balance General", "Activo Corriente"),

    # Inventarios / Existencias
    "inventarios": ("Inventarios", "Balance General", "Activo Corriente"),
    "inventario": ("Inventarios", "Balance General", "Activo Corriente"),
    "existencias": ("Inventarios", "Balance General", "Activo Corriente"),
    "inventario de mercaderia": ("Inventarios", "Balance General", "Activo Corriente"),
    "inventario de mercaderias": ("Inventarios", "Balance General", "Activo Corriente"),
    "materia prima": ("Inventarios", "Balance General", "Activo Corriente"),
    "materias primas": ("Inventarios", "Balance General", "Activo Corriente"),
    "productos terminados": ("Inventarios", "Balance General", "Activo Corriente"),
    "productos en proceso": ("Inventarios", "Balance General", "Activo Corriente"),
    "almacenes": ("Inventarios", "Balance General", "Activo Corriente"),
    "realizable": ("Inventarios", "Balance General", "Activo Corriente"),

    # Activos Fijos / Bienes de Uso
    "activos fijos": ("Propiedad Planta y Equipo Neto", "Balance General", "Activo No Corriente"),
    "activo fijo": ("Propiedad Planta y Equipo Neto", "Balance General", "Activo No Corriente"),
    "activo fijo neto": ("Propiedad Planta y Equipo Neto", "Balance General", "Activo No Corriente"),
    "bienes de uso": ("Propiedad Planta y Equipo Neto", "Balance General", "Activo No Corriente"),
    "propiedad planta y equipo": ("Propiedad Planta y Equipo Neto", "Balance General", "Activo No Corriente"),
    "propiedad planta y equipo neto": ("Propiedad Planta y Equipo Neto", "Balance General", "Activo No Corriente"),
    "inmuebles maquinaria y equipo": ("Propiedad Planta y Equipo Neto", "Balance General", "Activo No Corriente"),
    "maquinaria y equipo": ("Propiedad Planta y Equipo Neto", "Balance General", "Activo No Corriente"),
    "edificaciones y terrenos": ("Propiedad Planta y Equipo Neto", "Balance General", "Activo No Corriente"),
    "vehiculos y equipo de computacion": ("Propiedad Planta y Equipo Neto", "Balance General", "Activo No Corriente"),

    # Intangibles y Otros Activos LP
    "activos intangibles": ("Intangibles y Otros Activos LP", "Balance General", "Activo No Corriente"),
    "activos intangibles y cargos diferidos": ("Intangibles y Otros Activos LP", "Balance General", "Activo No Corriente"),
    "intangibles": ("Intangibles y Otros Activos LP", "Balance General", "Activo No Corriente"),
    "cargos diferidos": ("Intangibles y Otros Activos LP", "Balance General", "Activo No Corriente"),
    "otros activos": ("Intangibles y Otros Activos LP", "Balance General", "Activo No Corriente"),
    "otros activos no corrientes": ("Intangibles y Otros Activos LP", "Balance General", "Activo No Corriente"),
    "inversiones permanentes": ("Intangibles y Otros Activos LP", "Balance General", "Activo No Corriente"),
    "gastos de organizacion": ("Intangibles y Otros Activos LP", "Balance General", "Activo No Corriente"),

    # Pasivo Corriente: Cuentas por pagar
    "cuentas por pagar a proveedores": ("Cuentas por Pagar Comerciales", "Balance General", "Pasivo Corriente"),
    "cuentas por pagar comerciales": ("Cuentas por Pagar Comerciales", "Balance General", "Pasivo Corriente"),
    "proveedores": ("Cuentas por Pagar Comerciales", "Balance General", "Pasivo Corriente"),
    "cuentas por pagar": ("Cuentas por Pagar Comerciales", "Balance General", "Pasivo Corriente"),
    "documentos por pagar comerciales": ("Cuentas por Pagar Comerciales", "Balance General", "Pasivo Corriente"),
    "deudas comerciales": ("Cuentas por Pagar Comerciales", "Balance General", "Pasivo Corriente"),

    # Pasivo Corriente: Deuda Financiera Corto Plazo
    "obligaciones financieras c": ("Deuda Financiera Corto Plazo", "Balance General", "Pasivo Corriente"),
    "obligaciones financieras cp": ("Deuda Financiera Corto Plazo", "Balance General", "Pasivo Corriente"),
    "obligaciones bancarias a corto plazo": ("Deuda Financiera Corto Plazo", "Balance General", "Pasivo Corriente"),
    "obligaciones bancarias corto plazo": ("Deuda Financiera Corto Plazo", "Balance General", "Pasivo Corriente"),
    "obligaciones bancarias cp": ("Deuda Financiera Corto Plazo", "Balance General", "Pasivo Corriente"),
    "obligaciones bancarias": ("Deuda Financiera Corto Plazo", "Balance General", "Pasivo Corriente"),
    "deuda financiera corto plazo": ("Deuda Financiera Corto Plazo", "Balance General", "Pasivo Corriente"),
    "prestamos bancarios cp": ("Deuda Financiera Corto Plazo", "Balance General", "Pasivo Corriente"),
    "prestamos bancarios a corto plazo": ("Deuda Financiera Corto Plazo", "Balance General", "Pasivo Corriente"),
    "creditos bancarios cp": ("Deuda Financiera Corto Plazo", "Balance General", "Pasivo Corriente"),
    "sobregiros bancarios": ("Deuda Financiera Corto Plazo", "Balance General", "Pasivo Corriente"),

    # Pasivo Corriente: Fiscales y Laborales
    "deudas fiscales y sociales": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),
    "cargas fiscales y sociales por pagar": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),
    "cargas fiscales y sociales": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),
    "impuestos por pagar": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),
    "debito fiscal": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),
    "debito fiscal iva": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),
    "iue por pagar": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),
    "it por pagar": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),
    "sueldos y salarios por pagar": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),
    "cargas sociales por pagar": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),
    "afp por pagar": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),
    "retenciones por pagar": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),
    "pasivos laborales e impuestos": ("Pasivos Laborales e Impuestos", "Balance General", "Pasivo Corriente"),

    # Pasivo Corriente: Otros
    "otras cuentas por pagar": ("Otros Pasivos Corto Plazo", "Balance General", "Pasivo Corriente"),
    "otros pasivos corrientes": ("Otros Pasivos Corto Plazo", "Balance General", "Pasivo Corriente"),
    "otros pasivos corto plazo": ("Otros Pasivos Corto Plazo", "Balance General", "Pasivo Corriente"),
    "anticipos de clientes": ("Otros Pasivos Corto Plazo", "Balance General", "Pasivo Corriente"),
    "ingresos diferidos": ("Otros Pasivos Corto Plazo", "Balance General", "Pasivo Corriente"),

    # Pasivo No Corriente: Deuda Financiera Largo Plazo
    "obligaciones financieras nc": ("Deuda Financiera Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "obligaciones financieras lp": ("Deuda Financiera Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "obligaciones bancarias a largo plazo": ("Deuda Financiera Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "obligaciones bancarias largo plazo": ("Deuda Financiera Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "obligaciones bancarias lp": ("Deuda Financiera Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "deuda financiera largo plazo": ("Deuda Financiera Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "prestamos bancarios lp": ("Deuda Financiera Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "prestamos bancarios a largo plazo": ("Deuda Financiera Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "bonos por pagar": ("Deuda Financiera Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "hipotecas por pagar": ("Deuda Financiera Largo Plazo", "Balance General", "Pasivo No Corriente"),

    # Pasivo No Corriente: Otros
    "prevision para beneficios sociales e indemnizaciones": ("Otros Pasivos Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "prevision para beneficios sociales": ("Otros Pasivos Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "beneficios sociales e indemnizaciones": ("Otros Pasivos Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "prevision para indemnizacion": ("Otros Pasivos Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "prevision para indemnizaciones": ("Otros Pasivos Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "provision para indemnizacion": ("Otros Pasivos Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "beneficios sociales lp": ("Otros Pasivos Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "indemnizaciones": ("Otros Pasivos Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "provisiones laborales": ("Otros Pasivos Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "otros pasivos largo plazo": ("Otros Pasivos Largo Plazo", "Balance General", "Pasivo No Corriente"),
    "otros pasivos no corrientes": ("Otros Pasivos Largo Plazo", "Balance General", "Pasivo No Corriente"),

    # Patrimonio
    "capital pagado": ("Capital Social", "Balance General", "Patrimonio"),
    "capital social": ("Capital Social", "Balance General", "Patrimonio"),
    "capital emitido": ("Capital Social", "Balance General", "Patrimonio"),
    "cuotas de capital": ("Capital Social", "Balance General", "Patrimonio"),
    "capital asignado": ("Capital Social", "Balance General", "Patrimonio"),

    "ajuste capital": ("Otras Reservas y Ajustes", "Balance General", "Patrimonio"),
    "ajuste de capital": ("Otras Reservas y Ajustes", "Balance General", "Patrimonio"),
    "reserva por reavaluo tecnico": ("Otras Reservas y Ajustes", "Balance General", "Patrimonio"),
    "reserva de reavaluo tecnico": ("Otras Reservas y Ajustes", "Balance General", "Patrimonio"),
    "ajuste de reservas patrimoniales": ("Otras Reservas y Ajustes", "Balance General", "Patrimonio"),
    "otras reservas y ajustes": ("Otras Reservas y Ajustes", "Balance General", "Patrimonio"),
    "otras reservas": ("Otras Reservas y Ajustes", "Balance General", "Patrimonio"),

    "reserva legal": ("Reservas Legales", "Balance General", "Patrimonio"),
    "reservas legales": ("Reservas Legales", "Balance General", "Patrimonio"),

    "resultados acumulados": ("Utilidades Retenidas", "Balance General", "Patrimonio"),
    "utilidades retenidas": ("Utilidades Retenidas", "Balance General", "Patrimonio"),
    "resultados no asignados": ("Utilidades Retenidas", "Balance General", "Patrimonio"),
    "ganancias reservadas": ("Utilidades Retenidas", "Balance General", "Patrimonio"),
    "perdidas acumuladas": ("Utilidades Retenidas", "Balance General", "Patrimonio"),
    "superavit acumulado": ("Utilidades Retenidas", "Balance General", "Patrimonio"),

    # Estado de Resultados: Ingresos
    "ingresos de ventas": ("Ingresos por Ventas", "Estado de Resultados", "Ingresos"),
    "ingresos por ventas": ("Ingresos por Ventas", "Estado de Resultados", "Ingresos"),
    "ventas netas": ("Ingresos por Ventas", "Estado de Resultados", "Ingresos"),
    "ventas": ("Ingresos por Ventas", "Estado de Resultados", "Ingresos"),
    "facturacion neta": ("Ingresos por Ventas", "Estado de Resultados", "Ingresos"),
    "ingresos operacionales": ("Ingresos por Ventas", "Estado de Resultados", "Ingresos"),
    "ingresos por servicios": ("Ingresos por Ventas", "Estado de Resultados", "Ingresos"),

    # Estado de Resultados: Costos
    "costos de ventas": ("Costo de Ventas", "Estado de Resultados", "Costos"),
    "costo de ventas": ("Costo de Ventas", "Estado de Resultados", "Costos"),
    "costo de lo vendido": ("Costo de Ventas", "Estado de Resultados", "Costos"),
    "costo de mercaderias vendidas": ("Costo de Ventas", "Estado de Resultados", "Costos"),
    "costo de mercaderia vendida": ("Costo de Ventas", "Estado de Resultados", "Costos"),
    "costo de produccion": ("Costo de Ventas", "Estado de Resultados", "Costos"),
    "costo de servicios prestados": ("Costo de Ventas", "Estado de Resultados", "Costos"),

    # Estado de Resultados: Gastos Operativos
    "gastos de administracion": ("Gastos de Administracion y Ventas", "Estado de Resultados", "Gastos Operativos"),
    "gastos de comercializacion": ("Gastos de Administracion y Ventas", "Estado de Resultados", "Gastos Operativos"),
    "gastos de administracion y ventas": ("Gastos de Administracion y Ventas", "Estado de Resultados", "Gastos Operativos"),
    "gastos administrativos": ("Gastos de Administracion y Ventas", "Estado de Resultados", "Gastos Operativos"),
    "gastos comerciales": ("Gastos de Administracion y Ventas", "Estado de Resultados", "Gastos Operativos"),
    "gastos operativos": ("Gastos de Administracion y Ventas", "Estado de Resultados", "Gastos Operativos"),
    "gastos de venta": ("Gastos de Administracion y Ventas", "Estado de Resultados", "Gastos Operativos"),

    # Estado de Resultados: Depreciación y Amortización
    "depreciacion de bienes de uso y amortizacion": ("Depreciacion y Amortizacion", "Estado de Resultados", "Gastos Operativos"),
    "depreciacion de bienes de uso": ("Depreciacion y Amortizacion", "Estado de Resultados", "Gastos Operativos"),
    "depreciacion y amortizacion": ("Depreciacion y Amortizacion", "Estado de Resultados", "Gastos Operativos"),
    "depreciaciones y amortizaciones": ("Depreciacion y Amortizacion", "Estado de Resultados", "Gastos Operativos"),
    "depreciacion bienes de uso": ("Depreciacion y Amortizacion", "Estado de Resultados", "Gastos Operativos"),
    "depreciacion activo fijo": ("Depreciacion y Amortizacion", "Estado de Resultados", "Gastos Operativos"),
    "depreciacion del ejercicio": ("Depreciacion y Amortizacion", "Estado de Resultados", "Gastos Operativos"),
    "depreciacion": ("Depreciacion y Amortizacion", "Estado de Resultados", "Gastos Operativos"),
    "amortizacion gastos de organizacion": ("Depreciacion y Amortizacion", "Estado de Resultados", "Gastos Operativos"),
    "amortizacion de intangibles": ("Depreciacion y Amortizacion", "Estado de Resultados", "Gastos Operativos"),
    "amortizacion": ("Depreciacion y Amortizacion", "Estado de Resultados", "Gastos Operativos"),

    # Estado de Resultados: Financieros
    "gastos e ingresos financieros": ("Gastos Financieros (Intereses)", "Estado de Resultados", "Gastos Financieros"),
    "gastos financieros": ("Gastos Financieros (Intereses)", "Estado de Resultados", "Gastos Financieros"),
    "gastos financieros intereses": ("Gastos Financieros (Intereses)", "Estado de Resultados", "Gastos Financieros"),
    "intereses bancarios": ("Gastos Financieros (Intereses)", "Estado de Resultados", "Gastos Financieros"),
    "intereses pagados": ("Gastos Financieros (Intereses)", "Estado de Resultados", "Gastos Financieros"),
    "costo financiero": ("Gastos Financieros (Intereses)", "Estado de Resultados", "Gastos Financieros"),

    # Estado de Resultados: Otros y AITB
    "diferencia de cambio mant valor": ("Otros Ingresos (Gastos) No Operacionales", "Estado de Resultados", "Otros"),
    "diferencia de cambio y mantenimiento de valor": ("Otros Ingresos (Gastos) No Operacionales", "Estado de Resultados", "Otros"),
    "diferencia de cambio": ("Otros Ingresos (Gastos) No Operacionales", "Estado de Resultados", "Otros"),
    "mantenimiento de valor": ("Otros Ingresos (Gastos) No Operacionales", "Estado de Resultados", "Otros"),
    "aitb": ("Otros Ingresos (Gastos) No Operacionales", "Estado de Resultados", "Otros"),
    "ajuste por inflacion y tenencia de bienes": ("Otros Ingresos (Gastos) No Operacionales", "Estado de Resultados", "Otros"),
    "otros ingresos gastos no operacionales": ("Otros Ingresos (Gastos) No Operacionales", "Estado de Resultados", "Otros"),
    "otros egresos": ("Otros Ingresos (Gastos) No Operacionales", "Estado de Resultados", "Otros"),
    "otros ingresos": ("Otros Ingresos (Gastos) No Operacionales", "Estado de Resultados", "Otros"),
    "otros ingresos no operacionales": ("Otros Ingresos (Gastos) No Operacionales", "Estado de Resultados", "Otros"),
    "otros gastos no operacionales": ("Otros Ingresos (Gastos) No Operacionales", "Estado de Resultados", "Otros"),

    # Estado de Resultados: Impuestos
    "impuesto sobre las utilidades": ("Impuesto sobre la Renta", "Estado de Resultados", "Impuestos"),
    "impuesto sobre las utilidades de las empresas": ("Impuesto sobre la Renta", "Estado de Resultados", "Impuestos"),
    "impuesto sobre las utilidades de las empresas iue": ("Impuesto sobre la Renta", "Estado de Resultados", "Impuestos"),
    "impuesto sobre las utilidades iue": ("Impuesto sobre la Renta", "Estado de Resultados", "Impuestos"),
    "iue": ("Impuesto sobre la Renta", "Estado de Resultados", "Impuestos"),
    "impuesto sobre la renta": ("Impuesto sobre la Renta", "Estado de Resultados", "Impuestos"),
    "provision para iue": ("Impuesto sobre la Renta", "Estado de Resultados", "Impuestos"),
    "gasto por iue": ("Impuesto sobre la Renta", "Estado de Resultados", "Impuestos"),
    "impuesto a las utilidades": ("Impuesto sobre la Renta", "Estado de Resultados", "Impuestos"),

    # Datos Complementarios
    "inversiones en activo fijo capex": ("Inversiones en Activo Fijo (CAPEX)", "Datos Complementarios", "Inversiones"),
    "inversiones de capital en activo fijo capex": ("Inversiones en Activo Fijo (CAPEX)", "Datos Complementarios", "Inversiones"),
    "inversiones en activo fijo": ("Inversiones en Activo Fijo (CAPEX)", "Datos Complementarios", "Inversiones"),
    "capex": ("Inversiones en Activo Fijo (CAPEX)", "Datos Complementarios", "Inversiones"),
    "adquisicion de bienes de uso": ("Inversiones en Activo Fijo (CAPEX)", "Datos Complementarios", "Inversiones"),

    "compras del periodo": ("Compras del Periodo", "Datos Complementarios", "Operaciones"),
    "compras brutas de mercaderias e insumos": ("Compras del Periodo", "Datos Complementarios", "Operaciones"),
    "compras": ("Compras del Periodo", "Datos Complementarios", "Operaciones"),
    "compras netas": ("Compras del Periodo", "Datos Complementarios", "Operaciones"),

    "flujo de efectivo actividades operativas feao": ("Flujo de Efectivo Actividades Operativas (FEAO)", "Datos Complementarios", "Flujos"),
    "flujo de efectivo de actividades de operacion feao": ("Flujo de Efectivo Actividades Operativas (FEAO)", "Datos Complementarios", "Flujos"),
    "flujo de efectivo actividades operativas": ("Flujo de Efectivo Actividades Operativas (FEAO)", "Datos Complementarios", "Flujos"),
    "feao": ("Flujo de Efectivo Actividades Operativas (FEAO)", "Datos Complementarios", "Flujos"),

    "dividendos distribuidos a socios": ("Dividendos Pagados", "Datos Complementarios", "Financiamiento"),
    "dividendos pagados en efectivo": ("Dividendos Pagados", "Datos Complementarios", "Financiamiento"),
    "dividendos pagados": ("Dividendos Pagados", "Datos Complementarios", "Financiamiento"),
    "dividendos": ("Dividendos Pagados", "Datos Complementarios", "Financiamiento"),

    "tasa impositiva efectiva": ("Tasa Impositiva Efectiva (%)", "Datos Complementarios", "Parametros"),
    "tasa impositiva": ("Tasa Impositiva Efectiva (%)", "Datos Complementarios", "Parametros"),
    "tasa iue": ("Tasa Impositiva Efectiva (%)", "Datos Complementarios", "Parametros"),

    "ufv cierre de gestion": ("UFV Cierre de Gestion", "Datos Complementarios", "Normativa Bolivia"),
    "ufv cierre": ("UFV Cierre de Gestion", "Datos Complementarios", "Normativa Bolivia"),
    "ufv": ("UFV Cierre de Gestion", "Datos Complementarios", "Normativa Bolivia"),
    "ufv al cierre": ("UFV Cierre de Gestion", "Datos Complementarios", "Normativa Bolivia")
}


def map_bolivian_account(raw_name: str, section_state: str = "", section_cat: str = "") -> Tuple[str, str, str]:
    """
    Traduce una cuenta contable expresada en terminología boliviana a los nombres estandarizados
    de FinCorp Analytics, asignando automáticamente su Estado y Categoría oficiales.
    """
    norm = normalize_text(raw_name)
    if not norm:
        return (raw_name.strip(), section_state or "Balance General", section_cat or "Activo Corriente")

    # 0. Desambiguación dependiente de la categoría activa (Corto Plazo vs Largo Plazo)
    if norm in ["deudas financieras", "deuda financiera", "obligaciones financieras", "prestamos bancarios", "deudas financieras corto plazo", "deudas financieras largo plazo"]:
        if "no corriente" in (section_cat or "").lower() or "largo" in (section_cat or "").lower():
            return ("Deuda Financiera Largo Plazo", "Balance General", "Pasivo No Corriente")
        else:
            return ("Deuda Financiera Corto Plazo", "Balance General", "Pasivo Corriente")

    if norm in ["otras cuentas por cobrar", "cuentas por cobrar a relacionadas", "cuentas por cobrar relacionadas", "cuentas por cobrar con entel s a", "cuentas por cobrar con entel"]:
        if "no corriente" in (section_cat or "").lower() or "largo" in (section_cat or "").lower():
            return ("Intangibles y Otros Activos LP", "Balance General", "Activo No Corriente")
        elif "entel" in norm:
            return ("Cuentas por Cobrar Comerciales", "Balance General", "Activo Corriente")
        else:
            return ("Otros Activos Corrientes", "Balance General", "Activo Corriente")

    if norm in ["deudas con companias relacionadas", "deudas con relacionadas", "cuentas por pagar a relacionadas", "cuentas por pagar con entel s a", "cuentas por pagar con entel"]:
        if "no corriente" in (section_cat or "").lower() or "largo" in (section_cat or "").lower():
            return ("Otros Pasivos Largo Plazo", "Balance General", "Pasivo No Corriente")
        elif "entel" in norm:
            return ("Cuentas por Pagar Comerciales", "Balance General", "Pasivo Corriente")
        else:
            return ("Otros Pasivos Corto Plazo", "Balance General", "Pasivo Corriente")

    if norm in ["servicio de asistencia tecnica", "asistencia tecnica"]:
        return ("Gastos de Administracion y Ventas", "Estado de Resultados", "Gastos Operativos")

    if norm in ["intereses ganados", "otros egresos", "ganancia por enajenacion de activos", "ganancia por enajenacion de activo", "resultado en inversiones permanentes"]:
        return ("Otros Ingresos (Gastos) No Operacionales", "Estado de Resultados", "Otros")

    # 1. Búsqueda exacta en diccionario
    if norm in SYNONYM_MAPPING:
        return SYNONYM_MAPPING[norm]

    # 2. Búsqueda por subcadenas prioritarias ordenadas por longitud descendente
    sorted_keys = sorted(SYNONYM_MAPPING.keys(), key=len, reverse=True)

    # Si se conoce la sección contable, buscar primero coincidencias que pertenezcan a dicha sección
    if section_state:
        for syn_key in sorted_keys:
            if ((len(syn_key) >= 4 or syn_key in ["iue", "dpf", "aitb", "feao"]) and (syn_key in norm)):
                acc, est, cat = SYNONYM_MAPPING[syn_key]
                if est.lower() == section_state.lower():
                    return (acc, est, cat)

    # Búsqueda general en caso de que no coincida la sección o no esté especificada
    for syn_key in sorted_keys:
        if ((len(syn_key) >= 4 or syn_key in ["iue", "dpf", "aitb", "feao"]) and (syn_key in norm)):
            return SYNONYM_MAPPING[syn_key]

    # 3. Clasificación inteligente por contexto de sección si no hubo coincidencia
    fallback_state = section_state or ("Balance General" if any(w in norm for w in ["activo", "pasivo", "capital", "reserva", "deuda"]) else "Estado de Resultados")
    fallback_cat = section_cat or ("Activo Corriente" if fallback_state == "Balance General" else "Gastos Operativos")
    
    return (raw_name.strip(), fallback_state, fallback_cat)


# =========================================================
# 4. CARGA Y PROCESAMIENTO DE ARCHIVOS (CSV Y EXCEL)
# =========================================================
def load_default_data(filepath: str = "data/sample_financial_data.csv") -> pd.DataFrame:
    """Carga el dataset de muestra predeterminado aplicando la limpieza numérica."""
    df = pd.read_csv(filepath)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Limpiar columnas de periodos numéricos
    years = get_years_from_df(df)
    for y in years:
        df[y] = df[y].apply(clean_numeric_value)
        
    return df


def detect_years_in_row(row_values: List[Any]) -> List[Tuple[int, str]]:
    """Identifica qué posiciones de una fila contienen etiquetas de años (ej. 2021, 2022, 2023)."""
    detected = []
    for idx, v in enumerate(row_values):
        s = str(v).strip()
        # Buscar 4 dígitos correspondientes a un año (ej. 2020 a 2035)
        m = re.search(r"\b(20[1-3][0-9])\b", s)
        if m:
            detected.append((idx, m.group(1)))
    return detected


def parse_raw_dataframe(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    Analiza un DataFrame crudo (proveniente de CSV o Excel) sin asumir una estructura fija:
    - Encuentra la fila de cabecera con los años de gestión.
    - Soporta Balance General y Estado de Resultados en una sola hoja (uno debajo del otro).
    - Detecta y extrae la fila de UFV de cabecera si existe.
    - Filtra subtotales y totales.
    - Mapea cuentas mediante el diccionario de sinónimos bolivianos.
    - Agrupa y suma cuentas coincidentes.
    """
    # Eliminar filas y columnas completamente vacías
    raw_df = raw_df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    if raw_df.empty:
        raise ValueError("El archivo subido no contiene datos legibles.")

    # 1. Caso A: Si ya tiene las columnas estándar 'Estado', 'Categoria', 'Cuenta'
    col_names_norm = [normalize_text(c) for c in raw_df.columns]
    if "cuenta" in col_names_norm:
        cuenta_col = raw_df.columns[col_names_norm.index("cuenta")]
        estado_col = raw_df.columns[col_names_norm.index("estado")] if "estado" in col_names_norm else None
        cat_col = raw_df.columns[col_names_norm.index("categoria")] if "categoria" in col_names_norm else None
        
        # Identificar columnas de años
        meta_cols = {cuenta_col}
        if estado_col: meta_cols.add(estado_col)
        if cat_col: meta_cols.add(cat_col)
        year_cols = [c for c in raw_df.columns if c not in meta_cols and re.search(r"\b(20[1-3][0-9])\b", str(c))]
        if not year_cols:
            year_cols = [c for c in raw_df.columns if c not in meta_cols]
            
        rows_processed = []
        ufv_data = {}

        for _, row in raw_df.iterrows():
            c_name = str(row[cuenta_col]).strip()
            if not c_name or c_name.lower() in ["nan", "none"]:
                continue

            # Detectar si es fila UFV (incluyendo posible errata UFYBOB)
            if "ufv" in normalize_text(c_name) or "ufy" in normalize_text(c_name):
                for y in year_cols:
                    y_clean = re.search(r"\b(20[1-3][0-9])\b", str(y))
                    y_key = y_clean.group(1) if y_clean else str(y)
                    ufv_data[y_key] = clean_numeric_value(row[y])
                continue

            # Omitir subtotales
            if is_subtotal_or_total_row(c_name):
                continue

            est_in = str(row[estado_col]).strip() if estado_col else ""
            cat_in = str(row[cat_col]).strip() if cat_col else ""

            acc_name, est_out, cat_out = map_bolivian_account(c_name, est_in, cat_in)
            
            row_dict = {"Estado": est_out, "Categoria": cat_out, "Cuenta": acc_name}
            for y in year_cols:
                y_clean = re.search(r"\b(20[1-3][0-9])\b", str(y))
                y_key = y_clean.group(1) if y_clean else str(y)
                row_dict[y_key] = clean_numeric_value(row[y])
            rows_processed.append(row_dict)

        clean_df = pd.DataFrame(rows_processed)
        
        # Imbuir UFV si fue detectada y no existe fila previa
        if ufv_data and not clean_df.empty:
            years_found = [c for c in clean_df.columns if c not in ["Estado", "Categoria", "Cuenta"]]
            ufv_row = {"Estado": "Datos Complementarios", "Categoria": "Normativa Bolivia", "Cuenta": "UFV Cierre de Gestion"}
            for y in years_found:
                ufv_row[y] = ufv_data.get(y, 2.40858)
            clean_df = pd.concat([clean_df, pd.DataFrame([ufv_row])], ignore_index=True)

        if clean_df.empty:
            raise ValueError("No se encontraron cuentas financieras válidas después del filtrado.")

        # Consolidar cuentas repetidas sumando sus valores
        years_cols_clean = [c for c in clean_df.columns if c not in ["Estado", "Categoria", "Cuenta"]]
        clean_df = clean_df.groupby(["Estado", "Categoria", "Cuenta"], as_index=False)[years_cols_clean].sum()
        return clean_df

    # 2. Caso B: Archivo con formato de reporte contable tradicional (una sola hoja, encabezados, etc.)
    # Buscar qué fila contiene los años
    header_row_idx = None
    years_detected = []
    
    for r_idx in range(min(15, len(raw_df))):
        row_vals = list(raw_df.iloc[r_idx])
        y_info = detect_years_in_row(row_vals)
        if len(y_info) >= 2:  # Al menos 2 años detectados
            header_row_idx = r_idx
            years_detected = y_info
            break

    if header_row_idx is None:
        # Fallback: tomar fila 0 y buscar años en sus columnas
        years_detected = detect_years_in_row(list(raw_df.columns))
        if len(years_detected) >= 1:
            header_row_idx = -1  # Columnas directas
        else:
            raise ValueError("No se pudieron detectar columnas de años válidas (ej. 2021, 2022, 2023) en el archivo.")

    # Extraer nombres de años y columnas
    year_map = {}  # col_idx -> year_str
    for c_idx, y_str in years_detected:
        year_map[c_idx] = y_str
    ordered_years = [y_str for _, y_str in years_detected]

    # Asumir que la primera columna de texto no numérica a la izquierda es la columna de Cuenta
    account_col_idx = 0
    for c_i in range(raw_df.shape[1]):
        if c_i not in year_map:
            account_col_idx = c_i
            break

    # Rastrear contexto de sección actual
    current_state = "Balance General"
    current_cat = "Activo Corriente"
    rows_parsed = []
    ufv_data = {}

    data_start_row = header_row_idx + 1 if header_row_idx >= 0 else 0

    for r_i in range(data_start_row, len(raw_df)):
        raw_cuenta = str(raw_df.iloc[r_i, account_col_idx]).strip()
        if not raw_cuenta or raw_cuenta.lower() in ["nan", "none", "0"]:
            continue

        norm_line = normalize_text(raw_cuenta)

        # A. Detectar UFV en cabecera o fila (incluyendo posible errata UFYBOB)
        if "ufv" in norm_line or "ufy" in norm_line:
            for c_i, y_str in year_map.items():
                val = clean_numeric_value(raw_df.iloc[r_i, c_i])
                if val > 0:
                    ufv_data[y_str] = val
            continue

        # Verificar si la fila tiene montos numéricos en los años
        has_numeric_data = False
        row_vals = {}
        for c_i, y_str in year_map.items():
            val = clean_numeric_value(raw_df.iloc[r_i, c_i])
            row_vals[y_str] = val
            if val != 0.0:
                has_numeric_data = True

        # B. Si la fila NO contiene datos numéricos en los periodos, es un encabezado o separador de sección
        if not has_numeric_data:
            if any(k in norm_line for k in ["estado de resultados", "ganancias y perdidas"]):
                current_state = "Estado de Resultados"
                current_cat = "Ingresos"
            elif any(k in norm_line for k in ["balance general", "situacion financiera"]):
                current_state = "Balance General"
                current_cat = "Activo Corriente"
            elif any(k in norm_line for k in ["activo no corriente", "activos no corrientes"]):
                current_state = "Balance General"
                current_cat = "Activo No Corriente"
            elif any(k in norm_line for k in ["activo corriente", "activos corrientes", "activo circulante", "activo"]):
                current_state = "Balance General"
                current_cat = "Activo Corriente"
            elif any(k in norm_line for k in ["pasivo no corriente", "pasivos no corrientes"]):
                current_state = "Balance General"
                current_cat = "Pasivo No Corriente"
            elif any(k in norm_line for k in ["pasivo corriente", "pasivos corrientes", "pasivo circulante", "pasivo"]):
                current_state = "Balance General"
                current_cat = "Pasivo Corriente"
            elif any(k in norm_line for k in ["patrimonio", "capital contable"]):
                current_state = "Balance General"
                current_cat = "Patrimonio"
            elif any(k in norm_line for k in ["costo de ventas", "costos"]):
                current_state = "Estado de Resultados"
                current_cat = "Costos"
            elif any(k in norm_line for k in ["gastos de operacion", "gastos operativos", "gastos de administracion"]):
                current_state = "Estado de Resultados"
                current_cat = "Gastos Operativos"
            elif any(k in norm_line for k in ["datos complementarios", "flujos"]):
                current_state = "Datos Complementarios"
                current_cat = "Operaciones"
            continue

        # C. Omitir subtotales y totales calculados, actualizando la sección contable activa
        if is_subtotal_or_total_row(raw_cuenta):
            if any(k in norm_line for k in ["activo corriente total", "total activo corriente"]):
                current_state = "Balance General"
                current_cat = "Activo No Corriente"
            elif any(k in norm_line for k in ["activo no corriente total", "total activo no corriente", "activo total", "total activo"]):
                current_state = "Balance General"
                current_cat = "Pasivo Corriente"
            elif any(k in norm_line for k in ["pasivo corriente total", "total pasivo corriente"]):
                current_state = "Balance General"
                current_cat = "Pasivo No Corriente"
            elif any(k in norm_line for k in ["pasivo no corriente total", "total pasivo no corriente", "total pasivo", "pasivo total"]):
                current_state = "Balance General"
                current_cat = "Patrimonio"
            elif any(k in norm_line for k in ["patrimonio neto total", "total patrimonio", "pasivo y patrimonio"]):
                current_state = "Estado de Resultados"
                current_cat = "Ingresos"
            elif any(k in norm_line for k in ["ganancia bruta", "utilidad bruta", "resultado bruto"]):
                current_state = "Estado de Resultados"
                current_cat = "Gastos Operativos"
            elif any(k in norm_line for k in ["ganancia operativa", "utilidad operativa", "resultado operativo", "ganancia de operacion"]):
                current_state = "Estado de Resultados"
                current_cat = "Gastos Financieros"
            elif any(k in norm_line for k in ["ganancia antes", "utilidad antes", "resultado antes"]):
                current_state = "Estado de Resultados"
                current_cat = "Impuestos"
            elif any(k in norm_line for k in ["ganancia neta", "utilidad neta", "resultado neto"]):
                current_state = "Datos Complementarios"
                current_cat = "Financiamiento"
            continue

        # D. Mapear cuenta mediante sinónimos bolivianos
        canon_name, canon_state, canon_cat = map_bolivian_account(raw_cuenta, current_state, current_cat)

        # Si toda la fila de años es 0.0 y no es una cuenta requerida, omitir
        if has_numeric_data:
            row_dict = {"Estado": canon_state, "Categoria": canon_cat, "Cuenta": canon_name}
            for y_str in ordered_years:
                row_dict[y_str] = row_vals.get(y_str, 0.0)
            rows_parsed.append(row_dict)

    final_df = pd.DataFrame(rows_parsed)
    if final_df.empty:
        raise ValueError("No fue posible extraer cuentas contables estructuradas del archivo.")

    # Imbuir UFV si fue detectada
    if ufv_data:
        ufv_row = {"Estado": "Datos Complementarios", "Categoria": "Normativa Bolivia", "Cuenta": "UFV Cierre de Gestion"}
        for y_str in ordered_years:
            ufv_row[y_str] = ufv_data.get(y_str, 2.40858)
        final_df = pd.concat([final_df, pd.DataFrame([ufv_row])], ignore_index=True)

    # Consolidar cuentas repetidas (ej. múltiples cuentas de inventario sumadas)
    final_df = final_df.groupby(["Estado", "Categoria", "Cuenta"], as_index=False)[ordered_years].sum()
    return final_df


def parse_uploaded_file(uploaded_file) -> pd.DataFrame:
    """
    Parsea de forma 100% flexible un archivo subido por el usuario (.xlsx, .xls o .csv):
    - Soporta hojas múltiples o una sola hoja unificada con Balance General y Estado de Resultados apilados.
    - Realiza limpieza exhaustiva de formatos numéricos bolivianos (puntos de miles, comas, negativos, guiones).
    - Elimina subtotales y totales para evitar doble conteo.
    - Traduce la nomenclatura a las cuentas estándar de FinCorp Analytics.
    """
    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):
        # Detectar encoding y delimitador
        try:
            raw_df = pd.read_csv(uploaded_file, encoding="utf-8")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            raw_df = pd.read_csv(uploaded_file, encoding="latin1")
        return parse_raw_dataframe(raw_df)

    elif filename.endswith((".xlsx", ".xls")):
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names

        if len(sheet_names) == 1:
            raw_df = pd.read_excel(uploaded_file, sheet_name=0)
            return parse_raw_dataframe(raw_df)
        else:
            # Archivo con múltiples hojas (ej. Balance, Resultados, etc.)
            complete_sheets = []
            partial_sheets = []

            for sheet in sheet_names:
                s_clean = normalize_text(sheet)
                try:
                    s_df = pd.read_excel(uploaded_file, sheet_name=sheet)
                except Exception:
                    continue
                if s_df.empty:
                    continue

                # Si no tiene columna Estado, sugerir por el nombre de la hoja
                if "Estado" not in s_df.columns:
                    if "balance" in s_clean or "situacion" in s_clean:
                        s_df["Estado"] = "Balance General"
                    elif "resultado" in s_clean or "perdida" in s_clean:
                        s_df["Estado"] = "Estado de Resultados"
                    else:
                        s_df["Estado"] = "Datos Complementarios"

                try:
                    cleaned_sheet = parse_raw_dataframe(s_df)
                    unique_states = set(cleaned_sheet["Estado"].unique())
                    if "Balance General" in unique_states and "Estado de Resultados" in unique_states:
                        complete_sheets.append((sheet, cleaned_sheet))
                    else:
                        partial_sheets.append((sheet, cleaned_sheet))
                except Exception:
                    pass

            # Si una o más hojas ya contienen tanto Balance como Resultados, priorizar la principal
            if complete_sheets:
                for s_name, s_clean_df in complete_sheets:
                    s_norm = normalize_text(s_name)
                    if any(k in s_norm for k in ["estados", "financieros", "eeff", "datos"]):
                        return s_clean_df
                # Retornar la primera hoja completa
                return complete_sheets[0][1]

            # Si las hojas son parciales (ej. Hoja 1 Balance, Hoja 2 Resultados), combinarlas
            if partial_sheets:
                combined_df = pd.concat([df for _, df in partial_sheets], ignore_index=True)
                years = get_years_from_df(combined_df)
                combined_df = combined_df.groupby(["Estado", "Categoria", "Cuenta"], as_index=False)[years].sum()
                return combined_df

            # Fallback
            raw_df = pd.read_excel(uploaded_file, sheet_name=0)
            return parse_raw_dataframe(raw_df)
    else:
        raise ValueError("Formato de archivo no soportado. Por favor suba un archivo Excel (.xlsx, .xls) o CSV (.csv).")


def get_years_from_df(df: pd.DataFrame) -> List[str]:
    """Obtiene y ordena cronológicamente las columnas que representan años fiscales."""
    if df is None or df.empty:
        return []
    meta_cols = {"estado", "categoria", "cuenta", "subcategoria", "descripcion", "id", "item"}
    year_cols = [str(c).strip() for c in df.columns if str(c).strip().lower() not in meta_cols]

    def sort_key(x):
        try:
            m = re.search(r"\b(20[1-3][0-9])\b", str(x))
            return (0, int(m.group(1)) if m else int(x))
        except ValueError:
            return (1, str(x))

    return sorted(year_cols, key=sort_key)


def get_account_series(df: pd.DataFrame, account_name: str) -> pd.Series:
    """
    Busca una cuenta en el DataFrame por coincidencia exacta o de subcadena y devuelve sus valores por año.
    Garantiza que siempre retorne una serie de números flotantes válidos.
    """
    if df is None or df.empty:
        return pd.Series(dtype=float)
    years = get_years_from_df(df)
    if "Cuenta" not in df.columns:
        return pd.Series(0.0, index=years, dtype=float)

    acc_clean = account_name.strip().lower()

    # 1. Búsqueda exacta primero
    exact_match = df[df["Cuenta"].str.strip().str.lower() == acc_clean]
    if not exact_match.empty:
        return exact_match[years].iloc[0].astype(float)

    # 2. Búsqueda por contención
    contains_match = df[df["Cuenta"].str.strip().str.lower().str.contains(acc_clean, regex=False)]
    if not contains_match.empty:
        return contains_match[years].iloc[0].astype(float)

    # 3. Búsqueda a través del diccionario de sinónimos
    norm_target = normalize_text(account_name)
    if norm_target in SYNONYM_MAPPING:
        canon_name, _, _ = SYNONYM_MAPPING[norm_target]
        canon_match = df[df["Cuenta"].str.strip().str.lower() == canon_name.lower()]
        if not canon_match.empty:
            return canon_match[years].iloc[0].astype(float)

    # Si no se encuentra, retornar ceros con el índice de años
    return pd.Series(0.0, index=years, dtype=float)


def validate_accounting_equation(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Valida la ecuación fundamental contable para cada gestión:
    Total Activos == Total Pasivos + Total Patrimonio
    """
    if df is None or df.empty or "Estado" not in df.columns:
        return {"is_balanced": False, "details": {}}

    from modules.financial_math import get_balance_totals

    years = get_years_from_df(df)
    totals = get_balance_totals(df)

    results = {}
    is_balanced = True

    for y in years:
        activos = totals["Activos"].get(y, 0.0)
        pasivos = totals["Pasivos"].get(y, 0.0)
        patrimonio = totals["Patrimonio"].get(y, 0.0)
        pasivo_mas_patrimonio = pasivos + patrimonio
        diff = abs(activos - pasivo_mas_patrimonio)

        status = diff < 1.0  # Tolerancia de 1.0 por redondeo en centavos
        if not status:
            is_balanced = False

        results[y] = {
            "Total Activos": activos,
            "Total Pasivos": pasivos,
            "Total Patrimonio": patrimonio,
            "Pasivo + Patrimonio": pasivo_mas_patrimonio,
            "Diferencia": diff,
            "Cuadrado": status
        }

    return {
        "is_balanced": is_balanced,
        "details": results
    }


def generate_template_csv() -> str:
    """Genera la plantilla CSV estándar oficial con nomenclaturas bolivianas."""
    content = """Estado,Categoria,Cuenta,2022,2023,2024
Balance General,Activo Corriente,Disponibilidades / Caja y Bancos,45000,38000,52000
Balance General,Activo Corriente,Inversiones Temporales (ICP),0,0,0
Balance General,Activo Corriente,Cuentas por Cobrar Comerciales,85000,110000,135000
Balance General,Activo Corriente,Inventarios (Existencias),95000,130000,155000
Balance General,Activo Corriente,Otras Cuentas por Cobrar y Gastos Anticipados,15000,17000,18000
Balance General,Activo No Corriente,Activos Fijos / Bienes de Uso Neto,320000,365000,410000
Balance General,Activo No Corriente,Activos Intangibles y Otros Activos LP,40000,40000,50000
Balance General,Pasivo Corriente,Cuentas por Pagar Comerciales (Proveedores),65000,85000,105000
Balance General,Pasivo Corriente,Obligaciones Financieras CP (Deuda CP),40000,50000,45000
Balance General,Pasivo Corriente,Deudas Fiscales y Sociales (Impuestos/Sueldos),25000,30000,35000
Balance General,Pasivo No Corriente,Obligaciones Financieras LP (Deuda LP),180000,205000,235000
Balance General,Pasivo No Corriente,Prevision para Indemnizacion y Otros LP,20000,20000,25000
Balance General,Patrimonio,Capital Social / Capital Pagado,150000,150000,150000
Balance General,Patrimonio,Reservas Legales,30000,38000,48000
Balance General,Patrimonio,Resultados Acumulados / Utilidades Retenidas,90000,122000,177000
Estado de Resultados,Ingresos,Ingresos de Ventas / Ventas Netas,620000,740000,890000
Estado de Resultados,Costos,Costo de Ventas / Costo de lo Vendido,390000,460000,545000
Estado de Resultados,Gastos Operativos,Gastos de Administracion y Comercializacion,110000,132000,155000
Estado de Resultados,Gastos Operativos,Depreciacion y Amortizacion,28000,33000,38000
Estado de Resultados,Gastos Financieros,Gastos Financieros (Intereses),22000,26000,29000
Estado de Resultados,Otros,Diferencia de Cambio / Mant. Valor / Otros,0,1000,2000
Estado de Resultados,Impuestos,Impuesto sobre las Utilidades (IUE),17500,22500,31250
Datos Complementarios,Inversiones,Inversiones en Activo Fijo (CAPEX),45000,78000,83000
Datos Complementarios,Operaciones,Compras del Periodo,390000,495000,570000
Datos Complementarios,Flujos,Flujo de Efectivo Actividades Operativas (FEAO),80500,60500,101750
Datos Complementarios,Financiamiento,Dividendos Pagados,15000,20000,25000
Datos Complementarios,Parametros,Tasa Impositiva Efectiva (%),25.0,25.0,25.0
Datos Complementarios,Normativa Bolivia,UFV Cierre de Gestion,2.40858,2.47895,2.54520
"""
    return content


def generate_template_excel() -> bytes:
    """Genera la plantilla Excel descargable con nomenclatura boliviana profesional."""
    filepath = "estados_financieros_modelo_bolivia.xlsx"
    if os.path.exists(filepath):
        with open(filepath, "rb") as f:
            return f.read()
    output = io.BytesIO()
    df = load_default_data()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Estados_Financieros")
    return output.getvalue()
