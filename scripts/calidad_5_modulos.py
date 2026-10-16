"""
ANÁLISIS MODULAR DE CALIDAD DE DATOS
====================================

Diseñado para Jupyter Notebook.

Módulos disponibles:
- Cirugías
- MOCE
- Urgencias
- Egresos UCI
- Egresos

Flujo:
1. Elegir módulo.
2. Seleccionar archivo CSV / Excel / Parquet.
3. Ejecutar reglas del módulo.
4. Elegir dónde guardar el Excel.
5. Generar:
   - Hoja "Reglas": resumen puntual por regla.
   - Hoja "Resultados": únicamente incumplimientos, con motivo.

IMPORTANTE
----------
- "NO APLICA" no entra en el denominador de una regla.
- Los resultados detallados contienen sólo "NO CUMPLE".
- El análisis de CURP es estructural; no consulta RENAPO.
- Para sexo se asume:
      dato M = masculino -> CURP H
      dato F = femenino  -> CURP M
- Edad contra CURP se calcula al ingreso usando años cumplidos.
"""

from __future__ import annotations

from pathlib import Path
from datetime import datetime, date
from collections import Counter
import csv
import json
import re
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd

# ============================================================================
# CONFIGURACIÓN GENERAL
# ============================================================================

TOLERANCIA_EDAD_SEMANAS_ANOS = 1.0

CURP_RE = re.compile(
    r"^[A-Z][AEIOUX][A-Z]{2}\d{6}[HM][A-Z]{5}[A-Z0-9]\d$",
    re.IGNORECASE,
)

CLUES_RE = re.compile(r"^..IMB......$", re.IGNORECASE)


# ============================================================================
# DEFINICIÓN DE MÓDULOS Y VARIABLES
# ============================================================================

MODULOS = {
    "Cirugías": {
        "codigo": "CIR",
        "variables": {
            "CVE_REGISTRO_CIR": ["CVE_REGISTRO_CIR"],
            "CVE_REGISTRO_URG": ["CVE_REGISTRO_URG"],
            "CVE_REGISTRO_HOS": ["CVE_REGISTRO_HOS"],
            "CLUES": ["CLUES"],
            "Fecha de nacimiento": ["Fecha de nacimiento", "FECHA DE NACIMIENTO"],
            "PAC_EDAD_ANO": ["PAC_EDAD_ANO"],
            "PAC_EDAD_SEMANA": ["PAC_EDAD_SEMANA"],
            "CURP": ["curp"],
            "Sexo": ["Sexo", "SEXO"],
            "Fecha solicitud": [
                "Fecha solicitud",
                "FECHA SOLICITUD",
                "FECHA DE SOLICITUD",
            ],
            "CIE10 - Diagnóstico principal de la solicitud": [
                "CIE10 - Diagnóstico principal de la solicitud",
                "CIE 10 - Diagnóstico principal de la solicitud",
                "CIE 10 – Diagnóstico principal de la solicitud",
            ],
            "Procedimiento principal de la solicitud": [
                "Procedimiento principal de la solicitud"
            ],
            "Fecha de programación": ["Fecha de programación", "FECHA DE PROGRAMACIÓN"],
            "Estatus de la solicitud": [
                "Estatus de la solicitud",
                "ESTATUS DE LA SOLICITUD",
            ],
            "Hora de entrada a sala": [
                "Hora de entrada a sala",
                "Hora de entrada a la sala",
            ],
            "Hora de inicio de cirugía": ["Hora de inicio de cirugía"],
            "Hora de término de cirugía": ["Hora de término de cirugía"],
            "Hora de salida de sala": ["Hora de salida de sala"],
            "Fecha de realización": ["Fecha de realización", "Fecha de Realización"],
        },
    },
    "MOCE": {
        "codigo": "MOC",
        "variables": {
            "CLUES": ["CLUES"],
            "FECHA_ATENCION": ["FECHA_ATENCION", "FECHA ATENCION", "FECHA_ATENCIÓN"],
            "FEC_NACIMIENTO": ["FEC_NACIMIENTO", "FECHA_NACIMIENTO"],
            "CVE_CURP": ["CVE_CURP", "CURP", "curp"],
            "PACIENTE_SEXO": ["PACIENTE_SEXO", "SEXO"],
            "PESO": ["PESO"],
            "TALLA": ["TALLA"],
            "PRESION_SISTOLICA": ["PRESION_SISTOLICA", "PRESIÓN_SISTÓLICA"],
            "PRESION_DIASISTOLICA": ["PRESION_DIASISTOLICA", "PRESIÓN_DIASTÓLICA"],
            "FRECUENCIA_CARDIACA": ["FRECUENCIA_CARDIACA"],
            "FRECUENCIA_RESPIRATORIA": ["FRECUENCIA_RESPIRATORIA"],
            "TEMPERATURA": ["TEMPERATURA"],
            "10_PRINCIPAL": ["10_PRINCIPAL", "CIE10_PRINCIPAL"],
        },
    },
    "Urgencias": {
        "codigo": "URG",
        "variables": {
            "CVE_REGISTRO_URG": ["CVE_REGISTRO_URG"],
            "CLUES": ["CLUES"],
            "FECHA_HORA_INICIO_TRIAGE": ["FECHA_HORA_INICIO_TRIAGE"],
            "FECHA_HORA_TERMINO_TRIAGE": ["FECHA_HORA_TERMINO_TRIAGE"],
            "FECHA_HORA_INGRESO": ["FECHA_HORA_INGRESO"],
            "DX_INGRESO": ["DX_INGRESO"],
            "FECHA_HORA_EGRESO": ["FECHA_HORA_EGRESO"],
            "DX_PRIN_EGRESO": ["DX_PRIN_EGRESO"],
        },
    },
    "Egresos UCI": {
        "codigo": "UCI",
        "variables": {
            "CVE_REGISTRO_UCI": ["CVE_REGISTRO_UCI"],
            "CVE_REGISTRO_HOS": ["CVE_REGISTRO_HOS"],
            "CLUES": ["CLUES"],
            "FECHA_HORA_INGRESO": ["FECHA_HORA_INGRESO"],
            "FECHA_HORA_INGRESO_UCI": ["FECHA_HORA_INGRESO_UCI"],
            "CVE_ESP_INGRESO": ["CVE_ESP_INGRESO"],
            "DESC_ESP_INGRESO": ["DESC_ESP_INGRESO"],
            "DX_INGRESO": ["DX_INGRESO"],
            "DX_PRIN_EGRESO": ["DX_PRIN_EGRESO"],
            "FECHA_HORA_EGRESO": ["FECHA_HORA_EGRESO"],
            "PAC_SEXO": ["PAC_SEXO"],
            "PAC_EDAD_ANO": ["PAC_EDAD_ANO"],
            "PAC_CURP": ["PAC_CURP", "CURP"],
        },
    },
    "Egresos": {
        "codigo": "EGR",
        "variables": {
            "CVE_REGISTRO_HOS": ["CVE_REGISTRO_HOS"],
            "CVE_REGISTRO_URG": ["CVE_REGISTRO_URG"],
            "CLUES": ["CLUES"],
            "FECHA_HORA_INGRESO": ["FECHA_HORA_INGRESO"],
            "PROCEDENCIA": ["PROCEDENCIA"],
            "DESC_ESP_INGRESO": ["DESC_ESP_INGRESO"],
            "DX_INGRESO": ["DX_INGRESO"],
            "FECHA_HORA_EGRESO": ["FECHA_HORA_EGRESO"],
            "DX_PRIN_EGRESO": ["DX_PRIN_EGRESO"],
            "PAC_FECHANAC": ["PAC_FECHANAC", "PAC_FECHA_NAC"],
            "PAC_SEXO": ["PAC_SEXO"],
            "PAC_EDAD_ANO": ["PAC_EDAD_ANO"],
            "PAC_EDAD_SEMANA": ["PAC_EDAD_SEMANA"],
            "PAC_CURP": ["PAC_CURP", "CURP"],
        },
    },
}


# ============================================================================
# CATÁLOGO DE REGLAS
# ============================================================================

REGLAS = {
    "Cirugías": [
        {
            "id": "CIR-01",
            "variable": "CVE_REGISTRO_CIR",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "CIR-02",
            "variable": "CVE_REGISTRO_URG",
            "regla": "Nulo permitido; si existe se acepta como identificador",
            "tipo": "optional",
        },
        {
            "id": "CIR-03",
            "variable": "CVE_REGISTRO_HOS",
            "regla": "Nulo permitido; si existe se acepta como identificador",
            "tipo": "optional",
        },
        {
            "id": "CIR-04",
            "variable": "CLUES",
            "regla": "No nula, 11 caracteres e IMB en posiciones 3-5",
            "tipo": "clues",
        },
        {
            "id": "CIR-05",
            "variable": "Fecha de nacimiento",
            "regla": "No nula y fecha válida",
            "tipo": "required_date",
        },
        {
            "id": "CIR-06",
            "variable": "Fecha de nacimiento",
            "regla": "Debe coincidir con YYMMDD de CURP",
            "tipo": "birth_curp",
            "curp": "CURP",
        },
        {
            "id": "CIR-07",
            "variable": "PAC_EDAD_ANO",
            "regla": "No puede ser nulo y debe ser numérico",
            "tipo": "required_number",
        },
        {
            "id": "CIR-08",
            "variable": "PAC_EDAD_SEMANA",
            "regla": "No puede ser nulo y debe ser numérico",
            "tipo": "required_number",
        },
        {
            "id": "CIR-09",
            "variable": "PAC_EDAD_ANO",
            "regla": "Debe ser consistente con PAC_EDAD_SEMANA",
            "tipo": "age_weeks",
            "weeks": "PAC_EDAD_SEMANA",
        },
        {
            "id": "CIR-10",
            "variable": "CURP",
            "regla": "No nula y estructura básica de CURP válida",
            "tipo": "curp",
        },
        {
            "id": "CIR-11",
            "variable": "Sexo",
            "regla": "No nulo y debe coincidir con sexo de CURP",
            "tipo": "sex_curp",
            "curp": "CURP",
        },
        {
            "id": "CIR-12",
            "variable": "Fecha solicitud",
            "regla": "No nula y fecha válida",
            "tipo": "required_date",
        },
        {
            "id": "CIR-13",
            "variable": "CIE10 - Diagnóstico principal de la solicitud",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "CIR-14",
            "variable": "Procedimiento principal de la solicitud",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "CIR-15",
            "variable": "Fecha de programación",
            "regla": "Nulo permitido; si existe debe ser fecha válida",
            "tipo": "optional_date",
        },
        {
            "id": "CIR-16",
            "variable": "Fecha de programación",
            "regla": "Si existe no puede ser anterior a Fecha solicitud",
            "tipo": "date_gte",
            "other": "Fecha solicitud",
        },
        {
            "id": "CIR-17",
            "variable": "Estatus de la solicitud",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "CIR-18",
            "variable": "Hora de entrada a sala",
            "regla": "Obligatoria y válida cuando Estatus=Realizada",
            "tipo": "required_if_realizada",
        },
        {
            "id": "CIR-19",
            "variable": "Hora de inicio de cirugía",
            "regla": "Obligatoria y válida cuando Estatus=Realizada",
            "tipo": "required_if_realizada",
        },
        {
            "id": "CIR-20",
            "variable": "Hora de término de cirugía",
            "regla": "Obligatoria y válida cuando Estatus=Realizada",
            "tipo": "required_if_realizada",
        },
        {
            "id": "CIR-21",
            "variable": "Hora de salida de sala",
            "regla": "Obligatoria y válida cuando Estatus=Realizada",
            "tipo": "required_if_realizada",
        },
        {
            "id": "CIR-22",
            "variable": "Fecha de realización",
            "regla": "Obligatoria y válida cuando Estatus=Realizada",
            "tipo": "required_if_realizada",
        },
        {
            "id": "CIR-23",
            "variable": "Horas de sala",
            "regla": "Entrada <= Inicio <= Término <= Salida",
            "tipo": "sequence",
            "fields": [
                "Hora de entrada a sala",
                "Hora de inicio de cirugía",
                "Hora de término de cirugía",
                "Hora de salida de sala",
            ],
        },
    ],
    "MOCE": [
        {
            "id": "MOC-01",
            "variable": "CLUES",
            "regla": "No nula, 11 caracteres e IMB en posiciones 3-5",
            "tipo": "clues",
        },
        {
            "id": "MOC-02",
            "variable": "FECHA_ATENCION",
            "regla": "No puede ser nula y debe ser fecha válida",
            "tipo": "required_date",
        },
        {
            "id": "MOC-03",
            "variable": "FEC_NACIMIENTO",
            "regla": "No puede ser nula y debe ser fecha válida",
            "tipo": "required_date",
        },
        {
            "id": "MOC-04",
            "variable": "FEC_NACIMIENTO",
            "regla": "Debe coincidir con YYMMDD de CVE_CURP",
            "tipo": "birth_curp",
            "curp": "CVE_CURP",
        },
        {
            "id": "MOC-05",
            "variable": "CVE_CURP",
            "regla": "No nula y estructura básica de CURP válida",
            "tipo": "curp",
        },
        {
            "id": "MOC-06",
            "variable": "PACIENTE_SEXO",
            "regla": "No nulo y debe coincidir con sexo de CURP",
            "tipo": "sex_curp",
            "curp": "CVE_CURP",
        },
        {
            "id": "MOC-07",
            "variable": "PESO",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "MOC-08",
            "variable": "TALLA",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "MOC-09",
            "variable": "PRESION_SISTOLICA",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "MOC-10",
            "variable": "PRESION_DIASISTOLICA",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "MOC-11",
            "variable": "FRECUENCIA_CARDIACA",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "MOC-12",
            "variable": "FRECUENCIA_RESPIRATORIA",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "MOC-13",
            "variable": "TEMPERATURA",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "MOC-14",
            "variable": "10_PRINCIPAL",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
    ],
    "Urgencias": [
        {
            "id": "URG-01",
            "variable": "CVE_REGISTRO_URG",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "URG-02",
            "variable": "CLUES",
            "regla": "No nula, 11 caracteres e IMB en posiciones 3-5",
            "tipo": "clues",
        },
        {
            "id": "URG-03",
            "variable": "FECHA_HORA_INICIO_TRIAGE",
            "regla": "No puede ser nula y debe ser fecha/hora válida",
            "tipo": "required_date",
        },
        {
            "id": "URG-04",
            "variable": "FECHA_HORA_TERMINO_TRIAGE",
            "regla": "No puede ser nula y debe ser fecha/hora válida",
            "tipo": "required_date",
        },
        {
            "id": "URG-05",
            "variable": "FECHA_HORA_INGRESO",
            "regla": "No puede ser nula y debe ser fecha/hora válida",
            "tipo": "required_date",
        },
        {
            "id": "URG-06",
            "variable": "DX_INGRESO",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "URG-07",
            "variable": "FECHA_HORA_EGRESO",
            "regla": "No puede ser nula y debe ser fecha/hora válida",
            "tipo": "required_date",
        },
        {
            "id": "URG-08",
            "variable": "DX_PRIN_EGRESO",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
    ],
    "Egresos UCI": [
        {
            "id": "UCI-01",
            "variable": "CVE_REGISTRO_UCI",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "UCI-02",
            "variable": "CVE_REGISTRO_HOS",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "UCI-03",
            "variable": "CLUES",
            "regla": "No nula, 11 caracteres e IMB en posiciones 3-5",
            "tipo": "clues",
        },
        {
            "id": "UCI-04",
            "variable": "FECHA_HORA_INGRESO",
            "regla": "No puede ser nula y debe ser fecha/hora válida",
            "tipo": "required_date",
        },
        {
            "id": "UCI-05",
            "variable": "FECHA_HORA_INGRESO_UCI",
            "regla": "No puede ser nula y debe ser fecha/hora válida",
            "tipo": "required_date",
        },
        {
            "id": "UCI-06",
            "variable": "CVE_ESP_INGRESO",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "UCI-07",
            "variable": "DESC_ESP_INGRESO",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "UCI-08",
            "variable": "DX_INGRESO",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "UCI-09",
            "variable": "DX_PRIN_EGRESO",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "UCI-10",
            "variable": "FECHA_HORA_EGRESO",
            "regla": "No puede ser nula y debe ser fecha/hora válida",
            "tipo": "required_date",
        },
        {
            "id": "UCI-11",
            "variable": "PAC_SEXO",
            "regla": "No nulo y debe coincidir con sexo de PAC_CURP",
            "tipo": "sex_curp",
            "curp": "PAC_CURP",
        },
        {
            "id": "UCI-12",
            "variable": "PAC_EDAD_ANO",
            "regla": "No puede ser nulo y debe ser numérico",
            "tipo": "required_number",
        },
        {
            "id": "UCI-13",
            "variable": "PAC_EDAD_ANO",
            "regla": "Debe coincidir con edad derivada de PAC_CURP al ingreso",
            "tipo": "age_curp",
            "curp": "PAC_CURP",
            "reference": "FECHA_HORA_INGRESO",
        },
        {
            "id": "UCI-14",
            "variable": "PAC_CURP",
            "regla": "No nula y estructura básica de CURP válida",
            "tipo": "curp",
        },
    ],
    "Egresos": [
        {
            "id": "EGR-01",
            "variable": "CVE_REGISTRO_HOS",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "EGR-02",
            "variable": "CVE_REGISTRO_URG",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "EGR-03",
            "variable": "CLUES",
            "regla": "No nula, 11 caracteres e IMB en posiciones 3-5",
            "tipo": "clues",
        },
        {
            "id": "EGR-04",
            "variable": "FECHA_HORA_INGRESO",
            "regla": "No puede ser nula y debe ser fecha/hora válida",
            "tipo": "required_date",
        },
        {
            "id": "EGR-05",
            "variable": "PROCEDENCIA",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "EGR-06",
            "variable": "DESC_ESP_INGRESO",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "EGR-07",
            "variable": "DX_INGRESO",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "EGR-08",
            "variable": "FECHA_HORA_EGRESO",
            "regla": "No puede ser nula y debe ser fecha/hora válida",
            "tipo": "required_date",
        },
        {
            "id": "EGR-09",
            "variable": "DX_PRIN_EGRESO",
            "regla": "No puede ser nulo",
            "tipo": "required",
        },
        {
            "id": "EGR-10",
            "variable": "PAC_FECHANAC",
            "regla": "No puede ser nula y debe ser fecha válida",
            "tipo": "required_date",
        },
        {
            "id": "EGR-11",
            "variable": "PAC_FECHANAC",
            "regla": "Debe coincidir con YYMMDD de PAC_CURP",
            "tipo": "birth_curp",
            "curp": "PAC_CURP",
        },
        {
            "id": "EGR-12",
            "variable": "PAC_SEXO",
            "regla": "No nulo y debe coincidir con sexo de PAC_CURP",
            "tipo": "sex_curp",
            "curp": "PAC_CURP",
        },
        {
            "id": "EGR-13",
            "variable": "PAC_EDAD_ANO",
            "regla": "No puede ser nulo y debe ser numérico",
            "tipo": "required_number",
        },
        {
            "id": "EGR-14",
            "variable": "PAC_EDAD_ANO",
            "regla": "Debe coincidir con edad derivada de PAC_CURP al ingreso",
            "tipo": "age_curp",
            "curp": "PAC_CURP",
            "reference": "FECHA_HORA_INGRESO",
        },
        {
            "id": "EGR-15",
            "variable": "PAC_EDAD_SEMANA",
            "regla": "No puede ser nulo y debe ser numérico",
            "tipo": "required_number",
        },
        {
            "id": "EGR-16",
            "variable": "PAC_EDAD_SEMANA",
            "regla": "Debe ser consistente con PAC_EDAD_ANO",
            "tipo": "weeks_age",
            "years": "PAC_EDAD_ANO",
        },
        {
            "id": "EGR-17",
            "variable": "PAC_CURP",
            "regla": "No nula y estructura básica de CURP válida",
            "tipo": "curp",
        },
    ],
}


# ============================================================================
# UTILIDADES
# ============================================================================


def es_nulo(v):
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except Exception:
        pass
    return isinstance(v, str) and not v.strip()


def texto(v):
    return None if es_nulo(v) else str(v).strip()


def numero(v):
    if es_nulo(v):
        return None
    try:
        return float(v)
    except Exception:
        return None


def fecha(v):
    if es_nulo(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v
    if isinstance(v, (datetime, date)):
        return pd.Timestamp(v)
    d = pd.to_datetime(v, errors="coerce", dayfirst=True)
    return None if pd.isna(d) else d


def valor_mostrar(v):
    if es_nulo(v):
        return "(NULO)"
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(v).strftime("%d/%m/%Y %H:%M:%S")
    return str(v)


def normalizar_columnas(df, modulo):
    lookup = {str(c).strip().casefold(): str(c) for c in df.columns}
    out = {}
    for canonica, aliases in MODULOS[modulo]["variables"].items():
        encontrada = None
        for alias in aliases:
            if alias.strip().casefold() in lookup:
                encontrada = lookup[alias.strip().casefold()]
                break
        out[canonica] = encontrada
    return out


def getv(row, campo, columnas):
    col = columnas.get(campo)
    return None if col is None else row.get(col)


def curp_estructura_valida(v):
    s = texto(v)
    return bool(s and CURP_RE.fullmatch(s.upper()))


def curp_yymmdd(v):
    s = texto(v)
    if not s or len(s) < 10:
        return None
    x = s.upper()[4:10]
    return x if re.fullmatch(r"\d{6}", x) else None


def curp_sexo(v):
    s = texto(v)
    if not s or len(s) < 11:
        return None
    x = s.upper()[10]
    return x if x in {"H", "M"} else None


def fecha_nacimiento_desde_curp(curp, referencia):
    """
    Deriva la fecha de nacimiento a partir de YYMMDD de CURP.
    Debido a que CURP no incluye el siglo de forma explícita en esos 6 caracteres,
    prueba 19YY y 20YY y selecciona una fecha plausible respecto a la referencia.
    """
    ref = fecha(referencia)
    ymd = curp_yymmdd(curp)
    if ref is None or ymd is None:
        return None

    yy, mm, dd = int(ymd[:2]), int(ymd[2:4]), int(ymd[4:6])
    candidatos = []

    for year in (1900 + yy, 2000 + yy):
        try:
            d = pd.Timestamp(year=year, month=mm, day=dd)
            if d <= ref:
                age = ref.year - d.year - ((ref.month, ref.day) < (d.month, d.day))
                if 0 <= age <= 130:
                    candidatos.append((d, age))
        except Exception:
            pass

    if not candidatos:
        return None

    # Si hay dos fechas plausibles, preferir la que produce la menor edad
    # sólo cuando ambas están dentro del rango humano; normalmente habrá una.
    candidatos.sort(key=lambda x: x[1])
    return candidatos[0]


def clave_registro(row, modulo, columnas, fila):
    preferencias = {
        "Cirugías": ["CVE_REGISTRO_CIR"],
        "MOCE": ["CVE_CURP", "CLUES"],
        "Urgencias": ["CVE_REGISTRO_URG"],
        "Egresos UCI": ["CVE_REGISTRO_UCI", "CVE_REGISTRO_HOS"],
        "Egresos": ["CVE_REGISTRO_HOS"],
    }
    for c in preferencias.get(modulo, []):
        v = getv(row, c, columnas)
        if not es_nulo(v):
            return texto(v)
    return f"FILA_{fila}"


# ============================================================================
# MOTOR DE REGLAS
# ============================================================================


def eval_regla(regla, row, columnas):
    tipo = regla["tipo"]
    variable = regla.get("variable")
    v = getv(row, variable, columnas) if variable not in {"Horas de sala"} else None

    if tipo == "required":
        if es_nulo(v):
            return "NO CUMPLE", f"{variable} está nulo o vacío.", v
        return "CUMPLE", "", v

    if tipo == "optional":
        if es_nulo(v):
            return "NO APLICA", "El valor puede ser nulo.", v
        return "CUMPLE", "", v

    if tipo == "required_number":
        if es_nulo(v):
            return "NO CUMPLE", f"{variable} está nulo.", v
        n = numero(v)
        if n is None:
            return "NO CUMPLE", f"{variable} no es numérico.", v
        return "CUMPLE", "", v

    if tipo == "required_date":
        if es_nulo(v):
            return "NO CUMPLE", f"{variable} está nulo.", v
        if fecha(v) is None:
            return "NO CUMPLE", f"{variable} no pudo interpretarse como fecha/hora.", v
        return "CUMPLE", "", v

    if tipo == "optional_date":
        if es_nulo(v):
            return "NO APLICA", "El valor puede ser nulo.", v
        if fecha(v) is None:
            return (
                "NO CUMPLE",
                f"{variable} está informado pero no es fecha/hora válida.",
                v,
            )
        return "CUMPLE", "", v

    if tipo == "clues":
        if es_nulo(v):
            return "NO CUMPLE", "CLUES está nula o vacía.", v
        s = texto(v).upper()
        motivos = []
        if len(s) != 11:
            motivos.append(f"tiene {len(s)} caracteres; se esperaban 11")
        if len(s) < 5 or s[2:5] != "IMB":
            observado = s[2:5] if len(s) >= 5 else s
            motivos.append(
                f"se esperaba IMB en posiciones 3-5 y se observó '{observado}'"
            )
        if motivos:
            return "NO CUMPLE", "CLUES " + "; ".join(motivos) + ".", v
        return "CUMPLE", "", v

    if tipo == "curp":
        if es_nulo(v):
            return "NO CUMPLE", f"{variable} está nula o vacía.", v
        if not curp_estructura_valida(v):
            return (
                "NO CUMPLE",
                "CURP no cumple la estructura básica esperada de 18 caracteres.",
                v,
            )
        return "CUMPLE", "", v

    if tipo == "birth_curp":
        if es_nulo(v):
            return "NO APLICA", f"No se compara porque {variable} está nula.", v
        d = fecha(v)
        cv = getv(row, regla["curp"], columnas)
        ymd = curp_yymmdd(cv)
        if d is None:
            return (
                "NO APLICA",
                f"No se compara porque {variable} no es una fecha válida.",
                v,
            )
        if ymd is None:
            return (
                "NO APLICA",
                f"No se compara porque {regla['curp']} no permite extraer YYMMDD.",
                v,
            )
        esperado = f"{d.year % 100:02d}{d.month:02d}{d.day:02d}"
        if esperado != ymd:
            return (
                "NO CUMPLE",
                f"{variable}={d.strftime('%d/%m/%Y')} equivale a {esperado}, "
                f"pero {regla['curp']} contiene {ymd}.",
                v,
            )
        return "CUMPLE", "", v

    if tipo == "sex_curp":
        if es_nulo(v):
            return "NO CUMPLE", f"{variable} está nulo.", v

        sexo_base = texto(v).strip().upper()

        # Soporta ambas codificaciones usadas en los módulos:
        # M / F y Masculino / Femenino.
        equivalencias = {
            "M": "H",
            "F": "M",
            "MASCULINO": "H",
            "FEMENINO": "M",
        }

        if sexo_base not in equivalencias:
            return (
                "NO CUMPLE",
                f"{variable}='{texto(v)}'. " "Se esperaba M, F, Masculino o Femenino.",
                v,
            )

        esperado_curp = equivalencias[sexo_base]

        cv = getv(row, regla["curp"], columnas)
        sx_curp = curp_sexo(cv)

        if sx_curp is None:
            return (
                "NO APLICA",
                f"No se compara porque {regla['curp']} no tiene sexo interpretable.",
                v,
            )

        if esperado_curp != sx_curp:
            return (
                "NO CUMPLE",
                f"{variable}='{texto(v)}'; "
                f"se esperaba '{esperado_curp}' en CURP, "
                f"pero se observó '{sx_curp}'.",
                v,
            )

        return "CUMPLE", "", v

    if tipo == "age_weeks":
        years = numero(v)
        weeks = numero(getv(row, regla["weeks"], columnas))
        if years is None or weeks is None:
            return "NO APLICA", "Se requieren edad en años y semanas numéricas.", v
        eq = weeks / 52.1775
        if abs(years - eq) > TOLERANCIA_EDAD_SEMANAS_ANOS:
            return (
                "NO CUMPLE",
                f"{weeks:g} semanas equivalen aproximadamente a {eq:.2f} años; "
                f"{variable}={years:g}.",
                v,
            )
        return "CUMPLE", "", v

    if tipo == "weeks_age":
        weeks = numero(v)
        years = numero(getv(row, regla["years"], columnas))
        if weeks is None or years is None:
            return (
                "NO APLICA",
                "Se requieren PAC_EDAD_SEMANA y PAC_EDAD_ANO numéricos.",
                v,
            )
        eq = weeks / 52.1775
        if abs(years - eq) > TOLERANCIA_EDAD_SEMANAS_ANOS:
            return (
                "NO CUMPLE",
                f"PAC_EDAD_SEMANA={weeks:g} equivale aproximadamente a {eq:.2f} años, "
                f"pero PAC_EDAD_ANO={years:g}.",
                v,
            )
        return "CUMPLE", "", v

    if tipo == "age_curp":
        age = numero(v)
        cv = getv(row, regla["curp"], columnas)
        ref = getv(row, regla["reference"], columnas)

        if age is None:
            return "NO APLICA", f"No se compara porque {variable} no es numérico.", v

        derivada = fecha_nacimiento_desde_curp(cv, ref)
        if derivada is None:
            return (
                "NO APLICA",
                (
                    f"No fue posible derivar edad desde {regla['curp']} "
                    f"respecto a {regla['reference']}."
                ),
                v,
            )

        nacimiento, edad_esperada = derivada
        if int(age) != int(edad_esperada):
            return (
                "NO CUMPLE",
                f"{variable}={age:g}; CURP implica nacimiento {nacimiento.strftime('%d/%m/%Y')} "
                f"y edad de {edad_esperada} años a la fecha de ingreso.",
                v,
            )
        return "CUMPLE", "", v

    if tipo == "date_gte":
        a = fecha(v)
        b = fecha(getv(row, regla["other"], columnas))
        if a is None or b is None:
            return (
                "NO APLICA",
                f"Se requieren {variable} y {regla['other']} válidas.",
                v,
            )
        if a < b:
            return (
                "NO CUMPLE",
                f"{variable}={a.strftime('%d/%m/%Y')} es anterior a "
                f"{regla['other']}={b.strftime('%d/%m/%Y')}.",
                v,
            )
        return "CUMPLE", "", v

    if tipo == "required_if_realizada":
        est = (texto(getv(row, "Estatus de la solicitud", columnas)) or "").upper()
        realizada = est in {"REALIZADA", "REALIZADO"}

        if not realizada:
            if es_nulo(v):
                return (
                    "NO APLICA",
                    f"Estatus='{est or '(NULO)'}'; obligatoriedad no aplica.",
                    v,
                )
            if fecha(v) is None:
                return (
                    "NO CUMPLE",
                    f"{variable} está informado pero no es fecha/hora válida.",
                    v,
                )
            return "CUMPLE", "", v

        if es_nulo(v):
            return "NO CUMPLE", f"Estatus='{est}' y {variable} está nulo.", v
        if fecha(v) is None:
            return "NO CUMPLE", f"{variable} no pudo interpretarse como fecha/hora.", v
        return "CUMPLE", "", v

    if tipo == "sequence":
        vals = [(f, fecha(getv(row, f, columnas))) for f in regla["fields"]]
        if any(d is None for _, d in vals):
            return (
                "NO APLICA",
                "Se requieren las cuatro fechas/horas válidas.",
                " | ".join(
                    f"{f}={valor_mostrar(getv(row, f, columnas))}" for f, _ in vals
                ),
            )
        ds = [d for _, d in vals]
        if not all(a <= b for a, b in zip(ds, ds[1:])):
            return (
                "NO CUMPLE",
                "Secuencia inválida. Se esperaba Entrada <= Inicio <= Término <= Salida.",
                " | ".join(f"{f}={d.strftime('%d/%m/%Y %H:%M:%S')}" for f, d in vals),
            )
        return (
            "CUMPLE",
            "",
            " | ".join(f"{f}={d.strftime('%d/%m/%Y %H:%M:%S')}" for f, d in vals),
        )

    return "NO APLICA", f"Tipo de regla no implementado: {tipo}", v


# ============================================================================
# LECTURA DE ARCHIVOS
# ============================================================================


def detectar_separador_csv(ruta):
    with ruta.open("r", encoding="utf-8-sig", errors="replace", newline="") as f:
        muestra = f.read(100_000)

    try:
        return csv.Sniffer().sniff(muestra, delimiters=",;|\t").delimiter
    except Exception:
        primera = muestra.splitlines()[0] if muestra.splitlines() else ""
        candidatos = [",", ";", "\t", "|"]
        return max(candidatos, key=lambda x: primera.count(x))


def leer_archivo(ruta):
    ext = ruta.suffix.lower()

    if ext == ".csv":
        sep = detectar_separador_csv(ruta)
        errores = []
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                print(f"CSV: separador={repr(sep)}, codificación={enc}")
                return pd.read_csv(
                    ruta,
                    sep=sep,
                    encoding=enc,
                    dtype=object,
                    engine="python",
                    on_bad_lines="warn",
                )
            except Exception as e:
                errores.append(f"{enc}: {e}")
        raise ValueError("No fue posible leer el CSV:\n" + "\n".join(errores))

    if ext in {".xlsx", ".xlsm"}:
        return pd.read_excel(ruta, dtype=object, engine="openpyxl")

    if ext == ".xls":
        return pd.read_excel(ruta, dtype=object)

    if ext in {".parquet", ".pq"}:
        return pd.read_parquet(ruta)

    raise ValueError(
        f"Formato no admitido: {ext}\n" "Use CSV, XLSX, XLSM, XLS o Parquet."
    )


ENTIDADES_CLUES = {
    "AS": "Aguascalientes",
    "BC": "Baja California",
    "BS": "Baja California Sur",
    "CC": "Campeche",
    "CS": "Chiapas",
    "CH": "Chihuahua",
    "CL": "Coahuila",
    "CM": "Colima",
    "DF": "Ciudad de México",
    "DG": "Durango",
    "GT": "Guanajuato",
    "GR": "Guerrero",
    "HG": "Hidalgo",
    "JC": "Jalisco",
    "MC": "México",
    "MN": "Michoacán",
    "MS": "Morelos",
    "NT": "Nayarit",
    "NL": "Nuevo León",
    "OC": "Oaxaca",
    "PL": "Puebla",
    "QT": "Querétaro",
    "QR": "Quintana Roo",
    "SP": "San Luis Potosí",
    "SL": "Sinaloa",
    "SR": "Sonora",
    "TC": "Tabasco",
    "TS": "Tamaulipas",
    "TL": "Tlaxcala",
    "VZ": "Veracruz",
    "YN": "Yucatán",
    "ZS": "Zacatecas",
}


def entidad_desde_clues(valor):
    if es_nulo(valor):
        return "SIN ENTIDAD", "Sin entidad"
    clues = texto(valor).strip().upper()
    codigo = clues[:2] if len(clues) >= 2 else ""
    return codigo or "SIN ENTIDAD", ENTIDADES_CLUES.get(codigo, codigo or "Sin entidad")


# ============================================================================
# ANÁLISIS
# ============================================================================


def analizar(modulo, ruta):
    print("=" * 90)
    print(f"ANÁLISIS DE CALIDAD - {modulo.upper()}")
    print("=" * 90)
    print(f"Fuente: {ruta}")

    df = leer_archivo(ruta)
    df.columns = [str(c).strip() for c in df.columns]
    columnas = normalizar_columnas(df, modulo)

    faltantes = [c for c, real in columnas.items() if real is None]
    if faltantes:
        raise ValueError(
            "Faltan variables requeridas para este módulo:\n\n- "
            + "\n- ".join(faltantes)
            + "\n\nColumnas disponibles:\n- "
            + "\n- ".join(map(str, df.columns))
        )

    reglas = REGLAS[modulo]

    stats = {
        r["id"]: {"evaluables": 0, "cumple": 0, "no_cumple": 0, "no_aplica": 0}
        for r in reglas
    }

    resultados = []
    stats_entidad_regla = {}
    registros_entidad = {}
    stats_unidad = {}

    print(f"Registros: {len(df):,}")
    print(f"Reglas: {len(reglas):,}")
    print()

    for i, row in df.iterrows():
        fila = i + 2
        key = clave_registro(row, modulo, columnas, fila)

        clues_valor = getv(row, "CLUES", columnas)
        clues_texto = "" if es_nulo(clues_valor) else texto(clues_valor).strip().upper()
        entidad_codigo, entidad_nombre = entidad_desde_clues(clues_valor)
        registros_entidad[entidad_codigo] = {
            "ENTIDAD": entidad_nombre,
            "REGISTROS": registros_entidad.get(entidad_codigo, {}).get("REGISTROS", 0)
            + 1,
        }

        unidad_key = (entidad_codigo, clues_texto)
        if unidad_key not in stats_unidad:
            stats_unidad[unidad_key] = {"MODULO": modulo, "ENTIDAD_CODIGO": entidad_codigo, "ENTIDAD": entidad_nombre, "CLUES": clues_texto, "REGISTROS": 0, "REGISTROS_CON_INCUMPLIMIENTO": 0, "NO_CUMPLE": 0}
        stats_unidad[unidad_key]["REGISTROS"] += 1
        fila_incumple = False
        for regla in reglas:
            estado, motivo, valor = eval_regla(regla, row, columnas)
            s = stats[regla["id"]]

            ek = (entidad_codigo, regla["id"])
            if ek not in stats_entidad_regla:
                stats_entidad_regla[ek] = {
                    "ENTIDAD": entidad_nombre,
                    "EVALUABLES": 0,
                    "CUMPLE": 0,
                    "NO_CUMPLE": 0,
                    "NO_APLICA": 0,
                }
            se = stats_entidad_regla[ek]

            if estado == "CUMPLE":
                se["EVALUABLES"] += 1
                se["CUMPLE"] += 1
                s["evaluables"] += 1
                s["cumple"] += 1
            elif estado == "NO CUMPLE":
                fila_incumple = True
                stats_unidad[unidad_key]["NO_CUMPLE"] += 1
                s["evaluables"] += 1
                s["no_cumple"] += 1
                se["EVALUABLES"] += 1
                se["NO_CUMPLE"] += 1

                resultados.append(
                    {
                        "MODULO": modulo,
                        "ENTIDAD_CODIGO": entidad_codigo,
                        "ENTIDAD": entidad_nombre,
                        "CLUES": clues_texto,
                        "FILA_ORIGEN": fila,
                        "CLAVE_REGISTRO": key,
                        "REGLA_ID": regla["id"],
                        "VARIABLE": regla["variable"],
                        "RESULTADO": "NO CUMPLE",
                        "VALOR_OBSERVADO": valor_mostrar(valor),
                        "MOTIVO": motivo,
                    }
                )
            else:
                s["no_aplica"] += 1
                se["NO_APLICA"] += 1

        if fila_incumple:
            stats_unidad[unidad_key]["REGISTROS_CON_INCUMPLIMIENTO"] += 1
        if (i + 1) % 10000 == 0:
            print(f"Procesados: {i+1:,}/{len(df):,}")

    reglas_out = []

    for regla in reglas:
        s = stats[regla["id"]]
        pct = (s["cumple"] / s["evaluables"] * 100) if s["evaluables"] else None

        if s["evaluables"] == 0:
            estado = "NO EVALUABLE"
        elif s["no_cumple"] == 0:
            estado = "CUMPLE"
        else:
            estado = "NO CUMPLE"

        reglas_out.append(
            {
                "REGLA_ID": regla["id"],
                "VARIABLE": regla["variable"],
                "REGLA": regla["regla"],
                "ESTADO": estado,
                "EVALUABLES": s["evaluables"],
                "CUMPLE": s["cumple"],
                "NO_CUMPLE": s["no_cumple"],
                "NO_APLICA": s["no_aplica"],
                "PORCENTAJE_CUMPLIMIENTO": round(pct, 2) if pct is not None else None,
            }
        )

    df_reglas = pd.DataFrame(reglas_out)
    df_resultados = pd.DataFrame(
        resultados,
        columns=[
            "MODULO",
            "ENTIDAD_CODIGO",
            "ENTIDAD",
            "CLUES",
            "FILA_ORIGEN",
            "CLAVE_REGISTRO",
            "REGLA_ID",
            "VARIABLE",
            "RESULTADO",
            "VALOR_OBSERVADO",
            "MOTIVO",
        ],
    )

    filas_er = []
    regla_por_id = {r["id"]: r for r in reglas}
    for (codigo, regla_id), se in stats_entidad_regla.items():
        evaluables = se["EVALUABLES"]
        no_cumple = se["NO_CUMPLE"]
        pct_inc = (no_cumple / evaluables * 100) if evaluables else None
        rr = regla_por_id[regla_id]
        filas_er.append(
            {
                "MODULO": modulo,
                "ENTIDAD_CODIGO": codigo,
                "ENTIDAD": se["ENTIDAD"],
                "REGLA_ID": regla_id,
                "VARIABLE": rr["variable"],
                "REGLA": rr["regla"],
                "EVALUABLES": evaluables,
                "CUMPLE": se["CUMPLE"],
                "NO_CUMPLE": no_cumple,
                "NO_APLICA": se["NO_APLICA"],
                "PORCENTAJE_INCUMPLIMIENTO": (
                    round(pct_inc, 2) if pct_inc is not None else None
                ),
            }
        )
    df_entidad_regla = pd.DataFrame(filas_er)

    filas_entidad = []
    for codigo, dato in registros_entidad.items():
        sub = df_entidad_regla[df_entidad_regla["ENTIDAD_CODIGO"] == codigo]
        ev = int(sub["EVALUABLES"].sum()) if not sub.empty else 0
        nc = int(sub["NO_CUMPLE"].sum()) if not sub.empty else 0
        if not df_resultados.empty:
            sub_hallazgos = df_resultados[df_resultados["ENTIDAD_CODIGO"] == codigo]
            registros_con_incumplimiento = int(sub_hallazgos["FILA_ORIGEN"].nunique())
        else:
            registros_con_incumplimiento = 0
        filas_entidad.append(
            {
                "MODULO": modulo,
                "ENTIDAD_CODIGO": codigo,
                "ENTIDAD": dato["ENTIDAD"],
                "REGISTROS": dato["REGISTROS"],
                "REGISTROS_CON_INCUMPLIMIENTO": registros_con_incumplimiento,
                "PORCENTAJE_REGISTROS_CON_INCUMPLIMIENTO": round(
                    registros_con_incumplimiento / dato["REGISTROS"] * 100, 2
                ) if dato["REGISTROS"] else None,
                "EVALUABLES": ev,
                "CUMPLE": int(sub["CUMPLE"].sum()) if not sub.empty else 0,
                "NO_CUMPLE": nc,
                "NO_APLICA": int(sub["NO_APLICA"].sum()) if not sub.empty else 0,
                "PORCENTAJE_INCUMPLIMIENTO": round(nc / ev * 100, 2) if ev else None,
            }
        )
    df_entidad = pd.DataFrame(filas_entidad)

    return df_reglas, df_resultados, len(df), df_entidad, df_entidad_regla, pd.DataFrame(list(stats_unidad.values()))


# ============================================================================
# EXPORTACIÓN A EXCEL
# ============================================================================


def exportar_excel(modulo, ruta_salida, reglas, resultados, total_registros):
    """
    Genera exactamente dos hojas:
      1. Reglas
      2. Resultados
    """
    with pd.ExcelWriter(ruta_salida, engine="xlsxwriter") as writer:
        reglas.to_excel(writer, sheet_name="Reglas", index=False)

        if resultados.empty:
            resultados_export = pd.DataFrame(
                [
                    {
                        "FILA_ORIGEN": "",
                        "CLAVE_REGISTRO": "",
                        "REGLA_ID": "",
                        "VARIABLE": "",
                        "RESULTADO": "SIN INCUMPLIMIENTOS",
                        "VALOR_OBSERVADO": "",
                        "MOTIVO": "No se detectaron incumplimientos en las reglas evaluadas.",
                    }
                ]
            )
        else:
            resultados_export = resultados

        resultados_export.to_excel(writer, sheet_name="Resultados", index=False)

        wb = writer.book
        ws_r = writer.sheets["Reglas"]
        ws_d = writer.sheets["Resultados"]

        header = wb.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#611232",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )

        wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        center = wb.add_format({"align": "center", "valign": "top"})
        percent = wb.add_format({"num_format": '0.00"%"', "align": "center"})
        fail = wb.add_format({"bg_color": "#FCE8E6", "font_color": "#9D2D39"})
        ok = wb.add_format({"bg_color": "#EAF4EE", "font_color": "#235B4E"})
        neutral = wb.add_format({"bg_color": "#F1F3F4", "font_color": "#5F6368"})

        for ws, data in [(ws_r, reglas), (ws_d, resultados_export)]:
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(1, len(data)), len(data.columns) - 1)
            for col_idx, col in enumerate(data.columns):
                ws.write(0, col_idx, col, header)

        widths_rules = {
            "REGLA_ID": 12,
            "VARIABLE": 34,
            "REGLA": 60,
            "ESTADO": 18,
            "EVALUABLES": 14,
            "CUMPLE": 12,
            "NO_CUMPLE": 14,
            "NO_APLICA": 14,
            "PORCENTAJE_CUMPLIMIENTO": 24,
        }
        for i, c in enumerate(reglas.columns):
            ws_r.set_column(i, i, widths_rules.get(c, 18), wrap)

        widths_results = {
            "FILA_ORIGEN": 13,
            "CLAVE_REGISTRO": 24,
            "REGLA_ID": 12,
            "VARIABLE": 34,
            "RESULTADO": 16,
            "VALOR_OBSERVADO": 42,
            "MOTIVO": 75,
        }
        for i, c in enumerate(resultados_export.columns):
            ws_d.set_column(i, i, widths_results.get(c, 18), wrap)

        if len(reglas):
            col_estado = reglas.columns.get_loc("ESTADO")
            ws_r.conditional_format(
                1,
                col_estado,
                len(reglas),
                col_estado,
                {
                    "type": "text",
                    "criteria": "equal to",
                    "value": "CUMPLE",
                    "format": ok,
                },
            )
            ws_r.conditional_format(
                1,
                col_estado,
                len(reglas),
                col_estado,
                {
                    "type": "text",
                    "criteria": "equal to",
                    "value": "NO CUMPLE",
                    "format": fail,
                },
            )
            ws_r.conditional_format(
                1,
                col_estado,
                len(reglas),
                col_estado,
                {
                    "type": "text",
                    "criteria": "equal to",
                    "value": "NO EVALUABLE",
                    "format": neutral,
                },
            )

            pct_col = reglas.columns.get_loc("PORCENTAJE_CUMPLIMIENTO")
            ws_r.set_column(pct_col, pct_col, 24, percent)

        if len(resultados_export):
            col_resultado = resultados_export.columns.get_loc("RESULTADO")
            ws_d.conditional_format(
                1,
                col_resultado,
                len(resultados_export),
                col_resultado,
                {
                    "type": "text",
                    "criteria": "containing",
                    "value": "NO CUMPLE",
                    "format": fail,
                },
            )

        # Comentario metodológico visible arriba mediante comentarios de celda.
        ws_r.write_comment(
            0,
            0,
            f"Módulo: {modulo}\n"
            f"Registros analizados: {total_registros:,}\n"
            "NO APLICA no se incluye en EVALUABLES.\n"
            "PORCENTAJE_CUMPLIMIENTO = CUMPLE / EVALUABLES x 100.",
        )


# ============================================================================
# PROCESAMIENTO DE LOS 5 ARCHIVOS DESDE UNA CARPETA
# ============================================================================

EXTENSIONES_VALIDAS = {".csv", ".xlsx", ".xlsm", ".xls", ".parquet", ".pq"}

# Palabras usadas para identificar cada archivo por su nombre.
# Se evalúan en este orden para evitar que "egresos uci" sea tomado como "egresos".
PATRONES_ARCHIVOS = {
    "Cirugías": [
        ("cirugia",),
        ("cirugias",),
        ("cirugía",),
        ("cirugías",),
    ],
    "MOCE": [
        ("moce",),
        ("moc e",),
    ],
    "Urgencias": [
        ("urgencia",),
        ("urgencias",),
    ],
    "Egresos UCI": [
        ("egreso", "uci"),
        ("egresos", "uci"),
        ("uci",),
    ],
    "Egresos": [
        ("egreso",),
        ("egresos",),
    ],
}


def normalizar_nombre_archivo(nombre):
    import unicodedata

    s = unicodedata.normalize("NFKD", nombre)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.lower().replace("_", " ").replace("-", " ")


def seleccionar_carpeta():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    ruta = filedialog.askdirectory(
        title="Seleccionar carpeta que contiene los 5 archivos"
    )
    root.destroy()

    return Path(ruta) if ruta else None


def identificar_archivos(carpeta):
    """
    Busca los cinco archivos dentro de la carpeta seleccionada.

    Reglas:
    - Sólo considera CSV, Excel y Parquet.
    - Ignora archivos temporales de Excel (~$).
    - Ignora archivos de salida que empiecen por calidad_.
    - Egresos UCI se identifica antes que Egresos.
    - Si un módulo tiene 0 o más de 1 candidato, se informa claramente.
    """
    archivos = [
        p
        for p in carpeta.iterdir()
        if p.is_file()
        and p.suffix.lower() in EXTENSIONES_VALIDAS
        and not p.name.startswith("~$")
        and not p.name.lower().startswith("calidad_")
    ]

    encontrados = {}
    usados = set()

    orden = ["Egresos UCI", "Cirugías", "MOCE", "Urgencias", "Egresos"]

    for modulo in orden:
        candidatos = []

        for archivo in archivos:
            if archivo in usados:
                continue

            nombre = normalizar_nombre_archivo(archivo.stem)

            for patron in PATRONES_ARCHIVOS[modulo]:
                patron_norm = tuple(normalizar_nombre_archivo(x) for x in patron)
                if all(p in nombre for p in patron_norm):
                    # Evitar que Egresos simple capture un archivo UCI.
                    if modulo == "Egresos" and "uci" in nombre:
                        continue
                    candidatos.append(archivo)
                    break

        if len(candidatos) == 1:
            encontrados[modulo] = candidatos[0]
            usados.add(candidatos[0])
        elif len(candidatos) == 0:
            encontrados[modulo] = None
        else:
            encontrados[modulo] = candidatos

    return encontrados, archivos


def validar_identificacion(encontrados):
    problemas = []

    for modulo in ["Cirugías", "MOCE", "Urgencias", "Egresos UCI", "Egresos"]:
        valor = encontrados.get(modulo)

        if valor is None:
            problemas.append(f"{modulo}: no se encontró archivo.")
        elif isinstance(valor, list):
            nombres = ", ".join(p.name for p in valor)
            problemas.append(f"{modulo}: se encontraron varios candidatos: {nombres}")

    if problemas:
        raise ValueError(
            "No fue posible identificar de forma única los 5 archivos.\n\n"
            + "\n".join(f"- {p}" for p in problemas)
            + "\n\n"
            "Sugerencia de nombres:\n"
            "  cirugias.xlsx\n"
            "  moce.xlsx\n"
            "  urgencias.xlsx\n"
            "  egresos_uci.xlsx\n"
            "  egresos.xlsx"
        )


def seleccionar_salida_consolidada(carpeta):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    nombre = f"calidad_5_modulos_{datetime.now():%Y%m%d_%H%M%S}.xlsx"

    ruta = filedialog.asksaveasfilename(
        title="Guardar análisis consolidado de los 5 módulos",
        initialdir=str(carpeta),
        defaultextension=".xlsx",
        initialfile=nombre,
        filetypes=[("Excel", "*.xlsx")],
    )
    root.destroy()

    return Path(ruta) if ruta else None


def exportar_excel_consolidado(ruta_salida, analisis):
    """
    Genera un solo Excel con 13 hojas: 3 resúmenes + 10 hojas de detalle:
      Cir_Reglas, Cir_Resultados
      MOCE_Reglas, MOCE_Resultados
      Urg_Reglas, Urg_Resultados
      UCI_Reglas, UCI_Resultados
      Egr_Reglas, Egr_Resultados
    """
    nombres = {
        "Cirugías": ("Cir_Reglas", "Cir_Resultados"),
        "MOCE": ("MOCE_Reglas", "MOCE_Resultados"),
        "Urgencias": ("Urg_Reglas", "Urg_Resultados"),
        "Egresos UCI": ("UCI_Reglas", "UCI_Resultados"),
        "Egresos": ("Egr_Reglas", "Egr_Resultados"),
    }

    with pd.ExcelWriter(ruta_salida, engine="xlsxwriter") as writer:
        wb = writer.book

        header = wb.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#611232",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        percent = wb.add_format({"num_format": '0.00"%"', "align": "center"})
        fail = wb.add_format({"bg_color": "#FCE8E6", "font_color": "#9D2D39"})
        ok = wb.add_format({"bg_color": "#EAF4EE", "font_color": "#235B4E"})
        neutral = wb.add_format({"bg_color": "#F1F3F4", "font_color": "#5F6368"})

        resumen_modulos = []
        resumen_entidades = []
        resumen_entidad_regla = []
        for modulo, info in analisis.items():
            r = info["reglas"]
            ev = int(r["EVALUABLES"].sum())
            nc = int(r["NO_CUMPLE"].sum())
            resumen_modulos.append(
                {
                    "MODULO": modulo,
                    "ARCHIVO": info["archivo"].name,
                    "REGISTROS": info["total"],
                    "REGLAS": len(r),
                    "REGLAS_CON_INCUMPLIMIENTO": int((r["NO_CUMPLE"] > 0).sum()),
                    "EVALUABLES": ev,
                    "CUMPLE": int(r["CUMPLE"].sum()),
                    "NO_CUMPLE": nc,
                    "NO_APLICA": int(r["NO_APLICA"].sum()),
                    "PORCENTAJE_INCUMPLIMIENTO": (
                        round(nc / ev * 100, 2) if ev else None
                    ),
                }
            )
            resumen_entidades.append(info["resumen_entidad"])
            resumen_entidad_regla.append(info["resumen_entidad_regla"])

        df_rm = pd.DataFrame(resumen_modulos)
        df_re = (
            pd.concat(resumen_entidades, ignore_index=True)
            if resumen_entidades
            else pd.DataFrame()
        )
        df_rer = (
            pd.concat(resumen_entidad_regla, ignore_index=True)
            if resumen_entidad_regla
            else pd.DataFrame()
        )

        for nombre, data in (
            ("Resumen_Modulos", df_rm),
            ("Resumen_Entidad", df_re),
            ("Resumen_Entidad_Regla", df_rer),
        ):
            data.to_excel(writer, sheet_name=nombre, index=False)
            ws = writer.sheets[nombre]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(1, len(data)), max(0, len(data.columns) - 1))
            for j, col in enumerate(data.columns):
                ws.write(0, j, col, header)
                ancho = 28 if col in {"ARCHIVO", "ENTIDAD", "REGLA", "VARIABLE"} else 18
                ws.set_column(j, j, ancho, wrap)
            if "PORCENTAJE_INCUMPLIMIENTO" in data.columns:
                j = data.columns.get_loc("PORCENTAJE_INCUMPLIMIENTO")
                ws.set_column(j, j, 26, percent)

        for modulo, info in analisis.items():
            reglas = info["reglas"]
            resultados = info["resultados"]
            total = info["total"]
            archivo = info["archivo"]

            hoja_reglas, hoja_resultados = nombres[modulo]

            reglas.to_excel(writer, sheet_name=hoja_reglas, index=False)

            if resultados.empty:
                resultados_export = pd.DataFrame(
                    [
                        {
                            "MODULO": modulo,
                            "ENTIDAD_CODIGO": "",
                            "ENTIDAD": "",
                            "CLUES": "",
                            "FILA_ORIGEN": "",
                            "CLAVE_REGISTRO": "",
                            "REGLA_ID": "",
                            "VARIABLE": "",
                            "RESULTADO": "SIN INCUMPLIMIENTOS",
                            "VALOR_OBSERVADO": "",
                            "MOTIVO": "No se detectaron incumplimientos en las reglas evaluadas.",
                        }
                    ]
                )
            else:
                resultados_export = resultados

            resultados_export.to_excel(writer, sheet_name=hoja_resultados, index=False)

            ws_r = writer.sheets[hoja_reglas]
            ws_d = writer.sheets[hoja_resultados]

            for ws, data in ((ws_r, reglas), (ws_d, resultados_export)):
                ws.freeze_panes(1, 0)
                ws.autofilter(0, 0, max(1, len(data)), len(data.columns) - 1)

                for col_idx, col in enumerate(data.columns):
                    ws.write(0, col_idx, col, header)

            widths_rules = {
                "REGLA_ID": 12,
                "VARIABLE": 34,
                "REGLA": 62,
                "ESTADO": 18,
                "EVALUABLES": 14,
                "CUMPLE": 12,
                "NO_CUMPLE": 14,
                "NO_APLICA": 14,
                "PORCENTAJE_CUMPLIMIENTO": 24,
            }

            for i, c in enumerate(reglas.columns):
                ws_r.set_column(i, i, widths_rules.get(c, 18), wrap)

            widths_results = {
                "MODULO": 18,
                "ENTIDAD_CODIGO": 16,
                "ENTIDAD": 24,
                "CLUES": 16,
                "FILA_ORIGEN": 13,
                "CLAVE_REGISTRO": 25,
                "REGLA_ID": 12,
                "VARIABLE": 34,
                "RESULTADO": 16,
                "VALOR_OBSERVADO": 42,
                "MOTIVO": 78,
            }

            for i, c in enumerate(resultados_export.columns):
                ws_d.set_column(i, i, widths_results.get(c, 18), wrap)

            if len(reglas):
                col_estado = reglas.columns.get_loc("ESTADO")

                ws_r.conditional_format(
                    1,
                    col_estado,
                    len(reglas),
                    col_estado,
                    {
                        "type": "text",
                        "criteria": "equal to",
                        "value": "CUMPLE",
                        "format": ok,
                    },
                )
                ws_r.conditional_format(
                    1,
                    col_estado,
                    len(reglas),
                    col_estado,
                    {
                        "type": "text",
                        "criteria": "equal to",
                        "value": "NO CUMPLE",
                        "format": fail,
                    },
                )
                ws_r.conditional_format(
                    1,
                    col_estado,
                    len(reglas),
                    col_estado,
                    {
                        "type": "text",
                        "criteria": "equal to",
                        "value": "NO EVALUABLE",
                        "format": neutral,
                    },
                )

                pct_col = reglas.columns.get_loc("PORCENTAJE_CUMPLIMIENTO")
                ws_r.set_column(pct_col, pct_col, 24, percent)

            if len(resultados_export):
                col_res = resultados_export.columns.get_loc("RESULTADO")
                ws_d.conditional_format(
                    1,
                    col_res,
                    len(resultados_export),
                    col_res,
                    {
                        "type": "text",
                        "criteria": "containing",
                        "value": "NO CUMPLE",
                        "format": fail,
                    },
                )

            ws_r.write_comment(
                0,
                0,
                f"Módulo: {modulo}\n"
                f"Archivo: {archivo.name}\n"
                f"Registros analizados: {total:,}\n"
                "NO APLICA no se incluye en EVALUABLES.\n"
                "PORCENTAJE_CUMPLIMIENTO = CUMPLE / EVALUABLES x 100.",
            )



def exportar_json_tablero(analisis, ruta_json=None, fecha_corte=None):
    """Publica únicamente agregados; histórico filtrable y una observación por semana ISO."""
    if ruta_json is None:
        ruta_json = Path(__file__).resolve().parents[1] / "tablero" / "data" / "calidad.json"
    ruta_json = Path(ruta_json)
    ruta_json.parent.mkdir(parents=True, exist_ok=True)
    ruta_tendencia = ruta_json.parent / "tendencia.json"
    catalogo_path = ruta_json.parent / "catalogos" / "unidades.json"
    catalogo = {}
    if catalogo_path.exists():
        with catalogo_path.open(encoding="utf-8") as f:
            catalogo = {str(x.get("clues", "")).strip().upper(): x for x in json.load(f).get("unidades", [])}

    def records(df):
        if df is None or df.empty:
            return []
        return json.loads(df.to_json(orient="records", force_ascii=False, date_format="iso"))

    modulos, entidades, entidad_reglas, unidades = [], [], [], []
    for modulo, info in analisis.items():
        rr = info["reglas"]
        hall = info["resultados"]
        total = int(info["total"])
        afectados = int(hall["FILA_ORIGEN"].nunique()) if not hall.empty else 0
        ev = int(rr["EVALUABLES"].sum())
        nc = int(rr["NO_CUMPLE"].sum())
        cumple = int(rr["CUMPLE"].sum())
        na = int(rr["NO_APLICA"].sum())
        if ev != cumple + nc:
            raise ValueError(f"Inconsistencia EVALUABLES en {modulo}")
        if afectados > total:
            raise ValueError(f"Registros afectados mayores al universo en {modulo}")
        modulos.append({"MODULO": modulo, "REGISTROS": total, "REGISTROS_CON_INCUMPLIMIENTO": afectados, "NO_CUMPLE": nc, "CUMPLE": cumple, "EVALUABLES": ev, "NO_APLICA": na})
        entidades.extend(records(info["resumen_entidad"]))
        entidad_reglas.extend(records(info["resumen_entidad_regla"]))
        for u in records(info["resumen_unidad"]):
            c = catalogo.get(u["CLUES"], {})
            u["NOMBRE_UNIDAD"] = c.get("nombre", "Nombre no disponible")
            u["MUNICIPIO"] = c.get("municipio", "")
            unidades.append(u)

    ahora = datetime.now()
    fecha_corte = fecha_corte or ahora.date()
    if isinstance(fecha_corte, str):
        fecha_corte = date.fromisoformat(fecha_corte)
    semana = f"{fecha_corte.isocalendar().year}-S{fecha_corte.isocalendar().week:02d}"
    payload = {
        "metadata": {"titulo": "Calidad de datos institucional", "generado": ahora.isoformat(timespec="seconds"), "fecha_corte": fecha_corte.isoformat(), "semana_iso": semana, "version": "6.0", "reglas_version": "v6-reglas"},
        "resumen_modulos": modulos,
        "resumen_entidad": entidades,
        "resumen_entidad_regla": entidad_reglas,
        "resumen_unidad": unidades,
    }
    # Históricos anteriores de v5 sólo tenían totales globales y NO son comparables
    # por módulo, entidad y regla. No se migran como si fueran series filtrables.
    cortes = []
    if ruta_tendencia.exists():
        try:
            with ruta_tendencia.open(encoding="utf-8") as f:
                cortes = [x for x in json.load(f).get("cortes", []) if x.get("version") == "6.0" and x.get("resumen_modulos") is not None]
        except (ValueError, OSError):
            pass
    corte = {"version": "6.0", "fecha_corte": payload["metadata"]["fecha_corte"], "semana_iso": semana, "generado": payload["metadata"]["generado"], "reglas_version": payload["metadata"]["reglas_version"], "resumen_modulos": modulos, "resumen_entidad": entidades, "resumen_entidad_regla": entidad_reglas}
    cortes = [x for x in cortes if x.get("semana_iso") != semana]
    cortes.append(corte)
    cortes.sort(key=lambda x: x["fecha_corte"])
    with ruta_json.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, allow_nan=False)
    with ruta_tendencia.open("w", encoding="utf-8") as f:
        json.dump({"version": "6.0", "cortes": cortes}, f, ensure_ascii=False, indent=2, allow_nan=False)
    return ruta_json, ruta_tendencia


# ============================================================================
# MAIN - NOTEBOOK / 5 MÓDULOS
# ============================================================================


def pedir_fecha_corte(nombre):
    """Fecha de referencia de los datos (no fecha de ejecución)."""
    while True:
        valor = input(f"Fecha de corte {nombre} (AAAA-MM-DD): ").strip()
        try:
            return date.fromisoformat(valor)
        except ValueError:
            print("Fecha inválida; ejemplo: 2026-10-11")


def procesar_carpeta(carpeta, etiqueta):
    encontrados, _ = identificar_archivos(carpeta)
    validar_identificacion(encontrados)
    analisis = {}
    for indice, modulo in enumerate(["Cirugías", "MOCE", "Urgencias", "Egresos UCI", "Egresos"], 1):
        archivo = encontrados[modulo]
        print(f"[{etiqueta}] {indice}/5 {modulo}: {archivo.name}")
        reglas, resultados, total, resumen_entidad, resumen_entidad_regla, resumen_unidad = analizar(modulo, archivo)
        analisis[modulo] = {
            "reglas": reglas, "resultados": resultados, "total": total,
            "archivo": archivo, "resumen_entidad": resumen_entidad,
            "resumen_entidad_regla": resumen_entidad_regla, "resumen_unidad": resumen_unidad,
        }
        print(f"  {total:,} registros | {len(resultados):,} evaluaciones incumplidas")
    return analisis


def main():
    print("=" * 80)
    print("CALIDAD DE DATOS — COMPARACIÓN DE DOS CORTES (5 MÓDULOS POR CORTE)")
    print("Se aplicarán las mismas reglas v6 a ambos conjuntos.")
    print("Los archivos pueden ser acumulativos: se comparan dos fotografías de calidad")
    print("del universo completo, NO sólo los registros nuevos de cada semana.")
    print("=" * 80)

    print("\nSelecciona la carpeta del CORTE ANTERIOR (5 archivos).")
    anterior = seleccionar_carpeta()
    if anterior is None:
        print("Cancelado")
        return
    fecha_anterior = pedir_fecha_corte("ANTERIOR")

    print("\nSelecciona la carpeta del CORTE ACTUAL (5 archivos).")
    actual = seleccionar_carpeta()
    if actual is None:
        print("Cancelado")
        return
    fecha_actual = pedir_fecha_corte("ACTUAL")
    if fecha_anterior.isocalendar()[:2] == fecha_actual.isocalendar()[:2]:
        raise ValueError("Ambos cortes pertenecen a la misma semana ISO; selecciona semanas distintas.")
    if fecha_actual <= fecha_anterior:
        raise ValueError("La fecha actual debe ser posterior a la anterior.")

    # Primero validar identificación de ambos conjuntos, antes de modificar archivos de salida.
    for carpeta in (anterior, actual):
        identificados, _ = identificar_archivos(carpeta)
        validar_identificacion(identificados)

    print("\nProcesando corte anterior...")
    analisis_anterior = procesar_carpeta(anterior, "ANTERIOR")
    print("\nProcesando corte actual...")
    analisis_actual = procesar_carpeta(actual, "ACTUAL")

    print("\nSelecciona dónde guardar el Excel PRIVADO del corte anterior.")
    salida_anterior = seleccionar_salida_consolidada(anterior)
    if not salida_anterior:
        print("Cancelado")
        return
    print("\nSelecciona dónde guardar el Excel PRIVADO del corte actual.")
    salida_actual = seleccionar_salida_consolidada(actual)
    if not salida_actual:
        print("Cancelado")
        return
    if salida_anterior.resolve() == salida_actual.resolve():
        raise ValueError("Elige rutas diferentes para los dos Excel; no se deben sobrescribir.")

    exportar_excel_consolidado(salida_anterior, analisis_anterior)
    exportar_excel_consolidado(salida_actual, analisis_actual)
    # El último corte exportado permanece como calidad.json actual.
    exportar_json_tablero(analisis_anterior, fecha_corte=fecha_anterior)
    ruta_json, ruta_tendencia = exportar_json_tablero(analisis_actual, fecha_corte=fecha_actual)

    def resumen(a):
        total = sum(int(x["total"]) for x in a.values())
        afectados = sum(int(x["resultados"]["FILA_ORIGEN"].nunique()) if not x["resultados"].empty else 0 for x in a.values())
        return total, afectados, (100 * (total - afectados) / total if total else None)
    r1, f1, q1 = resumen(analisis_anterior)
    r2, f2, q2 = resumen(analisis_actual)
    print("\n" + "=" * 80)
    print(f"Anterior ({fecha_anterior}): {r1:,} registros, {f1:,} afectados, calidad {q1:.2f}%")
    print(f"Actual   ({fecha_actual}): {r2:,} registros, {f2:,} afectados, calidad {q2:.2f}%")
    print(f"Variación de calidad: {q2 - q1:+.2f} puntos porcentuales")
    print(f"Histórico actualizado: {ruta_tendencia}")
    print(f"Tablero actual: {ruta_json}")
    print("NOTA: la variación no mide necesariamente correcciones sobre los mismos registros.")
    print("=" * 80)


if __name__ == "__main__":
    main()
