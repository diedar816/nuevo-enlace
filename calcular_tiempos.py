
# -*- coding: utf-8 -*-
import os
import pandas as pd
import numpy as np
from openpyxl.styles import numbers

print("== Inicio del script ==")

# -------------------------------------------------------------------
# Rutas
# -------------------------------------------------------------------
BASE_DIR = r"C:\Proyecto TICs\Nuevo Enlace"
ENTRADA  = os.path.join(BASE_DIR, "Tabla Modelo.XLSX")
SALIDA   = os.path.join(BASE_DIR, "Tabla_ Resultado.xlsx")

print(f"Carpeta base: {BASE_DIR}")
print(f"Archivo entrada: {ENTRADA}")

# -------------------------------------------------------------------
# Festivos Colombia 2025 (ejemplo)
# -------------------------------------------------------------------
HOLIDAYS_2025 = np.array([
    "2025-01-01","2025-01-06","2025-03-24","2025-04-17","2025-04-18","2025-05-01",
    "2025-06-02","2025-06-23","2025-06-30","2025-07-20","2025-08-07","2025-08-18",
    "2025-10-13","2025-11-03","2025-11-17","2025-12-08","2025-12-25"
], dtype="datetime64[D]")
HOLIDAYS_2025 = np.unique(HOLIDAYS_2025)

INCLUDE_END = True  # incluir día final en hábiles

# -------------------------------------------------------------------
# Utilidades
# -------------------------------------------------------------------
def detectar_columna(df, candidatos):
    reales = list(df.columns)
    mapa = {c.lower().strip(): c for c in reales}
    for cand in candidatos:
        k = cand.lower().strip()
        if k in mapa: 
            return mapa[k]
    for cand in candidatos:
        k = cand.lower().strip()
        for norm, real in mapa.items():
            if k in norm: 
                return real
    return None

def parsear_fecha_solo_date(serie):
    """Convierte a datetime.date (sin hora)."""
    dt = pd.to_datetime(serie, errors="coerce", dayfirst=True)
    return dt.dt.date

LONG_RAD = 14
def normalizar_radicado(x):
    """Texto fijo de 14 dígitos, sin notación científica."""
    if pd.isna(x):
        return None
    s = str(x).strip()
    try:
        if "e" in s.lower():
            s = str(int(float(s)))
        else:
            s_num = s.replace(",", "").replace(" ", "")
            if s_num.replace(".", "", 1).isdigit():
                s = str(int(float(s_num)))
    except:
        pass
    s_digits = "".join(ch for ch in s if ch.isdigit())
    return s_digits.zfill(LONG_RAD) if s_digits else s

def dias_habiles(a, b, include_end=INCLUDE_END):
    """Días hábiles (lun-vie) entre a y b excluyendo festivos (2025)."""
    if a is None or b is None or pd.isna(a) or pd.isna(b):
        return np.nan
    start = np.datetime64(a)  # a y b son datetime.date
    end   = np.datetime64(b)

    if include_end and end >= start:
        end += np.timedelta64(1, "D")
    elif include_end and end < start:
        start += np.timedelta64(1, "D")
        return -np.busday_count(end, start, weekmask="1111100", holidays=HOLIDAYS_2025)

    if end >= start:
        return np.busday_count(start, end, weekmask="1111100", holidays=HOLIDAYS_2025)
    else:
        return -np.busday_count(end, start, weekmask="1111100", holidays=HOLIDAYS_2025)

def dias_naturales(a, b, inclusive=False):
    """Días naturales entre a y b. Si inclusive=True, suma 1."""
    if a is None or b is None or pd.isna(a) or pd.isna(b):
        return np.nan
    diff = (b - a).days
    return diff + 1 if inclusive else diff

# -------------------------------------------------------------------
# Lectura
# -------------------------------------------------------------------
if not os.path.exists(ENTRADA):
    raise FileNotFoundError(f"No se encontró el archivo: {ENTRADA}")

print("Leyendo Excel...")
df = pd.read_excel(ENTRADA, engine="openpyxl")
print(f"Filas: {len(df)} | Columnas: {len(df.columns)}")

# -------------------------------------------------------------------
# Detección de columnas principales
# -------------------------------------------------------------------
col_radicado  = detectar_columna(df, ["numero_radicado","nro_radicado","radicado"])
col_dep       = detectar_columna(df, ["nombre_dependencia","dependencia"])
col_asigna    = detectar_columna(df, ["fecha_asignacion","fecha_asigna","f_asignacion","asignacion"])
col_caso_crea = detectar_columna(df, ["CASO_FECHA_CREACION","caso_fecha_creacion","fecha_creacion_caso","f_caso_crea"])

print("Columnas detectadas:")
print("  numero_radicado     =>", col_radicado)
print("  nombre_dependencia  =>", col_dep)
print("  fecha_asignacion    =>", col_asigna)
print("  CASO_FECHA_CREACION =>", col_caso_crea)

faltantes = [k for k,v in {
    "numero_radicado": col_radicado,
    "nombre_dependencia": col_dep,
    "fecha_asignacion": col_asigna,
    "CASO_FECHA_CREACION": col_caso_crea,
}.items() if v is None]
if faltantes:
    raise ValueError(f"Faltan columnas imprescindibles: {faltantes}")

# -------------------------------------------------------------------
# Normalizar fechas (sin hora) y radicado (texto de 14 dígitos)
# ⚠ Se sobreescriben las columnas originales para que NO salgan con hora
# -------------------------------------------------------------------
df[col_asigna]    = parsear_fecha_solo_date(df[col_asigna])
df[col_caso_crea] = parsear_fecha_solo_date(df[col_caso_crea])

df["_fecha_asignacion"]    = df[col_asigna]          # internas (date)
df["_caso_fecha_creacion"] = df[col_caso_crea]       # internas (date)

df["_rad_str"] = df[col_radicado].apply(normalizar_radicado).astype("string")

# -------------------------------------------------------------------
# Máscaras de dependencia
# -------------------------------------------------------------------
dep_upper = df[col_dep].astype(str).str.upper()
mask_vent = dep_upper.str.startswith("VENTANILLA", na=False)
mask_mesa = dep_upper.str.startswith("MESA",       na=False)

# -------------------------------------------------------------------
# PAREO VERTICAL (VENTANILLA ↔ MESA) POR RADICADO
# TOTAL_DIAS_VENTANILLA = fecha_asignacion (MESA) - fecha_asignacion (VENTANILLA)
# (Se usa internamente; NO se escribirá 'MESA_FECHA_REFERENCIA' en el Excel)
# -------------------------------------------------------------------
print("Construyendo referencia de fecha_asignacion en MESA por radicado...")
mesa_asig = (df.loc[mask_mesa, ["_rad_str", "_fecha_asignacion"]]
               .dropna()
               .groupby("_rad_str")["_fecha_asignacion"]
               .apply(list)
               .to_dict())

def mesa_ref_date(radicado, vent_date):
    """Elige la fecha_asignacion de MESA para el radicado.
    Preferimos la primera fecha MESA >= vent_date; si no hay, usamos la mínima."""
    fechas = mesa_asig.get(radicado, [])
    fechas_validas = [d for d in fechas if pd.notna(d)]
    if not fechas_validas:
        return pd.NaT
    posteriores = [d for d in fechas_validas if d >= vent_date]
    return min(posteriores) if posteriores else min(fechas_validas)

print("Calculando TOTAL_DIAS_VENTANILLA (hábiles y naturales)...")
df["TOTAL_DIAS_VENTANILLA_NATURALES"] = np.nan
df["TOTAL_DIAS_VENTANILLA_HABILES"]   = np.nan

pareos_logrados = 0  # contador solicitado

for i in df.index:
    if mask_vent.iloc[i] and pd.notna(df.loc[i, "_fecha_asignacion"]):
        rad = df.loc[i, "_rad_str"]
        vent_date = df.loc[i, "_fecha_asignacion"]
        mesa_date = mesa_ref_date(rad, vent_date)
        if pd.notna(mesa_date):
            pareos_logrados += 1
            # Naturales (exclusivos; cambia inclusive=True si lo requieres)
            df.at[i, "TOTAL_DIAS_VENTANILLA_NATURALES"] = dias_naturales(vent_date, mesa_date, inclusive=False)
            # Hábiles (INCLUDE_END controla inclusión del fin)
            df.at[i, "TOTAL_DIAS_VENTANILLA_HABILES"]   = dias_habiles(vent_date, mesa_date, include_end=INCLUDE_END)

df["TOTAL_DIAS_VENTANILLA"] = df["TOTAL_DIAS_VENTANILLA_HABILES"]

# 👉 Reporte solicitado: pareos logrados
print(f"Pareos VENT ↔ MESA logrados: {pareos_logrados}/{int(mask_vent.sum())}")

# -------------------------------------------------------------------
# HORIZONTAL EN MESA
# TIEMPO_EN_MESA_DE_CREACION = CASO_FECHA_CREACION (MESA) - fecha_asignacion (MESA)
# -------------------------------------------------------------------
print("Calculando TIEMPO_EN_MESA_DE_CREACION (hábiles y naturales)...")
df["TIEMPO_EN_MESA_DE_CREACION_NATURALES"] = np.nan
df["TIEMPO_EN_MESA_DE_CREACION_HABILES"]   = np.nan

mesa_rows = df.loc[mask_mesa].index

df.loc[mesa_rows, "TIEMPO_EN_MESA_DE_CREACION_NATURALES"] = [
    dias_naturales(a, b, inclusive=False) for a, b in zip(
        df.loc[mesa_rows, "_fecha_asignacion"],
        df.loc[mesa_rows, "_caso_fecha_creacion"]
    )
]
df.loc[mesa_rows, "TIEMPO_EN_MESA_DE_CREACION_HABILES"] = [
    dias_habiles(a, b, include_end=INCLUDE_END) for a, b in zip(
        df.loc[mesa_rows, "_fecha_asignacion"],
        df.loc[mesa_rows, "_caso_fecha_creacion"]
    )
]
df["TIEMPO_EN_MESA_DE_CREACION"] = df["TIEMPO_EN_MESA_DE_CREACION_HABILES"]

# -------------------------------------------------------------------
# ALERTAS
# -------------------------------------------------------------------
print("Generando alertas...")
ALERTA = []
for i, row in df.iterrows():
    avisos = []
    if mask_vent.iloc[i]:
        if pd.isna(row["TOTAL_DIAS_VENTANILLA_HABILES"]) and pd.notna(row["_fecha_asignacion"]):
            avisos.append("VENT: Sin MESA para pareo (mismo radicado)")
        if pd.notna(row["TOTAL_DIAS_VENTANILLA_HABILES"]) and row["TOTAL_DIAS_VENTANILLA_HABILES"] < 0:
            avisos.append("VENT: Diferencia negativa (MESA < VENT)")
        # Auditoría: si hubo múltiples MESA y se logró pareo
        cant_mesa = len([d for d in mesa_asig.get(row["_rad_str"], []) if pd.notna(d)])
        if cant_mesa > 1 and pd.notna(row["TOTAL_DIAS_VENTANILLA_HABILES"]):
            avisos.append(f"VENT: MESA múltiple ({cant_mesa}) → se eligió la más cercana ≥ VENT")

    if mask_mesa.iloc[i]:
        if pd.isna(row["_fecha_asignacion"]) or pd.isna(row["_caso_fecha_creacion"]):
            avisos.append("MESA: Fecha(s) faltante(s)")
        elif pd.notna(row["TIEMPO_EN_MESA_DE_CREACION_HABILES"]) and row["TIEMPO_EN_MESA_DE_CREACION_HABILES"] < 0:
            avisos.append("MESA: Diferencia negativa (CREACIÓN < ASIGNACIÓN)")
    ALERTA.append("; ".join(avisos) if avisos else "")
df["ALERTA"] = ALERTA

# Filas pertinentes
pertinentes = df["ALERTA"].astype(bool) | \
              df["TOTAL_DIAS_VENTANILLA"].notna() | \
              df["TIEMPO_EN_MESA_DE_CREACION"].notna()
df_pert = df.loc[pertinentes].copy()

# -------------------------------------------------------------------
# Resumen por radicado (vista rápida)
# -------------------------------------------------------------------
print("Construyendo resumen por radicado...")
tmp = df.loc[mask_vent | mask_mesa, [col_radicado, col_dep,
    "TOTAL_DIAS_VENTANILLA","TIEMPO_EN_MESA_DE_CREACION",
    "TOTAL_DIAS_VENTANILLA_NATURALES","TIEMPO_EN_MESA_DE_CREACION_NATURALES"]].copy()
tmp["tipo"] = np.where(tmp[col_dep].astype(str).str.upper().str.startswith("VENTANILLA"), "VENTANILLA", "MESA")

summary = (tmp
    .pivot_table(index=col_radicado,
                 columns="tipo",
                 values=["TOTAL_DIAS_VENTANILLA","TIEMPO_EN_MESA_DE_CREACION",
                         "TOTAL_DIAS_VENTANILLA_NATURALES","TIEMPO_EN_MESA_DE_CREACION_NATURALES"],
                 aggfunc="max")
    .reset_index()
)
# Asegurar que la primera columna se llama exactamente col_radicado
if summary.columns[0] != col_radicado:
    summary.rename(columns={summary.columns[0]: col_radicado}, inplace=True)

# Renombrar columnas multiíndice a plano
summary.columns = [
    (c if isinstance(c, str) else f"{c[0]}_{c[1]}")
    for c in summary.columns
]

# -------------------------------------------------------------------
# Auditoría columnas detectadas
# -------------------------------------------------------------------
detalles_cols = pd.DataFrame([
    {"campo_logico": "numero_radicado", "columna_en_archivo": col_radicado},
    {"campo_logico": "nombre_dependencia", "columna_en_archivo": col_dep},
    {"campo_logico": "fecha_asignacion", "columna_en_archivo": col_asigna},
    {"campo_logico": "CASO_FECHA_CREACION", "columna_en_archivo": col_caso_crea},
    {"campo_logico": "Regla TOTAL_DIAS_VENTANILLA", "columna_en_archivo": "MESA.fecha_asignacion - VENT.fecha_asignacion (mismo radicado)"},
    {"campo_logico": "Regla TIEMPO_EN_MESA_DE_CREACION", "columna_en_archivo": "MESA.CASO_FECHA_CREACION - MESA.fecha_asignacion"}
])

# -------------------------------------------------------------------
# Preparar tipos finales para Excel:
# - Fechas: ya están como date (sin hora) en columnas originales
# - numero_radicado: texto (string) y prefijo apóstrofe para Excel
# -------------------------------------------------------------------
df[col_radicado] = df["_rad_str"].astype("string")
if col_radicado in summary.columns:
    summary[col_radicado] = summary[col_radicado].astype("string")

# -------------------------------------------------------------------
# Guardar Excel con formato fecha y texto
# -------------------------------------------------------------------
print("Guardando Excel de salida...")
with pd.ExcelWriter(SALIDA, engine="openpyxl", datetime_format="YYYY-MM-DD", date_format="YYYY-MM-DD") as wr:
    # Escribir hojas (sin MESA_FECHA_REFERENCIA)
    df.to_excel(wr, index=False, sheet_name="Todos")
    df_pert.to_excel(wr, index=False, sheet_name="Registros_para_revisar")
    summary.to_excel(wr, index=False, sheet_name="Resumen_rad_duplicados")
    detalles_cols.to_excel(wr, index=False, sheet_name="Columnas_detectadas")

    wb = wr.book

    # === Formateo de fechas (sin hora) ===
    fmt_fecha = numbers.FORMAT_DATE_YYYYMMDD2  # yyyy-mm-dd

    def formatear_fechas(ws, nombres_columnas):
        header_idx = {cell.value: cell.column for cell in ws[1] if cell.value in nombres_columnas}
        for col_name, col_index in header_idx.items():
            col_letter = ws.cell(row=1, column=col_index).column_letter
            for r in range(2, ws.max_row + 1):
                ws[f"{col_letter}{r}"].number_format = fmt_fecha

    # Hoja "Todos": formatear las columnas de fecha originales y las internas
    ws = wb["Todos"]
    headers_todos = [cell.value for cell in ws[1]]
    fecha_cols_todos = [col_asigna, col_caso_crea] + \
                       [c for c in df.columns if c.startswith("_") and "fecha" in c.lower()]
    formatear_fechas(ws, set(fecha_cols_todos) & set(headers_todos))

    # Hoja "Registros_para_revisar": formatear fechas originales e internas si están
    ws = wb["Registros_para_revisar"]
    headers_rev = [cell.value for cell in ws[1]]
    fecha_cols_rev = [col_asigna, col_caso_crea] + \
                     [c for c in df_pert.columns if c.startswith("_") and "fecha" in c.lower()]
    formatear_fechas(ws, set(fecha_cols_rev) & set(headers_rev))

    # === Forzar numero_radicado como texto en Excel (prefijo apóstrofe) ===
    for sheet_name in ["Todos", "Resumen_rad_duplicados"]:
        ws = wb[sheet_name]
        headers = [cell.value for cell in ws[1]]
        if col_radicado in headers:
            # buscar letra de columna
            for cell in ws[1]:
                if cell.value == col_radicado:
                    col_letter = cell.column_letter
                    break
            for r in range(2, ws.max_row + 1):
                val = ws[f"{col_letter}{r}"].value
                if val is not None:
                    ws[f"{col_letter}{r}"].value = "'" + str(val)

print(f"== Listo: {SALIDA} ==")


