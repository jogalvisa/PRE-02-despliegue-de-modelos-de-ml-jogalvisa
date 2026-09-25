"""
Resultados de laboratorio normales
==================================

Replica en Python (pandas) la lógica de la query ResultadosNormales():

  1. Toma la fecha de hace N días (por defecto 3).
  2. Filtra los resultados validados ese día y de los códigos de servicio del proyecto.
  3. Descarta el examen completo (paciente + servicio + fecha_validacion + autorización)
     si AL MENOS UNO de sus componentes NO es 'NORMAL' o viene vacío (NULL).
  4. Deja una sola fila por examen (equivale al GROUP BY sin agregaciones).
  5. Marca interpretacion_resultado = 'NORMAL' y ordena por paciente, fecha y servicio.

Cómo correrlo en VS Code
------------------------
  pip install pandas openpyxl            # (+ pyodbc si vas a usar MODO = "sql")
  python resultados_normales.py          # o botón "Run Python File"

Elige la fuente de datos en la sección CONFIGURACIÓN:
  - "demo"  : datos de ejemplo embebidos (sirve para probar sin base de datos).
  - "excel" : lee un .xlsx o .csv exportado de ResultadosAyudasDiagnosticas.
  - "sql"   : ejecuta la query original contra SQL Server y además aplica la lógica en pandas.
"""

import datetime
from pathlib import Path

import pandas as pd

# ============================================================================
# CONFIGURACIÓN
# ============================================================================
MODO = "demo"  # "demo" | "excel" | "sql"

DIAS_ATRAS = 3
# Para reprocesar un día puntual pon la fecha aquí, p. ej. datetime.date(2026, 9, 21)
FECHA_FIJA = None

# Modo "excel": ruta del archivo con los datos crudos de la tabla
RUTA_ARCHIVO = Path(__file__).parent / "ResultadosAyudasDiagnosticas.xlsx"

# Modo "sql": cadena de conexión ODBC a SQL Server
CADENA_CONEXION = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=MI_SERVIDOR;"
    "DATABASE=MI_BASE;"
    "Trusted_Connection=yes;"
)

# Archivo de salida
RUTA_SALIDA = Path(__file__).parent / "resultados_normales_{fecha}.xlsx"

CODIGOS_SERVICIO = [
    '19303', '19490', '19934', '19915', '19940', '19299', '19290', '19805', '19958',
    '19313', '197321', '19792', '19933', '191701', '19516', '19283', '19177', '19332',
    '19522', '19827', '19505', '19157', '19224', '19966', '19780', '19749', '19891',
    '19534', '19493', '19855', '19492', '19140', '19821', '19964', '194925', '19285',
    '19330', '19331', '199151',
]

# Columnas que identifican un mismo examen (las del NOT EXISTS)
LLAVE_EXAMEN = ["numero_id_paciente", "codigo_servicio", "fecha_validacion", "numero_autorizacion"]

# Columnas de salida (las del SELECT / GROUP BY)
COLUMNAS_SALIDA = [
    "fecha_validacion",
    "numero_autorizacion",
    "tipo_id_paciente",
    "numero_id_paciente",
    "telefono_paciente",
    "Celular",
    "Direccion_Electronica",
    "codigo_servicio",
    "descripcion_servicio",
]


# ============================================================================
# FECHA Y QUERY ORIGINAL
# ============================================================================
def fecha_objetivo():
    """Fecha a procesar: FECHA_FIJA si está definida, si no hoy - DIAS_ATRAS."""
    if FECHA_FIJA is not None:
        return FECHA_FIJA
    return datetime.date.today() - datetime.timedelta(days=DIAS_ATRAS)


def ResultadosNormales(fecha=None):
    """Misma query del código original; la fecha y los códigos van como parámetros."""
    fecha = fecha or fecha_objetivo()
    marcadores = ",".join("?" for _ in CODIGOS_SERVICIO)
    query = f"""
    SELECT
        t.fecha_validacion,
        t.numero_autorizacion,
        t.tipo_id_paciente,
        t.numero_id_paciente,
        t.telefono_paciente,
        t.Celular,
        t.Direccion_Electronica,
        t.codigo_servicio,
        t.descripcion_servicio,
        'NORMAL' AS interpretacion_resultado
    FROM BI_Salud.ResultadosAyudasDiagnosticas t
    WHERE CAST(t.fecha_validacion AS DATE) = ?
      AND t.codigo_servicio IN ({marcadores})
      AND NOT EXISTS (
            SELECT 1
            FROM BI_Salud.ResultadosAyudasDiagnosticas t2
            WHERE t2.numero_id_paciente = t.numero_id_paciente
              AND t2.codigo_servicio = t.codigo_servicio
              AND t2.fecha_validacion = t.fecha_validacion
              AND t2.numero_autorizacion = t.numero_autorizacion
              AND (t2.interpretacion_resultado <> 'NORMAL' OR t2.interpretacion_resultado IS NULL)
      )
    GROUP BY
        t.tipo_id_paciente,
        t.numero_id_paciente,
        t.telefono_paciente,
        t.Celular,
        t.Direccion_Electronica,
        t.fecha_validacion,
        t.codigo_servicio,
        t.descripcion_servicio,
        t.numero_autorizacion
    ORDER BY
        t.numero_id_paciente,
        t.fecha_validacion,
        t.codigo_servicio
    """
    parametros = [fecha.isoformat(), *CODIGOS_SERVICIO]
    return query, parametros


# ============================================================================
# MISMA LÓGICA EN PANDAS
# ============================================================================
def filtrar_resultados_normales(df, fecha):
    """Aplica sobre un DataFrame crudo la misma lógica de la query."""
    df = df.copy()
    df["fecha_validacion"] = pd.to_datetime(df["fecha_validacion"])
    df["codigo_servicio"] = df["codigo_servicio"].astype(str).str.strip()

    # WHERE CAST(fecha_validacion AS DATE) = fecha AND codigo_servicio IN (...)
    base = df[
        (df["fecha_validacion"].dt.date == fecha)
        & (df["codigo_servicio"].isin(CODIGOS_SERVICIO))
    ]

    # NOT EXISTS: exámenes con algún componente distinto de NORMAL o vacío.
    # strip/upper imita la collation por defecto de SQL Server ('normal ' = 'NORMAL').
    interpretacion = df["interpretacion_resultado"].astype("string").str.strip().str.upper()
    no_normal = interpretacion.ne("NORMAL") | interpretacion.isna()
    # En SQL, NULL = NULL no coincide, así que filas con la llave incompleta nunca se descartan.
    examenes_malos = (
        df.loc[no_normal, LLAVE_EXAMEN]
        .dropna()
        .drop_duplicates()
        .assign(_descartar=True)
    )
    base = base.merge(examenes_malos, on=LLAVE_EXAMEN, how="left")
    base = base[base["_descartar"].isna()]

    # GROUP BY sin agregaciones = una fila por combinación distinta
    resultado = base[COLUMNAS_SALIDA].drop_duplicates()
    resultado["interpretacion_resultado"] = "NORMAL"

    # ORDER BY
    return resultado.sort_values(
        ["numero_id_paciente", "fecha_validacion", "codigo_servicio"]
    ).reset_index(drop=True)


# ============================================================================
# FUENTES DE DATOS
# ============================================================================
def datos_demo(fecha):
    """Datos de ejemplo con los casos que la lógica debe resolver."""
    hora = datetime.datetime.combine(fecha, datetime.time(9, 30))
    otro_dia = hora - datetime.timedelta(days=1)
    contacto = {"tipo_id_paciente": "CC", "telefono_paciente": "6041234567",
                "Celular": "3001234567", "Direccion_Electronica": "paciente@correo.com"}
    filas = [
        # 111: perfil con todos los componentes NORMAL -> SALE (1 sola fila)
        ("111", "A1", "19303", "Perfil lipídico", "NORMAL", hora),
        ("111", "A1", "19303", "Perfil lipídico", "NORMAL", hora),
        ("111", "A1", "19303", "Perfil lipídico", "normal ", hora),
        # 222: un componente ALTO -> NO sale
        ("222", "B1", "19303", "Perfil lipídico", "NORMAL", hora),
        ("222", "B1", "19303", "Perfil lipídico", "ALTO", hora),
        # 333: componente sin interpretación (NULL) -> NO sale
        ("333", "C1", "19490", "Hemograma", "NORMAL", hora),
        ("333", "C1", "19490", "Hemograma", None, hora),
        # 444: normal pero código fuera de la lista -> NO sale
        ("444", "D1", "99999", "Otro examen", "NORMAL", hora),
        # 555: normal pero validado otro día -> NO sale
        ("555", "E1", "19934", "Glicemia", "NORMAL", otro_dia),
        # 666: dos exámenes; uno normal -> SALE solo 19934, el 19915 tiene BAJO
        ("666", "F1", "19934", "Glicemia", "NORMAL", hora),
        ("666", "F2", "19915", "Creatinina", "BAJO", hora),
    ]
    columnas = ["numero_id_paciente", "numero_autorizacion", "codigo_servicio",
                "descripcion_servicio", "interpretacion_resultado", "fecha_validacion"]
    return pd.DataFrame(filas, columns=columnas).assign(**contacto)


def leer_archivo(ruta):
    ruta = Path(ruta)
    if ruta.suffix.lower() == ".csv":
        return pd.read_csv(ruta, dtype=str)
    return pd.read_excel(ruta, dtype=str)


def leer_sql(fecha):
    """Ejecuta la query original en SQL Server y devuelve el resultado."""
    import pyodbc  # solo se necesita en modo "sql"

    query, parametros = ResultadosNormales(fecha)
    with pyodbc.connect(CADENA_CONEXION) as conexion:
        cursor = conexion.cursor()
        cursor.execute(query, parametros)
        columnas = [c[0] for c in cursor.description]
        return pd.DataFrame.from_records(cursor.fetchall(), columns=columnas)


# ============================================================================
# EJECUCIÓN
# ============================================================================
def guardar(df, fecha):
    ruta = Path(str(RUTA_SALIDA).format(fecha=fecha.isoformat()))
    try:
        df.to_excel(ruta, index=False)
    except ImportError:  # sin openpyxl, se guarda en CSV
        ruta = ruta.with_suffix(".csv")
        df.to_csv(ruta, index=False, encoding="utf-8-sig")
    return ruta


def main():
    fecha = fecha_objetivo()
    print(f"Fecha procesada: {fecha}  |  Modo: {MODO}")

    if MODO == "sql":
        resultado = leer_sql(fecha)
    else:
        crudo = datos_demo(fecha) if MODO == "demo" else leer_archivo(RUTA_ARCHIVO)
        print(f"Filas crudas: {len(crudo)}")
        resultado = filtrar_resultados_normales(crudo, fecha)

    print(f"Exámenes 100% normales: {len(resultado)}")
    print(f"Pacientes distintos:    {resultado['numero_id_paciente'].nunique()}\n")
    with pd.option_context("display.max_columns", None, "display.width", 200):
        print(resultado.to_string(index=False) if len(resultado) else "(sin resultados)")

    ruta = guardar(resultado, fecha)
    print(f"\nArchivo generado: {ruta}")


if __name__ == "__main__":
    main()
