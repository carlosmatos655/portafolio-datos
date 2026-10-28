"""Genera el proyecto Power BI (PBIP: modelo TMDL + reporte PBIR) del caso Compras públicas del Perú.

Identidad "Noche Paracas" (modo oscuro). Un mismo plano de posiciones produce:
  - los fondos de página (HTML -> PNG con Edge headless): barra lateral, títulos, marcos de tarjetas, íconos;
  - los visuales de Power BI, que se colocan encima de esos marcos.

El año se elige en una tabla desconectada (sel_anio): así los gráficos con eje de año o mes
pueden mostrar el año anterior junto al elegido. Sin selección se usa el último año completo.

Uso: python powerbi/build_pbip.py            (reescribe ComprasPublicasPeru.* y los fondos)
     python powerbi/build_pbip.py --no-png   (sin regenerar los PNG)
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import time
import uuid
from pathlib import Path

HERE = Path(__file__).resolve().parent
NAME = "ComprasPublicasPeru"
DATA = (HERE.parent / "data" / "processed").resolve()
SM = HERE / f"{NAME}.SemanticModel"
RP = HERE / f"{NAME}.Report"
BG = HERE / "backgrounds"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
W, H = 1280, 720

SCH = "https://developer.microsoft.com/json-schemas/fabric"
V_VISUAL, V_PAGE, V_REPORT = "2.4.0", "2.0.0", "3.0.0"

# Identidad "Noche Paracas" (paleta categórica validada en fondo oscuro: CVD y contraste)
C = {"bg": "#13121D", "side": "#0C0B14", "card": "#1B1A29", "track": "#262540", "line": "#2C2A42", "sunk": "#16152A",
     "ink": "#ECEAF4", "muted": "#A9A6BF", "faint": "#74718F",
     "indigo": "#7C82D4", "oro": "#BA8620", "turquesa": "#139AB3", "terracota": "#D25F43",
     "neg": "#F08BA0", "pos": "#7FD19A", "old": "#4A4866", "dim": "#3E3C63", "terra_dim": "#7C3F37"}
FONT_TXT = "Segoe UI"
FONT_NUM = "Bahnschrift"

ICON = {
    "grid": "M4 4h7v7H4z M13 4h7v4h-7z M13 10h7v10h-7z M4 13h7v7H4z",
    "balanza": "M12 4v16 M6 20h12 M5 8h14 M5 8l-3 6a3 3 0 0 0 6 0z M19 8l-3 6a3 3 0 0 0 6 0z",
    "edificio": "M4 21V8l8-5 8 5v13 M3 21h18 M9 21v-6h6v6 M8 10h2 M14 10h2",
    "monedas": "M12 3c-4.4 0-8 1.3-8 3v12c0 1.7 3.6 3 8 3s8-1.3 8-3V6c0-1.7-3.6-3-8-3z M4 6c0 1.7 3.6 3 8 3s8-1.3 8-3 M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3",
    "documento": "M7 3h7l5 5v13H7z M14 3v5h5 M10 13h6 M10 17h6",
    "check": "M12 3a9 9 0 1 0 0 18a9 9 0 1 0 0-18 M8 12l3 3 5-6",
    "reloj": "M12 3a9 9 0 1 0 0 18a9 9 0 1 0 0-18 M12 7v5l3 2",
    "obra": "M3 21h18 M5 21V11l7-4 7 4v10 M9 21v-5h6v5 M12 3v4",
    "servicio": "M12 8a4 4 0 1 0 0 8a4 4 0 1 0 0-8 M12 2v3 M12 19v3 M2 12h3 M19 12h3 M4.9 4.9l2.1 2.1 M17 17l2.1 2.1 M4.9 19.1l2.1-2.1 M17 7l2.1-2.1",
    "bien": "M4 8l8-4 8 4v8l-8 4-8-4z M4 8l8 4 8-4 M12 12v8",
    "nacional": "M4 21V10l8-6 8 6v11 M8 21v-7 M12 21v-7 M16 21v-7 M3 21h18",
    "local": "M3 21h18 M6 21V9h6v12 M12 21V5h6v16 M8 12h2 M8 16h2 M14 8h2 M14 12h2 M14 16h2",
    "pin": "M12 21s-7-6-7-11a7 7 0 0 1 14 0c0 5-7 11-7 11z M12 7.5a2.5 2.5 0 1 0 0 5a2.5 2.5 0 1 0 0-5",
    "flecha": "M4 12h14 M13 6l6 6-6 6",
    "persona": "M12 4a3.5 3.5 0 1 0 0 7a3.5 3.5 0 1 0 0-7 M5 20c0-3.9 3.1-7 7-7s7 3.1 7 7",
    "grupo": "M9 5a3 3 0 1 0 0 6a3 3 0 1 0 0-6 M3 19c0-3.3 2.7-6 6-6s6 2.7 6 6 M16 5.5a2.5 2.5 0 1 1 0 5 M17 13c2.5.5 4 2.8 4 5.5",
    "alerta": "M12 3l9 16H3z M12 10v4 M12 17h.01",
    "manos": "M3 12l4-4 4 3 4-3 6 5 M7 8l-3 7 M17 8l3 7 M9 14l2 2 M12 13l2 2",
    "barras": "M3 20h18 M5 20v-6h4v6 M10 20V9h4v11 M15 20v-8h4v8",
    "candado": "M6 11h12v9H6z M8.5 11V8a3.5 3.5 0 0 1 7 0v3 M12 14.5v2.5",
    "chispa": "M12 3v4 M12 17v4 M3 12h4 M17 12h4 M6 6l2.5 2.5 M15.5 15.5L18 18 M6 18l2.5-2.5 M15.5 8.5L18 6",
}

# ---------------------------------------------------------------- plano general
SIDE_W = 218
X0, X1 = 242, 1256                       # contenido principal
KPI_Y, KPI_H, KPI_W, KPI_STEP = 80, 136, 244, 256
PANEL_Y, PANEL_B = 228, 680              # paneles: de y=228 a y=680

PAGES = [
    {"name": "panorama", "display": "Panorama", "icon": "grid", "sub": "Cuánto y en qué",
     "kicker": "01 · PANORAMA", "title": "En 2025 el Estado adjudicó S/ 84.8 mil millones; más de la mitad, en obras",
     "footer": "Fuente: OECE, SEACE en estándar OCDS (Open Contracting Data Registry, CC BY 4.0). Procesos por fecha de convocatoria; "
               "un año en curso se compara con el mismo periodo del año anterior."},
    {"name": "competencia", "display": "Competencia", "icon": "balanza", "sub": "Cuántos compiten",
     "kicker": "02 · COMPETENCIA", "title": "9 de cada 10 contrataciones directas tienen un solo participante",
     "footer": "Fuente: OECE, SEACE en estándar OCDS (Open Contracting Data Registry, CC BY 4.0). Tipo de procedimiento según el estándar OCDS."},
    {"name": "proveedores", "display": "Proveedores", "icon": "edificio", "sub": "Quién gana y dónde",
     "kicker": "03 · PROVEEDORES Y TERRITORIO", "title": "Los consorcios se llevan 6 de cada 10 soles de las compras públicas",
     "footer": "Fuente: OECE, SEACE en estándar OCDS (Open Contracting Data Registry, CC BY 4.0). En consorcios el monto se reparte entre sus miembros."},
]

# ---------------------------------------------------------------- modelo semántico (TMDL)

TABLES = {
    "dim_fecha": [("fecha", "dateTime", "date"), ("anio", "int64", "Int64.Type"), ("mes", "int64", "Int64.Type"),
                  ("mes_nombre", "string", "text"), ("trimestre", "string", "text")],
    "dim_entidad": [("entidad_id", "string", "text"), ("entidad", "string", "text"), ("departamento", "string", "text"),
                    ("nivel_gobierno", "string", "text")],
    "dim_metodo": [("metodo", "string", "text"), ("tipo_ocds", "string", "text")],
    "dim_categoria": [("categoria_cod", "string", "text"), ("categoria", "string", "text")],
    "dim_proveedor": [("proveedor_id", "string", "text"), ("proveedor", "string", "text"), ("tipo", "string", "text")],
    "fact_proceso": [("ocid", "string", "text"), ("fecha", "dateTime", "date"), ("entidad_id", "string", "text"),
                     ("metodo", "string", "text"), ("categoria_cod", "string", "text"), ("valor_referencial_pen", "double", "number"),
                     ("n_participantes", "double", "number"), ("monto_adjudicado_pen", "double", "number"),
                     ("dias_a_buena_pro", "double", "number"), ("adjudicado", "int64", "Int64.Type"), ("monto_atipico", "int64", "Int64.Type")],
    "fact_adjudicacion": [("ocid", "string", "text"), ("proveedor_id", "string", "text"), ("fecha_buena_pro", "dateTime", "date"),
                          ("monto_pen", "double", "number")],
}
# columnas de orden: (tabla, columna ordenada, columna de orden, expresión DAX si es calculada)
SORT_BY = {("dim_fecha", "mes_nombre"): "mes", ("dim_categoria", "categoria"): "orden"}
CALC_COLS = {"dim_categoria": [("orden", "int64", 'SWITCH ( [categoria_cod], "works", 1, "services", 2, 3 )')]}
RELS = [("fact_proceso", "fecha", "dim_fecha", "fecha"), ("fact_proceso", "entidad_id", "dim_entidad", "entidad_id"),
        ("fact_proceso", "metodo", "dim_metodo", "metodo"), ("fact_proceso", "categoria_cod", "dim_categoria", "categoria_cod"),
        ("fact_adjudicacion", "ocid", "fact_proceso", "ocid"), ("fact_adjudicacion", "proveedor_id", "dim_proveedor", "proveedor_id")]

SEL_ANIO = """
VAR _anios = SELECTCOLUMNS ( DISTINCT ( dim_fecha[anio] ), "Anio", dim_fecha[anio] )
RETURN
    ADDCOLUMNS (
        _anios,
        "Etiqueta",
            VAR _a = [Anio]
            VAR _m = CALCULATE ( MAX ( dim_fecha[mes] ), REMOVEFILTERS (), dim_fecha[anio] = _a )
            RETURN FORMAT ( _a, "0" ) & IF ( _m < 12, "*", "" )
    )"""


def yr(expr: str) -> str:
    """Envuelve una expresión para evaluarla en el año de análisis."""
    return f"VAR _y = [Anio] RETURN CALCULATE ( {expr}, REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y )"


def prev(expr: str) -> str:
    """Mismo periodo del año anterior (año completo si el de análisis está completo)."""
    return (f"VAR _y = [Anio] VAR _corte = [Fecha corte] RETURN CALCULATE ( {expr}, REMOVEFILTERS ( dim_fecha ), "
            f"dim_fecha[anio] = _y - 1, dim_fecha[fecha] <= EDATE ( _corte, -12 ) )")


def pill_var(m: str) -> str:
    return (f'VAR _v = {m} RETURN IF ( ISBLANK ( _v ), BLANK (), IF ( _v < 0, "▼ ", "▲ " ) & FORMAT ( ABS ( _v ), "0.0%" ) '
            f'& " vs " & [Anio anterior txt] )')


def var_color(m: str) -> str:
    return f'IF ( {m} < 0, "{C["neg"]}", "{C["pos"]}" )'


# (nombre, expresión DAX, formato, carpeta)
MEASURES = [
    # --- contexto
    ("Anio", """
VAR _completo =
    CALCULATE (
        MAXX ( FILTER ( VALUES ( dim_fecha[anio] ), CALCULATE ( DISTINCTCOUNT ( dim_fecha[mes] ) ) = 12 ), dim_fecha[anio] ),
        REMOVEFILTERS ()
    )
RETURN IF ( HASONEVALUE ( sel_anio[Anio] ), VALUES ( sel_anio[Anio] ), _completo )""", "0", "0 Contexto"),
    ("Fecha corte", "VAR _y = [Anio] RETURN CALCULATE ( MAX ( dim_fecha[fecha] ), REMOVEFILTERS (), dim_fecha[anio] = _y )", "yyyy-mm-dd", "0 Contexto"),
    ("Anio parcial", "MONTH ( [Fecha corte] ) < 12", "", "0 Contexto"),
    ("Anio anterior txt", 'FORMAT ( [Anio] - 1, "0" ) & IF ( [Anio parcial], "*", "" )', "", "0 Contexto"),
    ("Base procesos", "CALCULATE ( COUNTROWS ( fact_proceso ), REMOVEFILTERS () )", "#,0", "0 Contexto"),
    ("Base adjudicaciones", "CALCULATE ( COUNTROWS ( fact_adjudicacion ), REMOVEFILTERS () )", "#,0", "0 Contexto"),
    ("Base proveedores", "CALCULATE ( COUNTROWS ( dim_proveedor ), REMOVEFILTERS () )", "#,0", "0 Contexto"),
    ("Base entidades", "CALCULATE ( COUNTROWS ( dim_entidad ), REMOVEFILTERS () )", "#,0", "0 Contexto"),

    # --- procesos (fact_proceso, por fecha de convocatoria)
    ("Procesos", "COUNTROWS ( fact_proceso )", "#,0", "1 Panorama"),
    ("Monto", "DIVIDE ( SUM ( fact_proceso[monto_adjudicado_pen] ), 1e9 )", "#,0.0", "1 Panorama"),
    ("Procesos anio", yr("[Procesos]"), "#,0", "1 Panorama"),
    ("Procesos anio ant", prev("[Procesos]"), "#,0", "1 Panorama"),
    ("Monto anio", yr("[Monto]"), "#,0.0", "1 Panorama"),
    ("Monto anio ant", prev("[Monto]"), "#,0.0", "1 Panorama"),
    ("Var monto", "IF ( NOT ISBLANK ( [Monto anio ant] ), DIVIDE ( [Monto anio], [Monto anio ant] ) - 1 )", "+0.0%;-0.0%", "1 Panorama"),
    ("Var procesos", "IF ( NOT ISBLANK ( [Procesos anio ant] ), DIVIDE ( [Procesos anio], [Procesos anio ant] ) - 1 )", "+0.0%;-0.0%", "1 Panorama"),
    ("Pill monto", pill_var("[Var monto]"), "", "1 Panorama"),
    ("Pill procesos", pill_var("[Var procesos]"), "", "1 Panorama"),
    ("Color var monto", var_color("[Var monto]"), "", "1 Panorama"),
    ("Color var procesos", var_color("[Var procesos]"), "", "1 Panorama"),
    ("Adjudicados anio", yr("SUM ( fact_proceso[adjudicado] )"), "#,0", "1 Panorama"),
    ("% adjudicados", "DIVIDE ( [Adjudicados anio], [Procesos anio] )", "0.0%", "1 Panorama"),
    ("% adjudicados resto", "1 - [% adjudicados]", "0.0%", "1 Panorama"),
    ("Pill adjudicados", 'FORMAT ( [Adjudicados anio], "#,0" ) & " con buena pro"', "", "1 Panorama"),
    ("Dias buena pro", yr("AVERAGE ( fact_proceso[dias_a_buena_pro] )"), "0.0", "1 Panorama"),
    ("Dias escala", "MIN ( DIVIDE ( [Dias buena pro], 60 ), 1 )", "0.0%", "1 Panorama"),
    ("Dias escala resto", "1 - [Dias escala]", "0.0%", "1 Panorama"),
    ("Monto comparativo", """
VAR _y = [Anio]
VAR _a = SELECTEDVALUE ( dim_fecha[anio] )
RETURN SWITCH ( _a, _y, [Monto anio], _y - 1, [Monto anio ant] )""", "#,0.0", "1 Panorama"),
    ("Color comparativo", f'IF ( SELECTEDVALUE ( dim_fecha[anio] ) = [Anio], "{C["indigo"]}", "{C["old"]}" )', "", "1 Panorama"),
    ("Procesos mes", "VAR _y = [Anio] RETURN CALCULATE ( [Procesos], dim_fecha[anio] = _y )", "#,0", "1 Panorama"),
    ("% categoria", "DIVIDE ( [Monto anio], CALCULATE ( [Monto anio], REMOVEFILTERS ( dim_categoria ) ) )", "0.0%", "1 Panorama"),
    ("Color categoria", f'SWITCH ( SELECTEDVALUE ( dim_categoria[categoria] ), "Obras", "{C["terracota"]}", "Servicios", "{C["turquesa"]}", "Bienes", "{C["oro"]}", "{C["indigo"]}" )', "", "1 Panorama"),
    *[(f"Var {c}", f'CALCULATE ( [Var monto], dim_categoria[categoria] = "{c}" )', "+0.0%;-0.0%", "1 Panorama") for c in ("Obras", "Servicios", "Bienes")],
    *[(f"Color var {c}", var_color(f"[Var {c}]"), "", "1 Panorama") for c in ("Obras", "Servicios", "Bienes")],
    *[(f"% {c}", f'CALCULATE ( [% categoria], dim_categoria[categoria] = "{c}" )', "0.0%", "1 Panorama") for c in ("Obras", "Servicios", "Bienes")],
    ("Insight categoria", """
VAR _t = FILTER ( ADDCOLUMNS ( ALL ( dim_categoria[categoria] ), "v", [Var monto] ), NOT ISBLANK ( [v] ) )
VAR _min = MINX ( _t, [v] )
VAR _max = MAXX ( _t, [v] )
VAR _cmin = MAXX ( FILTER ( _t, [v] = _min ), dim_categoria[categoria] )
VAR _cmax = MAXX ( FILTER ( _t, [v] = _max ), dim_categoria[categoria] )
RETURN
    IF ( ISEMPTY ( _t ), "Primer año de la serie: no hay año anterior para comparar",
    IF ( _min < 0,
        "Mayor caída: " & _cmin & ", " & FORMAT ( _min, "0.0%" ) & " frente a " & [Anio anterior txt],
        "Mayor alza: " & _cmax & ", " & FORMAT ( _max, "+0.0%" ) & " frente a " & [Anio anterior txt] ) )""", "", "1 Panorama"),
    ("% convocatorias mes", """
VAR _anios = CALCULATETABLE ( FILTER ( VALUES ( dim_fecha[anio] ), CALCULATE ( DISTINCTCOUNT ( dim_fecha[mes] ) ) = 12 ), REMOVEFILTERS ( dim_fecha ) )
VAR _n = CALCULATE ( [Procesos], REMOVEFILTERS ( dim_fecha[anio] ), TREATAS ( _anios, dim_fecha[anio] ) )
VAR _t = CALCULATE ( [Procesos], REMOVEFILTERS ( dim_fecha ), TREATAS ( _anios, dim_fecha[anio] ) )
RETURN DIVIDE ( _n, _t )""", "0.0%", "1 Panorama"),
    ("Color mes", f"""
VAR _r = RANKX ( ALL ( dim_fecha[mes], dim_fecha[mes_nombre] ), [% convocatorias mes] )
RETURN IF ( _r = 12, "{C["terracota"]}", IF ( _r <= 3, "{C["indigo"]}", "{C["dim"]}" ) )""", "", "1 Panorama"),
    ("% nivel", "DIVIDE ( [Monto anio], CALCULATE ( [Monto anio], REMOVEFILTERS ( dim_entidad[nivel_gobierno] ) ) )", "0.0%", "1 Panorama"),
    *[(f"% {n}", f'CALCULATE ( [% nivel], dim_entidad[nivel_gobierno] = "Gobierno {n.lower()}" )', "0.0%", "1 Panorama")
      for n in ("Nacional", "Local", "Regional")],

    # --- competencia
    ("% un participante", """
DIVIDE (
    CALCULATE ( COUNTROWS ( fact_proceso ), fact_proceso[n_participantes] = 1 ),
    CALCULATE ( COUNTROWS ( fact_proceso ), NOT ISBLANK ( fact_proceso[n_participantes] ) )
)""", "0.0%", "2 Competencia"),
    ("% un participante anio", yr("[% un participante]"), "0.0%", "2 Competencia"),
    ("% un participante resto", "1 - [% un participante anio]", "0.0%", "2 Competencia"),
    ("Monto directo", 'CALCULATE ( [Monto anio], dim_metodo[tipo_ocds] = "Directo" )', "#,0.00", "2 Competencia"),
    ("% directo", "DIVIDE ( [Monto directo], [Monto anio] )", "0.0%", "2 Competencia"),
    ("% directo resto", "1 - [% directo]", "0.0%", "2 Competencia"),
    ("Pill directo", 'FORMAT ( [% directo], "0.0%" ) & " del monto"', "", "2 Competencia"),
    ("% directas un participante", 'CALCULATE ( [% un participante anio], dim_metodo[tipo_ocds] = "Directo" )', "0.0%", "2 Competencia"),
    ("% directas un participante resto", "1 - [% directas un participante]", "0.0%", "2 Competencia"),
    ("Pill directas", '"de " & FORMAT ( CALCULATE ( [Procesos anio], dim_metodo[tipo_ocds] = "Directo" ), "#,0" ) & " directas"', "", "2 Competencia"),
    ("Participantes abiertos", yr('CALCULATE ( AVERAGE ( fact_proceso[n_participantes] ), dim_metodo[tipo_ocds] = "Abierto" )'), "0.0", "2 Competencia"),
    ("Participantes directos", yr('CALCULATE ( AVERAGE ( fact_proceso[n_participantes] ), dim_metodo[tipo_ocds] = "Directo" )'), "0.0", "2 Competencia"),
    ("Participantes escala", "MIN ( DIVIDE ( [Participantes abiertos], 20 ), 1 )", "0.0%", "2 Competencia"),
    ("Participantes escala resto", "1 - [Participantes escala]", "0.0%", "2 Competencia"),
    ("Ratio participantes", '"Un proceso abierto reúne " & FORMAT ( DIVIDE ( [Participantes abiertos], [Participantes directos] ), "0" ) & " veces más participantes que uno directo"', "", "2 Competencia"),
    ("% un participante metodo", "IF ( [Procesos anio] >= 500, [% un participante anio] )", "0.0%", "2 Competencia"),
    ("Color metodo", f'IF ( [% un participante metodo] >= 0.5, "{C["terracota"]}", "{C["indigo"]}" )', "", "2 Competencia"),
    ("Directo M", "[Monto directo] * 1000", "#,0.0", "2 Competencia"),
    ("Procesos directos", 'IF ( [Monto directo] > 0, CALCULATE ( [Procesos anio], dim_metodo[tipo_ocds] = "Directo" ) )', "#,0", "2 Competencia"),

    # --- proveedores (fact_adjudicacion: una fila por ganador; consorcios repartidos)
    ("Adjudicado", "DIVIDE ( SUM ( fact_adjudicacion[monto_pen] ), 1e9 )", "#,0.00", "3 Proveedores"),
    ("Adjudicado anio", yr("[Adjudicado]"), "#,0.00", "3 Proveedores"),
    ("% tipo ganador", "DIVIDE ( [Adjudicado anio], CALCULATE ( [Adjudicado anio], REMOVEFILTERS ( dim_proveedor ) ) )", "0.0%", "3 Proveedores"),
    *[(f"% {k}", f'CALCULATE ( [% tipo ganador], dim_proveedor[tipo] = "{t}" )', "0.0%", "3 Proveedores")
      for k, t in (("Consorcio", "Consorcio"), ("Empresa", "Empresa"), ("Persona natural", "Persona natural"), ("Otro", "Otro (no domiciliado u otro)"))],
    *[(f"Promedio {k}", 'VAR _v = ' + yr(f'DIVIDE ( SUM ( fact_adjudicacion[monto_pen] ), COUNTROWS ( fact_adjudicacion ) ), dim_proveedor[tipo] = "{t}"')
       + ' RETURN "S/ " & FORMAT ( _v / 1e6, "0.00" ) & " M prom."', "", "3 Proveedores")
      for k, t in (("Consorcio", "Consorcio"), ("Empresa", "Empresa"), ("Persona natural", "Persona natural"), ("Otro", "Otro (no domiciliado u otro)"))],
    ("% Consorcio resto", "1 - [% Consorcio]", "0.0%", "3 Proveedores"),
    ("Pill consorcio", 'FORMAT ( CALCULATE ( DIVIDE ( CALCULATE ( [Adjudicado anio], dim_categoria[categoria] = "Obras" ), [Adjudicado anio] ), dim_proveedor[tipo] = "Consorcio" ), "0%" ) & " en obras"', "", "3 Proveedores"),
    ("Ganadores", yr("DISTINCTCOUNT ( fact_adjudicacion[proveedor_id] )"), "#,0", "3 Proveedores"),
    ("Pill ganadores", '"año " & [Anio] & IF ( [Anio parcial], "*", "" )', "", "3 Proveedores"),
    ("Empresa M", 'CALCULATE ( [Adjudicado anio], dim_proveedor[tipo] = "Empresa" ) * 1000', "#,0.0", "3 Proveedores"),
    ("% Top 10 empresas", """
VAR _y = [Anio]
VAR _t = CALCULATE ( [Adjudicado anio], REMOVEFILTERS ( dim_proveedor ) )
VAR _emp = CALCULATETABLE ( VALUES ( fact_adjudicacion[proveedor_id] ), dim_proveedor[tipo] = "Empresa", REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y )
RETURN DIVIDE ( SUMX ( TOPN ( 10, _emp, [Adjudicado anio] ), [Adjudicado anio] ), _t )""", "0.0%", "3 Proveedores"),
    ("% Top 10 empresas resto", "1 - [% Top 10 empresas]", "0.0%", "3 Proveedores"),
    ("HHI proveedores", """
VAR _y = [Anio]
VAR _t = CALCULATE ( [Adjudicado anio], REMOVEFILTERS ( dim_proveedor ) )
VAR _p = CALCULATETABLE ( VALUES ( fact_adjudicacion[proveedor_id] ), REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y )
RETURN SUMX ( _p, ( 100 * DIVIDE ( [Adjudicado anio], _t ) ) ^ 2 )""", "#,0.0", "3 Proveedores"),
    ("Pill HHI", '"HHI " & FORMAT ( [HHI proveedores], "#,0.0" )', "", "3 Proveedores"),
    ("% Lima", 'DIVIDE ( CALCULATE ( [Monto anio], dim_entidad[departamento] = "Lima" ), CALCULATE ( [Monto anio], REMOVEFILTERS ( dim_entidad[departamento] ) ) )', "0.0%", "3 Proveedores"),
    ("% Lima resto", "1 - [% Lima]", "0.0%", "3 Proveedores"),
    ("Departamento top 10", "IF ( RANKX ( ALLSELECTED ( dim_entidad[departamento] ), [Monto anio] ) <= 10, [Monto anio] )", "#,0.00", "3 Proveedores"),
    ("Color departamento", f'IF ( RANKX ( ALLSELECTED ( dim_entidad[departamento] ), [Monto anio] ) = 1, "{C["terracota"]}", "{C["terra_dim"]}" )', "", "3 Proveedores"),
]


def tmdl_table(name: str, cols: list[tuple[str, str, str]]) -> str:
    out = [f"table {name}", ""]
    for col, dtype, _ in cols:
        out += [f"\tcolumn {col}", f"\t\tdataType: {dtype}"]
        if dtype == "dateTime":
            out.append("\t\tformatString: yyyy-mm-dd")
        out += ["\t\tsummarizeBy: none", f"\t\tsourceColumn: {col}"]
        if (name, col) in SORT_BY:
            out.append(f"\t\tsortByColumn: {SORT_BY[(name, col)]}")
        out.append("")
    for col, dtype, expr in CALC_COLS.get(name, []):
        out += [f"\tcolumn {col} = {expr}", f"\t\tdataType: {dtype}", "\t\tisHidden", "\t\tsummarizeBy: none", ""]
    types = ", ".join(f'{{"{c}", type {t}}}' if t in ("date", "text", "number") else f'{{"{c}", {t}}}' for c, _, t in cols)
    out += [
        f"\tpartition {name} = m",
        "\t\tmode: import",
        "\t\tsource =",
        "\t\t\t\tlet",
        f'\t\t\t\t    Origen = Csv.Document ( File.Contents ( RutaDatos & "{name}.csv" ), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv] ),',
        "\t\t\t\t    Encabezados = Table.PromoteHeaders ( Origen, [PromoteAllScalars = true] ),",
        f'\t\t\t\t    Tipado = Table.TransformColumnTypes ( Encabezados, {{{types}}}, "en-US" )',
        "\t\t\t\tin",
        "\t\t\t\t    Tipado",
        "",
    ]
    return "\n".join(out)


def tmdl_sel_anio() -> str:
    out = ["table sel_anio", "",
           "\tcolumn Anio", "\t\tdataType: int64", "\t\tisNameInferred", "\t\tformatString: 0", "\t\tsummarizeBy: none", "\t\tsourceColumn: [Anio]", "",
           "\tcolumn Etiqueta", "\t\tdataType: string", "\t\tisNameInferred", "\t\tsummarizeBy: none", "\t\tsourceColumn: [Etiqueta]",
           "\t\tsortByColumn: Anio", "",
           "\tpartition sel_anio = calculated", "\t\tmode: import", "\t\tsource ="]
    out += [f"\t\t\t\t{ln}" for ln in SEL_ANIO.strip().splitlines()]
    out.append("")
    return "\n".join(out)


def tmdl_measures() -> str:
    out = ["table Medidas", "",
           "\tcolumn Columna1", "\t\tdataType: int64", "\t\tisHidden", "\t\tsummarizeBy: none", "\t\tsourceColumn: [Columna1]", ""]
    for name, expr, fmt, folder in MEASURES:
        lines = expr.strip().splitlines()
        if len(lines) == 1:
            out.append(f"\tmeasure '{name}' = {lines[0]}")
        else:
            out.append(f"\tmeasure '{name}' =")
            out += [f"\t\t\t{ln}" for ln in lines]
        if fmt:
            out.append(f"\t\tformatString: {fmt}")
        out += [f"\t\tdisplayFolder: {folder}", ""]
    out += ["\tpartition Medidas = calculated", "\t\tmode: import", "\t\tsource = ROW ( \"Columna1\", 0 )", ""]
    return "\n".join(out)


def write_model() -> None:
    if SM.exists():
        shutil.rmtree(SM)
    d = SM / "definition"
    (d / "tables").mkdir(parents=True)
    (SM / "definition.pbism").write_text(json.dumps({"$schema": f"{SCH}/item/semanticModel/definitionProperties/1.0.0/schema.json",
                                                     "version": "4.2", "settings": {}}, indent=2), encoding="utf-8")
    (d / "database.tmdl").write_text("database\n\tcompatibilityLevel: 1601\n", encoding="utf-8")
    refs = "\n".join(f"ref table {t}" for t in [*TABLES, "sel_anio", "Medidas"])
    (d / "model.tmdl").write_text(
        "model Model\n\tculture: en-US\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\n\tsourceQueryCulture: es-PE\n"
        "\tdataAccessOptions\n\t\tlegacyRedirects\n\t\treturnErrorValuesAsNull\n\n"
        "annotation PBI_QueryOrder = [\"RutaDatos\"]\n\nannotation __PBI_TimeIntelligenceEnabled = 0\n\n" + refs + "\n", encoding="utf-8")
    ruta = str(DATA) + "\\"
    (d / "expressions.tmdl").write_text(
        f'expression RutaDatos = "{ruta}" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]\n'
        "\n\tannotation PBI_ResultType = Text\n", encoding="utf-8")
    rels = []
    for ft, fc, tt, tc in RELS:
        rels += [f"relationship {uuid.uuid5(uuid.NAMESPACE_URL, ft + fc + tt)}", f"\tfromColumn: {ft}.{fc}", f"\ttoColumn: {tt}.{tc}", ""]
    (d / "relationships.tmdl").write_text("\n".join(rels), encoding="utf-8")
    for t, cols in TABLES.items():
        (d / "tables" / f"{t}.tmdl").write_text(tmdl_table(t, cols), encoding="utf-8")
    (d / "tables" / "sel_anio.tmdl").write_text(tmdl_sel_anio(), encoding="utf-8")
    (d / "tables" / "Medidas.tmdl").write_text(tmdl_measures(), encoding="utf-8")


# ---------------------------------------------------------------- tema

def theme() -> dict:
    txt = lambda size, color, font: {"fontSize": size, "fontFace": font, "color": color}
    no = [{"show": False}]
    return {
        "name": "Noche Paracas",
        "dataColors": [C["indigo"], C["oro"], C["turquesa"], C["terracota"], C["muted"], C["faint"], C["pos"], C["neg"]],
        "foreground": C["ink"], "foregroundNeutralSecondary": C["muted"], "foregroundNeutralTertiary": C["faint"],
        "background": C["card"], "backgroundLight": C["sunk"], "backgroundNeutral": C["line"], "tableAccent": C["indigo"],
        "good": C["pos"], "neutral": C["oro"], "bad": C["neg"],
        "textClasses": {
            "callout": txt(28, C["ink"], FONT_NUM), "title": txt(12, C["ink"], FONT_TXT), "header": txt(11, C["ink"], "Segoe UI Semibold"),
            "label": txt(10, C["muted"], FONT_TXT),
        },
        "visualStyles": {
            "*": {"*": {
                "background": [{"show": False}], "border": no, "dropShadow": no, "title": no, "visualHeader": no,
                "visualTooltip": [{"show": True, "background": {"solid": {"color": C["sunk"]}},
                                   "titleFontColor": {"solid": {"color": C["muted"]}}, "valueFontColor": {"solid": {"color": C["ink"]}}}],
                "padding": [{"top": 0, "bottom": 0, "left": 0, "right": 0}],
                "valueAxis": [{"gridlineColor": {"solid": {"color": C["line"]}}, "gridlineThickness": 1, "fontFamily": FONT_TXT,
                               "labelColor": {"solid": {"color": C["muted"]}}}],
                "categoryAxis": [{"fontFamily": FONT_TXT, "labelColor": {"solid": {"color": C["muted"]}}}],
                "legend": [{"fontFamily": FONT_TXT, "labelColor": {"solid": {"color": C["ink"]}}}],
            }},
            "page": {"*": {"background": [{"transparency": 100}], "outspace": [{"color": {"solid": {"color": C["bg"]}}}]}},
        },
    }


# ---------------------------------------------------------------- ayudas PBIR

def lit(v):
    if isinstance(v, bool):
        s = "true" if v else "false"
    elif isinstance(v, (int, float)):
        s = f"{v}D"
    else:
        s = f"'{v}'"
    return {"expr": {"Literal": {"Value": s}}}


def color(hex_: str):
    return {"solid": {"color": lit(hex_)}}


def measure(name: str):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": "Medidas"}}, "Property": name}}


def fx(mname: str):
    """Color tomado de una medida (formato condicional por valor de campo)."""
    return {"solid": {"color": {"expr": measure(mname)}}}


def column(table: str, col: str):
    return {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": col}}


def proj(field: dict, display: str | None = None):
    kind = "Measure" if "Measure" in field else "Column"
    ent = field[kind]["Expression"]["SourceRef"]["Entity"]
    prop = field[kind]["Property"]
    p = {"field": field, "queryRef": f"{ent}.{prop}", "nativeQueryRef": prop}
    if display:
        p["displayName"] = display
    return p


def container(name, x, y, w, h, visual, z=1000, filters=None):
    c = {"$schema": f"{SCH}/item/report/definition/visualContainer/{V_VISUAL}/schema.json", "name": name,
         "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z}, "visual": visual}
    if filters:
        c["filterConfig"] = {"filters": filters}
    return c


def card(name, x, y, w, h, mname, size=26, col=None, font=FONT_NUM, color_measure=None):
    objs = {"labels": [{"properties": {"fontSize": lit(size), "fontFamily": lit(font), "labelDisplayUnits": lit(1),
                                       "color": fx(color_measure) if color_measure else color(col or C["ink"])}}],
            "categoryLabels": [{"properties": {"show": lit(False)}}]}
    return container(name, x, y, w, h, {"visualType": "card", "query": {"queryState": {"Values": {"projections": [proj(measure(mname))]}}},
                                        "objects": objs, "drillFilterOtherVisuals": True})


def chart(name, x, y, w, h, vtype, buckets: dict, objects: dict | None = None, sort: tuple | None = None, filters=None):
    q = {"queryState": {k: {"projections": v} for k, v in buckets.items()}}
    if sort:
        field, direction = sort
        q["sortDefinition"] = {"sort": [{"field": field, "direction": direction}], "isDefaultSort": False}
    return container(name, x, y, w, h, {"visualType": vtype, "query": q, "objects": objects or {}, "drillFilterOtherVisuals": True},
                     filters=filters)


def series_colors(field: dict, pairs: list[tuple[str, str]]) -> list:
    """dataPoint por valor de la serie (el color sigue a la entidad, no a su posición)."""
    return [{"properties": {"fill": color(hex_)},
             "selector": {"data": [{"scopeId": {"Comparison": {"ComparisonKind": 0, "Left": field, "Right": {"Literal": {"Value": f"'{value}'"}}}}}]}}
            for value, hex_ in pairs]


def measure_colors(pairs: list[tuple[str, str]]) -> list:
    return [{"properties": {"fill": color(hex_)}, "selector": {"metadata": f"Medidas.{m}"}} for m, hex_ in pairs]


def fx_fill(mname: str, metadata: str | None = None) -> dict:
    sel = {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}
    if metadata:
        sel["metadata"] = f"Medidas.{metadata}"
    return {"properties": {"fill": fx(mname)}, "selector": sel}


NO_AXES = {"valueAxis": [{"properties": {"show": lit(False), "gridlineShow": lit(False)}}],
           "categoryAxis": [{"properties": {"show": lit(False), "innerPadding": lit(0), "axisType": lit("Categorical")}}],
           "legend": [{"properties": {"show": lit(False)}}]}


def bar_objects(fill_fx: str | None = None, fill: str | None = None, cat_width=None, label_size=9, cat_size=10, pad=None) -> dict:
    o = {
        "dataPoint": [fx_fill(fill_fx)] if fill_fx else [{"properties": {"fill": color(fill or C["indigo"])}}],
        "labels": [{"properties": {"show": lit(True), "color": color(C["ink"]), "fontSize": lit(label_size), "fontFamily": lit("Segoe UI Semibold"),
                                   "labelDisplayUnits": lit(1)}}],
        "valueAxis": [{"properties": {"show": lit(False), "gridlineShow": lit(False)}}],
        "categoryAxis": [{"properties": {"show": lit(True), "fontSize": lit(cat_size), "labelColor": color(C["ink"]), "fontFamily": lit(FONT_TXT)}}],
        "legend": [{"properties": {"show": lit(False)}}],
    }
    if cat_width:
        o["categoryAxis"][0]["properties"]["maxMarginFactor"] = lit(cat_width)
    if pad is not None:
        o["categoryAxis"][0]["properties"]["innerPadding"] = lit(pad)
    return o


def meter(name, x, y, w, mname, rest, col, h=12):
    """Barra de avance: valor + resto (pista) en una barra 100% apilada."""
    return chart(name, x, y, w, h, "hundredPercentStackedBarChart",
                 {"Category": [proj(column("Medidas", "Columna1"))], "Y": [proj(measure(mname)), proj(measure(rest))]},
                 {"dataPoint": measure_colors([(mname, col), (rest, C["track"])]), **NO_AXES,
                  "labels": [{"properties": {"show": lit(False)}}]})


def topn_filter(name, table, col, mname, n):
    """Filtro Top N de visual (lo resuelve el motor una vez; evita RANKX sobre miles de filas)."""
    src = lambda s: {"Column": {"Expression": {"SourceRef": {"Source": s}}, "Property": col}}
    return {"name": name, "field": column(table, col), "type": "TopN",
            "filter": {"Version": 2,
                       "From": [{"Name": "subquery", "Expression": {"Subquery": {"Query": {
                           "Version": 2,
                           "From": [{"Name": "d", "Entity": table, "Type": 0}, {"Name": "m", "Entity": "Medidas", "Type": 0}],
                           "Select": [{**src("d"), "Name": "field"}],
                           "OrderBy": [{"Direction": 2, "Expression": {"Measure": {"Expression": {"SourceRef": {"Source": "m"}}, "Property": mname}}}],
                           "Top": n}}}, "Type": 2},
                           {"Name": "d", "Entity": table, "Type": 0}],
                       "Where": [{"Condition": {"In": {"Expressions": [src("d")], "Table": {"SourceRef": {"Source": "subquery"}}}}}]}}


def slicer(name, x, y, w, h, table, col, horizontal=False, default=None):
    general = {}
    if horizontal:
        general["orientation"] = {"expr": {"Literal": {"Value": "1D"}}}
    if default is not None:
        general["filter"] = {"filter": {
            "Version": 2, "From": [{"Name": "s", "Entity": table, "Type": 0}],
            "Where": [{"Condition": {"In": {"Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": col}}],
                                            "Values": [[{"Literal": {"Value": f"'{default}'"}}]]}}}]}}
    objs = {"data": [{"properties": {"mode": lit("Basic" if horizontal else "Dropdown")}}],
            "header": [{"properties": {"show": lit(False)}}],
            "items": [{"properties": {"fontColor": color(C["ink"]), "background": color(C["sunk"]), "textSize": lit(9),
                                      "fontFamily": lit(FONT_TXT), "outline": lit("Frame"), "outlineColor": color(C["line"])}}],
            "selection": [{"properties": {"singleSelect": lit(default is not None), "selectAllCheckboxEnabled": lit(False)}}]}
    if general:
        objs["general"] = [{"properties": general}]
    return container(name, x, y, w, h, {"visualType": "slicer",
                                        "query": {"queryState": {"Values": {"projections": [proj(column(table, col))]}}},
                                        "objects": objs, "drillFilterOtherVisuals": True}, z=5000)


def nav_button(name, x, y, w, h, target):
    return container(name, x, y, w, h, {
        "visualType": "actionButton",
        "objects": {"icon": [{"properties": {"show": lit(False)}}], "outline": [{"properties": {"show": lit(False)}}],
                    "fill": [{"properties": {"show": lit(False)}}], "text": [{"properties": {"show": lit(False)}}]},
        "visualContainerObjects": {"visualLink": [{"properties": {"show": lit(True), "type": lit("PageNavigation"),
                                                                   "navigationSection": lit(target)}}]},
        "drillFilterOtherVisuals": True}, z=6000)


def table_objects(row_pad=4, size=10) -> dict:
    return {"grid": [{"properties": {"gridVertical": lit(False), "gridHorizontal": lit(True), "gridHorizontalColor": color(C["line"]),
                                     "rowPadding": lit(row_pad), "outlineColor": color(C["line"])}}],
            "columnHeaders": [{"properties": {"fontColor": color(C["muted"]), "backColor": color(C["card"]), "fontFamily": lit(FONT_TXT), "fontSize": lit(9)}}],
            "values": [{"properties": {"fontColor": color(C["ink"]), "backColor": color(C["card"]), "backColorSecondary": color(C["card"]),
                                       "fontSize": lit(size), "fontFamily": lit(FONT_TXT)}}],
            "total": [{"properties": {"totals": lit(False)}}]}


# ---------------------------------------------------------------- HTML de fondo (piezas)

def icon(name, col, size=20, sw=1.7):
    return (f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{col}" stroke-width="{sw}" '
            f'stroke-linecap="round" stroke-linejoin="round"><path d="{ICON[name]}"/></svg>')


def tile(name, col, size=36):
    return (f'<span style="width:{size}px;height:{size}px;border-radius:10px;background:{col}26;box-shadow:inset 0 0 0 1px {col}55;'
            f'display:flex;align-items:center;justify-content:center;flex-shrink:0">{icon(name, col, int(size * 0.55))}</span>')


def box(x, y, w, h, extra=""):
    return f'<div class="abs box" style="left:{x}px;top:{y}px;width:{w}px;height:{h}px;{extra}"></div>'


def text(x, y, html, cls="", style=""):
    return f'<div class="abs {cls}" style="left:{x}px;top:{y}px;{style}">{html}</div>'


def panel_bg(x, y, w, h, title, sub, ic, col):
    return (box(x, y, w, h) + f'<div class="abs" style="left:{x + 16}px;top:{y + 14}px;display:flex;align-items:center;gap:10px">'
            f'{icon(ic, col, 18)}<span style="display:flex;flex-direction:column;gap:1px"><span style="font-size:15px;font-weight:600">{title}</span>'
            f'<span class="mini">{sub}</span></span></div>')


def kpi_bg(i, label, ic, col, pill_text=None, note=None):
    x = X0 + i * KPI_STEP
    out = box(x, KPI_Y, KPI_W, KPI_H)
    out += (f'<div class="abs" style="left:{x + 14}px;top:{KPI_Y + 12}px;width:{KPI_W - 28}px;display:flex;align-items:center;gap:10px">'
            f'{tile(ic, col, 36)}<span class="small" style="line-height:1.25">{label}</span></div>')
    # pastilla (el texto lo pone un visual, salvo que sea fijo)
    out += (f'<div class="abs" style="left:{x + KPI_W - 126}px;top:{KPI_Y + 62}px;width:112px;height:22px;border-radius:11px;'
            f'background:rgba(236,234,244,.07);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;color:{C["muted"]}">'
            f'{pill_text or ""}</div>')
    if note:
        out += text(x + 14, KPI_Y + 112, note, "mini")
    return out


def kpi_vis(i, value_m, pill_m=None, pill_color_m=None):
    x = X0 + i * KPI_STEP
    v = [card(f"kpi{i}", x + 6, KPI_Y + 52, 120, 42, value_m, size=24)]
    if pill_m:
        v.append(card(f"kpipill{i}", x + KPI_W - 126, KPI_Y + 62, 112, 22, pill_m, size=8, font="Segoe UI Semibold",
                      col=C["muted"], color_measure=pill_color_m))
    return v


def kpi_meter(i, m, rest, col):
    x = X0 + i * KPI_STEP
    return meter(f"kpimeter{i}", x + 10, KPI_Y + 96, KPI_W - 20, m, rest, col, h=14)


# ---------------------------------------------------------------- páginas

def page_panorama():
    bg, v, inter = "", [], []
    bg += kpi_bg(0, "Monto adjudicado ·<br>mil millones de S/", "monedas", C["indigo"])
    v += kpi_vis(0, "Monto anio", "Pill monto", "Color var monto")
    v.append(chart("kpicomp", X0 + 8, KPI_Y + 92, KPI_W - 16, 40, "clusteredBarChart",
                   {"Category": [proj(column("dim_fecha", "anio"), "Año")], "Y": [proj(measure("Monto comparativo"), "Mil millones S/")]},
                   bar_objects(fill_fx="Color comparativo", cat_size=8, label_size=8, pad=20), sort=(column("dim_fecha", "anio"), "Ascending")))
    x = X0 + KPI_STEP
    bg += kpi_bg(1, "Procesos convocados", "documento", C["terracota"])
    v += kpi_vis(1, "Procesos anio", "Pill procesos", "Color var procesos")
    v.append(chart("kpimeses", x + 10, KPI_Y + 94, KPI_W - 20, 36, "areaChart",
                   {"Category": [proj(column("dim_fecha", "mes"), "Mes")], "Y": [proj(measure("Procesos mes"), "Procesos")]},
                   {"dataPoint": [{"properties": {"fill": color(C["indigo"] + "99")}}], **NO_AXES, "labels": [{"properties": {"show": lit(False)}}],
                    "categoryAxis": [{"properties": {"show": lit(False), "axisType": lit("Scalar"), "innerPadding": lit(20)}}],
                    "zoom": [{"properties": {"show": lit(False)}}]},
                   sort=(column("dim_fecha", "mes"), "Ascending")))
    bg += kpi_bg(2, "Procesos con buena pro", "check", C["turquesa"], note="de los procesos convocados")
    v += kpi_vis(2, "% adjudicados", "Pill adjudicados")
    v.append(kpi_meter(2, "% adjudicados", "% adjudicados resto", C["turquesa"]))
    bg += kpi_bg(3, "Días hasta la buena pro", "reloj", C["oro"], pill_text="promedio",
                 note='<span style="display:inline-flex;width:216px;justify-content:space-between"><span>0</span><span>60 días</span></span>')
    v += kpi_vis(3, "Dias buena pro")
    v.append(kpi_meter(3, "Dias escala", "Dias escala resto", C["oro"]))

    # panel izquierdo: categorías
    px, pw = X0, 490
    bg += panel_bg(px, PANEL_Y, pw, PANEL_B - PANEL_Y, "Monto adjudicado por categoría", "miles de millones de soles · año elegido frente al anterior", "monedas", C["indigo"])
    bg += text(px + 16, PANEL_Y + 50,
               f'<span style="display:inline-flex;align-items:center;gap:5px"><span style="width:10px;height:10px;border-radius:2px;background:{C["old"]}"></span>Año anterior</span>'
               f'<span style="display:inline-flex;align-items:center;gap:5px;margin-left:12px"><span style="width:10px;height:10px;border-radius:2px;'
               f'background:{C["indigo"]}"></span>Año elegido</span>', "mini")
    band_y, band_h = PANEL_Y + 70, 74
    for k, (cat, ic, col) in enumerate([("Obras", "obra", C["terracota"]), ("Servicios", "servicio", C["turquesa"]), ("Bienes", "bien", C["oro"])]):
        cy = band_y + k * band_h
        bg += text(px + 16, cy + band_h / 2 - 22, tile(ic, col, 44))
        if k:
            bg += f'<div class="abs rule" style="left:{px + 16}px;top:{cy}px;width:{pw - 32}px"></div>'
        v.append(card(f"var{cat}", px + pw - 92, int(cy + band_h / 2 - 16), 80, 32, f"Var {cat}", size=17, color_measure=f"Color var {cat}"))
    cat = column("dim_categoria", "categoria")
    o = bar_objects(cat_width=22, label_size=8, cat_size=10, pad=18)
    o["dataPoint"] = measure_colors([("Monto anio ant", C["old"]), ("Monto anio", C["indigo"])])
    v.append(chart("categorias", px + 70, band_y + 2, pw - 170, band_h * 3 - 4, "clusteredBarChart",
                   {"Category": [proj(cat, "Categoría")], "Y": [proj(measure("Monto anio ant"), "Año anterior"), proj(measure("Monto anio"), "Año elegido")]},
                   o, sort=(cat, "Ascending")))
    comp_y = band_y + band_h * 3 + 8
    bg += f'<div class="abs rule" style="left:{px + 16}px;top:{comp_y}px;width:{pw - 32}px"></div>'
    bg += text(px + 16, comp_y + 10, "Composición del monto del año", "mini")
    v.append(chart("composicion", px + 12, comp_y + 24, pw - 24, 48, "hundredPercentStackedBarChart",
                   {"Category": [proj(column("Medidas", "Columna1"))], "Series": [proj(cat, "Categoría")], "Y": [proj(measure("% categoria"), "% del monto")]},
                   {"dataPoint": series_colors(cat, [("Obras", C["terracota"]), ("Servicios", C["turquesa"]), ("Bienes", C["oro"])]), **NO_AXES,
                    "labels": [{"properties": {"show": lit(True), "color": color(C["side"]), "fontSize": lit(9), "labelPrecision": lit(1),
                                               "fontFamily": lit("Segoe UI Semibold")}}]}))
    ins_y = PANEL_B - 62
    bg += (f'<div class="abs" style="left:{px + 16}px;top:{ins_y}px;width:{pw - 32}px;height:46px;border-radius:10px;'
           f'background:linear-gradient(90deg,{C["oro"]}22,transparent);box-shadow:inset 3px 0 0 {C["oro"]}"></div>')
    bg += text(px + 28, ins_y + 12, icon("chispa", C["oro"], 22))
    v.append(card("insight", px + 58, ins_y + 8, pw - 86, 30, "Insight categoria", size=10, font="Segoe UI Semibold"))
    inter += [("slc_categoria", n, "NoFilter") for n in ("categorias", "composicion", "insight", "varObras", "varServicios", "varBienes")]

    # panel derecho arriba: meses
    rx, rw = X0 + pw + 12, X1 - X0 - pw - 12
    bg += panel_bg(rx, PANEL_Y, rw, 300, "Enero arranca casi parado", "% de las convocatorias del año · promedio de los años completos", "reloj", C["terracota"])
    v.append(chart("meses", rx + 10, PANEL_Y + 64, rw - 20, 226, "clusteredColumnChart",
                   {"Category": [proj(column("dim_fecha", "mes_nombre"), "Mes")], "Y": [proj(measure("% convocatorias mes"), "% del año")]},
                   {"dataPoint": [fx_fill("Color mes")],
                    "labels": [{"properties": {"show": lit(True), "color": color(C["ink"]), "fontSize": lit(9), "labelPrecision": lit(1),
                                               "fontFamily": lit("Segoe UI Semibold")}}],
                    "valueAxis": [{"properties": {"show": lit(False), "gridlineShow": lit(False)}}],
                    "categoryAxis": [{"properties": {"show": lit(True), "fontSize": lit(9), "labelColor": color(C["muted"]), "innerPadding": lit(28)}}],
                    "legend": [{"properties": {"show": lit(False)}}]},
                   sort=(column("dim_fecha", "mes_nombre"), "Ascending")))
    # panel derecho abajo: nivel de gobierno
    ny = PANEL_Y + 312
    bg += panel_bg(rx, ny, rw, PANEL_B - ny, "Quién compra: nivel de gobierno", "% del monto adjudicado en el año", "nacional", C["indigo"])
    nivel = column("dim_entidad", "nivel_gobierno")
    v.append(chart("nivel_barra", rx + 12, ny + 44, rw - 24, 38, "hundredPercentStackedBarChart",
                   {"Category": [proj(column("Medidas", "Columna1"))], "Series": [proj(nivel, "Nivel")], "Y": [proj(measure("% nivel"), "% del monto")]},
                   {"dataPoint": series_colors(nivel, [("Gobierno nacional", C["indigo"]), ("Gobierno local", C["turquesa"]), ("Gobierno regional", C["oro"])]),
                    **NO_AXES, "labels": [{"properties": {"show": lit(False)}}]}))
    bw = (rw - 32 - 16) // 3
    for k, (lab, ic, col) in enumerate([("Nacional", "nacional", C["indigo"]), ("Local", "local", C["turquesa"]), ("Regional", "pin", C["oro"])]):
        bx = rx + 16 + k * (bw + 8)
        bg += box(bx, ny + 84, bw, 44, f"background:{C['sunk']};box-shadow:none;border-radius:8px")
        bg += text(bx + 10, ny + 96, icon(ic, col, 20))
        bg += text(bx + 38, ny + 88, lab, "mini")
        v.append(card(f"nivel{lab}", bx + 34, ny + 100, 80, 26, f"% {lab}", size=15))
    inter += [("slc_nivel", "nivel_barra", "NoFilter")] + [("slc_nivel", f"nivel{n}", "NoFilter") for n in ("Nacional", "Local", "Regional")]
    return bg, v, inter


def page_competencia():
    bg, v, inter = "", [], []
    bg += kpi_bg(0, "Contratación directa ·<br>mil millones de S/", "flecha", C["terracota"], note="del monto adjudicado en el año")
    v += kpi_vis(0, "Monto directo", "Pill directo")
    v.append(kpi_meter(0, "% directo", "% directo resto", C["terracota"]))
    bg += kpi_bg(1, "Directas con un solo<br>participante", "persona", C["neg"], note="de las contrataciones directas")
    v += kpi_vis(1, "% directas un participante", "Pill directas")
    v.append(kpi_meter(1, "% directas un participante", "% directas un participante resto", C["neg"]))
    bg += kpi_bg(2, "Participantes en<br>procesos abiertos", "grupo", C["indigo"], pill_text="promedio",
                 note='<span style="display:inline-flex;width:216px;justify-content:space-between"><span>0</span><span>20 participantes</span></span>')
    v += kpi_vis(2, "Participantes abiertos")
    v.append(kpi_meter(2, "Participantes escala", "Participantes escala resto", C["indigo"]))
    bg += kpi_bg(3, "Procesos con un<br>solo participante", "alerta", C["oro"], pill_text="todos", note="de los procesos con dato")
    v += kpi_vis(3, "% un participante anio")
    v.append(kpi_meter(3, "% un participante anio", "% un participante resto", C["oro"]))

    px, pw = X0, 520
    bg += panel_bg(px, PANEL_Y, pw, PANEL_B - PANEL_Y, "Procesos con un solo participante, por procedimiento",
                   "% de los procesos del año · procedimientos con 500 procesos o más", "balanza", C["indigo"])
    met = column("dim_metodo", "metodo")
    v.append(chart("metodos", px + 12, PANEL_Y + 62, pw - 24, PANEL_B - PANEL_Y - 74, "clusteredBarChart",
                   {"Category": [proj(met, "Procedimiento")], "Y": [proj(measure("% un participante metodo"), "% con un participante")]},
                   bar_objects(fill_fx="Color metodo", cat_width=42, label_size=9, cat_size=10, pad=30),
                   sort=(measure("% un participante metodo"), "Descending")))

    rx, rw = X0 + pw + 12, X1 - X0 - pw - 12
    th = 162
    bg += panel_bg(rx, PANEL_Y, rw, th, "Abiertos frente a directos", "participantes registrados por proceso, en promedio", "grupo", C["indigo"])
    half = (rw - 32) // 2
    bg += text(rx + 16, PANEL_Y + 62, f'<span style="display:flex;align-items:center;gap:8px">{icon("grupo", C["indigo"], 18)}Procesos abiertos</span>', "mini")
    bg += text(rx + 16 + half + 16, PANEL_Y + 62, f'<span style="display:flex;align-items:center;gap:8px">{icon("persona", C["neg"], 18)}Contrataciones directas</span>', "mini")
    bg += f'<div class="abs" style="left:{rx + 16 + half}px;top:{PANEL_Y + 62}px;width:1px;height:58px;background:{C["line"]}"></div>'
    v.append(card("part_abiertos", rx + 10, PANEL_Y + 82, 110, 40, "Participantes abiertos", size=26, col=C["indigo"]))
    v.append(card("part_directos", rx + 26 + half, PANEL_Y + 82, 110, 40, "Participantes directos", size=26, col=C["neg"]))
    bg += f'<div class="abs rule" style="left:{rx + 16}px;top:{PANEL_Y + 128}px;width:{rw - 32}px"></div>'
    v.append(card("ratio", rx + 12, PANEL_Y + 132, rw - 24, 24, "Ratio participantes", size=9, font=FONT_TXT, col=C["muted"]))

    ey = PANEL_Y + th + 12
    bg += panel_bg(rx, ey, rw, PANEL_B - ey, "Dónde se concentra la contratación directa",
                   "6 entidades con más monto por contratación directa en el año", "edificio", C["terracota"])
    ent = column("dim_entidad", "entidad")
    v.append(chart("entidades", rx + 12, ey + 58, rw - 24, 190, "tableEx",
                   {"Values": [proj(ent, "Entidad"), proj(measure("Procesos directos"), "Procesos"), proj(measure("Directo M"), "Millones S/")]},
                   {**table_objects(row_pad=3, size=9), "columnWidth": [
                       {"properties": {"value": lit(w)}, "selector": {"metadata": m}}
                       for m, w in (("dim_entidad.entidad", 286), ("Medidas.Procesos directos", 66), ("Medidas.Directo M", 86))]},
                   sort=(measure("Directo M"), "Descending"),
                   filters=[topn_filter("top6", "dim_entidad", "entidad", "Directo M", 6)]))
    bg += text(rx + 16, PANEL_B - 28, '"Participantes" son los registrados en el SEACE, no necesariamente los que presentaron oferta.', "mini",
               f"width:{rw - 32}px")
    return bg, v, inter


def page_proveedores():
    bg, v, inter = "", [], []
    bg += kpi_bg(0, "Monto ganado por<br>consorcios", "manos", C["indigo"], note="del monto adjudicado en el año")
    v += kpi_vis(0, "% Consorcio", "Pill consorcio")
    v.append(kpi_meter(0, "% Consorcio", "% Consorcio resto", C["indigo"]))
    bg += kpi_bg(1, "Proveedores ganadores", "grupo", C["turquesa"], note="consorcios, empresas, personas, otros")
    v += kpi_vis(1, "Ganadores", "Pill ganadores")
    tipo = column("dim_proveedor", "tipo")
    TIPOS = [("Consorcio", C["indigo"]), ("Empresa", C["turquesa"]), ("Persona natural", C["oro"]), ("Otro (no domiciliado u otro)", C["faint"])]
    x = X0 + KPI_STEP
    v.append(chart("kpitipos", x + 10, KPI_Y + 96, KPI_W - 20, 14, "hundredPercentStackedBarChart",
                   {"Category": [proj(column("Medidas", "Columna1"))], "Series": [proj(tipo, "Tipo")], "Y": [proj(measure("% tipo ganador"), "% del monto")]},
                   {"dataPoint": series_colors(tipo, TIPOS), **NO_AXES, "labels": [{"properties": {"show": lit(False)}}]}))
    bg += kpi_bg(2, "Top 10 empresas", "barras", C["oro"], note="del monto adjudicado en el año")
    v += kpi_vis(2, "% Top 10 empresas", "Pill HHI")
    v.append(kpi_meter(2, "% Top 10 empresas", "% Top 10 empresas resto", C["oro"]))
    bg += kpi_bg(3, "Entidades con sede<br>en Lima", "pin", C["terracota"], pill_text="del monto", note="Lima concentra casi la mitad")
    v += kpi_vis(3, "% Lima")
    v.append(kpi_meter(3, "% Lima", "% Lima resto", C["terracota"]))

    px, pw = X0, 530
    bg += panel_bg(px, PANEL_Y, pw, PANEL_B - PANEL_Y, "A quién le compra el Estado", "% del monto adjudicado en el año, por tipo de ganador", "manos", C["indigo"])
    v.append(chart("tipos", px + 12, PANEL_Y + 48, pw - 24, 36, "hundredPercentStackedBarChart",
                   {"Category": [proj(column("Medidas", "Columna1"))], "Series": [proj(tipo, "Tipo")], "Y": [proj(measure("% tipo ganador"), "% del monto")]},
                   {"dataPoint": series_colors(tipo, TIPOS), **NO_AXES, "labels": [{"properties": {"show": lit(False)}}]}))
    tw = (pw - 32 - 18) // 4
    for k, (lab, key, col) in enumerate([("Consorcios", "Consorcio", C["indigo"]), ("Empresas", "Empresa", C["turquesa"]),
                                        ("Personas naturales", "Persona natural", C["oro"]), ("Otros", "Otro", C["faint"])]):
        bx = px + 16 + k * (tw + 6)
        bg += box(bx, PANEL_Y + 84, tw, 66, f"background:{C['sunk']};box-shadow:inset 0 2px 0 {col};border-radius:8px")
        bg += text(bx + 8, PANEL_Y + 90, lab, "mini")
        v.append(card(f"tipo{k}", bx + 2, PANEL_Y + 104, 80, 26, f"% {key}", size=16))
        v.append(card(f"prom{k}", bx + 2, PANEL_Y + 128, tw - 4, 18, f"Promedio {key}", size=8, font=FONT_TXT, col=C["muted"]))
    sy = PANEL_Y + 160
    bg += f'<div class="abs rule" style="left:{px + 16}px;top:{sy}px;width:{pw - 32}px"></div>'
    bg += text(px + 16, sy + 8, '<span style="font-size:13px;font-weight:600;color:#ECEAF4">Top 10 empresas proveedoras</span>')
    bg += text(px + pw - 200, sy + 10, "millones de S/ adjudicados en el año", "mini")
    prov = column("dim_proveedor", "proveedor")
    v.append(chart("top_empresas", px + 12, sy + 32, pw - 24, PANEL_B - sy - 42, "clusteredBarChart",
                   {"Category": [proj(prov, "Empresa")], "Y": [proj(measure("Empresa M"), "Millones S/")]},
                   bar_objects(fill=C["turquesa"], cat_width=48, label_size=9, cat_size=9, pad=34),
                   sort=(measure("Empresa M"), "Descending"),
                   filters=[topn_filter("top10", "dim_proveedor", "proveedor", "Empresa M", 10)]))

    rx, rw = X0 + pw + 12, X1 - X0 - pw - 12
    dh = 368
    bg += panel_bg(rx, PANEL_Y, rw, dh, "Monto por departamento de la entidad", "miles de millones de soles · top 10 del año", "pin", C["terracota"])
    dep = column("dim_entidad", "departamento")
    v.append(chart("departamentos", rx + 12, PANEL_Y + 58, rw - 24, dh - 68, "clusteredBarChart",
                   {"Category": [proj(dep, "Departamento")], "Y": [proj(measure("Departamento top 10"), "Mil millones S/")]},
                   bar_objects(fill_fx="Color departamento", cat_width=22, label_size=9, cat_size=10, pad=30),
                   sort=(measure("Departamento top 10"), "Descending")))
    ly = PANEL_Y + dh + 12
    bg += box(rx, ly, rw, PANEL_B - ly)
    bg += text(rx + 14, ly + 14, tile("candado", C["indigo"], 40))
    bg += text(rx + 66, ly + 10, f'<b style="color:{C["ink"]};font-weight:600">Datos protegidos:</b> las personas naturales aparecen anonimizadas '
               'y no se publica ningún correo ni teléfono. Lima incluye entidades nacionales que ejecutan en todo el país.', "mini",
               f"width:{rw - 84}px;line-height:1.45")
    return bg, v, inter


BUILDERS = {"panorama": page_panorama, "competencia": page_competencia, "proveedores": page_proveedores}

# ---------------------------------------------------------------- barra lateral (fondo + visuales)

CHAKANA = "M14 2H22V8H28V14H34V22H28V28H22V34H14V28H8V22H2V14H8V8H14Z"
NAV_Y = [62, 116, 170]
SL = {"anio_lbl": 256, "anio": 274, "cat_lbl": 316, "cat": 334, "niv_lbl": 374, "niv": 392, "nota": 428, "base": 470}


def sidebar_bg(active: str) -> str:
    nav = ""
    for i, p in enumerate(PAGES):
        on = p["name"] == active
        y = NAV_Y[i]
        nav += (f'<div class="abs" style="left:14px;top:{y}px;width:190px;height:48px;border-radius:10px;box-sizing:border-box;padding:0 10px;'
                f'display:flex;align-items:center;gap:10px;background:{"#1F1E33" if on else "transparent"};color:{C["ink"] if on else C["muted"]}">'
                + (f'<span style="position:absolute;left:-14px;top:10px;bottom:10px;width:3px;border-radius:0 3px 3px 0;background:{C["oro"]}"></span>' if on else "")
                + icon(p["icon"], C["indigo"] if on else C["faint"], 20)
                + f'<span style="display:flex;flex-direction:column;gap:1px"><span style="font-size:14px;font-weight:600">{p["display"]}</span>'
                f'<span style="font-size:11px;color:{C["muted"] if on else C["faint"]}">{p["sub"]}</span></span>'
                f'<span style="margin-left:auto;font-family:\'Fira Code\';font-size:11px;color:{C["oro"] if on else C["faint"]}">0{i + 1}</span></div>')
    lbl = f"font-family:'Fira Code';font-size:10px;letter-spacing:.12em;color:{C['faint']}"
    base = "".join(text(14, SL["base"] + 20 + k * 21, t, "small") for k, t in enumerate(["Procesos", "Adjudicaciones", "Proveedores", "Entidades"]))
    greca = (f'<svg width="190" height="34" viewBox="0 0 190 34"><defs><pattern id="g" width="38" height="34" patternUnits="userSpaceOnUse">'
             f'<path d="M1 32V20H8V13H15V6H23V13H30V20H37V32" fill="none" stroke="{C["indigo"]}" stroke-width="1.6"/>'
             f'<rect x="16.5" y="18" width="5" height="5" fill="{C["oro"]}"/><path d="M8 32V26H30V32" fill="none" stroke="{C["terracota"]}" stroke-width="1.6"/>'
             '</pattern></defs><rect width="190" height="34" fill="url(#g)"/></svg>')
    return f"""
<div class="abs" style="left:0;top:0;width:{SIDE_W}px;height:{H}px;background:{C["side"]};border-right:1px solid {C["line"]};box-sizing:border-box"></div>
<div class="abs" style="left:14px;top:16px;display:flex;align-items:center;gap:10px">
  <svg width="34" height="34" viewBox="0 0 36 36"><path d="{CHAKANA}" fill="none" stroke="{C["indigo"]}" stroke-width="2" stroke-linejoin="round"/><circle cx="18" cy="18" r="4.5" fill="none" stroke="{C["oro"]}" stroke-width="2"/></svg>
  <span style="display:flex;flex-direction:column;line-height:1.15"><span style="font-weight:700;font-size:14px">Compras Públicas</span><span class="mini">Perú · SEACE · 2024–2026</span></span>
</div>
{nav}
<div class="abs rule" style="left:14px;top:228px;width:190px"></div>
{text(14, 236, "FILTROS", style=lbl)}
{text(14, SL["anio_lbl"], "Año de convocatoria", "small", "top:" + str(SL["anio_lbl"]) + "px")}
{text(14, SL["cat_lbl"], "Categoría", "small")}
{text(14, SL["niv_lbl"], "Nivel de gobierno", "small")}
{text(14, SL["nota"], "Sin selección: último año completo. * año en curso (parcial).", "mini", "width:190px;font-size:10px;color:" + C["faint"])}
<div class="abs rule" style="left:14px;top:{SL["base"] - 8}px;width:190px"></div>
{text(14, SL["base"], "LA BASE · 2024–2026", style=lbl)}
{base}
<div class="abs" style="left:14px;top:612px">{greca}</div>
<div class="abs mini" style="left:14px;top:656px;line-height:1.4"><b style="color:{C["ink"]};font-weight:600">Carlos Matos Gavino</b><br>Data Analyst · Lima</div>"""


def sidebar_visuals(active: str) -> tuple[list, list]:
    v = [slicer("slc_anio", 14, SL["anio"] + 16 - 16, 190, 34, "sel_anio", "Etiqueta", horizontal=True, default="2025"),
         slicer("slc_categoria", 14, SL["cat"], 190, 30, "dim_categoria", "categoria"),
         slicer("slc_nivel", 14, SL["niv"], 190, 30, "dim_entidad", "nivel_gobierno")]
    for k, m in enumerate(["Base procesos", "Base adjudicaciones", "Base proveedores", "Base entidades"]):
        v.append(card(f"base{k}", 120, SL["base"] + 17 + k * 21, 84, 20, m, size=9, font="Segoe UI Semibold"))
    for i, p in enumerate(PAGES):
        if p["name"] != active:
            v.append(nav_button(f"nav_{p['name']}", 14, NAV_Y[i], 190, 48, p["name"]))
    return v, []


# ---------------------------------------------------------------- fondos (HTML -> PNG)

def background_html(p: dict, body: str) -> str:
    return f"""<!doctype html><html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500&family=Fira+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
html,body{{margin:0;width:{W}px;height:{H}px;overflow:hidden;background:{C["bg"]};color:{C["ink"]};font-family:'Fira Sans','Segoe UI',sans-serif}}
.abs{{position:absolute}} .box{{background:{C["card"]};border-radius:12px;box-shadow:inset 0 0 0 1px {C["line"]};box-sizing:border-box}}
.small{{font-size:12px;color:{C["muted"]}}} .mini{{font-size:11px;color:{C["muted"]}}} .rule{{height:1px;background:{C["line"]}}}
</style></head><body>
{sidebar_bg(p["name"])}
<div class="abs" style="left:{X0}px;top:18px;font-family:'Fira Code';font-size:11px;letter-spacing:.14em;color:{C["oro"]}">{p["kicker"]}</div>
<div class="abs" style="left:{X0}px;top:36px;font-size:25px;font-weight:600;letter-spacing:-.01em;white-space:nowrap">{p["title"]}</div>
<div class="abs" style="right:{W - X1}px;top:44px;display:flex;align-items:center;gap:6px;font-size:11px;color:{C["muted"]};padding:5px 10px;border-radius:999px;box-shadow:inset 0 0 0 1px {C["line"]}">
<span style="width:6px;height:6px;border-radius:50%;background:{C["pos"]}"></span>Datos abiertos OCDS</div>
{body}
<div class="abs" style="left:{X0}px;top:692px;font-size:10.5px;color:{C["faint"]}">{p["footer"]}</div>
</body></html>"""


def render_backgrounds(bodies: dict) -> None:
    BG.mkdir(exist_ok=True)
    for p in PAGES:
        html = BG / f"bg_{p['name']}.html"
        html.write_text(background_html(p, bodies[p["name"]]), encoding="utf-8")
        png = BG / f"bg_{p['name']}.png"
        png.unlink(missing_ok=True)
        subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=2",
                        "--virtual-time-budget=6000", f"--window-size={W},{H}", f"--screenshot={png}", html.as_uri()],
                       check=True, capture_output=True, timeout=120)
        # Edge puede devolver el control antes de escribir el PNG
        for _ in range(60):
            if png.exists() and png.stat().st_size > 0:
                break
            time.sleep(0.5)
        print("fondo", png.name, png.stat().st_size)


# ---------------------------------------------------------------- reporte (PBIR)

def write_report(built: dict) -> None:
    if RP.exists():
        shutil.rmtree(RP)
    d = RP / "definition"
    res = RP / "StaticResources" / "RegisteredResources"
    res.mkdir(parents=True)
    (d / "pages").mkdir(parents=True)
    j = lambda path, obj: (path.parent.mkdir(parents=True, exist_ok=True), path.write_text(json.dumps(obj, indent=2, ensure_ascii=False), encoding="utf-8"))
    j(RP / "definition.pbir", {"$schema": f"{SCH}/item/report/definitionProperties/2.0.0/schema.json", "version": "4.0",
                               "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})
    j(res / "NocheParacas.json", theme())
    images = []
    for p in PAGES:
        src = BG / f"bg_{p['name']}.png"
        if src.exists():
            shutil.copy(src, res / src.name)
            images.append({"name": src.name, "path": src.name, "type": "Image"})
    ver = {"visual": V_VISUAL, "report": V_REPORT, "page": V_PAGE}
    j(d / "report.json", {
        "$schema": f"{SCH}/item/report/definition/report/{V_REPORT}/schema.json",
        "themeCollection": {"customTheme": {"name": "NocheParacas.json", "reportVersionAtImport": ver, "type": "RegisteredResources"}},
        "resourcePackages": [{"name": "RegisteredResources", "type": "RegisteredResources",
                              "items": [{"name": "NocheParacas.json", "path": "NocheParacas.json", "type": "CustomTheme"}, *images]}],
        "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True,
                     "useEnhancedTooltips": True},
    })
    j(d / "version.json", {"$schema": f"{SCH}/item/report/definition/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    j(d / "pages" / "pages.json", {"$schema": f"{SCH}/item/report/definition/pagesMetadata/1.0.0/schema.json",
                                   "pageOrder": [p["name"] for p in PAGES], "activePageName": PAGES[0]["name"]})
    for p in PAGES:
        visuals, inter = built[p["name"]]
        page = {"$schema": f"{SCH}/item/report/definition/page/{V_PAGE}/schema.json", "name": p["name"], "displayName": p["display"],
                "displayOption": "FitToPage", "height": H, "width": W,
                "objects": {"background": [{"properties": {
                    "image": {"image": {"name": lit(f"bg_{p['name']}.png"),
                                        "url": {"expr": {"ResourcePackageItem": {"PackageName": "RegisteredResources", "PackageType": 1, "ItemName": f"bg_{p['name']}.png"}}},
                                        "scaling": lit("Fit")}},
                    "transparency": lit(0)}}],
                    "outspace": [{"properties": {"color": color(C["bg"])}}]}}
        if inter:
            page["visualInteractions"] = [{"source": s, "target": t, "type": ty} for s, t, ty in inter]
        pd_ = d / "pages" / p["name"]
        j(pd_ / "page.json", page)
        for vis in visuals:
            j(pd_ / "visuals" / vis["name"] / "visual.json", vis)
    j(HERE / f"{NAME}.pbip", {"$schema": f"{SCH}/pbip/pbipProperties/1.0.0/schema.json", "version": "1.0",
                              "artifacts": [{"report": {"path": f"{NAME}.Report"}}], "settings": {"enableAutoRecovery": True}})


if __name__ == "__main__":
    bodies, built = {}, {}
    for p in PAGES:
        bg, vis, inter = BUILDERS[p["name"]]()
        sv, si = sidebar_visuals(p["name"])
        bodies[p["name"]] = bg
        built[p["name"]] = (vis + sv, inter + si)
    write_model()
    if "--no-png" not in sys.argv:
        render_backgrounds(bodies)
    write_report(built)
    print("PBIP listo:", HERE / f"{NAME}.pbip")
