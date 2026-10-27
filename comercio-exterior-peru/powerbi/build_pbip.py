"""Genera el proyecto Power BI (PBIP: modelo TMDL + reporte PBIR) del caso Comercio exterior del Perú.

Un mismo plano de posiciones (LAYOUT) produce:
  - los fondos de página (HTML -> PNG con Edge headless): barra lateral, títulos, tarjetas, íconos e ilustraciones;
  - los visuales de Power BI, que se colocan encima de esas tarjetas.

Uso: python powerbi/build_pbip.py            (reescribe ComercioExteriorPeru.* y los fondos)
     python powerbi/build_pbip.py --no-png   (sin regenerar los PNG)
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import uuid
from pathlib import Path

HERE = Path(__file__).resolve().parent
NAME = "ComercioExteriorPeru"
DATA = (HERE.parent / "data" / "processed").resolve()
SM = HERE / f"{NAME}.SemanticModel"
RP = HERE / f"{NAME}.Report"
BG = HERE / "backgrounds"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
W, H = 1280, 720

SCH = "https://developer.microsoft.com/json-schemas/fabric"
V_VISUAL, V_PAGE, V_REPORT = "2.4.0", "2.0.0", "3.0.0"

# Identidad "Tierra y Pacífico"
C = {"bg": "#F4EFE7", "card": "#FBF8F3", "line": "#E3DCD0", "ink": "#1E1B18", "ink2": "#5E574E", "muted": "#8A8178",
     "cobre": "#C2562B", "pacifico": "#00819E", "lucuma": "#B88400", "cochinilla": "#9E2A4F", "indigo": "#5B5FA8", "muna": "#4E7A2E",
     "track": "#EFE8DD"}

ICON = {
    "exportar": "M3 15h18l-2.5 5h-13z M6 15V9h6v6 M12 11h5v4 M8 9V5h3v4",
    "importar": "M4 8l8-4 8 4v8l-8 4-8-4z M4 8l8 4 8-4 M12 12v8",
    "balanza": "M12 4v16 M7 20h10 M5 8h14 M5 8l-3 6h6z M19 8l-3 6h6z",
    "calendario": "M4 6h16v14H4z M4 10h16 M8 3v4 M16 3v4 M8 17l3-3 2 2 3-3",
    "pastel": "M12 3a9 9 0 1 0 9 9h-9z M14.5 3.4A9 9 0 0 1 20.6 9.5h-6.1z",
    "podio": "M3 20h18 M5 20v-6h4v6 M10 20V9h4v11 M15 20v-8h4v8",
    "mundo": "M12 3a9 9 0 1 0 0 18a9 9 0 1 0 0-18 M3 12h18 M12 3c3 3 3 15 0 18 M12 3c-3 3-3 15 0 18",
    "diana": "M12 3a9 9 0 1 0 0 18a9 9 0 1 0 0-18 M12 8a4 4 0 1 0 0 8a4 4 0 1 0 0-8 M12 11.5v1",
    "mina": "M4 20l9-9 M13 4c3.5 0 7 3.5 7 7 M13 4c-2.5 1-4.5 3-5.5 5.5 M20 11c-1 2.5-3 4.5-5.5 5.5 M10 9l5 5",
    "brote": "M12 21v-9 M12 12c0-4.5 3-7.5 8-7.5 0 4.5-3 7.5-8 7.5 M12 15c0-3.5-2.5-6-7-6 0 3.5 2.5 6 7 6",
    "cacao": "M12 3c4.5 2 6.5 6 6.5 9s-2 7-6.5 9c-4.5-2-6.5-6-6.5-9s2-7 6.5-9z M12 3v18",
    "fruta": "M12 4c3 0 4 4 5 7 2 5-1 10-5 10s-7-5-5-10c1-3 2-7 5-7z M12 13a2.6 2.6 0 1 0 0 5.2a2.6 2.6 0 1 0 0-5.2",
}

# ---------------------------------------------------------------- layout
X0, CW = 104, 1148          # contenido: de x=104 a x=1252
KPI_Y, KPI_H = 92, 100
MAIN_Y, MAIN_H = 204, 488


def kpi_x(i: int) -> int:
    return X0 + i * 290


PAGES = [
    {
        "name": "balanza", "display": "Balanza comercial", "nav": 0,
        "kicker": "PERÚ · COMERCIO EXTERIOR · 2015–2026",
        "title": "El Perú exportó 94 mil millones de dólares en 2025, un récord",
        "slicers": [("anio", "dim_fecha", "anio", "Año")],
        "kpis": [
            ("Exportaciones (mil M US$)", "exportar", "cobre", "KPI Export", "KPI Export nota"),
            ("Importaciones (mil M US$)", "importar", "pacifico", "KPI Import", "KPI Import nota"),
            ("Superávit (mil M US$)", "balanza", "lucuma", "KPI Superavit", "KPI Superavit nota"),
            ("Export. año en curso (mil M US$)", "calendario", "cochinilla", "KPI Export YTD", "KPI Export YTD nota"),
        ],
        "kpi_unit": "mil M US$",
        "cards": [
            {"x": X0, "y": MAIN_Y, "w": 723, "h": MAIN_H, "title": "Exportaciones e importaciones anuales",
             "sub": "miles de millones US$ FOB · BCRP",
             "legend": [("Exportaciones", "cobre"), ("Importaciones", "pacifico")]},
            {"x": 839, "y": MAIN_Y, "w": 413, "h": 170, "title": "Qué se exportó en el año", "sub": "participación del valor FOB"},
            {"x": 839, "y": 386, "w": 413, "h": 306, "title": "Importaciones por uso", "sub": "% del valor FOB importado en el año"},
        ],
        "extras": "balanza",
        "footer": "Fuente: BCRP, series de balanza comercial (valores FOB). El año en curso tiene datos hasta el último mes publicado.",
    },
    {
        "name": "socios", "display": "Socios comerciales", "nav": 1,
        "kicker": "SOCIOS COMERCIALES · EXPORTACIONES",
        "title": "China compra uno de cada tres dólares que exporta el Perú",
        "slicers": [("anio", "dim_fecha", "anio", "Año"), ("continente", "dim_pais", "continente", "Continente")],
        "kpis": [
            ("Peso de China", "pastel", "cochinilla", "KPI China", "KPI China nota"),
            ("Top 5 destinos", "podio", "cobre", "KPI Top5", "KPI Top5 nota"),
            ("Asia como destino", "mundo", "lucuma", "KPI Asia", "KPI Asia nota"),
            ("Concentración (HHI, 0–10 000)", "diana", "indigo", "KPI HHI", "KPI HHI nota"),
        ],
        "kpi_unit": "",
        "cards": [
            {"x": X0, "y": MAIN_Y, "w": 608, "h": MAIN_H, "title": "Top 10 destinos de exportación", "sub": "miles de millones US$ · ONU Comtrade"},
            {"x": 724, "y": MAIN_Y, "w": 528, "h": 228, "title": "China gana peso; Estados Unidos lo pierde",
             "sub": "% de las exportaciones peruanas por año",
             "legend": [("China", "cochinilla"), ("Estados Unidos", "pacifico")], "legend_dy": 36},
            {"x": 724, "y": 444, "w": 528, "h": 248, "title": "Exportaciones por continente", "sub": "% del total del año"},
        ],
        "extras": "socios",
        "footer": "Fuente: ONU Comtrade (API pública), Perú como reportante, valores US$. Concentración: índice Herfindahl-Hirschman.",
    },
    {
        "name": "productos", "display": "Productos", "nav": 2,
        "kicker": "PRODUCTOS · EXPORTACIONES",
        "title": "La minería pesa dos tercios; el cacao y la fruta crecen más rápido",
        "slicers": [("anio", "dim_fecha", "anio", "Año"), ("sector", "dim_producto", "sector", "Sector")],
        "kpis": [
            ("Minería y metales", "mina", "cobre", "KPI Mineria", "KPI Mineria nota"),
            ("Agro y alimentos (mil M US$)", "brote", "muna", "KPI Agro", "KPI Agro nota"),
            ("Cacao, crecimiento en 5 años", "cacao", "cochinilla", "KPI Cacao", "KPI Cacao nota"),
            ("Frutas y frutos secos (mil M US$)", "fruta", "lucuma", "KPI Frutas", "KPI Frutas nota"),
        ],
        "kpi_unit": "",
        "cards": [
            {"x": X0, "y": MAIN_Y, "w": 374, "h": MAIN_H, "title": "Exportaciones por sector", "sub": "% del total del año"},
            {"x": 491, "y": MAIN_Y, "w": 374, "h": MAIN_H, "title": "Los 8 productos más exportados", "sub": "capítulo HS · miles de millones US$"},
            {"x": 878, "y": MAIN_Y, "w": 374, "h": MAIN_H, "title": "Los que más crecieron en 5 años",
             "sub": "productos con más de US$ 100 M exportados en el año"},
        ],
        "extras": "productos",
        "footer": "Fuente: ONU Comtrade (API pública), capítulos HS a 2 dígitos. Sectores agrupados por el autor.",
    },
]

# ---------------------------------------------------------------- semantic model (TMDL)

TABLES = {
    "dim_fecha": [("fecha", "dateTime", "date"), ("anio", "int64", "Int64.Type"), ("mes", "int64", "Int64.Type"),
                  ("mes_nombre", "string", "text"), ("trimestre", "string", "text"), ("anio_mes", "string", "text")],
    "dim_flujo": [("flujo_cod", "string", "text"), ("flujo", "string", "text")],
    "dim_pais": [("pais_cod", "int64", "Int64.Type"), ("iso3", "string", "text"), ("iso2", "string", "text"),
                 ("pais", "string", "text"), ("continente", "string", "text")],
    "dim_producto": [("capitulo", "string", "text"), ("producto", "string", "text"), ("sector", "string", "text"),
                     ("capitulo_producto", "string", "text")],
    "fact_balanza_mensual": [("fecha", "dateTime", "date"), ("flujo_cod", "string", "text"), ("grupo", "string", "text"),
                             ("valor_musd", "double", "number")],
    "fact_comercio_socios": [("fecha", "dateTime", "date"), ("flujo_cod", "string", "text"), ("pais_cod", "int64", "Int64.Type"),
                             ("valor_usd", "double", "number")],
    "fact_comercio_productos": [("fecha", "dateTime", "date"), ("flujo_cod", "string", "text"), ("capitulo", "string", "text"),
                                ("valor_usd", "double", "number")],
}
RELS = [("fact_balanza_mensual", "fecha", "dim_fecha", "fecha"), ("fact_balanza_mensual", "flujo_cod", "dim_flujo", "flujo_cod"),
        ("fact_comercio_socios", "fecha", "dim_fecha", "fecha"), ("fact_comercio_socios", "flujo_cod", "dim_flujo", "flujo_cod"),
        ("fact_comercio_socios", "pais_cod", "dim_pais", "pais_cod"),
        ("fact_comercio_productos", "fecha", "dim_fecha", "fecha"), ("fact_comercio_productos", "flujo_cod", "dim_flujo", "flujo_cod"),
        ("fact_comercio_productos", "capitulo", "dim_producto", "capitulo")]

# (nombre, expresión DAX, formato, carpeta)
MEASURES = [
    # --- año de análisis: el seleccionado, o el último completo si no hay selección
    ("Anio Balanza", """
VAR _completo =
    CALCULATE (
        MAXX (
            FILTER ( VALUES ( dim_fecha[anio] ), CALCULATE ( DISTINCTCOUNT ( fact_balanza_mensual[fecha] ) ) = 12 ),
            dim_fecha[anio]
        ),
        REMOVEFILTERS ( dim_fecha )
    )
RETURN IF ( HASONEVALUE ( dim_fecha[anio] ), VALUES ( dim_fecha[anio] ), _completo )""", "0", "0 Contexto"),
    ("Anio Comtrade", """
VAR _ultimo = CALCULATE ( YEAR ( MAX ( fact_comercio_socios[fecha] ) ), REMOVEFILTERS () )
VAR _sel = SELECTEDVALUE ( dim_fecha[anio] )
RETURN IF ( NOT ISBLANK ( _sel ) && _sel <= _ultimo, _sel, _ultimo )""", "0", "0 Contexto"),

    # --- Balanza (BCRP, millones -> miles de millones)
    ("Balanza mil M", "DIVIDE ( SUM ( fact_balanza_mensual[valor_musd] ), 1000 )", "#,0.0", "1 Balanza"),
    ("Exportaciones mil M", "CALCULATE ( [Balanza mil M], dim_flujo[flujo_cod] = \"X\" )", "#,0.0", "1 Balanza"),
    ("Importaciones mil M", "CALCULATE ( [Balanza mil M], dim_flujo[flujo_cod] = \"M\" )", "#,0.0", "1 Balanza"),
    ("Export anual completo", "IF ( CALCULATE ( DISTINCTCOUNT ( fact_balanza_mensual[fecha] ) ) = 12, [Exportaciones mil M] )", "#,0.0", "1 Balanza"),
    ("Import anual completo", "IF ( CALCULATE ( DISTINCTCOUNT ( fact_balanza_mensual[fecha] ) ) = 12, [Importaciones mil M] )", "#,0.0", "1 Balanza"),
    ("Export anio", "VAR _y = [Anio Balanza] RETURN CALCULATE ( [Exportaciones mil M], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y )", "#,0.0", "1 Balanza"),
    ("Import anio", "VAR _y = [Anio Balanza] RETURN CALCULATE ( [Importaciones mil M], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y )", "#,0.0", "1 Balanza"),
    ("Export anio anterior", "VAR _y = [Anio Balanza] RETURN CALCULATE ( [Exportaciones mil M], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y - 1 )", "#,0.0", "1 Balanza"),
    ("Import anio anterior", "VAR _y = [Anio Balanza] RETURN CALCULATE ( [Importaciones mil M], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y - 1 )", "#,0.0", "1 Balanza"),
    ("Superavit anio", "[Export anio] - [Import anio]", "#,0.0", "1 Balanza"),
    ("% grupo en el anio", """
VAR _y = [Anio Balanza]
VAR _v = CALCULATE ( [Balanza mil M], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y )
VAR _t = CALCULATE ( [Balanza mil M], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y, REMOVEFILTERS ( fact_balanza_mensual[grupo] ) )
RETURN DIVIDE ( _v, _t )""", "0.0%", "1 Balanza"),
    ("% export grupo", "CALCULATE ( [% grupo en el anio], dim_flujo[flujo_cod] = \"X\" )", "0.0%", "1 Balanza"),
    ("% import grupo", "CALCULATE ( [% grupo en el anio], dim_flujo[flujo_cod] = \"M\" )", "0.0%", "1 Balanza"),
    ("% Tradicionales", "CALCULATE ( [% export grupo], fact_balanza_mensual[grupo] = \"Tradicionales\" )", "0.0%", "1 Balanza"),
    ("% No tradicionales", "CALCULATE ( [% export grupo], fact_balanza_mensual[grupo] = \"No tradicionales\" )", "0.0%", "1 Balanza"),
    ("KPI Export", "[Export anio]", "#,0.0", "1 Balanza"),
    ("KPI Export nota", "VAR _v = DIVIDE ( [Export anio], [Export anio anterior] ) - 1 RETURN FORMAT ( _v, \"+0.0%;-0.0%\" ) & \"  vs \" & ( [Anio Balanza] - 1 ) & \" · año \" & [Anio Balanza]", "", "1 Balanza"),
    ("KPI Import", "[Import anio]", "#,0.0", "1 Balanza"),
    ("KPI Import nota", "VAR _v = DIVIDE ( [Import anio], [Import anio anterior] ) - 1 RETURN FORMAT ( _v, \"+0.0%;-0.0%\" ) & \"  vs \" & ( [Anio Balanza] - 1 ) & \" · año \" & [Anio Balanza]", "", "1 Balanza"),
    ("KPI Superavit", "[Superavit anio]", "#,0.0", "1 Balanza"),
    ("KPI Superavit nota", """
VAR _y = [Anio Balanza]
VAR _max =
    MAXX (
        FILTER ( ALL ( dim_fecha[anio] ), CALCULATE ( DISTINCTCOUNT ( fact_balanza_mensual[fecha] ) ) = 12 ),
        CALCULATE ( [Exportaciones mil M] - [Importaciones mil M] )
    )
RETURN IF ( [Superavit anio] >= _max, "Récord de la serie 2015–" & _y, "Récord: " & FORMAT ( _max, "#,0.0" ) & " mil M" )""", "", "1 Balanza"),
    ("Ultimo mes", "CALCULATE ( MAX ( fact_balanza_mensual[fecha] ), REMOVEFILTERS () )", "yyyy-mm-dd", "1 Balanza"),
    ("KPI Export YTD", """
VAR _f = [Ultimo mes]
RETURN CALCULATE ( [Exportaciones mil M], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = YEAR ( _f ), dim_fecha[mes] <= MONTH ( _f ) )""", "#,0.0", "1 Balanza"),
    ("KPI Export YTD nota", """
VAR _f = [Ultimo mes]
VAR _ant = CALCULATE ( [Exportaciones mil M], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = YEAR ( _f ) - 1, dim_fecha[mes] <= MONTH ( _f ) )
VAR _mes = LOWER ( LOOKUPVALUE ( dim_fecha[mes_nombre], dim_fecha[fecha], _f ) )
RETURN FORMAT ( DIVIDE ( [KPI Export YTD], _ant ) - 1, "+0.0%;-0.0%" ) & "  ene–" & _mes & " " & YEAR ( _f ) & " vs " & ( YEAR ( _f ) - 1 )""", "", "1 Balanza"),

    # --- Socios (Comtrade, US$ -> miles de millones)
    ("Socios mil M", "DIVIDE ( SUM ( fact_comercio_socios[valor_usd] ), 1e9 )", "#,0.0", "2 Socios"),
    ("Export socios", "CALCULATE ( [Socios mil M], dim_flujo[flujo_cod] = \"X\" )", "#,0.0", "2 Socios"),
    ("Import socios", "CALCULATE ( [Socios mil M], dim_flujo[flujo_cod] = \"M\" )", "#,0.0", "2 Socios"),
    ("Export socios anio", "VAR _y = [Anio Comtrade] RETURN CALCULATE ( [Export socios], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y )", "#,0.0", "2 Socios"),
    ("Import socios anio", "VAR _y = [Anio Comtrade] RETURN CALCULATE ( [Import socios], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y )", "#,0.0", "2 Socios"),
    ("% export pais", "DIVIDE ( [Export socios anio], CALCULATE ( [Export socios anio], REMOVEFILTERS ( dim_pais ) ) )", "0.0%", "2 Socios"),
    ("% import pais", "DIVIDE ( [Import socios anio], CALCULATE ( [Import socios anio], REMOVEFILTERS ( dim_pais ) ) )", "0.0%", "2 Socios"),
    ("Export top 10", "IF ( RANKX ( ALLSELECTED ( dim_pais ), [Export socios anio] ) <= 10, [Export socios anio] )", "#,0.0", "2 Socios"),
    ("Proveedor top 3 %", "IF ( RANKX ( ALL ( dim_pais ), [Import socios anio] ) <= 3, [% import pais] )", "0.0%", "2 Socios"),
    ("% China por anio", "DIVIDE ( CALCULATE ( [Export socios], dim_pais[pais] = \"China\" ), CALCULATE ( [Export socios], REMOVEFILTERS ( dim_pais ) ) )", "0.0%", "2 Socios"),
    ("% EE. UU. por anio", "DIVIDE ( CALCULATE ( [Export socios], dim_pais[pais] = \"Estados Unidos\" ), CALCULATE ( [Export socios], REMOVEFILTERS ( dim_pais ) ) )", "0.0%", "2 Socios"),
    ("KPI China", "CALCULATE ( [% export pais], dim_pais[pais] = \"China\" )", "0.0%", "2 Socios"),
    ("KPI China nota", """
VAR _primero = CALCULATE ( YEAR ( MIN ( fact_comercio_socios[fecha] ) ), REMOVEFILTERS () )
VAR _v = CALCULATE ( [% China por anio], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _primero )
RETURN "Era " & FORMAT ( _v, "0.0%" ) & " en " & _primero & " · año " & [Anio Comtrade]""", "", "2 Socios"),
    ("KPI Top5", """
VAR _t = CALCULATE ( [Export socios anio], REMOVEFILTERS ( dim_pais ) )
RETURN DIVIDE ( SUMX ( TOPN ( 5, ALL ( dim_pais ), [Export socios anio] ), [Export socios anio] ), _t )""", "0.0%", "2 Socios"),
    ("KPI Top5 nota", "\"entre \" & COUNTROWS ( FILTER ( ALL ( dim_pais ), [Export socios anio] > 0 ) ) & \" destinos\"", "", "2 Socios"),
    ("% Top 10", "DIVIDE ( SUMX ( TOPN ( 10, ALL ( dim_pais ), [Export socios anio] ), [Export socios anio] ), CALCULATE ( [Export socios anio], REMOVEFILTERS ( dim_pais ) ) )", "0.0%", "2 Socios"),
    ("KPI Asia", "CALCULATE ( [% export pais], dim_pais[continente] = \"Asia\" )", "0.0%", "2 Socios"),
    ("KPI Asia nota", "\"Norteamérica: \" & FORMAT ( CALCULATE ( [% export pais], dim_pais[continente] = \"Norteamérica\" ), \"0.0%\" )", "", "2 Socios"),
    ("KPI HHI", """
VAR _y = [Anio Comtrade]
VAR _t = CALCULATE ( SUM ( fact_comercio_socios[valor_usd] ), dim_flujo[flujo_cod] = "X", REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y, REMOVEFILTERS ( dim_pais ) )
RETURN
    SUMX (
        ALL ( dim_pais[pais_cod] ),
        VAR _v = CALCULATE ( SUM ( fact_comercio_socios[valor_usd] ), dim_flujo[flujo_cod] = "X", REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y )
        RETURN ( 100 * DIVIDE ( _v, _t ) ) ^ 2
    )""", "#,0", "2 Socios"),
    ("KPI HHI nota", "VAR _h = [KPI HHI] RETURN \"Concentración \" & IF ( _h > 2500, \"alta\", IF ( _h > 1500, \"moderada\", \"baja\" ) ) & \" · año \" & [Anio Comtrade]", "", "2 Socios"),
    ("China + EE. UU.", "CALCULATE ( [% export pais], dim_pais[pais] IN { \"China\", \"Estados Unidos\" } )", "0.0%", "2 Socios"),

    # --- Productos (Comtrade por capítulo HS)
    ("Productos mil M", "DIVIDE ( CALCULATE ( SUM ( fact_comercio_productos[valor_usd] ), dim_flujo[flujo_cod] = \"X\" ), 1e9 )", "#,0.0", "3 Productos"),
    ("Export prod anio", "VAR _y = [Anio Comtrade] RETURN CALCULATE ( [Productos mil M], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y )", "#,0.0", "3 Productos"),
    ("Export prod hace 5", "VAR _y = [Anio Comtrade] RETURN CALCULATE ( [Productos mil M], REMOVEFILTERS ( dim_fecha ), dim_fecha[anio] = _y - 5 )", "#,0.0", "3 Productos"),
    ("% export producto", "DIVIDE ( [Export prod anio], CALCULATE ( [Export prod anio], REMOVEFILTERS ( dim_producto ) ) )", "0.0%", "3 Productos"),
    ("Crecimiento anual 5a", "VAR _a = [Export prod hace 5] VAR _b = [Export prod anio] RETURN IF ( _a > 0, ( _b / _a ) ^ ( 1 / 5 ) - 1 )", "+0.0%;-0.0%", "3 Productos"),
    ("Crecimiento 5a", "VAR _a = [Export prod hace 5] VAR _b = [Export prod anio] RETURN IF ( _a > 0 && _b >= 0.1, _b / _a - 1 )", "+0%;-0%", "3 Productos"),
    ("Producto top 8", "IF ( RANKX ( ALLSELECTED ( dim_producto ), [Export prod anio] ) <= 8, [Export prod anio] )", "#,0.0", "3 Productos"),
    ("Top crecimiento", """
VAR _g = [Crecimiento 5a]
RETURN IF ( NOT ISBLANK ( _g ) && RANKX ( FILTER ( ALLSELECTED ( dim_producto ), NOT ISBLANK ( [Crecimiento 5a] ) ), [Crecimiento 5a] ) <= 5, _g )""", "+0%;-0%", "3 Productos"),
    ("Hace 5 anios (top)", "IF ( NOT ISBLANK ( [Top crecimiento] ), [Export prod hace 5] )", "#,0.00", "3 Productos"),
    ("En el anio (top)", "IF ( NOT ISBLANK ( [Top crecimiento] ), [Export prod anio] )", "#,0.00", "3 Productos"),
    ("KPI Mineria", "CALCULATE ( [% export producto], dim_producto[sector] = \"Minería y metales\" )", "0.0%", "3 Productos"),
    ("KPI Mineria nota", "\"US$ \" & FORMAT ( CALCULATE ( [Export prod anio], dim_producto[sector] = \"Minería y metales\" ), \"#,0.0\" ) & \" mil M · año \" & [Anio Comtrade]", "", "3 Productos"),
    ("KPI Agro", "CALCULATE ( [Export prod anio], dim_producto[sector] = \"Agro y alimentos\" )", "#,0.0", "3 Productos"),
    ("KPI Agro nota", "FORMAT ( CALCULATE ( [Crecimiento anual 5a], dim_producto[sector] = \"Agro y alimentos\" ), \"+0.0%;-0.0%\" ) & \" anual en 5 años\"", "", "3 Productos"),
    ("KPI Cacao", "CALCULATE ( [Export prod anio], dim_producto[capitulo] = \"18\" ) / CALCULATE ( [Export prod hace 5], dim_producto[capitulo] = \"18\" ) - 1", "+0%;-0%", "3 Productos"),
    ("KPI Cacao nota", "\"De US$ \" & FORMAT ( CALCULATE ( [Export prod hace 5], dim_producto[capitulo] = \"18\" ), \"0.00\" ) & \" a \" & FORMAT ( CALCULATE ( [Export prod anio], dim_producto[capitulo] = \"18\" ), \"0.00\" ) & \" mil M\"", "", "3 Productos"),
    ("KPI Frutas", "CALCULATE ( [Export prod anio], dim_producto[capitulo] = \"08\" )", "#,0.0", "3 Productos"),
    ("KPI Frutas nota", """
VAR _v = CALCULATE ( [Export prod anio], dim_producto[capitulo] = "08" )
VAR _r = COUNTROWS ( FILTER ( ALL ( dim_producto ), [Export prod anio] > _v ) ) + 1
RETURN "Puesto " & _r & " de " & COUNTROWS ( ALL ( dim_producto ) ) & " productos\"""", "", "3 Productos"),
]


def tmdl_table(name: str, cols: list[tuple[str, str, str]]) -> str:
    out = [f"table {name}", ""]
    for col, dtype, _ in cols:
        out += [f"\tcolumn {col}", f"\t\tdataType: {dtype}"]
        if dtype == "dateTime":
            out.append("\t\tformatString: yyyy-mm-dd")
        out += ["\t\tsummarizeBy: none", f"\t\tsourceColumn: {col}", ""]
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
    refs = "\n".join(f"ref table {t}" for t in [*TABLES, "Medidas"])
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
    (d / "tables" / "Medidas.tmdl").write_text(tmdl_measures(), encoding="utf-8")


# ---------------------------------------------------------------- theme

def theme() -> dict:
    txt = lambda size, color, font: {"fontSize": size, "fontFace": font, "color": color}
    no = [{"show": False}]
    return {
        "name": "Tierra y Pacífico",
        "dataColors": [C["cobre"], C["pacifico"], C["lucuma"], C["cochinilla"], C["indigo"], C["muna"], C["ink2"], C["muted"]],
        "foreground": C["ink"], "foregroundNeutralSecondary": C["ink2"], "foregroundNeutralTertiary": C["muted"],
        "background": C["card"], "backgroundLight": C["bg"], "backgroundNeutral": C["line"], "tableAccent": C["cobre"],
        "good": C["muna"], "neutral": C["lucuma"], "bad": C["cochinilla"],
        "textClasses": {
            "callout": txt(28, C["ink"], "DIN"), "title": txt(12, C["ink"], "Georgia"), "header": txt(11, C["ink"], "Segoe UI Semibold"),
            "label": txt(10, C["ink2"], "Segoe UI"),
        },
        "visualStyles": {
            "*": {"*": {
                "background": [{"show": False}], "border": no, "dropShadow": no, "title": no, "visualHeader": no,
                "visualTooltip": [{"show": True}], "padding": [{"top": 0, "bottom": 0, "left": 0, "right": 0}],
                "valueAxis": [{"gridlineColor": {"solid": {"color": C["line"]}}, "gridlineThickness": 1, "fontFamily": "Segoe UI", "labelColor": {"solid": {"color": C["ink2"]}}}],
                "categoryAxis": [{"fontFamily": "Segoe UI", "labelColor": {"solid": {"color": C["ink2"]}}}],
                "legend": [{"fontFamily": "Segoe UI", "labelColor": {"solid": {"color": C["ink"]}}}],
            }},
            "page": {"*": {"background": [{"transparency": 100}], "outspace": [{"color": {"solid": {"color": C["bg"]}}}]}},
            "slicer": {"*": {"background": [{"show": True, "color": {"solid": {"color": C["card"]}}, "transparency": 0}],
                             "border": [{"show": True, "color": {"solid": {"color": "#D6CDBF"}}, "radius": 4}]}},
        },
    }


# ---------------------------------------------------------------- PBIR helpers

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


def container(name, x, y, w, h, visual, z=1000):
    return {"$schema": f"{SCH}/item/report/definition/visualContainer/{V_VISUAL}/schema.json", "name": name,
            "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z}, "visual": visual}


def card(name, x, y, w, h, mname, size=28, col=None, font="DIN", bold=False):
    objs = {"labels": [{"properties": {"fontSize": lit(size), "fontFamily": lit(font), "color": color(col or C["ink"]), "bold": lit(bold)}}],
            "categoryLabels": [{"properties": {"show": lit(False)}}]}
    return container(name, x, y, w, h, {"visualType": "card", "query": {"queryState": {"Values": {"projections": [proj(measure(mname))]}}},
                                        "objects": objs, "drillFilterOtherVisuals": True})


def slicer(name, x, y, w, h, table, col, label):
    return container(name, x, y, w, h, {
        "visualType": "slicer",
        "query": {"queryState": {"Values": {"projections": [proj(column(table, col), label)]}}},
        "objects": {"data": [{"properties": {"mode": lit("Dropdown")}}],
                    "header": [{"properties": {"show": lit(True), "fontColor": color(C["ink2"]), "textSize": lit(9), "fontFamily": lit("Segoe UI")}}],
                    "items": [{"properties": {"fontColor": color(C["ink"]), "textSize": lit(10), "fontFamily": lit("Segoe UI")}}]},
        "drillFilterOtherVisuals": True}, z=5000)


def chart(name, x, y, w, h, vtype, buckets: dict, objects: dict | None = None, sort: tuple | None = None):
    q = {"queryState": {k: {"projections": v} for k, v in buckets.items()}}
    if sort:
        field, direction = sort
        q["sortDefinition"] = {"sort": [{"field": field, "direction": direction}], "isDefaultSort": False}
    return container(name, x, y, w, h, {"visualType": vtype, "query": q, "objects": objects or {}, "drillFilterOtherVisuals": True})


def series_colors(field: dict, pairs: list[tuple[str, str]]) -> list:
    """dataPoint por valor de la serie (el color sigue a la entidad, no a su posición)."""
    out = []
    for value, hex_ in pairs:
        out.append({"properties": {"fill": color(hex_)},
                    "selector": {"data": [{"scopeId": {"Comparison": {"ComparisonKind": 0, "Left": field, "Right": {"Literal": {"Value": f"'{value}'"}}}}}]}})
    return out


def measure_colors(pairs: list[tuple[str, str]]) -> list:
    return [{"properties": {"fill": color(hex_)}, "selector": {"metadata": f"Medidas.{m}"}} for m, hex_ in pairs]


def bar_objects(fill: str, fmt_units="0", label_fmt=None, cat_width=None) -> dict:
    o = {
        "dataPoint": [{"properties": {"fill": color(fill)}}],
        "labels": [{"properties": {"show": lit(True), "color": color(C["ink"]), "fontSize": lit(9), "fontFamily": lit("Segoe UI Semibold")}}],
        "valueAxis": [{"properties": {"show": lit(False), "gridlineShow": lit(False)}}],
        "categoryAxis": [{"properties": {"show": lit(True), "fontSize": lit(10), "labelColor": color(C["ink"]), "fontFamily": lit("Segoe UI")}}],
        "legend": [{"properties": {"show": lit(False)}}],
    }
    if cat_width:
        o["categoryAxis"][0]["properties"]["maxMarginFactor"] = lit(cat_width)
    return o


# ---------------------------------------------------------------- visuals per page

def kpi_visuals(page) -> list:
    out = []
    for i, (_, _, col, m, nota) in enumerate(page["kpis"]):
        x = kpi_x(i)
        out.append(card(f"kpi{i}", x + 8, KPI_Y + 24, 178, 44, m, size=26))
        out.append(card(f"kpinota{i}", x + 8, KPI_Y + 68, 178, 24, nota, size=9, col=C["ink2"], font="Segoe UI"))
    return out


def visuals_balanza(p) -> tuple[list, list]:
    v = kpi_visuals(p)
    anio = column("dim_fecha", "anio")
    v.append(chart("linea", 116, 262, 699, 420, "lineChart",
                   {"Category": [proj(anio, "Año")], "Y": [proj(measure("Export anual completo"), "Exportaciones"), proj(measure("Import anual completo"), "Importaciones")]},
                   {"dataPoint": measure_colors([("Export anual completo", C["cobre"]), ("Import anual completo", C["pacifico"])]),
                    "lineStyles": [{"properties": {"strokeWidth": lit(2), "showMarker": lit(True), "markerSize": lit(3)}}],
                    "legend": [{"properties": {"show": lit(False)}}],
                    "labels": [{"properties": {"show": lit(False)}}],
                    "categoryAxis": [{"properties": {"axisType": lit("Categorical"), "fontSize": lit(10)}}],
                    "valueAxis": [{"properties": {"fontSize": lit(10), "start": lit(0)}}]}))
    grupo = column("fact_balanza_mensual", "grupo")
    v.append(chart("export_grupos", 853, 254, 385, 40, "hundredPercentStackedBarChart",
                   {"Series": [proj(grupo, "Grupo")], "Y": [proj(measure("% export grupo"), "% del año")]},
                   {"dataPoint": series_colors(grupo, [("Tradicionales", C["cobre"]), ("No tradicionales", C["lucuma"]), ("Otros", C["ink2"])]),
                    "legend": [{"properties": {"show": lit(False)}}], "labels": [{"properties": {"show": lit(False)}}],
                    "valueAxis": [{"properties": {"show": lit(False)}}], "categoryAxis": [{"properties": {"show": lit(False)}}]}))
    v.append(card("pct_trad", 853, 316, 180, 36, "% Tradicionales", size=22))
    v.append(card("pct_notrad", 1045, 316, 180, 36, "% No tradicionales", size=22))
    v.append(chart("import_uso", 853, 440, 385, 118, "clusteredBarChart",
                   {"Category": [proj(grupo, "Uso")], "Y": [proj(measure("% import grupo"), "% del año")]},
                   bar_objects(C["pacifico"], cat_width=42), sort=(measure("% import grupo"), "Descending")))
    v.append(chart("proveedores", 853, 590, 236, 100, "tableEx",
                   {"Values": [proj(column("dim_pais", "pais"), "Proveedor"), proj(measure("Proveedor top 3 %"), "% importado")]},
                   {"grid": [{"properties": {"gridVertical": lit(False), "gridHorizontal": lit(True), "gridHorizontalColor": color(C["line"]),
                                             "rowPadding": lit(2), "outlineColor": color(C["line"])}}],
                    "columnHeaders": [{"properties": {"fontColor": color(C["ink2"]), "backColor": color(C["card"]), "fontFamily": lit("Segoe UI"), "fontSize": lit(9)}}],
                    "values": [{"properties": {"fontColor": color(C["ink"]), "backColor": color(C["card"]), "backColorSecondary": color(C["card"]), "fontSize": lit(10)}}],
                    "total": [{"properties": {"totals": lit(False)}}]},
                   sort=(measure("Proveedor top 3 %"), "Descending")))
    inter = [("slc_anio", "linea", "NoFilter")]
    return v, inter


def visuals_socios(p) -> tuple[list, list]:
    v = kpi_visuals(p)
    pais = column("dim_pais", "pais")
    v.append(chart("top10", 116, 250, 584, 360, "clusteredBarChart",
                   {"Category": [proj(pais, "País")], "Y": [proj(measure("Export top 10"), "Exportaciones")]},
                   bar_objects(C["cobre"], cat_width=28), sort=(measure("Export top 10"), "Descending")))
    v.append(card("china_eeuu", 116, 640, 180, 40, "China + EE. UU.", size=22))
    v.append(card("top10_pct", 330, 640, 180, 40, "% Top 10", size=22))
    anio = column("dim_fecha", "anio")
    v.append(chart("china_linea", 736, 270, 504, 152, "lineChart",
                   {"Category": [proj(anio, "Año")], "Y": [proj(measure("% China por anio"), "China"), proj(measure("% EE. UU. por anio"), "Estados Unidos")]},
                   {"dataPoint": measure_colors([("% China por anio", C["cochinilla"]), ("% EE. UU. por anio", C["pacifico"])]),
                    "lineStyles": [{"properties": {"strokeWidth": lit(2), "showMarker": lit(True), "markerSize": lit(3)}}],
                    "legend": [{"properties": {"show": lit(False)}}],
                    "categoryAxis": [{"properties": {"axisType": lit("Categorical"), "fontSize": lit(9)}}],
                    "valueAxis": [{"properties": {"fontSize": lit(9), "start": lit(0)}}]}))
    cont = column("dim_pais", "continente")
    v.append(chart("continentes", 736, 488, 504, 96, "hundredPercentStackedBarChart",
                   {"Series": [proj(cont, "Continente")], "Y": [proj(measure("% export pais"), "% del total")]},
                   {"dataPoint": series_colors(cont, [("Asia", C["cochinilla"]), ("Norteamérica", C["pacifico"]), ("Europa", C["indigo"]),
                                                      ("Latinoamérica", C["muna"]), ("Oceanía", C["muted"]), ("África", C["muted"]), ("Otros", C["muted"])]),
                    "legend": [{"properties": {"show": lit(True), "position": lit("Bottom"), "fontSize": lit(9)}}],
                    "labels": [{"properties": {"show": lit(True), "color": color("#FFFFFF"), "fontSize": lit(9), "labelPrecision": lit(1)}}],
                    "valueAxis": [{"properties": {"show": lit(False)}}], "categoryAxis": [{"properties": {"show": lit(False)}}]}))
    inter = [("slc_anio", "china_linea", "NoFilter"), ("slc_continente", "china_linea", "NoFilter")]
    return v, inter


def visuals_productos(p) -> tuple[list, list]:
    v = kpi_visuals(p)
    sector = column("dim_producto", "sector")
    v.append(chart("sectores", 116, 250, 350, 430, "clusteredBarChart",
                   {"Category": [proj(sector, "Sector")], "Y": [proj(measure("% export producto"), "% del total")],
                    "Tooltips": [proj(measure("Crecimiento anual 5a"), "Crecimiento anual (5 años)")]},
                   bar_objects(C["cobre"], cat_width=40), sort=(measure("% export producto"), "Descending")))
    prod = column("dim_producto", "producto")
    v.append(chart("top8", 503, 250, 350, 430, "clusteredBarChart",
                   {"Category": [proj(prod, "Producto")], "Y": [proj(measure("Producto top 8"), "Exportaciones")]},
                   bar_objects(C["lucuma"], cat_width=55), sort=(measure("Producto top 8"), "Descending")))
    v.append(chart("crecimiento", 890, 262, 350, 250, "tableEx",
                   {"Values": [proj(prod, "Producto"), proj(measure("Top crecimiento"), "Crecimiento en 5 años")]},
                   {"grid": [{"properties": {"gridVertical": lit(False), "gridHorizontal": lit(True), "gridHorizontalColor": color(C["line"]),
                                             "rowPadding": lit(6), "outlineColor": color(C["line"])}}],
                    "columnHeaders": [{"properties": {"fontColor": color(C["ink2"]), "backColor": color(C["card"]), "fontFamily": lit("Segoe UI"), "fontSize": lit(9), "wordWrap": lit(True)}}],
                    "values": [{"properties": {"fontColor": color(C["ink"]), "backColor": color(C["card"]), "backColorSecondary": color(C["card"]), "fontSize": lit(10), "wordWrap": lit(True)}}],
                    "total": [{"properties": {"totals": lit(False)}}]},
                   sort=(measure("Top crecimiento"), "Descending")))
    return v, []


BUILDERS = {"balanza": visuals_balanza, "socios": visuals_socios, "productos": visuals_productos}


# ---------------------------------------------------------------- backgrounds (HTML -> PNG)

def svg_icon(d: str, colr: str, size: int, stroke=1.5) -> str:
    return (f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{colr}" stroke-width="{stroke}" '
            f'stroke-linecap="round" stroke-linejoin="round"><path d="{d}"/></svg>')


TOCAPU = ('<svg width="20" height="320"><defs><pattern id="t" width="20" height="40" patternUnits="userSpaceOnUse">'
          f'<rect x="2" y="2" width="16" height="16" fill="none" stroke="{C["cobre"]}" stroke-width="1.5"/>'
          f'<rect x="7" y="7" width="6" height="6" fill="{C["lucuma"]}"/>'
          f'<path d="M2 22 L10 30 L18 22 M2 38 L10 30 L18 38" fill="none" stroke="{C["cochinilla"]}" stroke-width="1.5"/>'
          '</pattern></defs><rect width="20" height="320" fill="url(#t)"/></svg>')
INTI = ('<svg width="36" height="36" viewBox="0 0 36 36"><circle cx="18" cy="18" r="7" fill="none" stroke="#B88400" stroke-width="2"/>'
        '<circle cx="18" cy="18" r="2.5" fill="#C2562B"/><path d="M18 3v5 M18 28v5 M3 18h5 M28 18h5 M7.4 7.4l3.5 3.5 M25.1 25.1l3.5 3.5 '
        'M7.4 28.6l3.5-3.5 M25.1 10.9l3.5-3.5" stroke="#B88400" stroke-width="2" stroke-linecap="round"/></svg>')

EXTRAS = {
    "balanza": f"""
<div class="abs small" style="left:853px;top:300px;display:flex;align-items:center;gap:6px"><span style="width:8px;height:8px;border-radius:2px;background:#C2562B"></span>Tradicionales</div>
<div class="abs small" style="left:1045px;top:300px;display:flex;align-items:center;gap:6px"><span style="width:8px;height:8px;border-radius:2px;background:#B88400"></span>No tradicionales</div>
<div class="abs mini" style="left:853px;top:357px">cobre, oro, zinc</div>
<div class="abs mini" style="left:1045px;top:357px">agro, pesca, textil</div>
<div class="abs rule" style="left:853px;top:566px;width:385px"></div>
<div class="abs small" style="left:853px;top:572px">Principales proveedores</div>
<svg class="abs" style="left:1110px;top:600px" width="120" height="60" viewBox="0 0 92 46" fill="none" stroke-linecap="round" stroke-linejoin="round">
  <path d="M8 44V8h4v36 M8 8h44 M12 8l8 8 M40 8v10" stroke="#1E1B18" stroke-width="1.4"/>
  <rect x="34" y="18" width="12" height="7" stroke="#00819E" stroke-width="1.4"/><rect x="50" y="30" width="16" height="7" stroke="#C2562B" stroke-width="1.4"/>
  <rect x="68" y="30" width="16" height="7" stroke="#B88400" stroke-width="1.4"/><rect x="58" y="22" width="16" height="7" stroke="#9E2A4F" stroke-width="1.4"/>
  <path d="M2 44h88" stroke="#8A8178" stroke-width="1.4"/></svg>
<div class="abs mini" style="left:1100px;top:664px;width:140px;text-align:center">Puerto del Callao</div>""",
    "socios": f"""
<div class="abs rule" style="left:116px;top:618px;width:584px"></div>
<div class="abs small" style="left:116px;top:624px">China + Estados Unidos</div>
<div class="abs small" style="left:330px;top:624px">Top 10 destinos</div>
<div class="abs rule" style="left:736px;top:600px;width:504px"></div>
<svg class="abs" style="left:736px;top:614px" width="230" height="56" viewBox="0 0 200 44" fill="none" stroke-linecap="round" stroke-linejoin="round">
  <path d="M2 36c8-4 16-4 24 0s16 4 24 0 16-4 24 0 16 4 24 0 16-4 24 0 16 4 24 0 16-4 24 0 16 4 24 0" stroke="#00819E" stroke-width="1.6"/>
  <path d="M14 26 C 70 2, 130 2, 186 26" stroke="#9E2A4F" stroke-width="1.6" stroke-dasharray="3 4"/>
  <circle cx="14" cy="26" r="3.5" fill="#C2562B"/><circle cx="186" cy="26" r="3.5" fill="#9E2A4F"/>
  <path d="M92 16h16l-3 4h-10z M97 16v-5h5v5" stroke="#1E1B18" stroke-width="1.6"/></svg>
<div class="abs" style="left:984px;top:618px;width:250px;font-size:12px;line-height:1.4"><b>Callao → Asia</b><br>
<span style="color:#5E574E">Más de la mitad de lo exportado va a Asia, al otro lado del Pacífico.</span></div>""",
    "productos": f"""
<div class="abs" style="left:890px;top:584px;width:350px;height:96px;background:#F4EFE7;border-radius:6px"></div>
<svg class="abs" style="left:904px;top:596px" width="58" height="70" viewBox="0 0 44 52" fill="none" stroke-linecap="round" stroke-linejoin="round">
  <path d="M22 4c10 5 15 14 15 22s-5 17-15 22C12 43 7 34 7 26S12 9 22 4z" stroke="#9E2A4F" stroke-width="1.8"/>
  <path d="M22 4v44 M14.5 10c2.5 5 2.5 27 0 32 M29.5 10c-2.5 5-2.5 27 0 32" stroke="#9E2A4F" stroke-width="1.4"/>
  <path d="M22 4c0-2 2-3 4-3" stroke="#4E7A2E" stroke-width="1.8"/></svg>
<div class="abs" style="left:976px;top:600px;width:250px;font-size:12px;line-height:1.45">En 2024 el Perú exportó <b>más fruta (US$ 6.4 mil M)</b> que harina de pescado y café juntos (US$ 3.3 mil M).</div>""",
}


def background_html(p: dict) -> str:
    nav = ""
    for i, (num, label) in enumerate([("01", "BALANZA"), ("02", "SOCIOS"), ("03", "PRODUCTOS")]):
        on = i == p["nav"]
        nav += (f'<div style="width:56px;height:56px;display:flex;flex-direction:column;align-items:center;justify-content:center;'
                f'color:{"#F4EFE7" if on else "#B9B0A4"};{"border-bottom:2px solid #C2562B;" if on else ""}font-family:Bahnschrift">'
                f'<span style="font-size:18px;font-weight:600">{num}</span><span style="font-size:10px;letter-spacing:.06em">{label}</span></div>')
    kpis = ""
    for i, (label, icon, col, _, _) in enumerate(p["kpis"]):
        x = kpi_x(i)
        kpis += (f'<div class="abs cardbox" style="left:{x}px;top:{KPI_Y}px;width:278px;height:{KPI_H}px;border-top:3px solid {C[col]}"></div>'
                 f'<div class="abs small" style="left:{x + 8}px;top:{KPI_Y + 10}px;width:178px;text-align:center">{label}</div>'
                 f'<div class="abs" style="left:{x + 206}px;top:{KPI_Y + 21}px;width:58px;height:58px;border-radius:50%;border:1.5px solid {C[col]};'
                 f'display:flex;align-items:center;justify-content:center">{svg_icon(ICON[icon], C[col], 32)}</div>')
    cards = ""
    for c in p["cards"]:
        cards += (f'<div class="abs cardbox" style="left:{c["x"]}px;top:{c["y"]}px;width:{c["w"]}px;height:{c["h"]}px"></div>'
                  f'<div class="abs" style="left:{c["x"] + 16}px;top:{c["y"] + 13}px;font-family:Georgia;font-size:17px">{c["title"]}</div>'
                  f'<div class="abs mini" style="left:{c["x"] + 16}px;top:{c["y"] + 35}px">{c["sub"]}</div>')
        if c.get("legend"):
            lx = c["x"] + c["w"] - 16
            items = "".join(f'<span style="display:inline-flex;align-items:center;gap:6px;margin-left:14px"><span style="width:14px;height:3px;background:{C[k]}"></span>{t}</span>'
                            for t, k in c["legend"])
            cards += f'<div class="abs" style="right:{W - lx}px;top:{c["y"] + c.get("legend_dy", 16)}px;font-size:12px">{items}</div>'
    slicer_note = ""
    return f"""<!doctype html><html><head><meta charset="utf-8"><style>
html,body{{margin:0;width:{W}px;height:{H}px;overflow:hidden;background:{C["bg"]};color:{C["ink"]};font-family:'Segoe UI',sans-serif}}
.abs{{position:absolute}} .cardbox{{background:{C["card"]};border:1px solid {C["line"]};border-radius:6px;box-sizing:border-box}}
.small{{font-size:12px;color:{C["ink2"]}}} .mini{{font-size:11px;color:{C["ink2"]}}} .rule{{height:1px;background:{C["line"]}}}
</style></head><body>
<div class="abs" style="left:0;top:0;width:76px;height:{H}px;background:{C["ink"]};display:flex;flex-direction:column;align-items:center;padding-top:20px;box-sizing:border-box;gap:6px">
<div style="margin-bottom:14px">{INTI}</div>{nav}<div style="margin-top:auto">{TOCAPU}</div></div>
<div class="abs" style="left:{X0}px;top:22px;font-family:Bahnschrift;font-size:12px;letter-spacing:.14em;color:{C["cochinilla"]};font-weight:600">{p["kicker"]}</div>
<div class="abs" style="left:{X0}px;top:40px;width:820px;font-family:Georgia;font-size:26px;line-height:1.1;white-space:nowrap">{p["title"]}</div>
{kpis}{cards}{EXTRAS[p["extras"]]}{slicer_note}
<div class="abs mini" style="left:{X0}px;top:698px">{p["footer"]}</div>
<div class="abs mini" style="right:28px;top:698px">Carlos Matos Gavino · Data Analyst</div>
</body></html>"""


def render_backgrounds() -> None:
    BG.mkdir(exist_ok=True)
    for p in PAGES:
        html = BG / f"bg_{p['name']}.html"
        html.write_text(background_html(p), encoding="utf-8")
        png = BG / f"bg_{p['name']}.png"
        subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=2",
                        f"--window-size={W},{H}", f"--screenshot={png}", html.as_uri()], check=True, capture_output=True, timeout=120)
        print("fondo", png.name, png.stat().st_size)


# ---------------------------------------------------------------- report (PBIR)

def write_report() -> None:
    if RP.exists():
        shutil.rmtree(RP)
    d = RP / "definition"
    res = RP / "StaticResources" / "RegisteredResources"
    res.mkdir(parents=True)
    (d / "pages").mkdir(parents=True)
    j = lambda path, obj: (path.parent.mkdir(parents=True, exist_ok=True), path.write_text(json.dumps(obj, indent=2, ensure_ascii=False), encoding="utf-8"))
    j(RP / "definition.pbir", {"$schema": f"{SCH}/item/report/definitionProperties/2.0.0/schema.json", "version": "4.0",
                               "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})
    j(res / "TierraPacifico.json", theme())
    images = []
    for p in PAGES:
        src = BG / f"bg_{p['name']}.png"
        if src.exists():
            shutil.copy(src, res / src.name)
            images.append({"name": src.name, "path": src.name, "type": "Image"})
    ver = {"visual": V_VISUAL, "report": V_REPORT, "page": V_PAGE}
    j(d / "report.json", {
        "$schema": f"{SCH}/item/report/definition/report/{V_REPORT}/schema.json",
        "themeCollection": {"customTheme": {"name": "TierraPacifico.json", "reportVersionAtImport": ver, "type": "RegisteredResources"}},
        "resourcePackages": [{"name": "RegisteredResources", "type": "RegisteredResources",
                              "items": [{"name": "TierraPacifico.json", "path": "TierraPacifico.json", "type": "CustomTheme"}, *images]}],
        "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True,
                     "useEnhancedTooltips": True},
    })
    j(d / "version.json", {"$schema": f"{SCH}/item/report/definition/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    j(d / "pages" / "pages.json", {"$schema": f"{SCH}/item/report/definition/pagesMetadata/1.0.0/schema.json",
                                   "pageOrder": [p["name"] for p in PAGES], "activePageName": PAGES[0]["name"]})
    for p in PAGES:
        pd_ = d / "pages" / p["name"]
        visuals, inter = BUILDERS[p["name"]](p)
        # Segmentadores arriba a la derecha, de derecha a izquierda.
        sx = X0 + CW
        for key, table, col, label in reversed(p["slicers"]):
            wv = 128 if key != "anio" else 96
            sx -= wv
            visuals.append(slicer(f"slc_{key}", sx, 24, wv, 56, table, col, label))
            sx -= 10
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
        j(pd_ / "page.json", page)
        for vis in visuals:
            j(pd_ / "visuals" / vis["name"] / "visual.json", vis)
    j(HERE / f"{NAME}.pbip", {"$schema": f"{SCH}/pbip/pbipProperties/1.0.0/schema.json", "version": "1.0",
                              "artifacts": [{"report": {"path": f"{NAME}.Report"}}], "settings": {"enableAutoRecovery": True}})


if __name__ == "__main__":
    write_model()
    if "--no-png" not in sys.argv:
        render_backgrounds()
    write_report()
    print("PBIP listo:", HERE / f"{NAME}.pbip")
