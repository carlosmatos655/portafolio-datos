"""ETL - Comercio exterior del Perú.

Extrae datos públicos de dos APIs oficiales, los limpia con Pandas y escribe un modelo estrella
(constelación) en data/processed/ listo para SQL y Power BI.

Fuentes
  - BCRP (Banco Central de Reserva del Perú), API de series estadísticas: balanza comercial
    mensual, valores FOB en millones de US$.
  - ONU Comtrade, API pública (preview): comercio anual del Perú por país socio y por capítulo
    del Sistema Armonizado (HS), en US$.

Uso
  python etl/etl.py            # descarga (con caché en data/raw/) y transforma
  python etl/etl.py --refresh  # vuelve a descargar todo
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import pandas as pd
import requests
from babel import Locale
from babel.core import get_global

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
OUT = ROOT / "data" / "processed"
REFRESH = "--refresh" in sys.argv

PERU = 604
YEARS = range(2015, 2025)  # Comtrade publica el año completo con ~1 año de rezago
MONTH_FROM, MONTH_TO = "2015-1", "2026-12"

# Series BCRP de la balanza comercial (millones US$ FOB).
BCRP_SERIES = {
    "PN38715BM": ("Exportación", "Tradicionales"),
    "PN38716BM": ("Exportación", "No tradicionales"),
    "PN38717BM": ("Exportación", "Otros"),
    "PN38719BM": ("Importación", "Bienes de consumo"),
    "PN38720BM": ("Importación", "Insumos"),
    "PN38721BM": ("Importación", "Bienes de capital"),
    "PN38722BM": ("Importación", "Otros bienes"),
}
FLOWS = {"X": "Exportación", "M": "Importación"}
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic"]

# Capítulos del Sistema Armonizado: nombre corto en español y sector.
CAPITULOS = {
    "01": "Animales vivos", "02": "Carne", "03": "Pescados y mariscos", "04": "Lácteos, huevos y miel",
    "05": "Otros productos de origen animal", "06": "Plantas y flores", "07": "Hortalizas",
    "08": "Frutas y frutos secos", "09": "Café, té y especias", "10": "Cereales", "11": "Productos de molinería",
    "12": "Semillas y frutos oleaginosos", "13": "Gomas y resinas vegetales", "14": "Materias trenzables vegetales",
    "15": "Grasas y aceites", "16": "Conservas de carne y pescado", "17": "Azúcares y confitería",
    "18": "Cacao y sus preparaciones", "19": "Preparaciones de cereales", "20": "Conservas de hortalizas y frutas",
    "21": "Preparaciones alimenticias diversas", "22": "Bebidas", "23": "Harina de pescado y alimento animal",
    "24": "Tabaco", "25": "Sal, azufre, piedras y cemento", "26": "Minerales metalíferos",
    "27": "Combustibles y aceites minerales", "28": "Productos químicos inorgánicos",
    "29": "Productos químicos orgánicos", "30": "Productos farmacéuticos", "31": "Abonos",
    "32": "Tintes, pinturas y tintas", "33": "Aceites esenciales y cosméticos", "34": "Jabones y ceras",
    "35": "Materias albuminoideas y colas", "36": "Explosivos", "37": "Productos fotográficos",
    "38": "Productos químicos diversos", "39": "Plásticos", "40": "Caucho", "41": "Cueros y pieles",
    "42": "Manufacturas de cuero", "43": "Peletería", "44": "Madera", "45": "Corcho", "46": "Cestería",
    "47": "Pasta de celulosa", "48": "Papel y cartón", "49": "Productos editoriales", "50": "Seda",
    "51": "Lana y pelo fino", "52": "Algodón", "53": "Otras fibras vegetales", "54": "Filamentos sintéticos",
    "55": "Fibras sintéticas discontinuas", "56": "Guata, fieltro y cordelería", "57": "Alfombras",
    "58": "Tejidos especiales", "59": "Tejidos recubiertos", "60": "Tejidos de punto",
    "61": "Prendas de vestir de punto", "62": "Prendas de vestir, excepto de punto",
    "63": "Otros artículos textiles", "64": "Calzado", "65": "Sombreros", "66": "Paraguas y bastones",
    "67": "Plumas y flores artificiales", "68": "Manufacturas de piedra y cemento", "69": "Cerámica",
    "70": "Vidrio", "71": "Oro, plata y piedras preciosas", "72": "Hierro y acero",
    "73": "Manufacturas de hierro o acero", "74": "Cobre", "75": "Níquel", "76": "Aluminio",
    "78": "Plomo", "79": "Zinc", "80": "Estaño", "81": "Otros metales comunes", "82": "Herramientas",
    "83": "Manufacturas diversas de metal", "84": "Máquinas y artefactos mecánicos",
    "85": "Máquinas y aparatos eléctricos", "86": "Material ferroviario", "87": "Vehículos",
    "88": "Aeronaves", "89": "Barcos", "90": "Instrumentos de óptica y medida", "91": "Relojería",
    "92": "Instrumentos musicales", "93": "Armas y municiones", "94": "Muebles", "95": "Juguetes y deportes",
    "96": "Manufacturas diversas", "97": "Objetos de arte", "99": "No clasificado",
}


def sector(cap: str) -> str:
    c = int(cap)
    if c in (26, 71, 74, 75, 78, 79, 80, 81) or c == 25:
        return "Minería y metales"
    if c == 27:
        return "Hidrocarburos"
    if c in (3, 16, 23):
        return "Pesca"
    if 1 <= c <= 24:
        return "Agro y alimentos"
    if 50 <= c <= 65:
        return "Textil y confecciones"
    if 28 <= c <= 40:
        return "Químicos y plásticos"
    if c in (72, 73, 76, 82, 83, 84, 85, 86, 87, 88, 89, 90):
        return "Metalmecánica y maquinaria"
    return "Otros"


def get_json(url: str, cache: Path, pause: float = 0.0) -> dict:
    """GET con caché en disco y reintentos ante 429/5xx."""
    if cache.exists() and not REFRESH:
        return json.loads(cache.read_text(encoding="utf-8"))
    for intento in range(5):
        r = requests.get(url, timeout=60, headers={"User-Agent": "portafolio-datos/1.0 (uso educativo)"})
        if r.status_code == 429 or r.status_code >= 500:
            time.sleep(2 ** intento)
            continue
        r.raise_for_status()
        data = r.json()
        cache.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        time.sleep(pause)
        return data
    raise RuntimeError(f"Falló la descarga: {url}")


# ---------------------------------------------------------------- extract

def extract_bcrp() -> pd.DataFrame:
    codes = "-".join(BCRP_SERIES)
    url = f"https://estadisticas.bcrp.gob.pe/estadisticas/series/api/{codes}/json/{MONTH_FROM}/{MONTH_TO}"
    data = get_json(url, RAW / "bcrp_balanza.json")
    rows = []
    for p in data["periods"]:
        mes_txt, anio = p["name"].split(".")
        mes = [m.lower()[:3] for m in MESES].index(mes_txt.lower()[:3].replace("sep", "set")) + 1
        for code, val in zip(BCRP_SERIES, p["values"]):
            if val in ("n.d.", "", None):
                continue
            flujo, grupo = BCRP_SERIES[code]
            rows.append({"fecha": f"{anio}-{mes:02d}-01", "flujo": flujo, "grupo": grupo, "valor_musd": float(val)})
    return pd.DataFrame(rows)


def comtrade(params: str, cache: str) -> list[dict]:
    url = f"https://comtradeapi.un.org/public/v1/preview/C/A/HS?reporterCode={PERU}&{params}"
    return get_json(url, RAW / cache, pause=1.5).get("data", [])


def extract_comtrade() -> tuple[pd.DataFrame, pd.DataFrame]:
    socios, capitulos = [], []
    for anio in YEARS:
        for code in FLOWS:
            for r in comtrade(f"period={anio}&cmdCode=TOTAL&flowCode={code}", f"ct_socios_{anio}_{code}.json"):
                if r["partnerCode"] != 0:
                    socios.append({"anio": anio, "flujo_cod": code, "pais_cod": r["partnerCode"], "valor_usd": r["primaryValue"]})
            for r in comtrade(f"period={anio}&partnerCode=0&cmdCode=AG2&flowCode={code}", f"ct_capitulos_{anio}_{code}.json"):
                # El peso (netWgt) no viene en este nivel de agregación, solo el valor.
                capitulos.append({"anio": anio, "flujo_cod": code, "capitulo": r["cmdCode"], "valor_usd": r["primaryValue"]})
    return pd.DataFrame(socios), pd.DataFrame(capitulos)


def extract_paises() -> pd.DataFrame:
    data = get_json("https://comtradeapi.un.org/files/v1/app/reference/partnerAreas.json", RAW / "ct_partner_areas.json")
    return pd.DataFrame(data["results"])


# ---------------------------------------------------------------- transform

def continente(iso2: str | None) -> str:
    if not iso2:
        return "Otros"
    if iso2 in ("US", "CA", "MX"):
        return "Norteamérica"
    # Prefijo de zona horaria más frecuente (Chile tiene Pacific/Easter, pero es America/...).
    prefs = [z.split("/")[0] for z in get_global("territory_zones").get(iso2, [])]
    prio = ["America", "Europe", "Asia", "Africa", "Australia", "Pacific", "Indian", "Atlantic"]
    # Empates (Ecuador/Galápagos, España/Ceuta/Canarias): gana el continente principal.
    pref = max(prefs, key=lambda p: (prefs.count(p), -prio.index(p) if p in prio else -99)) if prefs else ""
    return {"America": "Latinoamérica", "Europe": "Europa", "Asia": "Asia", "Africa": "África",
            "Australia": "Oceanía", "Pacific": "Oceanía", "Indian": "Asia", "Atlantic": "Europa"}.get(pref, "Otros")


def build(bcrp: pd.DataFrame, socios: pd.DataFrame, caps: pd.DataFrame, areas: pd.DataFrame) -> dict[str, pd.DataFrame]:
    es = Locale("es")

    # dim_fecha: calendario mensual para los hechos mensuales y anuales.
    fechas = pd.date_range(f"{min(YEARS)}-01-01", bcrp["fecha"].max(), freq="MS")
    dim_fecha = pd.DataFrame({"fecha": fechas})
    dim_fecha["anio"] = dim_fecha.fecha.dt.year
    dim_fecha["mes"] = dim_fecha.fecha.dt.month
    dim_fecha["mes_nombre"] = dim_fecha.mes.map(lambda m: MESES[m - 1])
    dim_fecha["trimestre"] = "T" + dim_fecha.fecha.dt.quarter.astype(str)
    dim_fecha["anio_mes"] = dim_fecha.fecha.dt.strftime("%Y-%m")

    dim_flujo = pd.DataFrame({"flujo_cod": list(FLOWS), "flujo": list(FLOWS.values())})

    # dim_pais: solo países que aparecen en los hechos; nombre en español (CLDR) y continente.
    used = set(socios.pais_cod)
    a = areas[areas.PartnerCode.isin(used)].copy()
    a["iso2"] = a.PartnerCodeIsoAlpha2.where(a.PartnerCodeIsoAlpha2.str.len() == 2)
    sin_iso = {"Other Asia, nes": "Otros Asia (n.e.p.)", "Areas, nes": "Zonas no especificadas", "Free Zones": "Zonas francas"}
    a["pais"] = [es.territories.get(i, d) if isinstance(i, str) else sin_iso.get(d, d) for i, d in zip(a.iso2, a.PartnerDesc)]
    a["continente"] = a.iso2.map(continente)
    dim_pais = a.rename(columns={"PartnerCode": "pais_cod", "PartnerCodeIsoAlpha3": "iso3"})[["pais_cod", "iso3", "iso2", "pais", "continente"]]
    dim_pais = dim_pais.drop_duplicates("pais_cod").sort_values("pais")

    dim_producto = pd.DataFrame({"capitulo": list(CAPITULOS), "producto": list(CAPITULOS.values())})
    dim_producto["sector"] = dim_producto.capitulo.map(sector)
    dim_producto["capitulo_producto"] = dim_producto.capitulo + " · " + dim_producto.producto

    # Hechos. Comtrade es anual: se fecha al 1 de enero del año para usar la misma dimensión.
    fact_mensual = bcrp.merge(dim_flujo, on="flujo")[["fecha", "flujo_cod", "grupo", "valor_musd"]]
    fact_socios = socios.assign(fecha=pd.to_datetime(socios.anio.astype(str) + "-01-01").dt.strftime("%Y-%m-%d"))
    fact_socios = fact_socios[fact_socios.pais_cod.isin(dim_pais.pais_cod)][["fecha", "flujo_cod", "pais_cod", "valor_usd"]]
    fact_productos = caps.assign(fecha=pd.to_datetime(caps.anio.astype(str) + "-01-01").dt.strftime("%Y-%m-%d"))
    fact_productos = fact_productos[fact_productos.capitulo.isin(dim_producto.capitulo)][["fecha", "flujo_cod", "capitulo", "valor_usd"]]

    dim_fecha["fecha"] = dim_fecha.fecha.dt.strftime("%Y-%m-%d")
    return {"dim_fecha": dim_fecha, "dim_flujo": dim_flujo, "dim_pais": dim_pais, "dim_producto": dim_producto,
            "fact_balanza_mensual": fact_mensual, "fact_comercio_socios": fact_socios, "fact_comercio_productos": fact_productos}


def checks(t: dict[str, pd.DataFrame]) -> None:
    """Controles de calidad: claves únicas, sin huérfanos y totales coherentes entre fuentes."""
    assert t["dim_pais"].pais_cod.is_unique and t["dim_producto"].capitulo.is_unique and t["dim_fecha"].fecha.is_unique
    for f, k, d in [("fact_comercio_socios", "pais_cod", "dim_pais"), ("fact_comercio_productos", "capitulo", "dim_producto")]:
        assert t[f][k].isin(t[d][k]).all(), f"huérfanos en {f}"
    for f in ("fact_balanza_mensual", "fact_comercio_socios", "fact_comercio_productos"):
        assert t[f].fecha.isin(t["dim_fecha"].fecha).all(), f"fechas fuera de calendario en {f}"
    # BCRP vs Comtrade: exportaciones 2023 deberían diferir poco (metodologías distintas).
    b = t["fact_balanza_mensual"]
    bx = b[(b.flujo_cod == "X") & b.fecha.str.startswith("2023")].valor_musd.sum()
    cx = t["fact_comercio_productos"].query("flujo_cod == 'X' and fecha == '2023-01-01'").valor_usd.sum() / 1e6
    print(f"Exportaciones 2023 - BCRP: {bx:,.0f} M US$ | Comtrade: {cx:,.0f} M US$ | dif {100 * (cx / bx - 1):+.1f}%")


def main() -> None:
    RAW.mkdir(parents=True, exist_ok=True)
    OUT.mkdir(parents=True, exist_ok=True)
    bcrp = extract_bcrp()
    socios, caps = extract_comtrade()
    tablas = build(bcrp, socios, caps, extract_paises())
    checks(tablas)
    for name, df in tablas.items():
        df.to_csv(OUT / f"{name}.csv", index=False, encoding="utf-8")
        print(f"{name:26s} {len(df):>6,} filas")


if __name__ == "__main__":
    main()
