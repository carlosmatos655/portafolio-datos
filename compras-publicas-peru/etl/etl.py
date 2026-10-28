"""ETL - Compras públicas del Perú (SEACE, estándar OCDS).

Fuente: registro de datos del Open Contracting Partnership, que publica las contrataciones del
OECE (antes OSCE) en formato OCDS. Licencia CC BY 4.0.
https://data.open-contracting.org/en/publication/135

Cada archivo anual agrupa los procesos por la fecha de su ÚLTIMA actualización, no por la de su
convocatoria: un proceso convocado en 2025 y actualizado en 2026 solo está en el archivo 2026.
Por eso se leen 2024, 2025 y 2026, se conserva la versión más reciente de cada proceso (ocid) y
recién entonces se filtra por año de convocatoria.

Privacidad: no se publica ningún correo ni teléfono, y los proveedores que son personas naturales
(RUC 10 o 15) se muestran como "Persona natural" con un identificador anónimo.

Uso
  python etl/etl.py   # lee data/raw/*.csv.tar.gz y escribe data/processed/*.csv
"""
from __future__ import annotations

import hashlib
import tarfile
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
OUT = ROOT / "data" / "processed"
FILES = (2024, 2025, 2026)          # años de actualización a leer
YEARS = ("2024", "2025", "2026")    # años de convocatoria a analizar (2026 parcial)

CATEGORIA = {"goods": "Bienes", "services": "Servicios", "works": "Obras"}
METODO_OCDS = {"open": "Abierto", "selective": "Selectivo", "direct": "Directo", "limited": "Limitado"}
PERSONA_NATURAL = ("10", "15")   # RUC de personas naturales; 16 y 17 son códigos que el SEACE da a consorcios
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic"]


def read(year: int, name: str, usecols: list[str]) -> pd.DataFrame:
    """Lee un CSV del .tar.gz anual sin descomprimir todo el archivo."""
    with tarfile.open(RAW / f"{year}.csv.tar.gz") as tar:
        f = tar.extractfile(f"{year}/{name}.csv")
        df = pd.read_csv(f, usecols=usecols, low_memory=False, encoding="utf-8")
    return df.assign(archivo=year)


def read_all(name: str, usecols: list[str]) -> pd.DataFrame:
    return pd.concat([read(y, name, usecols) for y in FILES], ignore_index=True)


def anonimo(ruc_id: str) -> str:
    return "PN-" + hashlib.sha1(ruc_id.encode()).hexdigest()[:10]


def nivel_gobierno(nombre: str) -> str:
    """Aproximación por el nombre de la entidad (el dato no trae el nivel de gobierno)."""
    n = nombre.upper()
    if "MUNICIPALIDAD" in n:
        return "Gobierno local"
    if "GOBIERNO REGIONAL" in n or n.startswith("REGION ") or "DIRECCION REGIONAL" in n or "UNIDAD DE GESTION EDUCATIVA LOCAL" in n:
        return "Gobierno regional"
    return "Gobierno nacional"


def build() -> dict[str, pd.DataFrame]:
    # --- procesos: versión más reciente de cada ocid
    main = read_all("main", ["ocid", "date", "tender_datePublished", "buyer_id", "buyer_name", "tender_procurementMethod",
                             "tender_procurementMethodDetails", "tender_mainProcurementCategory", "tender_value_amount_PEN",
                             "tender_numberOfTenderers", "tender_title"])
    main = main.sort_values("date").drop_duplicates("ocid", keep="last")
    main["anio"] = main.tender_datePublished.str[:4]
    main = main[main.anio.isin(YEARS) & main.buyer_id.notna()].copy()
    version = main[["ocid", "archivo"]]
    print("procesos:", main.anio.value_counts().sort_index().to_dict())

    def same_version(df: pd.DataFrame) -> pd.DataFrame:
        """Solo filas de la misma versión (archivo) elegida para cada proceso."""
        return df.merge(version, left_on=["main_ocid", "archivo"], right_on=["ocid", "archivo"]).drop(columns="ocid")

    # --- adjudicaciones, en soles
    aw = same_version(read_all("awards", ["id", "date", "value_amount", "value_currency", "main_ocid"]))
    fx = same_version(read_all("awards_value_exchangeRates", ["rate", "currency", "main_ocid", "awards_id"]))
    fx = fx[fx.currency == "PEN"].drop_duplicates(["main_ocid", "awards_id"])[["main_ocid", "awards_id", "rate"]]
    aw = aw.merge(fx, left_on=["main_ocid", "id"], right_on=["main_ocid", "awards_id"], how="left")
    aw["monto_pen"] = aw.value_amount.where(aw.value_currency == "PEN", aw.value_amount * aw.rate)
    sin_tc = aw.monto_pen.isna() & aw.value_amount.notna()
    print(f"adjudicaciones: {len(aw):,} | en otra moneda sin tipo de cambio (excluidas): {sin_tc.sum()}")
    aw = aw[aw.monto_pen.notna() & (aw.monto_pen > 0)]

    # --- ganadores: el monto se reparte en partes iguales entre los miembros de un consorcio
    sup = same_version(read_all("awards_suppliers", ["id", "name", "main_ocid", "awards_id"]))
    sup = sup.merge(aw[["main_ocid", "id", "date", "monto_pen"]].rename(columns={"id": "awards_id", "date": "fecha_buena_pro"}),
                    on=["main_ocid", "awards_id"])
    sup["n_miembros"] = sup.groupby(["main_ocid", "awards_id"]).id.transform("count")
    sup["monto_pen"] = sup.monto_pen / sup.n_miembros
    ruc = sup.id.str.replace("PE-RUC-", "", regex=False)
    natural = ruc.str[:2].isin(PERSONA_NATURAL)
    # Muchos consorcios figuran con el RUC 10 de un miembro: cuenta el nombre, pero el RUC se sigue anonimizando.
    consorcio = ruc.str[:2].isin(("16", "17")) | sup.name.fillna("").str.upper().str.contains("CONSORCIO")
    sup["tipo"] = "Otro (no domiciliado u otro)"
    sup.loc[ruc.str[:2] == "20", "tipo"] = "Empresa"
    sup.loc[natural, "tipo"] = "Persona natural"
    sup.loc[consorcio, "tipo"] = "Consorcio"
    sup["proveedor_id"] = sup.id.where(~natural, sup.id.map(anonimo))
    # Con RUC de persona natural no se publica el nombre (algunos consorcios llevan el de un miembro).
    sup["proveedor"] = sup.name.where(~natural, "Persona natural").where(~(natural & consorcio), "Consorcio")

    # --- participantes registrados (el campo numberOfTenderers del SEACE cuenta participantes, no ofertas)
    ten = same_version(read_all("tender_tenderers", ["id", "main_ocid"]))
    n_post = ten.groupby("main_ocid").id.nunique()
    main["n_participantes"] = main.tender_numberOfTenderers.fillna(main.ocid.map(n_post))

    # --- entidades compradoras con su departamento (tabla parties, solo el rol buyer)
    par = []
    for y in FILES:
        with tarfile.open(RAW / f"{y}.csv.tar.gz") as tar:
            f = tar.extractfile(f"{y}/parties.csv")
            for chunk in pd.read_csv(f, usecols=["id", "roles", "address_department"], chunksize=500_000, low_memory=False):
                b = chunk[chunk.roles.fillna("").str.contains("buyer")]
                par.append(b[["id", "address_department"]])
    dep = pd.concat(par).dropna().drop_duplicates("id", keep="last").set_index("id").address_department

    # ---------------------------------------------------------------- tablas del modelo
    fecha = pd.to_datetime(main.tender_datePublished.str[:10])
    first_aw = aw.dropna(subset=["date"]).assign(date=lambda d: d.date.astype(str).str[:10]).groupby("main_ocid").date.min()
    adjudicado = aw.groupby("main_ocid").monto_pen.sum()
    fact_proceso = pd.DataFrame({
        "ocid": main.ocid,
        "fecha": fecha.dt.strftime("%Y-%m-%d"),
        "entidad_id": main.buyer_id,
        "metodo": main.tender_procurementMethodDetails.fillna("Sin dato"),
        "categoria_cod": main.tender_mainProcurementCategory,
        "valor_referencial_pen": main.tender_value_amount_PEN,
        "n_participantes": main.n_participantes,
        "monto_adjudicado_pen": main.ocid.map(adjudicado),
    })
    dias = (pd.to_datetime(main.ocid.map(first_aw)) - fecha).dt.days
    fact_proceso["dias_a_buena_pro"] = dias.where(dias.between(0, 730))
    fact_proceso["adjudicado"] = fact_proceso.monto_adjudicado_pen.notna().astype(int)
    # Atípico: adjudica más del triple del valor referencial (suele ser un lote mal cargado); se marca, no se borra.
    fact_proceso["monto_atipico"] = (fact_proceso.monto_adjudicado_pen > 3 * fact_proceso.valor_referencial_pen).astype(int)

    fact_adjudicacion = pd.DataFrame({
        # El id de la adjudicación del SEACE incluye el RUC del ganador: no se publica.
        "ocid": sup.main_ocid, "proveedor_id": sup.proveedor_id,
        "fecha_buena_pro": sup.fecha_buena_pro.str[:10], "monto_pen": sup.monto_pen.round(2),
    })
    fact_adjudicacion = fact_adjudicacion[fact_adjudicacion.ocid.isin(fact_proceso.ocid)]

    ent = main.drop_duplicates("buyer_id", keep="last")
    dim_entidad = pd.DataFrame({"entidad_id": ent.buyer_id, "entidad": ent.buyer_name.str.strip(),
                                "departamento": ent.buyer_id.map(dep).fillna("SIN DATO").str.title(),
                                "nivel_gobierno": ent.buyer_name.fillna("").map(nivel_gobierno)})
    prov = sup.drop_duplicates("proveedor_id", keep="last")
    dim_proveedor = pd.DataFrame({"proveedor_id": prov.proveedor_id, "proveedor": prov.proveedor.str.strip(), "tipo": prov.tipo})
    met = main.drop_duplicates("tender_procurementMethodDetails")
    dim_metodo = pd.DataFrame({"metodo": met.tender_procurementMethodDetails.fillna("Sin dato"),
                               "tipo_ocds": met.tender_procurementMethod.map(METODO_OCDS)}).drop_duplicates("metodo")
    # Los procedimientos de la Ley 32069 llegan sin tipo OCDS: se asigna por su naturaleza.
    sin_tipo = dim_metodo.tipo_ocds.isna()
    dim_metodo.loc[sin_tipo, "tipo_ocds"] = dim_metodo.loc[sin_tipo, "metodo"].map(
        lambda m: "Directo" if "Directa" in m else ("Sin dato" if m == "Sin dato" else "Abierto"))
    dim_categoria = pd.DataFrame({"categoria_cod": list(CATEGORIA), "categoria": list(CATEGORIA.values())})
    dias_cal = pd.date_range(f"{YEARS[0]}-01-01", fecha.max(), freq="D")
    dim_fecha = pd.DataFrame({"fecha": dias_cal.strftime("%Y-%m-%d"), "anio": dias_cal.year, "mes": dias_cal.month,
                              "mes_nombre": [MESES[m - 1] for m in dias_cal.month], "trimestre": "T" + dias_cal.quarter.astype(str)})

    return {"dim_fecha": dim_fecha, "dim_entidad": dim_entidad, "dim_proveedor": dim_proveedor, "dim_metodo": dim_metodo,
            "dim_categoria": dim_categoria, "fact_proceso": fact_proceso, "fact_adjudicacion": fact_adjudicacion}


def checks(t: dict[str, pd.DataFrame]) -> None:
    """Claves únicas, sin huérfanos y sin datos de contacto ni nombres de personas naturales."""
    for d, k in [("dim_entidad", "entidad_id"), ("dim_proveedor", "proveedor_id"), ("dim_metodo", "metodo"), ("dim_fecha", "fecha")]:
        assert t[d][k].is_unique, f"clave duplicada en {d}"
    p, a = t["fact_proceso"], t["fact_adjudicacion"]
    assert p.ocid.is_unique
    assert p.entidad_id.isin(t["dim_entidad"].entidad_id).all() and p.metodo.isin(t["dim_metodo"].metodo).all()
    assert p.fecha.isin(t["dim_fecha"].fecha).all() and a.proveedor_id.isin(t["dim_proveedor"].proveedor_id).all()
    pn = t["dim_proveedor"].query("tipo == 'Persona natural'")
    assert (pn.proveedor == "Persona natural").all() and pn.proveedor_id.str.startswith("PN-").all()
    anon = t["dim_proveedor"].query("proveedor_id.str.startswith('PN-')", engine="python")
    assert anon.proveedor.isin(["Persona natural", "Consorcio"]).all(), "nombre publicado con RUC de persona natural"
    # Ningún campo publicado debe contener un RUC de persona natural (10/15/16/17 + 9 dígitos).
    for name, df in t.items():
        for c in df.select_dtypes("object"):
            assert not df[c].astype(str).str.contains(r"(?<!\d)1[05]\d{9}(?!\d)").any(), f"RUC de persona natural en {name}.{c}"
    # El monto repartido entre ganadores debe cuadrar con el adjudicado por proceso.
    adj = p[p.adjudicado == 1].set_index("ocid").monto_adjudicado_pen
    gan = a.groupby("ocid").monto_pen.sum().reindex(adj.index).fillna(0)
    print(f"monto adjudicado sin ganador identificado: {(adj - gan).clip(lower=0).sum() / adj.sum():.2%}")
    print(f"monto en procesos atípicos (> 3 x referencial): {p.loc[p.monto_atipico == 1, 'monto_adjudicado_pen'].sum() / adj.sum():.2%}")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    t = build()
    checks(t)
    for name, df in t.items():
        df.to_csv(OUT / f"{name}.csv", index=False, encoding="utf-8")
        print(f"{name:20s} {len(df):>8,} filas")


if __name__ == "__main__":
    main()
