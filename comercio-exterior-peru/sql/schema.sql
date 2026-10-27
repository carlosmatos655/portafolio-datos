-- Modelo estrella (constelación) del comercio exterior del Perú. SQLite.
-- Tres tablas de hechos comparten las dimensiones de fecha y flujo.

DROP TABLE IF EXISTS fact_balanza_mensual;
DROP TABLE IF EXISTS fact_comercio_socios;
DROP TABLE IF EXISTS fact_comercio_productos;
DROP TABLE IF EXISTS dim_fecha;
DROP TABLE IF EXISTS dim_flujo;
DROP TABLE IF EXISTS dim_pais;
DROP TABLE IF EXISTS dim_producto;

CREATE TABLE dim_fecha (
    fecha       TEXT PRIMARY KEY,      -- primer día del mes, YYYY-MM-DD
    anio        INTEGER NOT NULL,
    mes         INTEGER NOT NULL,
    mes_nombre  TEXT NOT NULL,
    trimestre   TEXT NOT NULL,
    anio_mes    TEXT NOT NULL
);

CREATE TABLE dim_flujo (
    flujo_cod   TEXT PRIMARY KEY,      -- X exportación, M importación
    flujo       TEXT NOT NULL
);

CREATE TABLE dim_pais (
    pais_cod    INTEGER PRIMARY KEY,   -- código M49 de la ONU
    iso3        TEXT,
    iso2        TEXT,
    pais        TEXT NOT NULL,
    continente  TEXT NOT NULL
);

CREATE TABLE dim_producto (
    capitulo           TEXT PRIMARY KEY,   -- capítulo del Sistema Armonizado (2 dígitos)
    producto           TEXT NOT NULL,
    sector             TEXT NOT NULL,
    capitulo_producto  TEXT NOT NULL
);

-- BCRP, mensual, millones de US$ FOB por grupo (tradicionales, insumos, bienes de capital...).
CREATE TABLE fact_balanza_mensual (
    fecha       TEXT NOT NULL REFERENCES dim_fecha(fecha),
    flujo_cod   TEXT NOT NULL REFERENCES dim_flujo(flujo_cod),
    grupo       TEXT NOT NULL,
    valor_musd  REAL NOT NULL
);

-- ONU Comtrade, anual (fecha = 1 de enero), US$ por país socio.
CREATE TABLE fact_comercio_socios (
    fecha       TEXT NOT NULL REFERENCES dim_fecha(fecha),
    flujo_cod   TEXT NOT NULL REFERENCES dim_flujo(flujo_cod),
    pais_cod    INTEGER NOT NULL REFERENCES dim_pais(pais_cod),
    valor_usd   REAL NOT NULL
);

-- ONU Comtrade, anual (fecha = 1 de enero), US$ por capítulo HS.
CREATE TABLE fact_comercio_productos (
    fecha       TEXT NOT NULL REFERENCES dim_fecha(fecha),
    flujo_cod   TEXT NOT NULL REFERENCES dim_flujo(flujo_cod),
    capitulo    TEXT NOT NULL REFERENCES dim_producto(capitulo),
    valor_usd   REAL NOT NULL
);

CREATE INDEX ix_socios_fecha ON fact_comercio_socios(fecha, flujo_cod);
CREATE INDEX ix_productos_fecha ON fact_comercio_productos(fecha, flujo_cod);
CREATE INDEX ix_balanza_fecha ON fact_balanza_mensual(fecha, flujo_cod);
