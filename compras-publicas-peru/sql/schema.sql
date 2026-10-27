-- Modelo estrella de las compras públicas del Perú (SEACE / OCDS). SQLite.
-- Dos tablas de hechos: el proceso de contratación y la adjudicación a cada ganador.

DROP TABLE IF EXISTS fact_adjudicacion;
DROP TABLE IF EXISTS fact_proceso;
DROP TABLE IF EXISTS dim_fecha;
DROP TABLE IF EXISTS dim_entidad;
DROP TABLE IF EXISTS dim_proveedor;
DROP TABLE IF EXISTS dim_metodo;
DROP TABLE IF EXISTS dim_categoria;

CREATE TABLE dim_fecha (
    fecha       TEXT PRIMARY KEY,      -- YYYY-MM-DD
    anio        INTEGER NOT NULL,
    mes         INTEGER NOT NULL,
    mes_nombre  TEXT NOT NULL,
    trimestre   TEXT NOT NULL
);

CREATE TABLE dim_entidad (
    entidad_id      TEXT PRIMARY KEY,  -- código SEACE de la entidad compradora
    entidad         TEXT NOT NULL,
    departamento    TEXT NOT NULL,
    nivel_gobierno  TEXT NOT NULL      -- aproximado por el nombre de la entidad
);

CREATE TABLE dim_proveedor (
    proveedor_id  TEXT PRIMARY KEY,    -- RUC de empresas; id anónimo (PN-...) para personas naturales
    proveedor     TEXT NOT NULL,
    tipo          TEXT NOT NULL        -- Empresa / Persona natural / Otro
);

CREATE TABLE dim_metodo (
    metodo     TEXT PRIMARY KEY,       -- procedimiento de selección del SEACE
    tipo_ocds  TEXT NOT NULL           -- Abierto / Selectivo / Directo / Limitado (estándar OCDS)
);

CREATE TABLE dim_categoria (
    categoria_cod  TEXT PRIMARY KEY,
    categoria      TEXT NOT NULL       -- Bienes / Servicios / Obras
);

-- Un proceso de contratación por fila, fechado por su convocatoria.
CREATE TABLE fact_proceso (
    ocid                   TEXT PRIMARY KEY,
    fecha                  TEXT NOT NULL REFERENCES dim_fecha(fecha),
    entidad_id             TEXT NOT NULL REFERENCES dim_entidad(entidad_id),
    metodo                 TEXT NOT NULL REFERENCES dim_metodo(metodo),
    categoria_cod          TEXT REFERENCES dim_categoria(categoria_cod),
    valor_referencial_pen  REAL,
    n_participantes        REAL,       -- participantes registrados (no ofertas presentadas)
    monto_adjudicado_pen   REAL,
    dias_a_buena_pro       REAL,
    adjudicado             INTEGER NOT NULL,
    monto_atipico          INTEGER NOT NULL  -- adjudica más del triple del valor referencial
);

-- Una fila por ganador; en consorcios el monto se reparte en partes iguales.
CREATE TABLE fact_adjudicacion (
    ocid             TEXT NOT NULL REFERENCES fact_proceso(ocid),
    proveedor_id     TEXT NOT NULL REFERENCES dim_proveedor(proveedor_id),
    fecha_buena_pro  TEXT,
    monto_pen        REAL NOT NULL
);

CREATE INDEX ix_proceso_fecha ON fact_proceso(fecha);
CREATE INDEX ix_adj_ocid ON fact_adjudicacion(ocid);
CREATE INDEX ix_adj_prov ON fact_adjudicacion(proveedor_id);
