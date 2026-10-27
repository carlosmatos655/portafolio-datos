-- Consultas de análisis sobre compras.db (SQLite). Ejecutar con: python etl/run_sql.py
-- Cada bloque lleva una línea de título (name) que run_sql.py usa para imprimirlo.
-- 2026 tiene datos parciales: las comparaciones anuales usan 2024 y 2025.

-- name: 1. Resumen anual: procesos, adjudicación y tiempo hasta la buena pro
SELECT d.anio,
       COUNT(*) AS procesos,
       SUM(p.adjudicado) AS adjudicados,
       ROUND(100.0 * SUM(p.adjudicado) / COUNT(*), 1) AS pct_adjudicados,
       ROUND(SUM(p.monto_adjudicado_pen) / 1e9, 2) AS adjudicado_mil_m,
       ROUND(AVG(p.dias_a_buena_pro), 1) AS dias_promedio_buena_pro
FROM fact_proceso p
JOIN dim_fecha d ON d.fecha = p.fecha
GROUP BY d.anio
ORDER BY d.anio;

-- name: 2. Monto adjudicado por categoría, 2024 vs 2025
WITH c AS (
    SELECT cat.categoria,
           SUM(CASE WHEN d.anio = 2024 THEN p.monto_adjudicado_pen END) AS v2024,
           SUM(CASE WHEN d.anio = 2025 THEN p.monto_adjudicado_pen END) AS v2025
    FROM fact_proceso p
    JOIN dim_fecha d ON d.fecha = p.fecha
    JOIN dim_categoria cat ON cat.categoria_cod = p.categoria_cod
    GROUP BY cat.categoria
)
SELECT categoria,
       ROUND(v2024 / 1e9, 2) AS mil_m_2024,
       ROUND(v2025 / 1e9, 2) AS mil_m_2025,
       ROUND(100.0 * (v2025 / v2024 - 1), 1) AS variacion_pct,
       ROUND(100.0 * v2025 / SUM(v2025) OVER (), 1) AS peso_2025_pct
FROM c
ORDER BY v2025 DESC;

-- name: 3. Competencia por tipo de procedimiento (estándar OCDS), 2025
SELECT m.tipo_ocds,
       COUNT(*) AS procesos,
       ROUND(100.0 * SUM(p.monto_adjudicado_pen) / SUM(SUM(p.monto_adjudicado_pen)) OVER (), 1) AS pct_monto,
       ROUND(AVG(p.n_participantes), 1) AS participantes_promedio,
       ROUND(100.0 * SUM(CASE WHEN p.n_participantes = 1 THEN 1 ELSE 0 END)
             / SUM(CASE WHEN p.n_participantes IS NOT NULL THEN 1 ELSE 0 END), 1) AS pct_un_participante
FROM fact_proceso p
JOIN dim_fecha d ON d.fecha = p.fecha
JOIN dim_metodo m ON m.metodo = p.metodo
WHERE d.anio = 2025
GROUP BY m.tipo_ocds
ORDER BY procesos DESC;

-- name: 4. Procedimientos con más procesos de un solo participante, 2025 (mínimo 500 procesos)
SELECT p.metodo,
       COUNT(*) AS procesos,
       ROUND(100.0 * SUM(CASE WHEN p.n_participantes = 1 THEN 1 ELSE 0 END)
             / SUM(CASE WHEN p.n_participantes IS NOT NULL THEN 1 ELSE 0 END), 1) AS pct_un_participante,
       ROUND(AVG(p.n_participantes), 1) AS participantes_promedio
FROM fact_proceso p
JOIN dim_fecha d ON d.fecha = p.fecha
WHERE d.anio = 2025
GROUP BY p.metodo
HAVING COUNT(*) >= 500
ORDER BY pct_un_participante DESC;

-- name: 5. Top 10 empresas proveedoras del Estado, 2025, y su participación acumulada
WITH g AS (
    SELECT pr.proveedor, SUM(a.monto_pen) AS monto
    FROM fact_adjudicacion a
    JOIN fact_proceso p ON p.ocid = a.ocid
    JOIN dim_fecha d ON d.fecha = p.fecha
    JOIN dim_proveedor pr ON pr.proveedor_id = a.proveedor_id
    WHERE d.anio = 2025 AND pr.tipo = 'Empresa'
    GROUP BY pr.proveedor_id, pr.proveedor
),
t AS (SELECT SUM(a.monto_pen) AS total
      FROM fact_adjudicacion a JOIN fact_proceso p ON p.ocid = a.ocid JOIN dim_fecha d ON d.fecha = p.fecha
      WHERE d.anio = 2025)
SELECT RANK() OVER (ORDER BY monto DESC) AS puesto,
       SUBSTR(proveedor, 1, 50) AS proveedor,
       ROUND(monto / 1e6, 1) AS millones_soles,
       ROUND(100.0 * monto / t.total, 2) AS pct_total,
       ROUND(100.0 * SUM(monto) OVER (ORDER BY monto DESC) / t.total, 2) AS acumulado_pct
FROM g, t
ORDER BY monto DESC
LIMIT 10;

-- name: 6. Concentración de proveedores (índice Herfindahl-Hirschman, 0-10 000)
WITH s AS (
    SELECT d.anio, a.proveedor_id, SUM(a.monto_pen) AS monto
    FROM fact_adjudicacion a
    JOIN fact_proceso p ON p.ocid = a.ocid
    JOIN dim_fecha d ON d.fecha = p.fecha
    GROUP BY d.anio, a.proveedor_id
)
SELECT anio,
       COUNT(*) AS proveedores_ganadores,
       ROUND(SUM(POWER(100.0 * monto / tot, 2)), 1) AS hhi
FROM (SELECT anio, monto, SUM(monto) OVER (PARTITION BY anio) AS tot FROM s)
GROUP BY anio
ORDER BY anio;

-- name: 7. Personas naturales frente a empresas, 2025
SELECT pr.tipo,
       COUNT(DISTINCT a.proveedor_id) AS proveedores,
       COUNT(*) AS adjudicaciones,
       ROUND(100.0 * SUM(a.monto_pen) / SUM(SUM(a.monto_pen)) OVER (), 1) AS pct_monto,
       ROUND(AVG(a.monto_pen) / 1e3, 1) AS miles_soles_promedio
FROM fact_adjudicacion a
JOIN fact_proceso p ON p.ocid = a.ocid
JOIN dim_fecha d ON d.fecha = p.fecha
JOIN dim_proveedor pr ON pr.proveedor_id = a.proveedor_id
WHERE d.anio = 2025
GROUP BY pr.tipo
ORDER BY pct_monto DESC;

-- name: 8. Monto adjudicado por departamento de la entidad compradora, 2025 (top 10)
SELECT e.departamento,
       COUNT(*) AS procesos,
       ROUND(SUM(p.monto_adjudicado_pen) / 1e9, 2) AS mil_m,
       ROUND(100.0 * SUM(p.monto_adjudicado_pen) / SUM(SUM(p.monto_adjudicado_pen)) OVER (), 1) AS pct
FROM fact_proceso p
JOIN dim_fecha d ON d.fecha = p.fecha
JOIN dim_entidad e ON e.entidad_id = p.entidad_id
WHERE d.anio = 2025
GROUP BY e.departamento
ORDER BY SUM(p.monto_adjudicado_pen) DESC
LIMIT 10;

-- name: 9. Nivel de gobierno, 2025
SELECT e.nivel_gobierno,
       COUNT(*) AS procesos,
       ROUND(100.0 * SUM(p.monto_adjudicado_pen) / SUM(SUM(p.monto_adjudicado_pen)) OVER (), 1) AS pct_monto
FROM fact_proceso p
JOIN dim_fecha d ON d.fecha = p.fecha
JOIN dim_entidad e ON e.entidad_id = p.entidad_id
WHERE d.anio = 2025
GROUP BY e.nivel_gobierno
ORDER BY pct_monto DESC;

-- name: 10. Estacionalidad: procesos convocados por mes (promedio 2024-2025)
SELECT d.mes,
       d.mes_nombre,
       ROUND(COUNT(*) / 2.0) AS procesos_promedio,
       ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 1) AS pct_del_anio
FROM fact_proceso p
JOIN dim_fecha d ON d.fecha = p.fecha
WHERE d.anio IN (2024, 2025)
GROUP BY d.mes, d.mes_nombre
ORDER BY d.mes;
