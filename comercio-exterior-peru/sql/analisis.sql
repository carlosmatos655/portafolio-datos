-- Consultas de análisis sobre comercio.db (SQLite). Ejecutar con: python etl/run_sql.py
-- Cada bloque lleva una línea de título (name) que run_sql.py usa para imprimirlo.

-- name: 1. Balanza comercial anual (BCRP) y variación interanual
WITH anual AS (
    SELECT d.anio,
           SUM(CASE WHEN f.flujo_cod = 'X' THEN f.valor_musd END) AS exportaciones,
           SUM(CASE WHEN f.flujo_cod = 'M' THEN f.valor_musd END) AS importaciones,
           COUNT(DISTINCT d.mes) AS meses
    FROM fact_balanza_mensual f
    JOIN dim_fecha d ON d.fecha = f.fecha
    GROUP BY d.anio
)
SELECT anio,
       meses,
       ROUND(exportaciones) AS export_musd,
       ROUND(importaciones) AS import_musd,
       ROUND(exportaciones - importaciones) AS balanza_musd,
       CASE WHEN meses = 12 AND LAG(meses) OVER (ORDER BY anio) = 12   -- 2026 aún incompleto
            THEN ROUND(100.0 * (exportaciones / LAG(exportaciones) OVER (ORDER BY anio) - 1), 1) END AS var_export_pct
FROM anual
ORDER BY anio;

-- name: 2. Top 10 destinos de exportación 2024 y su participación
WITH x AS (
    SELECT p.pais, p.continente, f.valor_usd
    FROM fact_comercio_socios f
    JOIN dim_pais p ON p.pais_cod = f.pais_cod
    WHERE f.flujo_cod = 'X' AND f.fecha = '2024-01-01'
)
SELECT RANK() OVER (ORDER BY valor_usd DESC) AS puesto,
       pais,
       continente,
       ROUND(valor_usd / 1e6) AS musd,
       ROUND(100.0 * valor_usd / SUM(valor_usd) OVER (), 1) AS participacion_pct,
       ROUND(100.0 * SUM(valor_usd) OVER (ORDER BY valor_usd DESC) / SUM(valor_usd) OVER (), 1) AS acumulado_pct
FROM x
ORDER BY valor_usd DESC
LIMIT 10;

-- name: 3. Peso de China en las exportaciones, 2015-2024
SELECT d.anio,
       ROUND(SUM(CASE WHEN p.pais = 'China' THEN f.valor_usd END) / 1e6) AS china_musd,
       ROUND(100.0 * SUM(CASE WHEN p.pais = 'China' THEN f.valor_usd END) / SUM(f.valor_usd), 1) AS china_pct,
       ROUND(100.0 * SUM(CASE WHEN p.pais = 'Estados Unidos' THEN f.valor_usd END) / SUM(f.valor_usd), 1) AS eeuu_pct
FROM fact_comercio_socios f
JOIN dim_pais p ON p.pais_cod = f.pais_cod
JOIN dim_fecha d ON d.fecha = f.fecha
WHERE f.flujo_cod = 'X'
GROUP BY d.anio
ORDER BY d.anio;

-- name: 4. Concentración de destinos (índice Herfindahl-Hirschman, 0-10 000)
SELECT d.anio,
       ROUND(SUM(s.cuota * s.cuota)) AS hhi,
       CASE WHEN SUM(s.cuota * s.cuota) > 2500 THEN 'Alta'
            WHEN SUM(s.cuota * s.cuota) > 1500 THEN 'Moderada'
            ELSE 'Baja' END AS concentracion
FROM (
    SELECT f.fecha, 100.0 * f.valor_usd / SUM(f.valor_usd) OVER (PARTITION BY f.fecha) AS cuota
    FROM fact_comercio_socios f
    WHERE f.flujo_cod = 'X'
) s
JOIN dim_fecha d ON d.fecha = s.fecha
GROUP BY d.anio
ORDER BY d.anio;

-- name: 5. Exportaciones por sector, 2019 vs 2024 y crecimiento anual compuesto
WITH s AS (
    SELECT pr.sector,
           SUM(CASE WHEN f.fecha = '2019-01-01' THEN f.valor_usd END) AS v2019,
           SUM(CASE WHEN f.fecha = '2024-01-01' THEN f.valor_usd END) AS v2024
    FROM fact_comercio_productos f
    JOIN dim_producto pr ON pr.capitulo = f.capitulo
    WHERE f.flujo_cod = 'X'
    GROUP BY pr.sector
)
SELECT sector,
       ROUND(v2019 / 1e6) AS musd_2019,
       ROUND(v2024 / 1e6) AS musd_2024,
       ROUND(100.0 * (POWER(v2024 / v2019, 1.0 / 5) - 1), 1) AS cagr_pct,
       ROUND(100.0 * v2024 / SUM(v2024) OVER (), 1) AS peso_2024_pct
FROM s
ORDER BY v2024 DESC;

-- name: 6. Productos de exportación que más crecieron (2019-2024, mínimo 100 M US$ en 2024)
WITH p AS (
    SELECT pr.capitulo_producto,
           SUM(CASE WHEN f.fecha = '2019-01-01' THEN f.valor_usd END) AS v2019,
           SUM(CASE WHEN f.fecha = '2024-01-01' THEN f.valor_usd END) AS v2024
    FROM fact_comercio_productos f
    JOIN dim_producto pr ON pr.capitulo = f.capitulo
    WHERE f.flujo_cod = 'X'
    GROUP BY pr.capitulo_producto
)
SELECT capitulo_producto,
       ROUND(v2019 / 1e6) AS musd_2019,
       ROUND(v2024 / 1e6) AS musd_2024,
       ROUND(100.0 * (v2024 / v2019 - 1)) AS crecimiento_pct
FROM p
WHERE v2024 >= 1e8 AND v2019 > 0
ORDER BY v2024 / v2019 DESC
LIMIT 8;

-- name: 7. Composición de las importaciones por uso (BCRP), 2025
SELECT f.grupo,
       ROUND(SUM(f.valor_musd)) AS musd,
       ROUND(100.0 * SUM(f.valor_musd) / SUM(SUM(f.valor_musd)) OVER (), 1) AS peso_pct
FROM fact_balanza_mensual f
JOIN dim_fecha d ON d.fecha = f.fecha
WHERE f.flujo_cod = 'M' AND d.anio = 2025
GROUP BY f.grupo
ORDER BY musd DESC;

-- name: 8. Estacionalidad: exportaciones no tradicionales promedio por mes (2015-2025)
SELECT d.mes,
       d.mes_nombre,
       ROUND(AVG(f.valor_musd)) AS promedio_musd
FROM fact_balanza_mensual f
JOIN dim_fecha d ON d.fecha = f.fecha
WHERE f.grupo = 'No tradicionales' AND d.anio BETWEEN 2015 AND 2025
GROUP BY d.mes, d.mes_nombre
ORDER BY d.mes;
