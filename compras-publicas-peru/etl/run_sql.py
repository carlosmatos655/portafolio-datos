"""Carga data/processed/*.csv en SQLite (sql/schema.sql) y ejecuta sql/analisis.sql.

Uso: python etl/run_sql.py   -> crea data/compras.db e imprime cada consulta como tabla.
"""
import sqlite3
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / "data" / "compras.db"
TABLES = ["dim_fecha", "dim_entidad", "dim_proveedor", "dim_metodo", "dim_categoria",
          "fact_proceso", "fact_adjudicacion"]


def main() -> None:
    con = sqlite3.connect(DB)
    con.executescript((ROOT / "sql" / "schema.sql").read_text(encoding="utf-8"))
    for t in TABLES:
        df = pd.read_csv(ROOT / "data" / "processed" / f"{t}.csv", dtype={"entidad_id": str})
        df.to_sql(t, con, if_exists="append", index=False)
    con.commit()

    pd.set_option("display.width", 160)
    for block in (ROOT / "sql" / "analisis.sql").read_text(encoding="utf-8").split("-- name:")[1:]:
        title, _, query = block.partition("\n")
        print(f"\n### {title.strip()}\n")
        print(pd.read_sql_query(query, con).to_string(index=False))
    con.close()


if __name__ == "__main__":
    main()
