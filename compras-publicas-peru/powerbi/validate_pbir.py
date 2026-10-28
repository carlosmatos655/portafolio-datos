"""Valida los JSON del reporte PBIR contra sus esquemas oficiales (microsoft/json-schemas).

Uso: python powerbi/validate_pbir.py   -> imprime cada error de esquema; sale con 1 si hay alguno.
Los esquemas se descargan de GitHub y se guardan en powerbi/.schema-cache/ (ignorado por git).
"""
import json
import sys
from pathlib import Path

import requests
from jsonschema import Draft7Validator
from referencing import Registry, Resource

HERE = Path(__file__).resolve().parent
CACHE = HERE / ".schema-cache"
PREFIX = "https://developer.microsoft.com/json-schemas/"
RAW = "https://raw.githubusercontent.com/microsoft/json-schemas/main/"


def fetch(uri: str) -> dict:
    local = CACHE / uri.replace(PREFIX, "").replace("/", "_")
    if not local.exists():
        CACHE.mkdir(exist_ok=True)
        r = requests.get(uri.replace(PREFIX, RAW), timeout=60)
        r.raise_for_status()
        local.write_text(r.text, encoding="utf-8")
    return json.loads(local.read_text(encoding="utf-8"))


registry = Registry(retrieve=lambda uri: Resource.from_contents(fetch(uri)))
errors = 0
for f in sorted((HERE).glob("*.Report/**/*.json")):
    data = json.loads(f.read_text(encoding="utf-8"))
    sch = data.get("$schema")
    if not sch:
        continue
    v = Draft7Validator(fetch(sch), registry=registry)
    for e in v.iter_errors(data):
        errors += 1
        print(f"{f.relative_to(HERE)} :: {'/'.join(map(str, e.absolute_path))} :: {e.message[:300]}")
print("errores:", errors)
sys.exit(1 if errors else 0)
