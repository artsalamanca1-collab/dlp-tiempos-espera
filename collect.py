#!/usr/bin/env python3
"""
Recoge los tiempos de espera de Disneyland Paris y los anade al CSV del dia.

Fuente principal: themeparks.wiki  (standby + single rider + premier access)
Fuente de respaldo: queue-times.com (solo standby)

Uso:
    python collect.py              # una muestra, la anade a data/YYYY-MM-DD.csv
    python collect.py --dry-run    # la imprime por pantalla, no escribe
"""

import argparse
import csv
import json
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

TP_BASE = "https://api.themeparks.wiki/v1"
QT_BASE = "https://queue-times.com"
DEST_SLUG = "disneylandparis"
QT_PARKS = {4: "Disneyland Park", 28: "Disney Adventure World"}
PARIS = ZoneInfo("Europe/Paris")
UA = "dlp-wait-times/1.0 (proyecto personal)"
DATA_DIR = Path(__file__).parent / "data"

COLUMNS = [
    "timestamp_utc", "fecha", "hora", "minuto", "parque", "atraccion",
    "atraccion_id", "estado", "tipo_cola", "espera_min", "hora_retorno",
    "precio", "fuente",
]


def get_json(url, timeout=25):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def stamp():
    now_utc = datetime.now(timezone.utc)
    now_paris = now_utc.astimezone(PARIS)
    return {
        "timestamp_utc": now_utc.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "fecha": now_paris.strftime("%Y-%m-%d"),
        "hora": now_paris.hour,
        "minuto": now_paris.minute,
        "_dia": now_paris.strftime("%Y-%m-%d"),
    }


# ---------------------------------------------------------------- themeparks

def find_parks():
    """Resuelve los IDs de los dos parques por slug, sin hardcodear UUIDs."""
    data = get_json(f"{TP_BASE}/destinations")
    for dest in data.get("destinations", []):
        if dest.get("slug") == DEST_SLUG:
            return [(p["id"], p["name"]) for p in dest.get("parks", [])]
    raise RuntimeError(f"No se encontro el destino con slug '{DEST_SLUG}'")


def collect_themeparks(ts):
    rows = []
    for park_id, park_name in find_parks():
        live = get_json(f"{TP_BASE}/entity/{park_id}/live")
        for entry in live.get("liveData") or []:
            if entry.get("entityType") != "ATTRACTION":
                continue
            base = {
                "timestamp_utc": ts["timestamp_utc"],
                "fecha": ts["fecha"],
                "hora": ts["hora"],
                "minuto": ts["minuto"],
                "parque": park_name,
                "atraccion": entry.get("name", ""),
                "atraccion_id": entry.get("id", ""),
                "estado": entry.get("status", ""),
                "fuente": "themeparks.wiki",
            }
            queue = entry.get("queue") or {}
            if not queue:
                rows.append({**base, "tipo_cola": "SIN_COLA", "espera_min": "",
                             "hora_retorno": "", "precio": ""})
                continue
            for qtype, q in queue.items():
                q = q or {}
                precio = ""
                p = q.get("price")
                if isinstance(p, dict) and p.get("amount") is not None:
                    precio = f"{p['amount'] / 100:.2f} {p.get('currency', '')}".strip()
                rows.append({
                    **base,
                    "tipo_cola": qtype,
                    "espera_min": q.get("waitTime") if q.get("waitTime") is not None else "",
                    "hora_retorno": q.get("returnStart") or q.get("state") or "",
                    "precio": precio,
                })
    return rows


# --------------------------------------------------------------- queue-times

def collect_queuetimes(ts):
    rows = []
    for park_id, park_name in QT_PARKS.items():
        data = get_json(f"{QT_BASE}/parks/{park_id}/queue_times.json")
        rides = list(data.get("rides") or [])
        for land in data.get("lands") or []:
            rides.extend(land.get("rides") or [])
        for ride in rides:
            rows.append({
                "timestamp_utc": ts["timestamp_utc"],
                "fecha": ts["fecha"],
                "hora": ts["hora"],
                "minuto": ts["minuto"],
                "parque": park_name,
                "atraccion": ride.get("name", ""),
                "atraccion_id": ride.get("id", ""),
                "estado": "OPERATING" if ride.get("is_open") else "CLOSED",
                "tipo_cola": "STANDBY",
                "espera_min": ride.get("wait_time", ""),
                "hora_retorno": "",
                "precio": "",
                "fuente": "queue-times.com",
            })
    return rows


# --------------------------------------------------------------------- main

def append(rows, dia):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    path = DATA_DIR / f"{dia}.csv"
    nuevo = not path.exists()
    with path.open("a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        if nuevo:
            w.writeheader()
        w.writerows(rows)
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    ts = stamp()
    rows, error = [], None
    try:
        rows = collect_themeparks(ts)
        if not rows:
            raise RuntimeError("themeparks.wiki devolvio 0 atracciones")
    except Exception as e:
        error = e
        print(f"[aviso] themeparks.wiki fallo ({e}); probando queue-times", file=sys.stderr)
        try:
            rows = collect_queuetimes(ts)
        except Exception as e2:
            print(f"[error] las dos fuentes fallaron: {e} / {e2}", file=sys.stderr)
            return 1

    if args.dry_run:
        abiertas = [r for r in rows if r["tipo_cola"] == "STANDBY" and r["espera_min"] != ""]
        for r in sorted(abiertas, key=lambda x: -int(x["espera_min"]))[:15]:
            print(f"{r['parque'][:22]:<24} {r['atraccion'][:38]:<40} {r['espera_min']:>3} min")
        print(f"\n{len(rows)} filas ({ts['fecha']} {ts['hora']:02d}:{ts['minuto']:02d} Paris)")
        return 0

    path = append(rows, ts["_dia"])
    print(f"{len(rows)} filas anadidas a {path.name}" + (" [respaldo]" if error else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
