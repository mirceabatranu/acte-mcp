"""Configurarea: un fisier JSON in ~/.acte-mcp/config.json.

Parolele de aplicatie Gmail NU stau aici, ci in Windows Credential Manager (vezi parole.py).
Starea descarcarii (ce mesaje s-au procesat) sta in ~/.acte-mcp/stare.sqlite, pe discul
local: SQLite nu e sigur pe un disc de retea.
"""

from __future__ import annotations

import json
import os
from datetime import date
from pathlib import Path

HOME = Path(os.environ.get("ACTE_MCP_HOME", Path.home() / ".acte-mcp"))
CONFIG_PATH = HOME / "config.json"

DEFAULTS: dict = {
    # Radacina serverului local. Recomandat calea UNC (\\NAS\share), care merge si cand
    # unitatea Z: nu e conectata (de ex. in sarcina programata).
    "radacina": "Z:\\",
    # Casutele Gmail de descarcat; parola de aplicatie se salveaza cu `acte-mcp parola <email>`.
    "conturi": [],
    # Adresele cabinetului: nu identifica niciodata un client (la forward se cauta expeditorul original).
    "adrese_proprii": [],
    # Excelul cu clientii (foaia "Clienti", antet pe randul 4). Gol = doar dupa numele folderelor.
    "excel_clienti": "",
    # Mesajele primite inainte de aceasta data nu se descarca.
    "de_la": "2026-11-01",
    # Luna de lucru = luna mesajului minus atatea luni (1: mesaj din noiembrie -> octombrie).
    "luna_minus": 1,
    # Unde ajung fisierele, relativ la folderul clientului / la radacina.
    "folder_client": "Acte contabile\\{an}\\{luna}",
    "folder_neclasificat": "_De verificat\\{an}\\{luna}",
    # Folder tehnic pe radacina: jurnalul descarcarilor si marcajele "procesat", vizibile
    # de pe orice calculator care vede serverul.
    "folder_tehnic": "_acte-mcp",
    # Mapare manuala adresa sau domeniu -> numele folderului clientului de pe server.
    "mapare_manuala": {},
    # Expeditori ignorati (adresa sau domeniu, se potriveste ca subsir).
    "ignora_expeditori": [
        "monitorizarejuridica@wolterskluwer.ro",
        "newsletter@ceccar.ro",
        "noreply@aboutyou.ro",
        "easybox@sameday.ro",
        "gazeta@iabilete.ro",
        "office@eurobestteam.ro",
        "linkedin.com",
        "fancourier.ro",
        "office@sagasoft.ro",
        "stripe.com",
        "timedoctor.com",
        "acfe.com",
        "merchantpro",
        "mailer-daemon",
    ],
    "raport": {
        # Raportul zilnic si alertele se trimit de pe primul cont, catre aceasta adresa.
        "catre": "",
        # Ora de la care se trimite raportul pe ziua precedenta.
        "ora": 7,
        # Alerta de eroare se repeta cel mult o data la atatea minute.
        "alerta_minute": 120,
    },
}


def load() -> dict:
    cfg = json.loads(json.dumps(DEFAULTS))
    if CONFIG_PATH.exists():
        user = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        for k, v in user.items():
            if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                cfg[k].update(v)
            else:
                cfg[k] = v
    cfg["conturi"] = [c.strip().lower() for c in cfg["conturi"] if c.strip()]
    cfg["adrese_proprii"] = sorted({a.strip().lower() for a in cfg["adrese_proprii"] + cfg["conturi"]})
    return cfg


def save(cfg: dict) -> Path:
    HOME.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    return CONFIG_PATH


def root(cfg: dict) -> Path:
    return Path(cfg["radacina"])


def tech_dir(cfg: dict) -> Path:
    return root(cfg) / cfg["folder_tehnic"]


def work_month(d: date, minus: int) -> tuple[int, int]:
    """Luna de lucru: luna datei minus `minus` luni."""
    y, m = d.year, d.month - minus
    while m < 1:
        m += 12
        y -= 1
    return y, m


def month_path(template: str, year: int, month: int) -> Path:
    return Path(template.format(an=f"{year:04d}", luna=f"{month:02d}"))
