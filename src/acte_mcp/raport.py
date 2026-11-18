"""Raportul zilnic si alertele, pe e-mail (SMTP Gmail, cu aceeasi parola de aplicatie)."""

from __future__ import annotations

import json
import logging
import os
import smtplib
from collections import defaultdict
from datetime import date, datetime, timedelta
from email.message import EmailMessage

from . import config, parole

log = logging.getLogger("acte_mcp")
SMTP_HOST = os.environ.get("ACTE_MCP_SMTP_HOST", "smtp.gmail.com")


def send(cfg: dict, subject: str, body: str) -> None:
    sender = cfg["conturi"][0]
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = cfg["raport"]["catre"]
    msg["Subject"] = subject
    msg.set_content(body)
    last = None
    for _ in range(2):
        try:
            with smtplib.SMTP_SSL(SMTP_HOST, 465, timeout=60) as s:
                s.login(sender, parole.get(sender))
                s.send_message(msg)
            return
        except Exception as exc:  # noqa: BLE001
            last = exc
    raise RuntimeError(f"Nu am putut trimite e-mailul: {last}")


def daily(cfg: dict, db, day: date) -> tuple[str, str]:
    """Subiectul si corpul raportului pentru o zi (fisierele salvate in ziua respectiva)."""
    start, end = day.isoformat(), (day + timedelta(days=1)).isoformat()
    files = db.conn.execute(
        "SELECT * FROM fisiere WHERE creat_la >= ? AND creat_la < ? ORDER BY client, cale", (start, end)).fetchall()
    saved = [f for f in files if f["stare"] == "salvat"]
    dups = [f for f in files if f["stare"] == "dublura"]
    errors = db.conn.execute("SELECT * FROM mesaje WHERE stare='eroare' ORDER BY primit").fetchall()
    msgs = db.conn.execute(
        "SELECT stare, COUNT(*) n FROM mesaje WHERE actualizat >= ? AND actualizat < ? GROUP BY stare",
        (start, end)).fetchall()
    links = db.conn.execute(
        "SELECT * FROM mesaje WHERE linkuri IS NOT NULL AND actualizat >= ? AND actualizat < ?", (start, end)).fetchall()

    lines = [f"Raport descarcare atasamente — {day:%d.%m.%Y}", ""]
    warn: list[str] = []
    lines.append("CONTROL (mesaje din Gmail din ziua respectiva, comparate cu cele procesate):")
    for cont in cfg["conturi"]:
        c = db.get_meta(f"control:{cont}:{day.isoformat()}")
        if not c:
            lines.append(f"  {cont}: nu s-a putut verifica (nicio rulare reusita dupa miezul noptii)")
            warn.append(f"control lipsa pentru {cont}")
            continue
        n_pend = len(c["neprocesate"])
        lines.append(f"  {cont}: {c['in_gmail']} mesaje in Gmail, {c['in_gmail'] - n_pend} procesate, "
                     f"{n_pend} neprocesate (verificat la {c['verificat_la']})")
        if n_pend:
            warn.append(f"{n_pend} mesaje neprocesate pe {cont}")
    lines.append("")
    stari = {r["stare"]: r["n"] for r in msgs}
    lines.append(f"Mesaje examinate: {sum(stari.values())} (cu atasamente: {stari.get('ok', 0)}, "
                 f"fara: {stari.get('fara_atasamente', 0)}, ignorate: {stari.get('ignorat', 0)}, "
                 f"trimise altora: {stari.get('trimis_altora', 0)})")
    lines.append(f"Fisiere salvate pe server: {len(saved)}; dubluri sarite: {len(dups)}")
    lines.append("")

    by_client = defaultdict(list)
    for f in saved:
        by_client[f["client"] or "_De verificat"].append(f)
    if saved:
        lines.append("FISIERE SALVATE, PE CLIENTI:")
        for client in sorted(by_client, key=lambda k: (k == "_De verificat", k)):
            items = by_client[client]
            lines.append(f"  {client} — {len(items)} fisiere")
            if len(saved) <= 80:
                for f in items:
                    lines.append(f"    - {os.path.basename(f['cale'])} (luna {f['luna']}, {f['cont']})")
        lines.append("")

    unclassified = db.conn.execute(
        "SELECT expeditor, subiect, motiv, fisiere FROM mesaje WHERE stare='ok' AND client IS NULL "
        "AND actualizat >= ? AND actualizat < ?", (start, end)).fetchall()
    if unclassified:
        warn.append(f"{len(unclassified)} mesaje fara firma sigura (in _De verificat)")
        lines.append("DE VERIFICAT (fara firma sigura):")
        for m in unclassified:
            lines.append(f"  - {m['expeditor']} | {m['subiect'][:80]} | {m['motiv']}")
        lines.append("")
    if links:
        warn.append(f"{len(links)} mesaje cu linkuri de descarcat manual")
        lines.append("LINKURI (Drive / WeTransfer etc.) — fisierele NU sunt in mail, se descarca manual:")
        for m in links:
            lines.append(f"  - {m['expeditor']} | {m['subiect'][:80]}")
            for u in json.loads(m["linkuri"])[:5]:
                lines.append(f"      {u}")
        lines.append("")
    if errors:
        warn.append(f"{len(errors)} mesaje cu eroare")
        lines.append("ERORI (se reincearca automat la fiecare rulare):")
        for m in errors:
            lines.append(f"  - {m['cont']} | {m['primit']} | {m['expeditor']} | {m['subiect'] or ''} | "
                         f"{m['incercari']} incercari | {m['eroare']}")
        lines.append("")
    last = db.get_meta("ultima_rulare") or {}
    lines.append(f"Ultima rulare: {last.get('inceput', '-')} — {last.get('rezultat', '-')}")
    lines.append("Daca nu primiti acest raport intr-o zi, descarcarea automata nu ruleaza: verificati "
                 "calculatorul serverului si sarcina 'BBA Descarcare mail'.")

    status = "atentie: " + "; ".join(warn) if warn else f"ok, {len(saved)} fisiere"
    return f"Descarcare acte — {day:%d.%m.%Y} — {status}", "\n".join(lines)


def after_run(cfg: dict, db, result: dict) -> None:
    now = datetime.now()
    rcfg = cfg["raport"]
    if not rcfg.get("catre"):
        return
    # Alerta imediata la eroare, cel mult o data la `alerta_minute`.
    if result.get("erori"):
        last = db.get_meta("ultima_alerta")
        if not last or now - datetime.fromisoformat(last) > timedelta(minutes=int(rcfg["alerta_minute"])):
            body = ("Descarcarea automata a atasamentelor a intampinat erori:\n\n"
                    + "\n".join(f"- {e}" for e in result["erori"])
                    + "\n\nSe reincearca automat la fiecare rulare. Mesajele neprocesate NU se pierd: "
                      "vor fi descarcate la prima rulare reusita.")
            try:
                send(cfg, f"ALERTA descarcare acte — {now:%d.%m.%Y %H:%M}", body)
                db.set_meta("ultima_alerta", now.isoformat(timespec="seconds"))
            except Exception:  # noqa: BLE001
                log.exception("Alerta")
    # Raportul zilnic pe ziua de ieri, o singura data, dupa ora configurata.
    today = date.today().isoformat()
    if now.hour >= int(rcfg["ora"]) and db.get_meta("raport_trimis") != today:
        subject, body = daily(cfg, db, date.today() - timedelta(days=1))
        send(cfg, subject, body)
        db.set_meta("raport_trimis", today)
