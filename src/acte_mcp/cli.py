"""Linia de comanda.

  acte-mcp                      porneste serverul MCP (pentru Claude)
  acte-mcp configurare          creeaza / modifica ~/.acte-mcp/config.json
  acte-mcp parola <email>       salveaza parola de aplicatie Gmail (in Windows Credential Manager)
  acte-mcp test                 verifica serverul, Excelul, Gmail (IMAP) si trimiterea de e-mail
  acte-mcp descarca             o rulare a descarcarii (o face si sarcina programata)
  acte-mcp raport [--trimite]   arata (sau trimite) raportul pe ziua de ieri
  acte-mcp instaleaza-sarcina   creeaza sarcina Windows care descarca la cateva minute
"""

from __future__ import annotations

import argparse
import getpass
import json
import logging
import subprocess
import sys
from datetime import date, timedelta
from logging.handlers import RotatingFileHandler
from pathlib import Path

from . import __version__, config

TASK_NAME = "BBA Descarcare mail"


def setup_logging(console: bool = True) -> None:
    config.HOME.mkdir(parents=True, exist_ok=True)
    handlers: list[logging.Handler] = [
        RotatingFileHandler(config.HOME / "descarcare.log", maxBytes=5_000_000, backupCount=5, encoding="utf-8")]
    if console:
        handlers.append(logging.StreamHandler())
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", handlers=handlers)


def _ask(prompt: str, default: str) -> str:
    val = input(f"{prompt} [{default}]: ").strip()
    return val or default


def cmd_configurare(args) -> int:
    cfg = config.load()
    print(f"Configurare in {config.CONFIG_PATH} (Enter = pastreaza valoarea dintre paranteze)\n")
    cfg["radacina"] = _ask("Radacina serverului (recomandat \\\\NAS\\share, sau Z:\\)", cfg["radacina"])
    cfg["conturi"] = [c.strip() for c in _ask("Conturile Gmail, separate prin virgula", ", ".join(cfg["conturi"])).split(",")]
    cfg["excel_clienti"] = _ask("Calea catre Clienti_Contabilitate.xlsx (gol = fara)", cfg["excel_clienti"] or "")
    cfg["de_la"] = _ask("Descarca mesajele primite incepand cu (AAAA-LL-ZZ)", cfg["de_la"])
    cfg["raport"]["catre"] = _ask("Raportul zilnic si alertele merg la", cfg["raport"]["catre"])
    user = {k: cfg[k] for k in ("radacina", "conturi", "excel_clienti", "de_la", "raport", "adrese_proprii",
                                "mapare_manuala", "ignora_expeditori", "luna_minus", "folder_client",
                                "folder_neclasificat", "folder_tehnic")}
    print(f"\nSalvat: {config.save(user)}")
    print("Urmeaza: acte-mcp parola <email> pentru fiecare cont, apoi acte-mcp test")
    return 0


def cmd_parola(args) -> int:
    from . import parole

    print("Parola de aplicatie se genereaza din contul Google: Gestionati contul > Securitate >\n"
          "Verificare in doi pasi > Parole pentru aplicatii. Are 16 litere.")
    value = getpass.getpass(f"Parola de aplicatie pentru {args.email}: ").strip()
    if len(value.replace(" ", "")) < 16:
        print("Parola pare prea scurta (are 16 litere). Nu am salvat nimic.")
        return 1
    parole.save(args.email, value)
    print("Salvata in Windows Credential Manager.")
    return 0


def cmd_test(args) -> int:
    from . import parole, sync
    from .clienti import Registry

    cfg = config.load()
    ok = True
    root = config.root(cfg)
    print(f"Server: {root}")
    try:
        sync.check_root(cfg)
        print("  accesibil si se poate scrie.")
    except Exception as exc:  # noqa: BLE001
        print(f"  EROARE: {exc}")
        ok = False
    try:
        reg = Registry.load(cfg, root)
        mapped = [c for c in reg.clients if c.folder and (c.emailuri or c.domenii)]
        print(f"Clienti: {len(reg.folders)} foldere pe server, {len(mapped)} cu adrese de e-mail din Excel.")
        if reg.unmapped:
            print(f"  Din Excel fara folder pe server ({len(reg.unmapped)}): {', '.join(reg.unmapped[:20])}")
    except Exception as exc:  # noqa: BLE001
        print(f"Clienti: EROARE la citirea Excelului: {exc}")
        ok = False
    for cont in cfg["conturi"]:
        print(f"Gmail {cont}:")
        if not parole.has(cont):
            print(f"  lipseste parola: acte-mcp parola {cont}")
            ok = False
            continue
        try:
            imap = sync.connect(cont)
            folder = sync.all_mail_folder(imap)
            imap.select(folder, readonly=True)
            n = len(sync.search_since(imap, date.today() - timedelta(days=7)))
            imap.logout()
            print(f"  IMAP ok, dosar {folder}, {n} mesaje in ultimele 7 zile.")
        except Exception as exc:  # noqa: BLE001
            print(f"  EROARE IMAP: {exc}")
            ok = False
    sender = cfg["conturi"][0] if cfg["conturi"] else None
    if sender and parole.has(sender):
        import smtplib

        from .raport import SMTP_HOST

        try:
            with smtplib.SMTP_SSL(SMTP_HOST, 465, timeout=30) as s:
                s.login(sender, parole.get(sender))
            print(f"Trimitere e-mail (SMTP) de pe {sender}: ok.")
        except Exception as exc:  # noqa: BLE001
            print(f"Trimitere e-mail: EROARE {exc}")
            ok = False
    print("\nTotul e in regula." if ok else "\nSunt probleme de rezolvat (vezi mai sus).")
    return 0 if ok else 1


def cmd_descarca(args) -> int:
    from . import sync

    setup_logging(console=True)
    result = sync.run(conturi=[args.cont] if args.cont else None)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("rezultat") in ("ok", "sarit") else 1


def cmd_raport(args) -> int:
    from . import raport, sync

    cfg = config.load()
    day = date.fromisoformat(args.zi) if args.zi else date.today() - timedelta(days=1)
    subject, body = raport.daily(cfg, sync.Db(), day)
    print(subject + "\n\n" + body)
    if args.trimite:
        raport.send(cfg, subject, body)
        print(f"\nTrimis la {cfg['raport']['catre']}.")
    return 0


def cmd_sarcina(args) -> int:
    exe = Path(sys.executable).parent / "acte-sync.exe"
    if not exe.exists():
        # instalat cu uv tool: scripturile sunt in ~/.local/bin
        exe = Path.home() / ".local" / "bin" / "acte-sync.exe"
    if not exe.exists():
        print(f"Nu gasesc acte-sync.exe (am cautat langa {sys.executable} si in ~/.local/bin).")
        return 1
    cmd = ["schtasks", "/Create", "/F", "/TN", TASK_NAME, "/SC", "MINUTE", "/MO", str(args.minute),
           "/TR", f'"{exe}"']
    print("Rulez: " + " ".join(cmd))
    r = subprocess.run(cmd, capture_output=True, text=True)
    print(r.stdout or r.stderr)
    if r.returncode == 0:
        print(f"Sarcina '{TASK_NAME}' ruleaza la {args.minute} minute cat timp utilizatorul e conectat.\n"
              "Ca sa ruleze si fara utilizator conectat: Task Scheduler > sarcina > Properties > "
              "'Run whether user is logged on or not' (cere parola contului Windows) si folositi in "
              "configurare calea UNC a serverului in loc de Z:.")
    return r.returncode


def main_sync() -> int:
    """Punctul de intrare fara fereastra (acte-sync.exe), pentru sarcina programata."""
    from . import sync

    setup_logging(console=False)
    try:
        result = sync.run()
    except Exception:  # noqa: BLE001
        logging.getLogger("acte_mcp").exception("Rulare esuata")
        return 1
    return 0 if result.get("rezultat") in ("ok", "sarit") else 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="acte-mcp", description="Acte din Gmail pe serverul local + server MCP")
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="cmd")
    sub.add_parser("configurare", help="creeaza / modifica configurarea").set_defaults(func=cmd_configurare)
    p = sub.add_parser("parola", help="salveaza parola de aplicatie Gmail")
    p.add_argument("email")
    p.set_defaults(func=cmd_parola)
    sub.add_parser("test", help="verifica serverul, Excelul, Gmail si e-mailul").set_defaults(func=cmd_test)
    p = sub.add_parser("descarca", help="o rulare a descarcarii")
    p.add_argument("--cont", help="doar acest cont")
    p.set_defaults(func=cmd_descarca)
    p = sub.add_parser("raport", help="raportul pe o zi (implicit ieri)")
    p.add_argument("--zi", help="AAAA-LL-ZZ")
    p.add_argument("--trimite", action="store_true", help="il si trimite pe e-mail")
    p.set_defaults(func=cmd_raport)
    p = sub.add_parser("instaleaza-sarcina", help="sarcina Windows care descarca periodic")
    p.add_argument("--minute", type=int, default=5)
    p.set_defaults(func=cmd_sarcina)
    sub.add_parser("serve", help="porneste serverul MCP (implicit)")

    args = parser.parse_args(argv)
    if getattr(args, "func", None):
        return args.func(args)
    from .server import run

    run()
    return 0


if __name__ == "__main__":
    sys.exit(main())
