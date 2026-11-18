"""Uneltele MCP: actele clientilor de pe serverul local (Z:) si descarcarea din Gmail.

Merge pe orice calculator care vede serverul. Informatiile despre originea fiecarui
fisier (expeditor, subiect, data) si marcajele "procesat" stau in folderul tehnic de pe
server (_acte-mcp), deci se vad de pe toate calculatoarele.
"""

from __future__ import annotations

import getpass
import json
import os
import shutil
import socket
from datetime import date, datetime
from pathlib import Path
from typing import Literal

from mcp.server.fastmcp import FastMCP

from . import config, extract
from .clienti import Registry

mcp = FastMCP(
    "acte",
    instructions=(
        "Acces la actele clientilor cabinetului de contabilitate, de pe serverul local. Fiecare client "
        "are un folder '<numar>. <Nume firma>' cu actele in 'Acte contabile/<AAAA>/<LL>', unde luna e "
        "LUNA DE LUCRU (luna in care s-a primit mailul minus 1). Atasamentele din Gmail ajung aici "
        "automat; ce nu are firma sigura sta in '_De verificat/<AAAA>/<LL>'. Flux tipic: acte_clienti "
        "-> acte_lista(client, luna) -> acte_citeste(cale) -> pregatesti importul cu uneltele SAGA -> "
        "dupa import, acte_marcheaza(cale, 'importat_saga'). Actele deja marcate importat_saga nu se "
        "mai importa. acte_muta modifica serverul: doar cu acordul utilizatorului."
    ),
)

_journal_cache: dict = {"key": None, "files": {}, "links": []}


def _cfg() -> dict:
    return config.load()


def _root() -> Path:
    root = config.root(_cfg())
    if not root.exists():
        raise RuntimeError(f"Serverul nu e accesibil de pe acest calculator: {root}. "
                           "Verifica unitatea Z: sau 'radacina' din ~/.acte-mcp/config.json.")
    return root


def _registry() -> Registry:
    cfg = _cfg()
    try:
        return Registry.load(cfg, _root())
    except Exception:  # noqa: BLE001 - Excel lipsa / blocat: mergem doar pe foldere
        return Registry([], [p for p in _folders()], cfg.get("mapare_manuala"))


def _folders() -> list[str]:
    from .clienti import list_folders

    return list_folders(_root())


def _inside(path: Path) -> Path:
    root = _root().resolve()
    p = path if path.is_absolute() else root / path
    p = p.resolve()
    if p != root and root not in p.parents:
        raise ValueError(f"Calea e in afara serverului: {path}")
    return p


def _rel(path: Path) -> str:
    return os.path.relpath(path, _root())


def _key(rel: str) -> str:
    return rel.replace("/", "\\").lower()


def _journal() -> tuple[dict, list]:
    """Originea fisierelor, din jurnalele lunare scrise la descarcare."""
    folder = config.tech_dir(_cfg()) / "jurnal"
    files = sorted(folder.glob("*.jsonl")) if folder.exists() else []
    key = tuple((f.name, f.stat().st_mtime, f.stat().st_size) for f in files)
    if key != _journal_cache["key"]:
        by_path, links = {}, []
        for f in files:
            for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
                try:
                    e = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if e.get("tip") == "fisier" and e.get("stare") == "salvat":
                    by_path[_key(e["cale"])] = e
                elif e.get("tip") == "link":
                    links.append(e)
                elif e.get("tip") == "mutat":
                    old = by_path.pop(_key(e["din"]), None)
                    if old:
                        by_path[_key(e["in"])] = {**old, "cale": e["in"], "client": e.get("client")}
        _journal_cache.update(key=key, files=by_path, links=links)
    return _journal_cache["files"], _journal_cache["links"]


def _marks() -> dict:
    path = config.tech_dir(_cfg()) / "procesate.jsonl"
    out: dict = {}
    if path.exists():
        for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
            try:
                e = json.loads(line)
            except json.JSONDecodeError:
                continue
            out.setdefault(_key(e["cale"]), []).append(e)
    return out


def _default_month() -> str:
    y, m = config.work_month(date.today(), int(_cfg()["luna_minus"]))
    return f"{y:04d}-{m:02d}"


def _month(luna: str | None) -> tuple[int, int, str]:
    luna = luna or _default_month()
    y, m = luna.split("-")
    return int(y), int(m), f"{int(y):04d}-{int(m):02d}"


def _one_folder(client: str) -> str:
    hits = _registry().resolve(client)
    if not hits:
        raise ValueError(f"Nu gasesc niciun folder de client pentru '{client}'. Foloseste acte_clienti.")
    if len(hits) > 1:
        raise ValueError(f"'{client}' se potriveste cu mai multe foldere: {hits}. Fii mai precis.")
    return hits[0]


def _describe(path: Path, journal: dict, marks: dict) -> dict:
    rel = _rel(path)
    st = path.stat()
    d = {"cale": rel, "nume": path.name, "marime": st.st_size,
         "modificat": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")}
    j = journal.get(_key(rel))
    if j:
        d["origine"] = {k: j.get(k) for k in ("cont", "primit", "expeditor", "expeditori_originali", "subiect",
                                              "nume_original", "forward", "motiv") if j.get(k)}
    m = marks.get(_key(rel))
    if m:
        d["marcaje"] = [{k: e.get(k) for k in ("stare", "nota", "la", "de")} for e in m]
        imported = False
        for e in m:  # in ordine; "anulat" retrage marcajul anterior
            if e.get("stare") == "importat_saga":
                imported = True
            elif e.get("stare") == "anulat":
                imported = False
        d["importat_saga"] = imported
    return d


# ---------------------------------------------------------------- unelte

@mcp.tool()
async def acte_status() -> dict:
    """Starea: serverul e accesibil, cand a rulat ultima descarcare din Gmail si cu ce rezultat."""
    cfg = _cfg()
    out: dict = {"radacina": cfg["radacina"], "luna_de_lucru_implicita": _default_month(),
                 "config": str(config.CONFIG_PATH)}
    try:
        out["foldere_clienti"] = len(_folders())
        out["server_accesibil"] = True
    except RuntimeError as exc:
        out["server_accesibil"] = False
        out["mesaj"] = str(exc)
    db_path = config.HOME / "stare.sqlite"
    if db_path.exists():
        from .sync import Db

        db = Db(db_path)
        out["descarcare_pe_acest_calculator"] = True
        out["ultima_rulare"] = db.get_meta("ultima_rulare")
        out["mesaje_cu_eroare"] = db.conn.execute("SELECT COUNT(*) FROM mesaje WHERE stare='eroare'").fetchone()[0]
        out["ultima_reusita_pe_cont"] = {c: db.get_meta(f"ultima_ok:{c}") for c in cfg["conturi"]}
    else:
        out["descarcare_pe_acest_calculator"] = False
        files = sorted((config.tech_dir(cfg) / "jurnal").glob("*.jsonl")) if out.get("server_accesibil") else []
        if files:
            out["jurnal_actualizat"] = datetime.fromtimestamp(files[-1].stat().st_mtime).strftime("%Y-%m-%d %H:%M")
    return out


@mcp.tool()
async def acte_clienti(cauta: str | None = None) -> dict:
    """Clientii de pe server: folderul, CUI-ul si adresele de e-mail (din Excelul de clienti).

    Args:
        cauta: filtru dupa nume, numarul folderului sau CUI (optional).
    """
    reg = _registry()
    clients = [c for c in reg.clients if c.folder]
    if cauta:
        wanted = set(reg.resolve(cauta))
        clients = [c for c in clients if c.folder in wanted]
    return {"numar": len(clients), "clienti": [c.public() for c in sorted(clients, key=lambda c: c.folder)],
            "din_excel_fara_folder": reg.unmapped if not cauta else None}


@mcp.tool()
async def acte_lista(client: str, luna: str | None = None, doar_neimportate: bool = False) -> dict:
    """Actele unui client pentru o luna de lucru, cu originea lor (mail, expeditor, subiect)
    si marcajele (ex. importat_saga).

    Args:
        client: numele firmei, numarul folderului (ex. "12"), CUI-ul sau numele folderului.
            "_De verificat" pentru fisierele fara firma sigura.
        luna: luna de lucru "AAAA-LL" (implicit luna trecuta, adica luna de lucru curenta).
        doar_neimportate: true = doar actele care nu sunt marcate importat_saga.
    """
    cfg = _cfg()
    y, m, luna = _month(luna)
    if client.strip().lower().startswith("_de verificat"):
        folder = _root() / config.month_path(cfg["folder_neclasificat"], y, m)
        name = "_De verificat"
    else:
        name = _one_folder(client)
        folder = _root() / name / config.month_path(cfg["folder_client"], y, m)
    if not folder.exists():
        return {"client": name, "luna": luna, "folder": _rel(folder), "numar": 0, "fisiere": [],
                "mesaj": "Folderul lunii nu exista (niciun act inca)."}
    journal, _ = _journal()
    marks = _marks()
    items = [_describe(p, journal, marks) for p in sorted(folder.rglob("*"))
             if p.is_file() and not p.name.endswith(".partial") and not p.name.startswith("~$")]
    if doar_neimportate:
        items = [i for i in items if not i.get("importat_saga")]
    c = _registry().by_folder(name)
    return {"client": name, "cui": c.cui if c else "", "luna": luna, "folder": _rel(folder),
            "numar": len(items), "fisiere": items}


@mcp.tool()
async def acte_citeste(cale: str, pagini: str | None = None, ca_imagine: bool = False) -> list:
    """Continutul unui act: text (PDF cu text, Excel, CSV, XML e-Factura, DOCX) sau imagini
    (poze, PDF scanat) pe care le poti citi direct.

    Args:
        cale: calea din acte_lista (relativa la server) sau calea completa.
        pagini: la PDF, ex. "1-3" sau "1,4" (implicit primele pagini).
        ca_imagine: true = trimite paginile PDF ca imagini chiar daca au text (ex. bonuri, tabele).
    """
    path = _inside(Path(cale))
    if not path.is_file():
        raise FileNotFoundError(f"Nu exista fisierul: {cale}")
    return extract.read(path, pagini, ca_imagine)


@mcp.tool()
async def acte_cauta(text: str, luna: str | None = None, client: str | None = None, limita: int = 100) -> dict:
    """Cauta acte dupa numele fisierului, expeditor sau subiectul mailului.

    Args:
        text: ce cauti (ex. "extras", "factura 123", "@furnizor.ro").
        luna: limiteaza la o luna de lucru "AAAA-LL" (optional).
        client: limiteaza la un client (optional).
    """
    cfg = _cfg()
    needle = text.lower().strip()
    journal, _ = _journal()
    marks = _marks()
    root = _root()
    folders = [_one_folder(client)] if client else _folders() + [cfg["folder_neclasificat"].split("\\")[0]]
    hits = []
    for name in folders:
        base = root / name
        if luna:
            y, m, _ = _month(luna)
            tmpl = cfg["folder_neclasificat"] if name.startswith("_") else cfg["folder_client"]
            base = root / config.month_path(tmpl, y, m) if name.startswith("_") else base / config.month_path(tmpl, y, m)
        if not base.exists():
            continue
        for p in base.rglob("*"):
            if not p.is_file():
                continue
            j = journal.get(_key(_rel(p)), {})
            hay = " ".join([p.name, j.get("expeditor", ""), j.get("subiect", ""), " ".join(j.get("expeditori_originali") or [])]).lower()
            if needle in hay:
                hits.append(_describe(p, journal, marks))
                if len(hits) >= limita:
                    return {"numar": len(hits), "trunchiat": True, "fisiere": hits}
    return {"numar": len(hits), "fisiere": hits}


@mcp.tool()
async def acte_linkuri(luna_primire: str | None = None) -> dict:
    """Mesajele cu linkuri Drive/WeTransfer/Dropbox (fisiere care NU sunt atasate si trebuie
    descarcate manual).

    Args:
        luna_primire: "AAAA-LL" a primirii mailului (optional).
    """
    _, links = _journal()
    if luna_primire:
        links = [l for l in links if (l.get("primit") or "").startswith(luna_primire)]
    return {"numar": len(links), "mesaje": links[-200:]}


@mcp.tool()
async def acte_marcheaza(cale: str, stare: Literal["importat_saga", "verificat", "ignorat", "anulat"],
                         nota: str | None = None) -> dict:
    """Marcheaza un act (ex. dupa importul in SAGA), ca sa nu fie prelucrat de doua ori.
    Marcajul se vede de pe toate calculatoarele. "anulat" retrage un marcaj gresit.

    Args:
        cale: calea actului (din acte_lista).
        stare: importat_saga | verificat | ignorat | anulat.
        nota: ex. "import F_... in SAGA, factura 123".
    """
    path = _inside(Path(cale))
    if not path.is_file():
        raise FileNotFoundError(f"Nu exista fisierul: {cale}")
    entry = {"cale": _rel(path), "stare": stare, "nota": nota or "", "la": datetime.now().strftime("%Y-%m-%d %H:%M"),
             "de": f"{getpass.getuser()}@{socket.gethostname()}"}
    target = config.tech_dir(_cfg()) / "procesate.jsonl"
    target.parent.mkdir(parents=True, exist_ok=True)
    with open(target, "a", encoding="utf-8") as fh:
        fh.write(json.dumps(entry, ensure_ascii=False) + "\n")
    return {"marcat": entry}


@mcp.tool()
async def acte_muta(cale: str, client: str, luna: str | None = None, confirma: bool = False) -> dict:
    """Muta un act in folderul altui client / altei luni (ex. din _De verificat). Modifica
    serverul: apeleaza intai fara confirma, arata utilizatorului, apoi cu confirma=true.

    Args:
        cale: calea actului.
        client: clientul destinatie (nume, numar, CUI).
        luna: luna de lucru destinatie "AAAA-LL" (implicit luna din calea actuala, altfel luna de lucru).
        confirma: true = muta efectiv.
    """
    cfg = _cfg()
    src = _inside(Path(cale))
    if not src.is_file():
        raise FileNotFoundError(f"Nu exista fisierul: {cale}")
    name = _one_folder(client)
    if not luna:
        parts = src.parent.parts
        luna = f"{parts[-2]}-{parts[-1]}" if len(parts) >= 2 and parts[-2].isdigit() and parts[-1].isdigit() else None
    y, m, luna = _month(luna)
    dest_dir = _root() / name / config.month_path(cfg["folder_client"], y, m)
    dest = dest_dir / src.name
    n = 2
    while dest.exists():
        dest = dest_dir / f"{src.stem}_{n}{src.suffix}"
        n += 1
    plan = {"din": _rel(src), "in": _rel(dest), "client": name, "luna": luna}
    if not confirma:
        return {"plan": plan, "pas_urmator": "Arata utilizatorului mutarea; apoi acte_muta(..., confirma=true)."}
    dest_dir.mkdir(parents=True, exist_ok=True)
    shutil.move(str(src), str(dest))
    from .sync import append_journal

    append_journal(cfg, {"tip": "mutat", **plan, "la": datetime.now().strftime("%Y-%m-%d %H:%M"),
                         "de": f"{getpass.getuser()}@{socket.gethostname()}"})
    return {"mutat": plan}


@mcp.tool()
async def acte_descarca_acum() -> dict:
    """Porneste acum descarcarea din Gmail (altfel ruleaza automat la cateva minute).
    Merge doar pe calculatorul unde e configurata descarcarea."""
    import asyncio

    if not (config.HOME / "stare.sqlite").exists() and not config.CONFIG_PATH.exists():
        return {"rezultat": "indisponibil",
                "motiv": "Descarcarea nu e configurata pe acest calculator; ruleaza pe calculatorul serverului."}
    from . import sync

    return await asyncio.to_thread(sync.run)


def run() -> None:
    mcp.run()
