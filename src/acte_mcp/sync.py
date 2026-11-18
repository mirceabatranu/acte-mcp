"""Descarcarea atasamentelor din Gmail pe server, prin IMAP.

Garantii:
- Casuta se deschide DOAR IN CITIRE (EXAMINE) si mesajele se citesc cu BODY.PEEK:
  nimic nu se marcheaza ca citit, nu se muta, nu se sterge.
- Se cauta in "Toate mesajele" (All Mail), deci si mesajele arhivate sau cu etichete.
- Fiecare rulare reia ultimele 3 zile in plus fata de ultima rulare reusita; mesajele deja
  procesate (dupa X-GM-MSGID, un id stabil Gmail) se sar. Nu raman goluri, nu apar dubluri.
- Un mesaj se marcheaza "procesat" abia dupa ce toate fisierele lui sunt scrise pe disc.
  Daca ceva pica, se reia la urmatoarea rulare.
- Fisierele se scriu prin redenumire atomica (.partial -> nume final) si nu se suprascriu.
- Acelasi fisier (acelasi continut) in acelasi folder nu se scrie de doua ori.
"""

from __future__ import annotations

import hashlib
import imaplib
import json
import logging
import os
import re
import socket
import sqlite3
import time
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from email.utils import parsedate_to_datetime
from pathlib import Path

from . import config, parole
from .clienti import Client, Registry
from .mailparse import Parsed, analyze

log = logging.getLogger("acte_mcp")

IMAP_HOST = os.environ.get("ACTE_MCP_IMAP_HOST", "imap.gmail.com")
OVERLAP_DAYS = 3
FINAL = ("ok", "fara_atasamente", "trimis_altora", "ignorat", "ciorna")
BAD_CHARS = re.compile(r'[/\\:*?"<>|\x00-\x1f]')


# ---------------------------------------------------------------- stare locala

class Db:
    def __init__(self, path: Path | None = None):
        path = path or (config.HOME / "stare.sqlite")
        path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(path, timeout=30)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS mesaje (
                cont TEXT, msgid TEXT, primit TEXT, expeditor TEXT, subiect TEXT,
                stare TEXT, incercari INTEGER DEFAULT 0, eroare TEXT, client TEXT,
                motiv TEXT, fisiere INTEGER DEFAULT 0, linkuri TEXT, actualizat TEXT,
                PRIMARY KEY (cont, msgid));
            CREATE TABLE IF NOT EXISTS fisiere (
                id INTEGER PRIMARY KEY AUTOINCREMENT, cont TEXT, msgid TEXT, nume_original TEXT,
                cale TEXT, sha256 TEXT, marime INTEGER, client TEXT, luna TEXT, stare TEXT,
                creat_la TEXT);
            CREATE INDEX IF NOT EXISTS fisiere_sha ON fisiere (sha256);
            CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT);
            """
        )
        self.conn.commit()

    def get_meta(self, k: str, default=None):
        row = self.conn.execute("SELECT v FROM meta WHERE k=?", (k,)).fetchone()
        return json.loads(row["v"]) if row else default

    def set_meta(self, k: str, v) -> None:
        self.conn.execute("INSERT OR REPLACE INTO meta (k, v) VALUES (?, ?)", (k, json.dumps(v, ensure_ascii=False)))
        self.conn.commit()

    def done(self, cont: str, msgid: str) -> bool:
        row = self.conn.execute("SELECT stare FROM mesaje WHERE cont=? AND msgid=?", (cont, msgid)).fetchone()
        return bool(row and row["stare"] in FINAL)

    def record(self, cont: str, msgid: str, **fields) -> None:
        fields["actualizat"] = _now()
        row = self.conn.execute("SELECT 1 FROM mesaje WHERE cont=? AND msgid=?", (cont, msgid)).fetchone()
        if row:
            sets = ", ".join(f"{k}=?" for k in fields)
            self.conn.execute(f"UPDATE mesaje SET {sets} WHERE cont=? AND msgid=?", (*fields.values(), cont, msgid))
        else:
            cols = ", ".join(["cont", "msgid", *fields])
            self.conn.execute(f"INSERT INTO mesaje ({cols}) VALUES ({', '.join('?' * (len(fields) + 2))})",
                              (cont, msgid, *fields.values()))
        self.conn.commit()

    def fail(self, cont: str, msgid: str, err: str, **fields) -> None:
        row = self.conn.execute("SELECT incercari FROM mesaje WHERE cont=? AND msgid=?", (cont, msgid)).fetchone()
        n = (row["incercari"] if row else 0) + 1
        self.record(cont, msgid, stare="eroare", incercari=n, eroare=err[:500], **fields)

    def sha_in_dir(self, sha: str, folder: str) -> str | None:
        for row in self.conn.execute("SELECT cale FROM fisiere WHERE sha256=? AND stare='salvat'", (sha,)):
            if str(Path(row["cale"]).parent) == folder and Path(row["cale"]).exists():
                return row["cale"]
        return None


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


@contextmanager
def run_lock(path: Path):
    """Un singur proces de descarcare odata (sarcina programata vs. comanda din Claude)."""
    import msvcrt  # noqa: PLC0415 - doar pe Windows

    path.parent.mkdir(parents=True, exist_ok=True)
    fh = open(path, "a+")
    try:
        try:
            msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            yield False
            return
        try:
            yield True
        finally:
            fh.seek(0)
            msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
    finally:
        fh.close()


# ---------------------------------------------------------------- IMAP

def connect(cont: str, password: str | None = None) -> imaplib.IMAP4_SSL:
    imap = imaplib.IMAP4_SSL(IMAP_HOST, 993, timeout=120)
    imap.login(cont, password or parole.get(cont))
    return imap


def all_mail_folder(imap: imaplib.IMAP4_SSL) -> str:
    """Numele (localizat) al dosarului "Toate mesajele", dupa atributul \\All."""
    typ, data = imap.list()
    if typ != "OK":
        raise RuntimeError("Nu pot lista dosarele IMAP")
    for line in data:
        if isinstance(line, tuple):
            line = b" ".join(x for x in line if isinstance(x, bytes))
        text = line.decode("utf-8", "replace")
        m = re.match(r'\((?P<flags>[^)]*)\)\s+(?:"[^"]*"|NIL)\s+(?P<name>.+)$', text)
        if m and "\\All" in m.group("flags"):
            return m.group("name").strip()
    raise RuntimeError(
        "Nu gasesc dosarul 'Toate mesajele' prin IMAP. In Gmail: Setari > Redirectionare si POP/IMAP > "
        "activeaza IMAP; la Etichete, bifeaza 'Afiseaza in IMAP' pentru 'Toate mesajele'."
    )


def _imap_date(d: date) -> str:
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    return f"{d.day:02d}-{months[d.month - 1]}-{d.year}"


def search_since(imap, since: date, before: date | None = None) -> list[bytes]:
    crit = ["SINCE", _imap_date(since)]
    if before:
        crit += ["BEFORE", _imap_date(before)]
    typ, data = imap.uid("SEARCH", *crit)
    if typ != "OK":
        raise RuntimeError(f"Cautare IMAP esuata: {data}")
    return data[0].split() if data and data[0] else []


@dataclass
class Meta:
    uid: bytes
    msgid: str
    size: int
    labels: str
    internal: datetime | None


_META_RE = {
    "uid": re.compile(rb"\bUID (\d+)"),
    "msgid": re.compile(rb"X-GM-MSGID (\d+)"),
    "size": re.compile(rb"RFC822\.SIZE (\d+)"),
    "labels": re.compile(rb"X-GM-LABELS \((.*?)\)(?= [A-Z]|\)$|$)"),
    "internal": re.compile(rb'INTERNALDATE "([^"]+)"'),
}


def fetch_meta(imap, uids: list[bytes]) -> list[Meta]:
    out: list[Meta] = []
    for i in range(0, len(uids), 200):
        chunk = b",".join(uids[i:i + 200]).decode()
        typ, data = imap.uid("FETCH", chunk, "(X-GM-MSGID X-GM-LABELS INTERNALDATE RFC822.SIZE)")
        if typ != "OK":
            raise RuntimeError(f"FETCH esuat: {data}")
        for item in data:
            if item is None or item == b")":
                continue
            line = b" ".join(x for x in item if isinstance(x, bytes)) if isinstance(item, tuple) else item
            m = {k: r.search(line) for k, r in _META_RE.items()}
            if not m["uid"] or not m["msgid"]:
                continue
            internal = None
            if m["internal"]:
                try:
                    internal = datetime.strptime(m["internal"].group(1).decode(), "%d-%b-%Y %H:%M:%S %z")
                except ValueError:
                    internal = None
            out.append(Meta(uid=m["uid"].group(1), msgid=m["msgid"].group(1).decode(),
                            size=int(m["size"].group(1)) if m["size"] else 0,
                            labels=m["labels"].group(1).decode("utf-8", "replace") if m["labels"] else "",
                            internal=internal))
    return out


def fetch_raw(imap, uid: bytes) -> bytes:
    typ, data = imap.uid("FETCH", uid.decode(), "(BODY.PEEK[])")
    if typ != "OK":
        raise RuntimeError(f"Nu pot descarca mesajul: {data}")
    for item in data:
        if isinstance(item, tuple) and len(item) >= 2 and isinstance(item[1], bytes):
            return item[1]
    raise RuntimeError("Mesajul nu a fost intors de server (posibil sters intre timp)")


# ---------------------------------------------------------------- clasificare

def classify(p: Parsed, reg: Registry, own: set[str]) -> tuple[Client | None, str]:
    """Firma mesajului: dupa expeditor (la forward: expeditorul original), apoi dupa text."""
    own_sender = p.expeditor in own
    senders = p.expeditori_originali if (p.forward or own_sender) else [p.expeditor]
    names = " ".join(a.name for a in p.atasamente)
    variants: list[Client] = []
    for addr in senders:
        if addr in own:
            continue
        cands = reg.by_address(addr)
        if len(cands) > 1:
            narrowed = reg.find_in_text(f"{p.subiect}\n{p.corp}\n{names}", cands)
            cands = narrowed if len(narrowed) == 1 else cands
        if len(cands) == 1:
            c = cands[0]
            if not c.folder:
                return None, f"clientul {c.nume} nu are folder pe server"
            return c, f"dupa adresa {addr}"
        variants += [c for c in cands if c not in variants]
    if variants:
        return None, "variante: " + ", ".join(c.nume for c in variants)
    if p.forward or own_sender:
        hits = [c for c in reg.find_in_text(f"{p.subiect}\n{p.nota_forward}") if c.folder]
        if len(hits) == 1:
            return hits[0], "dupa numele/CUI din subiect sau nota de forward"
        if len(hits) > 1:
            return None, "variante: " + ", ".join(c.nume for c in hits)
        return None, "forward fara firma identificabila"
    return None, f"expeditor necunoscut ({p.expeditor})"


# ---------------------------------------------------------------- scriere pe disc

def clean_name(name: str, limit: int = 120) -> str:
    name = BAD_CHARS.sub("-", name).strip(" .")
    if len(name) > limit:
        stem, dot, ext = name.rpartition(".")
        if dot and len(ext) <= 8:
            name = stem[: limit - len(ext) - 1].rstrip(" .") + "." + ext
        else:
            name = name[:limit]
    return name or "fisier"


def write_unique(folder: Path, name: str, data: bytes, sha: str) -> tuple[Path, bool]:
    """Scrie fisierul fara sa suprascrie. Intoarce (cale, scris_acum).
    Daca exista deja un fisier cu acelasi nume si acelasi continut, nu-l mai scrie."""
    folder.mkdir(parents=True, exist_ok=True)
    stem, dot, ext = name.rpartition(".")
    if not dot:
        stem, ext = name, ""
    n = 1
    while True:
        cand = folder / (name if n == 1 else f"{stem}_{n}{'.' + ext if ext else ''}")
        if not cand.exists():
            break
        if cand.stat().st_size == len(data) and hashlib.sha256(cand.read_bytes()).hexdigest() == sha:
            return cand, False
        n += 1
    tmp = cand.with_name(cand.name + ".partial")
    with open(tmp, "wb") as fh:
        fh.write(data)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp, cand)
    if cand.stat().st_size != len(data):
        raise OSError(f"Marime gresita dupa scriere: {cand}")
    return cand, True


def append_journal(cfg: dict, entry: dict) -> None:
    folder = config.tech_dir(cfg) / "jurnal"
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / f"{datetime.now():%Y-%m}.jsonl"
    with open(path, "a", encoding="utf-8") as fh:
        fh.write(json.dumps(entry, ensure_ascii=False) + "\n")


def check_root(cfg: dict) -> None:
    root = config.root(cfg)
    if not root.exists():
        raise RuntimeError(f"Serverul nu e accesibil: {root}")
    probe = config.tech_dir(cfg) / ".test-scriere"
    probe.parent.mkdir(parents=True, exist_ok=True)
    probe.write_text(_now(), encoding="utf-8")
    probe.unlink()


# ---------------------------------------------------------------- procesare

def process_message(cfg: dict, cont: str, meta: Meta, raw: bytes, reg: Registry, db: Db) -> dict:
    own = set(cfg["adrese_proprii"])
    p = analyze(raw, own)
    received = meta.internal.astimezone() if meta.internal else None
    if received is None:
        try:
            received = parsedate_to_datetime(str(p.msg.get("Date"))).astimezone()
        except Exception:  # noqa: BLE001
            received = datetime.now().astimezone()
    base = dict(primit=received.strftime("%Y-%m-%d %H:%M"), expeditor=p.expeditor, subiect=p.subiect[:300],
                linkuri=json.dumps(p.linkuri) if p.linkuri else None)

    if p.expeditor == cont and cont not in p.destinatari:
        db.record(cont, meta.msgid, stare="trimis_altora", **base)
        return {"stare": "trimis_altora"}
    if any(x.lower() in p.expeditor for x in cfg["ignora_expeditori"]):
        db.record(cont, meta.msgid, stare="ignorat", **base)
        return {"stare": "ignorat"}
    if not p.atasamente:
        db.record(cont, meta.msgid, stare="fara_atasamente", **base)
        if p.linkuri and p.expeditor not in own:
            append_journal(cfg, {"tip": "link", "cont": cont, "msgid": meta.msgid, "primit": base["primit"],
                                 "expeditor": p.expeditor, "subiect": p.subiect, "linkuri": p.linkuri})
        return {"stare": "fara_atasamente", "linkuri": len(p.linkuri)}

    client, motiv = classify(p, reg, own)
    year, month = config.work_month(received.date(), int(cfg["luna_minus"]))
    luna = f"{year:04d}-{month:02d}"
    root = config.root(cfg)
    if client:
        folder = root / client.folder / config.month_path(cfg["folder_client"], year, month)
    else:
        folder = root / config.month_path(cfg["folder_neclasificat"], year, month)

    saved, dup = 0, 0
    for att in p.atasamente:
        sha = hashlib.sha256(att.data).hexdigest()
        prior = db.sha_in_dir(sha, str(folder))
        name = clean_name(f"{received:%Y-%m-%d}_{att.name}")
        if prior:
            path, new = Path(prior), False
        else:
            path, new = write_unique(folder, name, att.data, sha)
        state = "salvat" if new else "dublura"
        saved += new
        dup += not new
        db.conn.execute(
            "INSERT INTO fisiere (cont, msgid, nume_original, cale, sha256, marime, client, luna, stare, creat_la)"
            " VALUES (?,?,?,?,?,?,?,?,?,?)",
            (cont, meta.msgid, att.name, str(path), sha, att.size, client.folder if client else None, luna, state, _now()))
        db.conn.commit()
        append_journal(cfg, {
            "tip": "fisier", "stare": state, "cont": cont, "msgid": meta.msgid, "primit": base["primit"],
            "expeditor": p.expeditor, "expeditori_originali": p.expeditori_originali, "forward": p.forward,
            "subiect": p.subiect, "nume_original": att.name, "din_eml": att.din_eml,
            "cale": os.path.relpath(path, root), "sha256": sha, "marime": att.size,
            "client": client.folder if client else None, "cui": client.cui if client else "",
            "luna": luna, "motiv": motiv,
        })
    if p.linkuri and p.expeditor not in own:
        append_journal(cfg, {"tip": "link", "cont": cont, "msgid": meta.msgid, "primit": base["primit"],
                             "expeditor": p.expeditor, "subiect": p.subiect, "linkuri": p.linkuri})
    db.record(cont, meta.msgid, stare="ok", client=client.folder if client else None, motiv=motiv,
              fisiere=len(p.atasamente), eroare=None, **base)
    return {"stare": "ok", "salvate": saved, "dubluri": dup, "client": client.folder if client else None}


def sync_account(cfg: dict, cont: str, reg: Registry, db: Db, imap=None) -> dict:
    stats = {"cont": cont, "mesaje_noi": 0, "fisiere_salvate": 0, "dubluri": 0, "erori": 0, "neclasificate": 0}
    own_imap = imap is None
    imap = imap or connect(cont)
    try:
        folder = all_mail_folder(imap)
        typ, _ = imap.select(folder, readonly=True)
        if typ != "OK":
            raise RuntimeError(f"Nu pot deschide dosarul {folder}")
        floor = date.fromisoformat(cfg["de_la"])
        last_ok = db.get_meta(f"ultima_ok:{cont}")
        since = floor
        if last_ok:
            since = max(floor, datetime.fromisoformat(last_ok).date() - timedelta(days=OVERLAP_DAYS))
        uids = search_since(imap, since)
        metas = [m for m in fetch_meta(imap, uids) if not db.done(cont, m.msgid)]
        for m in metas:
            if m.internal and m.internal.astimezone().date() < floor:
                continue
            if "\\Draft" in m.labels:
                db.record(cont, m.msgid, stare="ciorna")
                continue
            stats["mesaje_noi"] += 1
            try:
                raw = fetch_raw(imap, m.uid)
                r = process_message(cfg, cont, m, raw, reg, db)
                stats["fisiere_salvate"] += r.get("salvate", 0)
                stats["dubluri"] += r.get("dubluri", 0)
                if r.get("stare") == "ok" and not r.get("client"):
                    stats["neclasificate"] += r.get("salvate", 0)
            except (imaplib.IMAP4.abort, socket.error, OSError) as exc:
                # Conexiune sau disc: oprim contul; se reia la urmatoarea rulare.
                db.fail(cont, m.msgid, f"{type(exc).__name__}: {exc}")
                stats["erori"] += 1
                raise
            except Exception as exc:  # noqa: BLE001 - un mesaj stricat nu opreste restul
                log.exception("Mesaj %s/%s", cont, m.msgid)
                db.fail(cont, m.msgid, f"{type(exc).__name__}: {exc}")
                stats["erori"] += 1
        control(imap, cont, db)
        db.set_meta(f"ultima_ok:{cont}", datetime.now().isoformat(timespec="seconds"))
    finally:
        if own_imap:
            try:
                imap.logout()
            except Exception:  # noqa: BLE001
                pass
    return stats


def control(imap, cont: str, db: Db) -> None:
    """Numaratoarea de control pe ziua de ieri: mesaje in Gmail vs. mesaje procesate."""
    today = date.today()
    y = today - timedelta(days=1)
    uids = search_since(imap, y, today)
    metas = fetch_meta(imap, uids)
    pending = [m.msgid for m in metas if not db.done(cont, m.msgid)]
    db.set_meta(f"control:{cont}:{y.isoformat()}", {"in_gmail": len(metas), "neprocesate": pending,
                                                     "verificat_la": _now()})


def run(cfg: dict | None = None, conturi: list[str] | None = None, db: Db | None = None,
        imap_factory=None) -> dict:
    """O rulare completa: toate conturile, apoi raportul (daca e ora) sau alerta (la eroare)."""
    from . import raport

    cfg = cfg or config.load()
    db = db or Db()
    started = time.time()
    with run_lock(config.HOME / "sync.lock") as got:
        if not got:
            return {"rezultat": "sarit", "motiv": "o alta descarcare e in curs"}
        result: dict = {"inceput": _now(), "conturi": [], "erori": []}
        try:
            check_root(cfg)
            reg = Registry.load(cfg, config.root(cfg))
        except Exception as exc:  # noqa: BLE001
            result["erori"].append(f"pregatire: {exc}")
            reg = None
        if reg is not None:
            for cont in conturi or cfg["conturi"]:
                try:
                    imap = imap_factory(cont) if imap_factory else None
                    result["conturi"].append(sync_account(cfg, cont, reg, db, imap))
                except Exception as exc:  # noqa: BLE001
                    log.exception("Cont %s", cont)
                    result["erori"].append(f"{cont}: {type(exc).__name__}: {exc}")
        result["durata_sec"] = round(time.time() - started, 1)
        result["rezultat"] = "ok" if not result["erori"] else "eroare"
        db.set_meta("ultima_rulare", result)
        log.info("Rulare: %s", json.dumps(result, ensure_ascii=False))
        try:
            raport.after_run(cfg, db, result)
        except Exception:  # noqa: BLE001
            log.exception("Raport")
        return result
