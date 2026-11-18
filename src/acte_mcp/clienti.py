"""Clientii: folderele de pe server + Excelul cu adrese de e-mail.

Folderele clientilor sunt pe radacina, cu numele "<numar>. <Nume firma>".
Excelul (Clienti_Contabilitate.xlsx, foaia "Clienti", antet pe randul 4) da, pentru
fiecare firma: B nume, C CUI, F e-mail, K alte e-mailuri / domenii.
Un client din Excel se leaga de folderul cu acelasi nume (fara numar, SRL, diacritice).
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

# Domenii publice: nu identifica o firma, doar adresa exacta conteaza.
GENERIC_DOMAINS = {
    "gmail.com", "googlemail.com", "yahoo.com", "yahoo.ro", "yahoo.co.uk", "hotmail.com",
    "outlook.com", "live.com", "icloud.com", "me.com", "aol.com", "mail.com", "gmx.com",
    "gmx.net", "protonmail.com", "proton.me", "zoho.com", "msn.com", "ymail.com",
}

_LEGAL = {"srl", "srld", "pfa", "ii", "if", "sa", "sca", "snc", "asociatia", "sc", "ong", "ltd", "gmbh"}
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
DOMAIN_RE = re.compile(r"(?<![A-Za-z0-9.\-@])@?([A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,})")


def strip_diacritics(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))


def norm(text: str) -> str:
    """Numele firmei comparabil: fara numarul din fata, forma juridica, diacritice, punctuatie."""
    t = strip_diacritics(text or "").lower()
    t = re.sub(r"^\s*\d+\s*[.)\-]\s*", "", t)
    t = t.replace("s.r.l.", "srl").replace("s.r.l", "srl").replace("s.a.", "sa").replace("srl-d", "srld")
    words = [w for w in re.split(r"[^a-z0-9]+", t) if w and w not in _LEGAL]
    return " ".join(words)


def norm_cui(text) -> str:
    return re.sub(r"\D", "", str(text or ""))


@dataclass
class Client:
    nume: str
    cui: str = ""
    folder: str | None = None  # numele folderului de pe server
    emailuri: set[str] = field(default_factory=set)
    domenii: set[str] = field(default_factory=set)
    contact: str = ""

    def public(self) -> dict:
        return {"nume": self.nume, "cui": self.cui, "folder": self.folder,
                "emailuri": sorted(self.emailuri), "domenii": sorted(self.domenii)}


def list_folders(root: Path) -> list[str]:
    """Folderele de client de pe radacina (fara cele tehnice, care incep cu _ sau .)."""
    if not root.exists():
        return []
    return sorted(p.name for p in root.iterdir() if p.is_dir() and not p.name.startswith(("_", ".", "$")))


def load_excel(path: str | Path) -> list[Client]:
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = None
    for name in wb.sheetnames:
        if strip_diacritics(name).lower().startswith("clien"):
            ws = wb[name]
            break
    ws = ws or wb.worksheets[0]
    out: list[Client] = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        cells = list(row) + [None] * 12
        nume = str(cells[1] or "").strip()
        if not nume or nume.lower().startswith("exemplu"):
            continue
        c = Client(nume=nume, cui=norm_cui(cells[2]), contact=str(cells[4] or "").strip())
        for cell in (cells[5], cells[10]):
            text = str(cell or "")
            for e in EMAIL_RE.findall(text):
                c.emailuri.add(e.lower())
            rest = EMAIL_RE.sub(" ", text)
            for d in DOMAIN_RE.findall(rest):
                d = d.lower()
                if d not in GENERIC_DOMAINS:
                    c.domenii.add(d)
        out.append(c)
    wb.close()
    return out


class Registry:
    """Leaga adresele de e-mail de folderele clientilor."""

    def __init__(self, clients: list[Client], folders: list[str], manual: dict[str, str] | None = None):
        self.folders = folders
        self.clients = clients
        self.manual = {k.lower().lstrip("@"): v for k, v in (manual or {}).items()}
        self.unmapped: list[str] = []
        by_norm: dict[str, list[str]] = {}
        for f in folders:
            by_norm.setdefault(norm(f), []).append(f)
        for c in clients:
            n = norm(c.nume)
            hits = by_norm.get(n, [])
            if len(hits) != 1 and len(n) >= 4:
                # Potrivire mai lejera: unul din nume il contine pe celalalt, unic.
                hits = [f for f in folders if n and (n in norm(f) or (len(norm(f)) >= 4 and norm(f) in n))]
            if len(hits) == 1:
                c.folder = hits[0]
            else:
                self.unmapped.append(c.nume)
        # Folderele fara rand in Excel devin clienti fara adrese (se pot gasi dupa nume).
        mapped = {c.folder for c in clients if c.folder}
        for f in folders:
            if f not in mapped:
                self.clients.append(Client(nume=re.sub(r"^\s*\d+\s*[.)\-]\s*", "", f), folder=f))

    @classmethod
    def load(cls, cfg: dict, root: Path) -> "Registry":
        clients: list[Client] = []
        if cfg.get("excel_clienti"):
            clients = load_excel(cfg["excel_clienti"])
        return cls(clients, list_folders(root), cfg.get("mapare_manuala"))

    def by_folder(self, folder: str) -> Client | None:
        return next((c for c in self.clients if c.folder == folder), None)

    def by_address(self, addr: str) -> list[Client]:
        addr = addr.lower().strip()
        domain = addr.split("@")[-1]
        for key in (addr, domain):
            if key in self.manual:
                c = self.by_folder(self.manual[key])
                return [c] if c else [Client(nume=self.manual[key], folder=self.manual[key])]
        exact = [c for c in self.clients if addr in c.emailuri]
        if exact:
            return exact
        if domain in GENERIC_DOMAINS:
            return []
        return [c for c in self.clients if domain in c.domenii or any(e.endswith("@" + domain) for e in c.emailuri)]

    def find_in_text(self, text: str, among: list[Client] | None = None) -> list[Client]:
        """Clientii al caror CUI sau nume apare in text."""
        pool = among if among is not None else self.clients
        t = " " + norm(text) + " "
        digits = " " + re.sub(r"[^0-9]+", " ", text or "") + " "
        hits = []
        for c in pool:
            n = norm(c.nume)
            if (c.cui and len(c.cui) >= 4 and f" {c.cui} " in digits) or (len(n) >= 4 and f" {n} " in t):
                hits.append(c)
        return hits

    def resolve(self, query: str) -> list[str]:
        """Folderul (sau folderele) care se potrivesc cu un nume, numar, CUI sau nume de folder."""
        q = (query or "").strip()
        if q in self.folders:
            return [q]
        if q.isdigit():
            pref = [f for f in self.folders if re.match(rf"^\s*{q}\s*[.)\-]", f)]
            if pref:
                return pref
            by_cui = [c.folder for c in self.clients if c.cui == q and c.folder]
            if by_cui:
                return by_cui
        n = norm(q)
        exact = [f for f in self.folders if norm(f) == n]
        if exact:
            return exact
        return [f for f in self.folders if n and n in norm(f)]
