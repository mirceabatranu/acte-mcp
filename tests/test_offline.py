"""Test complet fara internet: un server IMAP simulat, un server local intr-un folder
temporar, un Excel de clienti. Ruleaza: python -m pytest -q"""

from __future__ import annotations

import asyncio
import imaplib
import io
import json
import os
import sys
import tempfile
from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path

import pytest

TMP = Path(tempfile.mkdtemp(prefix="acte-test-"))
os.environ["ACTE_MCP_HOME"] = str(TMP / "home")

from acte_mcp import config, raport, server, sync  # noqa: E402

ROOT = TMP / "Z"
FOLDERS = ["1. Alfa SRL", "2. Beta Consult SRL", "3. Gamma", "4. Vinia SRL", "5. Podgoria SRL"]


def _pdf() -> bytes:
    objs = [b"<</Type/Catalog/Pages 2 0 R>>", b"<</Type/Pages/Kids[3 0 R]/Count 1>>",
            b"<</Type/Page/Parent 2 0 R/MediaBox[0 0 300 200]/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>",
            None, b"<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>"]
    stream = b"BT /F1 14 Tf 20 100 Td (Factura 123 total 1210 lei TVA) Tj ET"
    objs[3] = b"<</Length %d>>stream\n" % len(stream) + stream + b"\nendstream"
    out, offs = bytearray(b"%PDF-1.4\n"), []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj" % i + o + b"endobj\n"
    x = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    for o in offs:
        out += b"%010d 00000 n \n" % o
    out += b"trailer<</Size %d/Root 1 0 R>>\nstartxref\n%d\n%%%%EOF" % (len(objs) + 1, x)
    return bytes(out)


PDF_TEXT = _pdf()


def make_excel(path: Path) -> None:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Clienți"
    ws["A1"] = "Lista clienti"
    for i, h in enumerate(["Nr", "Nume firmă", "CUI", "", "Contact", "E-mail", "Telefon", "", "Link documente", "", "Alte e-mailuri / domenii"]):
        ws.cell(row=4, column=i + 1, value=h)
    rows = [
        ["1", "ALFA SRL", "RO11111111", "", "Ana", "contabil@alfa.ro", "", "", "", "", ""],
        ["2", "Beta Consult S.R.L.", "22222222", "", "Ion", "ion@beta.ro", "", "", "", "", ""],
        ["3", "GAMMA SRL", "33333333", "", "", "", "", "", "", "", "@gamma.com"],
        ["4", "VINIA SRL", "44444444", "", "", "", "", "", "", "", "vinuri.ro"],
        ["5", "PODGORIA SRL", "55555555", "", "", "", "", "", "", "", "vinuri.ro"],
        ["6", "Exemplu SRL", "1", "", "", "x@x.ro", "", "", "", "", ""],
        ["7", "Fara Folder SRL", "77777777", "", "", "a@fara.ro", "", "", "", "", ""],
    ]
    for r, row in enumerate(rows, start=5):
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=v)
    wb.save(path)


def png_bytes(size=(400, 300)) -> bytes:
    from PIL import Image

    buf = io.BytesIO()
    import random

    img = Image.new("RGB", size)
    img.putdata([(random.randrange(256), random.randrange(256), random.randrange(256)) for _ in range(size[0] * size[1])])
    img.save(buf, format="PNG")
    return buf.getvalue()


def mail(frm, to, subject, body, atts=(), eml: EmailMessage | None = None) -> bytes:
    m = EmailMessage()
    m["From"], m["To"], m["Subject"] = frm, to, subject
    m["Date"] = "Tue, 17 Nov 2026 10:00:00 +0200"
    m.set_content(body)
    for name, data, mt in atts:
        maintype, subtype = mt.split("/")
        m.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)
    if eml is not None:
        m.add_attachment(eml)
    return m.as_bytes()


class FakeImap:
    """Raspunde ca Gmail la comenzile folosite de sync (LIST, SELECT, UID SEARCH/FETCH)."""

    def __init__(self, messages: list[dict]):
        self.messages = messages
        self.readonly = None
        self.fail_once: set[int] = set()
        self.fetched_body: list[int] = []

    def list(self):
        return "OK", [b'(\\HasNoChildren) "/" "INBOX"', b'(\\All \\HasNoChildren) "/" "[Gmail]/Toate mesajele"']

    def select(self, folder, readonly=False):
        assert folder == '"[Gmail]/Toate mesajele"'
        self.readonly = readonly
        return "OK", [b"10"]

    def uid(self, cmd, *args):
        if cmd == "SEARCH":
            since = datetime.strptime(args[1], "%d-%b-%Y").date()
            before = datetime.strptime(args[3], "%d-%b-%Y").date() if len(args) > 3 else date.max
            ids = [str(m["uid"]).encode() for m in self.messages if since <= m["date"].date() < before]
            return "OK", [b" ".join(ids)]
        if cmd == "FETCH":
            assert self.readonly is True, "casuta trebuie deschisa doar in citire"
            uids = [int(x) for x in args[0].split(",")]
            if "BODY.PEEK[]" in args[1]:
                u = uids[0]
                if u in self.fail_once:
                    self.fail_once.discard(u)
                    return "NO", [b"temporar"]
                m = next(x for x in self.messages if x["uid"] == u)
                self.fetched_body.append(u)
                return "OK", [(f"1 (UID {u} BODY[] {{{len(m['raw'])}}}".encode(), m["raw"]), b")"]
            out = []
            for m in self.messages:
                if m["uid"] in uids:
                    d = m["date"].strftime("%d-%b-%Y %H:%M:%S +0200")
                    out.append(f'{m["uid"]} (X-GM-MSGID {m["uid"] + 9000} X-GM-LABELS ({m.get("labels", "")}) '
                               f'UID {m["uid"]} INTERNALDATE "{d}" RFC822.SIZE {len(m["raw"])})'.encode())
            return "OK", out
        raise AssertionError(cmd)

    def logout(self):
        pass


@pytest.fixture(scope="module")
def env():
    for f in FOLDERS:
        (ROOT / f).mkdir(parents=True)
    make_excel(TMP / "clienti.xlsx")
    cfg = config.load()
    cfg.update(radacina=str(ROOT), excel_clienti=str(TMP / "clienti.xlsx"), conturi=["cabinet@gmail.com"],
               adrese_proprii=["office@cabinet.ro", "paul@cabinet.ro"])
    cfg["raport"]["ora"] = 0
    cfg["raport"]["catre"] = "office@cabinet.ro"
    config.save(cfg)
    cfg = config.load()

    pdf = PDF_TEXT
    when = datetime(2026, 11, 17, 10, 0)
    inner = EmailMessage()
    inner["From"], inner["To"], inner["Subject"] = "Contabil Gamma <plati@gamma.com>", "x@y.ro", "extras"
    inner.set_content("vezi atasat")
    inner.add_attachment(b"Data;Suma\n01.10.2026;100\n", maintype="text", subtype="csv", filename="extras oct.csv")
    msgs = [
        dict(uid=1, raw=mail("Ana <contabil@alfa.ro>", "cabinet@gmail.com", "facturi octombrie", "Buna",
                             [("factura 123.pdf", pdf, "application/pdf"), ("image001.png", b"x" * 500, "image/png")])),
        dict(uid=2, raw=mail("office@cabinet.ro", "cabinet@gmail.com", "Fwd: acte",
                             "pt contabilitate\n\n---------- Forwarded message ---------\nFrom: Ion Pop <ion@beta.ro>\nDate: ...\n\nsalut",
                             [("stat plata.xlsx", b"PK\x03\x04fake", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")])),
        dict(uid=3, raw=mail("paul@cabinet.ro", "cabinet@gmail.com", "FW: extras", "vezi mai jos", eml=inner)),
        dict(uid=4, raw=mail("cineva@gmail.com", "cabinet@gmail.com", "documente", "atasat",
                             [("poza bon.png", png_bytes(), "image/png")])),
        dict(uid=5, raw=mail("cabinet@gmail.com", "client@alfa.ro", "situatii", "atasat",
                             [("balanta.pdf", pdf, "application/pdf")])),
        dict(uid=6, raw=mail("news@linkedin.com", "cabinet@gmail.com", "job", "x", [("a.pdf", pdf, "application/pdf")])),
        dict(uid=7, raw=mail("contabil@alfa.ro", "cabinet@gmail.com", "link", "https://we.tl/t-abc123 descarcati")),
        dict(uid=8, raw=mail("contabil@alfa.ro", "cabinet@gmail.com", "Re: facturi", "din nou",
                             [("factura 123.pdf", pdf, "application/pdf")])),
        dict(uid=9, raw=mail("cabinet@gmail.com", "", "ciorna", "x", [("c.pdf", pdf, "application/pdf")]), labels="\\\\Draft"),
        dict(uid=10, raw=mail("maria@vinuri.ro", "cabinet@gmail.com", "Facturi Vinia", "pentru VINIA SRL",
                              [("f1.pdf", pdf + b"1", "application/pdf")])),
        dict(uid=11, raw=mail("maria@vinuri.ro", "cabinet@gmail.com", "facturi", "atasat",
                              [("f2.pdf", pdf + b"2", "application/pdf")])),
        dict(uid=12, raw=mail("contabil@alfa.ro", "cabinet@gmail.com", "vechi", "x", [("v.pdf", pdf, "application/pdf")]),
             date=datetime(2026, 10, 20, 9, 0)),
    ]
    for m in msgs:
        m.setdefault("date", when)
    fake = FakeImap(msgs)
    sent: list = []
    raport.send = lambda cfg, s, b: sent.append((s, b))
    return dict(cfg=cfg, fake=fake, sent=sent)


def run(env):
    return sync.run(env["cfg"], imap_factory=lambda cont: env["fake"])


def test_sync_first_run(env):
    env["fake"].fail_once.add(4)
    r = run(env)
    assert r["rezultat"] == "ok", r
    st = r["conturi"][0]
    assert st["erori"] == 1  # mesajul 4 a esuat temporar
    alfa = ROOT / "1. Alfa SRL" / "Acte contabile" / "2026" / "10"
    assert sorted(p.name for p in alfa.iterdir()) == ["2026-11-17_factura 123.pdf"]  # dublura din msg 8 sarita
    assert (ROOT / "2. Beta Consult SRL/Acte contabile/2026/10/2026-11-17_stat plata.xlsx").exists()
    assert (ROOT / "3. Gamma/Acte contabile/2026/10/2026-11-17_extras oct.csv").exists()  # din .eml atasat
    assert (ROOT / "4. Vinia SRL/Acte contabile/2026/10/2026-11-17_f1.pdf").exists()
    dv = ROOT / "_De verificat/2026/10"
    assert sorted(p.name for p in dv.iterdir()) == ["2026-11-17_f2.pdf"]  # vinuri.ro fara firma numita
    assert not list(ROOT.rglob("*balanta*")) and not list(ROOT.rglob("a.pdf")) and not list(ROOT.rglob("*c.pdf"))
    assert not list(ROOT.rglob("*v.pdf"))  # inainte de data de start
    db = sync.Db()
    stari = dict(db.conn.execute("SELECT msgid, stare FROM mesaje").fetchall())
    assert stari["9005"] == "trimis_altora" and stari["9006"] == "ignorat" and stari["9007"] == "fara_atasamente"
    assert stari["9009"] == "ciorna" and stari["9004"] == "eroare"
    motiv = db.conn.execute("SELECT motiv FROM mesaje WHERE msgid='9011'").fetchone()[0]
    assert "VINIA" in motiv and "PODGORIA" in motiv
    assert env["sent"], "raportul zilnic trebuie trimis"


def test_sync_second_run_retries_and_no_duplicates(env):
    before = len(env["fake"].fetched_body)
    r = run(env)
    assert r["rezultat"] == "ok", r
    assert env["fake"].fetched_body[before:] == [4]  # doar mesajul care esuase
    assert (ROOT / "_De verificat/2026/10/2026-11-17_poza bon.png").exists()
    r = run(env)
    assert r["conturi"][0]["mesaje_noi"] == 0


def test_report(env):
    db = sync.Db()
    subject, body = raport.daily(env["cfg"], db, date.today())
    assert "1. Alfa SRL" in body and "we.tl" in body and "DE VERIFICAT" in body


def test_mcp_tools(env):
    a = asyncio.run
    cl = a(server.acte_clienti())
    assert cl["numar"] == 5 and "Fara Folder SRL" in cl["din_excel_fara_folder"]
    lst = a(server.acte_lista("alfa", "2026-10"))
    assert lst["numar"] == 1 and lst["cui"] == "11111111"
    f = lst["fisiere"][0]
    assert f["origine"]["expeditor"] == "contabil@alfa.ro"
    text = a(server.acte_citeste(f["cale"]))
    assert "Factura 123" in text[1]
    img = a(server.acte_citeste(f["cale"], ca_imagine=True))
    assert any(type(x).__name__ == "Image" for x in img)
    a(server.acte_marcheaza(f["cale"], "importat_saga", "test"))
    assert a(server.acte_lista("1", "2026-10", doar_neimportate=True))["numar"] == 0
    a(server.acte_marcheaza(f["cale"], "anulat"))
    assert a(server.acte_lista("11111111", "2026-10", doar_neimportate=True))["numar"] == 1
    # mutare din _De verificat
    dv = a(server.acte_lista("_De verificat", "2026-10"))
    f2 = next(x for x in dv["fisiere"] if x["nume"].endswith("f2.pdf"))
    plan = a(server.acte_muta(f2["cale"], "Podgoria"))
    assert "mutat" not in plan and plan["plan"]["luna"] == "2026-10"
    a(server.acte_muta(f2["cale"], "Podgoria", confirma=True))
    moved = a(server.acte_lista("Podgoria", "2026-10"))
    assert moved["numar"] == 1 and moved["fisiere"][0]["origine"]["expeditor"] == "maria@vinuri.ro"
    assert a(server.acte_cauta("vinuri"))["numar"] == 2
    assert a(server.acte_linkuri())["numar"] == 1
    with pytest.raises(ValueError):
        a(server.acte_citeste("..\\..\\secret.txt"))
    st = a(server.acte_status())
    assert st["server_accesibil"] and st["descarcare_pe_acest_calculator"]
