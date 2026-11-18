"""Desfacerea unui mesaj: atasamente (inclusiv din mesajele .eml atasate), expeditorul
original la forward, linkurile Drive/WeTransfer."""

from __future__ import annotations

import html
import mimetypes
import re
from dataclasses import dataclass, field
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from email.utils import getaddresses

from .clienti import EMAIL_RE

FWD_SUBJECT = re.compile(r"^\s*(re:\s*)*(fwd?|fw|tr|wg|redirec\w*|inaintat|înaintat|transf\w*)\s*:", re.I)
FWD_MARKERS = re.compile(
    r"(-{3,}\s*(forwarded message|mesaj redirec\w+|original message|mesaj original|weitergeleitete nachricht)\s*-{3,}"
    r"|begin forwarded message:|^_{10,}\s*$)",
    re.I | re.M,
)
FROM_LINE = re.compile(r"^\s*[>*\s]*(from|de la|von|expeditor|de)\s*:\s*\**(.+)$", re.I | re.M)
LINK_RE = re.compile(
    r"https?://(?:drive\.google\.com|docs\.google\.com|wetransfer\.com|we\.tl|[\w.-]*dropbox\.com"
    r"|1drv\.ms|onedrive\.live\.com|[\w.-]*sharepoint\.com)/[^\s\"'<>)\]]+",
    re.I,
)
SKIP_NAMES = re.compile(r"^(winmail\.dat|smime\.p7s|signature\.asc)$", re.I)
SKIP_EXT = {".ics", ".vcf", ".p7s", ".asc"}
SIG_IMAGE = re.compile(r"^(image|outlook|logo)[\w\-]*\.(png|jpe?g|gif|bmp)$", re.I)


@dataclass
class Attachment:
    name: str
    data: bytes
    content_type: str
    din_eml: bool = False
    expeditor_eml: str = ""

    @property
    def size(self) -> int:
        return len(self.data)


@dataclass
class Parsed:
    msg: EmailMessage
    expeditor: str
    expeditor_nume: str
    destinatari: list[str]
    subiect: str
    corp: str
    forward: bool
    expeditori_originali: list[str]
    nota_forward: str
    atasamente: list[Attachment] = field(default_factory=list)
    ignorate: list[str] = field(default_factory=list)
    linkuri: list[str] = field(default_factory=list)


def parse_bytes(raw: bytes) -> EmailMessage:
    return BytesParser(policy=policy.default).parsebytes(raw)


def _addrs(msg, *headers) -> list[tuple[str, str]]:
    vals = []
    for h in headers:
        for v in msg.get_all(h, []) or []:
            vals.append(str(v))
    return [(n, a.lower()) for n, a in getaddresses(vals) if a and "@" in a]


def body_text(msg: EmailMessage) -> str:
    part = msg.get_body(preferencelist=("plain", "html"))
    if part is None:
        return ""
    try:
        text = part.get_content()
    except Exception:  # noqa: BLE001 - codare stricata
        payload = part.get_payload(decode=True) or b""
        text = payload.decode("utf-8", "replace")
    if part.get_content_subtype() == "html":
        text = re.sub(r"(?is)<(script|style).*?</\1>", " ", text)
        text = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", text)
        text = html.unescape(re.sub(r"<[^>]+>", " ", text))
    return text


def _safe_name(name: str, ctype: str, n: int) -> str:
    name = (name or "").strip().strip('"')
    if not name:
        ext = mimetypes.guess_extension(ctype) or ".bin"
        name = f"atasament{n}{ext}"
    return name


def _walk(part: EmailMessage, out: list[Attachment], skipped: list[str], depth: int, eml_from: str) -> None:
    ctype = part.get_content_type()
    if ctype == "message/rfc822":
        payload = part.get_payload()
        inner = payload[0] if isinstance(payload, list) and payload else None
        if inner is not None:
            frm = _addrs(inner, "From")
            _walk(inner, out, skipped, depth + 1, frm[0][1] if frm else eml_from)
        return
    if part.is_multipart():
        for sub in part.iter_parts():
            _walk(sub, out, skipped, depth, eml_from)
        return
    filename = part.get_filename()
    disp = part.get_content_disposition()
    if not filename and disp != "attachment":
        return  # corpul mesajului
    data = part.get_payload(decode=True) or b""
    name = _safe_name(filename or "", ctype, len(out) + len(skipped) + 1)
    lower = name.lower()
    ext = "." + lower.rsplit(".", 1)[-1] if "." in lower else ""
    if not data:
        skipped.append(f"{name} (gol)")
    elif SKIP_NAMES.match(lower) or ext in SKIP_EXT:
        skipped.append(name)
    elif ctype.startswith("image/") and (len(data) < 20_000 or (SIG_IMAGE.match(lower) and disp != "attachment" and len(data) < 100_000)):
        skipped.append(f"{name} (imagine de semnatura)")
    else:
        out.append(Attachment(name=name, data=data, content_type=ctype, din_eml=depth > 0, expeditor_eml=eml_from if depth > 0 else ""))


def analyze(raw: bytes, own: set[str]) -> Parsed:
    msg = parse_bytes(raw)
    frm = _addrs(msg, "From")
    sender_name, sender = (frm[0] if frm else ("", ""))
    rcpt = [a for _, a in _addrs(msg, "To", "Cc", "Bcc", "Delivered-To", "X-Original-To")]
    subject = str(msg.get("Subject", "") or "")
    text = body_text(msg)

    atts: list[Attachment] = []
    skipped: list[str] = []
    _walk(msg, atts, skipped, 0, "")

    marker = FWD_MARKERS.search(text)
    forward = bool(FWD_SUBJECT.match(subject) or marker or any(a.din_eml for a in atts))
    originals: list[str] = []
    note = ""
    if forward:
        tail = text[marker.start():] if marker else text
        note = text[: marker.start()] if marker else ""
        for m in FROM_LINE.finditer(tail):
            for e in EMAIL_RE.findall(m.group(2)):
                e = e.lower()
                if e not in own and e not in originals:
                    originals.append(e)
        for a in atts:
            if a.expeditor_eml and a.expeditor_eml not in own and a.expeditor_eml not in originals:
                originals.append(a.expeditor_eml)

    links = sorted(set(LINK_RE.findall(text)))
    return Parsed(msg=msg, expeditor=sender, expeditor_nume=sender_name, destinatari=rcpt, subiect=subject,
                  corp=text, forward=forward, expeditori_originali=originals, nota_forward=note[:2000],
                  atasamente=atts, ignorate=skipped, linkuri=links)
