#!/usr/bin/env python3
"""Build a searchable archive from LISTSERV monthly log files (.eml).

Each input .eml is an envelope from LISTSERV whose body is a monthly
"notebook" log: many list messages concatenated, each preceded by a line of
73 '=' characters and starting with its own RFC 822 headers.

Usage:  python3 -I tools/build_archive.py RAW_DIR OUTPUT_DIR [--name NAME] [--web-images]

  --name NAME    list name used in the report title (default: LIST)
  --web-images   store smaller web copies of large images (PNG flyers become
                 JPEG, max 1600 px) using macOS `sips`. The .eml originals keep
                 the full-quality images.

Reads RAW_DIR/*.eml (never modified). Writes to OUTPUT_DIR:
  archive.sqlite   messages + attachments tables, FTS5 full-text index
  messages.csv     one row per message (UTF-8 with BOM, opens in Excel)
  report.md        validation report
Standard library only; no network access.
"""
import csv
import email
import email.utils
import hashlib
import html
import re
import sqlite3
import sys
from collections import Counter, defaultdict
from datetime import timezone
from email import policy
from email.header import decode_header, make_header
from html.parser import HTMLParser
from pathlib import Path

SEP_RE = re.compile(rb"(?m)^={73}\r?\n(?=[A-Za-z][A-Za-z0-9-]*:)")
EXCEL_CELL_LIMIT = 32000  # Excel max is 32,767; leave headroom

# --------------------------------------------------------------------------
# HTML -> text


class _HTMLText(HTMLParser):
    BLOCK = {"p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4", "h5",
             "h6", "blockquote", "table", "ul", "ol", "hr"}
    SKIP = {"script", "style", "head", "title"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out, self.links, self._skip, self._quote = [], [], 0, 0
        self._href, self._anchor_start = None, 0

    def handle_starttag(self, tag, attrs):
        if tag in self.SKIP:
            self._skip += 1
        if tag == "blockquote":
            self._quote += 1
        if tag in self.BLOCK:
            self.out.append("\n")
        if tag == "a":
            href = (dict(attrs).get("href") or "").strip()
            if href.lower().startswith(("http://", "https://")):
                self.links.append(href)
                self._href, self._anchor_start = href, len(self.out)

    def handle_endtag(self, tag):
        if tag == "a" and self._href:
            anchor = "".join(self.out[self._anchor_start:]).strip()
            if self._href not in anchor:  # keep the URL visible in the text
                self.handle_data(f" <{self._href}>")
            self._href = None
        if tag in self.SKIP and self._skip:
            self._skip -= 1
        if tag == "blockquote" and self._quote:
            self._quote -= 1
        if tag in self.BLOCK:
            self.out.append("\n")

    def handle_data(self, data):
        if self._skip:
            return
        if self._quote:  # mark quoted html content like plain-text quoting
            data = "\n".join("> " + ln if ln.strip() else ln
                             for ln in data.split("\n"))
        self.out.append(data)


def html_to_text(src):
    p = _HTMLText()
    try:
        p.feed(src)
        p.close()
    except Exception:
        return re.sub(r"<[^>]+>", " ", html.unescape(src)), []
    text = "".join(p.out).replace("\xa0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip(), p.links


# --------------------------------------------------------------------------
# Header helpers


def hdr(msg, name, flags):
    """Return a decoded header value as str ('' if absent)."""
    try:
        v = msg.get(name)
        if v is None:
            return ""
        return re.sub(r"\s+", " ", str(v)).strip()
    except Exception:
        # Fall back to the raw value with RFC 2047 decoding attempted.
        flags.add(f"header_parse_error:{name}")
        for k, raw in msg.raw_items():
            if k.lower() == name.lower():
                try:
                    return re.sub(r"\s+", " ",
                                  str(make_header(decode_header(raw)))).strip()
                except Exception:
                    return re.sub(r"\s+", " ", str(raw)).strip()
        return ""


def norm_msgid(v):
    v = (v or "").strip()
    m = re.search(r"<([^<>]+)>", v)
    v = m.group(1) if m else v
    return v.strip().lower()


def all_msgids(v):
    return [x.strip().lower() for x in re.findall(r"<([^<>]+)>", v or "")]


SUBJ_PREFIX = re.compile(r"^\s*((re|fw|fwd|aw|sv)\s*(\[\d+\])?\s*:|\[[^\]]*\])\s*",
                         re.I)


def norm_subject(s):
    s = s or ""
    while True:
        n = SUBJ_PREFIX.sub("", s, count=1)
        if n == s:
            break
        s = n
    return re.sub(r"\s+", " ", s).strip().lower()


# --------------------------------------------------------------------------
# Body extraction


def part_text(part, flags):
    try:
        return part.get_content()
    except Exception:
        flags.add("charset_fallback")
        payload = part.get_payload(decode=True) or b""
        cs = part.get_content_charset() or "utf-8"
        try:
            return payload.decode(cs, errors="replace")
        except LookupError:
            return payload.decode("utf-8", errors="replace")


def keep_attachment(ctype, disp, size):
    """Which attachments get saved to disk (all are listed in the DB)."""
    if size == 0 or ctype in ("application/pkcs7-signature",
                              "application/x-pkcs7-signature",
                              "application/applefile"):  # Mac resource-fork stubs
        return False
    if ctype.startswith("image/"):
        # signature logos are small inline images; keep deliberate
        # attachments and large images (flyers, screenshots)
        return disp == "attachment" or size > 100_000
    return True


# --------------------------------------------------------------------------
# Attachment storage + text extraction

SAFE_EXT = re.compile(r"^[a-z0-9]{1,8}$")
EXT_BY_TYPE = {"application/pdf": "pdf", "application/msword": "doc",
               "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
               "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
               "image/jpeg": "jpg", "image/png": "png", "image/gif": "gif",
               "image/tiff": "tif", "text/calendar": "ics"}


def stored_ext(filename, ctype):
    ext = Path(filename or "").suffix.lower().lstrip(".")
    if SAFE_EXT.match(ext or ""):
        return ext
    return EXT_BY_TYPE.get(ctype, "bin")


MAX_XML = 50 * 1024 * 1024  # zip-bomb guard for docx/pptx members


def _ooxml_text(path, member_re, para_tag):
    import zipfile
    import xml.etree.ElementTree as ET
    out = []
    with zipfile.ZipFile(path) as z:
        names = sorted((n for n in z.namelist() if re.fullmatch(member_re, n)),
                       key=lambda n: [int(x) if x.isdigit() else x
                                      for x in re.split(r"(\d+)", n)])
        for n in names:
            if z.getinfo(n).file_size > MAX_XML:
                raise ValueError(f"{n} too large")
            root = ET.fromstring(z.read(n))
            for p in root.iter():
                if p.tag.endswith("}" + para_tag):
                    t = "".join(x.text or "" for x in p.iter() if x.tag.endswith("}t"))
                    if t.strip():
                        out.append(t)
    return "\n".join(out)


def extract_attachment_text(path, ext):
    """Return (text, status). Never executes or renders the file."""
    try:
        if ext == "pdf":
            try:
                import logging
                from pypdf import PdfReader
                logging.getLogger("pypdf").setLevel(logging.ERROR)
            except ImportError:
                return "", "skipped_no_pypdf"
            reader = PdfReader(str(path))
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except Exception:
                    return "", "encrypted"
            text = "\n".join((pg.extract_text() or "") for pg in reader.pages)
            return text.strip(), "ok" if text.strip() else "no_text_layer"
        if ext == "docx":
            t = _ooxml_text(path, r"word/(document|footnotes|endnotes)\.xml", "p")
            return t, "ok" if t else "empty"
        if ext == "pptx":
            t = _ooxml_text(path, r"ppt/slides/slide\d+\.xml", "p")
            return t, "ok" if t else "empty"
        if ext in ("doc", "rtf") and Path("/usr/bin/textutil").exists():
            # macOS built-in converter: reads text, never runs macros
            import subprocess
            res = subprocess.run(["/usr/bin/textutil", "-convert", "txt", "-stdout", str(path)],
                                 capture_output=True, timeout=60)
            t = res.stdout.decode("utf-8", "replace").strip()
            return t, "ok" if t else "empty"
        if ext in ("txt", "csv", "ics"):
            return path.read_bytes().decode("utf-8", "replace"), "ok"
        return "", "not_supported"
    except Exception as e:
        return "", f"error:{type(e).__name__}"


WEB_MAX_PX = 1600
WEB_MIN_BYTES = 200_000  # smaller images are left alone


def _sips(*args):
    import subprocess
    return subprocess.run(["/usr/bin/sips", *args], capture_output=True, text=True,
                          timeout=120, check=False)


def alpha_is_used(path):
    """True if a PNG with an alpha channel has any non-opaque pixel.

    Many design-tool exports carry an alpha channel that is fully opaque;
    those are safe to convert to JPEG. Decodes via sips to a 32-bit BMP."""
    import struct, tempfile
    with tempfile.TemporaryDirectory() as td:
        bmp = Path(td) / "a.bmp"
        if _sips("-s", "format", "bmp", str(path), "--out", str(bmp)).returncode != 0 \
                or not bmp.is_file():
            return True
        d = bmp.read_bytes()
    try:
        off, hsize = struct.unpack("<II", d[10:18])
        bpp, comp = struct.unpack("<HI", d[28:34])
        if bpp != 32:
            return False  # no alpha written
        amask = struct.unpack("<I", d[66:70])[0] if comp == 3 and hsize >= 56 else 0xFF000000
        if amask not in (0xFF, 0xFF00, 0xFF0000, 0xFF000000):
            return True
        idx = {0xFF: 0, 0xFF00: 1, 0xFF0000: 2, 0xFF000000: 3}[amask]
        alpha = d[off + idx::4]
        return alpha.count(255) != len(alpha)
    except struct.error:
        return True


def make_web_image(path):
    """Replace a large stored PNG/JPEG with a smaller web copy, if that helps.

    Opaque PNGs become JPEG (quality 80); PNGs with transparency stay PNG and
    are only resized. Returns the (possibly new) path and content type, or
    None if the file was left as-is."""
    ext = path.suffix.lower()
    if ext not in (".png", ".jpg", ".jpeg") or path.stat().st_size < WEB_MIN_BYTES:
        return None
    info = _sips("-g", "hasAlpha", "-g", "pixelWidth", "-g", "pixelHeight", str(path)).stdout
    alpha = "hasAlpha: yes" in info and ext == ".png" and alpha_is_used(path)
    dims = [int(x) for x in re.findall(r"pixel(?:Width|Height): (\d+)", info)]
    if len(dims) != 2:
        return None
    to_jpeg = ext == ".png" and not alpha
    tmp = path.with_name(path.stem + ".web" + (".jpg" if to_jpeg or ext != ".png" else ".png"))
    args = []
    if to_jpeg or ext != ".png":
        args += ["-s", "format", "jpeg", "-s", "formatOptions", "80"]
    if max(dims) > WEB_MAX_PX:
        args += ["-Z", str(WEB_MAX_PX)]
    if not args:
        return None
    r = _sips(*args, str(path), "--out", str(tmp))
    if r.returncode != 0 or not tmp.is_file() or tmp.stat().st_size >= path.stat().st_size * 0.9:
        tmp.unlink(missing_ok=True)
        return None
    new = path.with_suffix(".jpg" if tmp.suffix == ".jpg" else path.suffix)
    path.chmod(0o644)
    path.unlink()
    tmp.rename(new)
    new.chmod(0o444)
    return new, ("image/jpeg" if new.suffix == ".jpg" else "image/png")


def store_attachments(rows, out, web_images=False):
    """Write kept attachments once each (named by SHA-256) and extract text."""
    base = out / "attachments"
    files = {}
    for r in rows:
        for a in r.get("_atts", []):
            data = a.pop("_data", None)
            a["sha256"] = a["stored_path"] = ""
            if not data:
                continue
            h = hashlib.sha256(data).hexdigest()
            a["sha256"] = h
            if h not in files:
                ext = stored_ext(a["filename"], a["content_type"])
                rel = Path("attachments") / h[:2] / f"{h}.{ext}"
                p = out / rel
                p.parent.mkdir(parents=True, exist_ok=True)
                p.write_bytes(data)
                p.chmod(0o444)
                files[h] = {"sha256": h, "stored_path": str(rel), "ext": ext,
                            "content_type": a["content_type"], "size_bytes": len(data),
                            "first_filename": a["filename"], "first_msg_key": r["msg_key"],
                            "times_attached": 0}
            files[h]["times_attached"] += 1
            a["stored_path"] = files[h]["stored_path"]
    if web_images:
        for f in files.values():
            f["original_bytes"] = f["size_bytes"]
            if not f["content_type"].startswith("image/"):
                continue
            res = make_web_image(out / f["stored_path"])
            if res:
                newp, ctype = res
                f["stored_path"] = str(newp.relative_to(out))
                f["content_type"] = ctype
                f["size_bytes"] = newp.stat().st_size
        by_sha = {h: f["stored_path"] for h, f in files.items()}
        for r in rows:
            for a in r.get("_atts", []):
                if a.get("sha256") in by_sha:
                    a["stored_path"] = by_sha[a["sha256"]]
    for f in files.values():
        if f["content_type"].startswith("image/"):
            f["text"], f["text_status"] = "", "image"
        else:
            f["text"], f["text_status"] = extract_attachment_text(out / f["stored_path"], f["ext"])
    return files


def extract_body(msg, flags):
    plain, htmls, atts, html_links = [], [], [], []
    for part in msg.walk():
        if part.is_multipart():
            continue
        ctype = part.get_content_type()
        disp = (part.get_content_disposition() or "").lower()
        fname = part.get_filename()
        is_att = (disp == "attachment" or fname is not None
                  or part.get_content_maintype() not in ("text",))
        if is_att:
            try:
                data = part.get_payload(decode=True) or b""
            except Exception:
                data = b""
                flags.add("attachment_decode_error")
            atts.append({"filename": fname or "", "content_type": ctype,
                         "size_bytes": len(data),
                         "inline": disp != "attachment",
                         "_data": data if keep_attachment(ctype, disp, len(data)) else None})
            continue
        if ctype == "text/plain":
            plain.append(part_text(part, flags))
        elif ctype == "text/html":
            htmls.append(part_text(part, flags))
        else:  # text/calendar, text/enriched, etc.
            atts.append({"filename": fname or "", "content_type": ctype,
                         "size_bytes": len(part.get_payload(decode=True) or b""),
                         "inline": True})
    for h in htmls:
        html_links.extend(html_to_text(h)[1])
    if any(p.strip() for p in plain):
        return "\n\n".join(p.strip("\n") for p in plain), "plain", atts, html_links
    if htmls:
        return "\n\n".join(html_to_text(h)[0] for h in htmls), "html", atts, html_links
    flags.add("no_text_body")
    return "", "none", atts, html_links


# --------------------------------------------------------------------------
# Cleaning (quoted replies, list footer, signatures)

FOOTER_START = re.compile(
    r"(Discussion Listserv is a benefit of membership"
    r"|Click on this link to unsubscribe from (this|the [\w-]+) list"
    r"|To unsubscribe from the [\w-]+ list"
    r"|^#{20,}\s*$)", re.I | re.M)
REPLY_HDR_PATTERNS = [
    re.compile(r"^\s*-{2,}\s*Original Message\s*-{2,}", re.I),
    re.compile(r"^\s*-{2,}\s*Reply message\s*-{2,}", re.I),
    re.compile(r"^\s*Begin forwarded message:", re.I),
    re.compile(r"^\s*-{2,}\s*Forwarded message\s*-{2,}", re.I),
]
ON_WROTE = re.compile(r"^\s*(On|Le|Am|El)\b.{0,300}?(wrote|écrit|schrieb|escribió)\s*:\s*$",
                      re.I | re.S)
# e.g. "At 11:16 AM 1/15/2019, Jane Example wrote:" or
#      "John Sample <...> wrote on 6/3/2020 at 8:14:"
WROTE_DATED = re.compile(r"^\s*\S.{0,250}?(\b\d{1,2}:\d{2}\b|\b(19|20)\d\d\b).{0,200}?"
                         r"\bwrote\b[^\n]{0,40}:\s*$", re.I | re.S)
OUTLOOK_FROM =re.compile(r"^\s*\**\s*From:\s*\**\s*\S", re.I)
OUTLOOK_NEXT = re.compile(r"^\s*\**\s*(Sent|Date|To|Subject):", re.I)
MOBILE_SIG = re.compile(r"^\s*(Sent from my |Sent from Mail for |Get Outlook for |"
                        r"Sent from Yahoo Mail|Sent via )", re.I)
SIGNOFF = re.compile(r"^\s*(best|best regards|best wishes|all the best|regards|kind regards|"
                     r"warm regards|warmly|warmest|thanks|thank you|many thanks|thx|"
                     r"cheers|sincerely|respectfully|cordially|take care|"
                     r"be well|peace|with gratitude|gratefully|appreciatively|"
                     r"in gratitude|onward|stay well)\b[^\n]{0,25}[,.!]?\s*$", re.I)


def clean_body(text, flags):
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    cut = len(lines)

    # 1. list footer
    for i, ln in enumerate(lines):
        if FOOTER_START.search(ln):
            j = i
            # also drop a short separator line just above ("-----", "___")
            while j > 0 and re.fullmatch(r"\s*[-_=*#]{3,}\s*|\s*", lines[j - 1]):
                j -= 1
            cut = min(cut, j)
            break

    # 2. start of quoted original message
    for i in range(min(cut, len(lines))):
        ln = lines[i]
        if any(p.search(ln) for p in REPLY_HDR_PATTERNS):
            if "forward" in ln.lower() and i < 3:
                continue  # body *is* the forward; keep it
            cut = min(cut, i)
            break
        if (re.match(r"^\s*(On|Le|Am|El)\b", ln) and any(
                ON_WROTE.match(" ".join(lines[i:i + n])) for n in (1, 2, 3, 4))) or (
                i > 0 and any(WROTE_DATED.match(" ".join(lines[i:i + n])) for n in (1, 2))):
            cut = min(cut, i)
            break
        if re.fullmatch(r"\s*_{10,}\s*", ln) and any(
                OUTLOOK_FROM.match(x) for x in lines[i + 1:i + 3]):
            cut = min(cut, i)
            break
        if OUTLOOK_FROM.match(ln) and any(
                OUTLOOK_NEXT.match(x) for x in lines[i + 1:i + 5]) and i > 0:
            cut = min(cut, i)
            break
    if cut < len(lines):
        flags.add("quote_or_footer_removed")
    lines = lines[:cut]

    # 3. '>' quoted lines
    before = len(lines)
    lines = [ln for ln in lines if not ln.lstrip().startswith(">")]
    if len(lines) != before:
        flags.add("quoted_lines_removed")

    # 4. signatures: standard delimiter or mobile footer
    for i, ln in enumerate(lines):
        if ln in ("-- ", "--") or MOBILE_SIG.match(ln):
            lines = lines[:i]
            flags.add("signature_removed")
            break

    # 5. heuristic sign-off: keep sign-off + next non-blank line (name)
    nonblank = [i for i, ln in enumerate(lines) if ln.strip()]
    if len(nonblank) > 3:
        tail_start = nonblank[max(0, len(nonblank) - 15)]
        for i in range(tail_start, len(lines)):
            if SIGNOFF.match(lines[i]) and i > nonblank[0]:
                k = i + 1
                while k < len(lines) and not lines[k].strip():
                    k += 1
                end = k + 1 if k < len(lines) else k
                if end < len(lines) and any(x.strip() for x in lines[end:]):
                    lines = lines[:end]
                    flags.add("signature_trimmed_heuristic")
                break

    out = "\n".join(ln.rstrip() for ln in lines)
    out = re.sub(r"\n{3,}", "\n\n", out).strip()

    # 6. pure forwards / quote-only posts: nothing left of the author's own
    #    words, so keep the forwarded content (footer removed, '>' stripped)
    if not out and text.strip():
        m = FOOTER_START.search(text)
        body = text[:m.start()] if m else text
        body = "\n".join(re.sub(r"^\s*(>\s?)+", "", ln) for ln in body.split("\n"))
        out = re.sub(r"\n{3,}", "\n\n", body).strip()
        flags.add("clean_is_forwarded_content")
    return repair_mojibake(out, flags)


MOJIBAKE = re.compile("â€|Ã[\u0080-¿©¨¡³±§¼ºª]|Â[ -¿ ]")
MOJI_FIXED = {"â€œ": "“", "â€�": "”", "â€™": "’", "â€˜": "‘",
              "â€”": "—", "â€“": "–", "â€¦": "…", "â€¢": "•"}


def repair_mojibake(text, flags):
    """Undo UTF-8-read-as-Windows-1252 garbling present in some source posts."""
    if not MOJIBAKE.search(text):
        return text
    for bad, good in MOJI_FIXED.items():
        text = text.replace(bad, good)
    fixed_lines = []
    for ln in text.split("\n"):
        if MOJIBAKE.search(ln):
            try:
                ln = ln.encode("cp1252").decode("utf-8")
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass
        fixed_lines.append(ln)
    flags.add("mojibake_repaired")
    return "\n".join(fixed_lines)


URL_RE = re.compile(r"""(?i)\b(?:https?://|www\.)[^\s<>"'\]\)]+""")


def extract_links(text, html_links):
    found = []
    for u in URL_RE.findall(text) + html_links:
        u = u.rstrip(".,;:!?*>")
        if u.lower().startswith("www."):
            u = "http://" + u
        if "lists.apa.org" in u.lower() and "unsubscribe" in u.lower():
            continue
        if u not in found:
            found.append(u)
    return found


# --------------------------------------------------------------------------


def split_log(raw):
    """Return list of byte chunks, one per list message."""
    m = re.search(rb"\r?\n\r?\n", raw)
    body = raw[m.end():] if m else raw
    parts = SEP_RE.split(body)
    return [p for p in parts[1:]] , parts[0]


def parse_file(path):
    raw = path.read_bytes()
    log = re.search(r"LOG(\d{4})", path.name).group(1)
    chunks, preamble = split_log(raw)
    rows, problems = [], []
    if preamble.strip(b"=\r\n \t"):
        problems.append(f"{path.name}: {len(preamble)} bytes before first separator")
    for seq, chunk in enumerate(chunks, 1):
        key = f"LOG{log}-{seq:04d}"
        flags = set()
        try:
            msg = email.message_from_bytes(chunk, policy=policy.default)
        except Exception as e:  # extremely rare; keep the row
            rows.append({"msg_key": key, "source_file": path.name, "log_month": log,
                         "seq_in_file": seq, "flags": f"parse_error:{type(e).__name__}",
                         "body_full": chunk.decode("utf-8", "replace")})
            continue
        if msg.defects:
            flags.add("mime_defects")
        date_raw = hdr(msg, "Date", flags)
        date_iso = date_utc = ""
        try:
            dt = email.utils.parsedate_to_datetime(date_raw)
            date_iso = dt.isoformat()
            if dt.tzinfo is None:
                flags.add("date_no_timezone")
                dt = dt.replace(tzinfo=timezone.utc)
            date_utc = dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            yy, mm = int(log[:2]) + 2000, int(log[2:])
            months_off = (dt.year - yy) * 12 + (dt.month - mm)
            if months_off not in (0, -1, 1) or (months_off != 0 and
                                                 not (dt.day >= 25 or dt.day <= 5)):
                flags.add("date_outside_log_month")
        except Exception:
            flags.add("date_missing_or_unparseable")

        from_raw = hdr(msg, "From", flags)
        from_name, from_email = email.utils.parseaddr(from_raw)
        from_email = from_email.lower()
        if not from_email:
            flags.add("from_missing")
        subject = hdr(msg, "Subject", flags)
        if not subject:
            flags.add("subject_missing")
        message_id = norm_msgid(hdr(msg, "Message-ID", flags))
        if not message_id:
            flags.add("message_id_missing")
        in_reply_to = norm_msgid(hdr(msg, "In-Reply-To", flags))
        references = " ".join(all_msgids(hdr(msg, "References", flags)))
        recips = {h: hdr(msg, h, flags) for h in ("To", "Cc", "X-To", "X-cc")}

        body, fmt, atts, html_links = extract_body(msg, flags)
        clean = clean_body(body, flags)
        if body.strip() and not clean:
            flags.add("clean_body_empty")
        # links come from the author's own text only (not footers, quotes or
        # signatures); for HTML bodies the URLs were inlined during conversion
        links = extract_links(clean, [])

        rows.append({
            "msg_key": key, "source_file": path.name, "log_month": log,
            "seq_in_file": seq, "message_id": message_id,
            "date": date_iso, "date_utc": date_utc, "date_raw": date_raw,
            "from_name": from_name, "from_email": from_email,
            "reply_to": hdr(msg, "Reply-To", flags),
            "to": recips["To"], "cc": recips["Cc"],
            "x_to": recips["X-To"], "x_cc": recips["X-cc"],
            "subject": subject, "subject_norm": norm_subject(subject),
            "in_reply_to": in_reply_to, "references": references,
            "body_format": fmt, "body_full": body, "body_clean": clean,
            "body_full_chars": len(body), "body_clean_chars": len(clean),
            "links": "\n".join(links), "link_count": len(links),
            "attachment_count": len(atts),
            "attachment_names": "; ".join(a["filename"] or f"(unnamed {a['content_type']})"
                                          for a in atts),
            "_atts": atts, "flags": "",
            "_flags": flags,
        })
    return rows, problems


def _sig_norm(ln):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]*>", "", ln)).strip().lower()


def strip_sender_signatures(rows, min_posts=4, share=0.4, tail=15):
    """Learn each sender's recurring trailing lines and trim them.

    A line counts as signature for a sender if it appears in the last `tail`
    non-blank lines of at least `share` of that sender's posts. Trailing
    runs of such lines (and blank lines) are removed from body_clean.
    """
    by_sender = defaultdict(list)
    for r in rows:
        if r.get("from_email") and r.get("body_clean"):
            by_sender[r["from_email"]].append(r)
    for sender, posts in by_sender.items():
        if len(posts) < min_posts:
            continue
        counts = Counter()
        for r in posts:
            lines = [ln for ln in r["body_clean"].split("\n") if ln.strip()]
            counts.update({_sig_norm(ln) for ln in lines[-tail:]})
        sig = {k for k, v in counts.items()
               if v >= share * len(posts) and len(k) >= 2}
        if not sig:
            continue
        for r in posts:
            lines = r["body_clean"].split("\n")
            nb = [i for i, ln in enumerate(lines) if ln.strip()]
            end = len(lines)
            # earliest line in the tail where the rest of the message is
            # mostly (>=60%) this sender's known signature lines
            for pos, i in enumerate(nb[-tail:]):
                if _sig_norm(lines[i]) not in sig:
                    continue
                rest = nb[len(nb) - len(nb[-tail:]) + pos:]
                hits = sum(1 for j in rest if _sig_norm(lines[j]) in sig)
                if hits >= 0.6 * len(rest):
                    end = i
                    break
            # never strip the whole message
            if end < len(lines) and any(ln.strip() for ln in lines[:end]):
                r["body_clean"] = "\n".join(lines[:end]).rstrip()
                r["_flags"].add("signature_removed_sender_pattern")


def finalize_clean(rows):
    for r in rows:
        if "_atts" not in r:
            continue
        links = extract_links(r["body_clean"], [])
        r["links"], r["link_count"] = "\n".join(links), len(links)
        r["body_clean_chars"] = len(r["body_clean"])


def assign_threads(rows):
    by_id = {}
    for r in rows:
        if r.get("message_id") and r["message_id"] not in by_id:
            by_id[r["message_id"]] = r
    parent = {}
    method = {}
    for r in rows:
        k = r["msg_key"]
        cand = [r.get("in_reply_to")] + list(reversed((r.get("references") or "").split()))
        for c in cand:
            if c and c in by_id and by_id[c]["msg_key"] != k:
                parent[k] = by_id[c]["msg_key"]
                method[k] = "headers"
                break
    # subject fallback for replies with no resolvable parent: attach to the
    # most recent earlier message with the same normalised subject within 60 days
    by_subj = defaultdict(list)
    for r in sorted(rows, key=lambda r: r.get("date_utc") or ""):
        k = r["msg_key"]
        s = r.get("subject_norm") or ""
        is_reply = bool(re.match(r"^\s*(re|aw|sv)\s*:", r.get("subject") or "", re.I))
        if k not in parent and is_reply and len(s) >= 4 and by_subj[s]:
            prev = by_subj[s][-1]
            if (r.get("date_utc") and prev.get("date_utc") and
                    _days_between(prev["date_utc"], r["date_utc"]) <= 60):
                parent[k] = prev["msg_key"]
                method[k] = "subject"
        if s:
            by_subj[s].append(r)

    def root(k, seen=None):
        seen = seen or set()
        while k in parent and k not in seen:
            seen.add(k)
            k = parent[k]
        return k
    for r in rows:
        k = r["msg_key"]
        r["parent_key"] = parent.get(k, "")
        r["thread_key"] = root(k)
        r["thread_method"] = method.get(k, "root")
    sizes = Counter(r["thread_key"] for r in rows)
    for r in rows:
        r["thread_size"] = sizes[r["thread_key"]]


def _days_between(a, b):
    from datetime import datetime
    f = "%Y-%m-%dT%H:%M:%SZ"
    return abs((datetime.strptime(b, f) - datetime.strptime(a, f)).days)


def mark_duplicates(rows):
    first_by_id, first_by_fp = {}, {}
    for r in sorted(rows, key=lambda r: r["msg_key"]):
        r["duplicate_of"] = ""
        mid = r.get("message_id")
        if mid:
            if mid in first_by_id:
                r["duplicate_of"] = first_by_id[mid]
                r["_flags"].add("duplicate_message_id")
            else:
                first_by_id[mid] = r["msg_key"]
        fp = hashlib.sha1("|".join([
            r.get("from_email", ""), r.get("date_utc", ""), r.get("subject_norm", ""),
            re.sub(r"\s+", " ", r.get("body_clean", ""))[:2000]]).encode()).hexdigest()
        if fp in first_by_fp and not r["duplicate_of"]:
            r["duplicate_of"] = first_by_fp[fp]
            r["_flags"].add("suspected_duplicate_content")
        first_by_fp.setdefault(fp, r["msg_key"])


COLUMNS = ["msg_key", "source_file", "log_month", "seq_in_file", "message_id",
           "date", "date_utc", "date_raw", "from_name", "from_email", "reply_to",
           "to", "cc", "x_to", "x_cc", "subject", "subject_norm",
           "in_reply_to", "references", "parent_key", "thread_key",
           "thread_method", "thread_size", "body_format",
           "body_full_chars", "body_clean_chars", "body_clean", "body_full",
           "links", "link_count", "attachment_count", "attachment_names",
           "duplicate_of", "flags"]


def write_outputs(rows, out, web_images=False):
    out.mkdir(parents=True, exist_ok=True)
    dbp = out / "archive.sqlite"
    if dbp.exists():
        dbp.unlink()
    import shutil
    if (out / "attachments").exists():
        shutil.rmtree(out / "attachments")
    files = store_attachments(rows, out, web_images)
    db = sqlite3.connect(dbp)
    coldefs = ", ".join(f'"{c}" {"INTEGER" if c in ("seq_in_file","thread_size","body_full_chars","body_clean_chars","link_count","attachment_count") else "TEXT"}'
                        + (" PRIMARY KEY" if c == "msg_key" else "") for c in COLUMNS)
    db.execute(f"CREATE TABLE messages ({coldefs})")
    db.execute("CREATE TABLE attachments (msg_key TEXT, filename TEXT, content_type TEXT, "
               "size_bytes INTEGER, inline INTEGER, sha256 TEXT, stored_path TEXT)")
    db.execute("CREATE TABLE attachment_files (sha256 TEXT PRIMARY KEY, stored_path TEXT, "
               "content_type TEXT, size_bytes INTEGER, first_filename TEXT, first_msg_key TEXT, "
               "times_attached INTEGER, text_status TEXT, text_chars INTEGER, text TEXT)")
    db.executemany(f"INSERT INTO messages VALUES ({','.join('?'*len(COLUMNS))})",
                   [[r.get(c, "") for c in COLUMNS] for r in rows])
    db.executemany("INSERT INTO attachments VALUES (?,?,?,?,?,?,?)",
                   [(r["msg_key"], a["filename"], a["content_type"], a["size_bytes"],
                     int(a["inline"]), a.get("sha256", ""), a.get("stored_path", ""))
                    for r in rows for a in r.get("_atts", [])])
    db.executemany("INSERT INTO attachment_files VALUES (?,?,?,?,?,?,?,?,?,?)",
                   [(f["sha256"], f["stored_path"], f["content_type"], f["size_bytes"],
                     f["first_filename"], f["first_msg_key"], f["times_attached"],
                     f["text_status"], len(f["text"]), f["text"]) for f in files.values()])
    db.execute("CREATE INDEX idx_att_msg ON attachments(msg_key)")
    db.execute("CREATE INDEX idx_att_sha ON attachments(sha256)")
    db.execute("CREATE VIRTUAL TABLE attachment_fts USING fts5(first_filename, text, "
               "content='attachment_files', content_rowid='rowid', tokenize='porter unicode61')")
    db.execute("INSERT INTO attachment_fts(rowid, first_filename, text) "
               "SELECT rowid, first_filename, text FROM attachment_files")
    for c in ("date_utc", "from_email", "thread_key", "message_id", "log_month"):
        db.execute(f'CREATE INDEX idx_{c} ON messages("{c}")')
    db.execute("CREATE VIRTUAL TABLE messages_fts USING fts5(subject, from_name, body_clean, "
               "content='messages', content_rowid='rowid', tokenize='porter unicode61')")
    db.execute("INSERT INTO messages_fts(rowid, subject, from_name, body_clean) "
               "SELECT rowid, subject, from_name, body_clean FROM messages")
    db.commit()

    # review list: every saved non-image file and each message it came in
    by_key = {r["msg_key"]: r for r in rows}
    with open(out / "attachments_review.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["filename", "content_type", "size_kb", "times_attached", "text_status",
                    "first_sent_date", "first_sender", "first_subject", "msg_keys",
                    "stored_path", "text_preview"])
        for (sha, path, ctype, size, fname, first_key, times, status, text) in db.execute(
                "SELECT sha256, stored_path, content_type, size_bytes, first_filename, "
                "first_msg_key, times_attached, text_status, text FROM attachment_files "
                "WHERE content_type NOT LIKE 'image/%' ORDER BY first_msg_key"):
            keys = [k for (k,) in db.execute(
                "SELECT msg_key FROM attachments WHERE sha256=? ORDER BY msg_key", (sha,))]
            fr = by_key.get(first_key, {})
            w.writerow([fname, ctype, round(size / 1024), times, status, fr.get("date", ""),
                        fr.get("from_name", ""), fr.get("subject", ""), " ".join(keys), path,
                        re.sub(r"\s+", " ", text or "")[:300]])
    db.close()

    csv_cols = COLUMNS + ["body_truncated_in_csv"]
    with open(out / "messages.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        w.writerow(csv_cols)
        for r in rows:
            vals, trunc = [], []
            for c in COLUMNS:
                v = r.get(c, "")
                if isinstance(v, str) and len(v) > EXCEL_CELL_LIMIT:
                    v = v[:EXCEL_CELL_LIMIT] + " …[TRUNCATED IN CSV – full text in archive.sqlite]"
                    trunc.append(c)
                vals.append(v)
            vals.append(";".join(trunc))
            w.writerow(vals)


def main(raw_dir, out_dir, name="LIST", web_images=False):
    raw_dir, out_dir = Path(raw_dir), Path(out_dir)
    files = sorted(raw_dir.glob("*.eml"))
    rows, problems, per_file = [], [], {}
    for p in files:
        try:
            r, pr = parse_file(p)
        except Exception as e:
            problems.append(f"{p.name}: FILE FAILED {type(e).__name__}: {e}")
            continue
        rows.extend(r)
        problems.extend(pr)
        per_file[p.name] = len(r)
    strip_sender_signatures(rows)
    finalize_clean(rows)
    assign_threads(rows)
    mark_duplicates(rows)
    for r in rows:
        r["flags"] = ";".join(sorted(r.pop("_flags", set()) | (set(r["flags"].split(";")) - {""})))
    write_outputs(rows, out_dir, web_images)
    write_report(rows, files, per_file, problems, out_dir, name, web_images)


def write_report(rows, files, per_file, problems, out, name="LIST", web_images=False):
    fc = Counter()
    for r in rows:
        for f in filter(None, r["flags"].split(";")):
            fc[f.split(":")[0]] += 1
    att_types = Counter()
    for r in rows:
        for a in r.get("_atts", []):
            att_types[a["content_type"]] += 1
    threads = Counter(r["thread_key"] for r in rows)
    lines = [f"# {name} archive build report", "",
             f"- Files processed: {len(per_file)} of {len(files)}",
             f"- Messages extracted: {len(rows)}",
             f"- Duplicates (same Message-ID): {fc['duplicate_message_id']}",
             f"- Suspected duplicates (same sender/time/subject/body): {fc['suspected_duplicate_content']}",
             f"- Threads: {len(threads)} (multi-message: {sum(1 for v in threads.values() if v > 1)})",
             f"- Thread links by headers: {sum(1 for r in rows if r['thread_method']=='headers')}, by subject fallback: {sum(1 for r in rows if r['thread_method']=='subject')}",
             f"- Unique senders: {len({r['from_email'] for r in rows if r.get('from_email')})}",
             f"- Messages with links: {sum(1 for r in rows if r.get('link_count'))}",
             "", "## File-level problems", ""]
    lines += [f"- {p}" for p in problems] or ["- none"]
    lines += ["", "## Flags", "", "| flag | messages |", "|---|---|"]
    lines += [f"| {k} | {v} |" for k, v in fc.most_common()]
    lines += ["", "## Messages per file", "", "| file | messages |", "|---|---|"]
    lines += [f"| {k} | {v} |" for k, v in per_file.items()]
    lines += ["", "## Attachment types", "", "| type | count |", "|---|---|"]
    lines += [f"| {k} | {v} |" for k, v in att_types.most_common()]
    db = sqlite3.connect(out / "archive.sqlite")
    lines += ["", "## Saved attachment files (deduplicated)", "",
              "| content type | text status | files | MB |", "|---|---|---|---|"]
    lines += [f"| {c} | {s} | {n} | {mb:.1f} |" for c, s, n, mb in db.execute(
        "SELECT content_type, text_status, count(*), sum(size_bytes)/1048576.0 "
        "FROM attachment_files GROUP BY 1,2 ORDER BY 3 DESC")]
    if web_images:
        lines += ["", "Images over 200 KB were stored as smaller web copies (max "
                  f"{WEB_MAX_PX} px; opaque PNG converted to JPEG). Full-quality originals "
                  "remain in the .eml files."]
    db.close()
    (out / "report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    args = sys.argv[1:]
    web = "--web-images" in args
    args = [a for a in args if a != "--web-images"]
    name = "LIST"
    if "--name" in args:
        i = args.index("--name")
        if i + 1 >= len(args):
            sys.exit(__doc__)
        name = args[i + 1]
        del args[i:i + 2]
    if len(args) != 2:
        sys.exit(__doc__)
    main(args[0], args[1], name, web)
