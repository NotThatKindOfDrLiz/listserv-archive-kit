#!/usr/bin/env python3
"""Local prototype of a LISTSERV archive website.

Serves ONLY on 127.0.0.1 (this computer). Reads the archives read-only.
Standard library only.

Usage:  python3 -I tools/prototype/server.py output/archive.sqlite [port] [announcements.sqlite]
Then open http://127.0.0.1:8013/

The discussion list and the announcements list are separate databases and
are never mixed: /api/... reads only the discussion archive, /api/ann/...
reads only the announcements archive.
"""
import json
import mimetypes
import re
import sqlite3
import sys
from collections import OrderedDict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

HERE = Path(__file__).resolve().parent
DB_PATH = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else None
PORT = int(sys.argv[2]) if len(sys.argv) > 2 else 8013
ANN_PATH = Path(sys.argv[3]).resolve() if len(sys.argv) > 3 else None
SOURCES = {"disc": DB_PATH}
if ANN_PATH:
    SOURCES["ann"] = ANN_PATH

# --------------------------------------------------------------------------
# privacy: what members see

EMAIL_RE = re.compile(r"(mailto:)?[\w.+'-]+@[\w-]+(\.[\w-]+)+")
PHONE_RE = re.compile(r"(?<![\w/=.-])(\+?1[\s.-]?)?\(?\d{3}\)?[\s.-]\d{3}[\s.-]\d{4}(?![\w/-])")


def display_name(name):
    """Sender name for display; never an email address."""
    name = (name or "").strip().strip('"')
    if not name or "@" in name:
        return "Member (name not provided)"
    return name


def redact(text):
    if not text:
        return text
    text = EMAIL_RE.sub("[email hidden]", text)
    return PHONE_RE.sub("[phone hidden]", text)


SUBJ_PREFIX = re.compile(r"^\s*((re|fw|fwd|aw|sv)\s*(\[\d+\])?\s*:|\[[^\]]*\])\s*", re.I)


def tidy_subject(s):
    s = s or ""
    while True:
        n = SUBJ_PREFIX.sub("", s, count=1)
        if n == s:
            break
        s = n
    return s.strip() or "(no subject)"


# --------------------------------------------------------------------------
# search

STOP = set("""a an and are as at be by can do does for from has have how i in
is it me my of on or our that the their them they this to was we what when
where which who why will with about any anyone messages message discussion
discussions people""".split())


def fts_query(q, mode):
    phrases = re.findall(r'"([^"]+)"', q or "")
    rest = re.sub(r'"[^"]*"', " ", q or "")
    words = [w for w in re.findall(r"[\w'’-]+", rest) if w.lower() not in STOP]
    terms = [t.replace('"', "").strip() for t in phrases + words]
    terms = [t for t in terms if t]
    if not terms:
        return None, []
    return (" AND " if mode == "AND" else " OR ").join(f'"{t}"' for t in terms), terms


def db(src="disc"):
    con = sqlite3.connect(f"file:{SOURCES[src]}?mode=ro", uri=True)
    con.row_factory = sqlite3.Row
    return con


def thread_meta(con, keys):
    if not keys:
        return {}
    qm = ",".join("?" * len(keys))
    meta = {}
    for r in con.execute(
            f"SELECT thread_key, count(*) n, min(date_utc) first, max(date_utc) last "
            f"FROM messages WHERE thread_key IN ({qm}) GROUP BY thread_key", keys):
        meta[r["thread_key"]] = {"thread_key": r["thread_key"], "count": r["n"],
                                 "first": r["first"], "last": r["last"]}
    for r in con.execute(f"SELECT msg_key, subject FROM messages WHERE msg_key IN ({qm})", keys):
        meta[r["msg_key"]]["subject"] = tidy_subject(r["subject"])
    for k in keys:
        names = [display_name(r[0]) for r in con.execute(
            "SELECT from_name FROM messages WHERE thread_key=? AND from_name<>'' "
            "GROUP BY from_name ORDER BY min(date_utc)", (k,))]
        meta[k]["participants"] = list(dict.fromkeys(names))[:6]
        meta[k]["participant_count"] = len(set(names))
    return meta


def date_filters(qs):
    where, params = [], []
    if qs.get("from"):
        where.append("m.date_utc >= ?")
        params.append(qs["from"] + "-01-01" if len(qs["from"]) == 4 else qs["from"])
    if qs.get("to"):
        where.append("m.date_utc < ?")
        params.append(str(int(qs["to"]) + 1) + "-01-01" if len(qs["to"]) == 4 else qs["to"])
    return where, params


def api_search(qs, src="disc"):
    q = qs.get("q", "")
    sort = qs.get("sort", "relevance")
    page = max(1, int(qs.get("page", "1") or 1))
    con = db(src)
    where, params = date_filters(qs)
    extra = (" AND " + " AND ".join(where)) if where else ""
    rows, terms, mode_used = [], [], "AND"
    for mode in ("AND", "OR"):
        mq, terms = fts_query(q, mode)
        if not mq:
            return {"query": q, "terms": [], "threads": [], "total_threads": 0,
                    "total_messages": 0, "documents": []}
        rows = con.execute(
            "SELECT m.msg_key, m.thread_key, m.date_utc, m.from_name, m.body_clean, "
            "bm25(messages_fts, 6.0, 1.0, 1.0) AS score, "
            "snippet(messages_fts, 2, char(1), char(2), ' … ', 28) AS snip "
            "FROM messages_fts JOIN messages m ON m.rowid = messages_fts.rowid "
            f"WHERE messages_fts MATCH ?{extra} ORDER BY score LIMIT 3000",
            [mq] + params).fetchall()
        mode_used = mode
        if rows:
            break
    threads = OrderedDict()
    for r in rows:
        t = threads.get(r["thread_key"])
        if t is None:
            snip = r["snip"] if "\x01" in (r["snip"] or "") else (r["body_clean"] or "")[:220]
            threads[r["thread_key"]] = {"score": r["score"], "hits": 1,
                                        "snippet": redact(snip), "snippet_from": display_name(r["from_name"]),
                                        "snippet_msg": r["msg_key"]}
        else:
            t["hits"] += 1
            if "\x01" not in t["snippet"] and "\x01" in (r["snip"] or ""):
                t.update(snippet=redact(r["snip"]), snippet_from=display_name(r["from_name"]),
                         snippet_msg=r["msg_key"])
    keys = list(threads)
    meta_all = thread_meta(con, keys) if sort != "relevance" else {}
    if sort == "newest":
        keys.sort(key=lambda k: meta_all[k]["last"] or "", reverse=True)
    elif sort == "oldest":
        keys.sort(key=lambda k: meta_all[k]["first"] or "")
    size = 20
    page_keys = keys[(page - 1) * size: page * size]
    meta = meta_all or thread_meta(con, page_keys)
    out = []
    for k in page_keys:
        d = dict(meta[k])
        d.update({x: threads[k][x] for x in ("hits", "snippet", "snippet_from", "snippet_msg")})
        out.append(d)

    # documents
    dq, _ = fts_query(q, mode_used)
    docs = []
    for r in con.execute(
            "SELECT f.sha256, f.first_filename, f.content_type, f.size_bytes, f.times_attached, "
            "f.first_msg_key, m.thread_key, m.subject, m.date_utc, m.from_name, "
            "snippet(attachment_fts, 1, char(1), char(2), ' … ', 28) AS snip "
            "FROM attachment_fts JOIN attachment_files f ON f.rowid = attachment_fts.rowid "
            "JOIN messages m ON m.msg_key = f.first_msg_key "
            f"WHERE attachment_fts MATCH ?{extra} ORDER BY bm25(attachment_fts) LIMIT 50",
            [dq] + params):
        docs.append({"sha": r["sha256"], "filename": r["first_filename"],
                     "content_type": r["content_type"], "size": r["size_bytes"],
                     "thread_key": r["thread_key"], "subject": tidy_subject(r["subject"]),
                     "date": r["date_utc"], "from_name": display_name(r["from_name"]),
                     "snippet": redact(r["snip"])})
    return {"query": q, "terms": terms, "mode": mode_used, "threads": out,
            "total_threads": len(keys), "total_messages": len(rows),
            "page": page, "pages": (len(keys) + size - 1) // size, "documents": docs}


def api_thread(key, src="disc"):
    con = db(src)
    msgs = con.execute(
        "SELECT msg_key, date, date_utc, from_name, subject, body_clean, body_full, links "
        "FROM messages WHERE thread_key=? ORDER BY date_utc, msg_key", (key,)).fetchall()
    if not msgs:
        return None
    out = []
    for m in msgs:
        atts = [{"sha": a["sha256"], "filename": a["filename"] or "(unnamed)",
                 "content_type": a["content_type"], "size": a["size_bytes"]}
                for a in con.execute(
                    "SELECT DISTINCT sha256, filename, content_type, size_bytes FROM attachments "
                    "WHERE msg_key=? AND stored_path<>''", (m["msg_key"],))]
        links = [u for u in (m["links"] or "").split("\n")
                 if u and not EMAIL_RE.search(u)]
        out.append({"msg_key": m["msg_key"], "date": m["date"], "date_utc": m["date_utc"],
                    "from_name": display_name(m["from_name"]),
                    "subject": m["subject"], "body": redact(m["body_clean"]),
                    "original": redact(m["body_full"]), "links": links,
                    "attachments": atts})
    meta = thread_meta(con, [key])[key]
    meta["messages"] = out
    return meta


def api_stats(src="disc"):
    con = db(src)
    one = lambda sql: con.execute(sql).fetchone()[0]
    years = [{"year": r[0], "messages": r[1], "threads": r[2]} for r in con.execute(
        "SELECT substr(date_utc,1,4) y, count(*), count(DISTINCT thread_key) FROM messages "
        "GROUP BY y ORDER BY y")]
    recent = [r[0] for r in con.execute(
        "SELECT thread_key FROM messages GROUP BY thread_key ORDER BY max(date_utc) DESC "
        f"LIMIT {10 if src == 'ann' else 6}")]
    active = [r[0] for r in con.execute(
        "SELECT thread_key FROM messages GROUP BY thread_key ORDER BY count(*) DESC LIMIT 6")]
    meta = thread_meta(con, list(dict.fromkeys(recent + active)))
    return {"messages": one("SELECT count(*) FROM messages"),
            "threads": one("SELECT count(DISTINCT thread_key) FROM messages"),
            "contributors": one("SELECT count(DISTINCT from_email) FROM messages WHERE from_email<>''"),
            "documents": one("SELECT count(*) FROM attachment_files WHERE content_type NOT LIKE 'image/%'"),
            "first": one("SELECT min(date_utc) FROM messages"),
            "last": one("SELECT max(date_utc) FROM messages"),
            "years": years,
            "recent": [meta[k] for k in recent], "active": [meta[k] for k in active]}


def api_browse(qs, src="disc"):
    con = db(src)
    year = qs.get("year", "")
    if not re.fullmatch(r"\d{4}", year):
        return {"error": "year required"}
    months = [{"month": r[0], "messages": r[1], "threads": r[2]} for r in con.execute(
        "SELECT substr(date_utc,1,7), count(*), count(DISTINCT thread_key) FROM messages "
        "WHERE substr(date_utc,1,4)=? GROUP BY 1 ORDER BY 1", (year,))]
    month = qs.get("month", "")
    threads = []
    if re.fullmatch(r"\d{4}-\d{2}", month):
        keys = [r[0] for r in con.execute(
            "SELECT thread_key FROM messages GROUP BY thread_key "
            "HAVING substr(min(date_utc),1,7)=? ORDER BY min(date_utc)", (month,))]
        meta = thread_meta(con, keys)
        threads = [meta[k] for k in keys]
    return {"year": year, "months": months, "month": month, "threads": threads}


# --------------------------------------------------------------------------
# HTTP


class Handler(BaseHTTPRequestHandler):
    server_version = "ListservArchivePrototype"

    def log_message(self, fmt, *args):
        pass

    def _send(self, code, body, ctype, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Security-Policy",
                         "default-src 'self'; img-src 'self' data:; style-src 'self' 'unsafe-inline'; "
                         "script-src 'self' 'unsafe-inline'; frame-ancestors 'none'")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj).encode(), "application/json; charset=utf-8")

    def do_GET(self):
        u = urlparse(self.path)
        qs = {k: v[0] for k, v in parse_qs(u.query).items()}
        try:
            if u.path in ("/", "/index.html"):
                return self._send(200, (HERE / "index.html").read_bytes(), "text/html; charset=utf-8")
            m = re.fullmatch(r"/api/(?:(ann)/)?(stats|search|browse|thread/(LOG\d{4}-\d{4}))", u.path)
            if m:
                src = m.group(1) or "disc"
                if src not in SOURCES:
                    return self._json({"error": "announcements archive not loaded"}, 404)
                if m.group(2) == "stats":
                    return self._json(api_stats(src))
                if m.group(2) == "search":
                    return self._json(api_search(qs, src))
                if m.group(2) == "browse":
                    return self._json(api_browse(qs, src))
                t = api_thread(m.group(3), src)
                return self._json(t) if t else self._json({"error": "not found"}, 404)
            m = re.fullmatch(r"/(?:(ann)/)?files/([0-9a-f]{64})", u.path)
            if m and (m.group(1) or "disc") in SOURCES:
                return self._file(m.group(2), qs.get("view") == "1", m.group(1) or "disc")
        except Exception as e:  # prototype: report, don't crash
            return self._json({"error": f"{type(e).__name__}: {e}"}, 500)
        self._json({"error": "not found"}, 404)

    def _file(self, sha, view, src="disc"):
        out_dir = SOURCES[src].parent
        r = db(src).execute("SELECT stored_path, content_type, first_filename FROM attachment_files "
                         "WHERE sha256=?", (sha,)).fetchone()
        if not r:
            return self._json({"error": "not found"}, 404)
        path = (out_dir / r["stored_path"]).resolve()
        if out_dir not in path.parents or not path.is_file():
            return self._json({"error": "not found"}, 404)
        ctype = r["content_type"]
        inline_ok = view and ctype in ("image/png", "image/jpeg", "image/jpg", "image/gif")
        name = re.sub(r'[^\w .()-]', "_", r["first_filename"] or path.name)[:120] or path.name
        if path.suffix and Path(name).suffix.lower() != path.suffix:
            name = Path(name).stem + path.suffix  # web copy, e.g. flyer.png -> flyer.jpg
        disp = "inline" if inline_ok else f'attachment; filename="{name}"'
        if not inline_ok:
            ctype = "application/octet-stream"
        self._send(200, path.read_bytes(), ctype, {"Content-Disposition": disp})


def main():
    if not all(p.is_file() for p in SOURCES.values() if p) or not DB_PATH:
        sys.exit(__doc__)
    srv = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print(f"Archive prototype running at http://127.0.0.1:{PORT}/  (Ctrl+C to stop)"
          + ("  [with announcements]" if ANN_PATH else ""), flush=True)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
