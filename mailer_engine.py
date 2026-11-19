# -*- coding: utf-8 -*-
"""
ArubaMailer - motore di invio.
Invio a blocchi, una mail per destinatario, con i limiti Aruba:
  - max 150 destinatari per singolo invio (qui sempre 1)
  - circa 250 email ogni 20 minuti (indicazione del supporto Aruba)
  - max 8000 destinatari al giorno (policy Aruba)
Stato salvato in SQLite: ogni indirizzo ha il suo esito, si riprende dopo interruzioni.
"""
import csv
import datetime as dt
import email
import email.utils
import imaplib
import os
import re
import smtplib
import socket
import sqlite3
import ssl
import sys
import threading
from email.message import EmailMessage

# ---------------------------------------------------------------- limiti
MAX_BATCH = 250                 # email per blocco (massimo)
MIN_INTERVAL_S = 4.0            # secondi minimi tra una email e la successiva
ARUBA_PER_20MIN = 250           # indicazione supporto Aruba
HARD_DAILY_CAP = 8000           # policy Aruba
DEFAULT_DAILY_CAP = 2000
MAX_ATTEMPTS = 3                # tentativi per errori temporanei
BLOCK_PAUSE_MIN = 60            # pausa se Aruba segnala blocco/policy
MAX_BLOCKS = 2                  # dopo questi blocchi la campagna si ferma
BATCH_GAP_MIN = 1.0             # pausa minima tra fine blocco e blocco successivo

EMAIL_RE = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?"
                      r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)*\.[A-Za-z]{2,}$")
FIND_EMAIL_RE = re.compile(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def min_batch_period_min(batch_size, interval_s):
    """Minuti minimi tra l'inizio di un blocco e il successivo."""
    by_rate = 20.0 * batch_size / ARUBA_PER_20MIN
    by_duration = batch_size * interval_s / 60.0 + BATCH_GAP_MIN
    return max(by_rate, by_duration)


# ---------------------------------------------------------------- lista
def _col_index(col):
    col = col.strip().upper()
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def load_addresses(path, column=""):
    """Legge indirizzi da .xlsx/.xlsm/.csv/.txt.
    column: vuoto = cerca in tutte le celle; lettera (es. 'B') o intestazione (es. 'Email').
    Ritorna (validi_in_ordine, scartati[(riga, valore)], duplicati)."""
    rows = []
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        for r in ws.iter_rows(values_only=True):
            rows.append(["" if v is None else str(v) for v in r])
        wb.close()
    elif ext == ".xls":
        raise ValueError("Formato .xls vecchio: aprilo in Excel e salvalo come .xlsx")
    else:
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            sample = f.read(4096)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=";,\t")
            except csv.Error:
                dialect = csv.excel
            rows = [list(r) for r in csv.reader(f, dialect)]

    col_idx = None
    start_row = 0
    column = (column or "").strip()
    if column:
        if re.fullmatch(r"[A-Za-z]{1,3}", column) and not (rows and column.lower() in [c.strip().lower() for c in rows[0]]):
            col_idx = _col_index(column)
        else:
            header = [c.strip().lower() for c in rows[0]] if rows else []
            if column.lower() not in header:
                raise ValueError("Colonna '%s' non trovata nella prima riga" % column)
            col_idx = header.index(column.lower())
            start_row = 1

    valid, invalid, seen, dups = [], [], set(), 0
    for i, r in enumerate(rows[start_row:], start=start_row + 1):
        cells = [r[col_idx]] if col_idx is not None and col_idx < len(r) else ([] if col_idx is not None else r)
        for raw in cells:
            v = raw.strip().strip("<>").replace("mailto:", "").strip()
            if not v:
                continue
            if col_idx is None and "@" not in v:
                continue  # in modalita automatica ignoro le celle che non sono email
            # una cella puo contenerne piu di una separate da ; o ,
            parts = [p.strip() for p in re.split(r"[;,\s]+", v) if p.strip()]
            for p in parts:
                if EMAIL_RE.match(p) and ".." not in p:
                    key = p.lower()
                    if key in seen:
                        dups += 1
                    else:
                        seen.add(key)
                        valid.append(key)
                else:
                    if col_idx is not None and i == 1 and "@" not in p:
                        continue  # probabilmente intestazione
                    invalid.append((i, p))
    return valid, invalid, dups


# ---------------------------------------------------------------- campagna (SQLite)
class Campaign:
    STATUSES = ("pending", "sent", "temp_error", "failed", "bounced")

    def __init__(self, db_path):
        self.db_path = db_path
        self.lock = threading.Lock()
        self.conn = sqlite3.connect(db_path, check_same_thread=False, isolation_level=None)
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("""CREATE TABLE IF NOT EXISTS recipients(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            status TEXT NOT NULL DEFAULT 'pending',
            attempts INTEGER NOT NULL DEFAULT 0,
            last_try TEXT, sent_at TEXT, response TEXT)""")
        self.conn.execute("CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS log(ts TEXT, msg TEXT)")

    def close(self):
        with self.lock:
            self.conn.close()

    # meta
    def set_meta(self, k, v):
        with self.lock:
            self.conn.execute("INSERT OR REPLACE INTO meta VALUES(?,?)", (k, str(v)))

    def get_meta(self, k, default=None):
        with self.lock:
            r = self.conn.execute("SELECT value FROM meta WHERE key=?", (k,)).fetchone()
        return r[0] if r else default

    def add_log(self, msg):
        with self.lock:
            self.conn.execute("INSERT INTO log VALUES(?,?)", (now_str(), msg))

    def import_addresses(self, addrs):
        with self.lock:
            cur = self.conn.cursor()
            cur.execute("BEGIN")
            added = 0
            for a in addrs:
                cur.execute("INSERT OR IGNORE INTO recipients(email) VALUES(?)", (a,))
                added += cur.rowcount
            cur.execute("COMMIT")
        return added

    def counts(self):
        with self.lock:
            rows = self.conn.execute("SELECT status, COUNT(*) FROM recipients GROUP BY status").fetchall()
        c = {s: 0 for s in self.STATUSES}
        c.update(dict(rows))
        c["total"] = sum(c[s] for s in self.STATUSES)
        return c

    def remaining(self):
        with self.lock:
            return self.conn.execute(
                "SELECT COUNT(*) FROM recipients WHERE status='pending' OR (status='temp_error' AND attempts<?)",
                (MAX_ATTEMPTS,)).fetchone()[0]

    def next_batch(self, n):
        with self.lock:
            rows = self.conn.execute(
                "SELECT email FROM recipients WHERE status='pending' "
                "OR (status='temp_error' AND attempts<?) "
                "ORDER BY CASE status WHEN 'pending' THEN 0 ELSE 1 END, id LIMIT ?",
                (MAX_ATTEMPTS, n)).fetchall()
        return [r[0] for r in rows]

    def mark(self, addr, status, response, count_attempt=True):
        ts = now_str()
        with self.lock:
            self.conn.execute(
                "UPDATE recipients SET status=?, response=?, last_try=?, "
                "attempts=attempts+?, sent_at=CASE WHEN ?='sent' THEN ? ELSE sent_at END WHERE email=?",
                (status, (response or "")[:500], ts, 1 if count_attempt else 0, status, ts, addr))

    def mark_bounced(self, addr, info):
        with self.lock:
            cur = self.conn.execute(
                "UPDATE recipients SET status='bounced', response=? WHERE email=? AND status IN ('sent','temp_error')",
                (("RIMBALZO: " + info)[:500], addr.lower()))
            return cur.rowcount

    def all_emails(self):
        with self.lock:
            return set(r[0] for r in self.conn.execute("SELECT email FROM recipients"))

    def sent_since(self, since):
        with self.lock:
            return self.conn.execute("SELECT COUNT(*) FROM recipients WHERE sent_at>=?",
                                     (since.strftime("%Y-%m-%d %H:%M:%S"),)).fetchone()[0]

    def oldest_sent_since(self, since):
        with self.lock:
            r = self.conn.execute("SELECT MIN(sent_at) FROM recipients WHERE sent_at>=?",
                                  (since.strftime("%Y-%m-%d %H:%M:%S"),)).fetchone()
        return parse_ts(r[0]) if r and r[0] else None

    def reset_failed_to_pending(self):
        with self.lock:
            cur = self.conn.execute("UPDATE recipients SET status='pending', attempts=0 WHERE status='temp_error'")
            return cur.rowcount

    def rows(self):
        with self.lock:
            return self.conn.execute(
                "SELECT email,status,attempts,sent_at,last_try,response FROM recipients ORDER BY id").fetchall()

    def export_xlsx(self, path):
        import openpyxl
        from openpyxl.styles import Font
        labels = {"pending": "Da inviare", "sent": "Inviata (accettata da Aruba)",
                  "temp_error": "Errore temporaneo", "failed": "Rifiutata", "bounced": "Rimbalzata"}
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Esiti"
        ws.append(["Email", "Esito", "Tentativi", "Inviata il", "Ultimo tentativo", "Risposta server"])
        for c in ws[1]:
            c.font = Font(bold=True)
        for e, s, a, sent, last, resp in self.rows():
            ws.append([e, labels.get(s, s), a, sent or "", last or "", resp or ""])
        for col, w in zip("ABCDEF", (38, 28, 10, 20, 20, 70)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"
        ws2 = wb.create_sheet("Riepilogo")
        c = self.counts()
        for k in ("total", "sent", "pending", "temp_error", "failed", "bounced"):
            ws2.append([labels.get(k, "Totale"), c[k]])
        ws2.column_dimensions["A"].width = 32
        wb.save(path)


def now_str():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_ts(s):
    return dt.datetime.strptime(s, "%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------- utilita Windows
def keep_awake(on):
    """Impedisce a Windows di andare in sospensione durante l'invio."""
    if os.name != "nt":
        return
    try:
        import ctypes
        ES_CONTINUOUS, ES_SYSTEM_REQUIRED = 0x80000000, 0x00000001
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | (ES_SYSTEM_REQUIRED if on else 0))
    except Exception:
        pass


def protect_password(pw):
    """Cifra la password con DPAPI di Windows (leggibile solo dal tuo utente)."""
    if os.name != "nt" or not pw:
        return ""
    import ctypes
    import ctypes.wintypes as wt
    import base64

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", wt.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]
    data = pw.encode("utf-8")
    inp = BLOB(len(data), ctypes.cast(ctypes.create_string_buffer(data, len(data)), ctypes.POINTER(ctypes.c_char)))
    out = BLOB()
    if not ctypes.windll.crypt32.CryptProtectData(ctypes.byref(inp), None, None, None, None, 0, ctypes.byref(out)):
        return ""
    enc = ctypes.string_at(out.pbData, out.cbData)
    ctypes.windll.kernel32.LocalFree(out.pbData)
    return base64.b64encode(enc).decode()


def unprotect_password(s):
    if os.name != "nt" or not s:
        return ""
    try:
        import ctypes
        import ctypes.wintypes as wt
        import base64

        class BLOB(ctypes.Structure):
            _fields_ = [("cbData", wt.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]
        data = base64.b64decode(s)
        inp = BLOB(len(data), ctypes.cast(ctypes.create_string_buffer(data, len(data)), ctypes.POINTER(ctypes.c_char)))
        out = BLOB()
        if not ctypes.windll.crypt32.CryptUnprotectData(ctypes.byref(inp), None, None, None, None, 0, ctypes.byref(out)):
            return ""
        dec = ctypes.string_at(out.pbData, out.cbData)
        ctypes.windll.kernel32.LocalFree(out.pbData)
        return dec.decode("utf-8")
    except Exception:
        return ""


# ---------------------------------------------------------------- SMTP
class Account:
    def __init__(self, sender_email, password, sender_name="", reply_to="",
                 smtp_host="smtps.aruba.it", smtp_port=465,
                 imap_host="imaps.aruba.it", imap_port=993, use_ssl=True):
        self.sender_email = sender_email.strip()
        self.password = password
        self.sender_name = sender_name.strip()
        self.reply_to = reply_to.strip()
        self.smtp_host = smtp_host.strip()
        self.smtp_port = int(smtp_port)
        self.imap_host = imap_host.strip()
        self.imap_port = int(imap_port)
        self.use_ssl = use_ssl

    def connect_smtp(self, timeout=60):
        ctx = ssl.create_default_context()
        if self.use_ssl:
            s = smtplib.SMTP_SSL(self.smtp_host, self.smtp_port, timeout=timeout, context=ctx)
        else:
            s = smtplib.SMTP(self.smtp_host, self.smtp_port, timeout=timeout)
            s.ehlo()
            if s.has_extn("starttls"):
                s.starttls(context=ctx)
                s.ehlo()
        if self.password:
            s.login(self.sender_email, self.password)
        return s

    def build_message(self, to_addr, content):
        msg = EmailMessage()
        msg["From"] = email.utils.formataddr((self.sender_name, self.sender_email)) if self.sender_name else self.sender_email
        msg["To"] = to_addr
        msg["Subject"] = content.subject
        msg["Date"] = email.utils.formatdate(localtime=True)
        domain = self.sender_email.split("@")[-1] if "@" in self.sender_email else None
        msg["Message-ID"] = email.utils.make_msgid(domain=domain)
        if self.reply_to:
            msg["Reply-To"] = self.reply_to
        msg.set_content(content.text, charset="utf-8")
        if content.html:
            msg.add_alternative(content.html, subtype="html", charset="utf-8")
            if content.images:
                html_part = msg.get_payload()[-1]
                for cid, data, sub in content.images:
                    html_part.add_related(data, maintype="image", subtype=sub, cid="<%s>" % cid,
                                          filename="%s.%s" % (cid, "jpg" if sub == "jpeg" else sub),
                                          disposition="inline")
        return msg


class Content:
    """Contenuto della mail: testo semplice sempre presente, HTML facoltativo (docx e/o logo)."""

    def __init__(self, subject, text, html=None, images=None):
        self.subject = subject
        self.text = text
        self.html = html
        self.images = images or []      # [(cid, bytes, subtype)]

    def size_estimate(self):
        return len(self.text.encode()) + len((self.html or "").encode()) + \
            int(sum(len(b) for _, b, _ in self.images) * 1.37)

    def preview_html(self):
        """HTML con le immagini incorporate, per l'anteprima nel browser."""
        import base64
        h = self.html or ("<pre style='font-family:Consolas,monospace;white-space:pre-wrap'>%s</pre>"
                          % self.text.replace("&", "&amp;").replace("<", "&lt;"))
        for cid, data, sub in self.images:
            h = h.replace("cid:%s" % cid, "data:image/%s;base64,%s" % (sub, base64.b64encode(data).decode()))
        return h


def build_content(subject, text="", docx_path="", logo_path=""):
    """Crea il Content. Se c'e' un .docx, testo e HTML vengono dal documento."""
    import docx_html as dh
    subject = (subject or "").strip()
    images = []
    html = None
    font, size = dh.FONT_STACK, 11
    if docx_path:
        if not os.path.exists(docx_path):
            raise ValueError("File Word non trovato: %s" % docx_path)
        body_html, text, font, size = dh.DocxConverter(docx_path).convert_with_images(images)
    else:
        body_html = dh.text_to_html(text) if logo_path else None
    if not text.strip() and not images:
        raise ValueError("Il messaggio è vuoto")
    logo = None
    if logo_path:
        data, sub, w, h = dh.load_logo(logo_path)
        images.insert(0, ("logo", data, sub))
        logo = ("logo", w, h)
    if body_html is not None:
        html = dh.wrap_html(body_html, font, size, logo)
    return Content(subject, text, html, images)


class BlockedError(Exception):
    """Aruba ha segnalato limite/policy: serve una pausa lunga."""


class FatalError(Exception):
    """Errore che richiede l'intervento dell'utente (password, mittente...)."""


def _err_text(code, msg):
    if isinstance(msg, bytes):
        msg = msg.decode("utf-8", "replace")
    return "%s %s" % (code, msg)


def _is_policy(text):
    t = text.lower()
    return any(k in t for k in ("policy", "rate", "too many", "limit", "exceeded", "blocked", "spam"))


def send_one(conn, account, addr, content):
    """Invia a un destinatario. Ritorna (status, risposta).
    Solleva BlockedError / FatalError / errori di connessione."""
    msg = account.build_message(addr, content)
    try:
        refused = conn.send_message(msg, from_addr=account.sender_email, to_addrs=[addr])
    except smtplib.SMTPRecipientsRefused as e:
        code, m = e.recipients.get(addr, (550, b"rifiutato"))
        text = _err_text(code, m)
        if code == 421 or (_is_policy(text) and "user" not in text.lower() and "mailbox" not in text.lower()):
            raise BlockedError(text)
        return ("failed" if 500 <= code < 600 else "temp_error"), text
    except smtplib.SMTPSenderRefused as e:
        text = _err_text(e.smtp_code, e.smtp_error)
        if e.smtp_code == 421 or 400 <= e.smtp_code < 500 or _is_policy(text):
            raise BlockedError(text)
        raise FatalError("Mittente rifiutato: " + text)
    except smtplib.SMTPDataError as e:
        text = _err_text(e.smtp_code, e.smtp_error)
        if 400 <= e.smtp_code < 500 and not _is_policy(text):
            return "temp_error", text
        raise BlockedError(text)
    if refused:
        code, m = refused.get(addr, (550, b"rifiutato"))
        return ("failed" if 500 <= code < 600 else "temp_error"), _err_text(code, m)
    return "sent", "250 OK accettata dal server"


# ---------------------------------------------------------------- invio a blocchi
class SendJob(threading.Thread):
    """
    settings: dict con
      content (Content), subject, body, docx, logo (per memoria), batch_size, interval_s, period_min, daily_cap,
      start_at (datetime|None), window_on, window_from (HH:MM), window_to (HH:MM)
    callback(kind, data): 'log' str, 'progress' dict counts, 'status' str, 'next' datetime|None, 'done' str
    """

    def __init__(self, campaign, account, settings, callback):
        super().__init__(daemon=True)
        self.c = campaign
        self.a = account
        self.s = dict(settings)
        self.cb = callback
        self.stop_event = threading.Event()
        self.blocks = 0
        self._validate()

    def _validate(self):
        s = self.s
        s["batch_size"] = max(1, min(MAX_BATCH, int(s["batch_size"])))
        s["interval_s"] = max(MIN_INTERVAL_S, float(s["interval_s"]))
        s["period_min"] = max(min_batch_period_min(s["batch_size"], s["interval_s"]), float(s["period_min"]))
        s["daily_cap"] = max(1, min(HARD_DAILY_CAP, int(s.get("daily_cap", DEFAULT_DAILY_CAP))))
        if not s["content"].subject:
            raise ValueError("Oggetto vuoto")

    def stop(self):
        self.stop_event.set()

    def log(self, m):
        self.c.add_log(m)
        self.cb("log", "%s  %s" % (dt.datetime.now().strftime("%d/%m %H:%M:%S"), m))

    def progress(self):
        self.cb("progress", self.c.counts())

    def sleep_until(self, when, label):
        self.cb("next", when)
        while not self.stop_event.is_set():
            rem = (when - dt.datetime.now()).total_seconds()
            if rem <= 0:
                break
            self.cb("status", "%s: tra %s" % (label, fmt_delta(rem)))
            self.stop_event.wait(min(1.0, rem))
        self.cb("next", None)
        return not self.stop_event.is_set()

    def adjust_window(self, when):
        """Sposta 'when' dentro la fascia oraria consentita."""
        if not self.s.get("window_on"):
            return when
        f = dt.datetime.strptime(self.s["window_from"], "%H:%M").time()
        t = dt.datetime.strptime(self.s["window_to"], "%H:%M").time()
        d = when
        for _ in range(3):
            start = dt.datetime.combine(d.date(), f)
            end = dt.datetime.combine(d.date(), t)
            if d < start:
                return start
            if d <= end - dt.timedelta(minutes=self.s["batch_size"] * self.s["interval_s"] / 60.0):
                return d
            d = dt.datetime.combine(d.date() + dt.timedelta(days=1), dt.time(0, 0))
        return when

    def run(self):
        keep_awake(True)
        result = "Completato"
        try:
            for k in ("subject", "body", "docx", "logo", "mode"):
                if k in self.s:
                    self.c.set_meta(k, self.s[k] or "")
            if not self.c.get_meta("started_at"):
                self.c.set_meta("started_at", now_str())
            s = self.s
            self.log("Avvio campagna: blocchi da %d, una email ogni %.0f s, un blocco ogni %.0f min, max %d/24h"
                     % (s["batch_size"], s["interval_s"], s["period_min"], s["daily_cap"]))
            next_start = s.get("start_at") or dt.datetime.now()
            if next_start < dt.datetime.now():
                next_start = dt.datetime.now()
            batch_no = 0
            while not self.stop_event.is_set():
                todo = self.c.next_batch(s["batch_size"])
                if not todo:
                    break
                # limite 24 ore (finestra mobile)
                since = dt.datetime.now() - dt.timedelta(hours=24)
                sent24 = self.c.sent_since(since)
                free = s["daily_cap"] - sent24
                if free <= 0:
                    oldest = self.c.oldest_sent_since(since)
                    wake = (oldest + dt.timedelta(hours=24, minutes=1)) if oldest else dt.datetime.now() + dt.timedelta(hours=1)
                    next_start = max(next_start, wake)
                    self.log("Limite giornaliero raggiunto (%d in 24h): riprendo il %s"
                             % (sent24, next_start.strftime("%d/%m %H:%M")))
                elif free < len(todo):
                    todo = todo[:free]
                next_start = self.adjust_window(next_start)
                if next_start > dt.datetime.now():
                    if not self.sleep_until(next_start, "Prossimo blocco"):
                        break
                    if free <= 0:
                        continue
                batch_no += 1
                batch_start = dt.datetime.now()
                self.log("Blocco %d: %d indirizzi (rimanenti totali %d)" % (batch_no, len(todo), self.c.remaining()))
                try:
                    self.send_batch(todo)
                    next_start = batch_start + dt.timedelta(minutes=s["period_min"])
                except BlockedError as e:
                    self.blocks += 1
                    self.log("ATTENZIONE: Aruba ha segnalato un limite/blocco: %s" % e)
                    if self.blocks >= MAX_BLOCKS:
                        result = "Fermato: Aruba ha bloccato l'invio %d volte. Attendi qualche ora o contatta l'assistenza Aruba." % self.blocks
                        break
                    next_start = dt.datetime.now() + dt.timedelta(minutes=BLOCK_PAUSE_MIN)
                    self.log("Pausa di sicurezza di %d minuti, poi riprendo." % BLOCK_PAUSE_MIN)
                except FatalError as e:
                    result = "Fermato: %s" % e
                    break
                self.progress()
            if self.stop_event.is_set() and result == "Completato":
                result = "Fermato dall'utente (puoi riprendere quando vuoi)"
            elif result == "Completato":
                c = self.c.counts()
                result = "Completato: %d inviate, %d rifiutate, %d errori temporanei non risolti" % (
                    c["sent"], c["failed"], c["temp_error"])
                self.c.set_meta("finished_at", now_str())
            self.log(result)
        except Exception as e:  # noqa
            result = "Errore imprevisto: %r" % e
            self.log(result)
        finally:
            keep_awake(False)
            self.progress()
            self.cb("done", result)

    def _connect(self):
        try:
            return self.a.connect_smtp()
        except smtplib.SMTPAuthenticationError as e:
            raise FatalError("Autenticazione fallita (email/password errati): %s" % _err_text(e.smtp_code, e.smtp_error))
        except smtplib.SMTPResponseException as e:
            text = _err_text(e.smtp_code, e.smtp_error)
            if _is_policy(text) or e.smtp_code == 421:
                raise BlockedError(text)
            raise

    def send_batch(self, todo):
        conn = None
        conn_errors = 0
        i = 0
        try:
            while i < len(todo):
                if self.stop_event.is_set():
                    return
                addr = todo[i]
                if conn is None:
                    try:
                        conn = self._connect()
                    except (OSError, smtplib.SMTPException) as e:
                        conn_errors += 1
                        self.log("Connessione non riuscita (%s). Riprovo tra 2 minuti." % e)
                        if conn_errors >= 5:
                            raise BlockedError("server non raggiungibile dopo 5 tentativi: %s" % e)
                        self.stop_event.wait(120)
                        continue
                try:
                    status, resp = send_one(conn, self.a, addr, self.s["content"])
                except (smtplib.SMTPServerDisconnected, OSError, socket.timeout) as e:
                    # connessione caduta: l'indirizzo NON e' segnato, si riprova dopo riconnessione
                    self.log("Connessione persa durante l'invio a %s (%s): riconnetto." % (addr, e))
                    conn = None
                    conn_errors += 1
                    if conn_errors >= 5:
                        raise BlockedError("troppe disconnessioni: %s" % e)
                    self.stop_event.wait(30)
                    continue
                except BlockedError:
                    self.c.mark(addr, "temp_error", "bloccato dal server, da riprovare", count_attempt=False)
                    raise
                conn_errors = 0
                self.c.mark(addr, status, resp)
                if status != "sent":
                    self.log("%s -> %s" % (addr, resp))
                i += 1
                if i % 10 == 0 or i == len(todo):
                    self.progress()
                if i < len(todo):
                    self.stop_event.wait(self.s["interval_s"])
        finally:
            if conn is not None:
                try:
                    conn.quit()
                except Exception:
                    pass


def fmt_delta(sec):
    sec = int(max(0, sec))
    h, r = divmod(sec, 3600)
    m, s = divmod(r, 60)
    return "%dh %02dm %02ds" % (h, m, s) if h else "%dm %02ds" % (m, s)


def estimate(n, batch_size, interval_s, period_min, start=None, daily_cap=DEFAULT_DAILY_CAP):
    """Stima (numero blocchi, fine prevista) senza fascia oraria."""
    batch_size = max(1, min(MAX_BATCH, int(batch_size)))
    interval_s = max(MIN_INTERVAL_S, float(interval_s))
    period_min = max(min_batch_period_min(batch_size, interval_s), float(period_min))
    if n <= 0:
        return 0, start or dt.datetime.now(), period_min
    blocks = (n + batch_size - 1) // batch_size
    start = start or dt.datetime.now()
    t = start
    sent_today, day_start = 0, start
    for b in range(blocks):
        size = min(batch_size, n - b * batch_size)
        if sent_today + size > daily_cap:
            t = max(t, day_start + dt.timedelta(hours=24))
            day_start, sent_today = t, 0
        end = t + dt.timedelta(seconds=(size - 1) * interval_s)
        sent_today += size
        t = t + dt.timedelta(minutes=period_min)
    return blocks, end, period_min


# ---------------------------------------------------------------- rimbalzi via IMAP
def check_bounces(account, campaign, since_date=None, log=lambda m: None):
    """Legge la Posta in arrivo e segna come 'bounced' gli indirizzi tornati indietro."""
    known = campaign.all_emails()
    if since_date is None:
        st = campaign.get_meta("started_at")
        since_date = parse_ts(st).date() if st else dt.date.today()
    ctx = ssl.create_default_context()
    m = imaplib.IMAP4_SSL(account.imap_host, account.imap_port, ssl_context=ctx)
    try:
        m.login(account.sender_email, account.password)
        m.select("INBOX", readonly=True)
        crit = '(SINCE "%s")' % since_date.strftime("%d-%b-%Y")
        typ, data = m.search(None, crit)
        ids = data[0].split() if data and data[0] else []
        log("Controllo %d messaggi in arrivo dal %s..." % (len(ids), since_date.strftime("%d/%m/%Y")))
        found = 0
        for num in ids:
            typ, hd = m.fetch(num, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT CONTENT-TYPE)])")
            h = email.message_from_bytes(hd[0][1]) if hd and hd[0] else None
            if h is None:
                continue
            frm = (h.get("From") or "").lower()
            subj = (h.get("Subject") or "").lower()
            ctype = (h.get("Content-Type") or "").lower()
            is_bounce = ("mailer-daemon" in frm or "postmaster" in frm or "report-type=delivery-status" in ctype
                         or any(k in subj for k in ("undeliver", "delivery status", "non recapitat", "mancata consegna",
                                                    "returned mail", "failure notice", "delivery failure",
                                                    "mail delivery failed", "impossibile recapitare")))
            if not is_bounce:
                continue
            typ, full = m.fetch(num, "(BODY.PEEK[])")
            msg = email.message_from_bytes(full[0][1])
            addrs = set()
            reason = ""
            for part in msg.walk():
                pct = part.get_content_type()
                if pct == "message/delivery-status":
                    payload = part.get_payload()
                    blocks = payload if isinstance(payload, list) else [part]
                    for blk in blocks:
                        txt = blk.as_string() if hasattr(blk, "as_string") else str(blk)
                        for line in txt.splitlines():
                            ll = line.lower()
                            if ll.startswith("final-recipient") or ll.startswith("original-recipient"):
                                a = line.split(";")[-1].strip().lower()
                                addrs.add(a)
                            if ll.startswith("diagnostic-code"):
                                reason = line.split(":", 1)[-1].strip()
                            elif ll.startswith("status") and not reason:
                                reason = line.strip()
            if not addrs:
                text = ""
                for part in msg.walk():
                    if part.get_content_maintype() == "text":
                        try:
                            text += part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", "replace")
                        except Exception:
                            pass
                for a in FIND_EMAIL_RE.findall(text):
                    a = a.lower()
                    if a in known and a != account.sender_email.lower():
                        addrs.add(a)
            for a in addrs:
                if a in known and campaign.mark_bounced(a, reason or (h.get("Subject") or "")):
                    found += 1
                    log("Rimbalzo: %s  %s" % (a, reason))
        log("Rimbalzi trovati e segnati: %d" % found)
        return found
    finally:
        try:
            m.logout()
        except Exception:
            pass
