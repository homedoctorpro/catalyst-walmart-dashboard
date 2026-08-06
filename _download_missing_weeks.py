#!/usr/bin/env python3
"""Download weekly Catalyst workbook attachments from ligneticsdata inbox to
this folder for posterity. Keys the week code off the ATTACHMENT FILENAME
(reliable) rather than the subject (which can be wrong on forwards). Read-only
mailbox access: does NOT mark messages read. Skips weeks already present locally.

Importable: build scripts call download_missing_weeks() before globbing the
weekly workbooks, so local builds can't run against a stale folder.
"""
import imaplib, email, re, os
from email.header import decode_header, make_header
from datetime import date, timedelta

WEEK_RE = re.compile(r"(20\d{4})")
ATTACH_EXTS = (".xlsx", ".xlsb")
HERE = os.path.dirname(os.path.abspath(__file__))

def dec(s):
    try: return str(make_header(decode_header(s))) if s else ""
    except Exception: return str(s)

def download_missing_weeks(verbose=True):
    """Fetch any weekly workbooks in the inbox not yet on disk.
    Returns list of (week, size, filename) saved. Raises on IMAP/auth failure —
    callers treat that as non-fatal (warn and build with what's on disk)."""
    import credentials as C
    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(C.EMAIL_USER, C.EMAIL_APP_PASSWORD)
    M.select("INBOX", readonly=True)
    since = (date.today() - timedelta(days=60)).strftime("%d-%b-%Y")
    typ, data = M.search(None, "SINCE", since, "SUBJECT", '"Weekly Sales"')
    ids = data[0].split() if data and data[0] else []

    seen_week = {}   # week -> saved path (dedupe within run)
    saved, skipped = [], []
    for num in ids:
        typ, raw = M.fetch(num, "(BODY.PEEK[])")
        if typ != "OK" or not raw or not raw[0]:
            continue
        msg = email.message_from_bytes(raw[0][1])
        subj = dec(msg.get("Subject"))
        att = None
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            fn = dec(part.get_filename())
            if fn and fn.lower().endswith(ATTACH_EXTS):
                # Same-thread replies carry other workbooks (endcap sales etc.)
                # whose names also contain a week code — only accept the real
                # weekly report by filename.
                if "weekly sales report" not in fn.lower():
                    continue
                payload = part.get_payload(decode=True)
                if payload:
                    att = (fn, payload)
                    break
        if not att:
            continue
        fn, payload = att
        # Week from FILENAME first (correct), subject only as fallback.
        m = WEEK_RE.search(fn) or WEEK_RE.search(subj)
        if not m:
            continue
        week = m.group(1)
        ext = ".xlsb" if fn.lower().endswith(".xlsb") else ".xlsx"
        out = os.path.join(HERE, f"{week} Weekly Sales Report Catalyst{ext}")
        if week in seen_week:
            skipped.append((week, "dup email this run"))
            continue
        if os.path.exists(out):
            skipped.append((week, f"already on disk ({os.path.getsize(out)} B)"))
            seen_week[week] = out
            continue
        with open(out, "wb") as f:
            f.write(payload)
        seen_week[week] = out
        saved.append((week, len(payload), fn))
    M.logout()

    if verbose:
        print("=== SAVED ===")
        for week, sz, fn in sorted(saved):
            print(f"  week {week}: {sz} B  (from attachment {fn!r})")
        print("=== SKIPPED ===")
        for week, why in sorted(skipped):
            print(f"  week {week}: {why}")
        if not saved:
            print("  (nothing new saved)")
    return saved

if __name__ == "__main__":
    download_missing_weeks()
