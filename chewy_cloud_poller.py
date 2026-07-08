#!/usr/bin/env python3
"""Poll the reports inbox over IMAP for Chewy "Brand Snapshot" PDFs.

Runs inside the chewy-report GitHub Action. Finds messages carrying a PDF
attachment whose filename contains "Brand Snapshot" (Jeff's monthly Chewy
snapshots, forwarded to the reports mailbox), and saves each one into the
snapshots folder inside the data-repo checkout so chewy_extract.py can parse it.

Idempotency: processed email Message-IDs are recorded in a state file inside the
data repo, so an email is never ingested twice.

Credentials: env EMAIL_USER, EMAIL_APP_PASSWORD (Gmail app password — grants
IMAP). Exit status is always 0; `found`/`files` are written to $GITHUB_OUTPUT.
"""
import os
import re
import sys
import email
import imaplib
import argparse
from email.header import decode_header, make_header
from datetime import date, timedelta

IMAP_HOST = "imap.gmail.com"
NAME_MATCH = "brand snapshot"   # case-insensitive substring of the PDF filename


def _decode(s):
    if not s:
        return ""
    try:
        return str(make_header(decode_header(s)))
    except Exception:
        return str(s)


def load_processed(state_file):
    if not os.path.exists(state_file):
        return set()
    with open(state_file, "r", encoding="utf-8") as f:
        return {line.strip() for line in f if line.strip()}


def append_processed(state_file, msg_id):
    with open(state_file, "a", encoding="utf-8") as f:
        f.write(msg_id + "\n")


def snapshot_pdfs(msg):
    """Yield (filename, bytes) for each 'Brand Snapshot' PDF attachment."""
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        fname = _decode(part.get_filename())
        if not fname or not fname.lower().endswith(".pdf"):
            continue
        if NAME_MATCH not in fname.lower():
            continue
        payload = part.get_payload(decode=True)
        if payload:
            yield fname, payload


def _safe(name):
    return re.sub(r"[^A-Za-z0-9._ ()-]", "_", name).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--inbox-dir", required=True,
                    help="Snapshot folder to save PDFs into (data-repo checkout).")
    ap.add_argument("--state-file", required=True,
                    help="File tracking processed Message-IDs (in the data repo).")
    ap.add_argument("--since-days", type=int, default=45)
    args = ap.parse_args()

    user = os.environ.get("EMAIL_USER")
    pw = os.environ.get("EMAIL_APP_PASSWORD")
    if not user or not pw:
        print("[chewy-poller] EMAIL_USER / EMAIL_APP_PASSWORD not set",
              file=sys.stderr)
        return _emit(False, [])

    os.makedirs(args.inbox_dir, exist_ok=True)
    processed = load_processed(args.state_file)
    saved = []

    M = imaplib.IMAP4_SSL(IMAP_HOST)
    try:
        M.login(user, pw)
        M.select("INBOX")
        since = (date.today() - timedelta(days=args.since_days)).strftime("%d-%b-%Y")
        # Any message with a PDF; we filter by attachment name below.
        typ, data = M.search(None, "SINCE", since)
        ids = data[0].split() if data and data[0] else []
        print(f"[chewy-poller] scanning {len(ids)} message(s) since {since}")

        for num in ids:
            typ, raw = M.fetch(num, "(RFC822)")
            if typ != "OK" or not raw or not raw[0]:
                continue
            msg = email.message_from_bytes(raw[0][1])
            msg_id = (msg.get("Message-ID") or "").strip() or f"uid:{num.decode()}"
            if msg_id in processed:
                continue
            pdfs = list(snapshot_pdfs(msg))
            if not pdfs:
                continue
            for fname, payload in pdfs:
                out = os.path.join(args.inbox_dir, _safe(fname))
                with open(out, "wb") as f:
                    f.write(payload)
                print(f"[chewy-poller] saved {os.path.basename(out)} "
                      f"({len(payload)} bytes) from {_decode(msg.get('Subject'))!r}")
                saved.append(os.path.basename(out))
            append_processed(args.state_file, msg_id)
            processed.add(msg_id)
            try:
                M.store(num, "+FLAGS", "\\Seen")
            except Exception:
                pass
    finally:
        try:
            M.logout()
        except Exception:
            pass

    return _emit(bool(saved), saved)


def _emit(found, files):
    print(f"[chewy-poller] found={found} files={len(files)}")
    gh_out = os.environ.get("GITHUB_OUTPUT")
    if gh_out:
        with open(gh_out, "a", encoding="utf-8") as f:
            f.write(f"found={'true' if found else 'false'}\n")
            f.write(f"files={';'.join(files)}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
