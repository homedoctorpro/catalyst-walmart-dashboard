#!/usr/bin/env python3
"""Poll the reports inbox over IMAP for new weekly Catalyst workbooks.

Runs inside the GitHub Actions weekly-report workflow. Finds messages whose
subject contains "Weekly Sales Report Catalyst" with an .xlsx/.xlsb attachment,
derives the 2026XX week code from the subject (or attachment filename), and
saves the attachment into the data-repo checkout as

    2026XX Weekly Sales Report Catalyst.<ext>

Idempotency: processed email Message-IDs are recorded in a state file inside the
data repo (committed by the workflow), so an email is never ingested twice even
if it was already opened/marked-read in the mailbox.

Credentials come from env: EMAIL_USER, EMAIL_APP_PASSWORD (the same Gmail app
password used to send the distro). Auth works for IMAP because Gmail app
passwords grant IMAP access.

Exit status is always 0; whether new data was found is reported via stdout and,
when running under Actions, the `found`/`weeks` outputs written to $GITHUB_OUTPUT.
"""
import os
import re
import sys
import email
import imaplib
import argparse
from email.header import decode_header, make_header

# IMAP subject anchor. Gmail's IMAP SUBJECT search is word-tokenized, not
# substring — "Report" does NOT match "Reports". So anchor only on the tokens
# that are invariant across Hailey's wording variants ("Weekly Sales Report
# Catalyst" vs "Weekly Sales Reports Catalyst"): "Weekly Sales". The real gate
# is downstream — a message only ingests if it also has a 2026XX week code and
# an xlsx/xlsb attachment — so a loose anchor can't pull in the wrong file.
IMAP_HOST = "imap.gmail.com"
SUBJECT_MATCH = "Weekly Sales"
WEEK_RE = re.compile(r"(2026\d{2})")
ATTACH_EXTS = (".xlsx", ".xlsb")


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


def find_attachment(msg):
    """Return (filename, bytes) for the first .xlsx/.xlsb attachment, or None."""
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        fname = _decode(part.get_filename())
        if fname and fname.lower().endswith(ATTACH_EXTS):
            payload = part.get_payload(decode=True)
            if payload:
                return fname, payload
    return None


def week_from(subject, filename):
    for source in (subject, filename or ""):
        m = WEEK_RE.search(source)
        if m:
            return m.group(1)
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--inbox-dir", required=True,
                    help="Directory to save attachments into (the data-repo checkout).")
    ap.add_argument("--state-file", required=True,
                    help="File tracking processed Message-IDs (lives in the data repo).")
    ap.add_argument("--since-days", type=int, default=30,
                    help="Only consider messages received within this many days.")
    args = ap.parse_args()

    user = os.environ.get("EMAIL_USER")
    pw = os.environ.get("EMAIL_APP_PASSWORD")
    if not user or not pw:
        print("[poller] EMAIL_USER / EMAIL_APP_PASSWORD not set", file=sys.stderr)
        return _emit(False, [])

    processed = load_processed(args.state_file)
    saved_weeks = []

    M = imaplib.IMAP4_SSL(IMAP_HOST)
    try:
        M.login(user, pw)
        M.select("INBOX")
        # Date-bounded subject search; keeps the working set small over time.
        from datetime import date, timedelta
        since = (date.today() - timedelta(days=args.since_days)).strftime("%d-%b-%Y")
        typ, data = M.search(None, "SINCE", since, "SUBJECT", f'"{SUBJECT_MATCH}"')
        ids = data[0].split() if data and data[0] else []
        print(f"[poller] {len(ids)} candidate message(s) since {since}")

        for num in ids:
            typ, raw = M.fetch(num, "(RFC822)")
            if typ != "OK" or not raw or not raw[0]:
                continue
            msg = email.message_from_bytes(raw[0][1])
            msg_id = (msg.get("Message-ID") or "").strip() or f"uid:{num.decode()}"
            if msg_id in processed:
                continue
            subject = _decode(msg.get("Subject"))
            att = find_attachment(msg)
            if not att:
                print(f"[poller] skip (no xlsx/xlsb attachment): {subject!r}")
                append_processed(args.state_file, msg_id)
                processed.add(msg_id)
                continue
            fname, payload = att
            week = week_from(subject, fname)
            if not week:
                print(f"[poller] skip (no 2026XX week code): subj={subject!r} file={fname!r}")
                append_processed(args.state_file, msg_id)
                processed.add(msg_id)
                continue
            ext = ".xlsb" if fname.lower().endswith(".xlsb") else ".xlsx"
            out = os.path.join(args.inbox_dir, f"{week} Weekly Sales Report Catalyst{ext}")
            with open(out, "wb") as f:
                f.write(payload)
            print(f"[poller] saved week {week}: {os.path.basename(out)} ({len(payload)} bytes) "
                  f"from subject {subject!r}")
            append_processed(args.state_file, msg_id)
            processed.add(msg_id)
            saved_weeks.append(week)
            # Mark read for tidiness (state file is the real dedupe guard).
            try:
                M.store(num, "+FLAGS", "\\Seen")
            except Exception:
                pass
    finally:
        try:
            M.logout()
        except Exception:
            pass

    return _emit(bool(saved_weeks), saved_weeks)


def _emit(found, weeks):
    weeks = sorted(set(weeks))
    print(f"[poller] found={found} weeks={','.join(weeks) if weeks else '-'}")
    gh_out = os.environ.get("GITHUB_OUTPUT")
    if gh_out:
        with open(gh_out, "a", encoding="utf-8") as f:
            f.write(f"found={'true' if found else 'false'}\n")
            f.write(f"weeks={','.join(weeks)}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
