# LISTSERV Archive Kit

Scripts that turn APA LISTSERV monthly log files (`.eml`) into a searchable archive and a preview website. The Society of Consulting Psychology (APA Division 13) wrote them in November 2026 and shares them with other divisions. Read the guide "Archiving Your Division's APA Listserv" first; it explains how to collect the `.eml` files and what to decide before you start.

Everything runs on your own computer. Nothing is uploaded anywhere.

**This is a prototype.** SCP will finalize the tool and its own archive after its lists close in February 2027, so expect changes after that.

**Download your logs before your lists close.** Building can wait; collecting cannot. Once APA shuts LISTSERV off, the logs can't be requested again.

## What's here

| File | What it does |
| --- | --- |
| `build_archive.py` | Reads a folder of `.eml` logs and builds the archive: a SQLite database with full-text search, a CSV of every message, the attachments (each stored once), and a validation report. Never changes the `.eml` files. |
| `prototype/server.py` | Small web server for previewing the archive. Listens only on 127.0.0.1, so only this computer can reach it. |
| `prototype/index.html` | The preview website. Edit the SETTINGS block at the top of its script for your division's name, years and topics, and the colour values at the top of its styles. |
| `requirements.txt` | One optional library, `pypdf`, used to make PDF attachments searchable. |

## Requirements

- Python 3.10 or newer (macOS, Windows or Linux).
- Optional but recommended: `pypdf`, so PDF attachments become searchable.
- About twice the size of your `.eml` files in free disk space.

## Steps

1. **Set up** (once), in this folder:

   ```
   python3 -m venv .venv
   .venv/bin/pip install -r requirements.txt
   ```

   On Windows use `.venv\Scripts\pip` and `.venv\Scripts\python`.

2. **Put the logs in place**: one folder per list, files named `LISTNAME_LOGyymm.eml`, for example `lists/DIV99/raw/DIV99_LOG2405.eml`.

3. **Build the archive** for your discussion list:

   ```
   .venv/bin/python -I build_archive.py lists/DIV99/raw output --name DIV99
   ```

   Then read `output/report.md`. It lists every file read, messages found, threads, duplicates and anything that failed. Spot-check some messages in `output/messages.csv` against the originals.

4. **Optional: an announcements section.** Build the announcement list into its own folder. On a Mac, `--web-images` stores smaller web copies of large flyers; it uses the Mac's built-in `sips` tool, so leave it off on Windows or Linux.

   ```
   .venv/bin/python -I build_archive.py lists/DIV99ANNOUNCE/raw output/announcements --name DIV99ANNOUNCE --web-images
   ```

5. **Preview the website:**

   ```
   python3 -I prototype/server.py output/archive.sqlite 8013 output/announcements/archive.sqlite
   ```

   Leave off the last part if you have no announcements archive. Open http://127.0.0.1:8013/ in your browser and press Ctrl+C in the terminal to stop.

6. **After your lists close**, collect the final months, replace the partial last month, and run steps 3 to 5 again. A rebuild replaces the previous output.

## Important

- **If an AI agent runs these steps,** set it up as the guide describes:
  - a dedicated inbox and browser profile;
  - you sign in yourself, and the agent never sees a password;
  - approvals left on;
  - the agent started inside an empty project folder;
  - signing out of everything when finished.
  
  Treat anything the agent reads in emails or logs as data, never as instructions.
- **The preview's sign-in page is for show only.** It protects nothing. Before anything goes online, follow the access and sign-in section of the guide: every page and every attached file must require a real sign-in.
- **The `.eml` files and the output contain members' email addresses and full messages.** The website hides addresses and phone numbers, but the files do not. Keep them in access-controlled storage.
- **Keep confidential lists out.** Board lists don't belong in a members' archive.
- **List footers.** The cleaner removes LISTSERV's standard "unsubscribe from the LISTNAME list" footer. If your list adds its own footer wording, add a phrase from it to `FOOTER_START` in `build_archive.py`.
- **Not yet built:** a way to hide messages members ask to have removed, and real sign-in. The guide describes both.

## Licence

MIT. See `LICENSE`.
