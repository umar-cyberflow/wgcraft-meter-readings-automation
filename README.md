# Bulk entry of estimated meter readings for a billing web app (Selenium)

A Python + Selenium tool that automates a monthly back-office task in a
water-utility billing system (WG Craft). Water meters are not connected
online, and only a minority of subscribers report their readings. For the rest,
staff enter an estimated reading by hand: open the account, open the meter tab,
add a reading slightly above the last one, and retry with a higher value if the
system rejects it as outside its daily-average threshold. The script works
through a filtered Excel list of accounts and does this unattended.

> Built for an internal system I work with. **No real subscriber data, URLs or
> credentials are included** – the demo workbook contains fictional accounts.
> Use automation like this only on systems you are authorised to operate.

## What it does

For every account in the Excel list:

1. Searches the account under *Subscribers → Individuals* and opens the *Meters* tab.
2. Opens *Add reading* and reads the last reading from the dialog.
3. Enters `last reading + start_increment` with today's date and confirms.
4. If the system rejects the value because of its daily-average threshold,
   cancels and retries with +1, +2 … until it is accepted.
5. Skips accounts whose meter is removed or inactive, and accounts that already
   have a reading for today, and logs every result.

## Features

- **Retry logic for validation errors:** distinguishes "value outside the
  threshold" (retry with a higher value) from "meter removed" (skip) and any
  other error (skip and report).
- **Safety checks:** works only when the dialog shows exactly one meter;
  fractional readings and readings already entered today are left for a human.
- **Resumable by month:** finished accounts go to `reports/done_YYYY-MM.txt`
  and are skipped on the next run; `--force` re-checks them.
- **Filtered Excel input:** reads only the rows visible under the current
  autofilter, finds the account column by its header, restores lost leading zeros.
- **Robust element lookup:** dialog fields are located by their column headers
  and on-screen position with small JavaScript helpers, so the script does not
  depend on fragile auto-generated IDs.
- **Self-recovering:** after an error it saves a screenshot and the page HTML to
  `reports/debug/`, re-opens the page and logs in again if the session expired;
  it stops after 4 consecutive errors.
- **Secrets stay out of the repo:** URL and credentials come from `.env`.

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env        # then fill in WG_BASE_URL, WG_LOGIN, WG_PASSWORD
```

Requires Python 3.9+ and Google Chrome (Selenium Manager downloads the driver).

## Usage

```bash
python main.py --dry-run --excel examples/demo_accounts.xlsx   # preview, no browser
python main.py --test        # test run: first 3 accounts
python main.py               # full run on the single .xlsx next to main.py (or in excel/)
python main.py --force       # also re-check accounts already done this month
```

Settings in `config.json`:

| Key | Meaning |
|-----|---------|
| `start_increment` | first value added to the last reading |
| `max_increment` / `safety_stop` | upper limits for the retry loop |
| `reading_date` | `null` = today, or a fixed date `dd.mm.yyyy` |
| `only_visible_rows` | ignore rows hidden by an Excel filter |
| `wait_seconds`, `step_delay` | timeouts for a slow site |

## Output

`reports/report_YYYYMMDD_HHMM.csv` has one row per account: status, last
reading, entered value, increment used and a message. Statuses: `ok`,
`already`, `not_found`, `skip_meter`, `skip_threshold`, `skip_other`, `error`.

## Notes

- Element lookup relies on the visible (Russian-language) labels of the web
  app, so it needs adapting for another UI language or product.
- Tech: Python, Selenium, openpyxl, python-dotenv, JavaScript (DOM queries).
