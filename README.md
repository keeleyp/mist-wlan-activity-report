# Mist WLAN Activity Report

A Python script that combines two views of a Juniper Mist org into one per-AP report:

1. **Which SSIDs actually reach each AP** — resolved the same way Mist itself pushes WLAN
   Templates down to APs (Template `applies`/`exceptions` site scoping, plus
   `filter_by_deviceprofile` / `deviceprofile_ids` matching).
2. **Whether that AP has had a real 5GHz client** in the last N hours, for one or more
   hour thresholds at once (e.g. 24h and 48h).

Cross-referencing the two surfaces the interesting case neither report alone shows: an AP
that **publishes a 5GHz SSID but has no actual 5GHz client activity** — a likely coverage or
hardware problem rather than just "nobody's there."

This combines the WLAN Template coverage logic from
[mist-ap-audit](https://github.com/keeleyp/mist-ap-audit)'s
`mist_deviceprofile_wlan_report.py` (there it's grouped by Site + Device Profile) with the
per-AP activity-check approach from
[mist-5ghz-client-check](https://github.com/keeleyp/mist-5ghz-client-check), applying the
WLAN resolution per-AP instead so both datasets share the same row.

## What It Does

1. Fetches WLAN Templates, WLANs, and site groups (org-level).
2. Fetches the site list, then AP stats site-by-site (site-level `stats/devices` is required
   for uptime, port stats, and radio stats — the org-level endpoint doesn't return them).
3. For every AP, resolves which WLANs reach it from its site + Device Profile against the
   Template rules.
4. For every AP, looks up its most recent 5GHz client session via a single targeted API call.
5. Flags each AP Yes/No per threshold, and flags the "publishes 5GHz but quiet" cases.
6. Builds a multi-sheet Excel workbook (autofilter + frozen header row on every sheet) and
   saves it locally.

### Why per-site AP stats, and why one call per AP for activity?

- Per-site: only `/sites/{site_id}/stats/devices` returns uptime, `port_stat` (wired port
  counters), and `radio_stat` (per-radio channel/power/bytes) — the org-level bulk endpoint
  omits them.
- Per-AP for activity: Mist's `clients/sessions/search` endpoint is a raw connection-event
  log, not a summary — on a busy org it can return **tens of millions of rows** in a 24–48h
  window (every roam/reassociation counts as a row), so pulling the whole org's session
  history and aggregating client-side isn't practical. Instead the script asks Mist directly,
  per AP: *"what's the most recent 5GHz session here?"* (`ap=<mac>&band=5&limit=1&sort=-timestamp`).
  That bounds the work to one call per AP no matter how much client churn the org has.

The trade-off is call *count*, not call *cost* — a large org needs a matching number of API
calls, which can exceed Mist's 5,000-calls/hour token limit. The script handles that
automatically (see [Rate-Limit Handling](#rate-limit-handling) below).

## Output

Saved as `Mist_WLAN_Activity_Report_<OrgName>_<timestamp>.xlsx`, with:

| Sheet | Content |
|---|---|
| All APs | One row per AP: identity/state, uptime, ETH0 + Band 5 radio stats, published SSIDs (all bands and 5GHz-only), and a Yes/No column per activity threshold |
| No 5GHz Clients - Nh | One sheet per threshold, listing only the APs with zero 5GHz clients in that window |
| Quiet in ALL Thresholds | APs with no 5GHz client across *every* threshold checked (includes disconnected APs, for which that's expected) |
| Possible Issues | The actionable subset: **connected** APs, publishing >=1 5GHz SSID, with no 5GHz client across every threshold. Disconnected APs are excluded here — they trivially have no clients, which isn't an RF/coverage issue |
| No WLANs Published | APs with **zero** SSIDs reaching them on any band — a config/coverage gap (no WLAN Template resolves to this AP's site + Device Profile), independent of client activity |
| AP x SSID Detail | One row per (AP, SSID) pair, with WLAN Template name, configured bands, and per-band rateset |

Every AP-level sheet includes: Name, Site Name, Site ID, Device Profile, MAC, Serial, Model,
Firmware, Status, Uptime (hours), ETH0 RX/TX Bytes, Band 5 Num WLANs/Channel/Bandwidth/Power/
RX/TX Bytes, Published SSIDs (+ count), Published 5GHz SSIDs (+ count), last 5GHz client
timestamp, and hours since.

Note: because each threshold is a "past N hours from now" window, they're nested — an AP
quiet for 48h is necessarily also quiet for 24h. "Quiet in ALL Thresholds" will always match
the sheet for your *largest* threshold; the smaller-threshold sheets are still useful to spot
APs that have freshly gone quiet.

Note: WLAN bands/rates shown are as *configured* on the WLAN — this doesn't cross-check
against each AP's actual `radio_config` (a radio can be disabled at the AP/profile level even
if the WLAN's `bands` field includes it). That's exactly what the "Possible Issues" sheet is
for: it points at APs where the configured coverage and the observed activity disagree.

## Prerequisites

- Python 3.8+
- A Mist API token with read access to the org (Org Settings → API Tokens in the Mist dashboard)

### Install dependencies

```bash
pip install -r requirements.txt
```

## Setup

1. Copy the example config and fill in your details:

```bash
cp mist_wlan_activity_report.ini.example mist_wlan_activity_report.ini
```

2. Edit `mist_wlan_activity_report.ini`:

```ini
[mist]
api_base = https://api.eu.mist.com/api/v1
org_id = YOUR_ORG_ID_HERE
api_token = YOUR_API_TOKEN_HERE

[output]
directory = ~

[thresholds]
hours = 24,48
```

- `api_base` — matches whichever Mist cloud region your org's dashboard is on (check the
  `manage.<region>.mist.com` hostname you log into). US is `https://api.mist.com/api/v1`, EU
  is `https://api.eu.mist.com/api/v1`.
- `org_id` — visible in the Mist dashboard URL, or via `GET /api/v1/self`.
- `api_token` — create under Org Settings → API Tokens. Read-only access is sufficient.
- `hours` — comma-separated default hour thresholds for the activity check. Override per-run
  with `--hours` on the command line if needed.

`mist_wlan_activity_report.ini` is gitignored — it holds your real token, so it's never committed.

## Usage

```bash
python3 mist_wlan_activity_report.py
```

Runs with the thresholds from the ini file's `[thresholds] hours` (default `24,48`). To
check different windows for a single run without editing the ini:

```bash
python3 mist_wlan_activity_report.py --hours 12 24 48 72
```

The script shows a pre-flight summary (org name, AP/site counts, estimated API calls, current
rate-limit headroom) and asks for confirmation before making any calls beyond the initial org
lookup.

## Rate-Limit Handling

Mist tokens are typically limited to 5,000 API calls/hour. This script uses one call per site
(AP stats) plus one call per AP (5GHz lookup), on top of a handful of org-level calls for
WLAN Templates/WLANs/site groups. Orgs with more sites+APs than the remaining hourly quota
will hit the limit mid-run. When that happens the script:

1. Checks remaining quota every 50 sites/APs (via `GET /self/usage`).
2. Pauses with a live countdown and prompts `w` (wait for the window to reset, then continue
   automatically) or `q` (quit without saving).

For orgs with a few thousand APs this typically finishes well within a single hour. For very
large orgs (tens of thousands of APs) expect the run to span multiple rate-limit windows — the
script handles this unattended if you choose `w` each time, so it's safe to kick off and leave
running.

## Configuration Reference

| Key | Section | Description |
|---|---|---|
| `api_base` | `[mist]` | Mist API base URL for your cloud region |
| `org_id` | `[mist]` | Organisation ID |
| `api_token` | `[mist]` | API token (keep secret — this file is gitignored) |
| `directory` | `[output]` | Output directory for the generated `.xlsx` report |
| `hours` | `[thresholds]` | Comma-separated default hour thresholds, e.g. `24,48` |

CLI flags:

| Flag | Default | Description |
|---|---|---|
| `--hours` / `-x` | from ini `[thresholds] hours` | One or more hour thresholds to check, space-separated (overrides the ini for this run only) |

## API Calls Used

- `GET /orgs/{org_id}` — org name
- `GET /orgs/{org_id}/stats` — AP/site counts for pre-flight
- `GET /orgs/{org_id}/templates` — WLAN Templates (applies/exceptions/filter_by_deviceprofile)
- `GET /orgs/{org_id}/wlans` — WLANs (SSID, bands, rateset, template_id)
- `GET /orgs/{org_id}/sitegroups` — site group → site ID expansion
- `GET /orgs/{org_id}/sites?limit=1000&page=N` — site list (page-based)
- `GET /sites/{site_id}/stats/devices?type=ap&limit=1000` — AP stats, one call per site
- `GET /orgs/{org_id}/clients/sessions/search?ap={mac}&band=5&start=X&end=Y&limit=1&sort=-timestamp` — most recent 5GHz session, one call per AP
- `GET /self/usage` — rate-limit headroom check

## Security

- Never commit `mist_wlan_activity_report.ini` — it contains your live API token. It's
  already listed in `.gitignore`.
- Use a read-only API token scoped to the org you're auditing.

## License

MIT
