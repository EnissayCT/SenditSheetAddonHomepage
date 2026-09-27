# Architecture - Sendit Google Sheets Add-on

Editor add-on (HTML sidebar + menu Extensions). Not a Workspace Card add-on.

## Components

- **Sidebar.html / Config.html / Welcome.html** — HtmlService UI
- **Code.gs** — auth, districts, create, bulk, sync, labels, template
- **Sendit API** — `https://app.sendit.ma/api/v1`

## Storage

- `PropertiesService` (user) — public key, secret key, pickup district id
- `CacheService` (user) — bearer token, 1 hour
- Hidden sheet `_Sendit_Districts_Cache` — city list for the open spreadsheet

## Auth

Config form keys → `POST /login` → token cached. First-time setup can load cities **before** keys are saved (`fetchDistrictsForConfig`).

## Create

Sheet row or sidebar form → validate name / phone / address / district → `POST /deliveries` → write code + mapped French status.

City match: district name, `ville`, or legacy `name - price DH`.

Phone: digits only, leading 0, exactly 10 characters. Template column is plain text.

## Sync

Manual only. Batch `getValues` / `setValues` / `setBackgrounds`. Skips `livré`, `retourné`, `annulé`, `refusé`.

## Labels

Selected rows → `POST /deliveries/getlabels` (`printFormat: 1`) → PDF URL.

## Scopes

```
spreadsheets.currentonly
script.container.ui
script.external_request
```

No Drive, no ScriptApp triggers, no web app / `doPost`.

## Status mapping

API codes (PENDING, DELIVERED, RETURNED, …) map to French sheet values (`en attente`, `livré`, `retourné`, …).
