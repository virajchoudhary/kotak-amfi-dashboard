# Deployment Verification Checks

Run these checks from the repository root before handing the build to Kotak server operations.

## Frontend

```powershell
npm install
npm run check:frontend
```

## Backend

```powershell
python -m venv .venv-deploy-check
.\.venv-deploy-check\Scripts\Activate.ps1
python -m pip install -r backend\requirements.lock.txt
python -m compileall backend
python -m pip check
python -m pip install pip-audit bandit
python -m pip_audit -r backend\requirements.lock.txt
$env:PYTHONUTF8 = "1"
python -m bandit backend\main.py backend\db.py backend\excel_engine.py backend\__init__.py
```

> **Bandit scope:** scan the application source files explicitly (above). Do **not** run
> `bandit -r backend`, because it recurses into `backend\venv` and scans tens of thousands of
> lines of third-party packages (and crashes the text formatter on non-ASCII characters under the
> Windows `cp1252` console). `PYTHONUTF8=1` avoids the encoding crash. The application source is
> expected to report **0 findings**.

To run the backend unit tests from a clean environment:

```powershell
python -m unittest discover -s backend\tests -p "test_*.py" -v
```

## Runtime configuration

Production must set at least:

```powershell
$env:APP_ENV = "production"
$env:AMFI_DB_PATH = "D:\amfi-dashboard-data\amfi.db"   # persistent storage, OUTSIDE the app directory
$env:AMFI_ALLOWED_HOSTS = "dashboard.kotak.example"     # REQUIRED, or Host-header validation is silently disabled
$env:AMFI_REQUIRE_PROXY_IDENTITY = "true"
$env:AMFI_IDENTITY_HEADER = "X-Forwarded-User"
$env:AMFI_ENABLE_DOCS = "false"
```

If React and FastAPI are hosted on separate HTTPS origins, also set `AMFI_ALLOWED_ORIGINS` to the
exact frontend origin list. Same-origin deployment should leave CORS origins empty in production.

### Full environment-variable reference

| Variable | Default | Effect |
| --- | --- | --- |
| `APP_ENV` | `development` | `prod`/`production` enables production-mode defaults (docs off, proxy identity required, empty CORS/host lists). |
| `AMFI_DB_PATH` | `backend/amfi.db` | SQLite database file path. **Must** be set in production to a persistent path outside the app directory (readiness fails otherwise). |
| `AMFI_ALLOWED_HOSTS` | empty in prod | Comma-separated allowed `Host` headers. **If unset in production, `TrustedHostMiddleware` is not registered and every Host header is accepted.** |
| `AMFI_ALLOWED_ORIGINS` | empty in prod | Comma-separated CORS origins. Empty ⇒ CORS middleware not added (correct for same-origin). Never set to `*`. |
| `AMFI_REQUIRE_PROXY_IDENTITY` | `true` in prod | When true, requests without the identity header are rejected with `401`. |
| `AMFI_IDENTITY_HEADER` | `X-Forwarded-User` | Header the app reads the authenticated user from. See **Reverse proxy & identity** below. |
| `AMFI_ENABLE_DOCS` | `false` in prod | Enables `/docs`, `/redoc`, `/openapi.json`, and the `/` → `/docs` redirect. Keep off in production. |
| `AMFI_MAX_UPLOAD_MB` | `25` | Max accepted upload size; larger uploads return `413`. |
| `AMFI_MAX_WORKBOOK_UNZIPPED_MB` | `100` | Zip-bomb guard: max uncompressed `.xlsx` size; larger returns `400`. |
| `AMFI_ALLOW_HTML_UPLOAD` | `false` | **Leave unset/false in production.** When true the backend parses raw upload bytes with `pandas.read_html` (untrusted-markup parsing). |
| `AMFI_HOST` | `127.0.0.1` | Bind address when launched via `python main.py`. Keep loopback/private; do not bind `0.0.0.0` on a directly reachable interface. |
| `AMFI_PORT` | `8000` | Bind port when launched via `python main.py`. |

## Reverse proxy & identity (REQUIRED)

The application does **not** authenticate users itself. It trusts the `AMFI_IDENTITY_HEADER`
(default `X-Forwarded-User`) on every request and treats its value as the authenticated user for
auditing and access. This is only safe behind a trusted, authenticating reverse proxy. The
deployment **must** satisfy all of the following:

1. The reverse proxy authenticates the user (SSO/Kerberos/OIDC/etc.) and **injects**
   `X-Forwarded-User` on each forwarded request.
2. The reverse proxy **strips or overwrites any client-supplied copy** of that header on inbound
   requests, so a client cannot forge it.
3. The FastAPI app is reachable **only** through the proxy — bind it to loopback or a private
   interface and firewall direct access. If a client can reach the app directly, it can send
   `X-Forwarded-User: anyone` and be authenticated/audited as that user. (`AMFI_REQUIRE_PROXY_IDENTITY`
   only enforces that the header is *present*, not that it is *authentic*.)
4. The proxy terminates TLS; the app speaks plain HTTP on its private bind address.

Because identity is derived entirely from this header, audit-log attribution is only trustworthy
when steps 1–3 are enforced.

## Database & storage constraints (SQLite — REQUIRED reading)

The backend uses an embedded **SQLite** database (WAL mode). It is acceptable **only** for a
single-server, single-instance deployment:

- **Single instance only.** Do not point more than one application instance at the same
  `AMFI_DB_PATH`. Two instances sharing one SQLite file (especially over a network/SMB share) risk
  lock contention and database corruption.
- **Single worker process.** Run the backend with exactly one worker (`uvicorn --workers 1`). The
  upload serialization lock is in-process only; multiple worker processes against the same file are
  **not** serialized by it and can hit `SQLITE_BUSY` under concurrent uploads. Multiple workers are
  not supported unless specifically tested and approved.
- **Persistent storage.** `AMFI_DB_PATH` must be on durable, non-ephemeral local storage (not a
  container overlay/tmpfs and not a network share).
- **Backups.** Back up the database file regularly, including the `-wal` and `-shm` sidecar files.
  This file holds all uploaded fund data.
- **Multi-instance / high-availability is not supported.** Horizontal scaling or HA requires
  migrating to a central database (e.g. PostgreSQL) first.

## Health, readiness & traffic gating

```powershell
Invoke-WebRequest https://<host>/api/health
Invoke-WebRequest https://<host>/api/readiness
```

- `/api/readiness` returns `200` when ready and `503` with an `errors` list when production
  configuration is incomplete (missing `AMFI_DB_PATH`, DB inside the app directory, missing
  `AMFI_ALLOWED_HOSTS`, docs enabled, or proxy identity not required).
- **Readiness gates the probe, not the process.** The app still serves traffic while readiness is
  `503`. The load balancer / orchestrator **must** be configured to pull the instance from rotation
  when `/api/readiness` returns `503`; otherwise a misconfigured instance will serve live traffic.

### Suggested production smoke tests

- `/api/readiness` returns `200 ready`; deliberately unset `AMFI_DB_PATH` / `AMFI_ALLOWED_HOSTS` and
  confirm `503` with the matching error.
- Reach the app host directly (bypassing the proxy) with a forged `X-Forwarded-User` and confirm the
  network controls make it unreachable; through the proxy with no upstream auth, confirm `401`.
- Upload a valid `.xlsx` → `200`. Negative uploads: `.csv` → `400`, oversize → `413`, corrupted zip
  → `400`, undetectable month → `400`. Confirm none return `500`.
- Force a reconciliation failure and confirm the previous month's data is unchanged.
- `/docs`, `/redoc`, `/openapi.json` → `404` in production.
