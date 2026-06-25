# SeePlaces Tools (web)

Internal dashboard for the SeePlaces workflow: a **live worklist** that watches for
new bookings (with sound + desktop alerts and a checkable, persistent to-do list),
a **manifest generator** that rebuilds the "Booked excursions list" per excursion,
plus the existing download / guide-report / reconcile / analytics tools.

The app logs in to SeePlaces with a headless Chrome browser (Selenium), so the
hosted version runs inside Docker with Chrome bundled.

## Deploy to Render

1. Push this repo to GitHub (already done).
2. In Render: **New + → Blueprint**, connect this repo. Render reads `render.yaml`
   and builds the `Dockerfile`.
3. Set the two environment variables when prompted (kept out of the repo):
   - `SEEPLACES_EMAIL` = `transport@btt.tn`
   - `SEEPLACES_PASSWORD` = `<the SeePlaces password>`
4. Deploy. You get a public URL like `https://seeplaces.onrender.com`.

### Notes on the Render free plan
- The instance **sleeps after ~15 min idle** (first request then takes ~30–60s).
- The filesystem is **ephemeral**: the worklist database (`worklist.db`) and any
  generated files reset on redeploy/sleep. Fine for shared, day-to-day use; for
  permanent history, upgrade to a paid plan with a persistent disk.
- Memory is 512 MB — enough for short Chrome download bursts.

## Run locally

```bash
pip install -r requirements.txt
python app.py          # desktop window (pywebview) or browser fallback
```

Or via Docker:

```bash
docker build -t seeplaces .
docker run -p 10000:10000 -e SEEPLACES_EMAIL=... -e SEEPLACES_PASSWORD=... seeplaces
# open http://localhost:10000
```

## Security

This deployment has **no login gate** (per current setup) and uses a single shared
SeePlaces account. Keep the URL private. `settings.json`, `uploads/`, `outputs/`,
and `*.db` are gitignored so credentials and customer data never reach the repo.
