"""
SeePlaces Tools — Flask backend + pywebview launcher
Run: python app.py
"""

import os, sys, json, queue, threading, time
from datetime import datetime
from werkzeug.utils import secure_filename
from flask import Flask, request, Response, stream_with_context, send_from_directory

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR   = os.path.join(BASE_DIR, "uploads")
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")

# Output folders — one per report type
DIR_DOWNLOADS     = os.path.join(BASE_DIR, "outputs", "1 - SeePlaces Downloads")
DIR_GUIDE_REPORTS = os.path.join(BASE_DIR, "outputs", "2 - Guide Reports")
DIR_RECONCILE     = os.path.join(BASE_DIR, "outputs", "3 - Reconciliation")

for d in (UPLOAD_DIR, DIR_DOWNLOADS, DIR_GUIDE_REPORTS, DIR_RECONCILE):
    os.makedirs(d, exist_ok=True)
sys.path.insert(0, BASE_DIR)

app  = Flask(__name__, static_folder=os.path.join(BASE_DIR, "static"))
PORT = 5174

_jobs: dict[str, queue.Queue] = {}
_io_lock = threading.Lock()

REPS_MAP = {
    "kasha":  "katarzyna.jaszcz@rep.itaka.pl",
    "aneta":  "aneta.dymek@rep.itaka.pl",
    "joanna": "joanna.kisiel@rep.itaka.pl",
}


# ── helpers ───────────────────────────────────────────────────────────────────

class _QueueStream:
    def __init__(self, q):
        self._q = q
    def write(self, text):
        if text and text.strip():
            self._q.put({"type": "log", "msg": text.rstrip()})
    def flush(self):
        pass

def _make_job():
    job_id = str(time.time_ns())
    q = queue.Queue()
    _jobs[job_id] = q
    return job_id, q

def _ts():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def _load_settings():
    # Hosted deployments inject credentials via env vars (kept out of the repo).
    # Strip whitespace — a trailing space/newline pasted into the host dashboard
    # is a common cause of "invalid username or password".
    env_email = (os.environ.get("SEEPLACES_EMAIL") or "").strip()
    env_pw    = (os.environ.get("SEEPLACES_PASSWORD") or "").strip()
    # If the email value is clearly not an address (e.g. the env var was set to
    # the key name by mistake), fall back to the known account email.
    if env_email and "@" not in env_email:
        env_email = "transport@btt.tn"
    if env_email and env_pw:
        return {"email": env_email, "password": env_pw}
    # Local fallback: settings.json (gitignored) if present.
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"email": env_email or "transport@btt.tn", "password": env_pw or ""}

def _save_settings(data: dict):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


# ── Static / index ────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")

@app.route("/healthz")
def healthz():
    return {"ok": True}

@app.route("/api/diag")
def api_diag():
    """Non-sensitive credential check: shows the configured email and the
    password *length* only (never the password), to verify env vars are set
    correctly on the host."""
    s = _load_settings()
    return {
        "email":            s.get("email", ""),
        "password_len":     len(s.get("password", "") or ""),
        "env_email_set":    bool(os.environ.get("SEEPLACES_EMAIL")),
        "env_password_set": bool(os.environ.get("SEEPLACES_PASSWORD")),
    }


# ── Settings ──────────────────────────────────────────────────────────────────

@app.route("/api/settings", methods=["GET"])
def get_settings():
    return _load_settings()

@app.route("/api/settings", methods=["POST"])
def post_settings():
    _save_settings(request.json)
    return {"ok": True}


# ── Assignment config (region rules + hotel overrides) ────────────────────────

@app.route("/api/assignment-config", methods=["GET"])
def get_assignment_config():
    import guide_report
    return guide_report.load_config()

@app.route("/api/assignment-config", methods=["POST"])
def post_assignment_config():
    import guide_report
    guide_report.save_config(request.json)
    return {"ok": True}


# ── File upload / download ────────────────────────────────────────────────────

@app.route("/api/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    if not f:
        return {"error": "no file"}, 400
    safe = f"{_ts()}_{secure_filename(f.filename)}"
    path = os.path.join(UPLOAD_DIR, safe)
    f.save(path)
    return {"path": path, "name": f.filename}

@app.route("/api/files/<filename>")
def serve_output(filename):
    for folder in (DIR_DOWNLOADS, DIR_GUIDE_REPORTS, DIR_RECONCILE):
        if os.path.exists(os.path.join(folder, filename)):
            return send_from_directory(folder, filename, as_attachment=True)
    return {"error": "File not found"}, 404

@app.route("/api/open-output")
def open_output():
    import subprocess
    folder_key = request.args.get("folder", "")
    folder_map = {
        "downloads":  DIR_DOWNLOADS,
        "guides":     DIR_GUIDE_REPORTS,
        "reconcile":  DIR_RECONCILE,
    }
    folder = folder_map.get(folder_key, os.path.join(BASE_DIR, "outputs"))
    subprocess.Popen(f'explorer "{folder}"')
    return {"ok": True}


# ── SSE stream ────────────────────────────────────────────────────────────────

@app.route("/api/stream/<job_id>")
def stream(job_id):
    q = _jobs.get(job_id)
    if not q:
        return {"error": "Job not found"}, 404

    def generate():
        while True:
            try:
                item = q.get(timeout=20)
                yield f"data: {json.dumps(item)}\n\n"
                if item.get("type") in ("done", "error"):
                    break
            except queue.Empty:
                # Frequent pings keep hosting proxies from closing the stream.
                yield f"data: {json.dumps({'type':'ping'})}\n\n"
        _jobs.pop(job_id, None)

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# ── Download SeePlaces report ─────────────────────────────────────────────────

@app.route("/api/download", methods=["POST"])
def api_download():
    data = request.json
    job_id, _ = _make_job()
    threading.Thread(target=_run_download, args=(job_id, data), daemon=True).start()
    return {"job_id": job_id}

def _run_download(job_id, data):
    q = _jobs.get(job_id)
    if not q: return
    with _io_lock:
        old, sys.stdout = sys.stdout, _QueueStream(q)
        sys.stderr = sys.stdout
        try:
            import seeplaces_downloader as dl
            dl.EMAIL         = data["email"]
            dl.PASSWORD      = data["password"]
            dl.OUTPUT_FOLDER = DIR_DOWNLOADS
            rep_names = data.get("reps", [])
            token   = dl.get_token()
            content = dl.download_report(token, data["dateFrom"], data["dateTo"],
                                          data["channel"],
                                          [REPS_MAP[n] for n in rep_names if n in REPS_MAP])
            path = dl.save_file(content, data["dateFrom"], data["dateTo"],
                                data["channel"], rep_names)
            name = os.path.basename(path)
            q.put({"type": "done", "filename": name, "msg": f"✅  Ready: {name}"})
        except Exception as e:
            q.put({"type": "error", "msg": f"❌  {e}"})
        finally:
            sys.stdout = sys.stderr = old


# ── Guide report (manual file) ────────────────────────────────────────────────

@app.route("/api/guide-report", methods=["POST"])
def api_guide_report():
    data = request.json
    job_id, _ = _make_job()
    threading.Thread(target=_run_guide_report, args=(job_id, data), daemon=True).start()
    return {"job_id": job_id}

def _run_guide_report(job_id, data):
    q = _jobs.get(job_id)
    if not q: return
    with _io_lock:
        old, sys.stdout = sys.stdout, _QueueStream(q)
        sys.stderr = sys.stdout
        try:
            import guide_report
            out_name = f"guide_report_{_ts()}.xlsx"
            summary  = guide_report.run(
                data["spPath"], os.path.join(DIR_GUIDE_REPORTS, out_name),
                date_from=data.get("dateFrom"), date_to=data.get("dateTo"),
            )
            q.put({"type": "done", "filename": out_name, "summary": summary, "msg": "✅  Done!"})
        except Exception as e:
            q.put({"type": "error", "msg": f"❌  {e}"})
        finally:
            sys.stdout = sys.stderr = old


# ── Guide report FULL (date range → auto login → download → generate) ─────────

@app.route("/api/guide-report-full", methods=["POST"])
def api_guide_report_full():
    data = request.json
    job_id, _ = _make_job()
    threading.Thread(target=_run_guide_report_full, args=(job_id, data), daemon=True).start()
    return {"job_id": job_id}

def _run_guide_report_full(job_id, data):
    q = _jobs.get(job_id)
    if not q: return
    with _io_lock:
        old, sys.stdout = sys.stdout, _QueueStream(q)
        sys.stderr = sys.stdout
        try:
            settings = _load_settings()
            if not settings.get("password"):
                raise RuntimeError("No password saved. Please save your credentials in the Hotels tab first.")

            import seeplaces_downloader as dl
            dl.EMAIL         = settings["email"]
            dl.PASSWORD      = settings["password"]
            dl.OUTPUT_FOLDER = UPLOAD_DIR

            q.put({"type": "step", "msg": "Logging in to SeePlaces…"})
            token = dl.get_token()

            q.put({"type": "step", "msg": "Downloading online report…"})
            content = dl.download_report(token, data["dateFrom"], data["dateTo"], "online", [])

            sp_path = os.path.join(UPLOAD_DIR, f"auto_online_{_ts()}.xlsx")
            with open(sp_path, "wb") as f:
                f.write(content)

            q.put({"type": "step", "msg": "Building guide report…"})
            import guide_report
            out_name = f"guide_report_{_ts()}.xlsx"
            summary  = guide_report.run(
                sp_path, os.path.join(DIR_GUIDE_REPORTS, out_name),
                date_from=data.get("dateFrom"), date_to=data.get("dateTo"),
            )
            q.put({"type": "done", "filename": out_name, "summary": summary, "msg": "✅  Done!"})
        except Exception as e:
            q.put({"type": "error", "msg": f"❌  {e}"})
        finally:
            sys.stdout = sys.stderr = old


# ── Analytics (date-range analysis → charts) ──────────────────────────────────

@app.route("/api/analyze-full", methods=["POST"])
def api_analyze_full():
    data = request.json
    job_id, _ = _make_job()
    threading.Thread(target=_run_analyze_full, args=(job_id, data), daemon=True).start()
    return {"job_id": job_id}

def _run_analyze_full(job_id, data):
    q = _jobs.get(job_id)
    if not q: return
    with _io_lock:
        old, sys.stdout = sys.stdout, _QueueStream(q)
        sys.stderr = sys.stdout
        try:
            settings = _load_settings()
            if not settings.get("password"):
                raise RuntimeError("No password saved. Go to the Hotels tab and save your credentials first.")

            import seeplaces_downloader as dl
            dl.EMAIL         = settings["email"]
            dl.PASSWORD      = settings["password"]
            dl.OUTPUT_FOLDER = UPLOAD_DIR

            q.put({"type": "step", "msg": "Logging in to SeePlaces…"})
            token = dl.get_token()

            channel = data.get("channel", "all")
            q.put({"type": "step", "msg": f"Downloading {channel} report…"})
            content = dl.download_report(token, data["dateFrom"], data["dateTo"], channel, [])

            sp_path = os.path.join(UPLOAD_DIR, f"auto_analyze_{_ts()}.xlsx")
            with open(sp_path, "wb") as f:
                f.write(content)

            q.put({"type": "step", "msg": "Analysing data…"})
            import guide_report
            analytics = guide_report.analyze(sp_path, data["dateFrom"], data["dateTo"], channel)
            q.put({"type": "done", "analytics": analytics, "msg": "✅  Done!"})
        except Exception as e:
            q.put({"type": "error", "msg": f"❌  {e}"})
        finally:
            sys.stdout = sys.stderr = old


# ── History ───────────────────────────────────────────────────────────────────

@app.route("/api/history", methods=["GET"])
def api_history():
    import history as hist
    limit = int(request.args.get("limit", 16))
    return hist.get_history(limit)


# ── Reconcile ─────────────────────────────────────────────────────────────────

@app.route("/api/reconcile", methods=["POST"])
def api_reconcile():
    data = request.json
    job_id, _ = _make_job()
    threading.Thread(target=_run_reconcile, args=(job_id, data), daemon=True).start()
    return {"job_id": job_id}

def _run_reconcile(job_id, data):
    q = _jobs.get(job_id)
    if not q: return
    with _io_lock:
        old, sys.stdout = sys.stdout, _QueueStream(q)
        sys.stderr = sys.stdout
        try:
            import reconcile as rc
            out_name  = f"reconciliation_{_ts()}.xlsx"
            summary   = rc.reconcile(data["pairs"],
                                     os.path.join(DIR_RECONCILE, out_name))
            q.put({"type": "done", "filename": out_name, "summary": summary, "msg": "✅  Done!"})
        except Exception as e:
            q.put({"type": "error", "msg": f"❌  {e}"})
        finally:
            sys.stdout = sys.stderr = old


# ── Live Worklist (auto download → detect new bookings → checklist) ───────────

@app.route("/api/worklist", methods=["GET"])
def api_worklist_get():
    import worklist
    return worklist.get_worklist()

@app.route("/api/worklist/entered", methods=["POST"])
def api_worklist_entered():
    import worklist
    data = request.json or {}
    return worklist.set_entered(data["key"], bool(data.get("entered")))

@app.route("/api/worklist/refresh", methods=["POST"])
def api_worklist_refresh():
    """Synchronous: log in, download, diff — and return the result directly.
    Fast HTTP login makes this ~2-3s, so no background job / SSE is needed
    (which also avoids cross-worker 'job not found' issues on hosted servers)."""
    import worklist
    data = request.json or {}
    try:
        worklist._meta_set("last_status", "running")
        settings = _load_settings()
        if not settings.get("password"):
            raise RuntimeError("No SEEPLACES_PASSWORD set on the server.")

        import seeplaces_downloader as dl
        dl.EMAIL         = settings["email"]
        dl.PASSWORD      = settings["password"]
        dl.OUTPUT_FOLDER = UPLOAD_DIR

        token   = dl.get_token()
        content = dl.download_report(token, data["dateFrom"], data["dateTo"], "all", [])

        report_path = os.path.join(UPLOAD_DIR, f"worklist_{_ts()}.xlsx")
        with open(report_path, "wb") as f:
            f.write(content)

        lines  = worklist.parse_lines(report_path)
        result = worklist.sync(lines)
        worklist._meta_set("last_report_path", report_path)
        worklist._meta_set("last_error", "")
        worklist._meta_set("last_status", "ok")

        return {
            "ok":              True,
            "new_count":       result["new_count"],
            "cancelled_count": result["cancelled_count"],
            "last_sync":       result["last_sync"],
        }
    except Exception as e:
        try:
            worklist._meta_set("last_error", f"{type(e).__name__}: {e}")
            worklist._meta_set("last_status", "error")
        except Exception:
            pass
        return {"ok": False, "error": str(e)}, 500

# ── Entry point ───────────────────────────────────────────────────────────────

def _port_in_use(port: int) -> bool:
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) == 0

def _run_flask():
    import logging
    logging.getLogger("werkzeug").setLevel(logging.ERROR)
    app.run(port=PORT, threaded=True, use_reloader=False, debug=False)

def _wait_until_ready(port: int, timeout: float = 30.0) -> bool:
    """Block until the Flask server actually accepts connections, or time out."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _port_in_use(port):
            return True
        time.sleep(0.1)
    return False

if __name__ == "__main__":
    # If something is already on our port, it's almost certainly a previous
    # instance of this app — just open a window pointing at it instead of
    # crashing the Flask thread on a bind error.
    already_running = _port_in_use(PORT)
    if not already_running:
        threading.Thread(target=_run_flask, daemon=True).start()
        if not _wait_until_ready(PORT):
            print(f"❌  Server failed to start on port {PORT}. "
                  f"Close any other copy of this app and try again.")
            sys.exit(1)

    url = f"http://localhost:{PORT}"
    try:
        import webview
        webview.create_window("SeePlaces Tools", url,
                               width=920, height=800, resizable=True)
        webview.start()
    except ImportError:
        import webbrowser
        print(f"Opening in browser → {url}")
        webbrowser.open(url)
        try:
            while True: time.sleep(1)
        except KeyboardInterrupt:
            pass
