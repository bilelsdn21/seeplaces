"""
SeePlaces - Supplier Report Downloader
=======================================
Requirements:
    pip install selenium requests webdriver-manager

Usage:
    python seeplaces_downloader.py
"""

import requests
import os
import time
import json
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ─────────────────────────────────────────────
# CONFIG — edit these
# ─────────────────────────────────────────────
EMAIL         = "transport@btt.tn"
PASSWORD      = "YOUR_PASSWORD_HERE"
OUTPUT_FOLDER = r"C:\Reports\SeePlaces"

LOGIN_URL  = "https://admin.seeplaces.com"
REPORT_URL = "https://admin.seeplaces.com/api/report/getsupplierreport"

# Sales channel -> saleScopeId UUID
SALES_CHANNELS = {
    "all":     None,
    "offline": ["fd3e5c7e-66f3-4f9b-bfaa-e5e744a1b35f"],
    "online":  ["166c5c1d-2c02-46be-905d-733add254c37"],
}

# Rep nicknames -> full emails
REPS = {
    "kasha":  "katarzyna.jaszcz@rep.itaka.pl",
    "aneta":  "aneta.dymek@rep.itaka.pl",
    "joanna": "joanna.kisiel@rep.itaka.pl",
}
# ─────────────────────────────────────────────


def get_token_via_browser() -> str:
    """Log in via browser and intercept the Bearer token from a real API call."""
    print("Opening browser to log in...")

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1280,800")
    options.add_argument("--log-level=3")
    # Keep memory low enough for small hosted instances (e.g. Render 512 MB).
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--disable-backgrounding-occluded-windows")
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    # Capture network traffic so we can read the Bearer token off a real API call.
    options.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    driver = webdriver.Chrome(options=options)

    try:
        driver.get(LOGIN_URL)
        wait = WebDriverWait(driver, 25)

        # Wait for the login form to render (SeePlaces now uses a Keycloak SSO page).
        pwd_field = wait.until(EC.presence_of_element_located((By.NAME, "password")))

        # Find the e-mail/username field — name varies (Keycloak: "username",
        # older panel: "email"). Try the likely selectors in order.
        user_field = None
        for by, sel in [(By.NAME, "username"), (By.NAME, "email"),
                        (By.ID, "username"),
                        (By.CSS_SELECTOR, "input[type='email']"),
                        (By.CSS_SELECTOR, "input[type='text']")]:
            els = driver.find_elements(by, sel)
            if els:
                user_field = els[0]
                break
        if user_field is None:
            raise RuntimeError("Could not find the e-mail/username field on the login page.")

        user_field.clear()
        user_field.send_keys(EMAIL)
        pwd_field.clear()
        pwd_field.send_keys(PASSWORD)

        # Submit login
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

        # Wait until redirected to the main app
        wait.until(lambda d: "sso.seeplaces.com" not in d.current_url)
        print("Logged in! Waiting for app to load...")
        time.sleep(4)

        # Primary: the SPA keeps the token in memory and sends it as a Bearer
        # header on its API calls. Visit a data page to trigger those calls,
        # then read the Authorization header from the network log.
        token = None
        try:
            driver.get("https://admin.seeplaces.com/bookings")
            time.sleep(6)
            token = _token_from_network_log(driver)
        except Exception:
            pass

        # Fallback: scan storage for anything JWT-shaped (older panel versions).
        if not token:
            token = driver.execute_script("""
                for (const store of [localStorage, sessionStorage]) {
                    for (const key of Object.keys(store)) {
                        const val = store.getItem(key);
                        if (val && val.includes('eyJ')) {
                            const parts = val.match(/eyJ[A-Za-z0-9_\\-\\.]+/);
                            if (parts) return parts[0];
                        }
                    }
                }
                return null;
            """)

        if token:
            print("Token extracted successfully.")
            return token
        else:
            raise RuntimeError("Could not extract token from browser.")

    except Exception as e:
        # Capture what the browser actually saw — distinguishes a slow page from
        # a Cloudflare/bot block (common when running from a datacenter IP).
        info = ""
        try:
            body = driver.find_element(By.TAG_NAME, "body").text[:300].replace("\n", " ")
            info = (f" | url={driver.current_url} | title={driver.title!r}"
                    f" | inputs={len(driver.find_elements(By.TAG_NAME,'input'))}"
                    f" | body={body!r}")
        except Exception:
            pass
        raise RuntimeError(f"Login failed ({type(e).__name__}){info}")

    finally:
        driver.quit()


def _token_from_network_log(driver) -> str | None:
    """Pull a Bearer token from the captured performance/network log."""
    try:
        logs = driver.get_log("performance")
    except Exception:
        return None
    for entry in logs:
        try:
            msg = json.loads(entry["message"])["message"]
            if msg.get("method") != "Network.requestWillBeSent":
                continue
            headers = msg["params"]["request"].get("headers", {})
            for hk, hv in headers.items():
                if hk.lower() == "authorization" and "bearer" in hv.lower():
                    return hv
        except Exception:
            continue
    return None


def download_report(token: str, date_from: str, date_to: str,
                    sales_channel: str, representatives: list) -> bytes:

    payload = {
        "searchFilter":          None,
        "offerNameSearchFilter": None,
        "bookingStatus":         None,
        "affiliateIDs":          None,
        "countryIDs":            None,
        "domainIds":             None,
        "groupBooking":          None,
        "guideLanguageIDs":      None,
        "iataCodes":             None,
        "offerDateFrom":         f"{date_from}T00:00:00.000Z",
        "offerDateTo":           f"{date_to}T23:59:59.000Z",
        "offerOwners":           None,
        "regionIDs":             None,
        "representatives":       representatives if representatives else None,
        "saleScopeIds":          SALES_CHANNELS.get(sales_channel),
        "salesChannel":          None,
        "supplierIDs":           None,
    }

    token = token.strip()
    if token.lower().startswith("bearer "):
        token = token[7:]

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/json",
        "Accept":        "*/*",
        "x-timezone":    "Europe/Paris",
        "Origin":        "https://admin.seeplaces.com",
        "Referer":       "https://admin.seeplaces.com/bookings",
    }

    print(f"Downloading report ({date_from} to {date_to}, channel: {sales_channel})...")
    resp = requests.post(REPORT_URL, json=payload, headers=headers)

    if resp.status_code == 401:
        print("Token rejected (401). Try again.")
        raise SystemExit(1)
    if resp.status_code == 400:
        print(f"Bad request (400): {resp.text}")
        raise SystemExit(1)

    resp.raise_for_status()
    return resp.content


def save_file(content: bytes, date_from: str, date_to: str,
              sales_channel: str, rep_names: list) -> str:
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    reps_label = "_".join(rep_names) if rep_names else "all"
    filename   = f"supplier_{date_from}_{date_to}_{sales_channel}_{reps_label}_{timestamp}.xlsx"
    filepath   = os.path.join(OUTPUT_FOLDER, filename)
    with open(filepath, "wb") as f:
        f.write(content)
    print(f"File saved: {filepath}")
    return filepath


def prompt_inputs():
    print("\n============================")
    print("  SeePlaces Report Downloader")
    print("============================\n")

    date_from = input("Date from (YYYY-MM-DD): ").strip()
    date_to   = input("Date to   (YYYY-MM-DD): ").strip()

    print("Sales channel: all / online / offline")
    sales_channel = input("Sales channel [all]: ").strip().lower() or "all"
    if sales_channel not in SALES_CHANNELS:
        print("Invalid channel, defaulting to 'all'")
        sales_channel = "all"

    print(f"Available reps: {', '.join(REPS.keys())} -- comma-separated, or blank for all")
    reps_input = input("Representatives: ").strip().lower()

    rep_names, representatives = [], []
    if reps_input:
        for name in [r.strip() for r in reps_input.split(",")]:
            if name in REPS:
                rep_names.append(name)
                representatives.append(REPS[name])
            else:
                print(f"Unknown rep '{name}', skipping.")

    return date_from, date_to, sales_channel, representatives, rep_names


def main():
    date_from, date_to, sales_channel, representatives, rep_names = prompt_inputs()

    try:
        token = get_token_via_browser()
    except Exception as e:
        print(f"Auto-login failed: {e}")
        print("Falling back to manual token.")
        print("(DevTools -> Network -> any request -> Headers -> authorization)")
        token = input("Paste Bearer token: ").strip()

    content = download_report(token, date_from, date_to, sales_channel, representatives)
    save_file(content, date_from, date_to, sales_channel, rep_names)
    print("Done!")


if __name__ == "__main__":
    main()
