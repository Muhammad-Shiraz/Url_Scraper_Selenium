import json
import os
import re
import time
import pandas as pd
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ── SETTINGS ────────────────────────────────────────────────
INPUT_XLSX      = "Groene Serie Rechtspersonens.xlsx"
OUTPUT_XLSX     = "Result_Groene Serie Rechtspersonens.xlsx"
PROGRESS_FILE   = "progress.json"
CAPTCHA_TIMEOUT = 60
# ────────────────────────────────────────────────────────────


def title_to_slug(title):
    slug = title.lower()
    slug = re.sub(r"[^a-z0-9\s]", " ", slug)
    slug = re.sub(r"\s+", "-", slug.strip())
    slug = re.sub(r"-+", "-", slug)
    return slug


def wait_for_slug_in_url(driver, slug, timeout=15):
    start = time.time()
    last_path = None
    while time.time() - start < timeout:
        time.sleep(0.5)
        current = driver.current_url.lower()
        path_only = current.split("?")[0]
        last_path = path_only
        if slug in current or slug in path_only:
            return current
    print(f"  Last URL checked: {last_path}")
    return None


def captcha_present(driver):
    signals = ["captcha", "recaptcha", "i'm not a robot", "unusual traffic"]
    page = driver.page_source.lower()
    return any(s in page for s in signals)


def wait_for_captcha(driver):
    print("⚠️  CAPTCHA detected — please solve it manually in the browser...")
    start = time.time()
    while time.time() - start < CAPTCHA_TIMEOUT:
        time.sleep(3)
        if not captcha_present(driver):
            print("✓ CAPTCHA solved, continuing.")
            return True
    print("✗ CAPTCHA timeout — skipping this title.")
    return False


def get_chromedriver():
    base = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base, "chromedriver.exe")
    if os.path.exists(path):
        return path
    raise FileNotFoundError("chromedriver.exe not found next to this script.")


def start_chrome():
    opts = Options()
    opts.add_argument("--start-maximized")
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.managed_default_content_settings.fonts": 2,
        "profile.managed_default_content_settings.media_stream": 2,
    }
    opts.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(service=Service(get_chromedriver()), options=opts)


def wait_for_login(driver):
    print("\n" + "=" * 60)
    print("  Opening inview.nl — please:")
    print("  1. Accept the cookie banner")
    print("  2. Log in when prompted")
    print("  Script continues automatically once the search box appears.")
    print("=" * 60)
    driver.get("https://www.inview.nl/zoeken")
    while True:
        try:
            boxes = driver.find_elements(
                By.CSS_SELECTOR,
                "input.wk-field-input[data-testid='search-bar-input-field']"
            )
            if boxes and boxes[0].is_displayed():
                print("✅ Logged in — starting search loop.\n")
                return
        except Exception:
            pass
        time.sleep(2)


# ────────────────────────────────────────────────────────────
# CORE SEARCH
# returns: [url]  → matched
#          []     → no match found
#          None   → hard error (captcha timeout / crash)
# ────────────────────────────────────────────────────────────
def search_title(driver, title):
    slug = title_to_slug(title)
    print(f"  Slug : {slug}")

    driver.get("https://www.inview.nl/zoeken")

    try:
        box = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((
                By.CSS_SELECTOR,
                "input.wk-field-input[data-testid='search-bar-input-field']"
            ))
        )
        print("  ✅Search box found")
    except Exception:
        print("  ✗ Search box not found within 20 s — skipping.")
        return None

    driver.execute_script("arguments[0].scrollIntoView(true);", box)

    driver.execute_script(
        """
        var el  = arguments[0];
        var val = arguments[1];
        var setter = Object.getOwnPropertyDescriptor(
                         window.HTMLInputElement.prototype, 'value').set;
        el.focus();
        setter.call(el, val);
        el.dispatchEvent(new Event('input',  { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        """,
        box, title
    )
    print("  Typed title into search box")
    time.sleep(0.3)

    try:
        btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.cg-search-button"))
        )
        btn.click()
        print("  Search button clicked")
        time.sleep(1)
    except Exception:
        box.send_keys(Keys.RETURN)
        print("  Search submitted with ENTER key")
        time.sleep(1)

    time.sleep(2)

    if captcha_present(driver):
        if not wait_for_captcha(driver):
            return None

    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.search-cluster"))
        )
        
    except Exception:
        print("  ⚠️  Timed out waiting for search clusters.")
        return []

    commentaar_cluster = None
    for c in driver.find_elements(By.CSS_SELECTOR, "div.search-cluster"):
        try:
            if c.get_attribute("data-e2e-cluster-name") == "Commentaar":
                commentaar_cluster = c
                print(f"  Cluster found")
                break
        except Exception:
            pass

    if not commentaar_cluster:
        print("  ✗ No 'Commentaar' cluster found.")
        return []

    links = []
    for item in commentaar_cluster.find_elements(
        By.CSS_SELECTOR, "[data-testid='search-results-item-title'] a"
    ):
        try:
            href = item.get_attribute("href") or ""
            if href.startswith("/"):
                href = "https://www.inview.nl" + href
            if href:
                links.append(href)
        except Exception:
            pass

    print(f"  Candidate links: {len(links)}")

    for idx, link in enumerate(links, 1):
        print(f"  [{idx}] {link}")
        driver.get(link)
        # print("First URL before waiting for slug:", driver.current_url)
        final_url = wait_for_slug_in_url(driver, slug, timeout=15)
        print(f"  URL after waiting for slug: {final_url}")
        if final_url:
            print(f"  ✅✅✅✅ Match: {final_url}")
            return [final_url]
        print("  ❌❌❌❌❌ Slug not found after redirect.")

    return []


# ────────────────────────────────────────────────────────────
# PROGRESS  — keyed by document_id (string)
#
# Each entry stored in progress.json:
#   {
#     "document_id": "8539994",
#     "title":       "Groene Serie ...",
#     "urls":        ["https://..."] | [] | null
#   }
# ────────────────────────────────────────────────────────────
def load_progress():
    """Returns dict[document_id -> entry_dict]."""
    if not os.path.exists(PROGRESS_FILE):
        return {}
    try:
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception as e:
        print(f"⚠️  Could not read {PROGRESS_FILE}: {e}")
    return {}


def save_progress(progress):
    """Atomic write — a crash never corrupts the file."""
    tmp = PROGRESS_FILE + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(progress, f, ensure_ascii=False, indent=2)
        os.replace(tmp, PROGRESS_FILE)          # atomic on Windows & Linux
    except Exception as e:
        print(f"⚠️  Could not save progress: {e}")
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


# ────────────────────────────────────────────────────────────
# INPUT
# ────────────────────────────────────────────────────────────
def read_input():
    """Returns list of dicts: [{'document_id': str, 'title': str}, ...]"""
    if not os.path.exists(INPUT_XLSX):
        raise FileNotFoundError(f"Input file not found: {INPUT_XLSX}")

    df = pd.read_excel(INPUT_XLSX)

    missing = [c for c in ("document_id", "primary_title") if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing column(s): {missing}. "
            f"Available columns: {list(df.columns)}"
        )

    df = df.dropna(subset=["document_id", "primary_title"]).copy()
    df["document_id"]   = df["document_id"].astype(str).str.strip()
    df["primary_title"] = df["primary_title"].astype(str).str.strip()

    return [
        {"document_id": row["document_id"], "title": row["primary_title"]}
        for _, row in df.iterrows()
    ]


# ────────────────────────────────────────────────────────────
# OUTPUT
# ────────────────────────────────────────────────────────────
def save_results(rows, progress):
    """
    rows     : list of {'document_id', 'title'}  (original input order)
    progress : dict[document_id -> entry]
    """
    wb = Workbook()
    ws = wb.active
    ws.append(["#", "Document ID", "Title", "URL"])

    for i, row in enumerate(rows, 1):
        doc_id = row["document_id"]
        title  = row["title"]
        entry  = progress.get(doc_id)

        if entry is None:
            url_cell = "NOT PROCESSED"
        elif entry["urls"] is None:
            url_cell = "SKIPPED (error)"
        elif len(entry["urls"]) == 0:
            url_cell = "NO MATCH"
        else:
            url_cell = entry["urls"][0]

        ws.append([i, doc_id, title, url_cell])

    wb.save(OUTPUT_XLSX)
    print(f"\n✅ Saved → {OUTPUT_XLSX}")


# ────────────────────────────────────────────────────────────
# MAIN
# ────────────────────────────────────────────────────────────
def main():
    rows     = read_input()          # [{'document_id', 'title'}, ...]
    progress = load_progress()       # {document_id: entry}
    done_ids = set(progress.keys())

    remaining = [r for r in rows if r["document_id"] not in done_ids]

    print(f"Total rows   : {len(rows)}")
    print(f"Already done : {len(done_ids)}")
    print(f"To process   : {len(remaining)}")

    if not remaining:
        print("\n✅ All titles already processed — saving Excel.")
        save_results(rows, progress)
        return

    driver = start_chrome()
    wait_for_login(driver)

    try:
        for i, row in enumerate(rows, 1):
            doc_id = row["document_id"]
            title  = row["title"]

            if doc_id in done_ids:
                print(f"[{i}/{len(rows)}] ⏩ SKIP  ID={doc_id}")
                continue

            print(f"\n[{i}/{len(rows)}] ID={doc_id}")
            print(f"  Title: {title[:80]}{'...' if len(title) > 80 else ''}")

            try:
                urls = search_title(driver, title)
            except Exception as e:
                print(f"  ✗ Unexpected error: {e}")
                urls = None

            # store result keyed by document_id — title is also saved
            # so progress.json is fully self-contained
            progress[doc_id] = {
                "document_id": doc_id,
                "title":       title,
                "urls":        urls,
            }
            done_ids.add(doc_id)
            save_progress(progress)   # atomic write after every row

            time.sleep(0.5)

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    save_results(rows, progress)


if __name__ == "__main__":
    main()