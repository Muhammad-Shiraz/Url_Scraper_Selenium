# import os
# import re
# import time
# import pandas as pd
# from openpyxl import Workbook
# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC


# # ── SETTINGS ────────────────────────────────────────────────
# INPUT_XLSX      = "Groene Serie Rechtspersonen.xlsx"
# OUTPUT_XLSX     = "Result_Groene Serie Rechtspersonen.xlsx"
# WAIT_SECONDS    = 7
# CAPTCHA_TIMEOUT = 60
# # ────────────────────────────────────────────────────────────


# # Convert title → slug (fallback matching logic)
# def title_to_slug(title):
#     slug = title.lower()
#     slug = re.sub(r"[^a-z0-9\s]", " ", slug)
#     slug = re.sub(r"\s+", "-", slug.strip())
#     slug = re.sub(r"-+", "-", slug)
#     return slug


# def captcha_present(driver):
#     signals = ["captcha", "recaptcha", "i'm not a robot", "unusual traffic"]
#     page = driver.page_source.lower()
#     return any(s in page for s in signals)


# def wait_for_captcha(driver):
#     print("⚠️ CAPTCHA detected. Solve it manually...")
#     start = time.time()

#     while time.time() - start < CAPTCHA_TIMEOUT:
#         time.sleep(3)
#         if not captcha_present(driver):
#             print("✓ CAPTCHA solved")
#             return True

#     print("✗ CAPTCHA timeout")
#     return False


# def get_chromedriver():
#     base = os.path.dirname(os.path.abspath(__file__))
#     path = os.path.join(base, "chromedriver.exe")
#     if os.path.exists(path):
#         return path
#     raise FileNotFoundError("chromedriver.exe not found")


# def start_chrome():
#     path = get_chromedriver()
#     opts = Options()
#     opts.add_argument("--start-maximized")
#     return webdriver.Chrome(service=Service(path), options=opts)


# # ────────────────────────────────────────────────────────────
# # 🔥 MAIN SEARCH LOGIC (CLICK FIRST RESULT + VERIFY)
# # ────────────────────────────────────────────────────────────
# def search_title(driver, title):
#     slug = title_to_slug(title)
#     print(f"\nSlug: {slug}")

#     # open page
#     driver.get("https://www.inview.nl/zoeken")
#     time.sleep(5)

#     # login check
#     if "login" in driver.current_url.lower():
#         print("⚠️ Please login manually")
#         input("Press ENTER after login...")

#     # ✅ FIXED: correct data-testid from 'search-overlay-input' → 'search-bar-input-field'
#     box = WebDriverWait(driver, 20).until(
#         EC.element_to_be_clickable(
#             (By.CSS_SELECTOR, "input.wk-field-input[data-testid='search-bar-input-field']")
#         )
#     )

#     # scroll into view
#     driver.execute_script("arguments[0].scrollIntoView(true);", box)
#     time.sleep(1)

#     # ✅ FIXED: use JS focus + React-safe value clear before typing
#     driver.execute_script("arguments[0].focus();", box)
#     time.sleep(0.3)
#     box.send_keys(Keys.CONTROL + "a")
#     box.send_keys(Keys.DELETE)
#     time.sleep(0.3)

#     # type slowly
#     for ch in title:
#         box.send_keys(ch)
#         time.sleep(0.02)

#     # ✅ FIXED: click the search button instead of pressing ENTER
#     # (more reliable for React forms that ignore keyboard submit)
#     try:
#         search_btn = driver.find_element(
#             By.CSS_SELECTOR, "button[data-e2e='cg-search-button']"
#         )
#         search_btn.click()
#     except:
#         box.send_keys(Keys.RETURN)  # fallback

#     time.sleep(WAIT_SECONDS)

#     # CAPTCHA check
#     if captcha_present(driver):
#         if not wait_for_captcha(driver):
#             return None

#     # wait for clusters
#     WebDriverWait(driver, 10).until(
#         EC.presence_of_element_located((By.CSS_SELECTOR, "div.search-cluster"))
#     )

#     clusters = driver.find_elements(By.CSS_SELECTOR, "div.search-cluster")

#     commentaar_cluster = None
#     for c in clusters:
#         try:
#             if c.get_attribute("data-e2e-cluster-name") == "Commentaar":
#                 commentaar_cluster = c
#                 break
#         except:
#             pass

#     if not commentaar_cluster:
#         print("No Commentaar cluster found")
#         return []

#     print("✓ Commentaar cluster found")

#     items = commentaar_cluster.find_elements(
#         By.CSS_SELECTOR,
#         "[data-testid='search-results-item-title'] a"
#     )

#     links = []
#     for item in items:
#         try:
#             print("Found item:", item.text)
#             href = item.get_attribute("href")
#             if href:
#                 if href.startswith("/"):
#                     href = "https://www.inview.nl" + href
#                 links.append(href)
#                 print("Full URL:", href)
#         except:
#             pass

#     print(f"Total links found: {len(links)}")

#     matched_urls = []
#     for idx, link in enumerate(links, 1):
#         print(f"\n[{idx}] Opening: {link}")
#         driver.get(link)
#         time.sleep(7)

#         final_url = driver.current_url
#         if slug in final_url.lower():
#             print("✓ MATCH FOUND")
#             matched_urls.append(final_url)
#             break
#         else:
#             print("✗ NO SLUG IN LINK")

#     if matched_urls:
#         print(f"\n✓ Total matched: {len(matched_urls)}")
#         return matched_urls
#     else:
#         print("\n✗ No matching URLs found")
#         return []

#     # ─────────────────────────────────────────────
#     # 🔥 STEP 2: OPEN EACH LINK + MATCH
#     # ─────────────────────────────────────────────
#     matched_urls = []

#     for idx, link in enumerate(links, 1):
#         print(f"\n[{idx}] Opening: {link}")

#         driver.get(link)
#         time.sleep(7)

#         final_url = driver.current_url
#         # page_text = driver.page_source.lower()

#         # check slug in URL OR page
#         if slug in final_url.lower():
#             # or slug[:30] in page_text:
#             print("✓ MATCH FOUND")

#             matched_urls.append(final_url)
#             print("Match 1 url then breaking loop to avoid multiple matches per title")
#             break  # stop after first match
            

#         else:
#             print("✗ NO SLUG IN LINK")

#     # ─────────────────────────────────────────────
#     # RETURN RESULT
#     # ─────────────────────────────────────────────
#     if matched_urls:
#         print(f"\n✓ Total matched: {len(matched_urls)}")
#         return matched_urls

#     else:
#         print("\n✗ No matching URLs found")
#         return []
# # ────────────────────────────────────────────────────────────
# # EXCEL INPUT
# # ────────────────────────────────────────────────────────────
# # def read_titles():
# #     df = pd.read_excel(INPUT_XLSX)
# #     return [str(x).strip() for x in df.iloc[:, 0].dropna()]

# def read_titles():
#     df = pd.read_excel(INPUT_XLSX)
#     return [str(x).strip() for x in df["primary_title"].dropna()]
# # ────────────────────────────────────────────────────────────
# # SAVE RESULTS
# # ────────────────────────────────────────────────────────────
# def save_results(titles, results):
#     wb = Workbook()
#     ws = wb.active

#     ws.append(["#", "Title", "URL"])

#     for i, t in enumerate(titles, 1):
#         data = results.get(t)

#         if data is None:
#             ws.append([i, t, "SKIPPED"])
#         elif len(data) == 0:
#             ws.append([i, t, "NO MATCH"])
#         else:
#             ws.append([i, t, data[0]])

#     wb.save(OUTPUT_XLSX)
#     print(f"\nSaved → {OUTPUT_XLSX}")


# # ────────────────────────────────────────────────────────────
# # MAIN
# # ────────────────────────────────────────────────────────────
# def main():
#     titles = read_titles()
#     driver = start_chrome()

#     results = {}

#     try:
#         for i, title in enumerate(titles, 1):
#             print(f"\n[{i}/{len(titles)}] {title}")

#             try:
#                 res = search_title(driver, title)
#                 results[title] = res
#             except Exception as e:
#                 print("Error:", e)
#                 results[title] = None

#             time.sleep(2)

#     finally:
#         driver.quit()

#     save_results(titles, results)


# if __name__ == "__main__":
#     main()


































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
INPUT_XLSX      = "Groene Serie Rechtspersonen.xlsx"
OUTPUT_XLSX     = "Result_Groene Serie Rechtspersonen.xlsx"
PROGRESS_FILE   = "progress.json"   # ⚡ resume from here on restart
CAPTCHA_TIMEOUT = 60
# ────────────────────────────────────────────────────────────


# Convert title → slug
def title_to_slug(title):
    slug = title.lower()
    slug = re.sub(r"[^a-z0-9\s]", " ", slug)
    slug = re.sub(r"\s+", "-", slug.strip())
    slug = re.sub(r"-+", "-", slug)
    return slug


# 🔥 NEW: wait for redirect to final slug URL
def wait_for_slug_in_url(driver, slug, timeout=15):
    start = time.time()

    while time.time() - start < timeout:
        current = driver.current_url.lower()

        # remove query params
        path_only = current.split("?")[0]

        if slug in current or slug in path_only:
            return current

        time.sleep(0.4)

    return None


def captcha_present(driver):
    signals = ["captcha", "recaptcha", "i'm not a robot", "unusual traffic"]
    page = driver.page_source.lower()
    return any(s in page for s in signals)


def wait_for_captcha(driver):
    print("⚠️ CAPTCHA detected. Solve it manually...")
    start = time.time()

    while time.time() - start < CAPTCHA_TIMEOUT:
        time.sleep(3)
        if not captcha_present(driver):
            print("✓ CAPTCHA solved")
            return True

    print("✗ CAPTCHA timeout")
    return False


def get_chromedriver():
    base = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base, "chromedriver.exe")
    if os.path.exists(path):
        return path
    raise FileNotFoundError("chromedriver.exe not found")


def start_chrome():
    path = get_chromedriver()
    opts = Options()
    opts.add_argument("--start-maximized")
    # ⚡ Block images, fonts, and media to load pages faster
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.managed_default_content_settings.fonts": 2,
        "profile.managed_default_content_settings.media_stream": 2,
    }
    opts.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(service=Service(path), options=opts)


# ────────────────────────────────────────────────────────────
# ONE-TIME LOGIN HANDLER (called once before the main loop)
# ────────────────────────────────────────────────────────────
def wait_for_login(driver):
    """
    Opens the search page once and waits for the user to:
      1. Accept the cookie banner
      2. Click login and fill the form
    Returns only when the search box is confirmed visible.
    No timeout — gives the user all the time they need.
    """
    print("\n" + "="*60)
    print("  Opening inview.nl — please:")
    print("  1. Accept the cookie banner")
    print("  2. Log in if prompted")
    print("  The script will continue automatically once logged in.")
    print("="*60)

    driver.get("https://www.inview.nl/zoeken")

    # Poll every 2 seconds until the search box appears
    # This gives unlimited time for cookie banner + login form
    while True:
        try:
            url = driver.current_url.lower()
            # If search box is visible → we're logged in and ready
            boxes = driver.find_elements(
                By.CSS_SELECTOR,
                "input.wk-field-input[data-testid='search-bar-input-field']"
            )
            if boxes and boxes[0].is_displayed():
                print("✅ Logged in! Starting search...\n")
                return
        except Exception:
            pass
        time.sleep(2)


# ────────────────────────────────────────────────────────────
# MAIN SEARCH LOGIC
# ────────────────────────────────────────────────────────────
def search_title(driver, title):
    slug = title_to_slug(title)
    print(f"\nSlug: {slug}")

    driver.get("https://www.inview.nl/zoeken")

    # ⚡ Wait for search box to be ready (already logged in, so this is quick)
    box = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "input.wk-field-input[data-testid='search-bar-input-field']")
        )
    )

    driver.execute_script("arguments[0].scrollIntoView(true);", box)
    # ⚡ Removed unnecessary 1s sleep after scrollIntoView

    # ⚡ Use JS to set value instantly (no char-by-char typing delay)
    driver.execute_script(
        """
        var el = arguments[0];
        var val = arguments[1];
        el.focus();
        var nativeInputSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        nativeInputSetter.call(el, val);
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        """,
        box, title
    )
    time.sleep(0.3)  # brief pause so React processes the value

    try:
        search_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[data-e2e='cg-search-button']")
            )
        )
        search_btn.click()
    except Exception:
        box.send_keys(Keys.RETURN)

    # ⚡ Smart wait: wait for URL to change (confirms search submitted), then wait for results
    try:
        WebDriverWait(driver, 5).until(
            lambda d: "zoeken" not in d.current_url or "?" in d.current_url
        )
    except:
        pass  # if URL didn't change, still try to find clusters

    if captcha_present(driver):
        if not wait_for_captcha(driver):
            return None

    # ⚡ Wait for search clusters (wrapped so a timeout doesn't crash the title)
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.search-cluster"))
        )
    except Exception:
        print("\u26a0\ufe0f Timed out waiting for search clusters")
        return []

    clusters = driver.find_elements(By.CSS_SELECTOR, "div.search-cluster")

    commentaar_cluster = None
    for c in clusters:
        try:
            if c.get_attribute("data-e2e-cluster-name") == "Commentaar":
                commentaar_cluster = c
                break
        except:
            pass

    if not commentaar_cluster:
        print("No Commentaar cluster found")
        return []

    print("✓ Commentaar cluster found")

    items = commentaar_cluster.find_elements(
        By.CSS_SELECTOR,
        "[data-testid='search-results-item-title'] a"
    )

    links = []
    for item in items:
        try:
            href = item.get_attribute("href")
            if href:
                if href.startswith("/"):
                    href = "https://www.inview.nl" + href
                links.append(href)
        except:
            pass

    print(f"Total links found: {len(links)}")

    # 🔥 MATCHING WITH FIXED REDIRECT HANDLING
    matched_urls = []

    for idx, link in enumerate(links, 1):
        print(f"\n[{idx}] Opening: {link}")

        driver.get(link)

        final_url = wait_for_slug_in_url(driver, slug, timeout=15)

        if final_url:
            print("✓ MATCH FOUND AFTER REDIRECT")
            matched_urls.append(final_url)
            break
        else:
            print("✗ NO SLUG IN LINK (after waiting)")

    if matched_urls:
        print(f"\n✓ Total matched: {len(matched_urls)}")
        return matched_urls
    else:
        print("\n✗ No matching URLs found")
        return []


# ────────────────────────────────────────────────────────────
# EXCEL INPUT
# ────────────────────────────────────────────────────────────
def read_titles():
    if not os.path.exists(INPUT_XLSX):
        raise FileNotFoundError(f"Input file not found: {INPUT_XLSX}")
    df = pd.read_excel(INPUT_XLSX)
    if "primary_title" not in df.columns:
        raise ValueError(f"Column 'primary_title' not found. Available: {list(df.columns)}")
    return [str(x).strip() for x in df["primary_title"].dropna()]


# ────────────────────────────────────────────────────────────
# SAVE RESULTS
# ────────────────────────────────────────────────────────────
def save_results(titles, results):
    wb = Workbook()
    ws = wb.active

    ws.append(["#", "Title", "URL"])

    for i, t in enumerate(titles, 1):
        data = results.get(t)

        if data is None:
            ws.append([i, t, "SKIPPED"])
        elif len(data) == 0:
            ws.append([i, t, "NO MATCH"])
        else:
            ws.append([i, t, data[0]])

    wb.save(OUTPUT_XLSX)
    print(f"\nSaved → {OUTPUT_XLSX}")


# ────────────────────────────────────────────────────────────
# SAVE PROGRESS (after each title)
# ────────────────────────────────────────────────────────────
def save_progress(results):
    # ⚡ Atomic write: save to temp file first, then rename
    # Prevents progress.json corruption if the script crashes mid-write
    tmp = PROGRESS_FILE + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        os.replace(tmp, PROGRESS_FILE)  # atomic on Windows
    except Exception as e:
        print(f"\u26a0\ufe0f Warning: could not save progress: {e}")
        if os.path.exists(tmp):
            os.remove(tmp)


# ────────────────────────────────────────────────────────────
# MAIN
# ────────────────────────────────────────────────────────────
def main():
    titles = read_titles()

    # ⚡ Load existing progress so we can resume after restart
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            results = json.load(f)
        already_done = set(results.keys())
        print(f"⚡ Resuming: {len(already_done)} titles already processed, {len(titles) - len(already_done)} remaining")
    else:
        results = {}
        already_done = set()

    # Find first unprocessed title index for display
    remaining = [t for t in titles if t not in already_done]
    if not remaining:
        print("✅ All titles already processed! Saving Excel output...")
        save_results(titles, results)
        return

    driver = start_chrome()

    # ⚡ Handle cookie banner + login ONCE before processing any titles
    # No timeout — waits until the search box is visible (login fully complete)
    wait_for_login(driver)

    try:
        for i, title in enumerate(titles, 1):
            # ⚡ Skip already-processed titles (both found and not-found)
            if title in already_done:
                print(f"[{i}/{len(titles)}] ⏩ SKIP (already done): {title[:60]}...")
                continue

            print(f"\n[{i}/{len(titles)}] {title}")

            try:
                res = search_title(driver, title)
                results[title] = res
            except Exception as e:
                print("Error:", e)
                results[title] = None

            # ⚡ Save after every title so restart resumes from here
            save_progress(results)

            time.sleep(0.5)

    finally:
        try:
            driver.quit()
        except Exception:
            pass  # browser may already be closed

    # ⚡ Always save Excel at the end (even after partial run)
    try:
        save_results(titles, results)
    except Exception as e:
        print(f"\u26a0\ufe0f Could not save Excel: {e}")


if __name__ == "__main__":
    main()










