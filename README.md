# 📚 Groene Serie — Title to URL Mapper

Automated tool that reads legal document titles from an Excel file and finds their exact URLs on [inview.nl](https://www.inview.nl) using Selenium. Supports **resume on restart**, **CAPTCHA detection**, and **crash-safe progress saving**.

---

## ✨ Features

| Feature | Details |
|---|---|
| ⚡ **Fast input** | Fills the search box instantly via JavaScript (no character-by-character typing) |
| ⏩ **Resume on restart** | Skips already-processed titles using `progress.json` — picks up exactly where it left off |
| 🍪 **Cookie & login aware** | Waits indefinitely for you to accept the cookie banner and log in before starting |
| 🛡️ **CAPTCHA detection** | Pauses and waits for you to solve CAPTCHAs manually, then continues automatically |
| 💾 **Crash-safe saving** | Progress saved after every title using atomic writes — JSON never gets corrupted |
| 🚫 **Blocks images/fonts** | Chrome configured to skip loading images, fonts, and media for faster page loads |
| 📊 **Excel output** | Final results saved as `.xlsx` with title, matched URL, or status (`NO MATCH` / `SKIPPED`) |

---

## 📋 Requirements

### Python
- Python 3.8 or higher

### Dependencies
Install all required packages:
```bash
pip install selenium pandas openpyxl
```

### ChromeDriver
- Download [ChromeDriver](https://chromedriver.chromium.org/downloads) matching your Chrome version
- Place `chromedriver.exe` in the **same folder** as `app.py`

---

## 📁 File Structure

```
Title_to_link/
│
├── app.py                              ← Main script
├── chromedriver.exe                    ← ChromeDriver (you provide this)
│
├── Groene Serie Rechtspersonen.xlsx    ← INPUT: titles to search
├── Result_Groene Serie Rechtspersonen.xlsx  ← OUTPUT: titles + URLs (auto-created)
│
├── progress.json                       ← Auto-created: resume state
└── README.md                           ← This file
```

---

## ⚙️ Configuration

Edit the settings at the top of `app.py`:

```python
INPUT_XLSX      = "Groene Serie Rechtspersonen.xlsx"   # Your input Excel file
OUTPUT_XLSX     = "Result_Groene Serie Rechtspersonen.xlsx"  # Output file name
PROGRESS_FILE   = "progress.json"                      # Resume file (auto-managed)
CAPTCHA_TIMEOUT = 60                                   # Seconds to wait for CAPTCHA solve
```

### Input Excel Format
Your Excel file must have a column named **`primary_title`** containing the titles to search:

| primary_title |
|---|
| Groene Serie Rechtspersonen, A Kernoverzicht bij: Burgerlijk Wetboek Boek 2, Artikel 225 |
| Groene Serie Rechtspersonen, 1 Wetsgeschiedenis bij: ... |
| ... |

---

## 🚀 How to Run

```bash
cd d:\Legsiys\Title_to_link
python app.py
```

### What happens step by step:

1. **Progress loaded** — any titles already in `progress.json` are skipped automatically
2. **Chrome opens** — the browser navigates to inview.nl
3. **Waiting for login** — the terminal shows:
   ```
   ============================================================
     Opening inview.nl — please:
     1. Accept the cookie banner
     2. Log in if prompted
     The script will continue automatically once logged in.
   ============================================================
   ```
4. **You act in the browser** — accept cookies, click login, fill the form
5. **Auto-detected** — as soon as the search box is visible, the script prints `✅ Logged in!` and starts
6. **Titles processed** — each title is searched, the matching URL is found and saved
7. **Results exported** — Excel file written at the end

---

## 🔄 Resuming After a Restart

Simply run `python app.py` again. The script will:
- Load `progress.json`
- Print how many titles are already done and how many remain
- Skip done titles instantly with `⏩ SKIP`
- Ask you to log in once, then continue from where it stopped

> **Note:** Titles saved as `null` in `progress.json` were tried and failed (no match found). They are also skipped on restart. To retry them, remove those entries from `progress.json` manually.

---

## 📊 Output Format

The result Excel file has 3 columns:

| # | Title | URL |
|---|---|---|
| 1 | Groene Serie ... Artikel 225 | `https://www.inview.nl/document/...` |
| 2 | Groene Serie ... Artikel 224 | NO MATCH |
| 3 | Groene Serie ... Artikel 223 | SKIPPED |

| Status | Meaning |
|---|---|
| URL | Match found successfully |
| `NO MATCH` | Searched but no matching URL found |
| `SKIPPED` | Error occurred during processing |

---

## ⚠️ Troubleshooting

### ChromeDriver version mismatch
```
SessionNotCreatedException: Message: session not created...
```
→ Download the ChromeDriver version that matches your Chrome. Check Chrome version at `chrome://settings/help`.

### Column not found error
```
ValueError: Column 'primary_title' not found. Available: ['Title', ...]
```
→ Rename the title column in your Excel to `primary_title`.

### CAPTCHA loop
→ The script pauses and prints `⚠️ CAPTCHA detected. Solve it manually...`. Solve it in the browser — the script continues automatically.

### Script crashed mid-run
→ Just run again. `progress.json` is saved atomically after every title, so at most one title needs to be retried.

### First title always fails
→ This was a known bug (fixed). Make sure you are using the latest version of `app.py`. The login step now happens **before** the first title is processed.

---

## 🔧 How the Search Works

For each title the script:
1. Navigates to `https://www.inview.nl/zoeken`
2. Fills the search box instantly via JavaScript
3. Clicks the search button (falls back to `ENTER` if button not found)
4. Waits for the `Commentaar` result cluster to appear
5. Opens each result link and checks if the slug from the title appears in the final URL
6. Saves the first matching URL and moves on

---

## 📝 Notes

- The script blocks images, fonts, and media in Chrome to speed up page loads
- Progress is saved **after every single title** — you can safely close the terminal at any time
- The `progress.json` file is your source of truth — do not delete it mid-run
- To start completely fresh, delete `progress.json` before running

---

## 📄 License

Internal tool — Legsiys.
