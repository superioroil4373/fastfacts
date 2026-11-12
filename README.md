# 🏆 Fast Facts — Automated Trivia Leaderboard

> **A robust Google Apps Script system designed to automate a trivia league leaderboard.** 
> Directly integrates with the StrawPoll API to fetch results, dynamically score participants based on rolling ranks, resolve player aliases, and automatically generate PDF leaderboards.

---

## ✨ Features

*   **📊 Automated API Scoring:** Pulls voting data directly from StrawPoll to eliminate manual data entry.
*   **📈 Dynamic Rank Multipliers:** Awards points based on a player's *current* rank on the board, ensuring underdogs catch up faster. First-to-answer correct submissions earn double points.
*   **🤝 Alias Resolution Engine:** Uses a dedicated `Aliases` tab and fuzzy-matching to ensure "Dave", "David", and "David S." all correctly apply to the same canonical leaderboard entry.
*   **🛡️ Deduplication:** Automatically ignores multiple submissions from the same user, counting only their earliest recorded vote.
*   **⏪ Historical Snapshots & Factory Resets:** Automatically backs up the leaderboard state prior to every poll. Includes a dynamic `recalculateAllPolls()` function capable of perfectly rebuilding the entire season's history from scratch in the event of an error.
*   **📄 Automated PDF Export:** Generates a beautifully formatted PDF of the updated leaderboard and saves it to Google Drive after every successful scoring run.

---

## 🚀 Setup Instructions

### 1️⃣ Google Sheets Preparation
Create a new Google Sheet and rename the first tab to your desired leaderboard name (e.g., `2026 Scores`). 

Open the Apps Script editor (`Extensions` > `Apps Script`) and paste the contents of `Code.gs`. Run the `setupSheets()` function once to automatically generate all the required tracking tabs:
*   `Config`: Where you paste the Poll IDs and declare the Correct Answer.
*   `Aliases`: Map player nicknames to their official leaderboard name.
*   `Poll Log`: An audit trail of every vote cast.
*   `Snapshots`: Automatically generated historical states.
*   `Script Log`: Developer diagnostic logs.

### 2️⃣ Environment Variables (Script Properties) — ⚠️ REQUIRED
For security, this script **does not** allow hardcoded API keys. If you paste your keys directly into the `Code.gs` file, you risk exposing them if you ever share your project. The script will intentionally throw an error and refuse to run until you set up secure Script Properties.

**To set these up:**
1. In the Apps Script Editor, go to **Project Settings** (the ⚙️ gear icon on the left menu).
2. Scroll down to **Script Properties** and click **Add script property**.
3. Add the following two properties exactly as written:
   *   **Property:** `STRAWPOLL_API_KEY` | **Value:** *Your StrawPoll V3 API Key*
   *   **Property:** `SHEET_ID` | **Value:** *The long string found in your Google Sheet's URL*
4. Click **Save script properties**.

### 3️⃣ Webhook Automation (Optional)
To have the script run instantly when a StrawPoll closes:
1. Deploy the script as a **Web App** (Execute as: `Me`, Access: `Anyone`).
2. Copy the resulting Web App URL.
3. In StrawPoll, set up a webhook that pings this URL upon the `poll.closed` event.

*(Alternatively, you can schedule a time-driven trigger in Apps Script to run `checkPolls()` daily).*

---

## 🛠️ Maintenance & Emergency Restore

If the leaderboard data ever becomes corrupted, you can easily trigger a factory reset:
1. Run `backupState()` from the Apps Script editor.
2. Run `recalculateAllPolls()`. 

The script will locate the oldest known historical snapshot, wipe the board, paste the historical baseline, and flawlessly recalculate every subsequent poll. If the recalculation fails, simply run `restoreState()` to undo the changes and revert to your backup.
