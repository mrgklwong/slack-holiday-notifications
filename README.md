# slack-holiday-notification-using-excel

# 🌴 **Daily Leave Report Bot** 🌴
**Automate leave tracking and send cheerful Slack notifications**

---

### ✨ **Features**
- 🛠️ Reads employee leave entries from an Excel file
- 🎉 Sends super friendly Slack messages with tropical emojis 🌺
- 📅 Checks for current/pending leaves using smart date comparisons
- 🎯 Compiles reports for today's active leaves
- 🚀 Works with common leave types (Annual, Sick Leave)

---

## 🚀 Getting Started

### 1. Prerequisites
Make sure you have:
- [Node.js](https://nodejs.org) installed
- [ExcelJS](https://www.npmjs.com/package/exceljs) (installed via npm)
- A Slack Webhook URL (see [Slack API Setup](https://api.slack.com/messaging/webhooks))
- Onedrive excel URL

### 2. Installation
```bash
# 1. Clone or download repository
git clone this repo

# 2. Install dependencies
npm install exceljs axios date-fns

# 3. Setup environment variables:
cp .env.example .env

### 3. How to set up and update?
1. Download the excel
2. Name it leaves.xlsx drop it into your one drive (check the format)
3. Get the Shared link to the excel (save for later)
4. Get your slack api webhook (save for later)
5. Create github actions secerts (ONE_DRIVE_DOWNLOAD_URL and SLACK_WEBHOOK_URL)

Github actions will run at 9am everyday. 
