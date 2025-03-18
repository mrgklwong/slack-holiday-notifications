# slack-holiday-notification-using-excel

[![Daily Leave Report](https://github.com/mrgklwong/slack-holiday-notifications/actions/workflows/daily-leave-report.yml/badge.svg)](https://github.com/mrgklwong/slack-holiday-notifications/actions/workflows/daily-leave-report.yml)


# 🌴 **Daily Leave Report Bot** 🌴
**Automate leave tracking and send cheerful Slack notifications**

---

### ✨ **Features**
- 🛠️ Reads employee leave entries from an Excel file downloaded from MS Onedrive
- 📅 Checks for current/pending leaves using smart date comparisons
- 🎯 Compiles reports for today's active leaves
- 🎉 Sends super friendly Slack messages with tropical emojis 🌺
- 🚀 Works with common leave types (Annual, Sick Leave)

---

## 🚀 Getting Started

### 1. Prerequisites
Make sure you have:
- [Node.js](https://nodejs.org) installed
- A Slack Webhook URL (see [Slack API Setup](https://api.slack.com/messaging/webhooks))
- Password, Secret and File Location

### 2. Installation
```bash
# 1. Clone or download repository
git clone this repo

# 2. Install dependencies
npm ci
```
### 3. Github Actions Secrets
Please set up the followings secrets in your github actions
```

```


### 4. How to set up and update?
1. Update the excel file on Onedrive
2. ....
3. Profit

Github actions will run at 9am everyday. 
