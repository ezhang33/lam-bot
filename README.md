# LAM Bot 🤖

A Discord bot that automatically manages tournament servers using Google Sheets. Perfect for Science Olympiad and other competitive events.

## What It Does ✨

- **Reads your Google Sheet** → Automatically invites users and assigns roles
- **Creates channels** → Dynamic building/event channels based on your data
- **Manages permissions** → Role-based access to channels
- **Handles help tickets** → Zone-based slacker assignment system
- **Building welcome messages** → Shows all events in each building

## Quick Setup 🚀

### 1. Install

```bash
pip install -r requirements.txt
```

### 2. Create `.env` file

```env
DISCORD_TOKEN=your_bot_token
GUILD_ID=your_server_id
GSPREAD_CREDS=path/to/credentials.json
RESET_SERVER=false
```

### 3. Google Sheets Setup

1. Create a Google Cloud project
2. Enable Google Sheets API
3. Create service account → Download JSON credentials
4. Share your sheet with the service account email

### 4. Run

```bash
python lam_bot.py
```

## Sheet Format 📊

| Discord ID | Name | Master Role | First Event | Building 1   | Room 1 |
| ---------- | ---- | ----------- | ----------- | ------------ | ------ |
| 123456789  | John | Volunteer   | Astronomy   | Science Hall | 101    |
| jane_doe   | Jane | Slacker     | Chemistry   | Lab Building | 205    |

## Commands 💬

- `/entertemplate <sheet_link>` - Connect to your Google Sheet
- `/sync` - Manually sync users (Admin only)
- `/login` - Get roles by entering your email
- `/help` - Show all commands

## Bot Permissions Needed 🔐

- Manage Roles
- Manage Channels
- Manage Nicknames
- Send Messages
- Create Invites
- Read Message History

**Important:** Enable "Server Members Intent" in Discord Developer Portal

## How It Works 🔄

1. Bot reads your Google Sheet every minute
2. Invites new users and assigns roles automatically
3. Creates building categories and event channels
4. Sets up help ticket system with zone assignments
5. Manages permissions based on roles

## Support 💡

Check the console output for detailed logs. Most issues are permission-related - make sure the bot role is high in your server's role hierarchy.

---

_Built for tournament management with automated onboarding and organization_
