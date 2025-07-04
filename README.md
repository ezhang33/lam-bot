# Discord LAM Bot 🤖

An automated Discord onboarding bot that integrates with Google Sheets for Science Olympiad team management.

## Features ✨

### 📋 **Automated Onboarding**
- Monitors Google Sheets for new Discord user IDs
- Sends personalized invite links to new members
- Automatically assigns roles when users join

### 🎭 **Dual Role Assignment**
- Assigns both "Master Role" (team role) and "First Event" (event role)
- Creates roles automatically if they don't exist
- Custom color mapping for team roles

### 📝 **Nickname Management**
- Sets nicknames in format: `Name (First Event)`
- Uses Google Sheet name if available, Discord username as fallback
- Updates nicknames for existing members

### 🏢 **Building & Channel Organization**
- Creates categories for each building from "Building 1" column
- Creates building chat channels: `[building]-chat` (restricted to people with events in that building)
- Creates event-specific channels: `[event]-[building]-[room]` (restricted to event role members)
- Automatically manages permissions so users only see relevant channels

## Required Bot Permissions 🛡️

Your Discord bot needs these permissions:
- ✅ **Send Messages**
- ✅ **Manage Roles**
- ✅ **Manage Nicknames**
- ✅ **Manage Channels**
- ✅ **Create Instant Invite**
- ✅ **Read Message History**

## Google Sheets Format 📊

Your Google Sheet should have these columns:
- `Discord ID` - User's Discord ID (required)
- `Name` - User's real name (optional, uses Discord username if empty)
- `Master Role` - Team role (e.g., "Volunteer", "Slacker")
- `First Event` - Event specialization (e.g., "Astronomy", "Chemistry Lab")
- `Building 1` - Building location (e.g., "Science Building", "Main Hall")
- `Room 1` - Room number/name (optional)

## Installation 🚀

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Set up environment variables:**
   ```bash
   cp env_template .env
   ```

3. **Configure your `.env` file:**
   ```env
   DISCORD_TOKEN=your_discord_bot_token_here
   GSPREAD_CREDS=path/to/service_account_credentials.json
   SHEET_ID=your_google_spreadsheet_id_here
   SHEET_NAME=Sheet1
   GUILD_ID=your_discord_guild_id_here
   AUTO_CREATE_ROLES=true
   DEFAULT_ROLE_COLOR=light_gray
   ```

4. **Run the bot:**
   ```bash
   python lam_bot.py
   ```

## Color Scheme 🎨

### Team Roles (Custom Colors):
- **"Slacker"** → 🟡 Yellow
- **"Volunteer"** → 🔵 Blue
- **"Lead Event Supervisor"** → 🟢 Green
- **"Photographer"** → 🔴 Red

### Event Roles & Others:
- **All event roles** → ⚫ Gray
- **Other roles** → ⚫ Gray (default)

## Example Server Structure 🏗️

```
📁 Science Building
   🔒 science-building-chat (only people with events in Science Building)
   🔒 astronomy-science-building-room101 (only Astronomy role)
   🔒 chemistry-lab-science-building-room205 (only Chemistry Lab role)

📁 Main Hall
   🔒 main-hall-chat (only people with events in Main Hall)
   🔒 forensics-main-hall-auditorium (only Forensics role)
   🔒 write-it-do-it-main-hall-room303 (only Write It Do It role)
```

## How It Works 🔄

1. **Every minute**, the bot reads your Google Sheet
2. **For each user:**
   - If not in server: Sends invite DM and queues roles
   - If in server: Assigns missing roles and updates nickname
   - Creates building structure and channels as needed
3. **When users join:**
   - Assigns queued roles automatically
   - Sets nickname: `Name (First Event)`
   - Gets access to their event-specific channels

## Troubleshooting 🔧

### Common Issues:
- **"No permission to create role"** → Add "Manage Roles" permission
- **"No permission to set nickname"** → Add "Manage Nicknames" permission
- **"No permission to create channel"** → Add "Manage Channels" permission
- **"Sheet not found"** → Check SHEET_NAME in .env file

### Console Output:
The bot provides detailed logging:
```
✅ Assigned role Volunteer to John
📝 Set nickname for John: 'John Smith (Astronomy)'
🏢 Created category: 'Science Building'
📺 Created building chat: '#science-building-chat' (restricted)
🔒 Added Astronomy access to #science-building-chat
📺 Created channel: '#astronomy-science-building-room101' (restricted to Astronomy)
```

## Support 💡

For issues or questions:
1. Check the console output for error messages
2. Verify bot permissions in Discord
3. Ensure Google Sheets credentials are correct
4. Check that all required columns exist in your sheet

---

**Built for Science Olympiad team management** 🏆