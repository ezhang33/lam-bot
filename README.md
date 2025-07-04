# Discord LAM Bot 🤖

An advanced Discord onboarding and server management bot that integrates with Google Sheets for Science Olympiad team management. Features automated member onboarding, dynamic channel creation, role management, and comprehensive server organization.

## 🆕 **Latest Features**
- 🔑 **Universal Slacker Access** - Slacker role gets access to ALL channels
- 📋 **Alphabetical Building Organization** - Building categories automatically sorted A-Z
- 👥 **Discord Handle Support** - Works with usernames, not just Discord IDs
- 🏠 **Static Channel Structure** - Welcome, Tournament Officials, and Volunteers categories
- 📋 **Forum Support** - Help channel as a forum for Q&A threads
- 🧨 **Server Reset** - Complete server reset functionality for testing
- 📝 **Smart Nickname Handling** - 32-character limit compliance

## Features ✨

### 📋 **Automated Onboarding**
- Monitors Google Sheets for new Discord users (supports IDs and handles)
- Sends personalized invite links to new members (lands in #welcome channel)
- Automatically assigns roles when users join
- Updates existing members with missing roles or incorrect nicknames

### 🎭 **Advanced Role Management**
- Assigns both "Master Role" (team role) and "First Event" (event role)
- Creates roles automatically with custom colors
- **Slacker role gets universal access** to all channels including restricted ones
- Role hierarchy management for proper permissions

### 📝 **Smart Nickname Management**
- Sets nicknames in format: `Name (First Event)` (truncated to 32 chars)
- Uses Google Sheet name if available, Discord username as fallback
- Updates nicknames for existing members automatically
- Handles Discord's nickname limitations gracefully

### 🏢 **Comprehensive Server Organization**

#### **Static Categories (Always Present):**
- 👋 **Welcome** - Public welcome channel for new members
- 📋 **Tournament Officials** - Private channels for Slacker role only:
  - `#slacker`, `#links`, `#scoring`, `#awards-ceremony`
- 🙋 **Volunteers** - Public channels for everyone:
  - `#general`, `#useful-links`, `#random`, `#help` (forum)

#### **Dynamic Building Structure:**
- 🏢 **Building categories** created from "Building 1" column (sorted alphabetically)
- 💬 **Building chat channels**: `[building]-chat` (restricted + Slacker access)
- 🎯 **Event-specific channels**: `[event]-[building]-[room]` (event role + Slacker access)
- 📋 **Automatic alphabetical sorting** of building categories

### 🔑 **Universal Slacker Access**
The Slacker role is special and gets access to:
- ✅ All Tournament Officials channels (private)
- ✅ All Volunteers channels (public)
- ✅ All building chat channels (restricted)
- ✅ All event-specific channels (restricted)
- ✅ Forum channels with full thread permissions

## Discord ID vs Handle Support 👥

The bot now supports multiple formats in the "Discord ID" column:

### ✅ **Supported Formats:**
- `1390416498077601822` (Discord ID - always works)
- `username#1234` (old Discord format - works for server members)
- `username` (new Discord format - works for server members)
- `Display Name` (server display name - works for server members)

### 📝 **Usage Notes:**
- **For existing server members**: Any format works
- **For new invites**: Discord ID required (can't search handles for non-members)
- **Automatic conversion**: Bot shows `🔍 Found user by handle 'username' -> ID: 123456789`

## Required Bot Permissions 🛡️

Your Discord bot needs these permissions:
- ✅ **Send Messages**
- ✅ **Manage Roles** (including role hierarchy placement)
- ✅ **Manage Nicknames** 
- ✅ **Manage Channels** (text, voice, forum, categories)
- ✅ **Create Instant Invite**
- ✅ **Read Message History**
- ✅ **Create Public Threads** (for forum access)
- ✅ **Send Messages in Threads**

### 🔒 **Important Permission Notes:**
- Bot role must be **higher than user roles** to manage nicknames
- Server owners **cannot** have nicknames changed by bots (Discord limitation)
- Enable **"Server Members Intent"** in Discord Developer Portal

## Google Sheets Format 📊

Your Google Sheet should have these columns:

| Column | Description | Example | Required |
|--------|-------------|---------|----------|
| `Discord ID` | Discord ID or handle | `ezhang3relace` or `1234567890` | ✅ Yes |
| `Name` | User's real name | `John Smith` | Optional |
| `Master Role` | Team role | `Volunteer`, `Slacker` | Optional |
| `First Event` | Event specialization | `Astronomy`, `Chemistry Lab` | Optional |
| `Building 1` | Building location | `Science Building`, `Main Hall` | Optional |
| `Room 1` | Room number/name | `101`, `Auditorium` | Optional |

## Installation 🚀

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Create environment file:**
   Create a `.env` file with your configuration:

3. **Configure your `.env` file:**
   ```env
   # Discord Configuration
   DISCORD_TOKEN=your_discord_bot_token_here
   GUILD_ID=your_discord_guild_id_here
   
   # Google Sheets Configuration
   GSPREAD_CREDS=gspread_creds/your_credentials_file.json
   SHEET_ID=your_google_spreadsheet_id_here
   SHEET_NAME=Sheet1
   
   # Optional Settings
   AUTO_CREATE_ROLES=true
   DEFAULT_ROLE_COLOR=light_gray
   
   # ⚠️ DANGER ZONE: Server Reset (use with caution!)
   RESET_SERVER=false
   ```

4. **Set up Google Sheets API:**
   - Enable Google Sheets API in Google Cloud Console
   - Create service account credentials
   - Share your Google Sheet with the service account email

5. **Run the bot:**
   ```bash
   python lam_bot.py
   ```

## Server Reset 🧨

**⚠️ DANGER ZONE**: Complete server reset functionality for testing/cleanup.

### To Enable:
```env
RESET_SERVER=true
```

### What Gets Reset:
- 🗑️ **All channels** (text, voice, forum) → Deleted
- 🗑️ **All categories** → Deleted  
- 🗑️ **All custom roles** → Deleted (keeps @everyone and bot roles)
- 📝 **All nicknames** → Reset to Discord usernames

### Safety Features:
- ⏰ 3-second warning with Ctrl+C to cancel
- 🛡️ Protects @everyone, bot roles, and bot nicknames
- 📊 Detailed progress logging and summary report
- 🔄 Automatically rebuilds structure after reset

## Color Scheme 🎨

### Team Roles (Custom Colors):
- **"Slacker"** → 🟠 Orange (universal access)
- **"Volunteer"** → 🔵 Blue
- **"Lead Event Supervisor"** → 🟡 Yellow
- **"Photographer"** → 🔴 Red
- **"Arbitrations"** → 🟢 Green
- **"Social Media"** → 🟣 Magenta

### Event Roles & Others:
- **All event roles** → ⚫ Light Gray (default)
- **Other roles** → ⚫ Light Gray (default)

## Server Structure Example 🏗️

```
👋 Welcome
   📺 #welcome (public - landing spot for new invites)

📋 Tournament Officials (🔒 Slacker only)
   📺 #slacker
   📺 #links  
   📺 #scoring
   📺 #awards-ceremony

🙋 Volunteers (public)
   📺 #general
   📺 #useful-links
   📺 #random
   📋 #help (forum - for Q&A threads)

🏢 Biology Lab (alphabetical order)
   🔒 #biology-lab-chat (restricted + Slacker)
   🔒 #ecology-biology-lab-room101 (Ecology role + Slacker)
   🔒 #botany-biology-lab-greenhouse (Botany role + Slacker)

🏢 Chemistry Building  
   🔒 #chemistry-building-chat (restricted + Slacker)
   🔒 #chemistry-lab-chemistry-building-205 (Chemistry Lab role + Slacker)

🏢 Physics Hall
   🔒 #physics-hall-chat (restricted + Slacker)
   🔒 #astronomy-physics-hall-observatory (Astronomy role + Slacker)
```

## How It Works 🔄

### **On Bot Startup:**
1. 🧨 **Server reset** (if enabled)
2. 🏗️ **Creates static structure** (Welcome, Tournament Officials, Volunteers)
3. 📋 **Organizes categories alphabetically** (static first, then buildings A-Z)
4. 🔑 **Grants Slacker universal access** to all existing channels
5. ⏰ **Starts member sync task** (runs every minute)

### **Every Minute (Sync Task):**
1. 📊 **Reads Google Sheet** and processes all rows
2. **For each user:**
   - 🔍 **Resolves Discord ID** (from ID or handle)
   - **If not in server**: Sends invite DM (to #welcome) and queues roles
   - **If in server**: Assigns missing roles, updates nickname, creates channels
3. 🏢 **Creates building structure** as needed
4. 📋 **Maintains alphabetical order** of building categories

### **When Users Join:**
1. ✅ **Assigns queued roles** automatically  
2. 📝 **Sets nickname**: `Name (First Event)` (32 char limit)
3. 🔑 **Gets access** to their event-specific channels

## Console Output Examples 📺

### **Normal Operation:**
```bash
🔄 Running member sync...
📊 Found 25 rows in spreadsheet
👥 Guild has 12 members
🔍 Found user by handle 'john_doe' -> ID: 1234567890
✅ Assigned role Volunteer to John
✅ Assigned role Astronomy to John  
📝 Updated nickname for John: 'John Smith (Astronomy)'
🏢 Created category: 'Science Building'
📺 Created building chat: '#science-building-chat' (restricted)
🔒 Added Astronomy access to #science-building-chat
🔑 Added Slacker access to #science-building-chat
📺 Created channel: '#astronomy-science-building-room101' (restricted to Astronomy)
🔑 Added Slacker access to #astronomy-science-building-room101
📋 Categories organized: Static categories first, then buildings alphabetically
✅ Sync complete. Processed 23 valid Discord IDs from 25 rows.
```

### **Server Reset:**
```bash
⚠️ ⚠️ ⚠️  STARTING COMPLETE SERVER RESET  ⚠️ ⚠️ ⚠️
🧨 This will delete EVERYTHING and reset all nicknames!
⏰ Starting in 3 seconds... (Ctrl+C to cancel)
🗑️ Starting server reset...
📝 Reset nickname for John Smith
🗑️ Deleted text channel: #general
🗑️ Deleted category: Science Building  
🗑️ Deleted role: Volunteer
🧨 SERVER RESET COMPLETE!
📊 Summary:
   • 5 nicknames reset
   • 12 text channels deleted  
   • 4 categories deleted
   • 8 roles deleted
🏗️ Server is now completely clean and ready for fresh setup!
```

## Troubleshooting 🔧

### **Permission Issues:**
- ❌ **"No permission to set nickname"** → Check "Manage Nicknames" + role hierarchy
- ❌ **"No permission to create role"** → Add "Manage Roles" permission
- ❌ **"No permission to create channel"** → Add "Manage Channels" permission
- ❌ **"Privileged intents required"** → Enable "Server Members Intent" in Discord Developer Portal

### **Google Sheets Issues:**
- ❌ **"Sheet not found"** → Check `SHEET_NAME` in .env file
- ❌ **"Permission denied"** → Enable Google Sheets API and share sheet with service account
- ❌ **"Credentials error"** → Check `GSPREAD_CREDS` path and file validity

### **Handle/ID Issues:**
- ⚠️ **"Could not find user with handle"** → Use Discord ID for non-server members
- ⚠️ **"User not found"** → Verify Discord ID is correct
- 🔍 **Handle resolution working** → Look for "Found user by handle" messages

### **Bot Not Responding:**
1. Check console for error messages
2. Verify bot is online in Discord  
3. Ensure bot has required permissions
4. Check environment variables are loaded correctly

## Advanced Configuration ⚙️

### **Environment Variables:**
```env
# Required
DISCORD_TOKEN=your_bot_token
GUILD_ID=123456789012345678
GSPREAD_CREDS=path/to/credentials.json
SHEET_ID=your_sheet_id

# Optional
SHEET_NAME=Sheet1                    # Default: "Sheet1"
AUTO_CREATE_ROLES=true               # Default: true
DEFAULT_ROLE_COLOR=light_gray        # Default: light_gray

# Danger Zone
RESET_SERVER=false                   # Default: false
```

### **Color Options:**
`blue`, `red`, `green`, `purple`, `orange`, `yellow`, `teal`, `pink`, `light_gray`, `dark_gray`, `black`, `white`

---

## Support 💡

### **Getting Help:**
1. 📺 **Check console output** for detailed error messages and progress
2. 🔒 **Verify bot permissions** in Discord server settings  
3. ✅ **Test Google Sheets access** and API credentials
4. 📋 **Ensure required columns** exist in your spreadsheet
5. 🔍 **Use Discord handles** for easier user identification

### **Common Solutions:**
- **Bot stuck**: Check "Server Members Intent" is enabled
- **Roles not assigning**: Verify bot role hierarchy
- **Channels not creating**: Check "Manage Channels" permission
- **Forum not working**: Bot will provide manual creation instructions

---

**🏆 Built for Science Olympiad team management with enterprise-grade automation**

*Features universal access control, dynamic organization, and comprehensive server management*