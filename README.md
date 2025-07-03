# Discord Bot with Google Sheets Integration

A Discord bot that can read data from Google Sheets spreadsheets and respond to commands in your Discord server.

## Features

- 📊 Read data from Google Sheets
- 🔍 Search for specific terms in spreadsheets
- 📋 Get spreadsheet information
- 🎯 Support for custom ranges and multiple sheets
- 💬 Beautiful Discord embeds for data display

## Commands

- `!read_sheet [range]` - Read data from Google Sheets (default: Sheet1!A1:Z100)
- `!sheet_info` - Get information about the configured Google Sheet
- `!search_sheet "term" [range]` - Search for a specific term in the sheet
- `!help_sheets` - Show all available commands

## Setup Instructions

### 1. Discord Bot Setup

1. Go to [Discord Developer Portal](https://discord.com/developers/applications)
2. Create a new application and bot
3. Copy the bot token
4. Invite the bot to your server with appropriate permissions

### 2. Google Sheets API Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Google Sheets API
4. Create a service account:
   - Go to "Credentials" → "Create Credentials" → "Service Account"
   - Download the JSON key file
   - Rename it to `service_account_key.json` and place it in the project root
5. Share your Google Sheet with the service account email address (give it "Viewer" permissions)

### 3. Installation

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Set up environment variables:
   ```bash
   cp .env.example .env
   ```
   Edit `.env` and add your tokens:
   ```
   DISCORD_TOKEN=your_discord_bot_token_here
   GOOGLE_SPREADSHEET_ID=your_google_spreadsheet_id_here
   ```

### 4. Getting Your Spreadsheet ID

The spreadsheet ID is found in your Google Sheets URL:
```
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit#gid=0
```

### 5. Run the Bot

```bash
python bot.py
```

## Authentication Options

### Option 1: Service Account Key File (Recommended)
Place your `service_account_key.json` file in the project root.

### Option 2: Environment Variable
Set the `GOOGLE_SERVICE_ACCOUNT_JSON` environment variable with the entire JSON content.

## Usage Examples

```
!read_sheet                    # Read default range (Sheet1!A1:Z100)
!read_sheet Sheet1!A1:C10     # Read specific range
!search_sheet "john doe"       # Search for "john doe" in default range
!search_sheet "sales" Sheet1!A1:D50  # Search in specific range
!sheet_info                    # Get spreadsheet information
```

## Error Handling

The bot includes comprehensive error handling for:
- Google Sheets API errors
- Discord API errors
- Authentication issues
- Data formatting problems

## Security Notes

- Never commit your `.env` file or `service_account_key.json` to version control
- Use environment variables in production
- Ensure your Google Sheet permissions are properly configured
- Only share your spreadsheet with the service account email

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is open source and available under the MIT License.