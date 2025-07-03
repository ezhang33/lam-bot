#!/usr/bin/env python3
"""
Quick setup script for the Discord bot with Google Sheets integration
"""

import os
import sys

def create_env_file():
    """Create a .env file with placeholder values"""
    env_content = """# Discord Bot Configuration
DISCORD_TOKEN=your_discord_bot_token_here

# Google Sheets Configuration
GOOGLE_SPREADSHEET_ID=your_google_spreadsheet_id_here

# Option 1: Google Service Account JSON as environment variable
# GOOGLE_SERVICE_ACCOUNT_JSON={"type": "service_account", "project_id": "your-project", ...}

# Option 2: Use service_account_key.json file instead (place it in the project root)
"""
    
    if not os.path.exists('.env'):
        with open('.env', 'w') as f:
            f.write(env_content)
        print("✅ Created .env file with placeholder values")
    else:
        print("⚠️  .env file already exists")

def check_dependencies():
    """Check if required dependencies are installed"""
    try:
        import discord
        import googleapiclient
        import google.oauth2
        print("✅ All required dependencies are installed")
        return True
    except ImportError as e:
        print(f"❌ Missing dependency: {e}")
        print("Please run: pip install -r requirements.txt")
        return False

def main():
    print("🤖 Discord Bot with Google Sheets - Quick Setup")
    print("=" * 50)
    
    # Check Python version
    if sys.version_info < (3, 8):
        print("❌ Python 3.8+ is required")
        sys.exit(1)
    
    print(f"✅ Python {sys.version_info.major}.{sys.version_info.minor} detected")
    
    # Create .env file
    create_env_file()
    
    # Check dependencies
    if not check_dependencies():
        print("\n📦 To install dependencies, run:")
        print("pip install -r requirements.txt")
        return
    
    print("\n🚀 Setup complete! Next steps:")
    print("1. Edit the .env file with your Discord bot token and Google Sheets ID")
    print("2. Set up Google Sheets API credentials (see README.md)")
    print("3. Run the bot with: python bot.py")
    print("\n📚 For detailed instructions, see README.md")

if __name__ == "__main__":
    main() 