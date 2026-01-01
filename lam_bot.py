# lam_bot.py
import os
import asyncio
import discord
from discord.ext import commands, tasks
from discord import app_commands
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv
from datetime import datetime, timedelta
import json
import base64

load_dotenv()

TOKEN         = os.getenv("DISCORD_TOKEN")
SERVICE_EMAIL = os.getenv("SERVICE_EMAIL")
GSPCREDS      = os.getenv("GSPREAD_CREDS")
GSPCREDS      = json.loads(base64.b64decode(GSPCREDS).decode('utf-8'))
# GSPCREDS      = "/etc/secrets/gspread_creds.json"
SHEET_ID      = os.getenv("SHEET_ID")  # Optional - can be set via /entertemplate command
SHEET_FILE_NAME = os.getenv("SHEET_FILE_NAME", "[TEMPLATE] Socal State")  # Name of the Google Sheet file to look for
SHEET_PAGE_NAME = os.getenv("SHEET_PAGE_NAME", "Sheet1")  # Name of the worksheet/tab within the sheet
AUTO_CREATE_ROLES = os.getenv("AUTO_CREATE_ROLES", "true").lower() == "true"
DEFAULT_ROLE_COLOR = os.getenv("DEFAULT_ROLE_COLOR", "light_gray")  # blue, red, green, purple, etc.

# ⚠️ ⚠️ ⚠️  DANGER ZONE: COMPLETE SERVER RESET  ⚠️ ⚠️ ⚠️
# Set to True to COMPLETELY RESET the server on bot startup
# WARNING: This will permanently delete ALL channels, categories, roles, and reset all nicknames!
# This is IRREVERSIBLE! Use only for testing or complete server reset!
RESET_SERVER = os.getenv("RESET_SERVER", "false").lower() == "true"

intents = discord.Intents.default()
intents.members = True
intents.message_content = True

class LamBot(commands.Bot):
    def __init__(self):
        super().__init__(command_prefix='!', intents=intents)

bot = LamBot()

# Set up gspread client
scope = [
    "https://www.googleapis.com/auth/spreadsheets",  # Full spreadsheet access (read & write)
    "https://www.googleapis.com/auth/drive.readonly"  # Needed to search for sheets
]
creds = ServiceAccountCredentials.from_json_keyfile_name(GSPCREDS, scope)
gc = gspread.authorize(creds)

# Sheet connection is now handled dynamically via /entertemplate command only
sheet = None
spreadsheet = None

# Note: SHEET_ID is still available as environment variable but won't auto-connect
# Use /entertemplate command to connect to sheets dynamically
print("📋 Bot starting - will attempt to load cached sheet connection or use /entertemplate command")

# Store pending role assignments and user info for users who haven't joined yet
pending_users = {}  # Changed from pending_roles to store more info

# Track chapter role names globally
chapter_role_names = set()

# Track active help tickets for re-pinging
active_help_tickets = {}  # thread_id -> ticket_info

# Cache configuration
CACHE_FILE = "bot_cache.json"

def save_cache(data):
    """Save cache data to JSON file"""
    try:
        with open(CACHE_FILE, 'w') as f:
            json.dump(data, f, indent=2)
        print(f"✅ Saved cache to {CACHE_FILE}")
    except Exception as e:
        print(f"❌ Error saving cache: {e}")

def load_cache():
    """Load cache data from JSON file"""
    try:
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, 'r') as f:
                data = json.load(f)
            print(f"✅ Loaded cache from {CACHE_FILE}")
            return data
        else:
            print(f"📄 No cache file found at {CACHE_FILE}")
            return {}
    except Exception as e:
        print(f"❌ Error loading cache: {e}")
        return {}

def clear_cache():
    """Clear the cache file"""
    try:
        if os.path.exists(CACHE_FILE):
            os.remove(CACHE_FILE)
            print(f"🗑️ Cleared cache file {CACHE_FILE}")
        else:
            print(f"📄 No cache file to clear")
    except Exception as e:
        print(f"❌ Error clearing cache: {e}")

async def load_spreadsheet_from_cache():
    """Try to load spreadsheet connection from cache"""
    global sheet, spreadsheet
    
    cache = load_cache()
    spreadsheet_id = cache.get("spreadsheet_id")
    worksheet_name = cache.get("worksheet_name", SHEET_PAGE_NAME)
    
    if not spreadsheet_id:
        print("📋 No cached spreadsheet connection found")
        return False
    
    try:
        print(f"🔄 Attempting to connect to cached spreadsheet: {spreadsheet_id}")
        spreadsheet = gc.open_by_key(spreadsheet_id)
        sheet = spreadsheet.worksheet(worksheet_name)
        
        # Test the connection by getting the first row
        headers = sheet.row_values(1)
        print(f"✅ Successfully connected to cached spreadsheet: '{spreadsheet.title}'")
        print(f"📊 Worksheet: '{sheet.title}' with {len(headers)} columns")
        return True
        
    except Exception as e:
        print(f"❌ Failed to connect to cached spreadsheet: {e}")
        print("🧹 Clearing invalid cache...")
        clear_cache()
        return False

async def handle_rate_limit(coro, operation_name, max_retries=3, default_delay=0.1):
    """
    Helper function to handle rate limits for Discord API calls.
    
    Args:
        coro: Coroutine to execute
        operation_name: Name of the operation for logging
        max_retries: Maximum number of retries (default: 3)
        default_delay: Default delay after successful operation in seconds (default: 0.1)
    
    Returns:
        Result of the coroutine, or None if all retries failed
    """
    retry_count = 0
    while retry_count < max_retries:
        try:
            result = await coro
            # Small delay after successful operation to avoid rate limits
            await asyncio.sleep(default_delay)
            return result
        except discord.HTTPException as e:
            error_msg = str(e)
            if e.status == 429 or "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                retry_count += 1
                if retry_count < max_retries:
                    # Extract retry_after from response if available
                    retry_after = 1.0  # Default 1 second
                    if hasattr(e, 'retry_after') and e.retry_after:
                        retry_after = float(e.retry_after)
                    elif isinstance(e.response, dict) and 'retry_after' in e.response:
                        retry_after = float(e.response['retry_after'])
                    
                    print(f"⚠️ Rate limited on {operation_name}, waiting {retry_after}s before retry {retry_count}/{max_retries}...")
                    await asyncio.sleep(retry_after)
                else:
                    print(f"❌ Rate limited on {operation_name} after {max_retries} retries, giving up")
                    return None
            else:
                # Re-raise non-rate-limit errors
                raise
        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                retry_count += 1
                if retry_count < max_retries:
                    print(f"⚠️ Rate limited on {operation_name}, waiting 1s before retry {retry_count}/{max_retries}...")
                    await asyncio.sleep(1.0)
                else:
                    print(f"❌ Rate limited on {operation_name} after {max_retries} retries, giving up")
                    return None
            else:
                # Re-raise non-rate-limit errors
                raise
    
    return None

async def get_or_create_role(guild, role_name):
    """Get a role by name, or create it if it doesn't exist"""
    role = discord.utils.get(guild.roles, name=role_name)
    if role:
        return role
    
    # Check if auto-creation is enabled
    if not AUTO_CREATE_ROLES:
        print(f"⚠️ Role '{role_name}' not found and auto-creation is disabled")
        return None
    
    # Role doesn't exist, create it
    try:
        # Special case: :( role gets full permissions
        if role_name == ":(":
            max_retries = 3
            retry_count = 0
            while retry_count < max_retries:
                try:
                    role = await guild.create_role(
                        name=":(",
                        permissions=discord.Permissions.all(),
                        color=discord.Color.purple(),
                        reason="Auto-created :( role for ezhang."
                    )
                    print(f"🆕 Created :( role with full permissions")
                    await asyncio.sleep(0.1)  # Small delay to avoid rate limits
                    return role
                except discord.HTTPException as e:
                    error_msg = str(e)
                    if e.status == 429 or "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                        retry_count += 1
                        if retry_count < max_retries:
                            retry_after = 1.0
                            if hasattr(e, 'retry_after') and e.retry_after:
                                retry_after = float(e.retry_after)
                            print(f"⚠️ Rate limited creating :( role, waiting {retry_after}s before retry {retry_count}/{max_retries}...")
                            await asyncio.sleep(retry_after)
                        else:
                            print(f"❌ Rate limited creating :( role after {max_retries} retries, giving up")
                            return None
                    else:
                        raise  # Re-raise non-rate-limit errors
                except Exception as e:
                    error_msg = str(e)
                    if "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                        retry_count += 1
                        if retry_count < max_retries:
                            print(f"⚠️ Rate limited creating :( role, waiting 1s before retry {retry_count}/{max_retries}...")
                            await asyncio.sleep(1.0)
                        else:
                            print(f"❌ Rate limited creating :( role after {max_retries} retries, giving up")
                            return None
                    else:
                        raise  # Re-raise non-rate-limit errors
            return None
        
        # Custom color mapping for specific roles
        custom_role_colors = {
            # Team roles only
            "Slacker": discord.Color.orange(),
            "Awards": discord.Color.yellow(),
            "Volunteer": discord.Color.blue(),
            "Lead Event Supervisor": discord.Color.yellow(),
            "Photographer": discord.Color.red(),
            "Arbitrations": discord.Color.green(),
            "Social Media": discord.Color.magenta(),
            "VIPer": discord.Color.green(),
        }
        
        # Check if this role has a custom color
        if role_name in custom_role_colors:
            role_color = custom_role_colors[role_name]
            color_name = {
                discord.Color.yellow(): "yellow",
                discord.Color.blue(): "blue", 
                discord.Color.green(): "green",
                discord.Color.red(): "red"
            }.get(role_color, "custom")
        elif role_name == "Unaffiliated" or role_name in chapter_role_names:
            # Chapter roles should be green
            role_color = discord.Color.green()
            color_name = "green"
        else:
            # Use default color mapping
            color_map = {
                "blue": discord.Color.blue(),
                "red": discord.Color.red(),
                "green": discord.Color.green(),
                "purple": discord.Color.purple(),
                "orange": discord.Color.orange(),
                "yellow": discord.Color.yellow(),
                "teal": discord.Color.teal(),
                "pink": discord.Color.magenta(),
                "light_gray": discord.Color.light_gray(),
                "dark_gray": discord.Color.dark_gray(),
                "black": discord.Color.from_rgb(0, 0, 0),
                "white": discord.Color.from_rgb(255, 255, 255)
            }
            
            role_color = color_map.get(DEFAULT_ROLE_COLOR.lower(), discord.Color.light_gray())
            color_name = DEFAULT_ROLE_COLOR
        
        max_retries = 3
        retry_count = 0
        while retry_count < max_retries:
            try:
                role = await guild.create_role(
                    name=role_name,
                    color=role_color,
                    reason="Auto-created by LAM Bot for onboarding"
                )
                print(f"🆕 Created new role: '{role_name}' (color: {color_name})")
                
                # If we just created the Slacker role, ensure it has access to Tournament Officials channels
                if role_name == "Slacker":
                    await ensure_slacker_tournament_officials_access(guild, role)
                
                # Note: Test folder search is now handled in setup_building_structure after channels are created
                # to ensure the target channel exists when we try to post the message
                
                # Small delay after creating role to avoid rate limits
                await asyncio.sleep(0.1)
                return role
            except discord.HTTPException as e:
                error_msg = str(e)
                if e.status == 429 or "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                    retry_count += 1
                    if retry_count < max_retries:
                        retry_after = 1.0
                        if hasattr(e, 'retry_after') and e.retry_after:
                            retry_after = float(e.retry_after)
                        print(f"⚠️ Rate limited creating role '{role_name}', waiting {retry_after}s before retry {retry_count}/{max_retries}...")
                        await asyncio.sleep(retry_after)
                    else:
                        print(f"❌ Rate limited creating role '{role_name}' after {max_retries} retries, giving up")
                        return None
                else:
                    raise  # Re-raise non-rate-limit errors
            except Exception as e:
                error_msg = str(e)
                if "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                    retry_count += 1
                    if retry_count < max_retries:
                        print(f"⚠️ Rate limited creating role '{role_name}', waiting 1s before retry {retry_count}/{max_retries}...")
                        await asyncio.sleep(1.0)
                    else:
                        print(f"❌ Rate limited creating role '{role_name}' after {max_retries} retries, giving up")
                        return None
                else:
                    raise  # Re-raise non-rate-limit errors
        
        return None
    except discord.Forbidden:
        print(f"❌ No permission to create role '{role_name}'")
        return None
    except Exception as e:
        print(f"❌ Error creating role '{role_name}': {e}")
        return None

async def get_or_create_category(guild, category_name):
    """Get a category by name, or create it if it doesn't exist"""
    category = discord.utils.get(guild.categories, name=category_name)
    if category:
        print(f"✅ DEBUG: Found existing category: '{category_name}'")
        return category
    
    try:
        category = await handle_rate_limit(
            guild.create_category(
                name=category_name,
                reason="Auto-created by LAM Bot for building organization"
            ),
            f"creating category '{category_name}'"
        )
        if category:
            print(f"🏢 DEBUG: Created NEW category: '{category_name}' (ID: {category.id})")
        return category
    except discord.Forbidden:
        print(f"❌ No permission to create category '{category_name}'")
        return None
    except Exception as e:
        print(f"❌ Error creating category '{category_name}': {e}")
        return None

async def get_or_create_channel(guild, channel_name, category, event_role=None, is_building_chat=False):
    """Get a channel by name, or create it if it doesn't exist"""
    channel = discord.utils.get(guild.text_channels, name=channel_name)
    if channel:
        print(f"✅ DEBUG: Found existing channel: #{channel_name} (ID: {channel.id})")
        return channel
    
    print(f"🔍 DEBUG: Channel #{channel_name} not found, creating new one...")
    
    try:
        # Set up permissions
        overwrites = {}
        
        # Give Slacker role access only to static channels (not building/event channels)
        slacker_role = discord.utils.get(guild.roles, name="Slacker")
        static_categories = ["Welcome", "Tournament Officials", "Chapters", "Volunteers"]
        if slacker_role and category and category.name in static_categories:
            overwrites[slacker_role] = discord.PermissionOverwrite(
                read_messages=True,
                send_messages=True,
                read_message_history=True
            )
        
        if event_role:
            # Event-specific channel: only event role can see it (plus Slacker)
            overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)
            overwrites[event_role] = discord.PermissionOverwrite(
                read_messages=True,
                send_messages=True,
                read_message_history=True
            )
        elif is_building_chat:
            # Building chat channel: restricted by default, roles will be added later
            overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)
        
        channel = await handle_rate_limit(
            guild.create_text_channel(
                name=channel_name,
                category=category,
                overwrites=overwrites,
                reason="Auto-created by LAM Bot for event organization"
            ),
            f"creating channel '{channel_name}'"
        )
        
        if channel:
            if event_role:
                print(f"📺 DEBUG: Created NEW channel: '#{channel_name}' (ID: {channel.id}, restricted to {event_role.name})")
            elif is_building_chat:
                print(f"📺 DEBUG: Created NEW building chat: '#{channel_name}' (ID: {channel.id}, restricted)")
            else:
                print(f"📺 DEBUG: Created NEW channel: '#{channel_name}' (ID: {channel.id})")
        return channel
    except discord.Forbidden:
        print(f"❌ No permission to create channel '{channel_name}'")
        return None
    except Exception as e:
        print(f"❌ Error creating channel '{channel_name}': {e}")
        return None

async def sort_building_categories_alphabetically(guild):
    """Sort all building categories alphabetically, keeping non-building categories at the top"""
    try:
        # Get all categories
        all_categories = guild.categories
        print(f"📋 DEBUG: Sorting {len(all_categories)} total categories")
        
        # Separate building categories from static categories
        static_categories = ["Welcome", "Tournament Officials", "Chapters", "Volunteers"]
        building_categories = []
        other_categories = []
        
        for category in all_categories:
            if category.name in static_categories:
                other_categories.append(category)
                print(f"📋 DEBUG: Static category: '{category.name}'")
            else:
                building_categories.append(category)
                print(f"🏢 DEBUG: Building category: '{category.name}'")
        
        # Sort building categories alphabetically
        building_categories.sort(key=lambda cat: cat.name.lower())
        print(f"📋 DEBUG: Sorted {len(building_categories)} building categories alphabetically")
        
        # Calculate positions: static categories first, then building categories
        position = 0
        
        # Position static categories first in the correct order
        desired_order = ["Welcome", "Tournament Officials", "Chapters", "Volunteers"]
        ordered_static_categories = []
        
        # Sort other_categories by desired order
        for desired_name in desired_order:
            for category in other_categories:
                if category.name == desired_name:
                    ordered_static_categories.append(category)
                    break
        
        # Add any remaining static categories that weren't in the desired order
        for category in other_categories:
            if category not in ordered_static_categories:
                ordered_static_categories.append(category)
        
        for category in ordered_static_categories:
            if category.position != position:
                result = await handle_rate_limit(
                    category.edit(position=position, reason="Organizing categories"),
                    f"moving category '{category.name}'"
                )
                if result is not None:
                    print(f"📋 Moved category '{category.name}' to position {position}")
            position += 1
        
        # Position building categories alphabetically after static ones
        for category in building_categories:
            if category.position != position:
                result = await handle_rate_limit(
                    category.edit(position=position, reason="Organizing building categories alphabetically"),
                    f"moving building category '{category.name}'"
                )
                if result is not None:
                    print(f"🏢 Moved building category '{category.name}' to position {position}")
            position += 1
            
        print("📋 Categories organized: Static categories first, then buildings alphabetically")
        
    except Exception as e:
        print(f"⚠️ Error organizing categories: {e}")

def sanitize_for_discord(text):
    """Sanitize text to be valid for Discord channel names"""
    # Replace spaces with hyphens and remove/replace invalid characters
    return text.lower().replace(' ', '-').replace('/', '-').replace('\\', '-').replace(':', '-').replace('*', '-').replace('?', '-').replace('"', '').replace('<', '').replace('>', '').replace('|', '-')

async def setup_building_structure(guild, building, first_event, room=None):
    """Set up category and channels for a building and event"""
    print(f"🏗️ DEBUG: Setting up building structure - Building: '{building}', Event: '{first_event}', Room: '{room}'")
    
    # Skip creating building chat for priority/custom roles (only create for actual event roles)
    priority_roles = [":(", "Volunteer", "Lead Event Supervisor", "Social Media", "Photographer", "Arbitrations", "Awards", "Slacker", "VIPer"]
    if first_event and first_event in priority_roles:
        print(f"⏭️ Skipping building structure creation for priority role '{first_event}' in {building} (only event roles get building structures)")
        return
    
    # Create or get the building category
    category_name = building
    print(f"🏢 DEBUG: Getting/creating category: '{category_name}'")
    category = await get_or_create_category(guild, category_name)
    if not category:
        return
    
    # Get Slacker role to ensure access
    slacker_role = discord.utils.get(guild.roles, name="Slacker")
    
    # Create general building chat channel (restricted to people with events in this building)
    building_chat_name = f"{sanitize_for_discord(building)}-chat"
    print(f"📺 DEBUG: Getting/creating building chat: '{building_chat_name}'")
    building_chat = await get_or_create_channel(guild, building_chat_name, category, is_building_chat=True)
    
    # Check if this is a newly created building chat (no messages yet) and send welcome message
    if building_chat:
        try:
            # Check if the channel has any messages (to avoid sending duplicate welcome messages)
            messages = [message async for message in building_chat.history(limit=1)]
            if not messages:
                print(f"📝 Building chat #{building_chat.name} appears to be new, sending welcome message...")
                await send_building_welcome_message(guild, building_chat, building)
        except Exception as e:
            print(f"⚠️ Error checking/sending welcome message for #{building_chat.name}: {e}")
    
    # Create event-specific channel if we have the info
    if first_event:
        # Get or create the event role
        event_role = await get_or_create_role(guild, first_event)
        if event_role:
            # Skip giving building chat access to Slacker role
            # (Slackers use the existing "slacker" channel in Tournament Officials)
            if first_event.lower() != "slacker":
                # Add the event role to the building chat permissions
                await add_role_to_building_chat(building_chat, event_role)
                
                # Create channel name: [First Event] - [Building] [Room]
                if room:
                    channel_name = f"{sanitize_for_discord(first_event)}-{sanitize_for_discord(building)}-{sanitize_for_discord(room)}"
                else:
                    channel_name = f"{sanitize_for_discord(first_event)}-{sanitize_for_discord(building)}"
                
                print(f"📺 DEBUG: Getting/creating event channel: '{channel_name}' for role '{event_role.name}'")
                
                # Create event channel
                event_channel = await get_or_create_channel(guild, channel_name, category, event_role)
                
                # After creating the event channel, search for test materials if this is an event-specific role
                priority_roles = [":(", "Volunteer", "Lead Event Supervisor", "Social Media", "Photographer", "Arbitrations", "Awards", "Slacker"]
                if first_event not in priority_roles:
                    print(f"🚀 DEBUG: Starting test folder search after channel creation for: {first_event}")
                    # This is an event-specific role, search for test folder now that channel exists
                    asyncio.create_task(search_and_share_test_folder(guild, first_event))

async def search_and_share_test_folder(guild, role_name):
    """Search for test materials folder and share with event participants"""
    try:
        print(f"🔍 DEBUG: Starting search for test materials for event: {role_name}")
        
        # Check if we have a connected spreadsheet to get the folder ID
        if not spreadsheet:
            print(f"❌ DEBUG: No spreadsheet connected, cannot search for test folder for {role_name}")
            return
        
        print(f"✅ DEBUG: Spreadsheet connected, ID: {spreadsheet.id}")
        
        # Import Drive API
        from googleapiclient.discovery import build
        
        # Build Drive API service
        drive_service = build('drive', 'v3', credentials=creds)
        
        # Get the parent folder ID of the connected spreadsheet
        sheet_metadata = drive_service.files().get(fileId=spreadsheet.id, fields='parents').execute()
        parent_folders = sheet_metadata.get('parents', [])
        
        if not parent_folders:
            print(f"❌ DEBUG: Could not find parent folder for the spreadsheet")
            return
        
        parent_folder_id = parent_folders[0]
        print(f"✅ DEBUG: Found parent folder ID: {parent_folder_id}")
        
        # Search for "Tests" folder in the parent directory
        tests_query = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='Tests'"
        tests_results = drive_service.files().list(q=tests_query, fields='files(id, name)').execute()
        tests_folders = tests_results.get('files', [])
        
        if not tests_folders:
            print(f"❌ DEBUG: No 'Tests' folder found in the parent directory")
            print(f"🔍 DEBUG: Searching for any folders in parent directory...")
            # List all folders in parent to help debug
            all_folders_query = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"
            all_folders_results = drive_service.files().list(q=all_folders_query, fields='files(id, name)').execute()
            all_folders = all_folders_results.get('files', [])
            print(f"📁 DEBUG: Found folders: {[f['name'] for f in all_folders]}")
            return
        
        tests_folder_id = tests_folders[0]['id']
        print(f"✅ DEBUG: Found Tests folder: {tests_folder_id}")
        
        # Search for the event-specific folder within Tests
        event_query = f"'{tests_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='{role_name}'"
        event_results = drive_service.files().list(q=event_query, fields='files(id, name, webViewLink)').execute()
        event_folders = event_results.get('files', [])
        
        if not event_folders:
            print(f"❌ DEBUG: No folder found for event '{role_name}' in Tests directory")
            print(f"🔍 DEBUG: Searching for any folders in Tests directory...")
            # List all folders in Tests to help debug
            all_test_folders_query = f"'{tests_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"
            all_test_folders_results = drive_service.files().list(q=all_test_folders_query, fields='files(id, name)').execute()
            all_test_folders = all_test_folders_results.get('files', [])
            print(f"📁 DEBUG: Found test folders: {[f['name'] for f in all_test_folders]}")
            return
        
        event_folder = event_folders[0]
        event_folder_id = event_folder['id']
        print(f"✅ DEBUG: Found test folder for {role_name}: {event_folder_id}")
        
        # Get all files in the event folder
        files_query = f"'{event_folder_id}' in parents and trashed=false"
        files_results = drive_service.files().list(q=files_query, fields='files(id, name, webViewLink, mimeType)').execute()
        files = files_results.get('files', [])
        
        if not files:
            print(f"❌ DEBUG: No files found in {role_name} test folder")
            return
        
        print(f"✅ DEBUG: Found {len(files)} files in {role_name} test folder")
        
        # Find the appropriate channel to post to (event-specific channel)
        target_channel = None
        
        print(f"🔍 DEBUG: Looking for channel containing '{role_name.lower().replace(' ', '-')}'")
        print(f"📺 DEBUG: Available channels: {[c.name for c in guild.text_channels]}")
        
        # Look for event-specific channels that contain the role name
        for channel in guild.text_channels:
            print(f"🔍 DEBUG: Checking channel #{channel.name} in category {channel.category.name if channel.category else 'None'}")
            if (role_name.lower().replace(' ', '-') in channel.name.lower() and 
                channel.category and 
                channel.category.name not in ["Welcome", "Tournament Officials", "Volunteers"]):
                target_channel = channel
                print(f"✅ DEBUG: Found target channel: #{channel.name}")
                break
        
        if not target_channel:
            print(f"❌ DEBUG: Could not find appropriate channel for {role_name}")
            print(f"🔍 DEBUG: Searched for channels containing: '{role_name.lower().replace(' ', '-')}'")
            return
        
        # Check if test materials message already exists in pinned messages
        pinned_messages = await target_channel.pins()
        test_materials_exists = False
        
        for message in pinned_messages:
            if message.embeds and message.embeds[0].title and f"📚 Test Materials for {role_name}" in message.embeds[0].title:
                test_materials_exists = True
                print(f"✅ DEBUG: Test materials message already pinned in #{target_channel.name}, skipping")
                break
        
        if test_materials_exists:
            return
        
        # Create embed for the test materials
        embed = discord.Embed(
            title=f"📚 Test Materials for {role_name}",
            description=f"Access your event-specific test materials and resources!\nPlease DO NOT share these materials with ANYBODY else (not even volunteers from different events).",
            color=discord.Color.green()
        )
        
        # Create file links as bullet points
        file_links = []
        for file in files:
            file_name = file['name']
            file_link = file['webViewLink']
            
            # Determine file type emoji
            mime_type = file.get('mimeType', '')
            if 'pdf' in mime_type:
                emoji = "📄"
            elif 'document' in mime_type:
                emoji = "📝"
            elif 'spreadsheet' in mime_type:
                emoji = "📊"
            elif 'presentation' in mime_type:
                emoji = "📖"
            elif 'image' in mime_type:
                emoji = "🖼️"
            elif 'folder' in mime_type:
                emoji = "📁"
            else:
                emoji = "📎"
            
            file_links.append(f"• {emoji} [**{file_name}**]({file_link})")
        
        # Split into chunks if too long for Discord (2000 character limit per field)
        files_text = "\n".join(file_links)
        if len(files_text) > 1900:  # Leave some buffer
            # Split into multiple fields
            chunk_size = 1900
            chunks = []
            current_chunk = ""
            
            for link in file_links:
                if len(current_chunk + link + "\n") > chunk_size:
                    chunks.append(current_chunk.strip())
                    current_chunk = link + "\n"
                else:
                    current_chunk += link + "\n"
            
            if current_chunk.strip():
                chunks.append(current_chunk.strip())
            
            # Add chunks as separate fields
            for i, chunk in enumerate(chunks):
                field_name = "📋 Test Materials" if i == 0 else f"📋 Test Materials (continued {i+1})"
                embed.add_field(name=field_name, value=chunk, inline=False)
        else:
            embed.add_field(name="📋 Test Materials", value=files_text, inline=False)

        # Post the embed to the channel
        message = await target_channel.send(embed=embed)
        print(f"📚 Shared test materials for {role_name} in #{target_channel.name}")
        
        # Pin the message so it's always visible
        try:
            await message.pin()
            print(f"📌 Pinned test materials message in #{target_channel.name}")
        except discord.Forbidden:
            print(f"⚠️ No permission to pin message in #{target_channel.name}")
        except Exception as pin_error:
            print(f"⚠️ Error pinning message in #{target_channel.name}: {pin_error}")
        
        # Check if scoring message already exists in pinned messages
        scoring_message_exists = False
        for msg in pinned_messages:
            if msg.embeds and msg.embeds[0].title and "📊 Score Input Instructions" in msg.embeds[0].title:
                scoring_message_exists = True
                print(f"✅ DEBUG: Scoring message already pinned in #{target_channel.name}, skipping")
                break
        
        if not scoring_message_exists:
            # Create scoring instructions embed
            scoring_embed = discord.Embed(
                title="📊 Score Input Instructions",
                description="**IMPORTANT**: All event supervisors must input scores through the official scoring portal!",
                color=discord.Color.blue()
            )
            
            scoring_embed.add_field(
                name="🔗 Scoring Portal",
                value="[**Click here to access the scoring system**](https://scoring.duosmium.org/login)",
                inline=False
            )
            
            scoring_embed.add_field(
                name="📋 Instructions",
                value="• Use your supervisor credentials to log in\n• Select the correct tournament and event\n• Input all team scores accurately\n• Double-check scores before submitting\n• Contact admin if you have login issues",
                inline=False
            )
            
            scoring_embed.add_field(
                name="⚠️ Important Notes",
                value="• Scores must be entered promptly after each event\n• Do not share your login credentials\n• Report any technical issues immediately",
                inline=False
            )
            
            # Post the scoring embed
            scoring_message = await target_channel.send(embed=scoring_embed)
            print(f"📊 Shared scoring instructions for {role_name} in #{target_channel.name}")
            
            # Pin the scoring message
            try:
                await scoring_message.pin()
                print(f"📌 Pinned scoring instructions message in #{target_channel.name}")
            except discord.Forbidden:
                print(f"⚠️ No permission to pin scoring message in #{target_channel.name}")
            except Exception as pin_error:
                print(f"⚠️ Error pinning scoring message in #{target_channel.name}: {pin_error}")
        
    except Exception as e:
        print(f"❌ Error searching for test folder for {role_name}: {e}")

async def setup_chapter_structure(guild, chapter_name):
    """Set up channels for a chapter"""
    print(f"📖 DEBUG: Setting up chapter structure - Chapter: '{chapter_name}'")
    
    # Add to global chapter role names set
    global chapter_role_names
    chapter_role_names.add(chapter_name)
    
    # Get or create the Chapters category
    chapters_category = await get_or_create_category(guild, "Chapters")
    if not chapters_category:
        return
    
    # Sanitize chapter name for Discord channel
    channel_name = sanitize_for_discord(chapter_name)
    
    # Create chapter channel
    chapter_channel = await get_or_create_channel(guild, channel_name, chapters_category)
    
    # Get or create the chapter role
    chapter_role = await get_or_create_role(guild, chapter_name)
    
    if chapter_channel and chapter_role:
        # Set up permissions so only chapter members can see the channel
        try:
            overwrites = chapter_channel.overwrites
            # Hide from @everyone
            overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)
            # Give chapter role access
            overwrites[chapter_role] = discord.PermissionOverwrite(
                read_messages=True,
                send_messages=True,
                read_message_history=True
            )
            
            # Give Slacker role access too
            slacker_role = discord.utils.get(guild.roles, name="Slacker")
            if slacker_role:
                overwrites[slacker_role] = discord.PermissionOverwrite(
                    read_messages=True,
                    send_messages=True,
                    read_message_history=True
                )
            
            await handle_rate_limit(
                chapter_channel.edit(overwrites=overwrites, reason=f"Set up {chapter_name} chapter permissions"),
                f"editing chapter channel '{channel_name}' permissions"
            )
            print(f"📖 Set up permissions for #{channel_name} chapter channel")
            
            # Sort chapter channels after creating a new one
            await sort_chapter_channels_alphabetically(guild)
            
        except Exception as e:
            print(f"❌ Error setting up permissions for #{channel_name}: {e}")

async def sort_chapter_channels_alphabetically(guild):
    """Sort chapter channels alphabetically with unaffiliated at the bottom"""
    try:
        # Find the Chapters category
        chapters_category = discord.utils.get(guild.categories, name="Chapters")
        if not chapters_category:
            print("⚠️ Chapters category not found")
            return
        
        # Get all text channels in the Chapters category
        chapter_channels = [channel for channel in chapters_category.text_channels]
        if len(chapter_channels) <= 1:
            print("📖 Not enough chapter channels to sort")
            return
        
        print(f"📖 Sorting {len(chapter_channels)} chapter channels alphabetically...")
        
        # Separate unaffiliated from other channels
        unaffiliated_channels = [ch for ch in chapter_channels if ch.name == "unaffiliated"]
        other_channels = [ch for ch in chapter_channels if ch.name != "unaffiliated"]
        
        # Sort other channels alphabetically
        other_channels.sort(key=lambda ch: ch.name.lower())
        
        # Combine: other channels first, then unaffiliated at the bottom
        final_order = other_channels + unaffiliated_channels
        
        # Update positions within the category
        for i, channel in enumerate(final_order):
            if channel.position != i:
                try:
                    result = await handle_rate_limit(
                        channel.edit(position=i, reason="Sorting chapter channels alphabetically"),
                        f"moving channel '{channel.name}'"
                    )
                    if result is not None:
                        print(f"📖 Moved #{channel.name} to position {i}")
                except Exception as e:
                    print(f"❌ Error moving #{channel.name}: {e}")
        
        print("✅ Chapter channels sorted alphabetically (unaffiliated at bottom)")
        
    except Exception as e:
        print(f"❌ Error sorting chapter channels: {e}")

async def search_and_share_useful_links(guild):
    """Search for Useful Links folder and share with volunteers"""
    try:
        print(f"🔍 DEBUG: Searching for Useful Links folder")
        
        # Check if we have a connected spreadsheet to get the folder ID
        if not spreadsheet:
            print(f"❌ DEBUG: No spreadsheet connected, cannot search for Useful Links folder")
            return
        
        print(f"✅ DEBUG: Spreadsheet connected, ID: {spreadsheet.id}")
        
        # Import Drive API
        from googleapiclient.discovery import build
        
        # Build Drive API service
        drive_service = build('drive', 'v3', credentials=creds)
        
        # Get the parent folder ID of the connected spreadsheet
        sheet_metadata = drive_service.files().get(fileId=spreadsheet.id, fields='parents').execute()
        parent_folders = sheet_metadata.get('parents', [])
        
        if not parent_folders:
            print(f"❌ DEBUG: Could not find parent folder for the spreadsheet")
            return
        
        parent_folder_id = parent_folders[0]
        print(f"✅ DEBUG: Found parent folder ID: {parent_folder_id}")
        
        # Search for "Useful Links" folder in the parent directory
        useful_links_query = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='Useful Links'"
        useful_links_results = drive_service.files().list(q=useful_links_query, fields='files(id, name, webViewLink)').execute()
        useful_links_folders = useful_links_results.get('files', [])
        
        if not useful_links_folders:
            print(f"❌ DEBUG: No 'Useful Links' folder found in the parent directory")
            return
        
        useful_links_folder = useful_links_folders[0]
        useful_links_folder_id = useful_links_folder['id']
        print(f"✅ DEBUG: Found Useful Links folder: {useful_links_folder_id}")
        
        # Get all files in the Useful Links folder
        files_query = f"'{useful_links_folder_id}' in parents and trashed=false"
        files_results = drive_service.files().list(q=files_query, fields='files(id, name, webViewLink, mimeType)').execute()
        files = files_results.get('files', [])
        
        if not files:
            print(f"❌ DEBUG: No files found in Useful Links folder")
            return
        
        print(f"✅ DEBUG: Found {len(files)} files in Useful Links folder")
        
        # Find the volunteers useful-links channel
        target_channel = discord.utils.get(guild.text_channels, name="useful-links")
        
        if not target_channel:
            print(f"❌ DEBUG: Could not find useful-links channel")
            return
        
        print(f"✅ DEBUG: Found target channel: #{target_channel.name}")
        
        # Check if useful links message already exists in pinned messages
        pinned_messages = await target_channel.pins()
        useful_links_exists = False
        
        for message in pinned_messages:
            if message.embeds and message.embeds[0].title and "🔗 Useful Links & Resources" in message.embeds[0].title:
                useful_links_exists = True
                print(f"✅ DEBUG: Useful links message already pinned in #{target_channel.name}, skipping")
                break
        
        if useful_links_exists:
            return
        
        # Create embed for the useful links
        embed = discord.Embed(
            title="🔗 Useful Links & Resources",
            description="Access important links and resources for volunteers!",
            color=discord.Color.green()
        )
        
        # Create file links as bullet points
        file_links = []
        for file in files:
            file_name = file['name']
            file_link = file['webViewLink']
            
            # Determine file type emoji
            mime_type = file.get('mimeType', '')
            if 'pdf' in mime_type:
                emoji = "📄"
            elif 'document' in mime_type:
                emoji = "📝"
            elif 'spreadsheet' in mime_type:
                emoji = "📊"
            elif 'presentation' in mime_type:
                emoji = "📖"
            elif 'image' in mime_type:
                emoji = "🖼️"
            elif 'folder' in mime_type:
                emoji = "📁"
            else:
                emoji = "📎"
            
            file_links.append(f"• {emoji} [**{file_name}**]({file_link})")
        
        # Split into chunks if too long for Discord (2000 character limit per field)
        files_text = "\n".join(file_links)
        if len(files_text) > 1900:  # Leave some buffer
            # Split into multiple fields
            chunk_size = 1900
            chunks = []
            current_chunk = ""
            
            for link in file_links:
                if len(current_chunk + link + "\n") > chunk_size:
                    chunks.append(current_chunk.strip())
                    current_chunk = link + "\n"
                else:
                    current_chunk += link + "\n"
            
            if current_chunk.strip():
                chunks.append(current_chunk.strip())
            
            # Add chunks as separate fields
            for i, chunk in enumerate(chunks):
                field_name = "📋 Useful Links" if i == 0 else f"📋 Useful Links (continued {i+1})"
                embed.add_field(name=field_name, value=chunk, inline=False)
        else:
            embed.add_field(name="📋 Useful Links", value=files_text, inline=False)
        
        # Post the embed to the channel
        message = await target_channel.send(embed=embed)
        print(f"🔗 Shared useful links in #{target_channel.name}")
        
        # Pin the message so it's always visible
        try:
            await message.pin()
            print(f"📌 Pinned useful links message in #{target_channel.name}")
        except discord.Forbidden:
            print(f"⚠️ No permission to pin message in #{target_channel.name}")
        except Exception as pin_error:
            print(f"⚠️ Error pinning message in #{target_channel.name}: {pin_error}")
        
    except Exception as e:
        print(f"❌ Error searching for Useful Links folder: {e}")

async def add_slacker_access(channel, slacker_role):
    """Add Slacker role access to a channel"""
    if not channel or not slacker_role:
        return
    
    try:
        # Get current overwrites
        overwrites = channel.overwrites
        
        # Add Slacker role with full permissions
        overwrites[slacker_role] = discord.PermissionOverwrite(
            read_messages=True,
            send_messages=True,
            read_message_history=True
        )
        
        # Update channel permissions
        await handle_rate_limit(
            channel.edit(overwrites=overwrites, reason=f"Added {slacker_role.name} access to all channels"),
            f"editing channel '{channel.name}' permissions"
        )
        print(f"🔑 Added {slacker_role.name} access to #{channel.name}")
        
    except discord.Forbidden:
        print(f"❌ No permission to edit channel permissions for #{channel.name}")
    except Exception as e:
        print(f"❌ Error updating channel permissions for #{channel.name}: {e}")

async def ensure_slacker_tournament_officials_access(guild, slacker_role):
    """Ensure Slacker role has access to Tournament Officials channels"""
    if not slacker_role:
        return
    
    print(f"🔑 Ensuring {slacker_role.name} access to Tournament Officials channels...")
    
    # Get Tournament Officials category
    tournament_officials_category = discord.utils.get(guild.categories, name="Tournament Officials")
    if not tournament_officials_category:
        print("⚠️ Tournament Officials category not found, skipping access setup")
        return
    
    # List of Tournament Officials channels
    official_channels = ["slacker", "links", "scoring", "awards-ceremony"]
    
    added_count = 0
    for channel_name in official_channels:
        channel = discord.utils.get(guild.text_channels, name=channel_name)
        if channel and channel.category == tournament_officials_category:
            try:
                await add_slacker_access(channel, slacker_role)
                added_count += 1
            except Exception as e:
                print(f"❌ Error adding Slacker access to #{channel_name}: {e}")
    
    print(f"✅ Added {slacker_role.name} access to {added_count} Tournament Officials channels")

async def send_building_welcome_message(guild, building_chat, building):
    """Send an initial welcome message to a building chat with all events in that building"""
    if not building_chat or not building:
        return
    
    try:
        # Get all events in this building
        building_events = await get_building_events(building)
        
        if not building_events:
            print(f"⚠️ No events found for building '{building}', skipping welcome message")
            return
        
        # Create the welcome message
        embed = discord.Embed(
            title=f"🏢 Welcome to {building}!",
            description=f"This is the general chat for everyone with events in **{building}**.",
            color=discord.Color.blue()
        )
        
        # Add events list
        events_text = ""
        for event, room in building_events:
            if room:
                events_text += f"• **{event}** - {room}\n"
            else:
                events_text += f"• **{event}**\n"
        
        embed.add_field(
            name="📋 Events in this building:",
            value=events_text,
            inline=False
        )
        
        embed.add_field(
            name="💬 How to use this chat:",
            value="• Coordinate with other events in your building\n• Share building-specific information\n• Ask questions about the venue\n• Connect with nearby events",
            inline=False
        )
        
        embed.set_footer(text="Each event also has its own dedicated channel for event-specific discussions.")
        
        # Send the message
        message = await building_chat.send(embed=embed)
        print(f"🏢 Sent welcome message to #{building_chat.name} for building '{building}'")
        
        # Pin the message so it's always visible
        try:
            await message.pin()
            print(f"📌 Pinned welcome message in #{building_chat.name}")
        except discord.Forbidden:
            print(f"⚠️ Could not pin welcome message in #{building_chat.name} (missing permissions)")
        except Exception as e:
            print(f"⚠️ Error pinning welcome message in #{building_chat.name}: {e}")
            
    except Exception as e:
        print(f"❌ Error sending welcome message to #{building_chat.name}: {e}")


async def add_role_to_building_chat(channel, role):
    """Add a role to a building chat channel permissions"""
    if not channel or not role:
        return
    
    try:
        # Get current overwrites
        overwrites = channel.overwrites
        
        # Set @everyone to not see the channel
        overwrites[channel.guild.default_role] = discord.PermissionOverwrite(read_messages=False)
        
        # Add the role with permissions to see and participate
        overwrites[role] = discord.PermissionOverwrite(
            read_messages=True,
            send_messages=True,
            read_message_history=True
        )
        
        # Update channel permissions
        await handle_rate_limit(
            channel.edit(overwrites=overwrites, reason=f"Added {role.name} to building chat access"),
            f"editing building chat '{channel.name}' permissions"
        )
        print(f"🔒 Added {role.name} access to #{channel.name}")
        
    except discord.Forbidden:
        print(f"❌ No permission to edit channel permissions for #{channel.name}")
    except Exception as e:
        print(f"❌ Error updating channel permissions for #{channel.name}: {e}")

async def reset_server_for_guild(guild):
    """⚠️ DANGER: Completely reset the server by deleting all channels, categories, roles, and nicknames"""
    if not guild:
        print("❌ Guild not provided!")
        return
    
    print("⚠️ ⚠️ ⚠️  STARTING COMPLETE SERVER RESET  ⚠️ ⚠️ ⚠️")
    print("🧨 This will delete EVERYTHING and reset all nicknames!")
    print("⏰ Starting in 3 seconds... (Ctrl+C to cancel)")
    
    await asyncio.sleep(3)
    
    print("🗑️ Starting server reset...")
    
    # Reset all member nicknames
    print("📝 Resetting all member nicknames...")
    nickname_count = 0
    for member in guild.members:
        if member.nick and not member.bot:  # Don't reset bot nicknames
            try:
                await handle_rate_limit(
                    member.edit(nick=None, reason="Server reset - clearing nickname"),
                    f"resetting nickname for {member}"
                )
                nickname_count += 1
                print(f"📝 Reset nickname for {member.display_name}")
            except discord.Forbidden:
                print(f"❌ No permission to reset nickname for {member.display_name}")
            except Exception as e:
                print(f"⚠️ Error resetting nickname for {member.display_name}: {e}")
    print(f"✅ Reset {nickname_count} nicknames")
    
    # Delete all text channels
    print("🗑️ Deleting all text channels...")
    channel_count = 0
    for channel in guild.text_channels:
        try:
            await channel.delete(reason="Server reset")
            channel_count += 1
            print(f"🗑️ Deleted text channel: #{channel.name}")
        except discord.Forbidden:
            print(f"❌ No permission to delete channel #{channel.name}")
        except Exception as e:
            print(f"⚠️ Error deleting channel #{channel.name}: {e}")
    
    # Delete all voice channels
    print("🗑️ Deleting all voice channels...")
    voice_count = 0
    for channel in guild.voice_channels:
        try:
            await channel.delete(reason="Server reset")
            voice_count += 1
            print(f"🗑️ Deleted voice channel: {channel.name}")
        except discord.Forbidden:
            print(f"❌ No permission to delete voice channel {channel.name}")
        except Exception as e:
            print(f"⚠️ Error deleting voice channel {channel.name}: {e}")
    
    # Delete all forum channels
    print("🗑️ Deleting all forum channels...")
    forum_count = 0
    for channel in guild.channels:
        if hasattr(channel, 'type') and channel.type == discord.ChannelType.forum:
            try:
                await channel.delete(reason="Server reset")
                forum_count += 1
                print(f"🗑️ Deleted forum channel: #{channel.name}")
            except discord.Forbidden:
                print(f"❌ No permission to delete forum #{channel.name}")
            except Exception as e:
                print(f"⚠️ Error deleting forum #{channel.name}: {e}")
    
    # Delete all categories
    print("🗑️ Deleting all categories...")
    category_count = 0
    for category in guild.categories:
        try:
            await category.delete(reason="Server reset")
            category_count += 1
            print(f"🗑️ Deleted category: {category.name}")
        except discord.Forbidden:
            print(f"❌ No permission to delete category {category.name}")
        except Exception as e:
            print(f"⚠️ Error deleting category {category.name}: {e}")
    
    # Delete all custom roles (keep @everyone and bot roles)
    print("🗑️ Deleting all custom roles...")
    role_count = 0
    for role in guild.roles:
        # Skip @everyone, bot roles, and roles higher than bot's highest role
        if (role.name != "@everyone" and 
            not role.managed and 
            role < guild.me.top_role):
            try:
                await role.delete(reason="Server reset")
                role_count += 1
                print(f"🗑️ Deleted role: {role.name}")
            except discord.Forbidden:
                print(f"❌ No permission to delete role {role.name}")
            except Exception as e:
                print(f"⚠️ Error deleting role {role.name}: {e}")
    
    print("🧨 SERVER RESET COMPLETE!")
    print(f"📊 Summary:")
    print(f"   • {nickname_count} nicknames reset")
    print(f"   • {channel_count} text channels deleted")
    print(f"   • {voice_count} voice channels deleted") 
    print(f"   • {forum_count} forum channels deleted")
    print(f"   • {category_count} categories deleted")
    print(f"   • {role_count} roles deleted")
    print("🏗️ Server is now completely clean and ready for fresh setup!")

async def post_welcome_instructions(welcome_channel):
    """Post welcome instructions and login information to the welcome channel"""
    try:
        # Check if there are already messages in the channel
        async for message in welcome_channel.history(limit=10):
            # If there are messages, check if any are from the bot with the welcome embed
            if message.author == bot.user and message.embeds:
                for embed in message.embeds:
                    if embed.title and "Welcome to the Science Olympiad Server" in embed.title:
                        print(f"✅ Welcome instructions already posted in #{welcome_channel.name}")
                        return
        
        # Create welcome embed
        embed = discord.Embed(
            title="🎉 Welcome to the Science Olympiad Server!",
            description="Thank you for joining our Science Olympiad community! This server helps coordinate events, volunteers, and communication.",
            color=discord.Color.blue()
        )
        
        embed.add_field(
            name="🔐 Getting Started - Login Required",
            value="**To access all channels and get your roles, you need to login:**\n\n"
                  "1️⃣ Type `/login` in any channel\n"
                  "2️⃣ Enter your email address when prompted\n"
                  "3️⃣ Get instant access to your assigned channels!\n\n"
                  "✅ You'll automatically receive:\n"
                  "• Your assigned roles\n"
                  "• Access to relevant channels\n"
                  "• Your building and room information\n"
                  "• Updated nickname with your event",
            inline=False
        )
        
        embed.add_field(
            name="📍 What You Can Do Right Now",
            value="Even before logging in, you can:\n"
                  "• Read announcements in this channel\n"
                  "• Browse volunteer channels for general info\n"
                  "• Ask questions in the help forum\n"
                  "• Start sobbing uncontrollably",
            inline=False
        )
        
        embed.add_field(
            name="❓ Need Help?",
            value="• **Can't find your email?** Contact an admin\n"
                  "• **Questions about your assignment?** Ask in volunteer channels\n"
                  "• **Technical problems?** Mention an admin or moderator\n"
                  "• **Don't know who the admins/moderators are?** Contact Edward Zhang\n"
                  "• **Edward's ghosting you?** LOL gg. Maybe try sending him $5. Jkjk you should contact David Zheng or Brian Lam instead",
            inline=False
        )
        
        embed.add_field(
            name="🎯 Important Notes",
            value="• Your email must be in our system to login\n"
                  "• Each email can only be linked to one Discord account\n"
                  "• Your nickname will be updated to show your event\n"
                  "• Channels will appear based on your assigned roles",
            inline=False
        )
        
        embed.set_footer(text="Use /login to get started! • Questions? Ask in volunteer channels")
        
        # Send the welcome message
        await welcome_channel.send(embed=embed)
        print(f"📋 Posted welcome instructions to #{welcome_channel.name}")
        
    except Exception as e:
        print(f"❌ Error posting welcome instructions: {e}")

async def post_welcome_tldr(welcome_channel):
    """Post welcome instructions and login information to the welcome channel"""
    try:
        # Check if TLDR message already exists
        async for message in welcome_channel.history(limit=10):
            if message.author == bot.user and message.embeds:
                for embed in message.embeds:
                    if embed.title and "TLDR: TYPE" in embed.title:
                        print(f"✅ Welcome TLDR already posted in #{welcome_channel.name}")
                        return
        
        # Create welcome embed
        embed = discord.Embed(
            title="TLDR: TYPE `/login` TO GET STARTED",
            description="Read below message for more info",
            color=discord.Color.blue()
        )
        
        # Send the welcome message
        await welcome_channel.send(embed=embed)
        print(f"📋 Posted welcome tldr to #{welcome_channel.name}")
        
    except Exception as e:
        print(f"❌ Error posting welcome tldr: {e}")

async def setup_static_channels_for_guild(guild):
    """Create static categories and channels for Tournament Officials and Volunteers"""
    if not guild:
        print("❌ Guild not provided!")
        return
    
    print(f"🏗️ Setting up static channels for {guild.name}...")
    
    # Get or create Slacker role for permissions
    slacker_role = await get_or_create_role(guild, "Slacker")
    # Get or create Awards role for awards-ceremony access
    awards_role = await get_or_create_role(guild, "Awards")
    
    # Welcome Category
    print("👋 Setting up Welcome category...")
    welcome_category = await get_or_create_category(guild, "Welcome")
    if welcome_category:
        # Create welcome channel (visible to everyone)
        welcome_channel = await get_or_create_channel(guild, "welcome", welcome_category)
        
        # Post welcome instructions
        if welcome_channel:
            await post_welcome_tldr(welcome_channel)
            await post_welcome_instructions(welcome_channel)
    
    # Tournament Officials Category
    print("📋 Setting up Tournament Officials category...")
    tournament_officials_category = await get_or_create_category(guild, "Tournament Officials")
    if tournament_officials_category:
        # Create channels in Tournament Officials category (restricted to Slacker role only)
        official_channels = ["slacker", "links", "scoring", "awards-ceremony"]
        for channel_name in official_channels:
            # Check if channel already exists
            channel = discord.utils.get(guild.text_channels, name=channel_name)
            if not channel:
                # Create new channel with restricted permissions
                try:
                    overwrites = {}
                    # Hide from @everyone
                    overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)
                    # Give Slacker role full access
                    if slacker_role:
                        overwrites[slacker_role] = discord.PermissionOverwrite(
                            read_messages=True,
                            send_messages=True,
                            read_message_history=True
                        )
                    # Give Awards role access to awards-ceremony channel
                    if channel_name == "awards-ceremony" and awards_role:
                        overwrites[awards_role] = discord.PermissionOverwrite(
                            read_messages=True,
                            send_messages=True,
                            read_message_history=True
                        )
                    
                    channel = await handle_rate_limit(
                        guild.create_text_channel(
                            name=channel_name,
                            category=tournament_officials_category,
                            overwrites=overwrites,
                            reason="Auto-created by LAM Bot - Tournament Officials only"
                        ),
                        f"creating channel '{channel_name}'"
                    )
                    if channel_name == "awards-ceremony":
                        print(f"📺 Created restricted channel: '#{channel_name}' (Slacker + Awards)")
                    else:
                        print(f"📺 Created restricted channel: '#{channel_name}' (Slacker only)")
                    
                    # Ensure Slacker access is properly added after channel creation
                    if slacker_role:
                        try:
                            await add_slacker_access(channel, slacker_role)
                            print(f"✅ Ensured Slacker access to #{channel_name}")
                        except Exception as e:
                            print(f"❌ Error ensuring Slacker access to #{channel_name}: {e}")
                    
                    # Ensure Awards role access to awards-ceremony channel
                    if channel_name == "awards-ceremony" and awards_role:
                        try:
                            await add_slacker_access(channel, awards_role)  # Reuse the same function
                            print(f"✅ Ensured Awards role access to #{channel_name}")
                        except Exception as e:
                            print(f"❌ Error ensuring Awards access to #{channel_name}: {e}")
                            
                except discord.Forbidden:
                    print(f"❌ No permission to create channel '{channel_name}'")
                except Exception as e:
                    print(f"❌ Error creating channel '{channel_name}': {e}")
            else:
                # Update existing channel to be restricted
                print(f"✅ Channel '#{channel_name}' already exists, updating permissions...")
                try:
                    overwrites = channel.overwrites
                    # Hide from @everyone
                    overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)
                    # Give Slacker role full access
                    if slacker_role:
                        overwrites[slacker_role] = discord.PermissionOverwrite(
                            read_messages=True,
                            send_messages=True,
                            read_message_history=True
                        )
                    # Give Awards role access to awards-ceremony channel
                    if channel_name == "awards-ceremony" and awards_role:
                        overwrites[awards_role] = discord.PermissionOverwrite(
                            read_messages=True,
                            send_messages=True,
                            read_message_history=True
                        )
                    
                    await handle_rate_limit(
                        channel.edit(overwrites=overwrites, reason="Updated to restrict to Slacker role only"),
                        f"editing channel '{channel_name}' permissions"
                    )
                    if channel_name == "awards-ceremony":
                        print(f"🔒 Updated #{channel_name} to be Slacker + Awards")
                    else:
                        print(f"🔒 Updated #{channel_name} to be Slacker-only")
                except Exception as e:
                    print(f"❌ Error updating permissions for #{channel_name}: {e}")
                
                # Ensure Slacker access is properly added after channel creation/update
                if slacker_role:
                    try:
                        await add_slacker_access(channel, slacker_role)
                        print(f"✅ Ensured Slacker access to #{channel_name}")
                    except Exception as e:
                        print(f"❌ Error ensuring Slacker access to #{channel_name}: {e}")
                
                # Ensure Awards role access to awards-ceremony channel
                if channel_name == "awards-ceremony" and awards_role:
                    try:
                        await add_slacker_access(channel, awards_role)  # Reuse the same function
                        print(f"✅ Ensured Awards role access to #{channel_name}")
                    except Exception as e:
                        print(f"❌ Error ensuring Awards access to #{channel_name}: {e}")
    
    # Chapters Category  
    print("📖 Setting up Chapters category...")
    chapters_category = await get_or_create_category(guild, "Chapters")
    
    # Volunteers Category  
    print("🙋 Setting up Volunteers category...")
    volunteers_category = await get_or_create_category(guild, "Volunteers")
    if volunteers_category:
        # Create regular text channels
        volunteer_text_channels = ["general", "useful-links", "random"]
        for channel_name in volunteer_text_channels:
            channel = await get_or_create_channel(guild, channel_name, volunteers_category)
        
        # Create forum channel for "help"
        help_channel = None
        
        # Try to find existing forum channel first
        for channel in guild.channels:
            if channel.name == "help" and hasattr(channel, 'type') and channel.type == discord.ChannelType.forum:
                help_channel = channel
                break
        
        if not help_channel:
            # Try to create forum channel
            try:
                overwrites = {}
                # Give Slacker role access to forum
                if slacker_role:
                    overwrites[slacker_role] = discord.PermissionOverwrite(
                        read_messages=True,
                        send_messages=True,
                        read_message_history=True,
                        create_public_threads=True,
                        send_messages_in_threads=True
                    )
                
                # Try different methods to create forum
                if hasattr(guild, 'create_forum_channel'):
                    help_channel = await handle_rate_limit(
                        guild.create_forum_channel(
                            name="help",
                            category=volunteers_category,
                            overwrites=overwrites,
                            reason="Auto-created LAM Bot forum channel"
                        ),
                        "creating forum channel 'help'"
                    )
                    if help_channel:
                        print(f"📺 Created forum channel: '#{help_channel.name}' ✅")
                elif hasattr(guild, 'create_forum'):
                    help_channel = await handle_rate_limit(
                        guild.create_forum(
                            name="help",
                            category=volunteers_category,
                            overwrites=overwrites,
                            reason="Auto-created by LAM Bot - Volunteers help forum"
                        ),
                        "creating forum 'help'"
                    )
                    print(f"📺 Created forum channel: '#{help_channel.name}' ✅")
                else:
                    print("⚠️ Forum creation not supported in this py-cord version")
                    print("📝 Please manually create a forum channel named 'help' in the Volunteers category")
                    print("   1. Right-click the Volunteers category")
                    print("   2. Create Channel → Forum")
                    print("   3. Name it 'help'")
                    print("   4. The bot will add permissions automatically on next restart")
                    
            except AttributeError:
                print("⚠️ Forum channels not supported in this py-cord version")
                print("📝 Please manually create a forum channel named 'help' in the Volunteers category")
            except discord.Forbidden:
                print(f"❌ No permission to create forum channel 'help'")
                print("📝 Please manually create a forum channel named 'help' in the Volunteers category")
            except Exception as e:
                print(f"❌ Error creating forum channel 'help': {e}")
                print("📝 Please manually create a forum channel named 'help' in the Volunteers category")
        else:
            print(f"✅ Forum channel 'help' already exists")
            
        # Slacker access to help channel will be handled automatically by the static category logic
    
    print("✅ Finished setting up static channels")

async def move_bot_role_to_top_for_guild(guild):
    """Move the bot's role to the highest possible position and make it teal"""
    if not guild:
        print("❌ Guild not provided!")
        return
    
    # Check if bot has required permissions
    bot_member = guild.me
    if not bot_member.guild_permissions.manage_roles:
        print("❌ Bot missing 'Manage Roles' permission! Cannot move bot role to top.")
        return
    
    print("🤖 Moving bot role to top and making it teal...")
    
    try:
        # Find the bot's role
        bot_role = None
        for role in guild.roles:
            if role.managed and role.members and bot.user in role.members:
                bot_role = role
                break
        
        if not bot_role:
            print("⚠️ Could not find bot's role!")
            return
        
        print(f"🤖 Found bot role: '{bot_role.name}' (current position: {bot_role.position})")
        
        # Calculate the highest position the bot can reach
        # Bot can only move to positions below roles that are above it and unmovable
        max_possible_position = len(guild.roles) - 1  # Highest possible position
        
        # Check if there are unmovable roles above us
        higher_unmovable_roles = []
        for role in guild.roles:
            if role.position > bot_role.position and (role.managed or role == guild.default_role):
                higher_unmovable_roles.append(role)
        
        if higher_unmovable_roles:
            # Can't go above unmovable roles, so go just below the lowest unmovable role
            max_possible_position = min([r.position for r in higher_unmovable_roles]) - 1
        
        # Try to move the bot role to the highest possible position and make it teal
        changes_made = False
        
        # Try to update color to teal if it's not already
        if bot_role.color != discord.Color.teal():
            try:
                await handle_rate_limit(
                    bot_role.edit(color=discord.Color.teal(), reason="Making bot role teal"),
                    "editing bot role color"
                )
                print(f"🎨 Changed bot role color to teal")
                changes_made = True
            except discord.Forbidden:
                print(f"⚠️ Cannot change bot role color automatically (Discord restriction)")
                print(f"💡 To make the role teal: Go to Server Settings → Roles → {bot_role.name} → Change color to teal")
            except Exception as e:
                print(f"⚠️ Could not change bot role color: {e}")
        else:
            print(f"✅ Bot role already teal colored")
        
        # Try to move to highest position if not already there
        if bot_role.position != max_possible_position:
            try:
                await handle_rate_limit(
                    bot_role.edit(position=max_possible_position, reason="Moving bot role to top"),
                    "moving bot role position"
                )
                print(f"📈 Moved bot role to position {max_possible_position} (highest possible)")
                changes_made = True
            except discord.Forbidden:
                print(f"⚠️ Cannot move bot role automatically (Discord restriction)")
                print(f"💡 To move to top: Go to Server Settings → Roles → Drag {bot_role.name} to the top")
            except Exception as e:
                print(f"⚠️ Could not move bot role to top: {e}")
        else:
            print(f"✅ Bot role already at position {bot_role.position}")
        
        # Check final status and provide summary
        roles_above_bot = [r for r in guild.roles if r.position > bot_role.position and r.name != "@everyone"]
        is_teal_now = bot_role.color == discord.Color.teal()
        
        if not roles_above_bot and is_teal_now:
            print(f"✅ Bot role '{bot_role.name}' is perfectly optimized! (Top position + Teal color)")
        else:
            print(f"⚠️ Bot role '{bot_role.name}' needs manual adjustments:")
            
            if roles_above_bot:
                print(f"   📋 Position: {len(roles_above_bot)} roles still above the bot")
                for role in roles_above_bot[:3]:  # Show first 3
                    print(f"      • {role.name}")
                if len(roles_above_bot) > 3:
                    print(f"      • ... and {len(roles_above_bot)-3} more")
            else:
                print(f"   ✅ Position: At top (#{bot_role.position})")
                
            if not is_teal_now:
                print(f"   🎨 Color: Needs to be changed to teal manually")
            else:
                print(f"   ✅ Color: Already teal")
                
            print(f"\n💡 Manual steps needed:")
            print(f"   1. Go to Server Settings → Roles")
            if roles_above_bot:
                print(f"   2. Drag '{bot_role.name}' to the VERY TOP")
            if not is_teal_now:
                print(f"   {'3' if roles_above_bot else '2'}. Click '{bot_role.name}' → Change color to teal")
            print(f"   {'4' if (roles_above_bot and not is_teal_now) else ('3' if (roles_above_bot or not is_teal_now) else '2')}. Use /fixbotrole to verify, then /organizeroles to organize other roles")
            
    except Exception as e:
        print(f"❌ Error moving bot role to top: {e}")

async def organize_role_hierarchy_for_guild(guild):
    """Organize roles in priority order: lambot, Slacker, Arbitrations, Photographer, Social Media, Lead Event Supervisor, Volunteer, :(, then others"""
    if not guild:
        print("❌ Guild not provided!")
        return
    
    # Check if bot has required permissions
    bot_member = guild.me
    if not bot_member.guild_permissions.manage_roles:
        print("❌ Bot missing 'Manage Roles' permission! Cannot organize role hierarchy.")
        print("💡 Please give the bot 'Manage Roles' permission in Server Settings → Roles")
        return
    
    # Define the priority order (higher index = higher priority/position)
    priority_roles = [
        ":(",  # Lowest priority (position 1)
        "Volunteer",
        "Lead Event Supervisor", 
        "Social Media",
        "Photographer",
        "Arbitrations",
        "Awards",
        "Slacker",
        # Bot role will be handled separately as highest priority
    ]
    
    print("📋 Organizing role hierarchy...")
    
    try:
        # Get all roles except @everyone
        all_roles = [role for role in guild.roles if role.name != "@everyone"]
        
        # Find the bot's role
        bot_role = None
        for role in all_roles:
            if role.managed and role.members and bot.user in role.members:
                bot_role = role
                break
        
        if not bot_role:
            print("⚠️ Could not find bot's role!")
            return
        
        print(f"🤖 Bot role: '{bot_role.name}' (current position: {bot_role.position})")
        
        # Separate roles into priority roles, chapter roles, and other roles
        priority_role_objects = []
        chapter_roles = []
        other_roles = []
        unmovable_roles = []
        
        for role in all_roles:
            if role == bot_role:
                continue  # Handle bot role separately
            elif role.position >= bot_role.position:
                # Can't move roles that are higher than or equal to bot's current position
                unmovable_roles.append(role)
                continue
            elif role.name in priority_roles:
                priority_role_objects.append(role)
            elif role.name == "Unaffiliated" or role.name in chapter_role_names:
                # This is a chapter role
                chapter_roles.append(role)
            else:
                other_roles.append(role)
        
        if unmovable_roles:
            print(f"⚠️ Cannot move {len(unmovable_roles)} roles (higher than bot): {', '.join([r.name for r in unmovable_roles])}")
            print("💡 Move the bot's role higher in Server Settings → Roles to manage these roles")
        
        # Sort priority roles according to the defined order
        priority_role_objects.sort(key=lambda r: priority_roles.index(r.name) if r.name in priority_roles else 999)
        
        # Sort chapter roles alphabetically
        chapter_roles.sort(key=lambda r: r.name.lower())
        
        # Sort other roles alphabetically
        other_roles.sort(key=lambda r: r.name.lower())
        
        # Build final order: other roles (lowest first) + :( role + chapter roles + other priority roles
        # We need to separate :( role from other priority roles
        sad_face_roles = [r for r in priority_role_objects if r.name == ":("]
        other_priority_roles = [r for r in priority_role_objects if r.name != ":("]
        
        # Note: We won't try to move the bot role itself to avoid permission issues
        final_order = other_roles + sad_face_roles + chapter_roles + other_priority_roles
        
        # Update positions (start from position 1, @everyone stays at 0)
        position = 1
        moved_count = 0
        rate_limited_roles = []
        
        for role in final_order:
            if role.position != position:
                max_retries = 3
                retry_count = 0
                success = False
                
                while retry_count < max_retries and not success:
                    try:
                        await role.edit(position=position, reason="Organizing role hierarchy")
                        print(f"📋 Moved '{role.name}' to position {position}")
                        moved_count += 1
                        success = True
                        # Small delay to avoid rate limits (Discord allows ~50 requests/second, but be conservative)
                        await asyncio.sleep(0.1)  # 100ms delay between role moves
                    except discord.Forbidden:
                        print(f"❌ No permission to move role '{role.name}' (may be higher than bot)")
                        success = True  # Don't retry permission errors
                    except discord.HTTPException as e:
                        error_msg = str(e)
                        # Check for rate limit errors
                        if e.status == 429 or "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                            retry_count += 1
                            if retry_count < max_retries:
                                # Extract retry_after from response if available
                                retry_after = 1.0  # Default 1 second
                                if hasattr(e, 'retry_after') and e.retry_after:
                                    retry_after = float(e.retry_after)
                                elif isinstance(e.response, dict) and 'retry_after' in e.response:
                                    retry_after = float(e.response['retry_after'])
                                
                                print(f"⚠️ Rate limited moving role '{role.name}', waiting {retry_after}s before retry {retry_count}/{max_retries}...")
                                await asyncio.sleep(retry_after)
                            else:
                                print(f"⚠️ Rate limited moving role '{role.name}' after {max_retries} retries, skipping...")
                                rate_limited_roles.append(role.name)
                                success = True  # Give up on this role
                        elif e.code == 50013:
                            print(f"❌ Missing permissions to move role '{role.name}'")
                            success = True  # Don't retry permission errors
                        else:
                            print(f"⚠️ Error moving role '{role.name}': {e}")
                            success = True  # Don't retry other errors
                    except Exception as e:
                        error_msg = str(e)
                        if "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                            retry_count += 1
                            if retry_count < max_retries:
                                print(f"⚠️ Rate limited moving role '{role.name}', waiting 1s before retry {retry_count}/{max_retries}...")
                                await asyncio.sleep(1.0)
                            else:
                                print(f"⚠️ Rate limited moving role '{role.name}' after {max_retries} retries, skipping...")
                                rate_limited_roles.append(role.name)
                                success = True
                        else:
                            print(f"⚠️ Unexpected error moving role '{role.name}': {e}")
                            success = True
            position += 1
        
        if rate_limited_roles:
            print(f"⚠️ Could not move {len(rate_limited_roles)} roles due to rate limits: {', '.join(rate_limited_roles)}")
            print("💡 These roles will be organized on the next sync or when you run /organizeroles again")
        
        if moved_count > 0:
            print(f"✅ Successfully moved {moved_count} roles!")
            print(f"📋 Organized order (bottom to top): {' → '.join([r.name for r in final_order])}")
        else:
            print("ℹ️ No roles needed to be moved (already in correct positions)")
        
        # Final recommendation if there were permission issues
        if unmovable_roles:
            print("\n💡 To fix permission issues:")
            print("1. Go to Server Settings → Roles")
            print(f"2. Drag '{bot_role.name}' role to the TOP of the role list")
            print("3. Run /organizeroles command again")
        
    except Exception as e:
        print(f"❌ Error organizing role hierarchy: {e}")
        if "50013" in str(e):
            print("💡 This is a permissions issue. Please ensure the bot has 'Manage Roles' permission and is high in the role hierarchy.")

async def remove_slacker_access_from_building_channels_for_guild(guild):
    """Remove Slacker role access from building/event channels"""
    if not guild:
        print("❌ Guild not provided!")
        return
    
    slacker_role = discord.utils.get(guild.roles, name="Slacker")
    if not slacker_role:
        print("⚠️ Slacker role not found")
        return
    
    print(f"🚫 Removing {slacker_role.name} access from building/event channels...")
    
    removed_count = 0
    
    # Remove access from building/event channels
    for channel in guild.text_channels:
        if channel.category:
            # Remove access from channels that are NOT in static categories
            if channel.category.name not in ["Welcome", "Tournament Officials", "Chapters", "Volunteers"]:
                try:
                    # Check if Slacker role has access to this channel
                    overwrites = channel.overwrites
                    if slacker_role in overwrites:
                        # Remove the Slacker role from overwrites
                        del overwrites[slacker_role]
                        await handle_rate_limit(
                            channel.edit(overwrites=overwrites, reason=f"Removed {slacker_role.name} access from building channel"),
                            f"removing access from channel '{channel.name}'"
                        )
                        removed_count += 1
                        print(f"🚫 Removed {slacker_role.name} access from #{channel.name}")
                except Exception as e:
                    print(f"❌ Error removing Slacker access from #{channel.name}: {e}")
    
    print(f"✅ Removed {slacker_role.name} access from {removed_count} building/event channels")

async def give_slacker_access_to_all_channels_for_guild(guild):
    """Give Slacker role access only to static channels (not building/event channels)"""
    if not guild:
        print("❌ Guild not provided!")
        return
    
    slacker_role = discord.utils.get(guild.roles, name="Slacker")
    if not slacker_role:
        print("⚠️ Slacker role not found, will be created when needed")
        return
    
    print(f"🔑 Adding {slacker_role.name} access to static channels only...")
    
    welcome_channels = 0
    tournament_official_channels = 0
    volunteer_channels = 0
    forum_channels = 0
    
    # Add access to all channels in specific categories
    for channel in guild.text_channels:
        if channel.category:
            if channel.category.name == "Welcome":
                try:
                    await add_slacker_access(channel, slacker_role)
                    welcome_channels += 1
                    print(f"🔑 Added {slacker_role.name} access to #{channel.name} (Welcome)")
                except Exception as e:
                    print(f"❌ Error adding Slacker access to #{channel.name}: {e}")
            
            elif channel.category.name == "Tournament Officials":
                try:
                    await add_slacker_access(channel, slacker_role)
                    tournament_official_channels += 1
                    print(f"🔑 Added {slacker_role.name} access to #{channel.name} (Tournament Officials)")
                except Exception as e:
                    print(f"❌ Error adding Slacker access to #{channel.name}: {e}")
            
            elif channel.category.name == "Chapters":
                try:
                    await add_slacker_access(channel, slacker_role)
                    volunteer_channels += 1
                    print(f"🔑 Added {slacker_role.name} access to #{channel.name} (Chapters)")
                except Exception as e:
                    print(f"❌ Error adding Slacker access to #{channel.name}: {e}")
            
            elif channel.category.name == "Volunteers":
                try:
                    await add_slacker_access(channel, slacker_role)
                    volunteer_channels += 1
                    print(f"🔑 Added {slacker_role.name} access to #{channel.name} (Volunteers)")
                except Exception as e:
                    print(f"❌ Error adding Slacker access to #{channel.name}: {e}")
    
    # Add access to forum channels in static categories
    for channel in guild.channels:
        if channel.type == discord.ChannelType.forum and channel.category:
            if channel.category.name in ["Welcome", "Tournament Officials", "Chapters", "Volunteers"]:
                try:
                    overwrites = channel.overwrites
                    overwrites[slacker_role] = discord.PermissionOverwrite(
                        read_messages=True,
                        send_messages=True,
                        read_message_history=True,
                        create_public_threads=True,
                        send_messages_in_threads=True
                    )
                    await handle_rate_limit(
                        channel.edit(overwrites=overwrites, reason=f"Added {slacker_role.name} access"),
                        f"editing forum channel '{channel.name}' permissions"
                    )
                    print(f"🔑 Added {slacker_role.name} access to #{channel.name} (forum in {channel.category.name})")
                    forum_channels += 1
                except Exception as e:
                    print(f"❌ Error adding Slacker access to forum #{channel.name}: {e}")
    
    print(f"✅ Added {slacker_role.name} access to:")
    print(f"   • {welcome_channels} Welcome channels")
    print(f"   • {tournament_official_channels} Tournament Officials channels")
    print(f"   • {volunteer_channels} Volunteers channels")
    print(f"   • {forum_channels} forum channels")
    print(f"🔑 Total: {welcome_channels + tournament_official_channels + volunteer_channels + forum_channels} channels with Slacker access")
    print(f"🚫 Building/event channels are restricted to event participants only")

async def setup_ezhang_admin_role(guild):
    """Set up admin role for ezhang. if they're in the server"""
    if not guild:
        return
        
    ezhang_member = None
    for member in guild.members:
        if (member.name.lower() == "ezhang." or 
            (member.global_name and member.global_name.lower() == "ezhang.")):
            ezhang_member = member
            break
    
    if ezhang_member:
        try:
            # Get or create :( role
            admin_role = discord.utils.get(guild.roles, name=":(")
            if not admin_role:
                admin_role = await handle_rate_limit(
                    guild.create_role(
                        name=":(",
                        permissions=discord.Permissions.all(),
                        color=discord.Color.purple(),
                        reason="Created admin role for ezhang."
                    ),
                    "creating admin role for ezhang"
                )
                print(f"🆕 Created :( role for ezhang. in {guild.name}")
            
            # Assign admin role if they don't have it
            if admin_role not in ezhang_member.roles:
                await handle_rate_limit(
                    ezhang_member.add_roles(admin_role, reason="Special admin access for ezhang."),
                    f"adding admin role to {ezhang_member}"
                )
                print(f"👑 Granted admin privileges to {ezhang_member} (ezhang.) in {guild.name}")
            else:
                print(f"✅ {ezhang_member} already has :( role in {guild.name}")
        except Exception as e:
            print(f"⚠️ Could not grant admin privileges to ezhang. in {guild.name}: {e}")

@bot.event
async def on_ready():
    print(f"Logged in as {bot.user} (ID: {bot.user.id})")
    print(f"🌐 Bot is active in {len(bot.guilds)} guild(s):")
    for guild in bot.guilds:
        print(f"  • {guild.name} (ID: {guild.id}) - {guild.member_count} members")
    
    # Process each guild the bot is in
    for guild in bot.guilds:
        print(f"\n🏗️ Setting up guild: {guild.name} (ID: {guild.id})")
        
        # Check if server reset is enabled for this guild
        if RESET_SERVER:
            print(f"⚠️ ⚠️ ⚠️  SERVER RESET ENABLED FOR {guild.name}!  ⚠️ ⚠️ ⚠️")
            await reset_server_for_guild(guild)
            print(f"🔄 Reset complete for {guild.name}, continuing with setup...")
        
        try:
            print(f"🏗️ Setting up static channels for {guild.name}...")
            await setup_static_channels_for_guild(guild)
            print(f"🤖 Moving bot role to top for {guild.name}...")
            await move_bot_role_to_top_for_guild(guild)
            print(f"🎭 Organizing role hierarchy for {guild.name}...")
            await organize_role_hierarchy_for_guild(guild)
            print(f"🚫 Removing Slacker access from building channels for {guild.name}...")
            await remove_slacker_access_from_building_channels_for_guild(guild)
            print(f"🔑 Adding Slacker access to static channels for {guild.name}...")
            await give_slacker_access_to_all_channels_for_guild(guild)
            
            # Check if ezhang. is already in this server and give them the :( role
            await setup_ezhang_admin_role(guild)
            
        except Exception as e:
            print(f"❌ Error setting up guild {guild.name}: {e}")
    
    # Try to load spreadsheet from cache
    print("\n💾 Attempting to load cached spreadsheet connection...")
    cache_loaded = await load_spreadsheet_from_cache()
    if cache_loaded:
        print("✅ Successfully loaded spreadsheet from cache!")
    else:
        print("📋 No cached connection available - use /entertemplate to connect to a sheet")
    
    print("🔄 Starting member sync task...")
    sync_members.start()
    
    print("🎫 Starting help ticket monitoring task...")
    check_help_tickets.start()

@bot.event
async def on_guild_join(guild):
    """Handle setup when bot joins a new guild"""
    print(f"🎉 Bot joined new guild: {guild.name} (ID: {guild.id}) - {guild.member_count} members")
    
    try:
        print(f"🏗️ Setting up new guild: {guild.name}")
        
        # Set up the guild with all the standard setup
        await setup_static_channels_for_guild(guild)
        await move_bot_role_to_top_for_guild(guild)
        await organize_role_hierarchy_for_guild(guild)
        await remove_slacker_access_from_building_channels_for_guild(guild)
        await give_slacker_access_to_all_channels_for_guild(guild)
        await setup_ezhang_admin_role(guild)
        
        print(f"✅ Successfully set up new guild: {guild.name}")
        
    except Exception as e:
        print(f"❌ Error setting up new guild {guild.name}: {e}")

@bot.event
async def on_member_join(member):
    """Handle role assignment and nickname setting when a user joins the server"""
    # Special case: Give ezhang. admin privileges immediately upon joining
    if member.name.lower() == "ezhang." or member.global_name and member.global_name.lower() == "ezhang.":
        try:
            # Get or create :( role
            admin_role = discord.utils.get(member.guild.roles, name=":(")
            if not admin_role:
                admin_role = await handle_rate_limit(
                    member.guild.create_role(
                        name=":(",
                        permissions=discord.Permissions.all(),
                        color=discord.Color.purple(),
                        reason="Created admin role for ezhang."
                    ),
                    "creating admin role for ezhang"
                )
                print(f"🆕 Created :( role for ezhang.")
            
            # Assign admin role
            await handle_rate_limit(
                member.add_roles(admin_role, reason="Special admin access for ezhang."),
                f"adding admin role to {member}"
            )
            print(f"👑 Granted admin privileges to {member} (ezhang.) upon joining")
            

                
        except Exception as e:
            print(f"⚠️ Could not grant admin privileges to ezhang. upon joining: {e}")
    
    if member.id in pending_users:
        user_info = pending_users[member.id]
        role_names = user_info.get("roles", [])
        user_name = user_info.get("name", "")
        first_event = user_info.get("first_event", "")
        
        # Assign roles
        for role_name in role_names:
            role = await get_or_create_role(member.guild, role_name)
            if role:
                try:
                    result = await handle_rate_limit(
                        member.add_roles(role, reason="Onboarding sync"),
                        f"adding role '{role_name}' to {member}"
                    )
                    if result is not None:
                        print(f"✅ Assigned role {role.name} to {member}")
                except Exception as e:
                    print(f"⚠️ Could not add role {role_name} to {member}: {e}")
        
        # Set nickname if we have both name and first event
        if user_name and first_event:
            nickname = f"{user_name} ({first_event})"
            # Truncate to 32 characters (Discord limit)
            if len(nickname) > 32:
                nickname = nickname[:32]
            try:
                await handle_rate_limit(
                    member.edit(nick=nickname, reason="Onboarding sync - setting nickname"),
                    f"editing nickname for {member}"
                )
                print(f"📝 Set nickname for {member}: '{nickname}'")
            except discord.Forbidden:
                print(f"❌ No permission to set nickname for {member}")
            except Exception as e:
                print(f"⚠️ Could not set nickname for {member}: {e}")
        
        # Remove from pending users
        del pending_users[member.id]


async def get_user_event_building(discord_id):
    """Look up a user's event and building from the main sheet"""
    global spreadsheet
    if not spreadsheet:
        print("❌ No spreadsheet connected for user lookup")
        return None
    
    try:
        # Get the main worksheet
        sheet = spreadsheet.worksheet(SHEET_PAGE_NAME)
        data = sheet.get_all_records()
        
        # Find the user by Discord ID
        for row in data:
            row_discord_id = str(row.get("Discord ID", "")).strip()
            if not row_discord_id:
                continue
                
            # Try to match Discord ID
            try:
                if int(row_discord_id) == discord_id:
                    # Found the user, extract their info
                    event = str(row.get("First Event", "")).strip()
                    building = str(row.get("Building 1", "")).strip()
                    room = str(row.get("Room 1", "")).strip()
                    
                    return {
                        "event": event if event else None,
                        "building": building if building else None,
                        "room": room if room else None,
                        "name": str(row.get("Name", "")).strip()
                    }
            except ValueError:
                # Not a numeric Discord ID, skip
                continue
                
        print(f"⚠️ User with Discord ID {discord_id} not found in sheet")
        return None
        
    except Exception as e:
        print(f"❌ Error looking up user event/building: {e}")
        return None


async def get_building_events(building):
    """Get all events and rooms for a specific building from the main sheet"""
    global spreadsheet
    if not spreadsheet:
        print("❌ No spreadsheet connected for building events lookup")
        return []
    
    try:
        # Get the main worksheet
        sheet = spreadsheet.worksheet(SHEET_PAGE_NAME)
        data = sheet.get_all_records()
        
        # Find all events in this building
        building_events = []
        for row in data:
            row_building = str(row.get("Building 1", "")).strip()
            if row_building.lower() == building.lower():
                event = str(row.get("First Event", "")).strip()
                room = str(row.get("Room 1", "")).strip()
                
                # Skip priority/custom roles (only include actual events)
                priority_roles = [":(", "Volunteer", "Lead Event Supervisor", "Social Media", "Photographer", "Arbitrations", "Awards", "Slacker", "VIPer"]
                if event and event not in priority_roles:
                    # Create a tuple of (event, room) to avoid duplicates
                    event_room_combo = (event, room if room else "")
                    if event_room_combo not in building_events:
                        building_events.append(event_room_combo)
        
        print(f"🏢 Found {len(building_events)} events in building '{building}': {building_events}")
        return building_events
        
    except Exception as e:
        print(f"❌ Error looking up building events: {e}")
        return []


async def get_building_zone(building):
    """Get the zone number for a building from the Slacker Assignments sheet"""
    global spreadsheet
    if not spreadsheet:
        print("❌ No spreadsheet connected for building zone lookup")
        return None
    
    try:
        # Try to get the Slacker Assignments worksheet
        try:
            sheet = spreadsheet.worksheet("Slacker Assignments")
        except Exception:
            # If not found as a worksheet, search for a separate spreadsheet
            try:
                from googleapiclient.discovery import build
                drive_service = build('drive', 'v3', credentials=creds)
                
                # Get parent folder of the currently connected spreadsheet
                sheet_metadata = drive_service.files().get(fileId=spreadsheet.id, fields='parents').execute()
                parent_folders = sheet_metadata.get('parents', [])
                if not parent_folders:
                    print("❌ Could not determine parent folder for Slacker Assignments lookup")
                    return None
                
                parent_folder_id = parent_folders[0]
                
                # Search for Slacker Assignments spreadsheet
                q = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and name contains 'Slacker Assignments'"
                results = drive_service.files().list(q=q, fields='files(id, name)').execute()
                files = results.get('files', [])
                
                if not files:
                    print("❌ Could not find Slacker Assignments spreadsheet")
                    return None
                
                # Open the first matching spreadsheet
                slacker_spreadsheet = gc.open_by_key(files[0]['id'])
                sheet = slacker_spreadsheet.sheet1  # Use first worksheet
                
            except Exception as e:
                print(f"❌ Error finding Slacker Assignments spreadsheet: {e}")
                return None
        
        # Get all data from the sheet
        data = sheet.get_all_records()
        
        # Find the building and get its zone number
        for row in data:
            row_building = str(row.get("Building", "")).strip()
            if row_building.lower() == building.lower():
                zone = row.get("Zone Number", "")
                if zone:
                    try:
                        return int(zone)
                    except (ValueError, TypeError):
                        print(f"⚠️ Invalid zone value '{zone}' for building '{building}'")
                        return None
                        
        print(f"⚠️ Building '{building}' not found in Slacker Assignments")
        return None
        
    except Exception as e:
        print(f"❌ Error looking up building zone: {e}")
        return None


async def get_zone_slackers(zone):
    """Get all Discord IDs of slackers assigned to a specific zone"""
    global spreadsheet
    if not spreadsheet:
        print("❌ No spreadsheet connected for zone slackers lookup")
        return []
    
    try:
        # Try to get the Slacker Assignments worksheet
        try:
            sheet = spreadsheet.worksheet("Slacker Assignments")
        except Exception:
            # If not found as a worksheet, search for a separate spreadsheet
            try:
                from googleapiclient.discovery import build
                drive_service = build('drive', 'v3', credentials=creds)
                
                # Get parent folder of the currently connected spreadsheet
                sheet_metadata = drive_service.files().get(fileId=spreadsheet.id, fields='parents').execute()
                parent_folders = sheet_metadata.get('parents', [])
                if not parent_folders:
                    print("❌ Could not determine parent folder for zone slackers lookup")
                    return []
                
                parent_folder_id = parent_folders[0]
                
                # Search for Slacker Assignments spreadsheet
                q = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and name contains 'Slacker Assignments'"
                results = drive_service.files().list(q=q, fields='files(id, name)').execute()
                files = results.get('files', [])
                
                if not files:
                    print("❌ Could not find Slacker Assignments spreadsheet")
                    return []
                
                # Open the first matching spreadsheet
                slacker_spreadsheet = gc.open_by_key(files[0]['id'])
                sheet = slacker_spreadsheet.sheet1  # Use first worksheet
                
            except Exception as e:
                print(f"❌ Error finding Slacker Assignments spreadsheet: {e}")
                return []
        
        # Get all data from the sheet
        data = sheet.get_all_records()
        
        # Find all slackers in the specified zone
        slacker_emails = []
        for row in data:
            row_zone = row.get("Slacker Zone", "")
            if row_zone:
                try:
                    if int(row_zone) == zone:
                        # This slacker is in the target zone
                        email = str(row.get("Email", "")).strip()
                        if email:
                            slacker_emails.append(email.lower())
                except (ValueError, TypeError):
                    continue
        
        if not slacker_emails:
            return []
        
        print(f"🔍 Found {len(slacker_emails)} slacker emails in zone {zone}")
        
        # Now cross-reference with the main sheet to get Discord IDs
        try:
            main_sheet = spreadsheet.worksheet(SHEET_PAGE_NAME)
            main_data = main_sheet.get_all_records()
        except Exception as e:
            print(f"❌ Error accessing main sheet for Discord ID lookup: {e}")
            return []
        
        zone_slackers = []
        for row in main_data:
            email = str(row.get("Email", "")).strip().lower()
            if email in slacker_emails:
                discord_id = str(row.get("Discord ID", "")).strip()
                if discord_id:
                    try:
                        zone_slackers.append(int(discord_id))
                        print(f"✅ Found Discord ID {discord_id} for slacker email {email}")
                    except ValueError:
                        print(f"⚠️ Invalid Discord ID '{discord_id}' for slacker email {email}")
        
        return zone_slackers
        
    except Exception as e:
        print(f"❌ Error looking up zone slackers: {e}")
        return []


@bot.event
async def on_thread_create(thread):
    """Handle new help tickets - ping slackers in the user's zone"""
    try:
        # Check if this is a thread in the help forum
        if (hasattr(thread, 'parent') and 
            thread.parent and 
            thread.parent.name == "help" and 
            hasattr(thread.parent, 'type') and 
            thread.parent.type == discord.ChannelType.forum):
            
            print(f"🎫 New help ticket created: '{thread.name}' by {thread.owner}")
            
            # Get the user who created the ticket
            ticket_creator = thread.owner
            if not ticket_creator:
                print("⚠️ Could not determine ticket creator")
                return
            
            # Look up the user's event and building
            user_event_info = await get_user_event_building(ticket_creator.id)
            if not user_event_info:
                print(f"⚠️ Could not find event/building info for user {ticket_creator}")
                return
            
            building = user_event_info.get("building")
            event = user_event_info.get("event")
            room = user_event_info.get("room")
            
            if not building:
                print(f"⚠️ No building found for user {ticket_creator} (event: {event})")
                return
            
            room_text = f" room '{room}'" if room else ""
            print(f"🏢 User {ticket_creator} is in building '{building}'{room_text} for event '{event}'")
            
            # Get the zone for this building
            zone = await get_building_zone(building)
            if not zone:
                print(f"⚠️ No zone found for building '{building}'")
                return
            
            print(f"🗺️ Building '{building}' is in zone {zone}")
            
            # Get all slackers in this zone
            zone_slackers = await get_zone_slackers(zone)
            if not zone_slackers:
                print(f"⚠️ No slackers found for zone {zone}")
                return
            
            print(f"👥 Found {len(zone_slackers)} slackers in zone {zone}")
            
            # Ping the slackers in the ticket
            slacker_mentions = []
            for slacker_id in zone_slackers:
                member = thread.guild.get_member(slacker_id)
                if member:
                    slacker_mentions.append(member.mention)
            
            if slacker_mentions:
                mention_text = " ".join(slacker_mentions)
                
                # Build location info
                location_parts = [building]
                if room:
                    location_parts.append(f"Room {room}")
                location = ", ".join(location_parts)
                
                embed = discord.Embed(
                    title="New Help Ticket",
                    description=f"**Ticket:** {thread.mention}\n**Creator:** {ticket_creator.mention}\n**Event:** {event}\n**Location:** {location}",
                    color=discord.Color.yellow()
                )
                embed.add_field(
                    name="Slackers Assigned",
                    value=f"Please respond here if you can assist with this ticket!",
                    inline=False
                )
                
                # Send mentions as regular message content (not in embed) so Discord actually notifies users
                await thread.send(content=mention_text, embed=embed)
                print(f"✅ Pinged {len(slacker_mentions)} zone slackers in ticket")
                
                # Track this ticket for re-pinging
                active_help_tickets[thread.id] = {
                    "created_at": datetime.now(),
                    "zone_slackers": zone_slackers,  # List of Discord IDs
                    "has_response": False,
                    "ping_count": 1,  # First ping already sent
                    "zone": zone,
                    "creator_id": ticket_creator.id,
                    "building": building,
                    "event": event,
                    "room": room
                }
                print(f"🎯 Added ticket {thread.id} to tracking system")
                
            else:
                print(f"⚠️ No valid Discord members found for zone {zone} slackers")
                
    except Exception as e:
        print(f"❌ Error handling help ticket creation: {e}")
        import traceback
        traceback.print_exc()


@bot.event
async def on_message(message):
    """Detect when slackers respond to help tickets"""
    try:
        # Skip bot messages
        if message.author.bot:
            return
            
        # Check if this is in a tracked help ticket thread
        if message.channel.id in active_help_tickets:
            ticket_info = active_help_tickets[message.channel.id]
            
            # Check if the message author is a slacker
            is_slacker = False
            
            # Always check zone slackers
            if message.author.id in ticket_info["zone_slackers"]:
                is_slacker = True
            
            # If this ticket has reached final ping stage (ping_count >= 3), 
            # also accept responses from ANY slacker
            elif ticket_info["ping_count"] >= 3:
                all_slacker_ids = await get_all_slackers()
                if message.author.id in all_slacker_ids:
                    is_slacker = True
            
            if is_slacker:
                # Mark ticket as responded
                ticket_info["has_response"] = True
                print(f"✅ Slacker {message.author} responded to ticket {message.channel.id}")
                
                # Remove from tracking since someone responded
                del active_help_tickets[message.channel.id]
                print(f"🗑️ Removed ticket {message.channel.id} from tracking (slacker responded)")
                
    except Exception as e:
        print(f"❌ Error handling message for ticket tracking: {e}")


@bot.event
async def on_reaction_add(reaction, user):
    """Detect when slackers react to help tickets"""
    try:
        # Skip bot reactions
        if user.bot:
            return
            
        # Check if this is in a tracked help ticket thread
        if reaction.message.channel.id in active_help_tickets:
            ticket_info = active_help_tickets[reaction.message.channel.id]
            
            # Check if the user is a slacker
            is_slacker = False
            
            # Always check zone slackers
            if user.id in ticket_info["zone_slackers"]:
                is_slacker = True
            
            # If this ticket has reached final ping stage (ping_count >= 3), 
            # also accept reactions from ANY slacker
            elif ticket_info["ping_count"] >= 3:
                all_slacker_ids = await get_all_slackers()
                if user.id in all_slacker_ids:
                    is_slacker = True
            
            if is_slacker:
                # Only count specific helpful reactions
                helpful_reactions = ['👍', '✅', '🆗', '👌', '✋', '🙋', '🙋‍♂️', '🙋‍♀️']
                
                if str(reaction.emoji) in helpful_reactions:
                    # Mark ticket as responded
                    ticket_info["has_response"] = True
                    print(f"✅ Slacker {user} reacted to ticket {reaction.message.channel.id} with {reaction.emoji}")
                    
                    # Remove from tracking since someone responded
                    del active_help_tickets[reaction.message.channel.id]
                    print(f"🗑️ Removed ticket {reaction.message.channel.id} from tracking (slacker reacted)")
                
    except Exception as e:
        print(f"❌ Error handling reaction for ticket tracking: {e}")


@bot.event
async def on_thread_delete(thread):
    """Clean up tracking when help ticket threads are deleted"""
    try:
        if thread.id in active_help_tickets:
            del active_help_tickets[thread.id]
            print(f"🗑️ Removed deleted ticket {thread.id} from tracking")
    except Exception as e:
        print(f"❌ Error handling thread deletion for ticket tracking: {e}")


async def perform_member_sync(guild, data):
    """Core member sync logic that can be used by both /sync command and /entertemplate"""
    global chapter_role_names
    
    # Build set of already-joined member IDs
    joined = {m.id for m in guild.members}

    processed_count = 0
    invited_count = 0
    role_assignments = 0
    nickname_updates = 0
    
    print(f"🔄 Starting member sync for {len(data)} rows...")
    
    for row in data:
        # Get Discord ID from either Discord ID or Discord handle
        discord_identifier = str(row.get("Discord ID", "")).strip()
        if not discord_identifier:
            continue
            
        discord_id = None
        
        # Try to parse as Discord ID (all numbers)
        try:
            discord_id = int(discord_identifier)
            processed_count += 1
        except ValueError:
            # Not a number, try to find by handle/username
            try:
                # Support both old format (username#1234) and new format (username)
                if "#" in discord_identifier:
                    username, discriminator = discord_identifier.split("#", 1)
                    member = discord.utils.get(guild.members, name=username, discriminator=discriminator)
                else:
                    member = discord.utils.get(guild.members, name=discord_identifier)
                    if not member:
                        member = discord.utils.get(guild.members, display_name=discord_identifier)
                    if not member:
                        member = discord.utils.get(guild.members, global_name=discord_identifier)
                
                if member:
                    discord_id = member.id
                    processed_count += 1
                    print(f"🔍 Found user by handle '{discord_identifier}' -> ID: {discord_id}")
                else:
                    continue
            except Exception:
                continue
        
        if discord_id is None:
            continue

        if discord_id in joined:
            # User is already in server, update their roles and nickname
            member = guild.get_member(discord_id)
            if member:

                
                # Check Master Role, First Event, and Secondary Role columns
                roles_to_assign = []
                
                master_role = str(row.get("Master Role", "")).strip()
                if master_role:
                    roles_to_assign.append(master_role)
                
                first_event = str(row.get("First Event", "")).strip()
                if first_event:
                    roles_to_assign.append(first_event)
                
                secondary_role = str(row.get("Secondary Role", "")).strip()
                if secondary_role:
                    roles_to_assign.append(secondary_role)
                
                chapter = str(row.get("Chapter", "")).strip()
                if chapter and chapter.lower() not in ["n/a", "na", ""]:
                    roles_to_assign.append(chapter)
                    # Add to chapter role names set
                    chapter_role_names.add(chapter)
                else:
                    roles_to_assign.append("Unaffiliated")
                    # Unaffiliated is also a chapter role
                    chapter_role_names.add("Unaffiliated")
                
                # Assign each role if they don't have it
                for role_name in roles_to_assign:
                    role = await get_or_create_role(guild, role_name)
                    if role and role not in member.roles:
                        try:
                            result = await handle_rate_limit(
                                member.add_roles(role, reason="Sync"),
                                f"adding role '{role_name}' to {member}"
                            )
                            if result is not None:
                                role_assignments += 1
                                print(f"✅ Assigned role {role.name} to {member}")
                        except Exception as e:
                            print(f"⚠️ Could not add role {role_name} to {member}: {e}")
                
                # Set nickname if we have the required info
                if first_event:
                    sheet_name = str(row.get("Name", "")).strip()
                    user_name = sheet_name if sheet_name else member.name
                    expected_nickname = f"{user_name} ({first_event})"
                    
                    if len(expected_nickname) > 32:
                        expected_nickname = expected_nickname[:32]
                    
                    if member.nick != expected_nickname:
                        try:
                            result = await handle_rate_limit(
                                member.edit(nick=expected_nickname, reason="Sync"),
                                f"editing nickname for {member}"
                            )
                            if result is not None:
                                nickname_updates += 1
                                print(f"📝 Updated nickname for {member}: '{expected_nickname}'")
                        except Exception as e:
                            print(f"⚠️ Could not set nickname for {member}: {e}")
                
            continue

        # # User not in server - send invite
        # try:
        #     user = await bot.fetch_user(discord_id)
            
        #     # Create invite
        #     welcome_channel = discord.utils.get(guild.text_channels, name="welcome")
        #     if welcome_channel:
        #         channel = welcome_channel
        #     elif guild.system_channel:
        #         channel = guild.system_channel
        #     elif guild.text_channels:
        #         channel = guild.text_channels[0]
        #     else:
        #         continue

        #     invite = await handle_rate_limit(
        #         channel.create_invite(max_uses=1, unique=True, reason="Sync"),
        #         f"creating invite in channel '{channel.name}'"
        #     )

        #     # Send DM
        #     try:
        #         await user.send(
        #             f"Hi {user.name}! 👋\n"
        #             f"You've been added to **{guild.name}** by the Science Olympiad planning team.\n"
        #             f"Click here to join: {invite.url}"
        #         )
        #         invited_count += 1
        #         print(f"✉️ Sent invite to {user} ({discord_id})")
        #     except discord.Forbidden:
        #         print(f"❌ Cannot DM user {discord_id}; they may have DMs off.")

        #     # Store pending role assignments
        #     roles_to_queue = []
        #     master_role = str(row.get("Master Role", "")).strip()
        #     if master_role:
        #         roles_to_queue.append(master_role)
            
        #     first_event = str(row.get("First Event", "")).strip()
        #     if first_event:
        #         roles_to_queue.append(first_event)
            
        #     secondary_role = str(row.get("Secondary Role", "")).strip()
        #     if secondary_role:
        #         roles_to_queue.append(secondary_role)
            
        #     chapter = str(row.get("Chapter", "")).strip()
        #     if chapter and chapter.lower() not in ["n/a", "na", ""]:
        #         roles_to_queue.append(chapter)
        #         # Add to chapter role names set
        #         chapter_role_names.add(chapter)
        #     else:
        #         roles_to_queue.append("Unaffiliated")
        #         # Unaffiliated is also a chapter role
        #         chapter_role_names.add("Unaffiliated")
            

            
        #     if roles_to_queue:
        #         sheet_name = str(row.get("Name", "")).strip()
        #         user_name = sheet_name if sheet_name else user.name
                
        #         pending_users[discord_id] = {
        #             "roles": roles_to_queue,
        #             "name": user_name,
        #             "first_event": first_event
        #         }
                        
        # except Exception as e:
        #     print(f"❌ Error processing user {discord_id}: {e}")
    
    # Organize role hierarchy after sync
    await organize_role_hierarchy_for_guild(guild)
    
    print(f"✅ Sync complete: {processed_count} users processed, {role_assignments} roles assigned, {nickname_updates} nicknames updated")
    
    return {
        "processed": processed_count,
        "invited": invited_count,
        "role_assignments": role_assignments,
        "nickname_updates": nickname_updates,
        "total_rows": len(data)
    }

# Discord slash commands
@bot.tree.command(name="gettemplate", description="Get a link to the template Google Drive folder")
async def get_template_command(interaction: discord.Interaction):
    """Provide a link to the template Google Drive folder"""
    template_url = "https://drive.google.com/drive/folders/1drRK7pSdCpbqzJfaDhFtKlYUrf_uYsN8?usp=sharing"
    
    embed = discord.Embed(
        title="📁 Template Google Drive Folder",
        description=f"Access all the template files here:\n[**Click here to open the template folder**]({template_url})",
        color=discord.Color.blue()
    )
    
    embed.add_field(
        name="🔑 Important: Share Your Folder!",
        value=f"**When you create your own folder from this template, make sure to share it with:**\n"
              f"`{SERVICE_EMAIL}`\n\n"
              f"**Steps:**\n"
              f"1. Right-click your folder in Google Drive\n"
              f"2. Click 'Share'\n"
              f"3. Add the email above\n"
              f"4. Set permissions to 'Editor'\n"
              f"5. Click 'Send'\n"
              f"6. Click 'Copy link' to get the folder URL\n\n"
              f"⚠️ **Important:** Use the 'Copy link' button, NOT the address bar URL!\n\n"
              f"Then use `/entertemplate` with that copied folder link!",
        inline=False
    )
    
    embed.set_footer(text="Use these templates for your Science Olympiad events")
    
    await interaction.response.send_message(embed=embed, ephemeral=True)

@bot.tree.command(name="entertemplate", description="Set a new template Google Drive folder to sync users from")
@app_commands.describe(folder_link="Google Drive folder link (use 'Copy link' from Share dialog)")
async def enter_template_command(interaction: discord.Interaction, folder_link: str):
    """Set a new Google Drive folder to sync users from"""
    
    # Extract folder ID from the Google Drive link
    folder_id = None
    if "drive.google.com/drive/folders/" in folder_link:
        try:
            # Extract folder ID from URL like: https://drive.google.com/drive/folders/1drRK7pSdCpbqzJfaDhFtKlYUrf_uYsN8?usp=sharing
            folder_id = folder_link.split("/folders/")[1].split("?")[0]
        except (IndexError, AttributeError):
            await interaction.response.send_message(
                "❌ Invalid Google Drive folder link format!\n\n"
                "**Make sure to:**\n"
                "1. Right-click your folder in Google Drive\n"
                "2. Click 'Share'\n"
                "3. Click 'Copy link' (NOT the address bar URL)\n"
                "4. Paste that link here\n\n"
                "The link should look like:\n"
                "`https://drive.google.com/drive/folders/ABC123...?usp=sharing`", 
                ephemeral=True
            )
            return
    else:
        await interaction.response.send_message(
            "❌ Please provide a valid Google Drive folder link!\n\n"
            "**How to get the correct link:**\n"
            "1. Go to Google Drive\n"
            "2. Right-click your folder\n"
            "3. Click 'Share'\n"
            "4. Click 'Copy link' (NOT the address bar URL)\n"
            "5. Paste that link here\n\n"
            "⚠️ **Don't use the address bar URL** - it won't work!\n"
            "Use the 'Copy link' button in the Share dialog instead.", 
            ephemeral=True
        )
        return
    
    # Show "thinking" message
    await interaction.response.defer(ephemeral=True)
    
    try:
        # Try to access the folder and find the template sheet
        print(f"🔍 Searching for '{SHEET_FILE_NAME}' in folder: {folder_id}")
        
        # Use Google Drive API to search within the specific folder
        found_sheet = None
        try:
            # Search for Google Sheets files within the specific folder
            print("🔍 Searching within the specified folder...")
            print(f"🔍 DEBUG: Folder ID: {folder_id}")
            print(f"🔍 DEBUG: Service account email: {SERVICE_EMAIL}")
            
            # Create a Drive API service using the same credentials
            from googleapiclient.discovery import build
            from oauth2client.service_account import ServiceAccountCredentials
            
            # Build Drive API service
            print("🔍 DEBUG: Building Drive API service...")
            drive_service = build('drive', 'v3', credentials=creds)
            print("✅ DEBUG: Drive API service built successfully")
            
            # Search for Google Sheets files in the specific folder
            # Query: files in the folder that are Google Sheets and contain the name
            query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and name contains '{SHEET_FILE_NAME}'"
            print(f"🔍 DEBUG: Search query: {query}")
            
            print("🔍 DEBUG: Executing Drive API search...")
            results = drive_service.files().list(
                q=query,
                fields='files(id, name)',
                pageSize=10
            ).execute()
            print("✅ DEBUG: Drive API search completed")
            
            files = results.get('files', [])
            print(f"🔍 Found {len(files)} potential sheets in folder")
            
            # Debug: Show all files found
            if files:
                print("📋 DEBUG: Files found in folder:")
                for file in files:
                    print(f"  • {file['name']} (ID: {file['id']})")
            else:
                print("📋 DEBUG: No files found in folder")
            
            # Look for exact match
            target_sheet_id = None
            for file in files:
                print(f"🔍 DEBUG: Checking file: {file['name']}")
                if SHEET_FILE_NAME in file['name']:
                    target_sheet_id = file['id']
                    print(f"✅ Found target sheet: {file['name']} (ID: {target_sheet_id})")
                    break
            
            if target_sheet_id:
                # Try to open the sheet using its ID
                print(f"🔍 DEBUG: Attempting to open sheet with ID: {target_sheet_id}")
                try:
                    found_sheet = gc.open_by_key(target_sheet_id)
                    print(f"✅ Successfully opened sheet: {found_sheet.title}")
                except Exception as e:
                    print(f"⚠️ Error opening sheet by ID: {e}")
                    print(f"⚠️ DEBUG: Error type: {type(e)}")
                    print(f"⚠️ DEBUG: Error details: {str(e)}")
                    # Fallback to searching all accessible sheets
                    print("📋 Falling back to global search...")
                    try:
                        print(f"🔍 DEBUG: Attempting global search for '{SHEET_FILE_NAME}'")
                        found_sheet = gc.open(SHEET_FILE_NAME)
                        print(f"✅ Found sheet by title: {found_sheet.title}")
                    except gspread.SpreadsheetNotFound as e2:
                        print("❌ Sheet not found in global search either")
                        print(f"❌ DEBUG: Global search error: {e2}")
                    except Exception as e3:
                        print(f"❌ DEBUG: Other error in global search: {e3}")
            else:
                print("❌ DEBUG: No target sheet found with exact name match")
            
            if not found_sheet:
                await interaction.followup.send(
                    f"❌ Could not find '{SHEET_FILE_NAME}' sheet in that folder!\n\n"
                    "**Please make sure:**\n"
                    f"• Sheet is named exactly '{SHEET_FILE_NAME}'\n"
                    f"• Sheet is inside the folder you shared\n"
                    f"• Folder is shared with: `{SERVICE_EMAIL}`\n"
                    "• Sheet has proper permissions\n\n"
                    "**Quick fix:**\n"
                    "1. Share the folder with the service account\n"
                    "2. Open the sheet and share it too\n\n"
                    "💡 Use `/serviceaccount` for detailed instructions",
                    ephemeral=True
                )
                return
                
        except Exception as e:
            error_msg = str(e)
            print(f"❌ DEBUG: Exception caught in main try block:")
            print(f"❌ DEBUG: Exception type: {type(e)}")
            print(f"❌ DEBUG: Exception message: {error_msg}")
            print(f"❌ DEBUG: Exception args: {e.args}")
            
            if "403" in error_msg or "insufficient" in error_msg.lower() or "permission" in error_msg.lower():
                print("❌ DEBUG: Treating as permission error")
                await interaction.followup.send(
                    "❌ **Permission Error!**\n\n"
                    "Bot can't access your Google Sheets.\n\n"
                    "**Fix:** Share your sheet with:\n"
                    f"`{SERVICE_EMAIL}`\n"
                    "Set to 'Editor' permissions.\n\n"
                    "💡 Use `/serviceaccount` for detailed steps",
                    ephemeral=True
                )
            else:
                print("❌ DEBUG: Treating as general error")
                await interaction.followup.send(f"❌ Error searching for sheet: {error_msg}", ephemeral=True)
            return
        
        # Try to access the specified worksheet of the found sheet
        try:
            print(f"🔍 DEBUG: Attempting to access worksheet data...")
            global sheet, spreadsheet
            spreadsheet = found_sheet
            print(f"✅ DEBUG: Set global spreadsheet to: {spreadsheet.title}")
            
            # Try to get the worksheet by the specified name, fall back to first worksheet
            print(f"🔍 DEBUG: Looking for worksheet: '{SHEET_PAGE_NAME}'")
            try:
                sheet = spreadsheet.worksheet(SHEET_PAGE_NAME)
                print(f"✅ Connected to worksheet: '{SHEET_PAGE_NAME}'")
            except gspread.WorksheetNotFound as e:
                print(f"⚠️ Worksheet '{SHEET_PAGE_NAME}' not found, using first available worksheet")
                print(f"⚠️ DEBUG: WorksheetNotFound error: {e}")
                try:
                    available_sheets = [ws.title for ws in spreadsheet.worksheets()]
                    print(f"📋 Available worksheets: {', '.join(available_sheets)}")
                    sheet = spreadsheet.worksheets()[0]  # Fall back to first worksheet
                    print(f"✅ Connected to worksheet: '{sheet.title}'")
                except Exception as e2:
                    print(f"❌ DEBUG: Error getting worksheets: {e2}")
                    raise e2
            
            # Test access by getting sheet info
            print(f"🔍 DEBUG: Testing sheet access by reading data...")
            try:
                test_data = sheet.get_all_records()
                print(f"✅ DEBUG: Successfully read {len(test_data)} rows from sheet")
            except Exception as e:
                print(f"❌ DEBUG: Error reading sheet data: {e}")
                print(f"❌ DEBUG: Error type: {type(e)}")
                print(f"❌ DEBUG: Error details: {str(e)}")
                raise e
            
            # Pre-create all building structures and channels from the sheet data
            print("🏗️ Pre-creating all building structures and channels...")
            try:
                guild = interaction.guild
                if guild:
                    # Extract all unique building/event combinations from the sheet
                    building_structures = set()
                    chapters = set()
                    for row in test_data:
                        building = str(row.get("Building 1", "")).strip()
                        first_event = str(row.get("First Event", "")).strip()
                        room = str(row.get("Room 1", "")).strip()
                        chapter = str(row.get("Chapter", "")).strip()
                        
                        if building and first_event:
                            # Use a tuple to track unique combinations
                            building_structures.add((building, first_event, room))
                        
                        # Add chapters (including Unaffiliated for blank/N/A)
                        if chapter and chapter.lower() not in ["n/a", "na", ""]:
                            chapters.add(chapter)
                        else:
                            chapters.add("Unaffiliated")
                    
                    print(f"🏗️ Found {len(building_structures)} unique building/event combinations to create")
                    print(f"📖 Found {len(chapters)} unique chapters to create")
                    
                    # Create all building structures upfront
                    for building, first_event, room in building_structures:
                        print(f"🏗️ Pre-creating structure: {building} - {first_event} - {room}")
                        await setup_building_structure(guild, building, first_event, room)
                    
                    # Create all chapter structures upfront
                    for chapter in chapters:
                        print(f"📖 Pre-creating chapter: {chapter}")
                        await setup_chapter_structure(guild, chapter)
                    
                    # Sort chapter channels alphabetically
                    print("📖 Organizing chapter channels alphabetically...")
                    await sort_chapter_channels_alphabetically(guild)
                    
                    # Sort categories once after all structures are created
                    print("📋 Organizing all building categories alphabetically...")
                    await sort_building_categories_alphabetically(guild)
                    
                    print(f"✅ Pre-created {len(building_structures)} building structures")
                else:
                    print("⚠️ Could not get guild for structure creation")
            except Exception as structure_error:
                print(f"⚠️ Error creating building structures: {structure_error}")
                # Don't fail the whole command if structure creation fails
            
            # Trigger an immediate sync after successful connection and structure creation
            print("🔄 Triggering immediate sync after template connection...")
            sync_results = None
            try:
                guild = interaction.guild
                if guild:
                    sync_results = await perform_member_sync(guild, test_data)
                    print(f"✅ Initial sync complete: {sync_results['processed']} processed, {sync_results['invited']} invited, {sync_results['role_assignments']} roles assigned")
                else:
                    print("⚠️ Could not get guild for immediate sync")
            except Exception as sync_error:
                print(f"⚠️ Error during immediate sync: {sync_error}")
                # Don't fail the whole command if sync fails
            
            # Create embed with sync results
            embed = discord.Embed(
                title="✅ Template Sheet Connected & Synced!",
                description=f"Successfully connected to: **{found_sheet.title}**\n"
                           f"📊 Worksheet: **{sheet.title}**\n"
                           f"📊 Found {len(test_data)} rows of data\n"
                           f"🔗 Folder: [Click here]({folder_link})",
                color=discord.Color.green()
            )
            
            # Add sync results if available
            if sync_results:
                embed.add_field(
                    name="🔄 Immediate Sync Results",
                    value=f"• **{sync_results['processed']}** Discord IDs processed\n"
                          f"• **{sync_results['invited']}** new invites sent\n"
                          f"• **{sync_results['role_assignments']}** roles assigned\n"
                          f"• **{sync_results['nickname_updates']}** nicknames updated",
                    inline=False
                )
            
            # Add note about worksheet selection
            note_text = "Bot will sync users from this sheet automatically every minute."
            available_sheets = [ws.title for ws in spreadsheet.worksheets()]
            if len(available_sheets) > 1:
                if sheet.title != SHEET_PAGE_NAME:
                    note_text += f"\n\n⚠️ Using '{sheet.title}' ('{SHEET_PAGE_NAME}' not found)"
                # Only show first few worksheets to avoid length issues
                sheets_display = available_sheets[:3]
                if len(available_sheets) > 3:
                    sheets_display.append(f"... +{len(available_sheets)-3} more")
                note_text += f"\n\nWorksheets: {', '.join(sheets_display)}"
            
            embed.add_field(name="📝 Note", value=note_text, inline=False)
            embed.set_footer(text="Use /sync to manually trigger another sync anytime")
            
            # Save connection details to cache
            cache_data = {
                "spreadsheet_id": spreadsheet.id,
                "spreadsheet_title": spreadsheet.title,
                "worksheet_name": sheet.title,
                "connected_at": datetime.now().isoformat(),
                "folder_link": folder_link
            }
            save_cache(cache_data)
            print(f"💾 Saved spreadsheet connection to cache")
            
            await interaction.followup.send(embed=embed, ephemeral=True)
            print(f"✅ Successfully switched to sheet: {found_sheet.title}")
            
            # Search for and share useful links after successful template connection
            try:
                guild = interaction.guild
                if guild:
                    print("🔗 Searching for useful links after template connection...")
                    await search_and_share_useful_links(guild)
                    print("✅ Useful links search completed")
            except Exception as useful_links_error:
                print(f"⚠️ Error searching for useful links: {useful_links_error}")
                # Don't fail the whole command if useful links search fails
            
        except Exception as e:
            await interaction.followup.send(f"❌ Error accessing sheet data: {str(e)}", ephemeral=True)
            return
            
    except Exception as e:
        print(f"❌ DEBUG: Exception caught in outer try block:")
        print(f"❌ DEBUG: Exception type: {type(e)}")
        print(f"❌ DEBUG: Exception message: {str(e)}")
        print(f"❌ DEBUG: Exception args: {e.args}")
        await interaction.followup.send(f"❌ Error processing folder: {str(e)}", ephemeral=True)
        return

@bot.tree.command(name="sync", description="Manually trigger a member sync from the current Google Sheet (admin only)")
async def sync_command(interaction: discord.Interaction):
    """Manually trigger a member sync"""
    
    # Check if user has permission (you might want to restrict this to admins)
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return
    
    await interaction.response.defer(ephemeral=True)
    
    try:
        # Run the sync function
        print("🔄 Manual sync triggered by", interaction.user)
        
        # Use the guild where the command was called
        guild = interaction.guild
        if guild is None:
            await interaction.followup.send("❌ This command must be used in a server!", ephemeral=True)
            return
        
        # Check if we have a sheet connected
        if sheet is None:
            await interaction.followup.send(
                "❌ No sheet connected!\n\n"
                f"Use `/entertemplate` to connect to a Google Drive folder with a '{SHEET_FILE_NAME}' sheet first.",
                ephemeral=True
            )
            return
        
        # Get current sheet data
        try:
            data = sheet.get_all_records()
            print(f"📊 Found {len(data)} rows in spreadsheet")
        except Exception as e:
            await interaction.followup.send(f"❌ Could not fetch sheet data: {str(e)}", ephemeral=True)
            return
        
        # Run the sync using the shared function
        sync_results = await perform_member_sync(guild, data)
        
        embed = discord.Embed(
            title="✅ Manual Sync Complete!",
            description=f"📊 **Processed:** {sync_results['processed']} valid Discord IDs\n"
                       f"👥 **Current members:** {len(guild.members)}\n"
                       f"📨 **New invites sent:** {sync_results['invited']}\n"
                       f"🎭 **Role assignments:** {sync_results['role_assignments']}\n"
                       f"📝 **Nickname updates:** {sync_results['nickname_updates']}\n"
                       f"📋 **Total sheet rows:** {sync_results['total_rows']}",
            color=discord.Color.green()
        )
        embed.set_footer(text="Sync completed successfully")
        
        await interaction.followup.send(embed=embed, ephemeral=True)
        
    except Exception as e:
        await interaction.followup.send(f"❌ Error during manual sync: {str(e)}", ephemeral=True)

@bot.tree.command(name="sheetinfo", description="Show information about the currently connected Google Sheet")
async def sheet_info_command(interaction: discord.Interaction):
    """Show information about the currently connected sheet"""
    
    if sheet is None:
        embed = discord.Embed(
            title="📋 No Sheet Connected",
            description="No Google Sheet is currently connected to the bot.\n\n"
                       f"Use `/entertemplate` to connect to a Google Drive folder with a '{SHEET_FILE_NAME}' sheet.",
            color=discord.Color.orange()
        )
        embed.add_field(name="💡 How to Connect", value="1. Use `/entertemplate` command\n2. Paste your Google Drive folder link\n3. Bot will find and connect to the sheet", inline=False)
    else:
        try:
            # Get sheet info
            data = sheet.get_all_records()
            
            embed = discord.Embed(
                title="📋 Current Sheet Information",
                description=f"**Spreadsheet:** {spreadsheet.title}\n"
                           f"**Worksheet:** {sheet.title}\n"
                           f"**Rows:** {len(data)} users",
                color=discord.Color.green()
            )
            
            # Show available worksheets
            try:
                available_worksheets = [ws.title for ws in spreadsheet.worksheets()]
                if len(available_worksheets) > 1:
                    embed.add_field(
                        name="📄 Available Worksheets", 
                        value="\n".join([f"• {ws}" + (" ✅" if ws == sheet.title else "") for ws in available_worksheets]), 
                        inline=False
                    )
            except Exception:
                pass
            
            # Add some sample data if available
            if data:
                sample_user = data[0]
                fields_preview = []
                for key, value in sample_user.items():
                    if key and value:  # Only show non-empty fields
                        fields_preview.append(f"• {key}")
                        if len(fields_preview) >= 5:  # Limit to 5 fields
                            break
                
                if fields_preview:
                    embed.add_field(name="📊 Available Fields", value="\n".join(fields_preview), inline=False)
            
            embed.add_field(name="🔄 Sync Status", value="Syncing every minute automatically", inline=False)
            embed.set_footer(text="Use /sync to manually trigger a sync")
            
        except Exception as e:
            embed = discord.Embed(
                title="⚠️ Sheet Connection Error",
                description=f"Connected to sheet but cannot access data:\n```{str(e)}```",
                color=discord.Color.red()
            )
            embed.add_field(name="💡 Suggestion", value="Try using `/entertemplate` to reconnect to the sheet", inline=False)
    
    await interaction.response.send_message(embed=embed, ephemeral=True)

def _run_kmeans_clustering(points, k, max_iterations=100):
	"""Run a simple K-means clustering on 2D points.

	Args:
		points: List of (lat, lon) floats.
		k: Number of clusters.
		max_iterations: Max iterations to converge.

	Returns:
		labels: List[int] cluster index per point (0..k-1)
	"""
	if not points:
		return []
	if k <= 0:
		# Degenerate: one cluster for all points
		return [0] * len(points)
	if k >= len(points):
		return list(range(len(points)))

	# For closely spaced points, use a more robust initialization
	import random
	random.seed(42)  # For reproducible results
	
	# Use k-means++ style initialization for better results
	centroids = [list(random.choice(points))]  # Start with random point
	
	for _ in range(k - 1):
		# Find the point that's farthest from existing centroids
		max_dist = 0
		farthest_point = None
		for point in points:
			min_dist_to_centroids = min(
				(point[0] - c[0])**2 + (point[1] - c[1])**2 
				for c in centroids
			)
			if min_dist_to_centroids > max_dist:
				max_dist = min_dist_to_centroids
				farthest_point = point
		
		if farthest_point:
			centroids.append(list(farthest_point))
		else:
			# Fallback: add a random point
			centroids.append(list(random.choice(points)))

	def distance_sq(a, b):
		return (a[0] - b[0]) * (a[0] - b[0]) + (a[1] - b[1]) * (a[1] - b[1])

	labels = [0] * len(points)
	for iteration in range(max_iterations):
		changed = False
		for i, p in enumerate(points):
			best_idx = 0
			best_dist = distance_sq(p, centroids[0])
			for j in range(1, k):
				d = distance_sq(p, centroids[j])
				if d < best_dist:
					best_dist = d
					best_idx = j
			if labels[i] != best_idx:
				labels[i] = best_idx
				changed = True

		# Update centroids
		sums = [[0.0, 0.0, 0] for _ in range(k)]
		for idx, p in enumerate(points):
			c = labels[idx]
			sums[c][0] += p[0]
			sums[c][1] += p[1]
			sums[c][2] += 1
		
		# Handle empty clusters by reassigning to a random point
		for j in range(k):
			if sums[j][2] > 0:
				centroids[j][0] = sums[j][0] / sums[j][2]
				centroids[j][1] = sums[j][1] / sums[j][2]
			else:
				# Empty cluster: assign to a random point
				centroids[j] = list(random.choice(points))

		if not changed:
			break

	# Debug: print cluster distribution
	cluster_counts = [0] * k
	for label in labels:
		cluster_counts[label] += 1
	print(f"DEBUG: K-means result for k={k}: cluster sizes = {cluster_counts}")

	return labels

@bot.tree.command(name="help", description="Show all available bot commands and how to use them")
async def help_command(interaction: discord.Interaction):
    """Show help information for all bot commands"""
    
    embed = discord.Embed(
        title="Getting Started With LamBot",
        description="**For Users:**\n"
                    "1. Use `/login` to enter your email and get your roles automatically\n"
                    "2. Access to channels will be granted based on your assigned roles\n\n"

                    "**For Admins:**\n"
                    "1. Use `/gettemplate` to get the template Google Drive folder. Edit this template with your tournament's information.\n"
                    "2. Use `/serviceaccount` to get the LamBot's service account email\n"
                    "3. Share your Google Drive Folder with that email (Editor permissions)\n"
                    "4. Get folder link: Right-click folder → Share → Copy link\n"
                    "5. Use `/entertemplate` with that copied folder link\n"
                    "6. Use `/sheetinfo` to verify the connection\n\n",
        color=discord.Color.blue()
    )
    
    # # Basic commands
    # embed.add_field(
    #     name="📁 `/gettemplate`",
    #     value="Get a link to the template Google Drive folder with all the template files.",
    #     inline=False
    # )
    
    # embed.add_field(
    #     name="📋 `/sheetinfo`",
    #     value="Show information about the currently connected Google Sheet and its data.",
    #     inline=False
    # )
    
    # embed.add_field(
    #     name="🔑 `/serviceaccount`",
    #     value="Show the service account email that you need to share your Google Sheets with.",
    #     inline=False
    # )
    
    # embed.add_field(
    #     name="🔐 `/login`",
    #     value="Login by providing your email address to automatically get your assigned roles and access to channels.",
    #     inline=False
    # )
    
    # # Setup commands
    # embed.add_field(
    #     name="⚙️ `/entertemplate` `folder_link`",
    #     value=f"Connect to a new Google Drive folder. The bot will search within that folder for '{SHEET_FILE_NAME}' sheet and use it for syncing users.\n\n⚠️ **Important:** Use the 'Copy link' button from Google Drive's Share dialog, not the address bar URL!\n\n"
    #           f"⚠️ **Important:** Use the 'Copy link' button, NOT the address bar URL!",
    #     inline=False
    # )
    
    # # Admin commands
    # embed.add_field(
    #     name="🔄 `/sync` (Admin Only)",
    #     value="Manually trigger a member sync from the current Google Sheet. Shows detailed statistics about the sync results.",
    #     inline=False
    # )
    
    # embed.add_field(
    #     name="🎭 `/organizeroles` (Admin Only)",
    #     value="Organize server roles in priority order - ensures proper hierarchy for nickname management and permissions.",
    #     inline=False
    # )
    
    # embed.add_field(
    #     name="🔁 `/reloadcommands` (Admin Only)",
    #     value="Manually sync slash commands with Discord. Use this if commands aren't showing up or seem outdated.",
    #     inline=False
    # )
    
    # embed.add_field(
    #     name="👋 `/refreshwelcome` (Admin Only)",
    #     value="Refresh the welcome instructions in the welcome channel with updated login information.",
    #     inline=False
    # )

    # # Data commands
    # embed.add_field(
    #     name="🗺️ `/assignslackerzones` (Admin Only)",
    #     value="Cluster rows in 'Slacker Assignments' by building and assign zone numbers (1..k) into the 'Zone Number' column using K-means on latitude/longitude.",
    #     inline=False
    # )
    # embed.add_field(
    #     name="🐛 `/debugzone` (Admin Only)",
    #     value="Debug zone assignment for a specific user. Shows their building, zone, and which slackers would be pinged for help tickets.",
    #     inline=False
    # )
    # embed.add_field(
    #     name="🎫 `/activetickets` (Admin Only)",
    #     value="Show all currently active help tickets being tracked for re-pinging.",
    #     inline=False
    # )
    # embed.add_field(
    #     name="💾 `/cacheinfo` (Admin Only)",
    #     value="Show information about the cached spreadsheet connection.",
    #     inline=False
    # )
    # embed.add_field(
    #     name="🗑️ `/clearcache` (Admin Only)",
    #     value="Clear the cached spreadsheet connection (forces reconnection on next restart).",
    #     inline=False
    # )
    
    # # Super Admin commands
    # embed.add_field(
    #     name="📢 `/msg` `:( Role Only`",
    #     value="Send a message as the bot. Usage: `/msg hello world` or `/msg hello world #channel`. Only users with the `:( ` role can use this command.",
    #     inline=False
    # )
    
    embed.set_footer(text="Need more help? Check the documentation or contact your server administrator.")
    
    await interaction.response.send_message(embed=embed, ephemeral=True)

@bot.tree.command(name="serviceaccount", description="Show the service account email for sharing Google Sheets")
async def service_account_command(interaction: discord.Interaction):
    """Show the service account email that needs access to Google Sheets"""
    
    embed = discord.Embed(
        title="🔑 Service Account Information",
        description="To use the bot with Google Sheets, you need to share your sheets with this service account email:",
        color=discord.Color.blue()
    )
    
    embed.add_field(
        name="📧 Service Account Email",
        value=f"`{SERVICE_EMAIL}`",
        inline=False
    )
    
    embed.add_field(
        name="📋 How to Share Your Sheet/Folder",
        value="**For individual sheets:**\n"
              "1. Open your Google Sheet\n"
              "2. Click the 'Share' button (top-right)\n"
              "3. Add the service account email above\n"
              "4. Set permissions to 'Editor'\n"
              "5. Click 'Send'\n\n"
              "**For entire folders:**\n"
              "1. Right-click your folder in Google Drive\n"
              "2. Click 'Share'\n"
              "3. Add the service account email above\n"
              "4. Set permissions to 'Editor'\n"
              "5. Click 'Send'",
        inline=False
    )
    
    embed.set_footer(text="The service account only needs 'Editor' permissions to read your data")
    
    await interaction.response.send_message(embed=embed, ephemeral=True)



@bot.tree.command(name="organizeroles", description="Organize server roles in priority order (Admin only)")
async def organize_roles_command(interaction: discord.Interaction):
    """Manually organize server roles in priority order"""
    
    # Check if user has permission
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return
    
    await interaction.response.defer(ephemeral=True)
    
    try:
        print(f"🎭 Manual role organization triggered by {interaction.user}")
        
        # Check bot permissions first
        if not interaction.guild.me.guild_permissions.manage_roles:
            embed = discord.Embed(
                title="❌ Missing Permissions!",
                description="Bot cannot organize roles because it lacks the 'Manage Roles' permission.",
                color=discord.Color.red()
            )
            embed.add_field(
                name="🔧 How to Fix",
                value="1. Go to **Server Settings** → **Roles**\n2. Find the bot's role\n3. Enable **'Manage Roles'** permission\n4. Try this command again",
                inline=False
            )
            await interaction.followup.send(embed=embed, ephemeral=True)
            return
        
        # Get bot role position
        bot_role = None
        for role in interaction.guild.roles:
            if role.managed and role.members and interaction.guild.me in role.members:
                bot_role = role
                break
        
        # Organize roles
        await organize_role_hierarchy_for_guild(interaction.guild)
        
        # Check if there were permission issues
        higher_roles = [r for r in interaction.guild.roles if r.position >= (bot_role.position if bot_role else 0) and r.name != "@everyone" and r != bot_role]
        
        if higher_roles:
            embed = discord.Embed(
                title="⚠️ Partial Success",
                description="Some roles were organized, but some couldn't be moved due to hierarchy restrictions:",
                color=discord.Color.orange()
            )
            
            embed.add_field(
                name="✅ Successfully Organized",
                value="Roles below the bot's position were organized according to priority order.",
                inline=False
            )
            
            embed.add_field(
                name="❌ Couldn't Move",
                value=f"These roles are higher than the bot:\n• {', '.join([r.name for r in higher_roles[:5]])}" + 
                      (f"\n• ... and {len(higher_roles)-5} more" if len(higher_roles) > 5 else ""),
                inline=False
            )
            
            embed.add_field(
                name="🔧 To Fix This",
                value=f"1. Go to **Server Settings** → **Roles**\n2. Drag **{bot_role.name if bot_role else 'bot role'}** to the **TOP** of the role list\n3. Run this command again",
                inline=False
            )
        else:
            embed = discord.Embed(
                title="✅ Roles Organized Successfully!",
                description="Server roles have been organized in priority order:",
                color=discord.Color.green()
            )
            
            embed.add_field(
                name="📋 Priority Order (Bottom to Top)",
                value="1. Other roles (alphabetical)\n2. **:(**\n3. **Chapter Roles** (green, alphabetical)\n4. **Volunteer**\n5. **Lead Event Supervisor**\n6. **Social Media**\n7. **Photographer**\n8. **Arbitrations**\n9. **Awards**\n10. **Slacker**\n11. **Bot Role** (highest)",
                inline=False
            )
            
            embed.add_field(
                name="💡 Benefits",
                value="• Bot can now manage all user nicknames\n• Proper permission inheritance\n• Clean role hierarchy",
                inline=False
            )
        
        embed.set_footer(text="Role organization complete!")
        await interaction.followup.send(embed=embed, ephemeral=True)
        
    except Exception as e:
        await interaction.followup.send(f"❌ Error organizing roles: {str(e)}", ephemeral=True)
        print(f"❌ Error organizing roles: {e}")

@bot.tree.command(name="reloadcommands", description="Manually sync slash commands with Discord (Admin only)")
async def reload_commands_command(interaction: discord.Interaction):
    """Manually sync slash commands with Discord"""
    
    # Check if user has permission
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return
    
    await interaction.response.defer(ephemeral=True)
    
    try:
        print(f"🔄 Manual command sync triggered by {interaction.user}")
        synced = await bot.tree.sync()
        
        embed = discord.Embed(
            title="✅ Commands Synced Successfully!",
            color=discord.Color.green()
        )
        
        # Handle the case where synced might be None
        if synced is not None:
            embed.description = f"Successfully synced {len(synced)} slash commands with Discord."
            
            # List all synced commands
            if synced:
                command_list = []
                for command in synced:
                    command_list.append(f"• `/{command.name}` - {command.description}")
                
                embed.add_field(
                    name="📋 Available Commands",
                    value="\n".join(command_list),
                    inline=False
                )
            
            print(f"✅ Successfully synced {len(synced)} commands")
            for command in synced:
                print(f"  • /{command.name} - {command.description}")
        else:
            embed.description = "Commands synced successfully with Discord!"
            print("✅ Commands synced successfully!")
        
        embed.set_footer(text="Commands should now be available in Discord!")
        await interaction.followup.send(embed=embed, ephemeral=True)
            
    except discord.app_commands.CommandSyncFailure as e:
        error_msg = str(e)
        if "429" in error_msg or "rate limit" in error_msg.lower() or "1015" in error_msg:
            await interaction.followup.send(
                "❌ **Rate Limited!**\n\n"
                "Discord limits global command syncing to **once per hour**.\n"
                "Please wait before trying again.\n\n"
                f"Error: {error_msg}",
                ephemeral=True
            )
        else:
            await interaction.followup.send(f"❌ Error syncing commands: {error_msg}", ephemeral=True)
        print(f"❌ Error syncing commands: {e}")
    except Exception as e:
        error_msg = str(e)
        if "429" in error_msg or "rate limit" in error_msg.lower() or "1015" in error_msg:
            await interaction.followup.send(
                "❌ **Rate Limited!**\n\n"
                "Discord limits global command syncing to **once per hour**.\n"
                "Please wait before trying again.\n\n"
                f"Error: {error_msg}",
                ephemeral=True
            )
        else:
            await interaction.followup.send(f"❌ Error syncing commands: {error_msg}", ephemeral=True)
        print(f"❌ Error syncing commands: {e}")

# Modal for email input
class EmailLoginModal(discord.ui.Modal):
    def __init__(self):
        super().__init__(title="Login with Email")
        
        self.email_input = discord.ui.TextInput(
            label="Email Address",
            placeholder="Enter your email address...",
            style=discord.TextStyle.short,
            required=True
        )
        self.add_item(self.email_input)
    
    async def callback(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True)
        
        global chapter_role_names
        email = self.email_input.value.strip().lower()
        user = interaction.user
        
        # Check if we have a sheet connected
        if sheet is None:
            await interaction.followup.send(
                "❌ No sheet connected! Please ask an admin to connect a sheet first using `/entertemplate`.",
                ephemeral=True
            )
            return
        
        try:
            # Get all sheet data
            data = sheet.get_all_records()
            
            # Find the user by email
            user_row = None
            row_index = None
            
            for i, row in enumerate(data):
                row_email = str(row.get("Email", "")).strip().lower()
                if row_email == email:
                    user_row = row
                    row_index = i + 2  # +2 because rows are 1-indexed and we skip header
                    break
            
            if not user_row:
                await interaction.followup.send(
                    f"❌ Email `{email}` not found!\n\n"
                    "Please make sure:\n"
                    "• You entered the correct email address\n"
                    "• There are no typos\n"
                    "• Your name is not David Zheng (he's banned)",
                    ephemeral=True
                )
                return
            
            # Check if Discord ID is already filled
            current_discord_id = str(user_row.get("Discord ID", "")).strip()
            if current_discord_id and current_discord_id != str(user.id):
                await interaction.followup.send(
                    f"⚠️ This email is already linked to a different Discord account!\n\n"
                    f"**Current Discord ID:** {current_discord_id}\n"
                    f"**Your Discord ID:** {user.id}\n\n"
                    "If this is an error, please contact an admin.",
                    ephemeral=True
                )
                return
            
            # Update the Discord ID in the sheet
            try:
                # Find the column letter for Discord ID
                headers = sheet.row_values(1)
                discord_id_col = None
                for i, header in enumerate(headers):
                    if header == "Discord ID":
                        discord_id_col = i + 1  # +1 because columns are 1-indexed
                        break
                
                if discord_id_col is None:
                    await interaction.followup.send(
                        "❌ 'Discord ID' column not found in the sheet!",
                        ephemeral=True
                    )
                    return
                
                # Convert column number to letter (A=1, B=2, etc.)
                col_letter = chr(ord('A') + discord_id_col - 1)
                cell_address = f"{col_letter}{row_index}"
                
                # Update the cell with the Discord ID
                sheet.update(cell_address, [[str(user.id)]])
                
                print(f"✅ Updated Discord ID for {email} to {user.id} in cell {cell_address}")
                
                # Trigger a sync
                guild = interaction.guild
                if guild:
                    updated_data = sheet.get_all_records()
                    sync_results = await perform_member_sync(guild, updated_data)
                    
                    # Get user info for response
                    user_name = str(user_row.get("Name", "")).strip()
                    first_event = str(user_row.get("First Event", "")).strip()
                    master_role = str(user_row.get("Master Role", "")).strip()
                    secondary_role = str(user_row.get("Secondary Role", "")).strip()
                    chapter = str(user_row.get("Chapter", "")).strip()
                    building = str(user_row.get("Building 1", "")).strip()
                    room = str(user_row.get("Room 1", "")).strip()

                    if(user_name == "David Zheng"):
                        embed = discord.Embed(
                            title="🤬 Oh god it's you again. Today better be a stress level -5 kind of day 😴",
                            description=f"Your Discord account has been linked to your email and roles have been assigned.",
                            color=discord.Color.green()
                        )
                    
                    elif(user_name == "Brian Lam"):
                        embed = discord.Embed(
                            title="❤️ Omg hi Brian I miss you. You are the LAM!!! 🐑",
                            description=f"Your Discord account has been linked to your email and roles have been assigned.",
                            color=discord.Color.green()
                        )
                    
                    elif(user_name == "Nikki Cheung"):
                        embed = discord.Embed(
                            title="🥑 Is it green? 🥑",
                            description=f"Your Discord account has been linked to your email and roles have been assigned.",
                            color=discord.Color.green()
                        )

                    elif(user_name == "Jinhuang Zhou"):
                        embed = discord.Embed(
                            title="🫵 Jinhuang Zhou. You are in trouble. Please report to the principal's office immediately.",
                            description=f"Your Discord account has been linked to your email and roles have been assigned.",
                            color=discord.Color.green()
                        )
                        
                    elif(user_name == "Satvik Kumar"):
                        embed = discord.Embed(
                            title="🌊 Hi Satvik when are we going surfing 🏄‍♂️",
                            description=f"Your Discord account has been linked to your email and roles have been assigned.",
                            color=discord.Color.green()
                        )
                        
                    else:
                        embed = discord.Embed(
                            title="✅ Successfully Logged In!",
                            description=f"Your Discord account has been linked to your email and roles have been assigned.",
                            color=discord.Color.green()
                        )
                        
                    # Build your information field
                    info_text = f"**Name:** {user_name or 'Not specified'}\n"
                    info_text += f"**Email:** {email}"
                    
                    if building and room:
                        info_text += f"\n**Location:** {building}, Room {room}"
                    elif building:
                        info_text += f"\n**Building:** {building}"
                    elif room:
                        info_text += f"\n**Room:** {room}"
                    
                    embed.add_field(
                        name="👤 Your Information",
                        value=info_text,
                        inline=False
                    )
                    
                    roles_assigned = []
                    if master_role:
                        roles_assigned.append(master_role)
                    if first_event != master_role:
                        roles_assigned.append(first_event)
                    if secondary_role and secondary_role not in roles_assigned:
                        roles_assigned.append(secondary_role)
                    
                    # Add chapter role
                    if chapter and chapter.lower() not in ["n/a", "na", ""]:
                        roles_assigned.append(chapter)
                        # Add to chapter role names set
                        chapter_role_names.add(chapter)
                    else:
                        roles_assigned.append("Unaffiliated")
                        # Unaffiliated is also a chapter role
                        chapter_role_names.add("Unaffiliated")
                    
                    if roles_assigned:
                        embed.add_field(
                            name="🎭 Roles Assigned",
                            value="\n".join([f"• {role}" for role in roles_assigned]),
                            inline=False
                        )
                    
                    embed.add_field(
                        name="🎉 What's Next?",
                        value="• You now have access to relevant channels\n"
                            "• Your nickname has been updated\n"
                            "• Check out the channels you can now see!",
                        inline=False
                    )
                    
                    embed.set_footer(text="Welcome to the team!")
                    
                    await interaction.followup.send(embed=embed, ephemeral=True)
                    
                else:
                    await interaction.followup.send(
                        "✅ Discord ID updated successfully, but could not trigger sync. Please contact an admin.",
                        ephemeral=True
                    )
                    
            except Exception as e:
                await interaction.followup.send(
                    f"❌ Error updating sheet: {str(e)}",
                    ephemeral=True
                )
                print(f"❌ Error updating sheet for {email}: {e}")
                
        except Exception as e:
            await interaction.followup.send(
                f"❌ Error accessing sheet: {str(e)}",
                ephemeral=True
            )
            print(f"❌ Error accessing sheet in login: {e}")

@bot.tree.command(name="login", description="Login by providing your email address to get your assigned roles")
async def login_command(interaction: discord.Interaction):
    """Login with email to get assigned roles"""
    
    # Show the modal
    modal = EmailLoginModal()
    await interaction.response.send_modal(modal)


@bot.tree.command(name="assignslackerzones", description="Assign zone numbers per building in 'Slacker Assignments' using K-means (Admin only)")
async def assign_slacker_zones_command(interaction: discord.Interaction):
    """Read 'Slacker Assignments' worksheet, cluster by building into K zones, write labels to 'Zone Number' column."""
    # Admin only
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)

    # Verify spreadsheet connection
    global spreadsheet
    if spreadsheet is None:
        await interaction.followup.send(
            "❌ No spreadsheet connected! Use `/entertemplate` first to connect your sheet.",
            ephemeral=True
        )
        return

    # Open the worksheet or find a separate spreadsheet in the same Drive folder
    worksheet_name = "Slacker Assignments"
    ws = None
    try:
        ws = spreadsheet.worksheet(worksheet_name)
    except Exception:
        # If not a tab in the current spreadsheet, search the parent Drive folder for a spreadsheet named like it
        try:
            from googleapiclient.discovery import build
            drive_service = build('drive', 'v3', credentials=creds)
            # Get parent folder of the currently connected spreadsheet
            sheet_metadata = drive_service.files().get(fileId=spreadsheet.id, fields='parents').execute()
            parent_folders = sheet_metadata.get('parents', [])
            if not parent_folders:
                await interaction.followup.send("❌ Could not determine parent folder to search for 'Slacker Assignments' sheet.", ephemeral=True)
                return
            parent_folder_id = parent_folders[0]
            # Search spreadsheets in same folder
            q = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and name contains '{worksheet_name}'"
            results = drive_service.files().list(q=q, fields='files(id, name)').execute()
            files = results.get('files', [])
            if not files:
                await interaction.followup.send(f"❌ Could not find a spreadsheet named '{worksheet_name}' in the same folder as the template.", ephemeral=True)
                return
            target = files[0]
            other_sheet = gc.open_by_key(target['id'])
            # Prefer a worksheet named exactly worksheet_name; otherwise first tab
            try:
                ws = other_sheet.worksheet(worksheet_name)
            except Exception:
                ws = other_sheet.worksheets()[0]
        except Exception as e2:
            await interaction.followup.send(f"❌ Could not locate '{worksheet_name}' in the same Drive folder: {str(e2)}", ephemeral=True)
            return

    # Fetch data
    try:
        headers = ws.row_values(1)
        rows = ws.get_all_records()
    except Exception as e:
        await interaction.followup.send(f"❌ Could not read worksheet data: {str(e)}", ephemeral=True)
        return

    # Normalize header names for lookups
    def _find_col_index(name_candidates):
        for i, h in enumerate(headers):
            if not h:
                continue
            for cand in name_candidates:
                if h.strip().lower() == cand:
                    return i + 1  # 1-indexed
        return None

    building_col_index = _find_col_index(["building"])
    coords_col_index = _find_col_index(["coordinates"])
    lat_col_index = _find_col_index(["latitude"]) 
    lon_col_index = _find_col_index(["longitude"]) 
    num_zones_col_index = _find_col_index(["number of zones"])
    zones_col_index = _find_col_index(["zone number"])
    num_zones = 0

    # Create zones column if missing
    if zones_col_index is None:
        try:
            new_col_idx = len(headers) + 1
            # Column letter (simple A..Z mapping consistent with rest of file usage)
            col_letter = chr(ord('A') + new_col_idx - 1)
            # Use user's preferred header name
            ws.update(f"{col_letter}1", [["zone number"]])
            headers.append("zone number")
            zones_col_index = new_col_idx
        except Exception as e:
            await interaction.followup.send(f"❌ Could not create 'Zone Number' column: {str(e)}", ephemeral=True)
            return

    # Build data per building
    from collections import defaultdict
    building_points = defaultdict(list)  # building -> list of (row_idx_1_based, (lat, lon))
    
    # Find the global K value from any row that has it
    global_k = None
    for row in rows:
        lower_row = { (k.strip().lower() if isinstance(k, str) else k): v for k, v in row.items() }
        k_raw = lower_row.get("number of zones", lower_row.get("zones count", lower_row.get("num zones", lower_row.get("k"))))
        if k_raw is not None and str(k_raw).strip() != "":
            try:
                global_k = int(float(k_raw))
                break  # Found it, use this value for all buildings
            except Exception:
                continue

    def _parse_float(val):
        try:
            if isinstance(val, str):
                val = val.strip()
                if not val:
                    return None
            return float(val)
        except Exception:
            return None

    for idx, row in enumerate(rows, start=2):  # data starts at row 2
        # Case-insensitive row access
        lower_row = { (k.strip().lower() if isinstance(k, str) else k): v for k, v in row.items() }
        building = str(lower_row.get("building", lower_row.get("building 1", ""))).strip()
        if not building:
            continue

        lat = _parse_float(lower_row.get("latitude", lower_row.get("lat")))
        lon = _parse_float(lower_row.get("longitude", lower_row.get("lon", lower_row.get("lng"))))

        if (lat is None or lon is None) and ("coordinates" in lower_row and lower_row["coordinates"]):
            coord_str = str(lower_row["coordinates"]).strip()
            if "," in coord_str:
                parts = [p.strip() for p in coord_str.split(",")]
                if len(parts) >= 2:
                    if lat is None:
                        lat = _parse_float(parts[0])
                    if lon is None:
                        lon = _parse_float(parts[1])

        if lat is None or lon is None:
            continue

        building_points[building].append((idx, (lat, lon)))

    if not building_points:
        await interaction.followup.send("⚠️ No valid location rows found to cluster.", ephemeral=True)
        return

    # Compute clusters and prepare updates
    updates = []  # list of (row_index, zone_label_str)
    k_to_use = global_k if global_k is not None and global_k > 0 else 1
    
    # Collect ALL points from ALL buildings for global clustering
    all_points = []
    all_items = []  # (building, row_idx, point)
    
    for bldg, items in building_points.items():
        for row_idx, point in items:
            all_points.append(point)
            all_items.append((bldg, row_idx, point))
    
    # Run K-means on ALL points together
    labels = _run_kmeans_clustering(all_points, k_to_use)
    
    # Debug: count cluster distribution by building
    building_zone_counts = {}  # building -> {zone: count}
    for i, (bldg, row_idx, point) in enumerate(all_items):
        zone = labels[i] + 1  # Convert to 1-based
        if bldg not in building_zone_counts:
            building_zone_counts[bldg] = {}
        building_zone_counts[bldg][zone] = building_zone_counts[bldg].get(zone, 0) + 1
    
    debug_info = []
    for bldg, zone_counts in building_zone_counts.items():
        zones = sorted(zone_counts.keys())
        counts = [zone_counts[z] for z in zones]
        debug_info.append(f"{bldg}: zones {zones} (counts {counts})")
    
    # Create updates
    for i, (bldg, row_idx, point) in enumerate(all_items):
        zone_label = str(labels[i] + 1)  # Convert to 1-based
        updates.append((row_idx, zone_label))

    # Apply updates (per cell to minimize risk of range mistakes)
    zones_col_letter = chr(ord('A') + zones_col_index - 1)
    updated = 0
    for row_idx, value in updates:
        try:
            ws.update(f"{zones_col_letter}{row_idx}", [[value]])
            updated += 1
        except Exception:
            pass

    # Summarize K used per building (limit for brevity)
    # Send debug info first
    debug_text = "\n".join(debug_info[:5])  # Limit to first 5 buildings
    if len(debug_info) > 5:
        debug_text += f"\n... and {len(debug_info) - 5} more buildings"
    
    await interaction.followup.send(
        f"✅ Assigned {k_to_use} zones for {len(updates)} rows across {len(building_points)} buildings in '{worksheet_name}'.",
        ephemeral=True
    )


@tasks.loop(minutes=1)
async def sync_members():
    """Every minute, read spreadsheet and invite any new Discord IDs."""
    print("🔄 Running member sync...")
    
    # Check if we have a sheet connected
    if sheet is None:
        print("⚠️ No sheet connected - use /entertemplate to connect to a sheet")
        return
    
    try:
        # Fetch all rows as list of dicts
        data = sheet.get_all_records()
        print(f"📊 Found {len(data)} rows in spreadsheet")
    except Exception as e:
        print("❌ Could not fetch sheet:", e)
        return

    # Sync with all guilds the bot is in
    total_processed = 0
    for guild in bot.guilds:
        try:
            print(f"🔄 Syncing members for guild: {guild.name}")
            sync_results = await perform_member_sync(guild, data)
            total_processed += sync_results['processed']
            print(f"✅ Sync complete for {guild.name}. Processed {sync_results['processed']} valid Discord IDs.")
        except Exception as e:
            print(f"❌ Error syncing guild {guild.name}: {e}")
    
    print(f"✅ Total sync complete. Processed {total_processed} valid Discord IDs across {len(bot.guilds)} guilds.")


@tasks.loop(minutes=1)
async def check_help_tickets():
    """Every minute, check for unresponded help tickets and re-ping if needed."""
    if not active_help_tickets:
        return
        
    print(f"🎫 Checking {len(active_help_tickets)} active help tickets...")
    
    current_time = datetime.now()
    tickets_to_remove = []
    
    for thread_id, ticket_info in active_help_tickets.items():
        try:
            # Calculate time since last ping
            time_since_created = current_time - ticket_info["created_at"]
            
            # Check if 5 minutes have passed since creation/last ping
            if time_since_created >= timedelta(minutes=5):
                # Find the thread across all guilds
                thread = None
                for guild in bot.guilds:
                    thread = guild.get_thread(thread_id)
                    if thread:
                        break
                
                if not thread:
                    print(f"⚠️ Thread {thread_id} not found in any guild, removing from tracking")
                    tickets_to_remove.append(thread_id)
                    continue
                
                # Check if thread is still active/not archived
                if thread.archived or thread.locked:
                    print(f"🗄️ Thread {thread_id} is archived/locked, removing from tracking")
                    tickets_to_remove.append(thread_id)
                    continue
                
                # Check ping count limit (max 3 pings)
                if ticket_info["ping_count"] >= 3:
                    print(f"⏹️ Thread {thread_id} reached max ping limit, removing from tracking")
                    tickets_to_remove.append(thread_id)
                    continue
                
                # Re-ping the slackers
                await send_ticket_repings(thread, ticket_info)
                
                # Update ping count and reset timer
                ticket_info["ping_count"] += 1
                ticket_info["created_at"] = current_time
                print(f"🔄 Re-pinged ticket {thread_id} (ping #{ticket_info['ping_count']})")
                
        except Exception as e:
            print(f"❌ Error checking ticket {thread_id}: {e}")
            tickets_to_remove.append(thread_id)
    
    # Clean up invalid tickets
    for thread_id in tickets_to_remove:
        if thread_id in active_help_tickets:
            del active_help_tickets[thread_id]
            print(f"🗑️ Removed invalid ticket {thread_id} from tracking")


async def get_all_slackers():
    """Get Discord IDs of ALL slackers from the Slacker Assignments sheet"""
    global spreadsheet
    if not spreadsheet:
        print("❌ No spreadsheet connected for all slackers lookup")
        return []
    
    try:
        # Try to get the Slacker Assignments worksheet
        try:
            sheet = spreadsheet.worksheet("Slacker Assignments")
        except Exception:
            # If not found as a worksheet, search for a separate spreadsheet
            try:
                from googleapiclient.discovery import build
                drive_service = build('drive', 'v3', credentials=creds)
                
                # Get parent folder of the currently connected spreadsheet
                sheet_metadata = drive_service.files().get(fileId=spreadsheet.id, fields='parents').execute()
                parent_folders = sheet_metadata.get('parents', [])
                if not parent_folders:
                    print("❌ Could not determine parent folder for all slackers lookup")
                    return []
                
                parent_folder_id = parent_folders[0]
                
                # Search for Slacker Assignments spreadsheet
                q = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and name contains 'Slacker Assignments'"
                results = drive_service.files().list(q=q, fields='files(id, name)').execute()
                files = results.get('files', [])
                
                if not files:
                    print("❌ Could not find Slacker Assignments spreadsheet")
                    return []
                
                # Open the first matching spreadsheet
                slacker_spreadsheet = gc.open_by_key(files[0]['id'])
                sheet = slacker_spreadsheet.sheet1  # Use first worksheet
                
            except Exception as e:
                print(f"❌ Error finding Slacker Assignments spreadsheet: {e}")
                return []
        
        # Get all data from the sheet
        data = sheet.get_all_records()
        
        # Find all slacker emails (anyone with a "Slacker Zone" value)
        slacker_emails = []
        for row in data:
            slacker_zone = row.get("Slacker Zone", "")
            if slacker_zone:  # Has a slacker zone assigned
                email = str(row.get("Email", "")).strip()
                if email:
                    slacker_emails.append(email.lower())
        
        if not slacker_emails:
            print("⚠️ No slacker emails found in Slacker Assignments")
            return []
        
        print(f"🔍 Found {len(slacker_emails)} total slacker emails")
        
        # Now cross-reference with the main sheet to get Discord IDs
        try:
            main_sheet = spreadsheet.worksheet(SHEET_PAGE_NAME)
            main_data = main_sheet.get_all_records()
        except Exception as e:
            print(f"❌ Error accessing main sheet for Discord ID lookup: {e}")
            return []
        
        all_slackers = []
        for row in main_data:
            email = str(row.get("Email", "")).strip().lower()
            if email in slacker_emails:
                discord_id = str(row.get("Discord ID", "")).strip()
                if discord_id:
                    try:
                        all_slackers.append(int(discord_id))
                    except ValueError:
                        print(f"⚠️ Invalid Discord ID '{discord_id}' for slacker email {email}")
        
        print(f"✅ Found {len(all_slackers)} total slacker Discord IDs")
        return all_slackers
        
    except Exception as e:
        print(f"❌ Error looking up all slackers: {e}")
        return []


async def send_ticket_repings(thread, ticket_info):
    """Send re-ping message for a help ticket"""
    try:
        ping_count = ticket_info["ping_count"] + 1
        
        # For final ping (3rd ping), get ALL slackers instead of just zone slackers
        if ping_count >= 3:
            print(f"🚨 Final ping for ticket {thread.id} - getting ALL slackers")
            all_slacker_ids = await get_all_slackers()
            slacker_mentions = []
            for slacker_id in all_slacker_ids:
                member = thread.guild.get_member(slacker_id)
                if member:
                    slacker_mentions.append(member.mention)
        else:
            # Regular ping - just zone slackers
            slacker_mentions = []
            for slacker_id in ticket_info["zone_slackers"]:
                member = thread.guild.get_member(slacker_id)
                if member:
                    slacker_mentions.append(member.mention)
        
        if not slacker_mentions:
            print(f"⚠️ No valid slackers found for re-ping in ticket {thread.id}")
            return
        
        mention_text = " ".join(slacker_mentions)
        
        # Build location info
        location_parts = [ticket_info["building"]]
        if ticket_info["room"]:
            location_parts.append(f"Room {ticket_info['room']}")
        location = ", ".join(location_parts)
        
        # Different field names for final ping vs regular ping
        if ping_count < 3:
            embed = discord.Embed(
                title=f"Still Need Help!",
                description=f"**Event:** {ticket_info['event']}\n**Location:** {location}",
                color=discord.Color.orange()
            )
            embed.add_field(
                name="Slackers Assigned",
                value=f"This ticket still needs assistance!",
                inline=False
            )
        else:
            embed = discord.Embed(
                title=f"Final Call!",
                description=f"**Event:** {ticket_info['event']}\n**Location:** {location}",
                color=discord.Color.red()
            )
            embed.add_field(
                name="ALL SLACKERS",
                value=f"This ticket still needs assistance!",
                inline=False
            )
            embed.add_field(
                name="\nFinal Ping",
                value="This is the final automatic ping. Please respond if you can help!",
                inline=False
            )
        
        # Send mentions as regular message content (not in embed) so Discord actually notifies users
        await thread.send(content=mention_text, embed=embed)
        print(f"📢 Sent re-ping #{ping_count} for ticket {thread.id}")
        
    except Exception as e:
        print(f"❌ Error sending re-ping for ticket {thread.id}: {e}")


@bot.tree.command(name="debugzone", description="Debug zone assignment for a user (Admin only)")
@app_commands.describe(user="The user to debug zone assignment for")
async def debug_zone_command(interaction: discord.Interaction, user: discord.Member):
    """Debug command to test zone assignment for a specific user"""
    # Admin only
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)

    try:
        # Look up the user's event and building
        user_event_info = await get_user_event_building(user.id)
        if not user_event_info:
            await interaction.followup.send(f"❌ Could not find event/building info for {user.mention}")
            return

        building = user_event_info.get("building")
        event = user_event_info.get("event")
        room = user_event_info.get("room")
        name = user_event_info.get("name")

        if not building:
            await interaction.followup.send(f"❌ No building found for {user.mention} (event: {event})")
            return

        # Get the zone for this building
        zone = await get_building_zone(building)
        if not zone:
            await interaction.followup.send(f"❌ No zone found for building '{building}'")
            return

        # Get all slackers in this zone
        zone_slackers = await get_zone_slackers(zone)

        # Create embed with debug info
        embed = discord.Embed(
            title="🐛 Zone Debug Info",
            description=f"Debug information for {user.mention}",
            color=discord.Color.blue()
        )
        
        # Build location info for debug display
        location_parts = [building]
        if room:
            location_parts.append(f"Room {room}")
        location = ", ".join(location_parts)
        
        embed.add_field(name="User Info", value=f"**Name:** {name}\n**Event:** {event}\n**Location:** {location}", inline=False)
        embed.add_field(name="Zone Assignment", value=f"**Zone:** {zone}", inline=False)
        
        if zone_slackers:
            slacker_mentions = []
            for slacker_id in zone_slackers:
                member = interaction.guild.get_member(slacker_id)
                if member:
                    slacker_mentions.append(member.mention)
                else:
                    slacker_mentions.append(f"<@{slacker_id}> (not in server)")
            
            embed.add_field(
                name=f"Zone {zone} Slackers ({len(zone_slackers)} total)",
                value="\n".join(slacker_mentions) if slacker_mentions else "No valid slackers found",
                inline=False
            )
        else:
            embed.add_field(name="Zone Slackers", value="No slackers found for this zone", inline=False)

        await interaction.followup.send(embed=embed)

    except Exception as e:
        await interaction.followup.send(f"❌ Error during debug: {str(e)}")
        print(f"❌ Debug zone error: {e}")
        import traceback
        traceback.print_exc()


@bot.tree.command(name="activetickets", description="Show all active help tickets being tracked (Admin only)")
async def active_tickets_command(interaction: discord.Interaction):
    """Debug command to show all active help tickets"""
    # Admin only
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)

    try:
        if not active_help_tickets:
            await interaction.followup.send("✅ No active help tickets being tracked.")
            return

        embed = discord.Embed(
            title="🎫 Active Help Tickets",
            description=f"Currently tracking {len(active_help_tickets)} help tickets",
            color=discord.Color.blue()
        )

        for thread_id, ticket_info in list(active_help_tickets.items())[:10]:  # Limit to 10 for display
            # Get thread info
            thread = interaction.guild.get_thread(thread_id)
            thread_name = thread.name if thread else f"Thread {thread_id} (not found)"
            
            # Calculate time since creation
            time_elapsed = datetime.now() - ticket_info["created_at"]
            minutes_elapsed = int(time_elapsed.total_seconds() / 60)
            
            # Build location info
            location_parts = [ticket_info["building"]]
            if ticket_info["room"]:
                location_parts.append(f"Room {ticket_info['room']}")
            location = ", ".join(location_parts)
            
            embed.add_field(
                name=f"🎫 {thread_name}",
                value=f"**Event:** {ticket_info['event']}\n**Location:** {location}\n**Zone:** {ticket_info['zone']}\n**Pings:** {ticket_info['ping_count']}\n**Time:** {minutes_elapsed}m ago",
                inline=True
            )

        if len(active_help_tickets) > 10:
            embed.add_field(
                name="📋 Note",
                value=f"Showing first 10 of {len(active_help_tickets)} active tickets",
                inline=False
            )

        await interaction.followup.send(embed=embed)

    except Exception as e:
        await interaction.followup.send(f"❌ Error fetching active tickets: {str(e)}")
        print(f"❌ Active tickets error: {e}")
        import traceback
        traceback.print_exc()


@bot.tree.command(name="cacheinfo", description="Show cached spreadsheet connection info (Admin only)")
async def cache_info_command(interaction: discord.Interaction):
    """Show information about the cached spreadsheet connection"""
    # Admin only
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)

    try:
        cache = load_cache()
        
        if not cache:
            await interaction.followup.send("📄 No cache file found.")
            return

        embed = discord.Embed(
            title="💾 Cache Information",
            description="Current cached spreadsheet connection details",
            color=discord.Color.blue()
        )

        # Basic connection info
        if cache.get("spreadsheet_id"):
            embed.add_field(
                name="📊 Spreadsheet",
                value=f"**Title:** {cache.get('spreadsheet_title', 'Unknown')}\n**ID:** `{cache.get('spreadsheet_id')}`",
                inline=False
            )

        if cache.get("worksheet_name"):
            embed.add_field(
                name="📋 Worksheet",
                value=cache.get("worksheet_name"),
                inline=True
            )

        if cache.get("connected_at"):
            try:
                connected_time = datetime.fromisoformat(cache.get("connected_at"))
                embed.add_field(
                    name="🕐 Connected At",
                    value=connected_time.strftime("%Y-%m-%d %H:%M:%S"),
                    inline=True
                )
            except:
                embed.add_field(name="🕐 Connected At", value="Unknown format", inline=True)

        if cache.get("folder_link"):
            embed.add_field(
                name="📁 Folder Link",
                value=f"[Open in Drive]({cache.get('folder_link')})",
                inline=False
            )

        # Cache file info
        if os.path.exists(CACHE_FILE):
            file_size = os.path.getsize(CACHE_FILE)
            embed.add_field(
                name="📄 Cache File",
                value=f"**Path:** `{CACHE_FILE}`\n**Size:** {file_size} bytes",
                inline=False
            )

        await interaction.followup.send(embed=embed)

    except Exception as e:
        await interaction.followup.send(f"❌ Error reading cache: {str(e)}")
        print(f"❌ Cache info error: {e}")


@bot.tree.command(name="clearcache", description="Clear the cached spreadsheet connection (Admin only)")
async def clear_cache_command(interaction: discord.Interaction):
    """Clear the cached spreadsheet connection"""
    # Admin only
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)

    try:
        # Check if cache exists
        cache_exists = os.path.exists(CACHE_FILE)
        
        if not cache_exists:
            await interaction.followup.send("📄 No cache file found to clear.")
            return

        # Clear the cache
        clear_cache()
        
        # Also clear the current connection
        global sheet, spreadsheet
        sheet = None
        spreadsheet = None

        embed = discord.Embed(
            title="🗑️ Cache Cleared",
            description="Cached spreadsheet connection has been cleared.\nUse `/entertemplate` to reconnect to a sheet.",
            color=discord.Color.orange()
        )

        await interaction.followup.send(embed=embed)
        print("🧹 Admin cleared the cache via command")

    except Exception as e:
        await interaction.followup.send(f"❌ Error clearing cache: {str(e)}")
        print(f"❌ Clear cache error: {e}")


@bot.tree.command(name="msg", description="Send a message as the bot (:( role only)")
@app_commands.describe(
    message="The message to send",
    channel="Channel to send to (optional, defaults to current channel)"
)
async def msg_command(interaction: discord.Interaction, message: str, channel: discord.TextChannel = None):
    """Send a message as the bot - restricted to :( role only"""
    
    # Check if user has the :( role
    sad_face_role = discord.utils.get(interaction.user.roles, name=":(")
    if not sad_face_role:
        await interaction.response.send_message("❌ You need the `:( ` role to use this command!", ephemeral=True)
        return
    
    # Use current channel if no channel specified
    target_channel = channel or interaction.channel
    
    # Check if bot has permission to send messages in the target channel
    if not target_channel.permissions_for(interaction.guild.me).send_messages:
        await interaction.response.send_message(f"❌ I don't have permission to send messages in {target_channel.mention}!", ephemeral=True)
        return
    
    try:
        # Send the message to the target channel
        await target_channel.send(message)
        
        # Confirm to the user (privately)
        if target_channel == interaction.channel:
            await interaction.response.send_message("✅ Message sent!", ephemeral=True)
        else:
            await interaction.response.send_message(f"✅ Message sent to {target_channel.mention}!", ephemeral=True)
        
        # Log the action
        print(f"📢 {interaction.user} used /msg in {interaction.guild.name}: '{message}' → #{target_channel.name}")
        
    except discord.Forbidden:
        await interaction.response.send_message(f"❌ I don't have permission to send messages in {target_channel.mention}!", ephemeral=True)
    except Exception as e:
        await interaction.response.send_message(f"❌ Error sending message: {str(e)}", ephemeral=True)
        print(f"❌ Error in /msg command: {e}")

if __name__ == "__main__":
    bot.run(TOKEN)