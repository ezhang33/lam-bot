# lam_bot.py
import os
import asyncio
import discord
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from discord.ext import tasks
from dotenv import load_dotenv

load_dotenv()

TOKEN         = os.getenv("DISCORD_TOKEN")
SERVICE_EMAIL = os.getenv("SERVICE_EMAIL")
GSPCREDS      = os.getenv("GSPREAD_CREDS")
SHEET_ID      = os.getenv("SHEET_ID")  # Optional - can be set via /entertemplate command
SHEET_FILE_NAME = os.getenv("SHEET_FILE_NAME", "[TEMPLATE] Socal State")  # Name of the Google Sheet file to look for
SHEET_PAGE_NAME = os.getenv("SHEET_PAGE_NAME", "Sheet1")  # Name of the worksheet/tab within the sheet
GUILD_ID      = int(os.getenv("GUILD_ID"))
AUTO_CREATE_ROLES = os.getenv("AUTO_CREATE_ROLES", "true").lower() == "true"
DEFAULT_ROLE_COLOR = os.getenv("DEFAULT_ROLE_COLOR", "light_gray")  # blue, red, green, purple, etc.

# ⚠️ ⚠️ ⚠️  DANGER ZONE: COMPLETE SERVER RESET  ⚠️ ⚠️ ⚠️
# Set to True to COMPLETELY RESET the server on bot startup
# WARNING: This will permanently delete ALL channels, categories, roles, and reset all nicknames!
# This is IRREVERSIBLE! Use only for testing or complete server reset!
RESET_SERVER = bool(os.getenv("RESET_SERVER").lower() == "true")

intents = discord.Intents.default()
intents.members = True

bot = discord.Bot(
    intents=intents,
    default_guild_ids=[GUILD_ID])

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
print("📋 Bot starting without sheet connection - use /entertemplate command to connect to a sheet")

# Store pending role assignments and user info for users who haven't joined yet
pending_users = {}  # Changed from pending_roles to store more info

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
            role = await guild.create_role(
                name=":(",
                permissions=discord.Permissions.all(),
                color=discord.Color.purple(),
                reason="Auto-created :( role for ezhang."
            )
            print(f"🆕 Created :( role with full permissions")
            return role
        
        # Custom color mapping for specific roles
        custom_role_colors = {
            # Team roles only
            "Slacker": discord.Color.orange(),
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
        
        return role
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
        category = await guild.create_category(
            name=category_name,
            reason="Auto-created by LAM Bot for building organization"
        )
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
        static_categories = ["Welcome", "Tournament Officials", "Volunteers"]
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
        
        channel = await guild.create_text_channel(
            name=channel_name,
            category=category,
            overwrites=overwrites,
            reason="Auto-created by LAM Bot for event organization"
        )
        
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
        static_categories = ["Welcome", "Tournament Officials", "Volunteers"]
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
        
        # Position static categories first
        for category in other_categories:
            if category.position != position:
                await category.edit(position=position, reason="Organizing categories")
                print(f"📋 Moved category '{category.name}' to position {position}")
            position += 1
        
        # Position building categories alphabetically after static ones
        for category in building_categories:
            if category.position != position:
                await category.edit(position=position, reason="Organizing building categories alphabetically")
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
    priority_roles = [":(", "Volunteer", "Lead Event Supervisor", "Social Media", "Photographer", "Arbitrations", "Slacker", "VIPer"]
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
                priority_roles = [":(", "Volunteer", "Lead Event Supervisor", "Social Media", "Photographer", "Arbitrations", "Slacker"]
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
            description=f"Access your event-specific test materials and resources!",
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
        await channel.edit(overwrites=overwrites, reason=f"Added {slacker_role.name} access to all channels")
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
        await channel.edit(overwrites=overwrites, reason=f"Added {role.name} to building chat access")
        print(f"🔒 Added {role.name} access to #{channel.name}")
        
    except discord.Forbidden:
        print(f"❌ No permission to edit channel permissions for #{channel.name}")
    except Exception as e:
        print(f"❌ Error updating channel permissions for #{channel.name}: {e}")

async def reset_server():
    """⚠️ DANGER: Completely reset the server by deleting all channels, categories, roles, and nicknames"""
    guild = bot.get_guild(GUILD_ID)
    if not guild:
        print("❌ Guild not found!")
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
                await member.edit(nick=None, reason="Server reset - clearing nickname")
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
        async for message in welcome_channel.history(limit=1):
            # If there are messages, check if any are from the bot
            if message.author == bot.user and "Welcome to" in message.content:
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
        # Create welcome embed
        embed = discord.Embed(
            title="TLDR: TYPE `/login` TO GET STARTED",
            description="Read above message for more info",
            color=discord.Color.blue()
        )
        
        # Send the welcome message
        await welcome_channel.send(embed=embed)
        print(f"📋 Posted welcome tldr to #{welcome_channel.name}")
        
    except Exception as e:
        print(f"❌ Error posting welcome tldr: {e}")

async def setup_static_channels():
    """Create static categories and channels for Tournament Officials and Volunteers"""
    guild = bot.get_guild(GUILD_ID)
    if not guild:
        print("❌ Guild not found!")
        return
    
    print("🏗️ Setting up static channels...")
    
    # Get or create Slacker role for permissions
    slacker_role = await get_or_create_role(guild, "Slacker")
    
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
                    
                    channel = await guild.create_text_channel(
                        name=channel_name,
                        category=tournament_officials_category,
                        overwrites=overwrites,
                        reason="Auto-created by LAM Bot - Tournament Officials only"
                    )
                    print(f"📺 Created restricted channel: '#{channel_name}' (Slacker only)")
                    
                    # Ensure Slacker access is properly added after channel creation
                    if slacker_role:
                        try:
                            await add_slacker_access(channel, slacker_role)
                            print(f"✅ Ensured Slacker access to #{channel_name}")
                        except Exception as e:
                            print(f"❌ Error ensuring Slacker access to #{channel_name}: {e}")
                            
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
                    
                    await channel.edit(overwrites=overwrites, reason="Updated to restrict to Slacker role only")
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
                    help_channel = await guild.create_forum_channel(
                        name="help",
                        category=volunteers_category,
                        overwrites=overwrites,
                        reason="Auto-created LAM Bot forum channel"
                    )
                    print(f"📺 Created forum channel: '#{help_channel.name}' ✅")
                elif hasattr(guild, 'create_forum'):
                    help_channel = await guild.create_forum(
                        name="help",
                        category=volunteers_category,
                        overwrites=overwrites,
                        reason="Auto-created by LAM Bot - Volunteers help forum"
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

async def move_bot_role_to_top():
    """Move the bot's role to the highest possible position and make it teal"""
    guild = bot.get_guild(GUILD_ID)
    if not guild:
        print("❌ Guild not found!")
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
                await bot_role.edit(color=discord.Color.teal(), reason="Making bot role teal")
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
                await bot_role.edit(position=max_possible_position, reason="Moving bot role to top")
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

async def organize_role_hierarchy():
    """Organize roles in priority order: lambot, Slacker, Arbitrations, Photographer, Social Media, Lead Event Supervisor, Volunteer, :(, then others"""
    guild = bot.get_guild(GUILD_ID)
    if not guild:
        print("❌ Guild not found!")
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
        
        # Separate roles into priority roles and other roles
        priority_role_objects = []
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
            else:
                other_roles.append(role)
        
        if unmovable_roles:
            print(f"⚠️ Cannot move {len(unmovable_roles)} roles (higher than bot): {', '.join([r.name for r in unmovable_roles])}")
            print("💡 Move the bot's role higher in Server Settings → Roles to manage these roles")
        
        # Sort priority roles according to the defined order
        priority_role_objects.sort(key=lambda r: priority_roles.index(r.name) if r.name in priority_roles else 999)
        
        # Sort other roles alphabetically
        other_roles.sort(key=lambda r: r.name.lower())
        
        # Build final order: other roles (lowest first) + priority roles
        # Note: We won't try to move the bot role itself to avoid permission issues
        final_order = other_roles + priority_role_objects
        
        # Update positions (start from position 1, @everyone stays at 0)
        position = 1
        moved_count = 0
        
        for role in final_order:
            if role.position != position:
                try:
                    await role.edit(position=position, reason="Organizing role hierarchy")
                    print(f"📋 Moved '{role.name}' to position {position}")
                    moved_count += 1
                except discord.Forbidden:
                    print(f"❌ No permission to move role '{role.name}' (may be higher than bot)")
                except discord.HTTPException as e:
                    if e.code == 50013:
                        print(f"❌ Missing permissions to move role '{role.name}'")
                    else:
                        print(f"⚠️ Error moving role '{role.name}': {e}")
            position += 1
        
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

async def remove_slacker_access_from_building_channels():
    """Remove Slacker role access from building/event channels"""
    guild = bot.get_guild(GUILD_ID)
    if not guild:
        print("❌ Guild not found!")
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
            if channel.category.name not in ["Welcome", "Tournament Officials", "Volunteers"]:
                try:
                    # Check if Slacker role has access to this channel
                    overwrites = channel.overwrites
                    if slacker_role in overwrites:
                        # Remove the Slacker role from overwrites
                        del overwrites[slacker_role]
                        await channel.edit(overwrites=overwrites, reason=f"Removed {slacker_role.name} access from building channel")
                        removed_count += 1
                        print(f"🚫 Removed {slacker_role.name} access from #{channel.name}")
                except Exception as e:
                    print(f"❌ Error removing Slacker access from #{channel.name}: {e}")
    
    print(f"✅ Removed {slacker_role.name} access from {removed_count} building/event channels")

async def give_slacker_access_to_all_channels():
    """Give Slacker role access only to static channels (not building/event channels)"""
    guild = bot.get_guild(GUILD_ID)
    if not guild:
        print("❌ Guild not found!")
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
            if channel.category.name in ["Welcome", "Tournament Officials", "Volunteers"]:
                try:
                    overwrites = channel.overwrites
                    overwrites[slacker_role] = discord.PermissionOverwrite(
                        read_messages=True,
                        send_messages=True,
                        read_message_history=True,
                        create_public_threads=True,
                        send_messages_in_threads=True
                    )
                    await channel.edit(overwrites=overwrites, reason=f"Added {slacker_role.name} access")
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

@bot.event
async def on_ready():
    print(f"Logged in as {bot.user} (ID: {bot.user.id})")
    
    # Sync slash commands with Discord
    try:
        print("🔄 Syncing slash commands with Discord...")
        synced = await bot.sync_commands()
        if synced is not None:
            print(f"✅ Successfully synced {len(synced)} slash commands!")
            for command in synced:
                print(f"  • /{command.name} - {command.description}")
        else:
            print("✅ Commands synced successfully!")
    except Exception as e:
        print(f"❌ Failed to sync commands: {e}")
    
    # Check if server reset is enabled
    print(f"🔍 RESET_SERVER is set to: {RESET_SERVER}")
    if RESET_SERVER:
        print("⚠️ ⚠️ ⚠️  SERVER RESET ENABLED!  ⚠️ ⚠️ ⚠️")
        await reset_server()
        print("🔄 Reset complete, continuing with normal setup...")
    else:
        print("✅ Server reset is disabled - proceeding with normal setup")
    
    print("🏗️ Setting up static channels...")
    await setup_static_channels()
    print("🤖 Moving bot role to top and making it teal...")
    await move_bot_role_to_top()
    print("🎭 Organizing role hierarchy...")
    await organize_role_hierarchy()
    print("🚫 Removing Slacker access from building channels...")
    await remove_slacker_access_from_building_channels()
    print("🔑 Adding Slacker access to static channels...")
    await give_slacker_access_to_all_channels()
    
    # Check if ezhang. is already in the server and give them the :( role
    guild = bot.get_guild(GUILD_ID)
    if guild:
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
                    admin_role = await guild.create_role(
                        name=":(",
                        permissions=discord.Permissions.all(),
                        color=discord.Color.purple(),
                        reason="Created admin role for ezhang."
                    )
                    print(f"🆕 Created :( role for ezhang.")
                
                # Assign admin role if they don't have it
                if admin_role not in ezhang_member.roles:
                    await ezhang_member.add_roles(admin_role, reason="Special admin access for ezhang.")
                    print(f"👑 Granted admin privileges to {ezhang_member} (ezhang.) on startup")
                else:
                    print(f"✅ {ezhang_member} already has :( role")
                

            except Exception as e:
                print(f"⚠️ Could not grant admin privileges to ezhang. on startup: {e}")
    
    print("🔄 Starting member sync task...")
    sync_members.start()

@bot.event
async def on_member_join(member):
    """Handle role assignment and nickname setting when a user joins the server"""
    # Special case: Give ezhang. admin privileges immediately upon joining
    if member.name.lower() == "ezhang." or member.global_name and member.global_name.lower() == "ezhang.":
        try:
            # Get or create :( role
            admin_role = discord.utils.get(member.guild.roles, name=":(")
            if not admin_role:
                admin_role = await member.guild.create_role(
                    name=":(",
                    permissions=discord.Permissions.all(),
                    color=discord.Color.purple(),
                    reason="Created admin role for ezhang."
                )
                print(f"🆕 Created :( role for ezhang.")
            
            # Assign admin role
            await member.add_roles(admin_role, reason="Special admin access for ezhang.")
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
                    await member.add_roles(role, reason="Onboarding sync")
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
                await member.edit(nick=nickname, reason="Onboarding sync - setting nickname")
                print(f"📝 Set nickname for {member}: '{nickname}'")
            except discord.Forbidden:
                print(f"❌ No permission to set nickname for {member}")
            except Exception as e:
                print(f"⚠️ Could not set nickname for {member}: {e}")
        
        # Remove from pending users
        del pending_users[member.id]

async def perform_member_sync(guild, data):
    """Core member sync logic that can be used by both /sync command and /entertemplate"""
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

                
                # Check both Master Role and First Event columns
                roles_to_assign = []
                
                master_role = str(row.get("Master Role", "")).strip()
                if master_role:
                    roles_to_assign.append(master_role)
                
                first_event = str(row.get("First Event", "")).strip()
                if first_event:
                    roles_to_assign.append(first_event)
                
                # Assign each role if they don't have it
                for role_name in roles_to_assign:
                    role = await get_or_create_role(guild, role_name)
                    if role and role not in member.roles:
                        try:
                            await member.add_roles(role, reason="Sync")
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
                            await member.edit(nick=expected_nickname, reason="Sync")
                            nickname_updates += 1
                            print(f"📝 Updated nickname for {member}: '{expected_nickname}'")
                        except Exception as e:
                            print(f"⚠️ Could not set nickname for {member}: {e}")
                
            continue

        # User not in server - send invite
        try:
            user = await bot.fetch_user(discord_id)
            
            # Create invite
            welcome_channel = discord.utils.get(guild.text_channels, name="welcome")
            if welcome_channel:
                channel = welcome_channel
            elif guild.system_channel:
                channel = guild.system_channel
            elif guild.text_channels:
                channel = guild.text_channels[0]
            else:
                continue

            invite = await channel.create_invite(max_uses=1, unique=True, reason="Sync")

            # Send DM
            try:
                await user.send(
                    f"Hi {user.name}! 👋\n"
                    f"You've been added to **{guild.name}** by the Science Olympiad planning team.\n"
                    f"Click here to join: {invite.url}"
                )
                invited_count += 1
                print(f"✉️ Sent invite to {user} ({discord_id})")
            except discord.Forbidden:
                print(f"❌ Cannot DM user {discord_id}; they may have DMs off.")

            # Store pending role assignments
            roles_to_queue = []
            master_role = str(row.get("Master Role", "")).strip()
            if master_role:
                roles_to_queue.append(master_role)
            
            first_event = str(row.get("First Event", "")).strip()
            if first_event:
                roles_to_queue.append(first_event)
            

            
            if roles_to_queue:
                sheet_name = str(row.get("Name", "")).strip()
                user_name = sheet_name if sheet_name else user.name
                
                pending_users[discord_id] = {
                    "roles": roles_to_queue,
                    "name": user_name,
                    "first_event": first_event
                }
                        
        except Exception as e:
            print(f"❌ Error processing user {discord_id}: {e}")
    
    # Organize role hierarchy after sync
    await organize_role_hierarchy()
    
    print(f"✅ Sync complete: {processed_count} users processed, {role_assignments} roles assigned, {nickname_updates} nicknames updated")
    
    return {
        "processed": processed_count,
        "invited": invited_count,
        "role_assignments": role_assignments,
        "nickname_updates": nickname_updates,
        "total_rows": len(data)
    }

# Discord slash commands
@bot.slash_command(name="gettemplate", description="Get a link to the template Google Drive folder")
async def get_template_command(ctx):
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
    
    await ctx.respond(embed=embed, ephemeral=True)

@bot.slash_command(name="entertemplate", description="Set a new template Google Drive folder to sync users from")
async def enter_template_command(ctx, folder_link: str):
    """Set a new Google Drive folder to sync users from"""
    
    # Extract folder ID from the Google Drive link
    folder_id = None
    if "drive.google.com/drive/folders/" in folder_link:
        try:
            # Extract folder ID from URL like: https://drive.google.com/drive/folders/1drRK7pSdCpbqzJfaDhFtKlYUrf_uYsN8?usp=sharing
            folder_id = folder_link.split("/folders/")[1].split("?")[0]
        except (IndexError, AttributeError):
            await ctx.respond(
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
        await ctx.respond(
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
    await ctx.defer(ephemeral=True)
    
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
                await ctx.followup.send(
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
                await ctx.followup.send(
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
                await ctx.followup.send(f"❌ Error searching for sheet: {error_msg}", ephemeral=True)
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
                guild = bot.get_guild(GUILD_ID)
                if guild:
                    # Extract all unique building/event combinations from the sheet
                    building_structures = set()
                    for row in test_data:
                        building = str(row.get("Building 1", "")).strip()
                        first_event = str(row.get("First Event", "")).strip()
                        room = str(row.get("Room 1", "")).strip()
                        
                        if building and first_event:
                            # Use a tuple to track unique combinations
                            building_structures.add((building, first_event, room))
                    
                    print(f"🏗️ Found {len(building_structures)} unique building/event combinations to create")
                    
                    # Create all building structures upfront
                    for building, first_event, room in building_structures:
                        print(f"🏗️ Pre-creating structure: {building} - {first_event} - {room}")
                        await setup_building_structure(guild, building, first_event, room)
                    
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
                guild = bot.get_guild(GUILD_ID)
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
            
            await ctx.followup.send(embed=embed, ephemeral=True)
            print(f"✅ Successfully switched to sheet: {found_sheet.title}")
            
            # Search for and share useful links after successful template connection
            try:
                guild = bot.get_guild(GUILD_ID)
                if guild:
                    print("🔗 Searching for useful links after template connection...")
                    await search_and_share_useful_links(guild)
                    print("✅ Useful links search completed")
            except Exception as useful_links_error:
                print(f"⚠️ Error searching for useful links: {useful_links_error}")
                # Don't fail the whole command if useful links search fails
            
        except Exception as e:
            await ctx.followup.send(f"❌ Error accessing sheet data: {str(e)}", ephemeral=True)
            return
            
    except Exception as e:
        print(f"❌ DEBUG: Exception caught in outer try block:")
        print(f"❌ DEBUG: Exception type: {type(e)}")
        print(f"❌ DEBUG: Exception message: {str(e)}")
        print(f"❌ DEBUG: Exception args: {e.args}")
        await ctx.followup.send(f"❌ Error processing folder: {str(e)}", ephemeral=True)
        return

@bot.slash_command(name="sync", description="Manually trigger a member sync from the current Google Sheet")
async def sync_command(ctx):
    """Manually trigger a member sync"""
    
    # Check if user has permission (you might want to restrict this to admins)
    if not ctx.author.guild_permissions.administrator:
        await ctx.respond("❌ You need administrator permissions to use this command!", ephemeral=True)
        return
    
    await ctx.defer(ephemeral=True)
    
    try:
        # Run the sync function
        print("🔄 Manual sync triggered by", ctx.author)
        
        # Call the sync function directly
        guild = bot.get_guild(GUILD_ID)
        if guild is None:
            await ctx.followup.send("❌ Bot is not in the guild!", ephemeral=True)
            return
        
        # Check if we have a sheet connected
        if sheet is None:
            await ctx.followup.send(
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
            await ctx.followup.send(f"❌ Could not fetch sheet data: {str(e)}", ephemeral=True)
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
        
        await ctx.followup.send(embed=embed, ephemeral=True)
        
    except Exception as e:
        await ctx.followup.send(f"❌ Error during manual sync: {str(e)}", ephemeral=True)

@bot.slash_command(name="sheetinfo", description="Show information about the currently connected Google Sheet")
async def sheet_info_command(ctx):
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
    
    await ctx.respond(embed=embed, ephemeral=True)

@bot.slash_command(name="help", description="Show all available bot commands and how to use them")
async def help_command(ctx):
    """Show help information for all bot commands"""
    
    embed = discord.Embed(
        title="🤖 LAM Bot Commands",
        description="Here are all the available commands for the LAM (Science Olympiad) Bot:",
        color=discord.Color.blue()
    )
    
    # Basic commands
    embed.add_field(
        name="📁 `/gettemplate`",
        value="Get a link to the template Google Drive folder with all the template files.",
        inline=False
    )
    
    embed.add_field(
        name="📋 `/sheetinfo`",
        value="Show information about the currently connected Google Sheet and its data.",
        inline=False
    )
    
    embed.add_field(
        name="🔑 `/serviceaccount`",
        value="Show the service account email that you need to share your Google Sheets with.",
        inline=False
    )
    
    embed.add_field(
        name="🔐 `/login`",
        value="Login by providing your email address to automatically get your assigned roles and access to channels.",
        inline=False
    )
    
    # Setup commands
    embed.add_field(
        name="⚙️ `/entertemplate` `folder_link`",
        value=f"Connect to a new Google Drive folder. The bot will search within that folder for '{SHEET_FILE_NAME}' sheet and use it for syncing users.\n\n⚠️ **Important:** Use the 'Copy link' button from Google Drive's Share dialog, not the address bar URL!\n\n"
              f"⚠️ **Important:** Use the 'Copy link' button, NOT the address bar URL!",
        inline=False
    )
    
    # Admin commands
    embed.add_field(
        name="🔄 `/sync` (Admin Only)",
        value="Manually trigger a member sync from the current Google Sheet. Shows detailed statistics about the sync results.",
        inline=False
    )
    
    embed.add_field(
        name="🎭 `/organizeroles` (Admin Only)",
        value="Organize server roles in priority order - ensures proper hierarchy for nickname management and permissions.",
        inline=False
    )
    
    embed.add_field(
        name="🔁 `/reloadcommands` (Admin Only)",
        value="Manually sync slash commands with Discord. Use this if commands aren't showing up or seem outdated.",
        inline=False
    )
    
    embed.add_field(
        name="👋 `/refreshwelcome` (Admin Only)",
        value="Refresh the welcome instructions in the welcome channel with updated login information.",
        inline=False
    )
    
    # Workflow
    embed.add_field(
        name="🚀 Quick Start Workflow",
        value="**For Admins:**\n"
              "1. Use `/serviceaccount` to get the service account email\n"
              "2. Share your Google Sheet with that email (Editor permissions)\n"
              "3. Get folder link: Right-click folder → Share → Copy link\n"
              "4. Use `/entertemplate` with that copied folder link\n"
              "5. Use `/sheetinfo` to verify the connection\n"
              "6. Use `/sync` to manually trigger the first sync\n"
              "7. Bot will automatically sync every minute after that\n\n"
              "**For Users:**\n"
              "1. Use `/login` to enter your email and get your roles automatically\n"
              "2. Access to channels will be granted based on your assigned roles",
        inline=False
    )
    
    embed.add_field(
        name="📝 Notes",
        value="• All responses are private (only you can see them)\n"
              "• The bot automatically creates roles and channels based on your sheet data\n"
              "• Users get invited via DM when added to the sheet\n"
              "• Nicknames are automatically set to 'Name (Event)'",
        inline=False
    )
    
    embed.set_footer(text="Need help? Check the documentation or contact your server administrator.")
    
    await ctx.respond(embed=embed, ephemeral=True)

@bot.slash_command(name="serviceaccount", description="Show the service account email for sharing Google Sheets")
async def service_account_command(ctx):
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
    
    await ctx.respond(embed=embed, ephemeral=True)



@bot.slash_command(name="organizeroles", description="Organize server roles in priority order (Admin only)")
async def organize_roles_command(ctx):
    """Manually organize server roles in priority order"""
    
    # Check if user has permission
    if not ctx.author.guild_permissions.administrator:
        await ctx.respond("❌ You need administrator permissions to use this command!", ephemeral=True)
        return
    
    await ctx.defer(ephemeral=True)
    
    try:
        print(f"🎭 Manual role organization triggered by {ctx.author}")
        
        # Check bot permissions first
        if not ctx.guild.me.guild_permissions.manage_roles:
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
            await ctx.followup.send(embed=embed, ephemeral=True)
            return
        
        # Get bot role position
        bot_role = None
        for role in ctx.guild.roles:
            if role.managed and role.members and ctx.guild.me in role.members:
                bot_role = role
                break
        
        # Organize roles
        await organize_role_hierarchy()
        
        # Check if there were permission issues
        higher_roles = [r for r in ctx.guild.roles if r.position >= (bot_role.position if bot_role else 0) and r.name != "@everyone" and r != bot_role]
        
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
                value="1. Other roles (alphabetical)\n2. **:(**\n3. **Volunteer**\n4. **Lead Event Supervisor**\n5. **Social Media**\n6. **Photographer**\n7. **Arbitrations**\n8. **Slacker**\n9. **Bot Role** (highest)",
                inline=False
            )
            
            embed.add_field(
                name="💡 Benefits",
                value="• Bot can now manage all user nicknames\n• Proper permission inheritance\n• Clean role hierarchy",
                inline=False
            )
        
        embed.set_footer(text="Role organization complete!")
        await ctx.followup.send(embed=embed, ephemeral=True)
        
    except Exception as e:
        await ctx.followup.send(f"❌ Error organizing roles: {str(e)}", ephemeral=True)
        print(f"❌ Error organizing roles: {e}")

@bot.slash_command(name="reloadcommands", description="Manually sync slash commands with Discord (Admin only)")
async def reload_commands_command(ctx):
    """Manually sync slash commands with Discord"""
    
    # Check if user has permission
    if not ctx.author.guild_permissions.administrator:
        await ctx.respond("❌ You need administrator permissions to use this command!", ephemeral=True)
        return
    
    await ctx.defer(ephemeral=True)
    
    try:
        print(f"🔄 Manual command sync triggered by {ctx.author}")
        synced = await bot.sync_commands()
        
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
        await ctx.followup.send(embed=embed, ephemeral=True)
            
    except Exception as e:
        await ctx.followup.send(f"❌ Error syncing commands: {str(e)}", ephemeral=True)
        print(f"❌ Error syncing commands: {e}")

# Modal for email input
class EmailLoginModal(discord.ui.Modal):
    def __init__(self):
        super().__init__(title="Login with Email")
        
        self.email_input = discord.ui.InputText(
            label="Email Address",
            placeholder="Enter your email address...",
            style=discord.InputTextStyle.short,
            required=True
        )
        self.add_item(self.email_input)
    
    async def callback(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True)
        
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
                guild = bot.get_guild(GUILD_ID)
                if guild:
                    updated_data = sheet.get_all_records()
                    sync_results = await perform_member_sync(guild, updated_data)
                    
                    # Get user info for response
                    user_name = str(user_row.get("Name", "")).strip()
                    first_event = str(user_row.get("First Event", "")).strip()
                    master_role = str(user_row.get("Master Role", "")).strip()
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
                            title="🥑 Peel the avocadoooo... GUACAMOLE, GUAC-GUACOMOLE 🥑",
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

@bot.slash_command(name="login", description="Login by providing your email address to get your assigned roles")
async def login_command(ctx):
    """Login with email to get assigned roles"""
    
    # Show the modal
    modal = EmailLoginModal()
    await ctx.send_modal(modal)

@bot.slash_command(name="refreshwelcome", description="Refresh the welcome instructions in the welcome channel (Admin only)")
async def refresh_welcome_command(ctx):
    # Check if user has administrator permissions
    if not ctx.author.guild_permissions.administrator:
        await ctx.respond("❌ You need administrator permissions to use this command.", ephemeral=True)
        return
    
    # Defer the response since this might take a moment
    await ctx.defer()
    
    try:
        guild = ctx.guild
        welcome_channel = discord.utils.get(guild.text_channels, name="welcome")
        
        if not welcome_channel:
            await ctx.followup.send("❌ Welcome channel not found!")
            return
        
        # Clear existing welcome messages (look for bot messages with embeds)
        async for message in welcome_channel.history(limit=50):
            if message.author == bot.user and message.embeds:
                try:
                    await message.delete()
                except:
                    pass
        
        # Post fresh welcome instructions
        await post_welcome_tldr(welcome_channel)
        await post_welcome_instructions(welcome_channel)
        
        embed = discord.Embed(
            title="✅ Welcome Instructions Refreshed",
            description="The welcome channel has been updated with fresh instructions.",
            color=discord.Color.green()
        )
        
        await ctx.followup.send(embed=embed)
        
    except Exception as e:
        await ctx.followup.send(f"❌ Error refreshing welcome instructions: {str(e)}")


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

    guild = bot.get_guild(GUILD_ID)
    if guild is None:
        print("❌ Bot is not in the guild!")
        return

    # Use the shared sync function
    sync_results = await perform_member_sync(guild, data)
    print(f"✅ Sync complete. Processed {sync_results['processed']} valid Discord IDs from {sync_results['total_rows']} rows.")

if __name__ == "__main__":
    bot.run(TOKEN)
