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
GSPCREDS      = os.getenv("GSPREAD_CREDS")
SHEET_ID      = os.getenv("SHEET_ID")
SHEET_NAME    = os.getenv("SHEET_NAME", "Sheet1")  # Default to "Sheet1" if not specified
GUILD_ID      = int(os.getenv("GUILD_ID"))
AUTO_CREATE_ROLES = os.getenv("AUTO_CREATE_ROLES", "true").lower() == "true"
DEFAULT_ROLE_COLOR = os.getenv("DEFAULT_ROLE_COLOR", "light_gray")  # blue, red, green, purple, etc.

# ⚠️ ⚠️ ⚠️  DANGER ZONE: COMPLETE SERVER RESET  ⚠️ ⚠️ ⚠️
# Set to True to COMPLETELY RESET the server on bot startup
# WARNING: This will permanently delete ALL channels, categories, roles, and reset all nicknames!
# This is IRREVERSIBLE! Use only for testing or complete server reset!
RESET_SERVER = os.getenv("RESET_SERVER", "false").lower() == "true"

intents = discord.Intents.default()
intents.members = True

bot = discord.Bot(intents=intents)

# Set up gspread client
scope = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
creds = ServiceAccountCredentials.from_json_keyfile_name(GSPCREDS, scope)
gc = gspread.authorize(creds)

# Open the specific sheet by name
try:
    spreadsheet = gc.open_by_key(SHEET_ID)
    sheet = spreadsheet.worksheet(SHEET_NAME)
    print(f"✅ Connected to sheet: '{SHEET_NAME}'")
except gspread.WorksheetNotFound:
    print(f"❌ Sheet '{SHEET_NAME}' not found!")
    print("Available sheets:")
    for worksheet in spreadsheet.worksheets():
        print(f"  - {worksheet.title}")
    exit(1)
except Exception as e:
    print(f"❌ Error connecting to Google Sheets: {e}")
    exit(1)

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
        # Custom color mapping for specific roles
        custom_role_colors = {
            # Team roles only
            "Slacker": discord.Color.orange(),
            "Volunteer": discord.Color.blue(),
            "Lead Event Supervisor": discord.Color.yellow(),
            "Photographer": discord.Color.red(),
            "Arbitrations": discord.Color.green(),
            "Social Media": discord.Color.magenta(),
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
        return category
    
    try:
        category = await guild.create_category(
            name=category_name,
            reason="Auto-created by LAM Bot for building organization"
        )
        print(f"🏢 Created category: '{category_name}'")
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
        # Add Slacker role to existing channel if it exists
        slacker_role = discord.utils.get(guild.roles, name="Slacker")
        if slacker_role and channel:
            await add_slacker_access(channel, slacker_role)
        return channel
    
    try:
        # Set up permissions
        overwrites = {}
        
        # Always give Slacker role access to all channels
        slacker_role = discord.utils.get(guild.roles, name="Slacker")
        if slacker_role:
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
            print(f"📺 Created channel: '#{channel_name}' (restricted to {event_role.name})")
        elif is_building_chat:
            print(f"📺 Created building chat: '#{channel_name}' (restricted)")
        else:
            print(f"📺 Created channel: '#{channel_name}'")
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
        
        # Separate building categories from static categories
        static_categories = ["Welcome", "Tournament Officials", "Volunteers"]
        building_categories = []
        other_categories = []
        
        for category in all_categories:
            if category.name in static_categories:
                other_categories.append(category)
            else:
                building_categories.append(category)
        
        # Sort building categories alphabetically
        building_categories.sort(key=lambda cat: cat.name.lower())
        
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

async def setup_building_structure(guild, building, first_event, room=None):
    """Set up category and channels for a building and event"""
    # Create or get the building category
    category_name = building
    category = await get_or_create_category(guild, category_name)
    if not category:
        return
    
    # Get Slacker role to ensure access
    slacker_role = discord.utils.get(guild.roles, name="Slacker")
    
    # Create general building chat channel (restricted to people with events in this building)
    building_chat_name = f"{building.lower().replace(' ', '-')}-chat"
    building_chat = await get_or_create_channel(guild, building_chat_name, category, is_building_chat=True)
    
    # Ensure Slacker has access to building chat
    if building_chat and slacker_role:
        await add_slacker_access(building_chat, slacker_role)
    
    # Create event-specific channel if we have the info
    if first_event:
        # Get or create the event role
        event_role = await get_or_create_role(guild, first_event)
        if event_role:
            # Add the event role to the building chat permissions
            await add_role_to_building_chat(building_chat, event_role)
            
            # Create channel name: [First Event] - [Building] [Room]
            if room:
                channel_name = f"{first_event.lower().replace(' ', '-')}-{building.lower().replace(' ', '-')}-{room.lower().replace(' ', '-')}"
            else:
                channel_name = f"{first_event.lower().replace(' ', '-')}-{building.lower().replace(' ', '-')}"
            
            # Create event channel with Slacker access
            event_channel = await get_or_create_channel(guild, channel_name, category, event_role)
            if event_channel and slacker_role:
                await add_slacker_access(event_channel, slacker_role)
    
    # Sort building categories alphabetically after creating new ones
    await sort_building_categories_alphabetically(guild)

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

async def add_role_to_building_chat(channel, role):
    """Add a role to a building chat channel permissions"""
    if not channel or not role:
        return
    
    try:
        # Get current overwrites
        overwrites = channel.overwrites
        
        # Set @everyone to not see the channel
        overwrites[channel.guild.default_role] = discord.PermissionOverwrite(read_messages=False)
        
        # Always ensure Slacker role has access
        slacker_role = discord.utils.get(channel.guild.roles, name="Slacker")
        if slacker_role:
            overwrites[slacker_role] = discord.PermissionOverwrite(
                read_messages=True,
                send_messages=True,
                read_message_history=True
            )
        
        # Add the role with permissions to see and participate
        overwrites[role] = discord.PermissionOverwrite(
            read_messages=True,
            send_messages=True,
            read_message_history=True
        )
        
        # Update channel permissions
        await channel.edit(overwrites=overwrites, reason=f"Added {role.name} to building chat access")
        print(f"🔒 Added {role.name} access to #{channel.name}")
        if slacker_role:
            print(f"🔑 Ensured {slacker_role.name} access to #{channel.name}")
        
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

async def setup_static_channels():
    """Create static categories and channels for Tournament Officials and Volunteers"""
    guild = bot.get_guild(GUILD_ID)
    if not guild:
        print("❌ Guild not found!")
        return
    
    print("🏗️ Setting up static channels...")
    
    # Get or create Slacker role for permissions
    slacker_role = discord.utils.get(guild.roles, name="Slacker")
    
    # Welcome Category
    print("👋 Setting up Welcome category...")
    welcome_category = await get_or_create_category(guild, "Welcome")
    if welcome_category:
        # Create welcome channel (visible to everyone)
        welcome_channel = await get_or_create_channel(guild, "welcome", welcome_category)
        if welcome_channel and slacker_role:
            await add_slacker_access(welcome_channel, slacker_role)
    
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
    
    # Volunteers Category  
    print("🙋 Setting up Volunteers category...")
    volunteers_category = await get_or_create_category(guild, "Volunteers")
    if volunteers_category:
        # Create regular text channels
        volunteer_text_channels = ["general", "useful-links", "random"]
        for channel_name in volunteer_text_channels:
            channel = await get_or_create_channel(guild, channel_name, volunteers_category)
            if channel and slacker_role:
                await add_slacker_access(channel, slacker_role)
        
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
            
        # Add Slacker access to help channel (whether it was just created or already existed)
        if help_channel and slacker_role:
            try:
                overwrites = help_channel.overwrites
                overwrites[slacker_role] = discord.PermissionOverwrite(
                    read_messages=True,
                    send_messages=True,
                    read_message_history=True,
                    create_public_threads=True,
                    send_messages_in_threads=True
                )
                await help_channel.edit(overwrites=overwrites, reason=f"Added {slacker_role.name} access")
                print(f"🔑 Added {slacker_role.name} access to #{help_channel.name} (forum)")
            except Exception as e:
                print(f"❌ Error adding Slacker access to forum #{help_channel.name}: {e}")
    
    print("✅ Finished setting up static channels")

async def give_slacker_access_to_all_channels():
    """Give Slacker role access to all existing channels"""
    guild = bot.get_guild(GUILD_ID)
    if not guild:
        print("❌ Guild not found!")
        return
    
    slacker_role = discord.utils.get(guild.roles, name="Slacker")
    if not slacker_role:
        print("⚠️ Slacker role not found, will be created when needed")
        return
    
    print(f"🔑 Adding {slacker_role.name} access to all channels...")
    
    static_channels = 0
    building_channels = 0
    forum_channels = 0
    
    # Categorize and add access to text channels
    static_categories = ["Welcome", "Tournament Officials", "Volunteers"]
    
    for channel in guild.text_channels:
        try:
            await add_slacker_access(channel, slacker_role)
            
            # Count by type
            if channel.category and channel.category.name in static_categories:
                static_channels += 1
            else:
                building_channels += 1
                
        except Exception as e:
            print(f"❌ Error adding Slacker access to #{channel.name}: {e}")
    
    # Add access to forum channels
    for channel in guild.channels:
        if channel.type == discord.ChannelType.forum:
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
                print(f"🔑 Added {slacker_role.name} access to #{channel.name} (forum)")
                forum_channels += 1
            except Exception as e:
                print(f"❌ Error adding Slacker access to forum #{channel.name}: {e}")
    
    print(f"✅ Added {slacker_role.name} access to:")
    print(f"   • {static_channels} static channels")
    print(f"   • {building_channels} building channels") 
    print(f"   • {forum_channels} forum channels")
    print(f"🔑 Total: {static_channels + building_channels + forum_channels} channels with Slacker access")

@bot.event
async def on_ready():
    print(f"Logged in as {bot.user} (ID: {bot.user.id})")
    
    # Check if server reset is enabled
    if RESET_SERVER:
        print("⚠️ ⚠️ ⚠️  SERVER RESET ENABLED!  ⚠️ ⚠️ ⚠️")
        await reset_server()
        print("🔄 Reset complete, continuing with normal setup...")
    
    print("🏗️ Setting up static channels...")
    await setup_static_channels()
    print("📋 Organizing categories alphabetically...")
    await sort_building_categories_alphabetically(bot.get_guild(GUILD_ID))
    print("🔑 Giving Slacker role access to all channels...")
    await give_slacker_access_to_all_channels()
    print("🔄 Starting member sync task...")
    sync_members.start()

@bot.event
async def on_member_join(member):
    """Handle role assignment and nickname setting when a user joins the server"""
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

@tasks.loop(minutes=1)
async def sync_members():
    """Every minute, read spreadsheet and invite any new Discord IDs."""
    print("🔄 Running member sync...")
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

    # Build set of already-joined member IDs
    joined = {m.id for m in guild.members}
    print(f"👥 Guild has {len(guild.members)} members")

    processed_count = 0
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
                    # Old format: username#1234
                    username, discriminator = discord_identifier.split("#", 1)
                    member = discord.utils.get(guild.members, name=username, discriminator=discriminator)
                else:
                    # New format: just username, or global display name
                    member = discord.utils.get(guild.members, name=discord_identifier)
                    if not member:
                        # Try by display name
                        member = discord.utils.get(guild.members, display_name=discord_identifier)
                    if not member:
                        # Try by global display name
                        member = discord.utils.get(guild.members, global_name=discord_identifier)
                
                if member:
                    discord_id = member.id
                    processed_count += 1
                    print(f"🔍 Found user by handle '{discord_identifier}' -> ID: {discord_id}")
                else:
                    print(f"⚠️ Could not find user with handle '{discord_identifier}'")
                    continue
            except Exception as e:
                print(f"❌ Error processing identifier '{discord_identifier}': {e}")
                continue
        
        if discord_id is None:
            continue

        if discord_id in joined:
            # User is already in server, check if they need role assignment and nickname
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
                            await member.add_roles(role, reason="Onboarding sync")
                            print(f"✅ Assigned role {role.name} to {member}")
                        except Exception as e:
                            print(f"⚠️ Could not add role {role_name} to {member}: {e}")
                
                # Set nickname if we have the required info
                if first_event:
                    # Get name from sheet if available, otherwise use Discord username
                    sheet_name = str(row.get("Name", "")).strip()
                    user_name = sheet_name if sheet_name else member.name
                    expected_nickname = f"{user_name} ({first_event})"
                    
                    # Truncate to 32 characters (Discord limit)
                    if len(expected_nickname) > 32:
                        expected_nickname = expected_nickname[:32]
                    
                    # Only update if nickname is different
                    if member.nick != expected_nickname:
                        try:
                            await member.edit(nick=expected_nickname, reason="Onboarding sync - updating nickname")
                            print(f"📝 Updated nickname for {member}: '{expected_nickname}'")
                        except discord.Forbidden:
                            print(f"❌ No permission to set nickname for {member}")
                        except Exception as e:
                            print(f"⚠️ Could not set nickname for {member}: {e}")
                
                # Set up building structure and channels
                building = str(row.get("Building 1", "")).strip()
                room = str(row.get("Room 1", "")).strip()
                if building and first_event:
                    await setup_building_structure(guild, building, first_event, room)
            continue

        # Fetch user object
        try:
            user = await bot.fetch_user(discord_id)
        except discord.NotFound:
            print(f"⚠️ User {discord_id} not found.")
            continue
        except Exception as e:
            print(f"❌ Error fetching user {discord_id}: {e}")
            continue

        # Create a single-use invite to the welcome channel (fallback to system channel or first text channel)
        welcome_channel = discord.utils.get(guild.text_channels, name="welcome")
        if welcome_channel:
            channel = welcome_channel
        elif guild.system_channel:
            channel = guild.system_channel
        elif guild.text_channels:
            channel = guild.text_channels[0]
        else:
            print(f"❌ No suitable channel found for invite in guild {guild.name}")
            continue

        invite = await channel.create_invite(max_uses=1, unique=True, reason="Onboarding from sheet")

        # DM the invite
        try:
            await user.send(
                f"Hi {user.name}! 👋\n"
                f"You've been added to **{guild.name}** by the Science Olympiad planning team.\n"
                f"Click here to join: {invite.url}"
            )
            print(f"✉️ Sent invite to {user} ({discord_id})")
        except discord.Forbidden:
            print(f"❌ Cannot DM user {discord_id}; they may have DMs off.")

        # Store the role assignments for when they join
        roles_to_queue = []
        
        master_role = str(row.get("Master Role", "")).strip()
        if master_role:
            roles_to_queue.append(master_role)
        
        first_event = str(row.get("First Event", "")).strip()
        if first_event:
            roles_to_queue.append(first_event)
        
        if roles_to_queue:
            # Get name from sheet if available, otherwise use Discord username
            sheet_name = str(row.get("Name", "")).strip()
            user_name = sheet_name if sheet_name else user.name
            
            pending_users[discord_id] = {
                "roles": roles_to_queue,
                "name": user_name,
                "first_event": first_event
            }
            roles_text = "', '".join(roles_to_queue)
            nickname_preview = f"{user_name} ({first_event})" if first_event else user_name
            # Truncate preview to 32 characters (Discord limit)
            if len(nickname_preview) > 32:
                nickname_preview = nickname_preview[:32]
            print(f"📝 Queued roles '{roles_text}' for {user} when they join")
            print(f"📝 Will set nickname: '{nickname_preview}'")
            
            # Set up building structure and channels
            building = str(row.get("Building 1", "")).strip()
            room = str(row.get("Room 1", "")).strip()
            if building and first_event:
                await setup_building_structure(guild, building, first_event, room)
    
    print(f"✅ Sync complete. Processed {processed_count} valid Discord IDs from {len(data)} rows.")

if __name__ == "__main__":
    bot.run(TOKEN)
