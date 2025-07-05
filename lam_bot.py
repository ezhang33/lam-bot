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
RESET_SERVER = False

intents = discord.Intents.default()
intents.members = True

bot = discord.Bot(
    intents=intents,
    default_guild_ids=[GUILD_ID])

# Set up gspread client
scope = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
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

async def perform_member_sync(guild, data):
    """Core member sync logic that can be used by both /sync command and /entertemplate"""
    # Build set of already-joined member IDs
    joined = {m.id for m in guild.members}

    processed_count = 0
    invited_count = 0
    role_assignments = 0
    nickname_updates = 0
    
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
                
                # Set up building structure
                building = str(row.get("Building 1", "")).strip()
                room = str(row.get("Room 1", "")).strip()
                if building and first_event:
                    await setup_building_structure(guild, building, first_event, room)
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
                    
                # Set up building structure
                building = str(row.get("Building 1", "")).strip()
                room = str(row.get("Room 1", "")).strip()
                if building and first_event:
                    await setup_building_structure(guild, building, first_event, room)
                        
        except Exception as e:
            print(f"❌ Error processing user {discord_id}: {e}")
    
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
              f"4. Set permissions to 'Viewer'\n"
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
                    "Set to 'Viewer' permissions.\n\n"
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
            
            # Trigger an immediate sync after successful connection
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
    
    # Setup commands
    embed.add_field(
        name="⚙️ `/entertemplate` `folder_link`",
        value=f"Connect to a new Google Drive folder. The bot will search within that folder for '{SHEET_FILE_NAME}' sheet and use it for syncing users.\n\n⚠️ **Important:** Use the 'Copy link' button from Google Drive's Share dialog, not the address bar URL!",
        inline=False
    )
    
    # Admin commands
    embed.add_field(
        name="🔄 `/sync` (Admin Only)",
        value="Manually trigger a member sync from the current Google Sheet. Shows detailed statistics about the sync results.",
        inline=False
    )
    
    embed.add_field(
        name="🔁 `/reloadcommands` (Admin Only)",
        value="Manually sync slash commands with Discord. Use this if commands aren't showing up or seem outdated.",
        inline=False
    )
    
    # Workflow
    embed.add_field(
        name="🚀 Quick Start Workflow",
        value="1. Use `/serviceaccount` to get the service account email\n"
              "2. Share your Google Sheet with that email (Viewer permissions)\n"
              "3. Get folder link: Right-click folder → Share → Copy link\n"
              "4. Use `/entertemplate` with that copied folder link\n"
              "5. Use `/sheetinfo` to verify the connection\n"
              "6. Use `/sync` to manually trigger the first sync\n"
              "7. Bot will automatically sync every minute after that",
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
              "4. Set permissions to 'Viewer'\n"
              "5. Click 'Send'\n\n"
              "**For entire folders:**\n"
              "1. Right-click your folder in Google Drive\n"
              "2. Click 'Share'\n"
              "3. Add the service account email above\n"
              "4. Set permissions to 'Viewer'\n"
              "5. Click 'Send'",
        inline=False
    )
    
    embed.set_footer(text="The service account only needs 'Viewer' permissions to read your data")
    
    await ctx.respond(embed=embed, ephemeral=True)

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
