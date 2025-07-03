import os
import discord
from discord.ext import commands
from google_sheets_service import GoogleSheetsService
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Bot configuration
TOKEN = os.getenv('DISCORD_TOKEN')
SPREADSHEET_ID = os.getenv('GOOGLE_SPREADSHEET_ID')

# Bot setup
intents = discord.Intents.default()
intents.message_content = True
bot = commands.Bot(command_prefix='!', intents=intents)

# Initialize Google Sheets service
sheets_service = GoogleSheetsService()

@bot.event
async def on_ready():
    print(f'{bot.user} has connected to Discord!')
    print(f'Bot is ready to read from Google Sheets')

@bot.command(name='read_sheet')
async def read_sheet(ctx, range_name: str = 'Sheet1!A1:Z100'):
    """
    Read data from Google Sheets
    Usage: !read_sheet [range] (default: Sheet1!A1:Z100)
    Example: !read_sheet Sheet1!A1:C10
    """
    try:
        # Read data from Google Sheets
        data = sheets_service.read_sheet(SPREADSHEET_ID, range_name)
        
        if not data:
            await ctx.send("No data found in the specified range.")
            return
        
        # Format the data for Discord
        response = "**Google Sheets Data:**\n```\n"
        
        # Add headers if they exist
        if data:
            headers = data[0]
            response += " | ".join(str(cell) for cell in headers) + "\n"
            response += "-" * (len(" | ".join(str(cell) for cell in headers))) + "\n"
            
            # Add data rows (limit to prevent message being too long)
            for row in data[1:10]:  # Show max 10 rows
                response += " | ".join(str(cell) if cell else "" for cell in row) + "\n"
            
            if len(data) > 11:  # 1 header + 10 data rows
                response += f"... and {len(data) - 11} more rows\n"
        
        response += "```"
        
        # Discord has a 2000 character limit for messages
        if len(response) > 2000:
            response = response[:1950] + "...\n```\n*Data truncated due to length*"
        
        await ctx.send(response)
        
    except Exception as e:
        await ctx.send(f"Error reading from Google Sheets: {str(e)}")

@bot.command(name='sheet_info')
async def sheet_info(ctx):
    """
    Get information about the configured Google Sheet
    """
    try:
        sheet_info = sheets_service.get_sheet_info(SPREADSHEET_ID)
        
        embed = discord.Embed(
            title="Google Sheet Information",
            color=discord.Color.blue()
        )
        
        embed.add_field(name="Title", value=sheet_info.get('title', 'N/A'), inline=False)
        embed.add_field(name="Spreadsheet ID", value=SPREADSHEET_ID, inline=False)
        
        sheets = sheet_info.get('sheets', [])
        if sheets:
            sheet_names = [sheet['properties']['title'] for sheet in sheets]
            embed.add_field(name="Available Sheets", value=", ".join(sheet_names), inline=False)
        
        await ctx.send(embed=embed)
        
    except Exception as e:
        await ctx.send(f"Error getting sheet info: {str(e)}")

@bot.command(name='search_sheet')
async def search_sheet(ctx, search_term: str, range_name: str = 'Sheet1!A1:Z1000'):
    """
    Search for a specific term in the Google Sheet
    Usage: !search_sheet "search term" [range]
    Example: !search_sheet "john doe" Sheet1!A1:D100
    """
    try:
        data = sheets_service.read_sheet(SPREADSHEET_ID, range_name)
        
        if not data:
            await ctx.send("No data found in the specified range.")
            return
        
        # Search for the term
        matches = []
        for row_idx, row in enumerate(data):
            for col_idx, cell in enumerate(row):
                if search_term.lower() in str(cell).lower():
                    matches.append({
                        'row': row_idx + 1,
                        'col': col_idx + 1,
                        'value': cell,
                        'full_row': row
                    })
        
        if not matches:
            await ctx.send(f"No matches found for '{search_term}'")
            return
        
        # Format results
        response = f"**Search Results for '{search_term}':**\n```\n"
        
        for match in matches[:5]:  # Show max 5 matches
            response += f"Row {match['row']}, Col {match['col']}: {match['value']}\n"
            response += f"Full row: {' | '.join(str(cell) for cell in match['full_row'])}\n"
            response += "-" * 50 + "\n"
        
        if len(matches) > 5:
            response += f"... and {len(matches) - 5} more matches\n"
        
        response += "```"
        
        await ctx.send(response)
        
    except Exception as e:
        await ctx.send(f"Error searching sheet: {str(e)}")

@bot.command(name='help_sheets')
async def help_sheets(ctx):
    """
    Show help for Google Sheets commands
    """
    embed = discord.Embed(
        title="Google Sheets Bot Commands",
        description="Available commands for reading Google Sheets",
        color=discord.Color.green()
    )
    
    embed.add_field(
        name="!read_sheet [range]", 
        value="Read data from Google Sheets\nExample: `!read_sheet Sheet1!A1:C10`", 
        inline=False
    )
    
    embed.add_field(
        name="!sheet_info", 
        value="Get information about the configured Google Sheet", 
        inline=False
    )
    
    embed.add_field(
        name="!search_sheet \"term\" [range]", 
        value="Search for a specific term in the sheet\nExample: `!search_sheet \"john doe\"`", 
        inline=False
    )
    
    embed.add_field(
        name="!help_sheets", 
        value="Show this help message", 
        inline=False
    )
    
    await ctx.send(embed=embed)

if __name__ == '__main__':
    if not TOKEN:
        print("Error: DISCORD_TOKEN not found in environment variables")
        exit(1)
    
    if not SPREADSHEET_ID:
        print("Error: GOOGLE_SPREADSHEET_ID not found in environment variables")
        exit(1)
    
    bot.run(TOKEN) 