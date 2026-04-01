import discord
from discord import app_commands
from discord.ext import commands, tasks
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, Alignment, PatternFill
import io
import os
import json
import asyncio
import calendar
import logging
from datetime import datetime
from dotenv import load_dotenv

# ─────────────────────────────────────────────
#  LOGGING SETUP
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
log = logging.getLogger("monsterbot")

load_dotenv()

# ─────────────────────────────────────────────
#  CONFIG & PATHS
# ─────────────────────────────────────────────
TOKEN              = os.getenv("DISCORD_TOKEN", "").strip()
GUILD_ID           = int(os.getenv("GUILD_ID", 0))
TICKET_CAT_ID      = int(os.getenv("TICKET_CATEGORY_ID", 0))
STAFF_ROLE_ID      = int(os.getenv("STAFF_ROLE_ID", 0))
LOG_CHANNEL_ID     = int(os.getenv("LOG_CHANNEL_ID", 0))
LEADERBOARD_CH_ID  = int(os.getenv("LEADERBOARD_CHANNEL_ID", 0))
PAYMENT_CH_ID      = int(os.getenv("PAYMENT_CHANNEL_ID", 0))

# Bestandsnamen
TEMPLATE_PATH      = "Clipfarming_Template.xlsx"
DATA_FILE          = "monthly_data.json"
STATE_FILE         = "bot_state.json"

OPEN_DAYS = {1, 2, 3, 4}
MONTHS_NL = ["Januari","Februari","Maart","April","Mei","Juni","Juli","Augustus","September","Oktober","November","December"]

# ─────────────────────────────────────────────
#  EXCEL TEMPLATE GENERATOR (Voor Railway)
# ─────────────────────────────────────────────
def ensure_excel_template():
    """Maakt de template aan als deze niet bestaat of aangepast moet worden."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Clip verdiensten"

    # Styling
    header_fill = PatternFill(start_color="7B2FF7", end_color="7B2FF7", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    headers = ["Discord Naam", "E-mail", "Platform", "Link naar Clip", "Views (Kies lijst)", "Verdiensten"]
    ws.append(headers)
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    # Prijslijst info in de sheet
    ws["G2"], ws["G3"], ws["G4"] = "PRIJSLIJST:", "10k - 500k: €1 per 10k", "1M+ views: €250,00"
    
    # Dropdown: 10.000 t/m 500.000 (stappen van 10k) + 1.000.000
    steps = [str(i) for i in range(10000, 510000, 10000)]
    steps.append("1000000")
    dv = DataValidation(type="list", formula1=f'"{",".join(steps)}"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("E2:E50") # Pas toe op de eerste 50 rijen

    ws.column_dimensions['D'].width = 40
    ws.column_dimensions['E'].width = 20
    wb.save(TEMPLATE_PATH)
    log.info("Excel template gegenereerd/gecontroleerd.")

# ─────────────────────────────────────────────
#  LOGICA & HELPERS
# ─────────────────────────────────────────────
def calc_earnings(views):
    """Nieuwe regels: 1M = 250 euro. Anders 1 euro per 10k."""
    if views >= 1_000_000:
        return 250.0
    return float(views // 10_000)

def is_staff(member):
    return any(role.id == STAFF_ROLE_ID for role in member.roles)

def is_bot_open():
    if datetime.now().day in OPEN_DAYS: return True
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r") as f: return json.load(f).get("bot_enabled", False)
    return False

# ─────────────────────────────────────────────
#  EMBEDS
# ─────────────────────────────────────────────
def build_prijslijst_embed():
    embed = discord.Embed(title="💰 Nieuwe Prijslijst", color=0x00c9a7)
    embed.add_field(name="Tarieven", value="• 10k - 500k views: **€1,- per 10.000**\n• 1.000.000+ views: **€250,- Jackpot!**", inline=False)
    embed.set_footer(text="Gebruik de dropdown in het Excel-bestand.")
    return embed

def build_summary_embed(data, user):
    embed = discord.Embed(title="Overzicht Inzending", description=f"Ingediend door: {user.mention}", color=0x7b2ff7)
    for i, c in enumerate(data["clips"][:10], 1):
        embed.add_field(name=f"Clip {i}", value=f"{c['views']:,} views -> **€{c['earning']:.2f}**", inline=True)
    embed.add_field(name="TOTAAL", value=f"## €{data['total']:.2f}", inline=False)
    return embed

# ─────────────────────────────────────────────
#  VIEWS & MODALS
# ─────────────────────────────────────────────
class PaymentProcessView(discord.ui.View):
    def __init__(self, user_id, bedrag):
        super().__init__(timeout=None)
        self.user_id, self.bedrag = user_id, bedrag

    @discord.ui.button(label="Markeer als Betaald", style=discord.ButtonStyle.success, emoji="💸")
    async def mark_paid(self, interaction, button):
        if not is_staff(interaction.user): return
        await interaction.message.delete()
        await interaction.response.send_message(f"✅ Betaling van {self.bedrag} aan <@{self.user_id}> is verwerkt.")

class StaffApprovalView(discord.ui.View):
    def __init__(self, data, user_id, methode, gegevens):
        super().__init__(timeout=None)
        self.data, self.user_id, self.methode, self.gegevens = data, user_id, methode, gegevens

    @discord.ui.button(label="Goedkeuren", style=discord.ButtonStyle.success)
    async def approve(self, interaction, button):
        if not is_staff(interaction.user): return
        pay_ch = interaction.guild.get_channel(PAYMENT_CH_ID)
        if pay_ch:
            emb = discord.Embed(title="Nieuwe Uitbetaling", color=0x2ecc71)
            emb.add_field(name="Gebruiker", value=f"<@{self.user_id}>")
            emb.add_field(name="Bedrag", value=f"€{self.data['total']:.2f}")
            emb.add_field(name="Info", value=f"{self.methode}: {self.gegevens}")
            await pay_ch.send(embed=emb, view=PaymentProcessView(self.user_id, f"€{self.data['total']:.2f}"))
        
        await interaction.response.send_message("Goedgekeurd! Kanaal sluit over 10 sec.")
        await asyncio.sleep(10)
        await interaction.channel.delete()

class PaymentModal(discord.ui.Modal, title="Betaalgegevens"):
    methode = discord.ui.TextInput(label="Methode (bijv. PayPal)", placeholder="PayPal", required=True)
    gegevens = discord.ui.TextInput(label="Account (bijv. e-mail)", placeholder="naam@mail.com", required=True)

    async def on_submit(self, interaction):
        cat = interaction.guild.get_channel(TICKET_CAT_ID)
        overwrites = {
            interaction.guild.default_role: discord.PermissionOverwrite(read_messages=False),
            interaction.user: discord.PermissionOverwrite(read_messages=True, send_messages=True, attach_files=True),
            interaction.guild.me: discord.PermissionOverwrite(read_messages=True, send_messages=True)
        }
        chan = await interaction.guild.create_text_channel(name=f"ticket-{interaction.user.name}", category=cat, overwrites=overwrites)
        chan.topic = f"User:{interaction.user.id} | {self.methode.value}:{self.gegevens.value}"
        
        with open(TEMPLATE_PATH, "rb") as f:
            await chan.send(f"{interaction.user.mention} Vul dit bestand in:", file=discord.File(f, TEMPLATE_PATH), embed=build_prijslijst_embed())
        await interaction.response.send_message(f"Ticket aangemaakt: {chan.mention}", ephemeral=True)

# ─────────────────────────────────────────────
#  BOT EVENTS & COMMANDS
# ─────────────────────────────────────────────
intents = discord.Intents.default()
intents.message_content = True
intents.members = True
bot = commands.Bot(command_prefix="!", intents=intents)

@bot.event
async def on_ready():
    ensure_excel_template()
    await bot.tree.sync()
    log.info(f"{bot.user} is online!")

@bot.tree.command(name="uitbetaling", description="Vraag je verdiensten aan")
async def uitbetaling(interaction: discord.Interaction):
    if not is_bot_open() and not is_staff(interaction.user):
        return await interaction.response.send_message("De bot is momenteel gesloten.", ephemeral=True)
    await interaction.response.send_modal(PaymentModal())

@bot.event
async def on_message(message):
    if message.author.bot or not message.attachments: return
    if not message.channel.category or message.channel.category.id != TICKET_CAT_ID: return

    for attach in message.attachments:
        if attach.filename.endswith(".xlsx"):
            raw = await attach.read()
            wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
            ws = wb.active
            
            clips = []
            total_money = 0.0
            
            # Lees rijen 2 t/m 20
            for row in ws.iter_rows(min_row=2, max_row=20, max_col=5):
                if not row[4].value: continue
                try:
                    v = int(row[4].value)
                    earn = calc_earnings(v)
                    clips.append({"views": v, "earning": earn})
                    total_money += earn
                except: continue
            
            if clips:
                # Info uit topic halen
                methode, gegevens = "Onbekend", "Onbekend"
                try:
                    parts = message.channel.topic.split("|")
                    if len(parts) > 1: methode, gegevens = parts[1].split(":", 1)
                except: pass

                data = {"clips": clips, "total": total_money}
                view = StaffApprovalView(data, message.author.id, methode, gegevens)
                await message.channel.send(embed=build_summary_embed(data, message.author), view=view)

bot.run(TOKEN)
