import discord
from discord import app_commands
from discord.ext import commands, tasks
import openpyxl
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
    handlers=[
        logging.FileHandler("monsterbot.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger("monsterbot")

load_dotenv()

# ─────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────
def get_env_int(key):
    val = os.getenv(key, "0").strip()
    try:
        return int(val)
    except ValueError:
        log.error(f"Ongeldige waarde voor {key}: '{val}'")
        return 0

TOKEN              = os.getenv("DISCORD_TOKEN", "").strip()
GUILD_ID           = get_env_int("GUILD_ID")
TICKET_CAT_ID      = get_env_int("TICKET_CATEGORY_ID")
STAFF_ROLE_ID      = get_env_int("STAFF_ROLE_ID")
LOG_CHANNEL_ID     = get_env_int("LOG_CHANNEL_ID")
LEADERBOARD_CH_ID  = get_env_int("LEADERBOARD_CHANNEL_ID")
PAYMENT_CH_ID      = get_env_int("PAYMENT_CHANNEL_ID")
TEMPLATE_PATH      = os.path.join(os.path.dirname(__file__), "Clipfarming_Template.xlsx")
DATA_FILE          = os.path.join(os.path.dirname(__file__), "monthly_data.json")
STATE_FILE         = os.path.join(os.path.dirname(__file__), "bot_state.json")

# Bot is beschikbaar op dag 1 t/m 4 van elke maand
OPEN_DAYS = {1, 2, 3, 4}

MONTHS_NL = [
    "Januari","Februari","Maart","April","Mei","Juni",
    "Juli","Augustus","September","Oktober","November","December"
]

if not TOKEN:
    log.critical("DISCORD_TOKEN ontbreekt in .env - bot kan niet starten")
    raise SystemExit(1)

# ─────────────────────────────────────────────
#  INTENTS & BOT
# ─────────────────────────────────────────────
intents = discord.Intents.default()
intents.message_content = True
intents.members = True

bot  = commands.Bot(command_prefix="!", intents=intents)
tree = bot.tree

# ─────────────────────────────────────────────
#  STATE (bot aan/uit toggle)
# ─────────────────────────────────────────────
def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {"bot_enabled": True}

def save_state(state):
    try:
        with open(STATE_FILE, "w") as f:
            json.dump(state, f, indent=2)
    except Exception as e:
        log.error(f"Kon state niet opslaan: {e}")

def is_bot_open():
    """Bot is open als: handmatig ingeschakeld OF het zijn de eerste 4 dagen van de maand."""
    state = load_state()
    if not state.get("bot_enabled", True):
        return False
    return datetime.now().day in OPEN_DAYS

# ─────────────────────────────────────────────
#  LEADERBOARD DATA
# ─────────────────────────────────────────────
def calc_earnings(views):
    if views >= 1_000_000:
        return 250.0
    return float(views // 10_000)

def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r") as f:
                return json.load(f)
        except Exception as e:
            log.error(f"Kon data niet laden: {e}")
    return {}

def save_data(data):
    try:
        with open(DATA_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        log.error(f"Kon data niet opslaan: {e}")

def get_month_key():
    now = datetime.now()
    return f"{now.year}-{now.month:02d}"

def add_to_leaderboard(discord_naam, user_id, views, earnings):
    data = load_data()
    key  = get_month_key()
    if key not in data:
        data[key] = {}
    uid = str(user_id)
    if uid not in data[key]:
        data[key][uid] = {"naam": discord_naam, "views": 0, "earnings": 0.0}
    data[key][uid]["views"]    += max(0, int(views))
    data[key][uid]["earnings"] += max(0.0, float(earnings))
    data[key][uid]["naam"]      = str(discord_naam)[:50]
    save_data(data)

# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
async def send_log(guild, embed, file=None):
    if not LOG_CHANNEL_ID:
        return
    try:
        ch = guild.get_channel(LOG_CHANNEL_ID)
        if not ch:
            return
        if file:
            await ch.send(embed=embed, file=file)
        else:
            await ch.send(embed=embed)
    except discord.Forbidden:
        log.warning("Geen toegang tot log-kanaal")
    except Exception as e:
        log.error(f"send_log fout: {e}")

def is_staff(member):
    if not STAFF_ROLE_ID:
        return False
    role = member.guild.get_role(STAFF_ROLE_ID)
    return role in member.roles if role else False

def sanitize(text, max_len=200):
    """Verwijder gevaarlijke tekens en beperk lengte."""
    if not text:
        return ""
    clean = str(text).strip()
    clean = clean.replace("@everyone", "").replace("@here", "")
    return clean[:max_len]

def parse_submission(file_bytes):
    """Parse het ingediende Excel bestand veilig."""
    if len(file_bytes) > 5 * 1024 * 1024:  # max 5MB
        log.warning("Ingediend bestand te groot (>5MB)")
        return None
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    except Exception as e:
        log.warning(f"Excel kon niet worden geopend: {e}")
        return None

    results      = []
    total_views  = 0
    total        = 0.0
    discord_naam = None
    email        = None

    for sheet_name in wb.sheetnames[:12]:  # max 12 sheets
        try:
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=4, max_row=18, max_col=6):
                naam_cell    = row[0].value if len(row) > 0 else None
                email_cell   = row[1].value if len(row) > 1 else None
                platform_val = row[2].value if len(row) > 2 else None
                link_val     = row[3].value if len(row) > 3 else None
                views_val    = row[4].value if len(row) > 4 else None

                if not views_val or str(views_val).strip() in ("", "Views (werkelijk)"):
                    continue
                try:
                    views = int(str(views_val).replace(".", "").replace(",", ""))
                    if views < 0 or views > 100_000_000:
                        continue
                except (ValueError, OverflowError):
                    continue

                if not discord_naam and naam_cell and "[Vul" not in str(naam_cell):
                    discord_naam = sanitize(naam_cell, 50)
                if not email and email_cell and "[Vul" not in str(email_cell):
                    email = sanitize(email_cell, 100)

                floored = (views // 10_000) * 10_000
                earning = calc_earnings(floored)
                total  += earning
                total_views += floored

                link = sanitize(link_val, 200)
                # Alleen http(s) links toestaan
                if link and not link.startswith(("http://", "https://")):
                    link = "Ongeldige link"

                results.append({
                    "platform": sanitize(platform_val, 30),
                    "link":     link,
                    "views":    views,
                    "floored":  floored,
                    "earning":  earning,
                })

                if len(results) >= 20:
                    break
        except Exception as e:
            log.warning(f"Fout bij lezen sheet {sheet_name}: {e}")
            continue

    if not results:
        return None

    return {
        "discord_naam": discord_naam or "Onbekend",
        "email":        email or "Onbekend",
        "clips":        results,
        "total_views":  total_views,
        "total":        round(total, 2),
    }

def build_summary_embed(data, user):
    embed = discord.Embed(
        title="Betalingsoverzicht - Ingediend",
        description=(
            f"**Discord:** {sanitize(data['discord_naam'])}\n"
            f"**E-mail:** {sanitize(data['email'])}\n"
            f"**Ingediend door:** {user.mention}\n"
            f"**Datum:** {datetime.now().strftime('%d %B %Y')}"
        ),
        color=0x7b2ff7,
        timestamp=datetime.utcnow(),
    )
    clips = data["clips"]
    lines = []
    for i, c in enumerate(clips[:20], 1):
        afgerond = f" *(afgerond naar {c['floored']:,})*" if c["views"] != c["floored"] else ""
        link_display = c['link'][:60] + ('...' if len(c['link']) > 60 else '')
        lines.append(
            f"`{i}.` **{c['platform']}** | {c['views']:,} views{afgerond} - EUR {c['earning']:.2f}\n"
            f"   {link_display}"
        )
    embed.add_field(name=f"Clips ({len(clips)})", value="\n".join(lines) if lines else "Geen clips", inline=False)
    embed.add_field(name="Totaal te betalen", value=f"**EUR {data['total']:.2f}**", inline=False)
    embed.set_footer(text="MONSTERBOT - Wacht op goedkeuring van staff")
    return embed

def build_prijslijst_embed():
    embed = discord.Embed(
        title="Prijslijst",
        description=(
            "EUR 1 per 10.000 views - altijd afgerond naar beneden.\n"
            "Views gelden voor een termijn van 1 maand.\n"
            "Voorbeeld: 16.000 views = 10.000 = EUR 1"
        ),
        color=0x00c9a7,
    )
    lines = [
        "10.000 views - EUR 1",
        "20.000 views - EUR 2  (per 10k erbij)",
        "50.000 views - EUR 5",
        "100.000 views - EUR 10",
        "250.000 views - EUR 25",
        "500.000 views - EUR 50",
        "1.000.000 views - EUR 250",
    ]
    embed.add_field(name="Tarieven", value="\n".join(lines), inline=False)
    return embed

def bot_closed_embed():
    now      = datetime.now()
    last_day = calendar.monthrange(now.year, now.month)[1]
    next_month = MONTHS_NL[now.month % 12]
    return discord.Embed(
        title="Bot momenteel gesloten",
        description=(
            f"Betalingen kunnen alleen in de **eerste 4 dagen van de maand** worden ingediend.\n\n"
            f"De bot opent weer op **1 {next_month}**."
        ),
        color=0xff4444,
    )

# ─────────────────────────────────────────────
#  BETAALVOORKEUR MODAL
# ─────────────────────────────────────────────
class PaymentModal(discord.ui.Modal, title="Betaalvoorkeur"):
    betaalmethode = discord.ui.TextInput(
        label="Betaalmethode",
        placeholder="Bijv. PayPal, Tikkie, Bankoverschrijving...",
        required=True,
        max_length=50,
    )
    betaalgegevens = discord.ui.TextInput(
        label="Betaalgegevens (e-mail/IBAN/telefoon)", 
        placeholder="Bijv. naam@email.com of NL00BANK0123456789",
        required=True,
        max_length=150,
    )
    async def on_submit(self, interaction):
        methode  = sanitize(self.betaalmethode.value, 50)
        gegevens = sanitize(self.betaalgegevens.value, 150)
        await create_ticket(interaction, methode, gegevens)

# ─────────────────────────────────────────────
#  TICKET AANMAKEN
# ─────────────────────────────────────────────
@tree.command(name="uitbetaling", description="Maak een betalingsticket aan om je clip-verdiensten in te dienen.")
@app_commands.guilds(discord.Object(id=GUILD_ID))
async def uitbetaling(interaction):
    # Tijdslot check — alleen eerste 4 dagen of als staff het open heeft gezet
    if not is_bot_open() and not is_staff(interaction.user):
        await interaction.response.send_message(embed=bot_closed_embed(), ephemeral=True)
        return

    guild    = interaction.guild
    user     = interaction.user
    category = guild.get_channel(TICKET_CAT_ID)

    if not category or not isinstance(category, discord.CategoryChannel):
        await interaction.response.send_message("Ticket-categorie niet gevonden. Vraag een admin.", ephemeral=True)
        log.error(f"TICKET_CATEGORY_ID ({TICKET_CAT_ID}) niet gevonden in guild")
        return

    existing = discord.utils.get(category.text_channels, name=f"ticket-{user.name.lower().replace(' ', '-')}")
    if existing:
        await interaction.response.send_message(f"Je hebt al een open ticket: {existing.mention}", ephemeral=True)
        return

    await interaction.response.send_modal(PaymentModal())

async def create_ticket(interaction, betaalmethode, betaalgegevens):
    await interaction.response.defer(ephemeral=True)
    guild      = interaction.guild
    user       = interaction.user
    category   = guild.get_channel(TICKET_CAT_ID)
    staff_role = guild.get_role(STAFF_ROLE_ID)

    if not category:
        await interaction.followup.send("Fout: categorie niet gevonden.", ephemeral=True)
        return

    try:
        overwrites = {
            guild.default_role: discord.PermissionOverwrite(read_messages=False),
            user:               discord.PermissionOverwrite(read_messages=True, send_messages=True, attach_files=True),
            guild.me:           discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_channels=True),
        }
        if staff_role:
            overwrites[staff_role] = discord.PermissionOverwrite(read_messages=True, send_messages=True)

        channel = await guild.create_text_channel(
            name=f"ticket-{user.name.lower().replace(' ', '-')}",
            category=category,
            overwrites=overwrites,
            topic=f"Ticket:{user.id}|{betaalmethode}:{betaalgegevens}|{datetime.now().strftime('%d/%m/%Y')}",
        )
    except discord.Forbidden:
        await interaction.followup.send("Fout: ik heb geen rechten om kanalen aan te maken.", ephemeral=True)
        log.error("Geen rechten om ticket-kanaal aan te maken")
        return
    except Exception as e:
        await interaction.followup.send("Er ging iets mis bij het aanmaken van je ticket.", ephemeral=True)
        log.error(f"create_ticket fout: {e}")
        return

    month_name = MONTHS_NL[datetime.now().month - 1]
    try:
        with open(TEMPLATE_PATH, "rb") as f:
            file_data = f.read()
    except FileNotFoundError:
        await interaction.followup.send("Fout: Excel template niet gevonden op de server.", ephemeral=True)
        log.critical(f"Template niet gevonden: {TEMPLATE_PATH}")
        return

    excel_file = discord.File(
        io.BytesIO(file_data),
        filename=f"Clipfarming_Verdiensten_{month_name}_{datetime.now().year}.xlsx"
    )

    embed = discord.Embed(
        title="Betalingsticket aangemaakt",
        description=(
            f"Hey {user.mention}! Download het Excel-bestand hierboven en volg de stappen:\n\n"
            "**1.** Vul het Excel-bestand in\n"
            "**2.** Upload het ingevulde Excel-bestand in dit ticket\n"
            "**3.** Staff verwerkt de betaling en komt zo snel mogelijk bij je terug\n\n"
            f"**Betaalmethode:** {betaalmethode}\n"
            f"**Betaalgegevens:** {betaalgegevens}"
        ),
        color=0x7b2ff7,
        timestamp=datetime.utcnow(),
    )
    embed.set_footer(text="MONSTERBOT - Ticket systeem")

    try:
        msg = await channel.send(
            content=f"{user.mention}" + (f" {staff_role.mention}" if staff_role else ""),
            file=excel_file,
            embeds=[embed, build_prijslijst_embed()],
            view=TicketCloseView(user.id),
        )
        await msg.pin()
    except Exception as e:
        log.error(f"Fout bij sturen ticketbericht: {e}")

    log_embed = discord.Embed(
        title="Ticket aangemaakt",
        description=(
            f"**Gebruiker:** {user.mention} (`{user}`)\n"
            f"**Kanaal:** {channel.mention}\n"
            f"**Betaalmethode:** {betaalmethode}\n"
            f"**Betaalgegevens:** {betaalgegevens}\n"
            f"**Datum:** {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        ),
        color=0x7b2ff7,
        timestamp=datetime.utcnow(),
    )
    log_embed.set_footer(text="MONSTERBOT - Ticket Log")
    await send_log(guild, log_embed)
    log.info(f"Ticket aangemaakt voor {user} ({user.id})")
    await interaction.followup.send(f"Ticket aangemaakt: {channel.mention}", ephemeral=True)

# ─────────────────────────────────────────────
#  TICKET SLUITEN
# ─────────────────────────────────────────────
class TicketCloseView(discord.ui.View):
    def __init__(self, owner_id):
        super().__init__(timeout=None)
        self.owner_id = owner_id

    @discord.ui.button(label="Ticket sluiten", style=discord.ButtonStyle.danger, custom_id="close_ticket")
    async def close_ticket(self, interaction, button):
        if not is_staff(interaction.user) and interaction.user.id != self.owner_id:
            await interaction.response.send_message("Geen toegang.", ephemeral=True)
            return
        await interaction.response.send_message("Ticket wordt gesloten in 5 seconden...")
        await asyncio.sleep(5)
        try:
            await interaction.channel.delete(reason=f"Gesloten door {interaction.user}")
        except Exception as e:
            log.error(f"Kon ticket niet sluiten: {e}")
            
# ─────────────────────────────────────────────
#  BETAAL KANAAL KNOP (VOOR STAFF)
# ─────────────────────────────────────────────
class PaymentProcessView(discord.ui.View):
    def __init__(self, user_id, bedrag, discord_naam):
        super().__init__(timeout=None)
        self.user_id = user_id
        self.bedrag = bedrag
        self.discord_naam = discord_naam

    @discord.ui.button(label="Betalen", style=discord.ButtonStyle.primary, emoji="💸", custom_id="mark_as_paid")
    async def mark_as_paid(self, interaction: discord.Interaction, button: discord.ui.Button):
        if not is_staff(interaction.user):
            await interaction.response.send_message("Alleen staff mag dit doen.", ephemeral=True)
            return

        # 1. Log de actie
        log_embed = discord.Embed(
            title="✅ Betaling Voltooid",
            description=(
                f"**Stafflid:** {interaction.user.mention}\n"
                f"**Ontvanger:** <@{self.user_id}> ({self.discord_naam})\n"
                f"**Bedrag:** {self.bedrag}\n"
                f"**Status:** Uitbetaald en verwijderd uit lijst."
            ),
            color=0x2ecc71,
            timestamp=datetime.utcnow()
        )
        await send_log(interaction.guild, log_embed)

        # 2. Verwijder het bericht uit het kanaal
        await interaction.message.delete()
        
        # 3. Optioneel: Bevestiging naar het stafflid (ephemeral)
        log.info(f"Betaling aan {self.discord_naam} voltooid door {interaction.user}")
# ─────────────────────────────────────────────
#  STAFF APPROVAL KNOPPEN
# ─────────────────────────────────────────────
class StaffApprovalView(discord.ui.View):
    def __init__(self, total, total_views, discord_naam, user_id, betaalmethode, betaalgegevens):
        super().__init__(timeout=None)
        self.total          = total
        self.total_views    = total_views
        self.discord_naam   = discord_naam
        self.user_id        = user_id
        self.betaalmethode  = betaalmethode
        self.betaalgegevens = betaalgegevens

    @discord.ui.button(label="Goedkeuren", style=discord.ButtonStyle.success, custom_id="approve_pay")
    async def approve(self, interaction, button):
        if not is_staff(interaction.user):
            await interaction.response.send_message("Alleen staff.", ephemeral=True)
            return

        embed = discord.Embed(
            title="Goedgekeurd",
            description=(
                f"**Bedrag:** EUR {self.total:.2f}\n"
                f"**Goedgekeurd door:** {interaction.user.mention}\n"
                "Ticket sluit in 15 seconden."
            ),
            color=0x00c9a7,
        )
        await interaction.response.send_message(embed=embed)
        self.stop()

        # Betaalkanaal
        if PAYMENT_CH_ID:
            try:
                pay_ch = interaction.guild.get_channel(PAYMENT_CH_ID)
                if pay_ch:
                    pay_embed = discord.Embed(title="Nieuwe uitbetaling vereist", color=0x00c9a7, timestamp=datetime.utcnow())
                    pay_embed.add_field(name="Naam",           value=self.discord_naam,         inline=True)
                    pay_embed.add_field(name="Discord",        value=f"<@{self.user_id}>",      inline=True)
                    pay_embed.add_field(name="Bedrag",         value=f"EUR {self.total:.2f}",   inline=True)
                    pay_embed.add_field(name="Betaalmethode",  value=self.betaalmethode,        inline=True)
                    pay_embed.add_field(name="Betaalgegevens", value=self.betaalgegevens,       inline=True)
                    pay_embed.add_field(name="Goedgekeurd door", value=interaction.user.mention, inline=True)
                    pay_embed.set_footer(text="Klik op de knop hieronder als het is overgemaakt.")

                    # HIER VOEGEN WE DE NIEUWE VIEW TOE:
                    pay_view = PaymentProcessView(
                        user_id=self.user_id, 
                        bedrag=f"EUR {self.total:.2f}", 
                        discord_naam=self.discord_naam
                    )
                    
                    await pay_ch.send(embed=pay_embed, view=pay_view)
            except Exception as e:
                log.error(f"Fout bij sturen naar betaalkanaal: {e}")

        add_to_leaderboard(self.discord_naam, self.user_id, self.total_views, self.total)

        log_embed = discord.Embed(
            title="Betaling goedgekeurd",
            description=(
                f"**Door:** {interaction.user.mention}\n"
                f"**Bedrag:** EUR {self.total:.2f}\n"
                f"**Datum:** {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            ),
            color=0x00c9a7,
            timestamp=datetime.utcnow(),
        )
        log_embed.set_footer(text="MONSTERBOT - Ticket Log")
        await send_log(interaction.guild, log_embed)
        log.info(f"Betaling goedgekeurd door {interaction.user} - EUR {self.total:.2f}")

        await asyncio.sleep(15)
        try:
            await interaction.channel.delete(reason="Betaling goedgekeurd")
        except Exception as e:
            log.error(f"Kon kanaal niet verwijderen na goedkeuring: {e}")

    @discord.ui.button(label="Afwijzen", style=discord.ButtonStyle.danger, custom_id="reject_pay")
    async def reject(self, interaction, button):
        if not is_staff(interaction.user):
            await interaction.response.send_message("Alleen staff.", ephemeral=True)
            return
        await interaction.response.send_modal(RejectModal())

    @discord.ui.button(label="Nakijken", style=discord.ButtonStyle.secondary, custom_id="review_pay")
    async def review(self, interaction, button):
        if not is_staff(interaction.user):
            await interaction.response.send_message("Alleen staff.", ephemeral=True)
            return
        try:
            # Hernoem kanaal met prefix zodat het duidelijk is
            new_name = f"nakijken-{interaction.channel.name.replace('ticket-', '')}"
            await interaction.channel.edit(name=new_name)
        except Exception as e:
            log.warning(f"Kon kanaal niet hernoemen: {e}")

        embed = discord.Embed(
            title="Moet worden nagekeken",
            description=(
                f"Dit ticket is gemarkeerd als **te nakijken** door {interaction.user.mention}.\n"
                "Staff zal dit zo snel mogelijk controleren."
            ),
            color=0xf0a500,
            timestamp=datetime.utcnow(),
        )
        await interaction.response.send_message(embed=embed)

        log_embed = discord.Embed(
            title="Ticket gemarkeerd: Nakijken",
            description=(
                f"**Door:** {interaction.user.mention}\n"
                f"**Kanaal:** {interaction.channel.mention}\n"
                f"**Datum:** {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            ),
            color=0xf0a500,
            timestamp=datetime.utcnow(),
        )
        log_embed.set_footer(text="MONSTERBOT - Ticket Log")
        await send_log(interaction.guild, log_embed)

class RejectModal(discord.ui.Modal, title="Afwijzingsreden"):
    reden = discord.ui.TextInput(
        label="Waarom wordt dit afgewezen?",
        style=discord.TextStyle.paragraph,
        placeholder="Bijv. views niet geverifieerd, verkeerd template...",
        required=True,
        max_length=500,
    )
    async def on_submit(self, interaction):
        reden_clean = sanitize(self.reden.value, 500)
        embed = discord.Embed(
            title="Afgewezen",
            description=f"**Reden:** {reden_clean}\n\nPas je inzending aan en probeer opnieuw.",
            color=0xff4444,
        )
        await interaction.response.send_message(embed=embed)
        log_embed = discord.Embed(
            title="Betaling afgewezen",
            description=(
                f"**Door:** {interaction.user.mention}\n"
                f"**Reden:** {reden_clean}\n"
                f"**Datum:** {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            ),
            color=0xff4444,
            timestamp=datetime.utcnow(),
        )
        log_embed.set_footer(text="MONSTERBOT - Ticket Log")
        await send_log(interaction.guild, log_embed)
        log.info(f"Betaling afgewezen door {interaction.user}: {reden_clean[:50]}")

# ─────────────────────────────────────────────
#  EXCEL UPLOAD VERWERKEN
# ─────────────────────────────────────────────
@bot.event
async def on_message(message):
    if message.author.bot:
        return
    if not message.channel.category or message.channel.category.id != TICKET_CAT_ID:
        await bot.process_commands(message)
        return

    for attachment in message.attachments:
        if not attachment.filename.lower().endswith((".xlsx", ".xls")):
            continue
        if attachment.size > 5 * 1024 * 1024:
            await message.channel.send("Bestand te groot (max 5MB).")
            continue

        try:
            file_bytes = await attachment.read()
        except Exception as e:
            log.error(f"Kon bijlage niet lezen: {e}")
            await message.channel.send("Kon het bestand niet lezen. Probeer opnieuw.")
            continue

        data = parse_submission(file_bytes)

        if not data:
            embed = discord.Embed(
                title="Bestand niet leesbaar",
                description="Gebruik het meegeleverde template en vul de Views-kolom correct in.",
                color=0xff8c00,
            )
            await message.channel.send(embed=embed)
            continue

        # Betaalinfo uit channel topic
        topic          = message.channel.topic or ""
        betaalmethode  = "Onbekend"
        betaalgegevens = "Onbekend"
        try:
            parts = topic.split("|")
            if len(parts) >= 2:
                pay_raw = parts[1].strip()
                if ":" in pay_raw:
                    betaalmethode, betaalgegevens = pay_raw.split(":", 1)
                    betaalmethode  = sanitize(betaalmethode, 50)
                    betaalgegevens = sanitize(betaalgegevens, 150)
        except Exception:
            pass

        staff_role = message.guild.get_role(STAFF_ROLE_ID)
        view = StaffApprovalView(
            total=data["total"],
            total_views=data["total_views"],
            discord_naam=data["discord_naam"],
            user_id=message.author.id,
            betaalmethode=betaalmethode,
            betaalgegevens=betaalgegevens,
        )

        try:
            await message.channel.send(
                content=(f"{staff_role.mention} - nieuw verzoek ter beoordeling!" if staff_role else ""),
                embed=build_summary_embed(data, message.author),
                view=view,
            )
        except Exception as e:
            log.error(f"Fout bij sturen summary: {e}")
            continue

        # Log + bijlage
        log_embed = discord.Embed(
            title="Excel ingediend",
            description=(
                f"**Gebruiker:** {message.author.mention}\n"
                f"**Naam:** {data['discord_naam']}\n"
                f"**Clips:** {len(data['clips'])}\n"
                f"**Totaal:** EUR {data['total']:.2f}\n"
                f"**Datum:** {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            ),
            color=0xf0a500,
            timestamp=datetime.utcnow(),
        )
        log_embed.set_footer(text="MONSTERBOT - Ticket Log")
        if LOG_CHANNEL_ID:
            try:
                log_ch = message.guild.get_channel(LOG_CHANNEL_ID)
                if log_ch:
                    excel_file = discord.File(
                        io.BytesIO(file_bytes),
                        filename=f"{message.author.name}_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx"
                    )
                    await log_ch.send(embed=log_embed, file=excel_file)
            except Exception as e:
                log.error(f"Fout bij log Excel upload: {e}")

        log.info(f"Excel ingediend door {message.author} - EUR {data['total']:.2f}")

    await bot.process_commands(message)

# ─────────────────────────────────────────────
#  LEADERBOARD
# ─────────────────────────────────────────────
async def post_leaderboard(guild, month_key):
    if not LEADERBOARD_CH_ID:
        return
    try:
        channel = guild.get_channel(LEADERBOARD_CH_ID)
        if not channel:
            return
        data = load_data()
        if month_key not in data or not data[month_key]:
            log.info(f"Geen leaderboard data voor {month_key}")
            return
        entries    = sorted(data[month_key].values(), key=lambda x: x["views"], reverse=True)[:10]
        year, month = month_key.split("-")
        month_name  = MONTHS_NL[int(month) - 1]
        embed = discord.Embed(
            title=f"Leaderboard - {month_name} {year}",
            description="Top 10 clippers van de maand.",
            color=0xf0a500,
            timestamp=datetime.utcnow(),
        )
        medals = ["1","2","3","4","5","6","7","8","9","10"]
        views_lines = []
        earn_lines  = []
        for i, entry in enumerate(entries):
            m = medals[i] if i < len(medals) else f"{i+1}."
            views_lines.append(f"#{m} **{entry['naam']}** - {entry['views']:,} views")
            earn_lines.append( f"#{m} **{entry['naam']}** - EUR {entry['earnings']:.2f}")
        embed.add_field(name="Views",       value="\n".join(views_lines) or "Geen data", inline=True)
        embed.add_field(name="Verdiensten", value="\n".join(earn_lines)  or "Geen data", inline=True)
        embed.set_footer(text="MONSTERBOT - Maandelijks Leaderboard")
        await channel.send(embed=embed)
        data.pop(month_key, None)
        save_data(data)
        log.info(f"Leaderboard gepost voor {month_key}")
    except Exception as e:
        log.error(f"post_leaderboard fout: {e}")

@tree.command(name="leaderboard", description="[Staff] Post het leaderboard van de huidige maand.")
@app_commands.guilds(discord.Object(id=GUILD_ID))
async def leaderboard_cmd(interaction):
    if not is_staff(interaction.user):
        await interaction.response.send_message("Alleen staff.", ephemeral=True)
        return
    await interaction.response.defer(ephemeral=True)
    await post_leaderboard(interaction.guild, get_month_key())
    await interaction.followup.send("Leaderboard gepost!", ephemeral=True)

@tasks.loop(hours=1)
async def check_monthly_leaderboard():
    try:
        now      = datetime.now()
        last_day = calendar.monthrange(now.year, now.month)[1]
        if now.day == last_day and now.hour == 22:
            for guild in bot.guilds:
                await post_leaderboard(guild, get_month_key())
    except Exception as e:
        log.error(f"check_monthly_leaderboard fout: {e}")

# ─────────────────────────────────────────────
#  STAFF COMMANDS
# ─────────────────────────────────────────────
@tree.command(name="bot_aan", description="[Staff] Zet de bot aan voor alle gebruikers.")
@app_commands.guilds(discord.Object(id=GUILD_ID))
async def bot_aan(interaction):
    if not is_staff(interaction.user):
        await interaction.response.send_message("Alleen staff.", ephemeral=True)
        return
    state = load_state()
    state["bot_enabled"] = True
    save_state(state)
    embed = discord.Embed(
        title="Bot ingeschakeld",
        description="Alle gebruikers kunnen nu tickets aanmaken.",
        color=0x00c9a7,
    )
    await interaction.response.send_message(embed=embed)
    log.info(f"Bot ingeschakeld door {interaction.user}")

@tree.command(name="bot_uit", description="[Staff] Zet de bot uit voor alle gebruikers.")
@app_commands.guilds(discord.Object(id=GUILD_ID))
async def bot_uit(interaction):
    if not is_staff(interaction.user):
        await interaction.response.send_message("Alleen staff.", ephemeral=True)
        return
    state = load_state()
    state["bot_enabled"] = False
    save_state(state)
    embed = discord.Embed(
        title="Bot uitgeschakeld",
        description="Alleen staff kan nu nog tickets aanmaken.",
        color=0xff4444,
    )
    await interaction.response.send_message(embed=embed)
    log.info(f"Bot uitgeschakeld door {interaction.user}")

@tree.command(name="bot_status", description="[Staff] Bekijk de huidige status van de bot.")
@app_commands.guilds(discord.Object(id=GUILD_ID))
async def bot_status(interaction):
    if not is_staff(interaction.user):
        await interaction.response.send_message("Alleen staff.", ephemeral=True)
        return
    state    = load_state()
    enabled  = state.get("bot_enabled", True)
    open_now = is_bot_open()
    now      = datetime.now()
    embed = discord.Embed(
        title="Bot Status",
        color=0x00c9a7 if open_now else 0xff4444,
    )
    embed.add_field(name="Handmatige schakelaar", value="AAN" if enabled else "UIT", inline=True)
    embed.add_field(name="Huidige dag",           value=f"Dag {now.day}", inline=True)
    embed.add_field(name="Open dagen",            value="Dag 1 t/m 4 van de maand", inline=True)
    embed.add_field(name="Status voor gebruikers", value="OPEN" if open_now else "GESLOTEN", inline=False)
    await interaction.response.send_message(embed=embed, ephemeral=True)

@tree.command(name="betaald", description="[Staff] Markeer ticket als betaald en sluit het.")
@app_commands.guilds(discord.Object(id=GUILD_ID))
@app_commands.describe(bedrag="Uitbetaald bedrag in euros")
async def betaald(interaction, bedrag: str = ""):
    if not is_staff(interaction.user):
        await interaction.response.send_message("Alleen staff.", ephemeral=True)
        return
    channel = interaction.channel
    if not channel.category or channel.category.id != TICKET_CAT_ID:
        await interaction.response.send_message("Dit is geen ticket-kanaal.", ephemeral=True)
        return
    try:
        bedrag_clean = sanitize(bedrag, 20)
        embed = discord.Embed(
            title="Betaling verwerkt",
            description=(
                f"**Bedrag:** EUR {bedrag_clean}\n"
                f"**Door:** {interaction.user.mention}\n"
                f"**Datum:** {datetime.now().strftime('%d %B %Y')}\n\n"
                "Ticket sluit in 10 seconden."
            ),
            color=0x00c9a7,
            timestamp=datetime.utcnow(),
        )
        await interaction.response.send_message(embed=embed)
        await asyncio.sleep(10)
        await channel.delete(reason=f"Betaald door {interaction.user}")
    except Exception as e:
        log.error(f"betaald command fout: {e}")

@tree.command(name="afwijzen", description="[Staff] Wijs een betalingsverzoek af met reden.")
@app_commands.guilds(discord.Object(id=GUILD_ID))
@app_commands.describe(reden="Reden voor afwijzing")
async def afwijzen(interaction, reden: str = "Geen reden opgegeven"):
    if not is_staff(interaction.user):
        await interaction.response.send_message("Alleen staff.", ephemeral=True)
        return
    channel = interaction.channel
    if not channel.category or channel.category.id != TICKET_CAT_ID:
        await interaction.response.send_message("Dit is geen ticket-kanaal.", ephemeral=True)
        return
    embed = discord.Embed(
        title="Betalingsverzoek Afgewezen",
        description=f"**Reden:** {sanitize(reden, 500)}\n**Door:** {interaction.user.mention}",
        color=0xff4444,
        timestamp=datetime.utcnow(),
    )
    await interaction.response.send_message(embed=embed)

# ─────────────────────────────────────────────
#  ERROR HANDLERS
# ─────────────────────────────────────────────
@tree.error
async def on_app_command_error(interaction, error):
    log.error(f"Slash command fout ({interaction.command.name if interaction.command else 'unknown'}): {error}")
    try:
        if not interaction.response.is_done():
            await interaction.response.send_message("Er ging iets mis. Probeer het later opnieuw.", ephemeral=True)
        else:
            await interaction.followup.send("Er ging iets mis.", ephemeral=True)
    except Exception:
        pass

@bot.event
async def on_error(event, *args, **kwargs):
    log.exception(f"Onverwachte fout in event '{event}'")

# ─────────────────────────────────────────────
#  BOT READY
# ─────────────────────────────────────────────
@bot.event
async def on_ready():
    log.info(f"Ingelogd als {bot.user} ({bot.user.id})")
    try:
        guild  = discord.Object(id=GUILD_ID)
        synced = await tree.sync(guild=guild)
        log.info(f"{len(synced)} slash commands gesynchroniseerd")
    except Exception as e:
        log.error(f"Sync mislukt: {e}")
    check_monthly_leaderboard.start()

bot.run(TOKEN, log_handler=None)
