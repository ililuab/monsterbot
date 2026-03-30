import discord
from discord import app_commands
from discord.ext import commands
import openpyxl
import io
import os
import asyncio
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

TOKEN          = os.getenv("DISCORD_TOKEN")
GUILD_ID       = int(os.getenv("GUILD_ID", "0"))
TICKET_CAT_ID  = int(os.getenv("TICKET_CATEGORY_ID", "0"))
STAFF_ROLE_ID  = int(os.getenv("STAFF_ROLE_ID", "0"))
LOG_CHANNEL_ID = int(os.getenv("LOG_CHANNEL_ID", "0"))
TEMPLATE_PATH  = os.path.join(os.path.dirname(__file__), "Clipfarming_Template.xlsx")

PRICES = [
    (1_000_000, 200),
    (500_000,    50),
    (250_000,    25),
    (100_000,    12),
    (50_000,      5),
    (10_000,      1),
]

MONTHS_NL = [
    "Januari","Februari","Maart","April","Mei","Juni",
    "Juli","Augustus","September","Oktober","November","December"
]

intents = discord.Intents.default()
intents.message_content = True
intents.members = True

bot = commands.Bot(command_prefix="!", intents=intents)
tree = bot.tree


async def send_log(guild: discord.Guild, embed: discord.Embed):
    """Stuur een log-bericht naar het log-kanaal."""
    if not LOG_CHANNEL_ID:
        return
    channel = guild.get_channel(LOG_CHANNEL_ID)
    if channel:
        await channel.send(embed=embed)


def calc_earnings(views: int) -> float:
    for threshold, amount in PRICES:
        if views >= threshold:
            return float(amount)
    return 0.0


def parse_submission(file_bytes: bytes) -> dict | None:
    """Parse een ingevulde Excel en geef samenvatting terug."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception:
        return None

    results = []
    total = 0.0
    discord_naam = None
    email = None

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Zoek datarijen: rij 4 t/m 18 (1-indexed)
        for row in ws.iter_rows(min_row=4, max_row=18):
            naam_cell    = row[0].value  # A
            email_cell   = row[1].value  # B
            platform_val = row[2].value  # C
            link_val     = row[3].value  # D
            views_val    = row[4].value  # E

            if views_val and str(views_val).strip() not in ("", "Views"):
                try:
                    views = int(str(views_val).replace(".", "").replace(",", ""))
                except ValueError:
                    continue

                if not discord_naam and naam_cell and "[Vul" not in str(naam_cell):
                    discord_naam = str(naam_cell).strip()
                if not email and email_cell and "[Vul" not in str(email_cell):
                    email = str(email_cell).strip()

                earning = calc_earnings(views)
                total += earning
                results.append({
                    "sheet":    sheet_name,
                    "platform": str(platform_val or "—"),
                    "link":     str(link_val or "—"),
                    "views":    views,
                    "earning":  earning,
                })

    return {
        "discord_naam": discord_naam or "Onbekend",
        "email":        email or "Onbekend",
        "clips":        results,
        "total":        total,
    }


def build_summary_embed(data: dict, user: discord.User) -> discord.Embed:
    embed = discord.Embed(
        title="📋 Betalingsoverzicht — Ingediend",
        description=(
            f"**Discord:** {data['discord_naam']}\n"
            f"**E-mail:** {data['email']}\n"
            f"**Ingediend door:** {user.mention}\n"
            f"**Datum:** {datetime.now().strftime('%d %B %Y')}"
        ),
        color=0x7b2ff7,
        timestamp=datetime.utcnow(),
    )

    clips = data["clips"]
    if clips:
        lines = []
        for i, c in enumerate(clips[:20], 1):
            lines.append(
                f"`{i}.` **{c['platform']}** | {c['views']:,} views → **€{c['earning']:.2f}**\n"
                f"   🔗 {c['link'][:60]}{'...' if len(c['link']) > 60 else ''}"
            )
        embed.add_field(name=f"📹 Clips ({len(clips)})", value="\n".join(lines), inline=False)
    else:
        embed.add_field(name="📹 Clips", value="Geen clips gevonden in het bestand.", inline=False)

    embed.add_field(
        name="💰 Totaal te betalen",
        value=f"## €{data['total']:.2f}",
        inline=False,
    )
    embed.set_footer(text="MONSTERBOT • Wacht op goedkeuring van staff")
    return embed


def build_prijslijst_embed() -> discord.Embed:
    embed = discord.Embed(
        title="📊 Prijslijst",
        description="Views worden **geteld over 1 maand**. Alleen views binnen deze termijn tellen mee.",
        color=0x00c9a7,
    )
    lines = [
        "🔹 **10.000** views → €1",
        "🔹 **50.000** views → €5",
        "🔹 **100.000** views → €12",
        "🔹 **250.000** views → €25",
        "🔹 **500.000** views → €50",
        "🔹 **1.000.000** views → €200",
    ]
    embed.add_field(name="Tarieven", value="\n".join(lines), inline=False)
    return embed


# ─────────────────────────────────────────────
#  TICKET AANMAKEN
# ─────────────────────────────────────────────
@tree.command(
    name="uitbetaling",
    description="Maak een betalingsticket aan om je clip-verdiensten in te dienen.",
)
@app_commands.guilds(discord.Object(id=GUILD_ID))
async def uitbetaling(interaction: discord.Interaction):
    guild    = interaction.guild
    user     = interaction.user
    category = guild.get_channel(TICKET_CAT_ID)

    if not category or not isinstance(category, discord.CategoryChannel):
        await interaction.response.send_message(
            "❌ Ticket-categorie niet gevonden. Vraag een admin om de bot opnieuw in te stellen.",
            ephemeral=True,
        )
        return

    # Check: heeft deze user al een open ticket?
    existing = discord.utils.get(
        category.text_channels,
        name=f"ticket-{user.name.lower().replace(' ', '-')}"
    )
    if existing:
        await interaction.response.send_message(
            f"❌ Je hebt al een open ticket: {existing.mention}", ephemeral=True
        )
        return

    await interaction.response.defer(ephemeral=True)

    staff_role  = guild.get_role(STAFF_ROLE_ID)
    overwrites  = {
        guild.default_role: discord.PermissionOverwrite(read_messages=False),
        user:               discord.PermissionOverwrite(
                                read_messages=True,
                                send_messages=True,
                                attach_files=True,
                            ),
        guild.me:           discord.PermissionOverwrite(
                                read_messages=True,
                                send_messages=True,
                                manage_channels=True,
                            ),
    }
    if staff_role:
        overwrites[staff_role] = discord.PermissionOverwrite(
            read_messages=True, send_messages=True
        )

    channel = await guild.create_text_channel(
        name=f"ticket-{user.name.lower().replace(' ', '-')}",
        category=category,
        overwrites=overwrites,
        topic=f"Betalingsticket van {user} | {datetime.now().strftime('%d/%m/%Y')}",
    )

    # Welkomstbericht + Excel bijlage
    embed = discord.Embed(
        title="🎫 Betalingsticket aangemaakt",
        description=(
            f"Hey {user.mention}! Welkom bij jouw betalingsticket.\n\n"
            "**📌 Zo werkt het:**\n"
            "1️⃣ Download het Excel-bestand hieronder\n"
            "2️⃣ Vul je **Discord naam**, **e-mail**, **platform**, **link** en **views** in\n"
            "3️⃣ Upload het ingevulde bestand terug in dit ticket\n"
            "4️⃣ Staff beoordeelt en verwerkt de betaling\n\n"
            "⚠️ **Let op:** Views tellen alleen als ze binnen **1 maand** zijn behaald!\n"
        ),
        color=0x7b2ff7,
        timestamp=datetime.utcnow(),
    )
    embed.set_footer(text="MONSTERBOT • Ticket systeem")

    with open(TEMPLATE_PATH, "rb") as f:
        file_data = f.read()

    month_name = MONTHS_NL[datetime.now().month - 1]
    file = discord.File(
        io.BytesIO(file_data),
        filename=f"Clipfarming_Verdiensten_{month_name}_{datetime.now().year}.xlsx"
    )

    prijslijst_embed = build_prijslijst_embed()

    view = TicketCloseView(user.id)
    msg  = await channel.send(
        content=f"{user.mention}" + (f" {staff_role.mention}" if staff_role else ""),
        embeds=[embed, prijslijst_embed],
        file=file,
        view=view,
    )
    await msg.pin()

    # Log aanmaken ticket
    log_embed = discord.Embed(
        title="🎫 Ticket aangemaakt",
        description=(
            f"**Gebruiker:** {user.mention} (`{user}`)\n"
            f"**Kanaal:** {channel.mention}\n"
            f"**Datum:** {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        ),
        color=0x7b2ff7,
        timestamp=datetime.utcnow(),
    )
    log_embed.set_footer(text="MONSTERBOT • Ticket Log")
    await send_log(guild, log_embed)

    await interaction.followup.send(
        f"✅ Ticket aangemaakt: {channel.mention}", ephemeral=True
    )


# ─────────────────────────────────────────────
#  TICKET SLUITEN BUTTON
# ─────────────────────────────────────────────
class TicketCloseView(discord.ui.View):
    def __init__(self, owner_id: int):
        super().__init__(timeout=None)
        self.owner_id = owner_id

    @discord.ui.button(label="🔒 Ticket sluiten", style=discord.ButtonStyle.danger, custom_id="close_ticket")
    async def close_ticket(self, interaction: discord.Interaction, button: discord.ui.Button):
        guild      = interaction.guild
        staff_role = guild.get_role(STAFF_ROLE_ID)
        is_staff   = staff_role and staff_role in interaction.user.roles
        is_owner   = interaction.user.id == self.owner_id

        if not is_staff and not is_owner:
            await interaction.response.send_message("❌ Geen toegang.", ephemeral=True)
            return

        await interaction.response.send_message("🔒 Ticket wordt gesloten in 5 seconden...")
        await asyncio.sleep(5)
        await interaction.channel.delete(reason=f"Ticket gesloten door {interaction.user}")


# ─────────────────────────────────────────────
#  BETAALD SLASH COMMAND (staff only)
# ─────────────────────────────────────────────
@tree.command(
    name="betaald",
    description="[Staff] Markeer ticket als betaald en sluit het.",
)
@app_commands.guilds(discord.Object(id=GUILD_ID))
@app_commands.describe(bedrag="Uitbetaald bedrag in euro's (bijv. 12.50)")
async def betaald(interaction: discord.Interaction, bedrag: str = ""):
    guild      = interaction.guild
    staff_role = guild.get_role(STAFF_ROLE_ID)

    if not staff_role or staff_role not in interaction.user.roles:
        await interaction.response.send_message("❌ Alleen staff kan dit gebruiken.", ephemeral=True)
        return

    channel = interaction.channel
    if not (channel.category and channel.category.id == TICKET_CAT_ID):
        await interaction.response.send_message("❌ Dit is geen ticket-kanaal.", ephemeral=True)
        return

    embed = discord.Embed(
        title="✅ Betaling verwerkt",
        description=(
            f"**Bedrag:** €{bedrag}\n"
            f"**Verwerkt door:** {interaction.user.mention}\n"
            f"**Datum:** {datetime.now().strftime('%d %B %Y')}\n\n"
            "Dit ticket wordt in 10 seconden gesloten."
        ),
        color=0x00c9a7,
        timestamp=datetime.utcnow(),
    )
    await interaction.response.send_message(embed=embed)
    await asyncio.sleep(10)
    await channel.delete(reason=f"Betaald door {interaction.user} — €{bedrag}")


# ─────────────────────────────────────────────
#  AFWIJZEN SLASH COMMAND (staff only)
# ─────────────────────────────────────────────
@tree.command(
    name="afwijzen",
    description="[Staff] Wijs een betalingsverzoek af met een reden.",
)
@app_commands.guilds(discord.Object(id=GUILD_ID))
@app_commands.describe(reden="Reden voor afwijzing")
async def afwijzen(interaction: discord.Interaction, reden: str = "Geen reden opgegeven"):
    guild      = interaction.guild
    staff_role = guild.get_role(STAFF_ROLE_ID)

    if not staff_role or staff_role not in interaction.user.roles:
        await interaction.response.send_message("❌ Alleen staff kan dit gebruiken.", ephemeral=True)
        return

    channel = interaction.channel
    if not (channel.category and channel.category.id == TICKET_CAT_ID):
        await interaction.response.send_message("❌ Dit is geen ticket-kanaal.", ephemeral=True)
        return

    embed = discord.Embed(
        title="❌ Betalingsverzoek Afgewezen",
        description=(
            f"**Reden:** {reden}\n"
            f"**Afgewezen door:** {interaction.user.mention}\n\n"
            "Pas je Excel aan en dien opnieuw in, of neem contact op met staff."
        ),
        color=0xff4444,
        timestamp=datetime.utcnow(),
    )
    await interaction.response.send_message(embed=embed)


# ─────────────────────────────────────────────
#  AUTOMATISCH EXCEL LEZEN BIJ UPLOAD
# ─────────────────────────────────────────────
@bot.event
async def on_message(message: discord.Message):
    if message.author.bot:
        return

    # Alleen in ticket-kanalen
    if not (message.channel.category and message.channel.category.id == TICKET_CAT_ID):
        await bot.process_commands(message)
        return

    for attachment in message.attachments:
        if attachment.filename.endswith((".xlsx", ".xls")):
            try:
                file_bytes = await attachment.read()
                data = parse_submission(file_bytes)
            except Exception:
                data = None

            if not data or not data["clips"]:
                embed = discord.Embed(
                    title="⚠️ Bestand niet leesbaar",
                    description=(
                        "Het Excel-bestand kon niet worden verwerkt of bevat geen clips.\n"
                        "Zorg dat je het **meegeleverde template** gebruikt en de Views-kolom hebt ingevuld."
                    ),
                    color=0xff8c00,
                )
                await message.channel.send(embed=embed)
                continue

            summary_embed = build_summary_embed(data, message.author)
            staff_role    = message.guild.get_role(STAFF_ROLE_ID)

            view = StaffApprovalView(data["total"])
            await message.channel.send(
                content=(f"{staff_role.mention} — nieuw betalingsverzoek ter beoordeling!" if staff_role else ""),
                embed=summary_embed,
                view=view,
            )

            # Log inzending
            log_embed = discord.Embed(
                title="📥 Excel ingediend",
                description=(
                    f"**Gebruiker:** {message.author.mention} (`{message.author}`)\n"
                    f"**Discord naam:** {data['discord_naam']}\n"
                    f"**E-mail:** {data['email']}\n"
                    f"**Clips:** {len(data['clips'])}\n"
                    f"**Totaal:** €{data['total']:.2f}\n"
                    f"**Kanaal:** {message.channel.mention}\n"
                    f"**Datum:** {datetime.now().strftime('%d/%m/%Y %H:%M')}"
                ),
                color=0xf0a500,
                timestamp=datetime.utcnow(),
            )
            log_embed.set_footer(text="MONSTERBOT • Ticket Log")
            # Stuur log met het ingediende Excel bestand als bijlage
            if LOG_CHANNEL_ID:
                log_channel = message.guild.get_channel(LOG_CHANNEL_ID)
                if log_channel:
                    excel_file = discord.File(
                        io.BytesIO(file_bytes),
                        filename=f"{message.author.name}_{datetime.now().strftime('%d-%m-%Y_%H-%M')}_{attachment.filename}"
                    )
                    await log_channel.send(embed=log_embed, file=excel_file)

    await bot.process_commands(message)


# ─────────────────────────────────────────────
#  STAFF APPROVAL BUTTONS
# ─────────────────────────────────────────────
class StaffApprovalView(discord.ui.View):
    def __init__(self, total: float):
        super().__init__(timeout=None)
        self.total = total

    @discord.ui.button(label="✅ Goedkeuren & Betalen", style=discord.ButtonStyle.success, custom_id="approve_pay")
    async def approve(self, interaction: discord.Interaction, button: discord.ui.Button):
        guild      = interaction.guild
        staff_role = guild.get_role(STAFF_ROLE_ID)
        if not staff_role or staff_role not in interaction.user.roles:
            await interaction.response.send_message("❌ Alleen staff.", ephemeral=True)
            return

        embed = discord.Embed(
            title="✅ Goedgekeurd",
            description=(
                f"**Bedrag:** €{self.total:.2f}\n"
                f"**Goedgekeurd door:** {interaction.user.mention}\n"
                "Betaling wordt verwerkt. Ticket sluit in 15 seconden."
            ),
            color=0x00c9a7,
        )
        await interaction.response.send_message(embed=embed)
        self.stop()
        # Log goedkeuring
        log_embed = discord.Embed(
            title="✅ Betaling goedgekeurd",
            description=(
                f"**Goedgekeurd door:** {interaction.user.mention}\n"
                f"**Bedrag:** €{self.total:.2f}\n"
                f"**Kanaal:** {interaction.channel.mention}\n"
                f"**Datum:** {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            ),
            color=0x00c9a7,
            timestamp=datetime.utcnow(),
        )
        log_embed.set_footer(text="MONSTERBOT • Ticket Log")
        await send_log(interaction.guild, log_embed)
        await asyncio.sleep(15)
        await interaction.channel.delete(reason="Betaling goedgekeurd en verwerkt")

    @discord.ui.button(label="❌ Afwijzen", style=discord.ButtonStyle.danger, custom_id="reject_pay")
    async def reject(self, interaction: discord.Interaction, button: discord.ui.Button):
        guild      = interaction.guild
        staff_role = guild.get_role(STAFF_ROLE_ID)
        if not staff_role or staff_role not in interaction.user.roles:
            await interaction.response.send_message("❌ Alleen staff.", ephemeral=True)
            return

        await interaction.response.send_modal(RejectModal())


class RejectModal(discord.ui.Modal, title="Afwijzingsreden"):
    reden = discord.ui.TextInput(
        label="Waarom wordt dit afgewezen?",
        style=discord.TextStyle.paragraph,
        placeholder="Bijv. views niet geverifieerd, verkeerd template...",
        required=True,
    )

    async def on_submit(self, interaction: discord.Interaction):
        embed = discord.Embed(
            title="❌ Afgewezen",
            description=f"**Reden:** {self.reden.value}\n\nPas je inzending aan en probeer opnieuw.",
            color=0xff4444,
        )
        await interaction.response.send_message(embed=embed)
        # Log afwijzing
        log_embed = discord.Embed(
            title="❌ Betaling afgewezen",
            description=(
                f"**Afgewezen door:** {interaction.user.mention}\n"
                f"**Reden:** {self.reden.value}\n"
                f"**Kanaal:** {interaction.channel.mention}\n"
                f"**Datum:** {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            ),
            color=0xff4444,
            timestamp=datetime.utcnow(),
        )
        log_embed.set_footer(text="MONSTERBOT • Ticket Log")
        await send_log(interaction.guild, log_embed)


# ─────────────────────────────────────────────
#  BOT READY
# ─────────────────────────────────────────────
@bot.event
async def on_ready():
    print(f"✅ Ingelogd als {bot.user} ({bot.user.id})")
    try:
        guild  = discord.Object(id=GUILD_ID)
        synced = await tree.sync(guild=guild)
        print(f"✅ {len(synced)} slash commands gesynchroniseerd")
    except Exception as e:
        print(f"❌ Sync mislukt: {e}")


bot.run(TOKEN)

# Gemaakt door Luab
