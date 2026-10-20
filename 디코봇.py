import discord
from discord.ext import commands
from discord import app_commands
from typing import Optional
import random

# ==========================
# 봇 토큰
# ==========================
TOKEN = "MTUzMDUwMTI5NDcwNjY1NTQwNA.GRka--.WeXz_ZCy_Ev80C5WMTqnu_fh-j23laC9su1YMg"

# ==========================
# ANSI 색상
# ==========================
RESET = "\u001b[0m"

RED = "\u001b[1;31m"      # 실패
GREEN = "\u001b[1;32m"    # 성공
BLUE = "\u001b[1;34m"     # 공격
PURPLE = "\u001b[1;35m"   # 적 공격

# ==========================
# 봇 설정
# ==========================
intents = discord.Intents.all()
intents.message_content = True
intents.guilds = True
intents.members = True

bot = commands.Bot(command_prefix='/', intents=intents)


# ==========================
# 봇 실행
# ==========================
@bot.event
async def on_ready():
    print(f'Logged in as {bot.user}')
    await bot.tree.sync()
    print("슬래시 명령어 동기화 완료")


# ==========================
# /판정
# ==========================
@bot.tree.command(name="판정", description="D10을 굴려 7 이상이면 성공입니다")
@app_commands.describe(
    행동지문="행동에 대한 설명 (길게 입력 가능)"
)
async def 판정(interaction: discord.Interaction, 행동지문: str):

    roll = random.randint(1, 10)

    success = roll >= 7
    결과 = "성공" if success else "실패"

    color = GREEN if success else RED

    await interaction.response.send_message(
        f"""```ansi
{color}[{roll}/{결과}]{RESET}
{행동지문}
```"""
    )


# ==========================
# /공격
# ==========================
@bot.tree.command(
    name="공격",
    description="D30을 굴립니다 (검은 장갑 옵션 선택 시 +10)"
)
@app_commands.describe(
    행동지문="행동에 대한 설명 (길게 입력 가능)",
    검은장갑="선택하면 D30+10으로 굴립니다"
)
@app_commands.choices(
    검은장갑=[
        app_commands.Choice(name="착용", value="true")
    ]
)
async def 공격(
    interaction: discord.Interaction,
    행동지문: str,
    검은장갑: Optional[app_commands.Choice[str]] = None
):

    roll = random.randint(1, 30)

    if 검은장갑 is not None:
        total = roll + 10
        header = f"[{total} ({roll}+10)/검은장갑/공격]"
    else:
        header = f"[{roll}/공격]"

    await interaction.response.send_message(
        f"""```ansi
{BLUE}{header}{RESET}
{행동지문}
```"""
    )



# ==========================
# /적공격
# ==========================
@bot.tree.command(
    name="적공격",
    description="원하는 다이스와 보정치로 적의 공격을 굴립니다."
)
@app_commands.describe(
    다이스="예: 20이면 D20, 30이면 D30",
    보정치="추가 보정치 (없으면 0)",
    행동지문="행동에 대한 설명"
)
async def 적공격(
    interaction: discord.Interaction,
    다이스: app_commands.Range[int, 1, 1000],
    보정치: int = 0,
    행동지문: str = ""
):

    roll = random.randint(1, 다이스)
    total = roll + 보정치

    if 보정치 == 0:
        header = f"[{roll}/적공격]"
    else:
        header = f"[{total} ({roll}+{보정치})/적공격]"

    await interaction.response.send_message(
        f"""```ansi
{PURPLE}{header}{RESET}
{행동지문}
```"""
    )




# ==========================
# 봇 실행
# ==========================
bot.run(TOKEN)
