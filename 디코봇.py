"""
디스코드 축귀 봇 (사귀 vs 축귀인) - 영문 시스템 반영판
설치: pip install discord.py openpyxl
실행: 환경변수 DISCORD_TOKEN 설정 후  python raid_bot.py
엑셀 샘플 생성: python raid_bot.py sample

엑셀(raid.xlsx) 구조
  시트 "사귀"   : 이름 | 체력 | 체력공개(Y/N) | 공격보너스(D20+N의 N) | 막수 | 조(선택)
  시트 "영문"   : 영문명 | 횟수(상시효과는 비움) | 새김가능인원
  시트 "축귀인" : 이름 | 디스코드ID | 공격강화 | 방어강화 | 영문(쉼표 구분, 모사는 '모사=강신' 형식) | 영석 | 조(1~3) | 체력(시작 체력, 비우면 120)
  시트 "강화"   : 단계(1~10) | 성공확률(%) | 필요영석   (없으면 DEFAULT_TABLE 사용)
  ※ 시트의 공격강화/방어강화/영석은 '운영진이 지급한 값'이고, /강화로 올린 단계와 쓴 영석은 state.json에 따로 저장됩니다.

행동은 3종류: 기본 공격 / 기본 방어 / 영문
사귀의 막별 기믹은 GIMMICKS 와 prepare_act() 에서 고칩니다.
"""
import asyncio
import io
import json
import os
import random
import sys
import threading
import time
import urllib.request
from collections import defaultdict
from dataclasses import dataclass, field

import discord
from discord import app_commands
from openpyxl import Workbook, load_workbook

EXCEL_PATH = "raid.xlsx"
ACTION_WINDOW = 600        # 행동 가능 시간(초) = 10분
BASE_HP = 120
BONUS_PER_ENHANCE = 2      # 강화 1당 +2
MAX_SKILLS_PER_PLAYER = 2   # 한 사람이 새길 수 있는 영문 수
MAX_ENHANCE_TOTAL = 10     # 공격강화+방어강화 합계 상한 (확인 필요)
JINGJO_HP = 180            # 징조: 최대 체력

# 상시효과 영문 (행동으로 쓰지 않고 자동 적용)
PASSIVES = {"징조", "불굴", "인내", "집념", "연사", "모사"}
# 행동형 영문: 필요한 대상 수 / 값 입력 필요 여부
ACTIVE = {
    "각광": (2, False), "항거": (0, False), "사활": (0, False), "보호": (2, False),
    "중독": (0, False), "복수": (0, False), "혼란": (0, False), "강신": (0, False),
    "천도": (0, False), "행운": (0, False), "역살": (0, True), "억제": (0, False),
    "기도": (1, False), "수복": (0, False), "선포": (0, False), "분배": (1, True),
    "정화": (1, False), "역행": (1, False), "예지": (0, False),
}


def d(n): return random.randint(1, n)
def nd(k, n): return sum(d(n) for _ in range(k))


# ---------- 데이터 ----------
@dataclass
class Boss:
    name: str
    hp: int
    max_hp: int
    hp_public: bool
    atk_bonus: int
    acts: int
    poisons: list = field(default_factory=list)  # [값, 남은 틱]


@dataclass
class Player:
    name: str
    uid: int
    atk_up: int
    def_up: int
    skills: dict                 # 영문명 -> 남은 횟수 (상시효과는 -1)
    hp: int = BASE_HP
    max_hp: int = BASE_HP
    poison_left: int = 0         # 독 남은 턴 (사귀 3막 기믹)
    invul_until: int = 0         # 이 막까지 피해 무효 (사활)
    last_taken: int = 0          # 직전 막에 받은 피해 (복수)
    status_hit: bool = False     # 상태이상 피해를 입었는가 (혼란) - 외부에서 True로 설정
    curses: list = field(default_factory=list)
    hanggeo_dmg: int = 0         # 다음 막에 되돌려줄 피해 (항거)

    @property
    def alive(self): return self.hp > 0
    def has(self, s): return s in self.skills


@dataclass
class Action:
    uid: int
    kind: str                    # 공격 / 방어 / 영문
    skill: str = ""
    targets: list = field(default_factory=list)
    value: int = 0
    name: str = ""
    roll: int = 0                # /공격·/방어를 누른 순간 굴린 값
    det: str = ""               # 굴림 표시 (13+6)


@dataclass
class Ctx:                       # 한 막 처리 중 임시 상태
    dmg: int = 0
    reduce: int = 0
    defended: dict = field(default_factory=dict)
    shield: dict = field(default_factory=lambda: defaultdict(int))
    redirect: int | None = None
    spot: list = field(default_factory=list)


@dataclass
class Raid:
    boss: Boss
    players: dict
    channel: discord.abc.Messageable
    act: int = 0
    actions: dict = field(default_factory=dict)
    plan: dict = field(default_factory=dict)
    open: bool = False
    all_done: asyncio.Event = field(default_factory=asyncio.Event)
    task: asyncio.Task | None = None
    gimmick: dict = field(default_factory=dict)   # 이번 막 기믹 정보
    choices: dict = field(default_factory=dict)   # uid -> 왼쪽/오른쪽 (4막)
    deadline: float = 0.0          # 이번 막 행동 마감 시각 (time.monotonic 기준)


# ---------- 굴림 ----------
def _fmt(label, base, bonus):
    """주사위 표기 없이 숫자만. 보정이 있으면 13+6, 없으면 13"""
    return f"{base}" + (f"+{bonus}" if bonus else "")


def _join(parts):
    return " & ".join(parts)


def shown(det, v):
    """'13+6 = 19' 형태. 계산식이 곧 결과면(보정 없음) 숫자만"""
    return det if det == str(v) else f"{det} = {v}"


def atk_roll(p: Player, max_dice=False):
    """기본 공격: D20 (집념 D10+10), 연사면 2회, 강화 보정 +2/단계"""
    bonus = p.atk_up * BONUS_PER_ENHANCE
    total, parts = 0, []
    for _ in range(2 if p.has("연사") else 1):
        if p.has("집념"):
            base, label = (10 if max_dice else d(10)) + 10, "D10+10"
        else:
            base, label = (20 if max_dice else d(20)), "D20"
        total += base + bonus
        parts.append(_fmt(label, base, bonus))
    return total, _join(parts)


def def_roll(p: Player):
    """기본 방어: D20 (불굴 D10+10), 인내면 2회, 강화 보정 +2/단계"""
    bonus = p.def_up * BONUS_PER_ENHANCE
    total, parts = 0, []
    for _ in range(2 if p.has("인내") else 1):
        if p.has("불굴"):
            base, label = d(10) + 10, "D10+10"
        else:
            base, label = d(20), "D20"
        total += base + bonus
        parts.append(_fmt(label, base, bonus))
    return total, _join(parts)


def hp_str(q: Player) -> str:
    return f"{q.hp}/{q.max_hp}"


def heal(t: Player, amount: int) -> int:
    if not t.alive:
        return 0
    before = t.hp
    t.hp = min(t.max_hp, t.hp + amount)
    return t.hp - before


# ---------- 사귀 공격 (★ 미정 - 여기만 고치면 됩니다) ----------
# ===== 사귀 기믹 설정 (막별) =====
BASIC_TEXT = "10+D20"            # 기본공격(광역공격): 막 시작 때 공개
SMASH_TEXT = "20+D20"            # 강타: 기본은 비공개(???)
SMASH_REVEAL_ACTS = set()        # 강타 값을 공개할 막 번호를 여기에 넣으면 전조에 표시 (예: {3})
SHARE_LIMIT = 100                # 쉐어 데미지: 이 값 미만으로 때리면 발동
POISON_TURNS, POISON_DIE = 3, 15
# 막 번호 -> 기믹 종류. 적지 않은 막은 'aoe'(기본공격 광역)
GIMMICKS = {1: "smash", 2: "share", 3: "poison", 4: "side"}


def josa(word, with_final, without_final):
    """받침 유무에 따라 이/가, 을/를 등을 고름"""
    ch = word[-1]
    if "가" <= ch <= "힣" and (ord(ch) - 0xAC00) % 28:
        return with_final
    return without_final


def basic_roll(): return 10 + d(20)
def smash_roll(): return 20 + d(20)


def prepare_act(r: Raid) -> str:
    """전조 단계: 이번 막의 기믹을 정하고 전조 설명문을 반환. r.gimmick / r.plan 채움"""
    b = r.boss
    alive = [p for p in r.players.values() if p.alive]
    kind = GIMMICKS.get(r.act, "aoe")
    r.plan, r.choices = {}, {}
    if kind == "smash":
        t = random.choice(alive); v = smash_roll()
        r.plan = {t.uid: v}
        r.gimmick = {"kind": kind, "preview": f"[{t.name}]에게 강타 {v}"}
        return f"{b.name}{josa(b.name, '이', '가')} **[{t.name}]**을(를) 바라본다."
    if kind == "share":
        total = 50 + d(50)
        r.gimmick = {"kind": kind, "total": total,
                     "preview": f"쉐어 데미지 총 {total} (사귀에게 {SHARE_LIMIT} 이상 피해를 주면 회피)"}
        return (f"{b.name}{josa(b.name, '이', '가')} 축귀인들을 가늠한다.\n"
                f"이번 막에 **사귀에게 합계 {SHARE_LIMIT} 이상** 피해를 주지 못하면 "
                f"**쉐어 데미지(50+D50)**를 생존자 수로 나눠 맞는다.")
    if kind == "poison":
        t = random.choice(alive)
        r.gimmick = {"kind": kind, "target": t.uid,
                     "preview": f"[{t.name}]에게 독 (매 막 D{POISON_DIE}, {POISON_TURNS}턴)"}
        return (f"{b.name}{josa(b.name, '이', '가')} **[{t.name}]**을(를) 바라본다.\n"
                f"독이 퍼진다. {POISON_TURNS}턴 동안 매 막 D{POISON_DIE} 지속 피해.")
    if kind == "side":
        v = basic_roll()
        r.gimmick = {"kind": kind, "value": v, "preview": f"왼쪽을 고른 사람에게 광역 {v}"}
        return (f"{b.name}{josa(b.name, '이', '가')} 양쪽을 번갈아 노려본다.\n"
                f"**왼쪽**과 **오른쪽** 중 하나를 `/선택`하라. **왼쪽**을 고른 사람만 광역 피해를 입는다.")
    v = basic_roll()
    r.gimmick = {"kind": "aoe", "value": v, "preview": f"전원에게 광역 {v}"}
    return f"{b.name}{josa(b.name, '이', '가')} 광역 공격을 준비한다."


# ---------- 영문 처리 ----------
def do_action(r: Raid, c: Ctx, a: Action) -> str:
    p = r.players[a.uid]
    b = r.boss
    if a.kind == "공격":
        v, det = (a.roll, a.det) if a.det else atk_roll(p); c.dmg += v
        return f"- {p.name}: 기본 공격 : {shown(det, v)}"
    if a.kind == "방어":
        v, det = (a.roll, a.det) if a.det else def_roll(p); c.defended[p.uid] = v
        return f"- {p.name}: 기본 방어 : {shown(det, v)}"

    s = a.skill
    tg = [r.players[u] for u in a.targets]
    head = f"- {p.name}: [{s}]"
    if s == "각광":
        c.spot.append((p.uid, a.targets))
        return f"{head} {tg[0].name}, {tg[1].name}의 피해를 절반으로 줄여 대신 받는다"
    if s == "항거":
        c.redirect = p.uid
        return f"{head} 이번 막의 모든 피해를 대신 받는다"
    if s == "사활":
        p.invul_until = r.act + 1
        return f"{head} {r.act}~{r.act + 1}막 동안 피해를 입지 않는다"
    if s == "보호":
        v, det = def_roll(p)
        for t in tg: c.shield[t.uid] += v
        return f"{head} 방어값 {shown(det, v)} → {tg[0].name}, {tg[1].name}에게 보호막 {v}"
    if s == "중독":
        v = d(10) + 5; c.dmg += v; b.poisons.append([v, 2])
        return f"{head} {v} 피해 (이후 2막 동안 매 막 {v})"
    if s == "복수":
        extra = p.last_taken // 2; v, det = atk_roll(p); c.dmg += v + extra
        return f"{head} 기본 공격 : {shown(det, v)} + 복수 {extra} = {v + extra}"
    if s == "혼란":
        v, det = atk_roll(p, max_dice=True); c.dmg += v; p.status_hit = False
        return f"{head} 기본 공격 최대치 : {shown(det, v)}"
    if s == "강신":
        v = 30 + d(30); c.dmg += v; return f"{head} 30+D30 → {v}"
    if s == "천도":
        v = nd(2, 25); c.dmg += v; return f"{head} 2D25 → {v}"
    if s == "행운":
        x = (d(10), d(15), d(20)); v = sum(x); c.dmg += v
        return f"{head} D10+D15+D20 {x} → {v}"
    if s == "역살":
        p.hp -= a.value; v, det = atk_roll(p); c.dmg += v + a.value
        return f"{head} 체력 -{a.value}, 기본 공격 : {shown(det, v)} + {a.value} = {v + a.value}"
    if s == "억제":
        v = d(30); c.dmg += v; c.reduce += v
        return f"{head} D30 → {v} 피해, 사귀 공격값 -{v}"
    if s == "기도":
        n = heal(tg[0], 10 + d(15))
        return f"{head} {tg[0].name} 체력 +{n} ({hp_str(tg[0])})"
    if s == "수복":
        r15 = d(15); v = 10 + r15; c.dmg += v; n = heal(p, r15)
        return f"{head} 10+D15 → {v} 피해, 본인 체력 +{n} ({hp_str(p)})"
    if s == "선포":
        amt = 5 + d(10)
        got = [f"{q.name} +{heal(q, amt)} ({hp_str(q)})" for q in r.players.values() if q.alive]
        return f"{head} 전원 체력 +{amt} ({', '.join(got)})"
    if s == "분배":
        p.hp -= a.value; n = heal(tg[0], a.value)
        return f"{head} 체력 -{a.value} ({hp_str(p)}), {tg[0].name} 체력 +{n} ({hp_str(tg[0])})"
    if s == "정화":
        t = tg[0]
        if not t.curses:
            return f"{head} {t.name}에게 해제할 저주가 없다"
        cur = a.name if a.name in t.curses else t.curses[0]
        t.curses.remove(cur)
        return f"{head} {t.name}의 저주 '{cur}' 해제"
    if s == "역행":
        return f"{head} {tg[0].name}의 [{a.name}] 사용 횟수 +1"
    if s == "예지":
        return ""          # 예지는 사용 즉시 채널에 알리므로 막 결과에는 넣지 않음
    return f"{head} (미구현)"


def resolve_act(r: Raid):
    """행동 → 사귀 피해 → 사귀 공격 순으로 처리. (행동 로그, 공격 결과 로그, 총 피해, 사귀 쓰러짐 여부) 반환"""
    b = r.boss
    c = Ctx()
    log = []
    for dot in b.poisons[:]:
        c.dmg += dot[0]; dot[1] -= 1
        log.append(f"- 중독: 사귀에게 지속 피해 {dot[0]}")
        if dot[1] <= 0: b.poisons.remove(dot)
    for p in r.players.values():
        if p.hanggeo_dmg:
            c.dmg += p.hanggeo_dmg
            log.append(f"- {p.name}: [항거] 되돌려주는 피해 {p.hanggeo_dmg}")
            p.hanggeo_dmg = 0
    for a in r.actions.values():
        line = do_action(r, c, a)
        if line:
            log.append(line)
    if not r.actions:
        log.append("- 아무도 행동하지 않았다.")

    b.hp = max(0, b.hp - c.dmg)
    if b.hp <= 0:
        return log, [], c.dmg, True

    # ---- 사귀 공격 ----
    res = []
    kind = r.gimmick.get("kind", "aoe")
    alive_now = [q for q in r.players.values() if q.alive]
    base = {}
    if kind == "smash":
        base = dict(r.plan)
    elif kind == "share":
        if c.dmg >= SHARE_LIMIT:
            res.append(f"- 쉐어 데미지 회피! (사귀에게 {c.dmg} 피해, 기준 {SHARE_LIMIT})")
        else:
            total = r.gimmick["total"]
            share = total // max(1, len(alive_now))
            base = {q.uid: share for q in alive_now}
            res.append(f"- 쉐어 데미지 발동! (사귀에게 {c.dmg} 피해 < {SHARE_LIMIT}) "
                       f"총 {total}을(를) {len(alive_now)}명이 {share}씩 나눠 맞는다")
    elif kind == "side":
        left = []
        for q in alive_now:
            side = r.choices.get(q.uid)
            if side is None:
                side = random.choice(["왼쪽", "오른쪽"])
                res.append(f"- {q.name}: 선택하지 않아 {side}으로 정해졌다")
            if side == "왼쪽": left.append(q)
        res.append("- 왼쪽: " + (", ".join(q.name for q in left) or "없음")
                   + " / 오른쪽: " + (", ".join(q.name for q in alive_now if q not in left) or "없음"))
        base = {q.uid: r.gimmick["value"] for q in left}
    elif kind == "aoe":
        base = {q.uid: r.gimmick["value"] for q in alive_now}
    plan = {u: max(0, v - c.reduce) for u, v in base.items()}
    if c.reduce and base:
        res.append(f"- 억제로 사귀 공격값 -{c.reduce}")
    inc = dict(plan)
    if c.redirect is not None and inc:              # 항거: 전부 대신 받음
        inc = {c.redirect: sum(inc.values())}
    for sid, tids in c.spot:                        # 각광: 절반으로 줄여 대신 받음
        for t in tids:
            if t in inc and t != sid:
                inc[sid] = inc.get(sid, 0) + inc.pop(t) // 2
    for p in r.players.values():
        p.last_taken = 0
    for uid, dmg in inc.items():
        p = r.players[uid]
        if not p.alive: continue
        notes = []
        if p.invul_until >= r.act:
            dmg, notes = 0, ["피해 무효"]
        else:
            if uid in c.defended:
                dd = min(dmg, c.defended[uid]); dmg -= dd; notes.append(f"방어 -{dd}")
            sh = min(c.shield.get(uid, 0), dmg)
            if sh: dmg -= sh; notes.append(f"보호막 -{sh}")
        p.hp = max(0, p.hp - dmg)
        p.last_taken = dmg
        if uid == c.redirect: p.hanggeo_dmg = dmg
        extra = f" ({', '.join(notes)})" if notes else ""
        res.append(f"- {b.name}의 강한 공격으로 **{p.name} 체력 -{dmg}** → {hp_str(p)}{extra}")
        if not p.alive:
            res.append(f"  {p.name}은(는) 행동 불능이 되었다.")

    # ---- 독 (3막 기믹): 방어와 상관없이 들어가는 지속 피해, 정화로 해제 가능 ----
    for q in r.players.values():
        q.status_hit = False
    if kind == "poison":
        t = r.players[r.gimmick["target"]]
        if t.alive and "독" not in t.curses:
            t.curses.append("독"); t.poison_left = POISON_TURNS
    for q in r.players.values():
        if q.alive and "독" in q.curses and q.poison_left > 0:
            dd = 0 if q.invul_until >= r.act else d(POISON_DIE)
            q.hp = max(0, q.hp - dd); q.poison_left -= 1
            if dd: q.status_hit = True
            if q.poison_left <= 0: q.curses.remove("독")
            res.append(f"- 독: **{q.name} 체력 -{dd}**" + (" (피해 무효)" if not dd else "") + f" → {hp_str(q)} (남은 {q.poison_left}턴)")
            if not q.alive:
                res.append(f"  {q.name}은(는) 행동 불능이 되었다.")
    return log, res, c.dmg, False


# ---------- 진행 ----------
raids: dict[int, Raid] = {}   # 채널ID -> 그 채널에서 진행 중인(또는 끝난) 축귀


def running(r):
    return bool(r and r.task and not r.task.done())


def maybe_done(r: Raid):
    """생존자 전원이 행동(4막은 선택까지) 마쳤으면 10분을 기다리지 않고 진행"""
    alive = [u for u, q in r.players.items() if q.alive]
    if all(u in r.actions for u in alive) and (
            r.gimmick.get("kind") != "side" or all(u in r.choices for u in alive)):
        r.all_done.set()


def hp_text(r: Raid):
    return f"{r.boss.hp}/{r.boss.max_hp}" if r.boss.hp_public else "???"


def remain_text(r: Raid) -> str:
    sec = max(0, int(r.deadline - time.monotonic()))
    return f"{sec // 60}분 {sec % 60}초"


COLOR_INTRO, COLOR_ACT, COLOR_HIT, COLOR_WIN, COLOR_END = 0x8E44AD, 0x3498DB, 0xE74C3C, 0xF1C40F, 0x95A5A6


def clip(text, limit=4000):
    return text if len(text) <= limit else text[:limit - 3] + "..."


def hp_bar(hp, mx, n=10):
    filled = round(n * hp / mx) if mx else 0
    return "🟥" * filled + "⬛" * (n - filled)


async def send_long(ch, text, limit=1900):
    """디스코드 메시지 2000자 제한 때문에 줄 단위로 나눠 보냄 (인원이 많아도 안전)"""
    buf = ""
    for line in text.split("\n"):
        if buf and len(buf) + len(line) + 1 > limit:
            await ch.send(buf)
            buf = ""
        buf += line + "\n"
    if buf.strip():
        await ch.send(buf)


async def run_raid(r: Raid):
    b = r.boss
    while r.act < b.acts and b.hp > 0 and any(p.alive for p in r.players.values()):
        r.act += 1
        r.actions.clear(); r.all_done.clear()
        desc = prepare_act(r)
        r.open = True
        r.deadline = time.monotonic() + ACTION_WINDOW
        intro = discord.Embed(
            title=f"[{r.act}막] {b.name}{josa(b.name, '이', '가')} 등장한다",
            description=desc,
            color=COLOR_INTRO)
        intro.add_field(name="사귀 체력", value=hp_text(r), inline=True)
        intro.add_field(name="기본공격 (광역)", value=BASIC_TEXT, inline=True)
        intro.add_field(name="강타", value=SMASH_TEXT if r.act in SMASH_REVEAL_ACTS else "???", inline=True)
        intro.add_field(name="행동 가능 시간", value="10분", inline=True)
        intro.set_footer(text="/공격  /방어  /영문명" + ("  /선택" if r.gimmick.get("kind") == "side" else ""))
        await r.channel.send(embed=intro)
        try:
            await asyncio.wait_for(r.all_done.wait(), timeout=ACTION_WINDOW)
        except asyncio.TimeoutError:
            pass
        r.open = False

        acts, res, total, boss_down = resolve_act(r)
        em = discord.Embed(title=f"[{r.act}막] 축귀인의 행동",
                           description=clip("\n".join(acts) or "아무도 행동하지 않았다."), color=COLOR_ACT)
        em.add_field(name="사귀 피해값", value=f"**-{total}**", inline=True)
        if b.hp_public:
            em.add_field(name="사귀 체력", value=f"{b.hp}/{b.max_hp}", inline=True)
        await r.channel.send(embed=em)
        if boss_down:
            await r.channel.send(embed=discord.Embed(
                title=f"{b.name} 쓰러짐!", description="축귀가 조기 종료되었습니다.", color=COLOR_WIN))
            return
        hit = discord.Embed(title=f"[{r.act}막] 공격 결과",
                            description=clip("\n".join(res) or "아무도 피해를 입지 않았다."), color=COLOR_HIT)
        hit.add_field(
            name="현재 체력",
            value=clip("\n".join(
                f"{'💀' if not p.alive else '❤️'} **{p.name}** {p.hp}/{p.max_hp}" + (" (행동 불능)" if not p.alive else "")
                for p in r.players.values()), 1024),
            inline=False)
        await r.channel.send(embed=hit)
    await r.channel.send(embed=discord.Embed(title="축귀 종료", color=COLOR_END))


# ---------- 엑셀 / 구글 시트 ----------
SHEET_ID = os.environ.get("SHEET_ID")   # 설정하면 raid.xlsx 대신 구글 시트를 읽음


def fetch_workbook():
    if SHEET_ID:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
        with urllib.request.urlopen(url, timeout=20) as resp:
            return load_workbook(io.BytesIO(resp.read()), data_only=True)
    return load_workbook(EXCEL_PATH, data_only=True)


# ---------- 강화 (영석) ----------
STATE_PATH = "state.json"   # 봇이 올린 강화 단계와 사용한 영석을 저장 (시트는 건드리지 않음)
state_lock = threading.Lock()
# 시트에 "강화" 탭이 없을 때 쓰는 기본값 (임시값 - 시트에서 직접 정하세요)  단계: (성공확률%, 필요영석)
DEFAULT_TABLE = {1: (95, 1), 2: (90, 1), 3: (80, 2), 4: (70, 2), 5: (60, 3),
                 6: (50, 4), 7: (40, 5), 8: (30, 6), 9: (20, 8), 10: (10, 10)}


def load_state():
    try:
        with open(STATE_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_state():
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=1)


state = load_state()   # {디스코드ID: {"atk": 올린공격강화, "def": 올린방어강화, "spent": 쓴영석}}

# 구글 시트에 직접 기록하는 방식 (Apps Script 웹앱 주소). 설정하면 시트가 '현재 값'의 기준이 되고 state.json은 쓰지 않음
SHEET_WEBHOOK = os.environ.get("SHEET_WEBHOOK")
WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "")


def bonus(uid):
    return {} if SHEET_WEBHOOK else state.get(str(uid), {})


def post_enhance(payload):
    """Apps Script에 강화 결과를 보내 시트에 기록하게 함. 성공하면 {'atk','def','stones'} 반환"""
    payload["secret"] = WEBHOOK_SECRET
    req = urllib.request.Request(SHEET_WEBHOOK, data=json.dumps(payload).encode("utf-8"),
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        res = json.loads(resp.read().decode("utf-8"))
    if not res.get("ok"):
        raise RuntimeError(f"시트 기록 실패: {res.get('error', '알 수 없는 오류')}")
    return res


def cell(row, i):
    return row[i] if len(row) > i else None


def read_enhance_table(wb):
    if "강화" in wb.sheetnames:
        t = {}
        for x in wb["강화"].iter_rows(min_row=2, values_only=True):
            if x[0] not in (None, ""):
                t[int(x[0])] = (float(x[1]), int(x[2]))
        if t:
            return t
    return DEFAULT_TABLE


def read_player(uid):
    """시트에서 본인 줄을 찾아 현재 강화/영석 상태를 계산 (시트 값 + 봇이 올린 값)"""
    wb = fetch_workbook()
    for x in wb["축귀인"].iter_rows(min_row=2, values_only=True):
        if x[0] and x[1] and int(x[1]) == uid:
            st = bonus(uid)
            return {
                "name": x[0],
                "atk": int(cell(x, 2) or 0) + st.get("atk", 0),
                "def": int(cell(x, 3) or 0) + st.get("def", 0),
                "stones": int(cell(x, 5) or 0) - st.get("spent", 0),
                "table": read_enhance_table(wb),
            }
    return None


def try_enhance(uid, kind):
    """kind: 'atk' 또는 'def'. 성공하면 영석 소모+강화, 실패하면 아무것도 변하지 않음"""
    with state_lock:
        info = read_player(uid)
        if not info:
            return "이번 명단에 없는 사용자입니다. 운영진에게 문의하세요."
        level = info["atk"] + info["def"]
        if level >= MAX_ENHANCE_TOTAL:
            return f"이미 최대 강화({MAX_ENHANCE_TOTAL}강)입니다."
        nxt = level + 1
        if nxt not in info["table"]:
            return f"강화 시트에 {nxt}강 정보가 없습니다."
        prob, cost = info["table"][nxt]
        if info["stones"] < cost:
            return f"영석이 부족합니다. ({nxt}강 필요 {cost}개 / 보유 {info['stones']}개)"
        label = "기본 공격" if kind == "atk" else "기본 방어"
        success = random.random() * 100 < prob
        if SHEET_WEBHOOK:                         # 시트에 직접 기록 (실패하면 아무것도 변하지 않음)
            res = post_enhance({
                "uid": str(uid), "name": info["name"], "kind": label,
                "result": "성공" if success else "실패", "cost": cost,
                "atk_add": 1 if success and kind == "atk" else 0,
                "def_add": 1 if success and kind == "def" else 0,
                "max_total": MAX_ENHANCE_TOTAL})
            left = res["stones"]
        else:                                     # state.json에 기록
            st = state.setdefault(str(uid), {"atk": 0, "def": 0, "spent": 0})
            st["spent"] += cost                   # 성공/실패와 관계없이 영석 소모
            if success:
                st[kind] += 1
            save_state()
            left = info["stones"] - cost
        if success:
            return (f"✨ **{info['name']}**: {label} **{nxt}강 성공!** "
                    f"(영석 -{cost}, 남은 영석 {left}개)")
        return (f"💥 **{info['name']}**: {label} {nxt}강 실패... "
                f"영석 {cost}개가 소모되었습니다. (남은 영석 {left}개)")


def load_raid(path, channel, group=None):
    """group을 주면 축귀인 시트의 '조' 열(7번째 칸)이 같은 사람만 참가.
    사귀 시트에도 '조' 열(6번째 칸)이 있으면 그 조의 줄을, 없으면 첫 줄을 사용."""
    wb = fetch_workbook()
    boss_rows = [r for r in wb["사귀"].iter_rows(min_row=2, values_only=True) if r[0]]
    b = next((r for r in boss_rows if group is not None and cell(r, 5) not in (None, "")
              and int(cell(r, 5)) == group), boss_rows[0])
    boss = Boss(b[0], int(b[1]), int(b[1]), str(b[2]).upper() == "Y", int(b[3] or 0), int(b[4] or 8))
    defs = {}
    for x in wb["영문"].iter_rows(min_row=2, values_only=True):
        if x[0]:
            defs[x[0]] = (int(x[1]) if x[1] not in (None, "") else -1, int(x[2] or 1))
    warns, holders, players = [], defaultdict(list), {}
    for x in wb["축귀인"].iter_rows(min_row=2, values_only=True):
        if not x[0]: continue
        in_group = group is None or int(cell(x, 6) or 0) == group   # 새김 인원은 전체 기준, 참가는 해당 조만
        name, uid = x[0], int(x[1])
        au, du = int(x[2] or 0), int(x[3] or 0)
        st = bonus(uid)                       # /강화로 올린 단계를 더함 (시트 직접 기록 방식이면 없음)
        au += st.get("atk", 0); du += st.get("def", 0)
        if au + du > MAX_ENHANCE_TOTAL:
            warns.append(f"{name}: 강화 합계 {au + du} > {MAX_ENHANCE_TOTAL}")
        skills = {}
        raw = str(x[4] or "").replace("，", ",").replace("、", ",")   # 전각 쉼표도 인식
        entries = [t for t in raw.split(",") if t.strip()]
        if len(entries) > MAX_SKILLS_PER_PLAYER:
            warns.append(f"{name}: 영문 {len(entries)}개 (최대 {MAX_SKILLS_PER_PLAYER}개)")
        for s in raw.split(","):
            s = s.strip()
            if not s: continue
            if s.startswith("모사="):
                src = s[3:]
                if src not in defs:
                    warns.append(f"{name}: 모사 대상 '{src}' 없음"); continue
                skills["모사"] = -1; skills[src] = defs[src][0]
                holders["모사"].append(name)
            elif s in defs:
                skills[s] = defs[s][0]; holders[s].append(name)
            else:
                warns.append(f"{name}: 알 수 없는 영문 '{s}'")
        if not in_group:
            continue
        p = Player(name, uid, au, du, skills)
        start_hp = int(cell(x, 7) or BASE_HP)          # 시트 '체력' 열 = 시작 체력 = 전체 체력
        p.hp = p.max_hp = start_hp
        if "징조" in skills: p.max_hp = max(start_hp, JINGJO_HP)   # 징조: 최대 체력만 늘고 현재 체력은 그대로
        players[uid] = p
    for s, names in holders.items():
        if s in defs and len(names) > defs[s][1]:
            warns.append(f"[{s}] 새김 인원 초과: {', '.join(names)} (최대 {defs[s][1]}인)")
    if not players:
        raise ValueError("참가 인원이 없습니다. 축귀인 시트의 '조' 열을 확인하세요.")
    return Raid(boss, players, channel), warns


def make_sample():
    wb = Workbook()
    ws = wb.active; ws.title = "사귀"
    ws.append(["이름", "체력", "체력공개", "공격보너스", "막수", "조"])
    ws.append(["사귀", 200, "N", 0, 4, 1])
    ws.append(["사귀", 200, "N", 0, 4, 2])
    ws.append(["사귀", 200, "N", 0, 4, 3])
    ws = wb.create_sheet("영문")
    ws.append(["영문명", "횟수", "새김가능인원"])
    for row in [("각광", 2, 2), ("항거", 1, 2), ("사활", 1, 2), ("보호", 2, 2), ("징조", None, 1),
                ("불굴", None, 1), ("인내", None, 1), ("중독", 1, 2), ("복수", 2, 2), ("혼란", 2, 2),
                ("강신", 1, 2), ("천도", 2, 2), ("행운", 3, 2), ("역살", 1, 2), ("억제", 2, 2),
                ("집념", None, 1), ("연사", None, 1), ("기도", 3, 3), ("수복", 1, 2), ("선포", 2, 2),
                ("분배", 1, 2), ("정화", 2, 2), ("역행", 1, 1), ("예지", 1, 1), ("모사", None, 1)]:
        ws.append(row)
    ws = wb.create_sheet("축귀인")
    ws.append(["이름", "디스코드ID", "공격강화", "방어강화", "영문", "영석", "조", "체력"])
    ws.append(["단", 123456789012345678, 0, 0, "강신,기도", 10, 1, 120])
    ws = wb.create_sheet("강화")
    ws.append(["단계", "성공확률(%)", "필요영석"])
    for lv, (pr, co) in DEFAULT_TABLE.items():
        ws.append([lv, pr, co])
    wb.save(EXCEL_PATH)
    print("raid.xlsx 샘플 생성 완료")


# ---------- 디스코드 ----------
intents = discord.Intents.default()
client = discord.Client(intents=intents)
tree = app_commands.CommandTree(client)


@client.event
async def on_ready():
    await tree.sync()
    print("ready")


축귀 = app_commands.Group(name="축귀", description="축귀 진행 (운영진용)",
                          default_permissions=discord.Permissions(manage_guild=True))
tree.add_command(축귀)


@축귀.command(name="시작", description="시트를 읽어 이 채널에서 축귀를 시작합니다")
@app_commands.describe(조="참가할 조 번호 (시트의 '조' 열). 조를 나누지 않으면 비워두세요")
async def start(i: discord.Interaction, 조: int = None):
    if running(raids.get(i.channel_id)):
        return await i.response.send_message("이 채널에서 이미 축귀가 진행 중입니다.", ephemeral=True)
    await i.response.defer()    # 시트를 읽는 동안 디스코드 3초 제한에 걸리지 않게 먼저 응답
    try:
        new, warns = await asyncio.to_thread(load_raid, EXCEL_PATH, i.channel, 조)
    except Exception as e:
        return await i.followup.send(f"엑셀 로드 실패: {e}")
    raids[i.channel_id] = raid = new
    msg = "축귀를 시작합니다." if 조 is None else f"{조}조 축귀를 시작합니다."
    if warns:
        msg += "\n⚠ 엑셀 확인 필요\n" + "\n".join(f"- {w}" for w in warns)
    await i.followup.send(msg)
    raid.task = asyncio.create_task(run_raid(raid))


@축귀.command(name="막종료", description="이번 막의 행동 시간을 지금 끝내고 결과를 진행합니다")
async def end_act(i: discord.Interaction):
    r = raids.get(i.channel_id)
    if not running(r) or not r.open:
        return await i.response.send_message("지금은 행동 시간이 아닙니다.", ephemeral=True)
    waiting = [q.name for u, q in r.players.items() if q.alive and u not in r.actions]
    r.all_done.set()
    msg = f"{r.act}막을 지금 종료합니다."
    if waiting:
        msg += f"\n행동하지 않은 사람: {', '.join(waiting)}"
    await i.response.send_message(msg)


@축귀.command(name="종료", description="이 채널에서 진행 중인 축귀를 종료합니다")
async def stop(i: discord.Interaction):
    raid = raids.get(i.channel_id)
    if running(raid):
        raid.task.cancel()
        return await i.response.send_message("축귀를 종료했습니다.")
    await i.response.send_message("진행 중인 축귀가 없습니다.", ephemeral=True)


async def submit(i: discord.Interaction, kind, skill="", members=(), value=None, name=None):
    """모든 행동 명령어가 공통으로 거치는 처리 (kind: 공격 / 방어 / 영문)"""
    async def no(msg): await i.response.send_message(msg, ephemeral=True)

    raid = raids.get(i.channel_id)
    if not raid or not raid.open:
        return await no("이 채널에서는 지금 행동할 수 없습니다.")
    p = raid.players.get(i.user.id)
    if not p: return await no("이번 축귀 참가자가 아닙니다.")
    if kind == "영문" and skill not in p.skills:
        return await no(f"[{skill}]은(는) 소지하지 않은 영문입니다.")
    if not p.alive: return await no("행동 불능 상태입니다.")
    if p.uid in raid.actions: return await no("이번 막에는 이미 행동했습니다.")

    a = Action(p.uid, kind)
    if kind == "공격": a.roll, a.det = atk_roll(p)      # 누르는 순간 굴려서 바로 보여줌
    elif kind == "방어": a.roll, a.det = def_roll(p)
    if kind == "영문":
        s = skill
        if s not in p.skills: return await no(f"[{s}]은(는) 소지하지 않은 영문입니다.")
        if s in PASSIVES: return await no("상시효과 영문은 따로 사용하지 않아도 자동 적용됩니다.")
        if p.skills[s] <= 0: return await no(f"[{s}] 사용 횟수가 남아있지 않습니다.")
        need, needval = ACTIVE[s]
        tids = [m.id for m in members if m]
        if len(tids) != need: return await no(f"[{s}]은(는) 대상 {need}명이 필요합니다.")
        if len(set(tids)) != len(tids): return await no("대상이 중복되었습니다.")
        for t in tids:
            tp = raid.players.get(t)
            if not tp: return await no("대상이 축귀 참가자가 아닙니다.")
            if not tp.alive: return await no("행동 불능인 대상은 지정할 수 없습니다.")
            if s in ("각광", "보호") and t == p.uid: return await no("본인은 지정할 수 없습니다.")
        if s == "사활" and p.hp > 50: return await no("체력이 50 이하일 때만 사용할 수 있습니다.")
        if s == "혼란" and not p.status_hit: return await no("상태 이상 피해를 입었을 때만 사용할 수 있습니다.")
        if needval:
            if not value or not (1 <= value <= 50): return await no("값은 1~50 사이로 입력하세요.")
            if value >= p.hp: return await no("현재 체력보다 적게 입력하세요.")
            a.value = value
        if s == "역행":
            tp = raid.players[tids[0]]
            if not name or tp.skills.get(name, -1) < 0 or name in PASSIVES:
                return await no("대상이 가진 행동형 영문 이름을 '이름'에 입력하세요.")
            a.name = name; tp.skills[name] += 1          # 즉시 +1 (대상이 이번 막에 쓸 수 있게)
        if s == "정화": a.name = name or ""
        a.skill, a.targets = s, tids
        p.skills[s] -= 1
        if s == "예지":
            txt = raid.gimmick.get("preview", "알 수 없음")
            await i.response.send_message(f"예지: 이번 전조 → {txt} (억제 등 이번 막 효과 적용 전)\n남은 시간: {remain_text(raid)}", ephemeral=True)
            raid.actions[p.uid] = a
            await raid.channel.send(f"🔮 **{p.name}**: [예지] 이번 전조의 결과를 미리 읽었다")   # 모두에게 공개
            maybe_done(raid)
            return
    raid.actions[p.uid] = a
    if kind == "영문":
        em = discord.Embed(title=f"✨ 영문 [{skill}]", description="사용을 선택했습니다.", color=0x9B59B6)
    else:
        atk = kind == "공격"
        em = discord.Embed(title="⚔️ 기본 공격" if atk else "🛡️ 기본 방어",
                           description=f"# {a.roll}", color=0xE67E22 if atk else 0x3498DB)
        if a.det != str(a.roll):
            em.add_field(name="굴림 내역", value=f"{a.det} = {a.roll}", inline=True)
    em.add_field(name="남은 시간", value=remain_text(raid), inline=True)
    em.set_author(name=p.name)
    em.set_footer(text=f"현재 체력 {p.hp}/{p.max_hp}")
    await i.response.send_message(embed=em, ephemeral=True)
    maybe_done(raid)


# ----- 기본 행동 -----
@tree.command(name="공격", description="이번 막에 기본 공격을 합니다")
async def cmd_attack(i: discord.Interaction):
    await submit(i, "공격")


@tree.command(name="방어", description="이번 막에 기본 방어를 합니다")
async def cmd_defend(i: discord.Interaction):
    await submit(i, "방어")


# ----- 영문: 영문마다 자기만의 명령어. 입력칸도 그 영문에 필요한 것만 나옴 -----
def add_simple_skill(skill):                      # 추가 입력이 없는 영문
    @tree.command(name=skill, description=f"영문 [{skill}]을(를) 사용합니다")
    async def _cmd(i: discord.Interaction):
        await submit(i, "영문", skill)


for _sk in ["항거", "사활", "중독", "복수", "혼란", "강신", "천도", "행운", "억제", "수복", "선포", "예지"]:
    add_simple_skill(_sk)


@tree.command(name="각광", description="영문 [각광]을(를) 사용합니다")
@app_commands.describe(대상1="첫 번째 대상", 대상2="두 번째 대상")
async def sk_gakgwang(i: discord.Interaction, 대상1: discord.Member, 대상2: discord.Member):
    await submit(i, "영문", "각광", (대상1, 대상2))


@tree.command(name="보호", description="영문 [보호]를 사용합니다")
@app_commands.describe(대상1="첫 번째 대상", 대상2="두 번째 대상")
async def sk_boho(i: discord.Interaction, 대상1: discord.Member, 대상2: discord.Member):
    await submit(i, "영문", "보호", (대상1, 대상2))


@tree.command(name="기도", description="영문 [기도]를 사용합니다")
@app_commands.describe(대상="회복할 대상")
async def sk_gido(i: discord.Interaction, 대상: discord.Member):
    await submit(i, "영문", "기도", (대상,))


@tree.command(name="분배", description="영문 [분배]를 사용합니다")
@app_commands.describe(대상="체력을 나눠줄 대상", 값="차감할 체력 (1~50)")
async def sk_bunbae(i: discord.Interaction, 대상: discord.Member, 값: int):
    await submit(i, "영문", "분배", (대상,), value=값)


@tree.command(name="정화", description="영문 [정화]를 사용합니다")
@app_commands.describe(대상="정화할 대상", 이름="해제할 저주 이름 (비우면 첫 번째 저주)")
async def sk_jeonghwa(i: discord.Interaction, 대상: discord.Member, 이름: str = None):
    await submit(i, "영문", "정화", (대상,), name=이름)


@tree.command(name="역행", description="영문 [역행]을(를) 사용합니다")
@app_commands.describe(대상="횟수를 늘려줄 대상", 이름="대상이 가진 영문 이름")
async def sk_yeokhaeng(i: discord.Interaction, 대상: discord.Member, 이름: str):
    await submit(i, "영문", "역행", (대상,), name=이름)


@tree.command(name="역살", description="영문 [역살]을(를) 사용합니다")
@app_commands.describe(값="차감할 체력 (1~50)")
async def sk_yeoksal(i: discord.Interaction, 값: int):
    await submit(i, "영문", "역살", value=값)


@tree.command(name="선택", description="왼쪽 또는 오른쪽을 선택합니다 (선택이 필요한 막에서만)")
@app_commands.choices(방향=[app_commands.Choice(name="왼쪽", value="왼쪽"),
                            app_commands.Choice(name="오른쪽", value="오른쪽")])
async def cmd_side(i: discord.Interaction, 방향: app_commands.Choice[str]):
    async def no(msg): await i.response.send_message(msg, ephemeral=True)
    raid = raids.get(i.channel_id)
    if not raid or not raid.open: return await no("이 채널에서는 지금 선택할 수 없습니다.")
    if raid.gimmick.get("kind") != "side": return await no("이번 막에는 선택할 것이 없습니다.")
    p = raid.players.get(i.user.id)
    if not p: return await no("이번 축귀 참가자가 아닙니다.")
    if not p.alive: return await no("행동 불능 상태입니다.")
    if p.uid in raid.choices: return await no("이미 선택했습니다.")
    raid.choices[p.uid] = 방향.value
    await i.response.send_message(f"{p.name}: **{방향.value}** 선택 완료\n남은 시간: {remain_text(raid)}", ephemeral=True)
    maybe_done(raid)


@tree.command(name="남은시간", description="이번 막의 남은 행동 시간을 확인합니다")
async def remaining(i: discord.Interaction):
    r = raids.get(i.channel_id)
    if not r or not r.open:
        return await i.response.send_message("지금은 행동 시간이 아닙니다.", ephemeral=True)
    waiting = [q.name for u, q in r.players.items()
               if q.alive and (u not in r.actions or (r.gimmick.get("kind") == "side" and u not in r.choices))]
    await i.response.send_message(
        f"{r.act}막 남은 시간: {remain_text(r)}\n아직 행동하지 않은 사람: {', '.join(waiting) or '없음'}",
        ephemeral=True)


@tree.command(name="강화", description="영석을 사용해 기본 공격 또는 기본 방어를 강화합니다")
@app_commands.describe(종류="강화할 대상")
@app_commands.choices(종류=[
    app_commands.Choice(name="기본 공격", value="atk"),
    app_commands.Choice(name="기본 방어", value="def")])
async def enhance(i: discord.Interaction, 종류: app_commands.Choice[str]):
    if any(running(r) and i.user.id in r.players for r in raids.values()):
        return await i.response.send_message("참가 중인 축귀가 진행 중에는 강화할 수 없습니다.", ephemeral=True)
    await i.response.defer()    # 시트를 읽는 동안 시간 초과 방지
    try:
        msg = await asyncio.to_thread(try_enhance, i.user.id, 종류.value)
    except Exception as e:
        msg = f"강화 처리 중 오류: {e}"
    await i.followup.send(msg)


@tree.command(name="내정보", description="내 강화 단계와 영석 개수를 확인합니다")
async def myinfo(i: discord.Interaction):
    await i.response.defer(ephemeral=True)
    try:
        info = await asyncio.to_thread(read_player, i.user.id)
    except Exception as e:
        return await i.followup.send(f"조회 중 오류: {e}")
    if not info:
        return await i.followup.send("이번 명단에 없는 사용자입니다.")
    lv = info["atk"] + info["def"]
    nxt = ""
    if lv < MAX_ENHANCE_TOTAL and (lv + 1) in info["table"]:
        pr, co = info["table"][lv + 1]
        nxt = f"\n다음 {lv + 1}강: 성공확률 {pr:g}% / 필요 영석 {co}개"
    r = raids.get(i.channel_id)
    rp = r.players.get(i.user.id) if r else None
    mine = ""
    if rp:
        mine = "\n내 영문: " + (", ".join(f"{k}({'상시' if n < 0 else str(n) + '회'})" for k, n in rp.skills.items()) or "없음")
    await i.followup.send(
        f"**{info['name']}**\n기본 공격 +{info['atk']}강 / 기본 방어 +{info['def']}강 (합계 {lv}/{MAX_ENHANCE_TOTAL})\n"
        f"보유 영석 {info['stones']}개{nxt}{mine}")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "sample":
        make_sample()
    else:
        client.run(os.environ["DISCORD_TOKEN"])
