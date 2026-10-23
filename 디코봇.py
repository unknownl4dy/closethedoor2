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
import re
import sys
import threading
import time
import urllib.request
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from dataclasses import dataclass, field

import discord
from discord import app_commands
from openpyxl import Workbook, load_workbook

EXCEL_PATH = "raid.xlsx"
NEXT_ACT_DELAY = 5        # 공격 결과 카드 이후 다음 막이 나오기 전 대기 시간(초)
ACTION_WINDOW = 600        # 행동 가능 시간(초) = 10분
BASE_HP = 120
BONUS_PER_ENHANCE = 2      # 강화 1당 +2
MAX_SKILLS_PER_PLAYER = 2   # 한 사람이 새길 수 있는 영문 수
MAX_ENHANCE = 10           # 기본 공격, 기본 방어 각각의 강화 상한
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
    skill_max: dict = field(default_factory=dict)   # 영문별 최대 사용 횟수 (시작 시 값)
    poison_left: int = 0         # 감염 남은 턴 (사귀 2막 기믹)
    fear_left: int = 0           # 공포 남은 막 수 (사귀 7막 기믹): 공격값 절반
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
    pre: dict = field(default_factory=dict)   # 영문을 쓰는 순간 미리 굴린 값


@dataclass
class Ctx:                       # 한 막 처리 중 임시 상태
    dmg: int = 0
    reduce: int = 0
    defended: dict = field(default_factory=dict)
    shield: dict = field(default_factory=lambda: defaultdict(int))
    redirect: int | None = None
    spot: list = field(default_factory=list)
    by: dict = field(default_factory=lambda: defaultdict(int))   # 사람별로 사귀에게 입히려던 피해 (빙의용)


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
    mpc: dict = field(default_factory=dict)       # 이번 막에 부른 MPC -> 부른 사람 이름
    mpc_note: dict = field(default_factory=dict)  # MPC -> 막 결과에 쓸 지문 (대상자 이름이 들어간 완성 문장)
    deadline: float = 0.0          # 이번 막 행동 마감 시각 (time.monotonic 기준)


# ---------- 굴림 ----------
def _fmt(die, fixed, bonus, level=0):
    """주사위 눈 + 고정값(집념·불굴의 +10) + 강화 보정(강화 단계 표시). 없는 항목은 생략: 13+6(강화3) / 5+10+12(강화6)"""
    return (f"{die}" + (f"+{fixed}" if fixed else "")
            + (f"+{bonus}(강화{level})" if bonus else ""))


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
            die, fixed = (10 if max_dice else d(10)), 10
        else:
            die, fixed = (20 if max_dice else d(20)), 0
        total += die + fixed + bonus
        parts.append(_fmt(die, fixed, bonus, p.atk_up))
    return total, _join(parts)


def def_roll(p: Player):
    """기본 방어: D20 (불굴 D10+10), 인내면 2회, 강화 보정 +2/단계"""
    bonus = p.def_up * BONUS_PER_ENHANCE
    total, parts = 0, []
    for _ in range(2 if p.has("인내") else 1):
        if p.has("불굴"):
            die, fixed = d(10), 10
        else:
            die, fixed = d(20), 0
        total += die + fixed + bonus
        parts.append(_fmt(die, fixed, bonus, p.def_up))
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
SMASH_TEXT = "D30+20"            # 강타: 기본은 비공개(???)
SMASH_REVEAL_ACTS = set()        # 강타 값을 공개할 막 번호를 여기에 넣으면 전조에 표시 (예: {3})
SHARE_LIMIT = 100                # 3막: 공격합이 이 값 미만이면 광역 피해
POISON_TURNS = 3                 # 감염(2막): 지속 막 수 (이번 막 포함), 피해는 5+D10
FEAR_LIMIT, FEAR_TURNS = 30, 3   # 공포(7막): 공격값이 이 값 이하이면 걸림 / 지속 막 수
# 막 번호 -> 기믹 종류. 적지 않은 막은 'aoe'(기본공격 광역)
GIMMICKS = {1: "smash", 2: "poison", 3: "share", 4: "side", 5: "fear", 6: "smash", 7: "possess", 8: "side"}


def josa(word, with_final, without_final):
    """받침 유무에 따라 이/가, 을/를 등을 고름"""
    ch = word[-1]
    if "가" <= ch <= "힣" and (ord(ch) - 0xAC00) % 28:
        return with_final
    return without_final


def red(*lines):
    """임베드 안에서 빨간 글씨처럼 보이게 하는 블록 (diff 코드블록의 '-' 줄은 빨간색)"""
    return "```diff\n" + "\n".join(f"- {x}" for x in lines) + "\n```"


def basic_roll(): return 10 + d(20)
def smash_roll(): return 20 + d(30)


def prepare_act(r: Raid) -> str:
    """전조 단계: 이번 막의 기믹을 정하고 전조 설명문을 반환. r.gimmick / r.plan 채움"""
    b = r.boss
    alive = [p for p in r.players.values() if p.alive]
    kind = GIMMICKS.get(r.act, "aoe")
    r.plan, r.choices = {}, {}
    if kind == "smash":
        t = random.choice(alive); v = smash_roll()
        r.plan = {t.uid: v}
        r.gimmick = {"kind": kind, "preview": f"[{t.name}]에게 강타 [{SMASH_TEXT}]"}
        return f"{b.name}{josa(b.name, '이', '가')} **[{t.name}]**을(를) 바라본다."
    if kind == "share":
        v = basic_roll()
        r.gimmick = {"kind": kind, "value": v,
                     "preview": f"이번 막 공격합 [{SHARE_LIMIT}] 미만이면 광역 피해 [{BASIC_TEXT}], 이상이면 피해 무효"}
        return (f"{b.name}{josa(b.name, '이', '가')} 몸을 웅크리며 힘을 모은다.\n"
                + red(f"이번 막 공격합 [{SHARE_LIMIT}] 이상 필요",
                      f"공격합이 [{SHARE_LIMIT}] 미만이면 광역 피해"))
    if kind == "poison":
        t = random.choice(alive)
        r.gimmick = {"kind": kind, "target": t.uid,
                     "preview": f"[{t.name}]에게 감염 [5+D10] ({POISON_TURNS}막 지속)"}
        return (f"보랏빛 안개가 **[{t.name}]**{josa(t.name, '을', '를')} 감싼다.\n"
                + red("[상태이상] 감염", f"{POISON_TURNS}막 동안 매 막 5+D10 피해"))
    if kind == "fear":
        t = random.choice(alive)
        r.gimmick = {"kind": kind, "target": t.uid,
                     "preview": f"[{t.name}]의 공격값 [{FEAR_LIMIT}] 이하일 경우 공포 ({FEAR_TURNS}막 동안 공격값 절반)"}
        return (f"{b.name}의 눈이 검게 변하며 **[{t.name}]**{josa(t.name, '을', '를')} 가리킨다.\n"
                + red(f"[{t.name}]의 공격값 [{FEAR_LIMIT}] 이하일 경우", "[상태이상] 공포"))
    if kind == "possess":
        picks = random.sample(alive, min(2, len(alive)))
        for q in picks:
            if "빙의" not in q.curses:
                q.curses.append("빙의")              # 저주로 취급: 정화로 해제 가능
        names = ", ".join(f"[{q.name}]" for q in picks)
        r.gimmick = {"kind": kind, "picked": [q.uid for q in picks],
                     "preview": f"{names}에게 빙의 (공격값이 자신을 제외한 팀원 2인에게 들어감)"}
        return (f"{b.name}{josa(b.name, '이', '가')} {names}에게 깃든다.\n"
                + red("빙의에 걸린 사람의 공격값은 사귀가 아닌",
                      "자신을 제외한 팀원 2인에게 랜덤으로 들어간다",
                      "빙의 상태에서는 기본 공격만 할 수 있다 (정화로 해제)"))
    if kind == "side":
        v = basic_roll(); hit = random.choice(["왼쪽", "오른쪽"])    # 어느 쪽을 칠지 전조 때 미리 정함 (결과 때 공개, 예지는 미리 공개)
        r.gimmick = {"kind": kind, "value": v, "hit": hit,
                     "preview": f"[{hit}]을 고른 사람에게 광역 [{BASIC_TEXT}]"}
        return (f"{b.name}{josa(b.name, '이', '가')} 높게 뛰어오른다.\n"
                "`/선택`으로 **왼쪽**, **오른쪽**을 고른다. 공격은 선택과 별개로 진행한다.\n"
                + red("사귀가 지정한 쪽을 고른 사람만 광역 피해를 입는다"))
    v = basic_roll()
    r.gimmick = {"kind": "aoe", "value": v, "preview": f"전원에게 광역 [{BASIC_TEXT}]"}
    return f"{b.name}{josa(b.name, '이', '가')} 광역 공격을 준비한다."


# ---------- 영문 처리 ----------
def preroll(p: Player, a: Action):
    """영문을 쓰는 순간 필요한 주사위를 모두 굴려 a.pre에 저장 (막 결과는 이 값을 그대로 사용)"""
    s = a.skill
    if s == "중독": a.pre = {"v": d(10) + 5}
    elif s == "복수":
        v, det = atk_roll(p); a.pre = {"v": v, "det": det, "extra": p.last_taken // 2}
    elif s == "혼란":
        v, det = atk_roll(p, max_dice=True); a.pre = {"v": v, "det": det}
    elif s == "강신": a.pre = {"v": 30 + d(30)}
    elif s == "천도": a.pre = {"v": nd(2, 25)}
    elif s == "행운":
        x = (d(10), d(15), d(20)); a.pre = {"v": sum(x), "x": x}
    elif s == "역살":
        v, det = atk_roll(p); a.pre = {"v": v, "det": det}
    elif s == "억제": a.pre = {"v": d(30)}
    elif s == "기도": a.pre = {"v": 10 + d(15)}
    elif s == "수복":
        r15 = d(15); a.pre = {"r15": r15, "v": 10 + r15}
    elif s == "선포": a.pre = {"v": 5 + d(10)}
    elif s == "보호":
        v, det = def_roll(p); a.pre = {"v": v, "det": det}


def big_value(a: Action) -> str:
    """영문 카드에 크게 보여줄 핵심 수치 (피해 / 회복량 / 보호막)"""
    s, pr = a.skill, a.pre
    if s in ("강신", "천도", "행운", "중독", "혼란", "억제", "수복", "보호"): return str(pr["v"])
    if s == "복수": return str(pr["v"] + pr["extra"])
    if s == "역살": return str(pr["v"] + a.value)
    if s in ("기도", "선포"): return f"+{pr['v']}"
    if s == "분배": return f"+{a.value}"
    return ""


def apply_heal(r: Raid, p: Player, a: Action):
    """기도·수복·선포·분배의 회복을 쓰는 즉시 적용 (최대 체력까지만). 결과는 a.pre['got']에 기록"""
    s = a.skill
    if s == "기도":
        t = r.players[a.targets[0]]; n = heal(t, a.pre["v"])
        a.pre["got"] = [f"{t.name} +{n} ({hp_str(t)})"]
    elif s == "수복":
        n = heal(p, a.pre["r15"])
        a.pre["got"] = [f"{p.name} +{n} ({hp_str(p)})"]
    elif s == "선포":
        amt = a.pre["v"]
        a.pre["got"] = [f"{q.name} +{heal(q, amt)} ({hp_str(q)})" for q in r.players.values() if q.alive]
    elif s == "분배":
        t = r.players[a.targets[0]]; n = heal(t, a.value)
        a.pre["got"] = [f"{t.name} +{n} ({hp_str(t)})"]


def preview_text(a: Action, names: list) -> str:
    """영문 사용 카드에 바로 보여줄 결과 (피해·회복·보호막 수치)"""
    s, pr, t = a.skill, a.pre, (names[0] if names else "")
    if s == "중독": return f"**피해 {pr['v']}** (이후 2막 동안 매 막 {pr['v']})"
    if s == "복수": return f"기본 공격 {shown(pr['det'], pr['v'])} + 복수 {pr['extra']} = **피해 {pr['v'] + pr['extra']}**"
    if s == "혼란": return f"기본 공격 최대치 {shown(pr['det'], pr['v'])} = **피해 {pr['v']}**"
    if s in ("강신", "천도"): return f"**피해 {pr['v']}**"
    if s == "행운": return f"{' + '.join(map(str, pr['x']))} = **피해 {pr['v']}**"
    if s == "역살": return f"체력 -{a.value}, 기본 공격 {shown(pr['det'], pr['v'])} + {a.value} = **피해 {pr['v'] + a.value}**"
    if s == "억제": return f"**피해 {pr['v']}**, 이번 막 사귀 공격값 -{pr['v']}"
    got = ", ".join(pr.get("got", []))
    if s == "기도": return f"체력 회복 → **{got}**"
    if s == "수복": return f"**피해 {pr['v']}**, 체력 회복 → **{got}**"
    if s == "선포": return f"전원 체력 회복 → **{got}**"
    if s == "보호": return f"방어값 {shown(pr['det'], pr['v'])} → 보호막 **{pr['v']}**"
    if s == "분배": return f"체력 -{a.value}, 체력 회복 → **{got}**"
    if s == "정화":
        return f"{t}의 저주 **'{pr['cleansed']}'** 해제" if pr.get("cleansed") else f"{t}에게 해제할 저주가 없다"
    return ""


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
        v, det = a.pre["v"], a.pre["det"]
        for t in tg: c.shield[t.uid] += v
        return f"{head} 방어값 {shown(det, v)} → {tg[0].name}, {tg[1].name}에게 보호막 {v}"
    if s == "중독":
        v = a.pre["v"]; c.dmg += v; b.poisons.append([v, 2])
        return f"{head} {v} 피해 (이후 2막 동안 매 막 {v})"
    if s == "복수":
        extra = a.pre["extra"]; v, det = a.pre["v"], a.pre["det"]; c.dmg += v + extra
        return f"{head} 기본 공격 : {shown(det, v)} + 복수 {extra} = {v + extra}"
    if s == "혼란":
        v, det = a.pre["v"], a.pre["det"]; c.dmg += v; p.status_hit = False
        return f"{head} 기본 공격 최대치 : {shown(det, v)}"
    if s == "강신":
        v = a.pre["v"]; c.dmg += v; return f"{head} 30+D30 → {v}"
    if s == "천도":
        v = a.pre["v"]; c.dmg += v; return f"{head} 2D25 → {v}"
    if s == "행운":
        x = a.pre["x"]; v = a.pre["v"]; c.dmg += v
        return f"{head} D10+D15+D20 {x} → {v}"
    if s == "역살":
        v, det = a.pre["v"], a.pre["det"]; c.dmg += v + a.value      # 체력 차감은 사용 시점에 이미 적용됨
        return f"{head} 체력 -{a.value}, 기본 공격 : {shown(det, v)} + {a.value} = {v + a.value}"
    if s == "억제":
        v = a.pre["v"]; c.dmg += v; c.reduce += v
        return f"{head} D30 → {v} 피해, 사귀 공격값 -{v}"
    if s == "기도":
        return f"{head} 체력 회복 → {', '.join(a.pre['got'])}"
    if s == "수복":
        v = a.pre["v"]; c.dmg += v
        return f"{head} 10+D15 → {v} 피해, 체력 회복 → {', '.join(a.pre['got'])}"
    if s == "선포":
        return f"{head} 전원 체력 회복 → {', '.join(a.pre['got'])}"
    if s == "분배":
        return f"{head} 체력 -{a.value}, 체력 회복 → {', '.join(a.pre['got'])}"
    if s == "정화":
        t = tg[0]
        cur = a.pre.get("cleansed")
        if not cur:
            return f"{head} {t.name}에게 해제할 저주가 없다"
        return f"{head} ✨ {t.name}의 저주 '{cur}'이(가) 해제되었다"
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
    for nm, who in r.mpc.items():
        log.append(f"- 🌟 {who}: MPC [{nm}] — {r.mpc_note.get(nm, MPC_INFO[nm][0])}")
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
        before = c.dmg
        line = do_action(r, c, a)
        c.by[a.uid] += c.dmg - before
        if line:
            log.append(line)
    if not r.actions:
        log.append("- 아무도 행동하지 않았다.")

    if r.gimmick.get("kind") == "fear":                 # 공포 판정용: 보정(공포·빙의) 전 대상의 공격값
        r.gimmick["seen"] = c.by.get(r.gimmick["target"], 0)

    # ---- 공포 (7막 기믹): 걸려 있는 동안 본인 공격값 절반 (소수점 버림) ----
    for q in r.players.values():
        if "공포" in q.curses and q.fear_left > 0 and q.alive:
            moved = c.by.get(q.uid, 0); keep = moved // 2
            if moved > 0:
                c.dmg -= moved - keep; c.by[q.uid] = keep
                log.append(f"- 😱 {q.name}: 공포! 공격값 {moved} → {keep}")
            q.fear_left -= 1
            if q.fear_left <= 0: q.curses.remove("공포")

    # ---- 빙의 (5막 기믹): 빙의에 걸린 사람의 공격값은 사귀 대신 팀원 2인에게 ----
    friendly = {}
    if r.gimmick.get("kind") == "possess":
        for u in r.gimmick.get("picked", []):         # 정화로 풀린 사람은 한 번 더 확실히 알림
            q = r.players[u]
            if "빙의" not in q.curses:
                log.append(f"- ✨ {q.name}: 정화로 빙의가 해제되어 공격이 정상적으로 사귀에게 들어간다")
        for q in r.players.values():
            if "빙의" not in q.curses or not q.alive:
                continue
            moved = c.by.get(q.uid, 0)
            mates = [m for m in r.players.values() if m.alive and m.uid != q.uid]
            if moved > 0 and mates:
                c.dmg -= moved
                tg = random.sample(mates, min(2, len(mates)))
                for m in tg:
                    friendly[m.uid] = friendly.get(m.uid, 0) + moved
                log.append(f"- 👻 {q.name}: 빙의! 공격값 {moved}이(가) 사귀 대신 {', '.join(m.name for m in tg)}에게 들어간다")
            else:
                log.append(f"- 👻 {q.name}: 빙의 상태지만 넘길 공격값이 없다")

    if "일도양단" in r.mpc:                              # MPC: 합산 공격값 1.5배 (소수점 버림)
        boosted = c.dmg * 3 // 2
        log.append(f"- 🌟 [일도양단] 합산 공격값 {c.dmg} → **{boosted}** (1.5배, 소수점 버림)")
        c.dmg = boosted

    b.hp = max(0, b.hp - c.dmg)
    if b.hp <= 0:
        return log, [], c.dmg, True

    # ---- 사귀 공격 ----
    res = []
    kind = r.gimmick.get("kind", "aoe")
    alive_now = [q for q in r.players.values() if q.alive]
    base = {}
    void_dmg = "천상불침" in r.mpc                        # MPC: 사귀의 이번 막 피해값 무효
    if void_dmg:
        res.append("- 🌟 [천상불침] 사귀의 피해값이 모두 무효가 되었다")
    elif kind == "smash":
        base = dict(r.plan)
    elif kind == "share":
        if c.dmg >= SHARE_LIMIT:
            res.append(f"- 공격합 {c.dmg} (기준 {SHARE_LIMIT} 이상) → 사귀의 피해가 무효가 되었다!")
        else:
            base = {q.uid: r.gimmick["value"] for q in alive_now}
            res.append(f"- 공격합 {c.dmg} (기준 {SHARE_LIMIT} 미만) → 광역 피해 발동!")
    elif kind == "side":
        hit_side = r.gimmick["hit"]
        got = []
        for q in alive_now:
            side = r.choices.get(q.uid)
            if side is None:
                side = random.choice(["왼쪽", "오른쪽"])
                res.append(f"- {q.name}: 선택하지 않아 {side}으로 정해졌다")
            if side == hit_side: got.append(q)
        other = "오른쪽" if hit_side == "왼쪽" else "왼쪽"
        res.append(f"- {b.name}{josa(b.name, '이', '가')} **{hit_side}**을 덮쳤다! "
                   f"{hit_side}: " + (", ".join(q.name for q in got) or "없음")
                   + f" / {other}: " + (", ".join(q.name for q in alive_now if q not in got) or "없음"))
        base = {q.uid: r.gimmick["value"] for q in got}
    elif kind == "possess":
        base = {}
    elif kind == "aoe":
        base = {q.uid: r.gimmick["value"] for q in alive_now}
    plan = {u: max(0, v - c.reduce) for u, v in base.items()}
    if c.reduce and base:
        res.append(f"- 억제로 사귀 공격값 -{c.reduce}")
    for u, v in friendly.items():
        plan[u] = plan.get(u, 0) + v
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
        who = "빙의된 팀원의 공격으로" if kind == "possess" else f"{b.name}의 강한 공격으로"
        res.append(f"- {who} **{p.name} 체력 -{dmg}** → {hp_str(p)}{extra}")
        if not p.alive:
            res.append(f"  {p.name}은(는) 행동 불능이 되었다.")

    if kind == "possess":                              # 빙의는 이번 막으로 끝
        for q in r.players.values():
            if "빙의" in q.curses:
                q.curses.remove("빙의")

    # ---- 독 (3막 기믹): 방어와 상관없이 들어가는 지속 피해, 정화로 해제 가능 ----
    for q in r.players.values():
        q.status_hit = False
    if kind == "poison":
        t = r.players[r.gimmick["target"]]
        if t.alive and "감염" not in t.curses:
            t.curses.append("감염"); t.poison_left = POISON_TURNS
    for q in r.players.values():
        if q.alive and "감염" in q.curses and q.poison_left > 0:
            dd = 0 if (q.invul_until >= r.act or void_dmg) else 5 + d(10)
            q.hp = max(0, q.hp - dd); q.poison_left -= 1
            if dd: q.status_hit = True
            if q.poison_left <= 0: q.curses.remove("감염")
            res.append(f"- 감염: **{q.name} 체력 -{dd}**" + (" (피해 무효)" if not dd else "") + f" → {hp_str(q)} (남은 {q.poison_left}턴)")
            if not q.alive:
                res.append(f"  {q.name}은(는) 행동 불능이 되었다.")
    if kind == "fear":                                 # 7막: 대상의 공격값(보정 전 원래 값)이 기준 이하면 공포
        t = r.players[r.gimmick["target"]]
        if t.alive:
            val = r.gimmick.get("seen", 0)
            if val <= FEAR_LIMIT:
                if "공포" not in t.curses: t.curses.append("공포")
                t.fear_left = FEAR_TURNS
                res.append(f"- 😱 {t.name}: 공격값 {val} (기준 {FEAR_LIMIT} 이하) → **공포**! {FEAR_TURNS}막 동안 공격값 절반")
            else:
                res.append(f"- {t.name}: 공격값 {val} (기준 {FEAR_LIMIT} 초과) → 공포를 이겨냈다")
    return log, res, c.dmg, False


# ---------- 진행 ----------
raids: dict[int, Raid] = {}   # 채널ID -> 그 채널에서 진행 중인(또는 끝난) 축귀


def running(r):
    return bool(r and r.task and not r.task.done())


# ===== MPC: 전체 조에서 각각 1번씩만 부를 수 있음 =====
MPC_INFO = {   # 이름: (지문, 효과 설명)
    "천상불침": ("의영공주의 토구가 모두를 지킨다.", "사귀의 이번 막 피해값 무효"),
    "일도양단": ("단이 불을 피우자, 바리공주가 나아갈 길을 알려준다.", "이번 막 유저 공격값 합산 1.5배 (소수점 버림)"),
    "고목봉춘": ("묵희가 {target}에게 비늘을 건넨다.", "대상자 1인의 체력 전체 회복"),
}
MPC_PATH = "mpc_state.json"


def load_mpc():
    try:
        with open(MPC_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_mpc():
    try:
        with open(MPC_PATH, "w", encoding="utf-8") as f:
            json.dump(mpc_used, f, ensure_ascii=False)
    except Exception:
        pass


mpc_used = load_mpc()   # {MPC 이름: 부른 채널ID} - 한 번 불리면 다른 조에서는 못 부름


def needs_choice(r: Raid) -> bool:
    return r.gimmick.get("kind") == "side" and "천상불침" not in r.mpc


def maybe_done(r: Raid):
    """생존자 전원이 행동(4막은 선택까지) 마쳤으면 10분을 기다리지 않고 진행"""
    alive = [u for u, q in r.players.items() if q.alive]
    if all(u in r.actions for u in alive) and (
            not needs_choice(r) or all(u in r.choices for u in alive)):
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
    """체력 막대: 비율에 따라 초록/노랑/빨강 칸 + 빈 칸. 체력이 남아 있으면 최소 1칸"""
    ratio = hp / mx if mx else 0
    filled = 0 if hp <= 0 else max(1, round(n * ratio))
    color = "🟩" if ratio > 0.5 else ("🟨" if ratio > 0.25 else "🟥")
    return color * filled + "⬛" * (n - filled)


def final_summary(r: Raid, title: str, desc: str, color: int):
    """축귀가 끝났을 때 전원의 최종 체력을 정리한 카드"""
    lines = [f"{'💀' if not q.alive else '❤️'} **{q.name}** {q.hp}/{q.max_hp}" + (" (행동 불능)" if not q.alive else "")
             for q in r.players.values()]
    em = discord.Embed(title=title, description=desc, color=color)
    em.add_field(name="최종 체력", value=clip("\n".join(lines), 1024), inline=False)
    return em


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
        r.actions.clear(); r.all_done.clear(); r.mpc.clear(); r.mpc_note.clear()
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
            await r.channel.send(embed=final_summary(
                r, f"{b.name} 쓰러짐!", "축귀가 조기 종료되었습니다.", COLOR_WIN))
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
        if r.act < b.acts and b.hp > 0 and any(q.alive for q in r.players.values()):
            await asyncio.sleep(NEXT_ACT_DELAY)          # 다음 막 전조까지 잠깐 쉬어가기
    if any(q.alive for q in r.players.values()):
        end_desc = f"{r.act}막까지 진행되었습니다."
    else:
        end_desc = "축귀인이 모두 행동 불능이 되었습니다."
    await r.channel.send(embed=final_summary(r, "축귀 종료", end_desc, COLOR_END))


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
        label = "기본 공격" if kind == "atk" else "기본 방어"
        level = info[kind]                      # 공격/방어는 각각 따로 1강~10강
        if level >= MAX_ENHANCE:
            return f"{label}은(는) 이미 최대 강화({MAX_ENHANCE}강)입니다."
        nxt = level + 1
        if nxt not in info["table"]:
            return f"강화 시트에 {nxt}강 정보가 없습니다."
        prob, cost = info["table"][nxt]
        if info["stones"] < cost:
            return f"영석이 부족합니다. ({label} {nxt}강 필요 {cost}개 / 보유 {info['stones']}개)"
        success = random.random() * 100 < prob
        if SHEET_WEBHOOK:                         # 시트에 직접 기록 (실패하면 아무것도 변하지 않음)
            res = post_enhance({
                "uid": str(uid), "name": info["name"], "kind": label,
                "result": "성공" if success else "실패", "cost": cost,
                "atk_add": 1 if success and kind == "atk" else 0,
                "def_add": 1 if success and kind == "def" else 0,
                "max_total": MAX_ENHANCE})
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
        if isinstance(x[1], float):
            warns.append(f"{name}: 디스코드ID가 숫자로 저장되어 끝자리가 깨졌을 수 있음 (ID 칸 서식을 '일반 텍스트'로 바꾸고 다시 붙여넣으세요)")
        au, du = int(x[2] or 0), int(x[3] or 0)
        st = bonus(uid)                       # /강화로 올린 단계를 더함 (시트 직접 기록 방식이면 없음)
        au += st.get("atk", 0); du += st.get("def", 0)
        if au > MAX_ENHANCE or du > MAX_ENHANCE:
            warns.append(f"{name}: 강화 단계 초과 (공격 {au}, 방어 {du} / 최대 {MAX_ENHANCE})")
        skills = {}
        raw = str(x[4] or "").replace("，", ",").replace("、", ",")   # 전각 쉼표도 인식
        entries = [t for t in raw.split(",") if t.strip()]
        if len(entries) > MAX_SKILLS_PER_PLAYER:
            warns.append(f"{name}: 영문 {len(entries)}개 (최대 {MAX_SKILLS_PER_PLAYER}개)")
        direct = {t.strip() for t in raw.split(",") if t.strip() and not t.strip().startswith("모사=")}
        for s in raw.split(","):
            s = s.strip()
            if not s: continue
            if s.startswith("모사="):
                src = s[3:].strip()
                if src not in defs:
                    warns.append(f"{name}: 모사 대상 '{src}' 없음"); continue
                if src in direct or src == "모사":      # 이미 가진 영문(또는 모사 자체)은 모사 불가
                    warns.append(f"{name}: 이미 가진 영문 '{src}'은(는) 모사할 수 없음 (모사 무시됨)"); continue
                skills["모사"] = -1; skills[src] = defs[src][0]
                holders["모사"].append(name)
            elif s in defs:
                skills[s] = defs[s][0]; holders[s].append(name)
            else:
                warns.append(f"{name}: 알 수 없는 영문 '{s}'")
        if not in_group:
            continue
        p = Player(name, uid, au, du, skills)
        p.skill_max = dict(skills)
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
    ws.append(["사귀", 500, "N", 0, 8, 1])
    ws.append(["사귀", 500, "N", 0, 8, 2])
    ws.append(["사귀", 500, "N", 0, 8, 3])
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
    ws.append(["단", "123456789012345678", 0, 0, "강신,기도", 10, 1, 120])   # ID는 반드시 텍스트로
    ws.column_dimensions["B"].width = 24
    for c in ws["B"][1:]:
        c.number_format = "@"
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
        return await i.response.send_message(embed=final_summary(raid, "축귀 종료", "운영진이 축귀를 종료했습니다.", COLOR_END))
    await i.response.send_message("진행 중인 축귀가 없습니다.", ephemeral=True)


SKILL_EFFECT = {
    "각광": "지정 2인의 피해를 대신 받는다 (피해 절반)",
    "항거": "이번 막의 모든 피해를 대신 받고, 다음 막에 그만큼 사귀를 공격",
    "사활": "이번 막 포함 2막 동안 피해를 입지 않는다",
    "보호": "지정 2인에게 방어값만큼 보호막 (본인 제외)",
    "중독": "사귀에게 D10+5 피해, 이번 막 포함 3막 지속",
    "복수": "직전에 받은 피해의 절반을 기본 공격에 더한다",
    "혼란": "이번 막 기본 공격 주사위가 최대로 나온다",
    "강신": "30+D30으로 공격",
    "천도": "2D25로 공격",
    "행운": "D10+D15+D20으로 공격",
    "역살": "체력을 {v} 차감하고, 그만큼 기본 공격에 더한다",
    "억제": "D30으로 공격, 이번 막 사귀 공격값이 그만큼 감소",
    "기도": "지정 대상의 체력을 10+D15 회복",
    "수복": "10+D15로 공격, D15만큼 본인 체력 회복",
    "선포": "전원의 체력을 5+D10 회복",
    "분배": "체력을 {v} 차감하고, 지정 대상을 그만큼 회복",
    "정화": "지정 대상의 저주를 1개 해제",
    "역행": "지정 대상의 [{n}] 사용 횟수 +1",
}
TARGET_LABEL = {"각광": "대신 받는 대상", "보호": "보호 대상", "기도": "치료 대상",
                "분배": "회복 대상", "정화": "정화 대상", "역행": "횟수 +1 대상"}


NOTE_DESC = "행동 묘사 (선택) 예: 칼을 휘두른다"


def clean_note(text):
    """지문: 마크다운이 깨지지 않게 특수문자를 빼고 길이를 제한"""
    if not text:
        return ""
    for ch in "*_~`|>":
        text = text.replace(ch, "")
    return text.strip()[:200]


async def submit(i: discord.Interaction, kind, skill="", members=(), value=None, name=None, note=None):
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
    if kind in ("방어", "영문") and "빙의" in p.curses:
        return await no("빙의 상태에서는 기본 공격 외에는 아무것도 할 수 없습니다. (정화로 빙의가 해제되면 가능합니다)")

    note = clean_note(note)
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
                return await no("대상이 가진 행동형 영문 이름을 '영문명'에 입력하세요.")
            a.name = name; tp.skills[name] += 1          # 즉시 +1 (대상이 이번 막에 쓸 수 있게)
        if s == "정화": a.name = name or ""
        a.skill, a.targets = s, tids
        p.skills[s] -= 1
        preroll(p, a)
        if s == "정화":                          # 저주 해제도 쓰는 즉시 반영
            t = raid.players[tids[0]]
            if t.curses:
                cur = a.name if a.name in t.curses else t.curses[0]
                t.curses.remove(cur)
                a.pre["cleansed"] = cur
            else:
                a.pre["cleansed"] = None
        if s in ("역살", "분배"):
            p.hp -= a.value                      # 본인 체력 차감은 쓰는 즉시 반영
        apply_heal(raid, p, a)                   # 회복도 쓰는 즉시 반영 (최대 체력까지만)
        if s == "예지":
            txt = raid.gimmick.get("preview", "알 수 없음")
            await i.response.send_message(f"예지: 이번 전조 → {txt} (억제 등 이번 막 효과 적용 전)\n남은 시간: {remain_text(raid)}", ephemeral=True)
            raid.actions[p.uid] = a
            await raid.channel.send(f"🔮 **{p.name}**: [예지] 이번 전조의 결과를 미리 읽었다" + (f"\n*{note}*" if note else ""))   # 모두에게 공개
            maybe_done(raid)
            return
    raid.actions[p.uid] = a
    if kind == "영문":
        effect = SKILL_EFFECT.get(skill, "").format(v=a.value, n=a.name)
        big = big_value(a)
        em = discord.Embed(title=f"✨ 영문 [{skill}]",
                           description=(f"# {big}\n" if big else "") + effect, color=0x9B59B6)
        if a.targets:
            em.add_field(name=TARGET_LABEL.get(skill, "대상"),
                         value=", ".join(raid.players[u].name for u in a.targets), inline=True)
        elif skill == "선포":
            em.add_field(name="대상", value="전원", inline=True)
        res_txt = preview_text(a, [raid.players[u].name for u in a.targets])
        if res_txt:
            em.add_field(name="결과", value=res_txt, inline=False)
        n_left = p.skills.get(skill, 0)
        em.add_field(name="남은 횟수", value=f"({n_left}/{max(n_left, p.skill_max.get(skill, n_left))})", inline=True)
    else:
        atk = kind == "공격"
        # 합산값을 가장 큰 글씨로, 바로 아래에 굴림 + 보정 내역
        detail = f"\n{a.det}" if a.det != str(a.roll) else ""
        em = discord.Embed(title="⚔️ 기본 공격" if atk else "🛡️ 기본 방어",
                           description=f"# {a.roll}{detail}", color=0xE67E22 if atk else 0x3498DB)
    em.add_field(name="남은 시간", value=remain_text(raid), inline=True)
    if note:
        em.description = f"*{note}*\n" + (em.description or "")      # 지문은 이탤릭으로 맨 위에
    em.set_author(name=p.name)
    em.set_footer(text=f"현재 체력 {hp_bar(p.hp, p.max_hp)} {p.hp}/{p.max_hp}")
    await i.response.send_message(embed=em)      # 모두에게 공개
    maybe_done(raid)


# ----- 기본 행동 -----
@tree.command(name="공격", description="이번 막에 기본 공격을 합니다")
@app_commands.describe(지문=NOTE_DESC)
async def cmd_attack(i: discord.Interaction, 지문: str = None):
    await submit(i, "공격", note=지문)


@tree.command(name="방어", description="이번 막에 기본 방어를 합니다")
@app_commands.describe(지문=NOTE_DESC)
async def cmd_defend(i: discord.Interaction, 지문: str = None):
    await submit(i, "방어", note=지문)


# ----- 영문: 영문마다 자기만의 명령어. 입력칸도 그 영문에 필요한 것만 나옴 -----
def add_simple_skill(skill):                      # 추가 입력이 없는 영문
    @tree.command(name=skill, description=f"영문 [{skill}]을(를) 사용합니다")
    @app_commands.describe(지문=NOTE_DESC)
    async def _cmd(i: discord.Interaction, 지문: str = None):
        await submit(i, "영문", skill, note=지문)


for _sk in ["항거", "사활", "중독", "복수", "혼란", "강신", "천도", "행운", "억제", "수복", "선포", "예지"]:
    add_simple_skill(_sk)


@tree.command(name="각광", description="영문 [각광]을(를) 사용합니다")
@app_commands.describe(대상자1="첫 번째 대상자", 대상자2="두 번째 대상자", 지문=NOTE_DESC)
async def sk_gakgwang(i: discord.Interaction, 대상자1: discord.Member, 대상자2: discord.Member, 지문: str = None):
    await submit(i, "영문", "각광", (대상자1, 대상자2), note=지문)


@tree.command(name="보호", description="영문 [보호]를 사용합니다")
@app_commands.describe(대상자1="첫 번째 대상자", 대상자2="두 번째 대상자", 지문=NOTE_DESC)
async def sk_boho(i: discord.Interaction, 대상자1: discord.Member, 대상자2: discord.Member, 지문: str = None):
    await submit(i, "영문", "보호", (대상자1, 대상자2), note=지문)


@tree.command(name="기도", description="영문 [기도]를 사용합니다")
@app_commands.describe(대상자="회복할 대상자", 지문=NOTE_DESC)
async def sk_gido(i: discord.Interaction, 대상자: discord.Member, 지문: str = None):
    await submit(i, "영문", "기도", (대상자,), note=지문)


@tree.command(name="분배", description="영문 [분배]를 사용합니다")
@app_commands.describe(대상자="체력을 나눠줄 대상자", 값="차감할 체력 (1~50)", 지문=NOTE_DESC)
async def sk_bunbae(i: discord.Interaction, 대상자: discord.Member, 값: int, 지문: str = None):
    await submit(i, "영문", "분배", (대상자,), value=값, note=지문)


@tree.command(name="정화", description="영문 [정화]를 사용합니다")
@app_commands.describe(대상자="정화할 대상자", 저주명="해제할 저주 이름 (비우면 첫 번째 저주)", 지문=NOTE_DESC)
async def sk_jeonghwa(i: discord.Interaction, 대상자: discord.Member, 저주명: str = None, 지문: str = None):
    await submit(i, "영문", "정화", (대상자,), name=저주명, note=지문)


@tree.command(name="역행", description="영문 [역행]을(를) 사용합니다")
@app_commands.describe(대상자="횟수를 늘려줄 대상", 영문명="대상자가 가진 영문 이름", 지문=NOTE_DESC)
async def sk_yeokhaeng(i: discord.Interaction, 대상자: discord.Member, 영문명: str, 지문: str = None):
    await submit(i, "영문", "역행", (대상자,), name=영문명, note=지문)


@tree.command(name="역살", description="영문 [역살]을(를) 사용합니다")
@app_commands.describe(값="차감할 체력 (1~50)", 지문=NOTE_DESC)
async def sk_yeoksal(i: discord.Interaction, 값: int, 지문: str = None):
    await submit(i, "영문", "역살", value=값, note=지문)


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
    await i.response.send_message("선택을 완료했습니다.", ephemeral=True)   # 본인에게만 (입력한 선택지가 채널에 노출되지 않게)
    await raid.channel.send(f"✅ **{p.name}**: 선택을 완료했습니다.")        # 채널 전체에는 선택 내용 없이 완료만 알림
    maybe_done(raid)


@tree.command(name="남은시간", description="이번 막의 남은 행동 시간을 확인합니다")
async def remaining(i: discord.Interaction):
    r = raids.get(i.channel_id)
    if not r or not r.open:
        return await i.response.send_message("지금은 행동 시간이 아닙니다.", ephemeral=True)
    waiting = [q.name for u, q in r.players.items()
               if q.alive and (u not in r.actions or (needs_choice(r) and u not in r.choices))]
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
    nxt = ""
    for key, nm in (("atk", "기본 공격"), ("def", "기본 방어")):
        lv = info[key]
        if lv < MAX_ENHANCE and (lv + 1) in info["table"]:
            nxt += f"\n다음 {nm} {lv + 1}강: 필요 영석 {info['table'][lv + 1][1]}개"
        elif lv >= MAX_ENHANCE:
            nxt += f"\n{nm}: 최대 강화"
    r = raids.get(i.channel_id)
    rp = r.players.get(i.user.id) if r else None
    mine = ""
    if rp:
        mine = "\n내 영문: " + (", ".join(f"{k}({'상시' if n < 0 else f'{n}/{max(n, rp.skill_max.get(k, n))}'})" for k, n in rp.skills.items()) or "없음")
    await i.followup.send(
        f"**{info['name']}**\n기본 공격 +{info['atk']}강 / 기본 방어 +{info['def']}강 (각 최대 {MAX_ENHANCE}강)\n"
        f"보유 영석 {info['stones']}개{nxt}{mine}")


# ===================== MPC 명령어 =====================
async def call_mpc(i: discord.Interaction, name: str, target=None):
    async def no(msg): await i.response.send_message(msg, ephemeral=True)

    raid = raids.get(i.channel_id)
    if not raid or not raid.open:
        return await no("이 채널에서는 지금 MPC를 부를 수 없습니다.")
    p = raid.players.get(i.user.id)
    if not p: return await no("이번 축귀 참가자가 아닙니다.")
    if not p.alive: return await no("행동 불능 상태입니다.")
    if raid.mpc:
        return await no(f"이번 막에는 이미 MPC [{', '.join(raid.mpc)}]을(를) 불렀습니다. 한 막에 MPC는 1개만 부를 수 있습니다.")
    if name in mpc_used:
        return await no(f"[{name}]은(는) 이미 사용되었습니다. (<#{mpc_used[name]}>)")
    tp = None
    if name == "고목봉춘":                                  # 대상자 확인은 사용 기록을 남기기 전에
        tp = raid.players.get(target.id) if target else None
        if not tp: return await no("대상자가 이번 축귀 참가자가 아닙니다.")
        if not tp.alive: return await no("행동 불능인 대상자는 지정할 수 없습니다.")
    # ↑ 확인과 ↓ 기록 사이에 await가 없어서, 두 조가 동시에 불러도 한쪽만 성공함
    mpc_used[name] = i.channel_id
    save_mpc()
    raid.mpc[name] = p.name
    flavor, effect = MPC_INFO[name]
    em = discord.Embed(title=f"🌟 MPC [{name}]", color=0xF1C40F)
    if name == "고목봉춘":                                  # 대상자 체력을 쓰는 즉시 전부 회복
        flavor = flavor.format(target=tp.name)
        before = tp.hp
        got = heal(tp, tp.max_hp - tp.hp)
        raid.mpc_note[name] = f"{flavor} (체력 +{got} → {hp_str(tp)})"
        em.description = f"*{flavor}*\n" + red(effect)
        em.add_field(name="회복 대상", value=tp.name, inline=True)
        em.add_field(name="회복 결과", value=f"{before} → {tp.hp}/{tp.max_hp} (+{got})", inline=True)
    else:
        em.description = f"*{flavor}*\n" + red(effect)
    em.set_author(name=p.name)
    em.set_footer(text="이 MPC는 이번 축귀에서 더 이상 부를 수 없습니다")
    await i.response.send_message(embed=em)
    maybe_done(raid)


def add_mpc(name):                                        # 대상자가 필요 없는 MPC
    @tree.command(name=name, description=f"MPC [{name}]을(를) 부릅니다 (전체 조에서 1번만)")
    async def _cmd(i: discord.Interaction):
        await call_mpc(i, name)


for _n in MPC_INFO:
    if _n != "고목봉춘":
        add_mpc(_n)


@tree.command(name="고목봉춘", description="MPC [고목봉춘]을(를) 부릅니다 (전체 조에서 1번만)")
@app_commands.describe(대상자="체력을 전부 회복할 대상자")
async def mpc_gomok(i: discord.Interaction, 대상자: discord.Member):
    await call_mpc(i, "고목봉춘", 대상자)


@축귀.command(name="mpc현황", description="MPC 사용 현황을 확인합니다")
async def mpc_status(i: discord.Interaction):
    lines = [f"- [{nm}] " + (f"사용됨 (<#{mpc_used[nm]}>)" if nm in mpc_used else "사용 가능") for nm in MPC_INFO]
    await i.response.send_message("\n".join(lines), ephemeral=True)


@축귀.command(name="mpc초기화", description="MPC 사용 기록을 초기화합니다 (새 축귀를 시작할 때)")
async def mpc_reset(i: discord.Interaction):
    mpc_used.clear()
    save_mpc()
    await i.response.send_message("MPC 사용 기록을 초기화했습니다. 모든 MPC를 다시 부를 수 있습니다.")


# ===================== 판정 =====================
JUDGE_TARGET = 7   # D10을 굴려 이 값 이상이면 성공


@tree.command(name="판정", description=f"D10을 굴려 {JUDGE_TARGET} 이상이면 성공입니다")
@app_commands.describe(지문=NOTE_DESC)
async def cmd_judge(i: discord.Interaction, 지문: str = None):
    v = d(10)
    ok = v >= JUDGE_TARGET
    note = clean_note(지문)
    r = raids.get(i.channel_id)
    rp = r.players.get(i.user.id) if r else None
    em = discord.Embed(title="🎲 판정",
                       description=(f"*{note}*\n" if note else "") + f"# {v}\n**{'성공' if ok else '실패'}**",
                       color=0x2ECC71 if ok else 0xE74C3C)
    em.set_author(name=rp.name if rp else i.user.display_name)
    em.set_footer(text=f"D10 · {JUDGE_TARGET} 이상 성공")
    await i.response.send_message(embed=em)


# ===================== 운세 (영석 지급) =====================
FORTUNE_DAILY_LIMIT = 3        # 1인당 하루(한국 시간 기준) 최대 횟수. 0이면 무제한
# (운세, 확률%, 영석 최소, 영석 최대, 한마디, 이모지, 카드 색)
FORTUNE_TABLE = [
    ("대길", 0.1, 10, 10, "하늘이 당신 편입니다. 오늘은 무엇을 해도 됩니다!", "🌟", 0xF1C40F),
    ("중길", 1,  8,  9,  "귀인이 곁에 있습니다. 좋은 일이 이어집니다.", "✨", 0x2ECC71),
    ("길",   5,  6,  7,  "순풍이 붑니다. 하던 대로만 하세요.", "🍀", 0x2ECC71),
    ("소길", 15, 4,  5,  "작은 행운이 따릅니다. 소소하게 좋은 하루.", "🌼", 0x3498DB),
    ("말길", 38.9, 3, 3,  "평범한 하루입니다. 서두르지 않으면 무난합니다.", "🌙", 0x3498DB),
    ("흉",   25, 2,  2,  "조금 불안한 기운이 있습니다. 조심하세요.", "☁️", 0x95A5A6),
    ("대흉", 15, 1,  1,  "오늘은 몸을 사리세요... 그래도 영석은 하나 챙겼습니다.", "⛈️", 0x7F8C8D),
]
KST = timezone(timedelta(hours=9))


def draw_fortune():
    pick = random.choices(FORTUNE_TABLE, weights=[x[1] for x in FORTUNE_TABLE])[0]
    return pick, random.randint(pick[2], pick[3])


def try_fortune(uid):
    """(결과 dict, 오류 문구) 반환. 영석 지급까지 처리"""
    with state_lock:
        info = read_player(uid)
        if not info:
            return None, "이번 명단에 없는 사용자입니다. 운영진에게 문의하세요."
        pick, n = draw_fortune()
        today = datetime.now(KST).strftime("%Y-%m-%d")
        if SHEET_WEBHOOK:                        # 시트에 직접 기록 (하루 한 번 검사도 시트에서 함)
            try:
                res = post_enhance({"action": "fortune", "uid": str(uid), "name": info["name"],
                                    "stones": n, "luck": pick[0], "date": today, "limit": FORTUNE_DAILY_LIMIT})
            except RuntimeError as e:
                return None, str(e).replace("시트 기록 실패: ", "")
            total, count = res["stones"], res.get("count")
        else:
            st = state.setdefault(str(uid), {"atk": 0, "def": 0, "spent": 0})
            rec = st.get("fortune")
            if isinstance(rec, dict) and rec.get("date") == today:
                used = rec.get("count", 0)
            elif rec == today:                       # 예전 형식(하루 한 번)으로 저장된 기록
                used = 1
            else:
                used = 0
            if FORTUNE_DAILY_LIMIT and used >= FORTUNE_DAILY_LIMIT:
                return None, f"오늘의 운세는 {FORTUNE_DAILY_LIMIT}번 모두 확인했습니다. 내일 다시 확인하세요."
            st["spent"] = st.get("spent", 0) - n     # 쓴 영석을 줄여서 지급
            st["fortune"] = {"date": today, "count": used + 1}
            save_state()
            total, count = info["stones"] + n, used + 1
        return {"name": info["name"], "pick": pick, "n": n, "total": total, "count": count}, None


@tree.command(name="운세", description="오늘의 운세를 확인하고 영석을 받습니다")
async def cmd_fortune(i: discord.Interaction):
    await i.response.defer()
    try:
        res, err = await asyncio.to_thread(try_fortune, i.user.id)
    except Exception as e:
        return await i.followup.send(f"운세 처리 중 오류: {e}", ephemeral=True)
    if err:
        return await i.followup.send(err, ephemeral=True)
    luck, _, _, _, msg, emoji, color = res["pick"]
    em = discord.Embed(title=f"{emoji} 오늘의 운세", description=f"# {luck}\n*{msg}*", color=color)
    em.add_field(name="획득 영석", value=f"+{res['n']}개", inline=True)
    em.add_field(name="보유 영석", value=f"{res['total']}개", inline=True)
    if FORTUNE_DAILY_LIMIT and res.get("count"):
        em.set_footer(text=f"오늘 {res['count']}/{FORTUNE_DAILY_LIMIT}번째 운세")
    em.set_author(name=res["name"])
    await i.followup.send(embed=em)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "sample":
        make_sample()
    else:
        client.run(os.environ["DISCORD_TOKEN"])
