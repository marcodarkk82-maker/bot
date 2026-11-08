# -*- coding: utf-8 -*-
"""TRONPICK AI SIGNAL BOT v2 - Crash/Mines statistik tahlil + Risk kalkulyatori (yagona fayl)."""
import asyncio
import html
import logging
import math
import os
import random
import re
import sqlite3
import sys
from contextlib import closing
from datetime import datetime, timezone
from logging.handlers import RotatingFileHandler

import numpy as np
from telegram import InlineKeyboardButton, InlineKeyboardMarkup as Kb, Update
from telegram.constants import ParseMode
from telegram.error import BadRequest, Forbidden, RetryAfter, TelegramError
from telegram.ext import (Application, CallbackQueryHandler, CommandHandler,
                          ContextTypes, MessageHandler, filters)


# ============================ SOZLAMALAR ============================
class config:
    try:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        BASE_DIR = os.getcwd()
    if not os.access(BASE_DIR, os.W_OK):
        BASE_DIR = os.path.expanduser("~")

    # Token: Railway Variables ichida TRONPICK_BOT_TOKEN nomi bilan saqlang.
    BOT_TOKEN = os.getenv("TRONPICK_BOT_TOKEN", "9099892515:AAFo4MN-vugKfxdNU2S93N2o0Oqm2kBffzI")

    # Admin Telegram ID lari (@userinfobot orqali bilib oling)
    ADMIN_IDS = [123456789]

    # Railway Volume ishlatsangiz: TRONPICK_DB_PATH=/data/tronpick_ai.db
    DB_PATH = os.getenv("TRONPICK_DB_PATH", os.path.join(BASE_DIR, "tronpick_ai.db"))
    LOG_PATH = os.path.join(BASE_DIR, "bot.log")

    # Nazariy hisoblar uchun faraz qilingan house edge (1% = 0.01)
    HOUSE_EDGE = 0.01


log = logging.getLogger("tronpick")

HEADER = "⚫🔵 <b>TRONPICK AI</b> 🔴⚫\n━━━━━━━━━━━━━━━━"
WARN = ("⚠️ <i>Bu bot statistik tahlil vositasi. Tasodifiy natijalarni oldindan "
        "bilmaydi, yutuqni kafolatlamaydi. Qimor moliyaviy xavf tug'diradi (18+).</i>")


def setup_logging():
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    try:
        fh = RotatingFileHandler(config.LOG_PATH, maxBytes=500_000, backupCount=2, encoding="utf-8")
        fh.setFormatter(fmt)
        root.addHandler(fh)
    except OSError:
        pass
    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    root.addHandler(sh)
    logging.getLogger("httpx").setLevel(logging.WARNING)


# ============================ DATABASE ============================
def now():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def _conn():
    c = sqlite3.connect(config.DB_PATH, timeout=15)
    c.row_factory = sqlite3.Row
    return c


def q_exec(sql, params=()):
    with closing(_conn()) as c:
        cur = c.execute(sql, params)
        c.commit()
        return cur.lastrowid


def q_all(sql, params=()):
    with closing(_conn()) as c:
        return c.execute(sql, params).fetchall()


def q_one(sql, params=()):
    with closing(_conn()) as c:
        return c.execute(sql, params).fetchone()


def init_db():
    with closing(_conn()) as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS users(
            user_id INTEGER PRIMARY KEY, username TEXT, first_name TEXT,
            first_seen TEXT, last_seen TEXT, actions INTEGER DEFAULT 0);
        CREATE TABLE IF NOT EXISTS crash_rounds(
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER,
            value REAL, created TEXT);
        CREATE TABLE IF NOT EXISTS mines_games(
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER,
            mines INTEGER, opened INTEGER, result TEXT, created TEXT);
        CREATE TABLE IF NOT EXISTS signals(
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER,
            kind TEXT, summary TEXT, created TEXT);
        CREATE INDEX IF NOT EXISTS ix_crash_user ON crash_rounds(user_id, id);
        CREATE INDEX IF NOT EXISTS ix_mines_user ON mines_games(user_id, id);
        CREATE INDEX IF NOT EXISTS ix_sig_user ON signals(user_id, id);
        """)
        c.commit()


def touch_user(update: Update):
    u = update.effective_user
    if not u:
        return
    t = now()
    q_exec("""INSERT INTO users(user_id,username,first_name,first_seen,last_seen,actions)
              VALUES(?,?,?,?,?,1)
              ON CONFLICT(user_id) DO UPDATE SET username=excluded.username,
              first_name=excluded.first_name, last_seen=excluded.last_seen,
              actions=users.actions+1""",
           (u.id, u.username or "", u.first_name or "", t, t))


def add_signal(uid, kind, summary):
    q_exec("INSERT INTO signals(user_id,kind,summary,created) VALUES(?,?,?,?)",
           (uid, kind, summary, now()))


def is_admin(uid):
    return uid in config.ADMIN_IDS


def esc(s):
    return html.escape(str(s or ""))


def B(text, data):
    return InlineKeyboardButton(text, callback_data=data)


HOME = [B("🏠 Menyu", "m:home")]


# ============================ UI HELPERS ============================
async def send(update: Update, text, kb=None):
    q = update.callback_query
    if q:
        try:
            await q.edit_message_text(text, reply_markup=kb, parse_mode=ParseMode.HTML)
        except BadRequest as e:
            if "not modified" not in str(e).lower():
                raise
    else:
        await update.effective_message.reply_text(text, reply_markup=kb, parse_mode=ParseMode.HTML)


def main_menu(uid):
    rows = [[B("🎮 CRASH AI", "m:crash"), B("💣 MINES AI", "m:mines")],
            [B("💎 GEMS 1.94x", "g:194"), B("💎 GEMS 1.46x", "g:146")],
            [B("🧮 RISK KALKULYATOR", "m:risk")],
            [B("📊 STATISTIKA", "m:stats"), B("📈 SIGNAL TARIXI", "m:hist")],
            [B("ℹ️ BOT HAQIDA", "m:about")]]
    if is_admin(uid):
        rows.append([B("🛠 ADMIN PANEL", "m:admin")])
    return Kb(rows)


def welcome_text(name):
    return (f"{HEADER}\n\n👋 Salom, <b>{esc(name)}</b>!\n\n"
            "🔵 Crash va Mines o'yinlari uchun <b>statistik tahlil</b> boti.\n"
            "🔴 Bo'limni tanlang:\n\n" + WARN)


# ============================ STATISTIK YORDAMCHILAR ============================
BUCKETS = [(1.0, 1.5, "1.00-1.49"), (1.5, 2.0, "1.50-1.99"), (2.0, 3.0, "2.00-2.99"),
           (3.0, 5.0, "3.00-4.99"), (5.0, 10.0, "5.00-9.99"), (10.0, float("inf"), "10.00+")]
NUM_RE = re.compile(r"\d+(?:[.,]\d+)?")


def bar(p, w=10):
    n = int(round(p * w))
    return "█" * n + "░" * (w - n)


def longest_run(mask):
    best = cur = 0
    for v in mask:
        cur = cur + 1 if v else 0
        best = max(best, cur)
    return best


def current_run(mask):
    c = 0
    for v in mask[::-1]:
        if v:
            c += 1
        else:
            break
    return c


def wilson(k, n, z=1.96):
    """Ehtimollik uchun 95% Wilson ishonch oralig'i."""
    if n == 0:
        return 0.0, 0.0
    p = k / n
    d = 1 + z * z / n
    c = p + z * z / (2 * n)
    r = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n))
    return (c - r) / d, (c + r) / d


def runs_test_z(mask):
    """Wald-Wolfowitz runs testi z-qiymati (None = ma'lumot yetarli emas)."""
    n = len(mask)
    n1 = int(np.sum(mask))
    n2 = n - n1
    if n < 20 or n1 == 0 or n2 == 0:
        return None
    runs = 1 + int(np.sum(mask[1:] != mask[:-1]))
    mu = 2 * n1 * n2 / n + 1
    var = (mu - 1) * (mu - 2) / (n - 1)
    if var <= 0:
        return None
    return (runs - mu) / math.sqrt(var)


# ============================ CRASH ============================
def crash_values(uid, n):
    rows = q_all("SELECT value FROM crash_rounds WHERE user_id=? ORDER BY id DESC LIMIT ?", (uid, n))
    return np.array([r["value"] for r in rows][::-1], dtype=float)


def crash_kb():
    return Kb([[B("➕ Natija kiritish", "c:add")],
               [B("📊 10", "c:an:10"), B("📊 20", "c:an:20"),
                B("📊 50", "c:an:50"), B("📊 100", "c:an:100")],
               [B("📉 Grafik", "c:chart"), B("📋 Jadval", "c:table"), B("🧪 Backtest", "c:bt")],
               [B("🗑 Tozalash", "c:clear")], HOME])


def crash_menu_text(uid):
    n = q_one("SELECT COUNT(*) c FROM crash_rounds WHERE user_id=?", (uid,))["c"]
    return (f"{HEADER}\n\n🎮 <b>CRASH AI</b>\n\n💾 Saqlangan raundlar: <b>{n}</b>\n\n"
            "➕ Natijalarni qo'lda kiriting (masalan: <code>1.52 3.40 2.01</code>)\n"
            "📊 Keyin oxirgi 10/20/50/100 raund tahlilini oling.\n"
            "🧪 Backtest: o'tgan ma'lumotda strategiyalar natijasi.")


def crash_text(uid, n):
    arr = crash_values(uid, n)
    cnt = len(arr)
    if cnt < 5:
        return None
    h = config.HOUSE_EDGE
    mask = arr < 2.0
    k2 = int(np.sum(arr >= 2.0))
    lo, hi = wilson(k2, cnt)
    p10, p25, p75, p90 = np.percentile(arr, [10, 25, 75, 90])
    L = [HEADER, "", "🚀 <b>CRASH AI ANALYSIS</b>", "",
         f"📊 Tahlil qilingan raundlar: <b>{cnt}</b>" + (f" (so'ralgan: {n})" if cnt < n else ""),
         f"📈 O'rtacha ko'rsatkich: <b>{np.mean(arr):.2f}x</b>",
         f"📍 Median: <b>{np.median(arr):.2f}x</b>",
         f"📏 Standart og'ish: <b>{np.std(arr):.2f}</b>",
         f"🔻 Min: {np.min(arr):.2f}x  🔺 Max: {np.max(arr):.2f}x",
         f"📐 P10 {p10:.2f} | P25 {p25:.2f} | P75 {p75:.2f} | P90 {p90:.2f}", "",
         "📦 <b>Taqsimot</b>"]
    for blo, bhi, label in BUCKETS:
        c = int(np.sum((arr >= blo) & (arr < bhi)))
        L.append(f"<code>{label:<10}</code> {bar(c / cnt)} {c / cnt * 100:4.1f}% ({c})")
    L += ["", "🎲 <b>Ehtimollik: empirik | nazariy | EV</b>"]
    for m in (1.5, 2, 3, 5, 10):
        emp = float(np.mean(arr >= m))
        th = min(1.0, (1 - h) / m)
        ev = emp * m - 1
        L.append(f"≥ {m}x: <b>{emp * 100:.1f}%</b> | {th * 100:.1f}% | EV {ev * 100:+.0f}%")
    L += [f"🎯 P(≥2x) 95% ishonch oralig'i: <b>{lo * 100:.0f}% - {hi * 100:.0f}%</b>",
          "", "🔬 <b>Tasodifiylik testlari</b>",
          f"Hozirgi 2x dan past seriya: <b>{current_run(mask)}</b> | eng uzuni: <b>{longest_run(mask)}</b>"]
    z = runs_test_z(mask)
    if z is None:
        L.append("Runs testi: kamida 20 raund kerak")
    else:
        verdict = "naqsh yo'q (tasodifiy)" if abs(z) < 1.96 else "chetlanish bor (kichik tanlanmada tasodif bo'lishi mumkin)"
        L.append(f"Runs testi z={z:+.2f}: {verdict}")
    if cnt >= 20 and np.std(arr) > 0:
        r = float(np.corrcoef(arr[:-1], arr[1:])[0, 1])
        band = 1.96 / math.sqrt(cnt)
        verdict = "bog'liqlik yo'q" if abs(r) < band else "kuchsiz chetlanish, ishonchli naqsh emas"
        L.append(f"Avtokorrelyatsiya r={r:+.2f} (chegara ±{band:.2f}): {verdict}")
    L += ["", f"<i>Nazariy qiymat {h * 100:.0f}% house edge farazi bilan: P(≥m) = {(1 - h):.2f}/m.</i>",
          "<i>Har raund mustaqil: o'tgan natijalar keyingisini belgilamaydi.</i>", "",
          "🎯 Keyingi raund: <b>aniq bashorat emas</b>",
          "⚠️ Ishonchlilik: <b>kafolatlanmaydi</b>"]
    add_signal(uid, "CRASH", f"n={cnt} o'rt={np.mean(arr):.2f} med={np.median(arr):.2f}")
    return "\n".join(L)


def crash_backtest_text(uid):
    arr = crash_values(uid, 500)
    n = len(arr)
    if n < 20:
        return None
    L = [HEADER, "", "🧪 <b>BACKTEST</b>", "",
         f"Ma'lumot: <b>{n}</b> raund. Tekis stavka (har raundda 1 birlik):", ""]
    for m in (1.5, 2, 3, 5):
        win = arr >= m
        roi = float(np.mean(np.where(win, m - 1, -1.0)))
        L.append(f"<b>{m}x</b>: urilish {win.mean() * 100:.0f}% | ROI {roi * 100:+.0f}% | "
                 f"eng uzun yutqazish seriyasi {longest_run(~win)}")
    m, stake, profit, peak, maxdd, maxstake = 2.0, 1.0, 0.0, 0.0, 0.0, 1.0
    for v in arr:
        if v >= m:
            profit += stake * (m - 1)
            stake = 1.0
        else:
            profit -= stake
            stake *= 2
        maxstake = max(maxstake, stake)
        peak = max(peak, profit)
        maxdd = max(maxdd, peak - profit)
    L += ["", "📈 <b>Martingale (x2 oshirish, 2.0x chiqish)</b>",
          f"Yakuniy natija: <b>{profit:+.1f}</b> birlik",
          f"Eng katta stavka: <b>{maxstake:.0f}</b> birlik | maks. tushish: <b>{maxdd:.1f}</b>",
          "", "<i>Bu o'tgan ma'lumotdagi natija. Kelajakni bildirmaydi: nazariy EV har doim "
          "house edge tufayli manfiy.</i>"]
    add_signal(uid, "BACKTEST", f"n={n} mart_x2 natija={profit:+.1f}")
    return "\n".join(L)


def crash_chart_text(uid):
    arr = crash_values(uid, 20)
    if len(arr) == 0:
        return None
    rows = []
    for i, v in enumerate(arr, 1):
        w = max(1, int(round(min(v, 10) / 10 * 12)))
        rows.append(f"{i:>2} {'█' * w:<12} {v:.2f}x")
    return (f"{HEADER}\n\n📉 <b>Oxirgi {len(arr)} raund grafigi</b>\n"
            f"<i>(10x dan yuqorisi kesilgan)</i>\n\n<pre>" + "\n".join(rows) + "</pre>")


def crash_table_text(uid):
    rows = q_all("SELECT id,value,created FROM crash_rounds WHERE user_id=? ORDER BY id DESC LIMIT 20", (uid,))
    if not rows:
        return None
    lines = [f"#{r['id']:<5} {r['value']:>8.2f}x  {r['created'][5:16]}" for r in rows]
    return f"{HEADER}\n\n📋 <b>Oxirgi 20 natija</b>\n\n<pre>" + "\n".join(lines) + "</pre>"


# ============================ RISK KALKULYATOR ============================
RISK_HELP = (f"{HEADER}\n\n🧮 <b>RISK KALKULYATOR</b>\n\n"
             "Stavkani har yutqazishda oshirish (Martingale) strategiyasining xavfini hisoblaydi "
             "va Monte-Karlo simulyatsiya qiladi (2000 ta sessiya x 300 stavka).\n\n"
             "5 ta sonni yuboring:\n"
             "<code>koeffitsient bazaviy_stavka oshirish balans max_qadam</code>\n\n"
             "Masalan: <code>1.94 1 2 100 10</code>\n"
             "• koeffitsient: yutganda to'lov (1.94x)\n"
             "• bazaviy_stavka: birinchi stavka\n"
             "• oshirish: yutqazgach stavka ko'paytmasi (2)\n"
             "• balans: umumiy mablag'\n"
             "• max_qadam: ketma-ket nechta yutqazishgacha (ixtiyoriy, 10)")


def martingale_sim(payout, base, mult, bankroll, max_steps, bets=300, sims=2000):
    h = config.HOUSE_EDGE
    p = min(0.999, (1 - h) / payout)
    rng = np.random.default_rng()
    bal = np.full(sims, float(bankroll))
    losses = np.zeros(sims, dtype=int)
    alive = np.ones(sims, dtype=bool)
    bust = np.zeros(sims, dtype=bool)
    for _ in range(bets):
        stake = base * np.power(mult, np.minimum(losses, 60))
        newbust = alive & ((stake > bal) | (losses >= max_steps))
        bust |= newbust
        alive &= ~newbust
        win = rng.random(sims) < p
        gain = np.where(win, stake * (payout - 1), -stake)
        bal = np.where(alive, bal + gain, bal)
        losses = np.where(alive, np.where(win, 0, losses + 1), losses)
    return p, bal, bust


def risk_text(uid, payout, base, mult, bankroll, max_steps):
    p, bal, bust = martingale_sim(payout, base, mult, bankroll, max_steps)
    q = 1 - p
    profit = bal - bankroll
    L = [HEADER, "", "🧮 <b>RISK NATIJASI</b>", "",
         f"⚙️ Koeff: {payout}x | baza: {base:g} | oshirish: x{mult:g} | balans: {bankroll:g} | max qadam: {max_steps}",
         f"🎲 Yutish ehtimoli (nazariy): <b>{p * 100:.1f}%</b>",
         f"📉 Har stavkada kutilgan natija: <b>{(p * payout - 1) * 100:+.2f}%</b> (house edge)", "",
         "⛓ <b>Ketma-ket yutqazish xavfi</b>", "<code>seriya  ehtimol    kerakli balans</code>"]
    for k in (5, 8, 10, 12):
        need = base * k if abs(mult - 1) < 1e-9 else base * (mult ** k - 1) / (mult - 1)
        L.append(f"<code>{k:>3} ta   {q ** k * 100:7.3f}%   {need:>12,.1f}</code>")
    L += ["", "🧪 <b>Simulyatsiya (300 stavka, 2000 sessiya)</b>",
          f"💥 Balans/limit tugagan sessiyalar: <b>{bust.mean() * 100:.1f}%</b>",
          f"📊 Yakuniy foyda: o'rtacha <b>{profit.mean():+.1f}</b> | median <b>{np.median(profit):+.1f}</b>",
          f"🔻 Eng yomon 5%: <b>{np.percentile(profit, 5):+.1f}</b>",
          f"✅ Foydada tugaganlar: <b>{np.mean(profit > 0) * 100:.1f}%</b>", "",
          "<i>Martingale kichik foydani ko'p marta beradi, lekin kam uchraydigan katta yo'qotish "
          "bilan hammasini qaytaradi. O'rtacha natija har doim manfiy.</i>"]
    add_signal(uid, "RISK", f"koef={payout} baza={base:g} x{mult:g} bust={bust.mean() * 100:.0f}%")
    return "\n".join(L)


def risk_kb():
    return Kb([[B("▶️ Namuna (1.94 1 2 100 10)", "r:demo")], [B("🔁 Yangi hisob", "m:risk")], HOME])


# ============================ GEMS STRATEGIYA SIMULYATORI ============================
GEMS = {"194": 1.94, "146": 1.46}
# (boshlang'ich stavka, yutqazgach oshirish %, max ketma-ket yutqazish, max stavka, maqsad foyda, max zarar)
GEMS_DEFAULT = (0.0001, 110.0, 10, 0.01, 0.001, 0.005)


def fmt(x):
    s = f"{x:.8f}".rstrip("0").rstrip(".")
    return "0" if s in ("", "-0") else s


def parse_nums(text):
    out = []
    for tok in NUM_RE.findall(text):
        try:
            out.append(float(tok.replace(",", ".")))
        except ValueError:
            pass
    return out


def gems_help(key):
    pay = GEMS[key]
    return (f"{HEADER}\n\n💎 <b>GEMS {pay}x — STRATEGIYA SIMULYATORI</b>\n\n"
            "Yutqazgach stavkani oshiruvchi strategiyani to'xtatish qoidalari bilan "
            "3000 ta sessiyada simulyatsiya qiladi.\n\n"
            "6 ta sonni yuboring:\n"
            "<code>boshlang'ich_stavka oshirish% max_yutqazish max_stavka maqsad_foyda max_zarar</code>\n\n"
            f"Standart: <code>{' '.join(fmt(v) for v in GEMS_DEFAULT)}</code>\n"
            "• oshirish%: 110 = yutqazgach stavka x1.10\n"
            "• yutganda stavka boshlang'ichga qaytadi\n"
            "• sessiya maqsad foyda, max zarar, max ketma-ket yutqazish yoki max stavkada to'xtaydi\n\n"
            "<i>Bu faqat simulyatsiya: bot saytda avtomatik stavka qilmaydi.</i>")


def gems_kb(key):
    return Kb([[B("▶️ Standart sozlama", f"g:run:{key}")],
               [B("🔁 Qayta kiritish", f"g:{key}")], HOME])


def gems_sim(payout, init, mult, max_losses, max_bet, target, max_loss, sessions=3000, max_bets=3000):
    p = min(0.999, (1 - config.HOUSE_EDGE) / payout)
    rng = np.random.default_rng()
    profit = np.zeros(sessions)
    stake = np.full(sessions, float(init))
    losses = np.zeros(sessions, dtype=int)
    bets = np.zeros(sessions, dtype=int)
    outcome = np.zeros(sessions, dtype=int)  # 0 davom, 1 maqsad, 2 zarar, 3 seriya, 4 max stavka, 5 tugamagan
    for _ in range(max_bets):
        act = outcome == 0
        if not act.any():
            break
        win = rng.random(sessions) < p
        profit = np.where(act, profit + np.where(win, stake * (payout - 1), -stake), profit)
        bets += act
        losses = np.where(act, np.where(win, 0, losses + 1), losses)
        stake = np.where(act, np.where(win, init, stake * mult), stake)
        t = act & (profit >= target)
        l = act & ~t & (profit <= -max_loss)
        s_ = act & ~t & ~l & (losses >= max_losses)
        b = act & ~t & ~l & ~s_ & (stake > max_bet)
        outcome[t] = 1
        outcome[l] = 2
        outcome[s_] = 3
        outcome[b] = 4
    outcome[outcome == 0] = 5
    return p, outcome, profit, bets


def gems_text(uid, key, prm):
    pay = GEMS[key]
    init, mult_pct, max_losses, max_bet, target, max_loss = prm
    mult = mult_pct / 100.0
    p, out, profit, bets = gems_sim(pay, init, mult, int(max_losses), max_bet, target, max_loss)
    n = len(out)
    share = lambda c: float(np.mean(out == c)) * 100
    q = 1 - p
    streak_cost = sum(init * mult ** i for i in range(int(max_losses)))
    L = [HEADER, "", f"💎 <b>GEMS {pay}x — NATIJA</b>", "",
         f"⚙️ Stavka {fmt(init)} | oshirish {mult_pct:g}% | max yutqazish {int(max_losses)}",
         f"⚙️ Max stavka {fmt(max_bet)} | maqsad +{fmt(target)} | max zarar -{fmt(max_loss)}",
         f"🎲 Yutish ehtimoli (nazariy faraz): <b>{p * 100:.1f}%</b>",
         f"📉 Har stavkada kutilgan natija: <b>{(p * pay - 1) * 100:+.2f}%</b>", "",
         f"⛓ {int(max_losses)} ta ketma-ket yutqazish ehtimoli (bir urinishda): <b>{q ** int(max_losses) * 100:.3f}%</b>",
         f"💸 Shu seriyada sarflanadigan jami: <b>{fmt(streak_cost)}</b>", "",
         f"🧪 <b>Sessiyalar ({n})</b>",
         f"✅ Maqsadga yetdi: <b>{share(1):.1f}%</b>",
         f"🛑 Zarar limiti: <b>{share(2):.1f}%</b>",
         f"⛓ Yutqazish seriyasi limiti: <b>{share(3):.1f}%</b>",
         f"📈 Max stavka limiti: <b>{share(4):.1f}%</b>",
         f"⏱ Tugamagan: <b>{share(5):.1f}%</b>", "",
         f"💰 Natija: o'rtacha <b>{fmt(float(profit.mean()))}</b> | median <b>{fmt(float(np.median(profit)))}</b>",
         f"🔻 Eng yomon 5%: <b>{fmt(float(np.percentile(profit, 5)))}</b>",
         f"🔢 O'rtacha stavkalar soni: <b>{bets.mean():.0f}</b>", "",
         "<i>Maqsadga yetish ko'p bo'lsa ham, kam uchraydigan katta yo'qotish o'rtachani manfiy qiladi. "
         "Haqiqiy o'yin ehtimoli farazdan farq qilishi mumkin.</i>"]
    add_signal(uid, "GEMS" + key, f"maqsad={share(1):.0f}% zarar={share(2):.0f}% o'rt={fmt(float(profit.mean()))}")
    return "\n".join(L)


# ============================ MINES ============================
def survive_prob(m, k, j):
    """k ta ochilgandan keyin keyingi j ta katak xavfsiz bo'lish ehtimoli."""
    total, safe = 25 - k, 25 - m - k
    if j > safe or j < 0:
        return 0.0
    return math.comb(safe, j) / math.comb(total, j)


def mines_pick_kb():
    nums = [1, 2, 3, 4, 5, 6, 8, 10, 12, 15, 18, 20, 22, 24]
    btns = [B(f"💣 {n}", f"mn:{n}") for n in nums]
    rows = [btns[i:i + 5] for i in range(0, len(btns), 5)]
    rows.append(HOME)
    return Kb(rows)


def new_game(context, m):
    context.user_data["mines"] = {"m": m, "mines": set(random.sample(range(25), m)),
                                  "opened": set(), "over": False, "showp": False,
                                  "hit": None, "result": None}


def finish_game(uid, g):
    q_exec("INSERT INTO mines_games(user_id,mines,opened,result,created) VALUES(?,?,?,?,?)",
           (uid, g["m"], len(g["opened"]), g["result"], now()))
    add_signal(uid, "MINES", f"mina={g['m']} ochildi={len(g['opened'])} natija={g['result']}")


def mines_text(g):
    m, k = g["m"], len(g["opened"])
    left = 25 - k
    L = [HEADER, "", "💣 <b>MINES AI</b> · <i>DEMO REJIM</i>", "",
         f"💣 Minalar: <b>{m}</b> | 💎 Ochilgan: <b>{k}</b> | ⬛ Yopiq: <b>{left}</b>"]
    if g["over"]:
        L.append("")
        L.append("🏆 <b>Yutuq (demo)!</b>" if g["result"] == "WIN" else "💥 <b>Mina topildi (demo)!</b>")
    else:
        pm = m / left
        L += ["", "📐 <b>Matematik ehtimollik</b>",
              f"Har bir yopiq katak: mina <b>{pm * 100:.1f}%</b> | xavfsiz <b>{(1 - pm) * 100:.1f}%</b>"]
        for j in (1, 2, 3):
            if j <= left - m:
                L.append(f"Keyingi {j} ta katak xavfsiz: {survive_prob(m, k, j) * 100:.1f}%")
    if k > 0:
        ps = survive_prob(m, 0, k)
        L.append(f"🧮 Nazariy koeffitsient ({k} qadam): <b>x{(1 - config.HOUSE_EDGE) / ps:.2f}</b>")
    L += ["", "⚠️ <i>Maydon bot tomonidan tasodifiy yaratilgan DEMO. Bu haqiqiy "
          "TRONPICK.IO o'yinidagi minalarni aniqlamaydi va xavfsiz katakni kafolatlamaydi. "
          "Boshlang'ich holatda barcha kataklar ehtimoli teng.</i>"]
    return "\n".join(L)


def mines_kb(g):
    m, k = g["m"], len(g["opened"])
    p = m / (25 - k)
    rows = []
    for r in range(5):
        row = []
        for c in range(5):
            i = r * 5 + c
            if g["over"]:
                if i in g["mines"]:
                    t = "💥" if i == g["hit"] else "💣"
                elif i in g["opened"]:
                    t = "💎"
                else:
                    t = "▫️"
            elif i in g["opened"]:
                t = "💎"
            elif g["showp"]:
                t = f"{p * 100:.0f}%"
            else:
                t = "⬛"
            row.append(B(t, f"mc:{i}"))
        rows.append(row)
    if not g["over"]:
        rows.append([B("💰 Yig'ish", "mx:cash"), B("🔢 Ehtimollik", "mx:prob")])
    rows.append([B("🔄 Qayta", "mx:new"), B("💣 Mines soni", "mx:pick")])
    rows.append(HOME)
    return Kb(rows)


async def mines_click(update, context, idx):
    q = update.callback_query
    uid = q.from_user.id
    g = context.user_data.get("mines")
    if not g:
        await q.answer("O'yin topilmadi. Mines AI ni qayta oching.", show_alert=True)
        return
    if g["over"] or idx in g["opened"]:
        await q.answer()
        return
    await q.answer()
    if idx in g["mines"]:
        g.update(over=True, hit=idx, result="LOSE")
        finish_game(uid, g)
    else:
        g["opened"].add(idx)
        if len(g["opened"]) == 25 - g["m"]:
            g.update(over=True, result="WIN")
            finish_game(uid, g)
    await send(update, mines_text(g), mines_kb(g))


# ============================ STATISTIKA ============================
def stats_text(uid):
    sa = q_one("SELECT COUNT(*) c FROM signals")["c"]
    su = q_one("SELECT COUNT(*) c FROM signals WHERE user_id=?", (uid,))["c"]
    cr = q_one("SELECT COUNT(*) c, AVG(value) a, MAX(value) mx FROM crash_rounds WHERE user_id=?", (uid,))
    mi = q_one("SELECT COUNT(*) c, COALESCE(SUM(result='WIN'),0) w FROM mines_games WHERE user_id=?", (uid,))
    u = q_one("SELECT * FROM users WHERE user_id=?", (uid,))
    L = [HEADER, "", "📊 <b>STATISTIKA</b>", "",
         f"🧠 Tahlillar: sizniki <b>{su}</b> | umumiy <b>{sa}</b>", "",
         "🎮 <b>Crash</b>",
         f"Raundlar: {cr['c']} | O'rtacha: {(cr['a'] or 0):.2f}x | Max: {(cr['mx'] or 0):.2f}x"]
    last = q_all("SELECT value FROM crash_rounds WHERE user_id=? ORDER BY id DESC LIMIT 5", (uid,))
    if last:
        L.append("Oxirgi: " + " · ".join(f"{r['value']:.2f}x" for r in last))
    wr = (mi["w"] / mi["c"] * 100) if mi["c"] else 0
    L += ["", "💣 <b>Mines (demo)</b>",
          f"O'yinlar: {mi['c']} | Yutuq: {mi['w']} | Yutqazish: {mi['c'] - mi['w']} | Win: {wr:.0f}%"]
    for r in q_all("SELECT mines,opened,result FROM mines_games WHERE user_id=? ORDER BY id DESC LIMIT 5", (uid,)):
        L.append(f"{'🏆' if r['result'] == 'WIN' else '💥'} mina {r['mines']} · ochildi {r['opened']}")
    if u:
        L += ["", "👤 <b>Faoliyat</b>",
              f"Harakatlar: {u['actions']}", f"Birinchi: {u['first_seen']}", f"Oxirgi: {u['last_seen']}"]
    days = {}
    for table, key in (("signals", "a"), ("crash_rounds", "c"), ("mines_games", "g")):
        for r in q_all(f"SELECT substr(created,1,10) d, COUNT(*) n FROM {table} "
                       "WHERE user_id=? GROUP BY d ORDER BY d DESC LIMIT 7", (uid,)):
            days.setdefault(r["d"], {"a": 0, "c": 0, "g": 0})[key] = r["n"]
    if days:
        L += ["", "📅 <b>Kunlik</b> (tahlil/crash/mines)"]
        for d in sorted(days, reverse=True)[:7]:
            v = days[d]
            L.append(f"<code>{d}</code> {v['a']}/{v['c']}/{v['g']}")
    return "\n".join(L)


def history_text(uid):
    rows = q_all("SELECT kind,summary,created FROM signals WHERE user_id=? ORDER BY id DESC LIMIT 15", (uid,))
    if not rows:
        return f"{HEADER}\n\n📈 <b>SIGNAL TARIXI</b>\n\nHozircha tahlil yo'q."
    L = [f"{r['created'][5:16]} · <b>{r['kind']}</b> · {esc(r['summary'])}" for r in rows]
    return f"{HEADER}\n\n📈 <b>SIGNAL TARIXI</b> (oxirgi 15)\n\n" + "\n".join(L)


ABOUT = (f"{HEADER}\n\nℹ️ <b>BOT HAQIDA</b>\n\n"
         "🔵 <b>TRONPICK AI</b> — Crash va Mines uchun statistik tahlil boti.\n\n"
         "• Crash: o'rtacha, median, persentil, taqsimot, ishonch oralig'i, tasodifiylik testlari, backtest.\n"
         "• Mines: demo maydon va matematik ehtimollik (kombinatorika).\n"
         "• Gems 1.94x va Gems 1.46x: stavka oshirish strategiyasining to'xtatish qoidalari bilan sessiya simulyatori.\n"
         "• Risk kalkulyator: stavka oshirish strategiyasining xavfini Monte-Karlo bilan hisoblaydi.\n\n"
         "🔴 <b>Muhim:</b> bot tasodifiy natijani oldindan bilmaydi, haqiqiy o'yindagi "
         "minalarni aniqlay olmaydi va hech qanday yutuq kafolatini bermaydi. "
         "Bu rasmiy TRONPICK.IO mahsuloti emas.\n\n" + WARN)


# ============================ ADMIN ============================
def admin_kb():
    return Kb([[B("👥 Foydalanuvchilar soni", "ad:count"), B("📊 Bot statistikasi", "ad:stats")],
               [B("📋 Ro'yxat", "ad:list"), B("📢 Xabar yuborish", "ad:bcast")],
               [B("🗄 Ma'lumotlar bazasi", "ad:db")], HOME])


def admin_back():
    return Kb([[B("⬅️ Admin panel", "m:admin")]])


def db_info_text():
    size = os.path.getsize(config.DB_PATH) / 1024 if os.path.exists(config.DB_PATH) else 0
    L = [f"{HEADER}\n\n🗄 <b>Ma'lumotlar bazasi</b>\n", f"Hajm: {size:.1f} KB"]
    for t in ("users", "crash_rounds", "mines_games", "signals"):
        L.append(f"{t}: {q_one(f'SELECT COUNT(*) c FROM {t}')['c']}")
    return "\n".join(L)


def db_kb():
    return Kb([[B("💾 Backup", "ad:backup"), B("🧹 VACUUM", "ad:vacuum")],
               [B("🗑 Signal tarixini tozalash", "ad:clr")],
               [B("⬅️ Admin panel", "m:admin")]])


def make_backup():
    path = os.path.join(config.BASE_DIR, f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
    with closing(sqlite3.connect(config.DB_PATH)) as src, closing(sqlite3.connect(path)) as dst:
        src.backup(dst)
    return path


async def do_broadcast(context, text):
    ok = fail = 0
    for r in q_all("SELECT user_id FROM users"):
        for attempt in range(2):
            try:
                await context.bot.send_message(r["user_id"], f"📢 TRONPICK AI\n\n{text}")
                ok += 1
                break
            except RetryAfter as e:
                await asyncio.sleep(e.retry_after + 1)
            except (Forbidden, BadRequest):
                fail += 1
                break
            except TelegramError as e:
                log.warning("Broadcast xato: %s", e)
                fail += 1
                break
        await asyncio.sleep(0.06)
    return ok, fail


# ============================ HANDLERS ============================
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    touch_user(update)
    context.user_data.clear()
    u = update.effective_user
    await send(update, welcome_text(u.first_name), main_menu(u.id))


async def cmd_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    touch_user(update)
    if not is_admin(update.effective_user.id):
        log.warning("Ruxsatsiz /admin: %s", update.effective_user.id)
        await update.effective_message.reply_text("⛔ Bu bo'lim faqat administrator uchun.")
        return
    await send(update, f"{HEADER}\n\n🛠 <b>ADMIN PANEL</b>", admin_kb())


async def cmd_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("state", None)
    await send(update, "❌ Bekor qilindi.", main_menu(update.effective_user.id))


async def on_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    uid = q.from_user.id
    data = q.data or ""
    touch_user(update)

    if data.startswith("ad:") and not is_admin(uid):
        log.warning("Ruxsatsiz admin urinish: %s", uid)
        await q.answer("⛔ Ruxsat yo'q", show_alert=True)
        return
    if data.startswith("mc:"):
        try:
            idx = int(data[3:])
        except ValueError:
            await q.answer()
            return
        if 0 <= idx < 25:
            await mines_click(update, context, idx)
        return
    if data == "mx:cash":
        g = context.user_data.get("mines")
        if not g or g["over"]:
            await q.answer()
        elif not g["opened"]:
            await q.answer("Avval kamida 1 ta katak oching.", show_alert=True)
        else:
            await q.answer()
            g.update(over=True, result="WIN")
            finish_game(uid, g)
            await send(update, mines_text(g), mines_kb(g))
        return

    await q.answer()
    context.user_data.pop("state", None)

    if data == "m:home":
        await send(update, welcome_text(q.from_user.first_name), main_menu(uid))
    elif data == "m:crash":
        await send(update, crash_menu_text(uid), crash_kb())
    elif data == "m:risk":
        context.user_data["state"] = "risk_input"
        await send(update, RISK_HELP, risk_kb())
    elif data in ("g:194", "g:146"):
        key = data[2:]
        context.user_data["state"] = "gems_input"
        context.user_data["gems_key"] = key
        await send(update, gems_help(key), gems_kb(key))
    elif data.startswith("g:run:"):
        key = data[6:]
        if key in GEMS:
            await send(update, gems_text(uid, key, GEMS_DEFAULT), gems_kb(key))
    elif data == "r:demo":
        await send(update, risk_text(uid, 1.94, 1.0, 2.0, 100.0, 10), risk_kb())
    elif data == "m:mines" or data == "mx:pick":
        await send(update, f"{HEADER}\n\n💣 <b>MINES AI</b>\n\nMinalar sonini tanlang (5x5 = 25 katak):\n\n"
                           "<i>Demo rejim: bot tasodifiy maydon yaratadi.</i>", mines_pick_kb())
    elif data.startswith("mn:"):
        try:
            m = int(data[3:])
        except ValueError:
            return
        if 1 <= m <= 24:
            new_game(context, m)
            g = context.user_data["mines"]
            await send(update, mines_text(g), mines_kb(g))
    elif data == "mx:new":
        g = context.user_data.get("mines")
        if g:
            new_game(context, g["m"])
            g = context.user_data["mines"]
            await send(update, mines_text(g), mines_kb(g))
        else:
            await send(update, "Minalar sonini tanlang:", mines_pick_kb())
    elif data == "mx:prob":
        g = context.user_data.get("mines")
        if g:
            g["showp"] = not g["showp"]
            await send(update, mines_text(g), mines_kb(g))
    elif data == "m:stats":
        await send(update, stats_text(uid), Kb([HOME]))
    elif data == "m:hist":
        await send(update, history_text(uid), Kb([HOME]))
    elif data == "m:about":
        await send(update, ABOUT, Kb([HOME]))
    elif data == "m:admin":
        if is_admin(uid):
            await send(update, f"{HEADER}\n\n🛠 <b>ADMIN PANEL</b>", admin_kb())
    # ---- Crash ----
    elif data == "c:add":
        context.user_data["state"] = "crash_input"
        await send(update, f"{HEADER}\n\n➕ <b>Natija kiritish</b>\n\n"
                           "Crash natijalarini yuboring (bo'sh joy bilan ajrating):\n"
                           "<code>1.52 3.40 2.01 1.00 12.5</code>\n\n"
                           "Bir xabarda 200 tagacha. Tugagach tugmani bosing.",
                   Kb([[B("📊 Tahlil (20)", "c:an:20")], [B("⬅️ Orqaga", "m:crash")]]))
    elif data.startswith("c:an:"):
        try:
            n = int(data[5:])
        except ValueError:
            return
        if n in (10, 20, 50, 100):
            t = crash_text(uid, n)
            await send(update, t or f"{HEADER}\n\n⚠️ Tahlil uchun kamida <b>5</b> ta natija kerak.", crash_kb())
    elif data == "c:bt":
        t = crash_backtest_text(uid)
        await send(update, t or f"{HEADER}\n\n⚠️ Backtest uchun kamida <b>20</b> ta natija kerak.", crash_kb())
    elif data == "c:chart":
        t = crash_chart_text(uid)
        await send(update, t or f"{HEADER}\n\nHozircha natija yo'q.", crash_kb())
    elif data == "c:table":
        t = crash_table_text(uid)
        await send(update, t or f"{HEADER}\n\nHozircha natija yo'q.", crash_kb())
    elif data == "c:clear":
        await send(update, "🗑 Barcha Crash natijalaringiz o'chiriladi. Davom etasizmi?",
                   Kb([[B("✅ Ha", "c:clearyes"), B("❌ Yo'q", "m:crash")]]))
    elif data == "c:clearyes":
        q_exec("DELETE FROM crash_rounds WHERE user_id=?", (uid,))
        await send(update, crash_menu_text(uid), crash_kb())
    # ---- Admin ----
    elif data == "ad:count":
        total = q_one("SELECT COUNT(*) c FROM users")["c"]
        today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        act = q_one("SELECT COUNT(*) c FROM users WHERE substr(last_seen,1,10)=?", (today,))["c"]
        new = q_one("SELECT COUNT(*) c FROM users WHERE substr(first_seen,1,10)=?", (today,))["c"]
        await send(update, f"{HEADER}\n\n👥 Jami: <b>{total}</b>\n🟢 Bugun faol: <b>{act}</b>\n🆕 Bugun yangi: <b>{new}</b>", admin_back())
    elif data == "ad:stats":
        L = [f"{HEADER}\n\n📊 <b>Bot statistikasi</b>\n"]
        for label, t in (("Foydalanuvchilar", "users"), ("Crash raundlar", "crash_rounds"),
                         ("Mines o'yinlar", "mines_games"), ("Tahlillar", "signals")):
            L.append(f"{label}: <b>{q_one(f'SELECT COUNT(*) c FROM {t}')['c']}</b>")
        await send(update, "\n".join(L), admin_back())
    elif data == "ad:list":
        rows = q_all("SELECT user_id,username,first_name,actions FROM users ORDER BY last_seen DESC LIMIT 30")
        L = [f"{r['user_id']} · {esc(r['first_name'])} " + (f"@{esc(r['username'])} " if r['username'] else "")
             + f"· {r['actions']}" for r in rows]
        await send(update, f"{HEADER}\n\n📋 <b>Oxirgi 30 foydalanuvchi</b>\n\n" + ("\n".join(L) or "Bo'sh"), admin_back())
    elif data == "ad:bcast":
        context.user_data["state"] = "bcast"
        await send(update, f"{HEADER}\n\n📢 Barcha foydalanuvchilarga yuboriladigan matnni yozing.\nBekor qilish: /cancel",
                   admin_back())
    elif data == "ad:db":
        await send(update, db_info_text(), db_kb())
    elif data == "ad:backup":
        path = None
        try:
            path = make_backup()
            with open(path, "rb") as f:
                await context.bot.send_document(q.message.chat_id, f, filename=os.path.basename(path))
        except Exception:
            log.exception("Backup xatosi")
            await context.bot.send_message(q.message.chat_id, "❌ Backup xatosi, bot.log ni tekshiring.")
        finally:
            if path and os.path.exists(path):
                os.remove(path)
    elif data == "ad:vacuum":
        with closing(_conn()) as c:
            c.execute("VACUUM")
        await send(update, "🧹 VACUUM bajarildi.\n\n" + db_info_text(), db_kb())
    elif data == "ad:clr":
        await send(update, "🗑 Barcha foydalanuvchilarning signal tarixi o'chiriladi. Tasdiqlaysizmi?",
                   Kb([[B("✅ Ha", "ad:clryes"), B("❌ Yo'q", "ad:db")]]))
    elif data == "ad:clryes":
        q_exec("DELETE FROM signals")
        await send(update, "✅ Signal tarixi tozalandi.\n\n" + db_info_text(), db_kb())


async def on_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    touch_user(update)
    uid = update.effective_user.id
    text = update.effective_message.text or ""
    st = context.user_data.get("state")

    if st == "crash_input":
        vals = []
        for tok in NUM_RE.findall(text):
            try:
                v = float(tok.replace(",", "."))
            except ValueError:
                continue
            if 1.0 <= v <= 1_000_000:
                vals.append(round(v, 2))
        vals = vals[:200]
        if not vals:
            await update.effective_message.reply_text("❌ To'g'ri son topilmadi (kamida 1.00). Qayta urinib ko'ring.")
            return
        t = now()
        with closing(_conn()) as c:
            c.executemany("INSERT INTO crash_rounds(user_id,value,created) VALUES(?,?,?)",
                          [(uid, v, t) for v in vals])
            c.commit()
        await update.effective_message.reply_text(
            f"✅ <b>{len(vals)}</b> ta natija saqlandi.\nYana kiriting yoki tahlilni tanlang:",
            parse_mode=ParseMode.HTML, reply_markup=crash_kb())
    elif st == "risk_input":
        nums = []
        for tok in NUM_RE.findall(text):
            try:
                nums.append(float(tok.replace(",", ".")))
            except ValueError:
                pass
        if len(nums) < 4:
            await update.effective_message.reply_text(
                "❌ Kamida 4 ta son kerak: koeffitsient, baza stavka, oshirish, balans.\n"
                "Masalan: 1.94 1 2 100 10")
            return
        payout, base, mult, bankroll = nums[:4]
        max_steps = int(nums[4]) if len(nums) > 4 else 10
        if not (1.01 <= payout <= 100 and base > 0 and 1.01 <= mult <= 10
                and bankroll >= base and 1 <= max_steps <= 30):
            await update.effective_message.reply_text(
                "❌ Qiymatlar oralig'ida emas.\nKoeff 1.01-100, oshirish 1.01-10, "
                "balans ≥ baza stavka, max qadam 1-30.")
            return
        await send(update, risk_text(uid, payout, base, mult, bankroll, max_steps), risk_kb())
    elif st == "gems_input":
        key = context.user_data.get("gems_key", "194")
        nums = parse_nums(text)
        if len(nums) < 6:
            await update.effective_message.reply_text(
                "❌ 6 ta son kerak: boshlang'ich_stavka oshirish% max_yutqazish max_stavka maqsad max_zarar.\n"
                "Standart: " + " ".join(fmt(v) for v in GEMS_DEFAULT))
            return
        init, mult_pct, ml, mb, tg, mloss = nums[:6]
        if not (init > 0 and 101 <= mult_pct <= 1000 and 1 <= ml <= 30 and mb > init and tg > 0 and mloss > 0):
            await update.effective_message.reply_text(
                "❌ Qiymatlar noto'g'ri.\nStavka > 0, oshirish 101-1000%, max yutqazish 1-30, "
                "max stavka > boshlang'ich stavka, maqsad va max zarar > 0.")
            return
        await send(update, gems_text(uid, key, (init, mult_pct, int(ml), mb, tg, mloss)), gems_kb(key))
    elif st == "bcast" and is_admin(uid):
        context.user_data.pop("state", None)
        await update.effective_message.reply_text("⏳ Yuborilmoqda...")
        ok, fail = await do_broadcast(context, text)
        log.info("Broadcast admin=%s ok=%s fail=%s", uid, ok, fail)
        await update.effective_message.reply_text(f"✅ Yuborildi: {ok}\n❌ Xato: {fail}", reply_markup=admin_back())
    else:
        await update.effective_message.reply_text("Menyuni ochish uchun /start bosing.")


async def on_error(update, context: ContextTypes.DEFAULT_TYPE):
    log.error("Xatolik yuz berdi", exc_info=context.error)
    try:
        if isinstance(update, Update) and update.effective_chat:
            await context.bot.send_message(update.effective_chat.id,
                                           "⚠️ Vaqtinchalik xatolik. /start bilan qayta urinib ko'ring.")
    except TelegramError:
        pass


def main():
    setup_logging()
    if not config.BOT_TOKEN or "PASTE" in config.BOT_TOKEN:
        print("❌ TRONPICK_BOT_TOKEN o'zgaruvchisini o'rnating (yoki class config ichiga yozing).")
        sys.exit(1)
    init_db()
    app = Application.builder().token(config.BOT_TOKEN).build()
    app.add_handler(CommandHandler(["start", "menu"], cmd_start))
    app.add_handler(CommandHandler("admin", cmd_admin))
    app.add_handler(CommandHandler("cancel", cmd_cancel))
    app.add_handler(CallbackQueryHandler(on_button))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.add_error_handler(on_error)
    log.info("TRONPICK AI bot ishga tushdi")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
