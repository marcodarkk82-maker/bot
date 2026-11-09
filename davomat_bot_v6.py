"""
70-maktab davomat boti — ReplyKeyboard versiyasi

/adminhide -> admin login/parol orqali kiradi; yagona bosh admin Telegram ID bilan belgilanadi.
Admin panelda: sinf qo'shish, har bir sinfga admin va o'quvchi qo'shish.
Fan o'qituvchilari mini ilovasi: hamma ko'radi, tahrirlash faqat bosh adminda (kod ichida, index.html avtomatik yaratiladi).
Barcha navigatsiya pastki menyu (ReplyKeyboard) orqali ishlaydi.

O'rnatish (Pydroid 3 -> Pip):
    pyTelegramBotAPI
    openpyxl
    python-docx      (Word hisobot uchun; bo'lmasa faqat Excel ishlaydi)
"""
import io
import os
import copy
import re
import json
import base64
import zlib
import secrets
import threading
import time
from datetime import datetime, timedelta, timezone, time as dtime

import telebot
from telebot import types
from telebot import apihelper
from telebot.apihelper import ApiTelegramException
from concurrent.futures import ThreadPoolExecutor

try:
    import requests
    NET_ERRORS = (requests.exceptions.RequestException, OSError)
except Exception:
    NET_ERRORS = (OSError,)
from openpyxl import Workbook, load_workbook

try:
    from docx import Document
    HAS_DOCX = True
except Exception:
    HAS_DOCX = False

# ================= CONFIG =================
# @BotFather dan olingan YANGI tokenni shu yerga qo'ying
TOKEN = os.environ.get("BOT_TOKEN", "").strip()  # Hosting Variables -> BOT_TOKEN
if not TOKEN:
    raise SystemExit("BOT_TOKEN topilmadi. Hosting Variables bo'limiga token kiriting.")

# Ustozlar (admin). Ariza shu odamlarga yuboriladi.
TEACHERS = {
    "teacher": {"label": "Sinf rahbari", "emoji": "👨‍🏫", "id": 1732314424},
    "school": {"label": "Maktab nazoratchisi", "emoji": "🏫", "id": 8355611778},
}
SUPER_ADMIN_ID = 8355611778
ADMINS = [SUPER_ADMIN_ID]  # yagona bosh admin

# Ustozlar BOSHQA telefondan /adminhide orqali kirishi uchun  login: parol
# (PAROLLARNI ALBATTA O'ZGARTIRING!)
ADMIN_ACCOUNTS = {}  # global super-loginlar o'chirilgan

# Birinchi sinf: eski ma'lumotlar (o'quvchilar, arizalar) shu sinfga ko'chadi.
# Yangi sinflar admin paneldan qo'shiladi ("🏫 Sinflar" -> "➕ Sinf qo'shish").
CLASS_NAME = "11-B"

# Mini ilova (fan o'qituvchilari): index.html ni HTTPS hostingga (GitHub Pages, Vercel...) joylang
# va shu yerga to'liq manzilini yozing. Masalan: "https://akmal.github.io/davomat/"
_RW_DOMAIN = os.environ.get("RAILWAY_PUBLIC_DOMAIN")     # Railway domen bersa, o'zi to'ladi
WEBAPP_URL = os.environ.get("WEBAPP_URL") or (f"https://{_RW_DOMAIN}/" if _RW_DOMAIN
                                              else "https://SIZNING-SAYT.github.io/davomat/")
# Bot ishga tushganda index.html ni o'zi yaratadi (GitHub Pages ga yuklash uchun).
# WEBAPP_PORT > 0 bo'lsa bot sahifani o'zi ham tarqatadi (HTTPS uchun tunnel/proksi kerak).
WEBAPP_PORT = int(os.environ.get("PORT", "0") or 0)   # Railway PORT ni o'zi beradi

MIN_WORDS = 5                # sabab uchun minimal so'z soni
MAX_REASON_LEN = 500

# Sababda shulardan biri uchrasa (so'zning bir qismi bo'lsa ham), "necha kun" so'raladi.
# Apostrof turlari (' ` ʻ ’) avtomatik tenglashtiriladi. Xohlasangiz yangi so'z qo'shing.
SICK_WORDS = (
    "kasal", "касал", "betob", "bemor", "xasta", "kasallan",
    "og'ri", "ogri", "og'rib", "ogrib", "оғри", "огри",
    "isitma", "harorat", "temperatura",
    "shamolla", "tumov", "yo'tal", "yotal", "yutal", "grip", "angina", "covid", "virus", "zotiljam",
    "qusdi", "qusish", "qusyap", "ich ket", "ichim ket", "bosh aylan", "migren", "qon bosim",
    "tobim yo'q", "tobim yoq", "tobim qoch", "tobi yo'q", "tobi yoq", "tobi qoch",
    "holim yo'q", "holim yoq", "holi yo'q", "holi yoq", "holsiz", "mazam yo'q", "mazam yoq",
    "o'zimni yomon", "ozimni yomon", "o'zini yomon", "ozini yomon",
    "dard", "jarohat", "allergiya", "operatsiya", "ukol", "dori ich", "dori ol",
    "shifokor", "vrach", "doktor", "shifoxona", "kasalxona", "bolnitsa", "poliklinika",
)
SICK_DAYS = [("1 kun", 1), ("2 kun", 2), ("3 kun", 3), ("4 kun", 4), ("1 hafta", 7), ("1 oy", 30)]

REMIND_AFTER_DAYS = (1, 2)   # kasallik tugagach 1-kun va 2-kun spravka so'raladi
REMIND_HOUR = 9              # eslatma soati (Toshkent vaqti)

AUTO_REPORT = True           # har kuni avtomatik hisobot yuborish
REPORT_TIME = dtime(10, 0)   # hisobot yuboriladigan vaqt

MAX_LOGIN_TRIES = 5          # login/parol xato urinish limiti
LOCK_MINUTES = 10

try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE_DIR = os.getcwd()
# Railway: Volume ulab, DATA_DIR=/data qiling, aks holda har deployda ma'lumot o'chadi
BASE_DIR = os.environ.get("DATA_DIR") or BASE_DIR
os.makedirs(BASE_DIR, exist_ok=True)

EXCEL_FILE = os.path.join(BASE_DIR, "davomat11b.xlsx")
USERS_FILE = os.path.join(BASE_DIR, "users.txt")
STUDENTS_FILE = os.path.join(BASE_DIR, "students.json")
REMINDERS_FILE = os.path.join(BASE_DIR, "reminders.json")
STATE_FILE = os.path.join(BASE_DIR, "state.json")
ADMIN_SESSIONS_FILE = os.path.join(BASE_DIR, "admin_sessions.json")
APPS_FILE = os.path.join(BASE_DIR, "arizalar.jsonl")   # arizalar bazasi (tez)
CLASSES_FILE = os.path.join(BASE_DIR, "classes.json")   # sinflar
ADMINS_FILE = os.path.join(BASE_DIR, "admins.json")     # sinf adminlari (login/parol)
SUBJECT_TEACHERS_FILE = os.path.join(BASE_DIR, "fan_oqituvchilari.json")   # fan o'qituvchilari

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("Asia/Tashkent")
except Exception:
    TZ = timezone(timedelta(hours=5))

HEADERS = ["Sana", "Vaqt", "UserID", "Ism", "Telefon", "Sinf", "Sabab",
           "Necha kun", "Kimga", "Holat", "Rasm"]
NCOLS = len(HEADERS)

try:
    apihelper.SESSION_TIME_TO_LIVE = 5 * 60   # ulanishni qayta ishlatish -> tezroq
except Exception:
    pass
bot = telebot.TeleBot(TOKEN, parse_mode=None, threaded=True, num_threads=8)
lock = threading.RLock()
user_data = {}       # chat_id -> jarayon holati
keyboard_routes = {} # chat_id -> {tugma matni: callback_data}
attempts = {}        # chat_id -> login urinishlari
CANCEL_TEXT = "❌ Bekor qilish"


# ================= UMUMIY YORDAMCHILAR =================
def now():
    return datetime.now(TZ)


def today_str(offset=0):
    return (now() - timedelta(days=offset)).strftime("%Y-%m-%d")


def _super_accounts():
    return {k.strip().lower(): v for k, v in ADMIN_ACCOUNTS.items()}


def _sessions():
    sess = load_json(ADMIN_SESSIONS_FILE, {})
    return sess if isinstance(sess, dict) else {}


def admin_scope(uid):
    """Faqat belgilangan Telegram ID bosh admin hisoblanadi."""
    if uid == SUPER_ADMIN_ID:
        return "*"
    login = _sessions().get(str(uid))
    if not login:
        return None
    if login in _super_accounts():
        return "*"
    acc = find_admin_by_login(login)
    if acc:
        return {acc["class_id"]}
    return None


def is_admin(uid):
    return admin_scope(uid) is not None


def is_super(uid):
    return admin_scope(uid) == "*"


def can_manage(uid, cid):
    sc = admin_scope(uid)
    return sc == "*" or (sc is not None and cid in sc)


def add_admin_session(uid, login):
    with lock:
        sess = _sessions()
        sess[str(uid)] = login.strip().lower()
        save_json(ADMIN_SESSIONS_FILE, sess)


def remove_admin_session(uid):
    with lock:
        sess = _sessions()
        if str(uid) in sess:
            del sess[str(uid)]
            save_json(ADMIN_SESSIONS_FILE, sess)


def check_admin_login(login, password):
    """To'g'ri bo'lsa login nomini qaytaradi, aks holda None."""
    login = login.strip().lower()
    pw_in = password.strip().encode("utf-8")
    pw = _super_accounts().get(login)
    if pw is not None:
        return login if secrets.compare_digest(pw.encode("utf-8"), pw_in) else None
    acc = find_admin_by_login(login)
    if acc and secrets.compare_digest(acc["password"].encode("utf-8"), pw_in):
        return acc["login"].lower()
    return None


def norm_text(t):
    t = t.lower()
    for ch in ("ʻ", "ʼ", "’", "‘", "`", "´", "′"):
        t = t.replace(ch, "'")
    return t


def is_sick_reason(text):
    t = norm_text(text)
    return any(w in t for w in SICK_WORDS)


class ColorInlineButton(types.InlineKeyboardButton):
    """Rangli inline tugma (Bot API 9.4: primary=ko'k, success=yashil, danger=qizil)."""
    def __init__(self, text, callback_data=None, style=None):
        super().__init__(text, callback_data=callback_data)
        self._style = style

    def to_dict(self):
        d = super().to_dict()
        if self._style:
            d["style"] = self._style
        return d


class ColorKeyboardButton(types.KeyboardButton):
    def __init__(self, text, request_contact=None, style=None):
        if request_contact:
            super().__init__(text, request_contact=True)
        else:
            super().__init__(text)
        self._style = style

    def to_dict(self):
        d = super().to_dict()
        if self._style:
            d["style"] = self._style
        return d


def auto_style(data):
    # qizil: bekor qilish / o'chirish / chiqish
    if data in ("cancel", "adm_logout") or data.startswith(("adm_sdel:", "adm_adel:")):
        return "danger"
    # yashil: asosiy amallar, qo'shish va holat/statistika
    if data in ("send", "login", "go", "phone_keep", "t_all", "my", "adm_stats", "adm_today",
                "adm_week", "adm_users", "adm_cadd") or data.startswith(("spr_", "adm_sadd:", "adm_aadd:")):
        return "success"
    # ko'k: qolgani (menyu, orqaga, hisobotlar...)
    return "primary"


def btn(text, data, style=None):
    return ColorInlineButton(text, callback_data=data, style=style or auto_style(data))


def safe_call(fn, *args, retries=2, **kwargs):
    """Telegram so'rovini bajaradi. 429 (flood) va tarmoq xatolarida qayta urinadi."""
    for i in range(retries + 1):
        try:
            return fn(*args, **kwargs)
        except ApiTelegramException as e:
            if getattr(e, "error_code", None) == 429 and i < retries:
                wait = 3
                try:
                    wait = int(e.result_json.get("parameters", {}).get("retry_after", 3))
                except Exception:
                    pass
                time.sleep(min(wait, 10) + 0.5)
                continue
            raise
        except NET_ERRORS:
            if i < retries:
                time.sleep(1.5 * (i + 1))
                continue
            raise


def _reply_markup(chat_id, kb):
    """Inline callback menyularini pastki ReplyKeyboard menyuga aylantiradi."""
    if kb is None:
        keyboard_routes.pop(chat_id, None)
        return None
    if isinstance(kb, types.ReplyKeyboardMarkup):
        return kb
    if isinstance(kb, types.InlineKeyboardMarkup):
        reply = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
        routes = {}
        for row in kb.keyboard:
            out_row = []
            for b in row:
                if getattr(b, "web_app", None):
                    out_row.append(types.KeyboardButton(b.text, web_app=b.web_app))
                else:
                    out_row.append(types.KeyboardButton(b.text))
                    if getattr(b, "callback_data", None):
                        routes[b.text] = b.callback_data
            if out_row:
                reply.row(*out_row)
        keyboard_routes[chat_id] = routes
        return reply
    return kb


def show(chat_id, text, kb=None, message_id=None):
    """Barcha menyularni pastki klaviaturada ko'rsatadi."""
    try:
        markup = _reply_markup(chat_id, kb)
        safe_call(bot.send_message, chat_id, text, reply_markup=markup)
    except Exception as e:
        print(f"send xatosi: {e}")


def send_long(chat_id, text, kb=None):
    chunks = [text[i:i + 3900] for i in range(0, len(text), 3900)] or [""]
    for i, chunk in enumerate(chunks):
        try:
            safe_call(bot.send_message, chat_id, chunk, reply_markup=_reply_markup(chat_id, kb) if i == len(chunks) - 1 else None)
        except Exception as e:
            print(f"send_long xatosi: {e}")


def send_bytes(chat_id, data, filename, caption=None):
    def _send():
        bio = io.BytesIO(data)          # har urinishda yangi bio
        bio.name = filename
        return bot.send_document(chat_id, bio, caption=caption)
    return safe_call(_send)


def safe_delete(chat_id, message_id):
    try:
        bot.delete_message(chat_id, message_id)
    except Exception:
        pass


def _answer_cb(call_id):
    try:
        bot.answer_callback_query(call_id)
    except Exception:
        pass


# ================= FAYLLAR =================
_json_cache = {}


def load_json(path, default):
    """Fayl o'zgarmagan bo'lsa diskdan qayta o'qimaydi (kesh)."""
    with lock:
        try:
            mt = os.stat(path).st_mtime_ns
        except Exception:
            return copy.deepcopy(default)
        c = _json_cache.get(path)
        if c and c[0] == mt:
            return copy.deepcopy(c[1])
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            print(f"{path} o'qishda xato: {e}")
            try:
                os.replace(path, path + ".buzilgan")
            except Exception:
                pass
            _json_cache.pop(path, None)
            return copy.deepcopy(default)
        _json_cache[path] = (mt, data)
        return copy.deepcopy(data)


def save_json(path, data):
    with lock:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        os.replace(tmp, path)
        try:
            _json_cache[path] = (os.stat(path).st_mtime_ns, copy.deepcopy(data))
        except Exception:
            _json_cache.pop(path, None)


_users = None


def _load_users():
    global _users
    with lock:
        if _users is None:
            _users = set()
            try:
                if os.path.exists(USERS_FILE):
                    with open(USERS_FILE, "r", encoding="utf-8") as f:
                        _users = {u.strip() for u in f.read().splitlines() if u.strip()}
            except Exception as e:
                print(f"users o'qishda xato: {e}")
        return _users


def save_user(user_id):
    try:
        with lock:
            users = _load_users()
            if str(user_id) not in users:
                users.add(str(user_id))
                with open(USERS_FILE, "a", encoding="utf-8") as f:
                    f.write(f"{user_id}\n")
    except Exception as e:
        print(f"save_user xatosi: {e}")


def read_users():
    with lock:
        return list(_load_users())


_apps = None


def _row_clean(r):
    r = list(r)[:NCOLS]
    r += [None] * (NCOLS - len(r))
    out = []
    for v in r:
        if v is None:
            out.append("")
        elif isinstance(v, (int, float, str)):
            out.append(v)
        else:
            out.append(str(v))
    return out


def _load_apps():
    """Arizalarni xotiraga bir marta yuklaydi. Eski Excel bo'lsa, bir marta ko'chiradi."""
    global _apps
    with lock:
        if _apps is not None:
            return _apps
        rows = []
        if os.path.exists(APPS_FILE):
            try:
                with open(APPS_FILE, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            r = json.loads(line)
                            if isinstance(r, list):
                                rows.append(_row_clean(r))
                        except Exception:
                            continue          # buzilgan qatorni o'tkazib yuboradi
            except Exception as e:
                print(f"arizalar o'qishda xato: {e}")
        elif os.path.exists(EXCEL_FILE):
            try:
                wb = load_workbook(EXCEL_FILE, read_only=True)
                ws = wb.active
                for r in ws.iter_rows(min_row=2, values_only=True):
                    if any(v is not None for v in r):
                        rows.append(_row_clean(r))
                wb.close()
                with open(APPS_FILE, "w", encoding="utf-8") as f:
                    for r in rows:
                        f.write(json.dumps(r, ensure_ascii=False) + "\n")
                print(f"Eski Excel'dan {len(rows)} ta ariza ko'chirildi.")
            except Exception as e:
                print(f"Eski Excel'ni ko'chirishda xato: {e}")
        _apps = rows
        return _apps


def save_application(rec, user_id):
    n = now()
    row = _row_clean([
        n.strftime("%Y-%m-%d"), n.strftime("%H:%M:%S"), user_id,
        rec["name"], rec["phone"], rec.get("class_name", ""), rec["reason"],
        rec.get("days_text", ""), rec["targets_text"], rec["status_text"],
        "Bor" if rec.get("file_id") else "Yo'q",
    ])
    with lock:
        rows = _load_apps()
        with open(APPS_FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")
        rows.append(row)


def read_rows():
    with lock:
        return list(_load_apps())


def all_xlsx(rows=None):
    """Arizalardan Excel faylni kerak bo'lganda yasaydi."""
    if rows is None:
        rows = read_rows()
    wb = Workbook()
    ws = wb.active
    ws.title = "Arizalar"
    ws.append(HEADERS)
    for r in rows:
        ws.append(r)
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


# ================= O'QUVCHILAR RO'YXATI =================
def make_prefix(name, used):
    base = re.sub(r"[^a-z0-9]", "", name.lower()) or "s"
    p, i = base, 1
    while p in used:
        i += 1
        p = f"{base}{i}"
    return p


def norm_name(s):
    return " ".join(s.lower().split())


def load_classes():
    return load_json(CLASSES_FILE, [])


def get_class(cid):
    for c in load_classes():
        if c["id"] == cid:
            return c
    return None


def class_name(cid):
    c = get_class(cid)
    return c["name"] if c else ""


def add_class(name):
    """(sinf, xato) qaytaradi."""
    with lock:
        classes = load_classes()
        key = norm_name(name)
        if any(norm_name(c["name"]) == key for c in classes):
            return None, "bunday sinf allaqachon bor"
        nums = [int(c["id"][1:]) for c in classes if c["id"][1:].isdigit()]
        cid = f"k{max(nums, default=0) + 1}"
        c = {"id": cid, "name": name, "prefix": make_prefix(name, {x["prefix"] for x in classes})}
        classes.append(c)
        save_json(CLASSES_FILE, classes)
        return c, None


def load_admin_accounts():
    return load_json(ADMINS_FILE, [])


def find_admin_by_login(login):
    l = login.strip().lower()
    for a in load_admin_accounts():
        if a["login"].lower() == l:
            return a
    return None


def class_admins(cid):
    return [a for a in load_admin_accounts() if a.get("class_id") == cid]


def add_admin_accounts(cid, lines):
    """Har qator: "Ism Familiya | Lavozim" (lavozim ixtiyoriy). Login/parol avtomatik."""
    added, skipped = [], []
    with lock:
        c = get_class(cid)
        if not c:
            return [], ["sinf topilmadi"]
        accs = load_admin_accounts()
        logins = {a["login"].lower() for a in accs} | set(_super_accounts())
        ids = {a["id"] for a in accs}
        n = len([a for a in accs if a.get("class_id") == cid])
        for line in lines:
            parts = [x.strip() for x in line.split("|", 1)]
            name = " ".join(parts[0].split())
            label = " ".join(parts[1].split()) if len(parts) > 1 and parts[1] else "Sinf rahbari"
            if not name:
                continue
            if len(name) > 60 or len(label) > 30:
                skipped.append(f"{name[:30]} (juda uzun)")
                continue
            n += 1
            login = f"u{c['prefix']}{n}"
            while login in logins:
                n += 1
                login = f"u{c['prefix']}{n}"
            logins.add(login)
            aid = secrets.token_hex(3)
            while aid in ids:
                aid = secrets.token_hex(3)
            ids.add(aid)
            acc = {"id": aid, "class_id": cid, "name": name, "label": label, "emoji": "👨‍🏫",
                   "login": login, "password": f"{secrets.randbelow(10 ** 6):06d}", "tg_id": None}
            accs.append(acc)
            added.append(acc)
        save_json(ADMINS_FILE, accs)
    return added, skipped


def remove_admin_account(cid, index):
    with lock:
        accs = load_admin_accounts()
        mine = [i for i, a in enumerate(accs) if a.get("class_id") == cid]
        if 1 <= index <= len(mine):
            gone = accs.pop(mine[index - 1])
            save_json(ADMINS_FILE, accs)
            sess = _sessions()
            keep = {k: v for k, v in sess.items() if v != gone["login"].lower()}
            if len(keep) != len(sess):
                save_json(ADMIN_SESSIONS_FILE, keep)
            return gone
    return None


def bind_admin_tg(login, uid):
    """Admin birinchi marta kirganda, shu Telegram akkaunt arizalar keladigan akkaunt bo'ladi."""
    with lock:
        accs = load_admin_accounts()
        for a in accs:
            if a["login"].lower() == login.lower() and not a.get("tg_id"):
                a["tg_id"] = uid
                save_json(ADMINS_FILE, accs)
                return True
    return False


def class_recips(cid, fallback=True):
    """Sinfga kelgan arizalar yuboriladigan ustozlar: {kalit: {label, emoji, id}}."""
    rec = {}
    for a in class_admins(cid):
        if a.get("tg_id"):
            lab = a.get("label") or "Sinf rahbari"
            if a.get("name") and a["name"] != lab:
                lab = f"{lab} — {a['name']}"
            rec[a["id"]] = {"label": lab, "emoji": a.get("emoji") or "👨‍🏫", "id": a["tg_id"]}
    if not rec and fallback:
        for i, uid in enumerate(ADMINS):
            rec[f"s{i}"] = {"label": "Admin", "emoji": "👨‍🏫", "id": uid}
    return rec


def ensure_classes():
    """Birinchi ishga tushirishda eski (bitta sinfli) ma'lumotni yangi tizimga ko'chiradi."""
    with lock:
        classes = load_json(CLASSES_FILE, None)
        if not classes:
            prefix = make_prefix(CLASS_NAME, set())
            classes = [{"id": "k1", "name": CLASS_NAME, "prefix": prefix}]
            save_json(CLASSES_FILE, classes)
            if not load_json(ADMINS_FILE, []):
                accs = []
                for n, t in enumerate(TEACHERS.values(), 1):
                    accs.append({
                        "id": secrets.token_hex(3), "class_id": "k1", "name": t["label"],
                        "label": t["label"], "emoji": t["emoji"],
                        "login": f"u{prefix}{n}", "password": f"{secrets.randbelow(10 ** 6):06d}",
                        "tg_id": t["id"],
                    })
                save_json(ADMINS_FILE, accs)
        students = load_json(STUDENTS_FILE, [])
        changed = False
        for st in students:
            if not st.get("class_id"):
                st["class_id"] = classes[0]["id"]
                changed = True
        if changed:
            save_json(STUDENTS_FILE, students)


def get_student(uid):
    for s in load_json(STUDENTS_FILE, []):
        if s.get("user_id") == uid:
            return s
    return None


def set_student_phone(uid, phone):
    with lock:
        students = load_json(STUDENTS_FILE, [])
        for s in students:
            if s.get("user_id") == uid:
                s["phone"] = phone
        save_json(STUDENTS_FILE, students)


def class_students(cid):
    return [s for s in load_json(STUDENTS_FILE, []) if s.get("class_id") == cid]


def add_students(lines, cid):
    added, skipped = [], []
    with lock:
        c = get_class(cid)
        if not c:
            return [], ["sinf topilmadi"]
        prefix = c["prefix"]
        students = load_json(STUDENTS_FILE, [])
        existing = {norm_name(s["name"]) for s in students if s.get("class_id") == cid}
        logins = {s["login"].lower() for s in students}
        nums = [int(s["login"][len(prefix):]) for s in students
                if s.get("class_id") == cid and s["login"].startswith(prefix)
                and s["login"][len(prefix):].isdigit()]
        n = max(nums, default=0)
        for line in lines:
            name = re.sub(r"^\s*\d+\s*[\.\)\-]\s*", "", line).strip()
            name = " ".join(name.split())
            if not name:
                continue
            if len(name.split()) < 2 or len(name) > 60:
                skipped.append(f"{name} (ism va familiya kerak)")
                continue
            if norm_name(name) in existing:
                skipped.append(f"{name} (allaqachon bor)")
                continue
            n += 1
            login = f"{prefix}{n:02d}"
            while login.lower() in logins:
                n += 1
                login = f"{prefix}{n:02d}"
            logins.add(login.lower())
            st = {
                "name": name,
                "login": login,
                "password": f"{secrets.randbelow(10000):04d}",
                "user_id": None,
                "phone": "",
                "class_id": cid,
            }
            students.append(st)
            existing.add(norm_name(name))
            added.append(st)
        save_json(STUDENTS_FILE, students)
    return added, skipped


def remove_student(index, cid):
    with lock:
        students = load_json(STUDENTS_FILE, [])
        mine = [i for i, s in enumerate(students) if s.get("class_id") == cid]
        if 1 <= index <= len(mine):
            gone = students.pop(mine[index - 1])
            save_json(STUDENTS_FILE, students)
            return gone
    return None


def edit_admin_account(cid, index, name, label, login, password):
    """Bosh admin sinf adminining ism/lavozim/login/parolini tahrirlaydi."""
    with lock:
        accs = load_admin_accounts()
        mine = [i for i, a in enumerate(accs) if a.get("class_id") == cid]
        if not (1 <= index <= len(mine)):
            return False, "Admin raqami topilmadi."
        pos = mine[index - 1]
        old_login = accs[pos]["login"].lower()
        new_login = login.strip().lower()
        if not re.fullmatch(r"[a-z0-9_.-]{3,32}", new_login):
            return False, "Login 3–32 belgi: lotin harfi, raqam, . _ -"
        if any(i != pos and a["login"].lower() == new_login for i, a in enumerate(accs)):
            return False, "Bu login band."
        if new_login in _super_accounts():
            return False, "Bu login band."
        if len(password) < 6:
            return False, "Parol kamida 6 belgidan iborat bo'lsin."
        accs[pos].update({"name": name.strip()[:60], "label": label.strip()[:30],
                          "login": new_login, "password": password})
        save_json(ADMINS_FILE, accs)
        if old_login != new_login:
            sess = _sessions()
            sess = {k: (new_login if v == old_login else v) for k, v in sess.items()}
            save_json(ADMIN_SESSIONS_FILE, sess)
        return True, "Admin ma'lumotlari yangilandi."


def edit_student(cid, index, name, password):
    with lock:
        students = load_json(STUDENTS_FILE, [])
        mine = [i for i, s in enumerate(students) if s.get("class_id") == cid]
        if not (1 <= index <= len(mine)):
            return False, "O'quvchi raqami topilmadi."
        if len(name.strip()) < 3 or len(name.strip()) > 60:
            return False, "Ism-familiya 3–60 belgidan iborat bo'lsin."
        if len(password) < 4:
            return False, "Parol kamida 4 belgidan iborat bo'lsin."
        students[mine[index - 1]]["name"] = " ".join(name.split())
        students[mine[index - 1]]["password"] = password
        save_json(STUDENTS_FILE, students)
        return True, "O'quvchi ma'lumotlari yangilandi."


def try_login(uid, login, password):
    with lock:
        students = load_json(STUDENTS_FILE, [])
        match = None
        for s in students:
            if s["login"].lower() == login.strip().lower() and s["password"] == password.strip():
                match = s
                break
        if not match:
            return None
        for s in students:
            if s.get("user_id") == uid:
                s["user_id"] = None
        match["user_id"] = uid
        save_json(STUDENTS_FILE, students)
        return dict(match)


def is_locked(uid):
    a = attempts.get(uid)
    return bool(a and a.get("until") and a["until"] > now())


def register_fail(uid):
    a = attempts.setdefault(uid, {"n": 0, "until": None})
    a["n"] += 1
    if a["n"] >= MAX_LOGIN_TRIES:
        a["until"] = now() + timedelta(minutes=LOCK_MINUTES)
        a["n"] = 0


# ================= KLAVIATURALAR =================
def kb_main(uid):
    kb = types.InlineKeyboardMarkup()
    if get_student(uid):
        kb.add(btn("📝 Ariza yuborish", "send"))
        kb.add(btn("📊 Mening arizalarim", "my"))
    else:
        kb.add(btn("🔑 Kirish (login/parol)", "login"))
    if webapp_ready():      # hamma ko'ra oladi (faqat ko'rish)
        kb.add(types.InlineKeyboardButton("👨‍🏫 Fan o'qituvchilari",
                                          web_app=types.WebAppInfo(url=webapp_url(view=True))))
    if is_admin(uid):
        kb.add(btn("🔐 Admin panel", "adm"))
    kb.add(btn("ℹ️ Yordam", "help"))
    return kb


def kb_back(target="menu", label="🏠 Bosh menyu"):
    kb = types.InlineKeyboardMarkup()
    kb.add(btn(label, target))
    return kb


def kb_cancel():
    kb = types.InlineKeyboardMarkup()
    kb.add(btn(CANCEL_TEXT, "cancel"))
    return kb


def scope_classes(uid):
    sc = admin_scope(uid)
    classes = load_classes()
    if sc == "*":
        return classes
    if not sc:
        return []
    return [c for c in classes if c["id"] in sc]


def scope_rows(uid):
    names = {c["name"] for c in scope_classes(uid)}
    return [r for r in read_rows() if r[5] in names]


def kb_admin(uid=None):
    kb = types.InlineKeyboardMarkup()
    kb.row(btn("🏫 Sinflar", "adm_cls"), btn("📊 Jami", "adm_stats"))
    kb.row(btn("📅 Bugun", "adm_today"), btn("📈 7 kunlik grafik", "adm_week"))
    kb.row(btn("📥 Barcha arizalar (Excel)", "adm_x:all"))
    if uid is not None and is_super(uid):
        kb.row(btn("👨‍🏫 Fan o'qituvchilari", "adm_tch"))
        kb.row(btn("👥 Bot a'zolari", "adm_users"))
    if uid is not None and uid != SUPER_ADMIN_ID:
        kb.row(btn("🚪 Admin paneldan chiqish", "adm_logout"))
    kb.row(btn("🏠 Bosh menyu", "menu"))
    return kb


def kb_classes(uid):
    kb = types.InlineKeyboardMarkup()
    students = load_json(STUDENTS_FILE, [])
    for c in scope_classes(uid):
        n = sum(1 for s in students if s.get("class_id") == c["id"])
        kb.row(btn(f"🏫 {c['name']} · {n} ta", f"adm_c:{c['id']}"))
    if is_super(uid):
        kb.row(btn("➕ Sinf qo'shish", "adm_cadd"))
    kb.row(btn("🔙 Admin panel", "adm"))
    return kb


def kb_class(cid, uid):
    kb = types.InlineKeyboardMarkup()
    kb.row(btn("🎓 O'quvchilar", f"adm_s:{cid}"))
    if is_super(uid):
        kb.row(btn("👨‍🏫 Adminlar", f"adm_a:{cid}"))
    kb.row(btn("📄 Bugun Excel", f"adm_r:{cid}:x:0"), btn("📝 Bugun Word", f"adm_r:{cid}:w:0"))
    kb.row(btn("📄 Kecha Excel", f"adm_r:{cid}:x:1"), btn("📝 Kecha Word", f"adm_r:{cid}:w:1"))
    kb.row(btn("📥 Barcha arizalar (Excel)", f"adm_x:{cid}"))
    kb.row(btn("🔙 Sinflar", "adm_cls"))
    return kb


def kb_students(cid):
    kb = types.InlineKeyboardMarkup()
    kb.row(btn("➕ Qo'shish", f"adm_sadd:{cid}"), btn("📋 Ro'yxat", f"adm_slist:{cid}"))
    kb.row(btn("🗑 O'chirish", f"adm_sdel:{cid}"), btn("✏️ Tahrirlash", f"adm_sedit:{cid}"))
    kb.row(btn("🔙 Sinf menyusi", f"adm_c:{cid}"))
    return kb


def kb_admins(cid):
    kb = types.InlineKeyboardMarkup()
    kb.row(btn("➕ Admin qo'shish", f"adm_aadd:{cid}"), btn("📋 Ro'yxat", f"adm_alist:{cid}"))
    kb.row(btn("🗑 O'chirish", f"adm_adel:{cid}"), btn("✏️ Tahrirlash", f"adm_aedit:{cid}"))
    kb.row(btn("🔙 Sinf menyusi", f"adm_c:{cid}"))
    return kb


def main_text(uid):
    st = get_student(uid)
    if st:
        cn = class_name(st.get("class_id"))
        return (f"👋 Assalomu alaykum, {st['name']}!\n"
                f"Bu bot {cn} sinf davomati uchun.\n\nQuyidagilardan birini tanlang:")
    return ("👋 Assalomu alaykum!\nBu bot sinf davomati uchun.\n\n"
            "Ariza yuborish uchun ustozingiz bergan login va parol bilan kiring.")


# ================= HISOBOTLAR =================
def day_rows(d, cname=None):
    return [r for r in read_rows() if str(r[0]) == d and (cname is None or r[5] == cname)]


def report_xlsx(d, cname, rows=None):
    if rows is None:
        rows = day_rows(d, cname)
    wb = Workbook()
    ws = wb.active
    ws.title = "Davomat"
    ws.append([f"{cname} sinf davomati — {d}"])
    ws.append([])
    ws.append(["№", "Ism familiya", "Telefon", "Sabab", "Necha kun", "Kimga", "Holat", "Vaqt", "Rasm"])
    for i, r in enumerate(rows, 1):
        ws.append([i, r[3], r[4], r[6], r[7], r[8], r[9], str(r[1])[:5], r[10]])
    for col, w in zip("ABCDEFGHI", (5, 25, 16, 45, 11, 28, 40, 8, 7)):
        ws.column_dimensions[col].width = w
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def report_docx(d, cname, rows=None):
    if not HAS_DOCX:
        return None
    if rows is None:
        rows = day_rows(d, cname)
    doc = Document()
    doc.add_heading(f"{cname} sinf davomati — {d}", level=1)
    doc.add_paragraph(f"Jami arizalar: {len(rows)} ta")
    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    for cell, h in zip(table.rows[0].cells, ["№", "Ism familiya", "Telefon", "Sabab", "Necha kun", "Holat"]):
        cell.text = h
    for i, r in enumerate(rows, 1):
        cells = table.add_row().cells
        for cell, v in zip(cells, [str(i), str(r[3]), str(r[4]), str(r[6]), str(r[7]), str(r[9])]):
            cell.text = v
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def send_report(chat_id, d, kind, cid):
    cname = class_name(cid)
    back = kb_back(f"adm_c:{cid}", "🔙 Sinf menyusi")
    rows = day_rows(d, cname)
    if not rows:
        show(chat_id, f"📄 {d} sanasida {cname} bo'yicha ariza yo'q.", back)
        return
    if kind == "x":
        send_bytes(chat_id, report_xlsx(d, cname, rows), f"davomat_{cname}_{d}.xlsx",
                   f"📄 {cname} davomati — {d}")
    else:
        data = report_docx(d, cname, rows)
        if data is None:
            show(chat_id, "⚠️ Word uchun Pip'dan python-docx ni o'rnating.", back)
            return
        send_bytes(chat_id, data, f"davomat_{cname}_{d}.docx", f"📝 {cname} davomati — {d}")


def check_daily_report():
    if not AUTO_REPORT:
        return
    n = now()
    d = n.strftime("%Y-%m-%d")
    if n.time() < REPORT_TIME:
        return
    state = load_json(STATE_FILE, {})
    if state.get("last_report") == d:
        return
    state["last_report"] = d
    save_json(STATE_FILE, state)

    for c in load_classes():
        cname = c["name"]
        rows = day_rows(d, cname)
        if rows:
            recips = class_recips(c["id"])
            xlsx_data = report_xlsx(d, cname, rows)
            word_data = report_docx(d, cname, rows)
        else:
            recips = class_recips(c["id"], fallback=False)
        for t in recips.values():
            try:
                if not rows:
                    safe_call(bot.send_message, t["id"], f"📅 {d}: bugun {cname} bo'yicha ariza kelmadi.")
                    continue
                send_bytes(t["id"], xlsx_data, f"davomat_{cname}_{d}.xlsx", f"📄 {cname} kunlik davomat — {d}")
                if word_data:
                    send_bytes(t["id"], word_data, f"davomat_{cname}_{d}.docx", f"📝 {cname} kunlik davomat — {d}")
            except Exception as e:
                print(f"Hisobot yuborishda xato ({t['id']}): {e}")


# ================= SPRAVKA ESLATMALARI =================
def add_reminder(uid, name, phone, recips, days, cname=""):
    base = (now() + timedelta(days=days)).date()
    due = []
    for k in REMIND_AFTER_DAYS:
        d = base + timedelta(days=k)
        due.append(datetime(d.year, d.month, d.day, REMIND_HOUR, 0, tzinfo=TZ).isoformat())
    with lock:
        rem = load_json(REMINDERS_FILE, [])
        rem.append({
            "id": secrets.token_hex(3), "user_id": uid, "name": name, "phone": phone,
            "recips": recips, "class_name": cname, "due": due, "sent": 0, "done": False,
        })
        save_json(REMINDERS_FILE, rem)


def check_reminders():
    to_send = []
    with lock:
        rem = load_json(REMINDERS_FILE, [])
        keep = []
        for r in rem:
            if r.get("done"):
                continue
            last_due = datetime.fromisoformat(r["due"][-1])
            if now() > last_due + timedelta(days=14):
                continue
            idx = r["sent"]
            if idx < len(r["due"]) and now() >= datetime.fromisoformat(r["due"][idx]):
                to_send.append((r["user_id"], r["id"], r["name"], idx + 1, len(r["due"])))
                r["sent"] = idx + 1
            keep.append(r)
        save_json(REMINDERS_FILE, keep)

    for uid, rid, name, num, total in to_send:
        kb = types.InlineKeyboardMarkup()
        kb.add(btn("📷 Spravka yuborish", f"spr_{rid}"))
        kb = _reply_markup(uid, kb)
        try:
            bot.send_message(
                uid,
                f"🩺 Assalomu alaykum, {name}!\n"
                f"Kasallik bo'yicha spravkangizni (rasm yoki fayl) yuborishingizni so'raymiz. "
                f"({num}/{total} eslatma)",
                reply_markup=kb,
            )
        except Exception as e:
            print(f"Eslatma yuborilmadi ({uid}): {e}")


def scheduler():
    while True:
        try:
            check_reminders()
            check_daily_report()
        except Exception as e:
            print(f"scheduler xatosi: {e}")
        time.sleep(60)


# ================= ARIZA OQIMI =================
def deliver(recips, text, file_id=None, ftype=None):
    def one(key):
        t = recips[key]
        try:
            if file_id and ftype == "photo":
                safe_call(bot.send_photo, t["id"], file_id, caption=text)
            elif file_id:
                safe_call(bot.send_document, t["id"], file_id, caption=text)
            else:
                safe_call(bot.send_message, t["id"], text)
            return True
        except Exception as e:
            print(f"{t['id']} ga yuborib bo'lmadi: {e}")
            return False

    keys = list(recips)
    if len(keys) <= 1:
        return {k: one(k) for k in keys}
    with ThreadPoolExecutor(max_workers=len(keys)) as ex:
        res = list(ex.map(one, keys))
    return dict(zip(keys, res))


def status_lines(results, recips):
    lines = []
    for key, ok in results.items():
        t = recips[key]
        lines.append(f"{t['emoji']} {t['label']} — {'✅ yuborildi' if ok else '❌ yuborilmadi'}")
    return lines


def application_text(st):
    text = (f"📌 Kelmagan — {st.get('class_name', '')}\n"
            f"👤 {st['name']}\n"
            f"📞 {st['phone']}\n"
            f"📝 {st['reason']}")
    if st.get("days_text"):
        text += f"\n🗓 Necha kun: {st['days_text']}"
    return text


def ask_phone(chat_id):
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.add(ColorKeyboardButton("📱 Telefon raqamni yuborish", request_contact=True, style="success"))
    kb.add(ColorKeyboardButton(CANCEL_TEXT, style="danger"))
    try:
        bot.send_message(chat_id, "📞 Telefon raqamingizni pastdagi tugma orqali yuboring:", reply_markup=kb)
    except Exception as e:
        print(f"ask_phone xatosi: {e}")


def ask_reason(chat_id, mid=None):
    user_data[chat_id]["step"] = "reason"
    show(chat_id, f"✍️ Kelmagan sababingizni yozing (kamida {MIN_WORDS} so'z):", kb_cancel(), mid)


def ask_days(chat_id, mid=None):
    user_data[chat_id]["step"] = "days"
    kb = types.InlineKeyboardMarkup(row_width=3)
    kb.add(*[btn(label, f"d_{n}") for label, n in SICK_DAYS])
    kb.add(btn(CANCEL_TEXT, "cancel"))
    show(chat_id, "🤒 Necha kun maktabga kela olmaysiz?", kb, mid)


def ask_file(chat_id, mid=None):
    st = user_data[chat_id]
    st["step"] = "file"
    extra = "\n(Kasallik bo'lsa spravka rasmini yuborishingiz mumkin.)" if st.get("sick") else ""
    kb = types.InlineKeyboardMarkup()
    kb.add(btn("⏭ O'tkazib yuborish", "file_skip"))
    kb.add(btn(CANCEL_TEXT, "cancel"))
    show(chat_id, f"📷 Kerak bo'lsa rasm yoki hujjat yuboring (ixtiyoriy).{extra}", kb, mid)


def ask_teacher(chat_id, mid=None):
    st = user_data[chat_id]
    recips = st["recips"]
    if len(recips) == 1:                      # tanlov yo'q — to'g'ridan-to'g'ri tasdiqlash
        st["targets"] = list(recips)
        show_confirm(chat_id, mid)
        return
    st["step"] = "teacher"
    kb = types.InlineKeyboardMarkup()
    for key, t in recips.items():
        kb.add(btn(f"{t['emoji']} {t['label']}", f"t_{key}"))
    if len(recips) > 1:
        kb.add(btn("✅ Hammasiga", "t_all"))
    kb.add(btn(CANCEL_TEXT, "cancel"))
    show(chat_id, "📨 Kimga yuborilsin?", kb, mid)


def show_confirm(chat_id, mid=None):
    st = user_data[chat_id]
    st["step"] = "confirm"
    names = ", ".join(st["recips"][k]["label"] for k in st["targets"])
    text = "📋 Ma'lumotlarni tekshiring:\n\n" + application_text(st)
    text += f"\n📎 Rasm/hujjat: {'bor' if st.get('file_id') else 'yo`q'}"
    text += f"\n📨 Kimga: {names}\n\nYuborilgandan keyin o'zgartirib bo'lmaydi."
    kb = types.InlineKeyboardMarkup()
    kb.add(btn("✅ Yuborish", "go"))
    kb.add(btn(CANCEL_TEXT, "cancel"))
    show(chat_id, text, kb, mid)


def after_reason(chat_id, mid=None):
    st = user_data[chat_id]
    st["sick"] = is_sick_reason(st["reason"])
    if st["sick"]:
        ask_days(chat_id, mid)
    else:
        ask_file(chat_id, mid)


def do_send(chat_id, mid):
    with lock:
        st = user_data.get(chat_id)
        if st and st.get("step") == "sending":
            return                      # ikkinchi bosish e'tiborsiz
        ok = bool(st) and st.get("step") == "confirm"
        if ok:
            st["step"] = "sending"
    if not ok:
        show(chat_id, "⚠️ Sessiya tugagan. Qaytadan boshlang.", kb_main(chat_id), mid)
        return
    if not get_student(chat_id):
        user_data.pop(chat_id, None)
        show(chat_id, "❌ Siz o'quvchilar ro'yxatida topilmadingiz.", kb_main(chat_id), mid)
        return

    chosen = {k: st["recips"][k] for k in st["targets"]}
    results = deliver(chosen, application_text(st), st.get("file_id"), st.get("ftype"))
    st["targets_text"] = ", ".join(t["label"] for t in chosen.values())
    st["status_text"] = "; ".join(
        f"{chosen[k]['label']}: {'yuborildi' if ok else 'yuborilmadi'}" for k, ok in results.items())

    try:
        save_application(st, chat_id)
    except Exception as e:
        print(f"Excel saqlashda xatolik: {e}")

    # Kasal bo'lsa va spravka rasmi biriktirilmagan bo'lsa, keyin qayta so'raladi
    if st.get("sick") and not st.get("file_id"):
        try:
            add_reminder(chat_id, st["name"], st["phone"], chosen, st.get("days", 1), st.get("class_name", ""))
        except Exception as e:
            print(f"Eslatma yaratishda xato: {e}")

    delivered = any(results.values())
    text = "📬 Yuborish holati:\n" + "\n".join(status_lines(results, chosen))
    if not delivered:
        text += "\n\n⚠️ Ariza saqlandi, lekin ustozga xabar yetib bormadi."
    if st.get("sick") and not st.get("file_id"):
        text += "\n\n🩺 Sog'ayganingizdan keyin bot sizdan spravka so'raydi."
    if st.get("sick"):
        text += "\n\n🤝 Tezroq sog'ayib keting! Sog'lig'ingiz mustahkam bo'lsin. Xayr! 👋"
    else:
        text += "\n\n🤝 Rahmat! Sizga omad va yaxshi kun tilaymiz. Xayr! 👋"

    user_data.pop(chat_id, None)
    show(chat_id, text, kb_main(chat_id), mid)


# ================= KOMANDALAR =================
@bot.message_handler(commands=["start", "menu"])
def cmd_start(message):
    try:
        chat_id = message.chat.id
        prev = user_data.pop(chat_id, None)
        save_user(chat_id)
        if prev and prev.get("step") == "phone":
            bot.send_message(chat_id, "🏠", reply_markup=types.ReplyKeyboardRemove())
        show(chat_id, main_text(message.from_user.id), kb_main(message.from_user.id))
    except Exception as e:
        print(f"/start xatosi: {e}")


@bot.message_handler(commands=["admin"])
def cmd_admin(message):
    # /admin yashirin: faqat ustozlarga ishlaydi, boshqalarga oddiy menyu ko'rinadi
    try:
        uid = message.from_user.id
        if is_admin(uid):
            show(message.chat.id, f"🔐 Admin panel", kb_admin(uid))
        else:
            show(message.chat.id, main_text(uid), kb_main(uid))
    except Exception as e:
        print(f"/admin xatosi: {e}")


@bot.message_handler(commands=["adminhide"])
def cmd_adminhide(message):
    # Ustozlar istalgan telefondan login/parol bilan admin panelga kiradi
    try:
        if message.chat.type != "private":
            return
        uid = message.from_user.id
        chat_id = message.chat.id
        if is_admin(uid):
            user_data.pop(chat_id, None)
            show(chat_id, f"🔐 Admin panel", kb_admin(uid))
            return
        if is_locked(uid):
            show(chat_id, f"⏳ Juda ko'p xato urinish. {LOCK_MINUTES} daqiqadan keyin urinib ko'ring.",
                 kb_back())
            return
        user_data[chat_id] = {"step": "adm_login"}
        show(chat_id, "🔐 Ustoz logini:", kb_cancel())
    except Exception as e:
        print(f"/adminhide xatosi: {e}")


# ================= FAN O'QITUVCHILARI (MINI ILOVA) =================
WEBAPP_HTML = r'''<!DOCTYPE html>
<html lang="uz">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>Fan o'qituvchilari</title>
<script src="https://telegram.org/js/telegram-web-app.js"></script>
<style>
:root{--ink:#1d2433;--mute:#5f6b80;--glass:rgba(255,255,255,.55);--edge:rgba(255,255,255,.9);--acc:#2f6bff;--ok:#12995c;--warn:#b87700;--bad:#d1383d}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
body{margin:0;min-height:100vh;font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;color:var(--ink);background:#e8eefc;overflow-x:hidden}
.bg{position:fixed;inset:0;z-index:-1;background:radial-gradient(60% 40% at 15% 8%,#a9c4ff,transparent 70%),radial-gradient(50% 40% at 92% 28%,#ffc9e3,transparent 70%),radial-gradient(60% 50% at 30% 96%,#b6f0dc,transparent 70%),#e8eefc}
.glass{background:var(--glass);backdrop-filter:blur(22px) saturate(170%);-webkit-backdrop-filter:blur(22px) saturate(170%);border:1px solid var(--edge);box-shadow:0 8px 30px rgba(40,60,120,.13),inset 0 1px 0 rgba(255,255,255,.95);border-radius:20px}
.wrap{max-width:560px;margin:0 auto;padding:16px 14px 110px}
header{display:flex;align-items:center;justify-content:space-between;margin:4px 2px 14px}
h1{font-size:22px;margin:0;font-weight:700;letter-spacing:-.2px}
.sub{font-size:13px;color:var(--mute);margin-top:2px}
.icon-btn{width:42px;height:42px;border-radius:14px;border:1px solid var(--edge);background:var(--glass);backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);font-size:18px;cursor:pointer;transition:transform .2s}
.icon-btn:active{transform:scale(.9)}
.search{width:100%;padding:14px 16px;font-size:16px;color:var(--ink);outline:none;margin-bottom:12px;transition:box-shadow .25s}
.search:focus{box-shadow:0 0 0 3px rgba(47,107,255,.25),0 8px 30px rgba(40,60,120,.13)}
.card{display:flex;gap:12px;padding:14px;margin-bottom:10px;cursor:pointer;transition:transform .25s cubic-bezier(.2,.8,.2,1)}
.card:active{transform:scale(.98)}
.card.new{animation:pop .45s cubic-bezier(.2,.9,.3,1.2)}
@keyframes pop{from{opacity:0;transform:translateY(14px) scale(.96)}to{opacity:1;transform:none}}
.av{flex:none;width:46px;height:46px;border-radius:15px;display:grid;place-items:center;font-weight:700;color:#fff;font-size:17px}
.info{flex:1;min-width:0}
.name{font-weight:650;font-size:16px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.meta{font-size:13.5px;color:var(--mute);margin-top:2px}
.note{font-size:13px;color:var(--mute);margin-top:6px}
.row{display:flex;align-items:center;gap:8px;margin-top:8px;flex-wrap:wrap}
.chip{font-size:12.5px;font-weight:600;padding:4px 10px;border-radius:99px;background:rgba(255,255,255,.7)}
.chip.ok{color:var(--ok)}.chip.warn{color:var(--warn)}.chip.bad{color:var(--bad)}
a.call{font-size:13.5px;color:var(--acc);text-decoration:none;font-weight:600}
.empty{text-align:center;color:var(--mute);padding:40px 10px}
.fab{position:fixed;left:50%;bottom:max(22px,env(safe-area-inset-bottom));transform:translateX(-50%);padding:15px 28px;border:0;border-radius:18px;background:var(--acc);color:#fff;font-size:16px;font-weight:650;box-shadow:0 10px 26px rgba(47,107,255,.4);cursor:pointer;transition:transform .2s}
.fab:active{transform:translateX(-50%) scale(.95)}
.scrim{position:fixed;inset:0;background:rgba(30,40,70,.28);opacity:0;pointer-events:none;transition:opacity .35s}
.scrim.on{opacity:1;pointer-events:auto}
.sheet{position:fixed;left:0;right:0;bottom:0;max-width:560px;margin:0 auto;border-radius:26px 26px 0 0;padding:18px 16px max(20px,env(safe-area-inset-bottom));transform:translateY(105%);transition:transform .4s cubic-bezier(.2,.85,.25,1);background:rgba(255,255,255,.72)}
.sheet.on{transform:none}
.sheet h2{margin:0 0 12px;font-size:19px}
label{display:block;font-size:13px;color:var(--mute);margin:10px 2px 5px}
input,select,textarea{width:100%;padding:13px 14px;font:inherit;font-size:16px;color:var(--ink);border:1px solid rgba(255,255,255,.95);background:rgba(255,255,255,.6);border-radius:14px;outline:none}
input:focus,select:focus,textarea:focus{box-shadow:0 0 0 3px rgba(47,107,255,.25)}
textarea{resize:none;height:64px}
.btns{display:flex;gap:10px;margin-top:16px}
.btn{flex:1;padding:14px;border:0;border-radius:14px;font-size:16px;font-weight:650;cursor:pointer;background:rgba(255,255,255,.75);color:var(--ink);transition:transform .2s}
.btn:active{transform:scale(.96)}
.btn.p{background:var(--acc);color:#fff}
.btn.d{color:var(--bad);flex:none;padding:14px 18px}
.err{color:var(--bad);font-size:13px;min-height:16px;margin-top:8px}
@media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
</style>
</head>
<body>
<div class="bg"></div>
<div class="wrap">
  <header>
    <div><h1>Fan o'qituvchilari</h1><div class="sub" id="count"></div></div>
    <button class="icon-btn" id="snd" aria-label="Ovozni yoqish yoki o'chirish">🔊</button>
  </header>
  <input class="glass search" id="q" type="search" placeholder="Ism yoki fan bo'yicha qidirish">
  <div id="list"></div>
</div>
<button class="fab" id="add">+ O'qituvchi qo'shish</button>

<div class="scrim" id="scrim"></div>
<div class="sheet glass" id="sheet" role="dialog" aria-modal="true">
  <h2 id="ttl">Yangi o'qituvchi</h2>
  <label for="f-name">Ism familiya</label><input id="f-name" autocomplete="off">
  <label for="f-subj">Fan</label><input id="f-subj" autocomplete="off" placeholder="Masalan: Matematika">
  <label for="f-tel">Telefon raqam</label><input id="f-tel" type="tel" placeholder="+998 90 123 45 67">
  <label for="f-tg">Telegram username (ixtiyoriy)</label><input id="f-tg" autocomplete="off" placeholder="@username">
  <label for="f-st">Ish faoliyati</label>
  <select id="f-st"><option>Faol</option><option>Ta'tilda</option><option>Ishdan ketgan</option></select>
  <label for="f-note">Izoh (soatlar, sinflar, tajriba)</label><textarea id="f-note"></textarea>
  <div class="err" id="err"></div>
  <div class="btns">
    <button class="btn d" id="del" style="display:none">O'chirish</button>
    <button class="btn" id="cancel">Bekor qilish</button>
    <button class="btn p" id="save">Saqlash</button>
  </div>
</div>

<script>
const tg = window.Telegram && Telegram.WebApp;
if (tg) { tg.ready(); tg.expand(); }
const $ = id => document.getElementById(id);
const esc = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

/* ---- Ovoz (fayl kerak emas, brauzer o'zi chiqaradi) ---- */
let soundOn = true, ac;
function tone(f, d = .09, t = 'sine', delay = 0, v = .08) {
  if (!soundOn) return;
  try {
    ac = ac || new (window.AudioContext || window.webkitAudioContext)();
    const o = ac.createOscillator(), g = ac.createGain(), s = ac.currentTime + delay;
    o.type = t; o.frequency.value = f;
    g.gain.setValueAtTime(v, s); g.gain.exponentialRampToValueAtTime(.0001, s + d);
    o.connect(g); g.connect(ac.destination); o.start(s); o.stop(s + d);
  } catch (e) {}
}
const sfx = {
  tap: () => tone(620, .07),
  open: () => { tone(440, .1); tone(660, .12, 'sine', .07); },
  ok: () => { tone(523, .12); tone(659, .12, 'sine', .09); tone(784, .18, 'sine', .18); },
  del: () => { tone(330, .14, 'triangle'); tone(220, .2, 'triangle', .1); },
  err: () => tone(180, .2, 'square', 0, .04)
};
const buzz = k => { try { tg && tg.HapticFeedback && (k === 'ok' ? tg.HapticFeedback.notificationOccurred('success') : tg.HapticFeedback.impactOccurred('light')); } catch (e) {} };

/* ---- Ma'lumot (hozir telefonda saqlanadi; serverga ulash uchun load/persist ni almashtiring) ---- */
const KEY = 'teachers_v1';
const botMode = new URLSearchParams(location.search).has('d');   // bot ichidan ochilganda true
const viewOnly = new URLSearchParams(location.search).get('m') === 'v';   // oddiy foydalanuvchi: faqat ko'rish
let teachers = [];
function load() {
  try { teachers = JSON.parse(localStorage.getItem(KEY)) || null; } catch (e) { teachers = null; }
  if (!teachers) teachers = [
    { id: 1, name: 'Dilnoza Karimova', subj: 'Matematika', tel: '+998901234567', st: 'Faol', note: '18 soat/hafta, 9-11 sinflar' },
    { id: 2, name: 'Bobur Rahmonov', subj: 'Fizika', tel: '+998935556677', st: 'Faol', note: '10 yillik tajriba' },
    { id: 3, name: 'Madina Yusupova', subj: 'Ingliz tili', tel: '+998977778899', st: "Ta'tilda", note: 'Mart oyida qaytadi' }
  ];
}
function persist() {
  if (botMode) { if (tg) { tg.MainButton.setText('Botga saqlash va yopish'); tg.MainButton.show(); } return; }
  try { localStorage.setItem(KEY, JSON.stringify(teachers)); } catch (e) {}
}

const colors = ['#4f7cff', '#e0559c', '#12a37f', '#e08a1e', '#8a5cf6', '#e5484d'];
const stClass = { 'Faol': 'ok', "Ta'tilda": 'warn', 'Ishdan ketgan': 'bad' };
let newId = null;

function render() {
  const q = $('q').value.trim().toLowerCase();
  const rows = teachers.filter(t => (t.name + ' ' + t.subj).toLowerCase().includes(q) && !(viewOnly && t.st === 'Ishdan ketgan'));
  $('count').textContent = rows.length + ' ta o\'qituvchi';
  $('list').innerHTML = rows.length ? rows.map(t => `
    <div class="card glass ${t.id === newId ? 'new' : ''}" data-id="${t.id}">
      <div class="av" style="background:${colors[t.id % colors.length]}">${esc(t.name.trim()[0] || '?').toUpperCase()}</div>
      <div class="info">
        <div class="name">${esc(t.name)}</div>
        <div class="meta">${esc(t.subj)}</div>
        <div class="row"><span class="chip ${stClass[t.st] || ''}">${esc(t.st)}</span>
          ${t.tel ? `<a class="call" href="tel:${esc(t.tel.replace(/[^+\d]/g, ''))}">${esc(t.tel)}</a>` : ''}
          ${t.tg ? `<a class="call" data-tg="${esc(t.tg)}" href="https://t.me/${esc(t.tg)}">@${esc(t.tg)}</a>` : ''}</div>
        ${t.note ? `<div class="note">${esc(t.note)}</div>` : ''}
      </div>
    </div>`).join('') : '<div class="empty">' + (viewOnly ? 'O\'qituvchi topilmadi.' : 'O\'qituvchi topilmadi. Pastdagi tugma bilan qo\'shing.') + '</div>';
  newId = null;
}

/* ---- Oyna ---- */
let editId = null;
function openSheet(id) {
  editId = id;
  const t = teachers.find(x => x.id === id) || { name: '', subj: '', tel: '', tg: '', st: 'Faol', note: '' };
  $('ttl').textContent = id ? 'Ma\'lumotni tahrirlash' : 'Yangi o\'qituvchi';
  $('f-name').value = t.name; $('f-subj').value = t.subj; $('f-tel').value = t.tel; $('f-tg').value = t.tg || '';
  $('f-st').value = t.st; $('f-note').value = t.note; $('err').textContent = '';
  $('del').style.display = id ? '' : 'none';
  $('scrim').classList.add('on'); $('sheet').classList.add('on');
  sfx.open(); buzz();
}
function closeSheet() { $('scrim').classList.remove('on'); $('sheet').classList.remove('on'); }

$('add').onclick = () => openSheet(null);
$('cancel').onclick = $('scrim').onclick = () => { sfx.tap(); closeSheet(); };
$('list').onclick = e => {
  const a = e.target.closest('a.call');
  if (a) {
    sfx.tap();
    if (a.dataset.tg && tg && tg.openTelegramLink) { e.preventDefault(); tg.openTelegramLink('https://t.me/' + a.dataset.tg); }
    return;
  }
  if (viewOnly) return;
  const c = e.target.closest('.card'); if (c) openSheet(+c.dataset.id);
};
$('q').oninput = render;
$('snd').onclick = () => { soundOn = !soundOn; $('snd').textContent = soundOn ? '🔊' : '🔇'; sfx.tap(); };

$('save').onclick = () => {
  const d = { name: $('f-name').value.trim(), subj: $('f-subj').value.trim(), tel: $('f-tel').value.trim(), tg: $('f-tg').value.trim().replace(/^@/, '').replace(/[^A-Za-z0-9_]/g, '').slice(0, 32), st: $('f-st').value, note: $('f-note').value.trim() };
  if (!d.name || !d.subj) { $('err').textContent = 'Ism va fanni kiriting.'; sfx.err(); return; }
  if (d.tel && !/^\+?[\d\s()-]{7,18}$/.test(d.tel)) { $('err').textContent = 'Telefon raqam noto\'g\'ri. Masalan: +998 90 123 45 67'; sfx.err(); return; }
  if (editId) Object.assign(teachers.find(x => x.id === editId), d);
  else { d.id = Date.now(); teachers.unshift(d); newId = d.id; }
  persist(); closeSheet(); render(); sfx.ok(); buzz('ok');
};
$('del').onclick = () => {
  const run = () => { teachers = teachers.filter(x => x.id !== editId); persist(); closeSheet(); render(); sfx.del(); buzz(); };
  tg && tg.showConfirm ? tg.showConfirm('Bu o\'qituvchi o\'chirilsinmi?', ok => ok && run()) : confirm('Bu o\'qituvchi o\'chirilsinmi?') && run();
};

/* ---- Bot bilan aloqa: ro'yxat URL ichida keladi, saqlashda botga yuboriladi ---- */
const b64u = u8 => btoa(String.fromCharCode(...u8)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
async function fromUrl() {
  try {
    let b = new URLSearchParams(location.search).get('d').replace(/-/g, '+').replace(/_/g, '/');
    b += '='.repeat((4 - b.length % 4) % 4);
    const bin = Uint8Array.from(atob(b), c => c.charCodeAt(0));
    const txt = await new Response(new Blob([bin]).stream().pipeThrough(new DecompressionStream('deflate'))).text();
    const arr = JSON.parse(txt);
    return Array.isArray(arr) ? arr : [];
  } catch (e) { return []; }
}
async function sendToBot() {
  const json = JSON.stringify(teachers);
  let out;
  try {
    const z = await new Response(new Blob([json]).stream().pipeThrough(new CompressionStream('deflate'))).arrayBuffer();
    out = 'z:' + b64u(new Uint8Array(z));
  } catch (e) { out = 'j:' + b64u(new TextEncoder().encode(json)); }
  if (out.length > 4000) { sfx.err(); tg.showAlert('Ro\'yxat juda katta. Izohlarni qisqartiring.'); return; }
  sfx.ok(); tg.sendData(out);
}
if (viewOnly) $('add').style.display = 'none';
(async () => {
  if (botMode) { teachers = await fromUrl(); if (tg) tg.MainButton.onClick(sendToBot); }
  else load();
  render();
})();
</script>
</body>
</html>
'''


def export_webapp():
    """index.html ni bot yoniga yozadi va kerak bo'lsa o'zi tarqatadi."""
    try:
        with open(os.path.join(BASE_DIR, "index.html"), "w", encoding="utf-8") as f:
            f.write(WEBAPP_HTML)
    except Exception as e:
        print(f"index.html yozilmadi: {e}")
    if WEBAPP_PORT:
        from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

        class _H(BaseHTTPRequestHandler):
            def do_GET(self):
                body = WEBAPP_HTML.encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def log_message(self, *a):
                pass

        threading.Thread(target=ThreadingHTTPServer(("0.0.0.0", WEBAPP_PORT), _H).serve_forever,
                         daemon=True).start()
        print(f"Mini ilova serveri: port {WEBAPP_PORT}")

TEACHER_STATUSES = ("Faol", "Ta'tilda", "Ishdan ketgan")
CLOSE_APP_TEXT = "❌ Yopish"


def teachers_list():
    data = load_json(SUBJECT_TEACHERS_FILE, [])
    return data if isinstance(data, list) else []


def webapp_ready():
    return "SIZNING-SAYT" not in WEBAPP_URL


def webapp_url(view=False):
    """Joriy ro'yxatni siqib URL ichiga qo'shadi. view=True: faqat ko'rish (hamma uchun)."""
    raw = json.dumps(teachers_list(), ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    d = base64.urlsafe_b64encode(zlib.compress(raw, 9)).decode().rstrip("=")
    return f"{WEBAPP_URL}?d={d}" + ("&m=v" if view else "")


def clean_teachers(items):
    """Mini ilovadan kelgan ma'lumotni tekshiradi va tozalaydi."""
    if not isinstance(items, list):
        raise ValueError("ro'yxat emas")
    out = []
    for i, t in enumerate(items[:200]):
        if not isinstance(t, dict):
            continue
        name = str(t.get("name", "")).strip()[:60]
        subj = str(t.get("subj", "")).strip()[:40]
        if not name or not subj:
            continue
        st = t.get("st") if t.get("st") in TEACHER_STATUSES else "Faol"
        tid = t.get("id") if isinstance(t.get("id"), int) else int(time.time() * 1000) + i
        out.append({
            "id": tid, "name": name, "subj": subj,
            "tel": re.sub(r"[^\d+\s()-]", "", str(t.get("tel", "")))[:20],
            "tg": re.sub(r"[^A-Za-z0-9_]", "", str(t.get("tg", "")))[:32],
            "st": st, "note": str(t.get("note", "")).strip()[:120],
        })
    return out


@bot.message_handler(content_types=["web_app_data"])
def on_web_app_data(message):
    try:
        uid = message.from_user.id
        chat_id = message.chat.id
        if not is_super(uid):          # faqat bosh admin saqlay oladi
            return
        kind, _, body = (message.web_app_data.data or "").partition(":")
        body += "=" * (-len(body) % 4)
        blob = base64.urlsafe_b64decode(body)
        if kind == "z":
            blob = zlib.decompressobj().decompress(blob, 1_000_000)
        items = clean_teachers(json.loads(blob.decode("utf-8")))
        save_json(SUBJECT_TEACHERS_FILE, items)
        bot.send_message(chat_id, f"✅ Saqlandi: {len(items)} ta o'qituvchi.",
                         reply_markup=types.ReplyKeyboardRemove())
        show(chat_id, "🔐 Admin panel", kb_admin(uid))
    except Exception as e:
        print(f"web_app_data xatosi: {e}")
        try:
            bot.send_message(message.chat.id, "❌ Ma'lumotni saqlab bo'lmadi. Qaytadan urinib ko'ring.",
                             reply_markup=types.ReplyKeyboardRemove())
        except Exception:
            pass


@bot.message_handler(func=lambda m: m.content_type == "text" and m.text == CLOSE_APP_TEXT)
def on_close_app(message):
    try:
        uid = message.from_user.id
        bot.send_message(message.chat.id, "🏠", reply_markup=types.ReplyKeyboardRemove())
        show(message.chat.id, "🔐 Admin panel" if is_admin(uid) else main_text(uid),
             kb_admin(uid) if is_admin(uid) else kb_main(uid))
    except Exception as e:
        print(f"yopish xatosi: {e}")


# ================= KONTAKT / RASM / MATN =================
@bot.message_handler(content_types=["contact"])
def on_contact(message):
    try:
        chat_id = message.chat.id
        st = user_data.get(chat_id)
        if not st or st.get("step") != "phone":
            show(chat_id, main_text(chat_id), kb_main(chat_id))
            return
        if message.contact.user_id != message.from_user.id:
            bot.send_message(chat_id, "❌ Iltimos, faqat o'zingizning raqamingizni yuboring.")
            return
        st["phone"] = message.contact.phone_number
        set_student_phone(chat_id, st["phone"])
        bot.send_message(chat_id, "✅ Raqam qabul qilindi.", reply_markup=types.ReplyKeyboardRemove())
        ask_reason(chat_id)
    except Exception as e:
        print(f"contact xatosi: {e}")


@bot.message_handler(content_types=["photo", "document"])
def on_file(message):
    try:
        chat_id = message.chat.id
        st = user_data.get(chat_id)
        step = st.get("step") if st else None
        if step not in ("file", "spravka"):
            if step in ("teacher", "confirm", "sending"):
                return                  # ikkinchi rasm (albom) — e'tiborsiz
            show(chat_id, main_text(chat_id), kb_main(chat_id))
            return

        if message.content_type == "photo":
            file_id, ftype = message.photo[-1].file_id, "photo"
        else:
            mime = (message.document.mime_type or "")
            if not (mime.startswith("image/") or mime == "application/pdf"):
                bot.send_message(chat_id, "❌ Faqat rasm yoki PDF yuboring.")
                return
            file_id, ftype = message.document.file_id, "document"

        if step == "file":
            st["file_id"], st["ftype"] = file_id, ftype
            bot.send_message(chat_id, "✅ Rasm qabul qilindi.")
            ask_teacher(chat_id)
            return

        # step == "spravka"
        rid = st.get("rid")
        entry = None
        with lock:
            rem = load_json(REMINDERS_FILE, [])
            for r in rem:
                if r["id"] == rid and r["user_id"] == chat_id and not r.get("done"):
                    r["done"] = True
                    entry = dict(r)
            save_json(REMINDERS_FILE, rem)
        user_data.pop(chat_id, None)
        if not entry:
            show(chat_id, "⚠️ Bu eslatma allaqachon yopilgan.", kb_main(chat_id))
            return
        caption = (f"🩺 Spravka — {entry.get('class_name', '')}\n👤 {entry['name']}\n📞 {entry['phone']}")
        recips = entry.get("recips") or {k: TEACHERS[k] for k in entry.get("targets", []) if k in TEACHERS}
        results = deliver(recips, caption, file_id, ftype)
        text = "📬 Spravka holati:\n" + "\n".join(status_lines(results, recips))
        text += "\n\n🤝 Rahmat! Sog'lig'ingiz doim yaxshi bo'lsin. Xayr! 👋"
        show(chat_id, text, kb_main(chat_id))
    except Exception as e:
        print(f"file xatosi: {e}")


def handle_admin_text(message, st):
    chat_id = message.chat.id
    uid = message.from_user.id
    step = st["step"]
    cid = st.get("cid")
    text = (message.text or "").strip()
    user_data.pop(chat_id, None)

    if step == "adm_cadd":
        back = kb_back("adm_cls", "🔙 Sinflar")
        if not is_super(uid):
            show(chat_id, "⚠️ Sinf qo'shish faqat bosh admin uchun.", back)
            return
        name = " ".join(text.split())
        if not name or len(name) > 20:
            show(chat_id, "❌ Sinf nomi 1–20 belgidan iborat bo'lsin.", back)
            return
        c, err = add_class(name)
        if err:
            show(chat_id, f"⚠️ {name}: {err}.", back)
            return
        show(chat_id, f"✅ {c['name']} sinfi qo'shildi.\n\nEndi unga admin va o'quvchilarni qo'shing.",
             kb_class(c["id"], uid))
        return

    if step in ("adm_aedit", "adm_sedit"):
        if step == "adm_aedit" and not is_super(uid):
            show(chat_id, "⚠️ Faqat bosh admin.", kb_back(f"adm_a:{cid}", "🔙 Adminlar"))
            return
        try:
            parts = [x.strip() for x in text.split("|")]
            if step == "adm_aedit" and len(parts) == 5:
                ok, msg = edit_admin_account(cid, int(parts[0]), *parts[1:])
                back = kb_back(f"adm_a:{cid}", "🔙 Adminlar")
            elif step == "adm_sedit" and len(parts) == 3:
                ok, msg = edit_student(cid, int(parts[0]), parts[1], parts[2])
                back = kb_back(f"adm_s:{cid}", "🔙 O'quvchilar")
            else:
                ok, msg = False, "Format noto'g'ri. Maydonlarni | belgisi bilan ajrating."
                back = kb_back(f"adm_c:{cid}", "🔙 Sinf menyusi")
            show(chat_id, ("✅ " if ok else "❌ ") + msg, back)
        except Exception as e:
            show(chat_id, f"❌ Xato: {e}", kb_back(f"adm_c:{cid}", "🔙 Sinf menyusi"))
        return

    # quyidagilar sinf bilan bog'liq
    if not cid or not get_class(cid) or not can_manage(uid, cid):
        show(chat_id, "⚠️ Sinf topilmadi yoki ruxsat yo'q.", kb_back("adm_cls", "🔙 Sinflar"))
        return

    if step == "adm_sadd":
        lines = [l for l in text.splitlines() if l.strip()]
        added, skipped = add_students(lines, cid)
        out = f"✅ {class_name(cid)} ga qo'shildi: {len(added)} ta\n"
        for s in added:
            out += f"• {s['name']} — login: {s['login']} | parol: {s['password']}\n"
        if skipped:
            out += f"\n⚠️ O'tkazib yuborildi: {len(skipped)} ta\n" + "\n".join(skipped)
        send_long(chat_id, out, kb_back(f"adm_s:{cid}", "🔙 O'quvchilar"))

    elif step == "adm_sdel":
        back = kb_back(f"adm_s:{cid}", "🔙 O'quvchilar")
        if text.isdigit():
            gone = remove_student(int(text), cid)
            if gone:
                show(chat_id, f"🗑 O'chirildi: {gone['name']}", back)
                return
        show(chat_id, "❌ Bunday raqam yo'q.", back)

    elif step in ("adm_aadd", "adm_adel"):
        back = kb_back(f"adm_a:{cid}", "🔙 Adminlar")
        if not is_super(uid):
            show(chat_id, "⚠️ Admin qo'shish/o'chirish faqat bosh admin uchun.", back)
            return
        if step == "adm_aadd":
            lines = [l for l in text.splitlines() if l.strip()]
            added, skipped = add_admin_accounts(cid, lines)
            out = f"✅ {class_name(cid)} ga admin qo'shildi: {len(added)} ta\n"
            for a in added:
                out += f"• {a['name']} ({a['label']}) — login: {a['login']} | parol: {a['password']}\n"
            if added:
                out += ("\nUstoz botda /adminhide yozib shu login va parol bilan kirsin. "
                        "Birinchi kirgan Telegram akkaunt shu sinf arizalari keladigan akkaunt bo'ladi.")
            if skipped:
                out += f"\n\n⚠️ O'tkazib yuborildi: {len(skipped)} ta\n" + "\n".join(skipped)
            send_long(chat_id, out, back)
        else:
            if text.isdigit():
                gone = remove_admin_account(cid, int(text))
                if gone:
                    show(chat_id, f"🗑 O'chirildi: {gone['name']}", back)
                    return
            show(chat_id, "❌ Bunday raqam yo'q.", back)


@bot.message_handler(content_types=["text"])
def on_text(message):
    try:
        chat_id = message.chat.id
        uid = message.from_user.id
        text = (message.text or "").strip()
        route = keyboard_routes.get(chat_id, {}).get(text)
        if route:
            from types import SimpleNamespace
            fake = SimpleNamespace(
                id="reply-keyboard",
                from_user=message.from_user,
                message=SimpleNamespace(chat=message.chat, message_id=message.message_id),
                data=route,
            )
            on_callback(fake)
            return
        st = user_data.get(chat_id)

        if not st:
            show(chat_id, main_text(uid), kb_main(uid))
            return
        step = st.get("step")

        if step in ("adm_cadd", "adm_sadd", "adm_sdel", "adm_sedit", "adm_aadd", "adm_adel", "adm_aedit"):
            if not is_admin(uid):
                user_data.pop(chat_id, None)
                show(chat_id, main_text(uid), kb_main(uid))
                return
            handle_admin_text(message, st)
            return

        if step in ("adm_login", "adm_pass"):
            if is_locked(uid):
                user_data.pop(chat_id, None)
                safe_delete(chat_id, message.message_id)
                show(chat_id, f"⏳ Juda ko'p xato urinish. {LOCK_MINUTES} daqiqadan keyin urinib ko'ring.",
                     kb_back())
                return
            if step == "adm_login":
                st["login"] = text
                st["step"] = "adm_pass"
                show(chat_id, "🔑 Parolni yozing:", kb_cancel())
            else:
                safe_delete(chat_id, message.message_id)   # parol chatda qolmasin
                acct = check_admin_login(st.get("login", ""), text)
                if acct:
                    attempts.pop(uid, None)
                    add_admin_session(uid, acct)
                    bind_admin_tg(acct, uid)
                    user_data.pop(chat_id, None)
                    show(chat_id, f"✅ Xush kelibsiz, ustoz!\n\n🔐 Admin panel", kb_admin(uid))
                else:
                    register_fail(uid)
                    st["step"] = "adm_login"
                    show(chat_id, "❌ Login yoki parol noto'g'ri. Loginni qayta yozing:", kb_cancel())
            return

        if step == "login":
            if is_locked(uid):
                user_data.pop(chat_id, None)
                show(chat_id, f"⏳ Juda ko'p xato urinish. {LOCK_MINUTES} daqiqadan keyin urinib ko'ring.",
                     kb_back())
                return
            st["login"] = text
            st["step"] = "password"
            show(chat_id, "🔐 Parolni yozing:", kb_cancel())

        elif step == "password":
            safe_delete(chat_id, message.message_id)   # parol chatda qolmasin
            if is_locked(uid):
                user_data.pop(chat_id, None)
                show(chat_id, f"⏳ Juda ko'p xato urinish. {LOCK_MINUTES} daqiqadan keyin urinib ko'ring.",
                     kb_back())
                return
            student = try_login(chat_id, st.get("login", ""), text)
            if not student:
                register_fail(uid)
                st["step"] = "login"
                show(chat_id, "❌ Login yoki parol noto'g'ri. Loginni qayta yozing:", kb_cancel())
                return
            attempts.pop(uid, None)
            user_data.pop(chat_id, None)
            show(chat_id, f"✅ Xush kelibsiz, {student['name']}!", kb_main(chat_id))

        elif step == "phone":
            if text == CANCEL_TEXT:
                user_data.pop(chat_id, None)
                bot.send_message(chat_id, "🚫 Bekor qilindi.", reply_markup=types.ReplyKeyboardRemove())
                show(chat_id, main_text(uid), kb_main(uid))
            else:
                bot.send_message(chat_id, "📱 Iltimos, pastdagi «Telefon raqamni yuborish» tugmasini bosing.")

        elif step == "reason":
            if len(text.split()) < MIN_WORDS:
                show(chat_id, f"❌ Kamida {MIN_WORDS} so'z yozing.", kb_cancel())
                return
            if len(text) > MAX_REASON_LEN:
                show(chat_id, f"❌ Juda uzun. {MAX_REASON_LEN} belgidan oshmasin.", kb_cancel())
                return
            st["reason"] = text
            after_reason(chat_id)

        elif step in ("days", "file", "teacher", "confirm", "spravka"):
            show(chat_id, "👆 Yuqoridagi tugmalardan foydalaning" +
                 (" yoki rasm yuboring." if step in ("file", "spravka") else "."), kb_cancel())

    except Exception as e:
        print(f"text xatosi: {e}")


# ================= ADMIN KO'RINISHLARI =================
CLASS_ACTIONS = ("adm_c", "adm_s", "adm_sadd", "adm_slist", "adm_sdel", "adm_sedit",
                 "adm_a", "adm_aadd", "adm_alist", "adm_adel", "adm_aedit", "adm_r", "adm_x")
SUPER_ONLY = ("adm_a", "adm_aadd", "adm_alist", "adm_adel", "adm_aedit")


def send_excel(chat_id, mid, rows, filename, back):
    try:
        send_bytes(chat_id, all_xlsx(rows), filename)
        show(chat_id, "📥 Excel fayl yuborildi.", back, mid)
    except Exception as e:
        print(f"Excel yuborish xatosi: {e}")
        show(chat_id, "❌ Faylni yuborib bo'lmadi.", back, mid)


def admin_view(chat_id, mid, data):
    uid = chat_id
    if not is_admin(uid):
        show(chat_id, main_text(uid), kb_main(uid), mid)
        return
    home = kb_back("adm", "🔙 Admin panel")

    if data == "adm":
        user_data.pop(chat_id, None)
        show(chat_id, "🔐 Admin panel", kb_admin(uid), mid)
        return

    if data == "adm_cls":
        user_data.pop(chat_id, None)
        text = "🏫 Sinflar" if scope_classes(uid) else "🏫 Hozircha sinf yo'q."
        show(chat_id, text, kb_classes(uid), mid)
        return

    if data == "adm_tch":
        if not is_super(uid):
            show(chat_id, "⚠️ Bu bo'lim faqat bosh admin uchun.", home, mid)
            return
        if not webapp_ready():
            show(chat_id, "⚠️ Avval kod tepasidagi WEBAPP_URL ga mini ilova manzilini yozing.", home, mid)
            return
        kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
        kb.add(types.KeyboardButton("📱 O'qituvchilar ilovasi", web_app=types.WebAppInfo(url=webapp_url())))
        kb.add(types.KeyboardButton(CLOSE_APP_TEXT))
        safe_call(bot.send_message, chat_id,
                  f"👨‍🏫 Fan o'qituvchilari: {len(teachers_list())} ta\n\n"
                  "Pastdagi tugma bilan ilovani oching. O'zgartirgach «Botga saqlash» tugmasini bosing.",
                  reply_markup=kb)
        return

    if data == "adm_cadd":
        if not is_super(uid):
            show(chat_id, "⚠️ Sinf qo'shish faqat bosh admin uchun.", home, mid)
            return
        user_data[chat_id] = {"step": "adm_cadd"}
        show(chat_id, "🏫 Yangi sinf nomini yozing (masalan: 10-A):",
             kb_back("adm_cls", "🔙 Bekor qilish"), mid)
        return

    if data == "adm_x:all":
        send_excel(chat_id, mid, scope_rows(uid), "barcha_arizalar.xlsx", home)
        return

    head, _, rest = data.partition(":")
    if head in CLASS_ACTIONS and rest:
        parts = rest.split(":")
        cid = parts[0]
        cls = get_class(cid)
        if not cls or not can_manage(uid, cid):
            show(chat_id, "⚠️ Sinf topilmadi yoki ruxsat yo'q.", kb_back("adm_cls", "🔙 Sinflar"), mid)
            return
        name = cls["name"]
        if head in SUPER_ONLY and not is_super(uid):
            show(chat_id, "⚠️ Bu bo'lim faqat bosh admin uchun.", kb_back(f"adm_c:{cid}", "🔙 Sinf menyusi"), mid)
            return

        if head == "adm_c":
            user_data.pop(chat_id, None)
            text = (f"🏫 {name}\n🎓 O'quvchilar: {len(class_students(cid))} ta\n"
                    f"👨‍🏫 Adminlar: {len(class_admins(cid))} ta")
            show(chat_id, text, kb_class(cid, uid), mid)

        elif head == "adm_s":
            user_data.pop(chat_id, None)
            show(chat_id, f"🎓 {name} o'quvchilari: {len(class_students(cid))} ta", kb_students(cid), mid)

        elif head == "adm_sadd":
            user_data[chat_id] = {"step": "adm_sadd", "cid": cid}
            show(chat_id,
                 f"➕ {name} o'quvchilar ro'yxatini yuboring.\nHar qatorda bitta ism-familiya:\n\n"
                 "Ali Valiyev\nVali Aliyev\n\nHar biriga login va parol avtomatik beriladi.",
                 kb_back(f"adm_s:{cid}", "🔙 Bekor qilish"), mid)

        elif head == "adm_slist":
            students = class_students(cid)
            if not students:
                show(chat_id, "📋 Ro'yxat bo'sh. Avval o'quvchilarni qo'shing.", kb_students(cid), mid)
                return
            out = f"📋 {name} ro'yxati ({len(students)} ta):\n\n"
            for i, st in enumerate(students, 1):
                mark = "🟢" if st.get("user_id") else "⚪"
                out += f"{i}. {mark} {st['name']} — {st['login']} / {st['password']}\n"
            out += "\n🟢 botga kirgan  ⚪ hali kirmagan"
            show(chat_id, "⬇️ Ro'yxat yuborildi.", kb_students(cid), mid)
            send_long(chat_id, out)

        elif head == "adm_sdel":
            user_data[chat_id] = {"step": "adm_sdel", "cid": cid}
            show(chat_id, "🗑 O'chirish uchun ro'yxatdagi tartib raqamini yozing (masalan: 5).",
                 kb_back(f"adm_s:{cid}", "🔙 Bekor qilish"), mid)

        elif head == "adm_sedit":
            user_data[chat_id] = {"step": "adm_sedit", "cid": cid}
            show(chat_id, "✏️ O'quvchini tahrirlash: tartib raqami | yangi ism-familiya | yangi parol\nMasalan: 2 | Ali Valiyev | 4821",
                 kb_back(f"adm_s:{cid}", "🔙 Bekor qilish"), mid)

        elif head == "adm_a":
            user_data.pop(chat_id, None)
            show(chat_id, f"👨‍🏫 {name} adminlari: {len(class_admins(cid))} ta", kb_admins(cid), mid)

        elif head == "adm_aadd":
            user_data[chat_id] = {"step": "adm_aadd", "cid": cid}
            show(chat_id,
                 f"➕ {name} uchun admin qo'shish.\nHar qatorda bitta ustoz (lavozim ixtiyoriy):\n\n"
                 "Karimov Aziz | Sinf rahbari\nSodiqova Nodira | Matematika o'qituvchisi\n\n"
                 "Har biriga login va parol avtomatik beriladi.",
                 kb_back(f"adm_a:{cid}", "🔙 Bekor qilish"), mid)

        elif head == "adm_alist":
            accs = class_admins(cid)
            if not accs:
                show(chat_id, "📋 Adminlar yo'q. Avval admin qo'shing.", kb_admins(cid), mid)
                return
            out = f"📋 {name} adminlari ({len(accs)} ta):\n\n"
            for i, a in enumerate(accs, 1):
                mark = "🟢" if a.get("tg_id") else "⚪"
                out += f"{i}. {mark} {a['name']} ({a['label']}) — {a['login']} / {a['password']}\n"
            out += "\n🟢 botga kirgan  ⚪ hali kirmagan"
            show(chat_id, "⬇️ Ro'yxat yuborildi.", kb_admins(cid), mid)
            send_long(chat_id, out)

        elif head == "adm_adel":
            user_data[chat_id] = {"step": "adm_adel", "cid": cid}
            show(chat_id, "🗑 O'chirish uchun adminning tartib raqamini yozing (masalan: 1).",
                 kb_back(f"adm_a:{cid}", "🔙 Bekor qilish"), mid)

        elif head == "adm_aedit":
            user_data[chat_id] = {"step": "adm_aedit", "cid": cid}
            show(chat_id, "✏️ Adminni tahrirlash: tartib raqami | ism-familiya | lavozim | login | parol\nMasalan: 1 | Ali Valiyev | Sinf rahbari | ali01 | YangiParol2026",
                 kb_back(f"adm_a:{cid}", "🔙 Bekor qilish"), mid)

        elif head == "adm_r" and len(parts) == 3:
            kind, off = parts[1], parts[2]
            if kind not in ("x", "w") or not off.isdigit():
                return
            show(chat_id, "⏳ Tayyorlanmoqda...", kb_back(f"adm_c:{cid}", "🔙 Sinf menyusi"), mid)
            send_report(chat_id, today_str(int(off)), kind, cid)

        elif head == "adm_x":
            rows = [r for r in read_rows() if r[5] == name]
            send_excel(chat_id, mid, rows, f"arizalar_{name}.xlsx", kb_back(f"adm_c:{cid}", "🔙 Sinf menyusi"))
        return

    rows = scope_rows(uid)

    if data == "adm_stats":
        text = f"📊 Jami arizalar: {len(rows)} ta"
        classes = scope_classes(uid)
        if len(classes) > 1:
            for c in classes:
                text += f"\n• {c['name']}: {sum(1 for r in rows if r[5] == c['name'])} ta"
        show(chat_id, text, home, mid)

    elif data == "adm_today":
        todays = [r for r in rows if str(r[0]) == today_str()]
        if not todays:
            show(chat_id, "📅 Bugun ariza yo'q.", home, mid)
        else:
            text = f"📅 Bugun: {len(todays)} ta\n\n"
            for r in todays:
                days = f" ({r[7]})" if r[7] else ""
                text += f"• {r[3]} [{r[5]}]{days} — {str(r[1])[:5]}\n   {r[6]}\n"
            show(chat_id, text[:3900], home, mid)

    elif data == "adm_week":
        counts = {today_str(i): 0 for i in range(6, -1, -1)}
        for r in rows:
            if str(r[0]) in counts:
                counts[str(r[0])] += 1
        peak = max(counts.values()) or 1
        text = "📈 Oxirgi 7 kun\n\n"
        for d, c in counts.items():
            bar = "█" * max(1, round(c / peak * 15)) if c else ""
            text += f"{d[5:]} | {bar} {c}\n"
        show(chat_id, text, home, mid)

    elif data == "adm_users":
        if not is_super(uid):
            show(chat_id, "⚠️ Bu bo'lim faqat bosh admin uchun.", home, mid)
        else:
            show(chat_id, f"👥 Botga kirganlar: {len(read_users())} ta", home, mid)


# ================= CALLBACKLAR =================
HELP_TEXT = (
    "ℹ️ Yordam\n\n"
    "1. Ustozingiz bergan login va parol bilan kiring\n"
    "2. «Ariza yuborish» tugmasini bosing\n"
    "3. Telefon raqamingizni yuboring\n"
    f"4. Sababni yozing (kamida {MIN_WORDS} so'z)\n"
    "5. Kerak bo'lsa rasm biriktiring\n"
    "6. Kimga yuborishni tanlang va tekshirib yuboring"
)


@bot.callback_query_handler(func=lambda c: True)
def on_callback(call):
    threading.Thread(target=_answer_cb, args=(call.id,), daemon=True).start()

    try:
        chat_id = call.message.chat.id
        mid = call.message.message_id
        uid = call.from_user.id
        data = call.data or ""
        st = user_data.get(chat_id)
        step = st.get("step") if st else None

        if data == "menu":
            user_data.pop(chat_id, None)
            show(chat_id, main_text(uid), kb_main(uid), mid)

        elif data == "cancel":
            user_data.pop(chat_id, None)
            show(chat_id, "🚫 Bekor qilindi.\n\n" + main_text(uid), kb_main(uid), mid)

        elif data == "help":
            show(chat_id, HELP_TEXT, kb_back(), mid)

        elif data == "login":
            if is_locked(uid):
                show(chat_id, f"⏳ Juda ko'p xato urinish. {LOCK_MINUTES} daqiqadan keyin urinib ko'ring.",
                     kb_back(), mid)
                return
            user_data[chat_id] = {"step": "login"}
            show(chat_id, "🔑 Loginingizni yozing:", kb_cancel(), mid)

        elif data == "send":
            student = get_student(uid)
            if not student:
                user_data[chat_id] = {"step": "login"}
                show(chat_id, "🔑 Avval kiring. Loginingizni yozing:", kb_cancel(), mid)
                return
            cid = student.get("class_id")
            user_data[chat_id] = {"step": "phone", "name": student["name"], "class_id": cid,
                                  "class_name": class_name(cid), "recips": class_recips(cid)}
            if student.get("phone"):
                kb = types.InlineKeyboardMarkup()
                kb.add(btn(f"✅ {student['phone']} — shu raqam", "phone_keep"))
                kb.add(btn("📱 Boshqa raqam", "phone_new"))
                kb.add(btn(CANCEL_TEXT, "cancel"))
                show(chat_id, f"👤 {student['name']}\n📞 Telefon raqamingiz shumi?", kb, mid)
            else:
                show(chat_id, f"👤 {student['name']}\nAriza boshlandi.", kb_cancel(), mid)
                ask_phone(chat_id)

        elif data == "phone_keep":
            student = get_student(uid)
            if step != "phone" or not student or not student.get("phone"):
                show(chat_id, "⚠️ Sessiya tugagan. Qaytadan boshlang.", kb_main(uid), mid)
                return
            st["phone"] = student["phone"]
            ask_reason(chat_id, mid)

        elif data == "phone_new":
            if step != "phone":
                show(chat_id, "⚠️ Sessiya tugagan. Qaytadan boshlang.", kb_main(uid), mid)
                return
            show(chat_id, "📞 Yangi raqamni yuboring.", kb_cancel(), mid)
            ask_phone(chat_id)

        elif data.startswith("d_"):
            if step != "days":
                show(chat_id, "⚠️ Sessiya tugagan. Qaytadan boshlang.", kb_main(uid), mid)
                return
            n = int(data[2:])
            st["days"] = n
            st["days_text"] = next((label for label, v in SICK_DAYS if v == n), f"{n} kun")
            ask_file(chat_id, mid)

        elif data == "file_skip":
            if step != "file":
                show(chat_id, "⚠️ Sessiya tugagan. Qaytadan boshlang.", kb_main(uid), mid)
                return
            ask_teacher(chat_id, mid)

        elif data.startswith("t_"):
            if step != "teacher":
                show(chat_id, "⚠️ Sessiya tugagan. Qaytadan boshlang.", kb_main(uid), mid)
                return
            key = data[2:]
            if key == "all":
                st["targets"] = list(st["recips"].keys())
            elif key in st["recips"]:
                st["targets"] = [key]
            else:
                return
            show_confirm(chat_id, mid)

        elif data == "go":
            do_send(chat_id, mid)

        elif data.startswith("spr_"):
            rid = data[4:]
            with lock:
                ok = any(r["id"] == rid and r["user_id"] == chat_id and not r.get("done")
                         for r in load_json(REMINDERS_FILE, []))
            if not ok:
                show(chat_id, "⚠️ Bu eslatma allaqachon yopilgan.", kb_main(uid), mid)
                return
            user_data[chat_id] = {"step": "spravka", "rid": rid}
            show(chat_id, "📷 Spravka rasmini (yoki PDF) yuboring:", kb_cancel(), mid)

        elif data == "my":
            mine = [r for r in read_rows() if str(r[2]) == str(uid)]
            if not mine:
                show(chat_id, "📊 Siz hali ariza yubormagansiz.", kb_back(), mid)
            else:
                text = f"📊 Jami arizalaringiz: {len(mine)} ta\n\nOxirgilari:\n"
                for r in mine[-5:][::-1]:
                    days = f" • {r[7]}" if r[7] else ""
                    text += f"\n📅 {r[0]} {str(r[1])[:5]}{days}\n📝 {r[6]}\n"
                show(chat_id, text[:3900], kb_back(), mid)

        elif data == "adm_logout":
            remove_admin_session(uid)
            user_data.pop(chat_id, None)
            show(chat_id, "🚪 Admin paneldan chiqdingiz.\n\n" + main_text(uid), kb_main(uid), mid)

        elif data.startswith("adm"):
            if not is_admin(uid):          # admin bo'limi yashirin
                show(chat_id, main_text(uid), kb_main(uid), mid)
                return
            admin_view(chat_id, mid, data)

        else:
            show(chat_id, main_text(uid), kb_main(uid), mid)

    except Exception as e:
        print(f"callback xatosi: {e}")
        try:
            show(call.message.chat.id, "❌ Xatolik yuz berdi. /start bosing.", kb_back(), call.message.message_id)
        except Exception:
            pass


# ================= ISHGA TUSHIRISH =================
if __name__ == "__main__":
    if any("ozgartir" in pw for pw in ADMIN_ACCOUNTS.values()):
        print("⚠️ OGOHLANTIRISH: ADMIN_ACCOUNTS dagi parollarni o'zgartiring!")
    ensure_classes()
    export_webapp()
    _load_apps()
    try:
        bot.set_my_commands([types.BotCommand("start", "🏠 Bosh menyu")])
        if webapp_ready():
            bot.set_chat_menu_button(
                menu_button=types.MenuButtonWebApp(
                    text="Fan o'qituvchilari",
                    web_app=types.WebAppInfo(url=webapp_url(view=True))
                )
            )
    except Exception as e:
        print(f"Komandalar o'rnatilmadi: {e}")
    threading.Thread(target=scheduler, daemon=True).start()
    print("Bot ishlayapti... (to'xtatish: Pydroid'da ■ tugmasi)")
    first = True
    while True:
        try:
            bot.infinity_polling(skip_pending=first, timeout=20, long_polling_timeout=20)
        except Exception as e:
            print(f"Polling to'xtadi, 5 soniyadan keyin qayta ulanaman: {e}")
        first = False
        time.sleep(5)
