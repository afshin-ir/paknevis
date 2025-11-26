import os
import re
import json
import uno
import unohelper
from com.sun.star.awt import MessageBoxButtons as MBButtons
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
from com.sun.star.awt import XTopWindowListener
import datetime
from urllib.parse import unquote, urlparse
from enum import Enum, auto # <-- وارد کردن کتابخانه Enum

ZWNJ = "\u200c"

# ---------- مسیر فایل‌های تنظیمات ----------
BASE_DIR = os.path.join(os.path.expanduser("~"), ".config", "libreoffice", "4", "user", "Scripts", "python")
CONFIG_FILE = os.path.join(BASE_DIR, "TextFixer.conf")
REPLACEMENTS_FILE = os.path.join(BASE_DIR, "DocumentList.json")
LOG_FILE = os.path.join(BASE_DIR, "TextFixer.log")

# ---------- تعریف Enum برای تنظیمات 
class FixOption(Enum):
    """این Enum تمام کلیدهای تنظیمات را به صورت ثابت و امن تعریف می‌کند."""
    FIX_K_Y = auto()
    FIX_PUNCT = auto()
    FIX_QUOTES = auto()
    FIX_NUMBERS_EN = auto()
    FIX_NUMBERS_AR = auto()
    FIX_HE_YE = auto()
    FIX_ME_NEMI = auto()
    FIX_PREFIX_VERBS = auto()
    FIX_SUFFIXES = auto()
    FIX_SPACES = auto()
    FIX_SPACE_BEFORE_PUNCT = auto()
    FIX_EXTRA_SPACES = auto()
    FIX_ELLIPSIS = auto()
    FIX_FAKE_HYPHENS = auto()
    FIX_DICT = auto()

    @classmethod
    def get_defaults(cls):
        """یک دیکشنری از مقادیر پیش‌فرض برای تمام گزینه‌ها برمی‌گرداند."""
        defaults = {option.name: True for option in cls}
        defaults[FixOption.FIX_DICT.name] = False
        return defaults

    @classmethod
    def get_dialog_items(cls):
        """لیستی از تاپل‌ها (کلید, برچسب) برای ساخت دیالوگ برمی‌گرداند."""
        return [
            (cls.FIX_K_Y.name, "تبدیل حرف ي و ك عربی به فارسی"),
            (cls.FIX_PUNCT.name, "تبدیل علائم سجاوندی انگلیسی به فارسی"),
            (cls.FIX_QUOTES.name, "گیومهٔ انگلیسی"),
            (cls.FIX_NUMBERS_EN.name, "اعداد انگلیسی"),
            (cls.FIX_NUMBERS_AR.name, "اعداد عربی"),
            (cls.FIX_HE_YE.name, "کسرهٔ اضافه"),
            (cls.FIX_ME_NEMI.name, "فاصلهٔ بعد از پیشوند افعال (مثل: می/نمی)"),
            (cls.FIX_PREFIX_VERBS.name, "فاصلهٔ بین اجزاء افعال پیشوندی"),
            (cls.FIX_SUFFIXES.name, "فاصلهٔ قبل از ضمایر ملکی (مثل: رفته ام)"),
            (cls.FIX_SPACES.name, "فاصلهٔ داخلی علائم سجاوندی"),
            (cls.FIX_SPACE_BEFORE_PUNCT.name, "فاصلهٔ قبل از علائم سجاوندی (با استثناء)"),
            (cls.FIX_EXTRA_SPACES.name, "فاصلهٔ اضافه بین واژه‌ها"),
            (cls.FIX_ELLIPSIS.name, "سه‌نقطهٔ تعلیق"),
            (cls.FIX_FAKE_HYPHENS.name, "تبدیل نیم‌فاصله‌های کاذب به نیم‌فاصلهٔ واقعی"),
            (cls.FIX_DICT.name, "غلط‌های املایی (بانک)")
        ]


# ---------- توابع کمکی ----------
# ثبت خطاها در فایل لاگ
def log_error(section, exc):
    msg = f"[{section}] {type(exc).__name__}: {exc}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass

# تبدیل اعداد انگلیسی به فارسی
def en_numbers_to_fa(text):
    return text.translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))

# تبدیل اعداد عربی به فارسی
def ar_numbers_to_fa(text):
    return text.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "۰۱۲۳۴۵۶۷۸۹"))

# بارگذاری فهرست جایگزینی واژه‌ها از فایل JSON
def load_replacements(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        replacements = {}
        for words in data.get("words", []):
            wrong = words.get("wrong")
            correct = words.get("correct")
            if wrong and correct:
                replacements[wrong] = correct
        return replacements
    except (FileNotFoundError, json.JSONDecodeError) as e:
        log_error("load_replacements", e)
        return {}

REPLACEMENTS = load_replacements(REPLACEMENTS_FILE)

# فهرست افعال ساده برای پردازش پیشوندها
simple_verbs = [
    "آمدن", "آوردن", "انداختن", "بردن", "بستن", "بودن", "خواستن",
    "خواندن", "خوردن", "دادن", "داشتن", "دانستن", "دیدن", "رفتن",
    "زدن", "شدن", "شستن", "شکستن", "شنیدن", "کردن", "گرفتن",
    "گشتن", "گفتن", "نوشتن", "یافتن",
]

# بارگذاری تنظیمات کاربر از فایل یا استفاده از پیش‌فرض‌ها
def load_config():
    defaults = FixOption.get_defaults()
    if not os.path.exists(CONFIG_FILE):
        return defaults
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            lines = f.readlines()
        for line in lines:
            if "=" not in line:
                continue
            key, val = line.strip().split("=", 1)
            if key in defaults:
                defaults[key] = val == "1"
    except Exception as e:
        log_error("load_config", e)
    return defaults

# ذخیره تنظیمات کاربر در فایل
def save_config(options):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            for k, v in options.items():
                f.write(f"{k}={'1' if v else '0'}\n")
    except Exception as e:
        log_error("save_config", e)

# تابع مرکزی برای مقداردهی اولیه دیکشنری گزارش
def get_initial_report_counts():
    return {
        "کاف عربی": 0, "ی عربی": 0, "ویرگول انگلیسی": 0, "نقطه‌ویرگول انگلیسی": 0,
        "علامت سؤال انگلیسی": 0, "گیومهٔ انگلیسی": 0, "اعداد انگلیسی": 0, "اعداد عربی": 0,
        "درصد انگلیسی": 0, "کسرهٔ اضافه": 0, "علامت پرسش تکراری": 0, "علامت تعجب تکراری": 0,
        "فاصلهٔ بعد از پیشوند افعال (مثل: می/نمی)": 0, "فاصلهٔ قبل از ضمایر ملکی (مثل: رفته ام)": 0,
        "فاصلهٔ قبل از پسوند جمع": 0, "فاصلهٔ اضافه بین واژه‌ها": 0, "فاصلهٔ داخلی علائم سجاوندی": 0,
        "فاصلهٔ قبل از علائم سجاوندی": 0, "فاصلهٔ بین اجزاء افعال پیشوندی": 0, "غلط‌های املایی (بانک)": 0,
        "سه‌نقطهٔ تعلیق": 0, "نیم‌فاصلهٔ کاذب": 0,
    }

# ---------- کلاس و دیالوگ ----------
class MyTopWindowListener(unohelper.Base, XTopWindowListener):
    def windowClosing(self, ev):
        try: ev.Source.dispose()
        except Exception as e: log_error("MyTopWindowListener.windowClosing", e)
    def windowClosed(self, ev): pass
    def windowActivated(self, ev): pass
    def windowDeactivated(self, ev): pass

# --- تابع اصلاح‌شده ---
def show_dialog(options):
    try:
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        dialog_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
        dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
        dialog.setModel(dialog_model)
        dialog.setTitle("گزینش اصلاح‌ها")

        items = FixOption.get_dialog_items()

        item_height = 15
        padding_top = 10
        padding_bottom = 30
        btn_height = 15
        dialog_height = padding_top + len(items) * item_height + padding_bottom
        dialog_width = 140
        dialog_model.setPropertyValue("Width", dialog_width)
        dialog_model.setPropertyValue("Height", dialog_height)
        dialog_model.setPropertyValue("PositionX", 100)
        dialog_model.setPropertyValue("PositionY", 100)

        y = padding_top
        for key, label in items:
            cb = dialog_model.createInstance("com.sun.star.awt.UnoControlCheckBoxModel")
            cb.setPropertyValue("PositionX", 10)
            cb.setPropertyValue("PositionY", y)
            cb.setPropertyValue("Width", 300)
            cb.setPropertyValue("Height", 12)
            cb.setPropertyValue("Label", label)
            cb.setPropertyValue("State", 1 if options.get(key, True) else 0)
            dialog_model.insertByName(key, cb)
            y += item_height

        btn_ok = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        btn_ok.setPropertyValue("PositionX", 10)
        btn_ok.setPropertyValue("PositionY", y)
        btn_ok.setPropertyValue("Width", 40)
        btn_ok.setPropertyValue("Height", btn_height)
        btn_ok.setPropertyValue("Label", "اجرا")
        btn_ok.setPropertyValue("PushButtonType", 1) # 1 = OK
        btn_ok.setPropertyValue("DefaultButton", True)
        dialog_model.insertByName("btn_ok", btn_ok)

        btn_cancel = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        btn_cancel.setPropertyValue("PositionX", 90)
        btn_cancel.setPropertyValue("PositionY", y)
        btn_cancel.setPropertyValue("Width", 40)
        btn_cancel.setPropertyValue("Height", btn_height)
        btn_cancel.setPropertyValue("Label", "انصراف")
        btn_cancel.setPropertyValue("PushButtonType", 2) # 2 = Cancel
        dialog_model.insertByName("btn_cancel", btn_cancel)

        dialog.createPeer(toolkit, None)
        try:
            peer = dialog.getPeer()
            listener = MyTopWindowListener()
            try: peer.addTopWindowListener(listener)
            except Exception as e1:
                try: peer.getContainerWindow().addTopWindowListener(listener)
                except Exception as e2: log_error("show_dialog - addTopWindowListener", e2)
        except Exception as e: log_error("show_dialog - getPeer", e)

        result = dialog.execute()
        
        # --- منطق اصلاح‌شده ---
        if result == 1:  # فقط اگر روی دکمه "اجرا" کلیک شد
            selected = {key: dialog.getControl(key).getState() == 1 for key, _ in items}
            save_config(selected)
            dialog.dispose()
            return True, selected  # بازگشت True و تنظیمات جدید
        else:  # اگر روی "انصراف" کلیک شد یا پنجره بسته شد
            dialog.dispose()
            return False, options.copy() # بازگشت False و تنظیمات اصلی

    except Exception as e:
        log_error("show_dialog", e)
        try: dialog.dispose()
        except: pass
        return False, options.copy() # در صورت خطا هم انصراف را برگردان

# ---------- توابع اصلاح متن ----------
def fix_k_y(text, report_counts):
    c_before = text.count("ك")
    if c_before: report_counts["کاف عربی"] += c_before; text = text.replace("ك", "ک")
    y_before = text.count("ي")
    if y_before: report_counts["ی عربی"] += y_before; text = text.replace("ي", "ی")
    return text

def fix_numbers_en_func(text, report_counts):
    text, n = re.subn(r"[0-9]", lambda m: en_numbers_to_fa(m.group(0)), text)
    report_counts["اعداد انگلیسی"] += n
    return text

def fix_numbers_ar_func(text, report_counts):
    text, n = re.subn(r"[٠-٩]", lambda m: ar_numbers_to_fa(m.group(0)), text)
    report_counts["اعداد عربی"] += n
    return text

def fix_punct(text, report_counts):
    punct_map = {",":"،",";":"؛","?":"؟","$":"﷼","%":"٪"}
    for en_punct, fa_punct in punct_map.items():
        n = text.count(en_punct)
        if n:
            if en_punct==",": report_counts["ویرگول انگلیسی"]+=n
            elif en_punct==";": report_counts["نقطه‌ویرگول انگلیسی"]+=n
            elif en_punct=="?": report_counts["علامت سؤال انگلیسی"]+=n
            elif en_punct=="%": report_counts["درصد انگلیسی"]+=n
            text = text.replace(en_punct, fa_punct)
    text, n_q = re.subn(r"؟{2,}", "؟", text); report_counts["علامت پرسش تکراری"] += n_q
    text, n_e = re.subn(r"!{2,}", "!", text); report_counts["علامت تعجب تکراری"] += n_e
    return text

def fix_quotes(text, report_counts):
    quote_chars = ['"', "'", '“', '”', '‘', '’']
    if not any(q in text for q in quote_chars): return text
    result, open_q, cnt = [], True, 0
    for ch in text:
        if ch in quote_chars: result.append("«" if open_q else "»"); open_q = not open_q; cnt +=1
        else: result.append(ch)
    if cnt%2==1 and result and result[-1]=="«": result[-1]="»"
    text="".join(result); report_counts["گیومهٔ انگلیسی"] += cnt//2
    return text

def fix_he_ye(text, report_counts):
    text, n = re.subn(r"(\S*ه)[\s\u200c]ی\b", lambda m: m.group(1)+"ٔ", text)
    report_counts["کسرهٔ اضافه"] += n
    return text

def fix_me_nemi(text, report_counts):
    VERB_SUFFIXES = ["م","ی","د","یم","ید","ند"]
    pattern = r"(?<!\u200c)\b(ن?می)(?:\s+)?([\u0600-\u06FF]+)\b"
    def replace_func(match):
        prefix, word_part = match.group(1), match.group(2)
        if any(word_part.endswith(s) for s in VERB_SUFFIXES) or word_part in ["شده","رفت","آمد","خورد","گشت","شد"]:
            report_counts["فاصلهٔ بعد از پیشوند افعال (مثل: می/نمی)"] +=1
            return prefix+ZWNJ+word_part
        return match.group(0)
    return re.sub(pattern, replace_func, text)

def fix_prefix_verbs(text, report_counts):
    prefixes = ["بر","در","فرو","فرا","باز","وا","ورا","ور"]
    block_words = ["می","نمی","خواهد","باید","که"]
    pattern = r"\b(" + "|".join(prefixes) + r")\s+([آ-ی]+)"
    def repl(m):
        prefix, next_word = m.group(1), m.group(2)
        if next_word in block_words or next_word not in simple_verbs: return m.group(0)
        report_counts["فاصلهٔ بین اجزاء افعال پیشوندی"] +=1
        return prefix+next_word
    return re.sub(pattern, repl, text)

def fix_ha_suffix(text, report_counts):
    ha_suffixes = ["ها", "های", "هایی", "هایم", "هایت", "هایش", "هایمان", "هایتان", "هایشان"]
    total_fixes = 0
    for suffix in ha_suffixes:
        pattern = rf"\b(\S+)\s+({suffix})\b"
        text, num_replacements = re.subn(pattern, rf"\1{ZWNJ}\2", text)
        total_fixes += num_replacements
    report_counts["فاصلهٔ قبل از پسوند جمع"] += total_fixes
    return text

def fix_pronominal_suffixes(text, report_counts):
    suffixes_pattern = r"(تر(?:ین)?|م|ت|ش|ام|ات|اش|ایم|اید|اند|مان|تان|شان)"
    def repl(match):
        word, suffix = match.group(1), match.group(2)
        report_counts["فاصلهٔ قبل از ضمایر ملکی (مثل: رفته ام)"] += 1
        if suffix in ["م", "ت", "ش"]: return f"{word}{suffix}"
        return f"{word}{ZWNJ}{suffix}"
    pattern = rf"(\S+)\s+{suffixes_pattern}\b"
    return re.sub(pattern, repl, text)

def fix_suffixes(text, report_counts):
    text = fix_ha_suffix(text, report_counts)
    text = fix_pronominal_suffixes(text, report_counts)
    return text

def fix_dict(text, report_counts):
    if not REPLACEMENTS: return text
    pattern = r"\b(" + "|".join(map(re.escape, REPLACEMENTS.keys())) + r")\b"
    def replace_match(m):
        report_counts["غلط‌های املایی (بانک)"] += 1
        return REPLACEMENTS[m.group(0)]
    return re.sub(pattern, replace_match, text)

def fix_spaces(text, report_counts):
    corrections = [(r"(?<=«)\s+",""), (r"\s+(?=»)", ""), (r"(?<=\()\s+",""), (r"\s+(?=\))",""), (r"(?<=\[)\s+",""), (r"\s+(?=\])",""), (r"(?<=\{)\s+",""), (r"\s+(?=\})",""), (r"(?<=⟨)\s+",""), (r"\s+(?=⟩)","")]
    for pat, rep in corrections:
        text, n = re.subn(pat, rep, text)
        report_counts["فاصلهٔ داخلی علائم سجاوندی"] += n
    return text

def fix_space_before_punct(text, report_counts):
    def repl(match):
        punct = match.group(1)
        if punct in "([«":
            if match.start()==0 or text[match.start()-1]==" ": return match.group(0)
            return " "+punct
        return punct
    new_text, n = re.subn(r"\s*([،؛:؟!.»\]\)\}])", repl, text)
    if n: report_counts["فاصلهٔ قبل از علائم سجاوندی"] += n
    return new_text

def fix_extra_spaces(text, report_counts):
    text, n1 = re.subn(r"\s+([،؛؟.\)»\]\}\⟩])", r"\1", text); report_counts["فاصلهٔ اضافه بین واژه‌ها"] += n1
    text, n2 = re.subn(r"[ ]{2,}", " ", text); report_counts["فاصلهٔ اضافه بین واژه‌ها"] += n2
    return text

def fix_ellipsis(text, report_counts):
    def replace_ellipsis(match):
        report_counts["سه‌نقطهٔ تعلیق"] +=1
        return "…"
    return re.subn(r"\.{3,}", replace_ellipsis, text)[0]

def fix_fake_hyphens_with_zwnj(text, report_counts):
    fake_chars = ['\u00AD','\u00AC','\u200F','\u2005','\uFEFF','\u200B','\u200D']
    total_count = 0
    for ch in fake_chars:
        n = text.count(ch)
        if n: text = text.replace(ch,ZWNJ); total_count += n
    report_counts["نیم‌فاصلهٔ کاذب"] += total_count
    return text

def fix_all(text, options, report_counts):
    pipeline = []
    if options.get(FixOption.FIX_K_Y.name, True): pipeline.append(fix_k_y)
    if options.get(FixOption.FIX_NUMBERS_EN.name, True): pipeline.append(fix_numbers_en_func)
    if options.get(FixOption.FIX_NUMBERS_AR.name, True): pipeline.append(fix_numbers_ar_func)
    if options.get(FixOption.FIX_PUNCT.name, True): pipeline.append(fix_punct)
    if options.get(FixOption.FIX_QUOTES.name, True): pipeline.append(fix_quotes)
    if options.get(FixOption.FIX_HE_YE.name, True): pipeline.append(fix_he_ye)
    if options.get(FixOption.FIX_ME_NEMI.name, True): pipeline.append(fix_me_nemi)
    if options.get(FixOption.FIX_PREFIX_VERBS.name, True): pipeline.append(fix_prefix_verbs)
    if options.get(FixOption.FIX_SUFFIXES.name, True): pipeline.append(fix_suffixes)
    if options.get(FixOption.FIX_DICT.name, True): pipeline.append(fix_dict)
    if options.get(FixOption.FIX_SPACES.name, True): pipeline.append(fix_spaces)
    if options.get(FixOption.FIX_SPACE_BEFORE_PUNCT.name, True): pipeline.append(fix_space_before_punct)
    if options.get(FixOption.FIX_EXTRA_SPACES.name, True): pipeline.append(fix_extra_spaces)
    if options.get(FixOption.FIX_ELLIPSIS.name, True): pipeline.append(fix_ellipsis)
    if options.get(FixOption.FIX_FAKE_HYPHENS.name, True): pipeline.append(fix_fake_hyphens_with_zwnj)
    
    for func in pipeline:
        text = func(text, report_counts)
    return text

# ---------- ماکروی اصلی ----------
# --- تابع اصلاح‌شده ---
def fix_text_full(event=None):
    try:
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.getCurrentComponent()
        if not doc or not doc.supportsService("com.sun.star.text.TextDocument"): return

        options = load_config()
        
        # --- تغییر در این خط ---
        # فراخوانی تابع و دریافت دو مقدار بازگشتی
        should_proceed, options = show_dialog(options)

        # --- این بخش جدید اضافه شده ---
        # اگر کاربر روی انصراف کلیک کرده بود، از تابع خارج شو
        if not should_proceed:
            return

        # --- بقیه کد بدون تغییر باقی می‌ماند ---
        report_counts = get_initial_report_counts()

        selections = doc.CurrentSelection
        has_nonempty_selection = False
        try: count = selections.getCount()
        except Exception: count = 0

        for i in range(count):
            try: sel = selections.getByIndex(i)
            except Exception: continue
            if not hasattr(sel, "String"): continue
            if sel.String and sel.String.strip(): has_nonempty_selection = True; break

        if has_nonempty_selection:
            for i in range(count):
                try: sel = selections.getByIndex(i)
                except Exception: continue
                if not hasattr(sel, "String"): continue
                old_text = sel.String
                new_text = fix_all(old_text, options, report_counts)
                if new_text != old_text: sel.String = new_text
        else:
            text = doc.Text
            cursor = text.createTextCursor()
            cursor.gotoStart(False)
            while True:
                cursor.gotoEndOfParagraph(True)
                old_text = cursor.getString()
                if old_text:
                    new_text = fix_all(old_text, options, report_counts)
                    if new_text != old_text: cursor.setString(new_text)
                if not cursor.gotoNextParagraph(False): break

        total = sum(report_counts.values())
        try:
            parent_win = doc.CurrentController.Frame.ContainerWindow
            mb = parent_win.getToolkit().createMessageBox(
                parent_win, MESSAGEBOX, MBButtons.BUTTONS_OK, "گزارش اصلاح متن",
                (f"مجموع اصلاحات: {en_numbers_to_fa(str(total))}\n" + "\n".join(f"{k}: {en_numbers_to_fa(str(v))}" for k,v in report_counts.items() if v>0) if total>0 else "هیچ اصلاحی لازم نبود.")
            )
            mb.execute()
        except Exception as e: log_error("fix_text_full - MessageBox", e)

        try:
            url = doc.URL
            if not url: folder, filename = os.path.expanduser("~"), "Untitled"
            else: folder, filename = os.path.dirname(unquote(urlparse(url).path)), os.path.basename(unquote(urlparse(url).path))
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            report_path = os.path.join(folder, f"Paknevis Report [{now}].txt")
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(f"نام فایل: {filename}\n\nمجموع اصلاحات: {en_numbers_to_fa(str(total))}\n")
                for k,v in report_counts.items():
                    if v>0: f.write(f"{k}: {en_numbers_to_fa(str(v))}\n")
        except Exception as e: log_error("fix_text_full - write report file", e)

    except Exception as e: log_error("fix_text_full", e)
