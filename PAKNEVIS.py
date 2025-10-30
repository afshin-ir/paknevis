import os
import re
import json
import uno
import unohelper
from com.sun.star.awt import MessageBoxButtons as MBButtons
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
from com.sun.star.awt import XTopWindowListener

BASE_DIR = os.path.join(os.path.expanduser("~"), ".config", "libreoffice", "4", "user", "Scripts", "python")
CONFIG_FILE = os.path.join(BASE_DIR, "TextFixer.conf")
REPLACEMENTS_FILE = os.path.join(BASE_DIR, "DocumentList.json")
ZWNJ = "\u200c"
LOG_FILE = os.path.join(BASE_DIR, "TextFixer.log")

def log_error(section, exc):
    msg = f"[{section}] {type(exc).__name__}: {exc}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass

def en_numbers_to_fa(text):
    return text.translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))

def ar_numbers_to_fa(text):
    return text.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "۰۱۲۳۴۵۶۷۸۹"))

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

simple_verbs = [
    "آمدن", "آوردن", "انداختن", "بردن", "بستن", "بودن", "خواستن",
    "خواندن", "خوردن", "دادن", "داشتن", "دانستن", "دیدن", "رفتن",
    "زدن", "شدن", "شستن", "شکستن", "شنیدن", "کردن", "گرفتن",
    "گشتن", "گفتن", "نوشتن", "یافتن",
]

def load_config():
    defaults = {key: True for key in [
        "fix_k_y", "fix_punct", "fix_quotes", "fix_numbers_en", "fix_numbers_ar",
        "fix_he_ye", "fix_me_nemi", "fix_prefix_verbs", "fix_suffixes", "fix_spaces", 
        "fix_extra_spaces", "fix_ellipsis", "fix_fake_hyphens"
    ]}
# تعداد واژه‌های این بانک فراوان است. بنابراین این گزینه به‌طور پیش‌فرض غیرفعال (تیک‌نخورده) است تا سرعت اجرای ماکرو را کند نکند.    
    defaults["fix_dict"] = False
    if not os.path.exists(CONFIG_FILE):
        return defaults
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            lines = f.readlines()
        for line in lines:
            if "=" not in line:
                continue
            key, val = line.strip().split("=", 1)
            defaults[key] = val == "1"
    except Exception as e:
        log_error("load_config", e)
    return defaults

def save_config(options):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            for k, v in options.items():
                f.write(f"{k}={'1' if v else '0'}\n")
    except Exception as e:
        log_error("save_config", e)

class MyTopWindowListener(unohelper.Base, XTopWindowListener):
    def windowClosing(self, ev):
        try:
            ev.Source.dispose()
        except Exception as e:
            log_error("MyTopWindowListener.windowClosing", e)
    def windowClosed(self, ev): pass
    def windowActivated(self, ev): pass
    def windowDeactivated(self, ev): pass

def show_dialog(options):
    try:
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)

        dialog_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
        dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
        dialog.setModel(dialog_model)
        dialog.setTitle("گزینش اصلاح‌ها")

        items = [
    ("fix_k_y", "تبدیل حرف ي و ك عربی به فارسی"),
    ("fix_punct", "تبدیل علائم سجاوندی انگلیسی به فارسی"),
    ("fix_quotes", "گیومهٔ انگلیسی"),
    ("fix_numbers_en", "اعداد انگلیسی"),
    ("fix_numbers_ar", "اعداد عربی"),
    ("fix_he_ye", "کسرهٔ اضافه"),
    ("fix_me_nemi", "فاصلهٔ قبل از پیشوند افعال (مثل: می/نمی)"),
    ("fix_prefix_verbs", "فاصلهٔ بین اجزاء افعال پیشوندی"),
    ("fix_suffixes", "فاصلهٔ قبل از ضمایر ملکی (مثل: رفته ام)"),
    ("fix_spaces", "فاصلهٔ داخلی علائم سجاوندی"),
    ("fix_space_before_punct", "فاصلهٔ قبل از علائم سجاوندی (با استثناء)"),
    ("fix_extra_spaces", "فاصلهٔ اضافه بین واژه‌ها"),
    ("fix_ellipsis", "سه‌نقطهٔ تعلیق"),
    ("fix_fake_hyphens", "تبدیل نیم‌فاصله‌های کاذب به نیم‌فاصلهٔ واقعی"),
    ("fix_dict", "غلط‌های املایی (بانک)")  # ← آخرین گزینه
]

        item_height = 15
        padding_top = 10
        padding_bottom = 30
        btn_height = 15
        dialog_height = padding_top + len(items) * item_height + padding_bottom
        dialog_width = 320
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
        btn_ok.setPropertyValue("Width", 70)
        btn_ok.setPropertyValue("Height", btn_height)
        btn_ok.setPropertyValue("Label", "اجرا")
        btn_ok.setPropertyValue("PushButtonType", 1)
        btn_ok.setPropertyValue("DefaultButton", True)
        dialog_model.insertByName("btn_ok", btn_ok)

        dialog.createPeer(toolkit, None)

        try:
            peer = dialog.getPeer()
            listener = MyTopWindowListener()
            try:
                peer.addTopWindowListener(listener)
            except Exception as e1:
                try:
                    peer.getContainerWindow().addTopWindowListener(listener)
                except Exception as e2:
                    log_error("show_dialog - addTopWindowListener", e2)
        except Exception as e:
            log_error("show_dialog - getPeer", e)

        result = dialog.execute()
        if result == 1:
            selected = {key: dialog.getControl(key).getState() == 1 for key, _ in items}
            save_config(selected)
        else:
            selected = options.copy()

        dialog.dispose()
        return selected
    except Exception as e:
        log_error("show_dialog", e)
        return options.copy()

# ===== توابع اصلاح متن =====

def fix_k_y(text, report_counts):
    c_before = text.count("ك")
    if c_before:
        report_counts["کاف عربی"] += c_before
        text = text.replace("ك", "ک")
    y_before = text.count("ي")
    if y_before:
        report_counts["ی عربی"] += y_before
        text = text.replace("ي", "ی")
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
    punct_map = {",": "،", ";": "؛", "?": "؟", "$": "﷼", "%": "٪"}
    for en_punct, fa_punct in punct_map.items():
        n = text.count(en_punct)
        if n:
            if en_punct == ",": report_counts["ویرگول انگلیسی"] += n
            elif en_punct == ";": report_counts["نقطه‌ویرگول انگلیسی"] += n
            elif en_punct == "?": report_counts["علامت سؤال انگلیسی"] += n
            elif en_punct == "%": report_counts["درصد انگلیسی"] += n
            text = text.replace(en_punct, fa_punct)
    text, n_q = re.subn(r"؟{2,}", "؟", text)
    report_counts["علامت پرسش تکراری"] += n_q
    text, n_e = re.subn(r"!{2,}", "!", text)
    report_counts["علامت تعجب تکراری"] += n_e
    return text

def fix_quotes(text, report_counts):
    quote_chars = ['"', "'", '“', '”', '‘', '’']
    if not any(q in text for q in quote_chars):
        return text
    result, open_q, cnt = [], True, 0
    for ch in text:
        if ch in quote_chars:
            result.append("«" if open_q else "»")
            open_q = not open_q
            cnt += 1
        else:
            result.append(ch)
    if cnt % 2 == 1 and result and result[-1] == "«":
        result[-1] = "»"
    text = "".join(result)
    report_counts["گیومهٔ انگلیسی"] += cnt // 2
    return text

def fix_he_ye(text, report_counts):
    text, n = re.subn(r"(\S*ه)[\s\u200c]ی\b", lambda m: m.group(1) + "ٔ", text)
    report_counts["کسرهٔ اضافه"] += n
    return text

def fix_me_nemi(text, report_counts):
    VERB_SUFFIXES = ["م", "ی", "د", "یم", "ید", "ند"]
    pattern = r"(?<!\u200c)\b(ن?می)(?:\s+)?([\u0600-\u06FF]+)\b"
    def replace_func(match):
        prefix = match.group(1)
        word_part = match.group(2)
        if any(word_part.endswith(suffix) for suffix in VERB_SUFFIXES):
            report_counts["فاصلهٔ قبل از پیشوند افعال (مثل: می/نمی)"] += 1
            return prefix + ZWNJ + word_part
        compound_verbs = ["شده", "رفت", "آمد", "خورد", "گشت", "شد"]
        if word_part in compound_verbs:
            report_counts["فاصلهٔ قبل از پیشوند افعال (مثل: می/نمی)"] += 1
            return prefix + ZWNJ + word_part
        return match.group(0)
    return re.sub(pattern, replace_func, text)

def fix_prefix_verbs(text, report_counts):
    prefixes = ["بر", "در", "فرو", "فرا", "باز", "وا", "ورا", "ور"]
    block_words = ["می", "نمی", "خواهد", "باید", "که"]
    pattern = r"\b(" + "|".join(prefixes) + r")\s+([آ-ی]+)"
    def repl(m):
        prefix = m.group(1)
        next_word = m.group(2)
        if next_word in block_words or next_word not in simple_verbs:
            return m.group(0)
        report_counts["فاصلهٔ بین اجزاء افعال پیشوندی"] += 1
        return prefix + next_word
    return re.sub(pattern, repl, text)

def fix_suffixes(text, report_counts):
    suffixes = r"(تر(?:ین)?|ها|م|ت|ش|ام|ات|اش|ایم|اید|اند|مان|تان|شان)"
    def fix_suffixes_func(m):
        word = m.group(1)
        suffix = m.group(2)
        if not re.search(r"\s", m.group(0)):
            return m.group(0)
        one_letter_suffixes = ["م", "ت", "ش"]
        two_letter_suffixes = ["ام", "ات", "اش"]
        plural_suffixes = ["مان", "تان", "شان"]
        if suffix in one_letter_suffixes:
            return word + suffix
        if suffix in two_letter_suffixes:
            return word + ZWNJ + suffix
        if suffix in plural_suffixes:
            if word.endswith("ه"):
                return word + ZWNJ + suffix
            return word + suffix
        return word + ZWNJ + suffix
    pattern_suffix = rf"(\S+)\s+{suffixes}\b"
    text, n2 = re.subn(pattern_suffix, fix_suffixes_func, text)
    report_counts["فاصلهٔ قبل از ضمایر ملکی (مثل: رفته ام)"] += n2
    return text

def fix_dict(text, report_counts):
    for wrong, correct in REPLACEMENTS.items():
        pat = r"\b" + re.escape(wrong) + r"\b"
        text, n = re.subn(pat, correct, text)
        report_counts["غلط‌های املایی (بانک)"] += n
    return text

def fix_spaces(text, report_counts):
    corrections = [
        (r"(?<=«)\s+", ""),
        (r"\s+(?=»)", ""),
        (r"(?<=\()\s+", ""),
        (r"\s+(?=\))", ""),
        (r"(?<=\[)\s+", ""),
        (r"\s+(?=\])", ""),
        (r"(?<=\{)\s+", ""),
        (r"\s+(?=\})", ""),
        (r"(?<=⟨)\s+", ""),
        (r"\s+(?=⟩)", "")
    ]
    for pat, rep in corrections:
        text, n = re.subn(pat, rep, text)
        report_counts["فاصلهٔ داخلی علائم سجاوندی"] += n
    return text

def fix_space_before_punct(text, report_counts):
    # حذف فاصلهٔ اضافی قبل از علائم سجاوندی، با استثناء (, [, «
    def repl(match):
        punct = match.group(1)
        if punct in "([«":
            if match.start() == 0:
                return punct
            elif text[match.start()-1] != " ":
                return " " + punct
            else:
                return match.group(0)
        else:
            return punct
    new_text, n = re.subn(r"\s*([،؛:؟!.»\]\)\}])", repl, text)
    if n:
        report_counts["فاصلهٔ قبل از علائم سجاوندی"] += n
    return new_text

def fix_extra_spaces(text, report_counts):
    text, n1 = re.subn(r"\s+([،؛؟.\)»\]\}\⟩])", r"\1", text)
    report_counts["فاصلهٔ اضافه بین واژه‌ها"] += n1
    text, n = re.subn(r"[ ]{2,}", " ", text)
    report_counts["فاصلهٔ اضافه بین واژه‌ها"] += n
    return text

def fix_ellipsis(text, report_counts):
    def replace_ellipsis(match):
        report_counts["سه‌نقطهٔ تعلیق"] += 1
        return "…"
    text, _ = re.subn(r"\.{3,}", replace_ellipsis, text)
    return text

def fix_fake_hyphens_with_zwnj(text, report_counts):
    fake_chars = {
        '\u00AD': "Soft Hyphen",
        '\u00AC': "Not Sign",
        '\u200F': "Right-to-Left Mark",
        '\u2005': "Four-Per-Em Space",
        '\uFEFF': "Zero Width No-Break Space",
        '\u200B': "Zero Width Space",
        '\u200D': "Zero Width Joiner",
    }
    total_count = 0
    for ch in fake_chars:
        n = text.count(ch)
        if n:
            text = text.replace(ch, ZWNJ)
            total_count += n
    report_counts["نیم‌فاصلهٔ کاذب"] += total_count
    return text

def fix_all(text, options, report_counts):
    pipeline = []
    if options.get("fix_k_y", True): pipeline.append(fix_k_y)
    if options.get("fix_numbers_en", True): pipeline.append(fix_numbers_en_func)
    if options.get("fix_numbers_ar", True): pipeline.append(fix_numbers_ar_func)
    if options.get("fix_punct", True): pipeline.append(fix_punct)
    if options.get("fix_quotes", True): pipeline.append(fix_quotes)
    if options.get("fix_he_ye", True): pipeline.append(fix_he_ye)
    if options.get("fix_me_nemi", True): pipeline.append(fix_me_nemi)
    if options.get("fix_prefix_verbs", True): pipeline.append(fix_prefix_verbs)
    if options.get("fix_suffixes", True): pipeline.append(fix_suffixes)
    if options.get("fix_dict", True): pipeline.append(fix_dict)
    if options.get("fix_spaces", True): pipeline.append(fix_spaces)
    if options.get("fix_space_before_punct", True): pipeline.append(fix_space_before_punct)
    if options.get("fix_extra_spaces", True): pipeline.append(fix_extra_spaces)
    if options.get("fix_ellipsis", True): pipeline.append(fix_ellipsis)
    if options.get("fix_fake_hyphens", True): pipeline.append(fix_fake_hyphens_with_zwnj)
    for func in pipeline:
        text = func(text, report_counts)
    return text

def fix_text_full(event=None):
    try:
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.getCurrentComponent()
        if not doc or not doc.supportsService("com.sun.star.text.TextDocument"):
            return
        options = load_config()
        options = show_dialog(options)
        report_counts = {k: 0 for k in [
            "کاف عربی", "ی عربی",
            "ویرگول انگلیسی", "نقطه‌ویرگول انگلیسی", "علامت سؤال انگلیسی",
            "گیومهٔ انگلیسی", "اعداد انگلیسی", "اعداد عربی",
            "درصد انگلیسی", "کسرهٔ اضافه", "علامت پرسش تکراری", "علامت تعجب تکراری",
            "فاصلهٔ قبل از پیشوند افعال (مثل: می/نمی)",
            "فاصلهٔ قبل از ضمایر ملکی (مثل: رفته ام)",
            "فاصلهٔ اضافه بین واژه‌ها",
            "فاصلهٔ داخلی علائم سجاوندی",
            "فاصلهٔ قبل از علائم سجاوندی",
            "فاصلهٔ بین اجزاء افعال پیشوندی",
            "غلط‌های املایی (بانک)", "سه‌نقطهٔ تعلیق", "نیم‌فاصلهٔ کاذب"
        ]}
        text = doc.Text
        cursor = text.createTextCursor()
        cursor.gotoStart(False)
        while True:
            cursor.gotoEndOfParagraph(True)
            old_text = cursor.getString()
            if old_text:
                new_text = fix_all(old_text, options, report_counts)
                if new_text != old_text:
                    cursor.setString(new_text)
            if not cursor.gotoNextParagraph(False):
                break
        total = sum(report_counts.values())
        try:
            parent_win = doc.CurrentController.Frame.ContainerWindow
            mb = parent_win.getToolkit().createMessageBox(
                parent_win, MESSAGEBOX, MBButtons.BUTTONS_OK,
                "گزارش اصلاح متن",
                (
                    f"مجموع اصلاحات: {en_numbers_to_fa(str(total))}\n"
                    + "\n".join(
                        f"{k}: {en_numbers_to_fa(str(v))}"
                        for k, v in report_counts.items() if v > 0
                    )
                    if total > 0
                    else "هیچ اصلاحی لازم نبود."
                )
            )
            mb.execute()
        except Exception as e:
            log_error("fix_text_full - MessageBox", e)
    except Exception as e:
        log_error("fix_text_full", e)

