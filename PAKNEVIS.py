import os
import re
import xml.etree.ElementTree as ET
import uno
import unohelper
from com.sun.star.awt import MessageBoxButtons as MBButtons
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
from com.sun.star.awt import XTopWindowListener

BASE_DIR = os.path.join(os.path.expanduser("~"), ".config", "libreoffice", "4", "user", "Scripts", "python")
CONFIG_FILE = os.path.join(BASE_DIR, "TextFixer.conf")
REPLACEMENTS_FILE = os.path.join(BASE_DIR, "DocumentList.xml")
ZWNJ = "\u200c"

def en_to_fa_numbers(text):
    return text.translate(str.maketrans("0123456789٠١٢٣٤٥٦٧٨٩", "۰۱۲۳۴۵۶۷۸۹۰۱۲۳۴۵۶۷۸۹"))

def load_replacements(path):
    try:
        tree = ET.parse(path)
        root = tree.getroot()
        replacements = {}
        ns = {"bl": "http://openoffice.org/2001/block-list"}
        for block in root.findall("bl:block", ns):
            wrong = block.get("{http://openoffice.org/2001/block-list}abbreviated-name")
            correct = block.get("{http://openoffice.org/2001/block-list}name")
            if wrong and correct:
                replacements[wrong] = correct
        return replacements
    except (FileNotFoundError, ET.ParseError):
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
        "fix_he_ye", "fix_me_nemi", "fix_prefix_verbs", "fix_suffixes", "fix_dict",
        "fix_spaces", "fix_extra_spaces", "fix_ellipsis"
    ]}
    if not os.path.exists(CONFIG_FILE):
        return defaults
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            lines = f.readlines()
        for line in lines:
            key, val = line.strip().split("=")
            defaults[key] = val == "1"
    except Exception:
        pass
    return defaults

def save_config(options):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            for k, v in options.items():
                f.write(f"{k}={'1' if v else '0'}\n")
    except Exception:
        pass

class MyTopWindowListener(unohelper.Base, XTopWindowListener):
    def windowClosing(self, ev):
        try:
            ev.Source.dispose()
        except Exception:
            pass
    def windowClosed(self, ev): pass
    def windowActivated(self, ev): pass
    def windowDeactivated(self, ev): pass

def show_dialog(options):
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
        ("fix_quotes", "فارسی‌سازی گیومه‌های انگلیسی"),
        ("fix_numbers_en", "تبدیل اعداد انگلیسی به فارسی"),
        ("fix_numbers_ar", "تبدیل اعداد عربی به فارسی"),
        ("fix_he_ye", "اصلاح کسرهٔ اضافه به ترجیح فرهنگستان"),
        ("fix_me_nemi", "اصلاح فاصله‌گذاری پیشوند افعال"),
        ("fix_prefix_verbs", "اصلاح افعال پیشوندیِ ساده"),
        ("fix_suffixes", "اصلاح فاصله‌گذاری پسوندها"),
        ("fix_dict", "تصحیح غلط‌های املایی (بانک)"),
        ("fix_spaces", "حذف فاصله قبل/بعد علائم"),
        ("fix_extra_spaces", "حذف فاصله‌های اضافی بین واژه‌ها"),
        ("fix_ellipsis", "اصلاح سه‌نقطهٔ تعلیق")
    ]

    item_height = 15
    padding_top = 10
    padding_bottom = 30
    btn_height = 15
    dialog_height = padding_top + len(items) * item_height + padding_bottom
    dialog_width = 300
    dialog_model.setPropertyValue("Width", dialog_width)
    dialog_model.setPropertyValue("Height", dialog_height)
    dialog_model.setPropertyValue("PositionX", 100)
    dialog_model.setPropertyValue("PositionY", 100)

    y = padding_top
    for key, label in items:
        cb = dialog_model.createInstance("com.sun.star.awt.UnoControlCheckBoxModel")
        cb.setPropertyValue("PositionX", 10)
        cb.setPropertyValue("PositionY", y)
        cb.setPropertyValue("Width", 250)
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
        except Exception:
            try:
                peer.getContainerWindow().addTopWindowListener(listener)
            except Exception:
                pass
    except Exception:
        pass

    result = dialog.execute()
    if result == 1:
        selected = {key: dialog.getControl(key).getState() == 1 for key, _ in items}
        save_config(selected)
    else:
        selected = options.copy()

    dialog.dispose()
    return selected

def fix_all(text, options, report_counts):
    if options.get("fix_k_y", True):
        c_before = text.count("ك")
        if c_before:
            report_counts["ك→ک"] += c_before
            text = text.replace("ك", "ک")
        y_before = text.count("ي")
        if y_before:
            report_counts["ي→ی"] += y_before
            text = text.replace("ي", "ی")

    if options.get("fix_numbers_en", True):
        text, n1 = re.subn(r"[0-9]", lambda m: en_to_fa_numbers(m.group(0)), text)
        report_counts["اعداد EN→FA"] += n1
    if options.get("fix_numbers_ar", True):
        text, n2 = re.subn(r"[٠-٩]", lambda m: en_to_fa_numbers(m.group(0)), text)
        report_counts["اعداد عربی→FA"] += n2

    if options.get("fix_punct", True):
        punct_map = {",": "،", ";": "؛", "?": "؟"}
        for en_punct, fa_punct in punct_map.items():
            n = text.count(en_punct)
            if n:
                report_counts["،؛؟"] += n
                text = text.replace(en_punct, fa_punct)
        text, n = re.subn(r"؟{2,}", "؟", text)
        report_counts["؟؟؟"] += n

    if options.get("fix_quotes", True) and any(q in text for q in ['"', "'", '“', '”', '‘', '’']):
        quote_chars = ['"', "'", '“', '”', '‘', '’']
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
        report_counts["گیومه"] += cnt // 2

    if options.get("fix_he_ye", True):
        text, n = re.subn(r"(\S*ه)[\s\u200c]ی\b", lambda m: m.group(1) + "ٔ", text)
        report_counts["ه ی → هٔ"] += n

    if options.get("fix_me_nemi", True):
        VERB_SUFFIXES = ["م", "ی", "د", "یم", "ید", "ند"]
        pattern = r"(?<!\u200c)\b(ن?می)(?:\s+)?([\u0600-\u06FF]+)\b"
        def replace_func(match):
            prefix = match.group(1)
            word_part = match.group(2)
            if any(word_part.endswith(suffix) for suffix in VERB_SUFFIXES):
                report_counts["نیم‌فاصله می/نمی"] += 1
                return prefix + ZWNJ + word_part
            compound_verbs = ["شده", "رفت", "آمد", "خورد", "گشت", "شد"]
            if word_part in compound_verbs:
                report_counts["نیم‌فاصله می/نمی"] += 1
                return prefix + ZWNJ + word_part
            return match.group(0)
        text = re.sub(pattern, replace_func, text)

    if options.get("fix_prefix_verbs", True):
        prefixes = ["بر", "در", "فرو", "فرا", "باز", "وا", "ورا", "ور"]
        block_words = ["می", "نمی", "خواهد", "باید", "که"]
        pattern = r"\b(" + "|".join(prefixes) + r")\s+([آ-ی]+)"
        def repl(m):
            prefix = m.group(1)
            next_word = m.group(2)
            if next_word in block_words or next_word not in simple_verbs:
                return m.group(0)
            report_counts["فعل پیشوندی"] += 1
            return prefix + next_word
        text = re.sub(pattern, repl, text)

    if options.get("fix_suffixes", True):
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
        report_counts["نیم‌فاصله پسوندها"] += n2

    if options.get("fix_dict", True):
        for wrong, correct in REPLACEMENTS.items():
            pat = r"\b" + re.escape(wrong) + r"\b"
            text, n = re.subn(pat, correct, text)
            report_counts["غلط‌های املایی (بانک)"] += n

    if options.get("fix_spaces", True):
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
            report_counts["فاصله قبل/بعد علائم"] += n

    if options.get("fix_extra_spaces", True):
        text, n1 = re.subn(r"\s+([،؛؟.\)»\]\}\⟩])", r"\1", text)
        report_counts["فاصله قبل از علائم"] += n1
        text, n = re.subn(r"[ ]{2,}", " ", text)
        report_counts["فاصله‌های اضافی"] += n

    if options.get("fix_ellipsis", True):
        def replace_ellipsis(match):
            report_counts["سه‌نقطهٔ تعلیق"] += 1
            return "…"
        text, _ = re.subn(r"\.{3,}", replace_ellipsis, text)

    return text

def fix_text_full(event=None):
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()
    if not doc or not doc.supportsService("com.sun.star.text.TextDocument"):
        return

    options = load_config()
    options = show_dialog(options)

    report_counts = {k: 0 for k in [
        "ك→ک", "ي→ی", "،؛؟", "گیومه", "اعداد EN→FA", "اعداد عربی→FA",
        "ه ی → هٔ", "؟؟؟", "فاصله قبل از علائم", "غلط‌های املایی (بانک)",
        "نیم‌فاصله پسوندها", "فاصله‌های اضافی", "فاصله قبل/بعد علائم",
        "نیم‌فاصله می/نمی", "فعل پیشوندی", "سه‌نقطهٔ تعلیق"
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
    if total > 0:
        lines = [f"{k}: {en_to_fa_numbers(str(v))}" for k, v in report_counts.items() if v > 0]
        report = f"مجموع اصلاحات: {en_to_fa_numbers(str(total))}\n" + "\n".join(lines)
        parent_win = doc.CurrentController.Frame.ContainerWindow
        mb = parent_win.getToolkit().createMessageBox(
            parent_win, MESSAGEBOX, MBButtons.BUTTONS_OK,
            "گزارش اصلاح متن", report
        )
        mb.execute()

