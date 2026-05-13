import io
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

# -------------------------
# Пути
# -------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
FONTS_DIR = BASE_DIR / "fonts"
ASSETS_DIR = BASE_DIR / "assets"
ACCESS_FILE = DATA_DIR / "access.json"

ASSET_PRODUCTS_FRONT_BG = ASSETS_DIR / "Ценник (ЕвроВизитка - 300ppi) (1).png"
ASSET_PRODUCTS_BACK_BG = ASSETS_DIR / "Ценник Бэк (ЕвроВизитка - 300ppi).png"
# Оставлено для совместимости со старым кодом/импортами: теперь это новый передний фон товаров.
ASSET_PRODUCTS_BG = ASSET_PRODUCTS_FRONT_BG
ASSET_TEA_BANK_BG = ASSETS_DIR / "Ценник Банки 70мм x 70мм - 300ppi Фон.png"
ASSET_TEA_BOX_BG = ASSETS_DIR / "Ценник Коробки 160мм x 20мм - 300ppi Фон.png"
ASSET_TIPS_FRONT_BG = ASSETS_DIR / "Главная Чаевые Фон.png"
ASSET_TIPS_BACK_BG = ASSETS_DIR / "Бэк Чаевые.png"

# -------------------------
# Цвета
# -------------------------
ORANGE = (0xF6 / 255, 0x76 / 255, 0x3C / 255)  # #F6763C
CREAM = (0xF4 / 255, 0xEF / 255, 0xE8 / 255)  # #F4EFE8
LINE = (0xC1 / 255, 0xBA / 255, 0xB1 / 255)  # #C1BAB1
QR_BG_HEX = "FEF6E9"
QR_FG_HEX = "231F20"

# -------------------------
# Категории чая / цветные фоны
# -------------------------
# Для каждой категории нужны 2 картинки в assets/:
#   bank - квадратная картинка для банок 70x70
#   box  - широкая картинка для коробок 160x20
TEA_CATEGORY_ASSETS = {
    "габа": {
        "hex": "F57044",
        "bank": ASSETS_DIR / "габа.png",
        "box": ASSETS_DIR / "габа (1).png",
    },
    "светлый улун": {
        "hex": "2AD604",
        "bank": ASSETS_DIR / "сву.png",
        "box": ASSETS_DIR / "сву (1).png",
    },
    "красный": {
        "hex": "DD0000",
        "bank": ASSETS_DIR / "красный.png",
        "box": ASSETS_DIR / "красный (1).png",
    },
    "жёлтый": {
        "hex": "F8DF00",
        "bank": ASSETS_DIR / "желт.png",
        "box": ASSETS_DIR / "желт (1).png",
    },
    "зеленый": {
        "hex": "9DEE6B",
        "bank": ASSETS_DIR / "зел.png",
        "box": ASSETS_DIR / "зел (1).png",
    },
    "белый": {
        "hex": "FEF6E9",
        "bank": ASSETS_DIR / "бел.png",
        "box": ASSETS_DIR / "бел (1).png",
    },
    "пуэр": {
        "hex": "B76A2F",
        "bank": ASSETS_DIR / "пуэр.png",
        "box": ASSETS_DIR / "пуэр (1).png",
    },
    "темный улун": {
        "hex": "406BEA",
        "bank": ASSETS_DIR / "темну.png",
        "box": ASSETS_DIR / "темну (1).png",
    },
    "травы": {
        "hex": "944FD5",
        "bank": ASSETS_DIR / "травы.png",
        "box": ASSETS_DIR / "травы (1).png",
    },
}

TEA_CATEGORY_ALIASES = {
    "шу пуэр": "пуэр",
    "шен пуэр": "пуэр",
    "лао шен пуэр": "пуэр",
    "фхдц": "темный улун",
    "дхп": "темный улун",
    "те гуань инь": "светлый улун",
    "дун дин": "светлый улун",
    "габа шен": "габа",
    "габа пуэр": "габа",
    "габа красный": "габа",
    "габа улун": "габа",
    "шу габа": "габа",
    "шен габа": "габа",
    "шэн габа": "габа",
}


# -------------------------
# Шрифты
# -------------------------
@dataclass
class Fonts:
    regular: str
    medium: str
    semibold: str
    bold: str


def register_unbounded_fonts() -> Fonts:
    bold_path = FONTS_DIR / "Unbounded-Bold.ttf"
    med_path = FONTS_DIR / "Unbounded-Medium.ttf"
    semi_path = FONTS_DIR / "Unbounded-SemiBold.ttf"
    reg_path = FONTS_DIR / "Unbounded-Regular.ttf"

    for p in [bold_path, med_path, semi_path]:
        if not p.exists():
            raise FileNotFoundError(f"Не найден шрифт: {p}.\nПоложи файл в папку fonts/")

    if not reg_path.exists():
        reg_path = med_path  # fallback

    pdfmetrics.registerFont(TTFont("Unbounded-Bold", str(bold_path)))
    pdfmetrics.registerFont(TTFont("Unbounded-Medium", str(med_path)))
    pdfmetrics.registerFont(TTFont("Unbounded-SemiBold", str(semi_path)))
    pdfmetrics.registerFont(TTFont("Unbounded-Regular", str(reg_path)))

    return Fonts(
        regular="Unbounded-Regular",
        medium="Unbounded-Medium",
        semibold="Unbounded-SemiBold",
        bold="Unbounded-Bold",
    )


def ensure_assets_exist():
    required = [
        ASSET_PRODUCTS_FRONT_BG,
        ASSET_PRODUCTS_BACK_BG,
        ASSET_TEA_BANK_BG,
        ASSET_TEA_BOX_BG,
        ASSET_TIPS_FRONT_BG,
        ASSET_TIPS_BACK_BG,
    ]

    for config in TEA_CATEGORY_ASSETS.values():
        required.append(config["bank"])
        required.append(config["box"])

    for p in required:
        if not p.exists():
            raise FileNotFoundError(f"Не найден ассет: {p}.\nПоложи файл в папку assets/.")


# -------------------------
# Утилиты
# -------------------------
def safe_filename(name: str, max_len: int = 80) -> str:
    name = (name or "").strip()
    name = re.sub(r"[\\/:*?\"<>|\n\r\t]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    if not name:
        name = "item"
    if len(name) > max_len:
        name = name[:max_len].rstrip()
    return name


def unique_names(base_names: List[str]) -> List[str]:
    used: Dict[str, int] = {}
    out = []
    for n in base_names:
        if n not in used:
            used[n] = 1
            out.append(n)
        else:
            used[n] += 1
            out.append(f"{n}_{used[n]}")
    return out


def parse_int_number(value) -> Optional[int]:
    if value is None:
        return None
    if isinstance(value, int):
        return int(value)
    if isinstance(value, float):
        return int(value) if value.is_integer() else None

    s = str(value).strip().replace(",", ".")
    if re.fullmatch(r"\d+(\.0+)?", s):
        return int(float(s))
    return None


def parse_float_number(value) -> Optional[float]:
    """Парсит цену как float: принимает 22, 22.50 и 22,50."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip().replace(",", ".")
    try:
        f = float(s)
        if f != f:  # NaN
            return None
        return f
    except (ValueError, OverflowError):
        return None


def category_from_price(price: float) -> str:
    if price <= 20:
        return "A"
    if price <= 35:
        return "A+"
    if price <= 55:
        return "A++"
    return "ПРЕМИУМ"


def hours_word(n: int) -> str:
    n = abs(int(n))
    n100 = n % 100
    n10 = n % 10
    if 11 <= n100 <= 14:
        return "часов"
    if n10 == 1:
        return "час"
    if 2 <= n10 <= 4:
        return "часа"
    return "часов"


def normalize_sentence_case(text: str) -> str:
    """
    Делает:
    - первая буква заглавная
    - после . ! ? следующая буква заглавная
    """
    s = re.sub(r"\s+", " ", (text or "").strip())
    if not s:
        return s

    parts = re.split(r"([.!?]+)", s)
    out = []
    for i in range(0, len(parts), 2):
        chunk = parts[i].strip()
        punct = parts[i + 1] if i + 1 < len(parts) else ""
        if chunk:
            m = re.search(r"[A-Za-zА-Яа-яЁё]", chunk)
            if m:
                idx = m.start()
                chunk = chunk[:idx] + chunk[idx].upper() + chunk[idx + 1 :]
        out.append(chunk + punct)

    res = " ".join([x.strip() for x in out if x.strip()])
    res = re.sub(r"\s+([.!?])", r"\1", res)
    return res.strip()


def img_size(path: Path) -> Tuple[int, int]:
    with Image.open(path) as im:
        return im.size


def pdf_with_background(page_w: int, page_h: int, bg_path: Path) -> Tuple[io.BytesIO, canvas.Canvas]:
    buff = io.BytesIO()
    c = canvas.Canvas(buff, pagesize=(page_w, page_h))
    c.drawImage(ImageReader(str(bg_path)), 0, 0, page_w, page_h, mask="auto")
    return buff, c


def text_width(font_name: str, size: int, text: str) -> float:
    return pdfmetrics.stringWidth(text, font_name, size)


def hex_to_rgb255(hex_color: str) -> Tuple[int, int, int]:
    hex_color = hex_color.strip().lstrip("#")
    if len(hex_color) != 6:
        raise ValueError(f"Некорректный HEX-цвет: {hex_color}")
    return tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))


def hex_to_reportlab_rgb(hex_color: str) -> Tuple[float, float, float]:
    r, g, b = hex_to_rgb255(hex_color)
    return (r / 255, g / 255, b / 255)


def normalize_tea_key(text: str) -> str:
    s = (text or "").strip().lower().replace("ё", "е")
    s = re.sub(r"[\u2010\u2011\u2012\u2013\u2014\u2212-]+", " ", s)
    s = re.sub(r"[^0-9a-zа-я\s]+", " ", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def tea_category_key(tea_type: str) -> Optional[str]:
    key = normalize_tea_key(tea_type)
    if not key:
        return None

    if key in TEA_CATEGORY_ALIASES:
        return TEA_CATEGORY_ALIASES[key]

    # Канонические категории и частые варианты ввода
    if key in {"габа", "gaba"} or key.startswith("габа ") or "габа" in key.split():
        return "габа"
    if key in {"светлый улун", "светлые улуны", "светлый улуны", "светл улун"}:
        return "светлый улун"
    if key in {"темный улун", "темные улуны", "темный улуны", "темн улун"}:
        return "темный улун"
    if key in {"красный", "красный чай"}:
        return "красный"
    if key in {"желтый", "желтый чай"}:
        return "жёлтый"
    if key in {"зеленый", "зеленый чай"}:
        return "зеленый"
    if key in {"белый", "белый чай"}:
        return "белый"
    if key in {"пуэр", "пуер"} or key.endswith(" пуэр") or key.endswith(" пуер"):
        return "пуэр"
    if key in {"травы", "травяной", "травяной чай", "травяные"}:
        return "травы"

    return None


def tea_asset_config(tea_type: str) -> Optional[Dict[str, object]]:
    key = tea_category_key(tea_type)
    if not key:
        return None
    return TEA_CATEGORY_ASSETS.get(key)


def tea_background_path(tea_type: str, kind: str) -> Path:
    config = tea_asset_config(tea_type)
    if config:
        path = config[kind]
        if isinstance(path, Path) and path.exists():
            return path

    # fallback на старые фоны, чтобы бот не падал при неизвестной категории
    return ASSET_TEA_BANK_BG if kind == "bank" else ASSET_TEA_BOX_BG


def tea_accent_rgb(tea_type: str) -> Tuple[float, float, float]:
    config = tea_asset_config(tea_type)
    if config:
        return hex_to_reportlab_rgb(str(config["hex"]))
    return ORANGE


def tea_accent_rgb255(tea_type: str) -> Tuple[int, int, int]:
    config = tea_asset_config(tea_type)
    if config:
        return hex_to_rgb255(str(config["hex"]))
    return (0xF5, 0x70, 0x44)


def break_long_word(word: str, font_name: str, size: int, max_width: float) -> List[str]:
    parts, buf = [], ""
    for ch in word:
        test = buf + ch
        if text_width(font_name, size, test) <= max_width:
            buf = test
        else:
            if buf:
                parts.append(buf)
                buf = ch
            else:
                parts.append(ch)
                buf = ""
    if buf:
        parts.append(buf)
    return parts


def wrap_lines(
    text: str,
    font_name: str,
    size: int,
    max_width: float,
    max_lines: int,
    allow_word_break: bool = True,
) -> Optional[List[str]]:
    """
    Разбивает текст на строки по словам.
    - allow_word_break=True: если встречается слово, которое не влезает, режем его по буквам.
    - allow_word_break=False: переносов по буквам НЕ делаем; если слово не влезает — возвращаем None.
    """
    text = re.sub(r"\s+", " ", (text or "").strip())
    if not text:
        return None

    words = text.split(" ")
    lines: List[str] = []
    current = ""

    def push(line: str):
        lines.append(line)

    i = 0
    while i < len(words):
        w = words[i]

        if text_width(font_name, size, w) > max_width:
            if not allow_word_break:
                return None
            pieces = break_long_word(w, font_name, size, max_width)
            words = words[:i] + pieces + words[i + 1 :]
            w = words[i]

        test = (current + " " + w).strip() if current else w
        if text_width(font_name, size, test) <= max_width:
            current = test
            i += 1
        else:
            push(current)
            current = ""
            if len(lines) >= max_lines:
                return None

    if current:
        push(current)

    return lines if len(lines) <= max_lines else None


def fit_text(
    text: str,
    font_name: str,
    max_size: int,
    min_size: int,
    max_width: float,
    max_lines: int,
    allow_word_break: bool = True,
) -> Tuple[int, List[str]]:
    """Подбирает максимально возможный размер шрифта, чтобы текст влез."""
    for size in range(max_size, min_size - 1, -1):
        lines = wrap_lines(text, font_name, size, max_width, max_lines, allow_word_break=allow_word_break)
        if lines:
            return size, lines

    size = min_size
    lines = wrap_lines(text, font_name, size, max_width, max_lines, allow_word_break=allow_word_break) or [
        ((text or "").strip()[:20] + "…").strip()
    ]
    return size, lines


def fit_text_in_box(
    text: str,
    font_name: str,
    max_size: int,
    min_size: int,
    max_width: float,
    max_lines: int,
    max_height: float,
    line_height: float,
    allow_word_break: bool = True,
) -> Tuple[int, List[str]]:
    """Как fit_text, но ещё учитывает высоту блока (max_height)."""
    for size in range(max_size, min_size - 1, -1):
        lines = wrap_lines(text, font_name, size, max_width, max_lines, allow_word_break=allow_word_break)
        if not lines:
            continue
        needed_h = size * line_height * len(lines)
        if needed_h <= max_height:
            return size, lines

    size = min_size
    lines = wrap_lines(text, font_name, size, max_width, max_lines, allow_word_break=allow_word_break) or [
        ((text or "").strip()[:20] + "…").strip()
    ]
    return size, lines


def draw_centered_multiline(
    c: canvas.Canvas,
    lines: List[str],
    font_name: str,
    font_size: int,
    center_x: float,
    center_y: float,
    color_rgb: Tuple[float, float, float],
    line_height: float = 1.2,
):
    c.setFillColorRGB(*color_rgb)
    c.setFont(font_name, font_size)
    block_h = font_size * line_height * len(lines)
    y = center_y + (block_h / 2) - font_size
    for line in lines:
        c.drawCentredString(center_x, y, line)
        y -= font_size * line_height


def draw_multiline_in_rect(
    c: canvas.Canvas,
    lines: List[str],
    font_name: str,
    font_size: int,
    x0: float,
    x1: float,
    center_y: float,
    color_rgb: Tuple[float, float, float],
    *,
    align: str = "center",
    line_height: float = 1.2,
):
    """
    Рисует многострочный текст внутри прямоугольника по X (x0..x1), по Y центрируется вокруг center_y.
    align: 'left' | 'center' | 'right'
    """
    c.setFillColorRGB(*color_rgb)
    c.setFont(font_name, font_size)
    block_h = font_size * line_height * len(lines)
    y = center_y + (block_h / 2) - font_size
    for line in lines:
        if align == "left":
            c.drawString(x0, y, line)
        elif align == "right":
            c.drawRightString(x1, y, line)
        else:
            c.drawCentredString((x0 + x1) / 2, y, line)
        y -= font_size * line_height


def fit_text_above_line(
    text: str,
    font_name: str,
    max_size: int,
    min_size: int,
    max_width: float,
    max_lines: int,
    *,
    y_top: float,
    y_line: float,
    clearance: float,
    line_height: float = 1.05,
    allow_word_break: bool = False,
) -> Tuple[int, List[str], float]:
    """
    Подбирает максимально крупный текст, который:
    - влезает по ширине/строкам
    - помещается в область от y_line+clearance до y_top
    - при этом низ глифов нижней строки выше y_line+clearance
    Возвращает: (size, lines, last_baseline_y)
    """
    text = re.sub(r"\s+", " ", (text or "").strip())
    if not text:
        return min_size, [""], y_line + clearance

    y_min_glyph = y_line + clearance

    for size in range(max_size, min_size - 1, -1):
        lines = wrap_lines(text, font_name, size, max_width, max_lines, allow_word_break=allow_word_break)
        if not lines:
            continue

        ascent = pdfmetrics.getAscent(font_name, size)
        descent = abs(pdfmetrics.getDescent(font_name, size))

        last_baseline = y_min_glyph + descent
        first_baseline = last_baseline + (len(lines) - 1) * (size * line_height)
        top_glyph = first_baseline + ascent

        if top_glyph <= y_top:
            return size, lines, last_baseline

    size = min_size
    lines = wrap_lines(text, font_name, size, max_width, max_lines, allow_word_break=allow_word_break) or [text]
    ascent = pdfmetrics.getAscent(font_name, size)
    descent = abs(pdfmetrics.getDescent(font_name, size))
    last_baseline = (y_line + clearance) + descent
    first_baseline = last_baseline + (len(lines) - 1) * (size * line_height)
    top_glyph = first_baseline + ascent

    if top_glyph > y_top:
        shift = top_glyph - y_top
        last_baseline -= shift

    return size, lines, last_baseline


def draw_multiline_above_line(
    c: canvas.Canvas,
    lines: List[str],
    font_name: str,
    font_size: int,
    center_x: float,
    last_baseline_y: float,
    color_rgb: Tuple[float, float, float],
    *,
    line_height: float = 1.05,
):
    """Рисует строки так, чтобы baseline последней строки был last_baseline_y (якорь снизу)."""
    c.setFillColorRGB(*color_rgb)
    c.setFont(font_name, font_size)
    first_baseline = last_baseline_y + (len(lines) - 1) * (font_size * line_height)
    y = first_baseline
    for line in lines:
        c.drawCentredString(center_x, y, line)
        y -= font_size * line_height


def draw_brand_ci(
    c: canvas.Canvas,
    fonts: Fonts,
    page_w: int,
    y: float,
    size: int,
    accent_rgb: Tuple[float, float, float] = ORANGE,
):
    """
    Рисует фирменную надпись как в исходной логике:
    буквы «Ч» и «И» — акцентным цветом, остальное — светлым.
    Используется для ценников товаров, чтобы их цвета текста остались как раньше.
    """
    font = fonts.medium
    seg1, seg2, seg3, seg4, seg5 = "Ч", "АЙНАЯ", " ", "И", "СТОРИЯ"
    w1 = text_width(font, size, seg1)
    w2 = text_width(font, size, seg2)
    w3 = text_width(font, size, seg3)
    w4 = text_width(font, size, seg4)
    w5 = text_width(font, size, seg5)
    total = w1 + w2 + w3 + w4 + w5
    x = (page_w - total) / 2

    c.setFont(font, size)
    c.setFillColorRGB(*accent_rgb)
    c.drawString(x, y, seg1)
    x += w1
    c.setFillColorRGB(*CREAM)
    c.drawString(x, y, seg2)
    x += w2
    c.drawString(x, y, seg3)
    x += w3
    c.setFillColorRGB(*accent_rgb)
    c.drawString(x, y, seg4)
    x += w4
    c.setFillColorRGB(*CREAM)
    c.drawString(x, y, seg5)


def draw_brand_white(
    c: canvas.Canvas,
    fonts: Fonts,
    page_w: int,
    y: float,
    size: int,
):
    """
    Рисует «ЧАЙНАЯ ИСТОРИЯ» полностью белым/светлым цветом.
    Используется для ценников чая, чтобы категория чая не меняла цвет букв.
    """
    font = fonts.medium
    text = "ЧАЙНАЯ ИСТОРИЯ"
    total = text_width(font, size, text)
    x = (page_w - total) / 2

    c.setFont(font, size)
    c.setFillColorRGB(*CREAM)
    c.drawString(x, y, text)

def format_price(price: float) -> str:
    if price == int(price):
        return f"{int(price)}₽"
    return f"{price:.2f}₽"


# -------------------------
# PDF генерация
# -------------------------
def make_pdf_products_two_sides(fonts: Fonts, name: str, price: float, hours: int) -> bytes:
    name = normalize_sentence_case(name)

    front_w, front_h = img_size(ASSET_PRODUCTS_FRONT_BG)
    back_w, back_h = img_size(ASSET_PRODUCTS_BACK_BG)

    # Основной размер PDF берём с переднего фона.
    # Если задний фон вдруг отличается по размеру, ReportLab переключит размер второй страницы.
    buff, c = pdf_with_background(front_w, front_h, ASSET_PRODUCTS_FRONT_BG)
    w, h = front_w, front_h

    # FRONT — логика, цвета и расположение текстов оставлены как раньше; поменян только фон.
    draw_brand_ci(c, fonts, w, y=585, size=36)

    size_name, lines_name = fit_text(name, fonts.bold, max_size=72, min_size=34, max_width=w - 140, max_lines=2)
    draw_centered_multiline(c, lines_name, fonts.bold, size_name, w / 2, 355, CREAM, line_height=1.15)

    price_text = format_price(price)
    size_price, lines_price = fit_text(
        price_text,
        fonts.semibold,
        max_size=60,
        min_size=28,
        max_width=w - 200,
        max_lines=1,
    )
    draw_centered_multiline(c, lines_price, fonts.semibold, size_price, w / 2, 80, ORANGE, line_height=1.0)

    c.showPage()

    # BACK — логика, цвета и расположение текстов оставлены как раньше; поменян только фон.
    if (back_w, back_h) != (w, h):
        c.setPageSize((back_w, back_h))
        w, h = back_w, back_h

    c.drawImage(ImageReader(str(ASSET_PRODUCTS_BACK_BG)), 0, 0, w, h, mask="auto")
    draw_brand_ci(c, fonts, w, y=585, size=36)

    title = "Срок хранения"
    size_t, lines_t = fit_text(title, fonts.bold, max_size=72, min_size=36, max_width=w - 160, max_lines=2)
    draw_centered_multiline(c, lines_t, fonts.bold, size_t, w / 2, 355, CREAM, line_height=1.05)

    phrase = f"{hours} {hours_word(hours)}"
    size_h, lines_h = fit_text(
        phrase,
        fonts.semibold,
        max_size=60,
        min_size=28,
        max_width=w - 200,
        max_lines=1,
    )
    draw_centered_multiline(c, lines_h, fonts.semibold, size_h, w / 2, 80, ORANGE, line_height=1.0)

    c.save()
    return buff.getvalue()

def make_pdf_tea_bank(fonts: Fonts, tea_type: str, name: str, price: float) -> bytes:
    bg_path = tea_background_path(tea_type, "bank")
    accent_rgb = tea_accent_rgb(tea_type)
    accent_rgb255 = tea_accent_rgb255(tea_type)

    w, h = img_size(bg_path)

    # Дизайн раньше был размечен в координатах 1654x1654.
    # Так текст не сломается, если фон будет 827x827 или 1654x1654.
    design_w = 1654
    design_h = 1654
    sx = w / design_w
    sy = h / design_h
    s = min(sx, sy)

    def x(v: float) -> float:
        return v * sx

    def y(v: float) -> float:
        return v * sy

    def fs(v: int) -> int:
        return max(1, int(round(v * s)))

    # Рамка-запас под обрез при печати.
    margin = fs(40)

    bg = Image.open(bg_path).convert("RGBA")
    bg_flat = Image.new("RGBA", bg.size, (*accent_rgb255, 255))
    bg_flat = Image.alpha_composite(bg_flat, bg)

    expanded = Image.new("RGBA", (w + margin * 2, h + margin * 2), (*accent_rgb255, 255))
    expanded.paste(bg_flat, (margin, margin))

    bg_bytes = io.BytesIO()
    expanded.save(bg_bytes, format="PNG")
    bg_bytes.seek(0)

    page_w = w + margin * 2
    page_h = h + margin * 2
    buff = io.BytesIO()
    c = canvas.Canvas(buff, pagesize=(page_w, page_h))
    c.drawImage(ImageReader(bg_bytes), 0, 0, page_w, page_h, mask="auto")

    # сдвигаем систему координат, чтобы (0,0) = левый нижний угол оригинального фона
    c.translate(margin, margin)

    # ---------- Бренд
    draw_brand_white(c, fonts, w, y=y(1505), size=fs(70))

    # ---------- Категория цены
    cat = category_from_price(price)
    size_cat, lines_cat = fit_text_in_box(
        cat,
        fonts.medium,
        max_size=fs(85),
        min_size=fs(18),
        max_width=w - x(220),
        max_lines=1,
        max_height=y(120),
        line_height=1.0,
        allow_word_break=False,
    )
    draw_centered_multiline(c, lines_cat, fonts.medium, size_cat, w / 2, y(210), CREAM, line_height=1.0)

    # ---------- Тип чая
    y_top = y(1425)
    y_line = y(855)
    clearance = y(36)
    size_tt, lines_tt, last_baseline = fit_text_above_line(
        tea_type,
        fonts.bold,
        max_size=fs(180),
        min_size=fs(34),
        max_width=w - x(240),
        max_lines=2,
        y_top=y_top,
        y_line=y_line,
        clearance=clearance,
        line_height=1.05,
        allow_word_break=False,
    )
    draw_multiline_above_line(c, lines_tt, fonts.bold, size_tt, w / 2, last_baseline, accent_rgb, line_height=1.05)

    # верхняя линия
    c.setStrokeColorRGB(*LINE)
    c.setLineWidth(fs(6))
    c.setLineCap(1)
    c.line(x(190), y(855), w - x(190), y(855))

    # ---------- Название
    top_y2 = y(840)
    bottom_y2 = y(615)
    box_h2 = top_y2 - bottom_y2
    center_y2 = (top_y2 + bottom_y2) / 2
    size_nm, lines_nm = fit_text_in_box(
        name,
        fonts.medium,
        max_size=fs(124),
        min_size=fs(22),
        max_width=w - x(260),
        max_lines=2,
        max_height=box_h2,
        line_height=1.2,
        allow_word_break=False,
    )
    draw_centered_multiline(c, lines_nm, fonts.medium, size_nm, w / 2, center_y2, CREAM, line_height=1.2)

    # нижняя линия
    c.setStrokeColorRGB(*LINE)
    c.setLineWidth(fs(6))
    c.setLineCap(1)
    c.line(x(260), y(600), w - x(260), y(600))

    # ---------- Цена
    price_text = format_price(price)
    size_p, lines_p = fit_text(
        price_text,
        fonts.bold,
        max_size=fs(110),
        min_size=fs(30),
        max_width=w - x(340),
        max_lines=1,
        allow_word_break=False,
    )
    draw_centered_multiline(c, lines_p, fonts.bold, size_p, w / 2, y(470), accent_rgb, line_height=1.0)

    c.save()
    return buff.getvalue()


def make_pdf_tea_box(fonts: Fonts, tea_type: str, name: str, price: float) -> bytes:
    bg_path = tea_background_path(tea_type, "box")
    accent_rgb = tea_accent_rgb(tea_type)

    w, h = img_size(bg_path)
    buff, c = pdf_with_background(w, h, bg_path)

    design_w = 1890
    design_h = 236
    sx = w / design_w
    sy = h / design_h
    s = min(sx, sy)

    def x(v: float) -> float:
        return v * sx

    def y(v: float) -> float:
        return v * sy

    def fs(v: int) -> int:
        return max(1, int(round(v * s)))

    x1, x2 = x(457), x(1524)

    c.setStrokeColorRGB(*LINE)
    c.setLineWidth(fs(6))
    c.setLineCap(1)
    c.line(x1, y(30), x1, h - y(30))
    c.line(x2, y(30), x2, h - y(30))

    pad = x(60)
    left_x0, left_x1 = pad, x1 - pad
    mid_x0, mid_x1 = x1 + pad, x2 - pad
    right_x0, right_x1 = x2 + pad, w - pad

    size_tt, lines_tt = fit_text(
        tea_type,
        fonts.bold,
        max_size=fs(72),
        min_size=fs(28),
        max_width=(left_x1 - left_x0),
        max_lines=2,
        allow_word_break=False,
    )
    draw_multiline_in_rect(
        c,
        lines_tt,
        fonts.bold,
        size_tt,
        left_x0,
        left_x1,
        h / 2,
        accent_rgb,
        align="left",
        line_height=1.0,
    )

    size_nm, lines_nm = fit_text(
        name,
        fonts.medium,
        max_size=fs(60),
        min_size=fs(24),
        max_width=(mid_x1 - mid_x0),
        max_lines=2,
    )
    draw_centered_multiline(
        c,
        lines_nm,
        fonts.medium,
        size_nm,
        (mid_x0 + mid_x1) / 2,
        h / 2 + y(3),
        CREAM,
        line_height=1.2,
    )

    price_text = format_price(price)
    size_p, lines_p = fit_text(
        price_text,
        fonts.bold,
        max_size=fs(60),
        min_size=fs(24),
        max_width=(right_x1 - right_x0),
        max_lines=1,
    )
    draw_multiline_in_rect(
        c,
        lines_p,
        fonts.bold,
        size_p,
        right_x0,
        right_x1,
        y(150),
        accent_rgb,
        align="right",
        line_height=1.0,
    )

    cat = category_from_price(price)
    size_c, lines_c = fit_text(
        cat,
        fonts.medium,
        max_size=fs(36),
        min_size=fs(18),
        max_width=(right_x1 - right_x0),
        max_lines=1,
    )
    draw_multiline_in_rect(
        c,
        lines_c,
        fonts.medium,
        size_c,
        right_x0,
        right_x1,
        y(80),
        CREAM,
        align="right",
        line_height=1.0,
    )

    c.save()
    return buff.getvalue()


def make_styled_qr_png(data: str, size_px: int = 180) -> bytes:
    import qrcode
    from qrcode.constants import ERROR_CORRECT_H

    try:
        from qrcode.image.styledpil import StyledPilImage
        from qrcode.image.styles.colormasks import SolidFillColorMask
        from qrcode.image.styles.eyedrawers import RoundedEyeDrawer
        from qrcode.image.styles.moduledrawers import RoundedModuleDrawer

        qr = qrcode.QRCode(version=None, error_correction=ERROR_CORRECT_H, box_size=10, border=1)
        qr.add_data(data)
        qr.make(fit=True)
        img = qr.make_image(
            image_factory=StyledPilImage,
            module_drawer=RoundedModuleDrawer(),
            eye_drawer=RoundedEyeDrawer(),
            color_mask=SolidFillColorMask(back_color=f"#{QR_BG_HEX}", front_color=f"#{QR_FG_HEX}"),
        ).convert("RGBA")
    except Exception:
        qr = qrcode.QRCode(version=None, error_correction=ERROR_CORRECT_H, box_size=10, border=1)
        qr.add_data(data)
        qr.make(fit=True)
        img = qr.make_image(fill_color=f"#{QR_FG_HEX}", back_color=f"#{QR_BG_HEX}").convert("RGBA")

    img = img.resize((size_px, size_px), Image.NEAREST)
    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()


def make_pdf_tips_two_sides(fonts: Fonts, person_name: str, goal: str, link: str) -> bytes:
    w, h = img_size(ASSET_TIPS_FRONT_BG)
    buff, c = pdf_with_background(w, h, ASSET_TIPS_FRONT_BG)

    size_n, lines_n = fit_text(
        person_name,
        fonts.bold,
        max_size=86,
        min_size=28,
        max_width=w - 120,
        max_lines=2,
        allow_word_break=False,
    )

    size_g, lines_g = fit_text(
        goal,
        fonts.regular,
        max_size=48,
        min_size=16,
        max_width=w - 140,
        max_lines=3,
        allow_word_break=False,
    )

    name_lh = 1.05
    goal_lh = 1.15

    base_gap = 20
    gap = base_gap
    if len(lines_n) >= 2 or len(lines_g) >= 2:
        gap = 14
    if len(lines_n) >= 2 and len(lines_g) >= 2:
        gap = 10

    name_h = size_n * name_lh * len(lines_n)
    goal_h = size_g * goal_lh * len(lines_g)
    group_center_y = 505
    name_center_y = group_center_y + (goal_h + gap) / 2
    goal_center_y = group_center_y - (name_h + gap) / 2

    safe_goal_bottom = 330
    goal_bottom = goal_center_y - goal_h / 2
    if goal_bottom < safe_goal_bottom:
        shift = safe_goal_bottom - goal_bottom
        name_center_y += shift
        goal_center_y += shift

    draw_centered_multiline(c, lines_n, fonts.bold, size_n, w / 2, name_center_y, CREAM, line_height=name_lh)
    draw_centered_multiline(c, lines_g, fonts.regular, size_g, w / 2, goal_center_y, ORANGE, line_height=goal_lh)

    box_left = 200
    box_bottom = 62
    box_size = 249
    qr_size = 180
    qr_left = int(box_left + (box_size - qr_size) / 2)
    qr_bottom = int(box_bottom + (box_size - qr_size) / 2)

    qr_png = make_styled_qr_png(link, size_px=qr_size)
    c.drawImage(ImageReader(io.BytesIO(qr_png)), qr_left, qr_bottom, qr_size, qr_size, mask="auto")

    c.showPage()
    c.drawImage(ImageReader(str(ASSET_TIPS_BACK_BG)), 0, 0, w, h, mask="auto")

    c.save()
    return buff.getvalue()


# -------------------------
# Excel шаблоны и чтение
# -------------------------
def build_xlsx_tea_template() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Чай"

    headers = ["Тип чая", "Наименование", "Цена (число)"]
    ws.append(headers)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="231F20")
    thin_side = Side(style="thin", color="C1BAB1")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 62

    for _ in range(2, 202):
        ws.append(["", "", ""])

    # Подсказка справа от таблицы — сотрудники видят её сразу,
    # но она не мешает чтению данных из колонок A:C.
    hint_title_fill = PatternFill("solid", fgColor="F6763C")
    hint_header_fill = PatternFill("solid", fgColor="231F20")
    hint_light_fill = PatternFill("solid", fgColor="F4EFE8")
    hint_font = Font(color="231F20")
    hint_bold_font = Font(bold=True, color="231F20")
    hint_white_bold = Font(bold=True, color="FFFFFF")

    ws.merge_cells("E1:F1")
    title_cell = ws["E1"]
    title_cell.value = "ПОДСКАЗКА ПО ЗАПОЛНЕНИЮ ЧАЙНЫХ ЦЕННИКОВ"
    title_cell.font = hint_white_bold
    title_cell.fill = hint_title_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    hint_rows = [
        ("Как заполнять", "В колонке «Тип чая» пишите один из вариантов ниже. Бот сам подберёт нужный цвет рамки и картинки."),
        ("Темные улуны", "Темный Улун, ФХДЦ, ДХП"),
        ("Светлые улуны", "Светлый Улун, Те Гуань Инь"),
        ("Красный", "Красный"),
        ("Пуэры", "Шу Пуэр, Лао Шен Пуэр, Шен Пуэр"),
        ("Жёлтый", "Жёлтый"),
        ("Зеленый", "Зеленый"),
        ("Габа", "Пишите просто «Габа». Если чай называется шу габа или шен габа — всё равно пишите только «Габа»."),
        ("Белый", "Белый"),
        ("Цена", "Пишите только число без букв и без ₽. Можно через точку или запятую: 22, 22.50 или 22,50."),
    ]

    for row_idx, (label, text) in enumerate(hint_rows, start=2):
        label_cell = ws.cell(row=row_idx, column=5)
        text_cell = ws.cell(row=row_idx, column=6)
        label_cell.value = label
        text_cell.value = text

        label_cell.font = hint_white_bold if row_idx == 2 else hint_bold_font
        text_cell.font = hint_white_bold if row_idx == 2 else hint_font
        label_cell.fill = hint_header_fill if row_idx == 2 else hint_light_fill
        text_cell.fill = hint_header_fill if row_idx == 2 else hint_light_fill
        label_cell.border = thin_border
        text_cell.border = thin_border
        label_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        text_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    for row_idx in range(1, 12):
        ws.row_dimensions[row_idx].height = 34 if row_idx != 1 else 42

    ws.freeze_panes = "A2"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def build_xlsx_products_template() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Товары"
    headers = ["Название", "Цена (число)", "Срок хранения (часы, число)"]
    ws.append(headers)

    header_font = Font(bold=True)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 28

    for _ in range(2, 202):
        ws.append(["", "", ""])

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def load_rows_tea(xlsx_bytes: bytes) -> List[Tuple[str, str, float]]:
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb.active
    rows = []

    for r in ws.iter_rows(min_row=2, values_only=True):
        tea_type, name, price = (r[0], r[1], r[2]) if len(r) >= 3 else (None, None, None)

        if (
            (tea_type is None or str(tea_type).strip() == "")
            and (name is None or str(name).strip() == "")
            and (price is None or str(price).strip() == "")
        ):
            continue

        tea_type_s = str(tea_type).strip() if tea_type is not None else ""
        name_s = str(name).strip() if name is not None else ""
        price_f = parse_float_number(price)

        if not tea_type_s:
            raise ValueError("В Excel для чая найдено пустое поле «Тип чая».")
        if not name_s:
            raise ValueError("В Excel для чая найдено пустое поле «Наименование».")
        if price_f is None or price_f < 0 or price_f > 1_000_000:
            raise ValueError("В Excel для чая «Цена» должна быть числом (0…1000000). Можно вводить 22.50 или 22,50.")

        rows.append((tea_type_s, name_s, price_f))

    if not rows:
        raise ValueError("Excel пустой: заполни хотя бы одну строку.")

    return rows


def load_rows_products(xlsx_bytes: bytes) -> List[Tuple[str, float, int]]:
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb.active
    rows = []

    for r in ws.iter_rows(min_row=2, values_only=True):
        name, price, hours = (r[0], r[1], r[2]) if len(r) >= 3 else (None, None, None)

        if (
            (name is None or str(name).strip() == "")
            and (price is None or str(price).strip() == "")
            and (hours is None or str(hours).strip() == "")
        ):
            continue

        name_s = str(name).strip() if name is not None else ""
        price_f = parse_float_number(price)
        hours_i = parse_int_number(hours)

        if not name_s:
            raise ValueError("В Excel для товаров найдено пустое поле «Название».")
        if price_f is None or price_f < 0 or price_f > 1_000_000:
            raise ValueError("В Excel для товаров «Цена» должна быть числом (0…1000000). Можно вводить 22.50 или 22,50.")
        if hours_i is None or hours_i < 0 or hours_i > 24 * 365:
            raise ValueError("В Excel для товаров «Срок хранения» должен быть числом часов (0…8760).")

        rows.append((name_s, price_f, hours_i))

    if not rows:
        raise ValueError("Excel пустой: заполни хотя бы одну строку.")

    return rows
