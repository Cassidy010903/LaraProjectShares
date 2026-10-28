import os
import sys
from openpyxl.styles import Font

GREEN_DARK = "083416"
GREEN = "2D6A35"
YELLOW = "EFBE22"
YELLOW_LIGHT = "FBEDB8"
GREY_DARK = "424242"
GREY_MID = "8E8E8E"
GREY_LIGHT = "E1E1E1"
WHITE = "FFFFFF"
FONT = "Arial"


def resource_path(relative):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


def add_logo(ws, anchor, height=30):
    try:
        from openpyxl.drawing.image import Image
        img = Image(resource_path(os.path.join("assets", "logo_black.png")))
        img.width = int(img.width * height / img.height)
        img.height = height
        ws.add_image(img, anchor)
    except Exception:
        pass


def add_title(ws, text, logo_anchor="D1"):
    ws["A1"].value = text
    ws["A1"].font = Font(name=FONT, bold=True, size=14, color=GREEN_DARK)
    ws.row_dimensions[1].height = 34
    add_logo(ws, logo_anchor)


TK_GREEN_DARK = "#" + GREEN_DARK
TK_GREEN = "#" + GREEN
TK_YELLOW = "#" + YELLOW
TK_GREY_DARK = "#" + GREY_DARK
TK_GREY_MID = "#" + GREY_MID
TK_GREY_LIGHT = "#" + GREY_LIGHT
TK_WHITE = "#" + WHITE
