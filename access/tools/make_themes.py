"""Generate the VBA palette modules: src/themes/light/modTheme.bas (the build
palette, constants used when the forms are created) and src/modules/modPalette.bas
(the named palettes users pick from, and PaletteOf, which maps each build color
to the chosen palette at run time).

VBA constants cannot call RGB(), so colors are written as Long values
(R + G*256 + B*65536). Edit the palettes here and re-run:

    python tools/make_themes.py [app_src_folder]

app_src_folder defaults to ../src next to this script. The script writes
<app_src_folder>/themes/light/modTheme.bas and <app_src_folder>/modules/modPalette.bas.

The script also prints a text-contrast check for every palette.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.join(HERE, "..", "src")
OUT = os.path.join(ROOT, "themes")

# Color roles: name -> (light, dark), the two base columns the palettes below
# start from.  Light reproduces the original web-app look;
# dark follows the "Auto Guard" dashboard style (navy surfaces, white CTAs).
P = {
    # Surfaces
    "CLR_BG":              ((255, 255, 255), (10, 14, 30)),      # page background
    "CLR_SURFACE":         ((255, 255, 255), (14, 24, 44)),      # cards, panels, list rows
    "CLR_INSET":           ((248, 250, 252), (6, 11, 24)),       # inset tiles, banners
    "CLR_DIALOG":          ((255, 255, 255), (14, 24, 44)),
    "CLR_ROW1":            ((241, 245, 249), (24, 37, 62)),      # tree level 1
    "CLR_ROW2":            ((248, 250, 252), (18, 30, 52)),      # tree level 2
    "CLR_INPUT_BG":        ((255, 255, 255), (6, 11, 24)),
    "CLR_CELL_BG":         ((226, 232, 240), (30, 44, 72)),      # attribute score/band cells
    "CLR_BORDER":          ((226, 232, 240), (33, 45, 72)),
    "CLR_INPUT_BORDER":    ((203, 213, 225), (48, 62, 94)),
    # Text
    "CLR_FG":              ((15, 23, 42), (241, 245, 249)),
    "CLR_MUTED_FG":        ((100, 116, 139), (148, 160, 184)),
    "CLR_ATTR_FG":         ((51, 65, 85), (203, 213, 225)),
    "CLR_WHITE":           ((255, 255, 255), (255, 255, 255)),   # text on colored fills
    "CLR_ERROR":           ((220, 38, 38), (248, 113, 113)),
    "CLR_DANGER_TEXT":     ((185, 28, 28), (248, 113, 113)),
    "CLR_WARN_TEXT":       ((180, 83, 9), (251, 146, 60)),
    "CLR_GREEN_600":       ((22, 163, 74), (74, 222, 128)),
    "CLR_AMBER_500":       ((245, 158, 11), (251, 191, 36)),
    # Accent
    "CLR_PRIMARY":         ((46, 109, 180), (59, 130, 246)),
    "CLR_PRIMARY_SOFT":    ((236, 243, 252), (23, 40, 78)),
    "CLR_PRIMARY_LINE":    ((191, 211, 236), (37, 64, 120)),
    "CLR_LINK":            ((46, 109, 180), (96, 165, 250)),
    "CLR_DESTRUCTIVE":     ((220, 38, 38), (220, 38, 38)),
    # Buttons
    "CLR_BTN_BG":          ((46, 109, 180), (255, 255, 255)),
    "CLR_BTN_FG":          ((255, 255, 255), (10, 14, 30)),
    "CLR_BTN_HOVER":       ((37, 91, 153), (226, 232, 240)),
    "CLR_BTN_PRESSED":     ((30, 75, 128), (203, 213, 225)),
    "CLR_OUT_BG":          ((255, 255, 255), (14, 24, 44)),
    "CLR_OUT_FG":          ((15, 23, 42), (241, 245, 249)),
    "CLR_OUT_HOVER":       ((241, 245, 249), (30, 44, 72)),
    "CLR_OUT_BORDER":      ((203, 213, 225), (226, 232, 240)),
    # Sidebar
    "CLR_SIDEBAR":         ((27, 58, 92), (20, 23, 43)),
    "CLR_SIDEBAR_LINE":    ((52, 79, 108), (38, 42, 66)),
    "CLR_NAV_FG":          ((187, 196, 206), (205, 210, 224)),
    "CLR_NAV_ACTIVE_BG":   ((61, 88, 121), (255, 255, 255)),
    "CLR_NAV_ACTIVE_FG":   ((255, 255, 255), (10, 14, 30)),
    "CLR_NAV_MUTED":       ((141, 156, 173), (130, 138, 160)),
    "CLR_BRAND":           ((27, 58, 92), (255, 255, 255)),
    "CLR_BLUE_300":        ((147, 197, 253), (147, 197, 253)),
    "CLR_AVATAR":          ((59, 130, 246), (59, 130, 246)),
    # Badges (bg, fg)
    "CLR_SUCCESS_BG":      ((220, 252, 231), (16, 58, 38)),
    "CLR_SUCCESS_FG":      ((22, 101, 52), (74, 222, 128)),
    "CLR_WARNING_BG":      ((254, 243, 199), (66, 45, 12)),
    "CLR_WARNING_FG":      ((146, 64, 14), (251, 191, 36)),
    "CLR_INFO_BG":         ((219, 234, 254), (23, 40, 84)),
    "CLR_INFO_FG":         ((30, 64, 175), (96, 165, 250)),
    "CLR_SECONDARY_BG":    ((226, 232, 240), (30, 41, 64)),
    "CLR_SECONDARY_FG":    ((15, 23, 42), (203, 213, 225)),
    "CLR_SEVA_BG":         ((254, 226, 226), (69, 16, 16)),
    "CLR_SEVA_FG":         ((153, 27, 27), (252, 165, 165)),
    "CLR_SEVB_BG":         ((255, 237, 213), (67, 30, 8)),
    "CLR_SEVB_FG":         ((154, 52, 18), (253, 186, 116)),
    # KPI tile accents (dark dashboard)
    "CLR_KPI1":            ((59, 130, 246), (59, 130, 246)),
    "CLR_KPI2":            ((129, 140, 248), (129, 140, 248)),
    "CLR_KPI3":            ((74, 222, 128), (110, 200, 140)),
    "CLR_KPI4":            ((248, 113, 113), (248, 113, 113)),
    "CLR_KPI5":            ((251, 146, 60), (251, 146, 60)),
    "CLR_KPI6":            ((244, 114, 182), (244, 114, 182)),
    # Print palette (reports are always printed on white paper)
    # Printed reports are black and white (grays only) in both themes.
    "PRINT_FG":            ((0, 0, 0), (0, 0, 0)),
    "PRINT_MUTED":         ((64, 64, 64), (64, 64, 64)),
    "PRINT_TITLE":         ((0, 0, 0), (0, 0, 0)),
    "PRINT_ACCENT":        ((0, 0, 0), (0, 0, 0)),
    "PRINT_ROW1":          ((230, 230, 230), (230, 230, 230)),
    "PRINT_PAPER":         ((255, 255, 255), (255, 255, 255)),   # opaque white behind report text
}

# Colors that are the same in every palette: status band pills (a band always
# reads the same) and the
# destructive button's hover shades. They are written as constants in
# modTheme and reserved: build colors must differ from these so a palette
# switch never repaints them.
FIXED = {
    "CLR_BAND_OUTSTANDING":    (147, 51, 234),   # purple-600
    "CLR_BAND_ABOVE":          (37, 99, 235),    # blue-600
    "CLR_BAND_AT":             (22, 163, 74),    # green-600
    "CLR_BAND_BELOW":          (245, 158, 11),   # amber-500
    "CLR_BAND_SIG_BELOW":      (185, 28, 28),    # red-700
    "CLR_BAND_NE":             (107, 114, 128),  # gray-500, Not Evaluated
    "CLR_DESTRUCTIVE_HOVER":   (185, 28, 28),
    "CLR_DESTRUCTIVE_PRESSED": (153, 27, 27),
}
RESERVED = set(FIXED.values())

WHITE, BLACK = (255, 255, 255), (0, 0, 0)

def mix(a, b, t):
    return tuple(round(x + (y - x) * t) for x, y in zip(a, b))

def accent_light(sidebar, primary):
    """Overrides for a light palette from its sidebar and accent colors."""
    return {
        "CLR_SIDEBAR": sidebar,
        "CLR_SIDEBAR_LINE": mix(sidebar, WHITE, 0.15),
        "CLR_NAV_FG": mix(sidebar, WHITE, 0.72),
        "CLR_NAV_ACTIVE_BG": mix(sidebar, WHITE, 0.18),
        "CLR_NAV_MUTED": mix(sidebar, WHITE, 0.5),
        "CLR_BRAND": sidebar,
        "CLR_PRIMARY": primary,
        "CLR_PRIMARY_SOFT": mix(primary, WHITE, 0.92),
        "CLR_PRIMARY_LINE": mix(primary, WHITE, 0.7),
        "CLR_LINK": primary,
        "CLR_BTN_BG": primary,
        "CLR_BTN_HOVER": mix(primary, BLACK, 0.17),
        "CLR_BTN_PRESSED": mix(primary, BLACK, 0.33),
        "CLR_AVATAR": mix(primary, WHITE, 0.15),
        "CLR_KPI1": primary,
    }

# The palettes users choose from (Appearance in the sidebar), in menu order.
# key: saved in the registry ("light" and "dark" are the original two, so a
# saved choice from before palettes still works). base: the column of P the
# palette starts from (0 light, 1 dark). dark: palettes built on the dark
# base also show the dark-only extras (accent bars).
PALETTES = [
    dict(key="light", name="Navy", base=0, over={}),
    dict(key="dark", name="Midnight", base=1, over={}),
    dict(key="slate", name="Slate", base=0, over=accent_light((51, 65, 85), (71, 85, 105))),
    dict(key="forest", name="Forest", base=0, over=accent_light((31, 61, 43), (47, 125, 79))),
    dict(key="harbor", name="Harbor", base=0, over=accent_light((15, 59, 76), (14, 116, 144))),
    # Submarine Program Manager's look: navy sidebar, teal accent. SPM's teal
    # (0, 132, 151) is a shade darker here so white text on it passes 4.5:1.
    dict(key="tealnavy", name="Teal Navy", base=0, over=accent_light((10, 42, 73), (0, 128, 147))),
    dict(key="graphite", name="Graphite", base=1, over={
        "CLR_BG": (12, 12, 14), "CLR_SURFACE": (24, 24, 27), "CLR_INSET": (9, 9, 11),
        "CLR_DIALOG": (24, 24, 27), "CLR_ROW1": (39, 39, 42), "CLR_ROW2": (31, 31, 34),
        "CLR_INPUT_BG": (9, 9, 11), "CLR_CELL_BG": (45, 45, 50), "CLR_BORDER": (50, 50, 56),
        "CLR_INPUT_BORDER": (70, 70, 78), "CLR_MUTED_FG": (161, 161, 170),
        "CLR_PRIMARY": (130, 130, 140), "CLR_PRIMARY_SOFT": (40, 40, 46), "CLR_PRIMARY_LINE": (70, 70, 80),
        "CLR_LINK": (212, 212, 216), "CLR_OUT_BG": (24, 24, 27), "CLR_OUT_HOVER": (39, 39, 42),
        "CLR_SIDEBAR": (18, 18, 20), "CLR_SIDEBAR_LINE": (40, 40, 44), "CLR_NAV_FG": (212, 212, 216),
        "CLR_NAV_MUTED": (130, 130, 140), "CLR_AVATAR": (113, 113, 122),
        "CLR_INFO_BG": (39, 39, 45), "CLR_INFO_FG": (212, 212, 216),
        "CLR_SECONDARY_BG": (39, 39, 42), "CLR_KPI1": (161, 161, 170), "CLR_KPI2": (212, 212, 216),
    }),
    dict(key="contrast", name="High Contrast", base=0, over={
        "CLR_FG": BLACK, "CLR_MUTED_FG": (64, 64, 64), "CLR_ATTR_FG": (32, 32, 32),
        "CLR_ROW1": (235, 235, 235), "CLR_ROW2": (245, 245, 245), "CLR_CELL_BG": (220, 220, 220),
        "CLR_BORDER": (118, 118, 118), "CLR_INPUT_BORDER": BLACK,
        "CLR_ERROR": (170, 0, 0), "CLR_DANGER_TEXT": (150, 0, 0), "CLR_WARN_TEXT": (130, 60, 0),
        "CLR_GREEN_600": (0, 110, 40), "CLR_AMBER_500": (160, 95, 0),
        "CLR_PRIMARY": (0, 64, 192), "CLR_PRIMARY_SOFT": (224, 234, 255), "CLR_PRIMARY_LINE": (0, 64, 192),
        "CLR_LINK": (0, 51, 160), "CLR_DESTRUCTIVE": (170, 0, 0),
        "CLR_BTN_BG": (0, 64, 192), "CLR_BTN_HOVER": (0, 48, 150), "CLR_BTN_PRESSED": (0, 36, 112),
        "CLR_OUT_FG": BLACK, "CLR_OUT_HOVER": (235, 235, 235), "CLR_OUT_BORDER": BLACK,
        "CLR_SIDEBAR": BLACK, "CLR_SIDEBAR_LINE": (90, 90, 90), "CLR_NAV_FG": WHITE,
        "CLR_NAV_ACTIVE_BG": (0, 64, 192), "CLR_NAV_MUTED": (200, 200, 200), "CLR_BRAND": BLACK,
        "CLR_AVATAR": (0, 64, 192),
        "CLR_SUCCESS_FG": (0, 80, 30), "CLR_WARNING_FG": (110, 50, 0), "CLR_INFO_FG": (0, 40, 140),
        "CLR_SECONDARY_FG": BLACK, "CLR_SEVA_FG": (120, 0, 0), "CLR_SEVB_FG": (120, 40, 0),
        "CLR_KPI1": (0, 64, 192),
    }),
]
for pal in PALETTES:
    assert not set(pal["over"]) - set(P), (pal["key"], set(pal["over"]) - set(P))
    pal["c"] = {k: pal["over"].get(k, P[k][pal["base"]]) for k in P}
    pal["dark"] = pal["base"] == 1
assert PALETTES[0]["key"] == "light" and not PALETTES[0]["over"], "palette 1 must be the build palette"

def to_long(rgb):
    r, g, b = rgb
    return r + g * 256 + b * 65536

def nudge(rgb, step):
    """A barely different shade: move channels by 1-2 toward mid-gray."""
    deltas = [(0, 0, 1), (0, 1, 0), (1, 0, 0), (0, 1, 1), (1, 0, 1), (1, 1, 0), (1, 1, 1),
              (0, 0, 2), (0, 2, 0), (2, 0, 0), (0, 2, 2), (2, 0, 2), (2, 2, 0), (2, 2, 2),
              (1, 2, 0), (2, 1, 0), (0, 1, 2), (0, 2, 1), (1, 0, 2), (2, 0, 1)]
    d = deltas[step]
    return tuple(c - k if c > 127 else c + k for c, k in zip(rgb, d))

# The forms are built once with the light palette; each color gets a unique
# value (nudged by a unit or two where the palette repeats a color) so that
# the runtime palette switch can tell every color role apart.
# Print colors are left alone: reports are never repainted.
ui = [k for k in P if not k.startswith("PRINT_")]
# Roles that look the same in both themes keep their exact color first.
ui.sort(key=lambda k: P[k][0] != P[k][1])
used, build = set(RESERVED), {}
for key in ui:
    rgb, step = P[key][0], 0
    cand = rgb
    while cand in used:
        cand = nudge(rgb, step)
        step += 1
    used.add(cand)
    build[key] = cand
for key in P:
    if key.startswith("PRINT_"):
        build[key] = P[key][0]

def const(key, rgb):
    v = to_long(rgb)
    return f"Public Const {key} As Long = {v}{' ' * max(1, 22 - len(key) - len(str(v)))}' RGB{rgb}"

lines = [
    "Option Compare Database",
    "Option Explicit",
    "",
    "' Build palette (the light theme).",
    "' Generated by tools/make_themes.py - edit the palette there.",
    "' Every UI color is unique, so modPalette.PaletteOf can map it to the",
    "' palette the user picks.",
    "",
    "Public Const THEME_DARK As Boolean = False   ' forms are always built light",
    "Public Const FONT_UI As String = \"Segoe UI\"",
    "",
]
for key in P:
    lines.append(const(key, build[key]))
lines += ["", "' The same in every palette (never repainted)."]
for key, rgb in FIXED.items():
    lines.append(const(key, rgb))
path = os.path.join(OUT, "light")
os.makedirs(path, exist_ok=True)
with open(os.path.join(path, "modTheme.bas"), "w", newline="\r\n") as f:
    f.write("\n".join(lines) + "\n")
print("wrote", os.path.join(path, "modTheme.bas"))

def proc_name(pal):
    return "Pal" + pal["key"].capitalize()

lines = [
    "Option Compare Database",
    "Option Explicit",
    "",
    "' The color palettes users pick from.",
    "' Generated by tools/make_themes.py - edit the palettes there.",
    "' Palette 1 is the build palette (modTheme). PaletteOf maps a build color",
    "' to palette i and returns any other color unchanged.",
    "",
    f"Public Const PALETTE_COUNT As Long = {len(PALETTES)}",
    "",
    "' Name saved in the registry for palette i.",
    "Public Function PaletteKey(ByVal i As Long) As String",
    "    Select Case i",
]
lines += [f"        Case {i}: PaletteKey = \"{p['key']}\"" for i, p in enumerate(PALETTES, 1)]
lines += [
    "    End Select",
    "End Function",
    "",
    "' Name shown in the Appearance dialog.",
    "Public Function PaletteName(ByVal i As Long) As String",
    "    Select Case i",
]
lines += [f"        Case {i}: PaletteName = \"{p['name']}\"" for i, p in enumerate(PALETTES, 1)]
lines += [
    "    End Select",
    "End Function",
    "",
    "' True for palettes with dark surfaces (they also show the dark-only extras).",
    "Public Function PaletteIsDark(ByVal i As Long) As Boolean",
    "    Select Case i",
    "        Case " + ", ".join(str(i) for i, p in enumerate(PALETTES, 1) if p["dark"]) + ": PaletteIsDark = True",
    "    End Select",
    "End Function",
    "",
    "' Swatch k of palette i for the Appearance dialog: 1 sidebar, 2 accent, 3 page.",
    "Public Function PaletteSwatch(ByVal i As Long, ByVal k As Long) As Long",
    "    Select Case i * 10 + k",
]
for i, p in enumerate(PALETTES, 1):
    for k, role in enumerate(("CLR_SIDEBAR", "CLR_PRIMARY", "CLR_BG"), 1):
        lines.append(f"        Case {i * 10 + k}: PaletteSwatch = {to_long(p['c'][role])}")
lines += [
    "    End Select",
    "End Function",
    "",
    "' Palette i's color for a build color; other colors are unchanged.",
    "Public Function PaletteOf(ByVal c As Long, ByVal i As Long) As Long",
    "    Select Case i",
]
lines += [f"        Case {i}: PaletteOf = {proc_name(p)}(c)" for i, p in enumerate(PALETTES, 1) if i > 1]
lines += [
    "        Case Else: PaletteOf = c",
    "    End Select",
    "End Function",
]
for p in PALETTES[1:]:
    fn = proc_name(p)
    lines += ["", f"' {p['name']}", f"Private Function {fn}(ByVal c As Long) As Long", "    Select Case c"]
    for key in ui:
        lb, target = build[key], p["c"][key]
        if target == lb or (p["base"] == 0 and target == P[key][0]):
            continue
        lv, tv = to_long(lb), to_long(target)
        lines.append(f"        Case {lv}: {fn} = {tv}{' ' * max(1, 20 - len(str(lv)) - len(str(tv)))}' {key}")
    lines += [f"        Case Else: {fn} = c", "    End Select", "End Function"]
os.makedirs(os.path.join(ROOT, "modules"), exist_ok=True)
with open(os.path.join(ROOT, "modules", "modPalette.bas"), "w", newline="\r\n") as f:
    f.write("\n".join(lines) + "\n")
print("wrote modPalette.bas")

# Text contrast (WCAG 2.1). Normal text needs 4.5:1; High Contrast aims for 7:1.
def lum(rgb):
    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (ch(v) for v in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b

def ratio(a, b):
    la, lb = sorted((lum(a), lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)

PAIRS = [("CLR_FG", "CLR_BG"), ("CLR_FG", "CLR_SURFACE"), ("CLR_MUTED_FG", "CLR_SURFACE"),
         ("CLR_ATTR_FG", "CLR_ROW1"), ("CLR_LINK", "CLR_SURFACE"), ("CLR_ERROR", "CLR_SURFACE"),
         ("CLR_BTN_FG", "CLR_BTN_BG"), ("CLR_OUT_FG", "CLR_OUT_BG"), ("CLR_NAV_FG", "CLR_SIDEBAR"),
         ("CLR_NAV_ACTIVE_FG", "CLR_NAV_ACTIVE_BG"), ("CLR_SUCCESS_FG", "CLR_SUCCESS_BG"),
         ("CLR_WARNING_FG", "CLR_WARNING_BG"), ("CLR_INFO_FG", "CLR_INFO_BG"),
         ("CLR_SEVA_FG", "CLR_SEVA_BG"), ("CLR_SEVB_FG", "CLR_SEVB_BG")]
for p in PALETTES:
    need = 7.0 if p["key"] == "contrast" else 4.5
    low = [f"{fg}/{bg} {ratio(p['c'][fg], p['c'][bg]):.1f}" for fg, bg in PAIRS
           if ratio(p["c"][fg], p["c"][bg]) < need]
    print(f"contrast {p['name']}: " + ("ok" if not low else "below %.1f: " % need + ", ".join(low)))
