import math
from PIL import Image, ImageDraw, ImageFont

# Excel geometry: col width 16 chars -> px ; row 20px ; chart cm->px @96dpi
COLPX = int(16 * 7) + 5      # 117
ROWPX = 20
def cm2px(cm): return int(cm / 2.54 * 96)

charts = [  # (name, anchor_col_letter, anchor_row, w_cm, h_cm)
    ("1 تراكمي/خطة", "C", 22, 16, 9),
    ("2 يومي P/T",    "O", 22, 16, 9),
    ("3 العائلات",    "C", 40, 16, 9),
    ("4 تصدير/محلي",  "O", 40, 16, 9),
    ("5 المواد",      "C", 58, 16, 9),
    ("6 الجودة",      "O", 58, 16, 9),
    ("7 العملاء",     "C", 76, 16, 10),
    ("8 الماكينات",   "O", 76, 16, 10),
]
NCOLS = 24  # A..X
rects = []
for name, colL, row, wcm, hcm in charts:
    c0 = ord(colL) - 65
    wpx, hpx = cm2px(wcm), cm2px(hcm)
    span = math.ceil(wpx / COLPX)
    rows_span = math.ceil(hpx / ROWPX)
    rects.append((name, c0, c0 + span, row, row + rows_span))

# overlap assertions (cols x rows grid)
bad = 0
for i in range(len(rects)):
    for j in range(i + 1, len(rects)):
        a, b = rects[i], rects[j]
        if a[1] < b[2] and b[1] < a[2] and a[3] < b[4] and b[3] < a[4]:
            print("OVERLAP:", a[0], b[0]); bad += 1
for name, c0, c1, r0, r1 in rects:
    print(f"{name}: cols {chr(65+c0)}..{chr(65+min(c1,25)-1)} rows {r0}..{r1}")
    if c1 > NCOLS: print("  !! extends beyond col X"); bad += 1
print("OVERLAPS:", bad)

# schematic (RTL: col A on the right)
W, H = 1400, 1000
img = Image.new("RGB", (W, H), "white")
dr = ImageDraw.Draw(img)
sx = W / (NCOLS + 1)
sy = H / 96.0
for c in range(NCOLS + 1):
    x = W - c * sx
    dr.line([(x, 0), (x, H)], fill="#eeeeee")
for r in range(0, 96, 2):
    dr.line([(0, r * sy), (W, r * sy)], fill="#eeeeee")
colors = ["#1976D2", "#0F9D58", "#7C3AED", "#DC2626", "#EA580C", "#CA8A04", "#0D9488", "#64748B"]
for k, (name, c0, c1, r0, r1) in enumerate(rects):
    x_right = W - c0 * sx
    x_left = W - c1 * sx
    dr.rectangle([x_left, r0 * sy, x_right, r1 * sy], outline=colors[k], width=3)
    dr.text((x_left + 8, r0 * sy + 6), name, fill=colors[k])
# KPI + selectors zone
dr.rectangle([W - 22 * sx, 4 * sy, W - 2 * sx, 20 * sy], outline="#0D47A1", width=2)
dr.text((W - 21 * sx, 4.5 * sy), "selectors + KPI cards + logo", fill="#0D47A1")
img.save("/tmp/layout_mock.png")
print("mock saved")
