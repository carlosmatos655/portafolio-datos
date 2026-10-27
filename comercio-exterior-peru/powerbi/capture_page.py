"""Captura la página visible de Power BI Desktop y la recorta al lienzo del reporte (1280×720 a escala).

Uso: python powerbi/capture_page.py <salida.png> [--width 1920]
Detecta el lienzo por la barra lateral oscura (#1E1B18) y el fondo hueso (#F4EFE7) del diseño.
"""
import ctypes
import sys

from PIL import ImageGrab

import time

user32 = ctypes.windll.user32
user32.SetProcessDPIAware()
out = sys.argv[1]
OUT_W = int(sys.argv[sys.argv.index("--width") + 1]) if "--width" in sys.argv else 1920

# Power BI al frente (la ventana desde la que se lanza el script lo taparía).
hwnd = user32.FindWindowW(None, "ComercioExteriorPeru")
if hwnd:
    user32.ShowWindow(hwnd, 3)
    user32.SetForegroundWindow(hwnd)
# Las ventanas visibles titulas "Claude" (la app desde la que se lanza) se minimizan un momento.
claude = []
EnumProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)


def _collect(h, _):
    buf = ctypes.create_unicode_buffer(64)
    user32.GetWindowTextW(h, buf, 64)
    if buf.value == "Claude" and user32.IsWindowVisible(h) and not user32.IsIconic(h):
        claude.append(h)
    return True


user32.EnumWindows(EnumProc(_collect), 0)
for h in claude:
    user32.ShowWindow(h, 6)
time.sleep(1.5)
img = ImageGrab.grab().convert("RGB")
for h in claude:
    user32.ShowWindow(h, 9)
W, H = img.size
px = img.load()


def near(c, ref, tol=14):
    return all(abs(a - b) <= tol for a, b in zip(c, ref))


RAIL = (30, 27, 24)
# La barra lateral del diseño es tinta cálida (R > B); los grises de Power BI son neutros (R = B).
# Su primera columna es el borde izquierdo del lienzo; su ancho es 76 de 1280 px.
warm = lambda c: near(c, RAIL, 6) and c[0] - c[2] >= 4
BONE = (244, 239, 231)
bone = lambda c: near(c, BONE, 8)
# A media altura, el primer píxel hueso es el borde derecho de la barra lateral (x = 76 en el diseño).
y = H // 2
dark = lambda c: max(c) < 70 and c[0] - c[2] >= 4  # tinta cálida de la barra, no el gris neutro de Power BI
# (el margen del lienzo también es hueso: se exige que a la izquierda haya barra oscura)
xb = next(x for x in range(60, W) if bone(px[x, y]) and bone(px[x + 6, y]) and all(dark(px[x - d, y]) for d in (8, 20, 40, 60)))
col = xb + 12
# Arriba y abajo: la columna junto a la barra es hueso de punta a punta del lienzo (encima está la
# cinta de Power BI y debajo la barra de pestañas, ambas oscuras).
top = next(yy for yy in range(H) if bone(px[col, yy]) and bone(px[col, yy + 6]))
bottom = next(yy for yy in range(top + 50, H) if not bone(px[col, yy]) and not bone(px[col, yy + 6]))
height = bottom - top
width = round(height * 1280 / 720)
left = xb - round(76 * width / 1280)
right = left + width - 1
crop = img.crop((left, top, right + 1, top + height))
crop = crop.resize((OUT_W, round(OUT_W * 720 / 1280)))
crop.save(out, optimize=True)
print(out, (left, top, right, top + height), crop.size)
