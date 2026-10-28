"""Captura la página visible de Power BI Desktop y la recorta al lienzo del reporte (1280×720 a escala).

Uso: python powerbi/capture_page.py <salida.png> [--width 1920]
Detecta el lienzo por la barra lateral del diseño (#0C0B14, un negro azulado): los grises oscuros
de la interfaz de Power BI son neutros (R = G = B), la barra lateral tiene más azul que rojo.
"""
import ctypes
import sys
import time

from PIL import ImageGrab

user32 = ctypes.windll.user32
user32.SetProcessDPIAware()
out = sys.argv[1]
OUT_W = int(sys.argv[sys.argv.index("--width") + 1]) if "--width" in sys.argv else 1920

# Power BI al frente; las ventanas "Claude" (desde donde se lanza) se minimizan durante la captura.
hwnd = user32.FindWindowW(None, "ComprasPublicasPeru")
if hwnd:
    user32.ShowWindow(hwnd, 3)
    user32.SetForegroundWindow(hwnd)
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

side = lambda c: max(c) < 40 and c[2] - c[0] >= 2 and c[2] >= c[1]   # barra lateral #0C0B14 (más azul que rojo)


def longest_run(x):
    best, cur, start, best_start = 0, 0, 0, 0
    for yy in range(H):
        if side(px[x, yy]):
            if cur == 0:
                start = yy
            cur += 1
            if cur > best:
                best, best_start = cur, start
        else:
            cur = 0
    return best, best_start


# La columna con la racha vertical más larga de "barra lateral" recorre el lienzo de arriba a abajo.
runs = [(longest_run(x), x) for x in range(40, W // 3, 3)]
(height, top), col = max(runs)
y = top + height // 2
xl = col
while xl > 0 and side(px[xl - 1, y]):
    xl -= 1
# El fondo del área principal también es azulado y la racha se corta en el borde de la barra lateral
# (x = 218 del diseño): si a la izquierda queda barra lateral, se retrocede su ancho.
if xl > round(218 * height / 720) and sum(side(px[xl - round(209 * height / 720), top + k * height // 20]) for k in range(1, 20)) > 12:
    xl -= round(218 * height / 720)
width = round(height * 1280 / 720)
crop = img.crop((xl, top, xl + width, top + height)).resize((OUT_W, round(OUT_W * 720 / 1280)))
crop.save(out, optimize=True)
print(out, (xl, top, xl + width, top + height), crop.size)
