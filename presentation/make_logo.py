#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_logo.py — векторная реконструкция знака «Кьево» (PNG, белый, прозрачный фон).

Геометрия по оригиналу 720×720:
  * окружность: центр (360,360), R≈238, штрих ≈10.5
  * двойная диагональ под ~52° (круче 45°): две параллельные линии с шагом
    ~52px; длинный «хвост» за пределами круга на право-вверх, короткий
    выступ на лево-вниз; торец «хвоста» закрыт короткой перемычкой.
Рендер 2× (1440px) с даунскейлом для чистого антинейсинга.
"""
import math
from PIL import Image, ImageDraw

S = 1440          # render size (2x)
O = 2.0           # upscale factor vs 720
WHITE = (255, 255, 255, 255)


def build():
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    sw = int(10.5 * O)

    cx = cy = 360 * O
    r = 238 * O
    d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=WHITE, width=sw)

    # Диагональ: угол 52° от горизонтали, направление вправо-вверх
    ang = math.radians(52)
    u = (math.cos(ang), -math.sin(ang))     # вдоль штриха
    p = (math.sin(ang), math.cos(ang))      # перпендикуляр (вправо-вниз)

    # центр линии проходит чуть левее-выше центра круга
    px0, py0 = 332 * O, 332 * O
    gap = 26 * O          # смещение каждой линии от оси
    t1, t2 = -262 * O, 302 * O   # несимметричные торцы (короткий/длинный)

    def L(off, t_start, t_end):
        x1 = px0 + t_start * u[0] + off * p[0]
        y1 = py0 + t_start * u[1] + off * p[1]
        x2 = px0 + t_end * u[0] + off * p[0]
        y2 = py0 + t_end * u[1] + off * p[1]
        d.line([x1, y1, x2, y2], fill=WHITE, width=sw)
        return (x1, y1), (x2, y2)

    a_top = L(-gap, t1, t2)[1]
    b_top = L(+gap, t1, t2)[1]
    # перемычка на верхнем (длинном) торце — замыкает «двойной штрих»
    d.line([a_top, b_top], fill=WHITE, width=sw)

    # обрезка «хвостов» (круглые края за пределами холста не нужны)
    im = im.resize((720, 720), Image.LANCZOS)
    return im


if __name__ == "__main__":
    im = build()
    im.save("assets/logo_white.png")
    # тёмная версия для светлых носителей (отчёты, печать)
    dark = im.copy()
    dark = dark.point(lambda v: v)  # альфа без изменений
    dk = Image.new("RGBA", (720, 720), (0, 0, 0, 0))
    a = dark.getchannel("A")
    dk.paste((26, 26, 26, 255), (0, 0), a)
    dk.save("assets/logo_dark.png")
    print("logo ok: assets/logo_white.png, assets/logo_dark.png")
