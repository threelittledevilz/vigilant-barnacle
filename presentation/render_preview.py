#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
render_preview.py — пиксельно-точный предпросмотр deck'а по scene.json.
Использует те же координаты и те же шрифты (Montserrat / Inter), что и PPTX.
Не нужно для самой презентации — только для быстрого ревью и QA.
"""
import json
import os
import sys

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
FONTDIR = os.path.expanduser("~/.fontsrc/node_modules/@expo-google-fonts")
DEJAVU = "/usr/share/fonts/truetype/dejavu"
SCALE = 144.0  # px per inch  -> 1920x1080

_FCACHE = {}
_CMAPS = {}


def _load(fname, weight):
    path = os.path.join(FONTDIR, weight, "%s_%s.ttf" % (fname, weight))
    if not os.path.exists(path):
        path = os.path.join(DEJAVU, "DejaVuSans%s.ttf" % ("-Bold" if weight[0] == "7" else ""))
    return path


FAMILIES = {
    "Montserrat": {"bold": "700Bold", "reg": "400Regular"},
    "Inter": {"bold": "700Bold", "reg": "400Regular"},
}


def _cmap(path):
    if path not in _CMAPS:
        try:
            from fontTools.ttLib import TTFont
            _CMAPS[path] = set(TTFont(path, lazy=True).getBestCmap() or {})
        except Exception:
            _CMAPS[path] = None
    return _CMAPS[path]


def _fallbacks(family, bold):
    wkey = "bold" if bold else "reg"
    out = [_load(family, FAMILIES[family][wkey])]
    for fam2, wk in FAMILIES.items():
        p = _load(fam2, wk[wkey] if wk is FAMILIES[family] else wk[wkey])
        if p not in out:
            out.append(p)
    out.append(os.path.join(DEJAVU, "DejaVuSans%s.ttf" % ("-Bold" if bold else "")))
    return out


def get_font(family, bold, px):
    key = (family, bold, int(px))
    if key not in _FCACHE:
        p = _fallbacks(family, bold)[0]
        _FCACHE[key] = {"primary": ImageFont.truetype(p, int(px)),
                        "primary_path": p,
                        "pool": [ImageFont.truetype(q, int(px)) for q in _fallbacks(family, bold)],
                        "pool_paths": _fallbacks(family, bold)}
    return _FCACHE[key]


def font_for_char(finfo, ch):
    path = finfo["primary_path"]
    cmap = _cmap(path)
    if cmap is not None and ord(ch) in cmap:
        return finfo["primary"]
    for f, p in zip(finfo["pool"], finfo["pool_paths"]):
        c2 = _cmap(p)
        if c2 is None or ord(ch) in c2:
            return f
    return finfo["primary"]


def hex2rgb(h):
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def lerp_color(c1, c2, t):
    return tuple(int(a + (b - a) * t) for a, b in zip(c1, c2))


def grad_color(stops, pos):
    if pos <= stops[0][0]:
        return hex2rgb(stops[0][1])
    for (p1, h1), (p2, h2) in zip(stops, stops[1:]):
        if pos <= p2:
            t = (pos - p1) / (p2 - p1) if p2 > p1 else 0
            return lerp_color(hex2rgb(h1), hex2rgb(h2), t)
    return hex2rgb(stops[-1][1])


def px(v):
    return v * SCALE


def put(img, layer):
    img.alpha_composite(layer)


def draw_shape(img, s):
    x0, y0 = px(s["x"]), px(s["y"])
    x1, y1 = px(s["x"] + s["w"]), px(s["y"] + s["h"])
    lay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(lay)
    kind = s["kind"]
    fill = s.get("fill")
    line = s.get("line")
    rad = px(s["radius"] * min(s["w"], s["h"])) if s.get("radius") else 0

    if s.get("grad") is not None:
        stops = sorted(s["grad"])
        gh = max(1, int(y1 - y0))
        gl = Image.new("RGBA", (1, gh))
        for i in range(gh):
            c = grad_color(stops, 100.0 * i / (gh - 1))
            gl.putpixel((0, i), c + (255,))
        gl = gl.resize((max(1, int(x1 - x0)), gh))
        lay.paste(gl, (int(x0), int(y0)))
        d = ImageDraw.Draw(lay)
    else:
        if fill:
            col = tuple(fill[0]) + (int(fill[1] * 2.55),)
            if kind == "ellipse":
                d.ellipse([x0, y0, x1, y1], fill=col)
            elif kind == "diamond":
                d.polygon([( (x0+x1)/2, y0), (x1, (y0+y1)/2), ((x0+x1)/2, y1), (x0, (y0+y1)/2)], fill=col)
            elif kind == "roundrect":
                d.rounded_rectangle([x0, y0, x1, y1], radius=rad, fill=col)
            else:
                d.rectangle([x0, y0, x1, y1], fill=col)
    if line:
        col = tuple(line[0]) + (int(line[1] * 2.55),)
        wpx = max(1, int(line[2] * SCALE / 72.0))
        if kind == "ellipse":
            d.ellipse([x0, y0, x1, y1], outline=col, width=wpx)
        elif kind == "roundrect":
            d.rounded_rectangle([x0, y0, x1, y1], radius=rad, outline=col, width=wpx)
        else:
            d.rectangle([x0, y0, x1, y1], outline=col, width=wpx)
    put(img, lay)


def measure(t, finfo, spc_px):
    w = 0.0
    for ch in t:
        f = font_for_char(finfo, ch)
        w += f.getlength(ch) + spc_px
    return w


def draw_text(img, t):
    bx, by, bw, bh = px(t["x"]), px(t["y"]), px(t["w"]), px(t["h"])
    # Разбивка: строки с переносом
    lines = []  # (list of (char, finfo, spc_px, size_px))
    for p in t["paras"]:
        if p.get("before"):
            lines.append(("space", p["before"] * 2))
        tokens = []
        for r in p["runs"]:
            finfo = get_font(r["f"], r.get("b"), r["px"] * 2)
            spc = r.get("spc", 0) / 100.0 * 2
            for ch in r["t"]:
                tokens.append((ch, finfo, spc, r["px"] * 2, r))
        # grouping into words (keep spaces as own tokens)
        words, cur = [], []
        for tk in tokens:
            if tk[0] == " ":
                if cur:
                    words.append(cur)
                cur = [tk]
            else:
                cur.append(tk)
        if cur:
            words.append(cur)
        # wrap
        cur_line, cur_w = [], 0.0
        for w_ in words:
            ww = measure("".join(c for c, *_ in w_), None, 0) if False else sum(
                font_for_char(f, c).getlength(c) + sp for c, f, sp, *_ in w_)
            add = ww
            if cur_line and cur_w + add > bw:
                lines.append(("line", cur_line))
                cur_line, cur_w = list(w_), ww
            else:
                cur_line.extend(w_)
                cur_w += add
        if cur_line:
            lines.append(("line", cur_line))
        ls = p.get("ls") or 1.2
        lines.append(("line_h", ls, bw, p.get("align", "left"), p.get("after")))
    # heights
    line_metrics = []
    total = 0.0
    for item in lines:
        if item[0] == "space":
            line_metrics.append((item, item[1]))
            total += item[1]
        elif item[0] == "line":
            chars = item[1]
            asc = max(font_for_char(f, c).getmetrics()[0] for c, f, *_ in chars)
            desc = max(font_for_char(f, c).getmetrics()[1] for c, f, *_ in chars)
            ls = None
            for later in lines[lines.index(item) + 1:]:
                if later[0] == "line_h":
                    ls = later[2] if False else later[1]
                    break
            ls = ls or 1.2
            h = (asc + desc) * ls
            line_metrics.append((item, h, asc, desc))
            total += h
        elif item[0] == "line_h":
            if item[4]:
                total += item[4] * 2
    y0 = by
    if t.get("anchor") == "middle":
        y0 = by + (bh - total) / 2.0
    d = ImageDraw.Draw(img)
    for lm in line_metrics:
        item = lm[0]
        if item[0] == "space":
            continue
        if item[0] == "line":
            chars = item[1]
            h, asc, desc = lm[1], lm[2], lm[3]
            # align: find next line_h
            idx = lines.index(item)
            align = "left"
            for later in lines[idx + 1:]:
                if later[0] == "line_h":
                    align = later[3]
                    break
            lw = sum(font_for_char(f, c).getlength(c) + sp for c, f, sp, *_ in chars)
            if align == "center":
                x = bx + max(0.0, (bw - lw) / 2.0)
            elif align == "right":
                x = bx + max(0.0, bw - lw)
            else:
                x = bx
            baseline = y0 + asc + (h - (asc + desc)) / 2.0
            lay = Image.new("RGBA", img.size, (0, 0, 0, 0))
            dl = ImageDraw.Draw(lay)
            for c, f, sp, size, r in chars:
                fch = font_for_char(f, c)
                col = (r["c"][0], r["c"][1], r["c"][2], int(r["c"][3] * 2.55))
                dl.text((x, baseline - fch.getmetrics()[0]), c, font=fch, fill=col)
                x += fch.getlength(c) + sp
            put(img, lay)
            y0 += h
        elif item[0] == "line_h":
            if item[4]:
                y0 += item[4] * 2


def render_scene(scene, outdir):
    os.makedirs(outdir, exist_ok=True)
    imgs = []
    for s in scene["slides"]:
        img = Image.new("RGBA", (int(px(scene["w"])), int(px(scene["h"]))), (10, 15, 29, 255))
        for sh in s["shapes"]:
            draw_shape(img, sh)
        for tx in s["texts"]:
            draw_text(img, tx)
        for pc in s.get("pics", []):
            p = os.path.join(HERE, pc["src"])
            im = Image.open(p).convert("RGBA")
            im = im.resize((int(px(pc["w"])), int(px(pc["h"]))), Image.LANCZOS)
            img.paste(im, (int(px(pc["x"])), int(px(pc["y"]))), im)
        p = os.path.join(outdir, "slide-%02d.png" % s["idx"])
        img.convert("RGB").save(p, "PNG")
        imgs.append(img.convert("RGB"))
        print("preview ->", p)
    # contact sheet 4x2
    tw, th = 480, 270
    sheet = Image.new("RGB", (tw * 4 + 30, th * 2 + 30), (24, 26, 34))
    for i, im in enumerate(imgs):
        im2 = im.resize((tw, th), Image.LANCZOS)
        x = 10 + (i % 4) * (tw + 5)
        y = 10 + (i // 4) * (th + 5)
        sheet.paste(im2, (x, y))
    sp = os.path.join(outdir, "contact-sheet.png")
    sheet.save(sp, "PNG")
    print("preview ->", sp)


if __name__ == "__main__":
    scene_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "scene.json")
    outdir = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "previews")
    with open(scene_path, encoding="utf-8") as fh:
        render_scene(json.load(fh), outdir)
