#!/usr/bin/env python3
"""Minimal WMF renderer for formula preview images (Equations).
Outputs PNG via PIL. Handles: text (EXTTEXT), lines, polygons, fonts (Times/Symbol)."""
import struct, sys, os, math
from PIL import Image, ImageDraw, ImageFont

FONT_DIR = '/home/user/learning/fonts'
FONTS = {}
def font_for(face, weight, italic, size):
    key = (face.lower(), weight >= 500, bool(italic), size)
    if key in FONTS: return FONTS[key]
    if face.lower().startswith('symbol') or face.lower() == 'mt extra':
        path = FONT_DIR + '/StandardSymbolsPS-fixed.ttf'
    else:
        base = 'NimbusRoman'
        if weight >= 500:
            style = 'BoldItalic' if italic else 'Bold'
        else:
            style = 'Italic' if italic else 'Regular'
        path = f'{FONT_DIR}/{base}-{style}.ttf'
    FONTS[key] = ImageFont.truetype(path, size)
    return FONTS[key]

def parse_wmf(path):
    data = open(path, 'rb').read()
    off = 0
    if len(data) > 22 and struct.unpack_from('<I', data, 0)[0] == 0x9AC6CDD7:
        ph = struct.unpack_from('<I H h h h h H', data, 0)
        off = 22
    hdr = struct.unpack_from('<HHH I H I H', data, off)
    pos = off + hdr[1]*2
    end = len(data)
    recs = []
    while pos + 6 <= end:
        size, func = struct.unpack_from('<IH', data, pos)
        if size == 0: break
        recs.append((func, data[pos+6:pos+size*2]))
        pos += size*2
    return recs

def render(path, out_path, scale=3.0, font_shrink=None):
    recs = parse_wmf(path)
    pos_shrink = 1.0
    if font_shrink is None:
        n_text = ext_zero = mv_count = 0
        for func, rec in recs:
            if func == 0x0A32:
                n_text += 1
                if len(rec) >= 4 and rec[:4] == b'\x00\x00\x00\x00':
                    ext_zero += 1
            elif func == 0x0214:
                mv_count += 1
        if n_text and ext_zero == n_text and mv_count >= n_text:
            font_shrink = pos_shrink = 1.0/3.0
        else:
            font_shrink = 1.0
    # --- collect primitives ---
    prims = []  # ('text', x, y, s, face, weight, italic, lfheight, charset) | ('line', x1,y1,x2,y2) | ('polyline', pts) | ('polygon', pts)
    fonts = {}  # index -> dict
    obj_table = {}  # handle -> dict (handles are 2-slot, reused)
    used_handles = set()
    def alloc_handle():
        h = 0
        while h in used_handles:
            h += 1
        used_handles.add(h)
        return h
    cur_font_idx = None
    text_color = (0, 0, 0)
    cur_pos = None
    text_align = 0
    for func, rec in recs:
        if func == 0x02FB:  # CREATEFONTINDIRECT (16-bit fields)
            if len(rec) < 18: continue
            lfHeight, lfWidth, lfEsc, lfOrient, lfWeight = struct.unpack_from('<hhhhh', rec, 0)
            italic = rec[10] if len(rec) > 10 else 0
            charset = rec[13] if len(rec) > 13 else 0
            pitchfam = rec[17] if len(rec) > 17 else 0
            name = rec[18:].split(b'\x00')[0].decode('latin1', 'replace').strip()
            obj_table[len(obj_table)] = {
                'face': name, 'weight': abs(lfWeight) if lfWeight else 400,
                'italic': bool(italic), 'height': abs(lfHeight) or 100,
                'charset': charset,
            }
        elif func == 0x0626:  # ESCAPE (e.g., MathType identifiers) - skip
            continue
        elif func == 0x012D:  # SELECTOBJECT
            idx = struct.unpack_from('<H', rec, 0)[0]
            if idx in obj_table:
                cur_font_idx = idx
        elif func == 0x01F0:  # DELETEOBJECT
            if len(rec) >= 2:
                h = struct.unpack_from('<H', rec, 0)[0]
                obj_table.pop(h, None)
                used_handles.discard(h)
        elif func == 0x0209:  # SETTEXTCOLOR
            b, g, r = rec[0], rec[1], rec[2]
            text_color = (r, g, b)
        elif func == 0x0214:  # MOVETO  (Y, X order per spec)
            y, x = struct.unpack_from('<hh', rec, 0)
            cur_pos = (x * pos_shrink, y * pos_shrink)
        elif func == 0x0213:  # LINETO
            y, x = struct.unpack_from('<hh', rec, 0)
            x *= pos_shrink; y *= pos_shrink
            if cur_pos:
                prims.append(('line', cur_pos[0], cur_pos[1], x, y))
            cur_pos = (x, y)
        elif func == 0x0325:  # POLYLINE
            n = struct.unpack_from('<H', rec, 0)[0]
            pts = []
            p = 2
            for _ in range(n):
                y, x = struct.unpack_from('<hh', rec, p); p += 4
                pts.append((x*pos_shrink, y*pos_shrink))
            for i in range(len(pts)-1):
                prims.append(('line', pts[i][0], pts[i][1], pts[i+1][0], pts[i+1][1]))
        elif func == 0x0324:  # POLYGON
            n = struct.unpack_from('<H', rec, 0)[0]
            pts = []
            p = 2
            for _ in range(n):
                y, x = struct.unpack_from('<hh', rec, p); p += 4
                pts.append((x*pos_shrink, y*pos_shrink))
            pts.append(pts[0])
            for i in range(len(pts)-1):
                prims.append(('line', pts[i][0], pts[i][1], pts[i+1][0], pts[i+1][1]))
        elif func == 0x0538:  # POLYPOLYGON
            npoly = struct.unpack_from('<H', rec, 0)[0]
            p = 2
            counts = []
            for _ in range(npoly):
                counts.append(struct.unpack_from('<H', rec, p)[0]); p += 2
            if len(rec) % 2 == 1: p += 1
            for c in counts:
                pts = []
                for _ in range(c):
                    y, x = struct.unpack_from('<hh', rec, p); p += 4
                    pts.append((x*pos_shrink, y*pos_shrink))
                if len(pts) > 1:
                    pts.append(pts[0])
                    for i in range(len(pts)-1):
                        prims.append(('line', pts[i][0], pts[i][1], pts[i+1][0], pts[i+1][1]))
        elif func == 0x0A32:  # EXTTEXTOUT
            if len(rec) < 8: continue
            y, x, slen, fopts = struct.unpack_from('<hhHH', rec, 0)
            p = 8
            if fopts & 0x02 or fopts & 0x0004:  # ETO_CLIPPED=4? ETO_OPAQUE=2
                p += 8
            s = rec[p:p+slen]
            p += slen
            if slen % 2 == 1: p += 1
            font = obj_table.get(cur_font_idx, {'face':'Times New Roman','weight':400,'italic':False,'height':100,'charset':0})
            # decode per font face (MathType may mark Symbol runs with cs=1/0)
            cs = font['charset']
            face_l = font['face'].lower()
            if cs == 2 or face_l.startswith('symbol') or face_l == 'mt extra':
                # SYMBOL_CHARSET: byte -> U+F000+byte
                text = ''.join(chr(0xF000 + b) for b in s)
            elif cs == 134:
                text = s.decode('gbk', 'replace')
            else:
                text = s.decode('latin1')
            # position: if (0,0) and cur_pos set (MOVETO style), use cur_pos
            if x == 0 and y == 0 and cur_pos is not None:
                x, y = cur_pos
            else:
                x, y = x * pos_shrink, y * pos_shrink
            prims.append(('text', x, y, text, font))
        elif func == 0x012E:  # SETTEXTALIGN
            text_align = struct.unpack_from('<H', rec, 0)[0]
        # everything else: skip
    # --- compute bbox ---
    xs, ys = [], []
    for pr in prims:
        if pr[0] == 'text':
            _, x, y, text, font = pr
            size1 = max(int(font['height'] * font_shrink), 8)
            fnt1 = font_for(font['face'], font['weight'], font['italic'], size1)
            xs += [x, x + fnt1.getlength(text)]
            ys += [y - font['height'] * font_shrink, y + font['height'] * font_shrink * 0.35]
        else:
            _, x1, y1, x2, y2 = pr
            xs += [x1, x2]; ys += [y1, y2]
    if not xs:
        xs, ys = [0, 100], [0, 100]
    pad = 30
    minx, maxx, miny, maxy = min(xs), max(xs), min(ys), max(ys)
    span = max(maxx - minx, maxy - miny, 1e-6)
    scale = min(scale, 5000.0 / span)
    W = int((maxx - minx) * scale) + pad*2
    H = int((maxy - miny) * scale) + pad*2
    W = max(W, 60); H = max(H, 60)
    img = Image.new('RGB', (W, H), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    def tx(x): return (x - minx) * scale + pad
    def ty(y): return (y - miny) * scale + pad
    for pr in prims:
        if pr[0] == 'line':
            _, x1, y1, x2, y2 = pr
            draw.line([(tx(x1), ty(y1)), (tx(x2), ty(y2))], fill=(20,20,20), width=max(1, int(2.2*scale)))
        else:
            _, x, y, text, font = pr
            size = max(int(font['height'] * font_shrink * scale), 8)
            fnt = font_for(font['face'], font['weight'], font['italic'], size)
            # WMF y is baseline-ish for text; draw with anchor at baseline
            draw.text((tx(x), ty(y)), text, font=fnt, fill=text_color, anchor='ls')
    img.save(out_path)
    return out_path

if __name__ == '__main__':
    src, dst = sys.argv[1], sys.argv[2]
    render(src, dst, scale=float(sys.argv[3]) if len(sys.argv) > 3 else 3.0,
           font_shrink=float(sys.argv[4]) if len(sys.argv) > 4 else 1.0)
