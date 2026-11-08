"""PDF 转 Word 核心：按版式重建，而不是用文本框硬拼。

思路：读出每个字的位置、字号、颜色，以及色块、分隔线、表格和图片，
再把它们还原成 Word 里原生、可编辑的元素：
  文字行  → 段落（精确行距 + 段前间距，让位置和原文一致）
  左右两端的文字（如公司名和日期）→ 同一段落里的制表位
  标题左侧的色条 → 段落左边框
  横向分隔线 → 段落下边框
  带底色的说明框 → 段落底纹 + 边框
  有框线的表格 → Word 表格
  图片 → 按原位置浮动的图片
"""
import io
import re
import sys
from statistics import median

import pdfplumber
import pypdfium2
import pypdfium2.raw as pdfium_c
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Emu, Pt, RGBColor

BULLETS = '•·●○■□◆◇▪▫–-*'
HANG_BULLETS = '•●○■□◆◇▪▫'
BOLD_HINTS = ('bold', 'medium', 'semibold', 'demi', 'heavy', 'black', 'w5', 'w6', 'w7', 'w8', 'w9')
DESC = 0.18          # 文字基线以下预留的高度（相对字号），用于换算行的位置


class ScannedPDF(Exception):
    """整份 PDF 几乎没有文字，判定为扫描件。"""


def default_font():
    if sys.platform == 'darwin':
        return 'PingFang SC'
    if sys.platform.startswith('win'):
        return 'Microsoft YaHei'
    return 'Noto Sans CJK SC'


# ---------- 读取 PDF ----------

def _color(c):
    """pdfplumber 的颜色 → RGB 十六进制；读不懂就返回 None。"""
    if c is None:
        return None
    if isinstance(c, (int, float)):
        c = (c,)
    c = tuple(c)
    if len(c) == 1:
        v = c[0]
        return '%02X%02X%02X' % ((round(v * 255),) * 3)
    if len(c) == 3:
        return ''.join('%02X' % round(v * 255) for v in c)
    if len(c) == 4:  # CMYK
        cc, m, y, k = c
        return ''.join('%02X' % round(255 * (1 - v) * (1 - k)) for v in (cc, m, y))
    return None


def _is_bold(fontname):
    name = fontname.lower().split('+')[-1]
    return any(h in name for h in BOLD_HINTS)


def _baseline(ch, page_h):
    return page_h - ch['matrix'][5]


def _inside(obj, box, pad=1.0):
    return (obj['x0'] >= box[0] - pad and obj['x1'] <= box[2] + pad
            and obj['top'] >= box[1] - pad and obj['bottom'] <= box[3] + pad)


def _make_lines(chars, page_h):
    """把字按基线分组成行，行内按较大的空白切成「段」（左边的标题、右边的日期等）。"""
    chars = [c for c in chars if abs(c['matrix'][1]) < 1e-3 and abs(c['matrix'][2]) < 1e-3]
    chars.sort(key=lambda c: (_baseline(c, page_h), c['x0']))
    rows = []
    for ch in chars:
        b = _baseline(ch, page_h)
        tol = max(1.5, 0.25 * ch['size'])
        if rows and abs(rows[-1]['b'] - b) <= tol:
            rows[-1]['chars'].append(ch)
        else:
            rows.append({'b': b, 'chars': [ch]})
    lines = []
    for r in rows:
        cs = sorted(r['chars'], key=lambda c: c['x0'])
        if not ''.join(c['text'] for c in cs).strip():
            continue
        size = median(c['size'] for c in cs if c['text'].strip())
        segs, cur, last_x1 = [], [], None
        for c in cs:
            if c['text'].strip() and last_x1 is not None and c['x0'] - last_x1 > max(2.5 * size, 18):
                segs.append(cur)
                cur = []
            cur.append(c)
            if c['text'].strip():
                last_x1 = c['x1']
        segs.append(cur)
        segments = []
        for s in segs:
            while s and not s[0]['text'].strip():
                s = s[1:]
            while s and not s[-1]['text'].strip():
                s = s[:-1]
            if s:
                segments.append({'x0': s[0]['x0'], 'x1': s[-1]['x1'], 'runs': _runs(s)})
        body = [c for c in cs if c['text'].strip()]
        lines.append({
            'kind': 'line', 'segments': segments,
            'x0': segments[0]['x0'], 'x1': segments[-1]['x1'],
            'top': min(c['top'] for c in body), 'bottom': max(c['bottom'] for c in body),
            'b': median(_baseline(c, page_h) for c in body),
            'size': max(c['size'] for c in body),
            'color': _color(body[0].get('non_stroking_color')),
            'text': ''.join(c['text'] for c in cs).strip(),
            'chars': body,
        })
    return lines


def _runs(chars):
    runs = []
    for c in chars:
        key = (c['fontname'], round(c['size'] * 2) / 2, _color(c.get('non_stroking_color')))
        if runs and runs[-1]['key'] == key:
            runs[-1]['text'] += c['text']
        else:
            runs.append({'key': key, 'text': c['text'], 'size': key[1],
                         'bold': _is_bold(c['fontname']), 'color': key[2]})
    for r in runs:
        r['text'] = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', r['text'])
    return runs


def _shapes(page, tables):
    """把矩形、曲线、线条分成：色条、带底色的框、横向分隔线。"""
    in_table = lambda o: any(_inside(o, t, 2) for t in tables)
    bars, boxes, rules = [], [], []
    for o in list(page.rects) + list(page.curves):
        if in_table(o):
            continue
        w, h = o['x1'] - o['x0'], o['bottom'] - o['top']
        fill = _color(o.get('non_stroking_color')) if o.get('fill') else None
        stroke = _color(o.get('stroking_color')) if o.get('stroke') else None
        if fill and w <= 6 and 4 <= h <= 40:
            bars.append({**o, 'color': fill})
        elif h <= 2.5 and w >= 40:
            rules.append({'x0': o['x0'], 'x1': o['x1'], 'y': (o['top'] + o['bottom']) / 2,
                          'w': max(h, 0.5), 'color': fill or stroke or '000000'})
        elif (fill and fill != 'FFFFFF' or stroke) and w >= 30 and h >= 8:
            boxes.append({'x0': o['x0'], 'x1': o['x1'], 'top': o['top'], 'bottom': o['bottom'],
                          'fill': fill if fill and fill != 'FFFFFF' else None,
                          'stroke': stroke, 'lw': o.get('linewidth') or 0.5})
    for ln in page.lines:
        if in_table(ln):
            continue
        if abs(ln['top'] - ln['bottom']) <= 1 and ln['x1'] - ln['x0'] >= 40:
            rules.append({'x0': ln['x0'], 'x1': ln['x1'], 'y': ln['top'],
                          'w': ln.get('linewidth') or 0.5,
                          'color': _color(ln.get('stroking_color')) or '000000'})
    # 框的边线不算分隔线
    rules = [r for r in rules if not any(b['x0'] - 2 <= r['x0'] and r['x1'] <= b['x1'] + 2
                                          and b['top'] - 2 <= r['y'] <= b['bottom'] + 2 for b in boxes)]
    return bars, boxes, rules


def _tables(page):
    out = []
    for t in page.find_tables():
        cells = [c for c in t.cells if c]
        xs = sorted({round(v, 1) for c in cells for v in (c[0], c[2])})
        ys = sorted({round(v, 1) for c in cells for v in (c[1], c[3])})
        if len(xs) < 2 or len(ys) < 3:
            continue
        rows = []
        for y0, y1 in zip(ys, ys[1:]):
            row = []
            for x0, x1 in zip(xs, xs[1:]):
                crop = page.crop((x0 + 0.5, y0 + 0.5, x1 - 0.5, y1 - 0.5))
                row.append(_make_lines(crop.chars, page.height))
            rows.append(row)
        edges = [e for e in page.edges if _inside(e, t.bbox, 2)]
        color = _color(edges[0].get('stroking_color')) if edges else '000000'
        lw = edges[0].get('linewidth', 0.5) if edges else 0.5
        out.append({'kind': 'table', 'bbox': t.bbox, 'top': t.bbox[1], 'bottom': t.bbox[3],
                    'xs': xs, 'ys': ys, 'rows': rows, 'color': color or '000000', 'lw': lw or 0.5})
    return out


def _images(pdf_path, index, page_h):
    out = []
    doc = pypdfium2.PdfDocument(pdf_path)
    try:
        page = doc[index]
        for obj in page.get_objects(filter=(pdfium_c.FPDF_PAGEOBJ_IMAGE,)):
            left, bottom, right, top = obj.get_bounds()
            if right - left < 4 or top - bottom < 4:
                continue
            try:
                pil = obj.get_bitmap(render=True).to_pil()
            except Exception:
                continue
            buf = io.BytesIO()
            pil.convert('RGBA' if pil.mode in ('RGBA', 'LA', 'P') else 'RGB').save(buf, 'PNG')
            out.append({'x0': left, 'x1': right, 'top': page_h - top, 'bottom': page_h - bottom,
                        'png': buf.getvalue()})
    finally:
        doc.close()
    return out


def _gutter(lines, page_w):
    """找页面中部一条几乎没有文字穿过、两边都有足够文字的竖向空白带，即分栏的栏间距。"""
    if len(lines) < 8:
        return None
    lo, hi = int(page_w * 0.3), int(page_w * 0.7)
    cover = [0] * (hi - lo)
    for ln in lines:
        xs = set()
        for c in ln['chars']:
            xs.update(range(max(lo, int(c['x0'])), min(hi, int(c['x1']) + 1)))
        for x in xs:
            cover[x - lo] += 1
    limit = max(1, len(lines) * 0.08)
    best, start = None, None
    for i, v in enumerate(cover + [limit + 1]):
        if v <= limit and start is None:
            start = i
        elif v > limit and start is not None:
            if i - start >= 8 and (best is None or i - start > best[1] - best[0]):
                best = (start, i)
            start = None
    if not best:
        return None
    g0, g1 = lo + best[0], lo + best[1]
    left = sum(1 for ln in lines if any(c['x1'] <= g0 + 1 for c in ln['chars']))
    right = sum(1 for ln in lines if any(c['x0'] >= g1 - 1 for c in ln['chars']))
    return (g0, g1) if left >= 4 and right >= 4 else None


def _regions(lines, chars, page):
    """按分栏把页面切成上下若干区域：通栏区域、双栏区域。"""
    g = _gutter(lines, page.width)
    if not g:
        return [{'type': 'single', 'lines': lines}]
    g0, g1 = g
    crossing = lambda ln: any(g0 + 0.5 < (c['x0'] + c['x1']) / 2 < g1 - 0.5 for c in ln['chars'])
    bands = []
    for ln in sorted(lines, key=lambda l: l['b']):
        kind = 'single' if crossing(ln) else 'cols'
        if bands and bands[-1]['type'] == kind:
            bands[-1]['lines'].append(ln)
        else:
            bands.append({'type': kind, 'lines': [ln]})
    out = []
    for band in bands:
        if band['type'] == 'single':
            out.append(band)
            continue
        bases = [l['b'] for l in band['lines']]
        cs = [c for c in chars if any(abs(_baseline(c, page.height) - b) <= max(1.5, 0.25 * c['size'])
                                      for b in bases)]
        lt = _make_lines([c for c in cs if (c['x0'] + c['x1']) / 2 <= g0 + 0.5], page.height)
        rt = _make_lines([c for c in cs if (c['x0'] + c['x1']) / 2 >= g1 - 0.5], page.height)
        band['top'] = min(l['top'] for l in band['lines'])
        band['bottom'] = max(l['bottom'] for l in band['lines'])
        band['cols'] = [{'x0': min(l['x0'] for l in lt), 'x1': max(l['x1'] for l in lt), 'lines': lt},
                        {'x0': min(l['x0'] for l in rt), 'x1': max(l['x1'] for l in rt), 'lines': rt}] \
            if lt and rt else None
        if band['cols'] is None:
            band['type'] = 'single'
        out.append(band)
    return out


FOOTER_ZONE = 0.92   # 页面最下方 8% 视为页脚区


def _footer_text(lines):
    return re.sub(r'\d+', '#', ' '.join(l['text'] for l in lines))


def analyse(pdf_path):
    pages, total_chars = [], 0
    with pdfplumber.open(pdf_path) as pdf:
        raw = []
        for i, page in enumerate(pdf.pages):
            tables = _tables(page)
            tboxes = [t['bbox'] for t in tables]
            chars = [c for c in page.chars if not any(_inside(c, b, 1) for b in tboxes)]
            total_chars += sum(1 for c in page.chars if c['text'].strip())
            bars, boxes, rules = _shapes(page, tboxes)
            lines = _make_lines(chars, page.height)
            raw.append((page, chars, {'w': page.width, 'h': page.height, 'lines': lines, 'tables': tables,
                                      'bars': bars, 'boxes': boxes, 'rules': rules,
                                      'images': _images(pdf_path, i, page.height)}))
        # 页脚：每页最下方都有、而且除数字（页码）外内容一致的文字，放进 Word 真正的页脚
        feet = [[l for l in pg['lines'] if l['top'] >= pg['h'] * FOOTER_ZONE] for _, _, pg in raw]
        use_footer = all(feet) and len({_footer_text(f) for f in feet}) == 1
        for (page, chars, pg), foot in zip(raw, feet):
            if use_footer:
                zone = pg['h'] * FOOTER_ZONE
                pg['footer'] = {'lines': foot, 'rules': [r for r in pg['rules'] if r['y'] >= zone]}
                pg['lines'] = [l for l in pg['lines'] if l not in foot]
                pg['rules'] = [r for r in pg['rules'] if r['y'] < zone]
                chars = [c for c in chars if c['top'] < zone - 1]
            pg['regions'] = _regions(pg['lines'], chars, page)
            pages.append(pg)
    if total_chars < 10 * max(1, len(pages)):
        raise ScannedPDF('这是扫描件，暂不支持')
    return pages


# ---------- 排版 ----------

def _group(lines, content_right):
    """把连续、被自动换行切开的行合并成一个段落。"""
    paras = []
    for ln in lines:
        p = paras[-1] if paras else None
        if p and p['box'] is ln['box'] and not ln['bar'] and len(ln['segments']) == 1 \
                and len(p['lines'][-1]['segments']) == 1 and abs(ln['size'] - p['size']) < 0.6 \
                and ln['color'] == p['lines'][-1]['color'] and ln['text'][0] not in BULLETS:
            prev = p['lines'][-1]
            gap = ln['b'] - prev['b']
            full = prev['x1'] >= content_right - 1.5 * prev['size']
            aligned = abs(ln['x0'] - p['lines'][0]['x0']) <= 2 or abs(ln['x0'] - p.get('hang', -99)) <= 2
            pitch_ok = 1.0 * ln['size'] <= gap <= 2.0 * ln['size'] and \
                (len(p['lines']) == 1 or abs(gap - p['pitch']) <= 1.5)
            if full and aligned and pitch_ok:
                p['lines'].append(ln)
                p['pitch'] = gap if len(p['lines']) == 2 else p['pitch']
                continue
        para = {'kind': 'para', 'lines': [ln], 'box': ln['box'], 'bar': ln['bar'],
                'size': ln['size'], 'top': ln['top'], 'pitch': None}
        first = ln['segments'][0]['runs'][0]['text']
        if first[:1] in HANG_BULLETS and len(ln['chars']) > 1:
            # 项目符号后文字的实际起点，用作悬挂缩进：换行后的文字和第一行文字对齐
            para['hang'] = ln['chars'][1]['x0']
        paras.append(para)
    return paras


def _border(tag, val='single', sz=4, space=0, color='000000'):
    el = OxmlElement(f'w:{tag}')
    el.set(qn('w:val'), val)
    el.set(qn('w:sz'), str(int(max(2, min(96, sz)))))
    el.set(qn('w:space'), str(int(max(0, min(31, round(space))))))
    el.set(qn('w:color'), color)
    return el


# Word 要求 pPr、tblPr 里的子元素按规范顺序排列，顺序错了 Word 会报「文件已损坏」
PPR_ORDER = ('pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd '
             'tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi '
             'adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc '
             'textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange').split()
TBLPR_ORDER = ('tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc '
               'tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription').split()


def _insert_ordered(parent, el, order):
    tag = el.tag.split('}')[1]
    idx = order.index(tag)
    for child in parent:
        ctag = child.tag.split('}')[1] if isinstance(child.tag, str) else ''
        if ctag in order and order.index(ctag) > idx:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


def _ppr_append(p, el):
    return _insert_ordered(p._p.get_or_add_pPr(), el, PPR_ORDER)


def _plain(p):
    for tag in ('autoSpaceDE', 'autoSpaceDN'):
        el = OxmlElement(f'w:{tag}')
        el.set(qn('w:val'), '0')
        _ppr_append(p, el)
    el = OxmlElement('w:widowControl')
    el.set(qn('w:val'), '0')
    _ppr_append(p, el)


def _set_spacing(p, before, line):
    pf = p.paragraph_format
    pf.space_before = Pt(max(0, round(before, 1)))
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(round(line, 1))


def _font(run, name, size, bold=False, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rFonts.set(qn(a), name)
    if color and color != '000000':
        run.font.color.rgb = RGBColor.from_string(color)


class Builder:
    def __init__(self, pages, font):
        self.pages, self.font = pages, font
        self.doc = Document()
        st = self.doc.styles['Normal']
        st.paragraph_format.space_after = Pt(0)
        st.paragraph_format.space_before = Pt(0)
        self._img_id = 100
        p0 = pages[0]
        xs0 = [ln['x0'] for pg in pages for ln in pg['lines']] + \
              [o['x0'] for pg in pages for o in pg['rules'] + pg['boxes'] + pg['bars']] + \
              [t['bbox'][0] for pg in pages for t in pg['tables']]
        xs1 = [ln['x1'] for pg in pages for ln in pg['lines']] + \
              [o['x1'] for pg in pages for o in pg['rules'] + pg['boxes']] + \
              [t['bbox'][2] for pg in pages for t in pg['tables']]
        self.ml = max(18, min(xs0) - 0.5)
        self.content_right = max(xs1)
        self.mr = max(18, p0['w'] - self.content_right - 0.5)
        tops = [min([ln['top'] for ln in pg['lines']] + [t['top'] for t in pg['tables']] or [72]) for pg in pages]
        bots = [max([ln['bottom'] for ln in pg['lines']] + [t['bottom'] for t in pg['tables']] or [0]) for pg in pages]
        self.mt = max(12, min(tops) - 10)
        self.mb = max(10, min(36, min(pg['h'] - b for pg, b in zip(pages, bots)) - 6))
        foot = p0.get('footer')
        if foot:
            ftop = min([r['y'] - 2 for r in foot['rules']] + [l['top'] - 3 for l in foot['lines']])
            self.mb = max(10, p0['h'] - ftop)
        sec = self.doc.sections[0]
        sec.page_width, sec.page_height = Pt(p0['w']), Pt(p0['h'])
        sec.left_margin, sec.right_margin = Pt(self.ml), Pt(self.mr)
        sec.top_margin, sec.bottom_margin = Pt(self.mt), Pt(self.mb)
        sec.header_distance = sec.footer_distance = Pt(0)
        body = self.doc.element.body
        for p in list(body.iterchildren(qn('w:p'))):
            body.remove(p)

    # 每一行在 Word 里占的「行框」：基线以下留 DESC，向上占满行高
    @staticmethod
    def _box(line, L):
        bottom = line['b'] + DESC * line['size']
        return bottom - L, bottom

    def build(self):
        self.container, self.in_footer = self.doc, False
        for i, pg in enumerate(self.pages):
            self._page(pg, first=(i == 0))
        foot = self.pages[0].get('footer')
        if foot:
            self._footer(self.pages[0], foot)
        return self.doc

    def _footer(self, pg, foot):
        footer = self.doc.sections[0].footer
        footer.is_linked_to_previous = False
        for p in list(footer.paragraphs):
            p._p.getparent().remove(p._p)
        self.container, self.in_footer = footer, True
        self.page_break, keep_first = False, self.first_para
        tops = [r['y'] - max(0.5, r['w']) / 2 - 0.7 for r in foot['rules']] + \
               [l['b'] + DESC * l['size'] - l['size'] * 1.3 for l in foot['lines']]
        bottoms = [r['y'] + 1 for r in foot['rules']] + [l['b'] + DESC * l['size'] for l in foot['lines']]
        self.cursor = min(tops)
        self.cl, self.cr = self.ml, self.content_right
        self._flow(foot['lines'], [{'kind': 'rule', 'top': r['y'], **r} for r in foot['rules']], pg, self.content_right)
        for sec in self.doc.sections:
            sec.footer_distance = Pt(max(0, pg['h'] - max(bottoms)))
        self.container, self.in_footer, self.first_para = self.doc, False, keep_first

    def _page(self, pg, first):
        regions = pg['regions']
        spans = []
        for r in regions:
            ls = r['lines']
            spans.append((min(l['top'] for l in ls), max(l['bottom'] for l in ls)) if ls else (0, 0))
        bounds = [-1e9] + [(spans[i][1] + spans[i + 1][0]) / 2 for i in range(len(spans) - 1)] + [1e9]
        others = [{'kind': 'rule', 'top': r['y'], **r} for r in pg['rules']] + list(pg['tables'])
        self.cursor = self.mt
        self.page_break = not first
        self.first_para = None
        for i, r in enumerate(regions):
            mine = [o for o in others if bounds[i] <= o['top'] < bounds[i + 1]]
            if r['type'] == 'single':
                self.cl, self.cr = self.ml, self.content_right
                self._flow(r['lines'], mine, pg, self.content_right)
            else:
                self._columns(r, mine, pg)
        if self.first_para is None:
            self.first_para = self._new_p()
        for im in pg['images']:
            self._float_image(self.first_para, im)

    def _flow(self, lines, others, pg, right_edge):
        for ln in lines:
            ln['box'] = next((b for b in pg['boxes'] if _inside(ln, (b['x0'], b['top'], b['x1'], b['bottom']), 1)), None)
            ln['bar'] = next((b for b in pg['bars'] if b['x1'] <= ln['x0'] + 1 and ln['x0'] - b['x1'] <= 20
                              and b['top'] <= ln['bottom'] and b['bottom'] >= ln['top']), None)
        items = _group(lines, right_edge) + others
        items.sort(key=lambda o: o['top'])
        prev_box = None
        for it in items:
            if it['kind'] == 'para':
                self._para(it, it['box'], prev_box)
                prev_box = it['box']
            elif it['kind'] == 'rule':
                self._rule(it)
                prev_box = None
            else:
                self._table(it)
                prev_box = None

    def _section_break(self):
        """插入连续分节符（Word 用分节来切换分栏），并把分节符所在的空段落压到最小。"""
        sec = self.doc.add_section(WD_SECTION.CONTINUOUS)
        p = self.doc.paragraphs[-1]
        _plain(p)
        _set_spacing(p, 0, 0.7)
        self.cursor += 0.7
        return sec

    @staticmethod
    def _set_cols(sec, widths=None, space=0):
        sectPr = sec._sectPr
        cols = sectPr.find(qn('w:cols'))
        if cols is None:
            cols = OxmlElement('w:cols')
            sectPr.append(cols)
        for child in list(cols):
            cols.remove(child)
        for a in list(cols.attrib):
            del cols.attrib[a]
        if not widths:
            cols.set(qn('w:space'), '720')
            return
        cols.set(qn('w:num'), str(len(widths)))
        cols.set(qn('w:space'), str(round(space * 20)))
        cols.set(qn('w:equalWidth'), '0')
        for i, w in enumerate(widths):
            c = OxmlElement('w:col')
            c.set(qn('w:w'), str(round(w * 20)))
            if i < len(widths) - 1:
                c.set(qn('w:space'), str(round(space * 20)))
            cols.append(c)

    def _columns(self, region, others, pg):
        left, right = region['cols']
        self._section_break()
        sec = self.doc.sections[-1]
        w1 = left['x1'] - self.ml
        space = right['x0'] - left['x1']
        w2 = self.content_right - right['x0']
        top, bottoms = self.cursor, []
        for ci, col in enumerate((left, right)):
            if ci:
                # 单独一个极矮的段落放分栏符，避免把空行带进下一栏
                br = self._new_p()
                _set_spacing(br, 0, 0.7)
                br.add_run().add_break(WD_BREAK.COLUMN)
            self.cursor = top + (0.7 if ci else 0)
            self.cl, self.cr = (self.ml, left['x1']) if ci == 0 else (right['x0'], self.content_right)
            mine = [o for o in others if (ci == 0) == (((o.get('x0', o.get('bbox', [0])[0])) +
                                                        (o.get('x1', o.get('bbox', [0, 0, 0])[2]))) / 2 < right['x0'])]
            self._flow(col['lines'], mine, pg, col['x1'])
            bottoms.append(self.cursor)
        self._set_cols(sec, [w1, w2], space)
        new = self._section_break()
        self._set_cols(new)
        self.cursor = max(bottoms) + 0.7

    def _new_p(self):
        p = self.container.add_paragraph()
        _plain(p)
        if self.page_break:
            p.paragraph_format.page_break_before = True
            self.page_break = False
        if self.first_para is None:
            self.first_para = p
        return p

    def _para(self, para, box, prev_box):
        lines = para['lines']
        first, last = lines[0], lines[-1]
        L = para['pitch'] or round(first['size'] * 1.3, 1)
        top, _ = self._box(first, L)
        _, bottom = self._box(last, L)
        p = self._new_p()
        pf = p.paragraph_format
        before = top - self.cursor
        if box is not None:
            pad_t = max(0, top - box['top'])
            pad_b = max(0, box['bottom'] - bottom)
            bw = max(0.5, box['lw'] if box['stroke'] else 0.5)
            if box is not prev_box:
                before = box['top'] - self.cursor
            else:
                pad_t = 0
            space_l = max(1, first['x0'] - box['x0'] - bw)
            pf.left_indent = Pt(first['x0'] - self.cl)
            pf.right_indent = Pt(max(0, self.cr - box['x1'] + space_l))
            color = box['stroke'] or box['fill']
            pBdr = OxmlElement('w:pBdr')
            for side, sp in (('top', pad_t), ('left', space_l), ('bottom', pad_b), ('right', space_l)):
                pBdr.append(_border(side, sz=bw * 8, space=sp, color=color))
            _ppr_append(p, pBdr)
            if box['fill']:
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), box['fill'])
                _ppr_append(p, shd)
            self.cursor = box['bottom'] if not (para is not None and False) else bottom
        else:
            self.cursor = bottom
            ind = first['x0'] - self.cl
            if 'hang' in para:
                pf.left_indent = Pt(para['hang'] - self.cl)
                pf.first_line_indent = Pt(first['x0'] - para['hang'])
            elif ind > 0.5:
                pf.left_indent = Pt(ind)
            if para['bar']:
                bar = para['bar']
                space = max(0, first['x0'] - bar['x1'])
                pf.left_indent = Pt(first['x0'] - self.cl)
                pBdr = OxmlElement('w:pBdr')
                pBdr.append(_border('left', sz=(bar['x1'] - bar['x0']) * 8, space=space, color=bar['color']))
                _ppr_append(p, pBdr)
        if before < 0:
            L = max(first['size'] * 1.05, L + before)
            before = 0
        _set_spacing(p, before, L)
        if len(lines) > 1:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            right = max(l['x1'] for l in lines[:-1])
            if box is None:
                pf.right_indent = Pt(max(0, self.cr - right - 0.5))
        segs = first['segments']
        if len(segs) > 1:
            for s in segs[1:]:
                if s is segs[-1] and s['x1'] > (self.cl + 0.55 * (self.cr - self.cl)):
                    pf.tab_stops.add_tab_stop(Pt(s['x1'] - self.cl - (pf.left_indent.pt if pf.left_indent else 0)),
                                              WD_TAB_ALIGNMENT.RIGHT)
                else:
                    pf.tab_stops.add_tab_stop(Pt(s['x0'] - self.cl - (pf.left_indent.pt if pf.left_indent else 0)))
        for li, ln in enumerate(lines):
            for si, seg in enumerate(ln['segments']):
                if si:
                    p.add_run('\t')
                for r in seg['runs']:
                    text = r['text']
                    if li and seg is ln['segments'][0] and r is seg['runs'][0]:
                        text = text.lstrip() if lines[li - 1]['text'][-1:].isascii() is False else text
                    if li == 0 and si == 0 and r is seg['runs'][0] and 'hang' in para:
                        # 项目符号用西文字体（中文字体会显示成全角大圆点），后面接制表符对齐到悬挂缩进
                        _font(p.add_run(text[0]), 'Arial', r['size'], r['bold'], r['color'])
                        _font(p.add_run('\t'), self.font, r['size'])
                        text = text[1:].lstrip()
                    if self.in_footer and re.search(r'\d', text):
                        self._page_fields(p, text, r)
                    else:
                        _font(p.add_run(text), self.font, r['size'], r['bold'], r['color'])
        return p

    def _page_fields(self, p, text, r):
        """页脚里的「1 / 3」「第 1 页」换成会自动更新的页码。"""
        m = re.search(r'(\d+)\s*/\s*(\d+)', text)
        parts = [(m.start(1), m.end(1), 'PAGE'), (m.start(2), m.end(2), 'NUMPAGES')] if m else \
            [(mm.start(), mm.end(), 'PAGE') for mm in re.finditer(r'\d+', text)][:1]
        pos = 0
        for a, b, instr in parts:
            if text[pos:a]:
                _font(p.add_run(text[pos:a]), self.font, r['size'], r['bold'], r['color'])
            run = p.add_run(text[a:b])
            _font(run, self.font, r['size'], r['bold'], r['color'])
            fld = OxmlElement('w:fldSimple')
            fld.set(qn('w:instr'), f' {instr} ')
            run._r.addprevious(fld)
            fld.append(run._r)
            pos = b
        if text[pos:]:
            _font(p.add_run(text[pos:]), self.font, r['size'], r['bold'], r['color'])

    def _rule(self, r):
        p = self._new_p()
        bw = max(0.5, r['w'])
        _set_spacing(p, r['y'] - bw / 2 - 0.7 - self.cursor, 0.7)
        pf = p.paragraph_format
        if r['x0'] - self.cl > 1:
            pf.left_indent = Pt(r['x0'] - self.cl)
        if self.cr - r['x1'] > 1:
            pf.right_indent = Pt(self.cr - r['x1'])
        pBdr = OxmlElement('w:pBdr')
        pBdr.append(_border('bottom', sz=bw * 8, space=0, color=r['color']))
        _ppr_append(p, pBdr)
        _font(p.add_run(''), self.font, 1)
        self.cursor = r['y'] + bw / 2

    def _table(self, t):
        gap = t['top'] - self.cursor
        if gap > 0.8 or self.page_break:
            sp = self._new_p()
            _set_spacing(sp, 0, max(0.7, gap))
        xs, ys = t['xs'], t['ys']
        tbl = self.doc.add_table(rows=len(ys) - 1, cols=len(xs) - 1)
        tblPr = tbl._tbl.tblPr
        ind = OxmlElement('w:tblInd')
        ind.set(qn('w:w'), str(round((xs[0] - self.cl) * 20)))
        ind.set(qn('w:type'), 'dxa')
        _insert_ordered(tblPr, ind, TBLPR_ORDER)
        borders = OxmlElement('w:tblBorders')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            borders.append(_border(side, sz=t['lw'] * 8, color=t['color']))
        _insert_ordered(tblPr, borders, TBLPR_ORDER)
        layout = OxmlElement('w:tblLayout')
        layout.set(qn('w:type'), 'fixed')
        _insert_ordered(tblPr, layout, TBLPR_ORDER)
        mar = parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:left w:w="0" w:type="dxa"/>'
                        f'<w:right w:w="0" w:type="dxa"/></w:tblCellMar>')
        _insert_ordered(tblPr, mar, TBLPR_ORDER)
        for ci, (x0, x1) in enumerate(zip(xs, xs[1:])):
            tbl.columns[ci].width = Pt(x1 - x0)
        for ri, (y0, y1) in enumerate(zip(ys, ys[1:])):
            row = tbl.rows[ri]
            row.height = Pt(y1 - y0)
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            for ci, (x0, x1) in enumerate(zip(xs, xs[1:])):
                cell = row.cells[ci]
                cell.width = Pt(x1 - x0)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                p = cell.paragraphs[0]
                _plain(p)
                lines = t['rows'][ri][ci]
                if not lines:
                    _set_spacing(p, 0, 10)
                    continue
                size = lines[0]['size']
                _set_spacing(p, 0, round(size * 1.3, 1))
                p.paragraph_format.left_indent = Pt(max(0, lines[0]['x0'] - x0))
                mid = (x0 + x1) / 2
                if all(abs((l['x0'] + l['x1']) / 2 - mid) < 3 for l in lines) and lines[0]['x0'] - x0 > 6:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.left_indent = Pt(0)
                for li, ln in enumerate(lines):
                    if li:
                        p.add_run().add_break()
                    for si, seg in enumerate(ln['segments']):
                        if si:
                            p.add_run(' ')
                        for r in seg['runs']:
                            _font(p.add_run(r['text']), self.font, r['size'], r['bold'], r['color'])
        self.cursor = t['bottom']

    def _float_image(self, p, im):
        """把图片按原位置浮动放在页面上（文字环绕），不占用正文排版。"""
        self._img_id += 1
        run = p.add_run()
        w, h = Emu(int(Pt(im['x1'] - im['x0']))), Emu(int(Pt(im['bottom'] - im['top'])))
        inline = self.doc.part.new_pic_inline(io.BytesIO(im['png']), w, h)
        graphic = inline.find(qn('a:graphic'))
        x, y = int(Pt(im['x0'])), int(Pt(im['top']))
        anchor = parse_xml(
            f'<wp:anchor {nsdecls("wp", "a", "pic", "r")} distT="0" distB="0" distL="114300" distR="0" '
            f'simplePos="0" relativeHeight="{self._img_id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{y}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{int(w)}" cy="{int(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapSquare wrapText="bothSides"/>'
            f'<wp:docPr id="{self._img_id}" name="图片 {self._img_id}"/>'
            f'<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            f'</wp:anchor>')
        anchor.append(graphic)
        drawing = OxmlElement('w:drawing')
        drawing.append(anchor)
        run._r.append(drawing)


def convert(pdf_path, docx_path, font=None):
    """转换一份 PDF；返回 (页数, 每页字数)。扫描件抛出 ScannedPDF。"""
    pages = analyse(pdf_path)
    Builder(pages, font or default_font()).build().save(docx_path)
    return len(pages)


if __name__ == '__main__':
    convert(sys.argv[1], sys.argv[2], *(sys.argv[3:4]))
    print('ok', sys.argv[2])
