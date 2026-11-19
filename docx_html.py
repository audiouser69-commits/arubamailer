# -*- coding: utf-8 -*-
"""
Conversione .docx -> HTML per email (stili inline, compatibile con Outlook/Gmail)
+ testo semplice alternativo + gestione logo.
Mantiene: grassetto, corsivo, sottolineato, barrato, apice/pedice, colore, evidenziato,
dimensione e tipo di carattere, allineamento, titoli, elenchi puntati/numerati (anche annidati),
tabelle, collegamenti, immagini incorporate, interruzioni di riga.
"""
import html
import io
import os

EMU_PER_PX = 9525
LOGO_MAX_W, LOGO_MAX_H = 180, 60        # dimensioni visualizzate del logo (px)
IMG_MAX_W = 640                          # larghezza massima immagini del documento
FONT_STACK = "Calibri, Arial, Helvetica, sans-serif"

HIGHLIGHT = {
    "YELLOW": "#ffff00", "BRIGHT_GREEN": "#00ff00", "TURQUOISE": "#00ffff", "PINK": "#ff00ff",
    "BLUE": "#0000ff", "RED": "#ff0000", "DARK_BLUE": "#000080", "TEAL": "#008080",
    "GREEN": "#008000", "VIOLET": "#800080", "DARK_RED": "#800000", "DARK_YELLOW": "#808000",
    "GRAY_50": "#808080", "GRAY_25": "#c0c0c0", "BLACK": "#000000", "WHITE": "#ffffff",
}
HEADING_SIZES = {1: 24, 2: 20, 3: 17, 4: 15, 5: 14, 6: 13}

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
V_NS = "urn:schemas-microsoft-com:vml"


def _w(tag):
    return "{%s}%s" % (W_NS, tag)


def esc(t):
    return html.escape(t, quote=False)


# ---------------------------------------------------------------- immagini
def prepare_image(data, max_w, max_h=None, scale=2):
    """Ritorna (bytes, subtype, larghezza_visualizzata, altezza_visualizzata).
    Converte qualsiasi formato letto da Pillow in PNG (se trasparente/pochi colori) o JPEG."""
    from PIL import Image, ImageOps
    im = Image.open(io.BytesIO(data))
    try:
        im.seek(0)
    except Exception:
        pass
    im = ImageOps.exif_transpose(im)
    w, h = im.size
    if w <= 0 or h <= 0:
        raise ValueError("Immagine non valida")
    r = min(1.0, max_w / float(w), (max_h / float(h)) if max_h else 1.0)
    dw, dh = max(1, int(round(w * r))), max(1, int(round(h * r)))
    # salvo a risoluzione doppia (schermi ad alta densita) ma mai oltre l'originale
    sw, sh = min(w, dw * scale), min(h, dh * scale)
    has_alpha = im.mode in ("RGBA", "LA", "P") and (im.mode != "P" or "transparency" in im.info)
    if has_alpha:
        im = im.convert("RGBA")
    else:
        im = im.convert("RGB")
    if (sw, sh) != im.size:
        im = im.resize((sw, sh), Image.LANCZOS)
    out = io.BytesIO()
    ncolors = len(im.getcolors(4096) or []) if im.size[0] * im.size[1] < 4_000_000 else 99999
    if has_alpha or (0 < ncolors < 4096):
        im.save(out, "PNG", optimize=True)
        sub = "png"
    else:
        im.save(out, "JPEG", quality=85, optimize=True, progressive=True)
        sub = "jpeg"
    return out.getvalue(), sub, dw, dh


def load_logo(path):
    """Legge un logo (png, jpg, gif, bmp, webp, tiff, ico...) e lo prepara per l'email."""
    if os.path.splitext(path)[1].lower() == ".svg":
        raise ValueError("Il formato SVG non è visualizzato dalla maggior parte dei programmi di posta: "
                         "esportalo in PNG.")
    with open(path, "rb") as f:
        data = f.read()
    return prepare_image(data, LOGO_MAX_W, LOGO_MAX_H)


# ---------------------------------------------------------------- docx
class DocxConverter:
    def __init__(self, path):
        import docx
        self.docx = docx
        self.doc = docx.Document(path)
        self.images = []        # (cid, bytes, subtype)
        self.text_lines = []
        self._num_fmt_cache = {}
        normal = None
        try:
            normal = self.doc.styles["Normal"]
        except KeyError:
            pass
        self.base_font = (normal.font.name if normal is not None and normal.font.name else None)
        self.base_size = (normal.font.size.pt if normal is not None and normal.font.size else 11)

    # ---- elenchi
    def _num_format(self, num_id, ilvl):
        key = (num_id, ilvl)
        if key in self._num_fmt_cache:
            return self._num_fmt_cache[key]
        fmt = "bullet"
        try:
            numbering = self.doc.part.numbering_part.element
            abs_id = None
            for num in numbering.findall(_w("num")):
                if num.get(_w("numId")) == str(num_id):
                    a = num.find(_w("abstractNumId"))
                    abs_id = a.get(_w("val")) if a is not None else None
                    break
            if abs_id is not None:
                for an in numbering.findall(_w("abstractNum")):
                    if an.get(_w("abstractNumId")) == abs_id:
                        for lvl in an.findall(_w("lvl")):
                            if lvl.get(_w("ilvl")) == str(ilvl):
                                nf = lvl.find(_w("numFmt"))
                                if nf is not None:
                                    fmt = nf.get(_w("val")) or "bullet"
                        break
        except Exception:
            pass
        self._num_fmt_cache[key] = fmt
        return fmt

    def _list_info(self, para):
        """(livello, 'ul'|'ol', stile css lista) oppure None."""
        pPr = para._p.pPr
        num_id = ilvl = None
        if pPr is not None and pPr.numPr is not None:
            if pPr.numPr.numId is not None:
                num_id = pPr.numPr.numId.val
            if pPr.numPr.ilvl is not None:
                ilvl = pPr.numPr.ilvl.val
        style_name = (para.style.name or "").lower() if para.style is not None else ""
        if num_id is None:
            # numerazione definita nello stile (es. "List Bullet", "Elenco puntato")
            st = para.style
            while st is not None and num_id is None:
                spPr = st.element.pPr
                if spPr is not None and spPr.numPr is not None and spPr.numPr.numId is not None:
                    num_id = spPr.numPr.numId.val
                    if spPr.numPr.ilvl is not None and ilvl is None:
                        ilvl = spPr.numPr.ilvl.val
                st = st.base_style
        if num_id is None or num_id == 0:
            if "list bullet" in style_name:
                return (0, "ul", "disc")
            if "list number" in style_name:
                return (0, "ol", "decimal")
            return None
        ilvl = ilvl or 0
        import re as _re
        m = _re.match(r"(list bullet|list number|elenco puntato|elenco numerato)\s*(\d)$", style_name)
        if m and ilvl == 0:
            ilvl = max(0, int(m.group(2)) - 1)
            fmt = self._num_format(num_id, 0)
            if "number" in style_name or "numerato" in style_name:
                return (ilvl, "ol", "decimal" if fmt == "bullet" else {"lowerLetter": "lower-alpha"}.get(fmt, "decimal"))
            return (ilvl, "ul", ["disc", "circle", "square"][ilvl % 3])
        fmt = self._num_format(num_id, ilvl)
        mapping = {"decimal": "decimal", "lowerLetter": "lower-alpha", "upperLetter": "upper-alpha",
                   "lowerRoman": "lower-roman", "upperRoman": "upper-roman", "decimalZero": "decimal-leading-zero"}
        if fmt in mapping:
            return (ilvl, "ol", mapping[fmt])
        return (ilvl, "ul", ["disc", "circle", "square"][ilvl % 3])

    # ---- run
    def _run_html(self, run, heading=False):
        parts = []
        el = run._element
        # immagini nel run
        for blip in el.iter("{%s}blip" % A_NS):
            rid = blip.get("{%s}embed" % R_NS)
            if rid:
                parts.append(self._image_html(rid, el))
        text = run.text or ""
        if not text:
            return "".join(parts)
        t = esc(text).replace("\t", "&emsp;").replace("\n", "<br>")
        f = run.font
        css = []
        if f.color is not None and f.color.type is not None and f.color.rgb is not None:
            css.append("color:#%s" % str(f.color.rgb))
        if f.size is not None and not heading:
            css.append("font-size:%.1fpt" % f.size.pt)
        if f.name and f.name != self.base_font:
            css.append("font-family:'%s', %s" % (f.name.replace("'", ""), FONT_STACK))
        try:
            hl = f.highlight_color
            if hl is not None and hl.name in HIGHLIGHT:
                css.append("background-color:%s" % HIGHLIGHT[hl.name])
        except Exception:
            pass
        # ombreggiatura carattere (w:shd)
        rPr = el.rPr
        if rPr is not None:
            shd = rPr.find(_w("shd"))
            if shd is not None:
                fill = shd.get(_w("fill"))
                if fill and fill.lower() not in ("auto", "ffffff"):
                    css.append("background-color:#%s" % fill)
        deco = []
        if run.underline:
            deco.append("underline")
        if f.strike or f.double_strike:
            deco.append("line-through")
        if deco:
            css.append("text-decoration:%s" % " ".join(deco))
        if f.all_caps:
            css.append("text-transform:uppercase")
        if f.small_caps:
            css.append("font-variant:small-caps")
        if css:
            t = '<span style="%s">%s</span>' % (";".join(css), t)
        if run.bold:
            t = "<strong>%s</strong>" % t
        elif run.bold is False and heading:
            t = '<span style="font-weight:normal">%s</span>' % t
        if run.italic:
            t = "<em>%s</em>" % t
        if f.superscript:
            t = "<sup>%s</sup>" % t
        elif f.subscript:
            t = "<sub>%s</sub>" % t
        parts.append(t)
        return "".join(parts)

    def _image_html(self, rid, run_el):
        try:
            part = self.doc.part.related_parts[rid]
            data = part.blob
        except Exception:
            return ""
        # dimensione impostata in Word
        want_w = None
        for ext in run_el.iter("{%s}extent" % WP_NS):
            try:
                want_w = int(ext.get("cx")) / EMU_PER_PX
            except Exception:
                pass
            break
        max_w = min(IMG_MAX_W, int(want_w)) if want_w else IMG_MAX_W
        try:
            b, sub, dw, dh = prepare_image(data, max_w)
        except Exception:
            return ""
        cid = "img%d" % (len(self.images) + 1)
        self.images.append((cid, b, sub))
        return ('<img src="cid:%s" width="%d" height="%d" alt="" '
                'style="border:0;max-width:100%%;height:auto;vertical-align:middle;">' % (cid, dw, dh))

    def _inline_html(self, para, heading=False):
        out = []
        text = []
        items = para.iter_inner_content() if hasattr(para, "iter_inner_content") else para.runs
        for item in items:
            if item.__class__.__name__ == "Hyperlink":
                inner = "".join(self._run_html(r, heading) for r in item.runs)
                addr = item.address or ""
                if item.fragment and not addr:
                    out.append(inner)
                else:
                    url = addr + (("#" + item.fragment) if item.fragment else "")
                    out.append('<a href="%s" style="color:#0563c1;text-decoration:underline;">%s</a>'
                               % (html.escape(url, quote=True), inner))
                    link_text = "".join(r.text for r in item.runs)
                    text.append(link_text + (" (%s)" % url if url and url not in link_text else ""))
                    continue
                text.append("".join(r.text for r in item.runs))
            else:
                out.append(self._run_html(item, heading))
                text.append(item.text or "")
        return "".join(out), "".join(text)

    def _para_style(self, para):
        css = []
        al = para.paragraph_format.alignment
        if al is None and para.style is not None:
            st = para.style
            while st is not None and al is None:
                al = st.paragraph_format.alignment
                st = st.base_style
        if al is not None:
            name = getattr(al, "name", str(al)).upper()
            if "CENTER" in name:
                css.append("text-align:center")
            elif "RIGHT" in name:
                css.append("text-align:right")
            elif "JUSTIFY" in name or "DISTRIBUTE" in name:
                css.append("text-align:justify")
        ind = para.paragraph_format.left_indent
        if ind:
            css.append("padding-left:%dpx" % int(ind.pt * 1.333))
        return css

    def _heading_level(self, para):
        name = (para.style.name or "") if para.style is not None else ""
        low = name.lower()
        if low == "title" or low == "titolo":
            return 1
        for prefix in ("heading ", "titolo "):
            if low.startswith(prefix):
                try:
                    return max(1, min(6, int(low[len(prefix):].strip())))
                except ValueError:
                    pass
        return 0

    def _render_paragraphs(self, paras, out, lines):
        """paras: lista di oggetti Paragraph o Table nell'ordine del documento."""
        stack = []      # [(tag, livello)]
        counters = {}

        def close_to(level):
            while stack and stack[-1][1] >= level:
                out.append("</li></%s>" % stack.pop()[0])

        for block in paras:
            if block.__class__.__name__ == "Table":
                close_to(0)
                self._render_table(block, out, lines)
                continue
            para = block
            li = self._list_info(para)
            inner, text = self._inline_html(para, heading=bool(self._heading_level(para)))
            if li:
                level, tag, ltype = li
                if stack and stack[-1][1] == level and stack[-1][0] == tag:
                    out.append("</li><li style=\"margin:0 0 4px 0;\">")
                else:
                    if stack and stack[-1][1] > level:
                        close_to(level + 1)
                    if stack and stack[-1][1] == level:
                        close_to(level)
                    out.append('<%s style="margin:0 0 10px 0;padding-left:28px;list-style-type:%s;">'
                               '<li style="margin:0 0 4px 0;">' % (tag, ltype))
                    stack.append((tag, level))
                css = self._para_style(para)
                out.append('<span style="%s">%s</span>' % (";".join(css), inner) if css else inner)
                for k in [k for k in counters if k > level]:
                    del counters[k]
                if tag == "ol":
                    counters[level] = counters.get(level, 0) + 1
                    bullet = "%d. " % counters[level]
                else:
                    counters.pop(level, None)
                    bullet = "- "
                lines.append("  " * level + bullet + text)
                continue
            close_to(0)
            counters.clear()
            css = self._para_style(para)
            hl = self._heading_level(para)
            if hl:
                css = ["font-size:%dpx" % HEADING_SIZES[hl], "font-weight:bold", "margin:16px 0 8px 0",
                       "line-height:1.25"] + css
                out.append('<p style="%s">%s</p>' % (";".join(css), inner or "&nbsp;"))
                lines.append("")
                lines.append(text.upper() if hl <= 2 else text)
                lines.append("")
            else:
                css = ["margin:0 0 10px 0"] + css
                out.append('<p style="%s">%s</p>' % (";".join(css), inner if inner.strip() else "&nbsp;"))
                lines.append(text)
        close_to(0)

    def _block_items(self, parent_el, parent_obj):
        from docx.table import Table
        from docx.text.paragraph import Paragraph
        items = []
        for child in parent_el.iterchildren():
            if child.tag == _w("p"):
                items.append(Paragraph(child, parent_obj))
            elif child.tag == _w("tbl"):
                items.append(Table(child, parent_obj))
            elif child.tag == _w("sdt"):  # controlli contenuto
                content = child.find(_w("sdtContent"))
                if content is not None:
                    items.extend(self._block_items(content, parent_obj))
        return items

    def _render_table(self, table, out, lines):
        out.append('<table role="presentation" cellpadding="6" cellspacing="0" border="0" '
                   'style="border-collapse:collapse;margin:0 0 12px 0;width:100%;">')
        for row in table.rows:
            out.append("<tr>")
            seen = set()
            row_text = []
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                span = cell._tc.grid_span if hasattr(cell._tc, "grid_span") else 1
                shade = ""
                tcPr = cell._tc.tcPr
                if tcPr is not None:
                    shd = tcPr.find(_w("shd"))
                    if shd is not None and shd.get(_w("fill")) not in (None, "auto", "FFFFFF", "ffffff"):
                        shade = "background-color:#%s;" % shd.get(_w("fill"))
                cell_out, cell_lines = [], []
                self._render_paragraphs(self._block_items(cell._tc, cell), cell_out, cell_lines)
                html_cell = "".join(cell_out).replace("margin:0 0 10px 0", "margin:0")
                out.append('<td valign="top"%s style="border:1px solid #cccccc;%s">%s</td>'
                           % (' colspan="%d"' % span if span > 1 else "", shade, html_cell))
                row_text.append(" ".join(l for l in cell_lines if l.strip()))
            out.append("</tr>")
            lines.append(" | ".join(row_text))
        out.append("</table>")
        lines.append("")

    def convert_with_images(self, images):
        """Come convert(), ma aggiunge le immagini del documento alla lista 'images'."""
        res = self.convert()
        images.extend(self.images)
        return res

    def convert(self):
        out, lines = [], []
        self._render_paragraphs(self._block_items(self.doc.element.body, self.doc), out, lines)
        # rimuovo righe vuote multiple dal testo
        text, blank = [], 0
        for l in lines:
            if l.strip():
                text.append(l.rstrip())
                blank = 0
            else:
                blank += 1
                if blank == 1 and text:
                    text.append("")
        font = FONT_STACK if not self.base_font else "'%s', %s" % (self.base_font.replace("'", ""), FONT_STACK)
        return "".join(out), "\n".join(text).strip() + "\n", font, self.base_size


# ---------------------------------------------------------------- composizione email
def wrap_html(body_html, font_family=FONT_STACK, font_size_pt=11, logo=None):
    """logo: (cid, larghezza, altezza) o None."""
    logo_row = ""
    if logo:
        cid, w, h = logo
        logo_row = ('<tr><td align="right" style="padding:0 0 14px 0;">'
                    '<img src="cid:%s" width="%d" height="%d" alt="" '
                    'style="display:inline-block;border:0;outline:none;text-decoration:none;'
                    'width:%dpx;height:%dpx;max-width:%dpx;"></td></tr>' % (cid, w, h, w, h, LOGO_MAX_W))
    return ('<!DOCTYPE html><html lang="it"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width, initial-scale=1"></head>'
            '<body style="margin:0;padding:0;background-color:#ffffff;">'
            '<table role="presentation" width="100%%" cellpadding="0" cellspacing="0" border="0">'
            '<tr><td align="center" style="padding:16px 12px;">'
            '<table role="presentation" width="100%%" cellpadding="0" cellspacing="0" border="0" '
            'style="max-width:680px;">%s'
            '<tr><td align="left" style="font-family:%s;font-size:%.1fpt;line-height:1.45;color:#222222;">'
            '%s</td></tr></table></td></tr></table></body></html>'
            % (logo_row, font_family, font_size_pt, body_html))


def text_to_html(text):
    paras = text.replace("\r\n", "\n").split("\n\n")
    return "".join('<p style="margin:0 0 10px 0;">%s</p>' % esc(p).replace("\n", "<br>") for p in paras)
