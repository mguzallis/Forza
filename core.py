#!/usr/bin/env python3
"""
Forza SDS formatter / rebrander - core engine.

extract()  : any SDS .docx  -> canonical model
render()   : canonical model + intake -> branded .docx built on a vertical shell

Design notes
------------
Nothing is copied blindly from the source document. Every paragraph and table is
re-emitted from parsed structure with Forza house formatting applied, which is
what prevents the subtle drift you get from carrying foreign XML across.

Preserved from source : text, bold/italic/underline/super/subscript, bullet and
                        numbered lists, tables (cell text + spans), inline images
Normalised            : font (Times New Roman), sizes, table style (Table Grid),
                        section heading form, spacing, hard page breaks (removed)
Regenerated           : title, top table, footer, DCN, headers/branding
"""

from __future__ import annotations

import io
import os
import re
import shutil
import zipfile
from dataclasses import dataclass, field

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

VERTICALS = ["Composites", "Construction", "Industrial",
             "Insulation", "Marine", "Transportation"]

SECTION_TITLES = {
    1: "Identification",
    2: "Hazard Identification",
    3: "Composition/Information on Ingredients",
    4: "First-Aid Measures",
    5: "Fire-Fighting Measures",
    6: "Accidental Release Measures",
    7: "Handling and Storage",
    8: "Exposure Controls / Personal Protection",
    9: "Physical and Chemical Properties",
    10: "Stability and Reactivity",
    11: "Toxicological Information",
    12: "Ecological Information",
    13: "Disposal Considerations",
    14: "Transport Information",
    15: "Regulatory Information",
    16: "Other Information",
}

# Any dash, any spacing, optional colon: "SECTION 1 - Identification", "Section 1:", "SECTION 1"
SECTION_RE = re.compile(r"^\s*SECTION\s+(\d{1,2})\s*[-\u2010-\u2015:.]?\s*(.*)$", re.I)

TNR = ('<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" '
       'w:hAnsi="Times New Roman" w:cs="Times New Roman"/>')


# ----------------------------------------------------------------- model ---
@dataclass
class Run:
    text: str = ""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    vert: str = ""            # "superscript" | "subscript" | ""
    image: str | None = None  # media filename once staged
    cx: int = 0
    cy: int = 0


@dataclass
class Para:
    runs: list = field(default_factory=list)
    list_kind: str = ""       # "bullet" | "number" | ""
    level: int = 0
    spacing_after: int | None = None   # dxa/20 units; None -> caller default

    @property
    def text(self):
        return "".join(r.text for r in self.runs)


@dataclass
class Table:
    grid: list = field(default_factory=list)   # column widths, dxa
    rows: list = field(default_factory=list)   # list[list[Cell]]


@dataclass
class Cell:
    paras: list = field(default_factory=list)
    span: int = 1
    width: int = 0


@dataclass
class Section:
    number: int
    title: str
    blocks: list = field(default_factory=list)  # Para | Table


@dataclass
class SDS:
    sections: list = field(default_factory=list)
    preamble: list = field(default_factory=list)
    media: dict = field(default_factory=dict)   # filename -> bytes
    detected: dict = field(default_factory=dict)


# ------------------------------------------------------------- extraction ---
def _xml_unescape(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&apos;", "'")
             .replace("&amp;", "&"))


def _xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;"))


def _blocks(body_xml):
    """Yield ('p'|'tbl', xml) for top-level children in document order."""
    i, n = 0, len(body_xml)
    while i < n:
        mp = body_xml.find("<w:p", i)
        mt = body_xml.find("<w:tbl>", i)
        if mp == -1 and mt == -1:
            return
        if mt == -1 or (mp != -1 and mp < mt):
            if body_xml[mp:mp + 5] not in ("<w:p>", "<w:p ") and body_xml[mp:mp + 6] != "<w:p/>":
                i = mp + 4
                continue
            if body_xml[mp:mp + 6] == "<w:p/>":
                yield "p", "<w:p/>"
                i = mp + 6
                continue
            end = body_xml.find("</w:p>", mp)
            if end == -1:
                return
            yield "p", body_xml[mp:end + 6]
            i = end + 6
        else:
            depth, j = 0, mt
            while True:
                nxt_o = body_xml.find("<w:tbl>", j + 1)
                nxt_c = body_xml.find("</w:tbl>", j + 1)
                if nxt_c == -1:
                    return
                if nxt_o != -1 and nxt_o < nxt_c:
                    depth += 1
                    j = nxt_o
                else:
                    if depth == 0:
                        yield "tbl", body_xml[mt:nxt_c + 8]
                        i = nxt_c + 8
                        break
                    depth -= 1
                    j = nxt_c


def _parse_runs(p_xml, rels, media_map):
    runs = []
    for m in re.finditer(r"<w:r(?:\s[^>]*)?>(.*?)</w:r>", p_xml, re.S):
        body = m.group(1)
        rpr = re.search(r"<w:rPr>(.*?)</w:rPr>", body, re.S)
        props = rpr.group(1) if rpr else ""
        bold = "<w:b/>" in props or '<w:b ' in props
        ital = "<w:i/>" in props or '<w:i ' in props
        und = "<w:u " in props
        vm = re.search(r'<w:vertAlign w:val="(\w+)"/>', props)
        vert = vm.group(1) if vm else ""

        emb = re.search(r'r:embed="([^"]+)"', body) or re.search(r'r:id="([^"]+)"', body)
        if emb and emb.group(1) in rels:
            tgt = rels[emb.group(1)]
            ext = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"/>', body)
            cx = int(ext.group(1)) if ext else 914400
            cy = int(ext.group(2)) if ext else 914400
            runs.append(Run(image=media_map.get(tgt, tgt), cx=cx, cy=cy))
            continue

        txt = "".join(_xml_unescape(t) for t in
                      re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", body, re.S))
        if not txt and re.search(r"<w:br\s*/>", body):
            runs.append(Run("\x00", bold, ital, und, vert))
            continue
        # <w:br/> inside a run becomes a paragraph split later; mark with \x00
        body_nobr = re.sub(r"<w:br\s*/>", "\x00", body)
        if "\x00" in body_nobr and txt:
            pieces = []
            for seg in re.split(r"\x00", body_nobr):
                pieces.append("".join(_xml_unescape(t) for t in
                                      re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", seg, re.S)))
            txt = "\x00".join(pieces)
        if txt:
            runs.append(Run(txt, bold, ital, und, vert))
    return runs


def _parse_para(p_xml, rels, media_map):
    ppr = re.search(r"<w:pPr>(.*?)</w:pPr>", p_xml, re.S)
    props = ppr.group(1) if ppr else ""
    kind, lvl = "", 0
    if "<w:numPr>" in props:
        kind = "bullet"
        lm = re.search(r'<w:ilvl w:val="(\d+)"/>', props)
        lvl = int(lm.group(1)) if lm else 0
    runs = _parse_runs(p_xml, rels, media_map)
    return Para(runs, kind, lvl)


def _parse_table(t_xml, rels, media_map):
    grid = [int(x) for x in re.findall(r'<w:gridCol w:w="(\d+)"', t_xml)]
    tbl = Table(grid=grid)
    depth = 0
    for m in re.finditer(r"<w:tr(?:\s[^>]*)?>(.*?)</w:tr>", t_xml, re.S):
        row_xml = m.group(1)
        cells = []
        for cm in re.finditer(r"<w:tc>(.*?)</w:tc>", row_xml, re.S):
            c_xml = cm.group(1)
            sm = re.search(r'<w:gridSpan w:val="(\d+)"/>', c_xml)
            wm = re.search(r'<w:tcW w:w="(\d+)"', c_xml)
            paras = [_parse_para(px, rels, media_map)
                     for kind, px in _blocks(c_xml) if kind == "p"]
            cells.append(Cell(paras,
                              int(sm.group(1)) if sm else 1,
                              int(wm.group(1)) if wm else 0))
        if cells:
            tbl.rows.append(cells)
    return tbl


def extract(path_or_bytes) -> SDS:
    """Parse any SDS .docx into the canonical model."""
    src = (zipfile.ZipFile(io.BytesIO(path_or_bytes))
           if isinstance(path_or_bytes, (bytes, bytearray))
           else zipfile.ZipFile(path_or_bytes))
    with src as z:
        doc = z.read("word/document.xml").decode("utf-8")
        try:
            rel_xml = z.read("word/_rels/document.xml.rels").decode("utf-8")
        except KeyError:
            rel_xml = ""
        rels = {m.group(1): m.group(2) for m in
                re.finditer(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rel_xml)}
        media = {}
        media_map = {}
        for name in z.namelist():
            if name.startswith("word/media/"):
                base = os.path.basename(name)
                media[base] = z.read(name)
                media_map["media/" + base] = base
                media_map[base] = base

    body = doc[doc.find("<w:body>"):]
    sds = SDS(media=media)
    current = None

    for kind, xml in _blocks(body):
        if kind == "tbl":
            tbl = _parse_table(xml, rels, media_map)
            (current.blocks if current else sds.preamble).append(tbl)
            continue

        para = _parse_para(xml, rels, media_map)
        raw = para.text.strip()
        m = SECTION_RE.match(raw) if raw else None
        # A heading is short and has no trailing sentence content
        if m and len(raw) < 90:
            num = int(m.group(1))
            if 1 <= num <= 16:
                title = m.group(2).strip(" -\u2013\u2014:") or SECTION_TITLES[num]
                current = Section(num, title)
                sds.sections.append(current)
                continue
        (current.blocks if current else sds.preamble).append(para)

    sds.detected = _detect(sds)
    inject_pictogram_images(sds)
    return sds


PICTOGRAM_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pictograms")

PICTOGRAM_FILES = {
    "1": "GHS01_exploding_bomb.png",
    "2": "GHS02_flame.png",
    "3": "GHS03_flame_over_circle.png",
    "4": "GHS04_gas_cylinder.png",
    "5": "GHS05_corrosion.png",
    "6": "GHS06_skull_and_crossbones.png",
    "7": "GHS07_exclamation_mark.png",
    "8": "GHS08_health_hazard.png",
    "9": "GHS09_environment.png",
}

_PICTO_CODE_RE = re.compile(r"GHS-?0*([1-9])\b", re.I)
_PICTO_KEYWORDS = [
    ("1", re.compile(r"exploding\s*bomb", re.I)),
    ("2", re.compile(r"\bflame\b(?!\s*over)", re.I)),
    ("3", re.compile(r"flame\s*over\s*circle|\boxidi[sz]", re.I)),
    ("4", re.compile(r"gas\s*cylinder", re.I)),
    ("5", re.compile(r"\bcorrosion\b", re.I)),
    ("6", re.compile(r"skull\s*(?:and|&)\s*crossbones", re.I)),
    ("7", re.compile(r"exclamation\s*mark", re.I)),
    ("8", re.compile(r"health\s*hazard", re.I)),
    ("9", re.compile(r"\benvironment\b", re.I)),
]
PICTOGRAM_LABEL_RE = re.compile(r"^(?:hazard\s+)?pictograms?\s*:?\s*(.*)$", re.I)
PICTOGRAM_SIDE_EMU = 902429  # ~0.987in square, matches the Forza reference convention


def _picto_codes(text: str) -> list:
    """Ordered, deduplicated GHS codes found in a short, already-scoped string.
    GHS-code matches ('GHS08') are tried first since they're unambiguous;
    keyword fallback ('Health Hazard') only runs if no code was found, and is
    only ever applied to text already scoped to a 'Pictograms' label - the
    same words appear constantly elsewhere in Section 2 hazard prose.
    """
    codes = []
    for m in _PICTO_CODE_RE.finditer(text):
        c = m.group(1)
        if c not in codes:
            codes.append(c)
    if not codes:
        hits = []
        for code, rx in _PICTO_KEYWORDS:
            m = rx.search(text)
            if m:
                hits.append((m.start(), code))
        hits.sort()
        for _, code in hits:
            if code not in codes:
                codes.append(code)
    return codes


def _picto_run(code: str, sds: SDS) -> Run:
    fname = PICTOGRAM_FILES[code]
    if fname not in sds.media:
        with open(os.path.join(PICTOGRAM_DIR, fname), "rb") as fh:
            sds.media[fname] = fh.read()
    return Run(image=fname, cx=PICTOGRAM_SIDE_EMU, cy=PICTOGRAM_SIDE_EMU)


def inject_pictogram_images(sds: SDS) -> SDS:
    """Where Section 2 lists hazard pictograms as text ('GHS08 (Health
    Hazard)') rather than actual images, replace that text with the real
    pictogram graphics. Detection is scoped strictly to whatever block is
    labeled 'Pictograms' - a block that already contains a real image is
    left untouched, since some sources (the Forza reference docs) already
    embed the actual artwork.
    """
    sec2 = next((s for s in sds.sections if s.number == 2), None)
    if not sec2:
        return sds

    i = 0
    while i < len(sec2.blocks):
        blk = sec2.blocks[i]
        if isinstance(blk, Para):
            m = PICTOGRAM_LABEL_RE.match(blk.text.strip())
            if m:
                has_image = any(r.image for r in blk.runs)
                trailing = m.group(1).strip()
                if not has_image and trailing:
                    codes = _picto_codes(trailing)
                    if codes:
                        blk.runs = ([Run("Hazard Pictograms: ", bold=True)] +
                                   [_picto_run(c, sds) for c in codes])
                elif not trailing and i + 1 < len(sec2.blocks):
                    nxt = sec2.blocks[i + 1]
                    if isinstance(nxt, Para) and not any(r.image for r in nxt.runs):
                        codes = _picto_codes(nxt.text)
                        if codes:
                            nxt.runs = [_picto_run(c, sds) for c in codes]
        elif isinstance(blk, Table):
            for row in blk.rows:
                if len(row) < 2:
                    continue
                label = " ".join(p.text for p in row[0].paras).strip()
                if re.match(r"^(?:hazard\s+)?pictograms?$", label, re.I):
                    has_image = any(r.image for p in row[1].paras for r in p.runs)
                    if not has_image:
                        value_text = " ".join(p.text for p in row[1].paras)
                        codes = _picto_codes(value_text)
                        if codes:
                            row[1].paras = [Para([_picto_run(c, sds) for c in codes])]
        i += 1
    return sds


def _detect(sds: SDS) -> dict:
    """Pull likely intake values out of the source so the UI can prefill.

    Section 1 identity fields are as likely to live in a table (OA4, OS2BT)
    as in bold-label paragraphs (the Forza reference doc, TAC850's Section 16).
    Both shapes are checked, with tolerant label matching, since real
    engineering drafts are not consistent about which they use.
    """
    out = {}
    label_pats = {
        "product_name": r"^product(\s*(name|identifier))?$",
        "other_means": r"other\s*means",
        "product_code": r"product\s*code",
        "recommended_use": r"recommended\s*use",
        "recommended_restrictions": r"recommended\s*restrictions|uses?\s*advised\s*against",
    }

    def check_label(label, value):
        for key, pat in label_pats.items():
            if key not in out and re.search(pat, label, re.I):
                out[key] = value
                return

    for sec in sds.sections:
        if sec.number != 1:
            continue
        for blk in sec.blocks:
            if isinstance(blk, Para):
                t = blk.text.strip()
                m = re.match(r"([A-Za-z][A-Za-z /]+?)\s*:\s*(.+)$", t)
                if m:
                    check_label(m.group(1).strip(), m.group(2).strip())
            elif isinstance(blk, Table):
                for row in blk.rows:
                    if len(row) >= 2:
                        label = " ".join(p.text for p in row[0].paras).strip()
                        value = " ".join(p.text for p in row[1].paras).strip()
                        if label and value:
                            check_label(label, value)

    for blk in sds.preamble:
        if isinstance(blk, Table):
            for row in blk.rows:
                if len(row) >= 2:
                    k = "".join(p.text for p in row[0].paras).strip().lower()
                    v = "".join(p.text for p in row[1].paras).strip()
                    if not v:
                        continue
                    if k.startswith("trade name"):
                        out.setdefault("trade_name", v)
                    elif k.startswith("sds"):
                        sm = re.match(r"(S-?\d+)\s*V?(\d+)?", v, re.I)
                        if sm:
                            out.setdefault("sds_number", sm.group(1))
                            if sm.group(2):
                                out.setdefault("version", sm.group(2))
                    elif k.startswith("date of issue") or k == "issue date":
                        out.setdefault("date_of_issue", normalize_date_string(v))
                    elif k.startswith("effective"):
                        out.setdefault("effective_date", normalize_date_string(v))
                    elif k.startswith("replaces"):
                        out.setdefault("replaces", v)
    return out


# ------------------------------------------------------------- rendering ---
def _rpr(run: Run, size=None, extra=""):
    p = TNR
    if run.bold:
        p += "<w:b/><w:bCs/>"
    if run.italic:
        p += "<w:i/><w:iCs/>"
    if run.underline:
        p += '<w:u w:val="single"/>'
    if run.vert:
        p += f'<w:vertAlign w:val="{run.vert}"/>'
    if size:
        p += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    return f"<w:rPr>{p}{extra}</w:rPr>"


def _emit_run(run: Run, rid_for):
    if run.image:
        rid = rid_for(run.image)
        if not rid:
            return ""
        cx, cy = run.cx or 914400, run.cy or 914400
        return (f'<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" '
                f'distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
                f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
                f'<wp:docPr id="{abs(hash(run.image)) % 90000 + 1000}" name="{run.image}"/>'
                f'<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/>'
                f'</wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.'
                f'openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr>'
                f'<pic:cNvPr id="{abs(hash(run.image)) % 90000 + 1000}" name="{run.image}"/>'
                f'<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/>'
                f'<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr>'
                f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
                f'</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
    if not run.text:
        return ""
    parts = run.text.split("\x00")
    out = ""
    for i, seg in enumerate(parts):
        if i:
            out += f"<w:r>{_rpr(run)}<w:br/></w:r>"
        if seg:
            out += (f"<w:r>{_rpr(run)}"
                    f'<w:t xml:space="preserve">{_xml_escape(seg)}</w:t></w:r>')
    return out


BULLET_NUMID = 991

BULLET_XML = (
    '<w:abstractNum w:abstractNumId="991"><w:multiLevelType w:val="hybridMultilevel"/>'
    + "".join(
        f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'<w:lvlText w:val="{ch}"/><w:lvlJc w:val="left"/><w:pPr>'
        f'<w:ind w:left="{720 + 360 * i}" w:hanging="360"/></w:pPr><w:rPr>'
        f'<w:rFonts w:ascii="{fnt}" w:hAnsi="{fnt}" w:hint="default"/></w:rPr></w:lvl>'
        for i, (ch, fnt) in enumerate(
            [("\uf0b7", "Symbol"), ("o", "Courier New"), ("\uf0a7", "Wingdings")] * 3)
    )
    + '</w:abstractNum>'
)

BULLET_NUM = '<w:num w:numId="991"><w:abstractNumId w:val="991"/></w:num>'


def _emit_para(p: Para, rid_for, spacing_after=120, keep_next=False):
    kn = "<w:keepNext/>" if keep_next else ""
    after = p.spacing_after if p.spacing_after is not None else spacing_after
    if p.list_kind:
        ppr = (f'<w:pPr><w:pStyle w:val="ListParagraph"/>{kn}<w:numPr>'
               f'<w:ilvl w:val="{min(p.level, 8)}"/><w:numId w:val="991"/></w:numPr>'
               f'<w:spacing w:after="40" w:line="240" w:lineRule="auto"/>'
               f'<w:contextualSpacing/><w:rPr>{TNR}</w:rPr></w:pPr>')
    else:
        ppr = (f'<w:pPr>{kn}<w:spacing w:after="{after}" w:line="240" '
               f'w:lineRule="auto"/><w:rPr>{TNR}</w:rPr></w:pPr>')
    runs = "".join(_emit_run(r, rid_for) for r in p.runs)
    return f"<w:p>{ppr}{runs}</w:p>"


def _emit_table(t: Table, rid_for):
    grid = t.grid or []
    if not grid and t.rows:
        cols = max(len(r) for r in t.rows)
        grid = [int(8996 / cols)] * cols
    total = sum(grid) or 8996
    if total > 9360:                       # keep inside Letter text column
        scale = 8996 / total
        grid = [max(400, int(g * scale)) for g in grid]

    tblpr = ('<w:tblPr><w:tblW w:w="0" w:type="auto"/>'
             '<w:tblBorders>'
             '<w:top w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:left w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:bottom w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:right w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:insideH w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:insideV w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
             '</w:tblBorders>'
             '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
             '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>'
             '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" '
             'w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr>')
    gridxml = "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{g}"/>' for g in grid) + "</w:tblGrid>"

    rows = ""
    for ri, row in enumerate(t.rows):
        cells = ""
        ci = 0
        for cell in row:
            span = max(1, cell.span)
            wide = sum(grid[ci:ci + span]) or (cell.width or 1500)
            ci += span
            spanxml = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
            body = "".join(
                _emit_para(p, rid_for, spacing_after=0) for p in cell.paras
            ) or f'<w:p><w:pPr><w:rPr>{TNR}</w:rPr></w:pPr></w:p>'
            cells += (f'<w:tc><w:tcPr><w:tcW w:w="{wide}" w:type="dxa"/>{spanxml}'
                      f'<w:vAlign w:val="center"/></w:tcPr>{body}</w:tc>')
        hdr = '<w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' if ri == 0 else \
              '<w:trPr><w:cantSplit/></w:trPr>'
        rows += f"<w:tr>{hdr}{cells}</w:tr>"
    spacer = f'<w:p><w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/>' \
             f'<w:rPr>{TNR}</w:rPr></w:pPr></w:p>'
    return f"<w:tbl>{tblpr}{gridxml}{rows}</w:tbl>{spacer}"


LABEL_RE = re.compile(r".*[:\u2013-]\s*$")


def _has_image(blk):
    return isinstance(blk, Para) and any(r.image for r in blk.runs)


def _is_label(blk):
    """A short line ending in a colon, i.e. a heading for what follows."""
    if not isinstance(blk, Para) or _has_image(blk):
        return False
    t = blk.text.strip()
    return bool(t) and len(t) < 120 and LABEL_RE.match(t) is not None


def _cohesion(blocks):
    """Decide which paragraphs must not be orphaned from what follows them.

    Hard page breaks are never inserted. Instead a paragraph is bound to the
    next block when splitting there would strand a label, so Word reflows the
    pair onto the following page by itself.
    """
    keep = [False] * len(blocks)
    for i, blk in enumerate(blocks):
        if i + 1 >= len(blocks):
            continue
        nxt = blocks[i + 1]
        if isinstance(blk, Table):
            continue
        # an image never separates from the image beside or above it
        if _has_image(blk) and (_has_image(nxt) or isinstance(nxt, Table)):
            keep[i] = True
            continue
        # a label stays with its table, its image, or its first value line
        if _is_label(blk):
            keep[i] = True
            continue
        # the line before a table stays with the table
        if isinstance(nxt, Table):
            keep[i] = True
            continue
    # a short list is a single visual unit: never split two or three items
    i = 0
    while i < len(blocks):
        blk = blocks[i]
        if isinstance(blk, Para) and blk.list_kind:
            j = i
            while (j < len(blocks) and isinstance(blocks[j], Para)
                   and blocks[j].list_kind):
                j += 1
            if 2 <= (j - i) <= 3:
                for k in range(i, j - 1):
                    keep[k] = True
            i = j
        else:
            i += 1
    # never let a bound run grow past five paragraphs, or a whole page shifts
    run = 0
    for i, k in enumerate(keep):
        if k:
            run += 1
            if run > 5:
                keep[i] = False
                run = 0
        else:
            run = 0
    return keep


def _emit_heading(num, title):
    txt = _xml_escape(f"SECTION {num} \u2013 {title}")
    rpr = (f'<w:rPr>{TNR}<w:b/><w:bCs/><w:kern w:val="0"/><w:sz w:val="36"/>'
           f'<w:szCs w:val="36"/><w:u w:val="thick"/></w:rPr>')
    return (f'<w:p><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120" '
            f'w:line="240" w:lineRule="auto"/>{rpr}</w:pPr>'
            f"<w:r>{rpr}<w:t>{txt}</w:t></w:r></w:p>")


# --------------------------------------------------------------- helpers ---
DATE_DOT_RE = re.compile(r"(?<!\d)(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?!\d)")


def normalize_date_string(s: str) -> str:
    """Convert dot-separated dates (08.18.2026) to slash-separated
    (08/18/2026). Everything else - including non-date text like 'N/A' or
    'Revision 6' in the Replaces field - passes through untouched.
    """
    if not s:
        return s
    return DATE_DOT_RE.sub(lambda m: f"{m.group(1)}/{m.group(2)}/{m.group(3)}", s)


def build_dcn(sds_number: str, version: str, trade_name: str) -> str:
    """S-0196 + 1 + R-OS86  ->  'S0196_V1 - R-OS86'  (en dash, per reference)."""
    core = re.sub(r"[^0-9A-Za-z]", "", str(sds_number))
    ver = re.sub(r"\D", "", str(version)) or "1"
    return f"{core}_V{ver} \u2013 {trade_name}".strip()


def build_sds_field(sds_number: str, version: str) -> str:
    """Top-table form keeps the hyphen: 'S-0196 V1'."""
    num = str(sds_number).strip()
    ver = re.sub(r"\D", "", str(version)) or "1"
    return f"{num} V{ver}"


S1_LABEL_PATTERNS = {
    "product_name": r"^product(\s*(name|identifier))?$",
    "other_means": r"other\s*means",
    "product_code": r"product\s*code",
    "recommended_use": r"recommended\s*use",
    "recommended_restrictions": r"recommended\s*restrictions|uses?\s*advised\s*against",
}


def canonical_section1(intake: dict) -> list:
    """Section 1 is always rendered in Forza's fixed house format, regardless
    of how the source document structured it. Only the five identity/use
    fields vary by product; the supplier block is constant company boilerplate
    and is never taken from the source, since engineering drafts often carry
    placeholder brackets there (e.g. '[Company name and address - insert]').
    """
    B = lambda t: Run(t, bold=True)
    return [
        Para([B("Product Name: "), Run(intake.get("product_name", ""))], spacing_after=120),
        Para([B("Other Means of Identification: "), Run(intake.get("other_means", ""))],
             spacing_after=120),
        Para([B("Product Code Number: "), Run(intake.get("product_code", ""))],
             spacing_after=120),
        Para([B("Recommended Use: "), Run(intake.get("recommended_use", ""))],
             spacing_after=120),
        Para([B("Recommended Restrictions: "), Run(intake.get("recommended_restrictions", ""))],
             spacing_after=200),
        Para([Run("Suppliers Details")], spacing_after=120),
        Para([B("Company: ")], spacing_after=20),
        Para([Run("Forza, Inc.")], spacing_after=0),
        Para([Run("3211 Nebraska Ave, Suite #300")], spacing_after=0),
        Para([Run("Council Bluffs, IA 51501, USA")], spacing_after=140),
        Para([B("Company Phone Number: ")], spacing_after=20),
        Para([Run("402-731-9300 (Available 8:00 am \u2013 4:30 pm CST)")], spacing_after=140),
        Para([B("Emergency Phone Number: ")], spacing_after=20),
        Para([Run("Chemtrec 1(800)-424-9300")], spacing_after=120),
    ]


PREPARED_BY_RE = re.compile(r"prepared\s*by", re.I)
FORZA_PREPARED_BY = "Forza, Inc."


def _s16_route(label: str, ver: str, date: str) -> str | None:
    """Decide what a Section 16 label wants: a version, a date, or both.

    'Revision' -> version. 'Issue Date' -> date. 'Revision date' contains
    both words but means a date, not a version, so a combined field is only
    assumed when the label has an explicit joiner ('/', '&', 'and') between
    the two concepts, e.g. 'Revision / Issue Date'.
    """
    has_date = bool(re.search(r"date", label, re.I))
    has_ver = bool(re.search(r"revision|version", label, re.I))
    combined = bool(re.search(r"(revision|version)\s*[/&]\s*\w*\s*date", label, re.I)
                    or re.search(r"(revision|version)\s+and\s+\w*\s*date", label, re.I))
    if combined:
        return f"V{ver} \u2013 {date}" if date else f"V{ver}"
    if has_date:
        return date or None
    if has_ver:
        return f"V{ver}"
    return None


PLACEHOLDER_BRACKET_RE = re.compile(r"\[[^\[\]]{1,80}\]")
PLACEHOLDER_KEYWORD_RE = re.compile(r"\b(TBD|PLACEHOLDER|FIXME|XXXX+)\b", re.I)

# Section 1 boilerplate is generated fresh every time and never sourced from
# the document, so it can never itself be short or contain a placeholder.
# Only sections 2-16 are worth scanning.
_EMPTY_SECTION_CHARS = 15   # below this with no table present -> "empty"
_SHORT_SECTION_CHARS = 60   # below this -> "unusually short", worth a glance


def _section_text_len(sec: Section) -> tuple:
    """(char_count, has_table) for a section's blocks."""
    chars, has_table = 0, False
    for blk in sec.blocks:
        if isinstance(blk, Table):
            has_table = True
            for row in blk.rows:
                for cell in row:
                    chars += sum(len(p.text) for p in cell.paras)
        else:
            chars += len(blk.text)
    return chars, has_table


def preflight(sds: SDS, intake: dict) -> dict:
    """Surface what might be wrong before the person downloads, rather than
    after. Nothing here blocks generation - it's a pre-flight check, not a
    validator - but a silent gap (a missing section, a leftover
    '[insert here]', a Section 16 with no date field to stamp) is much
    cheaper to catch here than after the document has gone out.
    """
    found = sorted(s.number for s in sds.sections)
    missing = [n for n in range(1, 17) if n not in found]

    empty, short, placeholders = [], [], []
    for sec in sds.sections:
        if sec.number == 1:
            continue  # always regenerated fresh; nothing to check
        chars, has_table = _section_text_len(sec)
        title = sec.title or SECTION_TITLES.get(sec.number, "")
        if chars < _EMPTY_SECTION_CHARS and not has_table:
            empty.append((sec.number, title))
        elif chars < _SHORT_SECTION_CHARS:
            short.append((sec.number, title, chars))

        for blk in sec.blocks:
            texts = ([p.text for row in blk.rows for c in row for p in c.paras]
                     if isinstance(blk, Table) else [blk.text])
            for t in texts:
                for rx in (PLACEHOLDER_BRACKET_RE, PLACEHOLDER_KEYWORD_RE):
                    for m in rx.finditer(t):
                        snippet = t[max(0, m.start() - 20):m.end() + 20].strip()
                        placeholders.append((sec.number, title, snippet))

    sec16 = next((s for s in sds.sections if s.number == 16), None)
    ver_found, date_found = False, False
    if sec16:
        for blk in sec16.blocks:
            rows = ([" ".join(p.text for p in row[0].paras).strip()
                     for row in blk.rows if row] if isinstance(blk, Table)
                    else [re.match(r"([A-Za-z][A-Za-z /]+?)\s*:\s*", blk.text.strip())])
            labels = rows if isinstance(blk, Table) else \
                     [m.group(1) for m in rows if m]
            for label in labels:
                has_date = bool(re.search(r"date", label, re.I))
                has_ver = bool(re.search(r"revision|version", label, re.I))
                date_found = date_found or has_date
                ver_found = ver_found or has_ver

    blank_optional = [f for f in ("recommended_use", "recommended_restrictions",
                                  "other_means", "product_code")
                      if not str(intake.get(f, "")).strip()]

    warnings = []
    if missing:
        warnings.append(f"Missing sections: {', '.join(map(str, missing))}.")
    if empty:
        names = ", ".join(f"{n} ({t})" for n, t, in empty)
        warnings.append(f"These sections parsed with no real content: {names}.")
    if short:
        names = ", ".join(f"{n} ({t}, {c} chars)" for n, t, c in short)
        warnings.append(f"Unusually short, worth a glance: {names}.")
    if placeholders:
        warnings.append(f"{len(placeholders)} placeholder-looking snippet(s) "
                        f"found (e.g. brackets or TBD/FIXME) - see detail below.")
    if not ver_found:
        warnings.append("Section 16 has no field that looks like a version/"
                        "revision - the version number will not appear there.")
    if not date_found:
        warnings.append("Section 16 has no field that looks like a date - "
                        "the date of issue will not appear there.")
    if blank_optional:
        warnings.append(f"Not filled in: {', '.join(blank_optional)}.")

    return dict(sections_found=found, sections_missing=missing,
                empty_sections=empty, short_sections=short,
                placeholders=placeholders,
                section16=dict(version_field_present=ver_found,
                               date_field_present=date_found),
                blank_optional_fields=blank_optional,
                images_total=len(sds.media), warnings=warnings)


def apply_intake(sds: SDS, intake: dict) -> SDS:
    """Stamp Section 16 date / version onto whatever shape the source used.

    Section 1 is not touched here: it is fully replaced at render time by
    canonical_section1(), regardless of what the source contained, so no
    in-place patching of it is needed or attempted.
    """
    ver = re.sub(r"\D", "", str(intake.get("version", "1"))) or "1"
    intake = dict(intake)
    intake["date_of_issue"] = normalize_date_string(intake.get("date_of_issue", ""))

    def set_cell(cell, text):
        cell.paras = [Para([Run(text)])]

    for sec in sds.sections:
        if sec.number == 16:
            prepared_found = False
            last_dated_idx = -1
            for i, blk in enumerate(sec.blocks):
                if isinstance(blk, Para):
                    t = blk.text.strip()
                    m = re.match(r"([A-Za-z][A-Za-z /]+?)\s*:\s*", t)
                    if not m:
                        continue
                    label = m.group(1).strip()
                    if PREPARED_BY_RE.search(label):
                        blk.runs = [Run(label + ": ", bold=True), Run(FORZA_PREPARED_BY)]
                        prepared_found = True
                        continue
                    val = _s16_route(label, ver, intake.get("date_of_issue", ""))
                    if val is not None:
                        blk.runs = [Run(label + ": ", bold=True), Run(val)]
                        last_dated_idx = i
                elif isinstance(blk, Table):
                    for row in blk.rows:
                        if len(row) < 2:
                            continue
                        label = " ".join(p.text for p in row[0].paras).strip()
                        if PREPARED_BY_RE.search(label):
                            set_cell(row[1], FORZA_PREPARED_BY)
                            prepared_found = True
                            continue
                        val = _s16_route(label, ver, intake.get("date_of_issue", ""))
                        if val is not None:
                            set_cell(row[1], val)
                            last_dated_idx = i

            if not prepared_found:
                # Some engineering drafts omit "Prepared by" entirely (TAC850
                # does). It must always appear, so insert it - right after the
                # revision/date line(s) if any were found, otherwise at the
                # top of the section.
                new_para = Para([Run("Prepared by: ", bold=True), Run(FORZA_PREPARED_BY)])
                sec.blocks.insert(last_dated_idx + 1 if last_dated_idx >= 0 else 0, new_para)
    return sds


TOP_ROWS = ["trade_name", "sds_field", "replaces", "date_of_issue", "effective_date"]
TOP_LABELS = ["Trade Name", "SDS #", "Replaces", "Date of Issue", "Effective Date"]


def render(sds: SDS, intake: dict, shell_path: str, out_path: str) -> str:
    """Inject the model into a vertical shell and write the finished .docx."""
    ver = re.sub(r"\D", "", str(intake.get("version", "1"))) or "1"
    values = {
        "trade_name": intake.get("trade_name", ""),
        "sds_field": build_sds_field(intake.get("sds_number", ""), ver),
        "replaces": normalize_date_string(intake.get("replaces", "")),
        "date_of_issue": normalize_date_string(intake.get("date_of_issue", "")),
        "effective_date": normalize_date_string(intake.get("effective_date", "")),
    }
    dcn = build_dcn(intake.get("sds_number", ""), ver, intake.get("trade_name", ""))

    with zipfile.ZipFile(shell_path) as z:
        parts = {n: z.read(n) for n in z.namelist()}

    # stage images, allocating fresh relationship ids
    rel_extra, ct_extra, used = "", "", {}
    next_id = [1]

    def rid_for(fname):
        if fname in used:
            return used[fname]
        if fname not in sds.media:
            return None
        rid = f"rIdImg{next_id[0]}"
        next_id[0] += 1
        used[fname] = rid
        parts[f"word/media/{fname}"] = sds.media[fname]
        return rid

    # ---- body
    doc = parts["word/document.xml"].decode("utf-8")
    chunks = []
    for sec in sorted(sds.sections, key=lambda s: s.number):
        chunks.append(_emit_heading(sec.number, sec.title or SECTION_TITLES.get(sec.number, "")))
        if sec.number == 1:
            # Always the fixed Forza house format, regardless of source shape.
            for p in canonical_section1(intake):
                chunks.append(_emit_para(p, rid_for))
            continue
        blocks = [b for b in sec.blocks if isinstance(b, Table) or b.runs]
        keep = _cohesion(blocks)
        for blk, kn in zip(blocks, keep):
            if isinstance(blk, Table):
                chunks.append(_emit_table(blk, rid_for))
            else:
                chunks.append(_emit_para(blk, rid_for, keep_next=kn))
    body_xml = "".join(chunks)

    marker = re.search(r"<w:p>(?:(?!</w:p>).)*\{\{BODY\}\}.*?</w:p>", doc, re.S)
    doc = doc[:marker.start()] + body_xml + doc[marker.end():]
    for key, val in values.items():
        doc = doc.replace("{{" + key.upper() + "}}", _xml_escape(val))
    doc = doc.replace("{{SDS_NUMBER}} V{{VERSION}}", _xml_escape(values["sds_field"]))
    doc = re.sub(r"\{\{[A-Z_]+\}\}", "", doc)
    parts["word/document.xml"] = doc.encode("utf-8")

    # ---- footer DCN
    ftr = parts["word/footer1.xml"].decode("utf-8")
    parts["word/footer1.xml"] = ftr.replace("{{DCN}}", _xml_escape(dcn)).encode("utf-8")

    # ---- relationships + content types for staged images
    rl = parts["word/_rels/document.xml.rels"].decode("utf-8")
    add = "".join(f'<Relationship Id="{rid}" Type="{R}/image" Target="media/{f}"/>'
                  for f, rid in used.items())
    parts["word/_rels/document.xml.rels"] = rl.replace("</Relationships>", add + "</Relationships>").encode("utf-8")

    num = parts["word/numbering.xml"].decode("utf-8")
    if 'w:numId="991"' not in num:
        # schema order: every abstractNum must precede every num
        first_num = num.find("<w:num ")
        if first_num == -1:
            num = num.replace("</w:numbering>", BULLET_XML + BULLET_NUM + "</w:numbering>")
        else:
            num = num[:first_num] + BULLET_XML + num[first_num:]
            num = num.replace("</w:numbering>", BULLET_NUM + "</w:numbering>")
    parts["word/numbering.xml"] = num.encode("utf-8")

    ct = parts["[Content_Types].xml"].decode("utf-8")
    for ext, mime in (("png", "image/png"), ("jpeg", "image/jpeg"), ("jpg", "image/jpeg"),
                      ("gif", "image/gif"), ("emf", "image/x-emf"), ("wmf", "image/x-wmf")):
        if f'Extension="{ext}"' not in ct:
            ct = ct.replace("</Types>", f'<Default Extension="{ext}" ContentType="{mime}"/></Types>')
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for name, data in parts.items():
            z.writestr(name, data)
    return out_path
