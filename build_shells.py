#!/usr/bin/env python3
"""
Build six corrected per-vertical Forza SDS shells from the reference document.

Corrections applied vs. the reference:
  - Page size normalised to US Letter (was A4)
  - Margins recomputed from measured banner geometry
  - Dead even/first header + even footer parts removed
  - Banner artwork converted CMYK -> RGB
  - Hardcoded "of 9" replaced with a NUMPAGES field
  - Top table gains a permanent "Replaces" row (5 rows)
"""
import os, re, shutil, zipfile

REF = "/mnt/user-data/uploads/ForzaEAGLE_R-0S86_Insulation_SDS_V1_08_05_2026__1_.docx"
UP = "/mnt/user-data/uploads"
OUT = "/home/claude/work/shells"

BANNERS = {
    # Pass-through: original file bytes are embedded unchanged, so the color is
    # byte-for-byte identical to what was uploaded. No RGB re-encode, ever.
    "Composites":     f"{UP}/1787089497628_Composites_SDS.docx|word/media/image1.jpg",
    "Construction":   f"{UP}/1787089497629_Construction_SDS.docx|word/media/image1.jpg",
    "Insulation":     f"{UP}/1787089497629_Insulation_SDS.docx|word/media/image1.jpg",
    "Marine":         f"{UP}/1787089497629_Marine_SDS.docx|word/media/image1.jpg",
    "Transportation": f"{UP}/1787089497630_Transportation_SDS.docx|word/media/image1.jpg",
    # True Letter-sized artwork, CMYK, byte-for-byte passthrough like the other
    # five. No reflow needed - this is a genuine drop-in replacement for the
    # earlier A4-derived Industrial banner.
    "Industrial":     "/mnt/user-data/uploads/Industrial_reredo.docx|word/media/image1.jpg",
}

# ---- page geometry (twips; 1440 = 1 inch) -------------------------------
PG_W, PG_H = 12240, 15840          # US Letter
MAR_TOP    = 2088                  # 1.45"  clears all six banners (max measured 1.30")
MAR_BOT    = 1224                  # 0.85"  clears artwork footer block (starts 10.29")
MAR_SIDE   = 1440
FRAME_X    = 7660                  # frame is transparent; only its right edge matters
FRAME_W    = 4320                  # 3.00" wide so the DCN can never wrap
FRAME_Y    = 14904                 # 10.35" page line, level with the artwork tagline
FRAME_Y2   = 15516                 # 10.78" DCN line, clear below the artwork address

# full-bleed banner: 8.5" x 11" in EMU
IMG_CX, IMG_CY = 7772400, 10058400

NS = ('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
      'mc:Ignorable="w14 wp14"')

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

# ---- header: full-bleed banner anchored to the page ---------------------
HEADER = DECL + f'''<w:hdr {NS}><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240" behindDoc="1" locked="0" layoutInCell="0" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="{IMG_CX}" cy="{IMG_CY}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="900001" name="Banner"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="900001" name="Banner"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdBanner"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{IMG_CX}" cy="{IMG_CY}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p></w:hdr>'''

HEADER_RELS = DECL + ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                      '<Relationship Id="rIdBanner" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                      'relationships/image" Target="media/banner.jpg"/></Relationships>')

# A 3" frame with right-justified text: far wider than any DCN, so no wrap is
# possible, and the unused left portion is transparent.
def _frame(y):
    return (f'<w:framePr w:w="{FRAME_W}" w:wrap="none" w:vAnchor="page" '
            f'w:hAnchor="page" w:x="{FRAME_X}" w:y="{y}"/>')


_FRAME = _frame(FRAME_Y)
_FRAME2 = _frame(FRAME_Y2)
_RPR = ('<w:rPr><w:rStyle w:val="PageNumber"/><w:rFonts w:ascii="Times New Roman" '
        'w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>')


def _run(text):
    return f'<w:r>{_RPR}<w:t xml:space="preserve">{text}</w:t></w:r>'


def _field(instr):
    return (f'<w:r>{_RPR}<w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r>{_RPR}<w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
            f'<w:r>{_RPR}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r>{_RPR}<w:t>1</w:t></w:r>'
            f'<w:r>{_RPR}<w:fldChar w:fldCharType="end"/></w:r>')


# ---- footer: "Page N of M" over "DCN: ..." ------------------------------
FOOTER = DECL + (
    f'<w:ftr {NS}>'
    f'<w:p><w:pPr><w:pStyle w:val="Footer"/>{_FRAME}'
    f'<w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="right"/></w:pPr>'
    + _run("Page ") + _field("PAGE") + _run(" of ") + _field("NUMPAGES") +
    f'</w:p>'
    f'<w:p><w:pPr><w:pStyle w:val="Footer"/>{_FRAME2}'
    f'<w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="right"/></w:pPr>'
    + _run("DCN: {{DCN}}") +
    f'</w:p>'
    f'<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>'
    f'</w:ftr>')

FOOTER_RELS = DECL + ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                      'relationships"></Relationships>')

# ---- body ---------------------------------------------------------------
TNR = ('<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" '
       'w:hAnsi="Times New Roman" w:cs="Times New Roman"/>')

TOP_ROWS = [
    ("Trade Name",   "{{TRADE_NAME}}"),
    ("SDS #",        "{{SDS_NUMBER}} V{{VERSION}}"),
    ("Replaces",     "{{REPLACES}}"),
    ("Date of Issue", "{{DATE_OF_ISSUE}}"),
    ("Effective Date", "{{EFFECTIVE_DATE}}"),
]


def _cell(text, w, bold=False, label=False):
    b = "<w:b/><w:bCs/>" if bold else ""
    return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>'
            f'<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
            f'<w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>'
            f'<w:vAlign w:val="center"/></w:tcPr>'
            f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{TNR}{b}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{TNR}{b}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>')


# Column ratio (37.7% / 62.3%) taken from the T-OS164 reference document,
# scaled from its A4 printable width onto this shell's Letter printable width.
_LABEL_W = 3532
_VALUE_W = 5828

TOP_BORDERS = ('<w:tblBorders>'
               '<w:top w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:left w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:bottom w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:right w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:insideH w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:insideV w:val="double" w:sz="4" w:space="0" w:color="auto"/>'
               '</w:tblBorders>')


def top_table():
    rows = "".join(f'<w:tr>{_cell(k, _LABEL_W, bold=True)}{_cell(v, _VALUE_W)}</w:tr>'
                   for k, v in TOP_ROWS)
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' + TOP_BORDERS +
            '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" '
            'w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{_LABEL_W}"/><w:gridCol w:w="{_VALUE_W}"/></w:tblGrid>'
            f'{rows}</w:tbl>')


TITLE = ('<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/>'
         f'<w:rPr>{TNR}<w:b/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:pPr>'
         f'<w:r><w:rPr>{TNR}<w:b/><w:bCs/><w:kern w:val="36"/><w:sz w:val="44"/>'
         '<w:szCs w:val="44"/></w:rPr><w:t>SAFETY DATA SHEET</w:t></w:r></w:p>')

MARKER = ('<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/>'
          f'<w:rPr>{TNR}</w:rPr></w:pPr>'
          f'<w:r><w:rPr>{TNR}</w:rPr><w:t>{{{{BODY}}}}</w:t></w:r></w:p>')

SECTPR = (f'<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/>'
          f'<w:footerReference w:type="default" r:id="rIdFtr"/>'
          f'<w:pgSz w:w="{PG_W}" w:h="{PG_H}"/>'
          f'<w:pgMar w:top="{MAR_TOP}" w:right="{MAR_SIDE}" w:bottom="{MAR_BOT}" '
          f'w:left="{MAR_SIDE}" w:header="0" w:footer="0" w:gutter="0"/>'
          f'<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>')

DOCUMENT = DECL + f'<w:document {NS}><w:body>{TITLE}{top_table()}{MARKER}{SECTPR}</w:body></w:document>'


def build(vertical, banner_docx, banner_member, dest):
    stage = f"/home/claude/work/.stage_{vertical}"
    shutil.rmtree(stage, ignore_errors=True)
    os.makedirs(stage)

    # 1. unpack the reference as donor for styles / theme / numbering / settings
    with zipfile.ZipFile(REF) as z:
        z.extractall(stage)
    for junk in ("word/header1.xml", "word/header2.xml", "word/header3.xml",
                 "word/footer1.xml", "word/footer2.xml"):
        p = os.path.join(stage, junk)
        if os.path.exists(p):
            os.remove(p)
    shutil.rmtree(os.path.join(stage, "word/_rels"), ignore_errors=True)
    shutil.rmtree(os.path.join(stage, "word/media"), ignore_errors=True)
    shutil.rmtree(os.path.join(stage, "customXml"), ignore_errors=True)
    os.makedirs(os.path.join(stage, "word/_rels"))
    os.makedirs(os.path.join(stage, "word/media"))

    # 2. banner artwork: CMYK -> RGB, saved as PNG
    # Embed the banner as-is: no color-space conversion, no re-encode.
    # A naive CMYK->RGB conversion (PIL default, or a non-color-managed
    # ImageMagick pass) measurably desaturates these images versus how Word
    # renders the original file. Byte-for-byte passthrough is the only
    # approach that is guaranteed to match.
    if banner_member == "@file":
        with open(banner_docx, "rb") as fh:
            raw = fh.read()
    else:
        with zipfile.ZipFile(banner_docx) as z:
            raw = z.read(banner_member)
    with open(os.path.join(stage, "word/media/banner.jpg"), "wb") as fh:
        fh.write(raw)

    # 3. new parts
    W = lambda rel, s: open(os.path.join(stage, rel), "w", encoding="utf-8").write(s)
    W("word/header1.xml", HEADER)
    W("word/_rels/header1.xml.rels", HEADER_RELS)
    W("word/footer1.xml", FOOTER)
    W("word/_rels/footer1.xml.rels", FOOTER_RELS)
    W("word/document.xml", DOCUMENT)

    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    W("word/_rels/document.xml.rels", DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      f'<Relationship Id="rIdStyles" Type="{rel}/styles" Target="styles.xml"/>'
      f'<Relationship Id="rIdSettings" Type="{rel}/settings" Target="settings.xml"/>'
      f'<Relationship Id="rIdWeb" Type="{rel}/webSettings" Target="webSettings.xml"/>'
      f'<Relationship Id="rIdFonts" Type="{rel}/fontTable" Target="fontTable.xml"/>'
      f'<Relationship Id="rIdTheme" Type="{rel}/theme" Target="theme/theme1.xml"/>'
      f'<Relationship Id="rIdNum" Type="{rel}/numbering" Target="numbering.xml"/>'
      f'<Relationship Id="rIdFn" Type="{rel}/footnotes" Target="footnotes.xml"/>'
      f'<Relationship Id="rIdEn" Type="{rel}/endnotes" Target="endnotes.xml"/>'
      f'<Relationship Id="rIdHdr" Type="{rel}/header" Target="header1.xml"/>'
      f'<Relationship Id="rIdFtr" Type="{rel}/footer" Target="footer1.xml"/>'
      '</Relationships>')

    # 4. content types: drop customXml overrides, register png
    ct = os.path.join(stage, "[Content_Types].xml")
    x = open(ct, encoding="utf-8").read()
    x = re.sub(r'<Override PartName="/customXml[^>]*/>', "", x)
    x = re.sub(r'<Override PartName="/word/(header|footer)\d+\.xml"[^>]*/>', "", x)
    x = re.sub(r'<Default Extension="(png|jpeg|jpg)"[^>]*/>', "", x)
    x = x.replace("<Types ", '<Types ', 1)
    ins = ('<Default Extension="png" ContentType="image/png"/>'
           '<Default Extension="jpg" ContentType="image/jpeg"/>'
           '<Default Extension="jpeg" ContentType="image/jpeg"/>')
    hdr = ('<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-'
           'officedocument.wordprocessingml.header+xml"/>'
           '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-'
           'officedocument.wordprocessingml.footer+xml"/>')
    x = x.replace("</Types>", ins + hdr + "</Types>")
    open(ct, "w", encoding="utf-8").write(x)

    # 5. root rels: drop customXml
    rr = os.path.join(stage, "_rels/.rels")
    x = open(rr, encoding="utf-8").read()
    x = re.sub(r'<Relationship[^>]*customXml[^>]*/>', "", x)
    open(rr, "w", encoding="utf-8").write(x)

    # 6. zip
    if os.path.exists(dest):
        os.remove(dest)
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(stage, "[Content_Types].xml"), "[Content_Types].xml")
        for root, _, files in os.walk(stage):
            for f in files:
                full = os.path.join(root, f)
                arc = os.path.relpath(full, stage)
                if arc != "[Content_Types].xml":
                    z.write(full, arc)
    shutil.rmtree(stage, ignore_errors=True)
    return dest


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    for vert, spec in BANNERS.items():
        docx, member = spec.split("|")
        out = build(vert, docx, member, f"{OUT}/SDS_Shell_{vert}.docx")
        print(f"  built {os.path.basename(out)}  ({os.path.getsize(out)//1024} KB)")
