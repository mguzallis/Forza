"""Draft-row export for the shared product-data repo (the `sds` tab).

Builds ONE row per generated SDS and hands it to the shared helper
`tools/master_repo.py` (copied into the app unchanged):

    master_repo.save_draft("sds", row)

This module never writes to data/, never touches another tab, and never
raises into the app: if the helper or token is missing, the caller gets a
clear message and the normal .docx / .zip outputs are unaffected.
"""
from __future__ import annotations

import re
from datetime import date

import core

TAB = "sds"

COLUMNS = [
    "Vertical", "Product Folder", "File Name", "Is Latest Version",
    "File Version", "File Date", "SDS #", "Header Date",
    "Legacy SDS No. / Layout", "Product Name (SDS)", "Product Code (SDS)",
    "Signal Word", "Classification", "Pictograms", "Hazard Statements",
    "Precautionary Statements", "Ingredients (Name | CAS | %)", "Handling",
    "Storage", "Disposal", "Parse Flags", "Notes",
]

# Standard GHS/CLP hazard-statement -> pictogram mapping. Used only when the
# source shows pictograms as images (no GHS0X text) so codes can't be read.
_H_TO_GHS = {
    "GHS01": r"H20[0-4]|H240|H241",
    "GHS02": r"H22[0-8]|H24[0-2]|H250|H251|H252|H260|H261",
    "GHS03": r"H27[0-2]",
    "GHS04": r"H280|H281",
    "GHS05": r"H290|H314|H318",
    "GHS06": r"H300|H301|H310|H311|H330|H331",
    "GHS07": r"H302|H312|H315|H319|H332|H335|H336",
    "GHS08": r"H304|H334|H340|H341|H350|H351|H360|H361|H362|H370|H371|H372|H373",
    "GHS09": r"H400|H410|H411",
}

_CODE_LINE = re.compile(r"^\s*((?:[HP]\d{3})(?:\s*\+\s*[HP]?\d{3})*)\s*[:\-–.]?\s*(.*)$")
_LINE_LABEL = re.compile(r"^\s*([A-Za-z][A-Za-z /&()\-]{2,60}?)\s*:\s*(.*)$")


def _lines(blocks) -> list:
    """Flatten Para blocks to text lines (manual breaks are \\x00)."""
    out = []
    for b in blocks:
        if isinstance(b, core.Para):
            for ln in b.text.split("\x00"):
                ln = ln.strip()
                if ln:
                    out.append(ln)
    return out


def _section(sds, n):
    return next((s for s in sds.sections if s.number == n), None)


def _labelled(lines, wanted_rx, stop_rx=None):
    """Return text under the first label line matching wanted_rx, up to the
    next label line. Handles 'Label: value' and 'Label:' + following lines."""
    got, grabbing = [], False
    for ln in lines:
        m = _LINE_LABEL.match(ln)
        if m and re.search(wanted_rx, m.group(1), re.I):
            grabbing = True
            if m.group(2).strip():
                got.append(m.group(2).strip())
            continue
        if grabbing:
            if _LINE_LABEL.match(ln) and not _CODE_LINE.match(ln):
                break
            got.append(ln)
    return "\n".join(got).strip()


def _statements(lines, letter) -> list:
    seen, out = set(), []
    for ln in lines:
        m = _CODE_LINE.match(ln)
        if not m or not m.group(1).startswith(letter):
            continue
        code = re.sub(r"\s+", "", m.group(1))
        text = m.group(2).strip()
        if not text:
            continue
        if text[-1] not in ".!?":
            text += "."
        item = f"{code} {text}"
        if item not in seen:
            seen.add(item)
            out.append(item)
    return out


def _pictograms(sec2, h_items) -> str:
    codes = []
    text = "\n".join(_lines(sec2.blocks)) if sec2 else ""
    for m in re.finditer(r"GHS\s?0?([1-9])", text):
        c = f"GHS0{m.group(1)}"
        if c not in codes:
            codes.append(c)
    if not codes:
        hcodes = " ".join(i.split()[0] for i in h_items)
        for ghs, rx in _H_TO_GHS.items():
            if re.search(rf"\b(?:{rx})\b", hcodes):
                codes.append(ghs)
    return ", ".join(sorted(codes))


def _ingredients(sec3) -> str:
    if not sec3:
        return ""
    for blk in sec3.blocks:
        if not isinstance(blk, core.Table) or len(blk.rows) < 2:
            continue
        cell = lambda c: " ".join(p.text.replace("\x00", " ") for p in c.paras).strip()
        head = [cell(c).lower() for c in blk.rows[0]]
        cas_i = next((i for i, h in enumerate(head) if "cas" in h), None)
        if cas_i is None:
            continue
        pct_i = next((i for i, h in enumerate(head)
                      if re.search(r"%|conc|weight|content", h)), None)
        out = []
        for row in blk.rows[1:]:
            vals = [cell(c) for c in row]
            if not vals or not vals[0]:
                continue
            pct = vals[pct_i] if pct_i is not None and pct_i < len(vals) else ""
            cas = vals[cas_i] if cas_i < len(vals) else ""
            out.append(f"{vals[0]} | {cas} | {pct}")
        if out:
            return "\n".join(out)
    return ""


def build_sds_row(sds, intake: dict, *, vertical: str, file_name: str) -> dict:
    """One `sds` row for the English document just generated.

    `sds` is the parsed (and intake-applied) core.SDS, `intake` is the dict
    passed to core.render (English SDS #), `file_name` is the English .docx
    name (the tab's current key).
    """
    s1 = _section(sds, 1)
    s2 = _section(sds, 2)
    s3 = _section(sds, 3)
    s7 = _section(sds, 7)
    s13 = _section(sds, 13)

    l2 = _lines(s2.blocks) if s2 else []
    haz = _statements(l2, "H")
    prec = _statements(l2, "P")

    l7 = _lines(s7.blocks) if s7 else []
    handling = _labelled(l7, r"handling")
    storage = _labelled(l7, r"storage")
    if not handling and not storage:
        handling = "\n".join(l7)

    sds_no = (intake.get("sds_number") or "").strip()
    ver = str(intake.get("version") or "").strip()
    row = {
        "Vertical": vertical,
        "Product Folder": "",
        "File Name": file_name,
        "Is Latest Version": "Yes",
        "File Version": f"V{ver}" if ver else "",
        "File Date": date.today().strftime("%m/%d/%Y"),
        "SDS #": sds_no,
        "Header Date": core.normalize_date_string(intake.get("date_of_issue") or ""),
        "Legacy SDS No. / Layout": "",
        "Product Name (SDS)": (intake.get("product_name") or intake.get("trade_name") or "").strip(),
        "Product Code (SDS)": (intake.get("product_code") or "").strip(),
        "Signal Word": _labelled(l2, r"signal\s*word"),
        "Classification": _labelled(l2, r"classification"),
        "Pictograms": _pictograms(s2, haz),
        "Hazard Statements": "\n".join(haz),
        "Precautionary Statements": "\n".join(prec),
        "Ingredients (Name | CAS | %)": _ingredients(s3),
        "Handling": handling,
        "Storage": storage,
        "Disposal": "\n".join(_lines(s13.blocks)) if s13 else "",
        "Parse Flags": "",
        "Notes": "",
    }
    assert list(row) == COLUMNS
    return row


def save_sds_draft(row: dict):
    """Call the shared helper. Returns (ok: bool, message: str). Never raises."""
    try:
        from tools import master_repo
    except Exception:
        return False, ("Draft saving isn't set up yet: tools/master_repo.py "
                       "is missing from the app.")
    try:
        result = master_repo.save_draft(TAB, row)
    except Exception as e:  # missing token, repo unreachable, unknown column...
        return False, f"Draft not saved: {e}"
    if isinstance(result, dict):
        status = (result.get("action") or result.get("status") or "saved")
    else:
        status = str(result) if result else "saved"
    return True, f"Draft {status} - {row['File Name']}"


def save_and_report(st, sds, intake: dict, *, vertical: str, file_name: str,
                    sds_number: str | None = None):
    """Build the English row, save it as a draft, and show the result.

    Called right after a successful generate, so it runs on every generation
    without needing a second button click (a nested button would vanish on
    Streamlit's rerun). Never raises and never blocks the downloads.
    """
    try:
        row_intake = dict(intake)
        if sds_number:
            row_intake["sds_number"] = sds_number
        # The key is the final PDF name (that is what gets saved as final),
        # so the draft written now is later replaced, not duplicated.
        key_name = re.sub(r"\.docx$", ".pdf", file_name, flags=re.I)
        row = build_sds_row(sds, row_intake, vertical=vertical, file_name=key_name)
        ok, msg = save_sds_draft(row)
    except Exception as e:  # noqa: BLE001 - must never break generation
        st.warning(f"Product data draft not saved: {e}")
        return
    (st.info if ok else st.warning)(msg)
