#!/usr/bin/env python3
"""Forza SDS Formatter / Rebrander - Streamlit front end.

Run:  streamlit run app.py
"""
import base64
import copy
import io
import os
import re
import tempfile
import zipfile
from datetime import date

import streamlit as st

from core import (VERTICALS, extract, apply_intake, render, preflight,
                  build_dcn, build_sds_field)

SHELL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shells")
ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
TODAY = date.today().strftime("%m/%d/%Y")
DOCX_MIME = ("application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document")

st.set_page_config(page_title="Forza SDS Formatter",
                   page_icon=os.path.join(ASSETS_DIR, "forza_favicon.png"),
                   layout="centered")

# Forza brand standards (Nov 2025): Regal Blue #1B3764 / Blaze Orange #F16022
# primary, Blue Velvet #09668D / Rusty Nail Orange #D35127 / Slate Grey
# #BFBFBF secondary. Kallisto Heavy is the real brand heading font - embedded
# below as a data URI so it loads with no extra static-file setup; Poppins
# Regular is the documented body font.
with open(os.path.join(ASSETS_DIR, "fonts", "kallisto-heavy.woff2"), "rb") as _fh:
    _KALLISTO_B64 = base64.b64encode(_fh.read()).decode("ascii")

st.markdown(f"""
<style>
@font-face {{
    font-family: 'Kallisto Heavy';
    src: url(data:font/woff2;base64,{_KALLISTO_B64}) format('woff2');
    font-weight: 400 900;
    font-style: normal;
}}
@import url('https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,400;0,500;0,600;0,700;0,800;1,600&display=swap');

html, body, [class*="css"], .stMarkdown, .stText, p, span, label, div {{
    font-family: 'Poppins', sans-serif;
}}
h1, h2, h3, .stRadio > label {{
    font-family: 'Kallisto Heavy', 'Poppins', sans-serif;
    color: #1B3764;
    letter-spacing: 0.01em;
}}
h1 {{ text-transform: uppercase; }}
.stButton > button[kind="primary"] {{
    font-family: 'Kallisto Heavy', 'Poppins', sans-serif;
    letter-spacing: 0.03em;
    text-transform: uppercase;
}}
.stAlert {{ border-radius: 6px; }}
</style>
""", unsafe_allow_html=True)

with open(os.path.join(ASSETS_DIR, "forza_logo_corporate.png"), "rb") as _fh:
    _LOGO_B64 = base64.b64encode(_fh.read()).decode("ascii")

st.markdown(
    f"""<div style="text-align:center;">
    <img src="data:image/png;base64,{_LOGO_B64}" width="420">
    </div>""",
    unsafe_allow_html=True)

st.markdown(
    """<div style="font-family:'Kallisto Heavy','Poppins',sans-serif;
    letter-spacing:0.1em; color:#1B3764;
    font-size:0.85rem; text-transform:uppercase; margin:14px 0 18px 0;">
    SDS Formatter &amp; Rebrand Tool</div>""",
    unsafe_allow_html=True)

st.caption("Drop in an R&D or Marketing SDS. Get back a branded, "
           "correctly formatted document.")

mode = st.radio(
    "Mode",
    ["Single file", "Rebrand across verticals", "Bulk reformat"],
    horizontal=True,
    help=("Single file: one source, one vertical.\n\n"
          "Rebrand across verticals: one formula, several verticals, each "
          "with its own Trade Name / SDS # / Product Code that you enter - "
          "never auto-derived, since a rebrand is not just a prefix swap.\n\n"
          "Bulk reformat: several different source SDSs, all onto the same "
          "vertical shell."))
st.divider()

st.markdown(
    """<div style="font-family:'Kallisto Heavy','Poppins',sans-serif;
    letter-spacing:0.04em; font-size:0.78rem; text-transform:uppercase;
    margin:-10px 0 16px 0; text-align:center;">
    <span style="color:#1B3764;">Purpose-Built </span>
    <span style="color:#BFBFBF;">Performance </span>
    <span style="color:#F16022;">Guaranteed Strength</span>
    </div>""",
    unsafe_allow_html=True)


def intake_fields(key_prefix: str, prefill: dict, section1_caption=True) -> dict:
    """Render the standard 11-field intake form and return the collected dict."""
    c1, c2 = st.columns(2)
    with c1:
        trade_name = st.text_input("Trade Name", prefill.get("trade_name", ""),
                                   key=f"{key_prefix}_tn")
        sds_number = st.text_input("SDS #", prefill.get("sds_number", ""),
                                   help="Format S-XXXX, e.g. S-0196", key=f"{key_prefix}_sds")
        version = st.text_input("Version Number", prefill.get("version", "1"),
                                key=f"{key_prefix}_ver")
    with c2:
        date_of_issue = st.text_input("Date of Issue",
                                      prefill.get("date_of_issue", TODAY),
                                      key=f"{key_prefix}_doi")
        replaces = st.text_input("Replaces (Date / Revision #)", prefill.get("replaces", ""),
                                 help="Freeform - whatever the prior SDS was labeled. "
                                      "E.g. 08/05/2026 / V1, Revision 6, New, or N/A.",
                                 key=f"{key_prefix}_rep")
        effective_date = st.text_input("Effective Date",
                                       prefill.get("effective_date", TODAY),
                                       key=f"{key_prefix}_eff")

    st.markdown("**Section 1**")
    product_name = st.text_input("Product Name", prefill.get("product_name", ""),
                                 key=f"{key_prefix}_pn")
    other_means = st.text_input("Other Means of Identification",
                                prefill.get("other_means", ""), key=f"{key_prefix}_om")
    product_code = st.text_input("Product Code Number", prefill.get("product_code", ""),
                                 key=f"{key_prefix}_pc")
    recommended_use = st.text_input("Recommended Use", prefill.get("recommended_use", ""),
                                    key=f"{key_prefix}_ru")
    recommended_restrictions = st.text_input(
        "Recommended Restrictions", prefill.get("recommended_restrictions", ""),
        key=f"{key_prefix}_rr")
    if section1_caption:
        st.caption("Section 1 always renders in Forza's fixed format. The supplier "
                  "block (company, phone, emergency contact) is constant and is "
                  "never taken from the source document.")

    ver_digits = re.sub(r"\D", "", version) or "1"
    return dict(trade_name=trade_name.strip(), sds_number=sds_number.strip(),
                version=ver_digits, replaces=replaces.strip(),
                date_of_issue=date_of_issue.strip(), effective_date=effective_date.strip(),
                product_name=product_name.strip(), other_means=other_means.strip(),
                product_code=product_code.strip(), recommended_use=recommended_use.strip(),
                recommended_restrictions=recommended_restrictions.strip())


def validate(intake: dict) -> list:
    problems = []
    if not intake["trade_name"]:
        problems.append("Trade Name is required.")
    if not re.match(r"^S-?\d{3,5}$", intake["sds_number"] or ""):
        problems.append("SDS # should look like S-0196.")
    return problems


def show_preflight(sds, intake, key_prefix=""):
    apply_intake(sds, intake)
    report = preflight(sds, intake)
    if not report["warnings"]:
        st.success("No issues found.")
    else:
        for w in report["warnings"]:
            st.warning(w)
        if report["placeholders"]:
            with st.expander(f"Placeholder detail ({len(report['placeholders'])})"):
                for num, title, snippet in report["placeholders"]:
                    st.text(f"Section {num} ({title}): ...{snippet}...")
    return report


def parse_upload(upload):
    raw = upload.read()
    try:
        return extract(raw), None
    except Exception as exc:                                # noqa: BLE001
        return None, str(exc)


# =========================================================== SINGLE FILE ===
if mode == "Single file":
    upload = st.file_uploader("Source SDS (.docx)", type=["docx"])
    if not upload:
        st.info("Upload an SDS to begin.")
        st.stop()

    sds, err = parse_upload(upload)
    if err:
        st.error(f"Could not read that document: {err}")
        st.stop()

    found = sorted(s.number for s in sds.sections)
    if len(found) == 16:
        st.success(f"Parsed all 16 sections, {len(sds.media)} embedded image(s).")
    else:
        missing = [n for n in range(1, 17) if n not in found]
        st.warning(f"Found sections {found}. Missing: {missing or 'none'}.")

    vertical = st.selectbox("Industry", VERTICALS, index=VERTICALS.index("Marine"))
    shell = os.path.join(SHELL_DIR, f"SDS_Shell_{vertical}.docx")

    intake = intake_fields("single", sds.detected)

    st.divider()
    p1, p2 = st.columns(2)
    p1.metric("SDS # field", build_sds_field(intake["sds_number"], intake["version"])
             if intake["sds_number"] else "-")
    p2.metric("Footer DCN", build_dcn(intake["sds_number"], intake["version"], intake["trade_name"])
             if intake["sds_number"] and intake["trade_name"] else "-")

    problems = validate(intake)
    for p in problems:
        st.warning(p)

    st.subheader("Pre-flight check")
    show_preflight(sds, intake)

    if st.button("Generate SDS", type="primary", disabled=bool(problems)):
        safe = re.sub(r"[^0-9A-Za-z_.-]", "_", intake["trade_name"])
        fname = f"Forza_{vertical}_{safe}_SDS_V{intake['version']}.docx"
        out = os.path.join(tempfile.mkdtemp(), fname)
        render(sds, intake, shell, out)
        with open(out, "rb") as fh:
            st.download_button("Download", fh.read(), file_name=fname, mime=DOCX_MIME)
        st.success("Done. Open in Word and press Ctrl+A then F9 to refresh page totals.")
        st.caption("Hard page breaks are intentionally not carried over.")

# =============================================== REBRAND ACROSS VERTICALS ===
elif mode == "Rebrand across verticals":
    st.caption("One formula, several verticals. Each vertical keeps the same "
              "hazard data (Sections 2-15) but gets its own identity - type "
              "the real Trade Name, SDS #, and Product Code for each; "
              "nothing here is auto-derived from another vertical's values.")
    upload = st.file_uploader("Source SDS - the base formula (.docx)", type=["docx"],
                              key="A_upload")
    if not upload:
        st.info("Upload the base SDS to rebrand across verticals.")
        st.stop()

    base_sds, err = parse_upload(upload)
    if err:
        st.error(f"Could not read that document: {err}")
        st.stop()

    found = sorted(s.number for s in base_sds.sections)
    st.success(f"Parsed {len(found)}/16 sections, {len(base_sds.media)} embedded image(s).")

    chosen = st.multiselect("Verticals to generate", VERTICALS)
    if not chosen:
        st.info("Pick at least one vertical.")
        st.stop()

    d = base_sds.detected
    shared_defaults = dict(date_of_issue=d.get("date_of_issue", TODAY),
                           effective_date=d.get("effective_date", TODAY),
                           recommended_use=d.get("recommended_use", ""),
                           recommended_restrictions=d.get("recommended_restrictions", ""))

    jobs, problems = {}, {}
    for v in chosen:
        with st.expander(f"{v}", expanded=True):
            intake = intake_fields(f"A_{v}", shared_defaults, section1_caption=False)
            jobs[v] = intake
            p = validate(intake)
            if p:
                problems[v] = p
                for msg in p:
                    st.warning(msg)

    st.subheader("Pre-flight check")
    for v, intake in jobs.items():
        st.markdown(f"**{v}**")
        show_preflight(copy.deepcopy(base_sds), intake)

    if st.button("Generate all verticals", type="primary", disabled=bool(problems)):
        buf = io.BytesIO()
        names = []
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zip_out:
            for v, intake in jobs.items():
                sds_copy = copy.deepcopy(base_sds)
                apply_intake(sds_copy, intake)
                safe = re.sub(r"[^0-9A-Za-z_.-]", "_", intake["trade_name"])
                fname = f"Forza_{v}_{safe}_SDS_V{intake['version']}.docx"
                out_path = os.path.join(tempfile.mkdtemp(), fname)
                render(sds_copy, intake, os.path.join(SHELL_DIR, f"SDS_Shell_{v}.docx"), out_path)
                zip_out.write(out_path, arcname=fname)
                names.append(fname)
        st.download_button("Download all (.zip)", buf.getvalue(),
                           file_name="Forza_SDS_multivertical.zip", mime="application/zip")
        st.success(f"Generated {len(names)} document(s): " + ", ".join(names))

# ============================================================ BULK REFORMAT ===
elif mode == "Bulk reformat":
    st.caption("Several different source SDSs, reformatted onto the same "
              "vertical. Each keeps its own identity - review the prefilled "
              "fields per file before generating.")
    uploads = st.file_uploader("Source SDS files (.docx)", type=["docx"],
                               accept_multiple_files=True, key="B_uploads")
    if not uploads:
        st.info("Upload two or more SDS files to reformat onto one vertical.")
        st.stop()

    vertical = st.selectbox("Industry", VERTICALS, index=VERTICALS.index("Marine"),
                            key="B_vertical")
    shell = os.path.join(SHELL_DIR, f"SDS_Shell_{vertical}.docx")

    parsed, jobs, problems = {}, {}, {}
    for f in uploads:
        s, err = parse_upload(f)
        if err:
            st.error(f"{f.name}: could not read ({err})")
            continue
        parsed[f.name] = s
        found = sorted(x.number for x in s.sections)
        with st.expander(f.name, expanded=False):
            st.caption(f"{len(found)}/16 sections parsed, {len(s.media)} embedded image(s).")
            intake = intake_fields(f"B_{f.name}", s.detected, section1_caption=False)
            jobs[f.name] = intake
            p = validate(intake)
            if p:
                problems[f.name] = p
                for msg in p:
                    st.warning(msg)

    st.subheader("Pre-flight check")
    for name, intake in jobs.items():
        st.markdown(f"**{name}**")
        show_preflight(copy.deepcopy(parsed[name]), intake)

    if st.button("Generate all files", type="primary", disabled=bool(problems)):
        buf = io.BytesIO()
        names = []
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zip_out:
            for fname_in, intake in jobs.items():
                s = parsed[fname_in]
                apply_intake(s, intake)
                safe = re.sub(r"[^0-9A-Za-z_.-]", "_", intake["trade_name"] or fname_in)
                out_name = f"Forza_{vertical}_{safe}_SDS_V{intake['version']}.docx"
                out_path = os.path.join(tempfile.mkdtemp(), out_name)
                render(s, intake, shell, out_path)
                zip_out.write(out_path, arcname=out_name)
                names.append(out_name)
        st.download_button("Download all (.zip)", buf.getvalue(),
                           file_name=f"Forza_{vertical}_batch.zip", mime="application/zip")
        st.success(f"Generated {len(names)} document(s): " + ", ".join(names))
