# Forza SDS Formatter

Formats or rebrands a Safety Data Sheet onto Forza's branded per-vertical templates.

## Run

```
pip install -r requirements.txt
streamlit run app.py
```

## Layout

```
forza_sds/
  app.py                 Streamlit UI
  core.py                extractor + renderer engine
  shells/                six branded templates, one per vertical
  build_shells.py        regenerates the shells from source artwork
```

## What it does

1. Reads any SDS .docx into a canonical 16-section model
2. Applies the intake answers to the header block and Section 1
3. Stamps Section 16 revision date and version when those fields exist
4. Renders into the chosen vertical shell

## Formatting rules applied

- Times New Roman throughout: title 22pt bold, section headings 18pt bold
  thick-underlined, body 12pt
- Section headings normalised to `SECTION n – Title` with an en dash
- All tables (the top block and every data table) use double-line borders,
  sz 4, matching the T-OS164 reference exactly. Header row repeats across
  pages, rows do not split.
- Hard page breaks are stripped and never regenerated. Instead the renderer
  applies keep-with-next so Word reflows units intact:
    - a label line (anything ending in a colon) stays with what follows it
    - the line before a table stays with the table
    - GHS pictograms never separate from each other or from their label
    - lists of two or three items are never split across a page
  A bound run is capped at five paragraphs so a whole page can never shift.
  Longer lists and large tables still rely on Word's own widow control; add a
  manual break in the rare case one lands badly.
- US Letter, margins top 1.45" / bottom 0.85" / sides 1.00"

## Pre-flight check

Runs automatically in the app after Section 16 has been stamped, so it
reports what will actually survive into the final document rather than
raw-source noise that's about to get fixed anyway. Checks:

- Missing sections (1-16)
- Sections that parsed with next to no content (heading found, body didn't)
- Placeholder text: bracketed text (`[insert here]`) and bare `TBD` /
  `PLACEHOLDER` / `FIXME` / `XXXX`
- Whether Section 16 has any field that looks like a version and any field
  that looks like a date - if neither exists, the version/date silently
  never appears there, which is worth knowing before the file goes out
- Recommended Use / Recommended Restrictions / Other Means / Product Code
  left blank

Nothing here blocks generation - it's a pre-flight check, not a validator -
but every one of these was a real, silent failure mode found while testing
against actual engineering drafts, so worth a glance every time.

## Hazard pictograms are always real images

Section 2 often lists pictograms as text - `GHS08 (Health Hazard)`, or a
table row `Pictograms | GHS02 (Flame), GHS04 (Gas Cylinder)`. Wherever that
text-only form appears, it's replaced with the actual GHS pictogram
graphics (bundled in `pictograms/`, nine standard red-diamond hazard
symbols). Detection matches the `GHS0X` code first since that's
unambiguous; a keyword fallback (`Health Hazard`, `Skull and Crossbones`,
etc.) only runs if no code is present, and only within whatever block is
already labeled "Pictograms" - the same words appear constantly elsewhere
in Section 2's hazard-classification prose and must never be touched there.

A block that already contains real embedded images (both Forza reference
docs already do) is left alone entirely - this only fires on genuinely
text-only pictogram listings.

## Batch modes

The app has three modes, selected at the top:

**Single file** - the original flow: one source, one vertical.

**Rebrand across verticals** - one formula, several verticals in one run.
Upload the base SDS once, pick the verticals, and fill in a per-vertical
form: Trade Name, SDS #, Product Code, Other Means, Recommended Use,
Recommended Restrictions. Nothing here is auto-derived from another
vertical's values or from a naive prefix-swap - a real rebrand gets a
genuinely distinct identity per vertical (different product code, different
positioning), not a relabeled SKU, so each vertical's identifiers are typed
in by hand. The underlying hazard data (Sections 2-15) is shared byte-for-
byte across every vertical generated, via an independent `copy.deepcopy` of
the parsed source per vertical - stamping one vertical's version/date can
never bleed into another's copy.

**Bulk reformat** - several different source SDSs, all onto the same
vertical shell. Each file gets its own prefilled, editable intake form and
its own pre-flight check; one file's fields never affect another's.

## Footer spacing

`Page N of M` and `DCN: ...` sit directly stacked with no gap between them -
two independently positioned frames, each 3" wide with right-justified
text so long trade names never wrap, offset by only a single line height.

Both batch modes produce a `.zip` with one `.docx` per job.

## Date format is always slashes

Any date fed in with dots (`08.18.2026`) is normalized to slashes
(`08/18/2026`) wherever it lands - the top table, Section 16, and the
Replaces field. Non-date text in the same field (`N/A`, `Revision 6`) is
left alone; only digit.digit.digit patterns that look like a date get
converted. This closes the original mismatch in the reference document,
which used slashes in the top table but dots in Section 16.

## Section 1 is always the fixed Forza format

Section 1 is never carried over from the source document, no matter what
structure or wording it used there. It is fully replaced at render time by a
canonical block matching the T-OS164 reference exactly: same five fields,
same wording, same spacing, plus the constant supplier block (company
address, phone, emergency contact) that Forza always uses.

Five values populate that block:
`product_name`, `other_means`, `product_code`, `recommended_use`,
`recommended_restrictions`. The first three come from the intake form. The
last two are auto-detected from the source (checking bold-label paragraphs
and table rows, tolerant of label variants like `Product` / `Product Name` /
`Product identifier`, or `Recommended Restrictions` / `Uses advised
against`) and shown as editable fields in the UI, since wording for these
two genuinely varies by product and is worth a human glance before it goes
out. A field genuinely absent from the source (TAC850 has no "Other Means"
row at all) is simply left blank rather than guessed at.

The supplier/company/phone block is always the fixed boilerplate text -
never pulled from the source - since engineering drafts often carry
placeholder brackets there (`[Company name and address - insert]`).

Section 16's "Prepared by" field is handled the same way: always forced to
`Forza, Inc.` regardless of what the source says, since this was also seen
holding a placeholder (`[PREPARER / DEPARTMENT]`) in a real draft. If the
field is missing from the source entirely (TAC850 has no such line at all),
it is inserted - directly after the revision/date line(s) if any were
found, otherwise at the top of the section - so it is always present, not
just corrected when it happens to exist.

## Section 16 field injection

Section 16 is not replaced wholesale like Section 1 - it keeps whatever
structure the source used (paragraph or table) and only the matching
revision/date fields are patched in place. Label routing: `Revision` ->
version, `Issue Date` -> date, `Revision date` -> date only despite
containing the word "revision", and only an explicit joiner (`Revision /
Issue Date`) is treated as wanting a combined version+date value.

## Numbering

Top table:  `S-0196 V1`   (hyphen kept, space before V)
Footer DCN: `S0196_V1 – R-OS86`   (hyphen dropped, en dash with spaces)

Both forms are taken from the reference document. Note that older SDSs in
circulation use at least two other DCN forms.

## Banner color fidelity

The five shells that use uploaded artwork unchanged (Composites, Construction,
Insulation, Marine, Transportation) embed the **original CMYK JPEG bytes
verbatim** - no RGB conversion, no re-encode. A naive CMYK->RGB conversion
(PIL's default `.convert("RGB")`, and several ImageMagick rendering intents
tested) measurably desaturates these files because it ignores the embedded
ICC profile. Byte-for-byte passthrough is the only approach proven to match.

Industrial now also uses a true Letter-sized source (`Industrial_reredo.docx`)
and is a byte-for-byte passthrough like the other five, no reflow needed. An
earlier revision required an A4 -> Letter reflow done in CMYK; that path is
gone now that a real Letter-sized asset exists, but if a future banner ever
needs a similar edit, keep it in CMYK end to end - do not let the image pass
through RGB at any point.

## Footer

Two independently positioned frames, both 10pt Times New Roman, right justified:

- `Page N of M` at 10.35" from the page top, level with the artwork tagline
- `DCN: ...` at 10.78", clear below the artwork address line

Each frame is 3" wide with right-justified text, so the DCN can never wrap to a
second line no matter how long the trade name is. The unused left portion of the
frame is transparent.

## After generating

Open in Word, select all and press F9 to refresh the page-count field.


## Corporate branding

The app UI follows Forza's Nov 2025 brand standards: Regal Blue (#1B3764)
and Blaze Orange (#F16022) primary colors, the real Kallisto Heavy font for
headings (embedded in `app.py` as a base64 data URI from
`assets/fonts/kallisto-heavy.woff2`, so it loads with no extra static-file
setup), and Poppins Regular for body text. The header is the corporate
logo lockup (eagle + "Forza" wordmark + tagline) on a plain white
background - a raster image, not retyped text, since the brand guide
explicitly prohibits recreating the logo in any font. Colors live in
`.streamlit/config.toml` (Streamlit's native theme, so buttons, focus
outlines, and selections all pick it up automatically) plus a small CSS
block in `app.py` for headings and the tricolor "Purpose-Built /
Performance / Guaranteed Strength" slogan mark under the mode selector.
