"""Shared helper for saving rows to the Forza product-data repo.

Every app (TDS Builder, SDS Formatter, Label Builder) copies this one file.

Rules (see README.md):
  * Apps only ever write DRAFT rows (drafts/<tab>.csv). Nothing reaches data/
    until the owner approves it.
  * Each app writes only its own tab.
  * Read access to data/ (the approved master) is always fine.

Setup in a Streamlit app (Settings > Secrets):
    GITHUB_TOKEN = "github_pat_..."        # fine-grained token, THIS repo only,
                                           # Contents: Read and write
    MASTER_REPO  = "mguzallis/forza-product-data"   # change if the repo moves
    MASTER_BRANCH = "main"
"""
import base64, csv, io, os

try:
    import requests
except ImportError:  # pragma: no cover
    requests = None

TABS = {
    "tds": "TDS", "sds": "SDS", "product_sds_map": "Product-SDS Map",
    "sds_ghs_status": "SDS GHS Status", "labels": "Labels",
}
# Primary key column(s) per tab (a row with the same key replaces the old one).
KEYS = {
    "tds": ["Product Code"],
    "sds": ["File Name"],            # becomes ["SDS #"] once all SDSs are on the current template
    "product_sds_map": ["Product Code"],
    "sds_ghs_status": ["File Name"],
    "labels": ["Product Code", "Size / Container", "Color"],
}
# Which app owns which tab (apps must not write another app's tab).
OWNER = {"tds": "TDS Builder", "sds": "SDS Formatter", "labels": "Label Builder",
         "product_sds_map": "Product Marketing", "sds_ghs_status": "Report"}


def _secret(name, default=None):
    try:
        import streamlit as st
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.environ.get(name, default)


class GitHubBackend:
    def __init__(self):
        self.token = _secret("GITHUB_TOKEN")
        self.repo = _secret("MASTER_REPO", "mguzallis/forza-product-data")
        self.branch = _secret("MASTER_BRANCH", "main")
        if not self.token:
            raise RuntimeError("GITHUB_TOKEN is not set")
        self.h = {"Authorization": f"Bearer {self.token}",
                  "X-GitHub-Api-Version": "2022-11-28"}

    def _url(self, path):
        return f"https://api.github.com/repos/{self.repo}/contents/{path}"

    def read(self, path):
        meta = requests.get(self._url(path), headers={**self.h, "Accept": "application/vnd.github+json"},
                            params={"ref": self.branch}, timeout=30)
        meta.raise_for_status()
        sha = meta.json()["sha"]
        raw = requests.get(self._url(path), headers={**self.h, "Accept": "application/vnd.github.raw+json"},
                           params={"ref": self.branch}, timeout=60)
        raw.raise_for_status()
        return raw.content.decode("utf-8-sig"), sha

    def write(self, path, text, sha, message):
        r = requests.put(self._url(path), headers=self.h, timeout=60, json={
            "message": message, "branch": self.branch, "sha": sha,
            "content": base64.b64encode(text.encode("utf-8")).decode()})
        r.raise_for_status()


def _parse(text):
    rd = csv.DictReader(io.StringIO(text))
    return list(rd), rd.fieldnames or []


def _dump(rows, header):
    out = io.StringIO()
    w = csv.DictWriter(out, fieldnames=header, lineterminator="\n", extrasaction="ignore")
    w.writeheader()
    w.writerows(rows)
    return out.getvalue()


def _same(a, b, keys):
    return all((a.get(k) or "").strip() == (b.get(k) or "").strip() for k in keys)


def read_master(tab, backend=None):
    """Approved rows (data/<tab>.csv) as a list of dicts."""
    b = backend or GitHubBackend()
    text, _ = b.read(f"data/{tab}.csv")
    return _parse(text)[0]


def save_draft(tab, row, backend=None, who="app"):
    """Add or replace one DRAFT row in drafts/<tab>.csv. Returns 'added' or 'replaced'."""
    if tab not in KEYS:
        raise ValueError(f"unknown tab {tab!r}")
    keys = KEYS[tab]
    missing = [k for k in keys if not (row.get(k) or "").strip()]
    if missing:
        raise ValueError(f"row is missing key column(s): {missing}")
    b = backend or GitHubBackend()
    path = f"drafts/{tab}.csv"
    text, sha = b.read(path)
    rows, header = _parse(text)
    if not header:  # empty draft file: take the header from the approved master
        header = _parse(b.read(f"data/{tab}.csv")[0])[1]
    unknown = [c for c in row if c not in header]
    if unknown:
        raise ValueError(f"columns not in {tab} spec: {unknown}")
    status = "added"
    for i, r in enumerate(rows):
        if _same(r, row, keys):
            rows[i] = {**r, **row}
            status = "replaced"
            break
    else:
        # Start from the approved row (if any) so fields the app did not
        # touch are carried over instead of being blanked on approval.
        base = next((m for m in read_master(tab, b) if _same(m, row, keys)), {})
        rows.append({**base, **row})
    b.write(path, _dump(rows, header), sha,
            f"Draft {status}: {tab} / {' | '.join(row.get(k, '') for k in keys)} ({who})")
    return status


def approve(tab, key_values, backend=None):
    """OWNER ONLY. Move the draft row(s) with these key values into data/<tab>.csv."""
    keys = KEYS[tab]
    wanted = dict(zip(keys, key_values))
    b = backend or GitHubBackend()
    dtext, dsha = b.read(f"drafts/{tab}.csv")
    drows, dheader = _parse(dtext)
    mtext, msha = b.read(f"data/{tab}.csv")
    mrows, mheader = _parse(mtext)
    picked = [r for r in drows if _same(r, wanted, keys)]
    if not picked:
        raise ValueError("no matching draft row")
    for p in picked:
        for i, m in enumerate(mrows):
            if _same(m, p, keys):
                mrows[i] = p          # the draft row is the full new row
                break
        else:
            mrows.append(p)
    rest = [r for r in drows if not _same(r, wanted, keys)]
    b.write(f"data/{tab}.csv", _dump(mrows, mheader), msha,
            f"Approved: {tab} / {' | '.join(key_values)}")
    _, dsha2 = b.read(f"drafts/{tab}.csv")
    b.write(f"drafts/{tab}.csv", _dump(rest, dheader), dsha2, f"Cleared approved draft: {tab}")
    return len(picked)
