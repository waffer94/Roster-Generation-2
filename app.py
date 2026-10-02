"""
Streamlit web UI for F.A.S.T. Rescue course tools.

Tab 1: Course Roster Builder  - Excel export -> Word rosters
Tab 2: Sign-Off Sheet Generator - Excel export -> TLMS sign-off Excel sheets

Run locally:   streamlit run app.py
Deploy free:   push repo to GitHub, deploy at https://share.streamlit.io
"""

import base64
import io
import os
import tempfile
import zipfile

import streamlit as st
import build_rosters as br
import build_signoff as bs

TEMPLATE_DIR = "templates"
MAPPING_PATH = bs.DEFAULT_MAPPING
LOGO_PATH = "assets/logo.png"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="F.A.S.T. Rescue - Training Tools",
                   page_icon="\U0001F691", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', system-ui, sans-serif; }
.block-container { padding-top: 1.5rem; max-width: 860px; }
#MainMenu, footer { visibility: hidden; }
.brand-bar {
    display: flex; align-items: center; gap: 16px;
    background: linear-gradient(135deg, #1e1e1e 0%, #3a3a3a 100%);
    border-radius: 14px; padding: 18px 22px; margin-bottom: 6px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.10);
}
.brand-bar img { width: 54px; height: 54px; border-radius: 50%; background:#fff; padding:2px; }
.brand-bar .t { color: #fff; font-size: 1.4rem; font-weight: 700; }
.brand-bar .s { color: #ccc; font-size: 0.85rem; margin-top: 2px; }
.section-head {
    font-size: 1.15rem; font-weight: 700; color: #1a1a1a;
    margin: 20px 0 4px; padding-bottom: 6px;
    border-bottom: 3px solid #E03127; display: inline-block;
}
.hint { color:#6b7280; font-size:0.88rem; margin: 0 0 16px; }
div[data-testid="stFileUploader"] {
    border: 2px dashed #d6431f44; border-radius: 12px;
    padding: 8px 12px; background: #fff8f7;
}
.stButton > button, .stDownloadButton > button {
    border-radius: 10px; font-weight: 600; padding: 0.45rem 1rem;
}
.stDownloadButton > button { border: 1px solid #E0312722; }
.card {
    border: 1px solid #e8e8e8; border-left: 4px solid #E03127;
    border-radius: 10px; padding: 12px 16px; margin-bottom: 14px;
    background: #fff; box-shadow: 0 1px 6px rgba(0,0,0,0.03);
}
.card.skip { border-left-color: #d4a017; }
.foot { color:#9aa0a6; font-size:0.78rem; text-align:center; margin-top:24px; }
div[data-testid="stTabs"] button[data-baseweb="tab"] {
    font-weight: 600; font-size: 0.95rem;
}
</style>
""", unsafe_allow_html=True)

_pw = st.secrets.get("password", None) if hasattr(st, "secrets") else None
if _pw and st.session_state.get("authed") is not True:
    st.markdown("#### \U0001F512 Password required")
    entered = st.text_input("Password", type="password", label_visibility="collapsed",
                            placeholder="Enter password")
    if entered == _pw:
        st.session_state["authed"] = True
        st.rerun()
    elif entered:
        st.error("Incorrect password.")
    st.stop()

logo_uri = ""
if os.path.exists(LOGO_PATH):
    logo_uri = "data:image/png;base64," + base64.b64encode(
        open(LOGO_PATH, "rb").read()).decode()

st.markdown(f"""
<div class="brand-bar">
  {'<img src="' + logo_uri + '"/>' if logo_uri else ''}
  <div>
    <div class="t">F.A.S.T. Rescue <span style="color:#E03127;">Training Tools</span></div>
    <div class="s">Course Roster Builder &nbsp;&middot;&nbsp; Sign-Off Sheet Generator</div>
  </div>
</div>
""", unsafe_allow_html=True)

tab_roster, tab_signoff = st.tabs(["\U0001F4CB  Course Rosters", "\u2705  Sign-Off Sheets"])


def _render_results(res, dl_key, zip_name, mime, label):
    st.divider()
    m1, m2 = st.columns(2)
    m1.metric("Created", res["made"]); m2.metric("Skipped", res["skip"])
    if res["files"]:
        st.download_button(f"\u2B07  Download all {label} (.zip)", res["zip"],
                           file_name=zip_name, mime="application/zip",
                           type="primary", use_container_width=True, key=dl_key)
    for s, t, d in res["log"]:
        cls = "card" if s == "ok" else "card skip"
        icon = "\u2705" if s == "ok" else "\u23ED\uFE0F"
        pre = "" if s == "ok" else "Skipped \u2014 "
        st.markdown(f'<div class="{cls}">{icon} <b>{pre}{t}</b><br>'
                    f'<span style="color:#6b7280">{d}</span></div>',
                    unsafe_allow_html=True)
    if res["files"]:
        with st.expander("Download individual files"):
            for i, (nm, d) in enumerate(res["files"]):
                st.download_button(nm, d, file_name=nm, mime=mime,
                                   key=f"{dl_key}_{i}", use_container_width=True)


with tab_roster:
    st.markdown('<div class="section-head">Course Roster Builder</div>', unsafe_allow_html=True)
    st.markdown('<div class="hint">Upload the Excel export \u2192 download print-ready Word rosters.</div>',
                unsafe_allow_html=True)
    with st.expander("\u2139\uFE0F  How it works"):
        st.markdown(
            "1. Export your courses to Excel.\n"
            "2. Upload the `.xlsx` below.\n"
            "3. Click **Generate** and download the Word files.\n\n"
            "The right template is chosen automatically from the course name "
            "(First Aid In-Class / Blended / Recertification, Working at Heights, "
            "Lift Truck, EWP, or General for anything else). "
            "Cancelled courses are skipped unless you opt in."
        )
    r_upload = st.file_uploader("Excel export", type=["xlsx"], key="r_upload")
    rc1, rc2 = st.columns(2)
    r_keep = rc1.toggle("Keep export order", help="Otherwise names are sorted A\u2013Z", key="r_keep")
    r_cancelled = rc2.toggle("Include cancelled", key="r_cancelled")
    r_go = st.button("Generate rosters", type="primary", use_container_width=True,
                     disabled=r_upload is None, key="r_go")
    if r_upload and r_go:
        tmap = br.load_template_map(TEMPLATE_DIR)
        if not tmap:
            st.error(f"No templates in `{TEMPLATE_DIR}/`.")
            st.stop()
        r_log, r_files = [], []
        with st.spinner("Building rosters \u2026"):
            with tempfile.TemporaryDirectory() as tmp:
                xp = os.path.join(tmp, "e.xlsx")
                with open(xp, "wb") as f: f.write(r_upload.getbuffer())
                od = os.path.join(tmp, "out"); os.makedirs(od)
                rows = br.read_export(xp)
                sfx = br._assign_suffixes(rows)
                made = 0
                for i, rec in enumerate(rows):
                    s, info, ct, n = br.build_one(rec, tmap, od,
                        sort_alpha=not r_keep, include_cancelled=r_cancelled,
                        fname_suffix=sfx[i])
                    if s == "ok":
                        made += 1
                        with open(info, "rb") as fh:
                            r_files.append((os.path.basename(info), fh.read()))
                        r_log.append(("ok", os.path.basename(info), f"{ct} \u00b7 {n} participants"))
                    else:
                        r_log.append(("skip", rec.get("customer") or "\u2014", info))
            zbuf = io.BytesIO()
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
                for nm, d in r_files: z.writestr(nm, d)
        st.session_state["r_res"] = {
            "files": r_files, "zip": zbuf.getvalue(), "log": r_log,
            "made": made, "skip": sum(1 for s, *_ in r_log if s == "skip"),
        }
    res = st.session_state.get("r_res")
    if res:
        _render_results(res, "r_dl", "rosters.zip", DOCX_MIME, "rosters")


with tab_signoff:
    st.markdown('<div class="section-head">Sign-Off Sheet Generator</div>', unsafe_allow_html=True)
    st.markdown('<div class="hint">Upload the sign-off export \u2192 download TLMS sign-off Excel sheets.</div>',
                unsafe_allow_html=True)
    with st.expander("\u2139\uFE0F  How it works"):
        st.markdown(
            "1. Export sign-off data to Excel.\n"
            "2. Upload the `.xlsx` below.\n"
            "3. Click **Generate** \u2014 each row produces a `.xlsx` with Sheet1 (user import) "
            "and Sheet2 (course completion).\n\n"
            "Course names are mapped to TLMS course codes using the bundled mapping table. "
            "To update the mapping, edit `mapping/Signoff_Mapping.xlsx` in the repo and push."
        )
    so_upload = st.file_uploader("Sign-off export", type=["xlsx"], key="so_upload")
    so_go = st.button("Generate sign-off sheets", type="primary", use_container_width=True,
                      disabled=so_upload is None, key="so_go")
    if so_upload and so_go:
        if not os.path.exists(MAPPING_PATH):
            st.error(f"Mapping file not found at `{MAPPING_PATH}`.")
            st.stop()
        mapping = bs.load_mapping(MAPPING_PATH)
        so_log, so_files = [], []
        with st.spinner("Building sign-off sheets \u2026"):
            with tempfile.TemporaryDirectory() as tmp:
                xp = os.path.join(tmp, "e.xlsx")
                with open(xp, "wb") as f: f.write(so_upload.getbuffer())
                od = os.path.join(tmp, "out"); os.makedirs(od)
                rows = bs.read_signoff_export(xp)
                made = 0
                for rec in rows:
                    s, info, n = bs.build_signoff(rec, mapping, od)
                    if s == "ok":
                        made += 1
                        with open(info, "rb") as fh:
                            so_files.append((os.path.basename(info), fh.read()))
                        so_log.append(("ok", os.path.basename(info), f"{n} participants"))
                    else:
                        so_log.append(("skip", rec.get("customer") or "\u2014", info))
            zbuf = io.BytesIO()
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
                for nm, d in so_files: z.writestr(nm, d)
        st.session_state["so_res"] = {
            "files": so_files, "zip": zbuf.getvalue(), "log": so_log,
            "made": made, "skip": sum(1 for s, *_ in so_log if s == "skip"),
        }
    sres = st.session_state.get("so_res")
    if sres:
        _render_results(sres, "so_dl", "signoff_sheets.zip", XLSX_MIME, "sign-off sheets")

st.markdown('<div class="foot">F.A.S.T. Rescue Incorporated \u00b7 Training Tools</div>',
            unsafe_allow_html=True)
