"""
MYNM Project Dashboard — Streamlit (English)

Run locally:
	streamlit run "maximo_streamlit/MYNM Dashboard.py"

Data sources:
- Reads from one of:
	1) Secret/Env MYNM_DATA_URL (CSV/Excel URL or local file path)
	2) Any local xlsx containing "mynm" or "tickets" (auto-discovery)
	3) File uploaded via the app when no source is found

Requirements: streamlit, pandas, plotly, requests, openpyxl
"""
from __future__ import annotations

import io
import os
import re
import glob
from pathlib import Path
from datetime import datetime, timedelta, timezone
from urllib.parse import quote_plus
from zoneinfo import ZoneInfo
from typing import Optional, Tuple, Dict, Any, List
import warnings

import numpy as np

import pandas as pd
import plotly.express as px
import streamlit as st



# -------------------------
# Page setup and branding
# -------------------------
st.set_page_config(
	page_title="MYNM Project Dashboard",
	page_icon="🟧",
	layout="wide",
	initial_sidebar_state="collapsed",
)

PRIMARY = "#FFD400"  # Bright yellow
COLOR_SEQ = (
	px.colors.qualitative.Vivid
	+ px.colors.qualitative.Safe
	+ px.colors.qualitative.Set2
	+ px.colors.qualitative.Pastel
)
ORANGE_SEQ = ["#FFD400", "#FFC300", "#FFDB4D", "#FFF07A"]

# Google Sheets file (strict: link-only via gviz CSV)
GOOGLE_SHEET_FILE_ID = "1zAEL1KVTYo-35va7bVdaXkneunIAZln0"


# -------------------------
# Utilities (secrets/time/formatting)
# -------------------------
def get_secret_env(key: str, default: Optional[str] = None) -> Optional[str]:
	try:
		if hasattr(st, "secrets") and key in st.secrets and str(st.secrets.get(key)).strip():
			return str(st.secrets.get(key))
	except Exception:
		pass
	v = os.environ.get(key)
	return v if (v is not None and str(v).strip()) else default


def _get_tz_name() -> str:
	return get_secret_env("MYNM_TZ", get_secret_env("DASH_TZ", "Asia/Riyadh")) or "Asia/Riyadh"


def _get_zoneinfo(tz_name: str):
	try:
		return ZoneInfo(tz_name)
	except Exception:
		return timezone(timedelta(hours=3))  # UTC+3 fallback


def now_local() -> datetime:
	return datetime.now(_get_zoneinfo(_get_tz_name()))


def now_local_naive() -> datetime:
	return now_local().replace(tzinfo=None)


def inject_brand_css() -> None:
	st.markdown(
		f"""
		<style>
		:root {{
			--accent: {PRIMARY};
			--shadow: 0 8px 18px rgba(0,0,0,0.08);
			--radius: 12px;
		}}
		body {{ direction: ltr; font-family: 'Segoe UI','Inter',system-ui,sans-serif; color:#111; }}
		header {{ visibility: hidden; }}
		.block-container {{ padding-top: 0.8rem; padding-bottom: 0.8rem; }}
		/* Ensure tabs start from the left */
		.stTabs, div[role="tablist"] {{ direction: ltr; }}
		div[role="tablist"] {{ justify-content: flex-start; }}
		.brand-header {{
			display:grid; grid-template-columns: 160px 1fr 160px; align-items:center;
			min-height: 110px; padding: .6rem 1rem; background: rgba(255,255,255,.92);
			border-bottom: 2px solid var(--accent); border-radius: 12px; backdrop-filter: blur(6px);
			margin-bottom: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.05);
		}}
		.logo-img {{ height:64px; width:64px; border-radius:10px; object-fit: contain; }}
		.brand-title {{ margin:0; font-size: 36px; color:#111; font-weight: 900; white-space: nowrap; }}
		.brand-sub {{ margin: 0; color:#111; font-size:14px; font-weight:800; }}
		.kpi-card {{ background:#fff; border-radius: 16px; box-shadow: var(--shadow); padding: 14px; border:1px solid rgba(0,0,0,.06); }}
		.kpi-title {{ color:#111; margin:0; font-weight:800; }}
		.kpi-value {{ margin:0; font-size:40px; font-weight:1000; color:#000; }}
		</style>
		""",
		unsafe_allow_html=True,
	)


def _file_to_data_uri(path: str) -> Optional[str]:
	try:
		with open(path, "rb") as f:
			data = f.read()
		import base64
		ext = os.path.splitext(path)[1].lower()
		mime = "image/png"
		if ext in (".jpg", ".jpeg"): mime = "image/jpeg"
		if ext == ".svg": mime = "image/svg+xml"
		if ext == ".webp": mime = "image/webp"
		b64 = base64.b64encode(data).decode()
		return f"data:{mime};base64,{b64}"
	except Exception:
		return None


def _gdrive_to_direct(u: str) -> str:
	if not u:
		return u
	if "drive.google.com/file/d/" in u:
		try:
			fid = u.split("/file/d/")[1].split("/")[0]
			return f"https://drive.google.com/uc?export=view&id={fid}"
		except Exception:
			return u
	return u



def _resolve_logo(default_url: str, env_keys: List[str], local_names: List[str], keywords: Optional[List[str]] = None) -> str:
	# 1) Local folders (project assets + Maximo_DEV/Logo)
	base_dir = os.path.dirname(__file__)
	project_root = os.path.dirname(base_dir)
	search_dirs = [
		os.path.join(base_dir, "assets"),
		os.path.join(project_root, "Maximo_DEV", "Logo"),
	]
	for folder in search_dirs:
		for n in local_names:
			p = os.path.join(folder, n)
			if os.path.exists(p):
				uri = _file_to_data_uri(p)
				if uri:
					return uri
		# بحث تقريبي بالكلمات الدلالية
		if keywords and os.path.isdir(folder):
			try:
				for fname in os.listdir(folder):
					fl = fname.lower()
					if any(kw.lower() in fl for kw in keywords) and os.path.splitext(fname)[1].lower() in [".png", ".jpg", ".jpeg", ".svg", ".webp"]:
						p = os.path.join(folder, fname)
						uri = _file_to_data_uri(p)
						if uri:
							return uri
			except Exception:
				pass
	# 2) Secrets/Env
	for k in env_keys:
		v = get_secret_env(k)
		if v:
			return _gdrive_to_direct(v)
	# 3) افتراضي
	return _gdrive_to_direct(default_url)


def brand_header() -> None:
	# Left: MYNM — Right: White Art
	right_logo = _resolve_logo(
		default_url="https://raw.githubusercontent.com/encharm/Font-Awesome-SVG-PNG/master/black/png/128/angellist-128.png",
		env_keys=["MYNM_RIGHT_LOGO", "WHITEART_LOGO"],
		local_names=["WA Logo.jpeg", "WA Logo.jpg", "white_art_logo.png", "white_art_logo.jpg", "white_art_logo.svg", "wa.png", "wa.jpg"],
		keywords=["white", "art", "wa"]
	)
	left_logo = _resolve_logo(
		default_url="https://upload.wikimedia.org/wikipedia/commons/4/44/BMW.svg",
		env_keys=["MYNM_LEFT_LOGO", "MYNM_CLIENT_LOGO"],
		local_names=["MYNM Logo.jpeg", "MYNM Logo.jpg", "mynm_logo.png", "mynm_logo.jpg", "mynm_logo.svg", "client_logo.png", "client_logo.jpg"],
		keywords=["mynm", "client", "bmw"]
	)
	now_str = now_local().strftime("%a, %d %b %Y – %H:%M %Z")
	st.markdown(
		f"""
		<div class='brand-header'>
			<div style='text-align:right;'>
				<img src="{left_logo}" class="logo-img" />
			</div>
			<div style='text-align:center;'>
				<h3 class='brand-title'>MYNM Project Dashboard</h3>
				<div class='brand-sub'>{now_str}</div>
			</div>
			<div style='text-align:left;'>
				<img src="{right_logo}" class="logo-img" />
			</div>
		</div>
		""",
		unsafe_allow_html=True,
	)


# -------------------------
# Data loading (Excel/CSV)
# -------------------------
def _parse_dates_safely(series: pd.Series) -> pd.Series:
	if series is None:
		return series
	s = series.astype(str).str.strip()
	if getattr(series, 'dtype', None) is not None and str(series.dtype).startswith('datetime'):
		return series
	fmts = [
		"%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y",
		"%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M", "%m/%d/%Y %H:%M",
	]
	best = (None, -1)
	for f in fmts:
		p = pd.to_datetime(s, format=f, errors="coerce")
		n = int(p.notna().sum())
		if n > best[1]:
			best = (p, n)
			if n == len(s):
				break
	if best[0] is not None and best[1] > 0:
		return best[0]
	with warnings.catch_warnings():
		warnings.simplefilter("ignore", category=UserWarning)
		return pd.to_datetime(s, errors="coerce", dayfirst=True)


# External URL reading is intentionally removed to avoid external API calls.


def _auto_find_local_excel() -> Optional[str]:
	return None


def _discover_excels(limit: int = 50) -> List[str]:
	return []


def _parse_excel_bytes(content: bytes, source_label: str) -> Dict[str, Any]:
	return {"west": None, "central": None, "mobilization_raw": None, "source": source_label, "error": None}


@st.cache_data(ttl=60, show_spinner=False)
def load_tickets_csv(csv_url: str) -> pd.DataFrame:
	"""Load tickets from a CSV URL (e.g., Google Sheets export)."""
	df = pd.read_csv(csv_url)
	# Normalize columns: strip headers
	df.columns = [str(c).strip() for c in df.columns]
	# Attempt to parse common date columns
	for c in ["Receiving date", "Receiving Date", "Create Date", "Created Date", "Response Date", "Rectification Date", "Due Date", "Target Date"]:
		if c in df.columns:
			df[c] = _parse_dates_safely(df[c])
	return df


def _gsheet_csv_url(file_id: str, sheet_name: str) -> str:
	return f"https://docs.google.com/spreadsheets/d/{file_id}/gviz/tq?tqx=out:csv&sheet={quote_plus(sheet_name)}"


@st.cache_data(ttl=120, show_spinner=False)
def load_gs_tickets(file_id: str) -> pd.DataFrame:
	"""Load West and Central sheets as CSV and combine with Region column."""
	west = pd.read_csv(_gsheet_csv_url(file_id, "MYN Tickets-West"))
	central = pd.read_csv(_gsheet_csv_url(file_id, "MYN Tickets-Central"))
	for df in (west, central):
		df.columns = [str(c).strip() for c in df.columns]
		for c in ["Receiving date", "Receiving Date", "Create Date", "Created Date", "Response Date", "Rectification Date", "Due Date", "Target Date"]:
			if c in df.columns:
				df[c] = _parse_dates_safely(df[c])
	west["Region"] = "West"
	central["Region"] = "Central"
	return pd.concat([west, central], ignore_index=True, sort=False)


@st.cache_data(ttl=120, show_spinner=False)
def load_mobilization_raw_csv(file_id: str) -> pd.DataFrame:
	"""Load Mobilization sheet as CSV without headers to preserve raw sections."""
	# Read twice: first with header=None for section parsing. If it fails in view, we can adapt.
	return pd.read_csv(_gsheet_csv_url(file_id, "Mobilization"), header=None)


def _workspace_root_dir() -> Path:
	# This file lives in maximo_streamlit/; project root is its parent.
	return Path(__file__).resolve().parents[1]


def _find_mobilization_excel_path() -> Optional[str]:
	"""Find a local .xlsx in the project root that contains a 'Mobilization' sheet."""
	root = _workspace_root_dir()
	# Search recursively so the file can live in subfolders; skip common large/irrelevant dirs.
	skip_parts = {".venv", "build", "output", "__pycache__"}
	xlsx_files: List[Path] = []
	for p in root.rglob("*.xlsx"):
		if p.name.startswith("~$"):
			continue
		parts = {part.lower() for part in p.parts}
		if parts & skip_parts:
			continue
		xlsx_files.append(p)
	xlsx_files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
	for p in xlsx_files:
		try:
			xls = pd.ExcelFile(p, engine="openpyxl")
			if any(str(s).strip().lower() == "mobilization" for s in xls.sheet_names):
				return str(p)
		except Exception:
			continue
	return None


@st.cache_data(ttl=120, show_spinner=False)
def load_mobilization_raw_excel(excel_path: str) -> pd.DataFrame:
	"""Load Mobilization sheet from local Excel (headerless) for section parsing."""
	return pd.read_excel(excel_path, sheet_name="Mobilization", header=None, engine="openpyxl")


def load_data() -> Dict[str, Any]:
	return {"west": None, "central": None, "mobilization_raw": None, "source": None, "error": "Disabled: link-only mode"}


def _columns_map(df: pd.DataFrame) -> Dict[str, str]:
	def _k(s: str) -> str:
		s = str(s).strip().lower()
		return "".join(ch for ch in s if ch.isalnum())
	return {_k(c): c for c in df.columns}


def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
	cmap = _columns_map(df)
	def _k(s: str) -> str:
		s = str(s).strip().lower()
		return "".join(ch for ch in s if ch.isalnum())
	for c in candidates:
		k = _k(c)
		if k in cmap:
			return cmap[k]
	return None


# -------------------------
# Tickets (West/Central)
# -------------------------
TICKET_SHEET_WEST = [
	"MYN Tickets-West",
	"myn tickets-west",
	"tickets-west",
	"west",
]
TICKET_SHEET_CENTRAL = [
	"MYN Tickets-Central",
	"myn tickets-central",
	"tickets-central",
	"central",
]
MOBILIZATION_SHEET = ["Mobilization", "mobilization"]


def _match_sheet_name(xls: pd.ExcelFile, names: List[str]) -> Optional[str]:
	def _norm(s: str) -> str:
		s = str(s).strip().lower()
		return "".join(ch for ch in s if ch.isalnum())

	target_norms = [_norm(n) for n in names]
	mapping = {sn: _norm(sn) for sn in xls.sheet_names}

	# 1) مطابقة تامة بعد التطبيع
	for real, key in mapping.items():
		if key in target_norms:
			return real
	# 2) احتواء جزئي
	for real, key in mapping.items():
		if any(t in key for t in target_norms):
			return real
	return None


def combine_tickets_frames(west: Optional[pd.DataFrame], central: Optional[pd.DataFrame]) -> pd.DataFrame:
	parts: List[pd.DataFrame] = []
	if west is not None: parts.append(west)
	if central is not None: parts.append(central)
	return pd.concat(parts, ignore_index=True, sort=False) if parts else pd.DataFrame()


def ticket_kpis(df: pd.DataFrame) -> Dict[str, Any]:
	# Prefer 'WO Status' then fallback to 'Status'
	status_col = find_col(df, ["WO Status", "Status"]) or ("WO Status" if "WO Status" in df.columns else None)
	recv_col = find_col(df, ["Receiving date", "Receiving Date", "Create Date", "Created Date"]) or None
	due_col = find_col(df, ["Due Date", "Target Date", "SLA Date"]) or None

	total = len(df)
	if status_col:
		st_u = df[status_col].astype(str).str.upper()
		closed_mask = st_u.str.contains("CLOSED") | st_u.str.contains("COMPLETED")
		open_mask = ~closed_mask
		closed_n = int(closed_mask.sum())
		open_n = int(open_mask.sum())
	else:
		closed_n = open_n = 0

	# Overdue: prefer Due/Target date; fallback to open tickets older than 48h from Receiving date
	overdue_n = 0
	if due_col and status_col:
		due = pd.to_datetime(df[due_col], errors="coerce")
		now_ts = pd.Timestamp.now(tz=None)
		overdue_n = int(((due.notna()) & (due < now_ts) & open_mask).sum())
	elif recv_col and status_col:
		recv = pd.to_datetime(df[recv_col], errors="coerce")
		cut = pd.Timestamp.now(tz=None) - pd.Timedelta(hours=48)
		overdue_n = int(((recv.notna()) & (recv < cut) & open_mask).sum())

	return {
		"total": total,
		"open": open_n,
		"closed": closed_n,
		"overdue": overdue_n,
	}


def kpi_card(title: str, value: str) -> None:
	st.markdown(
		f"""
		<div class='kpi-card'>
			<p class='kpi-title'>{title}</p>
			<p class='kpi-value'>{value}</p>
		</div>
		""",
		unsafe_allow_html=True,
	)


def _fmt(v: Any, unit: str = "") -> str:
	if v is None or (isinstance(v, float) and (v != v)):
		return "—"
	try:
		if unit == "%":
			return f"{float(v):.0f}%"
		if unit == "h":
			return f"{float(v):.1f} h"
		if unit == "int":
			return f"{int(v):,}"
		return str(v)
	except Exception:
		return str(v)


def operation_charts(df: pd.DataFrame) -> None:
	status_col = find_col(df, ["WO Status", "Status"]) or None
	recv_col = find_col(df, ["Receiving date", "Receiving Date", "Create Date", "Created Date"]) or None
	region_col = "Region" if "Region" in df.columns else None

	# 1) Bar: tickets by status
	if status_col:
		s = df[status_col].astype(str).str.strip()
		if not s.empty:
			vc = s.value_counts().reset_index()
			vc.columns = ["Status", "Count"]
			fig = px.bar(vc, x="Status", y="Count", text="Count", title="Tickets by Status", color_discrete_sequence=ORANGE_SEQ)
			fig.update_traces(textposition="outside", cliponaxis=False)
			fig.update_xaxes(tickangle=-30)
			st.plotly_chart(fig, use_container_width=True)

	# 2) Pie: ticket distribution by region
	if region_col:
		reg = df[region_col].astype(str).str.strip()
		if not reg.empty:
			vc = reg.value_counts().reset_index()
			vc.columns = ["Region", "Count"]
			fig = px.pie(vc, names="Region", values="Count", title="Ticket Distribution by Region", hole=0.45, color_discrete_sequence=COLOR_SEQ)
			st.plotly_chart(fig, use_container_width=True)

	# 3) Line: ticket creation trend over time
	if recv_col:
		ds = pd.to_datetime(df[recv_col], errors="coerce")
		valid = ds.notna()
		if valid.any():
			months = ds.dt.to_period("M").astype(str)
			cnt = months.value_counts().rename_axis("Month").reset_index(name="Count").sort_values("Month")
			fig = px.line(cnt, x="Month", y="Count", markers=True, title="Tickets Over Time", color_discrete_sequence=ORANGE_SEQ)
			st.plotly_chart(fig, use_container_width=True)


# -------------------------
# Mobilization sheet parsing (extract section tables)
# -------------------------
def _extract_table_after_title(df_raw: pd.DataFrame, title: str) -> Optional[pd.DataFrame]:
	# df_raw is without headers (header=None). Titles can be truncated (e.g., 'Team Hi') or merged.
	def _norm(s: Any) -> str:
		s = "" if s is None else str(s)
		s = s.strip().lower()
		return re.sub(r"[^a-z0-9\u0600-\u06FF]+", "", s)

	def _cell_matches_title(cell: Any, title_norm: str) -> bool:
		cn = _norm(cell)
		if not cn or not title_norm:
			return False
		# Prefer exact/containment matches first
		if cn == title_norm or title_norm in cn:
			return True
		# Allow truncated matches, but avoid generic tokens like "city" matching long titles.
		# Require the truncated token to be at least ~50% of the title length.
		min_len = max(4, int(len(title_norm) * 0.5))
		return (cn in title_norm) and (len(cn) >= min_len)

	def _row_matches_title(row: pd.Series, title_norm: str) -> bool:
		for x in row.tolist():
			if _cell_matches_title(x, title_norm):
				return True
		return False

	def _row_matches_any_section(row: pd.Series) -> bool:
		for sec in MOB_SECTIONS:
			sec_n = _norm(sec)
			if sec_n and _row_matches_title(row, sec_n):
				return True
		return False

	title_n = _norm(title)
	if not title_n:
		return None

	mask = df_raw.apply(lambda r: _row_matches_title(r, title_n), axis=1)
	idx = mask[mask].index.min() if mask.any() else None
	if idx is None:
		return None

	# Detect if header is on the SAME row as the title (common in Google Sheets CSV export)
	# Example: col0='Team Hi' and col1='Item' col2='Planned'...
	header_keywords = [
		"item",
		"city",
		"planned",
		"planed",
		"plannec",
		"onduty",
		"onsite",
		"purchased",
		"distributed",
		"rented",
		"occupied",
		"contracted",
		"signed",
		"remaining",
		"remain",
	]
	row0 = df_raw.iloc[idx]
	hits = 0
	for x in row0.tolist():
		xn = _norm(x)
		if not xn:
			continue
		if any(k in xn for k in header_keywords):
			hits += 1
	# Header row is either the title row (if it looks like header), otherwise the next non-empty row.
	header_row: Optional[int] = idx if hits >= 2 else None
	if header_row is None:
		for r in range(idx + 1, len(df_raw)):
			row = df_raw.iloc[r]
			non_na = int(row.notna().sum())
			if non_na >= 2:
				header_row = r
				break
			if non_na <= 2 and _row_matches_any_section(row):
				return None
	if header_row is None:
		return None

	header = df_raw.iloc[header_row].astype(str).str.strip().tolist()
	# If header row is also the title row, replace the first cell (title) with an index header.
	if header_row == idx and header and _cell_matches_title(header[0], title_n):
		header[0] = "n"
	table_rows: List[List[Any]] = []
	empty_streak = 0
	for r in range(header_row + 1, len(df_raw)):
		row = df_raw.iloc[r]
		non_na = int(row.notna().sum())
		if non_na == 0:
			empty_streak += 1
			if empty_streak >= 2:
				break
			continue
		empty_streak = 0
		# Stop when next section title starts (even if the title row contains headers too)
		if _row_matches_any_section(row):
			break
		table_rows.append(row.tolist())

	if not table_rows:
		return None

	tdf = pd.DataFrame(table_rows, columns=header)
	# Drop sequence column if present (n)
	for c in list(tdf.columns):
		if str(c).strip().lower() in {"n", "#", "no"}:
			tdf = tdf.drop(columns=[c])
	# Clean column names
	tdf.columns = [str(c).strip() for c in tdf.columns]
	# Drop empty/unnamed columns
	bad_cols = [c for c in tdf.columns if not str(c).strip() or str(c).strip().lower() in {"nan", "none"}]
	if bad_cols:
		tdf = tdf.drop(columns=bad_cols, errors="ignore")
	# Exclude any row containing 'Total' anywhere (case-insensitive)
	mask_total = tdf.apply(lambda rr: rr.astype(str).str.contains(r"\btotal\b", case=False, na=False).any(), axis=1)
	tdf = tdf[~mask_total].copy()
	return tdf.dropna(how="all").reset_index(drop=True)


MOB_SECTIONS = [
	"Team Hiring",
	"Manpower Coverage on Cities",
	"Cars",
	"Tools Readiness",
	"Uniform Readiness",
	"Accommodation Readiness",
	"Fire AMCs",
]

# Per-section metrics mapping: (planned candidates, completed candidates)
MOB_SECTION_METRICS: Dict[str, Tuple[List[str], List[str]]] = {
	"Team Hiring": (["Planned", "Planed", "Plannec"], ["On duty", "On Duty", "On duti"]),
	"Manpower Coverage on Cities": (["Planned", "Planed", "Plannec"], ["On duty", "On Duty", "On duti"]),
	"Cars": (["Planned", "Planed", "Plannec"], ["On Site", "On site", "On Duty", "Onsite"]),
	"Tools Readiness": (["Purchased", "Purchase"], ["distributed", "Distributed", "Distribute"]),
	"Uniform Readiness": (["Purchased", "Purchase"], ["distributed", "Distributed", "Distribute"]),
	"Accommodation Readiness": (["Rented", "Rent"], ["Occupied", "Occupy"]),
	"Fire AMCs": (["Contracted", "Contract"], ["Signed", "Sign"]),
}


def read_mobilization_from_raw(raw: pd.DataFrame) -> Dict[str, pd.DataFrame]:
	if raw is None or raw.empty:
		return {}
	out: Dict[str, pd.DataFrame] = {}
	for sec in MOB_SECTIONS:
		tbl = _extract_table_after_title(raw, sec)
		if tbl is not None and not tbl.empty:
			out[sec] = tbl
	return out


def mobilization_view(sections: Dict[str, pd.DataFrame]) -> None:
	if not sections:
		st.info("Could not extract Mobilization tables from the sheet.")
		return

	def _clean_mobil_df(df: pd.DataFrame) -> pd.DataFrame:
		mobil_df = df.copy()
		# Drop completely empty rows
		mobil_df = mobil_df.dropna(how="all")
		# Remove any row that contains 'Total' in any cell
		mask_total = mobil_df.apply(
			lambda rr: rr.astype(str).str.contains(r"\btotal\b", case=False, na=False).any(),
			axis=1,
		)
		mobil_df = mobil_df[~mask_total].copy()
		# Clean Status column
		if "Status" in mobil_df.columns:
			mobil_df["Status"] = mobil_df["Status"].fillna("Pending")
		return mobil_df

	def _compute_row_status(section_name: str, df: pd.DataFrame) -> Tuple[pd.DataFrame, Optional[str], Optional[str], str]:
		work = _clean_mobil_df(df)
		work.columns = [str(c).strip() for c in work.columns]
		dim_col = next(
			(c for c in work.columns if str(c).lower().startswith(("city", "item", "name", "category"))),
			work.columns[0],
		)
		# Remove empty dimension
		work = work[work[dim_col].notna()].copy()

		plan_cands, comp_cands = MOB_SECTION_METRICS.get(
			section_name,
			(["Planned", "Total", "Target"], ["Completed", "On duty", "On Duty", "On Site", "distributed", "Occupied", "Signed"]),
		)
		planned_col = find_col(work, plan_cands)
		completed_col = find_col(work, comp_cands)
		if not planned_col or not completed_col:
			work["Status"] = "Pending"
			return work, planned_col, completed_col, dim_col

		planned = pd.to_numeric(work[planned_col], errors="coerce")
		completed = pd.to_numeric(work[completed_col], errors="coerce")
		denom = planned.replace({0: np.nan})
		ratio = (completed / denom).replace([np.inf, -np.inf], np.nan)
		# If planned is missing/0: treat completed>0 as Completed, else Pending
		ratio = ratio.fillna((completed.fillna(0) > 0).astype(float))
		ratio = ratio.clip(lower=0)

		work["Status"] = pd.cut(
			ratio,
			bins=[-0.001, 0.0, 0.999999, 10.0],
			labels=["Pending", "In Progress", "Completed"],
			include_lowest=True,
		)
		work["Status"] = work["Status"].astype(str).replace({"nan": "Pending", "None": "Pending"})
		return work, planned_col, completed_col, dim_col

	for sec, df in sections.items():
		st.subheader(sec)
		work, planned_col, completed_col, dim_col = _compute_row_status(sec, df)
		if work.empty:
			st.info("No rows found (after excluding 'Total').")
			st.divider()
			continue

		# Always compute safe progress from Status distribution
		status_count = (
			work["Status"]
			.fillna("Pending")
			.astype(str)
			.value_counts()
			.reset_index()
		)
		status_count.columns = ["Status", "Count"]
		total = int(status_count["Count"].sum())
		completed = int(status_count.loc[status_count["Status"].str.lower() == "completed", "Count"].sum())
		progress = round((completed / total) * 100, 1) if total > 0 else 0

		c1, c2 = st.columns([1, 2])
		with c1:
			st.metric("Completion %", f"{progress:.1f}%")
			st.progress(min(max(progress / 100.0, 0.0), 1.0))
		with c2:
			# Reindex to stable order
			order = ["Completed", "In Progress", "Pending"]
			sdist = (
				status_count.assign(_ord=lambda d: d["Status"].map({k: i for i, k in enumerate(order)}))
				.sort_values(by="_ord", na_position="last")
				.drop(columns=["_ord"], errors="ignore")
			)
			fig = px.pie(
				sdist,
				names="Status",
				values="Count",
				title=f"{sec} — Status Distribution",
				hole=0.45,
				color="Status",
				color_discrete_map={"Completed": "#2E7D32", "In Progress": "#F9A825", "Pending": "#C62828"},
			)
			fig.update_traces(textposition="inside", textinfo="percent+label")
			st.plotly_chart(fig, use_container_width=True)

		show_cols: List[str] = [dim_col]
		if planned_col:
			show_cols.append(planned_col)
		if completed_col:
			show_cols.append(completed_col)
		show_cols.append("Status")
		st.dataframe(work[show_cols], use_container_width=True, hide_index=True)
		st.divider()


# -------------------------
# App entry
# -------------------------
def main() -> None:
	inject_brand_css()
	brand_header()

	tabs = st.tabs(["Operation", "Mobilization"])

	# Link-only mode: no local discovery or upload

	# --- Operation tab ---
	with tabs[0]:
		# Strict link-only: load both sheets from Google Sheets FILE_ID
		try:
			df = load_gs_tickets(GOOGLE_SHEET_FILE_ID)
			st.caption("Data source: Google Sheets (West + Central)")
			# Filters
			recv_col = find_col(df, ["Receiving date", "Receiving Date", "Create Date", "Created Date"]) or None
			if recv_col:
				ds = pd.to_datetime(df[recv_col], errors="coerce")
				min_d = ds.min()
				max_d = ds.max()
				if pd.notna(min_d) and pd.notna(max_d):
					start, end = st.slider("Date range (Receiving Date)", value=(min_d.to_pydatetime(), max_d.to_pydatetime()))
					mask = ds.between(start, end)
					df = df[mask].copy()

			k = ticket_kpis(df)
			c1, c2, c3, c4 = st.columns(4)
			with c1:
				kpi_card("Total Tickets", _fmt(k.get("total"), "int"))
			with c2:
				kpi_card("Open Tickets", _fmt(k.get("open"), "int"))
			with c3:
				kpi_card("Closed Tickets", _fmt(k.get("closed"), "int"))
			with c4:
				kpi_card("Overdue Tickets", _fmt(k.get("overdue"), "int"))

			operation_charts(df)
		except Exception as exc:
			st.error(f"Error loading Google Sheets tickets: {exc}")

	# --- Mobilization tab ---
	with tabs[1]:
		try:
			excel_path = _find_mobilization_excel_path()
			if excel_path:
				raw = load_mobilization_raw_excel(excel_path)
				st.caption(f"Mobilization source: local Excel ({os.path.basename(excel_path)})")
			else:
				raw = load_mobilization_raw_csv(GOOGLE_SHEET_FILE_ID)
				st.caption("Mobilization source: Google Sheets (fallback)")

			sections = read_mobilization_from_raw(raw)
			if not sections:
				st.info("Mobilization sections not detected in sheet. Showing raw table.")
				st.dataframe(raw, use_container_width=True, hide_index=True)
			else:
				mobilization_view(sections)
		except Exception as exc:
			st.error(f"Error loading Mobilization sheet: {exc}")


if __name__ == "__main__":
	main()

