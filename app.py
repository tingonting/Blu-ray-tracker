import random
import re
import urllib.parse
import datetime
import concurrent.futures
import openpyxl
import pandas as pd
import streamlit as st

# Set up page configuration & layout
st.set_page_config(
    page_title="Conor's Blu-ray Hub", page_icon="🎬", layout="wide"
)

# Custom styling — "Blu-ray Hub" theme
# Palette drawn from the format itself: Blu-ray's 405nm laser gives the violet,
# a disc's iridescent shimmer gives the teal, and cinema marquee bulbs give the amber.
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Manrope:wght@400;600;800&family=JetBrains+Mono:wght@500;700&display=swap');

    :root {
        --ink: #0A0B14;
        --panel: #14162A;
        --violet: #6C5CE8;
        --teal: #2DD4BF;
        --amber: #F0A83B;
        --paper: #EDECF5;
        --mute: #8B889E;
    }

    /* Content width — narrow on phones (unchanged from before), progressively
       roomier on iPad and desktop instead of sitting in a fixed 600px column
       with dead space on either side. layout="wide" above removes Streamlit's
       own centered-mode cap so this is the only width rule in play. */
    .block-container {
        max-width: 600px;
        margin: 0 auto;
        padding-top: 2rem;
    }
    @media (min-width: 768px) {
        .block-container { max-width: 820px; }
    }
    @media (min-width: 1200px) {
        .block-container { max-width: 1000px; }
    }

    .stApp {
        font-family: 'Manrope', sans-serif;
    }

    /* Body text only — deliberately NOT applied to span/div broadly, since
       Streamlit renders icons (expander arrows, checkmarks, etc.) as icon-font
       ligatures inside spans, and overriding their font breaks those glyphs. */
    .stApp p, .stApp label, .stApp li, .stApp td, .stApp th {
        font-family: 'Manrope', sans-serif;
    }

    /* Headline / section-header treatment */
    h1, h2, h3 {
        font-family: 'Bebas Neue', sans-serif !important;
        letter-spacing: 0.04em;
        text-transform: uppercase;
    }

    /* App title: violet-to-teal gradient, like a disc catching the light */
    .app-title {
        text-align: center;
        font-family: 'Bebas Neue', sans-serif;
        font-size: 2.6rem;
        letter-spacing: 0.06em;
        margin-bottom: 0px;
        background: linear-gradient(100deg, var(--violet) 20%, var(--teal) 80%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    .app-subtitle {
        text-align: center;
        color: var(--mute);
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.8rem;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        margin-top: -4px;
    }

    /* Sprocket-hole divider — a nod to physical film reels, standing in for st.divider() */
    .film-divider {
        display: flex;
        justify-content: space-between;
        margin: 18px 2px;
        opacity: 0.55;
    }
    .film-divider span {
        width: 7px;
        height: 7px;
        border-radius: 2px;
        background: var(--violet);
    }
    .film-divider span:nth-child(3n) { background: var(--teal); }

    /* Card/Metric styling */
    div[data-testid="stMetric"] {
        background: linear-gradient(160deg, rgba(108, 92, 232, 0.12), rgba(45, 212, 191, 0.05));
        border: 1px solid rgba(108, 92, 232, 0.25);
        padding: 10px;
        border-radius: 14px;
        text-align: center;
    }
    div[data-testid="stMetricValue"] {
        font-family: 'JetBrains Mono', monospace;
        color: var(--amber);
    }

    /* Winner card for random picker — marquee spotlight */
    .winner-box {
        position: relative;
        overflow: hidden;
        background: linear-gradient(135deg, rgba(108, 92, 232, 0.18), rgba(45, 212, 191, 0.1));
        border: 2px solid var(--amber);
        padding: 20px;
        border-radius: 18px;
        text-align: center;
        margin-top: 15px;
        margin-bottom: 15px;
        animation: winner-glow 4s ease-in-out infinite;
    }

    /* A light sweep across the card, like a marquee catching the light */
    .winner-box::before {
        content: "";
        position: absolute;
        top: 0;
        left: -60%;
        width: 40%;
        height: 100%;
        background: linear-gradient(100deg, transparent, rgba(255, 255, 255, 0.16), transparent);
        animation: winner-shimmer 3.6s ease-in-out infinite;
        pointer-events: none;
    }

    @keyframes winner-shimmer {
        0%   { left: -60%; }
        55%  { left: 130%; }
        100% { left: 130%; }
    }

    @keyframes winner-glow {
        0%, 100% { box-shadow: 0 4px 20px rgba(240, 168, 59, 0.22); border-color: var(--amber); }
        50%      { box-shadow: 0 4px 24px rgba(108, 92, 232, 0.3); border-color: var(--violet); }
    }

    /* Format badges — Blu-ray in violet (the laser color), 4K UHD in teal (the shimmer color) */
    .format-badge {
        display: inline-block;
        padding: 1px 9px;
        border-radius: 20px;
        font-size: 0.85em;
        font-weight: 700;
    }
    .format-badge.bluray {
        background: rgba(108, 92, 232, 0.18);
        color: #A79CFF;
    }
    .format-badge.uhd4k {
        background: rgba(45, 212, 191, 0.18);
        color: var(--teal);
    }
    .format-badge.other {
        background: rgba(139, 136, 158, 0.18);
        color: var(--mute);
    }

    /* Make buttons touch-friendly and prominent */
    .stButton button {
        width: 100%;
        border-radius: 12px;
        font-weight: 700;
        padding: 0.6rem 1rem;
        font-family: 'Manrope', sans-serif;
    }

    /* Form container look */
    div[data-testid="stForm"] {
        border: 1px solid rgba(108, 92, 232, 0.2);
        border-radius: 16px;
        padding: 15px;
        background-color: rgba(108, 92, 232, 0.04);
    }

    /* Dataframe responsiveness */
    div[data-testid="stDataFrame"] {
        width: 100%;
        border-radius: 12px;
    }

    /* Tabs: quiet until selected, then lit up in violet with a teal underline */
    button[data-baseweb="tab"] {
        font-family: 'Manrope', sans-serif;
        font-weight: 600;
        color: var(--mute);
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: var(--paper);
    }
    div[data-baseweb="tab-highlight"] {
        background-color: var(--teal) !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)


def film_divider():
  """Sprocket-hole divider — the app's signature element, standing in for st.markdown('---')."""
  st.markdown(
      '<div class="film-divider">' + "<span></span>" * 22 + "</div>",
      unsafe_allow_html=True,
  )

import shutil
import subprocess
import base64

try:
  import requests
except ImportError:
  requests = None

FILE_PATH = "Blu-ray_Collection_Tracker_v5.0.xlsx"
STAR_OPTIONS = ["⭐", "⭐⭐", "⭐⭐⭐", "⭐⭐⭐⭐", "⭐⭐⭐⭐⭐"]


def try_recalculate(filepath):
  """Best-effort recalculation of formulas (Dashboard stats, Franchise %,
  Consensus scores) using LibreOffice, since openpyxl can't compute formulas
  itself and wipes their cached values on save. Silently does nothing if
  LibreOffice ('soffice') isn't installed -- Excel will recalculate those
  cells automatically the next time you open the file there anyway."""
  soffice = shutil.which("soffice") or shutil.which("libreoffice")
  if not soffice:
    return False
  try:
    out_dir = "recalc_tmp"
    subprocess.run(
        [soffice, "--headless", "--convert-to", "xlsx", "--outdir", out_dir, filepath],
        capture_output=True,
        timeout=30,
        check=True,
    )
    converted = f"{out_dir}/{filepath.rsplit('/', 1)[-1]}"
    import os
    if os.path.exists(converted):
      shutil.move(converted, filepath)
      shutil.rmtree(out_dir, ignore_errors=True)
      return True
  except Exception:
    pass
  return False


def github_sync_configured():
  """True if GitHub persistence secrets have been set up (only relevant when
  deployed to the cloud -- harmless no-op when running locally)."""
  try:
    return "github" in st.secrets and "token" in st.secrets["github"]
  except Exception:
    return False


def tmdb_configured():
  """True if a free TMDb API key has been added to secrets. TMDb powers
  auto-fill on Add and the plot line on Tonight's Pick -- everything using
  it degrades gracefully to 'not shown' when this is False."""
  try:
    return "tmdb" in st.secrets and "api_key" in st.secrets["tmdb"]
  except Exception:
    return False


@st.cache_data(ttl=3600, show_spinner=False)
def tmdb_lookup_cached(title, year=None):
  """Searches TMDb for a title and pulls genre/director/runtime/plot/release
  year. Cached for an hour since this data barely changes. Returns None on
  any failure or no-match -- always safe to fall back to manual entry."""
  if requests is None or not tmdb_configured():
    return None
  try:
    api_key = st.secrets["tmdb"]["api_key"]
    params = {"api_key": api_key, "query": title}
    if year:
      params["year"] = int(year)
    search_resp = requests.get(
        "https://api.themoviedb.org/3/search/movie", params=params, timeout=8
    )
    results = search_resp.json().get("results", [])
    if not results:
      return None

    movie_id = results[0]["id"]
    detail_resp = requests.get(
        f"https://api.themoviedb.org/3/movie/{movie_id}",
        params={"api_key": api_key, "append_to_response": "credits"},
        timeout=8,
    )
    detail = detail_resp.json()

    genres = ", ".join(g["name"] for g in detail.get("genres", []))
    crew = detail.get("credits", {}).get("crew", [])
    director = next((c["name"] for c in crew if c.get("job") == "Director"), "")
    release_date = detail.get("release_date", "") or ""
    release_year = int(release_date[:4]) if release_date[:4].isdigit() else None

    return {
        "id": movie_id,
        "genre": genres,
        "director": director,
        "runtime": detail.get("runtime"),
        "plot": detail.get("overview", ""),
        "year": release_year,
        "release_date": release_date,
    }
  except Exception:
    return None


@st.cache_data(ttl=86400, show_spinner=False)
def tmdb_recommendations_cached(movie_id):
  """Films similar to the given TMDb movie ID. Returns a list of
  {'title':..., 'year':..., 'poster_path':...} dicts, or [] on failure --
  used to power the 'Because you loved...' Wishlist suggestions."""
  if requests is None or not tmdb_configured():
    return []
  try:
    api_key = st.secrets["tmdb"]["api_key"]
    resp = requests.get(
        f"https://api.themoviedb.org/3/movie/{movie_id}/recommendations",
        params={"api_key": api_key},
        timeout=8,
    )
    results = resp.json().get("results", [])
    out = []
    for r in results:
      release_date = r.get("release_date", "") or ""
      out.append({
          "title": r.get("title", ""),
          "year": int(release_date[:4]) if release_date[:4].isdigit() else None,
          "poster_path": r.get("poster_path"),
      })
    return out
  except Exception:
    return []


@st.cache_data(ttl=86400, show_spinner=False)
def tmdb_release_date_cached(title, year=None, region="GB"):
  """Used by 'On This Day' -- finds the film via search, then pulls the
  actual UK release date from TMDb's per-country release_dates endpoint
  (search's own release_date field is the primary/global date, usually US,
  which is why Transformers: The Movie was showing its June 1986 US date
  instead of its December 1986 UK one). Prefers a Theatrical release entry,
  falling back to Limited Theatrical, then Premiere, then whatever's there.
  Returns {'release_date': ..., 'poster_path': ...} -- the poster comes free
  from the same search call, no extra request needed. Cached for a day since
  a whole-collection scan is comparatively expensive."""
  if requests is None or not tmdb_configured():
    return None
  try:
    api_key = st.secrets["tmdb"]["api_key"]
    params = {"api_key": api_key, "query": title}
    if year:
      params["year"] = int(year)
    resp = requests.get(
        "https://api.themoviedb.org/3/search/movie", params=params, timeout=8
    )
    results = resp.json().get("results", [])
    if not results:
      return None
    movie_id = results[0]["id"]
    fallback_date = results[0].get("release_date") or None
    poster_path = results[0].get("poster_path")

    rd_resp = requests.get(
        f"https://api.themoviedb.org/3/movie/{movie_id}/release_dates",
        params={"api_key": api_key},
        timeout=8,
    )
    countries = rd_resp.json().get("results", [])
    region_entry = next((c for c in countries if c.get("iso_3166_1") == region), None)

    release_date = fallback_date
    if region_entry:
      dates = region_entry.get("release_dates", [])
      # TMDb release types: 1=Premiere, 2=Theatrical (limited), 3=Theatrical,
      # 4=Digital, 5=Physical, 6=TV -- prefer an actual cinema release date.
      for preferred_type in (3, 2, 1):
        match = next((d for d in dates if d.get("type") == preferred_type), None)
        if match and match.get("release_date"):
          release_date = match["release_date"][:10]
          break
      else:
        dated_entries = sorted(d["release_date"] for d in dates if d.get("release_date"))
        if dated_entries:
          release_date = dated_entries[0][:10]

    return {"release_date": release_date, "poster_path": poster_path}
  except Exception:
    pass
  return None


def sync_to_github(filepath):
  """Commits the workbook back to your GitHub repo so edits survive a
  Streamlit Cloud restart or redeploy -- the container's own disk isn't
  guaranteed to persist, but your repo is. Does nothing if GitHub secrets
  aren't configured (e.g. running locally on your PC)."""
  if requests is None or not github_sync_configured():
    return False
  try:
    cfg = st.secrets["github"]
    token = cfg["token"]
    repo = cfg["repo"]  # e.g. "yourusername/blu-ray-hub"
    branch = cfg.get("branch", "main")
    path = cfg.get("path", filepath)

    api_url = f"https://api.github.com/repos/{repo}/contents/{path}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github+json",
    }

    # GitHub requires the current file's sha to update it
    get_resp = requests.get(api_url, headers=headers, params={"ref": branch}, timeout=10)
    sha = get_resp.json().get("sha") if get_resp.status_code == 200 else None

    with open(filepath, "rb") as f:
      content_b64 = base64.b64encode(f.read()).decode("utf-8")

    payload = {
        "message": f"Update collection — {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}",
        "content": content_b64,
        "branch": branch,
    }
    if sha:
      payload["sha"] = sha

    put_resp = requests.put(api_url, headers=headers, json=payload, timeout=15)
    return put_resp.status_code in (200, 201)
  except Exception:
    return False


def save_and_sync(filepath):
  """The full save pipeline: recalculate formulas, then push to GitHub for
  durable cloud storage. Call this right after book.save(filepath)."""
  try_recalculate(filepath)
  synced = sync_to_github(filepath)
  if github_sync_configured() and not synced:
    st.warning(
        "Saved locally, but couldn't sync to GitHub — your change may not "
        "survive the app restarting. Check your GitHub token/secrets.",
        icon="⚠️",
    )


def decode_barcode(image_bytes):
  """Reads a UPC/EAN barcode from a photo. Returns the code as a string, or
  None if nothing could be decoded or the pyzbar library isn't installed."""
  try:
    from pyzbar.pyzbar import decode as zbar_decode
    from PIL import Image
    import io
    image = Image.open(io.BytesIO(image_bytes))
    results = zbar_decode(image)
    if results:
      return results[0].data.decode("utf-8")
  except ImportError:
    st.error(
        "Barcode scanning needs the pyzbar package installed "
        "(see requirements.txt / packages.txt)."
    )
  except Exception:
    pass
  return None


def clean_title_for_search(title):
  """Strips retailer-catalog junk like '(Blu-ray Disc)' or '[4K UHD Steelbook]'
  off a title before using it as a search query -- UPC product titles are
  full of this and it breaks a TMDb match otherwise."""
  cleaned = re.sub(
      r"[\(\[][^\)\]]*?(blu-?ray|dvd|4k|uhd|disc|steelbook|edition|region\s?[a-z0-9]+)[^\)\]]*?[\)\]]",
      "",
      title,
      flags=re.IGNORECASE,
  )
  return cleaned.strip(" -–—")


def lookup_barcode(code):
  """Looks up a UPC/EAN code via UPCitemdb's free trial endpoint (no API key
  needed). Returns {'title': ..., 'year': int or None} on a match, or None
  if nothing was found or the request failed -- always safe to fall back to
  manual entry. UPCitemdb doesn't have a dedicated release-year field, so
  the year is a best-effort guess pulled out of the title/description text,
  and may well be missing or wrong -- worth double-checking either way."""
  if requests is None:
    return None
  try:
    resp = requests.get(
        "https://api.upcitemdb.com/prod/trial/lookup",
        params={"upc": code},
        timeout=10,
    )
    data = resp.json()
    items = data.get("items", [])
    if not items:
      return None

    item = items[0]
    title = item.get("title", "")
    description = item.get("description", "") or ""

    year = None
    year_match = re.search(r"\b(19[5-9]\d|20[0-4]\d)\b", f"{title} {description}")
    if year_match:
      year = int(year_match.group(1))

    return {"title": title, "year": year}
  except Exception:
    pass
  return None


# ----------------------------- helpers -----------------------------------

def safe_year(value):
  """Turn a possibly-float/possibly-string year into a clean display string."""
  if pd.isna(value):
    return ""
  try:
    return str(int(float(value)))
  except (ValueError, TypeError):
    return str(value)


def star_index_from_rating(rating_val, default=2):
  """Map a stored rating (either '⭐⭐⭐' or a number) to a 0-4 index."""
  if pd.isna(rating_val):
    return default
  rating_str = str(rating_val).strip()
  if "⭐" in rating_str:
    return max(0, min(4, rating_str.count("⭐") - 1))
  try:
    return max(0, min(4, int(float(rating_val)) - 1))
  except (ValueError, TypeError):
    return default


def is_watched(value):
  return str(value).strip().lower() in ["yes", "y", "true"]


def watched_display(value):
  """Turn a raw Watched cell (which may be NaN/blank) into 'Yes' or 'No'."""
  return "Yes" if is_watched(value) else "No"


def format_badge_html(format_value):
  """Small colored pill for Format — violet for Blu-ray, teal for 4K UHD,
  matching the app's laser-blue / disc-shimmer palette."""
  if pd.isna(format_value):
    return ""
  label = str(format_value).strip()
  lower = label.lower()
  if "4k" in lower:
    css_class = "uhd4k"
  elif "blu" in lower:
    css_class = "bluray"
  else:
    css_class = "other"
  return f'<span class="format-badge {css_class}">{label}</span>'


def retailer_search_links(title):
  """Search-page links for a title, one tap instead of typing it out each
  time. HMV/Zavvi don't have a documented simple search URL, so those go via
  a site-scoped Google search (same idea as the IMDb/RT lookup links already
  in your Ratings sheet); Amazon's search URL format is well-established so
  that one links directly."""
  quoted = urllib.parse.quote_plus(title)
  return {
      "HMV": f"https://www.google.com/search?q=site:hmv.com+{quoted}",
      "Zavvi": f"https://www.google.com/search?q=site:zavvi.com+{quoted}",
      "Amazon": f"https://www.amazon.co.uk/s?k={quoted}",
  }


def rating_stars_display(rating_val):
  """Turn a raw Rating cell into a star string, or a neutral placeholder if unset."""
  if pd.isna(rating_val) or str(rating_val).strip() == "":
    return "Not rated yet"
  rating_str = str(rating_val).strip()
  if "⭐" in rating_str:
    return rating_str
  try:
    return STAR_OPTIONS[max(0, min(4, int(float(rating_val)) - 1))]
  except (ValueError, TypeError):
    return "Not rated yet"


def curated_browse_view(df):
  """A phone-friendly slice of the collection: just the columns someone
  browsing would actually want to see, with clean values -- not the full
  26-column spreadsheet."""
  wanted_cols = ["Title", "Year", "Format", "Watched", "Rating (1-5)", "Date Watched"]
  cols_present = [c for c in wanted_cols if c in df.columns]
  view = df[cols_present].copy()

  if "Year" in view.columns:
    view["Year"] = view["Year"].apply(safe_year)
  if "Watched" in view.columns:
    view["Watched"] = view["Watched"].apply(watched_display)
  if "Rating (1-5)" in view.columns:
    view["Rating (1-5)"] = view["Rating (1-5)"].apply(rating_stars_display)
    view = view.rename(columns={"Rating (1-5)": "Rating"})
  if "Date Watched" in view.columns:
    view["Date Watched"] = view["Date Watched"].apply(format_date_watched)
    view = view.rename(columns={"Date Watched": "Last Watched"})

  return view.reset_index(drop=True)


def style_format_column(df):
  """Colors the Format column violet for Blu-ray / teal for 4K UHD, matching
  the app's palette. Returns a Styler; falls back to the plain DataFrame if
  the installed pandas version doesn't support cell styling."""
  if "Format" not in df.columns:
    return df

  def _color(value):
    lower = str(value).lower()
    if "4k" in lower:
      return "color: #2DD4BF; font-weight: 700;"
    if "blu" in lower:
      return "color: #A79CFF; font-weight: 700;"
    return ""

  try:
    return df.style.map(_color, subset=["Format"])
  except AttributeError:
    try:
      return df.style.applymap(_color, subset=["Format"])
    except Exception:
      return df


def find_row_by_id(sheet, id_column, target_id, header_row=4):
  """Scan a sheet for the row whose id_column matches target_id."""
  for r in range(header_row, sheet.max_row + 1):
    if str(sheet.cell(row=r, column=id_column).value).strip() == str(
        target_id
    ).strip():
      return r
  return None


def get_col_index(sheet, header_name, header_row=4):
  """Finds a column's position by its header text instead of a hardcoded
  number, so writes stay correct even if the sheet's columns get reordered."""
  for c in range(1, sheet.max_column + 1):
    if str(sheet.cell(row=header_row, column=c).value).strip() == header_name:
      return c
  return None


def ensure_column(sheet, header_name, header_row=4):
  """Like get_col_index, but creates the column (writing the header text)
  if it doesn't exist yet -- used for 'Date Watched', which isn't part of
  the original spreadsheet."""
  col = get_col_index(sheet, header_name, header_row)
  if col:
    return col
  new_col = sheet.max_column + 1
  sheet.cell(row=header_row, column=new_col, value=header_name)
  return new_col


def format_date_watched(value):
  """Turns a stored Date Watched value into a clean dd/mm/yyyy string, or
  '' if it's blank/unset."""
  if pd.isna(value) or str(value).strip() == "":
    return ""
  if isinstance(value, (datetime.date, datetime.datetime)):
    return value.strftime("%d/%m/%Y")
  try:
    return pd.to_datetime(value).strftime("%d/%m/%Y")
  except Exception:
    return str(value)


def add_to_collection(fields):
  """Adds a new film to the Collection sheet. fields is a dict of
  {header_name: value} -- only columns that exist in the sheet get written,
  so this stays safe even if the sheet's structure changes."""
  book = openpyxl.load_workbook(FILE_PATH)
  sheet = book["Collection"]
  next_row = sheet.max_row + 1

  film_id_col = get_col_index(sheet, "Film ID") or 1
  sheet.cell(row=next_row, column=film_id_col, value=next_film_id(collection_df))

  for header_name, value in fields.items():
    if value in (None, ""):
      continue
    col = get_col_index(sheet, header_name)
    if col:
      sheet.cell(row=next_row, column=col, value=value)

  book.save(FILE_PATH)
  save_and_sync(FILE_PATH)


def delete_from_collection(film_id):
  """Removes a film from the Collection sheet entirely."""
  book = openpyxl.load_workbook(FILE_PATH)
  sheet = book["Collection"]
  target_row = find_row_by_id(sheet, id_column=1, target_id=film_id)
  if not target_row:
    return False
  sheet.delete_rows(target_row)
  book.save(FILE_PATH)
  save_and_sync(FILE_PATH)
  return True


def stars_to_number(star_string):
  """Turn '⭐⭐⭐' into 3. Falls back to None if it can't be parsed."""
  count = str(star_string).count("⭐")
  return count if count > 0 else None


def save_collection_update(film_id, watched_value, rating_value, date_watched=None):
  """Writes Watched + Rating to the Collection sheet, and mirrors the
  numeric rating into the Ratings sheet's 'Your Rating /5' column so the
  two stay in sync. date_watched is only written when the film is marked
  Watched -- pass None to leave it untouched (e.g. when unmarking)."""
  book = openpyxl.load_workbook(FILE_PATH)
  coll_sheet = book["Collection"]
  target_row = find_row_by_id(coll_sheet, id_column=1, target_id=film_id)
  if not target_row:
    return False

  coll_sheet.cell(row=target_row, column=7, value=watched_value)
  coll_sheet.cell(row=target_row, column=8, value=rating_value)

  if is_watched(watched_value) and date_watched:
    date_col = ensure_column(coll_sheet, "Date Watched")
    date_str = date_watched.strftime("%Y-%m-%d") if hasattr(date_watched, "strftime") else str(date_watched)
    coll_sheet.cell(row=target_row, column=date_col, value=date_str)

  if "Ratings" in book.sheetnames:
    ratings_sheet = book["Ratings"]
    ratings_row = find_row_by_id(ratings_sheet, id_column=1, target_id=film_id)
    numeric_rating = stars_to_number(rating_value)
    if ratings_row and numeric_rating is not None:
      ratings_sheet.cell(row=ratings_row, column=4, value=numeric_rating)

  book.save(FILE_PATH)
  save_and_sync(FILE_PATH)
  return True


# Load data with caching
@st.cache_data
def load_data():
  collection_df = pd.read_excel(FILE_PATH, sheet_name="Collection", skiprows=3)
  wishlist_df = pd.read_excel(FILE_PATH, sheet_name="Wishlist", skiprows=3)
  franchise_df = pd.read_excel(
      FILE_PATH, sheet_name="Franchise Tracker", skiprows=3
  )
  actors_df = pd.read_excel(FILE_PATH, sheet_name="Actors", skiprows=3)
  directors_df = pd.read_excel(FILE_PATH, sheet_name="Directors", skiprows=3)
  awards_df = pd.read_excel(FILE_PATH, sheet_name="Awards", skiprows=3)
  ratings_df = pd.read_excel(FILE_PATH, sheet_name="Ratings", skiprows=3)
  return (
      collection_df,
      wishlist_df,
      franchise_df,
      actors_df,
      directors_df,
      awards_df,
      ratings_df,
  )


def split_multi_value_counts(series):
  """Split comma-separated cells (e.g. 'Action, Comedy') and count each value."""
  exploded = (
      series.dropna()
      .astype(str)
      .str.split(",")
      .explode()
      .str.strip()
  )
  exploded = exploded[exploded != ""]
  return exploded.value_counts()


def next_film_id(collection_df):
  """Generate the next Film ID in the F0001-style sequence."""
  existing = (
      collection_df["Film ID"]
      .dropna()
      .astype(str)
      .str.extract(r"F(\d+)", expand=False)
      .dropna()
      .astype(int)
  )
  next_num = (existing.max() + 1) if not existing.empty else 1
  return f"F{next_num:04d}"


def add_to_wishlist(title, priority=3, target_price=None, notes=""):
  """Adds a film to the Wishlist sheet. Shared by the Wishlist form and the
  'Add to Wishlist' shortcut on a zero-result Search."""
  book = openpyxl.load_workbook(FILE_PATH)
  sheet = book["Wishlist"]
  next_row = sheet.max_row + 1
  sheet.cell(row=next_row, column=1, value=title)
  sheet.cell(row=next_row, column=3, value=priority)
  if target_price is not None:
    sheet.cell(row=next_row, column=4, value=target_price)
  sheet.cell(row=next_row, column=8, value="No")
  sheet.cell(row=next_row, column=9, value=notes)
  book.save(FILE_PATH)
  save_and_sync(FILE_PATH)


try:
  (
      collection_df,
      wishlist_df,
      franchise_df,
      actors_df,
      directors_df,
      awards_df,
      ratings_df,
  ) = load_data()

  valid_collection = collection_df.dropna(subset=["Title"]).copy()
  valid_wishlist = wishlist_df.dropna(subset=["Title"]).copy()
  total_collection = len(valid_collection)
  total_wishlist = len(valid_wishlist)
  total_directors = len(directors_df.dropna(subset=["Director"]))

  # Detect an optional price/cost column so the value stat only shows if it exists
  price_col = next(
      (c for c in ["Price", "Price (£)", "Cost", "Cost (£)"]
       if c in valid_collection.columns),
      None,
  )

  # App Header
  st.markdown(
      "<div class='app-title'>🎬 Conor's Blu-ray Hub</div>",
      unsafe_allow_html=True,
  )
  st.markdown(
      "<div class='app-subtitle'>Mobile Media Companion</div>",
      unsafe_allow_html=True,
  )
  if github_sync_configured():
    st.markdown(
        "<p style='text-align:center; font-size:0.75em; color:var(--teal); "
        "margin-top:4px;'>☁️ Cloud sync active</p>",
        unsafe_allow_html=True,
    )
  else:
    st.markdown(
        "<p style='text-align:center; font-size:0.75em; opacity:0.5; "
        "margin-top:4px;'>💾 Local storage only — edits may not survive a restart</p>",
        unsafe_allow_html=True,
    )
  film_divider()

  # Navigation Tabs
  app_mode = st.tabs([
      "🏠 Home",
      "🔍 Search",
      "📚 Update",
      "🛒 Wishlist",
      "📊 Stats",
      "🏆 Extras",
      "📅 On This Day",
  ])

  # --- TAB 1: HOME & RANDOM PICKER ---
  with app_mode[0]:
    st.markdown("### 📊 Library Stats")
    col1, col2, col3 = st.columns(3)
    with col1:
      st.metric(label="Owned", value=total_collection)
    with col2:
      st.metric(label="Wishlist", value=total_wishlist)
    with col3:
      st.metric(label="Directors", value=total_directors)

    if price_col:
      total_value = pd.to_numeric(
          valid_collection[price_col], errors="coerce"
      ).sum()
      st.markdown(
          f"<p style='text-align: center; color: gray; margin-top: -8px;'>"
          f"💰 Estimated collection value: <b>£{total_value:,.2f}</b></p>",
          unsafe_allow_html=True,
      )

    film_divider()
    st.markdown("### 🎲 Movie Night Decider")

    rcol1, rcol2 = st.columns(2)
    with rcol1:
      formats = (
          ["All"]
          + sorted(valid_collection["Format"].dropna().unique().tolist())
          if "Format" in valid_collection.columns
          else ["All"]
      )
      selected_format_filter = st.selectbox(
          "Format", formats, key="rand_format"
      )

    with rcol2:
      selected_watched_filter = st.selectbox(
          "Status", ["All", "Unwatched Only", "Watched Only"]
      )

    double_feature = st.toggle("🎬 Double Feature (pick 2 films)", value=False)

    pool_df = valid_collection.copy()
    if selected_format_filter != "All" and "Format" in pool_df.columns:
      pool_df = pool_df[pool_df["Format"] == selected_format_filter]

    if "Watched" in pool_df.columns:
      if selected_watched_filter == "Unwatched Only":
        pool_df = pool_df[~pool_df["Watched"].apply(is_watched)]
      elif selected_watched_filter == "Watched Only":
        pool_df = pool_df[pool_df["Watched"].apply(is_watched)]

    st.markdown(
        f"<p style='text-align: center; font-size: 0.9em; color:"
        f" gray;'>Available pool: <b>{len(pool_df)}</b> films</p>",
        unsafe_allow_html=True,
    )

    pick_label = "🎰 Pick a Double Feature!" if double_feature else "🎰 Pick a Random Movie!"
    if st.button(pick_label, type="primary"):
      needed = 2 if double_feature else 1
      if len(pool_df) >= needed:
        picked = pool_df.sample(n=needed)
        st.session_state["picked_movie_ids"] = picked.get(
            "Film ID", picked.index.to_series()
        ).astype(str).tolist()
      else:
        st.warning("Not enough movies match your filters.")
        st.session_state.pop("picked_movie_ids", None)

    if "picked_movie_ids" in st.session_state and "Film ID" in valid_collection.columns:
      picked_rows = valid_collection[
          valid_collection["Film ID"].astype(str).isin(
              st.session_state["picked_movie_ids"]
          )
      ]

      for _, picked_movie in picked_rows.iterrows():
        p_title = picked_movie.get("Title", "Unknown")
        p_year_str = safe_year(picked_movie.get("Year", ""))
        p_format_badge = format_badge_html(picked_movie.get("Format", "Blu-ray"))
        p_genre = picked_movie.get("Genre", "N/A")
        p_director = picked_movie.get("Director", "N/A")
        p_id = picked_movie.get("Film ID", "")
        p_watched_raw = picked_movie.get("Watched", "No")
        p_watched = watched_display(p_watched_raw)
        p_rating_display = rating_stars_display(picked_movie.get("Rating (1-5)"))
        p_date_watched = format_date_watched(picked_movie.get("Date Watched")) if "Date Watched" in picked_movie.index else ""
        p_last_watched_html = f"<br><b>Last watched:</b> {p_date_watched}" if p_date_watched else ""

        p_plot_html = ""
        if tmdb_configured():
          tmdb_pick_data = tmdb_lookup_cached(p_title, picked_movie.get("Year"))
          if tmdb_pick_data and tmdb_pick_data.get("plot"):
            p_plot_html = (
                f'<p style="font-size: 0.85em; margin-top: 10px; opacity: 0.75; '
                f'font-style: italic;">{tmdb_pick_data["plot"]}</p>'
            )

        st.markdown(
            f"""
                <div class="winner-box">
                    <h3 style="color: #F0A83B; margin-bottom: 2px;">🎉 Tonight's Pick:</h3>
                    <h2 style="margin: 0px;">{p_title} ({p_year_str})</h2>
                    <p style="font-size: 0.95em; margin-top: 8px; opacity: 0.85;">
                        {p_format_badge} &nbsp;<b>Director:</b> {p_director}<br>
                        <b>Genre:</b> {p_genre} | <b>Watched:</b> {p_watched}<br>
                        <b>Rating:</b> {p_rating_display}{p_last_watched_html}
                    </p>
                    {p_plot_html}
                </div>
                """,
            unsafe_allow_html=True,
        )

        with st.form(f"quick_picker_update_form_{p_id}"):
          st.markdown(f"**Quick Update — {p_title}:**")
          uq_col1, uq_col2 = st.columns(2)
          with uq_col1:
            quick_watched = st.selectbox(
                "Watched",
                ["Yes", "No"],
                index=0 if is_watched(p_watched) else 1,
                key=f"qp_watched_{p_id}",
            )
          with uq_col2:
            quick_stars = st.selectbox(
                "Rating",
                STAR_OPTIONS,
                index=2,
                key=f"qp_stars_{p_id}",
            )
          quick_date_watched = st.date_input(
              "Date watched (optional)",
              value=pd.to_datetime(picked_movie.get("Date Watched")).date()
                    if "Date Watched" in picked_movie.index and pd.notna(picked_movie.get("Date Watched"))
                    else datetime.date.today(),
              key=f"qp_date_{p_id}",
          )

          if st.form_submit_button("💾 Save to Excel"):
            if save_collection_update(p_id, quick_watched, quick_stars, quick_date_watched):
              st.cache_data.clear()
              st.success(f"Updated '{p_title}' successfully!")
              if quick_stars == STAR_OPTIONS[-1]:
                st.balloons()
              st.rerun()
            else:
              st.error("Couldn't find that film in the sheet to update.")

  # --- TAB 2: SEARCH ---
  with app_mode[1]:
    st.subheader("🔍 Search Library")
    scol1, scol2 = st.columns([2, 1])
    with scol1:
      query = st.text_input("Keyword search:")
    with scol2:
      search_column = st.selectbox(
          "In column", ["All"] + collection_df.columns.tolist()
      )

    if query:
      if search_column == "All":
        mask = collection_df.astype(str).apply(
            lambda x: x.str.contains(query, case=False, na=False)
        ).any(axis=1)
      else:
        mask = (
            collection_df[search_column]
            .astype(str)
            .str.contains(query, case=False, na=False)
        )
      results = collection_df[mask]
      if len(results) == 0:
        st.warning(f"No matches for '{query}' in your collection.")
        if st.button(f"🛒 Add '{query}' to Wishlist", key="search_add_to_wishlist"):
          add_to_wishlist(query)
          st.cache_data.clear()
          st.success(f"Added '{query}' to your Wishlist!")
          st.rerun()
      else:
        st.success(f"Found {len(results)} matches:")
        st.dataframe(style_format_column(curated_browse_view(results)), use_container_width=True, hide_index=True)
    else:
      st.info("Type a keyword above to find a film.")

  # --- TAB 3: COLLECTION & UPDATE ---
  with app_mode[2]:
    st.subheader("📚 Update Status & Rating")

    valid_collection_sorted = valid_collection.sort_values(by="Film ID")

    # Build a label -> Film ID map so selection never depends on re-parsing text
    option_to_id = {}
    for _, row in valid_collection_sorted.iterrows():
      title = row.get("Title", "")
      year_str = safe_year(row.get("Year", ""))
      label = f"{title} ({year_str})" if year_str else f"{title}"
      # De-duplicate identical labels (e.g. two copies of the same title/year)
      base_label, suffix = label, 2
      while base_label in option_to_id:
        base_label = f"{label} [{suffix}]"
        suffix += 1
      option_to_id[base_label] = row.get("Film ID", "")

    selected_option = st.selectbox(
        "Select Movie:", ["-- Select --"] + list(option_to_id.keys())
    )

    if selected_option != "-- Select --":
      selected_id = option_to_id[selected_option]
      movie_row = valid_collection_sorted[
          valid_collection_sorted["Film ID"].astype(str) == str(selected_id)
      ].iloc[0]

      selected_movie = movie_row["Title"]
      current_watched = movie_row.get("Watched", "No")
      current_rating_val = movie_row.get("Rating (1-5)", 3)
      default_star_index = star_index_from_rating(current_rating_val)

      with st.form("update_movie_form"):
        st.markdown(f"**Editing:** {selected_movie}")
        col_a, col_b = st.columns(2)
        with col_a:
          watched_status = st.selectbox(
              "Watched",
              ["Yes", "No"],
              index=0 if is_watched(current_watched) else 1,
          )
        with col_b:
          new_rating_stars = st.selectbox(
              "Stars",
              STAR_OPTIONS,
              index=default_star_index,
          )
        current_date_watched = movie_row.get("Date Watched") if "Date Watched" in movie_row.index else None
        new_date_watched = st.date_input(
            "Date watched (optional)",
            value=pd.to_datetime(current_date_watched).date() if pd.notna(current_date_watched) else datetime.date.today(),
        )

        if st.form_submit_button("💾 Save Changes"):
          if save_collection_update(selected_id, watched_status, new_rating_stars, new_date_watched):
            st.cache_data.clear()
            st.success(f"Saved '{selected_movie}'!")
            if new_rating_stars == STAR_OPTIONS[-1]:
              st.balloons()
            st.rerun()
          else:
            st.error("Couldn't find that film in the sheet to update.")

      with st.expander("🗑️ Remove this film from my collection"):
        st.warning(f"This permanently deletes '{selected_movie}' from your Collection sheet.")
        confirm_delete = st.checkbox("Yes, I'm sure", key=f"confirm_del_{selected_id}")
        if st.button("Remove permanently", key=f"remove_btn_{selected_id}", disabled=not confirm_delete):
          if delete_from_collection(selected_id):
            st.cache_data.clear()
            st.success(f"Removed '{selected_movie}' from your collection.")
            st.rerun()
          else:
            st.error("Couldn't find that film in the sheet to remove.")

    film_divider()
    st.markdown("**➕ Add a Film**")

    add_tab1, add_tab2 = st.tabs(["Type it in", "📷 Scan barcode"])

    with add_tab1:
      st.markdown("**1. Title & year**")
      mcol1, mcol2 = st.columns([2, 1])
      with mcol1:
        manual_title = st.text_input("Title *", key="manual_title_input")
      with mcol2:
        manual_year = st.number_input(
            "Year", min_value=1900, max_value=2100,
            value=datetime.date.today().year, step=1, key="manual_year_input",
        )

      if tmdb_configured():
        if st.button("🔎 Look up on TMDb", key="manual_tmdb_btn"):
          if manual_title:
            result = tmdb_lookup_cached(manual_title, manual_year)
            st.session_state["manual_tmdb_result"] = result
            if not result:
              st.warning("No TMDb match — you can still fill in the details manually below.")
          else:
            st.warning("Enter a title first.")
      else:
        st.caption("Add a free TMDb API key to auto-fill genre, director, runtime, and plot here.")

      tmdb_prefill = st.session_state.get("manual_tmdb_result")
      if tmdb_prefill:
        st.success(
            f"Found on TMDb: {tmdb_prefill.get('genre') or '—'} · "
            f"dir. {tmdb_prefill.get('director') or '—'} · "
            f"{tmdb_prefill.get('runtime') or '—'} min"
        )

      st.markdown("**2. Confirm & add**")
      with st.form("manual_add_form", clear_on_submit=True):
        mcol3, mcol4 = st.columns(2)
        with mcol3:
          m_format = st.selectbox("Format", ["Blu-ray", "4K UHD", "DVD", "Other"])
        with mcol4:
          m_runtime = st.number_input(
              "Runtime (min)", min_value=0, max_value=600,
              value=int(tmdb_prefill.get("runtime") or 0) if tmdb_prefill else 0,
              step=1,
          )
        m_genre = st.text_input("Genre", value=(tmdb_prefill.get("genre", "") if tmdb_prefill else ""))
        m_director = st.text_input("Director", value=(tmdb_prefill.get("director", "") if tmdb_prefill else ""))
        m_notes = st.text_area(
            "Notes / Plot", value=(tmdb_prefill.get("plot", "") if tmdb_prefill else ""), height=80,
        )

        if st.form_submit_button("💾 Add to Collection"):
          if manual_title:
            add_to_collection({
                "Title": manual_title,
                "Year": manual_year,
                "Format": m_format,
                "Genre": m_genre,
                "Director": m_director,
                "Runtime (min)": m_runtime if m_runtime else None,
                "Notes": m_notes,
                "Watched": "No",
            })
            st.session_state.pop("manual_tmdb_result", None)
            st.cache_data.clear()
            st.success(f"Added '{manual_title}' to your collection!")
            st.rerun()
          else:
            st.warning("Title is required.")

    with add_tab2:
      st.caption(
          "Tap below and choose 'Take Photo' — this opens your phone's real "
          "camera app (which defaults to the rear camera), not a browser "
          "preview. Works best with good lighting and the barcode filling "
          "most of the frame."
      )
      barcode_photo = st.file_uploader(
          "Scan barcode",
          type=["jpg", "jpeg", "png"],
          key="barcode_uploader",
      )

      if barcode_photo is not None:
        image_bytes = barcode_photo.getvalue()
        decoded_code = decode_barcode(image_bytes)
        if not decoded_code:
          st.error("Couldn't read a barcode in that photo — try again with the barcode closer and well-lit.")
        else:
          st.info(f"Barcode: {decoded_code}")
          upc_result = lookup_barcode(decoded_code)

          prefill_title = clean_title_for_search(upc_result.get("title", "")) if upc_result else ""
          prefill_year = upc_result.get("year") if upc_result else None

          # Chain a TMDb lookup off the (cleaned) UPC title to also pull
          # genre/director/runtime/plot, and a more reliable year if TMDb
          # has one -- the UPC database frequently doesn't.
          tmdb_result = None
          if prefill_title and tmdb_configured():
            tmdb_result = tmdb_lookup_cached(prefill_title, prefill_year)
            if tmdb_result and tmdb_result.get("year"):
              prefill_year = tmdb_result["year"]

          with st.form("barcode_add_form"):
            if upc_result:
              st.success("Found a match online — check it's right before adding:")
            else:
              st.warning("No online match for this barcode — enter the details manually.")

            if upc_result and not prefill_year:
              st.caption("⚠️ Couldn't confirm a release year — please check it.")
            if tmdb_result:
              st.caption(
                  f"TMDb: {tmdb_result.get('genre') or '—'} · "
                  f"dir. {tmdb_result.get('director') or '—'} · "
                  f"{tmdb_result.get('runtime') or '—'} min"
              )

            b_title = st.text_input("Title *", value=prefill_title)
            bcol1, bcol2 = st.columns(2)
            with bcol1:
              b_year = st.number_input(
                  "Year",
                  min_value=1900,
                  max_value=2100,
                  value=prefill_year if prefill_year else datetime.date.today().year,
                  step=1,
                  key="b_year",
              )
            with bcol2:
              b_format = st.selectbox("Format", ["Blu-ray", "4K UHD", "DVD", "Other"], key="b_format")
            b_genre = st.text_input("Genre", value=(tmdb_result.get("genre", "") if tmdb_result else ""))
            b_director = st.text_input("Director", value=(tmdb_result.get("director", "") if tmdb_result else ""))
            b_runtime = st.number_input(
                "Runtime (min)", min_value=0, max_value=600,
                value=int(tmdb_result.get("runtime") or 0) if tmdb_result else 0, step=1,
            )
            default_notes = f"Barcode: {decoded_code}"
            if tmdb_result and tmdb_result.get("plot"):
              default_notes = f"{tmdb_result['plot']}\n\nBarcode: {decoded_code}"
            b_notes = st.text_area("Notes / Plot", value=default_notes, height=80)

            if st.form_submit_button("💾 Add to Collection"):
              if b_title:
                add_to_collection({
                    "Title": b_title,
                    "Year": b_year,
                    "Format": b_format,
                    "Genre": b_genre,
                    "Director": b_director,
                    "Runtime (min)": b_runtime if b_runtime else None,
                    "Notes": b_notes,
                    "Watched": "No",
                })
                st.cache_data.clear()
                st.success(f"Added '{b_title}' to your collection!")
                st.rerun()
              else:
                st.warning("Title is required.")

    film_divider()
    st.markdown("**Your collection**")
    st.dataframe(
        style_format_column(curated_browse_view(valid_collection)),
        use_container_width=True,
        height=350,
        hide_index=True,
    )

  # --- TAB 4: WISHLIST ---
  with app_mode[3]:
    st.subheader("🛒 Wishlist")

    if tmdb_configured() and "Rating (1-5)" in valid_collection.columns:
      with st.expander("✨ Because you loved these..."):
        five_star = valid_collection.copy()
        five_star["_stars"] = five_star["Rating (1-5)"].apply(stars_to_number)
        five_star = five_star[five_star["_stars"] == 5].head(5)

        if five_star.empty:
          st.caption("Rate some films 5 stars from the Update tab and suggestions will show up here.")
        else:
          owned_titles = set(valid_collection["Title"].dropna().str.strip().str.lower())
          wishlist_titles = set(valid_wishlist["Title"].dropna().str.strip().str.lower()) if not valid_wishlist.empty else set()

          seen = set()
          suggestions = []
          for _, source_row in five_star.iterrows():
            source_lookup = tmdb_lookup_cached(source_row.get("Title", ""), source_row.get("Year"))
            if not source_lookup or not source_lookup.get("id"):
              continue
            for rec in tmdb_recommendations_cached(source_lookup["id"]):
              rec_title = rec.get("title", "")
              key = rec_title.strip().lower()
              if not rec_title or key in seen or key in owned_titles or key in wishlist_titles:
                continue
              seen.add(key)
              suggestions.append(rec)
            if len(suggestions) >= 10:
              break

          if not suggestions:
            st.caption("No new suggestions right now — try rating a few more films 5 stars.")
          else:
            for rec in suggestions[:10]:
              rcol1, rcol2 = st.columns([1, 3])
              with rcol1:
                if rec.get("poster_path"):
                  st.image(f"https://image.tmdb.org/t/p/w200{rec['poster_path']}")
              with rcol2:
                rec_label = f"{rec['title']} ({rec['year']})" if rec.get("year") else rec["title"]
                st.markdown(f"**{rec_label}**")
                if st.button("+ Add to Wishlist", key=f"rec_add_{rec['title']}"):
                  add_to_wishlist(rec["title"])
                  st.cache_data.clear()
                  st.success(f"Added '{rec['title']}' to your Wishlist!")
                  st.rerun()

    if not valid_wishlist.empty:
      # Deal alerts: cheapest found price has actually dropped to/below target
      has_price_cols = (
          "Target Price (£)" in valid_wishlist.columns
          and "Cheapest Found (£)" in valid_wishlist.columns
      )
      if has_price_cols:
        deals = valid_wishlist[
            valid_wishlist["Cheapest Found (£)"].notna()
            & valid_wishlist["Target Price (£)"].notna()
            & (valid_wishlist["Cheapest Found (£)"] <= valid_wishlist["Target Price (£)"])
        ]
        if not deals.empty:
          st.markdown("### 🔥 Deal Alerts")
          for _, deal_row in deals.iterrows():
            where = deal_row.get("Where", "")
            where_str = f" at {where}" if pd.notna(where) and where else ""
            st.success(
                f"**{deal_row.get('Title','')}** — found for "
                f"£{deal_row.get('Cheapest Found (£)'):.2f} "
                f"(target was £{deal_row.get('Target Price (£)'):.2f}){where_str}"
            )
          film_divider()

      for idx, row in valid_wishlist.iterrows():
        w_title = row.get("Title", "Untitled")
        w_priority = row.get("Priority (1-5)", "")
        w_target = row.get("Target Price (£)", None)
        w_cheapest = row.get("Cheapest Found (£)", None)

        price_bits = []
        if pd.notna(w_target):
          price_bits.append(f"Target: £{w_target:.2f}")
        if pd.notna(w_cheapest):
          price_bits.append(f"Cheapest found: £{w_cheapest:.2f}")
        price_line = " | ".join(price_bits) if price_bits else "No price tracked yet"

        wcol1, wcol2, wcol3 = st.columns([3, 1, 1])
        with wcol1:
          st.markdown(
              f"**{w_title}**  \n<span style='opacity:0.65;font-size:0.85em;'>"
              f"Priority: {w_priority} | {price_line}</span>",
              unsafe_allow_html=True,
          )
        with wcol2:
          buy_clicked = st.button("✅ Bought", key=f"buy_{idx}")
        with wcol3:
          delete_clicked = st.button("🗑️", key=f"del_{idx}")

        links = retailer_search_links(w_title)
        st.markdown(
            f"<span style='font-size:0.85em;'>🔗 Check price: "
            f"<a href='{links['HMV']}' target='_blank'>HMV</a> · "
            f"<a href='{links['Zavvi']}' target='_blank'>Zavvi</a> · "
            f"<a href='{links['Amazon']}' target='_blank'>Amazon</a></span>",
            unsafe_allow_html=True,
        )

        with st.expander("💷 Log a price check"):
          with st.form(f"price_form_{idx}", clear_on_submit=False):
            pcol1, pcol2 = st.columns(2)
            with pcol1:
              logged_price = st.number_input(
                  "Cheapest price found (£)",
                  min_value=0.0,
                  value=float(w_cheapest) if pd.notna(w_cheapest) else 0.0,
                  step=0.50,
                  key=f"price_input_{idx}",
              )
            with pcol2:
              logged_where = st.text_input(
                  "Where",
                  value=str(row.get("Where", "") or ""),
                  key=f"where_input_{idx}",
              )
            if st.form_submit_button("Save price"):
              book = openpyxl.load_workbook(FILE_PATH)
              wish_sheet = book["Wishlist"]
              excel_row = idx + 5
              wish_sheet.cell(row=excel_row, column=5, value=logged_price)
              wish_sheet.cell(row=excel_row, column=6, value=logged_where)
              wish_sheet.cell(
                  row=excel_row, column=7,
                  value=pd.Timestamp.today().strftime("%Y-%m-%d"),
              )
              book.save(FILE_PATH)
              save_and_sync(FILE_PATH)
              st.cache_data.clear()
              st.success(f"Logged price for '{w_title}'.")
              st.rerun()

        if buy_clicked:
          book = openpyxl.load_workbook(FILE_PATH)
          coll_sheet = book["Collection"]
          wish_sheet = book["Wishlist"]

          next_row = coll_sheet.max_row + 1
          new_film_id = next_film_id(collection_df)
          coll_sheet.cell(row=next_row, column=1, value=new_film_id)
          coll_sheet.cell(row=next_row, column=2, value=w_title)
          coll_sheet.cell(row=next_row, column=7, value="No")  # Watched

          # Remove the row from Wishlist (Excel row = df index + header offset)
          wish_sheet.delete_rows(idx + 5)

          book.save(FILE_PATH)
          save_and_sync(FILE_PATH)
          st.cache_data.clear()
          st.success(f"Moved '{w_title}' to your Collection! 🎉")
          st.rerun()

        if delete_clicked:
          book = openpyxl.load_workbook(FILE_PATH)
          wish_sheet = book["Wishlist"]
          wish_sheet.delete_rows(idx + 5)
          book.save(FILE_PATH)
          save_and_sync(FILE_PATH)
          st.cache_data.clear()
          st.info(f"Removed '{w_title}' from wishlist.")
          st.rerun()
    else:
      st.info("Your wishlist is empty — add something below!")

    st.markdown("### Add to Wishlist")
    with st.form("wishlist_form", clear_on_submit=True):
      new_title = st.text_input("Title *")
      new_priority = st.slider("Priority (1-5)", 1, 5, 3)
      new_target_price = st.number_input(
          "Target Price (£)", min_value=0.0, value=15.00, step=0.50
      )
      new_notes = st.text_input("Notes")

      if st.form_submit_button("💾 Add Film"):
        if new_title:
          add_to_wishlist(new_title, new_priority, new_target_price, new_notes)
          st.cache_data.clear()
          st.success(f"Added '{new_title}'!")
          st.rerun()
        else:
          st.warning("Title is required.")

  # --- TAB 5: STATS ---
  with app_mode[4]:
    st.subheader("📊 Collection Insights")

    if "Genre" in valid_collection.columns:
      st.markdown("**By Genre**")
      genre_counts = split_multi_value_counts(valid_collection["Genre"]).head(15)
      st.bar_chart(genre_counts)

    if "Year" in valid_collection.columns:
      st.markdown("**By Decade**")
      decades = (
          pd.to_numeric(valid_collection["Year"], errors="coerce")
          .dropna()
          .apply(lambda y: f"{int(y) // 10 * 10}s")
      )
      st.bar_chart(decades.value_counts().sort_index())

    if "Director" in valid_collection.columns:
      st.markdown("**Top Directors**")
      top_directors = split_multi_value_counts(valid_collection["Director"]).head(10)
      st.bar_chart(top_directors)

    if "Watched" in valid_collection.columns:
      watched_count = valid_collection["Watched"].apply(is_watched).sum()
      unwatched_count = total_collection - watched_count
      st.markdown("**Watched vs Unwatched**")
      st.bar_chart(pd.Series(
          {"Watched": watched_count, "Unwatched": unwatched_count}
      ))

  # --- TAB 6: EXTRAS ---
  with app_mode[5]:
    st.subheader("🏆 Franchises, Awards & Ratings")
    st.caption(
        "Franchise % Complete and Consensus scores are Excel formulas. "
        "If you've just made an edit and they look blank, open the file in "
        "Excel once to refresh them (this app tries to do it automatically "
        "if LibreOffice is installed on your machine)."
    )
    sub_tab1, sub_tab2, sub_tab3, sub_tab4 = st.tabs(["Franchises", "Awards", "Ratings", "People"])
    with sub_tab1:
      if (
          "Collection" in valid_collection.columns
          and "Collection" in franchise_df.columns
          and "Target Films" in franchise_df.columns
      ):
        owned_by_set = valid_collection["Collection"].dropna().value_counts()
        sets_progress = franchise_df.dropna(subset=["Collection"]).copy()
        sets_progress["Owned"] = (
            sets_progress["Collection"].map(owned_by_set).fillna(0).astype(int)
        )
        sets_progress["Target"] = pd.to_numeric(
            sets_progress["Target Films"], errors="coerce"
        ).fillna(sets_progress["Owned"])
        sets_progress["Missing"] = (
            (sets_progress["Target"] - sets_progress["Owned"]).clip(lower=0).astype(int)
        )
        sets_progress = sets_progress.sort_values(
            by="Missing", ascending=False
        )

        st.markdown("**📦 Box Set Completion**")
        for _, srow in sets_progress.iterrows():
          owned = int(srow["Owned"])
          target = int(srow["Target"])
          missing = int(srow["Missing"])
          pct = min(1.0, owned / target) if target > 0 else 0.0
          label = f"{srow['Collection']} — {owned}/{target}"
          if missing == 0 and target > 0:
            label += " ✅ Complete!"
          else:
            label += f" ({missing} to go)"
          st.markdown(label)
          st.progress(pct)
        film_divider()

      st.dataframe(franchise_df, use_container_width=True, height=300)
    with sub_tab2:
      if {"Title", "Nominations"}.issubset(awards_df.columns):
        nominated = awards_df[
            pd.to_numeric(awards_df["Nominations"], errors="coerce").fillna(0) > 0
        ].copy()
        nominated["Wins"] = pd.to_numeric(nominated.get("Wins", 0), errors="coerce").fillna(0).astype(int)
        nominated["Nominations"] = pd.to_numeric(nominated["Nominations"], errors="coerce").fillna(0).astype(int)
        nominated = nominated.sort_values(by=["Wins", "Nominations"], ascending=False)

        if nominated.empty:
          st.info("None of your films show any Oscar nominations in this sheet.")
        else:
          st.caption(f"{len(nominated)} films with at least one Oscar nomination.")
          for _, arow in nominated.iterrows():
            title = arow.get("Title", "Unknown")
            year_str = safe_year(arow.get("Year", ""))
            wins = int(arow.get("Wins", 0))
            noms = int(arow.get("Nominations", 0))
            categories = arow.get("Winning Categories") if wins > 0 else arow.get("All Nominated Categories")
            categories = categories if pd.notna(categories) else ""

            if wins > 0:
              headline = f"🏆 **{title}** ({year_str}) — {wins} win{'s' if wins != 1 else ''}, {noms} nomination{'s' if noms != 1 else ''}"
            else:
              headline = f"🎖️ **{title}** ({year_str}) — {noms} nomination{'s' if noms != 1 else ''}"
            st.markdown(headline)
            if categories:
              st.caption(categories)
      else:
        st.dataframe(awards_df, use_container_width=True, height=300)
    with sub_tab3:
      if "Rating (1-5)" in valid_collection.columns:
        rated = valid_collection.copy()
        rated["_stars"] = rated["Rating (1-5)"].apply(stars_to_number)
        rated = rated[rated["_stars"].notna()].sort_values(by="_stars", ascending=False)

        if rated.empty:
          st.info("No films rated yet — rate some from the Update tab and they'll show up here.")
        else:
          st.caption(f"{len(rated)} rated film{'s' if len(rated) != 1 else ''}, highest first.")
          for _, rrow in rated.iterrows():
            title = rrow.get("Title", "Unknown")
            year_str = safe_year(rrow.get("Year", ""))
            stars = rating_stars_display(rrow.get("Rating (1-5)"))
            st.markdown(f"{stars} — **{title}** ({year_str})")
      else:
        st.dataframe(ratings_df, use_container_width=True, height=300)

    with sub_tab4:
      person_type = st.radio("Browse by", ["Director", "Actor"], horizontal=True, key="people_browse_type")
      people_df = directors_df if person_type == "Director" else actors_df
      name_col = "Director" if person_type == "Director" else "Actor"

      if name_col not in people_df.columns or "Films Owned" not in people_df.columns:
        st.info(f"No {person_type.lower()} data found in this sheet.")
      else:
        people_sorted = people_df.dropna(subset=[name_col]).sort_values(by="Films Owned", ascending=False)
        options = [
            f"{row[name_col]} ({int(row['Films Owned'])} films)"
            for _, row in people_sorted.iterrows()
        ]
        name_lookup = dict(zip(options, people_sorted[name_col]))

        selected_person_option = st.selectbox(f"Select a {person_type.lower()}:", ["-- Select --"] + options)

        if selected_person_option != "-- Select --":
          selected_name = name_lookup[selected_person_option]
          person_row = people_sorted[people_sorted[name_col] == selected_name].iloc[0]

          pcol1, pcol2, pcol3 = st.columns(3)
          with pcol1:
            st.metric("Films Owned", int(person_row.get("Films Owned", 0)))
          with pcol2:
            st.metric("Earliest", safe_year(person_row.get("Earliest Film")))
          with pcol3:
            st.metric("Latest", safe_year(person_row.get("Latest Film")))

          film_ids_raw = person_row.get("Film IDs", "")
          film_ids = [f.strip() for f in str(film_ids_raw).split(",") if f.strip()] if pd.notna(film_ids_raw) else []

          if film_ids and "Film ID" in valid_collection.columns:
            person_films = valid_collection[valid_collection["Film ID"].astype(str).isin(film_ids)]
            st.dataframe(
                style_format_column(curated_browse_view(person_films)),
                use_container_width=True,
                hide_index=True,
            )
          else:
            films_text = person_row.get("Films in Collection", "")
            if pd.notna(films_text):
              st.markdown(films_text)

  # --- TAB 7: ON THIS DAY ---
  with app_mode[6]:
    st.subheader("📅 On This Day")
    today = datetime.date.today()
    st.caption(f"Films in your collection first released on {today.strftime('%B %d')}, across the years.")

    if not tmdb_configured():
      st.info(
          "This needs a free TMDb API key in your app's secrets "
          "(the same [tmdb] section used for auto-fill on Add) — "
          "see the Extras tab caption for the general secrets pattern."
      )
    else:
      with st.spinner("Checking release dates across your collection — cached after the first time, so this gets fast."):
        titles_years = [
            (row.get("Title", ""), row.get("Year") if pd.notna(row.get("Year")) else None)
            for _, row in valid_collection.iterrows()
            if row.get("Title")
        ]

        def _check(title_year):
          title, year = title_year
          result = tmdb_release_date_cached(title, year)
          if not result or not result.get("release_date"):
            return None
          try:
            d = datetime.datetime.strptime(result["release_date"], "%Y-%m-%d").date()
          except ValueError:
            return None
          if d.month == today.month and d.day == today.day:
            return (title, d.year, result.get("poster_path"))
          return None

        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
          results = list(executor.map(_check, titles_years))

      matches = sorted((r for r in results if r), key=lambda x: x[1])

      if matches:
        for title, rel_year, poster_path in matches:
          pcol1, pcol2 = st.columns([1, 3])
          with pcol1:
            if poster_path:
              st.image(f"https://image.tmdb.org/t/p/w200{poster_path}")
          with pcol2:
            st.markdown(f"🎬 **{title}**")
            st.caption(f"Released {today.strftime('%B %d')}, {rel_year}")
      else:
        st.info("Nothing in your collection was first released on this date — check back tomorrow!")

except FileNotFoundError:
  st.error(
      f"Couldn't find '{FILE_PATH}'. Make sure the spreadsheet is in the "
      "same folder as this app."
  )
except Exception as e:
  st.error(f"Error: {e}")
