import random
import openpyxl
import pandas as pd
import streamlit as st

# Set up page configuration & layout
st.set_page_config(
    page_title="Conor's Blu-ray Hub", page_icon="🎬", layout="centered"
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

    /* Main container and background styling */
    .stApp {
        max-width: 600px;
        margin: 0 auto;
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
        background: linear-gradient(135deg, rgba(108, 92, 232, 0.18), rgba(45, 212, 191, 0.1));
        border: 2px solid var(--amber);
        padding: 20px;
        border-radius: 18px;
        text-align: center;
        margin-top: 15px;
        margin-bottom: 15px;
        box-shadow: 0 4px 20px rgba(108, 92, 232, 0.15);
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


def find_row_by_id(sheet, id_column, target_id, header_row=4):
  """Scan a sheet for the row whose id_column matches target_id."""
  for r in range(header_row, sheet.max_row + 1):
    if str(sheet.cell(row=r, column=id_column).value).strip() == str(
        target_id
    ).strip():
      return r
  return None


def stars_to_number(star_string):
  """Turn '⭐⭐⭐' into 3. Falls back to None if it can't be parsed."""
  count = str(star_string).count("⭐")
  return count if count > 0 else None


def save_collection_update(film_id, watched_value, rating_value):
  """Writes Watched + Rating to the Collection sheet, and mirrors the
  numeric rating into the Ratings sheet's 'Your Rating /5' column so the
  two stay in sync."""
  book = openpyxl.load_workbook(FILE_PATH)
  coll_sheet = book["Collection"]
  target_row = find_row_by_id(coll_sheet, id_column=1, target_id=film_id)
  if not target_row:
    return False

  coll_sheet.cell(row=target_row, column=7, value=watched_value)
  coll_sheet.cell(row=target_row, column=8, value=rating_value)

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
        p_format = picked_movie.get("Format", "Blu-ray")
        p_genre = picked_movie.get("Genre", "N/A")
        p_director = picked_movie.get("Director", "N/A")
        p_id = picked_movie.get("Film ID", "")
        p_watched_raw = picked_movie.get("Watched", "No")
        p_watched = watched_display(p_watched_raw)
        p_rating_display = rating_stars_display(picked_movie.get("Rating (1-5)"))

        st.markdown(
            f"""
                <div class="winner-box">
                    <h3 style="color: #F0A83B; margin-bottom: 2px;">🎉 Tonight's Pick:</h3>
                    <h2 style="margin: 0px;">{p_title} ({p_year_str})</h2>
                    <p style="font-size: 0.95em; margin-top: 8px; opacity: 0.85;">
                        <b>Format:</b> {p_format} | <b>Director:</b> {p_director}<br>
                        <b>Genre:</b> {p_genre} | <b>Watched:</b> {p_watched}<br>
                        <b>Rating:</b> {p_rating_display}
                    </p>
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

          if st.form_submit_button("💾 Save to Excel"):
            if save_collection_update(p_id, quick_watched, quick_stars):
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
      st.success(f"Found {len(results)} matches:")
      st.dataframe(results, use_container_width=True)
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

        if st.form_submit_button("💾 Save Changes"):
          if save_collection_update(selected_id, watched_status, new_rating_stars):
            st.cache_data.clear()
            st.success(f"Saved '{selected_movie}'!")
            if new_rating_stars == STAR_OPTIONS[-1]:
              st.balloons()
            st.rerun()
          else:
            st.error("Couldn't find that film in the sheet to update.")

    film_divider()
    st.dataframe(collection_df, use_container_width=True, height=350)

  # --- TAB 4: WISHLIST ---
  with app_mode[3]:
    st.subheader("🛒 Wishlist")

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
          book = openpyxl.load_workbook(FILE_PATH)
          sheet = book["Wishlist"]
          next_row = sheet.max_row + 1
          sheet.cell(row=next_row, column=1, value=new_title)
          sheet.cell(row=next_row, column=3, value=new_priority)
          sheet.cell(row=next_row, column=4, value=new_target_price)
          sheet.cell(row=next_row, column=8, value="No")
          sheet.cell(row=next_row, column=9, value=new_notes)
          book.save(FILE_PATH)
          save_and_sync(FILE_PATH)
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
    sub_tab1, sub_tab2, sub_tab3 = st.tabs(["Franchises", "Awards", "Ratings"])
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
      st.dataframe(awards_df, use_container_width=True, height=300)
    with sub_tab3:
      st.caption("Your score plus external ratings (IMDb, RT, Letterboxd, Metacritic, TMDb).")
      st.dataframe(ratings_df, use_container_width=True, height=300)

except FileNotFoundError:
  st.error(
      f"Couldn't find '{FILE_PATH}'. Make sure the spreadsheet is in the "
      "same folder as this app."
  )
except Exception as e:
  st.error(f"Error: {e}")
