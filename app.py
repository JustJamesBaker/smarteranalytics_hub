from pathlib import Path
import base64
from io import StringIO

import numpy as np
import pandas as pd
import requests
import streamlit as st
import yfinance as yf
from pandas.tseries.offsets import MonthEnd


# =====================================
# CONFIG
# =====================================
BASE_DIR = Path(__file__).parent
DATA_FILE = BASE_DIR / "index_database.xlsx"
ALBION_LOGO_FILE = BASE_DIR / "albion_logo.png"
POWERED_BY_FILE = BASE_DIR / "Powered by SA.png"

TIME_SERIES_SHEET = "time_series"
MAPPING_SHEET = "mapping"

BRAND_ORANGE = "#f36f21"
BRAND_ORANGE_DARK = "#d65f17"
LIGHT_GREY = "#f3f3f3"
TEXT_GREY = "#555555"

DISPLAY_GROUPS_ABSOLUTE = [
    {
        "title": "Global Equity",
        "items": ["Global stocks"],
        "labels": {"Global stocks": ""},
    },
    {
        "title": "UK Equity",
        "items": ["UK stocks", "UK value stocks", "UK small stocks"],
        "labels": {
            "UK stocks": "Market",
            "UK value stocks": "Value",
            "UK small stocks": "Small",
        },
    },
    {
        "title": "Developed Equity",
        "items": ["Developed stocks", "Developed value stocks", "Developed small stocks"],
        "labels": {
            "Developed stocks": "Market",
            "Developed value stocks": "Value",
            "Developed small stocks": "Small",
        },
    },
    {
        "title": "Emerging Equity",
        "items": ["Emerging stocks", "Emerging value stocks", "Emerging small stocks"],
        "labels": {
            "Emerging stocks": "Market",
            "Emerging value stocks": "Value",
            "Emerging small stocks": "Small",
        },
    },
    {
        "title": "REITs",
        "items": ["Developed REITs"],
        "labels": {"Developed REITs": ""},
    },
    {
        "title": "",
        "items": [
            "Cash (GBP)",
            "UK Gilts (0-5)",
            "UK IL Gilts (0-5)",
            "Global GBP hedged bonds (0-5)",
        ],
        "labels": {
            "Cash (GBP)": "Cash",
            "UK Gilts (0-5)": "UK Gilts (0-5)",
            "UK IL Gilts (0-5)": "UK IL Gilts (0-5)",
            "Global GBP hedged bonds (0-5)": "Global GBP hedged bonds (0-5)",
        },
    },
]

DISPLAY_GROUPS_RELATIVE_MINOR = [
    {
        "title": "Relative to DM",
        "items": [
            "UK stocks",
            "Emerging stocks",
            "Developed REITs",
            "Developed value stocks",
            "Developed small stocks",
        ],
        "labels": {
            "UK stocks": "UK",
            "Emerging stocks": "EM",
            "Developed REITs": "REIT",
            "Developed value stocks": "DM Value",
            "Developed small stocks": "DM Small",
        },
    },
    {
        "title": "Relative to EM",
        "items": ["Emerging value stocks", "Emerging small stocks"],
        "labels": {
            "Emerging value stocks": "EM Value",
            "Emerging small stocks": "EM Small",
        },
    },
    {
        "title": "Relative to UK",
        "items": ["UK value stocks", "UK small stocks"],
        "labels": {
            "UK value stocks": "UK Value",
            "UK small stocks": "UK Small",
        },
    },
    {
        "title": "",
        "items": [
            "Cash (GBP)",
            "UK Gilts (0-5)",
            "UK IL Gilts (0-5)",
            "Global GBP hedged bonds (0-5)",
        ],
        "labels": {
            "Cash (GBP)": "Cash",
            "UK Gilts (0-5)": "UK Gilts (0-5)",
            "UK IL Gilts (0-5)": "UK IL Gilts (0-5)",
            "Global GBP hedged bonds (0-5)": "Global GBP hedged bonds (0-5)",
        },
    },
]

HORIZONS = {
    "20Y": "20 Year",
    "10Y": "10 Year",
    "5Y": "5 Year",
    "YTD": "YTD",
}

ASSET_CLASS_ALIASES = {
    "Global equity": "Global stocks",
    "Global stocks": "Global stocks",
    "World equity": "Global stocks",
    "World stocks": "Global stocks",
    "UK equity": "UK stocks",
    "UK stocks": "UK stocks",
    "UK value": "UK value stocks",
    "UK value stocks": "UK value stocks",
    "UK small": "UK small stocks",
    "UK small stocks": "UK small stocks",
    "Developed equity": "Developed stocks",
    "Developed stocks": "Developed stocks",
    "Developed value": "Developed value stocks",
    "Developed value stocks": "Developed value stocks",
    "Developed small": "Developed small stocks",
    "Developed small stocks": "Developed small stocks",
    "Emerging market equity": "Emerging stocks",
    "Emerging equity": "Emerging stocks",
    "Emerging stocks": "Emerging stocks",
    "Emerging value": "Emerging value stocks",
    "Emerging value stocks": "Emerging value stocks",
    "Emerging small": "Emerging small stocks",
    "Emerging small stocks": "Emerging small stocks",
    "Developed REITs": "Developed REITs",
    "REITs": "Developed REITs",
    "Cash GBP": "Cash (GBP)",
    "Cash (GBP)": "Cash (GBP)",
    "Short Gilt": "UK Gilts (0-5)",
    "UK Gilts (0-5)": "UK Gilts (0-5)",
    "Short IL Gilt": "UK IL Gilts (0-5)",
    "UK IL Gilts (0-5)": "UK IL Gilts (0-5)",
    "Global GBP Hedged bonds (0-5)": "Global GBP hedged bonds (0-5)",
    "Global GBP hedged bonds (0-5)": "Global GBP hedged bonds (0-5)",
    "Global Short Bond GBP": "Global GBP hedged bonds (0-5)",
}

ETF_DOWNLOAD_START = "2000-01-01"

MAJOR_GROWTH_BASE_MAP = {
    "Global stocks": None,
    "UK stocks": "Global stocks",
    "UK value stocks": "Global stocks",
    "UK small stocks": "Global stocks",
    "Developed stocks": "Global stocks",
    "Developed value stocks": "Global stocks",
    "Developed small stocks": "Global stocks",
    "Emerging stocks": "Global stocks",
    "Emerging value stocks": "Global stocks",
    "Emerging small stocks": "Global stocks",
    "Developed REITs": "Global stocks",
}

MINOR_GROWTH_BASE_MAP = {
    "Global stocks": None,
    "UK stocks": "Developed stocks",
    "UK value stocks": "UK stocks",
    "UK small stocks": "UK stocks",
    "Developed stocks": None,
    "Developed value stocks": "Developed stocks",
    "Developed small stocks": "Developed stocks",
    "Emerging stocks": "Developed stocks",
    "Emerging value stocks": "Emerging stocks",
    "Emerging small stocks": "Emerging stocks",
    "Developed REITs": "Developed stocks",
}

DEFENSIVE_BASE_MAP = {
    "Cash (GBP)": None,
    "UK Gilts (0-5)": "Cash (GBP)",
    "UK IL Gilts (0-5)": "Cash (GBP)",
    "Global GBP hedged bonds (0-5)": "Cash (GBP)",
}

INFLATION_SERIES_ALIASES = [
    "UK inflation",
    "UK CPIH",
    "CPIH",
    "UK CPI",
    "CPI",
    "Inflation",
    "UK RPI",
    "RPI",
]

ONS_CPI_INDEX_CSV_URL = (
    "https://www.ons.gov.uk/generator?format=csv&uri="
    "/economy/inflationandpriceindices/timeseries/d7bt/mm23"
)


# =====================================
# HELPERS
# =====================================
def img_to_base64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("utf-8")


def standardise_series(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce")
    max_abs = s.dropna().abs().max()
    if pd.notna(max_abs) and max_abs > 1.5:
        s = s / 100.0
    return s


def normalise_ticker(value: object) -> str:
    if pd.isna(value):
        return ""
    ticker = str(value).strip().upper()
    if ticker.lower() in {"", "nan", "none"}:
        return ""
    return ticker


def normalise_name(value: object) -> str:
    return str(value).strip().lower()


def format_pct(x: float) -> str:
    return "-" if pd.isna(x) else f"{x:.1%}"


def annualised_return_from_growth(growth: float, years: float) -> float:
    if pd.isna(growth) or growth <= 0 or years <= 0:
        return np.nan
    return growth ** (1 / years) - 1


def build_lookup_table(returns_df: pd.DataFrame) -> dict:
    if returns_df.empty:
        return {}
    return returns_df.set_index("asset_class").to_dict(orient="index")


def heat_colour(value: float, vmin: float, vmax: float) -> str:
    if pd.isna(value):
        return "#E9E9E9"
    if value < 0:
        return "#F3B5B5"

    if vmin == vmax:
        norm = 0.75 if value >= 0 else 0.25
    else:
        norm = (value - vmin) / (vmax - vmin)
        norm = min(max(norm, 0), 1)

    light = np.array([120, 255, 120])
    dark = np.array([0, 150, 0])
    rgb = (light * (1 - norm) + dark * norm).astype(int)
    return f"rgb({rgb[0]},{rgb[1]},{rgb[2]})"


def group_class_name(title: str) -> str:
    return "group-card group-card-standard"


def safe_relative_return(asset_return: float, base_return: float) -> float:
    if pd.isna(asset_return) or pd.isna(base_return):
        return np.nan
    if (1 + base_return) <= 0:
        return np.nan
    return ((1 + asset_return) / (1 + base_return)) - 1


def get_display_groups(is_relative_mode: bool, relative_detail_mode: str) -> list[dict]:
    if is_relative_mode and relative_detail_mode == "Minor":
        return DISPLAY_GROUPS_RELATIVE_MINOR
    return DISPLAY_GROUPS_ABSOLUTE


def convert_to_relative_returns(
    absolute_returns_df: pd.DataFrame,
    relative_detail_mode: str,
) -> pd.DataFrame:
    relative_df = absolute_returns_df.copy()
    growth_base_map = MAJOR_GROWTH_BASE_MAP if relative_detail_mode == "Major" else MINOR_GROWTH_BASE_MAP

    all_base_maps = {}
    all_base_maps.update(growth_base_map)
    all_base_maps.update(DEFENSIVE_BASE_MAP)

    asset_to_row = relative_df.set_index("asset_class")

    for horizon in HORIZONS.keys():
        for idx, row in relative_df.iterrows():
            asset = row["asset_class"]
            asset_val = row[horizon]
            base_asset = all_base_maps.get(asset)

            if base_asset is None or base_asset not in asset_to_row.index:
                relative_df.at[idx, horizon] = np.nan
                continue

            base_val = asset_to_row.at[base_asset, horizon]
            relative_df.at[idx, horizon] = safe_relative_return(asset_val, base_val)

    return relative_df


def find_inflation_column(ts: pd.DataFrame) -> str:
    cols = [str(c).strip() for c in ts.columns]
    lookup = {normalise_name(c): c for c in cols}

    for alias in INFLATION_SERIES_ALIASES:
        key = normalise_name(alias)
        if key in lookup:
            return lookup[key]

    raise KeyError(
        "Could not find a UK inflation series in time_series sheet. "
        f"Tried aliases: {', '.join(INFLATION_SERIES_ALIASES)}"
    )


def build_inflation_levels_from_timeseries(ts: pd.DataFrame) -> pd.Series:
    inflation_col = find_inflation_column(ts)
    ret = standardise_series(ts[inflation_col])
    ser = pd.Series(ret.values, index=ts["Date"], name="UK inflation").dropna().sort_index()

    if ser.empty:
        raise ValueError("Inflation series was found but contains no valid data.")

    levels = (1 + ser).cumprod()
    levels.name = "UK inflation"
    return levels


@st.cache_data(show_spinner=False)
def fetch_ons_raw_csv(csv_url: str) -> pd.DataFrame:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/146.0.0.0 Safari/537.36"
        ),
        "Accept": "text/csv,application/csv,text/plain,*/*",
        "Referer": "https://www.ons.gov.uk/",
    }

    resp = requests.get(csv_url, headers=headers, timeout=30)
    resp.raise_for_status()
    return pd.read_csv(StringIO(resp.text))


@st.cache_data(show_spinner=False)
def fetch_ons_cpi_index_series(csv_url: str) -> pd.Series:
    raw = fetch_ons_raw_csv(csv_url)
    raw.columns = [str(c).strip() for c in raw.columns]

    if raw.shape[1] < 2:
        raise ValueError("ONS CPI index CSV did not contain at least two columns.")

    date_col = raw.columns[0]
    value_col = raw.columns[1]

    s = raw[[date_col, value_col]].copy()
    s.columns = ["date_raw", "value_raw"]

    s["date_raw"] = s["date_raw"].astype(str).str.strip()
    s["value"] = pd.to_numeric(s["value_raw"], errors="coerce")

    s = s.dropna(subset=["value"]).copy()
    s["Date"] = pd.NaT

    mask_yyyy_mon = s["date_raw"].str.match(r"^\d{4}\s+[A-Z]{3}$", na=False)
    if mask_yyyy_mon.any():
        s.loc[mask_yyyy_mon, "Date"] = pd.to_datetime(
            "01 " + s.loc[mask_yyyy_mon, "date_raw"],
            format="%d %Y %b",
            errors="coerce",
        )

    mask_mon_yyyy = s["date_raw"].str.match(r"^[A-Z]{3}\s+\d{4}$", na=False)
    if mask_mon_yyyy.any():
        s.loc[mask_mon_yyyy, "Date"] = pd.to_datetime(
            "01 " + s.loc[mask_mon_yyyy, "date_raw"],
            format="%d %b %Y",
            errors="coerce",
        )

    mask_year = s["date_raw"].str.match(r"^\d{4}$", na=False)
    if mask_year.any():
        s.loc[mask_year, "Date"] = pd.to_datetime(
            s.loc[mask_year, "date_raw"] + "-01-01",
            format="%Y-%m-%d",
            errors="coerce",
        )

    s = s.dropna(subset=["Date"]).copy()
    s["Date"] = pd.to_datetime(s["Date"]) + MonthEnd(0)
    s = s.sort_values("Date")

    out = pd.Series(s["value"].values, index=s["Date"], name="ONS CPI index")
    out = out[~out.index.duplicated(keep="last")].sort_index()

    if out.empty:
        raise ValueError("ONS CPI index CSV produced no valid rows.")

    return out


def extend_inflation_levels_with_ons_index(
    existing_inflation_levels: pd.Series,
    ons_index_series: pd.Series,
) -> pd.Series:
    if existing_inflation_levels is None or existing_inflation_levels.dropna().empty:
        raise ValueError("Existing inflation level history is required.")

    wb = existing_inflation_levels.dropna().copy().sort_index()
    ons = ons_index_series.dropna().copy().sort_index()

    common_dates = wb.index.intersection(ons.index)
    if common_dates.empty:
        raise ValueError("No overlapping dates between workbook inflation and ONS CPI index.")

    anchor_date = common_dates.max()
    wb_anchor = float(wb.loc[anchor_date])
    ons_anchor = float(ons.loc[anchor_date])

    if pd.isna(ons_anchor) or ons_anchor == 0:
        raise ValueError("Invalid ONS anchor value.")

    scale_factor = wb_anchor / ons_anchor
    scaled_ons = ons * scale_factor

    combined = pd.concat(
        [
            wb[wb.index <= anchor_date],
            scaled_ons[scaled_ons.index > anchor_date],
        ]
    )
    combined = combined[~combined.index.duplicated(keep="last")].sort_index()
    combined.name = "UK inflation"
    return combined


def build_best_available_inflation_levels(
    ts: pd.DataFrame,
) -> tuple[pd.Series | None, str, str | None]:
    workbook_levels = build_inflation_levels_from_timeseries(ts)
    source_note = "Workbook time_series"

    try:
        ons_index = fetch_ons_cpi_index_series(ONS_CPI_INDEX_CSV_URL)
        extended_levels = extend_inflation_levels_with_ons_index(workbook_levels, ons_index)

        if extended_levels.index.max() > workbook_levels.index.max():
            return (
                extended_levels,
                "Workbook time_series + ONS CPI index extension",
                None,
            )

        return workbook_levels, source_note, "ONS fetched, but no extension beyond workbook end date."
    except Exception as exc:
        return workbook_levels, source_note, f"ONS extension failed: {type(exc).__name__}: {exc}"


def get_live_consistent_end_date(
    stitched_series_map: dict[str, pd.Series],
    live_diag: pd.DataFrame,
) -> pd.Timestamp:
    if live_diag is None or live_diag.empty:
        last_dates = []
        for ser in stitched_series_map.values():
            s = ser.dropna()
            if not s.empty:
                last_dates.append(pd.Timestamp(s.index.max()))
        if not last_dates:
            return pd.Timestamp.today().normalize()
        return max(last_dates)

    diag = live_diag.copy()
    live_rows = diag[diag["series_type"].isin(["stitched", "live_only"])].copy()

    if not live_rows.empty:
        live_last_dates = pd.to_datetime(live_rows["live_last_date"], errors="coerce").dropna()
        if not live_last_dates.empty:
            return pd.Timestamp(live_last_dates.min())

    last_dates = []
    for ser in stitched_series_map.values():
        s = ser.dropna()
        if not s.empty:
            last_dates.append(pd.Timestamp(s.index.max()))
    if not last_dates:
        return pd.Timestamp.today().normalize()
    return max(last_dates)


def get_dashboard_end_date(
    stitched_series_map: dict[str, pd.Series],
    live_diag: pd.DataFrame,
    inflation_levels: pd.Series | None,
    is_real_mode: bool,
) -> pd.Timestamp:
    end_date = get_live_consistent_end_date(stitched_series_map, live_diag)

    if is_real_mode and inflation_levels is not None and not inflation_levels.dropna().empty:
        inflation_last_date = pd.Timestamp(inflation_levels.dropna().index.max())
        end_date = min(end_date, inflation_last_date)

    return end_date


def convert_to_real_returns(
    nominal_returns_df: pd.DataFrame,
    inflation_returns_df: pd.DataFrame,
) -> pd.DataFrame:
    real_df = nominal_returns_df.copy()

    if inflation_returns_df.empty:
        return real_df

    infl_row = inflation_returns_df[inflation_returns_df["asset_class"] == "UK inflation"]
    if infl_row.empty:
        return real_df

    infl_row = infl_row.iloc[0]

    for horizon in HORIZONS.keys():
        infl_val = infl_row[horizon]
        for idx, row in real_df.iterrows():
            real_df.at[idx, horizon] = safe_relative_return(row[horizon], infl_val)

    return real_df


def get_methodology_paragraph(
    is_relative_mode: bool,
    relative_detail_mode: str,
    is_real_mode: bool,
    inflation_source_note: str,
) -> str:
    basis = "real" if is_real_mode else "nominal"
    prefix = f"This dashboard shows annualised {basis} GBP returns across the displayed horizons, with YTD shown cumulatively."

    if is_relative_mode:
        if relative_detail_mode == "Major":
            relative_text = (
                " In relative mode, growth assets are shown relative to Global stocks and defensive assets are shown relative to Cash (GBP)."
            )
        else:
            relative_text = (
                " In relative mode (Minor), UK, EM, REITs, Developed Value and Developed Small are shown relative to Developed stocks; "
                "EM Value and EM Small are shown relative to Emerging stocks; UK Value and UK Small are shown relative to UK stocks; "
                "and defensive assets are shown relative to Cash (GBP)."
            )
    else:
        relative_text = " In absolute mode, headline returns are shown directly for each asset class."

    inflation_text = ""
    if is_real_mode:
        inflation_text = (
            f" Real returns are calculated using UK inflation via (1+asset return)/(1+inflation return)-1. "
            f"Current inflation source: {inflation_source_note}."
        )

    source_text = (
        " Albion index series history is used as the preferred source where available. "
        "yfinance mappings are used to keep the live series as up to date as possible once index history ends, "
        "and those live extensions are periodically overwritten by Albion indices on a quarterly basis. "
        "More information on the Albion indices can be found at "
        '<a href="https://smartersuccess.net/indices" target="_blank">smartersuccess.net/indices</a>.'
    )

    return prefix + relative_text + inflation_text + source_text


def dataframe_to_csv_download(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def build_monthly_returns_from_levels(levels: pd.Series) -> pd.Series:
    s = levels.dropna().sort_index()
    if s.empty:
        return pd.Series(dtype=float, name="UK inflation")
    ret = s.pct_change()
    ret.name = "UK inflation"
    return ret


def format_diagnostic_table(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in ["index_last_date", "live_first_date", "live_last_date", "stitch_anchor_date"]:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce").dt.strftime("%d/%m/%Y")
            out[c] = out[c].fillna("")
    return out


def convert_pct_table_for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in ["YTD", "5Y", "10Y", "20Y"]:
        if c in out.columns:
            out[c] = out[c].map(lambda x: np.nan if pd.isna(x) else round(x * 100, 2))
    return out


# =====================================
# DATA LOAD
# =====================================
@st.cache_data(show_spinner=False)
def load_data(file_path: str, file_mtime: float):
    ts = pd.read_excel(file_path, sheet_name=TIME_SERIES_SHEET)
    mp = pd.read_excel(file_path, sheet_name=MAPPING_SHEET)

    ts = ts.copy()
    mp = mp.copy()

    ts.columns = [str(c).strip() for c in ts.columns]
    mp.columns = [str(c).strip() for c in mp.columns]

    ts.iloc[:, 0] = pd.to_datetime(ts.iloc[:, 0], errors="coerce")
    ts = ts.rename(columns={ts.columns[0]: "Date"})
    ts = ts.dropna(subset=["Date"]).sort_values("Date").reset_index(drop=True)

    rename_map = {}
    if len(mp.columns) >= 1:
        rename_map[mp.columns[0]] = "index_name"
    if len(mp.columns) >= 2:
        rename_map[mp.columns[1]] = "asset_class"
    if len(mp.columns) >= 3:
        rename_map[mp.columns[2]] = "live_fund_primary"
    if len(mp.columns) >= 4:
        rename_map[mp.columns[3]] = "live_fund_secondary"

    mp = mp.rename(columns=rename_map)

    if "index_name" in mp.columns:
        mp["index_name"] = mp["index_name"].astype(str).str.strip()

    if "asset_class" in mp.columns:
        mp["asset_class"] = mp["asset_class"].astype(str).str.strip()
        mp["asset_class"] = mp["asset_class"].replace(ASSET_CLASS_ALIASES)

    if "live_fund_primary" in mp.columns:
        mp["live_fund_primary"] = mp["live_fund_primary"].map(normalise_ticker)

    if "live_fund_secondary" in mp.columns:
        mp["live_fund_secondary"] = mp["live_fund_secondary"].map(normalise_ticker)

    valid_rows = mp[mp["index_name"].isin(ts.columns)].copy()
    return ts, valid_rows


def build_monthly_index_levels(ts: pd.DataFrame, mapping: pd.DataFrame) -> dict[str, pd.Series]:
    needed_assets = set(mapping["asset_class"].dropna().astype(str).unique())
    output = {}

    for _, row in mapping.iterrows():
        asset_class = row["asset_class"]
        index_name = row["index_name"]

        if asset_class not in needed_assets:
            continue
        if index_name not in ts.columns:
            continue

        ret = standardise_series(ts[index_name])
        ser = pd.Series(ret.values, index=ts["Date"], name=asset_class).dropna().sort_index()

        if ser.empty:
            continue

        levels = (1 + ser).cumprod()
        levels.name = asset_class
        output[asset_class] = levels

    return output


# =====================================
# YFINANCE
# =====================================
@st.cache_data(show_spinner=False)
def fetch_yf_prices(tickers: tuple[str, ...], start_date: str) -> pd.DataFrame:
    tickers = tuple(sorted({normalise_ticker(t) for t in tickers if normalise_ticker(t)}))
    if not tickers:
        return pd.DataFrame()

    data = yf.download(
        list(tickers),
        start=start_date,
        progress=False,
        auto_adjust=True,
        actions=False,
        threads=False,
        group_by="ticker",
    )

    if data is None or len(data) == 0:
        return pd.DataFrame()

    close_frames = []

    if isinstance(data.columns, pd.MultiIndex):
        for ticker in tickers:
            if ticker in data.columns.get_level_values(0):
                sub = data[ticker].copy()
                close_col = "Close" if "Close" in sub.columns else None
                if close_col is not None:
                    ser = pd.to_numeric(sub[close_col], errors="coerce").rename(ticker)
                    close_frames.append(ser)
    else:
        if "Close" in data.columns and len(tickers) == 1:
            close_frames.append(pd.to_numeric(data["Close"], errors="coerce").rename(tickers[0]))

    if not close_frames:
        return pd.DataFrame()

    out = pd.concat(close_frames, axis=1)
    out.index = pd.to_datetime(out.index).tz_localize(None)
    out = out.sort_index()
    return out


def get_price_series(prices_df: pd.DataFrame, ticker: str) -> pd.Series:
    if prices_df.empty or ticker not in prices_df.columns:
        return pd.Series(dtype=float)
    return pd.to_numeric(prices_df[ticker], errors="coerce").dropna().sort_index()


def pick_live_ticker_for_asset(
    row: pd.Series,
    prices_df: pd.DataFrame,
) -> tuple[str, str, str, pd.Series]:
    primary = normalise_ticker(row.get("live_fund_primary", ""))
    secondary = normalise_ticker(row.get("live_fund_secondary", ""))

    primary_ser = get_price_series(prices_df, primary) if primary else pd.Series(dtype=float)
    secondary_ser = get_price_series(prices_df, secondary) if secondary else pd.Series(dtype=float)

    if not primary_ser.empty:
        return primary, "primary", "Primary available in yfinance", primary_ser

    if not secondary_ser.empty:
        return secondary, "secondary", "Primary unavailable; using secondary", secondary_ser

    if primary and secondary:
        return "", "none", "Neither primary nor secondary returned price history", pd.Series(dtype=float)
    if primary:
        return "", "none", "Primary returned no price history and no secondary provided", pd.Series(dtype=float)
    if secondary:
        return "", "none", "No primary provided; secondary returned no price history", pd.Series(dtype=float)
    return "", "none", "No live fund tickers provided", pd.Series(dtype=float)


def nearest_on_or_before(series: pd.Series, target_date: pd.Timestamp) -> tuple[pd.Timestamp | None, float | None]:
    s = series.dropna()
    s = s[s.index <= target_date]
    if s.empty:
        return None, None
    return s.index[-1], float(s.iloc[-1])


def build_stitched_asset_series(
    monthly_levels: dict[str, pd.Series],
    mapping: pd.DataFrame,
    prices_df: pd.DataFrame,
) -> tuple[dict[str, pd.Series], pd.DataFrame]:
    needed_assets = set(mapping["asset_class"].dropna().astype(str).unique())
    stitched = {}
    diag_rows = []

    deduped_mapping = mapping.drop_duplicates(subset=["asset_class"]).copy()

    for asset_class in sorted(needed_assets):
        row_match = deduped_mapping[deduped_mapping["asset_class"] == asset_class]

        if asset_class not in monthly_levels:
            diag_rows.append(
                {
                    "asset_class": asset_class,
                    "live_fund_primary": "",
                    "live_fund_secondary": "",
                    "selected_ticker": "",
                    "selected_source": "none",
                    "series_type": "missing",
                    "index_last_date": pd.NaT,
                    "live_first_date": pd.NaT,
                    "live_last_date": pd.NaT,
                    "stitch_anchor_date": pd.NaT,
                    "note": "No index series found for asset class",
                }
            )
            continue

        index_levels = monthly_levels[asset_class].dropna().sort_index()
        row = row_match.iloc[0] if not row_match.empty else pd.Series(dtype=object)

        primary = normalise_ticker(row.get("live_fund_primary", ""))
        secondary = normalise_ticker(row.get("live_fund_secondary", ""))

        selected_ticker, selected_source, note, live_prices = pick_live_ticker_for_asset(
            row=row,
            prices_df=prices_df,
        )

        index_last_date = pd.Timestamp(index_levels.index.max())

        if live_prices.empty:
            stitched[asset_class] = index_levels.copy()
            diag_rows.append(
                {
                    "asset_class": asset_class,
                    "live_fund_primary": primary,
                    "live_fund_secondary": secondary,
                    "selected_ticker": selected_ticker,
                    "selected_source": selected_source,
                    "series_type": "index_only",
                    "index_last_date": index_last_date,
                    "live_first_date": pd.NaT,
                    "live_last_date": pd.NaT,
                    "stitch_anchor_date": pd.NaT,
                    "note": note,
                }
            )
            continue

        live_prices = live_prices.dropna().sort_index()
        live_first_date = pd.Timestamp(live_prices.index.min())
        live_last_date = pd.Timestamp(live_prices.index.max())

        anchor_date = index_last_date
        anchor_level = float(index_levels.loc[anchor_date])

        live_anchor_date, live_anchor_price = nearest_on_or_before(live_prices, anchor_date)

        if live_anchor_date is None or live_anchor_price is None:
            stitched[asset_class] = index_levels.copy()
            diag_rows.append(
                {
                    "asset_class": asset_class,
                    "live_fund_primary": primary,
                    "live_fund_secondary": secondary,
                    "selected_ticker": selected_ticker,
                    "selected_source": selected_source,
                    "series_type": "index_only",
                    "index_last_date": index_last_date,
                    "live_first_date": live_first_date,
                    "live_last_date": live_last_date,
                    "stitch_anchor_date": anchor_date,
                    "note": f"{note}. Live series has no price on or before the final index date, so no stitch was applied",
                }
            )
            continue

        live_extension = live_prices[live_prices.index > anchor_date].copy()

        if live_extension.empty:
            stitched[asset_class] = index_levels.copy()
            diag_rows.append(
                {
                    "asset_class": asset_class,
                    "live_fund_primary": primary,
                    "live_fund_secondary": secondary,
                    "selected_ticker": selected_ticker,
                    "selected_source": selected_source,
                    "series_type": "index_only",
                    "index_last_date": index_last_date,
                    "live_first_date": live_first_date,
                    "live_last_date": live_last_date,
                    "stitch_anchor_date": anchor_date,
                    "note": f"{note}. Live series exists but has no data after the final index date",
                }
            )
            continue

        live_levels = anchor_level * (live_extension / live_anchor_price)
        stitched_series = pd.concat([index_levels, live_levels])
        stitched_series = stitched_series[~stitched_series.index.duplicated(keep="last")].sort_index()
        stitched[asset_class] = stitched_series

        diag_rows.append(
            {
                "asset_class": asset_class,
                "live_fund_primary": primary,
                "live_fund_secondary": secondary,
                "selected_ticker": selected_ticker,
                "selected_source": selected_source,
                "series_type": "stitched",
                "index_last_date": index_last_date,
                "live_first_date": live_first_date,
                "live_last_date": live_last_date,
                "stitch_anchor_date": anchor_date,
                "note": f"{note}. Index history retained through its final date, then extended with live adjusted-close history",
            }
        )

    diag = pd.DataFrame(diag_rows)
    return stitched, diag


def calc_horizon_returns_from_levels(
    stitched_series_map: dict[str, pd.Series],
    end_date: pd.Timestamp,
) -> pd.DataFrame:
    out = []

    for asset_class, series in stitched_series_map.items():
        s = series.dropna().sort_index()
        s = s[s.index <= end_date]

        row = {"asset_class": asset_class}

        if s.empty:
            row["YTD"] = np.nan
            row["5Y"] = np.nan
            row["10Y"] = np.nan
            row["20Y"] = np.nan
            out.append(row)
            continue

        end_level = float(s.iloc[-1])

        prev_year_end = pd.Timestamp(end_date.year - 1, 12, 31)
        ytd_base = s[s.index <= prev_year_end]
        row["YTD"] = (end_level / float(ytd_base.iloc[-1]) - 1) if not ytd_base.empty else np.nan

        for label, years in [("5Y", 5), ("10Y", 10), ("20Y", 20)]:
            lookback_date = end_date - pd.DateOffset(years=years)
            base = s[s.index <= lookback_date]
            if base.empty:
                row[label] = np.nan
            else:
                growth = end_level / float(base.iloc[-1])
                row[label] = annualised_return_from_growth(growth, years)

        out.append(row)

    return pd.DataFrame(out)


# =====================================
# PAGE SETUP
# =====================================
st.set_page_config(page_title="Market dashboard", layout="wide")

st.markdown(
    """
    <style>
    /* =========================
       GLOBAL
       ========================= */
    html, body, .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stMarkdownContainer"],
    p, li, label, button, input, textarea, table {
        font-family: "Calibri Light", Calibri, "Segoe UI", Arial, sans-serif !important;
    }

    .stApp {
        background-color: white;
    }

    header[data-testid="stHeader"],
    div[data-testid="stToolbar"],
    div[data-testid="stDecoration"],
    div[data-testid="stStatusWidget"] {
        display: none !important;
        visibility: hidden !important;
        height: 0 !important;
    }

    #MainMenu {
        visibility: hidden !important;
    }

    [data-testid="stAppViewContainer"] > .main {
        padding-top: 0 !important;
    }

    .block-container {
        padding-top: 0.6rem !important;
        padding-bottom: 1rem !important;
        max-width: 1500px;
    }

    /* =========================
       HEADER
       ========================= */
    .dashboard-header {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 30px;
        margin-bottom: 10px;
    }

    .dashboard-title-wrap {
        flex: 1;
        min-width: 0;
    }

    .dashboard-title {
        font-size: 40px;
        font-weight: 500;
        margin-bottom: 2px;
        color: black;
        line-height: 1.15;
        padding-top: 0;
    }

    .header-logo {
        flex-shrink: 0;
        display: flex;
        align-items: flex-start;
        justify-content: flex-end;
        min-width: 180px;
        overflow: visible;
        padding-top: 6px;
    }

    .header-logo img {
        max-height: 50px;
        width: auto;
        object-fit: contain;
        display: block;
    }

    /* =========================
       TOOLBAR
       ========================= */
    .toolbar-title {
        font-size: 13px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #444;
        margin-bottom: 10px;
    }

    .toolbar-label {
        font-size: 13px;
        font-weight: 700;
        color: #444;
        margin-bottom: 6px;
    }

    .toolbar-label-muted {
        color: #888 !important;
    }

    .toolbar-meta {
        text-align: right;
        font-size: 13px;
        color: #555;
        padding-top: 28px;
        line-height: 1.2;
        white-space: nowrap;
    }

    /* =========================
       PERIOD / TILE LAYOUT
       ========================= */
    .period-shell {
        background: transparent;
        padding: 8px 8px 12px 8px;
        min-height: 100%;
    }

    .period-title {
        text-align: center;
        font-size: 28px;
        font-weight: 500;
        margin-bottom: 10px;
        color: black;
    }

    .group-card {
        padding: 10px 8px 8px 8px;
        margin-bottom: 10px;
        background: #e7e7e7;
    }

    .section-title {
        text-align: center;
        font-size: 18px;
        font-weight: 500;
        margin-top: 0;
        margin-bottom: 4px;
        color: black;
    }

    .section-title-empty {
        height: 4px;
        margin-bottom: 2px;
    }

    .section-subtitle {
        text-align: center;
        font-size: 13px;
        font-style: italic;
        color: #444;
        margin-bottom: 6px;
    }

    .big-tile, .small-tile {
        color: white;
        text-align: center;
        font-weight: 700;
        border-radius: 0;
        line-height: 1.2;
    }

    .big-tile {
        padding: 10px 8px;
        margin-bottom: 8px;
        font-size: 22px;
    }

    .small-tile {
        padding: 8px 6px;
        font-size: 18px;
        margin-bottom: 8px;
    }

    .tile-label {
        display: block;
        color: black !important;
        font-size: 13px;
        font-style: italic;
        font-weight: 500;
        margin-bottom: 4px;
    }

    .tile-label-plain {
        font-style: normal !important;
        font-weight: 700 !important;
    }

    .tile-label-on-colour {
        display: block;
        color: black !important;
        font-size: 13px;
        font-style: italic;
        font-weight: 500;
        margin-bottom: 4px;
    }

    .spacer {
        height: 2px;
    }

    /* =========================
       FOOTER / NOTES
       ========================= */
    .methodology-text {
        margin-top: 18px;
        margin-bottom: 10px;
        font-size: 13px;
        color: #555 !important;
        text-align: center;
        line-height: 1.45;
    }

    .methodology-text a {
        color: #d65f17;
        text-decoration: none;
    }

    .footer-bar {
        margin-top: 20px;
        display: flex;
        justify-content: center;
        align-items: center;
        width: 100%;
    }

    .footer-logo img {
        max-height: 48px;
        width: auto;
        object-fit: contain;
        display: block;
    }

    .diag-title {
        font-size: 20px;
        font-weight: 500;
        color: black !important;
        margin-top: 4px;
        margin-bottom: 8px;
    }

    .diag-note {
        font-size: 13px;
        color: #555 !important;
        margin-bottom: 14px;
        line-height: 1.4;
    }

    /* =========================
       SEGMENTED CONTROLS
       ========================= */
    [data-testid="stSegmentedControl"] {
        width: 100%;
    }

    [data-testid="stSegmentedControl"] [role="radiogroup"] {
        gap: 0.25rem;
    }

    [data-testid="stSegmentedControl"] [role="radiogroup"] label {
        background: #f2f2f2 !important;
        border: 1px solid #cfcfcf !important;
        border-radius: 4px !important;
        min-height: 2.35rem !important;
        opacity: 1 !important;
    }

    [data-testid="stSegmentedControl"] [role="radiogroup"] label,
    [data-testid="stSegmentedControl"] [role="radiogroup"] label p,
    [data-testid="stSegmentedControl"] [role="radiogroup"] label span {
        color: black !important;
        fill: black !important;
    }

    [data-testid="stSegmentedControl"] [role="radiogroup"] label[data-selected="true"] {
        background: #f36f21 !important;
        border-color: #f36f21 !important;
    }

    [data-testid="stSegmentedControl"] [role="radiogroup"] label[data-selected="true"],
    [data-testid="stSegmentedControl"] [role="radiogroup"] label[data-selected="true"] p,
    [data-testid="stSegmentedControl"] [role="radiogroup"] label[data-selected="true"] span {
        color: white !important;
        fill: white !important;
    }

    /* =========================
       BUTTONS
       ========================= */
    .stDownloadButton button,
    .stButton button {
        background: #f2f2f2 !important;
        color: black !important;
        border: 1px solid #cfcfcf !important;
        border-radius: 4px !important;
        font-weight: 500 !important;
    }

    .stDownloadButton button *,
    .stButton button * {
        color: black !important;
    }

    .stDownloadButton button:hover,
    .stButton button:hover {
        border-color: #f36f21 !important;
        color: black !important;
    }

    /* =========================
       TABS
       ========================= */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.25rem !important;
        background: #ffffff !important;
    }

    .stTabs button[role="tab"] {
        background: #f2f2f2 !important;
        border: 1px solid #d0d0d0 !important;
        border-radius: 4px 4px 0 0 !important;
    }

    .stTabs button[role="tab"],
    .stTabs button[role="tab"] *,
    .stTabs button[role="tab"] p,
    .stTabs button[role="tab"] span,
    .stTabs button[role="tab"] div,
    .stTabs button[role="tab"] .stMarkdown,
    .stTabs button[role="tab"] [data-testid="stMarkdownContainer"],
    .stTabs button[role="tab"] [data-testid="stMarkdownContainer"] * {
        color: #000000 !important;
        fill: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
        opacity: 1 !important;
    }

    .stTabs button[role="tab"][aria-selected="true"] {
        background: #f36f21 !important;
        border-color: #f36f21 !important;
    }

    .stTabs button[role="tab"][aria-selected="true"],
    .stTabs button[role="tab"][aria-selected="true"] *,
    .stTabs button[role="tab"][aria-selected="true"] p,
    .stTabs button[role="tab"][aria-selected="true"] span,
    .stTabs button[role="tab"][aria-selected="true"] div,
    .stTabs button[role="tab"][aria-selected="true"] .stMarkdown,
    .stTabs button[role="tab"][aria-selected="true"] [data-testid="stMarkdownContainer"],
    .stTabs button[role="tab"][aria-selected="true"] [data-testid="stMarkdownContainer"] * {
        color: #ffffff !important;
        fill: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        opacity: 1 !important;
    }

    .stTabs button[role="tab"]:hover {
        border-color: #f36f21 !important;
    }

    /* =========================
       METRICS / HEADINGS SAFETY
       ========================= */
    .stMetric,
    .stMetric * {
        color: black !important;
    }

    h1, h2, h3, h4, h5, h6,
    [data-testid="stHeading"],
    [data-testid="stHeading"] * {
        color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
    }

    /* Diagnostics toggle button only */
    .st-key-diagnostics_toggle_button button {
        background: #111827 !important;
        color: #ffffff !important;
        border: 1px solid #111827 !important;
        border-radius: 4px !important;
        font-weight: 600 !important;
        text-align: left !important;
        justify-content: flex-start !important;
        padding: 0.55rem 0.9rem !important;
        margin: 0 !important;
    }

    .st-key-diagnostics_toggle_button button * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }

    .st-key-diagnostics_toggle_button {
        margin-top: 8px !important;
        margin-bottom: 0 !important;
    }

    .st-key-diagnostics_toggle_button + div {
        margin-top: 0 !important;
        padding-top: 0 !important;
    }

    /* =====================================
   KILL EMPTY STREAMLIT SPACER BLOCKS
   ===================================== */

    /* Remove empty blocks globally (safe) */
    div[data-testid="stVerticalBlock"] > div:empty {
        display: none !important;
    }

    /* Specifically tighten diagnostics area */
    .diag-wrapper div[data-testid="stVerticalBlock"] > div:empty {
        display: none !important;
    }

    /* Also remove gap between children */
    .diag-wrapper div[data-testid="stVerticalBlock"] {
        gap: 0 !important;
    }
        div[data-testid="stVerticalBlock"] > div {
        margin-top: 0 !important;
        margin-bottom: 0 !important;
    }
    /* Remove internal spacing from Streamlit button container */
    div[data-testid="stButton"] {
        margin: 0 !important;
        padding: 0 !important;
    }

    /* Kill invisible wrapper spacing */
    div[data-testid="stButton"] > div {
        margin: 0 !important;
        padding: 0 !important;
    }

    /* Force button to have zero bottom gap */
    div[data-testid="stButton"] button {
        margin-bottom: 0 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =====================================
# HEADER
# =====================================
logo_html = ""
if ALBION_LOGO_FILE.exists():
    logo_html = f'<div class="header-logo"><img src="data:image/png;base64,{img_to_base64(ALBION_LOGO_FILE)}"></div>'

st.markdown(
    f"""
    <div class="dashboard-header">
        <div class="dashboard-title-wrap">
            <div class="dashboard-title">Market metrics dashboard</div>
        </div>
        {logo_html}
    </div>
    """,
    unsafe_allow_html=True,
)


# =====================================
# LOAD + PREP
# =====================================
if not DATA_FILE.exists():
    st.error(f"File not found: {DATA_FILE}")
    st.stop()

current_file_mtime = DATA_FILE.stat().st_mtime
previous_file_mtime = st.session_state.get("index_db_mtime")

if previous_file_mtime != current_file_mtime:
    st.session_state["index_db_mtime"] = current_file_mtime
    st.session_state.pop("live_diagnostics", None)

try:
    ts, mapping = load_data(str(DATA_FILE), current_file_mtime)
except Exception as exc:
    st.exception(exc)
    st.stop()

monthly_levels = build_monthly_index_levels(ts, mapping)
if not monthly_levels:
    st.error(
        "No mapped asset classes were found. Check that the mapping sheet asset_class "
        "names match the expected names in DISPLAY_GROUPS / ASSET_CLASS_ALIASES."
    )
    st.stop()

inflation_levels = None
inflation_source_note = "Workbook time_series"
inflation_debug_message = None

try:
    inflation_levels, inflation_source_note, inflation_debug_message = build_best_available_inflation_levels(ts)
except Exception as exc:
    st.warning(f"Inflation series could not be built. Real mode may be unavailable. Details: {exc}")
    inflation_levels = None

inflation_monthly_returns = pd.Series(dtype=float)
if inflation_levels is not None and not inflation_levels.dropna().empty:
    inflation_monthly_returns = build_monthly_returns_from_levels(inflation_levels)

primary_tickers = mapping.get("live_fund_primary", pd.Series(dtype=str)).dropna().tolist()
secondary_tickers = mapping.get("live_fund_secondary", pd.Series(dtype=str)).dropna().tolist()
all_live_tickers = tuple(sorted({normalise_ticker(x) for x in primary_tickers + secondary_tickers if normalise_ticker(x)}))

try:
    live_prices = fetch_yf_prices(all_live_tickers, ETF_DOWNLOAD_START)
except Exception as exc:
    st.warning(f"Live price download failed. Falling back to index-only results. Details: {exc}")
    live_prices = pd.DataFrame()

stitched_series_map, live_diag = build_stitched_asset_series(
    monthly_levels=monthly_levels,
    mapping=mapping,
    prices_df=live_prices,
)

if not stitched_series_map:
    st.error("No stitched or index-only asset series could be built.")
    st.stop()

display_mode = st.session_state.get("display_mode_toolbar", "Absolute")
return_basis = st.session_state.get("return_basis_toolbar", "Nominal")
relative_detail_mode = st.session_state.get("relative_basis_toolbar", "Major")

is_relative_mode = display_mode == "Relative"
is_real_mode = return_basis == "Real"
effective_real_mode = is_real_mode and inflation_levels is not None and not inflation_levels.dropna().empty

end_date = get_dashboard_end_date(
    stitched_series_map=stitched_series_map,
    live_diag=live_diag,
    inflation_levels=inflation_levels,
    is_real_mode=effective_real_mode,
)


# =====================================
# TOOLBAR
# =====================================
st.markdown('<div class="toolbar-title">Toolbar</div>', unsafe_allow_html=True)

toolbar_cols = st.columns([1.15, 1.15, 1.15, 3.55])

with toolbar_cols[0]:
    st.markdown('<div class="toolbar-label">Display mode</div>', unsafe_allow_html=True)
    display_mode = st.segmented_control(
        label="Display mode",
        options=["Absolute", "Relative"],
        default=st.session_state.get("display_mode_toolbar", "Absolute"),
        key="display_mode_toolbar",
        label_visibility="collapsed",
    ) or "Absolute"

is_relative_mode = display_mode == "Relative"

with toolbar_cols[1]:
    label_class = "toolbar-label" if is_relative_mode else "toolbar-label toolbar-label-muted"
    st.markdown(f'<div class="{label_class}">Relative basis</div>', unsafe_allow_html=True)
    relative_detail_mode = st.segmented_control(
        label="Relative basis",
        options=["Major", "Minor"],
        default=st.session_state.get("relative_basis_toolbar", "Major"),
        key="relative_basis_toolbar",
        label_visibility="collapsed",
    ) or st.session_state.get("relative_basis_toolbar", "Major")

with toolbar_cols[2]:
    st.markdown('<div class="toolbar-label">Return basis</div>', unsafe_allow_html=True)
    return_basis = st.segmented_control(
        label="Return basis",
        options=["Nominal", "Real"],
        default=st.session_state.get("return_basis_toolbar", "Nominal"),
        key="return_basis_toolbar",
        label_visibility="collapsed",
    ) or "Nominal"

is_real_mode = return_basis == "Real"
effective_real_mode = is_real_mode and inflation_levels is not None and not inflation_levels.dropna().empty

with toolbar_cols[3]:
    st.markdown(
        f"""
        <div class="toolbar-meta">
            Annualised {"real" if effective_real_mode else "nominal"} returns in GBP to <b>{end_date.strftime("%d/%m/%Y")}</b>
        </div>
        """,
        unsafe_allow_html=True,
    )

display_groups = get_display_groups(is_relative_mode, relative_detail_mode)

if is_real_mode and not effective_real_mode:
    st.warning("Real mode selected but no usable UK inflation series was found. Falling back to nominal results.")


# =====================================
# RETURNS PREP
# =====================================
absolute_returns_df = calc_horizon_returns_from_levels(stitched_series_map, end_date)

mode_label = "relative" if is_relative_mode else "absolute"
basis_label = "real" if effective_real_mode else "nominal"
relative_detail_label = (
    str(relative_detail_mode).lower()
    if is_relative_mode and relative_detail_mode is not None
    else ""
)

if is_relative_mode:
    nominal_display_returns_df = convert_to_relative_returns(
        absolute_returns_df=absolute_returns_df,
        relative_detail_mode=relative_detail_mode,
    )
else:
    nominal_display_returns_df = absolute_returns_df.copy()

inflation_returns_df = pd.DataFrame()
if inflation_levels is not None and not inflation_levels.dropna().empty:
    inflation_returns_df = calc_horizon_returns_from_levels(
        stitched_series_map={"UK inflation": inflation_levels},
        end_date=end_date,
    )

if effective_real_mode:
    returns_df = convert_to_real_returns(
        nominal_returns_df=nominal_display_returns_df,
        inflation_returns_df=inflation_returns_df,
    )
else:
    returns_df = nominal_display_returns_df.copy()

lookup = build_lookup_table(returns_df)


# =====================================
# RENDER
# =====================================
cols = st.columns(4)
period_order = ["20Y", "10Y", "5Y", "YTD"]

for col, period in zip(cols, period_order):
    period_vals = returns_df[period].dropna()
    vmin = period_vals.min() if len(period_vals) else -0.05
    vmax = period_vals.max() if len(period_vals) else 0.15

    with col:
        st.markdown('<div class="period-shell">', unsafe_allow_html=True)
        st.markdown(f'<div class="period-title">{HORIZONS[period]}</div>', unsafe_allow_html=True)

        for group in display_groups:
            title = group["title"]
            items = group["items"]
            labels = group["labels"]

            st.markdown(f'<div class="{group_class_name(title)}">', unsafe_allow_html=True)

            if title:
                st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="section-title-empty"></div>', unsafe_allow_html=True)

            if len(items) == 1:
                item = items[0]
                val = lookup.get(item, {}).get(period, np.nan)
                colour = heat_colour(val, vmin, vmax)
                label = labels.get(item, "")

                subtitle_html = f'<div class="section-subtitle">{label}</div>' if label else ""
                st.markdown(
                    f"""
                    {subtitle_html}
                    <div class="big-tile" style="background:{colour}; width:50%; margin-left:auto; margin-right:auto;">
                        {format_pct(val)}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            elif len(items) == 2:
                two_cols = st.columns(2)
                for idx, item in enumerate(items):
                    val = lookup.get(item, {}).get(period, np.nan)
                    colour = heat_colour(val, vmin, vmax)
                    label = labels.get(item, item)

                    with two_cols[idx]:
                        st.markdown(
                            f"""
                            <div class="small-tile" style="background:{colour}; min-height:74px;">
                                <span class="tile-label tile-label-plain">{label}</span>
                                {format_pct(val)}
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )

            elif len(items) == 3:
                broad = items[0]
                broad_val = lookup.get(broad, {}).get(period, np.nan)
                broad_colour = heat_colour(broad_val, vmin, vmax)
                broad_label = labels.get(broad, "")

                broad_label_html = f'<span class="tile-label-on-colour">{broad_label}</span>' if broad_label else ""

                st.markdown(
                    f"""
                    <div class="big-tile" style="background:{broad_colour};">
                        {broad_label_html}
                        {format_pct(broad_val)}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                small_cols = st.columns(2)
                for i, item in enumerate(items[1:]):
                    val = lookup.get(item, {}).get(period, np.nan)
                    colour = heat_colour(val, vmin, vmax)
                    label = labels.get(item, "")

                    with small_cols[i]:
                        st.markdown(
                            f"""
                            <div class="small-tile" style="background:{colour};">
                                <span class="tile-label">{label}</span>
                                {format_pct(val)}
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )

            elif len(items) == 4:
                row1 = st.columns(2)
                row2 = st.columns(2)

                for idx, item in enumerate(items):
                    val = lookup.get(item, {}).get(period, np.nan)
                    colour = heat_colour(val, vmin, vmax)
                    label = labels.get(item, item)

                    target_cols = row1 if idx < 2 else row2
                    target_col = target_cols[idx % 2]

                    with target_col:
                        st.markdown(
                            f"""
                            <div class="small-tile" style="background:{colour}; min-height:74px;">
                                <span class="tile-label tile-label-plain">{label}</span>
                                {format_pct(val)}
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )

            elif len(items) == 5:
                row1 = st.columns(2)
                row2 = st.columns(2)
                row3 = st.columns(1)

                for idx, item in enumerate(items):
                    val = lookup.get(item, {}).get(period, np.nan)
                    colour = heat_colour(val, vmin, vmax)
                    label = labels.get(item, item)

                    if idx < 2:
                        target_col = row1[idx]
                    elif idx < 4:
                        target_col = row2[idx - 2]
                    else:
                        target_col = row3[0]

                    with target_col:
                        width_style = ""
                        if idx == 4:
                            width_style = "width:50%; margin-left:auto; margin-right:auto;"

                        st.markdown(
                            f"""
                            <div class="small-tile" style="background:{colour}; min-height:74px; {width_style}">
                                <span class="tile-label tile-label-plain">{label}</span>
                                {format_pct(val)}
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )

            st.markdown("</div>", unsafe_allow_html=True)
            st.markdown('<div class="spacer"></div>', unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)


# =====================================
# METHODOLOGY FOOTER
# =====================================
st.markdown(
    f'<div class="methodology-text">{get_methodology_paragraph(is_relative_mode, relative_detail_mode, effective_real_mode, inflation_source_note)}</div>',
    unsafe_allow_html=True,
)


# =====================================
# FOOTER
# =====================================
if POWERED_BY_FILE.exists():
    st.markdown(
        f"""
        <div class="footer-bar">
            <div class="footer-logo">
                <img src="data:image/png;base64,{img_to_base64(POWERED_BY_FILE)}">
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =====================================
# DIAGNOSTICS
# =====================================
if "show_diagnostics" not in st.session_state:
    st.session_state["show_diagnostics"] = False

diag_label = (
    "▼ Diagnostics and underlying data"
    if st.session_state["show_diagnostics"]
    else "► Diagnostics and underlying data"
)

if st.button(diag_label, key="diagnostics_toggle_button", use_container_width=True):
    st.session_state["show_diagnostics"] = not st.session_state["show_diagnostics"]

if st.session_state["show_diagnostics"]:
    with st.container(border=True):
        st.markdown('<div class="diag-title">Live fund stitching diagnostics</div>', unsafe_allow_html=True)
        st.markdown(
            f"""
            <div class="diag-note">
                Index history is always preferred where available. Live yfinance history is only used after index history ends.
                Primary ticker is checked first and secondary ticker is used only where required.
                Current dashboard mode: <b>{mode_label}</b>{f" - <b>{relative_detail_label}</b>" if relative_detail_label else ""} - <b>{basis_label}</b>.
                Inflation source: <b>{inflation_source_note}</b>. Dashboard end date: <b>{end_date.strftime("%d/%m/%Y")}</b>.
            </div>
            """,
            unsafe_allow_html=True,
        )

        if inflation_debug_message:
            st.info(inflation_debug_message)

        tabs = st.tabs(["Live stitching", "Returns", "Inflation", "Mapping"])

        with tabs[0]:
            if not live_diag.empty:
                summary = (
                    live_diag["series_type"]
                    .fillna("unknown")
                    .value_counts(dropna=False)
                    .rename_axis("series_type")
                    .reset_index(name="count")
                )

                counts = summary.set_index("series_type")["count"].to_dict()
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Stitched", int(counts.get("stitched", 0)))
                c2.metric("Index only", int(counts.get("index_only", 0)))
                c3.metric("Live only", int(counts.get("live_only", 0)))
                c4.metric("Missing", int(counts.get("missing", 0)))

                st.subheader("Summary")
                st.dataframe(summary, width="stretch", hide_index=True)

                diag_show = format_diagnostic_table(live_diag)
                preferred_cols = [
                    "asset_class",
                    "live_fund_primary",
                    "live_fund_secondary",
                    "selected_ticker",
                    "selected_source",
                    "series_type",
                    "index_last_date",
                    "live_first_date",
                    "live_last_date",
                    "stitch_anchor_date",
                    "note",
                ]
                display_cols = [c for c in preferred_cols if c in diag_show.columns]

                st.subheader("Detailed diagnostics")
                st.dataframe(diag_show[display_cols], width="stretch", hide_index=True)

                csv_bytes = dataframe_to_csv_download(live_diag)
                st.download_button(
                    label="Download live diagnostics (CSV)",
                    data=csv_bytes,
                    file_name="live_fund_stitch_diagnostics.csv",
                    mime="text/csv",
                    use_container_width=False,
                    key="download_live_diagnostics_csv",
                )
            else:
                st.write("No live diagnostics available.")

        with tabs[1]:
            returns_tabs = st.tabs(["Displayed returns", "Nominal display", "Absolute returns"])

            with returns_tabs[0]:
                st.dataframe(convert_pct_table_for_display(returns_df), width="stretch", hide_index=True)

            with returns_tabs[1]:
                st.dataframe(convert_pct_table_for_display(nominal_display_returns_df), width="stretch", hide_index=True)

            with returns_tabs[2]:
                st.dataframe(convert_pct_table_for_display(absolute_returns_df), width="stretch", hide_index=True)

        with tabs[2]:
            c1, c2 = st.columns(2)

            with c1:
                if not inflation_returns_df.empty:
                    st.subheader("Inflation return table")
                    st.dataframe(convert_pct_table_for_display(inflation_returns_df), width="stretch", hide_index=True)
                else:
                    st.write("No inflation return table available.")

            with c2:
                if inflation_levels is not None and not inflation_levels.empty:
                    infl_last = pd.to_datetime(inflation_levels.index.max()).strftime("%d/%m/%Y")
                    jan26_val = inflation_monthly_returns.get(pd.Timestamp("2026-01-31"), np.nan)
                    st.metric("Inflation source", inflation_source_note)
                    st.metric("Inflation last date", infl_last)
                    st.metric("Jan-26 monthly return", "-" if pd.isna(jan26_val) else f"{jan26_val:.4%}")
                else:
                    st.write("No inflation series available.")

            if inflation_levels is not None and not inflation_levels.empty:
                levels_df = inflation_levels.reset_index()
                levels_df.columns = ["Date", "Inflation level"]
                levels_df["Date"] = pd.to_datetime(levels_df["Date"]).dt.strftime("%d/%m/%Y")

                monthly_df = inflation_monthly_returns.reset_index()
                monthly_df.columns = ["Date", "UK inflation monthly return"]
                monthly_df["Date"] = pd.to_datetime(monthly_df["Date"]).dt.strftime("%d/%m/%Y")
                monthly_df["UK inflation monthly return"] = monthly_df["UK inflation monthly return"].map(
                    lambda x: np.nan if pd.isna(x) else round(x * 100, 4)
                )

                lower_tabs = st.tabs(["Inflation levels tail", "Inflation monthly tail"])

                with lower_tabs[0]:
                    st.dataframe(levels_df.tail(24), width="stretch", hide_index=True)

                with lower_tabs[1]:
                    st.dataframe(monthly_df.tail(24), width="stretch", hide_index=True)
            else:
                st.write("No inflation series available.")

        with tabs[3]:
            st.subheader("Mapping table")
            st.dataframe(mapping, width="stretch", hide_index=True)