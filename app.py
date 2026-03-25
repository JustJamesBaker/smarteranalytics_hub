from pathlib import Path
import base64
from io import BytesIO, StringIO
from html.parser import HTMLParser
import re
import zipfile

import altair as alt
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
TEXT_GREY = "#555555"
LIGHT_GREY = "#f3f3f3"
MID_GREY = "#d9d9d9"
CHART_BG_GREY = "#f5f5f5"
WHITE = "#ffffff"
APP_TITLE = "smarteranalytics™ Hub"
PAGE_LABELS = {
    "Dashboard": "Market snapshot",
    "Charts": "Performance analysis",
    "Risk": "Risk analysis",
    "Yield": "Yield curves",
}
DISPLAY_NAME_OVERRIDES = {
    "Global stocks": "Global market",
    "UK stocks": "UK market",
    "Developed stocks": "Developed market",
    "Emerging stocks": "Emerging market",
    "UK value stocks": "UK value",
    "UK small stocks": "UK small",
    "Developed value stocks": "Developed value",
    "Developed small stocks": "Developed small",
    "Emerging value stocks": "Emerging value",
    "Emerging small stocks": "Emerging small",
    "Developed REITs": "REITs",
    "Global GBP hedged bonds (0-5)": "Global bonds (0-5, GBP)",
}

DASHBOARD_HORIZONS = {
    "20Y": "20 Year",
    "10Y": "10 Year",
    "5Y": "5 Year",
    "YTD": "YTD",
}

CHART_PERIODS = {
    "YTD": "YTD",
    "1Y": "1 Year",
    "3Y": "3 Year",
    "5Y": "5 Year",
    "10Y": "10 Year",
    "20Y": "20 Year",
}

RISK_PERIODS = {
    "1Y": "1 Year",
    "3Y": "3 Year",
    "5Y": "5 Year",
    "10Y": "10 Year",
    "20Y": "20 Year",
}

RETURNS_TABLE_PERIODS = ["YTD", "1Y", "3Y", "5Y", "7Y", "10Y", "15Y", "20Y", "25Y", "Period"]

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
            "Global GBP hedged bonds (0-5)": "Global bonds (0-5, GBP)",
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
            "Global GBP hedged bonds (0-5)": "Global bonds (0-5, GBP)",
        },
    },
]

ASSET_CLASS_ALIASES = {
    "Global equity": "Global stocks",
    "Global stocks": "Global stocks",
    "Global market": "Global stocks",
    "World equity": "Global stocks",
    "World stocks": "Global stocks",
    "UK equity": "UK stocks",
    "UK stocks": "UK stocks",
    "UK market": "UK stocks",
    "UK value": "UK value stocks",
    "UK value stocks": "UK value stocks",
    "UK small": "UK small stocks",
    "UK small stocks": "UK small stocks",
    "Developed equity": "Developed stocks",
    "Developed stocks": "Developed stocks",
    "Developed market": "Developed stocks",
    "Developed value": "Developed value stocks",
    "Developed value stocks": "Developed value stocks",
    "Developed small": "Developed small stocks",
    "Developed small stocks": "Developed small stocks",
    "Emerging market equity": "Emerging stocks",
    "Emerging equity": "Emerging stocks",
    "Emerging stocks": "Emerging stocks",
    "Emerging market": "Emerging stocks",
    "Emerging value": "Emerging value stocks",
    "Emerging value stocks": "Emerging value stocks",
    "Emerging small": "Emerging small stocks",
    "Emerging small stocks": "Emerging small stocks",
    "Developed REITs": "Developed REITs",
    "REITs": "Developed REITs",
    "Cash GBP": "Cash (GBP)",
    "Cash (GBP)": "Cash (GBP)",
    "UK Gilts": "UK Gilts (0-5)",
    "Short Gilt": "UK Gilts (0-5)",
    "UK Gilts (0-5)": "UK Gilts (0-5)",
    "UK IL Gilts": "UK IL Gilts (0-5)",
    "Short IL Gilt": "UK IL Gilts (0-5)",
    "UK IL Gilts (0-5)": "UK IL Gilts (0-5)",
    "Global bonds (0-5, GBP)": "Global GBP hedged bonds (0-5)",
    "Global GBP Hedged bonds (0-5)": "Global GBP hedged bonds (0-5)",
    "Global GBP hedged bonds (0-5)": "Global GBP hedged bonds (0-5)",
    "Global Short Bond GBP": "Global GBP hedged bonds (0-5)",
    "GSDB (GBP)": "Global GBP hedged bonds (0-5)",
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
BOE_YIELD_CURVE_ZIP_URL = "https://www.bankofengland.co.uk/-/media/boe/files/statistics/yield-curves/latest-yield-curve-data.zip"
DIVIDENDDATA_INDEX_LINKED_GILTS_URL = "https://www.dividenddata.co.uk/index-linked-gilts-prices-yields.py"

DEFAULT_ASSET_ORDER = [
    "Global stocks",
    "UK stocks",
    "UK value stocks",
    "UK small stocks",
    "Developed stocks",
    "Developed value stocks",
    "Developed small stocks",
    "Emerging stocks",
    "Emerging value stocks",
    "Emerging small stocks",
    "Developed REITs",
    "Cash (GBP)",
    "UK Gilts (0-5)",
    "UK IL Gilts (0-5)",
    "Global GBP hedged bonds (0-5)",
]

DEFAULT_CHART_ASSETS = [
    "Cash (GBP)",
    "UK stocks",
    "Developed stocks",
    "Developed value stocks",
    "Developed small stocks",
    "Emerging stocks",
    "Developed REITs",
    "Global GBP hedged bonds (0-5)",
    "UK IL Gilts (0-5)",
]

ASSET_COLOURS = {
    "Global stocks": "#c95b2b",
    "UK stocks": "#d71921",
    "UK value stocks": "#ef5350",
    "UK small stocks": "#8f1015",
    "Developed stocks": "#f36f21",
    "Developed value stocks": "#ff9a4d",
    "Developed small stocks": "#c95f18",
    "Emerging stocks": "#f4c542",
    "Emerging value stocks": "#ffd965",
    "Emerging small stocks": "#cfa400",
    "Developed REITs": "#2e8b57",
    "Cash (GBP)": "#4b5563",
    "UK Gilts (0-5)": "#1f77b4",
    "UK IL Gilts (0-5)": "#5dade2",
    "Global GBP hedged bonds (0-5)": "#0b5394",
}


# =====================================
# HELPERS
# =====================================
@st.cache_data(show_spinner=False)
def img_to_base64(path: str, file_mtime: float) -> str:
    return base64.b64encode(Path(path).read_bytes()).decode("utf-8")


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
    return "" if ticker.lower() in {"", "nan", "none"} else ticker


def normalise_name(value: object) -> str:
    return str(value).strip().lower()


def display_name(value: object) -> str:
    text = str(value).strip()
    return DISPLAY_NAME_OVERRIDES.get(text, text)


def format_pct(x: float) -> str:
    return "-" if pd.isna(x) else f"{x:.1%}"


def annualised_return_from_growth(growth: float, years: float) -> float:
    if pd.isna(growth) or growth <= 0 or years <= 0:
        return np.nan
    return growth ** (1 / years) - 1


def build_lookup_table(returns_df: pd.DataFrame) -> dict:
    return {} if returns_df.empty else returns_df.set_index("asset_class").to_dict(orient="index")


def heat_colour(value: float, vmin: float, vmax: float) -> str:
    if pd.isna(value):
        return "#E9E9E9"

    if value == 0:
        return "#f6f6f6"

    if value > 0:
        positive_max = max(float(vmax), 0.0)
        if positive_max <= 0:
            return "#f6f6f6"
        norm = min(max(value / positive_max, 0), 1)
        light = np.array([130, 130, 130])
        dark = np.array([0, 170, 95])
        rgb = (light * (1 - norm) + dark * norm).astype(int)
        return f"rgb({rgb[0]},{rgb[1]},{rgb[2]})"

    negative_min = min(float(vmin), 0.0)
    if negative_min >= 0:
        return "#f6f6f6"
    norm = min(max(abs(value) / abs(negative_min), 0), 1)
    light = np.array([130, 130, 130])
    dark = np.array([190, 30, 55])
    rgb = (light * (1 - norm) + dark * norm).astype(int)
    return f"rgb({rgb[0]},{rgb[1]},{rgb[2]})"


def rank_heat_colour(value: float, vmin: float, vmax: float, low_is_good: bool = False) -> str:
    if pd.isna(value):
        return "#E9E9E9"
    if vmax <= vmin:
        return "#f6f6f6"

    norm = (float(value) - float(vmin)) / (float(vmax) - float(vmin))
    norm = min(max(norm, 0), 1)
    if low_is_good:
        norm = 1 - norm

    start = np.array([190, 30, 55])
    end = np.array([0, 170, 95])
    rgb = (start * (1 - norm) + end * norm).astype(int)
    return f"rgb({rgb[0]},{rgb[1]},{rgb[2]})"


def correlation_heat_colour(value: float) -> str:
    if pd.isna(value):
        return "#E9E9E9"
    value = float(value)
    grey = np.array([138, 138, 138])
    green = np.array([0, 170, 95])
    red = np.array([190, 30, 55])

    if np.isclose(value, 1.0):
        rgb = grey
    elif value >= 0.5:
        norm = min(max((value - 0.5) / 0.49, 0), 1)
        rgb = (grey * (1 - norm) + red * norm).astype(int)
    else:
        norm = min(max(value / 0.5, 0), 1)
        rgb = (green * (1 - norm) + grey * norm).astype(int)

    return f"rgb({rgb[0]},{rgb[1]},{rgb[2]})"


def safe_relative_return(asset_return: float, base_return: float) -> float:
    if pd.isna(asset_return) or pd.isna(base_return) or (1 + base_return) <= 0:
        return np.nan
    return ((1 + asset_return) / (1 + base_return)) - 1


def get_display_groups(is_relative_mode: bool, relative_detail_mode: str) -> list[dict]:
    if is_relative_mode and relative_detail_mode == "Minor":
        return DISPLAY_GROUPS_RELATIVE_MINOR
    return DISPLAY_GROUPS_ABSOLUTE


def convert_to_relative_returns(absolute_returns_df: pd.DataFrame, relative_detail_mode: str) -> pd.DataFrame:
    relative_df = absolute_returns_df.copy()
    growth_base_map = MAJOR_GROWTH_BASE_MAP if relative_detail_mode == "Major" else MINOR_GROWTH_BASE_MAP
    all_base_maps = {**growth_base_map, **DEFENSIVE_BASE_MAP}
    asset_to_row = relative_df.set_index("asset_class")

    period_cols = [c for c in relative_df.columns if c != "asset_class"]
    for period in period_cols:
        for idx, row in relative_df.iterrows():
            asset = row["asset_class"]
            base_asset = all_base_maps.get(asset)
            if base_asset is None or base_asset not in asset_to_row.index:
                relative_df.at[idx, period] = np.nan
            else:
                relative_df.at[idx, period] = safe_relative_return(row[period], asset_to_row.at[base_asset, period])

    return relative_df


def find_inflation_column(ts: pd.DataFrame) -> str:
    lookup = {normalise_name(c): str(c).strip() for c in ts.columns}
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
    returns = standardise_series(ts[inflation_col])
    series = pd.Series(returns.values, index=ts["Date"], name="UK inflation").dropna().sort_index()
    if series.empty:
        raise ValueError("Inflation series was found but contains no valid data.")
    levels = (1 + series).cumprod()
    levels.name = "UK inflation"
    return levels


@st.cache_data(show_spinner=False, ttl=86400)
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
    response = requests.get(csv_url, headers=headers, timeout=30)
    response.raise_for_status()
    return pd.read_csv(StringIO(response.text))


@st.cache_data(show_spinner=False, ttl=86400)
def fetch_ons_cpi_index_series(csv_url: str) -> pd.Series:
    raw = fetch_ons_raw_csv(csv_url)
    out, _ = parse_ons_cpi_index_frame(raw)
    return out


def extend_inflation_levels_with_ons_index(existing_inflation_levels: pd.Series, ons_index_series: pd.Series) -> pd.Series:
    workbook = existing_inflation_levels.dropna().sort_index()
    ons = ons_index_series.dropna().sort_index()

    if workbook.empty:
        raise ValueError("Existing inflation level history is required.")

    common_dates = workbook.index.intersection(ons.index)
    if common_dates.empty:
        raise ValueError("No overlapping dates between workbook inflation and ONS CPI index.")

    anchor_date = common_dates.max()
    workbook_anchor = float(workbook.loc[anchor_date])
    ons_anchor = float(ons.loc[anchor_date])

    if pd.isna(ons_anchor) or ons_anchor == 0:
        raise ValueError("Invalid ONS anchor value.")

    scaled_ons = ons * (workbook_anchor / ons_anchor)
    combined = pd.concat([workbook[workbook.index <= anchor_date], scaled_ons[scaled_ons.index > anchor_date]])
    combined = combined[~combined.index.duplicated(keep="last")].sort_index()
    combined.name = "UK inflation"
    return combined


def build_best_available_inflation_levels(ts: pd.DataFrame) -> tuple[pd.Series | None, str, str | None]:
    workbook_levels = build_inflation_levels_from_timeseries(ts)
    try:
        ons_index = fetch_ons_cpi_index_series(ONS_CPI_INDEX_CSV_URL)
        extended = extend_inflation_levels_with_ons_index(workbook_levels, ons_index)
        if extended.index.max() > workbook_levels.index.max():
            return extended, "Workbook time_series + ONS CPI index extension", None
        return workbook_levels, "Workbook time_series", "ONS fetched, but no extension beyond workbook end date."
    except Exception as exc:
        return workbook_levels, "Workbook time_series", f"ONS extension failed: {type(exc).__name__}: {exc}"


def build_monthly_returns_from_levels(levels: pd.Series) -> pd.Series:
    series = levels.dropna().sort_index()
    if series.empty:
        return pd.Series(dtype=float, name="UK inflation")
    returns = series.pct_change()
    returns.name = "UK inflation"
    return returns


def get_live_consistent_end_date(stitched_series_map: dict[str, pd.Series], live_diag: pd.DataFrame) -> pd.Timestamp:
    if live_diag is not None and not live_diag.empty:
        live_rows = live_diag[live_diag["series_type"].isin(["stitched", "live_only"])].copy()
        if not live_rows.empty:
            live_last_dates = pd.to_datetime(live_rows["live_last_date"], errors="coerce").dropna()
            if not live_last_dates.empty:
                return pd.Timestamp(live_last_dates.min())

    last_dates = [pd.Timestamp(s.dropna().index.max()) for s in stitched_series_map.values() if not s.dropna().empty]
    return max(last_dates) if last_dates else pd.Timestamp.today().normalize()


def get_dashboard_end_date(
    stitched_series_map: dict[str, pd.Series],
    live_diag: pd.DataFrame,
    inflation_levels: pd.Series | None,
    is_real_mode: bool,
) -> pd.Timestamp:
    end_date = get_live_consistent_end_date(stitched_series_map, live_diag)
    if is_real_mode and inflation_levels is not None and not inflation_levels.dropna().empty:
        end_date = min(end_date, pd.Timestamp(inflation_levels.dropna().index.max()))
    return end_date


def convert_to_real_returns(nominal_returns_df: pd.DataFrame, inflation_returns_df: pd.DataFrame) -> pd.DataFrame:
    real_df = nominal_returns_df.copy()
    if inflation_returns_df.empty:
        return real_df

    infl_row = inflation_returns_df[inflation_returns_df["asset_class"] == "UK inflation"]
    if infl_row.empty:
        return real_df

    infl_row = infl_row.iloc[0]
    period_cols = [c for c in real_df.columns if c != "asset_class"]
    for period in period_cols:
        infl_val = infl_row.get(period, np.nan)
        for idx, row in real_df.iterrows():
            real_df.at[idx, period] = safe_relative_return(row[period], infl_val)

    return real_df


def get_methodology_paragraph(page_name: str, is_relative_mode: bool, relative_detail_mode: str, is_real_mode: bool, inflation_source_note: str) -> str:
    basis = "real" if is_real_mode else "nominal"

    if page_name == "Dashboard":
        prefix = f"This tab shows annualised {basis} GBP returns across the displayed horizons, with YTD shown cumulatively."
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
    else:
        prefix = (
            "This tab shows return in GBP from the common inception date to the selected end date, "
            "with the chart normalised to a starting value of £1.00. "
            "For the growth of wealth chart, daily live-fund history is preferred where available and monthly index history is stitched in before it to extend the chart further back."
        )
        relative_text = ""

    inflation_text = ""
    if is_real_mode:
        inflation_text = (
            f" Real returns are calculated using UK inflation via (1+asset return)/(1+inflation return)-1. "
            f"Current inflation source: {inflation_source_note}."
        )

    source_text = (
        " Albion index series history is used as the preferred source where available. "
        "For the dashboard and tables, live yfinance mappings are only used after index history ends. "
        "More information on the Albion indices can be found at "
        '<a href="https://smartersuccess.net/indices" target="_blank">smartersuccess.net/indices</a>.'
    )
    return prefix + relative_text + inflation_text + source_text


def dataframe_to_csv_download(df: pd.DataFrame) -> bytes:
    return prepare_dataframe_for_display(df).to_csv(index=False).encode("utf-8")


def prepare_dataframe_for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if out.empty:
        return out

    for col in out.columns:
        series = out[col]

        if pd.api.types.is_datetime64_any_dtype(series):
            out[col] = pd.to_datetime(series, errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
            continue

        if pd.api.types.is_timedelta64_dtype(series):
            out[col] = series.astype("string").fillna("")
            continue

        if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):
            non_null = series.dropna()
            if non_null.empty:
                out[col] = series.astype("string").fillna("")
                continue

            non_null_str = non_null.astype(str).str.strip()
            date_like_mask = (
                non_null_str.str.match(r"^\d{1,2}/\d{1,2}/\d{4}$", na=False)
                | non_null_str.str.match(r"^\d{4}-\d{2}-\d{2}$", na=False)
                | non_null_str.str.match(r"^\d{4}/\d{2}/\d{2}$", na=False)
            )
            if len(non_null_str) > 0 and date_like_mask.all():
                full_str = series.astype("string").fillna("").str.strip()
                parsed_dates = pd.Series(pd.NaT, index=out.index, dtype="datetime64[ns]")

                slash_mask = full_str.str.match(r"^\d{1,2}/\d{1,2}/\d{4}$", na=False)
                if slash_mask.any():
                    parsed_dates.loc[slash_mask] = pd.to_datetime(
                        full_str.loc[slash_mask],
                        format="%d/%m/%Y",
                        errors="coerce",
                    )

                iso_dash_mask = full_str.str.match(r"^\d{4}-\d{2}-\d{2}$", na=False)
                if iso_dash_mask.any():
                    parsed_dates.loc[iso_dash_mask] = pd.to_datetime(
                        full_str.loc[iso_dash_mask],
                        format="%Y-%m-%d",
                        errors="coerce",
                    )

                iso_slash_mask = full_str.str.match(r"^\d{4}/\d{2}/\d{2}$", na=False)
                if iso_slash_mask.any():
                    parsed_dates.loc[iso_slash_mask] = pd.to_datetime(
                        full_str.loc[iso_slash_mask],
                        format="%Y/%m/%d",
                        errors="coerce",
                    )

                out[col] = parsed_dates.dt.strftime("%d/%m/%Y").fillna("")
                continue

            numeric = pd.to_numeric(non_null, errors="coerce")
            if len(non_null) > 0 and numeric.notna().sum() == len(non_null):
                out[col] = pd.to_numeric(series, errors="coerce")
                continue

            out[col] = series.astype("string").fillna("")

    return out


def format_diag_date(value: object) -> str:
    ts = pd.to_datetime(value, errors="coerce")
    return "" if pd.isna(ts) else ts.strftime("%d/%m/%Y")


def format_diagnostic_table(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in ["index_last_date", "live_first_date", "live_last_date", "stitch_anchor_date"]:
        if col in out.columns:
            out[col] = pd.to_datetime(out[col], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
    return prepare_dataframe_for_display(out)


def build_series_summary(series: pd.Series | None) -> dict[str, object]:
    s = pd.Series(dtype=float) if series is None else series.dropna().sort_index()
    if s.empty:
        return {
            "points": 0,
            "first_date": pd.NaT,
            "last_date": pd.NaT,
            "last_value": np.nan,
        }
    return {
        "points": int(len(s)),
        "first_date": pd.Timestamp(s.index.min()),
        "last_date": pd.Timestamp(s.index.max()),
        "last_value": float(s.iloc[-1]),
    }


def parse_ons_cpi_index_frame(raw: pd.DataFrame) -> tuple[pd.Series, pd.DataFrame]:
    raw = raw.copy()
    raw.columns = [str(c).strip() for c in raw.columns]

    if raw.shape[1] < 2:
        raise ValueError("ONS CPI index CSV did not contain at least two columns.")

    date_col, value_col = raw.columns[:2]
    df = raw[[date_col, value_col]].copy()
    df.columns = ["date_raw", "value_raw"]

    df["date_raw"] = df["date_raw"].astype(str).str.strip().str.upper()
    df["value"] = pd.to_numeric(df["value_raw"], errors="coerce")
    df["Date"] = pd.NaT
    df["match_type"] = ""

    mask_yyyy_mon = df["date_raw"].str.match(r"^\d{4}\s+[A-Z]{3}$", na=False)
    if mask_yyyy_mon.any():
        df.loc[mask_yyyy_mon, "Date"] = pd.to_datetime(
            "01 " + df.loc[mask_yyyy_mon, "date_raw"],
            format="%d %Y %b",
            errors="coerce",
        )
        df.loc[mask_yyyy_mon, "match_type"] = "YYYY MON"

    mask_mon_yyyy = df["date_raw"].str.match(r"^[A-Z]{3}\s+\d{4}$", na=False)
    if mask_mon_yyyy.any():
        df.loc[mask_mon_yyyy, "Date"] = pd.to_datetime(
            "01 " + df.loc[mask_mon_yyyy, "date_raw"],
            format="%d %b %Y",
            errors="coerce",
        )
        df.loc[mask_mon_yyyy, "match_type"] = "MON YYYY"

    mask_year = df["date_raw"].str.match(r"^\d{4}$", na=False)
    if mask_year.any():
        df.loc[mask_year, "Date"] = pd.to_datetime(
            df.loc[mask_year, "date_raw"] + "-01-01",
            format="%Y-%m-%d",
            errors="coerce",
        )
        df.loc[mask_year, "match_type"] = "YYYY"

    parsed = df.dropna(subset=["value", "Date"]).copy()
    parsed["Date"] = pd.to_datetime(parsed["Date"]) + MonthEnd(0)
    parsed = parsed.sort_values("Date")

    out = pd.Series(parsed["value"].values, index=parsed["Date"], name="ONS CPI index")
    out = out[~out.index.duplicated(keep="last")].sort_index()

    if out.empty:
        raise ValueError("ONS CPI index CSV produced no valid rows.")
    return out, df


def build_ons_fetch_diagnostics(csv_url: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    try:
        raw = fetch_ons_raw_csv(csv_url)
        parsed_series, parsed_df = parse_ons_cpi_index_frame(raw)
        summary = pd.DataFrame(
            [
                {"metric": "Fetch status", "value": "OK"},
                {"metric": "Source URL", "value": csv_url},
                {"metric": "Raw rows", "value": int(len(raw))},
                {"metric": "Raw columns", "value": int(raw.shape[1])},
                {"metric": "Parsed numeric rows", "value": int(parsed_df["value"].notna().sum())},
                {"metric": "Parsed dated rows", "value": int(parsed_df["Date"].notna().sum())},
                {"metric": "Output rows", "value": int(len(parsed_series))},
                {"metric": "Output first date", "value": format_diag_date(parsed_series.index.min())},
                {"metric": "Output last date", "value": format_diag_date(parsed_series.index.max())},
                {"metric": "Output last value", "value": round(float(parsed_series.iloc[-1]), 4)},
            ]
        )

        preview = parsed_df.copy()
        preview["Date"] = pd.to_datetime(preview["Date"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
        return summary, preview
    except Exception as exc:
        summary = pd.DataFrame(
            [
                {"metric": "Fetch status", "value": "Failed"},
                {"metric": "Source URL", "value": csv_url},
                {"metric": "Error type", "value": type(exc).__name__},
                {"metric": "Error", "value": str(exc)},
            ]
        )
        return summary, pd.DataFrame()


@st.cache_data(show_spinner=False, ttl=43200)
def fetch_boe_yield_curve_zip(zip_url: str) -> bytes:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/146.0.0.0 Safari/537.36"
        )
    }
    response = requests.get(zip_url, headers=headers, timeout=45)
    response.raise_for_status()
    return response.content


@st.cache_data(show_spinner=False, ttl=43200)
def fetch_dividenddata_html(page_url: str) -> str:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/146.0.0.0 Safari/537.36"
        )
    }
    response = requests.get(page_url, headers=headers, timeout=45)
    response.raise_for_status()
    return response.text


class SimpleHTMLTableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.tables: list[list[list[str]]] = []
        self._in_table = False
        self._in_row = False
        self._in_cell = False
        self._current_table: list[list[str]] = []
        self._current_row: list[str] = []
        self._current_cell: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag == "table":
            self._in_table = True
            self._current_table = []
        elif tag == "tr" and self._in_table:
            self._in_row = True
            self._current_row = []
        elif tag in {"td", "th"} and self._in_row:
            self._in_cell = True
            self._current_cell = []

    def handle_data(self, data: str) -> None:
        if self._in_cell:
            self._current_cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag in {"td", "th"} and self._in_cell:
            cell_text = " ".join(part.strip() for part in self._current_cell if part.strip()).strip()
            self._current_row.append(cell_text)
            self._in_cell = False
            self._current_cell = []
        elif tag == "tr" and self._in_row:
            if self._current_row:
                self._current_table.append(self._current_row)
            self._in_row = False
            self._current_row = []
        elif tag == "table" and self._in_table:
            if self._current_table:
                self.tables.append(self._current_table)
            self._in_table = False
            self._current_table = []


class TextExtractingHTMLParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        text = str(data).strip()
        if text:
            self.parts.append(text)


def pick_boe_zip_member(member_names: list[str], target_text: str) -> str:
    target = target_text.strip().lower()
    candidates = [name for name in member_names if target in name.lower()]
    if not candidates:
        raise FileNotFoundError(f"Could not find '{target_text}' in BOE yield-curve zip.")
    return sorted(candidates)[0]


def parse_maturity_label(value: object) -> float:
    if pd.isna(value):
        return np.nan
    text = str(value).strip()
    numeric = pd.to_numeric(text, errors="coerce")
    if pd.notna(numeric):
        return float(numeric)
    match = re.search(r"(\d+(?:\.\d+)?)", text)
    return float(match.group(1)) if match else np.nan


def parse_time_to_maturity(value: object) -> float:
    if pd.isna(value):
        return np.nan
    text = str(value).strip().lower()
    if not text:
        return np.nan

    years = 0.0
    months = 0.0
    days = 0.0

    year_match = re.search(r"(\d+(?:\.\d+)?)\s*year", text)
    month_match = re.search(r"(\d+(?:\.\d+)?)\s*month", text)
    day_match = re.search(r"(\d+(?:\.\d+)?)\s*day", text)

    if year_match:
        years = float(year_match.group(1))
    if month_match:
        months = float(month_match.group(1))
    if day_match:
        days = float(day_match.group(1))

    if years == 0 and months == 0 and days == 0:
        numeric = pd.to_numeric(text, errors="coerce")
        return float(numeric) if pd.notna(numeric) else np.nan

    return years + (months / 12.0) + (days / 365.25)


def parse_dividenddata_text_fallback(html: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    parser = TextExtractingHTMLParser()
    parser.feed(html)
    text = "\n".join(parser.parts)

    row_pattern = re.compile(
        r"(?P<epic>[A-Z0-9]+)\s+.*?\s+(?P<maturity>\d+\s+years?(?:\s+\d+\s+days?)?)\s+£.*?(?P<real_yield>-?\d+(?:\.\d+)?)%",
        flags=re.IGNORECASE,
    )
    matches = list(row_pattern.finditer(text))
    if not matches:
        raise ValueError("DividendData fallback text parser could not find index-linked gilt rows.")

    rows = []
    for match in matches:
        maturity_text = match.group("maturity")
        real_yield = pd.to_numeric(match.group("real_yield"), errors="coerce")
        maturity_years = parse_time_to_maturity(maturity_text)
        if pd.isna(real_yield) or pd.isna(maturity_years):
            continue
        rows.append(
            {
                "maturity_years": maturity_years,
                "yield_percent": float(real_yield),
                "curve_type": "Real",
                "curve_date": pd.Timestamp.today().normalize(),
                "source": "DividendData",
                "epic": match.group("epic"),
                "time_to_maturity": maturity_text,
                "real_yield_raw": f"{match.group('real_yield')}%",
            }
        )

    out = pd.DataFrame(rows).dropna(subset=["maturity_years", "yield_percent"])
    out = out.sort_values("maturity_years").drop_duplicates(subset=["maturity_years"], keep="first")
    preview = out[["epic", "time_to_maturity", "real_yield_raw", "maturity_years", "yield_percent", "source"]].copy()
    return out[["maturity_years", "yield_percent", "curve_type", "curve_date", "source"]], preview


def parse_boe_spot_curve_workbook(workbook_bytes: bytes, curve_type: str) -> tuple[pd.DataFrame, pd.Timestamp, pd.DataFrame]:
    raw = pd.read_excel(BytesIO(workbook_bytes), sheet_name="4. spot curve", header=None)
    if raw.shape[0] < 6 or raw.shape[1] < 2:
        raise ValueError(f"BOE {curve_type} spot curve sheet is missing expected rows/columns.")

    candidate_rows = []
    for row_idx in range(min(10, len(raw))):
        parsed_row = raw.iloc[row_idx, 1:].map(parse_maturity_label)
        candidate_rows.append((row_idx, int(parsed_row.notna().sum()), parsed_row))

    maturity_row_idx, maturity_count, maturities = max(candidate_rows, key=lambda x: x[1])
    if maturity_count == 0:
        raise ValueError(f"BOE {curve_type} spot curve maturities could not be parsed.")

    data_rows = raw.iloc[maturity_row_idx + 1 :, :].copy()
    data_rows = data_rows.rename(columns={0: "date_raw"})
    data_rows["curve_date"] = pd.to_datetime(data_rows["date_raw"], errors="coerce")
    value_block = data_rows.iloc[:, 1 : 1 + len(maturities)].apply(pd.to_numeric, errors="coerce")

    valid_mask = data_rows["curve_date"].notna() & value_block.notna().any(axis=1)
    valid_rows = data_rows.loc[valid_mask].copy()
    valid_values = value_block.loc[valid_mask].copy()
    if valid_rows.empty:
        raise ValueError(f"BOE {curve_type} spot curve sheet had no valid dated rows.")

    latest_idx = valid_rows.index[-1]
    curve_date = pd.Timestamp(valid_rows.loc[latest_idx, "curve_date"])
    latest_values = valid_values.loc[latest_idx]

    points = pd.DataFrame(
        {
            "maturity_years": pd.to_numeric(maturities, errors="coerce"),
            "yield_percent": pd.to_numeric(latest_values, errors="coerce"),
            "curve_type": curve_type,
            "curve_date": curve_date,
        }
    ).dropna(subset=["maturity_years", "yield_percent"])

    preview = valid_rows[["curve_date"]].copy()
    preview["points_available"] = valid_values.notna().sum(axis=1).values
    preview["curve_type"] = curve_type
    preview["maturity_row_idx"] = maturity_row_idx + 1
    preview["curve_date"] = pd.to_datetime(preview["curve_date"], errors="coerce")
    return points.sort_values("maturity_years"), curve_date, preview


def fetch_dividenddata_real_yield_extension(page_url: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    html = fetch_dividenddata_html(page_url)
    try:
        parser = SimpleHTMLTableParser()
        parser.feed(html)

        target_rows = None
        header = None
        maturity_idx = None
        real_yield_idx = None
        for table_rows in parser.tables:
            if len(table_rows) < 3:
                continue
            max_cols = max(len(r) for r in table_rows[:3])
            candidate_headers = []
            for depth in (1, 2, 3):
                rows_to_merge = table_rows[:depth]
                merged = []
                for col_idx in range(max_cols):
                    parts = []
                    for row in rows_to_merge:
                        if col_idx < len(row):
                            cell = str(row[col_idx]).strip()
                            if cell:
                                parts.append(cell)
                    merged.append(" ".join(parts).strip())
                candidate_headers.append((depth, merged))

            for depth, merged_header in candidate_headers:
                header_lookup = {col.lower(): idx for idx, col in enumerate(merged_header)}
                maturity_idx = next((idx for col, idx in header_lookup.items() if "time to maturity" in col), None)
                real_yield_idx = next((idx for col, idx in header_lookup.items() if "real yield" in col), None)
                if maturity_idx is not None and real_yield_idx is not None:
                    target_rows = table_rows[depth:]
                    header = merged_header
                    break
            if target_rows is not None:
                break

        if not target_rows or header is None or maturity_idx is None or real_yield_idx is None:
            raise ValueError("DividendData table with 'Time to Maturity' and 'Real Yield' columns was not found.")

        body = target_rows
        table = pd.DataFrame(body, columns=header)
        maturity_col = header[maturity_idx]
        real_yield_col = header[real_yield_idx]
        out = pd.DataFrame(
            {
                "maturity_years": table[maturity_col].map(parse_time_to_maturity),
                "yield_percent": pd.to_numeric(
                    table[real_yield_col].astype(str).str.replace("%", "", regex=False).str.replace(",", "", regex=False),
                    errors="coerce",
                ),
                "curve_type": "Real",
                "curve_date": pd.Timestamp.today().normalize(),
                "source": "DividendData",
            }
        ).dropna(subset=["maturity_years", "yield_percent"])
        out = out.sort_values("maturity_years").drop_duplicates(subset=["maturity_years"], keep="first")

        preview = table[[maturity_col, real_yield_col]].copy()
        preview.columns = ["time_to_maturity", "real_yield_raw"]
        preview["maturity_years"] = preview["time_to_maturity"].map(parse_time_to_maturity)
        preview["yield_percent"] = pd.to_numeric(
            preview["real_yield_raw"].astype(str).str.replace("%", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce",
        )
        preview["source"] = "DividendData"
        return out, preview
    except Exception:
        return parse_dividenddata_text_fallback(html)


def build_boe_yield_curve_diagnostics(zip_url: str, dividenddata_url: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    try:
        zip_bytes = fetch_boe_yield_curve_zip(zip_url)
        with zipfile.ZipFile(BytesIO(zip_bytes)) as zf:
            members = zf.namelist()
            nominal_member = pick_boe_zip_member(members, "GLC Nominal daily data current month")
            real_member = pick_boe_zip_member(members, "GLC Real daily data current month")
            nominal_points, nominal_date, nominal_preview = parse_boe_spot_curve_workbook(
                zf.read(nominal_member),
                "Nominal",
            )
            real_points, real_date, real_preview = parse_boe_spot_curve_workbook(
                zf.read(real_member),
                "Real",
            )

        nominal_points["source"] = "BOE"
        real_points["source"] = "BOE"

        extension_points, extension_preview = fetch_dividenddata_real_yield_extension(dividenddata_url)
        short_real_extension = pd.DataFrame(columns=real_points.columns)
        if not real_points.empty and not extension_points.empty:
            real_min_maturity = float(real_points["maturity_years"].min())
            short_real_extension = extension_points[extension_points["maturity_years"] < real_min_maturity].copy()
            if not short_real_extension.empty:
                short_real_extension["curve_date"] = real_date

        real_combined = pd.concat([short_real_extension, real_points], ignore_index=True)
        real_combined = real_combined.sort_values("maturity_years").drop_duplicates(subset=["maturity_years"], keep="last")
        combined = pd.concat([nominal_points, real_combined], ignore_index=True)
        summary = pd.DataFrame(
            [
                {"metric": "Fetch status", "value": "OK"},
                {"metric": "Source URL", "value": zip_url},
                {"metric": "Extension URL", "value": dividenddata_url},
                {"metric": "Archive members", "value": int(len(members))},
                {"metric": "Nominal workbook", "value": nominal_member},
                {"metric": "Nominal latest date", "value": format_diag_date(nominal_date)},
                {"metric": "Nominal points", "value": int(len(nominal_points))},
                {"metric": "Real workbook", "value": real_member},
                {"metric": "Real latest date", "value": format_diag_date(real_date)},
                {"metric": "Real points", "value": int(len(real_points))},
                {"metric": "Real extension points", "value": int(len(short_real_extension))},
                {
                    "metric": "Real extension max maturity",
                    "value": "-" if short_real_extension.empty else round(float(short_real_extension["maturity_years"].max()), 3),
                },
            ]
        )
        preview = pd.concat([nominal_preview.tail(10), real_preview.tail(10), extension_preview.head(10)], ignore_index=True)
        preview["curve_date"] = pd.to_datetime(preview["curve_date"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
        return combined, summary, preview
    except Exception as exc:
        summary = pd.DataFrame(
            [
                {"metric": "Fetch status", "value": "Failed"},
                {"metric": "Source URL", "value": zip_url},
                {"metric": "Extension URL", "value": dividenddata_url},
                {"metric": "Error type", "value": type(exc).__name__},
                {"metric": "Error", "value": str(exc)},
            ]
        )
        return pd.DataFrame(columns=["maturity_years", "yield_percent", "curve_type", "curve_date"]), summary, pd.DataFrame()


def build_asset_coverage_table(
    mapping: pd.DataFrame,
    monthly_levels: dict[str, pd.Series],
    stitched_series_map: dict[str, pd.Series],
    chart_series_map: dict[str, pd.Series],
    live_diag: pd.DataFrame,
    chart_diag: pd.DataFrame,
) -> pd.DataFrame:
    mapping_view = mapping.drop_duplicates(subset=["asset_class"]).copy()
    rows = []

    for _, row in mapping_view.iterrows():
        asset_class = str(row.get("asset_class", "")).strip()
        if not asset_class:
            continue

        index_summary = build_series_summary(monthly_levels.get(asset_class))
        dashboard_summary = build_series_summary(stitched_series_map.get(asset_class))
        chart_summary = build_series_summary(chart_series_map.get(asset_class))

        live_row = (
            live_diag[live_diag["asset_class"] == asset_class].iloc[0].to_dict()
            if live_diag is not None and not live_diag.empty and asset_class in live_diag["asset_class"].values
            else {}
        )
        chart_row = (
            chart_diag[chart_diag["asset_class"] == asset_class].iloc[0].to_dict()
            if chart_diag is not None and not chart_diag.empty and asset_class in chart_diag["asset_class"].values
            else {}
        )

        rows.append(
            {
                "asset_class": asset_class,
                "index_name": row.get("index_name", ""),
                "live_fund_primary": row.get("live_fund_primary", ""),
                "live_fund_secondary": row.get("live_fund_secondary", ""),
                "selected_ticker": live_row.get("selected_ticker", ""),
                "dashboard_series_type": live_row.get("series_type", ""),
                "chart_series_type": chart_row.get("series_type", ""),
                "index_points": index_summary["points"],
                "index_first_date": index_summary["first_date"],
                "index_last_date": index_summary["last_date"],
                "dashboard_points": dashboard_summary["points"],
                "dashboard_last_date": dashboard_summary["last_date"],
                "chart_points": chart_summary["points"],
                "chart_first_date": chart_summary["first_date"],
                "chart_last_date": chart_summary["last_date"],
            }
        )

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    for col in [
        "index_first_date",
        "index_last_date",
        "dashboard_last_date",
        "chart_first_date",
        "chart_last_date",
    ]:
        out[col] = pd.to_datetime(out[col], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")

    return order_asset_rows(out)


def build_mapping_diagnostics_table(mapping: pd.DataFrame, ts: pd.DataFrame) -> pd.DataFrame:
    rows = []
    ts_columns = {str(c).strip() for c in ts.columns}
    asset_counts = mapping["asset_class"].value_counts(dropna=False).to_dict() if "asset_class" in mapping.columns else {}

    for _, row in mapping.iterrows():
        index_name = str(row.get("index_name", "")).strip()
        asset_class = str(row.get("asset_class", "")).strip()
        primary = normalise_ticker(row.get("live_fund_primary", ""))
        secondary = normalise_ticker(row.get("live_fund_secondary", ""))

        issues = []
        if not asset_class:
            issues.append("Missing asset_class")
        if not index_name:
            issues.append("Missing index_name")
        elif index_name not in ts_columns:
            issues.append("Index column missing from time_series")
        if asset_class and asset_counts.get(asset_class, 0) > 1:
            issues.append("Duplicate asset_class in mapping")
        if not primary and not secondary:
            issues.append("No live ticker")

        rows.append(
            {
                "asset_class": asset_class,
                "index_name": index_name,
                "live_fund_primary": primary,
                "live_fund_secondary": secondary,
                "status": "Issue" if issues else "OK",
                "issues": "; ".join(issues) if issues else "",
            }
        )

    return order_asset_rows(pd.DataFrame(rows))


def build_live_price_diagnostics(prices_df: pd.DataFrame) -> pd.DataFrame:
    if prices_df is None or prices_df.empty:
        return pd.DataFrame(columns=["ticker", "points", "first_date", "last_date", "last_price"])

    rows = []
    for ticker in sorted(prices_df.columns):
        series = pd.to_numeric(prices_df[ticker], errors="coerce").dropna().sort_index()
        if series.empty:
            rows.append(
                {
                    "ticker": ticker,
                    "points": 0,
                    "first_date": "",
                    "last_date": "",
                    "last_price": np.nan,
                }
            )
            continue
        rows.append(
            {
                "ticker": ticker,
                "points": int(len(series)),
                "first_date": format_diag_date(series.index.min()),
                "last_date": format_diag_date(series.index.max()),
                "last_price": round(float(series.iloc[-1]), 4),
            }
        )
    return pd.DataFrame(rows)


def build_return_anchor_table(
    series_map: dict[str, pd.Series],
    end_date: pd.Timestamp,
    period_keys: list[str],
    whole_period_start: pd.Timestamp,
) -> pd.DataFrame:
    rows = []
    years_map = {"1Y": 1, "3Y": 3, "5Y": 5, "7Y": 7, "10Y": 10, "15Y": 15, "20Y": 20, "25Y": 25}

    for asset_class, series in series_map.items():
        s = series.dropna().sort_index()
        if s.empty:
            continue

        end_anchor_date = pd.Timestamp(s.index.max()) if not s[s.index <= end_date].empty else pd.NaT
        for period_key in period_keys:
            if period_key == "YTD":
                base_date, base_level = nearest_level_on_or_before(s, pd.Timestamp(end_date.year - 1, 12, 31))
            elif period_key == "Period":
                base_date, base_level = nearest_level_on_or_before(s, whole_period_start)
            else:
                years = years_map.get(period_key)
                base_date, base_level = nearest_level_on_or_before(s, end_date - pd.DateOffset(years=years))

            rows.append(
                {
                    "asset_class": asset_class,
                    "period": period_key,
                    "base_date_used": format_diag_date(base_date),
                    "end_date_used": format_diag_date(end_anchor_date),
                    "base_level": np.nan if base_level is None else round(float(base_level), 6),
                    "end_level": round(float(s[s.index <= end_date].iloc[-1]), 6) if not s[s.index <= end_date].empty else np.nan,
                    "return_value": calc_period_return(s, end_date, period_key, whole_period_start=whole_period_start),
                }
            )

    out = pd.DataFrame(rows)
    if not out.empty:
        out["return_value"] = out["return_value"].map(lambda x: np.nan if pd.isna(x) else round(x * 100, 4))
    return order_asset_rows(out)


def years_between(start_date: pd.Timestamp, end_date: pd.Timestamp) -> float:
    return max((pd.Timestamp(end_date) - pd.Timestamp(start_date)).days / 365.25, 0)


def nearest_level_on_or_before(series: pd.Series, date: pd.Timestamp) -> tuple[pd.Timestamp | None, float | None]:
    s = series.dropna().sort_index()
    s = s[s.index <= date]
    if s.empty:
        return None, None
    return pd.Timestamp(s.index[-1]), float(s.iloc[-1])


def calc_period_return(series: pd.Series, end_date: pd.Timestamp, period_key: str, whole_period_start: pd.Timestamp | None = None) -> float:
    s = series.dropna().sort_index()
    s = s[s.index <= end_date]
    if s.empty:
        return np.nan

    end_level = float(s.iloc[-1])

    if period_key == "YTD":
        _, base_level = nearest_level_on_or_before(s, pd.Timestamp(end_date.year - 1, 12, 31))
        return (end_level / base_level - 1) if base_level and base_level > 0 else np.nan

    if period_key == "Period":
        if whole_period_start is not None:
            base_date, base_level = nearest_level_on_or_before(s, whole_period_start)
        else:
            base_date, base_level = pd.Timestamp(s.index.min()), float(s.iloc[0])

        if base_level is None or base_level <= 0:
            return np.nan

        years = years_between(base_date, s.index[-1])
        growth = end_level / base_level
        return annualised_return_from_growth(growth, years)

    years_map = {"1Y": 1, "3Y": 3, "5Y": 5, "7Y": 7, "10Y": 10, "15Y": 15, "20Y": 20, "25Y": 25}
    years = years_map.get(period_key)
    if years is None:
        return np.nan

    _, base_level = nearest_level_on_or_before(s, end_date - pd.DateOffset(years=years))
    if base_level is None or base_level <= 0:
        return np.nan

    growth = end_level / base_level
    return annualised_return_from_growth(growth, years)


@st.cache_data(show_spinner=False)
def load_data(file_path: str, file_mtime: float) -> tuple[pd.DataFrame, pd.DataFrame]:
    ts = pd.read_excel(file_path, sheet_name=TIME_SERIES_SHEET)
    mapping = pd.read_excel(file_path, sheet_name=MAPPING_SHEET)

    ts.columns = [str(c).strip() for c in ts.columns]
    mapping.columns = [str(c).strip() for c in mapping.columns]

    ts = ts.copy()
    mapping = mapping.copy()

    ts.iloc[:, 0] = pd.to_datetime(ts.iloc[:, 0], errors="coerce")
    ts = ts.rename(columns={ts.columns[0]: "Date"})
    ts = ts.dropna(subset=["Date"]).sort_values("Date").reset_index(drop=True)

    normalized_mapping_cols = {str(c).strip().lower(): str(c).strip() for c in mapping.columns}
    rename_map = {}
    explicit_name_map = {
        "index_name": "index_name",
        "asset_class": "asset_class",
        "live_fund_primary": "live_fund_primary",
        "live_fund_secondary": "live_fund_secondary",
        "growth/defensive": "growth_defensive",
        "growth_defensive": "growth_defensive",
    }
    for source_name, target_name in explicit_name_map.items():
        if source_name in normalized_mapping_cols:
            rename_map[normalized_mapping_cols[source_name]] = target_name

    if "index_name" not in rename_map.values() and len(mapping.columns) >= 1:
        rename_map[mapping.columns[0]] = "index_name"
    if "asset_class" not in rename_map.values() and len(mapping.columns) >= 2:
        rename_map[mapping.columns[1]] = "asset_class"
    if "live_fund_primary" not in rename_map.values() and len(mapping.columns) >= 3:
        rename_map[mapping.columns[2]] = "live_fund_primary"
    if "live_fund_secondary" not in rename_map.values() and len(mapping.columns) >= 4:
        rename_map[mapping.columns[3]] = "live_fund_secondary"
    mapping = mapping.rename(columns=rename_map)

    if "index_name" in mapping.columns:
        mapping["index_name"] = mapping["index_name"].astype(str).str.strip()

    if "asset_class" in mapping.columns:
        mapping["asset_class"] = mapping["asset_class"].astype(str).str.strip().replace(ASSET_CLASS_ALIASES)

    if "live_fund_primary" in mapping.columns:
        mapping["live_fund_primary"] = mapping["live_fund_primary"].map(normalise_ticker)

    if "live_fund_secondary" in mapping.columns:
        mapping["live_fund_secondary"] = mapping["live_fund_secondary"].map(normalise_ticker)

    if "growth_defensive" in mapping.columns:
        mapping["growth_defensive"] = pd.to_numeric(mapping["growth_defensive"], errors="coerce")

    valid_rows = mapping[mapping["index_name"].isin(ts.columns)].copy()
    return ts, valid_rows


@st.cache_data(show_spinner=False)
def build_monthly_index_levels(ts: pd.DataFrame, mapping: pd.DataFrame) -> dict[str, pd.Series]:
    output = {}
    for _, row in mapping.iterrows():
        asset_class = row["asset_class"]
        index_name = row["index_name"]

        if index_name not in ts.columns:
            continue

        returns = standardise_series(ts[index_name])
        series = pd.Series(returns.values, index=ts["Date"], name=asset_class).dropna().sort_index()
        if series.empty:
            continue

        levels = (1 + series).cumprod()
        levels.name = asset_class
        output[asset_class] = levels

    return output


@st.cache_data(show_spinner=False, ttl=43200)
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
                if "Close" in sub.columns:
                    close_frames.append(pd.to_numeric(sub["Close"], errors="coerce").rename(ticker))
    else:
        if "Close" in data.columns and len(tickers) == 1:
            close_frames.append(pd.to_numeric(data["Close"], errors="coerce").rename(tickers[0]))

    if not close_frames:
        return pd.DataFrame()

    out = pd.concat(close_frames, axis=1)
    out.index = pd.to_datetime(out.index).tz_localize(None)
    return out.sort_index()


def get_price_series(prices_df: pd.DataFrame, ticker: str) -> pd.Series:
    if prices_df.empty or ticker not in prices_df.columns:
        return pd.Series(dtype=float)
    return pd.to_numeric(prices_df[ticker], errors="coerce").dropna().sort_index()


def pick_live_ticker_for_asset(row: pd.Series, prices_df: pd.DataFrame) -> tuple[str, str, str, pd.Series]:
    primary = normalise_ticker(row.get("live_fund_primary", ""))
    secondary = normalise_ticker(row.get("live_fund_secondary", ""))

    primary_series = get_price_series(prices_df, primary) if primary else pd.Series(dtype=float)
    secondary_series = get_price_series(prices_df, secondary) if secondary else pd.Series(dtype=float)

    if not primary_series.empty:
        return primary, "primary", "Primary available in yfinance", primary_series
    if not secondary_series.empty:
        return secondary, "secondary", "Primary unavailable; using secondary", secondary_series
    if primary and secondary:
        return "", "none", "Neither primary nor secondary returned price history", pd.Series(dtype=float)
    if primary:
        return "", "none", "Primary returned no price history and no secondary provided", pd.Series(dtype=float)
    if secondary:
        return "", "none", "No primary provided; secondary returned no price history", pd.Series(dtype=float)
    return "", "none", "No live fund tickers provided", pd.Series(dtype=float)


def build_stitched_asset_series(monthly_levels: dict[str, pd.Series], mapping: pd.DataFrame, prices_df: pd.DataFrame) -> tuple[dict[str, pd.Series], pd.DataFrame]:
    stitched = {}
    diag_rows = []
    deduped_mapping = mapping.drop_duplicates(subset=["asset_class"], keep="last").copy()

    for asset_class in sorted(deduped_mapping["asset_class"].dropna().astype(str).unique()):
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

        row = row_match.iloc[0] if not row_match.empty else pd.Series(dtype=object)
        primary = normalise_ticker(row.get("live_fund_primary", ""))
        secondary = normalise_ticker(row.get("live_fund_secondary", ""))

        index_levels = monthly_levels[asset_class].dropna().sort_index()
        index_last_date = pd.Timestamp(index_levels.index.max())
        anchor_level = float(index_levels.loc[index_last_date])

        selected_ticker, selected_source, note, live_prices = pick_live_ticker_for_asset(row, prices_df)

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

        live_anchor_date, live_anchor_price = nearest_level_on_or_before(live_prices, index_last_date)
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
                    "stitch_anchor_date": index_last_date,
                    "note": f"{note}. Live series has no price on or before the final index date, so no stitch was applied",
                }
            )
            continue

        live_extension = live_prices[live_prices.index > index_last_date].copy()
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
                    "stitch_anchor_date": index_last_date,
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
                "stitch_anchor_date": index_last_date,
                "note": f"{note}. Index history retained through its final date, then extended with live adjusted-close history",
            }
        )

    return stitched, pd.DataFrame(diag_rows)


def build_chart_preferred_series(monthly_levels: dict[str, pd.Series], mapping: pd.DataFrame, prices_df: pd.DataFrame) -> tuple[dict[str, pd.Series], pd.DataFrame]:
    chart_series_map = {}
    diag_rows = []
    deduped_mapping = mapping.drop_duplicates(subset=["asset_class"], keep="last").copy()

    for asset_class in sorted(deduped_mapping["asset_class"].dropna().astype(str).unique()):
        row_match = deduped_mapping[deduped_mapping["asset_class"] == asset_class]
        row = row_match.iloc[0] if not row_match.empty else pd.Series(dtype=object)

        primary = normalise_ticker(row.get("live_fund_primary", ""))
        secondary = normalise_ticker(row.get("live_fund_secondary", ""))

        index_levels = monthly_levels.get(asset_class, pd.Series(dtype=float)).dropna().sort_index()
        selected_ticker, selected_source, note, live_prices = pick_live_ticker_for_asset(row, prices_df)

        if live_prices.empty and index_levels.empty:
            diag_rows.append(
                {
                    "asset_class": asset_class,
                    "live_fund_primary": primary,
                    "live_fund_secondary": secondary,
                    "selected_ticker": selected_ticker,
                    "selected_source": selected_source,
                    "series_type": "missing",
                    "index_last_date": pd.NaT,
                    "live_first_date": pd.NaT,
                    "live_last_date": pd.NaT,
                    "stitch_anchor_date": pd.NaT,
                    "note": "No index or live series available for chart series",
                }
            )
            continue

        if live_prices.empty and not index_levels.empty:
            chart_series_map[asset_class] = index_levels.copy()
            diag_rows.append(
                {
                    "asset_class": asset_class,
                    "live_fund_primary": primary,
                    "live_fund_secondary": secondary,
                    "selected_ticker": selected_ticker,
                    "selected_source": selected_source,
                    "series_type": "index_fallback",
                    "index_last_date": pd.Timestamp(index_levels.index.max()),
                    "live_first_date": pd.NaT,
                    "live_last_date": pd.NaT,
                    "stitch_anchor_date": pd.NaT,
                    "note": "No usable live history. Monthly index history used for chart series",
                }
            )
            continue

        live_series = live_prices.dropna().sort_index()
        live_first_date = pd.Timestamp(live_series.index.min())
        live_last_date = pd.Timestamp(live_series.index.max())

        if index_levels.empty:
            chart_series_map[asset_class] = live_series.copy()
            diag_rows.append(
                {
                    "asset_class": asset_class,
                    "live_fund_primary": primary,
                    "live_fund_secondary": secondary,
                    "selected_ticker": selected_ticker,
                    "selected_source": selected_source,
                    "series_type": "live_only",
                    "index_last_date": pd.NaT,
                    "live_first_date": live_first_date,
                    "live_last_date": live_last_date,
                    "stitch_anchor_date": live_first_date,
                    "note": f"{note}. No index history available, so daily live history is used on its own",
                }
            )
            continue

        live_anchor_date = live_first_date
        live_anchor_value = float(live_series.iloc[0])

        index_anchor_date, index_anchor_level = nearest_level_on_or_before(index_levels, live_anchor_date)
        if index_anchor_date is None or index_anchor_level is None or index_anchor_level <= 0:
            chart_series_map[asset_class] = live_series.copy()
            diag_rows.append(
                {
                    "asset_class": asset_class,
                    "live_fund_primary": primary,
                    "live_fund_secondary": secondary,
                    "selected_ticker": selected_ticker,
                    "selected_source": selected_source,
                    "series_type": "live_preferred",
                    "index_last_date": pd.Timestamp(index_levels.index.max()),
                    "live_first_date": live_first_date,
                    "live_last_date": live_last_date,
                    "stitch_anchor_date": live_anchor_date,
                    "note": f"{note}. Daily live history used for chart series; no valid index anchor before live start so no backward stitch applied",
                }
            )
            continue

        historical_index = index_levels[index_levels.index < live_anchor_date].copy()
        scaled_historical_index = historical_index * (live_anchor_value / index_anchor_level)

        combined = pd.concat([scaled_historical_index, live_series])
        combined = combined[~combined.index.duplicated(keep="last")].sort_index()
        combined.name = asset_class
        chart_series_map[asset_class] = combined

        diag_rows.append(
            {
                "asset_class": asset_class,
                "live_fund_primary": primary,
                "live_fund_secondary": secondary,
                "selected_ticker": selected_ticker,
                "selected_source": selected_source,
                "series_type": "chart_stitched",
                "index_last_date": pd.Timestamp(index_levels.index.max()),
                "live_first_date": live_first_date,
                "live_last_date": live_last_date,
                "stitch_anchor_date": live_anchor_date,
                "note": (
                    f"{note}. Monthly index history stitched in before the first live daily observation; "
                    "daily live history is prioritised from that point onward"
                ),
            }
        )

    return chart_series_map, pd.DataFrame(diag_rows)


def calc_horizon_returns_from_levels(stitched_series_map: dict[str, pd.Series], end_date: pd.Timestamp, period_keys: list[str]) -> pd.DataFrame:
    rows = []
    for asset_class, series in stitched_series_map.items():
        row = {"asset_class": asset_class}
        for period_key in period_keys:
            row[period_key] = calc_period_return(series, end_date, period_key)
        rows.append(row)
    return pd.DataFrame(rows)


def calc_whole_period_returns(stitched_series_map: dict[str, pd.Series], end_date: pd.Timestamp, whole_period_start: pd.Timestamp) -> pd.DataFrame:
    rows = []
    for asset_class, series in stitched_series_map.items():
        rows.append(
            {
                "asset_class": asset_class,
                "Period": calc_period_return(series, end_date, "Period", whole_period_start=whole_period_start),
            }
        )
    return pd.DataFrame(rows)


def merge_return_tables(*dfs: pd.DataFrame) -> pd.DataFrame:
    out = None
    for df in dfs:
        out = df if out is None else out.merge(df, on="asset_class", how="outer")
    return out if out is not None else pd.DataFrame(columns=["asset_class"])


def order_asset_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "asset_class" not in df.columns:
        return df
    order_map = {name: idx for idx, name in enumerate(DEFAULT_ASSET_ORDER)}
    return (
        df.assign(_order=df["asset_class"].map(order_map).fillna(9999))
        .sort_values(["_order", "asset_class"])
        .drop(columns="_order")
    )


def build_real_level_series(nominal_series: pd.Series, inflation_levels: pd.Series) -> pd.Series:
    nominal = nominal_series.dropna().sort_index()
    infl = inflation_levels.dropna().sort_index()
    if nominal.empty or infl.empty:
        return pd.Series(dtype=float)
    aligned_infl = infl.reindex(nominal.index, method="ffill")
    real = nominal / aligned_infl
    return real.dropna().sort_index()


def build_growth_of_wealth_df(
    chart_series_map: dict[str, pd.Series],
    selected_assets: list[str],
    end_date: pd.Timestamp,
    period_key: str,
    is_real_mode: bool,
    inflation_levels: pd.Series | None,
) -> pd.DataFrame:
    rows = []

    for asset in selected_assets:
        if asset not in chart_series_map:
            continue

        series = chart_series_map[asset].dropna().sort_index()
        series = series[series.index <= end_date]
        if series.empty:
            continue

        if is_real_mode:
            if inflation_levels is None or inflation_levels.dropna().empty:
                continue
            series = build_real_level_series(series, inflation_levels)
            series = series[series.index <= end_date]
            if series.empty:
                continue

        if period_key == "YTD":
            base_date, base_level = nearest_level_on_or_before(series, pd.Timestamp(end_date.year - 1, 12, 31))
        else:
            years = int(period_key.replace("Y", ""))
            base_date, base_level = nearest_level_on_or_before(series, end_date - pd.DateOffset(years=years))

        if base_date is None or base_level is None or base_level <= 0:
            continue

        chart_series = series[series.index >= base_date].copy() / base_level
        if chart_series.empty:
            continue

        chart_df = chart_series.reset_index()
        chart_df.columns = ["Date", "Growth"]
        chart_df["asset_class"] = asset
        rows.append(chart_df)

    return pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(columns=["Date", "Growth", "asset_class"])


def build_monthly_level_window(series: pd.Series, end_date: pd.Timestamp, years: int) -> pd.Series:
    s = series.dropna().sort_index()
    s = s[s.index <= end_date]
    if s.empty:
        return pd.Series(dtype=float)

    monthly = s.resample("ME").last().dropna()
    if monthly.empty:
        return pd.Series(dtype=float)

    base_date, _ = nearest_level_on_or_before(monthly, end_date - pd.DateOffset(years=years))
    if base_date is None:
        return pd.Series(dtype=float)

    return monthly[monthly.index >= base_date].copy()


def build_monthly_level_window_from_start(series: pd.Series, end_date: pd.Timestamp, start_date: pd.Timestamp) -> pd.Series:
    s = series.dropna().sort_index()
    s = s[s.index <= end_date]
    if s.empty:
        return pd.Series(dtype=float)

    monthly = s.resample("ME").last().dropna()
    if monthly.empty:
        return pd.Series(dtype=float)

    base_date, _ = nearest_level_on_or_before(monthly, start_date)
    if base_date is None:
        return pd.Series(dtype=float)

    return monthly[monthly.index >= base_date].copy()


def calc_annualised_volatility(series: pd.Series, end_date: pd.Timestamp, years: int) -> float:
    window = build_monthly_level_window(series, end_date, years)
    monthly_returns = window.pct_change().dropna()
    if monthly_returns.empty or len(monthly_returns) < 2:
        return np.nan
    return float(monthly_returns.std(ddof=1) * np.sqrt(12))


def build_monthly_return_window(series: pd.Series, end_date: pd.Timestamp, years: int) -> pd.Series:
    window = build_monthly_level_window(series, end_date, years)
    if window.empty:
        return pd.Series(dtype=float)
    return window.pct_change().dropna()


def build_monthly_return_window_from_start(series: pd.Series, end_date: pd.Timestamp, start_date: pd.Timestamp) -> pd.Series:
    window = build_monthly_level_window_from_start(series, end_date, start_date)
    if window.empty:
        return pd.Series(dtype=float)
    return window.pct_change().dropna()


def calc_tracking_error(series: pd.Series, benchmark_series: pd.Series, end_date: pd.Timestamp, years: int) -> float:
    asset_returns = build_monthly_return_window(series, end_date, years)
    benchmark_returns = build_monthly_return_window(benchmark_series, end_date, years)
    aligned = pd.concat([asset_returns.rename("asset"), benchmark_returns.rename("benchmark")], axis=1).dropna()
    if aligned.empty or len(aligned) < 2:
        return np.nan
    return float((aligned["asset"] - aligned["benchmark"]).std(ddof=1) * np.sqrt(12))


def calc_tracking_error_from_start(series: pd.Series, benchmark_series: pd.Series, end_date: pd.Timestamp, start_date: pd.Timestamp) -> float:
    asset_returns = build_monthly_return_window_from_start(series, end_date, start_date)
    benchmark_returns = build_monthly_return_window_from_start(benchmark_series, end_date, start_date)
    aligned = pd.concat([asset_returns.rename("asset"), benchmark_returns.rename("benchmark")], axis=1).dropna()
    if aligned.empty or len(aligned) < 2:
        return np.nan
    excess_returns = aligned["asset"] - aligned["benchmark"]
    return float(excess_returns.std(ddof=1) * np.sqrt(12))


def calc_downside_deviation(series: pd.Series, end_date: pd.Timestamp, years: int) -> float:
    monthly_returns = build_monthly_return_window(series, end_date, years)
    if monthly_returns.empty:
        return np.nan
    downside = monthly_returns.clip(upper=0)
    return float(np.sqrt((downside.pow(2)).mean()) * np.sqrt(12))


def calc_downside_deviation_from_start(series: pd.Series, end_date: pd.Timestamp, start_date: pd.Timestamp) -> float:
    monthly_returns = build_monthly_return_window_from_start(series, end_date, start_date)
    if monthly_returns.empty:
        return np.nan
    downside = monthly_returns.clip(upper=0)
    return float(np.sqrt((downside.pow(2)).mean()) * np.sqrt(12))


def calc_annualised_volatility_from_start(series: pd.Series, end_date: pd.Timestamp, start_date: pd.Timestamp) -> float:
    monthly_returns = build_monthly_return_window_from_start(series, end_date, start_date)
    if monthly_returns.empty or len(monthly_returns) < 2:
        return np.nan
    return float(monthly_returns.std(ddof=1) * np.sqrt(12))


def calc_worst_drawdown_since_inception(series: pd.Series) -> float:
    s = series.dropna().sort_index()
    if s.empty:
        return np.nan
    drawdown = (s / s.cummax()) - 1
    return float(drawdown.min())


def build_risk_summary_table(
    series_map: dict[str, pd.Series],
    asset_style_map: dict[str, float],
    selected_assets: list[str],
    end_date: pd.Timestamp,
    start_date: pd.Timestamp,
) -> pd.DataFrame:
    rows = []
    cash_series = series_map.get("Cash (GBP)", pd.Series(dtype=float))
    global_series = series_map.get("Global stocks", pd.Series(dtype=float))
    developed_series = series_map.get("Developed stocks", pd.Series(dtype=float))
    assets = selected_assets if selected_assets else list(series_map.keys())

    for asset in assets:
        if asset not in series_map:
            continue

        series = series_map[asset]
        period_return = calc_period_return(series, end_date, "Period", whole_period_start=start_date)
        period_vol = calc_annualised_volatility_from_start(series, end_date, start_date)
        downside_dev = calc_downside_deviation_from_start(series, end_date, start_date)
        worst_drawdown = calc_worst_drawdown_since_inception(series)

        growth_flag = asset_style_map.get(asset, np.nan)
        benchmark_asset = "Global stocks" if growth_flag == 1 else "Cash (GBP)"
        benchmark_series = global_series if benchmark_asset == "Global stocks" else cash_series
        benchmark_return = (
            calc_period_return(benchmark_series, end_date, "Period", whole_period_start=start_date)
            if not benchmark_series.empty
            else np.nan
        )
        tracking_error = (
            calc_tracking_error_from_start(series, benchmark_series, end_date, start_date)
            if not benchmark_series.empty and asset != benchmark_asset
            else np.nan
        )

        return_vol_ratio = (
            period_return / period_vol
            if pd.notna(period_return) and pd.notna(period_vol) and period_vol > 0
            else np.nan
        )
        sharpe_ratio = (
            (period_return - calc_period_return(cash_series, end_date, period_key)) / period_vol
            if not cash_series.empty and pd.notna(period_vol) and period_vol > 0
            else np.nan
        )
        information_ratio = (
            (period_return - benchmark_return) / tracking_error
            if pd.notna(benchmark_return) and pd.notna(tracking_error) and tracking_error > 0
            else np.nan
        )
        sortino_ratio = (
            period_return / downside_dev
            if pd.notna(period_return) and pd.notna(downside_dev) and downside_dev > 0
            else np.nan
        )
        calmar_ratio = (
            period_return / abs(worst_drawdown)
            if pd.notna(period_return) and pd.notna(worst_drawdown) and worst_drawdown < 0
            else np.nan
        )

        if asset == "Cash (GBP)":
            return_vol_ratio = np.nan
            sharpe_ratio = np.nan
            information_ratio = np.nan
            sortino_ratio = np.nan
            worst_drawdown = np.nan
            calmar_ratio = np.nan
            tracking_error = np.nan

        rows.append(
            {
                "asset_class": asset,
                "Period return": period_return,
                "Period vol": period_vol,
                "Return/vol ratio": return_vol_ratio,
                "Sharpe ratio": sharpe_ratio,
                "Information ratio": information_ratio,
                "Sortino ratio": sortino_ratio,
                "Worst drawdown": worst_drawdown,
                "Calmar ratio": calmar_ratio,
                "Tracking error": tracking_error,
            }
        )

    out = order_asset_rows(pd.DataFrame(rows))
    if not out.empty and "asset_class" in out.columns:
        cash_rows = out[out["asset_class"] == "Cash (GBP)"]
        non_cash_rows = out[out["asset_class"] != "Cash (GBP)"]
        out = pd.concat([non_cash_rows, cash_rows], ignore_index=True)
    return out


def build_correlation_matrix_table(
    series_map: dict[str, pd.Series],
    selected_assets: list[str],
    end_date: pd.Timestamp,
    start_date: pd.Timestamp,
) -> pd.DataFrame:
    assets = selected_assets if selected_assets else list(series_map.keys())
    return_frames = []

    for asset in assets:
        if asset not in series_map:
            continue
        monthly_returns = build_monthly_return_window_from_start(series_map[asset], end_date, start_date)
        if monthly_returns.empty:
            continue
        return_frames.append(monthly_returns.rename(asset))

    if not return_frames:
        return pd.DataFrame()

    returns_df = pd.concat(return_frames, axis=1).dropna(how="all")
    if returns_df.empty:
        return pd.DataFrame()

    corr = returns_df.corr()
    ordered_assets = [c for c in assets if c in corr.index and c != "Cash (GBP)"]
    if "Cash (GBP)" in corr.index:
        ordered_assets.append("Cash (GBP)")
    corr = corr.reindex(index=ordered_assets, columns=ordered_assets)

    for row_idx, row_name in enumerate(corr.index):
        for col_idx, col_name in enumerate(corr.columns):
            if col_idx > row_idx:
                corr.loc[row_name, col_name] = np.nan

    corr = corr.reset_index().rename(columns={"index": "asset_class"})
    return corr


def build_risk_metrics_table(
    series_map: dict[str, pd.Series],
    selected_assets: list[str],
    end_date: pd.Timestamp,
    period_keys: list[str],
) -> pd.DataFrame:
    rows = []
    assets = selected_assets if selected_assets else list(series_map.keys())

    for asset in assets:
        if asset not in series_map:
            continue

        row = {"asset_class": asset}
        series = series_map[asset]
        for period_key in period_keys:
            years = int(period_key.replace("Y", ""))
            row[f"{period_key} Return"] = calc_period_return(series, end_date, period_key)
            row[f"{period_key} Vol"] = calc_annualised_volatility(series, end_date, years)
        rows.append(row)

    return order_asset_rows(pd.DataFrame(rows))


def build_risk_scatter_df(
    series_map: dict[str, pd.Series],
    selected_assets: list[str],
    end_date: pd.Timestamp,
    period_key: str,
) -> pd.DataFrame:
    rows = []
    years = int(period_key.replace("Y", ""))
    assets = selected_assets if selected_assets else list(series_map.keys())

    for asset in assets:
        if asset not in series_map:
            continue

        series = series_map[asset]
        annual_return = calc_period_return(series, end_date, period_key)
        annual_vol = calc_annualised_volatility(series, end_date, years)
        if pd.isna(annual_return) or pd.isna(annual_vol):
            continue

        rows.append(
            {
                "asset_class": asset,
                "display_asset_class": display_name(asset),
                "annual_return": annual_return,
                "annual_volatility": annual_vol,
            }
        )

    return pd.DataFrame(rows)


def build_calendar_year_returns(stitched_series_map: dict[str, pd.Series], end_date: pd.Timestamp, years_back: int = 10) -> pd.DataFrame:
    last_complete_year = end_date.year - 1
    years = list(reversed(list(range(last_complete_year - years_back + 1, last_complete_year + 1))))
    rows = []

    for asset_class, series in stitched_series_map.items():
        row = {"asset_class": asset_class}
        s = series.dropna().sort_index()

        for year in years:
            _, end_level = nearest_level_on_or_before(s, pd.Timestamp(year, 12, 31))
            _, start_level = nearest_level_on_or_before(s, pd.Timestamp(year - 1, 12, 31))
            if end_level is None or start_level is None or start_level <= 0:
                row[str(year)] = np.nan
            else:
                row[str(year)] = (end_level / start_level) - 1

        rows.append(row)

    return order_asset_rows(pd.DataFrame(rows))


def format_pct_strings(df: pd.DataFrame, exclude_cols: set[str] | None = None) -> pd.DataFrame:
    exclude_cols = exclude_cols or {"asset_class"}
    out = df.copy()
    for col in out.columns:
        if col not in exclude_cols:
            out[col] = out[col].map(format_pct)
    return out


def build_html_table(
    df: pd.DataFrame,
    percent_cols: list[str] | None = None,
    conditional_cols: list[str] | None = None,
    header_wrap_cols: list[str] | None = None,
    invert_conditional_cols: list[str] | None = None,
    rank_conditional_cols: list[str] | None = None,
    decimal_cols: list[str] | None = None,
    correlation_conditional_cols: list[str] | None = None,
) -> str:
    if df.empty:
        return '<div class="table-shell"><div class="table-empty">No data available.</div></div>'

    display_df = df.copy()
    if "asset_class" in display_df.columns:
        display_df["asset_class"] = display_df["asset_class"].map(display_name)
        display_df = display_df.rename(columns={"asset_class": "Asset class"})

    cols = list(display_df.columns)
    percent_cols = set(percent_cols or [])
    conditional_cols = set(conditional_cols or [])
    header_wrap_cols = set(header_wrap_cols or [])
    invert_conditional_cols = set(invert_conditional_cols or [])
    rank_conditional_cols = set(rank_conditional_cols or [])
    decimal_cols = set(decimal_cols or [])
    correlation_conditional_cols = set(correlation_conditional_cols or [])
    if cols:
        asset_col_width = 28 if "Asset class" in cols else 0
        other_count = len(cols) - (1 if "Asset class" in cols else 0)
        other_width = ((100 - asset_col_width) / other_count) if other_count > 0 else 100
    else:
        asset_col_width = 0
        other_width = 100

    colgroup = "".join(
        [
            f'<col style="width:{asset_col_width if col == "Asset class" else other_width:.4f}%;">'
            for col in cols
        ]
    )

    thead = "".join(
        [
            f"<th>{col.replace(' ', '<br>') if col in header_wrap_cols else col}</th>"
            for col in cols
        ]
    )

    heat_bounds = {}
    for col in cols:
        source_col = "asset_class" if col == "Asset class" else col
        if source_col in conditional_cols and source_col in df.columns and pd.api.types.is_numeric_dtype(df[source_col]):
            numeric = pd.to_numeric(df[source_col], errors="coerce").dropna()
            heat_bounds[source_col] = (
                float(numeric.min()) if not numeric.empty else 0.0,
                float(numeric.max()) if not numeric.empty else 0.0,
            )

    body_rows = []
    for row_idx, (_, row) in enumerate(display_df.iterrows()):
        cells = []
        for col in cols:
            value = row[col]
            source_col = "asset_class" if col == "Asset class" else col
            cell_text = "-"
            cell_style = ""

            if pd.notna(value):
                if source_col in percent_cols:
                    cell_text = format_pct(float(value))
                elif source_col in decimal_cols:
                    cell_text = f"{float(value):.2f}"
                else:
                    cell_text = str(value)

            if source_col in heat_bounds and pd.notna(value):
                vmin, vmax = heat_bounds[source_col]
                if source_col in correlation_conditional_cols:
                    bg_colour = correlation_heat_colour(float(value))
                elif source_col in rank_conditional_cols:
                    bg_colour = rank_heat_colour(
                        float(value),
                        vmin,
                        vmax,
                        low_is_good=source_col in invert_conditional_cols,
                    )
                else:
                    heat_value = -float(value) if source_col in invert_conditional_cols else float(value)
                    heat_vmin = -vmax if source_col in invert_conditional_cols else vmin
                    heat_vmax = -vmin if source_col in invert_conditional_cols else vmax
                    bg_colour = heat_colour(heat_value, heat_vmin, heat_vmax)
                cell_style = (
                    f' style="background:{bg_colour};'
                    ' color:#ffffff; font-weight:600;"'
                )

            cells.append(f"<td{cell_style}>{cell_text}</td>")
        body_rows.append(f"<tr>{''.join(cells)}</tr>")

    tbody = "".join(body_rows)

    return f"""
    <div class="table-shell">
        <table class="custom-data-table">
            <colgroup>{colgroup}</colgroup>
            <thead>
                <tr>{thead}</tr>
            </thead>
            <tbody>
                {tbody}
            </tbody>
        </table>
    </div>
    """


def convert_pct_table_for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        if col != "asset_class" and pd.api.types.is_numeric_dtype(out[col]):
            out[col] = out[col].map(lambda x: np.nan if pd.isna(x) else round(x * 100, 2))
    return prepare_dataframe_for_display(out)


def get_chart_axis_format(period_key: str) -> str:
    if period_key in {"YTD", "1Y"}:
        return "%d/%m/%y"
    if period_key in {"3Y", "5Y"}:
        return "%b-%y"
    return "%Y"


def get_chart_right_padding_days(period_key: str) -> int:
    if period_key == "YTD":
        return 10
    if period_key == "1Y":
        return 18
    if period_key in {"3Y", "5Y"}:
        return 35
    return 120


def build_chart(chart_df: pd.DataFrame, selected_assets: list[str], period_key: str) -> alt.Chart:
    if chart_df.empty:
        return alt.Chart(pd.DataFrame({"x": [], "y": []})).mark_line()

    chart_df = chart_df.copy()
    chart_df["asset_class"] = chart_df["asset_class"].map(display_name)

    colour_domain = [display_name(a) for a in selected_assets if display_name(a) in chart_df["asset_class"].unique()]
    colour_range = [ASSET_COLOURS.get(a, "#1f77b4") for a in selected_assets if display_name(a) in chart_df["asset_class"].unique()]

    ymin = float(chart_df["Growth"].min())
    ymax = float(chart_df["Growth"].max())
    spread = max(ymax - ymin, 0.02)
    pad = spread * 0.10
    domain_min = ymin - pad
    domain_max = ymax + pad
    x_axis_format = get_chart_axis_format(period_key)

    return (
        alt.Chart(chart_df)
        .mark_line(strokeWidth=2.5)
        .encode(
            x=alt.X(
                "Date:T",
                title=None,
                axis=alt.Axis(
                    format=x_axis_format,
                    labelColor=TEXT_GREY,
                    labelFontSize=15,
                    tickColor=MID_GREY,
                    domainColor=MID_GREY,
                    grid=False,
                ),
            ),
            y=alt.Y(
                "Growth:Q",
                title=None,
                scale=alt.Scale(domain=[domain_min, domain_max], zero=False),
                axis=alt.Axis(
                    labelColor=TEXT_GREY,
                    labelFontSize=15,
                    tickColor=MID_GREY,
                    domainColor=MID_GREY,
                    grid=True,
                    gridColor=MID_GREY,
                    gridDash=[2, 4],
                    labelExpr="'£' + format(datum.value, ',.2f')",
                ),
            ),
            color=alt.Color(
                "asset_class:N",
                title="Asset class",
                scale=alt.Scale(domain=colour_domain, range=colour_range),
                legend=alt.Legend(
                    labelColor=TEXT_GREY,
                    titleColor=TEXT_GREY,
                    orient="bottom",
                    direction="horizontal",
                    columns=max(1, len(colour_domain)),
                    title=None,
                    symbolLimit=len(colour_domain),
                    labelLimit=180,
                ),
            ),
            tooltip=[
                alt.Tooltip("Date:T", title="Date"),
                alt.Tooltip("asset_class:N", title="Asset class"),
                alt.Tooltip("Growth:Q", title="Growth of wealth", format=".3f"),
            ],
        )
        .properties(
            height=420,
            width="container",
            padding={"left": 14, "top": 8, "right": max(28, int(get_chart_right_padding_days(period_key) * 0.8)), "bottom": 10},
        )
        .configure_view(stroke=None, fill=CHART_BG_GREY)
        .configure_axis(labelFont="Calibri", titleFont="Calibri")
        .configure_legend(
            labelFont="Calibri",
            titleFont="Calibri",
            fillColor=CHART_BG_GREY,
            strokeColor=CHART_BG_GREY,
        )
        .configure(background=CHART_BG_GREY)
    )


def build_risk_scatter_chart(risk_df: pd.DataFrame, selected_assets: list[str]) -> alt.Chart:
    if risk_df.empty:
        return alt.Chart(pd.DataFrame({"x": [], "y": []})).mark_circle()

    colour_domain = [
        display_name(a)
        for a in selected_assets
        if display_name(a) in risk_df["display_asset_class"].unique()
    ]
    colour_range = [
        ASSET_COLOURS.get(a, "#1f77b4")
        for a in selected_assets
        if display_name(a) in risk_df["display_asset_class"].unique()
    ]

    return (
        alt.Chart(risk_df)
        .mark_circle(size=180, opacity=0.95)
        .encode(
            x=alt.X(
                "annual_volatility:Q",
                title="Annualised volatility",
                axis=alt.Axis(
                    format=".0%",
                    labelColor=TEXT_GREY,
                    titleColor=TEXT_GREY,
                    labelFontSize=15,
                    tickColor=MID_GREY,
                    domainColor=MID_GREY,
                    grid=True,
                    gridColor=MID_GREY,
                    gridDash=[2, 4],
                ),
            ),
            y=alt.Y(
                "annual_return:Q",
                title="Annualised return",
                axis=alt.Axis(
                    format=".0%",
                    labelColor=TEXT_GREY,
                    titleColor=TEXT_GREY,
                    labelFontSize=15,
                    tickColor=MID_GREY,
                    domainColor=MID_GREY,
                    grid=True,
                    gridColor=MID_GREY,
                    gridDash=[2, 4],
                ),
            ),
            color=alt.Color(
                "display_asset_class:N",
                title=None,
                scale=alt.Scale(domain=colour_domain, range=colour_range),
                legend=alt.Legend(
                    labelColor=TEXT_GREY,
                    orient="bottom",
                    direction="horizontal",
                    columns=max(1, len(colour_domain)),
                    symbolLimit=len(colour_domain),
                    labelLimit=180,
                ),
            ),
            tooltip=[
                alt.Tooltip("display_asset_class:N", title="Asset class"),
                alt.Tooltip("annual_return:Q", title="Annualised return", format=".2%"),
                alt.Tooltip("annual_volatility:Q", title="Annualised volatility", format=".2%"),
            ],
        )
        .properties(height=420, width="container", padding={"left": 14, "top": 8, "right": 28, "bottom": 10})
        .configure_view(stroke=None, fill=CHART_BG_GREY)
        .configure_axis(labelFont="Calibri", titleFont="Calibri")
        .configure_legend(
            labelFont="Calibri",
            titleFont="Calibri",
            fillColor=CHART_BG_GREY,
            strokeColor=CHART_BG_GREY,
        )
        .configure(background=CHART_BG_GREY)
    )


def build_yield_curve_chart(yield_curve_df: pd.DataFrame) -> alt.Chart:
    if yield_curve_df.empty:
        return alt.Chart(pd.DataFrame({"x": [], "y": []})).mark_line()

    chart_df = yield_curve_df.copy()
    chart_df["curve_type"] = pd.Categorical(chart_df["curve_type"], categories=["Nominal", "Real", "Breakeven inflation"], ordered=True)
    chart_df["yield_decimal"] = pd.to_numeric(chart_df["yield_percent"], errors="coerce") / 100.0

    return (
        alt.Chart(chart_df)
        .mark_line(point=True, strokeWidth=2.8)
        .encode(
            x=alt.X(
                "maturity_years:Q",
                title="Maturity (years)",
                axis=alt.Axis(
                    labelColor=TEXT_GREY,
                    titleColor=TEXT_GREY,
                    labelFontSize=15,
                    tickColor=MID_GREY,
                    domainColor=MID_GREY,
                    grid=False,
                ),
            ),
            y=alt.Y(
                "yield_decimal:Q",
                title="Yield",
                axis=alt.Axis(
                    labelColor=TEXT_GREY,
                    titleColor=TEXT_GREY,
                    labelFontSize=15,
                    tickColor=MID_GREY,
                    domainColor=MID_GREY,
                    grid=True,
                    gridColor=MID_GREY,
                    gridDash=[2, 4],
                    format=".0%",
                ),
            ),
            color=alt.Color(
                "curve_type:N",
                title=None,
                scale=alt.Scale(domain=["Nominal", "Real", "Breakeven inflation"], range=["#c95b2b", "#1f77b4", "#2e8b57"]),
                legend=alt.Legend(
                    labelColor=TEXT_GREY,
                    orient="bottom",
                    direction="horizontal",
                    columns=3,
                ),
            ),
            tooltip=[
                alt.Tooltip("curve_type:N", title="Curve"),
                alt.Tooltip("curve_date:T", title="Date"),
                alt.Tooltip("maturity_years:Q", title="Maturity (years)", format=".1f"),
                alt.Tooltip("yield_decimal:Q", title="Yield", format=".2%"),
            ],
        )
        .properties(height=420, width="container", padding={"left": 14, "top": 8, "right": 28, "bottom": 10})
        .configure_view(stroke=None, fill=CHART_BG_GREY)
        .configure_axis(labelFont="Calibri", titleFont="Calibri")
        .configure_legend(
            labelFont="Calibri",
            titleFont="Calibri",
            fillColor=CHART_BG_GREY,
            strokeColor=CHART_BG_GREY,
        )
        .configure(background=CHART_BG_GREY)
    )


def build_yield_curve_display_df(yield_curve_df: pd.DataFrame, selected_series: list[str]) -> pd.DataFrame:
    if yield_curve_df.empty:
        return pd.DataFrame(columns=["maturity_years", "yield_percent", "curve_type", "curve_date"])

    base = yield_curve_df.copy()
    if selected_series:
        base = base[base["curve_type"].isin(selected_series)].copy()

    nominal = yield_curve_df[yield_curve_df["curve_type"] == "Nominal"][["maturity_years", "yield_percent", "curve_date"]].copy()
    real = yield_curve_df[yield_curve_df["curve_type"] == "Real"][["maturity_years", "yield_percent", "curve_date"]].copy()
    nominal = nominal.rename(columns={"yield_percent": "nominal_yield", "curve_date": "nominal_date"})
    real = real.rename(columns={"yield_percent": "real_yield", "curve_date": "real_date"})
    if not nominal.empty and not real.empty:
        nominal_sorted = nominal.sort_values("maturity_years").dropna(subset=["maturity_years", "nominal_yield"])
        real_sorted = real.sort_values("maturity_years").dropna(subset=["maturity_years", "real_yield"])
        interpolated_nominal = np.interp(
            real_sorted["maturity_years"].to_numpy(),
            nominal_sorted["maturity_years"].to_numpy(),
            nominal_sorted["nominal_yield"].to_numpy(),
        )
        bei = real_sorted.copy()
        bei["nominal_yield"] = interpolated_nominal
        bei["nominal_date"] = nominal_sorted["nominal_date"].max()
        bei["curve_type"] = "Breakeven inflation"
        bei["yield_percent"] = (((1 + (bei["nominal_yield"] / 100.0)) / (1 + (bei["real_yield"] / 100.0))) - 1) * 100.0
        bei["curve_date"] = bei[["nominal_date", "real_date"]].max(axis=1)
        bei = bei[["maturity_years", "yield_percent", "curve_type", "curve_date"]]
    else:
        bei = pd.DataFrame(columns=["maturity_years", "yield_percent", "curve_type", "curve_date"])

    combined = pd.concat([base, bei], ignore_index=True)
    if selected_series:
        combined = combined[combined["curve_type"].isin(selected_series)].copy()
    return combined.sort_values(["curve_type", "maturity_years"]).reset_index(drop=True)


# =====================================
# PAGE SETUP
# =====================================
st.set_page_config(page_title=APP_TITLE, layout="wide")

st.markdown(
    f"""
    <style>
    html, body, .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stMarkdownContainer"],
    p, li, label, button, input, textarea, table {{
        font-family: "Calibri Light", Calibri, "Segoe UI", Arial, sans-serif !important;
    }}

    .stApp {{ background-color: white; }}

    header[data-testid="stHeader"],
    div[data-testid="stToolbar"],
    div[data-testid="stDecoration"],
    div[data-testid="stStatusWidget"] {{
        display: none !important;
        visibility: hidden !important;
        height: 0 !important;
    }}

    #MainMenu {{ visibility: hidden !important; }}

    [data-testid="stAppViewContainer"] > .main {{ padding-top: 0 !important; }}

    .block-container {{
        padding-top: 0.25rem !important;
        padding-bottom: 1rem !important;
        max-width: 1500px;
    }}

    .top-header-grid {{
        display: grid;
        grid-template-columns: 1fr auto;
        align-items: center;
        gap: 18px;
        margin-bottom: 2px;
    }}

    .top-title-wrap {{
        display: flex;
        align-items: center;
    }}

    .dashboard-title {{
        font-size: 40px;
        font-weight: 500;
        margin-bottom: 2px;
        color: black;
        line-height: 1.15;
        padding-top: 0;
    }}

    .header-logo {{
        display: flex;
        justify-content: flex-end;
        align-items: center;
        padding-top: 0;
    }}

    .header-logo img {{
        max-height: 50px;
        width: auto;
        object-fit: contain;
        display: block;
    }}

    .toolbar-title {{
        font-size: 13px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #444;
        margin-bottom: 4px;
    }}

    .toolbar-label {{
        font-size: 13px;
        font-weight: 700;
        color: #444;
        margin-bottom: 6px;
    }}

    .toolbar-label-muted {{ color: #888 !important; }}

    .toolbar-meta {{
        text-align: right;
        font-size: 13px;
        color: {TEXT_GREY};
        padding-top: 22px;
        line-height: 1.2;
        white-space: nowrap;
    }}

    .period-shell {{
        background: transparent;
        padding: 8px 8px 12px 8px;
        min-height: 100%;
    }}

    .period-title {{
        text-align: center;
        font-size: 28px;
        font-weight: 500;
        margin-bottom: 10px;
        color: black;
    }}

    .group-card {{
        padding: 10px 8px 8px 8px;
        margin-bottom: 10px;
        background: #e7e7e7;
    }}

    .section-title {{
        text-align: center;
        font-size: 18px;
        font-weight: 500;
        margin-top: 0;
        margin-bottom: 4px;
        color: black;
    }}

    .section-title-empty {{
        height: 4px;
        margin-bottom: 2px;
    }}

    .section-subtitle {{
        text-align: center;
        font-size: 13px;
        font-style: italic;
        color: #444;
        margin-bottom: 6px;
    }}

    .big-tile, .small-tile {{
        color: white;
        text-align: center;
        font-weight: 700;
        border-radius: 0;
        line-height: 1.2;
    }}

    .big-tile {{
        padding: 10px 8px;
        margin-bottom: 8px;
        font-size: 22px;
    }}

    .small-tile {{
        padding: 8px 6px;
        font-size: 18px;
        margin-bottom: 8px;
    }}

    .tile-label, .tile-label-on-colour {{
        display: block;
        color: black !important;
        font-size: 13px;
        font-style: italic;
        font-weight: 500;
        margin-bottom: 4px;
    }}

    .tile-label-plain {{
        font-style: normal !important;
        font-weight: 700 !important;
    }}

    .spacer {{ height: 2px; }}

    .page-section-title {{
        font-size: 20px;
        font-weight: 600;
        color: {TEXT_GREY};
        margin: 6px 0 10px 0;
    }}

    .methodology-text {{
        margin-top: 18px;
        margin-bottom: 10px;
        font-size: 13px;
        color: {TEXT_GREY} !important;
        text-align: center;
        line-height: 1.45;
    }}

    .methodology-text a {{
        color: {BRAND_ORANGE_DARK};
        text-decoration: none;
    }}

    .footer-bar {{
        margin-top: 20px;
        display: flex;
        justify-content: center;
        align-items: center;
        width: 100%;
    }}

    .footer-logo img {{
        max-height: 48px;
        width: auto;
        object-fit: contain;
        display: block;
    }}

    .diag-title {{
        font-size: 20px;
        font-weight: 500;
        color: black !important;
        margin-top: 4px;
        margin-bottom: 8px;
    }}

    .diag-note {{
        font-size: 13px;
        color: {TEXT_GREY} !important;
        margin-bottom: 14px;
        line-height: 1.4;
    }}

    .stButton button {{
        border-radius: 6px !important;
        font-weight: 600 !important;
    }}

    [data-testid="collapsedControl"] {{
        display: none !important;
    }}

    section[data-testid="stSidebar"] {{
        min-width: 58px !important;
        max-width: 58px !important;
        transition: min-width 0.18s ease, max-width 0.18s ease !important;
        overflow-x: hidden !important;
        background: #ffffff !important;
        border-right: 1px solid {MID_GREY} !important;
    }}

    section[data-testid="stSidebar"]:hover {{
        min-width: 228px !important;
        max-width: 228px !important;
    }}

    section[data-testid="stSidebar"] > div:first-child {{
        width: 58px !important;
        transition: width 0.18s ease !important;
        overflow-x: hidden !important;
    }}

    section[data-testid="stSidebar"]:hover > div:first-child {{
        width: 228px !important;
    }}

    .sidebar-hamburger {{
        font-size: 24px;
        font-weight: 700;
        color: #111827;
        margin: 2px 0 14px 2px;
        line-height: 1;
    }}

    .sidebar-nav-title {{
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #6b7280;
        margin: 10px 0 8px 2px;
        white-space: nowrap;
    }}

    section[data-testid="stSidebar"] .sidebar-nav-title,
    section[data-testid="stSidebar"] div[data-testid="stButton"] {{
        opacity: 0;
        visibility: hidden;
        max-height: 0;
        overflow: hidden;
        pointer-events: none;
        margin: 0 !important;
        transition: opacity 0.12s ease, max-height 0.12s ease, visibility 0.12s ease;
    }}

    section[data-testid="stSidebar"]:hover .sidebar-nav-title,
    section[data-testid="stSidebar"]:hover div[data-testid="stButton"] {{
        opacity: 1;
        visibility: visible;
        max-height: 80px;
        pointer-events: auto;
    }}

    .stSidebar .stButton button {{
        min-height: 40px !important;
        width: 100% !important;
        justify-content: flex-start !important;
        padding-left: 12px !important;
    }}

    .stSidebar .stButton button[kind="primary"] {{
        background: {BRAND_ORANGE} !important;
        border: 1px solid {BRAND_ORANGE} !important;
        color: #ffffff !important;
    }}

    .stSidebar .stButton button[kind="secondary"] {{
        background: #f7f7f7 !important;
        border: 1px solid #d8d8d8 !important;
        color: #111111 !important;
    }}

    .stDownloadButton button {{
        background: #f2f2f2 !important;
        color: black !important;
        border: 1px solid #cfcfcf !important;
        border-radius: 4px !important;
        font-weight: 500 !important;
    }}

    .stDownloadButton button * {{
        color: black !important;
    }}

    .stDownloadButton button:hover,
    .stButton button:hover {{
        border-color: {BRAND_ORANGE} !important;
    }}

    .stTabs [data-baseweb="tab-list"] {{
        gap: 6px;
    }}

    .stTabs [data-baseweb="tab"] {{
        color: {TEXT_GREY} !important;
        background: #f3f3f3 !important;
        border: 0.2px solid {MID_GREY} !important;
        border-radius: 8px 8px 0 0 !important;
        padding: 0.45rem 0.8rem !important;
    }}

    .stTabs [aria-selected="true"] {{
        color: #111111 !important;
        background: #ffffff !important;
        border-bottom-color: #ffffff !important;
        font-weight: 600 !important;
    }}

    .stTabs [data-baseweb="tab"] * {{
        color: inherit !important;
        -webkit-text-fill-color: inherit !important;
    }}

    .stMetric, .stMetric * {{
        color: black !important;
    }}

    h1, h2, h3, h4, h5, h6,
    [data-testid="stHeading"],
    [data-testid="stHeading"] * {{
        color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
    }}

    .stMultiSelect label, .stSelectbox label {{
        color: {TEXT_GREY} !important;
    }}

    .st-key-diagnostics_toggle_button button {{
        background: #111827 !important;
        color: #ffffff !important;
        border: 1px solid #111827 !important;
        border-radius: 4px !important;
        font-weight: 600 !important;
        text-align: left !important;
        justify-content: flex-start !important;
        padding: 0.55rem 0.9rem !important;
        margin: 0 !important;
    }}

    .st-key-diagnostics_toggle_button button *,
    .st-key-diagnostics_toggle_button button span {{
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }}

    div[data-testid="stVerticalBlock"] > div:empty {{
        display: none !important;
    }}

    div[data-testid="stButton"],
    div[data-testid="stButton"] > div {{
        margin: 0 !important;
        padding: 0 !important;
    }}

    div[data-testid="stButton"] button {{
        margin-bottom: 0 !important;
    }}

    div[data-testid="stVegaLiteChart"] {{
        background: {CHART_BG_GREY} !important;
        border: 0.2px solid {MID_GREY} !important;
        border-radius: 12px !important;
        padding: 10px 18px 8px 12px !important;
        overflow: hidden !important;
    }}

    div[data-testid="stVegaLiteChart"] canvas,
    div[data-testid="stVegaLiteChart"] svg {{
        border-radius: 10px !important;
    }}

    .table-shell {{
        width: 100%;
        border: 0.2px solid {MID_GREY};
        border-radius: 12px;
        overflow: hidden;
        margin-bottom: 8px;
    }}

    .custom-data-table {{
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
    }}

    .custom-data-table thead th {{
        background: {LIGHT_GREY};
        color: {TEXT_GREY};
        border: 0.2px solid {MID_GREY};
        padding: 10px 12px;
        text-align: center;
        font-weight: 600;
        white-space: normal;
        line-height: 1.2;
        overflow: hidden;
        text-overflow: ellipsis;
    }}

    .custom-data-table tbody td {{
        color: {TEXT_GREY};
        border: 0.2px solid {MID_GREY};
        padding: 10px 12px;
        text-align: center;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }}

    .custom-data-table tbody tr:nth-child(odd) td {{
        background: {LIGHT_GREY};
    }}

    .custom-data-table tbody tr:nth-child(even) td {{
        background: {WHITE};
    }}

    .custom-data-table th:first-child,
    .custom-data-table td:first-child {{
        text-align: left;
    }}

    .table-empty {{
        padding: 16px;
        color: {TEXT_GREY};
        text-align: center;
        background: {WHITE};
    }}
    </style>
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

try:
    ts, mapping = load_data(str(DATA_FILE), current_file_mtime)
except Exception as exc:
    st.exception(exc)
    st.stop()

monthly_levels = build_monthly_index_levels(ts, mapping)
if not monthly_levels:
    st.error(
        "No mapped asset classes were found. Check that the mapping sheet asset_class names "
        "match the expected names in DISPLAY_GROUPS / ASSET_CLASS_ALIASES."
    )
    st.stop()

inflation_levels = None
inflation_source_note = "Workbook time_series"
inflation_debug_message = None

try:
    inflation_levels, inflation_source_note, inflation_debug_message = build_best_available_inflation_levels(ts)
except Exception as exc:
    st.warning(f"Inflation series could not be built. Real mode may be unavailable. Details: {exc}")

inflation_monthly_returns = (
    build_monthly_returns_from_levels(inflation_levels)
    if inflation_levels is not None and not inflation_levels.dropna().empty
    else pd.Series(dtype=float)
)

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

chart_series_map, chart_diag = build_chart_preferred_series(
    monthly_levels=monthly_levels,
    mapping=mapping,
    prices_df=live_prices,
)

if not stitched_series_map:
    st.error("No stitched or index-only asset series could be built.")
    st.stop()

available_assets = [a for a in DEFAULT_ASSET_ORDER if a in stitched_series_map or a in chart_series_map]
default_chart_assets = [a for a in DEFAULT_CHART_ASSETS if a in available_assets]
whole_period_start = pd.Timestamp("1989-07-31")
common_inception_text = whole_period_start.strftime("%d/%m/%Y")

asset_coverage_diag = build_asset_coverage_table(
    mapping=mapping,
    monthly_levels=monthly_levels,
    stitched_series_map=stitched_series_map,
    chart_series_map=chart_series_map,
    live_diag=live_diag,
    chart_diag=chart_diag,
)
mapping_diag = build_mapping_diagnostics_table(mapping, ts)
live_price_diag = build_live_price_diagnostics(live_prices)
ons_fetch_summary, ons_fetch_preview = build_ons_fetch_diagnostics(ONS_CPI_INDEX_CSV_URL)
yield_curve_df, boe_yield_summary, boe_yield_preview = build_boe_yield_curve_diagnostics(
    BOE_YIELD_CURVE_ZIP_URL,
    DIVIDENDDATA_INDEX_LINKED_GILTS_URL,
)
asset_style_map = (
    mapping.drop_duplicates(subset=["asset_class"], keep="last")
    .set_index("asset_class")
    .get("growth_defensive", pd.Series(dtype=float))
    .to_dict()
)


# =====================================
# TOP HEADER
# =====================================
logo_html = ""
if ALBION_LOGO_FILE.exists():
    logo_html = (
        f'<div class="header-logo"><img src="data:image/png;base64,'
        f'{img_to_base64(str(ALBION_LOGO_FILE), ALBION_LOGO_FILE.stat().st_mtime)}"></div>'
    )

if "top_page_selector" not in st.session_state:
    st.session_state["top_page_selector"] = "Dashboard"

with st.sidebar:
    st.markdown('<div class="sidebar-hamburger">&#9776;</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-nav-title">Navigation</div>', unsafe_allow_html=True)
    if st.button(
        PAGE_LABELS["Dashboard"],
        key="sidebar_page_dashboard_btn",
        use_container_width=True,
        type="primary" if st.session_state["top_page_selector"] == "Dashboard" else "secondary",
    ):
        st.session_state["top_page_selector"] = "Dashboard"
        st.rerun()
    if st.button(
        PAGE_LABELS["Charts"],
        key="sidebar_page_charts_btn",
        use_container_width=True,
        type="primary" if st.session_state["top_page_selector"] == "Charts" else "secondary",
    ):
        st.session_state["top_page_selector"] = "Charts"
        st.rerun()
    if st.button(
        PAGE_LABELS["Risk"],
        key="sidebar_page_risk_btn",
        use_container_width=True,
        type="primary" if st.session_state["top_page_selector"] == "Risk" else "secondary",
    ):
        st.session_state["top_page_selector"] = "Risk"
        st.rerun()
    if st.button(
        PAGE_LABELS["Yield"],
        key="sidebar_page_yield_btn",
        use_container_width=True,
        type="primary" if st.session_state["top_page_selector"] == "Yield" else "secondary",
    ):
        st.session_state["top_page_selector"] = "Yield"
        st.rerun()

st.markdown('<div class="top-header-grid">', unsafe_allow_html=True)
left_col, right_col = st.columns([2.2, 1.0])

with left_col:
    st.markdown(f'<div class="top-title-wrap"><div class="dashboard-title">{APP_TITLE}</div></div>', unsafe_allow_html=True)

with right_col:
    st.markdown(logo_html, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

page_name = st.session_state["top_page_selector"]


# =====================================
# CONTENT
# =====================================
if page_name == "Dashboard":
    display_mode = st.session_state.get("display_mode_toolbar", "Absolute")
    return_basis = st.session_state.get("return_basis_toolbar", "Nominal")
    relative_detail_mode = st.session_state.get("relative_basis_toolbar", "Major")

    is_relative_mode = display_mode == "Relative"
    is_real_mode = return_basis == "Real"
    effective_real_mode = is_real_mode and inflation_levels is not None and not inflation_levels.dropna().empty

    end_date_dashboard = get_dashboard_end_date(
        stitched_series_map=stitched_series_map,
        live_diag=live_diag,
        inflation_levels=inflation_levels,
        is_real_mode=effective_real_mode,
    )

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
        ) or "Major"

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
                Annualised {"real" if effective_real_mode else "nominal"} returns in GBP to <b>{end_date_dashboard.strftime("%d/%m/%Y")}</b>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if is_real_mode and not effective_real_mode:
        st.warning("Real mode selected but no usable UK inflation series was found. Falling back to nominal results.")

    display_groups = get_display_groups(is_relative_mode, relative_detail_mode)

    absolute_returns_df = order_asset_rows(
        calc_horizon_returns_from_levels(stitched_series_map, end_date_dashboard, list(DASHBOARD_HORIZONS.keys()))
    )

    if is_relative_mode:
        nominal_display_returns_df = order_asset_rows(
            convert_to_relative_returns(absolute_returns_df, relative_detail_mode=relative_detail_mode)
        )
    else:
        nominal_display_returns_df = absolute_returns_df.copy()

    inflation_returns_dashboard_df = (
        calc_horizon_returns_from_levels({"UK inflation": inflation_levels}, end_date_dashboard, list(DASHBOARD_HORIZONS.keys()))
        if inflation_levels is not None and not inflation_levels.dropna().empty
        else pd.DataFrame()
    )

    displayed_returns_dashboard_df = (
        order_asset_rows(convert_to_real_returns(nominal_display_returns_df, inflation_returns_dashboard_df))
        if effective_real_mode
        else nominal_display_returns_df.copy()
    )

    lookup = build_lookup_table(displayed_returns_dashboard_df)

    cols = st.columns(4)
    period_order = ["20Y", "10Y", "5Y", "YTD"]

    for col, period in zip(cols, period_order):
        period_vals = displayed_returns_dashboard_df[period].dropna()
        vmin = period_vals.min() if len(period_vals) else -0.05
        vmax = period_vals.max() if len(period_vals) else 0.15

        with col:
            st.markdown('<div class="period-shell">', unsafe_allow_html=True)
            st.markdown(f'<div class="period-title">{DASHBOARD_HORIZONS[period]}</div>', unsafe_allow_html=True)

            for group in display_groups:
                title = group["title"]
                items = group["items"]
                labels = group["labels"]

                st.markdown('<div class="group-card">', unsafe_allow_html=True)

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
                        with target_cols[idx % 2]:
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
                    row1 = st.columns(3)
                    row2 = st.columns(2)

                    for idx, item in enumerate(items):
                        val = lookup.get(item, {}).get(period, np.nan)
                        colour = heat_colour(val, vmin, vmax)
                        label = labels.get(item, item)

                        if idx < 3:
                            target_cols = row1
                            target_idx = idx
                        else:
                            target_cols = row2
                            target_idx = idx - 3

                        with target_cols[target_idx]:
                            st.markdown(
                                f"""
                                <div class="small-tile" style="background:{colour}; min-height:74px;">
                                    <span class="tile-label tile-label-plain">{label}</span>
                                    {format_pct(val)}
                                </div>
                                """,
                                unsafe_allow_html=True,
                            )

                st.markdown("</div>", unsafe_allow_html=True)
                st.markdown('<div class="spacer"></div>', unsafe_allow_html=True)

            st.markdown("</div>", unsafe_allow_html=True)

    st.markdown(
        f'<div class="methodology-text">{get_methodology_paragraph("Dashboard", is_relative_mode, relative_detail_mode, effective_real_mode, inflation_source_note)}</div>',
        unsafe_allow_html=True,
    )

elif page_name == "Charts":
    detail_return_basis = st.session_state.get("detail_return_basis_toolbar", "Nominal")
    detail_period = st.session_state.get("detail_period_toolbar", "YTD")

    is_real_mode = detail_return_basis == "Real"
    effective_real_mode = is_real_mode and inflation_levels is not None and not inflation_levels.dropna().empty

    end_date_charts = get_dashboard_end_date(
        stitched_series_map=stitched_series_map,
        live_diag=live_diag,
        inflation_levels=inflation_levels,
        is_real_mode=effective_real_mode,
    )

    st.markdown('<div class="toolbar-title">Toolbar</div>', unsafe_allow_html=True)
    toolbar_cols = st.columns([1.0, 2.35, 2.65])

    with toolbar_cols[0]:
        st.markdown('<div class="toolbar-label">Return basis</div>', unsafe_allow_html=True)
        detail_return_basis = st.segmented_control(
            label="Return basis",
            options=["Nominal", "Real"],
            default=st.session_state.get("detail_return_basis_toolbar", "Nominal"),
            key="detail_return_basis_toolbar",
            label_visibility="collapsed",
        ) or "Nominal"

    is_real_mode = detail_return_basis == "Real"
    effective_real_mode = is_real_mode and inflation_levels is not None and not inflation_levels.dropna().empty

    with toolbar_cols[1]:
        st.markdown('<div class="toolbar-label">Chart period</div>', unsafe_allow_html=True)
        detail_period = st.segmented_control(
            label="Chart period",
            options=list(CHART_PERIODS.keys()),
            default=st.session_state.get("detail_period_toolbar", "YTD"),
            key="detail_period_toolbar",
            label_visibility="collapsed",
        ) or "YTD"

    with toolbar_cols[2]:
        st.markdown(
            f"""
            <div class="toolbar-meta">
                Return in GBP from <b>{common_inception_text}</b> to <b>{end_date_charts.strftime("%d/%m/%Y")}</b>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if is_real_mode and not effective_real_mode:
        st.warning("Real mode selected but no usable UK inflation series was found. Falling back to nominal results.")

    saved_default_assets = st.session_state.get("detail_selected_assets", default_chart_assets)
    saved_default_assets = [a for a in saved_default_assets if a in available_assets] or default_chart_assets

    selected_assets = st.multiselect(
        "Asset classes",
        options=available_assets,
        default=saved_default_assets,
        key="detail_selected_assets",
    )

    if not selected_assets:
        st.info("Select at least one asset class to populate the chart and tables.")
        selected_assets = []

    nominal_returns_charts_df = order_asset_rows(
        merge_return_tables(
            calc_horizon_returns_from_levels(
                stitched_series_map,
                end_date_charts,
                ["YTD", "1Y", "3Y", "5Y", "7Y", "10Y", "15Y", "20Y", "25Y"],
            ),
            calc_whole_period_returns(stitched_series_map, end_date_charts, whole_period_start),
        )
    )

    inflation_returns_charts_df = (
        order_asset_rows(
            merge_return_tables(
                calc_horizon_returns_from_levels(
                    {"UK inflation": inflation_levels},
                    end_date_charts,
                    ["YTD", "1Y", "3Y", "5Y", "7Y", "10Y", "15Y", "20Y", "25Y"],
                ),
                calc_whole_period_returns({"UK inflation": inflation_levels}, end_date_charts, whole_period_start),
            )
        )
        if inflation_levels is not None and not inflation_levels.dropna().empty
        else pd.DataFrame()
    )

    displayed_returns_charts_df = (
        order_asset_rows(convert_to_real_returns(nominal_returns_charts_df, inflation_returns_charts_df))
        if effective_real_mode
        else nominal_returns_charts_df.copy()
    )

    calendar_year_df = build_calendar_year_returns(stitched_series_map, end_date_charts, years_back=10)

    if effective_real_mode and inflation_levels is not None and not inflation_levels.dropna().empty:
        inflation_calendar_year_df = build_calendar_year_returns({"UK inflation": inflation_levels}, end_date_charts, years_back=10)
        infl_lookup = inflation_calendar_year_df.set_index("asset_class").to_dict(orient="index").get("UK inflation", {})
        calendar_year_real_df = calendar_year_df.copy()
        for year_col in [c for c in calendar_year_real_df.columns if c != "asset_class"]:
            infl_val = infl_lookup.get(year_col, np.nan)
            calendar_year_real_df[year_col] = calendar_year_real_df[year_col].map(lambda x: safe_relative_return(x, infl_val))
        calendar_year_df = calendar_year_real_df

    growth_df = build_growth_of_wealth_df(
        chart_series_map=chart_series_map,
        selected_assets=selected_assets,
        end_date=end_date_charts,
        period_key=detail_period,
        is_real_mode=effective_real_mode,
        inflation_levels=inflation_levels,
    )

    st.markdown('<div class="page-section-title">Growth of wealth</div>', unsafe_allow_html=True)
    if growth_df.empty:
        st.info("No chart data available for the current selection.")
    else:
        st.altair_chart(build_chart(growth_df, selected_assets, detail_period), width="stretch")

    st.markdown('<div class="page-section-title">Annualised returns</div>', unsafe_allow_html=True)
    returns_display_df = (
        displayed_returns_charts_df[displayed_returns_charts_df["asset_class"].isin(selected_assets)][["asset_class"] + RETURNS_TABLE_PERIODS]
        if selected_assets else displayed_returns_charts_df[["asset_class"] + RETURNS_TABLE_PERIODS]
    )
    st.markdown(
        build_html_table(
            returns_display_df,
            percent_cols=RETURNS_TABLE_PERIODS,
            conditional_cols=RETURNS_TABLE_PERIODS,
        ),
        unsafe_allow_html=True,
    )

    st.markdown('<div class="page-section-title">Calendar year returns</div>', unsafe_allow_html=True)
    calendar_display_df = (
        calendar_year_df[calendar_year_df["asset_class"].isin(selected_assets)]
        if selected_assets else calendar_year_df
    )
    year_cols = [c for c in calendar_display_df.columns if c != "asset_class"]
    st.markdown(
        build_html_table(
            calendar_display_df[["asset_class"] + year_cols],
            percent_cols=year_cols,
            conditional_cols=year_cols,
        ),
        unsafe_allow_html=True,
    )

    st.markdown(
        f'<div class="methodology-text">{get_methodology_paragraph("Charts", False, "Major", effective_real_mode, inflation_source_note)}</div>',
        unsafe_allow_html=True,
    )

elif page_name == "Risk":
    risk_return_basis = st.session_state.get("risk_return_basis_toolbar", "Nominal")
    risk_period = st.session_state.get("risk_period_toolbar", "10Y")

    is_real_mode = risk_return_basis == "Real"
    effective_real_mode = is_real_mode and inflation_levels is not None and not inflation_levels.dropna().empty

    end_date_risk = get_dashboard_end_date(
        stitched_series_map=stitched_series_map,
        live_diag=live_diag,
        inflation_levels=inflation_levels,
        is_real_mode=effective_real_mode,
    )

    st.markdown('<div class="toolbar-title">Toolbar</div>', unsafe_allow_html=True)
    toolbar_cols = st.columns([1.0, 2.4, 2.6])

    with toolbar_cols[0]:
        st.markdown('<div class="toolbar-label">Return basis</div>', unsafe_allow_html=True)
        risk_return_basis = st.segmented_control(
            label="Return basis",
            options=["Nominal", "Real"],
            default=st.session_state.get("risk_return_basis_toolbar", "Nominal"),
            key="risk_return_basis_toolbar",
            label_visibility="collapsed",
        ) or "Nominal"

    is_real_mode = risk_return_basis == "Real"
    effective_real_mode = is_real_mode and inflation_levels is not None and not inflation_levels.dropna().empty

    with toolbar_cols[1]:
        st.markdown('<div class="toolbar-label">Risk period</div>', unsafe_allow_html=True)
        period_cols = st.columns(6)
        with period_cols[0]:
            st.button("YTD", key="risk_period_ytd_disabled", disabled=True, use_container_width=True)
        for idx, period_key in enumerate(RISK_PERIODS.keys(), start=1):
            with period_cols[idx]:
                if st.button(
                    period_key,
                    key=f"risk_period_{period_key}",
                    use_container_width=True,
                    type="primary" if risk_period == period_key else "secondary",
                ):
                    st.session_state["risk_period_toolbar"] = period_key
                    st.rerun()

    with toolbar_cols[2]:
        st.markdown(
            f"""
            <div class="toolbar-meta">
                Annualised risk / return in GBP to <b>{end_date_risk.strftime("%d/%m/%Y")}</b>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if is_real_mode and not effective_real_mode:
        st.warning("Real mode selected but no usable UK inflation series was found. Falling back to nominal results.")

    saved_risk_assets = st.session_state.get("risk_selected_assets", default_chart_assets)
    saved_risk_assets = [a for a in saved_risk_assets if a in available_assets] or default_chart_assets

    selected_assets = st.multiselect(
        "Asset classes",
        options=available_assets,
        default=saved_risk_assets,
        key="risk_selected_assets",
    )

    if not selected_assets:
        st.info("Select at least one asset class to populate the chart and table.")
        selected_assets = []

    risk_series_map = {}
    for asset_class, series in stitched_series_map.items():
        risk_series = series
        if effective_real_mode:
            if inflation_levels is None or inflation_levels.dropna().empty:
                continue
            risk_series = build_real_level_series(series, inflation_levels)
        risk_series_map[asset_class] = risk_series

    risk_scatter_df = build_risk_scatter_df(
        series_map=risk_series_map,
        selected_assets=selected_assets,
        end_date=end_date_risk,
        period_key=st.session_state.get("risk_period_toolbar", risk_period),
    )
    risk_table_df = build_risk_metrics_table(
        series_map=risk_series_map,
        selected_assets=selected_assets,
        end_date=end_date_risk,
        period_keys=list(RISK_PERIODS.keys()),
    )
    risk_summary_df = build_risk_summary_table(
        series_map=risk_series_map,
        asset_style_map=asset_style_map,
        selected_assets=selected_assets,
        end_date=end_date_risk,
        start_date=whole_period_start,
    )
    correlation_matrix_df = build_correlation_matrix_table(
        series_map=risk_series_map,
        selected_assets=selected_assets,
        end_date=end_date_risk,
        start_date=whole_period_start,
    )

    st.markdown('<div class="page-section-title">Volatility/return chart</div>', unsafe_allow_html=True)
    if risk_scatter_df.empty:
        st.info("No risk data available for the current selection.")
    else:
        st.altair_chart(build_risk_scatter_chart(risk_scatter_df, selected_assets), width="stretch")

    st.markdown('<div class="page-section-title">Volatility/return table</div>', unsafe_allow_html=True)
    risk_percent_cols = [c for c in risk_table_df.columns if c != "asset_class"]
    st.markdown(
        build_html_table(
            risk_table_df,
            percent_cols=risk_percent_cols,
            conditional_cols=risk_percent_cols,
            header_wrap_cols=risk_percent_cols,
            invert_conditional_cols=[c for c in risk_percent_cols if c.endswith("Vol")],
            rank_conditional_cols=[c for c in risk_percent_cols if c.endswith("Vol")],
        ),
        unsafe_allow_html=True,
    )

    st.markdown('<div class="page-section-title">Since inception risk metrics</div>', unsafe_allow_html=True)
    risk_summary_display_df = risk_summary_df.copy()
    ratio_cols = [
        "Return/vol ratio",
        "Sharpe ratio",
        "Information ratio",
        "Sortino ratio",
        "Calmar ratio",
    ]
    for col in ratio_cols:
        if col in risk_summary_display_df.columns:
            risk_summary_display_df[col] = risk_summary_display_df[col].map(
                lambda x: np.nan if pd.isna(x) else round(float(x), 2)
            )

    risk_summary_percent_cols = [
        c for c in ["Period return", "Period vol", "Worst drawdown", "Tracking error"] if c in risk_summary_display_df.columns
    ]
    risk_summary_conditional_cols = [c for c in risk_summary_display_df.columns if c != "asset_class"]
    st.markdown(
        build_html_table(
            risk_summary_display_df,
            percent_cols=risk_summary_percent_cols,
            conditional_cols=risk_summary_conditional_cols,
            header_wrap_cols=risk_summary_conditional_cols,
            invert_conditional_cols=["Period vol", "Tracking error"],
            rank_conditional_cols=["Period vol", "Worst drawdown", "Tracking error"],
            decimal_cols=ratio_cols,
        ),
        unsafe_allow_html=True,
    )

    st.markdown('<div class="page-section-title">Correlation matrix</div>', unsafe_allow_html=True)
    correlation_cols = [c for c in correlation_matrix_df.columns if c != "asset_class"]
    st.markdown(
        build_html_table(
            correlation_matrix_df,
            conditional_cols=correlation_cols,
            header_wrap_cols=correlation_cols,
            decimal_cols=correlation_cols,
            correlation_conditional_cols=correlation_cols,
        ),
        unsafe_allow_html=True,
    )

    risk_methodology = (
        f"This tab shows annualised {'real' if effective_real_mode else 'nominal'} return and annualised volatility in GBP "
        f"to <b>{end_date_risk.strftime('%d/%m/%Y')}</b>. Volatility is calculated from monthly returns and annualised using the square root of 12."
    )
    if effective_real_mode:
        risk_methodology += f" Current inflation source: {inflation_source_note}."
    risk_methodology += (
        " Glossary: Return/vol ratio is annualised return divided by annualised volatility. "
        "Sharpe ratio is excess return over Cash (GBP) divided by volatility. "
        "Information ratio is excess return versus the assigned benchmark divided by tracking error. "
        "Sortino ratio is return divided by downside deviation. "
        "Worst drawdown is the largest peak-to-trough fall since common inception. "
        "Calmar ratio is annualised return divided by the absolute worst drawdown. "
        "Tracking error is the annualised standard deviation of monthly excess returns versus the assigned benchmark. "
        "The assigned benchmark is Global market for growth assets and Cash (GBP) for defensive assets."
    )
    st.markdown(f'<div class="methodology-text">{risk_methodology}</div>', unsafe_allow_html=True)

else:
    st.markdown('<div class="page-section-title">Yield curves</div>', unsafe_allow_html=True)
    selected_yield_series = st.multiselect(
        "Curve series",
        options=["Nominal", "Real", "Breakeven inflation"],
        default=st.session_state.get("yield_selected_series", ["Nominal", "Real", "Breakeven inflation"]),
        key="yield_selected_series",
    )
    boe_fetch_ok = (
        not boe_yield_summary.empty
        and "Fetch status" in boe_yield_summary["metric"].values
        and boe_yield_summary.loc[boe_yield_summary["metric"] == "Fetch status", "value"].astype(str).iloc[0] == "OK"
    )
    yield_curve_display_df = build_yield_curve_display_df(yield_curve_df, selected_yield_series)
    has_curve_points = not yield_curve_display_df.empty

    if not boe_fetch_ok and not has_curve_points:
        st.warning("Bank of England yield-curve data could not be loaded. See diagnostics for details.")
    else:
        latest_dates = (
            yield_curve_display_df.groupby("curve_type")["curve_date"]
            .max()
            .dropna()
            .sort_index()
        ) if has_curve_points else pd.Series(dtype="datetime64[ns]")
        latest_text = " | ".join(
            [f"{curve}: {pd.Timestamp(curve_date).strftime('%d/%m/%Y')}" for curve, curve_date in latest_dates.items()]
        ) if not latest_dates.empty else ""
        if latest_text:
            st.markdown(
                f'<div class="toolbar-meta" style="text-align:left; padding-top:0; margin-bottom:10px;">Latest curve dates: <b>{latest_text}</b></div>',
                unsafe_allow_html=True,
            )
        if has_curve_points:
            st.altair_chart(build_yield_curve_chart(yield_curve_display_df), width="stretch")
        else:
            st.info("BOE fetch succeeded but no curve points were available to plot. See diagnostics for details.")

    yield_note = (
        "Latest nominal and real UK spot curves are fetched from the Bank of England yield-curve archive. "
        "The app reads the latest non-blank dated row from the '4. spot curve' sheet in the current-month nominal and real daily workbooks. "
        "Where DividendData provides shorter-maturity index-linked gilts than the Bank of England real curve, those short-end gilt real yields are prepended to extend the real curve only below the first BOE real maturity. "
        "Breakeven inflation is then calculated point-by-point as ((1+nominal yield)/(1+real yield))-1 on the maturities where both nominal and real yields are available."
    )
    st.markdown(f'<div class="methodology-text">{yield_note}</div>', unsafe_allow_html=True)


# =====================================
# FOOTER
# =====================================
if POWERED_BY_FILE.exists():
    st.markdown(
        f"""
        <div class="footer-bar">
            <div class="footer-logo">
                <img src="data:image/png;base64,{img_to_base64(str(POWERED_BY_FILE), POWERED_BY_FILE.stat().st_mtime)}">
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
    generic_end_date = get_dashboard_end_date(
        stitched_series_map=stitched_series_map,
        live_diag=live_diag,
        inflation_levels=inflation_levels,
        is_real_mode=False,
    )
    dashboard_anchor_diag = build_return_anchor_table(
        stitched_series_map,
        generic_end_date,
        ["YTD", "1Y", "3Y", "5Y", "10Y", "20Y", "Period"],
        whole_period_start,
    )
    chart_anchor_diag = build_return_anchor_table(
        chart_series_map,
        generic_end_date,
        ["YTD", "1Y", "3Y", "5Y", "10Y", "20Y", "Period"],
        whole_period_start,
    )
    inflation_anchor_diag = (
        build_return_anchor_table(
            {"UK inflation": inflation_levels},
            generic_end_date,
            ["YTD", "1Y", "3Y", "5Y", "10Y", "20Y", "Period"],
            whole_period_start,
        )
        if inflation_levels is not None and not inflation_levels.dropna().empty
        else pd.DataFrame()
    )

    with st.container(border=True):
        st.markdown('<div class="diag-title">Live fund stitching diagnostics</div>', unsafe_allow_html=True)
        st.markdown(
            f"""
            <div class="diag-note">
                Index history is always preferred where available. Live yfinance history is only used after index history ends for the dashboard and tables.
                The growth of wealth chart uses a separate cached series that prioritises daily live-fund history where available and stitches monthly index history before that to extend the chart backwards.
                Inflation source: <b>{inflation_source_note}</b>. Generic dashboard end date: <b>{generic_end_date.strftime("%d/%m/%Y")}</b>.
                Common inception date used for period table: <b>{common_inception_text}</b>.
            </div>
            """,
            unsafe_allow_html=True,
        )

        if inflation_debug_message:
            st.info(inflation_debug_message)

        tabs = st.tabs(["Overview", "Dashboard stitching", "Chart series", "Returns", "Inflation", "Yield curves", "Mapping & prices"])

        with tabs[0]:
            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("Mapped rows", int(len(mapping)))
            c2.metric("Unique asset classes", int(mapping["asset_class"].nunique()))
            c3.metric("Monthly index series", int(len(monthly_levels)))
            c4.metric("Dashboard series", int(len(stitched_series_map)))
            c5.metric("Chart series", int(len(chart_series_map)))
            c6.metric("Live tickers returned", int(len(live_prices.columns)) if not live_prices.empty else 0)

            st.subheader("Asset coverage summary")
            st.dataframe(prepare_dataframe_for_display(asset_coverage_diag), width="stretch", hide_index=True)
            st.download_button(
                label="Download asset coverage (CSV)",
                data=dataframe_to_csv_download(asset_coverage_diag),
                file_name="asset_coverage_diagnostics.csv",
                mime="text/csv",
                key="download_asset_coverage_csv",
            )

            missing_assets = asset_coverage_diag[
                (asset_coverage_diag["index_points"] == 0)
                | (asset_coverage_diag["dashboard_points"] == 0)
                | (asset_coverage_diag["chart_points"] == 0)
            ].copy()
            st.subheader("Coverage gaps")
            if missing_assets.empty:
                st.write("No major coverage gaps detected in mapped assets.")
            else:
                st.dataframe(prepare_dataframe_for_display(missing_assets), width="stretch", hide_index=True)

        with tabs[1]:
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
                st.dataframe(prepare_dataframe_for_display(summary), width="stretch", hide_index=True)

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
                st.dataframe(prepare_dataframe_for_display(diag_show[display_cols]), width="stretch", hide_index=True)

                st.download_button(
                    label="Download live diagnostics (CSV)",
                    data=dataframe_to_csv_download(live_diag),
                    file_name="live_fund_stitch_diagnostics.csv",
                    mime="text/csv",
                    use_container_width=False,
                    key="download_live_diagnostics_csv",
                )
            else:
                st.write("No live diagnostics available.")

        with tabs[2]:
            if not chart_diag.empty:
                summary = (
                    chart_diag["series_type"]
                    .fillna("unknown")
                    .value_counts(dropna=False)
                    .rename_axis("series_type")
                    .reset_index(name="count")
                )
                st.subheader("Summary")
                st.dataframe(prepare_dataframe_for_display(summary), width="stretch", hide_index=True)

                chart_show = format_diagnostic_table(chart_diag)
                st.subheader("Chart series diagnostics")
                st.dataframe(prepare_dataframe_for_display(chart_show), width="stretch", hide_index=True)
                st.download_button(
                    label="Download chart diagnostics (CSV)",
                    data=dataframe_to_csv_download(chart_diag),
                    file_name="chart_series_diagnostics.csv",
                    mime="text/csv",
                    key="download_chart_diagnostics_csv",
                )
            else:
                st.write("No chart series diagnostics available.")

        with tabs[3]:
            returns_tabs = st.tabs(["Dashboard returns", "Charts returns", "Calendar year returns", "Return anchors"])

            with returns_tabs[0]:
                dashboard_nominal = order_asset_rows(
                    calc_horizon_returns_from_levels(stitched_series_map, generic_end_date, list(DASHBOARD_HORIZONS.keys()))
                )
                st.dataframe(convert_pct_table_for_display(dashboard_nominal), width="stretch", hide_index=True)
                st.download_button(
                    label="Download dashboard returns (CSV)",
                    data=dataframe_to_csv_download(dashboard_nominal),
                    file_name="dashboard_returns_diagnostics.csv",
                    mime="text/csv",
                    key="download_dashboard_returns_diag_csv",
                )

            with returns_tabs[1]:
                charts_nominal = order_asset_rows(
                    merge_return_tables(
                        calc_horizon_returns_from_levels(
                            stitched_series_map,
                            generic_end_date,
                            ["YTD", "1Y", "3Y", "5Y", "7Y", "10Y", "15Y", "20Y", "25Y"],
                        ),
                        calc_whole_period_returns(stitched_series_map, generic_end_date, whole_period_start),
                    )
                )
                st.dataframe(convert_pct_table_for_display(charts_nominal), width="stretch", hide_index=True)
                st.download_button(
                    label="Download charts returns (CSV)",
                    data=dataframe_to_csv_download(charts_nominal),
                    file_name="charts_returns_diagnostics.csv",
                    mime="text/csv",
                    key="download_charts_returns_diag_csv",
                )

            with returns_tabs[2]:
                calendar_diag = build_calendar_year_returns(stitched_series_map, generic_end_date, 10)
                st.dataframe(convert_pct_table_for_display(calendar_diag), width="stretch", hide_index=True)
                st.download_button(
                    label="Download calendar returns (CSV)",
                    data=dataframe_to_csv_download(calendar_diag),
                    file_name="calendar_year_returns_diagnostics.csv",
                    mime="text/csv",
                    key="download_calendar_returns_diag_csv",
                )

            with returns_tabs[3]:
                anchor_tabs = st.tabs(["Dashboard anchor dates", "Chart anchor dates"])

                with anchor_tabs[0]:
                    st.dataframe(prepare_dataframe_for_display(dashboard_anchor_diag), width="stretch", hide_index=True)
                    st.download_button(
                        label="Download dashboard anchors (CSV)",
                        data=dataframe_to_csv_download(dashboard_anchor_diag),
                        file_name="dashboard_return_anchor_diagnostics.csv",
                        mime="text/csv",
                        key="download_dashboard_anchor_diag_csv",
                    )

                with anchor_tabs[1]:
                    st.dataframe(prepare_dataframe_for_display(chart_anchor_diag), width="stretch", hide_index=True)
                    st.download_button(
                        label="Download chart anchors (CSV)",
                        data=dataframe_to_csv_download(chart_anchor_diag),
                        file_name="chart_return_anchor_diagnostics.csv",
                        mime="text/csv",
                        key="download_chart_anchor_diag_csv",
                    )

        with tabs[4]:
            c1, c2 = st.columns(2)

            with c1:
                if inflation_levels is not None and not inflation_levels.dropna().empty:
                    infl_diag_table = order_asset_rows(
                        merge_return_tables(
                            calc_horizon_returns_from_levels(
                                {"UK inflation": inflation_levels},
                                generic_end_date,
                                ["YTD", "1Y", "3Y", "5Y", "7Y", "10Y", "15Y", "20Y", "25Y"],
                            ),
                            calc_whole_period_returns({"UK inflation": inflation_levels}, generic_end_date, whole_period_start),
                        )
                    )
                    st.subheader("Inflation return table")
                    st.dataframe(convert_pct_table_for_display(infl_diag_table), width="stretch", hide_index=True)
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

            st.subheader("ONS fetch diagnostics")
            st.dataframe(prepare_dataframe_for_display(ons_fetch_summary), width="stretch", hide_index=True)

            if not ons_fetch_preview.empty:
                ons_preview_show = ons_fetch_preview.copy()
                ons_preview_show["value"] = ons_preview_show["value"].map(lambda x: np.nan if pd.isna(x) else round(float(x), 4))
                st.dataframe(prepare_dataframe_for_display(ons_preview_show.tail(24)), width="stretch", hide_index=True)
                st.download_button(
                    label="Download ONS parse preview (CSV)",
                    data=dataframe_to_csv_download(ons_fetch_preview),
                    file_name="ons_cpi_parse_preview.csv",
                    mime="text/csv",
                    key="download_ons_parse_preview_csv",
                )

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

                lower_tabs = st.tabs(["Inflation levels tail", "Inflation monthly tail", "Inflation anchors"])

                with lower_tabs[0]:
                    st.dataframe(prepare_dataframe_for_display(levels_df.tail(24)), width="stretch", hide_index=True)

                with lower_tabs[1]:
                    st.dataframe(prepare_dataframe_for_display(monthly_df.tail(24)), width="stretch", hide_index=True)

                with lower_tabs[2]:
                    st.dataframe(prepare_dataframe_for_display(inflation_anchor_diag), width="stretch", hide_index=True)
            else:
                st.write("No inflation series available.")

        with tabs[5]:
            st.subheader("Bank of England yield-curve diagnostics")
            st.dataframe(prepare_dataframe_for_display(boe_yield_summary), width="stretch", hide_index=True)
            if not boe_yield_preview.empty:
                st.dataframe(prepare_dataframe_for_display(boe_yield_preview), width="stretch", hide_index=True)
                st.download_button(
                    label="Download BOE yield diagnostics (CSV)",
                    data=dataframe_to_csv_download(boe_yield_preview),
                    file_name="boe_yield_curve_diagnostics.csv",
                    mime="text/csv",
                    key="download_boe_yield_diag_csv",
                )
            if not yield_curve_df.empty:
                curve_preview = yield_curve_df.copy()
                curve_preview["curve_date"] = pd.to_datetime(curve_preview["curve_date"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                st.subheader("Latest BOE spot-curve points")
                st.dataframe(prepare_dataframe_for_display(curve_preview), width="stretch", hide_index=True)

        with tabs[6]:
            mapping_tabs = st.tabs(["Validated mapping", "Raw mapping", "Live prices"])

            with mapping_tabs[0]:
                st.subheader("Validated mapping health")
                st.dataframe(prepare_dataframe_for_display(mapping_diag), width="stretch", hide_index=True)
                st.download_button(
                    label="Download mapping diagnostics (CSV)",
                    data=dataframe_to_csv_download(mapping_diag),
                    file_name="mapping_diagnostics.csv",
                    mime="text/csv",
                    key="download_mapping_diag_csv",
                )

            with mapping_tabs[1]:
                st.subheader("Mapping table")
                st.dataframe(prepare_dataframe_for_display(mapping), width="stretch", hide_index=True)

            with mapping_tabs[2]:
                st.subheader("Live price coverage")
                if live_price_diag.empty:
                    st.write("No live price history available.")
                else:
                    st.dataframe(prepare_dataframe_for_display(live_price_diag), width="stretch", hide_index=True)
                    st.download_button(
                        label="Download live price coverage (CSV)",
                        data=dataframe_to_csv_download(live_price_diag),
                        file_name="live_price_coverage.csv",
                        mime="text/csv",
                        key="download_live_price_coverage_csv",
                    )
