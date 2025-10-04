import pandas as pd
import numpy as np

REQUIRED_COLS = ["date", "product", "region", "sales_amount", "cost", "customer_type"]

def load_and_validate(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    lc = {c.lower().strip(): c for c in df.columns}
    missing = [c for c in REQUIRED_COLS if c not in lc]
    if missing:
        raise ValueError(f"Missing columns: {missing}. Present: {list(df.columns)}")
    df = df.rename(columns={lc[c]: c for c in REQUIRED_COLS})
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["sales_amount"] = pd.to_numeric(df["sales_amount"], errors="coerce")
    df["cost"] = pd.to_numeric(df["cost"], errors="coerce")
    df = df.dropna(subset=["date", "sales_amount", "cost"]).copy()
    df["profit"] = df["sales_amount"] - df["cost"]
    df["margin"] = np.where(df["sales_amount"] > 0, df["profit"] / df["sales_amount"], np.nan)
    df["month"] = df["date"].dt.to_period("M").dt.to_timestamp()
    df["quarter"] = df["date"].dt.to_period("Q").dt.to_timestamp()
    return df

def kpi(df: pd.DataFrame) -> dict:
    total_rev = float(df["sales_amount"].sum())
    avg_rev   = float(df["sales_amount"].mean())
    total_profit = float(df["profit"].sum())
    avg_margin = df["margin"].mean(skipna=True)
    avg_margin = float(avg_margin) if pd.notna(avg_margin) else None

    by_month = df.groupby("month", as_index=False)["sales_amount"].sum().sort_values("month")
    growth = None
    if len(by_month) >= 2:
        last, prev = by_month.iloc[-1]["sales_amount"], by_month.iloc[-2]["sales_amount"]
        growth = float((last - prev) / prev) if prev else None

    return {
        "total_revenue": total_rev,
        "avg_revenue": avg_rev,
        "total_profit": total_profit,
        "avg_margin": avg_margin,
        "growth_mom": growth
    }

def monthly_trends(df: pd.DataFrame) -> pd.DataFrame:
    return df.groupby("month", as_index=False)[["sales_amount","profit"]].sum().sort_values("month")

def quarterly_trends(df: pd.DataFrame) -> pd.DataFrame:
    return df.groupby("quarter", as_index=False)[["sales_amount","profit"]].sum().sort_values("quarter")

def regional_breakdown(df: pd.DataFrame) -> pd.DataFrame:
    return df.groupby("region", as_index=False)[["sales_amount","profit"]].sum().sort_values("sales_amount", ascending=False)

def top_bottom_products(df: pd.DataFrame, n:int=10):
    by_prod = df.groupby("product", as_index=False)[["sales_amount","profit"]].sum()
    top = by_prod.sort_values("profit", ascending=False).head(n)
    bottom = by_prod.sort_values("profit", ascending=True).head(n)
    return top, bottom

# --- Extra aggregations for Full Excel ---
def by_customer_type(df: pd.DataFrame) -> pd.DataFrame:
    return df.groupby("customer_type", as_index=False)[["sales_amount","profit"]].sum().sort_values("sales_amount", ascending=False)

def product_month_pivot_profit(df: pd.DataFrame) -> pd.DataFrame:
    p = df.pivot_table(index="product", columns="month", values="profit", aggfunc="sum", fill_value=0)
    return p.reset_index()

def region_month_pivot_sales(df: pd.DataFrame) -> pd.DataFrame:
    p = df.pivot_table(index="region", columns="month", values="sales_amount", aggfunc="sum", fill_value=0)
    return p.reset_index()

def margins_describe(df: pd.DataFrame) -> pd.DataFrame:
    return df["margin"].describe(percentiles=[0.1,0.25,0.5,0.75,0.9]).to_frame(name="margin").reset_index(names=["stat"])

def data_dictionary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for col in df.columns:
        dtype = str(df[col].dtype)
        nnull = int(df[col].isna().sum())
        rows.append({"column": col, "dtype": dtype, "nulls": nnull})
    return pd.DataFrame(rows)

def monthly_growth_table(df: pd.DataFrame) -> pd.DataFrame:
    m = df.groupby("month", as_index=False)["sales_amount"].sum().sort_values("month")
    m["growth_mom"] = m["sales_amount"].pct_change()
    return m

# ---- Filters (append at end of file) ----
from typing import Iterable, Tuple

def get_filter_options(df: pd.DataFrame) -> dict:
    """Return unique lists and date bounds for UI controls."""
    if df.empty:
        return {"regions": [], "customer_types": [], "date_min": None, "date_max": None}
    return {
        "regions": sorted(df["region"].dropna().unique().tolist()),
        "customer_types": sorted(df["customer_type"].dropna().unique().tolist()),
        "date_min": pd.to_datetime(df["date"].min()).date(),
        "date_max": pd.to_datetime(df["date"].max()).date(),
    }

def apply_filters(
    df: pd.DataFrame,
    date_from=None,
    date_to=None,
    regions: Iterable[str] | None = None,
    customer_types: Iterable[str] | None = None,
) -> pd.DataFrame:
    """Return filtered dataframe by date range, regions, customer types."""
    if df is None or df.empty:
        return df
    out = df.copy()
    if date_from is not None:
        out = out[out["date"] >= pd.to_datetime(date_from)]
    if date_to is not None:
        out = out[out["date"] <= pd.to_datetime(date_to)]
    if regions:
        out = out[out["region"].isin(regions)]
    if customer_types:
        out = out[out["customer_type"].isin(customer_types)]
    return out

def product_month_pivot_profit_filtered(df: pd.DataFrame) -> pd.DataFrame:
    """Pivot Product × Month (profit). Safe for empty df."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["product"])
    p = df.pivot_table(index="product", columns="month", values="profit", aggfunc="sum", fill_value=0)
    return p.reset_index()
