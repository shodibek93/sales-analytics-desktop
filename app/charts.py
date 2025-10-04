import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

def revenue_trend_with_fit(monthly_df):
    fig, ax = plt.subplots()
    if len(monthly_df) == 0:
        ax.text(0.5, 0.5, "No data", ha="center", va="center")
        ax.axis("off")
        return fig
    x = np.arange(len(monthly_df))
    y = monthly_df["sales_amount"].to_numpy()
    ax.plot(monthly_df["month"], y, marker="o", label="Revenue")
    if len(x) >= 2:
        m, b = np.polyfit(x, y, 1)
        y_fit = m * x + b
        ax.plot(monthly_df["month"], y_fit, linestyle="--", label="Trend")
        ax.legend()
    ax.set_title("Revenue over time (with trend)")
    ax.set_xlabel("Month")
    ax.set_ylabel("Revenue")
    fig.autofmt_xdate()
    return fig

def regional_pie(reg_df):
    fig, ax = plt.subplots()
    if len(reg_df) == 0 or reg_df["sales_amount"].sum() == 0:
        ax.text(0.5, 0.5, "No data", ha="center", va="center")
        ax.axis("off")
        return fig
    ax.pie(reg_df["sales_amount"], labels=reg_df["region"], autopct="%1.1f%%")
    ax.set_title("Regional revenue share")
    return fig

def quarterly_trend_chart(q_df):
    fig, ax = plt.subplots()
    if len(q_df) == 0:
        ax.text(0.5, 0.5, "No data", ha="center", va="center")
        ax.axis("off")
        return fig
    ax.bar(q_df["quarter"].astype(str), q_df["sales_amount"])
    ax.set_title("Quarterly revenue")
    ax.set_xlabel("Quarter")
    ax.set_ylabel("Revenue")
    return fig

def margin_hist(df):
    fig, ax = plt.subplots()
    margins = df["margin"].dropna()
    if len(margins) == 0:
        ax.text(0.5, 0.5, "No data", ha="center", va="center")
        ax.axis("off")
        return fig
    ax.hist(margins, bins=30)
    ax.set_title("Profit margin distribution")
    ax.set_xlabel("Margin")
    ax.set_ylabel("Frequency")
    return fig


# ---- Heatmap (append at end of file) ----
def heatmap_product_month(pivot_df):
    fig, ax = plt.subplots()
    if pivot_df is None or pivot_df.empty or pivot_df.shape[1] <= 1:
        ax.text(0.5, 0.5, "No data", ha="center", va="center")
        ax.axis("off")
        return fig

    products = pivot_df["product"].astype(str).tolist()
    month_cols = [c for c in pivot_df.columns if c != "product"]
    data = pivot_df[month_cols].to_numpy(dtype=float)

    im = ax.imshow(data, aspect="auto")
    ax.set_title("Product × Month — Profit heatmap")
    ax.set_yticks(np.arange(len(products)))
    ax.set_yticklabels(products)

    def _fmt_month(c):
        try:
            ts = pd.to_datetime(c, errors="coerce")
            return ts.strftime("%Y-%m") if pd.notna(ts) else str(c)
        except Exception:
            return str(c)

    month_labels = [_fmt_month(c) for c in month_cols]
    ax.set_xticks(np.arange(len(month_cols)))
    ax.set_xticklabels(month_labels, rotation=45, ha="right")

    fig.colorbar(im, ax=ax, label="Profit")
    fig.tight_layout()
    return fig

