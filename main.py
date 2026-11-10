"""
Project 1: Spreadsheet automation report
Cleans messy_sales_2026.csv, summarizes it, and exports an Excel report + chart.

Run it:
    python clean_sales.py                      (uses messy_sales_2026.csv)
    python clean_sales.py some_other_file.csv  (uses a different file)

Requires: pandas (2.0 or newer), openpyxl, matplotlib
    pip install pandas openpyxl matplotlib
"""

import sys

import matplotlib

matplotlib.use("Agg")  # lets matplotlib save images without opening a window
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl.drawing.image import Image as XLImage

# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------
INPUT_FILE = sys.argv[1] if len(sys.argv) > 1 else "messy_sales_2026.csv"
OUTPUT_EXCEL = "clean_sales_report.xlsx"
OUTPUT_CHART = "sales_by_category.png"
MAX_REASONABLE_QTY = 50  # anything above this is treated as a data-entry error

# Lookup tables: messy value (lowercase, no spaces) -> clean value.
# These are the "fix the typos" dictionaries.
CATEGORY_MAP = {
    "electronics": "Electronics",
    "electonics": "Electronics",
    "groceries": "Groceries",
    "grocery": "Groceries",
    "beverages": "Beverages",
    "beverage": "Beverages",
    "drinks": "Beverages",
    "personalcare": "Personal Care",
    "schoolsupplies": "School Supplies",
    "supplies": "School Supplies",
}
BRANCH_MAP = {
    "mainbranch": "Main Branch",
    "main": "Main Branch",
    "downtown": "Downtown",
    "marketst": "Market Street",
    "marketstreet": "Market Street",
}
PAYMENT_MAP = {
    "cash": "Cash",
    "gcash": "GCash",
    "g-cash": "GCash",
    "maya": "Maya",
    "cod": "COD",
    "creditcard": "Credit Card",
    "card": "Credit Card",
}
NUMBER_WORDS = {
    "one": "1",
    "two": "2",
    "three": "3",
    "four": "4",
    "five": "5",
    "six": "6",
}

# A running list of every cleaning decision (great for your README).
cleaning_log = []


def log(message):
    """Print a message and remember it for the cleaning log."""
    print(message)
    cleaning_log.append(message)


# ----------------------------------------------------------------------------
# Step 1: Load the file
# ----------------------------------------------------------------------------
def load_data(path):
    # skiprows=2 skips the two title lines above the real header.
    # dtype=str keeps everything as text so we control how it is converted.
    # keep_default_na=False stops pandas from silently turning "N/A" into missing.
    df = pd.read_csv(path, skiprows=2, dtype=str, keep_default_na=False)

    # Tidy column names: "Unit Price " -> "unit_price"
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

    # Remove stray spaces from every cell
    df = df.apply(lambda col: col.str.strip())

    log(f"Loaded {len(df)} rows from {path}")

    # Remove the footer "TOTAL" row (and any row with no order id)
    before = len(df)
    df = df[(df["order_id"] != "") & (df["order_id"].str.upper() != "TOTAL")]
    log(f"Removed {before - len(df)} junk row(s) such as the TOTAL footer")
    return df.copy()


# ----------------------------------------------------------------------------
# Step 2: Clean each column
# ----------------------------------------------------------------------------
def standardize(series, mapping, column_name):
    """Lowercase + remove spaces, then look the value up in the mapping."""
    key = series.str.lower().str.replace(" ", "", regex=False)
    cleaned = key.map(mapping)
    unmapped = cleaned.isna().sum()
    if unmapped:
        log(f"WARNING: {unmapped} value(s) in '{column_name}' not recognised")
    return cleaned


def clean_columns(df):
    # Text columns
    df["category"] = standardize(df["category"], CATEGORY_MAP, "category")
    df["branch"] = standardize(df["branch"], BRANCH_MAP, "branch")
    df["payment_method"] = standardize(
        df["payment_method"], PAYMENT_MAP, "payment_method"
    )
    log("Standardized category, branch and payment_method spellings")

    # Customer names: fix capitalization, label blanks as Unknown
    df["customer_name"] = df["customer_name"].str.title()
    blanks = (df["customer_name"] == "").sum()
    df["customer_name"] = df["customer_name"].replace("", "Unknown")
    log(f"Filled {blanks} blank customer name(s) with 'Unknown'")

    # Dates: format="mixed" lets pandas read each date in its own format.
    # Caution: 05/03/2026 is ambiguous (May 3 or March 5?). pandas assumes
    # month-first. For real work you would confirm this with the data owner.
    df["date"] = pd.to_datetime(df["date"], format="mixed", errors="coerce")
    bad_dates = df["date"].isna().sum()
    log(f"Parsed dates ({bad_dates} could not be read)")

    # Prices: remove everything except digits and the decimal point,
    # so "₱1,250.00" and "PHP 150" both become numbers.
    df["unit_price"] = df["unit_price"].str.replace(r"[^\d.]", "", regex=True)
    df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce")  # "N/A" -> NaN

    # Fill missing prices using the typical price of the same product.
    missing_prices = df["unit_price"].isna().sum()
    typical_price = df.groupby("product")["unit_price"].transform("median")
    df["unit_price"] = df["unit_price"].fillna(typical_price)
    log(f"Filled {missing_prices} missing price(s) using each product's typical price")

    # Quantities: words -> numbers, then convert to numeric.
    df["qty"] = df["qty"].str.lower().replace(NUMBER_WORDS)
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce")

    # Negative quantities: legitimate for REFUND rows, typos everywhere else.
    is_refund = df["notes"].str.upper() == "REFUND"
    df["qty"] = df["qty"].abs()
    df["qty"] = df["qty"].where(~is_refund, -df["qty"])
    log(
        f"Kept {is_refund.sum()} refund row(s) as negative quantities; "
        "made other negative quantities positive"
    )

    # Outliers: quantities that are implausibly large become missing.
    outliers = (df["qty"].abs() > MAX_REASONABLE_QTY).sum()
    df.loc[df["qty"].abs() > MAX_REASONABLE_QTY, "qty"] = pd.NA
    log(
        f"Marked {outliers} outlier quantity(ies) above {MAX_REASONABLE_QTY} as missing"
    )
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce")

    return df


# ----------------------------------------------------------------------------
# Step 3: Remove duplicates and rows we cannot use
# ----------------------------------------------------------------------------
def remove_bad_rows(df):
    before = len(df)
    df = df.drop_duplicates(subset="order_id")
    log(f"Removed {before - len(df)} duplicate order(s)")

    before = len(df)
    df = df.dropna(subset=["qty", "unit_price"])
    log(f"Dropped {before - len(df)} row(s) with no usable quantity or price")
    return df


# ----------------------------------------------------------------------------
# Step 4: Summaries (your "pivot tables")
# ----------------------------------------------------------------------------
def build_summaries(df):
    df["total"] = df["qty"] * df["unit_price"]
    df["month"] = df["date"].dt.to_period("M").astype(str)

    def summarize(column):
        out = (
            df.groupby(column)["total"]
            .agg(orders="count", total_sales="sum")
            .sort_values("total_sales", ascending=False)
            .reset_index()
        )
        out["total_sales"] = out["total_sales"].round(2)
        return out

    return {
        "By Category": summarize("category"),
        "By Branch": summarize("branch"),
        "By Payment": summarize("payment_method"),
        "By Month": summarize("month").sort_values("month"),
        "Top 5 Products": summarize("product").head(5),
    }


# ----------------------------------------------------------------------------
# Step 5: Chart and Excel export
# ----------------------------------------------------------------------------
def make_chart(category_summary):
    ax = category_summary.plot(
        kind="bar",
        x="category",
        y="total_sales",
        legend=False,
        figsize=(8, 5),
        color="#4C78A8",
    )
    ax.set_title("Total Sales by Category")
    ax.set_xlabel("")
    ax.set_ylabel("Sales (PHP)")
    plt.xticks(rotation=30, ha="right")
    plt.tight_layout()
    plt.savefig(OUTPUT_CHART, dpi=120)
    plt.close()


def export_excel(df, summaries):
    clean_columns_order = [
        "order_id",
        "date",
        "customer_name",
        "branch",
        "product",
        "category",
        "qty",
        "unit_price",
        "total",
        "payment_method",
        "notes",
    ]

    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        df[clean_columns_order].sort_values("date").to_excel(
            writer, sheet_name="Clean Data", index=False
        )
        for sheet_name, table in summaries.items():
            table.to_excel(writer, sheet_name=sheet_name, index=False)

        # Put the chart next to the category summary
        writer.sheets["By Category"].add_image(XLImage(OUTPUT_CHART), "E2")

        # Cleaning log sheet
        pd.DataFrame({"Cleaning steps": cleaning_log}).to_excel(
            writer, sheet_name="Cleaning Log", index=False
        )

        # Make columns wide enough to read
        for sheet in writer.sheets.values():
            for column_cells in sheet.columns:
                width = max(len(str(cell.value or "")) for cell in column_cells[:50])
                sheet.column_dimensions[column_cells[0].column_letter].width = min(
                    width + 3, 60
                )

    log(f"Saved report to {OUTPUT_EXCEL}")


# ----------------------------------------------------------------------------
# Run everything
# ----------------------------------------------------------------------------
def main():
    df = load_data(INPUT_FILE)
    df = clean_columns(df)
    df = remove_bad_rows(df)
    summaries = build_summaries(df)

    print("\nFinal clean dataset:", df.shape[0], "rows")
    print("\nSales by category:\n", summaries["By Category"])
    print("\nSales by branch:\n", summaries["By Branch"])
    print("\nTop 5 products:\n", summaries["Top 5 Products"])

    make_chart(summaries["By Category"])
    export_excel(df, summaries)


if __name__ == "__main__":
    main()
