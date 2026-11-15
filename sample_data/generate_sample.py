
from pathlib import Path
import numpy as np
import pandas as pd


def generate_sample_excel():

    rng = np.random.default_rng(42)

    output_dir = Path(__file__).resolve().parent
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = output_dir / "business_sales.xlsx"

    products = [
        ("Laptop Pro", "Electronics", 125000),
        ("Smartphone X", "Electronics", 85000),
        ("Office Chair", "Furniture", 22000),
        ("Desk Premium", "Furniture", 35000),
        ("Wireless Mouse", "Accessories", 4500),
        ("Monitor 27", "Electronics", 42000),
        ("Keyboard Mechanical", "Accessories", 9500),
        ("Headphones", "Accessories", 12000)
    ]

    regions = [
        "Algiers",
        "Blida",
        "Oran",
        "Constantine",
        "Setif"
    ]

    customers = [
        "Amine Benali",
        "Sarah Haddad",
        "Yacine Merabet",
        "Lina Mansouri",
        "Nadia Kaci",
        "Karim Boudiaf",
        "Meriem Saidi",
        "Sofiane Ali"
    ]

    payment_methods = [
        "Cash",
        "Bank Transfer",
        "Credit Card",
        "Edahabia"
    ]

    statuses = [
        "Completed",
        "Completed",
        "Completed",
        "Pending",
        "Cancelled"
    ]

    records = []

    for i in range(1, 201):

        product, category, unit_price = products[
            rng.integers(0, len(products))
        ]

        quantity = int(rng.integers(1, 8))

        unit_price = int(
            unit_price * rng.uniform(0.90, 1.12)
        )

        revenue = quantity * unit_price

        order_date = (
            pd.Timestamp("2026-01-01")
            + pd.Timedelta(days=int(rng.integers(0, 300)))
        )

        records.append({
            "Order ID": f"ORD-{i:04d}",
            "Order Date": order_date,
            "Customer": rng.choice(customers),
            "Product": product,
            "Category": category,
            "Region": rng.choice(regions),
            "Quantity": quantity,
            "Unit Price (DZD)": unit_price,
            "Revenue (DZD)": revenue,
            "Payment Method": rng.choice(payment_methods),
            "Status": rng.choice(statuses)
        })

    sales_df = pd.DataFrame(records)

    # Add a few missing values for cleaning demonstrations
    missing_rows = rng.choice(
        sales_df.index,
        size=8,
        replace=False
    )

    sales_df.loc[missing_rows[:3], "Customer"] = None
    sales_df.loc[missing_rows[3:6], "Region"] = None
    sales_df.loc[missing_rows[6:], "Quantity"] = np.nan

    # Add duplicate records to test duplicate removal
    sales_df = pd.concat(
        [sales_df, sales_df.iloc[[4, 15, 28]]],
        ignore_index=True
    )

    products_df = pd.DataFrame(
        products,
        columns=["Product", "Category", "Base Price (DZD)"]
    )

    regional_df = (
        sales_df.groupby("Region", dropna=False)
        .agg(
            Total_Orders=("Order ID", "count"),
            Total_Revenue=("Revenue (DZD)", "sum"),
            Total_Quantity=("Quantity", "sum")
        )
        .reset_index()
    )

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl"
    ) as writer:

        sales_df.to_excel(
            writer,
            sheet_name="Sales Data",
            index=False
        )

        products_df.to_excel(
            writer,
            sheet_name="Products",
            index=False
        )

        regional_df.to_excel(
            writer,
            sheet_name="Regional Summary",
            index=False
        )

    print("Sample Excel file created successfully!")
    print(f"Location: {output_path}")
    print(f"Sales rows: {len(sales_df)}")
    print(f"Products rows: {len(products_df)}")
    print(f"Regional summary rows: {len(regional_df)}")


if __name__ == "__main__":
    generate_sample_excel()
