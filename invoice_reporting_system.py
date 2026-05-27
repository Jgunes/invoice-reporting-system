from pathlib import Path
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF

INPUT_FILE = "business_data.xlsx"
OUTPUT_DIR = Path("output")
REPORT_DATE = datetime.today().strftime("%Y-%m-%d")


def read_business_data(file_path: str):
    invoices = pd.read_excel(file_path, sheet_name="Invoices")
    sales = pd.read_excel(file_path, sheet_name="Sales")
    customers = pd.read_excel(file_path, sheet_name="Customers")
    products = pd.read_excel(file_path, sheet_name="Products")

    return invoices, sales, customers, products


def clean_data(invoices: pd.DataFrame, sales: pd.DataFrame):
    invoices = invoices.copy()
    sales = sales.copy()

    invoices["InvoiceDate"] = pd.to_datetime(invoices["InvoiceDate"], errors="coerce")
    invoices["DueDate"] = pd.to_datetime(invoices["DueDate"], errors="coerce")
    invoices["PaymentDate"] = pd.to_datetime(invoices["PaymentDate"], errors="coerce")
    sales["OrderDate"] = pd.to_datetime(sales["OrderDate"], errors="coerce")

    invoice_numeric = ["Quantity", "UnitPrice", "Subtotal", "VATAmount", "TotalAmount"]
    sales_numeric = ["Quantity", "UnitPrice", "Revenue", "Cost", "Profit"]

    for col in invoice_numeric:
        invoices[col] = pd.to_numeric(invoices[col], errors="coerce").fillna(0)

    for col in sales_numeric:
        sales[col] = pd.to_numeric(sales[col], errors="coerce").fillna(0)

    return invoices, sales


def analyze_business(invoices: pd.DataFrame, sales: pd.DataFrame):
    total_revenue = sales["Revenue"].sum()
    total_profit = sales["Profit"].sum()

    category_performance = (
        sales.groupby("Category", as_index=False)[["Revenue", "Profit"]]
        .sum()
        .sort_values("Revenue", ascending=False)
    )

    monthly_sales = (
        sales.assign(Month=sales["OrderDate"].dt.to_period("M").astype(str))
        .groupby("Month", as_index=False)[["Revenue", "Profit"]]
        .sum()
    )

    payment_status = (
        invoices.groupby("PaymentStatus", as_index=False)["TotalAmount"]
        .sum()
        .sort_values("TotalAmount", ascending=False)
    )

    summary = {
        "Report Date": REPORT_DATE,
        "Total Revenue": round(total_revenue, 2),
        "Total Profit": round(total_profit, 2),
        "Invoice Count": int(len(invoices)),
        "Sales Count": int(len(sales)),
    }

    return {
        "summary": summary,
        "category_performance": category_performance,
        "monthly_sales": monthly_sales,
        "payment_status": payment_status,
    }


def create_neon_charts(analysis: dict):
    OUTPUT_DIR.mkdir(exist_ok=True)

    plt.style.use("dark_background")

    category_data = analysis["category_performance"]

    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(
        category_data["Category"],
        category_data["Revenue"],
        width=0.38,
        color=["#00F5FF", "#FF00FF", "#39FF14", "#FFD700", "#FF4500"]
    )

    ax.set_title(
        "Revenue by Category",
        fontsize=22,
        color="#00F5FF",
        weight="bold"
    )

    ax.set_xlabel("Category", fontsize=12, color="white")
    ax.set_ylabel("Revenue", fontsize=12, color="white")

    ax.grid(color="#333333", linestyle="--", alpha=0.4)

    for bar in bars:
        yval = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            yval,
            f"{yval:,.0f}",
            ha="center",
            va="bottom",
            color="white",
            fontsize=10
        )

    chart_path = OUTPUT_DIR / "revenue_chart.png"

    plt.tight_layout()
    plt.savefig(chart_path, dpi=300, bbox_inches="tight")
    plt.close()

    return chart_path


def export_neon_pdf(analysis: dict, chart_path):
    pdf_path = OUTPUT_DIR / "neon_business_report.pdf"

    pdf = FPDF()
    pdf.add_page()

    pdf.set_fill_color(5, 5, 20)
    pdf.rect(0, 0, 210, 297, "F")

    pdf.set_text_color(0, 255, 255)
    pdf.set_font("Arial", "B", 24)
    pdf.cell(0, 18, "INVOICE REPORT SYSTEM", ln=True, align="C")

    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", "", 14)
    pdf.cell(0, 10, "Automated Business Dashboard", ln=True, align="C")

    stats = [
        ("Revenue", f"${analysis['summary']['Total Revenue']:,.2f}", (0,255,255)),
        ("Profit", f"${analysis['summary']['Total Profit']:,.2f}", (255,0,255)),
        ("Invoices", str(analysis['summary']['Invoice Count']), (57,255,20)),
        ("Sales", str(analysis['summary']['Sales Count']), (255,215,0)),
    ]

    x_positions = [10, 60, 110, 160]

    for (title, value, color), x in zip(stats, x_positions):
        pdf.set_draw_color(*color)
        pdf.set_line_width(1.5)
        pdf.rect(x, 40, 40, 28)

        pdf.set_xy(x, 46)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", "B", 10)
        pdf.cell(40, 5, title, align="C")

        pdf.set_xy(x, 55)
        pdf.set_text_color(*color)
        pdf.set_font("Arial", "B", 14)
        pdf.cell(40, 5, value, align="C")

    pdf.set_xy(15, 80)
    pdf.image(str(chart_path), w=180)

    pdf.output(str(pdf_path))

    return pdf_path


def main():
    print("Starting Neon Invoice Reporting System...")

    invoices, sales, customers, products = read_business_data(INPUT_FILE)

    invoices, sales = clean_data(invoices, sales)

    analysis = analyze_business(invoices, sales)

    chart_path = create_neon_charts(analysis)

    pdf_path = export_neon_pdf(analysis, chart_path)

    print("Workflow completed successfully!")
    print(f"Chart created: {chart_path}")
    print(f"PDF created: {pdf_path}")


if __name__ == "__main__":
    main()