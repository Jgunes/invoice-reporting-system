from pathlib import Path
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF


INPUT_FILE = "business_data.xlsx"
OUTPUT_DIR = Path("output")
REPORT_DATE = datetime.today().strftime("%Y-%m-%d")


def read_business_data(file_path: str):
    """Read all required sheets from the Excel workbook."""
    invoices = pd.read_excel(file_path, sheet_name="Invoices")
    sales = pd.read_excel(file_path, sheet_name="Sales")
    customers = pd.read_excel(file_path, sheet_name="Customers")
    products = pd.read_excel(file_path, sheet_name="Products")
    return invoices, sales, customers, products



def clean_data(invoices: pd.DataFrame, sales: pd.DataFrame):
    """Clean dates, text columns, numeric columns, and missing values."""
    invoices = invoices.copy()
    sales = sales.copy()

    invoices["InvoiceDate"] = pd.to_datetime(invoices["InvoiceDate"], errors="coerce")
    invoices["DueDate"] = pd.to_datetime(invoices["DueDate"], errors="coerce")
    invoices["PaymentDate"] = pd.to_datetime(invoices["PaymentDate"], errors="coerce")
    sales["OrderDate"] = pd.to_datetime(sales["OrderDate"], errors="coerce")

    text_columns_invoices = ["InvoiceID", "CustomerID", "CustomerName", "ProductID", "ProductName", "PaymentStatus"]
    for col in text_columns_invoices:
        invoices[col] = invoices[col].astype(str).str.strip()

    text_columns_sales = ["OrderID", "CustomerID", "CustomerName", "ProductID", "ProductName", "Category", "SalesChannel", "Region", "SalesRep"]
    for col in text_columns_sales:
        sales[col] = sales[col].astype(str).str.strip()

    invoice_numeric_cols = ["Quantity", "UnitPrice", "Subtotal", "VATRate", "VATAmount", "TotalAmount"]
    for col in invoice_numeric_cols:
        invoices[col] = pd.to_numeric(invoices[col], errors="coerce")

    sales_numeric_cols = ["Quantity", "UnitPrice", "Revenue", "Cost", "Profit"]
    for col in sales_numeric_cols:
        sales[col] = pd.to_numeric(sales[col], errors="coerce")

    invoices[invoice_numeric_cols] = invoices[invoice_numeric_cols].fillna(0)
    sales[sales_numeric_cols] = sales[sales_numeric_cols].fillna(0)

    return invoices, sales



def analyze_business(invoices: pd.DataFrame, sales: pd.DataFrame):
    """Create useful business insights from invoices and sales records."""
    today = pd.Timestamp(datetime.today().date())

    total_revenue = sales["Revenue"].sum()
    total_profit = sales["Profit"].sum()
    total_invoiced = invoices["TotalAmount"].sum()
    unpaid_amount = invoices.loc[invoices["PaymentStatus"].isin(["Unpaid", "Overdue"]), "TotalAmount"].sum()
    overdue_amount = invoices.loc[invoices["PaymentStatus"] == "Overdue", "TotalAmount"].sum()

    duplicate_invoices = invoices[invoices.duplicated(subset=["InvoiceID"], keep=False)].sort_values("InvoiceID")

    past_due_unpaid = invoices[(invoices["DueDate"] < today) & (invoices["PaymentStatus"] != "Paid")]

    top_customers = (
        sales.groupby("CustomerName", as_index=False)["Revenue"]
        .sum()
        .sort_values("Revenue", ascending=False)
        .head(10)
    )

    top_products = (
        sales.groupby("ProductName", as_index=False)[["Revenue", "Profit", "Quantity"]]
        .sum()
        .sort_values("Revenue", ascending=False)
        .head(10)
    )

    category_performance = (
        sales.groupby("Category", as_index=False)[["Revenue", "Profit", "Quantity"]]
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
        "Total Invoiced": round(total_invoiced, 2),
        "Unpaid Amount": round(unpaid_amount, 2),
        "Overdue Amount": round(overdue_amount, 2),
        "Duplicate Invoice Count": duplicate_invoices["InvoiceID"].nunique(),
        "Past Due Unpaid Count": len(past_due_unpaid),
    }

    analysis = {
        "summary": summary,
        "top_customers": top_customers,
        "top_products": top_products,
        "category_performance": category_performance,
        "monthly_sales": monthly_sales,
        "payment_status": payment_status,
        "duplicate_invoices": duplicate_invoices,
        "past_due_unpaid": past_due_unpaid,
    }

    return analysis



def create_charts(analysis: dict):
    """Generate business charts and save them as PNG files."""
    OUTPUT_DIR.mkdir(exist_ok=True)

    
    category_data = analysis["category_performance"]
    plt.figure(figsize=(10, 5))
    plt.bar(category_data["Category"], category_data["Revenue"])
    plt.title("Revenue by Product Category")
    plt.xlabel("Category")
    plt.ylabel("Revenue")
    plt.xticks(rotation=35, ha="right")
    plt.tight_layout()
    revenue_chart_path = OUTPUT_DIR / "revenue_by_category.png"
    plt.savefig(revenue_chart_path)
    plt.close()

    
    monthly_data = analysis["monthly_sales"]
    plt.figure(figsize=(10, 5))
    plt.plot(monthly_data["Month"], monthly_data["Revenue"], marker="o", label="Revenue")
    plt.plot(monthly_data["Month"], monthly_data["Profit"], marker="o", label="Profit")
    plt.title("Monthly Revenue and Profit")
    plt.xlabel("Month")
    plt.ylabel("Amount")
    plt.legend()
    plt.xticks(rotation=35, ha="right")
    plt.tight_layout()
    monthly_chart_path = OUTPUT_DIR / "monthly_revenue_profit.png"
    plt.savefig(monthly_chart_path)
    plt.close()

    
    payment_data = analysis["payment_status"]
    plt.figure(figsize=(8, 5))
    plt.bar(payment_data["PaymentStatus"], payment_data["TotalAmount"])
    plt.title("Invoice Amount by Payment Status")
    plt.xlabel("Payment Status")
    plt.ylabel("Total Amount")
    plt.tight_layout()
    payment_chart_path = OUTPUT_DIR / "payment_status.png"
    plt.savefig(payment_chart_path)
    plt.close()

    return revenue_chart_path, monthly_chart_path, payment_chart_path




def export_excel_report(invoices: pd.DataFrame, sales: pd.DataFrame, analysis: dict):
    """Export cleaned data and analysis tables into one Excel report."""
    excel_output = OUTPUT_DIR / "automated_business_report.xlsx"

    summary_df = pd.DataFrame(list(analysis["summary"].items()), columns=["Metric", "Value"])

    with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Dashboard", index=False)
        analysis["top_customers"].to_excel(writer, sheet_name="Top Customers", index=False)
        analysis["top_products"].to_excel(writer, sheet_name="Top Products", index=False)
        analysis["category_performance"].to_excel(writer, sheet_name="Category Performance", index=False)
        analysis["monthly_sales"].to_excel(writer, sheet_name="Monthly Sales", index=False)
        analysis["payment_status"].to_excel(writer, sheet_name="Payment Status", index=False)
        analysis["duplicate_invoices"].to_excel(writer, sheet_name="Duplicate Invoices", index=False)
        analysis["past_due_unpaid"].to_excel(writer, sheet_name="Past Due Unpaid", index=False)
        invoices.to_excel(writer, sheet_name="Cleaned Invoices", index=False)
        sales.to_excel(writer, sheet_name="Cleaned Sales", index=False)

    return excel_output


def export_pdf_report(analysis: dict, chart_paths):
    """Create a PDF report with summary metrics and charts."""
    pdf_output = OUTPUT_DIR / "automated_business_report.pdf"

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Arial", "B", 18)
    pdf.cell(0, 10, "Automated Invoice & Sales Report", ln=True)

    pdf.set_font("Arial", "", 11)
    pdf.cell(0, 8, f"Report date: {REPORT_DATE}", ln=True)
    pdf.ln(5)

    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 8, "Executive Summary", ln=True)
    pdf.set_font("Arial", "", 11)

    for metric, value in analysis["summary"].items():
        pdf.cell(0, 7, f"{metric}: {value}", ln=True)

    pdf.ln(4)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 8, "Business Insights", ln=True)
    pdf.set_font("Arial", "", 11)

    top_customer = analysis["top_customers"].iloc[0]
    top_product = analysis["top_products"].iloc[0]
    pdf.multi_cell(0, 7, f"Top customer by revenue: {top_customer['CustomerName']} ({top_customer['Revenue']:.2f})")
    pdf.multi_cell(0, 7, f"Top product by revenue: {top_product['ProductName']} ({top_product['Revenue']:.2f})")

    if analysis["summary"]["Duplicate Invoice Count"] > 0:
        pdf.multi_cell(0, 7, "Warning: Duplicate invoice numbers were detected. These should be reviewed before accounting approval.")

    if analysis["summary"]["Overdue Amount"] > 0:
        pdf.multi_cell(0, 7, "Recommendation: Follow up on overdue invoices to improve cash flow.")

    
    for chart_path in chart_paths:
        pdf.add_page()
        pdf.set_font("Arial", "B", 14)
        chart_title = chart_path.stem.replace("_", " ").title()
        pdf.cell(0, 10, chart_title, ln=True)
        pdf.image(str(chart_path), x=15, y=30, w=180)

    pdf.output(str(pdf_output))
    return pdf_output




def main():
    print("Starting automated invoice and sales reporting workflow...")

    invoices, sales, customers, products = read_business_data(INPUT_FILE)
    invoices, sales = clean_data(invoices, sales)
    analysis = analyze_business(invoices, sales)

    chart_paths = create_charts(analysis)
    excel_report = export_excel_report(invoices, sales, analysis)
    pdf_report = export_pdf_report(analysis, chart_paths)

    print("Workflow completed successfully.")
    print(f"Excel report created: {excel_report}")
    print(f"PDF report created: {pdf_report}")
    print("Charts created:")
    for chart in chart_paths:
        print(f"- {chart}")


if __name__ == "__main__":
    main()
