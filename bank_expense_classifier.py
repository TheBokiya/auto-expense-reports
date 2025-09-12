import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, NamedStyle
import json
import os

# === LOAD STATIC VALUES ===
STATIC_PATH = os.path.join(os.path.dirname(__file__), "static_values.json")
with open(STATIC_PATH, "r") as f:
    STATIC = json.load(f)

# === CLASSIFICATION & SUMMARIZATION ===
def classify_fs_line(details):
    details_lower = details.lower()
    if any(keyword in details_lower for keyword in STATIC["salary_keywords"]):
        return "Salary"
    elif any(keyword in details_lower for keyword in STATIC["utility_keywords"]):
        return "Utility"
    elif any(keyword in details_lower for keyword in STATIC["rent_keywords"]):
        return "Rent"
    elif any(keyword in details_lower for keyword in STATIC["ingredient_keywords"]):
        return "Ingredient"
    else:
        return "Other"

def summarize_item(details):
    details_lower = details.lower()
    if any(keyword in details_lower for keyword in STATIC["utility_keywords"]):
        return "Utilities"
    elif any(keyword in details_lower for keyword in STATIC["rent_keywords"]):
        return "Rent"
    elif any(name.lower() in details_lower for name in STATIC["names"]):
        matched_name = next(name for name in STATIC["names"] if name.lower() in details_lower)
        return f"Salary - {matched_name}"
    elif any(keyword in details_lower for keyword in STATIC["milk_keywords"]):
        return "Milk Purchase"
    elif any(keyword in details_lower for keyword in STATIC["coffee_keywords"]):
        return "Coffee Purchase"
    else:
        words = details.split()
        return " ".join(words[:5]) if words else "Unknown"

# === FILE PROCESSING ===
def process_file(filepath, currency):
    df = pd.read_excel(filepath, header=None)
    for i in range(10):
        if df.iloc[i].str.contains("Date", case=False, na=False).any():
            header_row = i
            break
    else:
        raise ValueError("Could not find header row with 'Date'")

    df = pd.read_excel(filepath, header=header_row)
    df = df.rename(columns=lambda x: str(x).strip().lower())

    date_col = next(col for col in df.columns if STATIC["column_names"]["date"] in col)
    details_col = next(col for col in df.columns if STATIC["column_names"]["details"] in col)
    money_out_col = next(col for col in df.columns if STATIC["column_names"]["money_out"] in col)

    # Filter out rows where "Money Out" is empty or zero
    df = df[df[money_out_col].notna() & (df[money_out_col] != 0)]

    output_rows = []
    for _, row in df.iterrows():
        try:
            date = pd.to_datetime(row[date_col])
            datecode = f"{date.month}{date.year}"
            details = str(row[details_col])
            fs_line = classify_fs_line(details)
            item = summarize_item(details)
            expense_usd = row[money_out_col] if currency == "USD" else 0
            expense_khr = row[money_out_col] if currency == "KHR" else 0

            output_rows.append({
                "Datecode": datecode,
                "Date": date.strftime("%Y-%m-%d"),
                "FS Line": fs_line,
                "Item": item,
                "Expense USD": expense_usd,
                "Expense KHR": expense_khr
            })
        except Exception as e:
            print(f"Error processing row: {e}")

    return pd.DataFrame(output_rows)

# === GUI ===
def run_gui():
    def load_khr():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            khr_path.set(path)

    def load_usd():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            usd_path.set(path)

    def generate_report():
        try:
            df_khr = process_file(khr_path.get(), "KHR") if khr_path.get() else pd.DataFrame()
            df_usd = process_file(usd_path.get(), "USD") if usd_path.get() else pd.DataFrame()
            final_df = pd.concat([df_khr, df_usd], ignore_index=True)

            final_df["Date"] = pd.to_datetime(final_df["Date"])
            final_df = final_df.sort_values(by="Date")
            final_df["Date"] = final_df["Date"].dt.strftime("%Y-%m-%d")

            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                # Save the DataFrame to Excel
                final_df.to_excel(save_path, index=False)

                # Apply accounting number format using openpyxl
                wb = load_workbook(save_path)
                ws = wb.active

                # Define accounting style
                accounting_style = NamedStyle(name="AccountingStyle", number_format="#,##0.00_);(#,##0.00)")

                # Apply the style to "Expense USD" and "Expense KHR" columns
                usd_col = ws["E"]  # Column E is "Expense USD"
                khr_col = ws["F"]  # Column F is "Expense KHR"

                for cell in usd_col[1:]:  # Skip the header row
                    cell.style = accounting_style

                for cell in khr_col[1:]:  # Skip the header row
                    cell.style = accounting_style

                # Save the workbook with formatting
                wb.save(save_path)

                messagebox.showinfo("Success", f"Expense report saved to:\n{save_path}")
            else:
                messagebox.showinfo("Cancelled", "Report generation cancelled.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    root = tk.Tk()
    root.title("Expense Report Generator")

    khr_path = tk.StringVar()
    usd_path = tk.StringVar()

    tk.Label(root, text="KHR Statement:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    tk.Entry(root, textvariable=khr_path, width=50).grid(row=0, column=1)
    tk.Button(root, text="Browse", command=load_khr).grid(row=0, column=2)

    tk.Label(root, text="USD Statement:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    tk.Entry(root, textvariable=usd_path, width=50).grid(row=1, column=1)
    tk.Button(root, text="Browse", command=load_usd).grid(row=1, column=2)

    tk.Button(root, text="Generate Report", command=generate_report, bg="green", fg="white").grid(row=2, column=1, pady=20)

    root.mainloop()

# === RUN ===
if __name__ == "__main__":
    run_gui()
