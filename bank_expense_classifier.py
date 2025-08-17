import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

# === CLASSIFICATION & SUMMARIZATION ===
def classify_fs_line(details):
    details_lower = details.lower()
    if any(name.lower() in details_lower for name in ["heng alisa", "morn monita", "luy solay"]):
        return "Salary"
    elif "utilities bill" in details_lower:
        return "Utility"
    elif "rental fee" in details_lower:
        return "Rent"
    elif any(keyword in details_lower for keyword in ["meiji", "gbs", "kirisu", "milk"]):
        return "Ingredient"
    else:
        return "Other"

def summarize_item(details):
    details_lower = details.lower()
    if "utilities bill" in details_lower:
        return "Utilities"
    elif "rental fee" in details_lower:
        return "Rent"
    elif any(name.lower() in details_lower for name in ["heng alisa", "morn monita", "luy solay"]):
        matched_name = next(name for name in ["Heng Alisa", "Morn Monita", "Luy Solay"] if name.lower() in details_lower)
        return f"Salary - {matched_name}"
    elif any(keyword in details_lower for keyword in ["meiji", "kirisu", "milk"]):
        return "Milk Purchase"
    elif "gbs" in details_lower:
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

    date_col = next(col for col in df.columns if "date" in col)
    details_col = next(col for col in df.columns if "detail" in col)
    money_out_col = next(col for col in df.columns if "money out" in col)

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
                final_df.to_excel(save_path, index=False)
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
