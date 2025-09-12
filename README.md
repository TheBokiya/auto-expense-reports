# 🧾 Expense Report Generator (Local Setup)

This is a personal automation tool to generate categorized expense reports from CSV, Excel, or PDF files.

---

## ⚙️ Setup Instructions

### 1. Clone the Repo (or navigate to your script folder)

```bash
git clone https://github.com/yourusername/expense-report-generator.git
cd expense-report-generator
```

### 2. Create a Virtual Environment (optional but recommended)

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Build a macOS App (Optional)

To package the script as a standalone macOS app, follow these steps:

1. Install PyInstaller:

   ```bash
   pip install pyinstaller
   ```

2. Run the following command to build the app:

   ```bash
   pyinstaller --windowed --add-data "static_values.json:." --name "Expense Report Generator" bank_expense_classifier.py
   ```

   - `--windowed`: Ensures the app runs without a terminal window.
   - `--add-data "static_values.json:."`: Includes the `static_values.json` file in the app bundle.
   - `--name "Expense Report Generator"`: Sets the name of the app.

3. After the build process, the app will be located in the `dist` directory as `Expense Report Generator.app`.

4. Double-click the `.app` file to run the application.

---

### Install Dependencies

```bash
pip install pandas openpyxl PyPDF2
```

## ▶️ Running the Script

### Basic Usage

```bash
python expense_report_generator.py
```
