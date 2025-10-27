import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import os

# -------- FILE SELECTION DIALOG --------
root = tk.Tk()
root.withdraw()  # Hide the root window
file_path = filedialog.askopenfilename(
    title="Select Excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not file_path:
    print("No file selected. Exiting.")
    exit()

# -------- CONFIG --------
email_column = "Primary Email"
office_column = "Office"

# -------- READ EXCEL --------
df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()

# -------- CASE-INSENSITIVE COLUMN MATCH --------
lower_cols = {col.lower(): col for col in df.columns}
required_cols = [email_column.lower(), office_column.lower()]
missing_cols = [col for col in required_cols if col not in lower_cols]

if missing_cols:
    print(f"Error: Missing required column(s): {', '.join(missing_cols)}")
    exit()

# Map original case column names
email_col_actual = lower_cols[email_column.lower()]
office_col_actual = lower_cols[office_column.lower()]

# -------- FILTER (CASE-INSENSITIVE CONTENT MATCH) --------
filtered_df = df[
    df[email_col_actual].astype(str).str.lower().str.contains("@gil-bar.com", na=False) |
    df[office_col_actual].astype(str).str.lower().str.contains("gbi", na=False)
]

# -------- WRITE TO NEW FILE --------
base, ext = os.path.splitext(file_path)
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
new_file_path = f"{base}_filtered_{timestamp}{ext}"

filtered_df.to_excel(new_file_path, sheet_name="Filtered", index=False)

print(f"Filtered data written to new file: {new_file_path}")
