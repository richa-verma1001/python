import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog
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
df.columns = df.columns.str.strip()  # remove spaces

# Convert columns to lowercase for case-insensitive matching
lower_cols = {col.lower(): col for col in df.columns}

# -------- VALIDATE COLUMNS (case-insensitive) --------
missing_cols = [col for col in [email_column.lower(), office_column.lower()] if col not in lower_cols]
if missing_cols:
    print(f"Error: Missing required column(s): {', '.join(missing_cols)}")
    exit()

# Map actual column names
email_col_actual = lower_cols[email_column.lower()]
office_col_actual = lower_cols[office_column.lower()]
email_match = simpledialog.askstring("Primary Email Key", "Enter the email match")
office_match = simpledialog.askstring("Office Key", "Enter the office match")

# -------- FILTER (case-insensitive) --------
filtered_df = df[
    df[email_col_actual].astype(str).str.lower().str.contains(email_match, na=False) |
    df[office_col_actual].astype(str).str.lower().str.contains(office_match, na=False)
]

# -------- WRITE TO NEW FILE --------
base, ext = os.path.splitext(file_path)
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
new_file_path = f"{base}_filtered_{timestamp}{ext}"

# Write both sheets — original + filtered
with pd.ExcelWriter(new_file_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Original Data", index=False)
    filtered_df.to_excel(writer, sheet_name="Filtered "+ office_match, index=False)

print(f"✅ Original and filtered data written to new file: {new_file_path}")
