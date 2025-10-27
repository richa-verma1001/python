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
email_column = "Primary Email"  # updated column name
office_column = "Office"

# -------- READ EXCEL --------
df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()  # remove leading/trailing spaces

# -------- VALIDATE COLUMNS --------
missing_cols = [col for col in [email_column, office_column] if col not in df.columns]
if missing_cols:
    print(f"Error: Missing required column(s): {', '.join(missing_cols)}")
    exit()

# -------- FILTER --------
filtered_df = df[df[email_column].str.contains("@gil-bar.com", na=False) |
                 df[office_column].str.contains("GBI", na=False)]

# -------- WRITE TO NEW FILE --------
base, ext = os.path.splitext(file_path)
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
new_file_path = f"{base}_filtered_{timestamp}{ext}"

filtered_df.to_excel(new_file_path, sheet_name="Filtered", index=False)

print(f"Filtered data written to new file: {new_file_path}")
