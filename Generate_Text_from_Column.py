import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import pandas as pd

def main():
    # Initialize Tkinter and hide the main window.
    root = tk.Tk()
    root.withdraw()

    # Ask the user to choose the Excel file.
    input_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    if not input_file:
        messagebox.showerror("Error", "No file selected.")
        return

    # Ask which column to process (by header name or column letter).
    col_input = simpledialog.askstring(
        "Input",
        "Enter the column name or column letter (e.g., 'Interventions' or 'A') to process:"
    )
    if not col_input:
        messagebox.showerror("Error", "No column specified.")
        return

    # Load the Excel file into a DataFrame.
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read the Excel file:\n{e}")
        return

    # Determine the column based on user input.
    if col_input.isalpha() and len(col_input) == 1:
        col_index = ord(col_input.upper()) - ord('A')
        if col_index < len(df.columns):
            col_name = df.columns[col_index]
        else:
            messagebox.showerror("Error", "Column letter is out of range.")
            return
    else:
        col_name = col_input
        if col_name not in df.columns:
            messagebox.showerror("Error", "Specified column not found in Excel file.")
            return

    # Concatenate all non-empty cell contents from the selected column.
    items = []
    for cell in df[col_name]:
        if pd.isnull(cell):
            continue
        items.append(str(cell).strip())

    paragraph = ", ".join(items)

    # Ask the user where to save the output text file.
    output_file = filedialog.asksaveasfilename(
        title="Save Output Text File",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")]
    )
    if not output_file:
        messagebox.showerror("Error", "No output file selected.")
        return

    # Save the paragraph to the text file.
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(paragraph)
        messagebox.showinfo("Success", f"Output successfully saved to:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save output file:\n{e}")

if __name__ == "__main__":
    main()
