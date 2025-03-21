import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import pandas as pd

def main():
    # Set up Tkinter root and hide main window.
    root = tk.Tk()
    root.withdraw()

    # Ask user to choose the Excel file
    input_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    if not input_file:
        messagebox.showerror("Error", "No file selected.")
        return

    # Ask which column to process (by header name or column letter)
    col_input = simpledialog.askstring(
        "Input",
        "Enter the column name or column letter (e.g., 'Interventions' or 'A') to process:"
    )
    if not col_input:
        messagebox.showerror("Error", "No column specified.")
        return

    # Load the Excel file into a DataFrame
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read the Excel file:\n{e}")
        return

    # Determine the correct column based on user input
    if col_input.isalpha() and len(col_input) == 1:
        # Convert a single letter (e.g., 'A') to a column index
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

    # Process the selected column:
    # - Split interventions separated by commas
    # - Remove extra whitespace
    # - Count each unique intervention per study (row)
    counts = {}
    for cell in df[col_name]:
        if pd.isnull(cell):
            continue
        # Split the cell by commas and trim whitespace
        interventions = [item.strip() for item in str(cell).split(',')]
        # Remove any empty strings
        interventions = [i for i in interventions if i]
        # Use a set to count each intervention only once per study
        for intervention in set(interventions):
            counts[intervention] = counts.get(intervention, 0) + 1

    # Sort the interventions in descending order of frequency
    sorted_interventions = sorted(counts.items(), key=lambda x: x[1], reverse=True)
    
    # Create a DataFrame with the results
    output_df = pd.DataFrame(sorted_interventions, columns=["Intervention", "Count"])

    # Ask the user where to save the output Excel file
    output_file = filedialog.asksaveasfilename(
        title="Save Output Excel File",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not output_file:
        messagebox.showerror("Error", "No output file selected.")
        return

    # Save the DataFrame to an Excel file
    try:
        output_df.to_excel(output_file, index=False)
        messagebox.showinfo("Success", f"Output successfully saved to:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save output file:\n{e}")

if __name__ == "__main__":
    main()
