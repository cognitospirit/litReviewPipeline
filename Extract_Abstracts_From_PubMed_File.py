import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from Bio import Medline

# Function to process the MEDLINE file and save to Excel
def process_pubmed_file(file_path):
    # Open the MEDLINE file with UTF-8 encoding
    try:
        with open(file_path, 'r', encoding='utf-8') as handle:
            records = Medline.parse(handle)
            records = list(records)  # Convert to list for easy iteration
    except UnicodeDecodeError:
        messagebox.showerror("Error", "Unable to read the file. Please make sure it's in UTF-8 format.")
        return
    
    # Create lists to hold the extracted data
    titles = []
    abstracts = []
    authors = []
    years = []

    # Loop through the records and extract the necessary fields
    for record in records:
        title = record.get('TI', 'No title available')
        abstract = record.get('AB', 'No abstract available')
        author_list = record.get('AU', 'No authors available')
        year = record.get('DP', 'No date available')
        
        # Handle multiple authors by joining them with commas
        authors_str = ', '.join(author_list) if isinstance(author_list, list) else 'No authors available'
        
        # Append to the lists
        titles.append(title)
        abstracts.append(abstract)
        authors.append(authors_str)
        years.append(year)
    
    # Create a Pandas DataFrame
    df = pd.DataFrame({
        'Title': titles,
        'Authors': authors,
        'Year': years,
        'Abstract': abstracts
    })
    
    # Ask the user where to save the Excel file
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    
    if save_path:
        # Write the DataFrame to an Excel file
        df.to_excel(save_path, index=False)
        messagebox.showinfo("Success", f"File saved successfully at {save_path}")
    else:
        messagebox.showwarning("Cancelled", "File save cancelled.")

# Function to open a file dialog and select the PubMed file
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("PubMed files", "*.txt *.medline")])
    if file_path:
        process_pubmed_file(file_path)
    else:
        messagebox.showwarning("Cancelled", "File selection cancelled.")

# Create the main window
root = tk.Tk()
root.title("PubMed to Excel Converter")

# Set the window size
root.geometry("400x200")

# Add a label and button to select the file
label = tk.Label(root, text="Select a PubMed File in MEDLINE Format", font=("Arial", 14))
label.pack(pady=20)

# Correct button definition
button = tk.Button(root, text="Select File", command=select_file, font=("Arial", 12))
button.pack(pady=10)

# Run the application
root.mainloop()
