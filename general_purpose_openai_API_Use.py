import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import openai
import os

# Initialize the OpenAI client (API key should be set as an environment variable)
#openai.api_key = os.getenv("OPENAI_API_KEY")  # Make sure your API key is properly set in the environment
openai.api_key = ''
# Arguments to control generation
generation_args = {
    "model": "gpt-4o",  # Use the appropriate model, e.g., "gpt-4", "gpt-3.5-turbo"
}

def generate_response(text, instruction):
    try:
        start_time = time.time()  # Record start time
        prompt_text = [
            {"role": "system", "content": "You are a helpful AI assistant."},
            {"role": "user", "content": instruction},
            {"role": "user", "content": text},
        ]
        
        # Fetch the response from OpenAI API
        chat_completion = openai.ChatCompletion.create(
            model=generation_args["model"],
            messages=prompt_text
        )
        
        # Extract the bot's response text
        answer = chat_completion.choices[0].message['content']
        elapsed_time = time.time() - start_time  # Calculate elapsed time
        print(f"Time taken for response: {elapsed_time:.2f} seconds")
        return answer
    
    except Exception as e:
        return f"Error: {e}"

def select_file():
    # Function to select a CSV or Excel file and process the text
    file_path = filedialog.askopenfilename(
        filetypes=[("CSV and Excel Files", "*.csv *.xlsx *.xls")]
    )
    if not file_path:
        messagebox.showerror("Error", "No file selected")
        return

    try:
        # Prompt the user for the column name
        column_name = simpledialog.askstring("Input", "Enter the column name to process:")
        if not column_name:
            messagebox.showerror("Error", "No column name provided")
            return

        # Prompt the user for the instruction
        instruction = simpledialog.askstring("Input", "Enter the instruction for the model:")
        if not instruction:
            messagebox.showerror("Error", "No instruction provided")
            return

        # Determine the file extension
        ext = os.path.splitext(file_path)[1].lower()

        # Load the file based on the extension
        if ext == '.csv':
            df = pd.read_csv(file_path)
        elif ext in ['.xls', '.xlsx']:
            df = pd.read_excel(file_path)
        else:
            messagebox.showerror("Error", "Unsupported file type")
            return

        if column_name not in df.columns:
            messagebox.showerror("Error", f"'{column_name}' column not found in the file")
            return

        results = []

        for i, text in enumerate(df[column_name], start=1):
            # Ensure the text is a valid string
            if pd.isna(text) or not isinstance(text, str):
                print(f"Skipping row {i}: Invalid data")
                results.append("")  # Append empty string for invalid data
                continue  # Skip if the text is NaN or not a string

            print(f"Processing row {i}/{len(df)}")
            result = generate_response(str(text), instruction)  # Convert text to string
            results.append(result)

        # Add the results to the DataFrame
        df['Result'] = results

        # Prompt the user to save the output file
        output_file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save the output file"
        )
        if output_file_path:
            # Save the DataFrame in Excel format
            df.to_excel(output_file_path, index=False)
            messagebox.showinfo("Success", f"Results have been exported to {output_file_path} successfully!")
        else:
            messagebox.showerror("Error", "No output file selected")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def main():
    # Create the GUI window
    root = tk.Tk()
    root.title("Text Processor")

    # Set up the window size and button
    root.geometry("400x200")
    button = tk.Button(root, text="Select File and Process Text", command=select_file)
    button.pack(expand=True)

    # Run the GUI loop
    root.mainloop()

if __name__ == "__main__":
    main()
