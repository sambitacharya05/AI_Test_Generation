import pandas as pd
import os

def excel_to_markdown(source_file: str, output_file: str):
    """
    Converts an Excel (.xlsx) file into a Markdown (.md) table with evenly spaced columns.

    This function reads data from an Excel file, processes the data to account for mixed-width 
    characters (such as those in Japanese and English), and writes the data into a Markdown file 
    formatted as a table with even spacing between columns and rows.

    Parameters:
    ----------
    source_file : str
        The file path to the source Excel (.xlsx) file. This file should contain the data you 
        want to export to Markdown.
    
    output_file : str
        The file path to the output Markdown (.md) file. If the file does not exist, it will be 
        created. If it does exist, its content will be overwritten with the new Markdown table.

    Functionality:
    -------------
    - The function reads the data from the specified Excel file using the pandas library.
    - It calculates the maximum display width for each column, accounting for mixed-width characters 
      (e.g., Japanese characters and English characters) to ensure proper alignment.
    - The function then formats each row, including the header, so that all columns are evenly spaced.
    - The formatted data is written into the specified Markdown file as a table.

    Notes:
    -----
    - The function handles text in both Japanese and English, ensuring that the output Markdown table 
      appears well-aligned, even with mixed character sets.
    - Each column's width is determined based on the widest entry in that column, ensuring consistent 
      spacing in the Markdown table.

    Example Usage:
    -------------
    Assuming you have an Excel file named `data.xlsx` and you want to convert it to `table.md`:

    ```python
    excel_to_markdown('data.xlsx', 'table.md')
    ```

    After running the above command, `table.md` will contain a Markdown table representation of the 
    data from `data.xlsx`, with properly aligned columns and rows.
    """
    # Check if the output file exists
    if not os.path.exists(output_file):
        # If it doesn't exist, create it
        with open(output_file, 'w', encoding='utf-8') as md_file:
            pass
    
    # Read the Excel file
    df = pd.read_excel(source_file)
    
    # Function to calculate the display width of a string, accounting for mixed-width characters
    def calc_display_width(text):
        return sum(2 if ord(char) > 255 else 1 for char in str(text))
    
    # Calculate the maximum width for each column
    col_widths = [max(calc_display_width(item) for item in df[col].tolist() + [col]) for col in df.columns]

    # Format each column width
    def format_row(row):
        formatted_row = []
        for i, item in enumerate(row):
            item_str = str(item)
            padding = col_widths[i] - calc_display_width(item_str)
            formatted_row.append(item_str + ' ' * padding)
        return "| " + " | ".join(formatted_row) + " |"

    # Open the output file for writing
    with open(output_file, 'w', encoding='utf-8') as md_file:
        # Write the header
        headers = df.columns.tolist()
        md_file.write(format_row(headers) + '\n')
        #md_file.write('|' + '|'.join(['-' * (col_widths[i] + 2) for i in range(len(headers))]) + '|\n')
        md_file.write('\n')
        
        # Write each row
        for _, row in df.iterrows():
            md_file.write(format_row(row.tolist()) + '\n')

# Example usage
# excel_to_markdown('input.xlsx', 'output.md')

"""
Indices:
Function Purpose:
Converts an Excel file into a neatly formatted Markdown table.
Parameters:
source_file: Path to the source .xlsx file.
output_file: Path to the output .md file.
Functionality:
Reads Excel data, calculates the appropriate column widths (considering both English and Japanese characters), and formats the output in a Markdown table.
Notes:
Handles mixed-language input gracefully, ensuring the output Markdown table is properly aligned.
Example Usage:
Shows how to call the function with sample file paths.
"""

input_excel = "/Users/sambit/Documents/python workspace/GenAI/Demo Data.xlsx"
output_md = "/Users/sambit/Documents/python workspace/GenAI/Demo Data.md"

excel_to_markdown(input_excel, output_md)
print("Successfully converted the input excel to markdown")
