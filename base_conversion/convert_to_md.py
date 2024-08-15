import pandas as pd
import os

def excel_to_markdown(source_file: str, output_file: str):
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
