# Filename: convert_to_md.py

import sys
from markitdown import MarkItDown

def convert_docx_to_md(input_path, output_path):
    """
    Convert a DOCX file to Markdown format.

    Args:
        input_path (str): The path to the input DOCX file.
        output_path (str): The path to the output Markdown file.
    """
    md = MarkItDown()
    result = md.convert(input_path)

    # Write the converted content to the output file with UTF-8 encoding
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(result.text_content)

def main():
    """
    Main function to execute the DOCX to Markdown conversion.
    
    Modes:
        1. If no command-line arguments are provided, use the default input path.
        2. If an input path is provided as a command-line argument, use it.
    """
    # Define the default input and output paths
    default_input_path = r"test.docx"
    output_path = "output.md"

    # Check if an input path is provided as a command-line argument
    if len(sys.argv) > 1:
        input_path = sys.argv[1]
    else:
        input_path = default_input_path

    # Perform the conversion
    convert_docx_to_md(input_path, output_path)
    print(f"Conversion completed. Markdown file saved as '{output_path}'.")

if __name__ == "__main__":
    main()
