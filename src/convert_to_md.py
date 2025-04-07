# Filename: convert_docx_to_md.py

import sys
import os
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

def get_default_paths():
    """
    Determine the default input and output paths.

    Returns:
        tuple: A tuple containing the default input path and output path.
    """
    # Define the default input path
    default_input_path = r"C:\Users\shade\OneDrive\000_gits\tec\docs\鑫图上海微TEC控制设备协议开发说明(V1.0.6).docx"

    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Define the output directory as a sibling 'output' directory
    output_dir = os.path.abspath(os.path.join(script_dir, '..', 'output'))

    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Define the output path
    output_path = os.path.join(output_dir, "output.md")
    
    config_dir=os.path.abspath(os.path.join(script_dir,'..','config'))
    config_file_path=os.path.join(config_dir,'to_md_local.txt')

    with open(config_file_path,'r',encoding='utf-8') as file:
        input_file=file.readline().strip()
    default_input_path=input_file
    return default_input_path, output_path

def main():
    """
    Main function to execute the DOCX to Markdown conversion.
    
    Modes:
        1. If no command-line arguments are provided, use the default input path.
        2. If an input path is provided as a command-line argument, use it.
    """
    # Get the default input and output paths
    default_input_path, output_path = get_default_paths()

    # Check if an input path is provided as a command-line argument
    if len(sys.argv) > 1:
        input_path = sys.argv[1]
    else:
        input_path = default_input_path

    # Perform the conversion
    convert_docx_to_md(input_path, output_path)
    print(f"Conversion completed. Markdown file saved at '{output_path}'.")

if __name__ == "__main__":
    main()
