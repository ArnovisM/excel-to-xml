import pandas as pd
import xml.etree.ElementTree as ET
import argparse
import re
import os

def clean_tag(tag):
    # Remove invalid characters for XML tags and ensure it starts with a letter if possible
    # For simplicity, we just remove non-alphanumeric chars. 
    # If the first char is a digit, prepend 'Field_'
    tag = re.sub(r'[^a-zA-Z0-9]', '', tag)
    if tag and tag[0].isdigit():
        tag = "Field_" + tag
    return tag

def convert_excel_to_xml(excel_path, xml_path):
    if not os.path.exists(excel_path):
        print(f"Error: File {excel_path} not found.")
        return

    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    root = ET.Element("Manifest")
    
    for _, row in df.iterrows():
        record = ET.SubElement(root, "Record")
        for col in df.columns:
            tag_name = clean_tag(str(col))
            if not tag_name:
                tag_name = "UnknownField"
            
            child = ET.SubElement(record, tag_name)
            val = row[col]
            child.text = str(val) if pd.notna(val) else ""
            
    tree = ET.ElementTree(root)
    try:
        ET.indent(tree, space="  ", level=0)
    except AttributeError:
        # ET.indent is available in Python 3.9+
        pass
        
    tree.write(xml_path, encoding="utf-8", xml_declaration=True)
    print(f"Successfully converted {excel_path} to {xml_path}")

def process_batch(file_paths, output_dir=None):
    """
    Convert multiple Excel files to XML.
    
    Args:
        file_paths (list): List of paths to Excel files.
        output_dir (str, optional): Directory to save XML files. 
                                    If None, saves in the same directory as input.
    """
    results = {"success": [], "failed": []}
    
    for file_path in file_paths:
        try:
            if not os.path.exists(file_path):
                results["failed"].append((file_path, "File not found"))
                continue
                
            # Determine output path
            base_name = os.path.basename(file_path)
            name_without_ext = os.path.splitext(base_name)[0]
            xml_name = name_without_ext + ".xml"
            
            if output_dir:
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                output_path = os.path.join(output_dir, xml_name)
            else:
                output_path = os.path.join(os.path.dirname(file_path), xml_name)
                
            convert_excel_to_xml(file_path, output_path)
            results["success"].append(file_path)
            
        except Exception as e:
            results["failed"].append((file_path, str(e)))
            print(f"Failed to convert {file_path}: {e}")
            
    return results

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert Excel to XML")
    parser.add_argument("input", nargs='+', help="Input Excel file path(s)")
    parser.add_argument("--output-dir", help="Output directory for XML files", default=None)
    
    args = parser.parse_args()
    
    # If single input and it looks like a file-to-file conversion (legacy support)
    if len(args.input) == 1 and args.output_dir and args.output_dir.endswith('.xml'):
        # This is a bit hacky to support the old "python main.py in.xlsx out.xml" style 
        # if the user tries to use it that way, but the new arg structure is different.
        # Let's stick to the new structure: input is a list, output is a dir.
        # But if the user wants specific output filename for single file, they might need the old function.
        # For now, let's just use process_batch for the CLI too.
        pass

    if len(args.input) == 1 and args.input[0].endswith('.xlsx') and args.output_dir and args.output_dir.endswith('.xml'):
         convert_excel_to_xml(args.input[0], args.output_dir)
    else:
         process_batch(args.input, args.output_dir)
