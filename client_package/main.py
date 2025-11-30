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

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert Excel to XML")
    parser.add_argument("input", help="Input Excel file path")
    parser.add_argument("output", help="Output XML file path")
    
    args = parser.parse_args()
    convert_excel_to_xml(args.input, args.output)
