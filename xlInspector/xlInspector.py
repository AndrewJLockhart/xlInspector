import argparse
import json
from oletools.olevba import VBA_Parser
import os

def analyze_file(filename, output_file):
    vba_parser = VBA_Parser(filename)
    vba_modules = vba_parser.extract_all_macros()
    
    if vba_modules:
        print(f"Found {len(vba_modules)} VBA module(s) in {filename}:")
        for module in vba_modules:
            print(f"Module name: {module.name}")
            print(f"Module code:\n{module.code}\n")
        
        # Save the output as JSON
        if output_file:
            output_data = {
                "filename": filename,
                "vba_modules": [
                    {
                        "name": module.name,
                        "code": module.code
                    }
                    for module in vba_modules
                ]
            }
            with open(output_file, "w") as f:
                json.dump(output_data, f, indent=4)
            print(f"Output saved to {output_file}")
    else:
        print(f"No VBA modules found in {filename}")

def main():
    parser = argparse.ArgumentParser(description="Tool for extracting VBA code from Office files")
    parser.add_argument("-f", "--filename", help="Path to the file to analyze")
    parser.add_argument("-o", "--output", help="Path to the output JSON file")
    args = parser.parse_args()
    
    if not args.filename:
        parser.error("Please provide the path to the file to analyze using the -f or --filename option.")
    if not args.output:
        parser.error("Please provide the path to the output JSON file using the -o or --output option.")
    
    if os.path.exists(args.filename):
        if not os.path.exists(args.output):
            analyze_file(args.filename, args.output)
        else:
            print(f"Output file {args.output} already exists.")
    else:
        print(f"File {args.filename} does not exist.")
    

if __name__ == "__main__":
    main()