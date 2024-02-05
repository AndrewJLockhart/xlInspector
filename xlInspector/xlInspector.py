import argparse
import json
from oletools.olevba import VBA_Parser
import os
import hashlib

def analyze_file(filename, output_file):
    

    vba_parser = VBA_Parser(filename)
    vba_modules = vba_parser.extract_all_macros()
    

    if vba_modules:
        print(f"Found {len(vba_modules)} VBA module(s) in {filename}:")
        for _, _, vba_filename, code in vba_modules:
            output_data= {
                "vbaObjName" : vba_filename,
                "vbaCode" : hashlib.sha3_512(code.encode()).hexdigest()
            }
            print(output_data)
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