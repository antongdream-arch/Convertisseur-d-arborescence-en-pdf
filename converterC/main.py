import sys
import logging
from pathlib import Path
from converters import txt_converter, docx_converter, xlsx_converter
import argparse


parser = argparse.ArgumentParser()
parser.add_argument("--debug", action="store_true", help="Enable highly detailed logs")
parser.add_argument("folder", help="Absolute or relative path to the source folder")
args = parser.parse_args()


log_level = logging.DEBUG if args.debug else logging.INFO

logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s')
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

if len(sys.argv) < 2:
    logging.error("Usage: python main.py <path_to_source_folder>")
    sys.exit(1)

input_folder = Path(sys.argv[1])

if not input_folder.exists():
    logging.error("The input folder does not exist")
    sys.exit(1)

output_folder = input_folder.with_name(input_folder.name + "_PDF")

def scanner():
    return input_folder.rglob("*")

def convert_files():

    test = "test"


    input_folder = Path(args.folder)
    for file in sorted(scanner()):
        if file.is_file() and file.stat().st_size > 0:
            if file.suffix.lower() == ".txt":
                txt_converter(file, input_folder, output_folder)
            elif file.suffix.lower() == ".docx":
                docx_converter(file, input_folder, output_folder)
            elif file.suffix.lower() == ".xlsx":
                xlsx_converter(file, input_folder, output_folder)

if __name__ == "__main__":
    convert_files()