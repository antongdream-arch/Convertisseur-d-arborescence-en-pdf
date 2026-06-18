import sys
import logging
import argparse
from pathlib import Path


from converters import txt_converter, docx_converter, xlsx_converter
parser = argparse.ArgumentParser()
parser.add_argument("--debug", action="store_true")
parser.add_argument("--folder", help="file to convert")
parser.add_argument("--logfile", default="app_history.log", help="path to sve the logfile")
parser.add_argument("folder", help="original file (Input)")
parser.add_argument("--output", help="the choose file (Output)")
args = parser.parse_args()
if args.debug:
    logging.basicConfig(level=logging.DEBUG)
else:
     log_level = logging.DEBUG if args.debug else logging.INFO
     logging.basicConfig(level=log_level, format='%(leave_name)s: %(message)s')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s: %(message)s',
    filename=args.logfile,
    encoding="utf-8"
)

logging.info(f"the programme started.The log are in : {args.logfile}")


def scanner(input_folder):
    """scan for files in input_folder"""
    return input_folder.rglob("*")


def convert_files():
    """launch the files conversion after the detection of the file's type"""
    input_folder = Path(args.folder)

    if not input_folder.exists() or not input_folder.is_dir():
        logging.error("The input folder does not exist or is not a valid directory")
        sys.exit(1)
    if args.output:
        output_folder = Path(args.output)
    else:
        output_folder = input_folder.with_name(input_folder.name + "_PDF")

    valid_files = sorted([f for f in scanner(input_folder) if f.is_file() and f.stat().st_size > 0])
    total_files = len(valid_files)

    if total_files == 0:
        logging.info("No valid files found.")
        return

    files_processed = 0

    for file in valid_files:
        if file.suffix.lower() == ".txt":
            txt_converter(file, input_folder, output_folder)
        elif file.suffix.lower() == ".docx":
            docx_converter(file, input_folder, output_folder)
        elif file.suffix.lower() == ".xlsx":
            xlsx_converter(file, input_folder, output_folder)

        files_processed += 1
        percent = int((files_processed / total_files) * 100)

        sys.stdout.write(f"\rProgress: {files_processed} / {total_files} files ({percent}%)")
        sys.stdout.flush()

    print("\n\nConversion Complete!")


if __name__ == "__main__":
    convert_files()