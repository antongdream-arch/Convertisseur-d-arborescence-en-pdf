import sys
import logging
import argparse
from pathlib import Path
from converters import txt_converter, docx_converter, xlsx_converter


def scanner(input_folder):
    return input_folder.rglob("*")


def convert_files():
    parser = argparse.ArgumentParser()
    # TODO why --debug or folder ?
    # TODO Add output directory
    # TODO Add logfile argument
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("folder")
    args = parser.parse_args()

    log_level = logging.DEBUG if args.debug else logging.INFO
    logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s')

    input_folder = Path(args.folder)

    if not input_folder.exists() or not input_folder.is_dir():
        logging.error("The input folder does not exist or is not a valid directory")
        sys.exit(1)

    output_folder = input_folder.with_name(input_folder.name + "_PDF")

    # TODO You don't need to sort to get number of files
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

        # TODO why you don't use logging system
        sys.stdout.write(f"\rProgress: {files_processed} / {total_files} files ({percent}%)")
        sys.stdout.flush()

    # TODO why you don't use logging system
    print("\n\nConversion Complete!")


if __name__ == "__main__":
    # TODO you need to parse arguments at the beginning
    convert_files()