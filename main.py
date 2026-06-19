import sys
import logging
import argparse
import os
import shutil
import time
from pathlib import Path
from converters import txt_converter, docx_converter, xlsx_converter


def scanner(_input_folder):
    """scan for files in input_folder"""
    return _input_folder.rglob("*")


def convert_files(_input_folder, _output_folder):
    valid_files = sorted([f for f in scanner(_input_folder) if f.is_file() and f.stat().st_size > 0])
    total_files = len(valid_files)

    if total_files == 0:
        logging.info("No valid files found.")
        return

    files_processed = 0

    for file in valid_files:
        if file.suffix.lower() == ".txt":
            logging.info(f"txt_converter({file}, {_input_folder}, {_output_folder})")
            txt_converter(file, _input_folder, _output_folder)
        elif file.suffix.lower() == ".docx":
            logging.info(f"docx_converter({file}, {_input_folder}, {_output_folder})")
            docx_converter(file, _input_folder, _output_folder)

        elif file.suffix.lower() == ".xlsx":
            logging.info(f"xlsx_converter({file}, {_input_folder}, {_output_folder})")
            xlsx_converter(file, _input_folder, _output_folder)
        else:
            logging.error(f"NOT CONVERTED {file}")

        files_processed += 1
        percent = int((files_processed / total_files) * 100)

        logging.debug(f"Progress: {files_processed} / {total_files} files ({percent}%)")
        time.sleep(2)

    logging.info("nConversion Complete!")


if __name__ == "__main__":
    """launch the files conversion after the detection of the file's type"""
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", required=True, help="Path to the source folder")
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--logfile", help="logfile")
    parser.add_argument("--output", help="Path to the custom output directory")
    args = parser.parse_args()

    log_level = logging.DEBUG if args.debug else logging.INFO

    logfile = Path(args.logfile)

    logging.basicConfig(
        level=log_level,
        format='%(asctime)s | %(levelname)s: %(message)s',
        encoding="utf-8",
        handlers=[
            logging.FileHandler(logfile),
            logging.StreamHandler(sys.stdout)
        ]
    )

    logging.info(f"The program started. Logs are saved in: {logfile}")

    input_folder = Path(args.folder)

    if not input_folder.exists() or not input_folder.is_dir():
        logging.error("The input folder does not exist or is not a valid directory")
        sys.exit(1)

    if args.output:
        output_folder = Path(args.output)
    else:
        output_folder = input_folder.with_name(input_folder.name + "_PDF")

    if os.path.isdir(output_folder):
        logging.info(f"Folder output exist. we delete it: {output_folder}")
        shutil.rmtree(output_folder)
    else:
        logging.info(f"Folder output doesn't exist: {output_folder}")

    output_folder.mkdir(parents=True, exist_ok=True)
    convert_files(input_folder, output_folder)
