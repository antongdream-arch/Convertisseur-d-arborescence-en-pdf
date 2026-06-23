import sys
import logging
import argparse
import os
import shutil
import time
from check_files import check_files
from pathlib import Path
from converters import txt_converter, docx_converter, doc_converter, xlsx_converter, copy_file


def scanner(_input_folder):
    """
    Recursively scans the input folder to find all items it contains.

    :param _input_folder: The path of the root folder to scan.
    :type _input_folder: Path
    :return: An iterator yielding the paths of all found elements (files and folders).
    :rtype: generator
    """
    return _input_folder.rglob("*")


def convert_files(_input_folder, _output_folder):
    """
    Filters, sorts, and converts valid files from the input folder to the output folder.

    This function ignores subdirectories and empty files. It checks the extension
    of each file and calls the appropriate converter (.txt, .docx, .doc, .xlsx).
    Overall progress is recorded in the logs.

    :param _input_folder: The root folder containing the original files to process.
    :type _input_folder: Path
    :param _output_folder: The destination folder where the generated PDFs will be saved.
    :type _output_folder: Path
    :return: No value returned.
    :rtype: None
    """
    valid_files = sorted([
        f for f in scanner(_input_folder)
        if f.is_file()
           and f.stat().st_size > 0
           and not f.name.startswith("~$")
           and not f.name.startswith(".")
    ])
    total_files = len(valid_files)

    if total_files == 0:
        logging.info("No valid files found.")
        return

    files_processed = 0

    for file in valid_files:
        try:
            if file.suffix.lower() == ".txt":
                logging.info(f"txt_converter({file}, {_input_folder}, {_output_folder})")
                txt_converter(file, _input_folder, _output_folder)

            elif file.suffix.lower() == ".docx":
                logging.info(f"docx_converter({file}, {_input_folder}, {_output_folder})")
                docx_converter(file, _input_folder, _output_folder)

            elif file.suffix.lower() == ".doc":
                logging.info(f"doc_converter({file}, {_input_folder}, {_output_folder})")
                doc_converter(file, _input_folder, _output_folder)

            elif file.suffix.lower() in [".xlsx", ".xls"]:
                logging.info(f"xlsx_converter({file}, {_input_folder}, {_output_folder})")
                xlsx_converter(file, _input_folder, _output_folder)

            elif file.suffix.lower() == ".pdf":
                logging.info(f"copy file({file}, {_input_folder}, {_output_folder})")
                copy_file(file, _input_folder, _output_folder)

            else:
                logging.info(f"I can't convert this file (not the good sufix) {file}")

        except Exception as e:

            logging.error(f"Error processing file {file}: {e}")
            sys.exit(1)
        finally:
            files_processed += 1
            percent = int((files_processed / total_files) * 100)
            logging.debug(f"Progress: {files_processed} / {total_files} files ({percent}%)")

        time.sleep(2)

    logging.info("\nConversion Complete!")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", required=True, help="Path to the source folder")
    parser.add_argument("--debug", action="store_true", help="Enable debug mode")
    parser.add_argument("--logfile", help="Path to the custom log file")
    parser.add_argument("--output", help="Path to the custom output directory")
    args = parser.parse_args()

    log_level = logging.DEBUG if args.debug else logging.INFO

    handlers = [logging.StreamHandler(sys.stdout)]

    if args.logfile:
        logfile = Path(args.logfile)
        handlers.append(logging.FileHandler(logfile, encoding="utf-8"))

    logging.basicConfig(
        level=log_level,
        format='%(asctime)s | %(levelname)s: %(message)s',
        handlers=handlers
    )

    if args.logfile:
        logging.info(f"The program started. Logs are saved in: {args.logfile}")
    else:
        logging.info("The program started. No log file specified, outputting to console only.")

    input_folder = Path(args.folder)

    if not input_folder.exists() or not input_folder.is_dir():
        logging.error("The input folder does not exist or is not a valid directory")
        sys.exit(1)

    if args.output:
        output_folder = Path(args.output)
    else:
        output_folder = input_folder.with_name(input_folder.name + "_PDF")

    try:
        if os.path.isdir(output_folder):
            logging.info(f"Folder output exists. We delete it: {output_folder}")
            shutil.rmtree(output_folder)
        else:
            logging.info(f"Folder output doesn't exist: {output_folder}")

        output_folder.mkdir(parents=True, exist_ok=True)

    except Exception as e:
        logging.error(f"Error setting up output folder: {e}")
        sys.exit(1)

    convert_files(input_folder, output_folder)

    try:
        check_files(input_folder, output_folder)
    except Exception as e:
        logging.error(f"Error during file verification: {e}")
        sys.exit(1)