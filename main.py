import sys
import logging
import argparse
import os
import shutil
import time
from check_files import check_files
from pathlib import Path
from converters import create_pdf, docx_converter, doc_converter, xlsx_converter, copy_file
from config_log_level import setup_configurable_logger

logger = logging.getLogger("UniversalConverter")


def ensure_utf8_and_read(file_path):
    """
    Reads a text file by testing multiple encodings.
    If the original encoding is not UTF-8, the file is converted and rewritten in UTF-8.

    :param file_path: The path to the source text file.
    :type file_path: Path
    :return: The lines of text from the file.
    :rtype: list[str]
    """
    encodings_to_try = ['utf-8', 'cp1252', 'iso-8859-1']
    content_lines = None
    used_encoding = None

    for enc in encodings_to_try:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                content_lines = f.readlines()
            used_encoding = enc
            break
        except UnicodeDecodeError:
            continue

    if content_lines is None:
        logger.error(f"Critical decoding failure for file: {file_path.name}")
        return []

    if used_encoding != 'utf-8':
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.writelines(content_lines)
            logger.info(f"File converted to UTF-8: {file_path.name} (Old encoding: {used_encoding})")
        except Exception as e:
            logger.error(f"Error rewriting {file_path.name} in UTF-8: {e}")

    return content_lines


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
        logger.info("No valid files found.")
        return

    files_processed = 0

    for file in valid_files:
        try:
            if file.suffix.lower() == ".txt":
                logger.info(f"Processing TXT: {file.name}")
                lines = ensure_utf8_and_read(file)
                if lines:
                    create_pdf(lines, file, _input_folder, _output_folder)

            elif file.suffix.lower() == ".docx":
                logger.info(f"docx_converter({file}, {_input_folder}, {_output_folder})")
                docx_converter(file, _input_folder, _output_folder)

            elif file.suffix.lower() == ".doc":
                logger.info(f"doc_converter({file}, {_input_folder}, {_output_folder})")
                doc_converter(file, _input_folder, _output_folder)

            elif file.suffix.lower() in [".xlsx", ".xls"]:
                logger.info(f"xlsx_converter({file}, {_input_folder}, {_output_folder})")
                xlsx_converter(file, _input_folder, _output_folder)

            elif file.suffix.lower() == ".pdf":
                logger.info(f"copy file({file}, {_input_folder}, {_output_folder})")
                copy_file(file, _input_folder, _output_folder)

            else:
                logger.info(f"I can't convert this file (not the good suffix) {file}")

        except Exception as e:
            logger.error(f"Error processing file {file}: {e}")
            sys.exit(1)
        finally:
            files_processed += 1
            percent = int((files_processed / total_files) * 100)
            logger.debug(f"Progress: {files_processed} / {total_files} files ({percent}%)")

        time.sleep(2)

    logger.info("\nConversion Complete!")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", required=True, help="Path to the source folder")
    parser.add_argument("--debug", action="store_true", help="Enable debug mode")
    parser.add_argument("--logfile", help="Path to the custom log file")
    parser.add_argument("--output", help="Path to the custom output directory")
    args = parser.parse_args()

    # Ceci attache les handlers à "UniversalConverter"
    setup_configurable_logger(debug_mode=args.debug, log_file_path=args.logfile)

    if args.logfile:
        logger.info(f"The program started. Logs are saved in: {args.logfile}")
    else:
        logger.info("The program started. No log file specified, outputting to console only.")

    input_folder = Path(args.folder)

    if not input_folder.exists() or not input_folder.is_dir():
        logger.error("The input folder does not exist or is not a valid directory")
        sys.exit(1)

    if args.output:
        output_folder = Path(args.output)
    else:
        output_folder = input_folder.with_name(input_folder.name + "_PDF")

    try:
        if os.path.isdir(output_folder):
            logger.info(f"Folder output exists. We delete it: {output_folder}")
            shutil.rmtree(output_folder)
        else:
            logger.info(f"Folder output doesn't exist: {output_folder}")

        output_folder.mkdir(parents=True, exist_ok=True)

    except Exception as e:
        logger.error(f"Failed to create output folder: {e}")
        sys.exit(1)

    convert_files(input_folder, output_folder)