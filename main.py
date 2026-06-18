import sys
import logging
import argparse
from pathlib import Path
from converters import txt_converter, docx_converter, xlsx_converter

def scanner(input_folder):
    """scan for files in input_folder"""
    return input_folder.rglob("*")

def convert_files():
    """launch the files conversion after the detection of the file's type"""
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", required=True, help="Path to the source folder")
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--logdir", help="Path to the log directory")
    parser.add_argument("--output", help="Path to the custom output directory")
    args = parser.parse_args()

    log_level = logging.DEBUG if args.debug else logging.INFO

    if args.logdir:
        log_folder = Path(args.logdir)
        log_folder.mkdir(parents=True, exist_ok=True)
        log_file_path = log_folder / "app_history.log"
    else:
        log_file_path = Path("app_history.log")

    logging.basicConfig(
        level=log_level,
        format='%(asctime)s | %(levelname)s: %(message)s',
        filename=log_file_path,
        encoding="utf-8"
    )

    logging.info(f"The program started. Logs are saved in: {log_file_path}")

    input_folder = Path(args.folder)

    if not input_folder.exists() or not input_folder.is_dir():
        logging.error("The input folder does not exist or is not a valid directory")
        sys.exit(1)

    if args.output:
        output_folder = Path(args.output)
    else:
        output_folder = input_folder.with_name(input_folder.name + "_PDF")

    output_folder.mkdir(parents=True, exist_ok=True)

    valid_files = sorted([f for f in scanner(input_folder) if f.is_file() and f.stat().st_size > 0])
    total_files = len(valid_files)

    if total_files == 0:
        logging.info("No valid files found.")
        print("No valid files found.")
        return

    print(f"\nSource: {input_folder}")
    print(f"Target: {output_folder}\n")

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