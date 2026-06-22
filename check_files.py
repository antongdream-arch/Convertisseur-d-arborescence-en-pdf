import logging
import sys
from pathlib import Path


def check_files(input_folder, output_folder):
    """
    Verifies that all convertible files AND existing PDFs from the input folder
    are correctly present as PDFs in the output folder.

    :param input_folder: The root folder containing the original files.
    :type input_folder: Path
    :param output_folder: The root folder where the generated/copied PDF files are saved.
    :type output_folder: Path
    :return: No value returned.
    :rtype: None
    """

    target_extensions = [".xls", ".xlsx", ".txt", ".docx", ".doc", ".pdf"]

    try:

        target_files = [
            f for f in input_folder.rglob("*")
            if f.is_file() and f.stat().st_size > 0 and f.suffix.lower() in target_extensions
        ]
    except Exception as e:
        logging.error(f"Verification aborted: Critical error reading input folder {input_folder}. Details: {e}")
        sys.exit(1)
        return

    total_targets = len(target_files)

    if total_targets == 0:
        logging.info("Verification skipped: No target files (convertible or PDF) were found.")
        return

    success_count = 0
    missing_files = []

    for file in target_files:
        try:
            relative_path = file.relative_to(input_folder)
            expected_pdf = (output_folder / relative_path).with_suffix(".pdf")


            if expected_pdf.exists():
                success_count += 1
            else:
                missing_files.append(file.name)

        except Exception as e:
            logging.error(f"Error while verifying file {file.name}: {e}")
            missing_files.append(f"{file.name} (Access/Path Error)")
            sys.exit(1)

    if success_count == total_targets:
        logging.info(f"Verification passed: All {total_targets} target files (converted and copied PDFs) are present.")
    else:
        missing_count = total_targets - success_count
        logging.warning(f"Verification failed: {missing_count} out of {total_targets} expected PDFs are missing!")

        for missing in missing_files:
            logging.debug(f"Missing expected PDF for original file: {missing}")