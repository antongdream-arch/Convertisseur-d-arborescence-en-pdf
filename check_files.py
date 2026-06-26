import logging
import sys
from pathlib import Path

logger = logging.getLogger("UniversalConverter")


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
            if f.is_file()
            and not f.name.startswith("~$")
            and not f.name.startswith(".")
            and f.suffix.lower() in target_extensions
            and f.stat().st_size > 0
        ]
    except Exception as e:
        logger.error(f"Verification aborted: Critical error reading input folder {input_folder}. Details: {e}")
        sys.exit(1)

    total_targets = len(target_files)

    if total_targets == 0:
        logger.info("Verification skipped: No target files (convertible or PDF) were found.")
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
            logger.error(f"Error while verifying file {file.name}: {e}")
            missing_files.append(f"{file.name} (Access/Path Error)")

    if success_count == total_targets:
        logger.info(f"Verification passed: All {total_targets} target files (converted and copied PDFs) are present.")
    else:
        missing_count = total_targets - success_count
        logger.warning(f"Verification failed: {missing_count} out of {total_targets} expected PDFs are missing!")

        for missing in missing_files:
            logger.debug(f"Missing expected PDF for original file: {missing}")