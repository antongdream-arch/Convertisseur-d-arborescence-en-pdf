import logging
from pathlib import Path

PDF_FONT = "Helvetica"
PDF_FONT_SIZE = 12
PDF_MARGIN_LEFT = 72
PDF_MARGIN_BOTTOM = 72
PDF_LINE_HEIGHT = 15


def setup_configurable_logger(debug_mode=False, log_file_path=None):
    """
    Configures and returns a custom logger for the application.

    This setup includes two handlers:
    1. A console handler that outputs INFO logs by default, or DEBUG logs if enabled.
    2. An optional file handler that records DEBUG level logs.
    It also clears any existing handlers to prevent duplicate log entries.

    :param debug_mode: A flag to set the console logging level to DEBUG instead of INFO. Defaults to False.
    :type debug_mode: bool
    :param log_file_path: Optional path to save the log file. If None, defaults to 'conversion_history.log'.
    :type log_file_path: Path or str
    :return: The configured logger instance ready to be used across the application.
    :rtype: logging.Logger
    """
    logger = logging.getLogger("UniversalConverter")
    logger.setLevel(logging.DEBUG)

    if logger.hasHandlers():
        logger.handlers.clear()

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG if debug_mode else logging.INFO)
    console_format = logging.Formatter('%(levelname)s: %(message)s')
    console_handler.setFormatter(console_format)
    logger.addHandler(console_handler)

    if log_file_path is None:
        dossier_project = Path(__file__).parent
        dossier_logs = dossier_project / "logs"
        dossier_logs.mkdir(exist_ok=True)
        log_file_path = dossier_logs / "conversion_history.log"

    try:
        file_handler = logging.FileHandler(log_file_path, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_format = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler.setFormatter(file_format)
        logger.addHandler(file_handler)
    except Exception as e:
        logger.warning(f"Failed to create log file at '{log_file_path}': {e}")
        logger.warning("Logs will only be displayed in the console.")

    return logger