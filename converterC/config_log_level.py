import logging


PDF_FONT = "Helvetica"
PDF_FONT_SIZE = 12
PDF_MARGIN_LEFT = 72
PDF_MARGIN_BOTTOM = 72
PDF_LINE_HEIGHT = 15


def setup_configurable_logger(debug_mode=False):
    """setup the logger"""
    logger = logging.getLogger("UniversalConverter")
    logger.setLevel(logging.DEBUG)


    if logger.hasHandlers():
        logger.handlers.clear()


    console_handler = logging.StreamHandler()

    console_handler.setLevel(logging.DEBUG if debug_mode else logging.INFO)
    console_format = logging.Formatter('%(levelname)s: %(message)s')
    console_handler.setFormatter(console_format)


    file_handler = logging.FileHandler("conversion_history.log", encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)

    file_format = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
    file_handler.setFormatter(file_format)

    
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    return logger