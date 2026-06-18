import logging


PDF_FONT = "Helvetica"
PDF_FONT_SIZE = 12
PDF_MARGIN_LEFT = 72
PDF_MARGIN_BOTTOM = 72
PDF_LINE_HEIGHT = 15

# TODO change filename
# TODO add comments with """
def setup_configurable_logger(debug_mode=False):
    # 1. Création du Logger principal
    logger = logging.getLogger("UniversalConverter")
    logger.setLevel(logging.DEBUG)  # Le logger global accepte tout

    # Évite de dupliquer les logs si la fonction est appelée deux fois
    if logger.hasHandlers():
        logger.handlers.clear()

    # 2. Configuration du Terminal (Console)
    console_handler = logging.StreamHandler()
    # Le terminal affiche INFO en temps normal, ou DEBUG si l'option est activée
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