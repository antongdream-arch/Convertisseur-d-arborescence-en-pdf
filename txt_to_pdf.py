import logging
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from config_log_level import PDF_FONT, PDF_FONT_SIZE, PDF_MARGIN_LEFT, PDF_MARGIN_BOTTOM, PDF_LINE_HEIGHT

logger = logging.getLogger("UniversalConverter")


def create_pdf(lines, source_file, input_folder, output_folder):
    """
    Converts a list of text lines into a PDF document using ReportLab.

    :param lines: Lines of text read from the source file.
    :type lines: list[str]
    :param source_file: The full path to the source text file.
    :type source_file: Path
    :param input_folder: The root folder containing the original text files.
    :type input_folder: Path
    :param output_folder: The root folder where the generated PDF files will be saved.
    :type output_folder: Path
    :return: No value returned.
    :rtype: None
    """
    try:
        relative_path = source_file.relative_to(input_folder)
        pdf_file = (output_folder / relative_path).with_suffix(".pdf")
        pdf_file.parent.mkdir(parents=True, exist_ok=True)

        doc = SimpleDocTemplate(
            str(pdf_file),
            pagesize=A4,
            rightMargin=72,
            leftMargin=PDF_MARGIN_LEFT,
            topMargin=72,
            bottomMargin=PDF_MARGIN_BOTTOM
        )

        styles = getSampleStyleSheet()
        custom_style = ParagraphStyle(
            'CustomStyle',
            parent=styles['Normal'],
            fontName=PDF_FONT,
            fontSize=PDF_FONT_SIZE,
            spaceAfter=4,
        )

        story = []
        for line in lines:
            clean_line = line.strip().replace('\t', '&nbsp;&nbsp;&nbsp;&nbsp;')

            if not clean_line:
                story.append(Spacer(1, PDF_LINE_HEIGHT))
                continue

            clean_line = clean_line.replace('<', '&lt;').replace('>', '&gt;')

            p = Paragraph(clean_line, custom_style)
            story.append(p)

        doc.build(story)
        logger.info(f"PDF generated successfully (TXT): {pdf_file.name}")

    except Exception as e:
        logger.error(f"Failed to generate PDF for text file {source_file.name}: {e}")