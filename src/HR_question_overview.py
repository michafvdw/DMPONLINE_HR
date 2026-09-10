
import argparse
import logging
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from dmponline import DMPonline


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def get_answer_text(question):
    """
    Convert the answer structure returned by DMPonline into readable text.

    DMPonline can return different answer structures depending on the
    question type. This function tries to handle the common cases:
      - free-text answers
      - option-based answers
      - multiple selected options
      - simple scalar answers
    """

    if not question.get('answered'):
        return None

    answer = question.get('answer')

    if answer is None:
        return ""

    # ---------------------------------------------------------------
    # String / scalar answer
    # ---------------------------------------------------------------

    if isinstance(answer, str):
        return answer

    if isinstance(answer, (int, float, bool)):
        return str(answer)

    # ---------------------------------------------------------------
    # Dictionary answer
    # ---------------------------------------------------------------

    if isinstance(answer, dict):

        # Option-based answers
        #
        # Example:
        # {
        #     "options": [
        #         {"text": "Yes"},
        #         {"text": "No"}
        #     ]
        # }
        options = answer.get('options')

        if options:
            option_texts = []

            for option in options:

                if isinstance(option, dict):
                    text = (
                        option.get('text')
                        or option.get('label')
                        or option.get('name')
                    )

                    if text:
                        option_texts.append(str(text))

                else:
                    option_texts.append(str(option))

            if option_texts:
                return "\n".join(option_texts)

        # Common text/value fields
        for key in (
            'text',
            'value',
            'answer',
            'content',
            'response'
        ):
            value = answer.get(key)

            if value is not None:
                return str(value)

        # -----------------------------------------------------------
        # If we don't recognise the structure, return it as text
        # rather than silently losing the answer.
        # -----------------------------------------------------------

        return str(answer)

    # ---------------------------------------------------------------
    # List / tuple answer
    # ---------------------------------------------------------------

    if isinstance(answer, (list, tuple)):

        values = []

        for item in answer:

            if isinstance(item, dict):
                value = (
                    item.get('text')
                    or item.get('label')
                    or item.get('name')
                    or item.get('value')
                )

                if value is not None:
                    values.append(str(value))
                else:
                    values.append(str(item))

            else:
                values.append(str(item))

        return "\n".join(values)

    # ---------------------------------------------------------------
    # Fallback
    # ---------------------------------------------------------------

    return str(answer)


def safe_sheet_name(name, used_names):
    """
    Create an Excel-compatible worksheet name.

    Excel:
      - limits sheet names to 31 characters
      - does not allow: \\ / * ? : [ ]
      - does not allow duplicate sheet names
    """

    if not name:
        name = "Section"

    # Remove invalid Excel characters
    invalid_characters = ['\\', '/', '*', '?', ':', '[', ']']

    for character in invalid_characters:
        name = name.replace(character, '-')

    name = name.strip()

    if not name:
        name = "Section"

    # Excel limits worksheet names to 31 characters
    name = name[:31]

    original_name = name
    counter = 2

    while name in used_names:

        suffix = f" ({counter})"

        name = original_name[:31 - len(suffix)] + suffix

        counter += 1

    used_names.add(name)

    return name


def format_worksheet(ws):
    """
    Apply formatting to a worksheet.
    """

    # ---------------------------------------------------------------
    # Colours / styles
    # ---------------------------------------------------------------

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    section_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAF7"
    )

    header_font = Font(
        bold=True,
        color="FFFFFF"
    )

    question_font = Font(
        bold=True
    )

    title_font = Font(
        bold=True,
        size=16
    )

    subtitle_font = Font(
        bold=True,
        size=11
    )

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    # ---------------------------------------------------------------
    # Title
    # ---------------------------------------------------------------

    ws["A1"].font = title_font

    # ---------------------------------------------------------------
    # Header row
    # ---------------------------------------------------------------

    header_row = 4

    for cell in ws[header_row]:

        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True
        )

        cell.border = thin_border

    ws.row_dimensions[header_row].height = 30

    # ---------------------------------------------------------------
    # Data rows
    # ---------------------------------------------------------------

    for row in ws.iter_rows(
        min_row=header_row + 1
    ):

        for cell in row:

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

            cell.border = thin_border

    # ---------------------------------------------------------------
    # Make question column bold
    # ---------------------------------------------------------------

    for row in range(
        header_row + 1,
        ws.max_row + 1
    ):

        # Question text is column C
        if ws.cell(row=row, column=3).value:
            ws.cell(
                row=row,
                column=3
            ).font = question_font

    # ---------------------------------------------------------------
    # Column widths
    # ---------------------------------------------------------------

    widths = {
        "A": 14,   # Question
        "B": 15,   # Type
        "C": 65,   # Question text
        "D": 100   # Answer
    }

    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    # ---------------------------------------------------------------
    # Freeze question headers
    # ---------------------------------------------------------------

    ws.freeze_panes = "A5"

    # ---------------------------------------------------------------
    # Autofilter
    # ---------------------------------------------------------------

    if ws.max_row >= 4:
        ws.auto_filter.ref = (
            f"A4:D{ws.max_row}"
        )


# ---------------------------------------------------------------------------
# Main Excel export
# ---------------------------------------------------------------------------

def question_overview(plan_id, api, output_file=None):
    """
    Export answered questions from a specific DMPonline plan.

    Creates one Excel worksheet per DMP section, plus an Overview sheet.
    The raw API response is used here because DMPonline.get_plan_v0()
    returns a pandas DataFrame and therefore does not expose the nested
    plan_content structure directly.
    """

    # Get the raw DMP response from the API.
    parsed = api.get(
        request='v0/plans?plan={}'.format(plan_id),
        params={'remove_tests': 'false'}
    )

    if not parsed:
        raise ValueError(
            f"No DMP found for plan ID {plan_id}"
        )

    # The API returns a list containing the plan.
    try:
        plan = parsed[0]
        plan_content = plan['plan_content'][0]
        sections = plan_content['sections']
    except (IndexError, KeyError, TypeError) as exc:
        raise ValueError(
            "Could not find the DMP sections in the API response."
        ) from exc

    workbook = Workbook()

    # Remove the default worksheet.
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    used_sheet_names = {"Overview"}

    # ------------------------------------------------------------------
    # Overview sheet
    # ------------------------------------------------------------------

    overview = workbook.create_sheet("Overview")

    overview["A1"] = "DMPonline DMP overview"
    overview["A1"].font = Font(bold=True, size=18)

    overview["A3"] = "Plan ID"
    overview["B3"] = plan_id

    # Add useful metadata when available.
    overview["A4"] = "DMP title"
    overview["B4"] = (
        plan.get("title")
        or plan.get("name")
        or ""
    )

    overview["A5"] = "Template"
    template = plan.get("template") or {}
    overview["B5"] = template.get("title", "")

    overview["A7"] = "Section"
    overview["B7"] = "Worksheet"
    overview["C7"] = "Answered questions"

    overview_header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    overview_header_font = Font(
        bold=True,
        color="FFFFFF"
    )

    overview_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    for cell in overview[7]:
        cell.fill = overview_header_fill
        cell.font = overview_header_font
        cell.alignment = Alignment(
            vertical="center",
            wrap_text=True
        )
        cell.border = overview_border

    overview_row = 8

    # ------------------------------------------------------------------
    # One worksheet per section
    # ------------------------------------------------------------------

    for section_index, section in enumerate(sections, start=1):

        section_number = section.get(
            "number",
            section_index
        )

        section_title = (
            section.get("title")
            or section.get("name")
            or f"Section {section_number}"
        )

        questions = section.get("questions") or []

        # Only include questions for which DMPonline says answered=True.
        answered_questions = []

        for question in questions:

            if not question.get("answered"):
                continue

            answer_text = get_answer_text(question)

            if answer_text is not None:
                answered_questions.append(
                    (question, answer_text)
                )

        # Create a valid and unique Excel worksheet name.
        sheet_name = safe_sheet_name(
            f"{section_number} - {section_title}",
            used_sheet_names
        )

        ws = workbook.create_sheet(sheet_name)

        # ------------------------------------------------------------------
        # Section heading
        # ------------------------------------------------------------------

        ws["A1"] = (
            f"Section {section_number}: {section_title}"
        )

        ws["A1"].font = Font(
            bold=True,
            size=16
        )

        ws.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=4
        )

        ws["A2"] = "Answered questions"
        ws["B2"] = len(answered_questions)

        # ------------------------------------------------------------------
        # Table headers
        # ------------------------------------------------------------------

        ws["A4"] = "Question"
        ws["B4"] = "Type"
        ws["C4"] = "Question"
        ws["D4"] = "Answer"

        row = 5

        for question, answer_text in answered_questions:

            ws.cell(
                row=row,
                column=1,
                value=question.get("number", "")
            )

            ws.cell(
                row=row,
                column=2,
                value=question.get("format", "")
            )

            ws.cell(
                row=row,
                column=3,
                value=question.get("text", "")
            )

            ws.cell(
                row=row,
                column=4,
                value=answer_text
            )

            row += 1

        # ------------------------------------------------------------------
        # Formatting
        # ------------------------------------------------------------------

        format_worksheet(ws)

        # Make rows containing long answers reasonably tall.
        for data_row in range(5, ws.max_row + 1):
            ws.row_dimensions[data_row].height = 75

        # ------------------------------------------------------------------
        # Add section to Overview
        # ------------------------------------------------------------------

        overview.cell(
            row=overview_row,
            column=1,
            value=f"{section_number}: {section_title}"
        )

        overview.cell(
            row=overview_row,
            column=2,
            value=sheet_name
        )

        overview.cell(
            row=overview_row,
            column=3,
            value=len(answered_questions)
        )

        for column in range(1, 4):
            overview.cell(
                row=overview_row,
                column=column
            ).border = overview_border

        overview_row += 1

    # ------------------------------------------------------------------
    # Format Overview
    # ------------------------------------------------------------------

    for row in overview.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

    overview.column_dimensions["A"].width = 55
    overview.column_dimensions["B"].width = 45
    overview.column_dimensions["C"].width = 22

    overview.freeze_panes = "A8"

    if overview_row > 8:
        overview.auto_filter.ref = (
            f"A7:C{overview_row - 1}"
        )

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------

    if not output_file:
        raise ValueError(
            "An output Excel filename is required. "
            "Use -o filename.xlsx"
        )

    extension = os.path.splitext(output_file)[-1].lower()

    if extension != ".xlsx":
        raise ValueError(
            "The output file must have an .xlsx extension."
        )

    workbook.save(output_file)

    logging.info(
        f"Excel report written to {output_file}"
    )


def main():
    parser = argparse.ArgumentParser(
        description="Export answered DMPonline questions to Excel."
    )

    parser.add_argument(
        "-i",
        "--id",
        dest="plan_id",
        required=True,
        help="DMPonline plan ID"
    )

    parser.add_argument(
        "-t",
        "--token",
        dest="token",
        required=True,
        help="DMPonline API token"
    )

    parser.add_argument(
        "-u",
        "--user",
        dest="token_user",
        required=True,
        help="DMPonline user email address"
    )

    parser.add_argument(
        "-o",
        "--output",
        required=True,
        help="Output Excel filename"
    )

    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(message)s"
    )

    # Create the DMPonline API client
    api = DMPonline(
        token=args.token,
        token_user=args.token_user
    )

    # Export the DMP to Excel
    question_overview(
        plan_id=args.plan_id,
        api=api,
        output_file=args.output
    )


if __name__ == "__main__":
    main()

