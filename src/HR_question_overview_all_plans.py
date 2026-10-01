import argparse
import logging
import os
import re

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from dmponline import DMPonline


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------


def get_answer_text(question):
    """
    Convert the answer structure returned by DMPonline into readable text.

    Handles:
        - free-text answers
        - scalar answers
        - option-based answers
        - multiple selected options
        - list/tuple answers
        - dictionary answers
    """

    if not question.get("answered"):
        return ""

    answer = question.get("answer")

    if answer is None:
        return ""

    # String / scalar answer
    if isinstance(answer, str):
        return answer

    if isinstance(answer, (int, float, bool)):
        return str(answer)

    # Dictionary answer
    if isinstance(answer, dict):

        # Option-based answers
        options = answer.get("options")

        if options:
            option_texts = []

            for option in options:
                if isinstance(option, dict):
                    text = (
                        option.get("text")
                        or option.get("label")
                        or option.get("name")
                        or option.get("value")
                    )

                    if text is not None:
                        option_texts.append(str(text))
                else:
                    option_texts.append(str(option))

            if option_texts:
                return "\n".join(option_texts)

        # Common text/value fields
        for key in (
            "text",
            "value",
            "answer",
            "content",
            "response",
        ):
            value = answer.get(key)

            if value is not None:
                return str(value)

        # Unknown structure: keep the information instead of losing it.
        return str(answer)

    # List / tuple answer
    if isinstance(answer, (list, tuple)):

        values = []

        for item in answer:
            if isinstance(item, dict):
                value = (
                    item.get("text")
                    or item.get("label")
                    or item.get("name")
                    or item.get("value")
                )

                if value is not None:
                    values.append(str(value))
            else:
                values.append(str(item))

        return "\n".join(values)

    return str(answer)


def find_section(sections, section_number):
    """Find a section by its section number."""

    for section in sections:
        number = section.get("number")

        if str(number) == str(section_number):
            return section

    return None


# ---------------------------------------------------------------------------
# Retrieve all DMPs
# ---------------------------------------------------------------------------


def get_all_plans(api, include_tests=True):
    """
    Retrieve all DMP plans available to the authenticated DMPonline
    organisation administrator.

    The DMPonline v0 API paginates the plans endpoint.
    """

    logging.info("Retrieving all DMPonline plans...")

    all_plans = []
    page = 1

    while True:
        logging.info("Retrieving plans page %s...", page)

        params = {
            "page": page,
            "remove_tests": "false" if include_tests else "true",
        }

        parsed = api.get(
            request="v0/plans",
            params=params,
        )

        if not parsed:
            break

        # Normally DMPonline v0 returns a list.
        # This also handles a dictionary containing a common
        # data/items/plans key.
        if isinstance(parsed, list):
            page_plans = parsed

        elif isinstance(parsed, dict):
            page_plans = (
                parsed.get("data")
                or parsed.get("items")
                or parsed.get("plans")
                or []
            )

        else:
            raise ValueError(
                f"Unexpected response type while retrieving page {page}: "
                f"{type(parsed).__name__}"
            )

        if not page_plans:
            break

        all_plans.extend(page_plans)

        logging.info(
            "Page %s: found %s plans (total so far: %s).",
            page,
            len(page_plans),
            len(all_plans),
        )

        page += 1

    logging.info("Total plans found: %s", len(all_plans))

    return all_plans


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------


def export_all_plans_to_one_excel(
    plans,
    api,
    section_number,
    output_file,
):
    """
    Export the selected section from ALL DMPonline plans into ONE Excel file.

    New Excel layout:

        A1                  B1              C1              D1
        DMP Name + ID       Question 1      Question 2      Question 3

        DMP A + ID          Answer 1        Answer 2        Answer 3
        DMP B + ID          Answer 1        Answer 2        Answer 3
        DMP C + ID          Answer 1        Answer 2        Answer 3

    Each DMP gets one row.

    The questions are placed horizontally across the columns.
    """

    logging.info(
        "Creating combined Excel workbook for %s plans...",
        len(plans),
    )

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "All DMPs"

    # -----------------------------------------------------------------------
    # First retrieve all DMP data.
    #
    # We need to know ALL questions before creating the Excel columns,
    # because different DMPs may contain different questions.
    # -----------------------------------------------------------------------

    all_questions = []
    question_keys = set()

    dmp_data = []

    successful = 0
    failed = 0

    for index, plan in enumerate(plans, start=1):

        plan_id = plan.get("id")
        plan_title = plan.get("title") or f"DMP {plan_id}"

        logging.info(
            "Processing plan %s/%s: ID=%s | %s",
            index,
            len(plans),
            plan_id,
            plan_title,
        )

        if not plan_id:
            logging.error("Skipping plan without an ID.")
            failed += 1
            continue

        try:
            # Retrieve the complete DMP.
            parsed = api.get(
                request="v0/plans",
                params={
                    "plan": plan_id,
                    "remove_tests": "false",
                },
            )

            if not parsed:
                raise ValueError(
                    f"No DMP content found for plan ID {plan_id}."
                )

            try:
                full_plan = parsed[0]
                plan_content = full_plan["plan_content"][0]
                sections = plan_content["sections"]

            except (IndexError, KeyError, TypeError) as exc:
                raise ValueError(
                    f"Could not find DMP sections for plan {plan_id}."
                ) from exc

            selected_section = find_section(
                sections,
                section_number,
            )

            if selected_section is None:
                available_sections = [
                    str(section.get("number"))
                    for section in sections
                ]

                raise ValueError(
                    f"Section {section_number} was not found. "
                    f"Available sections: "
                    f"{', '.join(available_sections)}"
                )

            questions = selected_section.get("questions") or []

            actual_plan_title = (
                full_plan.get("title")
                or plan_title
            )

            # ---------------------------------------------------------------
            # Store the questions and answers for this DMP.
            # ---------------------------------------------------------------

            answers_by_question = {}

            for question in questions:

                question_number = question.get(
                    "number",
                    "",
                )

                question_text = question.get(
                    "text",
                    "",
                )

                answer_text = get_answer_text(question)

                # Create a unique column key.
                #
                # Including the question number makes sure that questions
                # with identical wording are still distinguishable.
                question_key = (
                    f"{question_number}. {question_text}"
                    if question_number
                    else question_text
                )

                answers_by_question[question_key] = answer_text

                # Add this question to the global list only once.
                if question_key not in question_keys:
                    question_keys.add(question_key)

                    all_questions.append(
                        {
                            "key": question_key,
                            "number": question_number,
                            "text": question_text,
                        }
                    )

            dmp_data.append(
                {
                    "plan_id": plan_id,
                    "plan_title": actual_plan_title,
                    "answers": answers_by_question,
                }
            )

            successful += 1

            logging.info(
                "Plan %s processed: %s questions.",
                plan_id,
                len(questions),
            )

        except Exception as exc:
            failed += 1

            logging.error(
                "Could not process plan %s (%s): %s",
                plan_id,
                plan_title,
                exc,
            )

            # Continue with the next DMP.
            continue

    # -----------------------------------------------------------------------
    # Create Excel headers
    # -----------------------------------------------------------------------

    # First column contains DMP name + ID.
    worksheet.cell(
        row=1,
        column=1,
        value="DMP Name + ID",
    )

    # Every question gets its own column.
    for column_index, question in enumerate(
        all_questions,
        start=2,
    ):
        worksheet.cell(
            row=1,
            column=column_index,
            value=question["key"],
        )

    # -----------------------------------------------------------------------
    # Add DMP rows
    # -----------------------------------------------------------------------

    for row_index, dmp in enumerate(
        dmp_data,
        start=2,
    ):

        plan_id = dmp["plan_id"]
        plan_title = dmp["plan_title"]
        answers = dmp["answers"]

        # DMP name + ID
        dmp_name_id = f"{plan_title} ({plan_id})"

        worksheet.cell(
            row=row_index,
            column=1,
            value=dmp_name_id,
        )

        # Answers
        for column_index, question in enumerate(
            all_questions,
            start=2,
        ):

            question_key = question["key"]

            answer = answers.get(
                question_key,
                "",
            )

            worksheet.cell(
                row=row_index,
                column=column_index,
                value=answer,
            )

    # -----------------------------------------------------------------------
    # Formatting
    # -----------------------------------------------------------------------

    format_combined_worksheet(
        worksheet,
        len(all_questions),
    )

    # Make sure the output directory exists.
    os.makedirs(
        os.path.dirname(os.path.abspath(output_file)),
        exist_ok=True,
    )

    workbook.save(output_file)

    logging.info(
        "Combined Excel report written to: %s",
        os.path.abspath(output_file),
    )

    logging.info(
        "Successfully exported: %s/%s plans.",
        successful,
        len(plans),
    )

    logging.info(
        "Failed: %s plans.",
        failed,
    )


# ---------------------------------------------------------------------------
# Excel formatting
# ---------------------------------------------------------------------------


def format_combined_worksheet(ws, number_of_questions):
    """
    Format the new horizontal DMP worksheet.
    """

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    header_font = Font(
        bold=True,
        color="FFFFFF",
    )

    dmp_font = Font(
        bold=True,
    )

    thin_border = Border(
        left=Side(
            style="thin",
            color="D9D9D9",
        ),
        right=Side(
            style="thin",
            color="D9D9D9",
        ),
        top=Side(
            style="thin",
            color="D9D9D9",
        ),
        bottom=Side(
            style="thin",
            color="D9D9D9",
        ),
    )

    # -----------------------------------------------------------------------
    # Header row
    # -----------------------------------------------------------------------

    for cell in ws[1]:

        cell.fill = header_fill

        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

        cell.border = thin_border

    ws.row_dimensions[1].height = 80

    # -----------------------------------------------------------------------
    # Data rows
    # -----------------------------------------------------------------------

    for row in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row,
    ):

        for cell in row:

            cell.alignment = Alignment(
                vertical="top",
                horizontal="left",
                wrap_text=True,
            )

            cell.border = thin_border

        # Make DMP name + ID bold.
        row[0].font = dmp_font

    # -----------------------------------------------------------------------
    # Column widths
    # -----------------------------------------------------------------------

    # DMP Name + ID
    ws.column_dimensions["A"].width = 40

    # Question columns
    #
    # The questions can be long, so 35 gives a reasonable starting point.
    # Users can still manually resize them in Excel.
    for column_index in range(
        2,
        number_of_questions + 2,
    ):

        column_letter = ws.cell(
            row=1,
            column=column_index,
        ).column_letter

        ws.column_dimensions[column_letter].width = 35

    # -----------------------------------------------------------------------
    # Row heights
    # -----------------------------------------------------------------------

    for row_number in range(
        2,
        ws.max_row + 1,
    ):

        ws.row_dimensions[row_number].height = 100

    # -----------------------------------------------------------------------
    # Freeze panes
    #
    # This keeps both the DMP name and question headers visible while
    # scrolling through the spreadsheet.
    # -----------------------------------------------------------------------

    ws.freeze_panes = "B2"

    # -----------------------------------------------------------------------
    # Filter
    # -----------------------------------------------------------------------

    if ws.max_row >= 2 and number_of_questions >= 1:

        last_column = ws.cell(
            row=1,
            column=number_of_questions + 1,
        ).column_letter

        ws.auto_filter.ref = (
            f"A1:{last_column}{ws.max_row}"
        )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():

    parser = argparse.ArgumentParser(
        description=(
            "Export a selected section from ALL DMPonline plans "
            "into ONE horizontally structured Excel file."
        )
    )

    # API token
    parser.add_argument(
        "-t",
        "--token",
        dest="token",
        required=True,
        help="DMPonline API token",
    )

    # User email
    parser.add_argument(
        "-u",
        "--user",
        dest="token_user",
        required=True,
        help="DMPonline user email address",
    )

    # Section number
    parser.add_argument(
        "-s",
        "--section",
        dest="section_number",
        required=True,
        help="Section number to export, e.g. 1, 2, 3, 4 or 5",
    )

    # Output file
    parser.add_argument(
        "-o",
        "--output",
        dest="output_file",
        default="DMP_all_plans.xlsx",
        help=(
            "Excel file to create. "
            "Default: DMP_all_plans.xlsx"
        ),
    )

    # Whether to exclude test plans.
    parser.add_argument(
        "--exclude-tests",
        action="store_true",
        help="Exclude DMPonline test plans.",
    )

    args = parser.parse_args()

    # -----------------------------------------------------------------------
    # Logging
    # -----------------------------------------------------------------------

    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    # -----------------------------------------------------------------------
    # Validate output extension
    # -----------------------------------------------------------------------

    if os.path.splitext(
        args.output_file
    )[1].lower() != ".xlsx":

        raise ValueError(
            "The output file must have an .xlsx extension."
        )

    # -----------------------------------------------------------------------
    # Create DMPonline API client
    # -----------------------------------------------------------------------

    logging.info("Connecting to DMPonline...")

    api = DMPonline(
        token=args.token,
        token_user=args.token_user,
    )

    # -----------------------------------------------------------------------
    # Get all plans
    # -----------------------------------------------------------------------

    plans = get_all_plans(
        api=api,
        include_tests=not args.exclude_tests,
    )

    if not plans:

        logging.warning(
            "No plans were returned by the DMPonline API."
        )

        return

    # -----------------------------------------------------------------------
    # Export all plans into ONE workbook
    # -----------------------------------------------------------------------

    export_all_plans_to_one_excel(
        plans=plans,
        api=api,
        section_number=args.section_number,
        output_file=args.output_file,
    )


# ---------------------------------------------------------------------------
# Program entry point
# ---------------------------------------------------------------------------


if __name__ == "__main__":
    main()

