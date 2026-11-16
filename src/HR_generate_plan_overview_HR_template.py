import argparse
import logging
import os
import html
from html.parser import HTMLParser

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from dmponline import DMPonline


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Only DMPs using this template will be exported.
TARGET_TEMPLATE = "Hogeschool Rotterdam (HR) Template"


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------


class _HTMLTextExtractor(HTMLParser):
    """Extract visible text from HTML while keeping table/list items readable."""

    BLOCK_TAGS = {
        "br", "p", "div", "tr", "td", "th", "li", "ul", "ol",
        "table", "h1", "h2", "h3", "h4", "h5", "h6"
    }

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.skip_depth = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self.skip_depth += 1
        elif tag in self.BLOCK_TAGS and self.parts:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style") and self.skip_depth:
            self.skip_depth -= 1
        elif tag in self.BLOCK_TAGS:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self.skip_depth and data.strip():
            self.parts.append(data.strip())


def clean_answer_text(value):
    """Turn plain text or HTML answer content into clean readable text."""
    if value is None:
        return ""

    text = html.unescape(str(value))

    # Only invoke the HTML parser when the value appears to contain HTML.
    if "<" in text and ">" in text:
        parser = _HTMLTextExtractor()
        try:
            parser.feed(text)
            text = " ".join(parser.parts)
        except Exception:
            # If parsing fails, retain the original value rather than lose data.
            pass

    # Normalize whitespace while retaining useful line breaks.
    lines = [" ".join(line.split()) for line in text.splitlines()]
    return "\n".join(line for line in lines if line).strip()


def get_answer_text(question):
    """
    Convert the answer structure returned by DMPonline into readable text.

    HTML markup is removed so selected answers are exported as visible text.
    Handles free text, scalar values, selected options, lists and dictionaries.
    """

    if not question.get("answered"):
        return ""

    answer = question.get("answer")

    if answer is None:
        return ""

    if isinstance(answer, str):
        return clean_answer_text(answer)

    if isinstance(answer, (int, float, bool)):
        return str(answer)

    if isinstance(answer, dict):
        options = answer.get("options")

        if options:
            option_texts = []
            for option in options:
                if isinstance(option, dict):
                    value = (
                        option.get("text")
                        or option.get("label")
                        or option.get("name")
                        or option.get("value")
                    )
                else:
                    value = option

                if value is not None:
                    cleaned = clean_answer_text(value)
                    if cleaned:
                        option_texts.append(cleaned)

            if option_texts:
                return "\n".join(option_texts)

        for key in ("text", "value", "answer", "content", "response"):
            value = answer.get(key)
            if value is not None:
                return clean_answer_text(value)

        return clean_answer_text(answer)

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
                    cleaned = clean_answer_text(value)
                    if cleaned:
                        values.append(cleaned)
            else:
                cleaned = clean_answer_text(item)
                if cleaned:
                    values.append(cleaned)
        return "\n".join(values)

    return clean_answer_text(answer)


def get_comment_text(question):
    """Extract comment text from common DMPonline comment structures."""

    def extract_text(value):
        if value is None:
            return ""
        if isinstance(value, str):
            return value.strip()
        if isinstance(value, (int, float, bool)):
            return str(value)
        if isinstance(value, (list, tuple)):
            parts = [extract_text(item) for item in value]
            return "\n".join(part for part in parts if part)
        if isinstance(value, dict):
            for key in ("text", "comment", "content", "body", "value", "message"):
                if value.get(key) is not None:
                    return extract_text(value[key])
            for key in ("comments", "items", "data"):
                if value.get(key) is not None:
                    return extract_text(value[key])
        return ""

    found_comments = []

    def find_comment_fields(value):
        if isinstance(value, dict):
            for key, nested_value in value.items():
                if key.lower() in ("comments", "comment"):
                    text = extract_text(nested_value)
                    if text and text not in found_comments:
                        found_comments.append(text)
                elif isinstance(nested_value, (dict, list, tuple)):
                    find_comment_fields(nested_value)
        elif isinstance(value, (list, tuple)):
            for item in value:
                find_comment_fields(item)

    find_comment_fields(question)
    return "\n\n".join(found_comments)


# ---------------------------------------------------------------------------
# Template helper
# ---------------------------------------------------------------------------


def get_template_name(plan):
    """
    Try to retrieve the template name from a DMPonline plan.

    DMPonline responses can contain the template information in slightly
    different structures, so this function checks several possibilities.

    Returns:
        str: Template name if found, otherwise "".
    """

    template = plan.get("template")

    if template is None:
        return ""

    # ---------------------------------------------------------------
    # Template is simply a string
    # ---------------------------------------------------------------

    if isinstance(template, str):
        return template.strip()

    # ---------------------------------------------------------------
    # Template is a dictionary
    # ---------------------------------------------------------------

    if isinstance(template, dict):

        # Most likely fields
        for key in (
            "title",
            "name",
            "text",
            "label",
        ):

            value = template.get(key)

            if value:
                return str(value).strip()

        # Sometimes the name may be nested
        nested_template = template.get("template")

        if isinstance(nested_template, dict):

            for key in (
                "title",
                "name",
                "text",
                "label",
            ):

                value = nested_template.get(key)

                if value:
                    return str(value).strip()

        # Sometimes template information can contain an ID and title
        # inside another object.
        for value in template.values():

            if isinstance(value, dict):

                for key in (
                    "title",
                    "name",
                    "text",
                    "label",
                ):

                    nested_value = value.get(key)

                    if nested_value:
                        return str(nested_value).strip()

    return ""


def plan_uses_target_template(plan, target_template):
    """
    Check whether a DMP uses the requested template.

    Matching is case-insensitive and ignores leading/trailing whitespace.
    """

    template_name = get_template_name(plan)

    if not template_name:
        return False

    return (
        template_name.strip().lower()
        == target_template.strip().lower()
    )


# ---------------------------------------------------------------------------
# Retrieve all DMPs
# ---------------------------------------------------------------------------


def get_all_plans(api, include_tests=True):
    """
    Retrieve all DMP plans available to the authenticated DMPonline
    organisation administrator.

    The DMPonline v0 API paginates the plans endpoint.
    """

    logging.info(
        "Retrieving all DMPonline plans..."
    )

    all_plans = []

    page = 1

    while True:

        logging.info(
            "Retrieving plans page %s...",
            page,
        )

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
            "Page %s: found %s plans "
            "(total so far: %s).",
            page,
            len(page_plans),
            len(all_plans),
        )

        page += 1

    logging.info(
        "Total plans found: %s",
        len(all_plans),
    )

    return all_plans


# ---------------------------------------------------------------------------
# Filter DMPs by template
# ---------------------------------------------------------------------------


def filter_plans_by_template(plans, target_template):
    """
    Return only plans that use the requested DMPonline template.
    """

    logging.info(
        "Filtering plans using template: %s",
        target_template,
    )

    matching_plans = []

    template_counts = {}

    for plan in plans:

        template_name = get_template_name(plan)

        if template_name:

            template_counts[template_name] = (
                template_counts.get(template_name, 0) + 1
            )

        if plan_uses_target_template(
            plan,
            target_template,
        ):

            matching_plans.append(plan)

    # ---------------------------------------------------------------
    # Logging
    # ---------------------------------------------------------------

    logging.info(
        "Found %s plans using template '%s'.",
        len(matching_plans),
        target_template,
    )

    logging.info(
        "Templates found in retrieved plans:"
    )

    if template_counts:

        for template_name, count in sorted(
            template_counts.items()
        ):

            logging.info(
                "  %s: %s plans",
                template_name,
                count,
            )

    else:

        logging.warning(
            "No template information was found "
            "in the plan responses."
        )

    return matching_plans


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------


def export_all_plans_to_one_excel(
    plans,
    api,
    output_file,
):
    """
    Export all sections from the filtered DMPonline plans
    into ONE Excel file.

    Excel structure:

        DMP Name + ID | Question 1 | Question 2 | Question 3
        -----------------------------------------------------
        DMP A (123)   | Answer 1   | Answer 2   | Answer 3
        DMP B (456)   | Answer 1   | Answer 2   | Answer 3

    Each DMP gets one row.

    Each question gets one column.

    The answer is placed directly under its question.
    """

    logging.info(
        "Creating combined Excel workbook for %s plans...",
        len(plans),
    )

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "All DMPs"

    # -----------------------------------------------------------------------
    # Collect all DMP information
    # -----------------------------------------------------------------------

    all_questions = []

    question_keys = set()

    dmp_data = []

    successful = 0

    failed = 0

    # -----------------------------------------------------------------------
    # Process every DMP
    # -----------------------------------------------------------------------

    for index, plan in enumerate(
        plans,
        start=1,
    ):

        plan_id = plan.get("id")

        plan_title = (
            plan.get("title")
            or f"DMP {plan_id}"
        )

        template_name = get_template_name(plan)

        logging.info(
            "Processing plan %s/%s: ID=%s | %s | Template=%s",
            index,
            len(plans),
            plan_id,
            plan_title,
            template_name,
        )

        if not plan_id:

            logging.error(
                "Skipping plan without an ID."
            )

            failed += 1

            continue

        try:

            # ---------------------------------------------------------------
            # Retrieve the complete DMP
            # ---------------------------------------------------------------

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

            # ---------------------------------------------------------------
            # Extract plan content
            # ---------------------------------------------------------------

            try:

                full_plan = parsed[0]

                plan_content = full_plan[
                    "plan_content"
                ][0]

                sections = plan_content[
                    "sections"
                ]

            except (
                IndexError,
                KeyError,
                TypeError,
            ) as exc:

                raise ValueError(
                    f"Could not find DMP sections "
                    f"for plan {plan_id}."
                ) from exc

            # ---------------------------------------------------------------
            # Process questions from every section
            # ---------------------------------------------------------------

            actual_plan_title = (
                full_plan.get("title")
                or plan_title
            )

            answers_by_question = {}
            total_questions = 0

            for section in sections:

                questions = section.get("questions") or []

                for question in questions:

                    total_questions += 1

                    # The question text already contains its display number
                    # (for example, "1.1 ..."). Do not add another number.
                    question_key = str(
                        question.get("text") or ""
                    ).strip()

                    if not question_key:
                        question_key = str(
                            question.get("number") or "Unnamed question"
                        ).strip()

                    answer_text = get_answer_text(question)
                    comment_text = get_comment_text(question)

                    if comment_text:
                        if answer_text:
                            answer_text = (
                                f"{answer_text}\n\nComments:\n{comment_text}"
                            )
                        else:
                            answer_text = f"Comments:\n{comment_text}"

                    answers_by_question[question_key] = answer_text

                    if question_key not in question_keys:
                        question_keys.add(question_key)
                        all_questions.append(question_key)

            # ---------------------------------------------------------------
            # Store the DMP
            # ---------------------------------------------------------------

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
                total_questions,
            )

        except Exception as exc:

            failed += 1

            logging.error(
                "Could not process plan %s (%s): %s",
                plan_id,
                plan_title,
                exc,
            )

            # Continue processing other DMPs.
            continue

    # -----------------------------------------------------------------------
    # Create the header row
    # -----------------------------------------------------------------------

    # Column A:
    # DMP name + ID

    worksheet.cell(
        row=1,
        column=1,
        value="DMP Name + ID",
    )

    # Columns B onward:
    # One question per column.

    for column_index, question in enumerate(
        all_questions,
        start=2,
    ):

        worksheet.cell(
            row=1,
            column=column_index,
            value=question,
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

        # ---------------------------------------------------------------
        # Column A = DMP name + ID
        # ---------------------------------------------------------------

        dmp_name_id = (
            f"{plan_title} ({plan_id})"
        )

        worksheet.cell(
            row=row_index,
            column=1,
            value=dmp_name_id,
        )

        # ---------------------------------------------------------------
        # Columns B onward = answers
        # ---------------------------------------------------------------

        for column_index, question in enumerate(
            all_questions,
            start=2,
        ):

            answer = answers.get(
                question,
                "",
            )

            worksheet.cell(
                row=row_index,
                column=column_index,
                value=answer,
            )

    # -----------------------------------------------------------------------
    # Format worksheet
    # -----------------------------------------------------------------------

    format_combined_worksheet(
        worksheet,
        len(all_questions),
    )

    # -----------------------------------------------------------------------
    # Create output directory
    # -----------------------------------------------------------------------

    os.makedirs(
        os.path.dirname(
            os.path.abspath(output_file)
        ),
        exist_ok=True,
    )

    # -----------------------------------------------------------------------
    # Save workbook
    # -----------------------------------------------------------------------

    workbook.save(
        output_file
    )

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


def format_combined_worksheet(
    ws,
    number_of_questions,
):
    """
    Format the horizontally structured worksheet.
    """

    # -----------------------------------------------------------------------
    # Styles
    # -----------------------------------------------------------------------

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

    ws.row_dimensions[1].height = 100

    # -----------------------------------------------------------------------
    # Data rows
    # -----------------------------------------------------------------------

    for row in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row,
    ):

        for cell in row:

            cell.alignment = Alignment(
                horizontal="left",
                vertical="top",
                wrap_text=True,
            )

            cell.border = thin_border

        # DMP name + ID is bold.
        row[0].font = dmp_font

    # -----------------------------------------------------------------------
    # Column widths
    # -----------------------------------------------------------------------

    ws.column_dimensions["A"].width = 40

    for column_index in range(
        2,
        number_of_questions + 2,
    ):

        column_letter = ws.cell(
            row=1,
            column=column_index,
        ).column_letter

        ws.column_dimensions[
            column_letter
        ].width = 35

    # -----------------------------------------------------------------------
    # Row heights
    # -----------------------------------------------------------------------

    for row_number in range(
        2,
        ws.max_row + 1,
    ):

        ws.row_dimensions[
            row_number
        ].height = 100

    # -----------------------------------------------------------------------
    # Freeze panes
    # -----------------------------------------------------------------------

    ws.freeze_panes = "B2"

    # -----------------------------------------------------------------------
    # Auto filter
    # -----------------------------------------------------------------------

    if (
        ws.max_row >= 2
        and number_of_questions >= 1
    ):

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
            "Export all sections from "
            "DMPonline plans using the "
            f"'{TARGET_TEMPLATE}' template "
            "into ONE Excel file."
        )
    )

    # -----------------------------------------------------------------------
    # API token
    # -----------------------------------------------------------------------

    parser.add_argument(
        "-t",
        "--token",
        dest="token",
        required=True,
        help="DMPonline API token",
    )

    # -----------------------------------------------------------------------
    # User email
    # -----------------------------------------------------------------------

    parser.add_argument(
        "-u",
        "--user",
        dest="token_user",
        required=True,
        help="DMPonline user email address",
    )

    # -----------------------------------------------------------------------
    # Output file
    # -----------------------------------------------------------------------

    parser.add_argument(
        "-o",
        "--output",
        dest="output_file",
        default="DMP_HR_Template_plans.xlsx",
        help=(
            "Excel file to create. "
            "Default: DMP_HR_Template_plans.xlsx"
        ),
    )

    # -----------------------------------------------------------------------
    # Exclude test plans
    # -----------------------------------------------------------------------

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
            "The output file must have "
            "an .xlsx extension."
        )

    # -----------------------------------------------------------------------
    # Create DMPonline API client
    # -----------------------------------------------------------------------

    logging.info(
        "Connecting to DMPonline..."
    )

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
            "No plans were returned by "
            "the DMPonline API."
        )

        return

    # -----------------------------------------------------------------------
    # Filter by template
    # -----------------------------------------------------------------------

    filtered_plans = filter_plans_by_template(
        plans=plans,
        target_template=TARGET_TEMPLATE,
    )

    if not filtered_plans:

        logging.warning(
            "No DMPs were found using the template: %s",
            TARGET_TEMPLATE,
        )

        return

    # -----------------------------------------------------------------------
    # Export filtered plans into ONE workbook
    # -----------------------------------------------------------------------

    export_all_plans_to_one_excel(
        plans=filtered_plans,
        api=api,
        output_file=args.output_file,
    )


# ---------------------------------------------------------------------------
# Program entry point
# ---------------------------------------------------------------------------


if __name__ == "__main__":
    main()