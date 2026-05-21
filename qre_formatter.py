import pandas as pd
import re
from datetime import datetime
import os

# =====================================
# PROCESS FUNCTION
# =====================================

def process_qre(input_file):

    # =====================================
    # OUTPUT FOLDER
    # =====================================

    output_folder = "outputs"

    os.makedirs(output_folder, exist_ok=True)

    # =====================================
    # OUTPUT FILE NAME
    # =====================================

    output_file = os.path.join(
        output_folder,
        f"formatted_qre_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    # =====================================
    # READ EXCEL
    # =====================================

    df = pd.read_excel(input_file, header=None)

    # Skip first 2 template/header rows
    df = df.iloc[2:].reset_index(drop=True)

    # Remove fully blank rows
    df = df.dropna(how="all")

    # =====================================
    # OUTPUT COLUMNS
    # =====================================

    final_columns = [
        "Type ID",
        "Text",
        "Comment for Shopper",
        "Answer",
        "Could this be Not Applicable?",
        "Is answer Required?",
        "Question is Numbered?",
        "Answers Visualization:",
        "Comment:",
        "Comment Content Rule:",
        "Comment Requirement Rule:",
        "On Export Show Only the Selected Answer",
        "Grid title",
        "Grid Hidden Comment",
        "Grid Rows",
        "Grid Cols",
        "Grid Type",
        "Visibility Exceptions",
        "Visibility Default Mode",
        "Question Category 1",
        "Question Category 2",
        "Question Category 3"
    ]

    output_rows = []

    # =====================================
    # HELPER FUNCTIONS
    # =====================================

    def blank_row():
        return {col: "" for col in final_columns}

    def clean_text(text):
        return re.sub(r"\s+", " ", str(text)).strip()

    # =====================================
    # NOTE EXTRACTION
    # =====================================

    def extract_note(question):

        question = str(question)

        note = ""
        clean_question = question

        note_keywords = [
            "NOTE:",
            "Note:",
            "NOTES:",
            "Important:",
            "Instructions:"
        ]

        pattern = r"(.*?)(?:" + "|".join(note_keywords) + r")(.*)"

        match = re.search(
            pattern,
            question,
            re.IGNORECASE
        )

        if match:

            clean_question = clean_text(match.group(1))
            note = clean_text(match.group(2))

        return clean_question, note

    # =====================================
    # QUESTION TRACKERS
    # =====================================

    current_question = None
    current_answers = []
    current_format = None

    # =====================================
    # SAVE QUESTION FUNCTION
    # =====================================

    def save_question():

        nonlocal current_question
        nonlocal current_answers
        nonlocal current_format

        if not current_question:
            return

        question, note = extract_note(current_question)

        answers = "|".join([
            clean_text(a)
            for a in current_answers
            if pd.notna(a)
        ])

        comment_val = "4" if str(current_format).strip().lower() in [
            "text",
            "numeric",
            "monetary numeric"
        ] else "1"

        row = blank_row()

        row["Type ID"] = "Q"
        row["Text"] = question
        row["Comment for Shopper"] = note
        row["Answer"] = answers
        row["Could this be Not Applicable?"] = "n"
        row["Is answer Required?"] = "y"
        row["Question is Numbered?"] = "n"
        row["Answers Visualization:"] = "1"
        row["Comment:"] = comment_val
        row["Comment Content Rule:"] = "1"
        row["Comment Requirement Rule:"] = "1"
        row["On Export Show Only the Selected Answer"] = "y"

        output_rows.append(row)

    # =====================================
    # MAIN LOOP
    # =====================================

    for _, row in df.iterrows():

        qid = row[0] if len(row) > 0 else None
        question_text = row[1] if len(row) > 1 else None
        format_type = row[2] if len(row) > 2 else None
        answer = row[4] if len(row) > 4 else None

        # =====================================
        # DETECT SECTION ROW
        # =====================================

        is_section = (
            pd.notna(qid)
            and (
                pd.isna(question_text)
                or str(question_text).strip() == ""
            )
            and pd.isna(answer)
        )

        # =====================================
        # SECTION ROW
        # =====================================

        if is_section:

            save_question()

            # Add E only if output already contains data
            if len(output_rows) > 0:

                end_row = blank_row()
                end_row["Type ID"] = "E"

                output_rows.append(end_row)

            # =====================================
            # BUILD SECTION NAME
            # =====================================

            section_parts = []

            for val in [qid, question_text]:

                if pd.notna(val):

                    clean_val = clean_text(val)

                    if clean_val:

                        section_parts.append(clean_val)

            section_name = " ".join(section_parts)

            # Remove SECTION word if present
            section_name = re.sub(
                r"^SECTION\s*",
                "",
                section_name,
                flags=re.IGNORECASE
            )

            # =====================================
            # ADD SECTION ROW
            # =====================================

            section_row = blank_row()

            section_row["Type ID"] = "S"
            section_row["Text"] = section_name

            output_rows.append(section_row)

            current_question = None
            current_answers = []
            current_format = None

            continue

        # =====================================
        # QUESTION ROW
        # =====================================

        if pd.notna(question_text):

            save_question()

            current_question = clean_text(question_text)

            current_answers = []
            current_format = format_type

            if pd.notna(answer):

                current_answers.append(
                    clean_text(answer)
                )

        # =====================================
        # ADDITIONAL ANSWERS
        # =====================================

        elif pd.isna(question_text) and pd.notna(answer):

            current_answers.append(
                clean_text(answer)
            )

    # =====================================
    # SAVE LAST QUESTION
    # =====================================

    save_question()

    end_row = blank_row()
    end_row["Type ID"] = "E"

    output_rows.append(end_row)

    # =====================================
    # EXPORT OUTPUT
    # =====================================

    output_df = pd.DataFrame(
        output_rows,
        columns=final_columns
    )

    output_df.to_excel(
        output_file,
        index=False
    )

    return output_file