import re
from docx import Document
from openpyxl import Workbook, load_workbook


# ============================================================
# FILE CONFIGURATION
# ============================================================

WORD_FILE = "requirements.docx"
EXISTING_EXCEL = "existing_requirements.xlsx"

EXTRACTED_EXCEL = "extracted_tor_ids.xlsx"
COMPARISON_EXCEL = "tor_comparison.xlsx"
TEXT_REPORT = "tor_comparison.txt"


# ============================================================
# TOR ID PATTERN
# ============================================================

TOR_PATTERN = re.compile(
    r"^ACDE-TOR-\d+$",
    re.IGNORECASE
)


# ============================================================
# NORMALIZE TOR ID
# ============================================================

def normalize_tor_id(value):
    """
    Convert the value into a standard TOR ID.

    Example:
        acde-tor-123  -> ACDE-TOR-123
        ACDE-TOR-123  -> ACDE-TOR-123

    Returns None if the value is not a valid TOR ID.
    """

    if value is None:
        return None

    tor_id = str(value).strip().upper()

    if TOR_PATTERN.fullmatch(tor_id):
        return tor_id

    return None


# ============================================================
# EXTRACT TOR IDs FROM WORD DOCUMENT
# ============================================================

def extract_tor_ids_from_word(file_path):

    document = Document(file_path)

    tor_ids = set()

    for table_number, table in enumerate(
        document.tables,
        start=1
    ):

        id_column_index = None

        # ----------------------------------------------------
        # Find ID / TYPE column
        # ----------------------------------------------------

        for row_number, row in enumerate(table.rows):

            for column_number, cell in enumerate(row.cells):

                header = re.sub(
                    r"\s+",
                    "",
                    cell.text
                ).upper()

                if header == "ID/TYPE":

                    id_column_index = column_number

                    print(
                        f"Found ID / TYPE column: "
                        f"Table {table_number}, "
                        f"Column {column_number + 1}"
                    )

                    break

            if id_column_index is not None:
                break

        # ----------------------------------------------------
        # Ignore tables without ID / TYPE
        # ----------------------------------------------------

        if id_column_index is None:
            continue

        # ----------------------------------------------------
        # Read TOR IDs from ID / TYPE column
        # ----------------------------------------------------

        for row in table.rows:

            if id_column_index >= len(row.cells):
                continue

            cell_text = row.cells[
                id_column_index
            ].text.strip()

            # Find ACDE-TOR-123 anywhere in the cell
            matches = re.findall(
                r"\bACDE-TOR-\d+\b",
                cell_text,
                re.IGNORECASE
            )

            for match in matches:

                tor_ids.add(
                    match.upper()
                )

    return tor_ids


# ============================================================
# READ TOR IDs FROM EXISTING EXCEL
# ============================================================

def read_tor_ids_from_existing_excel(file_path):

    workbook = load_workbook(
        file_path,
        read_only=True,
        data_only=True
    )

    tor_ids = set()

    for worksheet in workbook.worksheets:

        # ----------------------------------------------------
        # ONLY READ COLUMN A
        # ----------------------------------------------------

        for row in worksheet.iter_rows(
            min_col=1,
            max_col=1
        ):

            value = row[0].value

            tor_id = normalize_tor_id(value)

            if tor_id is not None:
                tor_ids.add(tor_id)

    workbook.close()

    return tor_ids


# ============================================================
# CREATE EXTRACTED TOR EXCEL
# ============================================================

def create_extracted_excel(
    tor_ids,
    output_file
):

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "TOR IDs"

    worksheet.append([
        "TOR ID"
    ])

    for tor_id in sorted(tor_ids):

        worksheet.append([
            tor_id
        ])

    workbook.save(output_file)


# ============================================================
# CREATE COMPARISON EXCEL
# ============================================================

def create_comparison_excel(
    existing_ids,
    extracted_ids,
    output_file
):

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "TOR Comparison"

    worksheet.append([
        "TOR ID",
        "Status"
    ])

    # Combine all unique TOR IDs
    all_tor_ids = sorted(
        existing_ids | extracted_ids
    )

    for tor_id in all_tor_ids:

        # -----------------------------------------------
        # Present in both
        # -----------------------------------------------

        if (
            tor_id in existing_ids
            and tor_id in extracted_ids
        ):

            status = "Present in both"

        # -----------------------------------------------
        # Existing Excel only
        # -----------------------------------------------

        elif tor_id in existing_ids:

            status = (
                "Requirement and test case "
                "needs to be deleted"
            )

        # -----------------------------------------------
        # Extracted Excel only
        # -----------------------------------------------

        else:

            status = "Requirement is missing"

        worksheet.append([
            tor_id,
            status
        ])

    workbook.save(output_file)


# ============================================================
# CREATE TEXT REPORT
# ============================================================

def create_text_report(
    existing_ids,
    extracted_ids,
    output_file
):

    # Existing Excel but NOT extracted/current
    delete_requirements = sorted(
        existing_ids - extracted_ids
    )

    # Extracted/current but NOT existing Excel
    missing_requirements = sorted(
        extracted_ids - existing_ids
    )

    with open(
        output_file,
        "w",
        encoding="utf-8"
    ) as file:

        file.write(
            "TOR REQUIREMENT COMPARISON REPORT\n"
        )

        file.write(
            "=" * 60 + "\n\n"
        )

        # ------------------------------------------------
        # Existing Excel only
        # ------------------------------------------------

        file.write(
            "REQUIREMENT AND TEST CASE NEEDS TO BE DELETED\n"
        )

        file.write(
            "-" * 60 + "\n"
        )

        if delete_requirements:

            for tor_id in delete_requirements:
                file.write(
                    f"{tor_id}\n"
                )

        else:

            file.write("None\n")

        file.write("\n\n")

        # ------------------------------------------------
        # Extracted Excel only
        # ------------------------------------------------

        file.write(
            "REQUIREMENT IS MISSING\n"
        )

        file.write(
            "-" * 60 + "\n"
        )

        if missing_requirements:

            for tor_id in missing_requirements:
                file.write(
                    f"{tor_id}\n"
                )

        else:

            file.write("None\n")


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    print("=" * 60)
    print("TOR REQUIREMENT COMPARISON")
    print("=" * 60)

    # --------------------------------------------------------
    # 1. Extract TOR IDs from Word
    # --------------------------------------------------------

    print("\nReading Word document...")

    extracted_ids = extract_tor_ids_from_word(
        WORD_FILE
    )

    print(
        f"Extracted {len(extracted_ids)} unique TOR IDs."
    )

    # --------------------------------------------------------
    # 2. Create Excel containing extracted TOR IDs
    # --------------------------------------------------------

    create_extracted_excel(
        extracted_ids,
        EXTRACTED_EXCEL
    )

    print(
        f"Created: {EXTRACTED_EXCEL}"
    )

    # --------------------------------------------------------
    # 3. Read existing Excel - ONLY COLUMN A
    # --------------------------------------------------------

    print("\nReading existing Excel Column A...")

    existing_ids = read_tor_ids_from_existing_excel(
        EXISTING_EXCEL
    )

    print(
        f"Found {len(existing_ids)} unique TOR IDs "
        f"in existing Excel."
    )

    # --------------------------------------------------------
    # 4. Compare
    # --------------------------------------------------------

    create_comparison_excel(
        existing_ids,
        extracted_ids,
        COMPARISON_EXCEL
    )

    print(
        f"Created: {COMPARISON_EXCEL}"
    )

    # --------------------------------------------------------
    # 5. Create text report
    # --------------------------------------------------------

    create_text_report(
        existing_ids,
        extracted_ids,
        TEXT_REPORT
    )

    print(
        f"Created: {TEXT_REPORT}"
    )

    # --------------------------------------------------------
    # 6. Summary
    # --------------------------------------------------------

    present_in_both = (
        existing_ids & extracted_ids
    )

    delete_requirements = (
        existing_ids - extracted_ids
    )

    missing_requirements = (
        extracted_ids - existing_ids
    )

    print("\n" + "=" * 60)
    print("COMPARISON SUMMARY")
    print("=" * 60)

    print(
        f"Present in both: "
        f"{len(present_in_both)}"
    )

    print(
        f"Existing Excel only "
        f"(delete): "
        f"{len(delete_requirements)}"
    )

    print(
        f"Extracted Excel only "
        f"(missing): "
        f"{len(missing_requirements)}"
    )

    print("=" * 60)
