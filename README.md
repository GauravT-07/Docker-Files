import re
from docx import Document
from openpyxl import Workbook, load_workbook


# ============================================================
# CONFIGURATION
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
    r"\bACDE-TOR-\d+\b",
    re.IGNORECASE
)


# ============================================================
# HEADER NORMALIZATION
# ============================================================

def normalize_header(text):
    """
    Converts variations such as:

        ID / TYPE
        ID/TYPE
        ID /TYPE
        ID/ TYPE

    into:

        ID/TYPE
    """

    return re.sub(r"\s+", "", str(text)).upper()


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
        # Find the "ID / TYPE" column
        # ----------------------------------------------------

        for row_number, row in enumerate(table.rows):

            for column_number, cell in enumerate(row.cells):

                cell_text = cell.text.strip()

                if normalize_header(cell_text) == "ID/TYPE":

                    id_column_index = column_number

                    print(
                        f"Found ID / TYPE column in "
                        f"Table {table_number}, "
                        f"Row {row_number + 1}, "
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
        # Read TOR IDs from the identified column
        # ----------------------------------------------------

        for row in table.rows:

            if id_column_index >= len(row.cells):
                continue

            cell_text = row.cells[
                id_column_index
            ].text.strip()

            matches = TOR_PATTERN.findall(
                cell_text
            )

            for tor_id in matches:

                tor_ids.add(
                    tor_id.upper()
                )

    return tor_ids


# ============================================================
# READ TOR IDs FROM EXISTING EXCEL
# ============================================================

def read_tor_ids_from_excel(file_path):

    workbook = load_workbook(
        file_path,
        read_only=True,
        data_only=True
    )

    tor_ids = set()

    for worksheet in workbook.worksheets:

        # TOR ID is in the FIRST COLUMN
        for row in worksheet.iter_rows(
            min_col=1,
            max_col=1
        ):

            value = row[0].value

            if value is None:
                continue

            value = str(value).strip()

            matches = TOR_PATTERN.findall(
                value
            )

            for tor_id in matches:

                tor_ids.add(
                    tor_id.upper()
                )

    workbook.close()

    return tor_ids


# ============================================================
# CREATE EXCEL WITH EXTRACTED TOR IDs
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
    current_ids,
    output_file
):

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "TOR Comparison"

    worksheet.append([
        "TOR ID",
        "Status"
    ])

    # Combine IDs from both files
    all_tor_ids = sorted(
        existing_ids | current_ids
    )

    for tor_id in all_tor_ids:

        # -----------------------------------------------
        # Present in BOTH
        # -----------------------------------------------

        if (
            tor_id in existing_ids
            and tor_id in current_ids
        ):

            status = "Present in both"

        # -----------------------------------------------
        # Existing Excel ONLY
        # -----------------------------------------------

        elif tor_id in existing_ids:

            status = (
                "Requirement and test case "
                "needs to be deleted"
            )

        # -----------------------------------------------
        # Current/Extracted Excel ONLY
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
    current_ids,
    output_file
):

    # Existing Excel but NOT current/extracted Excel
    deleted_requirements = sorted(
        existing_ids - current_ids
    )

    # Current/extracted Excel but NOT existing Excel
    missing_requirements = sorted(
        current_ids - existing_ids
    )

    # Present in both
    common_requirements = sorted(
        existing_ids & current_ids
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
        # Existing only
        # ------------------------------------------------

        file.write(
            "1. REQUIREMENT AND TEST CASE NEEDS TO BE DELETED\n"
        )

        file.write(
            "-" * 60 + "\n"
        )

        if deleted_requirements:

            for tor_id in deleted_requirements:

                file.write(
                    f"{tor_id}\n"
                )

        else:

            file.write(
                "None\n"
            )

        file.write("\n")

        # ------------------------------------------------
        # Current only
        # ------------------------------------------------

        file.write(
            "2. REQUIREMENT IS MISSING\n"
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

            file.write(
                "None\n"
            )

        file.write("\n")

        # ------------------------------------------------
        # Common
        # ------------------------------------------------

        file.write(
            "3. PRESENT IN BOTH\n"
        )

        file.write(
            "-" * 60 + "\n"
        )

        if common_requirements:

            for tor_id in common_requirements:

                file.write(
                    f"{tor_id}\n"
                )

        else:

            file.write(
                "None\n"
            )


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

    current_ids = extract_tor_ids_from_word(
        WORD_FILE
    )

    print(
        f"Found {len(current_ids)} unique TOR IDs "
        f"in Word document."
    )

    # --------------------------------------------------------
    # 2. Create extracted TOR Excel
    # --------------------------------------------------------

    create_extracted_excel(
        current_ids,
        EXTRACTED_EXCEL
    )

    print(
        f"Created: {EXTRACTED_EXCEL}"
    )

    # --------------------------------------------------------
    # 3. Read TOR IDs from existing Excel
    # --------------------------------------------------------

    print("\nReading existing Excel...")

    existing_ids = read_tor_ids_from_excel(
        EXISTING_EXCEL
    )

    print(
        f"Found {len(existing_ids)} unique TOR IDs "
        f"in existing Excel."
    )

    # --------------------------------------------------------
    # 4. Compare both files
    # --------------------------------------------------------

    create_comparison_excel(
        existing_ids,
        current_ids,
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
        current_ids,
        TEXT_REPORT
    )

    print(
        f"Created: {TEXT_REPORT}"
    )

    # --------------------------------------------------------
    # 6. Print summary
    # --------------------------------------------------------

    common_count = len(
        existing_ids & current_ids
    )

    deleted_count = len(
        existing_ids - current_ids
    )

    missing_count = len(
        current_ids - existing_ids
    )

    print("\n" + "=" * 60)
    print("COMPARISON SUMMARY")
    print("=" * 60)

    print(
        f"Present in both: {common_count}"
    )

    print(
        "Existing Excel only "
        "(Requirement and test case needs to be deleted): "
        f"{deleted_count}"
    )

    print(
        "Current/Extracted Excel only "
        "(Requirement is missing): "
        f"{missing_count}"
    )

    print("=" * 60)
