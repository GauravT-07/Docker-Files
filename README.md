import re
from docx import Document


# ACDE-TOR followed by one or more digits
REQUIREMENT_PATTERN = re.compile(
    r"\bACDE-TOR-\d+\b",
    re.IGNORECASE
)


def normalize_header(text):
    """
    Normalize a table header so that variations like:

        ID / TYPE
        ID/TYPE
        ID /TYPE
        ID/ TYPE

    are treated as the same header.
    """
    return re.sub(r"\s+", "", text).upper()


def extract_requirement_ids(file_path):
    document = Document(file_path)

    requirement_ids = set()

    for table_number, table in enumerate(document.tables, start=1):

        id_column_index = None

        # ---------------------------------------------------------
        # Step 1: Find the column containing "ID / TYPE"
        # Search every row instead of assuming row 0 is the header.
        # ---------------------------------------------------------

        for row_number, row in enumerate(table.rows):

            for column_number, cell in enumerate(row.cells):

                cell_text = cell.text.strip()

                normalized = normalize_header(cell_text)

                if normalized == "ID/TYPE":
                    id_column_index = column_number

                    print(
                        f"Found 'ID / TYPE' in "
                        f"table {table_number}, "
                        f"row {row_number + 1}, "
                        f"column {column_number + 1}"
                    )

                    break

            if id_column_index is not None:
                break

        # ---------------------------------------------------------
        # Step 2: If this table does not contain ID / TYPE,
        # ignore the table.
        # ---------------------------------------------------------

        if id_column_index is None:
            continue

        # ---------------------------------------------------------
        # Step 3: Read every cell below the ID / TYPE column
        # ---------------------------------------------------------

        for row_number, row in enumerate(table.rows):

            # Make sure the column exists in this row
            if id_column_index >= len(row.cells):
                continue

            cell_text = row.cells[id_column_index].text.strip()

            # Find all ACDE-TOR-XXXXX patterns in the cell
            matches = REQUIREMENT_PATTERN.findall(cell_text)

            for requirement_id in matches:
                requirement_ids.add(requirement_id.upper())

    return sorted(requirement_ids)


if __name__ == "__main__":

    file_path = "requirements.docx"

    requirement_ids = extract_requirement_ids(file_path)

    print("\n" + "=" * 50)
    print(f"Found {len(requirement_ids)} unique requirement IDs")
    print("=" * 50)

    for requirement_id in requirement_ids:
        print(requirement_id)
