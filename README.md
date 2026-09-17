
# Docker-Files
This Repo contains Docker files
import re
from docx import Document


REQUIREMENT_PATTERN = re.compile(r"\bACDE-TOR-\d+\b", re.IGNORECASE)


def extract_requirement_ids(file_path):
    document = Document(file_path)

    requirement_ids = set()

    for table in document.tables:

        # Find the "ID / TYPE" column
        header_row = table.rows[0]

        id_column_index = None

        for index, cell in enumerate(header_row.cells):
            header = cell.text.strip().upper()

            if header == "ID / TYPE":
                id_column_index = index
                break

        # Skip this table if it doesn't have an ID / TYPE column
        if id_column_index is None:
            continue

        # Read values from the ID / TYPE column
        for row in table.rows[1:]:
            cell_text = row.cells[id_column_index].text.strip()

            matches = REQUIREMENT_PATTERN.findall(cell_text)

            for requirement_id in matches:
                requirement_ids.add(requirement_id.upper())

    return sorted(requirement_ids)


if __name__ == "__main__":
    file_path = "requirements.docx"

    requirement_ids = extract_requirement_ids(file_path)

    print(f"Found {len(requirement_ids)} requirement IDs:")

    for requirement_id in requirement_ids:
        print(requirement_id)

