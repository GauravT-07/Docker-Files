from openpyxl import load_workbook, Workbook


EXISTING_EXCEL = "existing_requirements.xlsx"
EXTRACTED_EXCEL = "extracted_tor_ids.xlsx"

COMPARISON_EXCEL = "tor_comparison.xlsx"
TEXT_REPORT = "tor_comparison.txt"


def read_column_a(file_path):
    """
    Read TOR IDs directly from Column A.
    Duplicate values are automatically removed.
    """

    workbook = load_workbook(
        file_path,
        read_only=True,
        data_only=True
    )

    tor_ids = set()

    for worksheet in workbook.worksheets:

        for row in worksheet.iter_rows(
            min_col=1,
            max_col=1
        ):

            value = row[0].value

            if value is None:
                continue

            value = str(value).strip()

            # Skip header
            if value.upper() in [
                "TOR ID",
                "TOR_ID",
                "TORID",
                "ID",
                "ID / TYPE"
            ]:
                continue

            tor_ids.add(value)

    workbook.close()

    return tor_ids


def compare_tor_ids(existing_ids, extracted_ids):

    # Existing but NOT extracted
    delete_ids = existing_ids - extracted_ids

    # Extracted but NOT existing
    missing_ids = extracted_ids - existing_ids

    # Present in both
    common_ids = existing_ids & extracted_ids

    return (
        common_ids,
        delete_ids,
        missing_ids
    )


def create_comparison_excel(
    existing_ids,
    extracted_ids
):

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "TOR Comparison"

    worksheet.append([
        "TOR ID",
        "Status"
    ])

    all_ids = sorted(
        existing_ids | extracted_ids
    )

    for tor_id in all_ids:

        if (
            tor_id in existing_ids
            and tor_id in extracted_ids
        ):

            status = "Present in both"

        elif tor_id in existing_ids:

            status = (
                "Requirement and test case "
                "needs to be deleted"
            )

        else:

            status = "Requirement is missing"

        worksheet.append([
            tor_id,
            status
        ])

    workbook.save(
        COMPARISON_EXCEL
    )


def create_text_report(
    delete_ids,
    missing_ids
):

    with open(
        TEXT_REPORT,
        "w",
        encoding="utf-8"
    ) as file:

        file.write(
            "TOR COMPARISON REPORT\n"
        )

        file.write(
            "=" * 60 + "\n\n"
        )

        file.write(
            "REQUIREMENT AND TEST CASE NEEDS TO BE DELETED\n"
        )

        file.write(
            "-" * 60 + "\n"
        )

        if delete_ids:

            for tor_id in sorted(delete_ids):

                file.write(
                    f"{tor_id}\n"
                )

        else:

            file.write("None\n")

        file.write("\n\n")

        file.write(
            "REQUIREMENT IS MISSING\n"
        )

        file.write(
            "-" * 60 + "\n"
        )

        if missing_ids:

            for tor_id in sorted(missing_ids):

                file.write(
                    f"{tor_id}\n"
                )

        else:

            file.write("None\n")


if __name__ == "__main__":

    print("Reading existing Excel...")

    existing_ids = read_column_a(
        EXISTING_EXCEL
    )

    print(
        f"Existing Excel TOR IDs: "
        f"{len(existing_ids)}"
    )

    print("Reading extracted Excel...")

    extracted_ids = read_column_a(
        EXTRACTED_EXCEL
    )

    print(
        f"Extracted Excel TOR IDs: "
        f"{len(extracted_ids)}"
    )

    # --------------------------------------------------------
    # Compare
    # --------------------------------------------------------

    common_ids, delete_ids, missing_ids = (
        compare_tor_ids(
            existing_ids,
            extracted_ids
        )
    )

    # --------------------------------------------------------
    # Create Excel
    # --------------------------------------------------------

    create_comparison_excel(
        existing_ids,
        extracted_ids
    )

    # --------------------------------------------------------
    # Create TXT
    # --------------------------------------------------------

    create_text_report(
        delete_ids,
        missing_ids
    )

    # --------------------------------------------------------
    # Print results
    # --------------------------------------------------------

    print("\n" + "=" * 60)
    print("RESULT")
    print("=" * 60)

    print(
        f"Present in both       : {len(common_ids)}"
    )

    print(
        f"Need to delete        : {len(delete_ids)}"
    )

    print(
        f"Requirement missing   : {len(missing_ids)}"
    )

    print("=" * 60)

    print(
        f"\nCreated: {COMPARISON_EXCEL}"
    )

    print(
        f"Created: {TEXT_REPORT}"
    )
