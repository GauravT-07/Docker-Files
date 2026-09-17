from openpyxl import load_workbook


EXISTING_EXCEL = "existing_requirements.xlsx"
EXTRACTED_EXCEL = "extracted_tor_ids.xlsx"


def inspect_excel(file_path):

    print("\n" + "=" * 70)
    print(f"FILE: {file_path}")
    print("=" * 70)

    workbook = load_workbook(
        file_path,
        read_only=True,
        data_only=True
    )

    print("Sheets:")

    for sheet_name in workbook.sheetnames:
        print(f"  - {sheet_name}")

    print("\nColumn A values:\n")

    for worksheet in workbook.worksheets:

        print(f"--- SHEET: {worksheet.title} ---")

        count = 0

        for row_number, row in enumerate(
            worksheet.iter_rows(
                min_col=1,
                max_col=1
            ),
            start=1
        ):

            value = row[0].value

            if value is not None:

                print(
                    f"Row {row_number}: "
                    f"value={repr(value)} "
                    f"type={type(value).__name__}"
                )

                count += 1

                # Only print first 20
                if count >= 20:
                    break

        print()


inspect_excel(EXISTING_EXCEL)

inspect_excel(EXTRACTED_EXCEL)
