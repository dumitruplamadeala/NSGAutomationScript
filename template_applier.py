import argparse
import shutil
import time
from pathlib import Path

import pythoncom
import xlwings as xw
from win32com.client import Dispatch


def get_used_size(sheet):
    used = sheet.api.UsedRange

    last_row = used.Row + used.Rows.Count - 1
    last_col = used.Column + used.Columns.Count - 1

    return last_row, last_col


def clear_template_data(sheet, last_col):
    """
    Clear values/formulas only.
    Template formatting and conditional formatting remain untouched.
    """
    template_last_row, template_last_col = get_used_size(sheet)

    if template_last_row < 2:
        return

    last_col = max(last_col, template_last_col)

    sheet.api.Range(
        sheet.api.Cells(2, 1),
        sheet.api.Cells(template_last_row, last_col),
    ).ClearContents()


def copy_values_and_formulas(
    source_sheet,
    target_sheet,
    last_row,
    last_col,
):
    """
    Copy contents only.

    This does NOT copy:
      - conditional formatting
      - borders
      - fills
      - cell styles
    """
    if last_row < 2:
        return

    source_range = source_sheet.api.Range(
        source_sheet.api.Cells(2, 1),
        source_sheet.api.Cells(last_row, last_col),
    )

    target_range = target_sheet.api.Range(
        target_sheet.api.Cells(2, 1),
        target_sheet.api.Cells(last_row, last_col),
    )

    target_range.Formula = source_range.Formula


def get_characters(cell, start, length):
    range_api = cell.api
    dispid = range_api._oleobj_.GetIDsOfNames("Characters")
    characters = range_api._oleobj_.Invoke(
        dispid,
        0,
        pythoncom.DISPATCH_PROPERTYGET,
        1,
        start,
        length,
    )
    return Dispatch(characters)


def get_char_style(cell, position):
    font = get_characters(cell, position, 1).Font

    return (
        font.Bold,
        font.Italic,
        font.Underline,
        font.Strikethrough,
        font.Color,
        font.Subscript,
        font.Superscript,
    )


def apply_char_style(cell, start, length, style):
    font = get_characters(cell, start, length).Font

    (
        font.Bold,
        font.Italic,
        font.Underline,
        font.Strikethrough,
        font.Color,
        font.Subscript,
        font.Superscript,
    ) = style


def copy_rich_text_cell(source_cell, target_cell):
    """
    Copy only character-level formatting for cells containing
    mixed rich text.

    Example:
        normal text + partially bold text
    """
    value = source_cell.value

    if not isinstance(value, str) or len(value) < 2:
        return

    try:
        first_style = get_char_style(source_cell, 1)
    except Exception:
        return

    styles = [first_style]
    mixed = False

    for position in range(2, len(value) + 1):
        style = get_char_style(source_cell, position)
        styles.append(style)

        if style != first_style:
            mixed = True

    # Entire cell has one style, so leave template formatting alone.
    if not mixed:
        return

    # Apply consecutive runs instead of formatting one character at a time.
    run_start = 1
    run_style = styles[0]

    for index in range(1, len(styles)):
        if styles[index] != run_style:
            apply_char_style(
                target_cell,
                run_start,
                index - run_start + 1,
                run_style,
            )

            run_start = index + 1
            run_style = styles[index]

    apply_char_style(
        target_cell,
        run_start,
        len(styles) - run_start + 1,
        run_style,
    )


def copy_rich_text(
    source_sheet,
    target_sheet,
    last_row,
    last_col,
):
    """
    Scan text cells and preserve only mixed character formatting.
    """
    for row in range(2, last_row + 1):
        for col in range(1, last_col + 1):

            source_cell = source_sheet.cells(row, col)

            if not isinstance(source_cell.value, str):
                continue

            target_cell = target_sheet.cells(row, col)

            try:
                copy_rich_text_cell(
                    source_cell,
                    target_cell,
                )
            except Exception as error:
                print(
                    f"Warning: rich text failed at "
                    f"{source_sheet.name}!R{row}C{col}: {error}"
                )


def set_action_wrap_text(sheet, last_row, last_col):
    if last_row < 2:
        return

    headers = [sheet.cells(1, col).value for col in range(1, last_col + 1)]

    try:
        action_col = next(
            col
            for col, header in enumerate(headers, start=1)
            if str(header).strip().casefold() == "action"
        )
    except StopIteration:
        print(f"Warning: Action column not found on '{sheet.name}'.")
        return

    sheet.api.Range(f"2:{last_row}").WrapText = False

    for row in range(2, last_row + 1):
        action = sheet.cells(row, action_col).value
        should_wrap = str(action).strip().casefold() in {
            "create",
            "remove",
            "update",
        }

        if should_wrap:
            sheet.api.Rows(row).WrapText = True


def merge_workbooks(source_file, template_file, output_file):
    source_file = Path(source_file).resolve()
    template_file = Path(template_file).resolve()
    output_file = Path(output_file).resolve()

    if output_file.exists():
        output_file.unlink()

    # Output starts as an exact copy of template.xlsx
    shutil.copy2(template_file, output_file)

    app = None
    wb_source = None
    wb_output = None

    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False

        wb_source = app.books.open(
            str(source_file),
            read_only=True,
            update_links=False,
        )

        wb_output = app.books.open(
            str(output_file),
            update_links=False,
        )

        template_sheets = {
            sheet.name
            for sheet in wb_output.sheets
        }

        for source_sheet in wb_source.sheets:

            sheet_name = source_sheet.name

            if sheet_name not in template_sheets:
                print(
                    f"Skipping '{sheet_name}': "
                    f"not present in template."
                )
                continue

            print(f"Processing: {sheet_name}")

            target_sheet = wb_output.sheets[sheet_name]

            last_row, last_col = get_used_size(source_sheet)

            print(
                f"  Source range: "
                f"A1 to R{last_row}C{last_col}"
            )

            # 1. Keep template structure/formatting.
            clear_template_data(
                target_sheet,
                last_col,
            )

            # 2. Copy values and formulas only.
            copy_values_and_formulas(
                source_sheet,
                target_sheet,
                last_row,
                last_col,
            )

            # 3. Restore partial/mixed rich text.
            copy_rich_text(
                source_sheet,
                target_sheet,
                last_row,
                last_col,
            )

            # 4. Wrap rows only for Create and Remove actions.
            set_action_wrap_text(
                target_sheet,
                last_row,
                last_col,
            )

        wb_output.save()

    finally:
        if wb_source:
            try:
                wb_source.close()
            except Exception:
                pass

        if wb_output:
            try:
                wb_output.close()
            except Exception:
                pass

        if app:
            try:
                app.quit()
            except Exception:
                pass


def main():
    start_time = time.time()

    parser = argparse.ArgumentParser(
        description="Apply Excel template while preserving rich text."
    )

    parser.add_argument(
        "input_file",
        help="Source Excel file",
    )

    parser.add_argument(
        "--template",
        default="assets/template.xlsx",
        help="Template file",
    )

    args = parser.parse_args()

    input_path = Path(args.input_file).resolve()
    template_path = Path(args.template).resolve()

    if not input_path.is_file():
        parser.error(f"Input file not found: {input_path}")

    if not template_path.is_file():
        parser.error(f"Template file not found: {template_path}")

    output_path = input_path.with_name(
        f"{input_path.stem}_templated{template_path.suffix}"
    )

    merge_workbooks(
        input_path,
        template_path,
        output_path,
    )

    print(f"Created: {output_path.name}")
    print(f"Finished in {time.time() - start_time:.2f}s")


if __name__ == "__main__":
    main()