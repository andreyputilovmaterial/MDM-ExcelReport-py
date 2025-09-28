# how to use:
# format_sheet(workbook, worksheet, nrows=len(df.index)+1, ncols=len(df.columns)+1)


def format_sheet(workbook, worksheet, nrows, ncols):
    # --- define formats ---
    align_left_top = workbook.add_format({
        "align": "left",
        "valign": "top"
    })

    align_left_bottom_bold = workbook.add_format({
        "align": "left",
        "valign": "bottom",
        "bold": True
    })

    # --- apply formats to first column (A) ---
    worksheet.set_column(0, 0, 35, align_left_top)  # col 0 = A

    # --- apply formats to first row (row 1) ---
    worksheet.set_row(0, 21, align_left_bottom_bold)  # row 0 = first row

    # --- set default width for all other columns ---
    if ncols > 1:
        worksheet.set_column(1, ncols - 1, 80)  # from col 1 (B) onwards