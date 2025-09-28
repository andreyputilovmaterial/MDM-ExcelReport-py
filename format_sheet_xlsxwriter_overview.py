

def format_sheet(workbook, worksheet, nrows, ncols,config):
    # --- formats ---
    font_white = workbook.add_format({"font_color": "#FFFFFF"})
    font_header_bold_gray = workbook.add_format({
        "font_color": "#666666",
        "bold": True,
        "font_size": 24
    })
    align_col1 = workbook.add_format({
        "align": "right",
        "valign": "top",
        "indent": 2
    })
    align_col2 = workbook.add_format({
        "align": "left",
        "valign": "top",
        "indent": 2
    })
    fill_white = workbook.add_format({
        "bg_color": "#FFFFFF"
    })
    # Note: borders are none by default, so no need to set them

    # --- column widths ---
    worksheet.set_column(0, 0, 35, align_col1)   # Column A
    worksheet.set_column(1, 1, 125, align_col2) # Column B

    # --- row heights ---
    worksheet.set_row(0, 65, font_white)            # Header row 1
    worksheet.set_row(1, 45, font_header_bold_gray) # Row 2
    for r in range(2, nrows):
        worksheet.set_row(r, 25)  # remaining rows, default alignment

    # --- header fill / font ---
    # In xlsxwriter, fills + font applied via formats at write-time
    # If you already wrote values, you can reapply format via conditional_format on all headers
    worksheet.conditional_format(0, 0, 0, ncols-1, {"type": "no_errors", "format": font_white})  # row 1
    worksheet.conditional_format(1, 1, 1, 1, {"type": "no_errors", "format": font_header_bold_gray}) # B2
