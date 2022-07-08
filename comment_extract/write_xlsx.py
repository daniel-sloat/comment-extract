import xlsxwriter


def create_formats(
    workbook: xlsxwriter.Workbook, run: list[str, str]
) -> list[xlsxwriter.workbook.Format, str]:
    """Encodes format-string to xlsxwriter workbook Format."""
    text, props = run
    run_format = {}
    for p in props:
        match p:
            case "b":
                run_format["bold"] = True
            case "i":
                run_format["italic"] = True
            case "u":
                run_format["underline"] = 1
            case "w":
                run_format["underline"] = 2
            case "s":
                run_format["font_strikeout"] = True
            case "z":  # No double-strikeout in Excel so strikeout and red text is used.
                run_format["font_strikeout"] = True
                run_format["font_color"] = "#FF0000"  # #FF0000 = red
            case "x":
                run_format["font_script"] = 1
            case "v":
                run_format["font_script"] = 2
    format = workbook.add_format(run_format)
    run = [format, text]
    return run


def prepare_comments_for_excel(
    workbook: xlsxwriter.Workbook,
    file_data: list[list[list[str, str]]],
) -> list[list[xlsxwriter.workbook.Format, str]]:
    """Flattens the data within "Comment Data" from Comment[Paragraph[Run]]] to Comment[Run]]
    by appending a newline break character at the end of each paragraph (except the last
    paragraph). This is because Excel has no concept of paragraphs like Word does."""
    fresh_comments = []
    for data in file_data:
        fresh_comment_data = {}
        for name, value in data.items():
            if name == "Comment Data":
                fresh_runs = []
                for paragraph in value:
                    for run in paragraph:
                        # Replaces formats in string form to a Format object for use by xlsxwriter.
                        fresh_runs.extend(create_formats(workbook, run))
                    if paragraph != data[name][-1]:
                        fresh_runs.extend(create_formats(workbook, ["\n", ""]))
                fresh_comment_data[name] = fresh_runs
            else:
                fresh_comment_data[name] = value
        fresh_comments.append(fresh_comment_data)
    return fresh_comments


def set_column_formats(
    workbook: xlsxwriter.Workbook,
    worksheet: xlsxwriter.workbook.Worksheet,
    add_columns: bool=True,
) -> None:
    """Set column width and cell formats"""
    text_wrap = workbook.add_format({"text_wrap": 1, "valign": "top", "align": "left"})
    align = workbook.add_format({"valign": "top", "align": "left"})
    date_format = workbook.add_format(
        {"num_format": "[$-en-US]m/d/yy h:mm AM/PM;@", "valign": "top", "align": "left"}
    )
    hidden = {"hidden": True}
    
    set_col_formats = [
            (5, align),
            (15, align),
            (5, align),
            (12, align),
            (5, align),
            (18, text_wrap, hidden),
            (5, align, hidden),
            (16, date_format, hidden),
            (18, text_wrap, hidden),
            (18, text_wrap),
            (80, text_wrap),
        ]
    
    if add_columns:
        set_col_formats.insert(10, (18, text_wrap))
        set_col_formats.insert(11, (18, text_wrap, hidden))
        set_col_formats.insert(13, (80, text_wrap))
        
    for no, col in enumerate(set_col_formats, 0):
        worksheet.set_column(no, no, *col)
    return None


def write_header(
    workbook: xlsxwriter.Workbook,
    worksheet: xlsxwriter.workbook.Worksheet,
    column_names: list[str],
) -> None:
    """Writes header cells with format."""
    header_format = workbook.add_format(
        {
            "bold": True,
            "text_wrap": True,
            "valign": "bottom",
            "border": 1,
        }
    )
    for col_num, value in enumerate(column_names):
        worksheet.write(0, col_num, value, header_format)
    return None


def write_rich_list(
    worksheet: xlsxwriter.workbook.Worksheet,
    row: int,
    col: int,
    rich_list: list[xlsxwriter.workbook.Format, str],
) -> xlsxwriter.workbook.Worksheet:
    """A write handler rich formatted lists for xlsxwriter."""
    if len(rich_list) == 2:
        return worksheet.write_string(row, col, rich_list[1])
    else:
        return worksheet.write_rich_string(row, col, *rich_list)


def write_data(
    worksheet: xlsxwriter.workbook.Worksheet,
    data: list[dict[str : [int | str | list[list[xlsxwriter.workbook.Format, str]]]]],
) -> None:
    """Writes data across columns and rows to worksheet."""
    for col_num, col in enumerate(data[0], 1):
        # Begin with row after header row, increase row
        for row in range(1, len(data) + 1):
            comment_no = row - 1
            worksheet.write(row, 0, row)
            worksheet.write(row, col_num, data[comment_no][col])
    return None


def create_workbook(
    file_data: list[dict[str : [int | str | list[list[list[str, str]]]]]],
    workbook_savepath: str = r"output/comments.xlsx",
    sheet_name: str = "Comments",
    add_columns: bool=False,
) -> str:
    """Create xlsx workbook from comment data, structured as described.

    Args:
        file_data (
            list[
                dict[
                    str : [int | str | list[list[list[str, str]]]]
                    ]
                ]
        ):
            Data must be structured as described.

        workbook_savepath (str, optional): Path and name of saved workbook.
            Defaults to r"output/workbook.xlsx".
        sheet_name (str, optional): Name of sheet. Defaults to "Comments".
    """
    workbook = xlsxwriter.Workbook(workbook_savepath)
    worksheet = workbook.add_worksheet(sheet_name)
    worksheet.add_write_handler(list, write_rich_list)

    set_column_formats(workbook, worksheet, add_columns=add_columns)
    xl_ready_file_data = prepare_comments_for_excel(workbook, file_data)
    column_names = ["Comment Number"] + list(xl_ready_file_data[0])

    write_header(workbook, worksheet, column_names)
    write_data(worksheet, xl_ready_file_data)

    # Autofilter and Freeze Panes
    max_row = len(xl_ready_file_data)
    max_col = len(column_names) - 1
    worksheet.autofilter(0, 0, max_row, max_col)
    worksheet.freeze_panes(1, 0)

    workbook.close()
    return workbook_savepath
