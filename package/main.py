import os
import pdfplumber
import pandas as pd
import re
import xlwt
import xlrd
from xlutils.copy import copy
import xlwings as xw
import xlsxwriter
import pkg_resources
from openpyxl import load_workbook
from xlwt import Formula
from openpyxl.drawing.image import Image
import argparse

phone_number_pattern = re.compile(
    r"(\+\d{1,3} \(\d{1,4}\) \d{1,4}-\d{1,4}-\d{1,4})|(\d{10})"
)


def clean_text(text):
    return text.replace("(cid:695)", "і")


# Get the path to the base.xls file in the package
BASE_FILE_PATH = pkg_resources.resource_filename(__name__, "base.xls")
CAR_IMAGE_FILE_PATH = pkg_resources.resource_filename(__name__, "car-image.png")


def extract_data_from_pdf(pdf_path, output_xls_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        tables = []
        for page in pdf.pages:
            text += page.extract_text()
            tables.append(page.extract_table())

    all_pdf_data = text.split("\n")
    pdf_data_without_dileverer = all_pdf_data[12:]
    dileverer_address = re.sub(r"Маршрут: [^\s]*", "", all_pdf_data[2]).strip()
    document_number_and_date = [
        line for line in all_pdf_data if line.startswith("Накладна")
    ][1]

    default_font_style = "font: name Calibri, height 220; "
    default_border_style = "border: top thin, bottom thin, left thin, right thin; "
    default_font_styles = xlwt.easyxf(default_font_style)

    rb = xlrd.open_workbook(BASE_FILE_PATH, formatting_info=True)

    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    base_table_start_end_rows = {"start": 12, "end": 26}

    max_lengths = [0] * len(tables[0][0])
    rows_content = []

    items_to_remove_from_headers_and_items = [3, 5, 6]

    for table in tables:
        if table and len(table):

            table_headers = [
                item
                for i, item in enumerate(table[0])
                if i not in items_to_remove_from_headers_and_items
            ]
            table_items = table[1:]

            table_headers[4] = "Ціна"
            table_headers[5] = "Сума"

            table_items_indexes_to_tranfrom_to_number = [3, 4]

            for table_header_index, table_header in enumerate(table_headers):
                w_sheet.write(
                    base_table_start_end_rows["start"],
                    table_header_index + 1,
                    table_header,
                    xlwt.easyxf(
                        f"font: name Calibri, height 220, bold on; {default_border_style}"
                    ),
                )
                max_lengths[table_header_index] = max(
                    max_lengths[table_header_index], len(str(table_header))
                )

            for table_item_index, table_item in enumerate(table_items):
                row_data = []
                for item_index, item_data in enumerate(
                    [
                        item
                        for i, item in enumerate(table_item)
                        if i not in items_to_remove_from_headers_and_items
                    ]
                ):
                    data_str = clean_text(item_data)

                    data_to_write = data_str
                    if (
                        data_str
                        and item_index in table_items_indexes_to_tranfrom_to_number
                    ):
                        data_to_write = float(
                            data_str.replace(" ", "").replace(",", ".")
                        )

                    if item_index == 5:
                        index = (
                            base_table_start_end_rows["start"] + table_item_index + 2
                        )
                        data_to_write = Formula(f"E{index}*F{index}")

                    w_sheet.write(
                        base_table_start_end_rows["start"] + 1 + table_item_index,
                        item_index + 1,
                        data_to_write,
                        xlwt.easyxf(f"{default_font_style} {default_border_style}"),
                    )
                    max_lengths[item_index] = max(
                        max_lengths[item_index], len(data_str)
                    )
                    row_data.append(data_str)
                rows_content.append(row_data)

            w_sheet.write(
                base_table_start_end_rows["start"] + len(table),
                4,
                "Загальна сума:",
                xlwt.easyxf(f"font: name Calibri, height 220, bold on;"),
            )

            base_items_table_start_index = base_table_start_end_rows["start"] + 2

            w_sheet.write(
                base_table_start_end_rows["start"] + len(table),
                6,
                Formula(
                    f"SUM(G{base_items_table_start_index}:G{base_items_table_start_index + len(table_items) - 1})"
                ),
                xlwt.easyxf(f"font: name Calibri, height 220, bold on;"),
            )

    for col_index, max_length in enumerate(max_lengths):
        w_sheet.col(col_index + 1).width = 256 * (max_length + 2)

    for row_index, row_data in enumerate(w_sheet.rows):
        height = 300
        if (
            row_index >= base_table_start_end_rows["start"]
            and row_index < base_table_start_end_rows["end"] - 1
        ):
            height = 600
        w_sheet.row(row_index).height = height

    wb.save(output_xls_path)

    wb = xw.Book(output_xls_path)
    sheet = wb.sheets["Sheet 1"]

    cell_location = "E3"

    sheet.pictures.add(
        CAR_IMAGE_FILE_PATH,
        name="Car-Image",
        update=True,
        left=sheet.range(cell_location).left,
        top=sheet.range(cell_location).top,
    )

    wb.save(output_xls_path)
    wb.close()


def main():
    parser = argparse.ArgumentParser(
        description="Extract data from PDF and save to Excel."
    )
    parser.add_argument("pdf_path", help="Path to the PDF file to extract data from.")
    parser.add_argument(
        "output_xls_path", help="Path for the output Excel file (.xls)."
    )

    args = parser.parse_args()

    extract_data_from_pdf(args.pdf_path, args.output_xls_path)


if __name__ == "__main__":
    main()
