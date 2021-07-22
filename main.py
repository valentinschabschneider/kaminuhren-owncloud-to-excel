from io import BytesIO
import logging
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

from openpyxl.drawing.image import Image as XlImage
from openpyxl import Workbook
from openpyxl.styles import Alignment

from PIL import Image

root = tk.Tk()
root.withdraw()

root_directory = Path(filedialog.askdirectory())

# only subfolders with naming scheme
clock_directories = [
    path for path in root_directory.glob("Kaminuhr-*") if path.is_dir()
]

# create Workbook
workbook = Workbook()
worksheet = workbook.worksheets[0]
worksheet.title = "Tabelle1"

for row_index, c_dir in enumerate(clock_directories, 1):

    name_cell = worksheet.cell(row_index, 1)
    name_cell.value = c_dir.name

    description_file_path = c_dir.joinpath("Beschreibung.txt")
    if not description_file_path.exists():
        logging.warn('"%s" does not exist!', description_file_path)
        description = ""
    else:
        description = open(description_file_path, encoding="utf-8").read()

    description_cell = worksheet.cell(row_index, 2)
    description_cell.alignment = Alignment(wrapText=True)
    description_cell.value = description

    qr_code_file_path = c_dir.joinpath(c_dir.name).with_suffix(".png")
    if not qr_code_file_path.exists():
        logging.warn('"%s" does not exist!', qr_code_file_path)
    else:
        qr_code = Image.open(qr_code_file_path)
        qr_code.thumbnail((256, 256), Image.ANTIALIAS)

        qr_code_cell = worksheet.cell(row_index, 3)

        temp_qr_code = BytesIO()
        qr_code.save(temp_qr_code, format="png")

        img = XlImage(temp_qr_code)
        worksheet.add_image(img, qr_code_cell.coordinate)

    try:
        clock_image_file_path = next(c_dir.glob("*.jpg"))
    except StopIteration:
        logging.warn("no clock image found in %s", c_dir.name)
    else:
        clock_image = Image.open(clock_image_file_path)
        clock_image.thumbnail((256, 256), Image.ANTIALIAS)

        clock_image_cell = worksheet.cell(row_index, 4)

        temp_clock_image = BytesIO()
        clock_image.save(temp_clock_image, format="jpeg")

        img = XlImage(temp_clock_image)
        worksheet.add_image(img, clock_image_cell.coordinate)

    worksheet.row_dimensions[row_index].height = 192

worksheet.column_dimensions["A"].width = 14
worksheet.column_dimensions["B"].width = 70
worksheet.column_dimensions["C"].width = 36
worksheet.column_dimensions["D"].width = 36


file_path = Path(
    filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")),
    )
)

workbook.save(file_path)
