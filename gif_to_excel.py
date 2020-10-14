import typer
from PIL import Image
from PIL import GifImagePlugin
import numpy as np
from pathlib import Path
import xlsxwriter
from tqdm.auto import trange
from matplotlib.colors import rgb2hex


def main(
    input_file: Path,
    output_file: Path,
    column_width: float = 0.1,
    row_height: float = 1.0,
):
    """
    Entry point and main logic for script to convert image to Excel file
    :param input_file: path of the image file to read
    :param output_file: path of the xlsx file to write
    :param column_width: width of the columns
    :param row_height: height of the rows
    """
    img = Image.open(input_file)
    workbook = xlsxwriter.Workbook(output_file)

    for frame in trange(0, img.n_frames, desc="frame"):
        img.seek(frame)
        img_array = np.array(img.convert(mode="RGB"))
        worksheet = workbook.add_worksheet(name=f"Frame {frame}")
        worksheet.set_default_row(row_height)
        worksheet.set_column(0, img_array.shape[1], column_width)
        for row in range(img_array.shape[0]):
            for col in range(img_array.shape[1]):
                color_string = rgb2hex(img_array[row, col, :3] / 255)
                cell_format = workbook.add_format()
                cell_format.set_bg_color(color_string)
                worksheet.write(row, col, None, cell_format)

    workbook.add_vba_project("./vbaProject.bin")
    workbook.close()


if __name__ == "__main__":
    typer.run(main)
