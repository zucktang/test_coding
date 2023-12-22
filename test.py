import openpyxl
from openpyxl.styles import PatternFill

def get_cell_background_color(sheet, cell):
    fill = sheet[cell].fill
    return fill.fgColor.rgb
   

def main():
    # Load the Excel workbook
    workbook = openpyxl.load_workbook('hw4.xlsx')

    # Select the desired sheet
    sheet = workbook['Sheet1']  # Replace 'Sheet1' with your sheet name

    # Specify the cell you want to check
    target_cell = 'A3'  # Replace with your cell reference

    # Get the background color of the cell
    background_color = get_cell_background_color(sheet, target_cell)

    if background_color is not None:
        print(f"The background color of cell {target_cell} is: {background_color}")
    else:
        print(f"Cell {target_cell} does not have a background color.")

if __name__ == "__main__":
    main()