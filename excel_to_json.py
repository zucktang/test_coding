import openpyxl
import json
from datetime import datetime
from openpyxl.styles import PatternFill

class State:
    HEADER, CAMPAIGN, INTEREST = range(3)

def is_yellow(cell):
    return cell.fill.start_color.index == 'FFFFFF00'  # Check for yellow background color

def process_sheet(sheet):
    state = State.HEADER
    campaign_data = {"data": [], "interest": []}

    for row in sheet.iter_rows(values_only=True):
        if state == State.HEADER and any(row):
            state = State.CAMPAIGN
        elif state == State.CAMPAIGN and "interest rate" in row:
            state = State.INTEREST
            continue
        elif state == State.INTEREST and not any(row):
            break  # Stop reading when encountering an empty row

        if state == State.INTEREST:
            campaign_data["interest"].append(list(row))
        elif state == State.CAMPAIGN and is_yellow(sheet.cell(row=sheet.max_row, column=1)):  # Check yellow background for the last row
            campaign_data["data"].append({
                "effdate": str(row[0]),
                "filename": str(row[1]),
                "remark": str(row[2])
            })

    return campaign_data

def read_excel(file_path):
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    result_dict = {}

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        campaign_data = process_sheet(sheet)
        result_dict[sheet_name] = {
            "data": campaign_data["data"],
            "interest": campaign_data["interest"]
        }

    return result_dict

def main():
    file_path = 'hw4.xlsx'
    result_dict = read_excel(file_path)
    json_string = json.dumps(result_dict, indent=2, ensure_ascii=False)
    
    with open('output.json', 'w', encoding='utf-8') as json_file:
        json_file.write(json_string)

if __name__ == "__main__":
    main()
