import openpyxl
import json
from datetime import datetime

class State:
    HEADER, CAMPAIGN, INTEREST = range(3)


def process_sheet(sheet):
    state = State.HEADER
    campaign_data = {}

    for row in sheet.iter_rows(values_only=True):

    return campaign_data

def read_excel(file_path):
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    result_dict = {}

    for sheet in workbook:
        campaign_data = process_sheet(sheet)
        result_dict.update(campaign_data)

    return result_dict

def main():
    file_path = 'hw4.xlsx'
    result_dict = read_excel(file_path)
    json_string = json.dumps(result_dict, indent=2, ensure_ascii=False)
    
    with open('output.json', 'w', encoding='utf-8') as json_file:
        json_file.write(json_string)

if __name__ == "__main__":
    main()
