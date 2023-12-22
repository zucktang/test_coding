import openpyxl
from datetime import datetime

def get_next_coordinate(cell_coordinate, i=1, is_row=True):
    col_letter, col_number = cell_coordinate[0], int(cell_coordinate[1:])
    if is_row:
        col_number += i
    else:
        col_letter = chr(ord(col_letter) + i)
    return f"{col_letter}{col_number}"

def extract_data_from_sheet(sheet):
    yellow_color = 'FFFFFF00'
    red_color = 'FFFF0000'
    
    campaign_cols, interest_rate_cols = {}, {}
    campaign_name, filename_coordinate, effdate_coordinate, remark_coordinate = '', '', '', ''

    for col in sheet.iter_cols():
        for cell in col:
            if cell.fill.fgColor.rgb == yellow_color:
                if 'campaign' in str(cell.value).lower():
                    campaign_name = cell.value
                if cell.value == 'effdate':
                    effdate_coordinate = cell.coordinate
                if cell.value == 'filename':
                    filename_coordinate = cell.coordinate
                if cell.value == 'remark':
                    remark_coordinate = cell.coordinate
                campaign_cols[cell.coordinate] = cell.value

            if cell.font.color.rgb == red_color:
                interest_rate_cols[cell.coordinate] = cell.value

    return campaign_name, filename_coordinate, effdate_coordinate, remark_coordinate, campaign_cols, interest_rate_cols

def process_interest_rate_cols(interest_rate_cols):
    interest_list = []

    for index, (key, value) in enumerate(interest_rate_cols.items()):
        data_in_rows = []
        for i in range(3):
            next_col = get_next_coordinate(key, i, False)
            if next_col in interest_rate_cols:
                data_in_rows.append(interest_rate_cols[next_col])
        interest_list.append(data_in_rows)

        if index == 2:
            break

    return interest_list

def process_campaign_cols(effdate_coordinate, filename_coordinate, remark_coordinate, campaign_cols):
    data_list = []

    loops = int(len(campaign_cols) / 6 - 1)
    i = 1
    while i <= loops:
        effdate = get_next_coordinate(effdate_coordinate, i)
        filename = get_next_coordinate(filename_coordinate, i)
        remark = get_next_coordinate(remark_coordinate, i)
        data_dict = {
            "effdate": campaign_cols[effdate].strftime("%Y/%m/%d"),
            "filename": campaign_cols[filename],
            "remark": campaign_cols[remark]
        }
        data_list.append(data_dict)
        i += 1

    return data_list

def read_excel(file_name):
    wb = openpyxl.load_workbook(file_name, data_only=True)
    result = {}
    for sheet in wb:
        (
            campaign_name,
            filename_coordinate,
            effdate_coordinate,
            remark_coordinate,
            campaign_cols,
            interest_rate_cols
        ) = extract_data_from_sheet(sheet)

        data_list = process_campaign_cols(effdate_coordinate, filename_coordinate, remark_coordinate, campaign_cols)
        interest_list = process_interest_rate_cols(interest_rate_cols)
        result[campaign_name] = {"data": data_list, "interest": interest_list}

    return result

def main():
    file_name = 'hw4.xlsx'
    result = read_excel(file_name)
    print(result)

if __name__ == "__main__":
    main()
