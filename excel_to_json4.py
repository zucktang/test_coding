import openpyxl
import datetime

def convert_excel_to_json(file_path):
    # อ่านไฟล์ Excel
    workbook = openpyxl.load_workbook(file_path)

    # เริ่มต้นด้วย state เริ่มต้น
    state = "initial"

    # ประกาศ list ของตาราง campaign
    campaign_data = []
    interest_rate_data = []

    # ประกาศตัวแปรสำหรับเก็บคอลัมน์ที่สนใจ
    column = None

    # วนลูปอ่านข้อมูลในไฟล์
    for row in workbook.active.iter_rows():
        # แปลงข้อมูลใน row เป็น list
        row_data = [cell.value for cell in row]
        # ตรวจสอบ state ปัจจุบัน
        if state == "initial":
            # ถ้าพบคอลัมน์ effdate, filename, remark ครบทั้งสามคอลัมน์ ให้เปลี่ยน state เป็น state campaign
            if "effdate" in row_data and "filename" in row_data and "remark" in row_data:
                state = "campaign"
        elif state == "campaign":
            # ถ้าพบคอลัมน์ interest rate ให้เปลี่ยน state เป็น state interest rate
            if "interest rate" in row_data:
                state = "interest rate"
        elif state == "interest rate":
            # ถ้าพบคอลัมน์อื่น ให้ข้ามไป
            if "effdate" not in row_data and "filename" not in row_data:
                continue

        # ดำเนินการตาม state ปัจจุบัน
        if state == "campaign":
            # ตรวจสอบคอลัมน์แรก
            if column is None:
                column = row[0]

            # ถ้าคอลัมน์ปัจจุบันเป็นคอลัมน์ที่เราสนใจ
            if column == row[0]:
                # เพิ่มข้อมูลลงใน list ของตาราง campaign
                campaign_data.append(row_data)

                # เลื่อนคอลัมน์ไปทีละคอลัมน์
                column = row[1]

    # แปลงข้อมูลของตาราง campaign เป็น JSON
    campaign_data_json = [
        {"effdate": row[0],
         "filename": row[1],
         "remark": row[2]}
        for row in campaign_data
    ]

    # แปลงข้อมูลของตาราง interest rate เป็น JSON
    interest_rate_data_json = [
        [cell.value for cell in row] for row in interest_rate_data
    ]

    # คืนค่าข้อมูล JSON ของทั้งสองตาราง
    return {"campaign": campaign_data_json, "interest rate": interest_rate_data_json}


if __name__ == "__main__":
    # ไฟล์ Excel ที่จะแปลง
    file_path = "hw4.xlsx"

    # แปลงไฟล์ Excel เป็น JSON
    json_data = convert_excel_to_json(file_path)

    # พิมพ์ข้อมูล JSON
    print(json_data)