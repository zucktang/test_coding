import json
import datetime


filename='hw4.xlsx'
with open(filename, "rb") as f:
    data = f.read()
    data_lines = data.split(b"\n")
    data_split = [line.split(b",") for line in data_lines]

print(data_split)

#   # สร้าง state machine
#   class StateMachine:
#     def __init__(self):
#       self.state = "campaign"
#       self.campaign = None
#       self.data = []
#       self.interest = []

#     def process_row(self, row):
#       if self.state == "campaign":
#         if "campaign" in row[0]:
#           self.campaign = row[0]
#         elif len(row) == 0:
#           self.state = "interest"
#         else:
#           self.data.append({
#             "effdate": datetime.datetime.strptime(row[0], "%Y/%m/%d"),
#             "filename": row[1],
#             "remark": row[2],
#           })
#       elif self.state == "interest":
#         self.interest.append(row)

#   # เริ่มต้น state machine
#   sm = StateMachine()

#   # วนลูปอ่านข้อมูลทีละ row
#   for row in data:
#     sm.process_row(row)

#   # สร้าง JSON ของผลลัพธ์
#   result = {
#     "data": sm.data,
#     "interest": sm.interest,
#   }

#   # คืนค่า JSON
#   return json.dumps(result)


# if __name__ == "__main__":
#   filename = "hw4.xlsx"
#   result = read_excel_to_json(filename)
#   print(result)