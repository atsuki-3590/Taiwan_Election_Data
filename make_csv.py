import pandas as pd


file_path = "第11屆平地原住民立法委員選舉候選人在臺北市各投開票所得票數一覽表.xlsx"
excel_file = pd.ExcelFile(file_path)

excel_file.sheet_names