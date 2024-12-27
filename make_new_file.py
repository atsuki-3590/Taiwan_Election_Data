import os
import glob
import warnings
import pandas as pd
import numpy as np
import xlsxwriter

# FutureWarning を抑制 (必要があれば)
warnings.simplefilter(action='ignore', category=FutureWarning)

input_folder = "第11届区域立法委員選挙"
output_folder = "new_第11届区域立法委員選挙"

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

for file_path in glob.glob(os.path.join(input_folder, "*.xlsx")):
    file_name = os.path.basename(file_path)
    new_file_name = file_name.replace(".xlsx", "_new.xlsx")
    output_file_path = os.path.join(output_folder, new_file_name)

    excel_file = pd.ExcelFile(file_path)
    
    with pd.ExcelWriter(output_file_path, engine="xlsxwriter") as writer:
        for sheet_name in excel_file.sheet_names:
            df = excel_file.parse(sheet_name)
            
            # ---- 元の処理ここから ----
            df.iloc[1] = df.iloc[1].fillna(df.iloc[0])
            df_cleaned = df.dropna(how="all")
            df_cleaned = df_cleaned.iloc[1:]  # タイトル行を削除
            df_cleaned.columns = df_cleaned.iloc[0]  # 最初の行をカラム名に
            df_cleaned = df_cleaned.iloc[1:]  # カラム名として使用した行を削除

            df_cleaned['鄉(鎮、市、區)別'] = df_cleaned['鄉(鎮、市、區)別'].ffill()
            df_cleaned = df_cleaned.dropna(subset=['村里別', '投開票所別'], how='all')

            # ★ ここを追加：数値として扱いたい列を明示的に to_numeric する
            #    変換できないセルは NaN → fillna(0) で最終的に 0 扱い
            df_cleaned['投票數C\nC=A+B'] = pd.to_numeric(df_cleaned['投票數C\nC=A+B'], errors='coerce').fillna(0)
            df_cleaned['選舉人數G\nG=E+F'] = pd.to_numeric(df_cleaned['選舉人數G\nG=E+F'], errors='coerce').fillna(0)

            # original_index の追加・グループ化
            df_cleaned['original_index'] = df_cleaned.index
            df_grouped = df_cleaned.groupby(['鄉(鎮、市、區)別', '村里別']).sum().reset_index()

            # 分母が0なら投票率=0 のロジック
            df_grouped['投票率H\nH=C÷G'] = np.where(
                df_grouped['選舉人數G\nG=E+F'] == 0,
                0,
                (df_grouped['投票數C\nC=A+B'] / df_grouped['選舉人數G\nG=E+F']) * 100
            )

            df_grouped = df_grouped.sort_values(by='original_index').drop(columns=['original_index'])
            # ---- 元の処理ここまで ----

            df_grouped.to_excel(writer, sheet_name=sheet_name, index=False)

            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            normal_format = workbook.add_format({'bold': False})
            for col_num, value in enumerate(df_grouped.columns):
                worksheet.write(0, col_num, value, normal_format)

    print(f"処理完了: {file_path} -> {output_file_path}")
