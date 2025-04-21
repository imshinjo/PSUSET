import glob
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

# 現在の月を取得して、前月の数字を計算
current_month = datetime.now().month
prev_month = current_month - 1 if current_month > 1 else 12 # 1月なら前月は12月
two_month_ago = prev_month - 1 if prev_month > 1 else 12

# 月と列の対応表を定義
month_column_4E = { # 各月の記入列の対応表
    4: "E", 5: "F", 6: "G", 7: "H", 8: "I", 9: "J",
    10: "K", 11: "L", 12: "M", 1: "N", 2: "O", 3: "P"
}
month_column_4D = { # 教員カラープリンタ印刷統計.xlsxの各月の記入列の対応表
    4: "D", 5: "E", 6: "F", 7: "G", 8: "H", 9: "I",
    10: "J", 11: "K", 12: "L", 1: "M", 2: "N", 3: "O"
}
prev_col_4E = month_column_4E[prev_month] # 先月（記入対象）列
prev_col_4D = month_column_4D[prev_month]
two_col_4E = month_column_4E[two_month_ago] # 先々月（比較対象）列
two_col_4D = month_column_4D[two_month_ago]



def file_handler(report_file): # 統計記入ファイルに対応した利用集計ファイルを返す関数
    
    if report_file != "Ricohスキャナ統計.xlsx":
        print("最新の機器カウンターレポートから値を取得します")

        file = glob.glob("./number_report/最新の機器カウンターレポート*") # 最新の CSV を取得（前方一致）
        csv_file = file[0] # CSVは1つだけの想定なので先頭のファイルを取得

        # === CSVを Shift-JIS → UTF-8 に変換 ===
        utf8_csv_path = csv_file.replace(".csv", "_utf8.csv")

        df = pd.read_csv(csv_file, encoding="shift_jis", comment="#", skip_blank_lines=True)
        df.to_csv(utf8_csv_path, encoding="utf-8", index=False)
        print(f"UTF-8 へ変換完了: {utf8_csv_path}")

        return utf8_csv_path
    
    else:
        print("機能×カラー別集計レポートから値を取得します")
        file = glob.glob("./number_report/機能×カラー別集計レポート*") # 最新の xls を取得（前方一致）
        excel_file = file[0]  # 先頭のファイルを取得

        return excel_file
    


def excel_hadler(report_file, sheet_name): # excelデータをインスタンス化したものと，値がある最後の行を返す
    workbook = load_workbook(report_file)
    sheet = workbook[sheet_name]

    # 統計記入ファイルのホスト情報列の最終行を取得
    last_row = sheet.max_row # シートの最終行を取得
    while sheet.cell(row=last_row, column=2).value is None and last_row > 1: # B列の,値がある最後の行を取得
        last_row -= 1  # 空欄なら1行ずつ上にさかのぼる

    return workbook, sheet, last_row



def fetch_data(report_file, sheet_name, prev_col, column_char, column_head):
    print(f"{sheet_name} シートに記入します")

    total_file = file_handler(report_file) # 利用集計CSVファイル

    stat_file = excel_hadler(f"./statistics_report/{report_file}", sheet_name)
    stat_workbook = stat_file[0] # 統計レポートのExcelファイル
    stat_sheet = stat_file[1] # 統計レポートExcelの記入対象シート    
    last_row = stat_file[2] # B列の,値がある最後の行
    

    if report_file != "Ricohスキャナ統計.xlsx":
        df_utf8 = pd.read_csv(total_file, encoding="utf-8")

        for row_idx in range(3, last_row + 1):  # 統計レポートの行を1つずつチェック
            host_name = stat_sheet[f"{column_char}{row_idx}"].value  # 統計レポートからホスト名取得

            if host_name:
                if host_name in df_utf8["Host Name"].values: # 集計レポートからホスト名を検索
                    value = df_utf8[df_utf8["Host Name"] == host_name][column_head].values[0] # host_nameと一致する行のcolumn_head列の値を取得
                    stat_sheet[f"{prev_col}{row_idx}"].value = value  # 前月の列に書き込み
                    print(f"{host_name} の {column_head} = {value} を {prev_col}{row_idx} に記入")
                
                else:
                    print(f"{host_name} が見つかりませんでした")


    else: # Ricohスキャナ統計.xlsxの場合
        # プリンタサーバから出力されるxlsファイルが破損しているため，手動で新規xlsxファイルにコピー＆ペーストする。
        # xlsファイルの機器別サマリーシートだけを，新規作成したxlsxファイルに貼り付けるのみで良い。

        total_workbook = load_workbook(total_file)  # `openpyxl` で `.xlsx` ファイルを開く
        total_sheet = total_workbook.active  # 新規作成したExcelのアクティブなシートを取得する


        for row_idx in range(3, last_row + 1):
            host_name = stat_sheet[f"{column_char}{row_idx}"].value  # 統計記入ファイルからホスト名取得

            if host_name:
                for row in range(7, 84):
                    find = 0
                    total_host_name = total_sheet[f"B{row}"].value  # 利用集計ファイルのB列（ホスト名）を取得
                    
                    if host_name == total_host_name:
                        value = total_sheet[f"{column_head}{row}"].value  # `column_head` の列の値を取得
                        stat_sheet[f"{prev_col}{row_idx}"].value = value # 統計記入ファイルの前月の列に記入
                        stat_sheet[f"Q{row_idx}"].value = value # 統計記入ファイルのQ列に記入（なのでRicohスキャナ統計はcompare_monthは実行しなくてよい）

                        print(f"{host_name} の {column_head} = {value} を {prev_col} と Q の {row_idx} に記入")
                        find = 1
                        break
                        
                if find == 0:
                    print(f"{host_name} が見つかりませんでした")


    stat_workbook.save(f"./statistics_report/{report_file}")
    print(f"{report_file}に保存完了")



def compare_month(report_file, sheet_name, prev_column, two_column, column):
    
    stat_file = excel_hadler(f"./statistics_report/{report_file}", sheet_name)
    stat_workbook = stat_file[0] # 統計レポートのExcelファイル
    stat_sheet = stat_file[1] # 統計レポートExcelの記入対象シート    
    last_row = stat_file[2] # B列の,値がある最後の行

    # 前月が4月の場合、別シートの該当列と計算
    if prev_month != 4:
        print(f"{report_file}: {prev_column}列 - {two_column}列 = 記入先: {column}列")

        for row_idx in range(3, last_row + 1):
            formula = f"={prev_column}{row_idx} - {two_column}{row_idx}"
            stat_sheet[f"{column}{row_idx}"].value = formula  

    else:
        old_sheet_name = f"{sheet_name}_OLD"  # 旧シートの名前
        print(f"{report_file}: {prev_column}列 - {old_sheet_name} の {two_column}列 = 記入先: {column}列")

        for row_idx in range(3, last_row + 1):
            formula = f"={prev_column}{row_idx} - {old_sheet_name}!{two_column}{row_idx}"  # 旧シートの "two_column" 列と計算
            stat_sheet[f"{column}{row_idx}"].value = formula

    stat_workbook.save(f"./statistics_report/{report_file}")
    print(f"{report_file}に保存完了")



def fill_in_report(report_file):
    print(f"{report_file}の記入処理を開始します")


    if report_file != "教員カラープリンタ印刷統計.xlsx": # ロビープリンタ印刷統計 教室等モノクロプリンタ印刷統計 Ricohスキャナ統計 の場合

        # 前月の列を取得
        print(f"前月 {prev_month} 月のデータを {prev_col_4E} 列に記入します")

        if report_file != "Ricohスキャナ統計.xlsx": # ロビープリンタ印刷統計 教室等モノクロプリンタ印刷統計 の場合

            fetch_data(report_file, "ALL", prev_col_4E, "D", "Printer: B&W") # (report_file, 統計レポートのシート名, prev_col, 統計レポートのホスト名の列, 集計レポートの対象ヘッダ名)

            compare_month(report_file, "ALL", prev_col_4E, two_col_4E, "Q")

        else: # Ricohスキャナ統計.xlsx の場合
            fetch_data(report_file, "ロビーカラー", prev_col_4E, "D", "AK")
            fetch_data(report_file, "ロビーモノクロ", prev_col_4E, "D", "AL")


    else: # 教員カラープリンタ印刷統計 の場合

        # 前月の列を取得
        print(f"前月 {prev_month} 月のデータを {prev_col_4D} 列に記入します")

        fetch_data(report_file, "モノクロ", prev_col_4D, "C", "Printer: B&W")
        fetch_data(report_file, "カラー", prev_col_4D, "C", "Printer: Full Color")
        
        compare_month(report_file, "モノクロ", prev_col_4D, two_col_4D, "P")
        compare_month(report_file, "カラー", prev_col_4D, two_col_4D, "P")

    print("--------------------")



fill_in_report("ロビープリンタ印刷統計.xlsx")
fill_in_report("教員カラープリンタ印刷統計.xlsx")
fill_in_report("教室等モノクロプリンタ印刷統計.xlsx")
fill_in_report("Ricohスキャナ統計.xlsx")
