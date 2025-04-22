import glob
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

current_month = datetime.now().month
prev_month = current_month - 1 if current_month > 1 else 12
two_month_ago = prev_month - 1 if prev_month > 1 else 12


def columns(report_file):

    month_column_4E = { 4: "E", 5: "F", 6: "G", 7: "H", 8: "I", 9: "J", 10: "K", 11: "L", 12: "M", 1: "N", 2: "O", 3: "P" } #月と列の対応を定義
    month_column_4D = { 4: "D", 5: "E", 6: "F", 7: "G", 8: "H", 9: "I", 10: "J", 11: "K", 12: "L", 1: "M", 2: "N", 3: "O" } #教員カラープリンタ用
    prev_col_4E = month_column_4E[prev_month] # 先月（記入対象）列
    prev_col_4D = month_column_4D[prev_month]
    two_col_4E = month_column_4E[two_month_ago] # 先々月（比較対象）列
    two_col_4D = month_column_4D[two_month_ago]

    if report_file != "教員カラープリンタ印刷統計.xlsx":
        prev_col = prev_col_4E
        two_col = two_col_4E
        diff_col = "Q"
        host_col = "D"
    else:
        prev_col = prev_col_4D
        two_col = two_col_4D
        diff_col = "P"
        host_col = "C"

    return prev_col, two_col, diff_col, host_col



def file_handler(report_file): # 統計記入ファイルに対応した利用集計ファイルを返す。CSVはUTF-8に変換する
    
    if report_file != "Ricohスキャナ統計.xlsx":
        print("最新の機器カウンターレポートから値を取得します")

        file = glob.glob("./number_report/最新の機器カウンターレポート*")
        csv_file = file[0]
        utf8_csv_path = csv_file.replace(".csv", "_utf8.csv")
        df = pd.read_csv(csv_file, encoding="shift_jis", comment="#", skip_blank_lines=True) # コメント行と空白行を無視する
        df.to_csv(utf8_csv_path, encoding="utf-8", index=False)

        return utf8_csv_path
    
    else:
        print("機能×カラー別集計レポートから値を取得します")

        file = glob.glob("./number_report/機能×カラー別集計レポート*.xlsx")
        excel_file = file[0]

        return excel_file
    


def excel_hadler(report_file, sheet_name=None, data_only=False): # excelデータをインスタンス化したものと，指定の列の値がある最後の行を返す
    if data_only:
        workbook = load_workbook(report_file, data_only=True)
    else:
        workbook = load_workbook(report_file)

    if sheet_name:
        sheet = workbook[sheet_name]
    else:
        sheet = workbook.active

    host_last_row = sum_last_row = sheet.max_row # シートの最終行を設定

    while sheet.cell(row=host_last_row, column=2).value is None and host_last_row > 1: # C列の,値がある最後の行を取得。P,Q列の記入時に使用する
        host_last_row -= 1  # 空欄なら1行ずつ上にさかのぼる

    while sheet.cell(row=sum_last_row, column=4).value is None and sum_last_row > 1: # E列の,値がある最後の行を取得。メールレポート作成時に使用する
        sum_last_row -= 1

    return workbook, sheet, host_last_row, sum_last_row



def fetch_data(report_file, sheet_name, column_head):
    print(f"{sheet_name} シートに記入します")

    total_file = file_handler(report_file) # 利用集計ファイル
    stat_file = excel_hadler(f"./statistics_report/{report_file}", sheet_name)
    stat_workbook = stat_file[0] # 統計レポートのExcelファイル
    stat_sheet = stat_file[1] # 統計レポートExcelの記入対象シート
    last_row = stat_file[2] # C列の値がある最後の行
    column = columns(report_file)
    prev_col = column[0] # 前月の列
    host_col = column[3] # ホスト名の列

    if report_file != "Ricohスキャナ統計.xlsx":
        df_utf8 = pd.read_csv(total_file, encoding="utf-8")

        for row_idx in range(3, last_row + 1):  # 統計レポートの行を1つずつチェック
            host_name = stat_sheet[f"{host_col}{row_idx}"].value  # 統計レポートからホスト名取得

            if host_name:
                if host_name in df_utf8["Host Name"].values: # 集計レポートからホスト名を検索

                    value = df_utf8[df_utf8["Host Name"] == host_name][column_head].values[0] # host_nameと一致する行のcolumn_head列の値を取得
                    stat_sheet[f"{prev_col}{row_idx}"].value = value  # 前月の列に書き込み
                    print(f"{host_name} の {column_head} = {value} を {prev_col}{row_idx} に記入")
                
                else:
                    print(f"{host_name} が見つかりませんでした")


    else: # Ricohスキャナ統計.xlsxの場合
        # プリンタサーバから出力されるxlsファイルが破損しているため，手動で新規xlsxファイルにコピー＆ペーストする。
        # xlsファイルの機器別サマリーシートだけを，新規作成したxlsxファイルに貼り付け，ファイル名を同じものにする。

        total_sheet = excel_hadler(total_file)[1]

        for row_idx in range(3, last_row + 1):
            host_name = stat_sheet[f"{host_col}{row_idx}"].value  # 統計記入ファイルからホスト名取得

            if host_name:
                for row in range(7, 84):
                    find = 0
                    total_host_name = total_sheet[f"B{row}"].value  # 利用集計ファイルのB列（ホスト名）を取得
                    
                    if host_name == total_host_name:

                        value = total_sheet[f"{column_head}{row}"].value  # column_head の列の値を取得
                        stat_sheet[f"{prev_col}{row_idx}"].value = value # 統計記入ファイルの前月の列に記入
                        stat_sheet[f"Q{row_idx}"].value = value # 統計記入ファイルのQ列に記入（なのでRicohスキャナ統計はcompare_monthは実行しなくてよい）

                        print(f"{host_name} の {column_head} = {value} を {prev_col} と Q の {row_idx} に記入")
                        find = 1
                        break
                        
                if find == 0:
                    print(f"{host_name} が見つかりませんでした")

    stat_workbook.save(f"./statistics_report/{report_file}")



def compare_month(report_file, sheet_name):

    stat_file = excel_hadler(f"./statistics_report/{report_file}", sheet_name)
    stat_workbook = stat_file[0]  # 統計レポートのExcelファイル
    stat_sheet = stat_file[1]  # 統計レポートExcelの記入対象シート
    last_row = stat_file[2]  # C列の値がある最後の行
    column = columns(report_file)
    prev_col = column[0]
    two_col = column[1]
    diff_col = column[2]

    # 前月が4月の場合、別シートの該当列と計算
    if prev_month != 4:
        print(f"{report_file}: {prev_col}列 - {two_col}列 = 記入先: {diff_col}列")

        for row_idx in range(3, last_row + 1):
            prev_value = stat_sheet[f"{prev_col}{row_idx}"].value
            two_value = stat_sheet[f"{two_col}{row_idx}"].value

            # 値を計算して直接代入
            diff_value = (prev_value if prev_value else 0) - (two_value if two_value else 0)
            stat_sheet[f"{diff_col}{row_idx}"].value = diff_value

    else:
        old_sheet_name = f"{sheet_name}_OLD"  # 旧シートの名前
        print(f"{report_file}: {prev_col}列 - {old_sheet_name} の {[two_col]}列 = 記入先: {diff_col}列")

        old_sheet = stat_workbook[old_sheet_name]  # 別シートを取得

        for row_idx in range(3, last_row + 1):
            prev_value = stat_sheet[f"{prev_col}{row_idx}"].value
            two_value = old_sheet[f"{two_col}{row_idx}"].value

            # 値を計算して直接代入
            diff_value = (prev_value if prev_value else 0) - (two_value if two_value else 0)
            stat_sheet[f"{diff_col}{row_idx}"].value = diff_value

    stat_workbook.save(f"./statistics_report/{report_file}")



def fill_in_report(report_file):
    print(f"{report_file}の記入処理を開始します")

    if report_file != "教員カラープリンタ印刷統計.xlsx": # ロビープリンタ印刷統計 教室等モノクロプリンタ印刷統計 Ricohスキャナ統計 の場合

        if report_file != "Ricohスキャナ統計.xlsx": # ロビープリンタ印刷統計 教室等モノクロプリンタ印刷統計 の場合

            fetch_data(report_file, "ALL", "Printer: B&W") # (report_file, 統計レポートのシート名, prev_col, 統計レポートのホスト名の列, 集計レポートの対象ヘッダ名)
            compare_month(report_file, "ALL")

        else: # Ricohスキャナ統計.xlsx の場合
            fetch_data(report_file, "ロビーカラー", "AK")
            fetch_data(report_file, "ロビーモノクロ", "AL")
            fetch_data(report_file, "教室カラー", "AL")
            fetch_data(report_file, "教室モノクロ", "AL")

    else: # 教員カラープリンタ印刷統計 の場合
        fetch_data(report_file, "モノクロ", "Printer: B&W")
        fetch_data(report_file, "カラー", "Printer: Full Color")
        compare_month(report_file, "モノクロ")
        compare_month(report_file, "カラー")
    print("--------------------")



def gen_text(report_file, sheet_name):

    workbook = excel_hadler(f"./statistics_report/{report_file}", sheet_name, data_only=True)
    ws = workbook[1]
    last_row = workbook[2] # E列の最終行
    column = columns(report_file) # 比較する列を取得
    two_col = column[1]
    prev_value = ws[f"{column[0]}{last_row}"].value # 前月の値を取得
    
    if prev_month != 4:
        two_value = ws[f"{two_col}{last_row}"].value
    else:
        old_ws = excel_hadler(f"./statistics_report/{report_file}", f"{sheet_name}_OLD", "E")[1]
        two_value = old_ws[f"{two_col}{last_row}"].value

    prev_value = prev_value if prev_value else 0 # Noneを0に変換
    two_value = two_value if two_value else 0

    if prev_value > two_value:
        word = "増加"
    elif prev_value < two_value:
        word = "減少"
        if prev_value == 0:
            word = "減少（利用なし）"
    else:
        word = "変化なし"
        if prev_value == 0:
            word = "変化なし（利用なし）"

    # Q列またはP列の最大値を持つ行を探す
    max_value = float('0')  # 初期化
    max_row_num = None  # 最大値の行番号を保持
    target_col = 17 if column[2] == "Q" else 16 # Q列の場合は17列目を，P列の場合は16列目を指定

    # 最大値を持つ行を特定
    for row in ws.iter_rows(min_row=2, max_row=last_row, min_col=target_col, max_col=target_col):
        value = row[0].value

        if value and isinstance(value, (int, float)):
        
            if value > max_value:
                max_value = value
                max_row_num = row[0].row

    if max_row_num: # 最大値を持つ行の A, B, C 列の値を取得
        
        if column[2] == "Q":
            place = f"{ws[f'A{max_row_num}'].value} {ws[f'B{max_row_num}'].value} {ws[f'C{max_row_num}'].value}"
        
        else:  # P列の場合は A, B のみ
            place = f"{ws[f'A{max_row_num}'].value} {ws[f'B{max_row_num}'].value}"
    else:
        place = "該当なし"

    return word, place, max_value



fill_in_report("Ricohスキャナ統計.xlsx")
fill_in_report("ロビープリンタ印刷統計.xlsx")
fill_in_report("教員カラープリンタ印刷統計.xlsx")
fill_in_report("教室等モノクロプリンタ印刷統計.xlsx")

class_print = gen_text("教室等モノクロプリンタ印刷統計.xlsx", "ALL")
teacher_print_color = gen_text("教員カラープリンタ印刷統計.xlsx", "モノクロ")
teacher_print_mono = gen_text("教員カラープリンタ印刷統計.xlsx", "カラー")
lobby_print = gen_text("ロビープリンタ印刷統計.xlsx", "ALL")
class_scan_color = gen_text("Ricohスキャナ統計.xlsx", "教室カラー")
lobby_scan_color = gen_text("Ricohスキャナ統計.xlsx", "ロビーカラー")
class_scan_mono = gen_text("Ricohスキャナ統計.xlsx", "教室モノクロ")
lobby_scan_mono = gen_text("Ricohスキャナ統計.xlsx", "ロビーモノクロ")

base_text = f"・プリンタ，スキャナ\n\
当月は, 教室プリンタは{class_print[0]} ,\
教員室プリンタは, カラーが{teacher_print_color[0]}, 白黒が{teacher_print_mono[0]}。\
ロビープリンタは{lobby_print[0]},\
教室プリンタのスキャナは, カラーが{class_scan_color[0]}, 白黒が{class_scan_mono[0]}。\
ロビープリンタのスキャナは, カラーが{lobby_scan_color[0]}, 白黒が{lobby_scan_mono[0]}。\n\
\
特に目立った利用は，\
教室プリンタ ( {class_print[1]} → {class_print[2]} 枚 ), \
教員室プリンタ ( カラー: {teacher_print_color[1]} → {teacher_print_color[2]} 枚，白黒: {teacher_print_mono[1]} → {teacher_print_mono[2]} 枚 ), \
ロビープリンタ ( {lobby_print[1]} → {lobby_print[2]} 枚 ), \
教室プリンタのスキャナ ( カラー: {class_scan_color[1]} → {class_scan_color[2]} 枚，白黒: {class_scan_mono[1]} → {class_scan_mono[2]} 枚 ), \
ロビープリンタのスキャナ ( カラー: {lobby_scan_color[1]} → {lobby_scan_color[2]} 枚，白黒: {lobby_scan_mono[1]} → {lobby_scan_mono[2]} 枚 )だった。"

print(base_text)
