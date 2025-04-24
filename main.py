import glob
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime


current_month = datetime.now().month
prev_month = current_month - 1 if current_month > 1 else 12
two_month_ago = prev_month - 1 if prev_month > 1 else 12


def excel_handler(report_file, sheet_name=None):
    # sheet
    workbook = load_workbook(report_file, data_only=True)
    sheet = workbook[sheet_name] if sheet_name else workbook.active

    # last low
    host_last_row = sum_last_row = sheet.max_row
    while sheet.cell(row=host_last_row, column=2).value is None and host_last_row > 1: # C列の,値がある最後の行を取得。P,Q列の記入時に使用する
        host_last_row -= 1
    while sheet.cell(row=sum_last_row, column=4).value is None and sum_last_row > 1: # E列の,値がある最後の行を取得。メールレポート作成時に使用する
        sum_last_row -= 1

    # column
    if report_file != "./statistics_report/教員カラープリンタ印刷統計.xlsx":
        month_col = { 4: "E", 5: "F", 6: "G", 7: "H", 8: "I", 9: "J", 10: "K", 11: "L", 12: "M", 1: "N", 2: "O", 3: "P" } #月と列の対応を定義
        prev_month_col = month_col[prev_month]
        two_month_col = month_col[two_month_ago]
        host_col = [ 3, "D" ]
        diff_col = [ 16, "Q" ]
    else: # 教員カラープリンタ印刷統計 の場合
        month_col = { 4: "D", 5: "E", 6: "F", 7: "G", 8: "H", 9: "I", 10: "J", 11: "K", 12: "L", 1: "M", 2: "N", 3: "O" } #教員カラープリンタ用
        prev_month_col = month_col[prev_month]
        two_month_col = month_col[two_month_ago]
        host_col = [ 2, "C" ]
        diff_col = [ 15, "P" ]

    # header
    excel_head = csv_head = None
    if report_file != "./statistics_report/教員カラープリンタ印刷統計.xlsx":
        if report_file != "./statistics_report/Ricohスキャナ統計.xlsx": # ロビープリンタ印刷統計 教室等モノクロプリンタ印刷統計 の場合
            csv_head = "Printer: B&W"

        else: # Ricohスキャナ統計.xlsx の場合
            if "カラー" in sheet_name:
                excel_head = "AK"
            if "モノクロ" in sheet_name:
                excel_head = "AL"

    else: # 教員カラープリンタ印刷統計 の場合
        if "カラー" in sheet_name:
            csv_head = "Printer: Full Color"
        if "モノクロ" in sheet_name:
            csv_head = "Printer: B&W"

    return report_file, sheet_name, workbook, sheet, host_last_row, sum_last_row, prev_month_col, two_month_col, host_col, diff_col, csv_head, excel_head



class commander:
    def __init__(self, report_file, sheet_name, workbook, sheet, host_last_row, sum_last_row, prev_month_col, two_month_col, host_col, diff_col, csv_head, excel_head):
        self.report_file = report_file
        self.sheet_name = sheet_name
        self.workbook = workbook
        self.sheet = sheet
        self.host_last_row = host_last_row
        self.sum_last_row = sum_last_row
        self.prev_month_col = prev_month_col
        self.two_month_col = two_month_col
        self.host_col = host_col
        self.diff_col = diff_col
        self.csv_head = csv_head
        self.excel_head = excel_head
        

    def fill_in_report(self):
        print(f"{self.report_file}の処理を開始します")

        if self.report_file != "./statistics_report/Ricohスキャナ統計.xlsx":

            file = glob.glob("./number_report/最新の機器カウンターレポート*")[0]
            reference_file = file.replace(".csv", "_utf8.csv")
            df = pd.read_csv(file, encoding="shift_jis", comment="#", skip_blank_lines=True) # コメント行と空白行を無視する
            df.to_csv(reference_file, encoding="utf-8", index=False)
            df_utf8 = pd.read_csv(reference_file, encoding="utf-8")

            for row_idx in range(3, self.host_last_row + 1):
                host_name = self.sheet[f"{self.host_col[1]}{row_idx}"].value  # 統計レポートからホスト名取得

                if host_name:
                    if host_name in df_utf8["Host Name"].values: # 集計レポートでホスト名を検索
                        value = df_utf8[df_utf8["Host Name"] == host_name][self.csv_head].values[0] # host_nameと一致する行のcsv_head列の値を取得
                    
                    else:
                        print(f"{host_name} が見つかりませんでした")
                        if prev_month != 4:
                            value = self.sheet[f"{self.two_month_col}{row_idx}"].value  # 先々月の列の値を取得
                        else:
                            old_sheet = excel_handler(self.report_file, f"{self.sheet_name}_OLD")[3] # 4月の記入時はOLDシートの3月の数値を参照
                            value = old_sheet[f"{self.two_month_col}{row_idx}"].value

                    self.sheet[f"{self.prev_month_col}{row_idx}"].value = value  # 前月の列に書き込み
                    print(f"{host_name} の {self.csv_head} = {value} を {self.prev_month_col}{row_idx} に記入")

        else: # Ricohスキャナ統計.xlsxの場合 プリンタサーバから出力されるxlsファイルが破損しているため，手動で新規xlsxファイルにコピー＆ペーストする。
            reference_file = glob.glob("./number_report/機能×カラー別集計レポート*.xlsx")[0]
            reference_sheet = excel_handler(reference_file)[3]

            for row_idx in range(3, self.host_last_row + 1):
                host_name = self.sheet[f"{self.host_col[1]}{row_idx}"].value

                if host_name:
                    for row in range(7, 84):
                        find = 0
                        total_host_name = reference_sheet[f"B{row}"].value  # 利用集計ファイルのB列（ホスト名）を取得
                        
                        if host_name == total_host_name:

                            value = reference_sheet[f"{self.excel_head}{row}"].value
                            self.sheet[f"{self.prev_month_col}{row_idx}"].value = value # 統計記入ファイルの前月の列に記入
                            self.sheet[f"Q{row_idx}"].value = value # 統計記入ファイルのQ列に記入（なのでRicohスキャナ統計はcompare_monthは実行しなくてよい）

                            print(f"{host_name} の {self.excel_head} = {value} を {self.prev_month_col} と Q の {row_idx} に記入")
                            find = 1
                            break
                            
                    if find == 0:
                        print(f"{host_name} が見つかりませんでした")
                        if prev_month != 4:
                            value = self.sheet[f"{self.two_month_col}{row_idx}"].value  # 先々月の列の値を取得
                        else:
                            old_sheet = excel_handler(self.report_file, f"{self.sheet_name}_OLD")[3] # 4月の記入時はOLDシートの3月の数値を参照
                            value = old_sheet[f"{self.two_month_col}{row_idx}"].value


        if prev_month != 4:
            print(f"{self.report_file}: {self.prev_month_col}列 - {self.two_month_col}列 = 記入先: {self.diff_col[1]}列")

            for row_idx in range(3, self.host_last_row + 1):

                prev_value = self.sheet[f"{self.prev_month_col}{row_idx}"].value
                two_value = self.sheet[f"{self.two_month_col}{row_idx}"].value

                diff_value = (prev_value if prev_value else 0) - (two_value if two_value else 0)
                self.sheet[f"{self.diff_col[1]}{row_idx}"].value = diff_value

        else:
            old_sheet_name = f"{self.sheet_name}_OLD"
            old_sheet = self.workbook[old_sheet_name]
            print(f"{self.report_file}: {self.prev_month_col}列 - {old_sheet_name} の {[self.two_month_col]}列 = 記入先: {self.diff_col[1]}列")

            for row_idx in range(3, self.host_last_row + 1):

                prev_value = self.sheet[f"{self.prev_month_col}{row_idx}"].value
                two_value = old_sheet[f"{self.two_month_col}{row_idx}"].value

                diff_value = (prev_value if prev_value else 0) - (two_value if two_value else 0)
                self.sheet[f"{self.diff_col[1]}{row_idx}"].value = diff_value

        self.workbook.save(self.report_file)
        print("----------")



    def gen_text(self):
        print(f"{self.report_file}のテキストを作成します")

        prev_value = self.sheet[f"{self.prev_month_col}{self.sum_last_row}"].value # 前月の値を取得
        
        if prev_month != 4:
            two_value = self.sheet[f"{self.two_month_col}{self.sum_last_row}"].value
        else:
            old_sheet = excel_handler(self.report_file, f"{self.sheet_name}_OLD")[3]
            two_value = old_sheet[f"{self.two_month_col}{self.sum_last_row}"].value

        prev_value = int(prev_value) if prev_value else 0  # Noneや文字列を整数に変換
        two_value = int(two_value) if two_value else 0

        if prev_value > two_value:
            word = "増加"
            if prev_value - two_value >= 20000:
                word = "大幅に増加"

        elif prev_value < two_value:
            word = "減少"
            if two_value - prev_value >= 20000:
                word = "大幅に減少"
            elif prev_value == 0:
                word = "減少（利用なし）"

        else:
            word = "変化なし"
            if prev_value == 0:
                word = "変化なし（利用なし）"

        # Q列またはP列の最大値を持つ行を探す
        max_value = float('0') # 初期化
        max_row_num = None

        # 最大値を持つ行を特定
        for row in self.sheet.iter_rows(min_row=2, max_row=self.sum_last_row, min_col=self.diff_col[0], max_col=self.diff_col[0]):
            value = row[0].value

            if value and isinstance(value, (int, float)):
            
                if value > max_value:
                    max_value = value
                    max_row_num = row[0].row


        if max_row_num: # 最大値を持つ行の 建屋，部屋名 を取得

            if self.report_file != "./statistics_report/ロビープリンタ印刷統計.xlsx":
                place = f"{self.sheet[f'A{max_row_num}'].value} {self.sheet[f'B{max_row_num}'].value}"
            
            else: # ロビープリンタ印刷統計.xlsxの場合はA,B,C列を取得
                place = f"{self.sheet[f'A{max_row_num}'].value} {self.sheet[f'B{max_row_num}'].value} {self.sheet[f'C{max_row_num}'].value}"
        else:
            place = "該当なし"
        print("----------")

        return word, place, max_value



# create instance
class_room = commander(*excel_handler("./statistics_report/教室等モノクロプリンタ印刷統計.xlsx", "ALL"))
teacher_color = commander(*excel_handler("./statistics_report/教員カラープリンタ印刷統計.xlsx", "カラー"))
teacher_mono = commander(*excel_handler("./statistics_report/教員カラープリンタ印刷統計.xlsx", "モノクロ"))
lobby = commander(*excel_handler("./statistics_report/ロビープリンタ印刷統計.xlsx", "ALL"))
ricoh_class_color = commander(*excel_handler("./statistics_report/Ricohスキャナ統計.xlsx", "教室カラー"))
ricoh_class_mono = commander(*excel_handler("./statistics_report/Ricohスキャナ統計.xlsx", "教室モノクロ"))
ricoh_lobby_color = commander(*excel_handler("./statistics_report/Ricohスキャナ統計.xlsx", "ロビーカラー"))
ricoh_lobby_mono = commander(*excel_handler("./statistics_report/Ricohスキャナ統計.xlsx", "ロビーモノクロ"))

# fill in report
instances = [ricoh_lobby_color, ricoh_lobby_mono, ricoh_class_color, ricoh_class_mono, lobby, teacher_color, teacher_mono, class_room]
for instance in instances:
    instance.fill_in_report()

# create mail report text
class_print = class_room.gen_text()
teacher_print_color = teacher_color.gen_text()
teacher_print_mono = teacher_mono.gen_text()
lobby_print = lobby.gen_text()
class_scan_color = ricoh_class_color.gen_text()
class_scan_mono = ricoh_class_mono.gen_text()
lobby_scan_color = ricoh_lobby_color.gen_text()
lobby_scan_mono = ricoh_lobby_mono.gen_text()

report_text = f"・プリンタ，スキャナ\n\
当月は, 教室プリンタは{class_print[0]}, \
教員室プリンタは, カラーが{teacher_print_color[0]}, 白黒が{teacher_print_mono[0]}。\
ロビープリンタは{lobby_print[0]}, \
教室プリンタのスキャナは, カラーが{class_scan_color[0]}, 白黒が{class_scan_mono[0]}。\
ロビープリンタのスキャナは, カラーが{lobby_scan_color[0]}, 白黒が{lobby_scan_mono[0]}。\n\
\
特に目立った利用は，\
教室プリンタ ( {class_print[1]} → {class_print[2]} 枚 ), \
教員室プリンタ ( カラー: {teacher_print_color[1]} → {teacher_print_color[2]} 枚，白黒: {teacher_print_mono[1]} → {teacher_print_mono[2]} 枚 ), \
ロビープリンタ ( {lobby_print[1]} → {lobby_print[2]} 枚 ), \
教室プリンタのスキャナ ( カラー: {class_scan_color[1]} → {class_scan_color[2]} 枚，白黒: {class_scan_mono[1]} → {class_scan_mono[2]} 枚 ), \
ロビープリンタのスキャナ ( カラー: {lobby_scan_color[1]} → {lobby_scan_color[2]} 枚，白黒: {lobby_scan_mono[1]} → {lobby_scan_mono[2]} 枚 )だった。"

print(report_text)
