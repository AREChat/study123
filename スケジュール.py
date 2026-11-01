import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

# ==========================================
# 🔧 設定エリア
# ==========================================
TARGET_YEAR = 2026   # 年
TARGET_MONTH = 12     # 月
LOG_ROWS = 20        # 行数
FILE_NAME = f"{TARGET_YEAR}年{TARGET_MONTH}月_毎日のスケジュール.xlsx"
# ==========================================

def create_stylish_schedule(year, month):
    # 日付リスト生成
    start_date = f"{year}-{month}-01"
    dates = pd.date_range(start=start_date, periods=32, freq='D')
    dates = [d for d in dates if d.month == month]
    japanese_days = ["月", "火", "水", "木", "金", "土", "日"]
    
    wb = openpyxl.Workbook()
    
    # --- デザイン定義 ---
    # フォント
    font_main = Font(name='Meiryo UI', size=11)
    font_bold = Font(name='Meiryo UI', size=11, bold=True)
    font_header_lg = Font(name='Meiryo UI', size=14, bold=True)
    font_link = Font(name='Meiryo UI', size=10, color="0000FF", underline="single")
    font_white_bold = Font(name='Meiryo UI', size=11, bold=True, color="FFFFFF")
    
    # 罫線
    border_thin = Border(left=Side(style='thin', color="BFBFBF"), 
                         right=Side(style='thin', color="BFBFBF"), 
                         top=Side(style='thin', color="BFBFBF"), 
                         bottom=Side(style='thin', color="BFBFBF"))
    border_box = Border(left=Side(style='medium', color="4472C4"), 
                        right=Side(style='medium', color="4472C4"), 
                        top=Side(style='medium', color="4472C4"), 
                        bottom=Side(style='medium', color="4472C4"))

    # 塗りつぶし色
    fill_header_navy = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid") # 濃い青
    fill_sub_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")  # 薄い青
    fill_work_green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")  # 仕事(緑)
    fill_rest_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # 休み(黄)
    fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")        # 無効(グレー)

    # ==========================================
    # 1. 使い方・注意事項シート
    # ==========================================
    ws_readme = wb.active
    ws_readme.title = "使い方・注意事項"
    ws_readme.sheet_view.showGridLines = False # 目盛り線を消して綺麗に
    
    ws_readme['A1'] = "📖 スケジュール管理表の使い方"
    ws_readme['A1'].font = Font(name='Meiryo UI', size=18, bold=True, color="44546A")
    
    instructions = [
        ("手順1", "「目次(設定)」シートでシフトを登録します。"),
        ("", "・C列で「仕事」「休み」を選択します。"),
        ("", "・休みを選ぶと時間は自動で「-」になります。仕事の場合は時間を入力してください。"),
        ("手順2", "日付ごとのシートへ移動します。", "目次の日付リンクをクリックして移動できます。"),
        ("手順3", "作業ログを入力します。", "開始・終了・内容を入力すると、右端に「OK」が表示されます。"),
        ("便利機能", "各シートの左上にある「🔙 目次へ戻る」を押すと、すぐに目次に戻れます。")
    ]
    
    current_row = 3
    for item in instructions:
        title = item[0]
        desc = item[1]
        
        if title:
            cell_t = ws_readme.cell(row=current_row, column=1, value=title)
            cell_t.font = Font(name='Meiryo UI', size=11, bold=True, color="FFFFFF")
            cell_t.fill = fill_header_navy
            cell_t.alignment = Alignment(horizontal='center', vertical='center')
        
        cell_d = ws_readme.cell(row=current_row, column=2, value=desc)
        cell_d.font = font_main
        cell_d.alignment = Alignment(vertical='center')
        
        ws_readme.row_dimensions[current_row].height = 20
        current_row += 1

    ws_readme.column_dimensions['A'].width = 15
    ws_readme.column_dimensions['B'].width = 80

    # ==========================================
    # 2. 目次(設定)シート
    # ==========================================
    ws_index = wb.create_sheet("目次(設定)")
    ws_index.sheet_view.showGridLines = False
    
    # ヘッダー作成
    headers = ["日付", "曜日", "シフト(選択)", "勤務開始", "勤務終了", "今日の目標"]
    for col, text in enumerate(headers, 1):
        cell = ws_index.cell(row=1, column=col, value=text)
        cell.font = font_white_bold
        cell.fill = fill_header_navy
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border_thin

    # プルダウン設定
    dv = DataValidation(type="list", formula1='"仕事,休み,その他"', allow_blank=True)
    ws_index.add_data_validation(dv)
    dv.add(f"C2:C{len(dates)+1}")

    for i, date in enumerate(dates):
        r = i + 2
        sheet_name = f"{date.month}月{date.day}日"
        day_str = japanese_days[date.weekday()]
        
        # A列: リンク
        c_date = ws_index.cell(row=r, column=1, value=sheet_name)
        c_date.hyperlink = f"#'{sheet_name}'!A1"
        c_date.font = font_link
        c_date.alignment = Alignment(horizontal='center')
        c_date.border = border_thin
        
        # B列: 曜日
        c_day = ws_index.cell(row=r, column=2, value=day_str)
        c_day.alignment = Alignment(horizontal='center')
        c_day.border = border_thin
        # 土日カラー
        if date.weekday() == 5: c_day.font = Font(name='Meiryo UI', color="0000FF", bold=True)
        if date.weekday() == 6: c_day.font = Font(name='Meiryo UI', color="FF0000", bold=True)

        # C列: シフト
        c_shift = ws_index.cell(row=r, column=3)
        c_shift.alignment = Alignment(horizontal='center')
        c_shift.border = border_thin

        # D, E列: 時刻数式
        formula = f'=IF(C{r}="休み", "-", "")'
        
        c_start = ws_index.cell(row=r, column=4, value=formula)
        c_start.alignment = Alignment(horizontal='center')
        c_start.border = border_thin
        
        c_end = ws_index.cell(row=r, column=5, value=formula)
        c_end.alignment = Alignment(horizontal='center')
        c_end.border = border_thin

        # F列: 今日の目標
        c_goal = ws_index.cell(row=r, column=6)
        c_goal.border = border_thin
        
    # 条件付き書式 (視認性向上)
    data_range = f"A2:E{len(dates)+1}"
    ws_index.conditional_formatting.add(data_range, FormulaRule(formula=['$C2="仕事"'], fill=fill_work_green))
    ws_index.conditional_formatting.add(data_range, FormulaRule(formula=['$C2="休み"'], fill=fill_rest_yellow))
    # 休みの場合、時刻欄をグレーアウト
    ws_index.conditional_formatting.add(f"D2:E{len(dates)+1}", FormulaRule(formula=['$C2="休み"'], fill=fill_gray))

    # 列幅調整
    ws_index.column_dimensions['A'].width = 15
    ws_index.column_dimensions['B'].width = 6
    ws_index.column_dimensions['C'].width = 12
    ws_index.column_dimensions['D'].width = 12
    ws_index.column_dimensions['E'].width = 12
    ws_index.column_dimensions['F'].width = 40
    ws_index.freeze_panes = "A2"

    # ==========================================
    # 3. 各日付の詳細シート
    # ==========================================
    for i, date in enumerate(dates):
        idx_row = i + 2
        sheet_name = f"{date.month}月{date.day}日"
        ws = wb.create_sheet(sheet_name)
        ws.sheet_view.showGridLines = False # 方眼紙っぽさを消す
        
        # --- 1行目: メインヘッダー ---
        ws.merge_cells("A1:C1")
        cell_title = ws['A1']
        cell_title.value = f"{year}年{date.month}月{date.day}日 ({japanese_days[date.weekday()]})"
        cell_title.font = font_header_lg
        cell_title.alignment = Alignment(horizontal='center', vertical='center') # 中央揃え
        
        # 土日カラー
        if date.weekday() == 5: cell_title.font = Font(name='Meiryo UI', size=14, bold=True, color="0000FF")
        if date.weekday() == 6: cell_title.font = Font(name='Meiryo UI', size=14, bold=True, color="FF0000")

        # 参照数式
        ws['D1'].value = f"='目次(設定)'!C{idx_row}"
        ws['E1'].value = f"='目次(設定)'!D{idx_row}"
        ws['F1'].value = f"='目次(設定)'!E{idx_row}"
        
        # デザイン適用
        for c_char in ['D', 'E', 'F']:
            cell = ws[f'{c_char}1']
            cell.font = font_bold
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border_box
            if c_char in ['E', 'F']: cell.number_format = 'h:mm'

        # ヘッダー背景色の条件付き書式
        header_range = "A1:F1"
        ws.conditional_formatting.add(header_range, FormulaRule(formula=['$D$1="仕事"'], fill=fill_work_green))
        ws.conditional_formatting.add(header_range, FormulaRule(formula=['$D$1="休み"'], fill=fill_rest_yellow))


    # --- 2行目: 今日の目標 ---
        ws.merge_cells("A2:C2")
        goal = ws["A2"]
        goal.value = f'=IF(\'目次(設定)\'!F{idx_row}="","今日の目標：","今日の目標："&\'目次(設定)\'!F{idx_row})'
        goal.font = Font(name="Meiryo UI", size=12, bold=True)
        goal.alignment = Alignment(horizontal="left", vertical="center")
        ws["A3"].value = "🔙 目次へ戻る"
        ws["A3"].hyperlink = "#'目次(設定)'!A1"
        ws["A3"].font = Font(name="Meiryo UI", size=9, color="808080", underline="single")
        ws['A2'].alignment = Alignment(horizontal='left', vertical='top')

        # --- 3行目: テーブル項目 ---
        table_headers = ["開始時刻", "終了時刻", "スケジュール内容", "完了"]
        for col, text in enumerate(table_headers, 1):
            cell = ws.cell(row=4, column=col, value=text)
            cell.font = font_white_bold
            cell.fill = fill_header_navy
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border_thin

        # --- 5行目以降: ログ入力欄 ---
        start_row = 5
        for r_idx in range(LOG_ROWS):
            r = start_row + r_idx
            
            # 罫線とフォント適用
            for c in range(1, 5):
                cell = ws.cell(row=r, column=c)
                cell.border = border_thin
                cell.font = font_main
            
            # 配置
            ws.cell(row=r, column=1).alignment = Alignment(horizontal='center') # 開始
            ws.cell(row=r, column=2).alignment = Alignment(horizontal='center') # 終了
            ws.cell(row=r, column=3).alignment = Alignment(horizontal='left')   # 内容
            ws.cell(row=r, column=4).alignment = Alignment(horizontal='center') # 完了
            
            # OK判定数式
            ws.cell(row=r, column=4).value = f'=IF(AND(A{r}<>"", B{r}<>"", C{r}<>""), "OK", "")'

        # 完了列（D列）でOKが出たら太字にする条件付き書式
            ws.conditional_formatting.add(f"D4:D{start_row+LOG_ROWS}", 
                                      FormulaRule(formula=['D4="OK"'], font=Font(bold=True, color="008000")))

        # --- 列幅とウィンドウ枠固定 ---
        ws.column_dimensions['A'].width = 13
        ws.column_dimensions['B'].width = 13
        ws.column_dimensions['C'].width = 50
        ws.column_dimensions['D'].width = 13 # シフト/完了
        ws.column_dimensions['E'].width = 15 # 開始
        ws.column_dimensions['F'].width = 15 # 終了
            
        ws.freeze_panes = "A5"

        # タブの色分け
        if date.weekday() == 5: ws.sheet_properties.tabColor = "0000FF" # 土
        elif date.weekday() == 6: ws.sheet_properties.tabColor = "FF0000" # 日

    wb.save(FILE_NAME)
    print(f"✅ 作成完了: {FILE_NAME}")
    print("すべての要望を反映し、デザインを調整しました。")

if __name__ == "__main__":
    create_stylish_schedule(TARGET_YEAR, TARGET_MONTH)