import os
import sys
import json
import openpyxl
from openpyxl.styles import Font, PatternFill
import pulp
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ================= ベースパス取得関数 =================
def get_base_path():
    """実行環境のベースパスを取得（PyInstallerのexe化に対応）"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

# ================= 外部設定ファイルの読み込み =================
config_path = os.path.join(get_base_path(), 'config.json')
try:
    with open(config_path, 'r', encoding='utf-8') as f:
        CONFIG = json.load(f)
except FileNotFoundError:
    print(f"❌ エラー: 設定ファイル [{config_path}] が見つかりません。作成してください。")
    sys.exit(1)
except json.JSONDecodeError:
    print(f"❌ エラー: [{config_path}] のフォーマットが正しくありません。JSONの構文を確認してください。")
    sys.exit(1)

# 動的なファイル設定の取得
FILES = CONFIG.get('FILE_SETTINGS', {})
SPREADSHEET_ID = FILES.get('SPREADSHEET_ID', '')
DOWNLOAD_TASKS = [
    (FILES.get('TARGET_MONTH_SHEET', ''), FILES.get('TARGET_LOCAL_FILE', 'Book1.xlsx')),
    (FILES.get('PREV_MONTH_SHEET', ''), FILES.get('PREV_LOCAL_FILE', 'BookLM.xlsx'))
]

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
CREDENTIALS_FILE = os.path.join(get_base_path(), 'credentials.json')
TOKEN_FILE = os.path.join(get_base_path(), 'token_drive_exporter.json')


# ================= 1. Google Drive API ログイン＆ダウンロード =================
def get_drive_service():
    """Google Drive API 認証"""
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(
            f'❌ [{CREDENTIALS_FILE}] が見つかりません。Google Client Secret JSON'
            ' ファイルをexeと同じ階層に配置してください。'
        )

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES
            )
            creds = flow.run_local_server(
                port=0,
                authorization_prompt_message=(
                    'ポップアップしたブラウザで承認を完了してください...'
                ),
                success_message=(
                    '✅ 承認が成功しました！プログラムに戻ってください。'
                ),
            )

        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    return build('drive', 'v3', credentials=creds)

def fetch_and_split_sheets():
    """全体エクスポート＆特定シート抽出処理"""
    service = get_drive_service()
    temp_excel_path = os.path.join(get_base_path(), '_temp_full_export.xlsx')

    print(f'📥 クラウドからデータを取得中... (Spreadsheet ID: {SPREADSHEET_ID[:10]}...)')

    mime_type_xlsx = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    request = service.files().export_media(fileId=SPREADSHEET_ID, mimeType=mime_type_xlsx)

    with open(temp_excel_path, 'wb') as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()

    print('✅ クラウドからのエクスポートが完了しました。指定シートの抽出・分割を行います...\n')

    for target_sheet, output_filename in DOWNLOAD_TASKS:
        if not target_sheet: continue # シート名が空の場合はスキップ
        
        out_path = os.path.join(get_base_path(), output_filename)
        wb = openpyxl.load_workbook(temp_excel_path)

        if target_sheet not in wb.sheetnames:
            print(f'⚠️ スプレッドシート内に [{target_sheet}] というシートが見つかりませんでした。スキップします。')
            continue

        for sheet_name in wb.sheetnames:
            if sheet_name != target_sheet:
                del wb[sheet_name]

        wb.save(out_path)
        print(f'💾 シート [{target_sheet}] の書式を保持して [{output_filename}] に保存しました！')

    if os.path.exists(temp_excel_path):
        os.remove(temp_excel_path)

    print('\n🎉 すべての指定シートの抽出・保存が完了しました！\n')


# ================= 2. PuLP 数理最適化ソルバー =================
def solve_schedule_with_pulp(file_path, prev_month_file_path=None):
    file_path = os.path.join(get_base_path(), file_path)
    if prev_month_file_path:
        prev_month_file_path = os.path.join(get_base_path(), prev_month_file_path)

    print("🚀 【PuLP + HiGHS 数理最適化ソルバー】を起動中...\n")

    if not os.path.exists(file_path):
        print(f"❌ 指定されたファイルが見つかりません: [{file_path}]")
        return

    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    name_col, team_col, day_start_col, header_row = None, None, None, None
    for r in range(1, 10):
        for c in range(1, 20):
            val = ws.cell(row=r, column=c).value
            if val == '名前':
                name_col = c; header_row = r
            elif val == 'チーム':
                team_col = c
            elif str(val) == '1' and header_row and r == header_row:
                day_start_col = c

    target_rows, employees, teams, prefilled, red_border_cells = [], [], [], {}, set()

    def check_red_border(cell):
        border = cell.border
        if not border: return False
        for side in [border.left, border.right, border.top, border.bottom]:
            if side and side.color and side.color.rgb:
                if str(side.color.rgb).upper() == 'FFFF0000': return True
        return False

    seen_employees = set()
    for r in range(header_row + 2, ws.max_row + 1):
        team = ws.cell(row=r, column=team_col).value
        raw_name = ws.cell(row=r, column=name_col).value
        name = str(raw_name or '').strip().replace(' ', '').replace('\u3000', '')

        if name in seen_employees: break

        if team in ['NOC', 'CSC'] and name:
            seen_employees.add(name)
            target_rows.append(r)
            employees.append(name)
            teams.append(team)

            emp_idx = len(employees) - 1
            for d in range(30):
                c = day_start_col + d
                cell = ws.cell(row=r, column=c)
                val_str = str(cell.value).strip() if cell.value is not None else ''
                has_red = check_red_border(cell)
                if has_red: red_border_cells.add((emp_idx, d))

                if has_red: prefilled[(emp_idx, d)] = val_str if val_str else '休'
                elif val_str != '': prefilled[(emp_idx, d)] = val_str

    # 前月データの読み込み
    prev_month_aug31_night = set()
    prev_month_aug31_ake = set()
    prev_month_last5_work = {}

    def is_work_shift_name(val):
        val = str(val or '').strip()
        return val not in ['', '休', '年']

    if prev_month_file_path and os.path.exists(prev_month_file_path):
        wb_lm = openpyxl.load_workbook(prev_month_file_path, data_only=True)
        ws_lm = wb_lm[wb_lm.sheetnames[0]]
        
        name_col_l, day_start_col_l, header_row_l = None, None, None
        for r in range(1, 10):
            for c in range(1, 20):
                val = ws_lm.cell(row=r, column=c).value
                if val == '名前': name_col_l = c; header_row_l = r
                elif str(val) == '1' and header_row_l and r == header_row_l: day_start_col_l = c

        lm_shifts_map = {}
        for r in range(header_row_l + 2, ws_lm.max_row + 1):
            raw_name = ws_lm.cell(row=r, column=name_col_l).value
            name = str(raw_name or '').strip().replace(' ', '').replace('\u3000', '')
            if name and name not in lm_shifts_map:
                shifts_aug = [str(ws_lm.cell(row=r, column=day_start_col_l + d).value or '').strip() for d in range(31)]
                lm_shifts_map[name] = shifts_aug

        for e, name in enumerate(employees):
            aug_shifts = lm_shifts_map.get(name, [''] * 31)
            aug31_val = aug_shifts[30] if len(aug_shifts) >= 31 else ''
            if aug31_val in ['N', 'N設置作業']: prev_month_aug31_night.add(e)
            elif aug31_val == '明': prev_month_aug31_ake.add(e)
            last5 = aug_shifts[26:31] if len(aug_shifts) >= 31 else [''] * 5
            prev_month_last5_work[e] = [1 if is_work_shift_name(s) else 0 for s in last5]

        print(f'📊 前月データの読み込み完了: 前月最終日夜勤(N) {len(prev_month_aug31_night)} 名、明け(明) {len(prev_month_aug31_ake)} 名。')

    def is_red_or_pink_color(hex_color):
        if not hex_color or len(hex_color) < 6: return False
        hex_rgb = hex_color[-6:].upper()
        if hex_rgb in ['FFFFFF', '000000', '00FFFFFF']: return False
        try:
            r, g, b = int(hex_rgb[0:2], 16), int(hex_rgb[2:4], 16), int(hex_rgb[4:6], 16)
            if b > 210 and b >= r: return False
            if r > 180 and (r - g > 15) and (r - b > 15): return True
            if any(kw in hex_rgb for kw in ['F4CC', 'C0CB', 'D9D9', 'ECEC', 'FFC0', 'FFD']): return True
        except ValueError: pass
        return False

    def is_holiday_check(d):
        c = day_start_col + d
        if str(ws.cell(row=header_row + 1, column=c).value or '').strip() in ['土', '日', '土曜', '日曜', 'Sat', 'Sun']: return True
        for r_check in range(1, header_row + 3):
            fill = ws.cell(row=r_check, column=c).fill
            if fill and fill.start_color and fill.start_color.rgb and is_red_or_pink_color(str(fill.start_color.rgb).upper()): return True
        return False

    num_emp, num_days = len(employees), 30
    holiday_dates = [d for d in range(num_days) if is_holiday_check(d)]
    target_public_rests = len(holiday_dates) if CONFIG.get('USE_DYNAMIC_PUBLIC_REST') else CONFIG.get('FIXED_PUBLIC_REST')

    shifts = [0, 1, 2, 3, 4]
    trainees = CONFIG.get('TRAINEES', [])
    primary_mentors = CONFIG.get('PRIMARY_MENTORS', [])
    night_partner_map = CONFIG.get('NIGHT_PARTNER_MAP', {})

    noc_seniors = [e for e in range(num_emp) if teams[e] == 'NOC' and employees[e] not in trainees]
    csc_seniors = [e for e in range(num_emp) if teams[e] == 'CSC' and employees[e] not in trainees]
    all_seniors = noc_seniors + csc_seniors
    trainee_indices = [e for e in range(num_emp) if employees[e] in trainees]
    mentor_indices = [e for e in range(num_emp) if employees[e] in primary_mentors]

    def build_model(diagnose_mode=False):
        prob = pulp.LpProblem('Shift_Scheduling', pulp.LpMaximize)
        x = pulp.LpVariable.dicts('x', (range(num_emp), range(num_days), shifts), cat=pulp.LpBinary)
        y_d2 = pulp.LpVariable.dicts('y_d2', (mentor_indices, range(num_days)), cat=pulp.LpBinary)
        penalties, diag_vars = [], {}

        for e in range(num_emp):
            for d in range(num_days): prob += pulp.lpSum([x[e][d][s] for s in shifts]) == 1

        shift_map = {'休': 0, 'D': 1, 'D2': 1, 'D3': 1, 'メ': 1, '保全': 1, 'CPN\n設置作業日': 1, 'CPN設置作業日': 1, 'N': 2, 'N設置作業': 2, '明': 3, '年': 4, 'PM年\nor早上がり': 4}
        for e in range(num_emp):
            for d in range(num_days):
                if (e, d) in prefilled and prefilled[(e, d)] in shift_map:
                    prob += x[e][d][shift_map[prefilled[(e, d)]]] == 1
                elif (e, d) not in prefilled:
                    prob += x[e][d][4] == 0

        for e in range(num_emp):
            prob += pulp.lpSum([x[e][d][0] for d in range(num_days)]) == target_public_rests
            if e in prev_month_aug31_night: prob += x[e][0][3] == 1
            else:
                if (e, 0) not in prefilled or prefilled[(e, 0)] != '明': prob += x[e][0][3] == 0
            for d in range(num_days - 1): prob += x[e][d + 1][3] == x[e][d][2]
            
            if e in prev_month_aug31_ake and (e, 0) not in prefilled: prob += x[e][0][0] + x[e][0][4] == 1
            for d in range(num_days - 1):
                if not ((e, d + 1) in prefilled): prob += x[e][d + 1][0] + x[e][d + 1][4] >= x[e][d][3]

        for e in range(num_emp):
            last5_history = prev_month_last5_work.get(e, [0] * 5)
            for m in range(1, 6):
                prob += sum(last5_history[5 - (6 - m) : 5]) + pulp.lpSum([x[e][d_i][1] + x[e][d_i][2] + x[e][d_i][3] for d_i in range(m)]) <= 5
            for d in range(num_days - 5):
                prob += pulp.lpSum([x[e][d + i][1] + x[e][d + i][2] + x[e][d + i][3] for i in range(6)]) <= 5

        for d in range(num_days):
            prob += pulp.lpSum([x[e][d][2] for e in noc_seniors]) == 1
            prob += pulp.lpSum([x[e][d][2] for e in csc_seniors]) == 1
            if d > 0: prob += pulp.lpSum([x[e][d][3] for e in all_seniors]) == 2
            
            d2_sum = pulp.lpSum([y_d2[m][d] for m in mentor_indices])
            if d in holiday_dates: prob += d2_sum == 0
            else: prob += d2_sum == 1
            
            for m in mentor_indices:
                prob += y_d2[m][d] <= x[m][d][1]
                if (m, d) in prefilled and str(prefilled[(m, d)]).strip() in {'CPN\n設置作業日', 'CPN設置作業日', '保全', 'メ'}:
                    prob += y_d2[m][d] == 0

        # ペナルティ＆報酬の追加
        for tr_name, mentor_name in night_partner_map.items():
            if tr_name in employees and mentor_name in employees:
                tr_idx, m_idx = employees.index(tr_name), employees.index(mentor_name)
                for d in range(num_days):
                    s_np = pulp.LpVariable(f'slack_np_{tr_idx}_{d}', cat=pulp.LpBinary)
                    prob += x[tr_idx][d][2] - x[m_idx][d][2] <= s_np
                    penalties.append(CONFIG.get('PENALTY_NIGHT_PARTNER_MISSING', 50000) * s_np)

        return prob, x, y_d2

    prob, x, y_d2 = build_model()
    threads_count = os.cpu_count() or 4

    print("\n⚡ 高性能ソルバーで計算を実行中...")
    try:
        solver = pulp.HiGHS(msg=True, timeLimit=600, gapRel=0.01, threads=threads_count)
        status = prob.solve(solver)
    except Exception:
        solver = pulp.PULP_CBC_CMD(msg=True, timeLimit=600, gapRel=0.01, threads=threads_count)
        status = prob.solve(solver)

    if pulp.LpStatus[status] in ['Optimal', 'Feasible']:
        print('🎉 解の導出に成功しました！最適シフト表を作成しました。\n')
    else:
        print('❌ 解が見つかりません：条件が厳しすぎます（論理的な競合）。')
        return

    out_file = os.path.join(get_base_path(), FILES.get('PULP_OUTPUT_FILE', 'Book1_PuLP診断版出力結果.xlsx'))
    result_schedule = [['' for _ in range(num_days)] for _ in range(num_emp)]
    val_map = {0: '休', 1: 'D', 2: 'N', 3: '明', 4: '年'}

    for e in range(num_emp):
        for d in range(num_days):
            if (e, d) in prefilled and (e, d) not in red_border_cells:
                result_schedule[e][d] = prefilled[(e, d)]
            else:
                for s in shifts:
                    if pulp.value(x[e][d][s]) is not None and pulp.value(x[e][d][s]) > 0.5:
                        result_schedule[e][d] = val_map[s]

    for d in range(num_days):
        for tr_i in trainee_indices:
            if result_schedule[tr_i][d] == 'D': result_schedule[tr_i][d] = 'D3'
        for m in mentor_indices:
            if pulp.value(y_d2[m][d]) is not None and pulp.value(y_d2[m][d]) > 0.5:
                result_schedule[m][d] = 'D2'

    wb_out = openpyxl.load_workbook(file_path)
    ws_out = wb_out[sheet_name]
    for e in range(num_emp):
        for d in range(num_days):
            if (e, d) not in prefilled or (e, d) in red_border_cells:
                ws_out.cell(row=target_rows[e], column=day_start_col + d).value = result_schedule[e][d]

    wb_out.save(out_file)
    print(f'📁 中間シフト表の出力が完了しました: {out_file}\n')


# ================= 3. シフト表役割後処理エンジン =================
def post_process_schedule(input_file, output_file):
    input_file = os.path.join(get_base_path(), input_file)
    output_file = os.path.join(get_base_path(), output_file)

    print("🚀 【シフト表役割後処理エンジン】を起動中...\n")
    if not os.path.exists(input_file):
        print(f"❌ 入力ファイルが見つかりません: [{input_file}]")
        return

    wb = openpyxl.load_workbook(input_file)
    ws = wb[wb.sheetnames[0]]

    # 構造解析
    name_col, team_col, day_start_col, header_row = None, None, None, None
    for r in range(1, 10):
        for c in range(1, 20):
            val = ws.cell(row=r, column=c).value
            if val == '名前': name_col = c; header_row = r
            elif val == 'チーム': team_col = c
            elif str(val) == '1' and header_row and r == header_row: day_start_col = c

    num_days = 30
    holiday_dates = set(d for d in range(num_days) if (d % 7) in [4, 5]) # 簡易的な週末判定
    trainees_set = set(CONFIG.get('TRAINEES', []))

    noc_trainees, all_target_employees = [], []
    for r in range(header_row + 2, ws.max_row + 1):
        team = ws.cell(row=r, column=team_col).value
        name = str(ws.cell(row=r, column=name_col).value or '').strip().replace(' ', '')
        if team in ['NOC', 'CSC'] and name:
            all_target_employees.append((name, r))
            if team == 'NOC' and name in trainees_set:
                noc_trainees.append((name, r))

    # 実習生の在メ自動変換
    trainee_matrix = {name: [str(ws.cell(row=r, column=day_start_col + d).value or '').strip() for d in range(num_days)] for name, r in noc_trainees}
    total_converted = 0

    for d in range(num_days):
        working_trainees = [name for name, _ in noc_trainees if 'D' in str(trainee_matrix[name][d]).upper()]
        if len(working_trainees) > 2:
            for name in working_trainees[2:]:
                print(f"   ⚡ {d+1:2d} 日目: [{name}] を [在メ] に調整します")
                trainee_matrix[name][d] = '在メ'
                total_converted += 1

    for name, r in noc_trainees:
        for d in range(num_days): ws.cell(row=r, column=day_start_col + d).value = trainee_matrix[name][d]

    # 週末のD2制限解除
    for name, r in all_target_employees:
        for d in holiday_dates:
            cell = ws.cell(row=r, column=day_start_col + d)
            if str(cell.value or '').strip() in ['D2', 'D2(メ)']:
                cell.value = 'D'

    wb.save(output_file)
    print(f"\n🎉 最終ファイルのエクスポートが完了しました: [{output_file}]")


# ================= 4. 実行コントロールセンター =================
def main():
    while True:
        print("\n" + "=" * 50)
        print("🌟 スマートシフト作成システム - メインメニュー 🌟")
        print("=" * 50)
        print("  [1] クラウドから最新データを取得")
        print("  [2] PuLP最適化計算を実行 (シフト生成)")
        print("  [3] 役割とルールの後処理 (最終調整)")
        print("-" * 50)
        print("  [4] 全自動一括実行 (1~3を連続実行)")
        print("  [0] システムを終了")
        print("=" * 50)

        choice = input("👉 実行したいステップの番号を入力してください: ").strip()

        target_file = FILES.get('TARGET_LOCAL_FILE', 'Book1.xlsx')
        prev_file = FILES.get('PREV_LOCAL_FILE', 'BookLM.xlsx')
        pulp_out = FILES.get('PULP_OUTPUT_FILE', 'Book1_PuLP診断版出力結果.xlsx')
        final_out = FILES.get('FINAL_OUTPUT_FILE', 'Book1_最终完成版班表.xlsx')

        if choice == '1':
            try:
                fetch_and_split_sheets()
                print(f"💡 【次へ】: [{target_file}] を微調してから、ステップ [2] を実行してください。")
            except Exception as e: print(f"❌ エラー: {e}")

        elif choice == '2':
            try:
                solve_schedule_with_pulp(target_file, prev_file)
                print(f"💡 【次へ】: [{pulp_out}] を確認し、問題なければステップ [3] を実行してください。")
            except Exception as e: print(f"❌ エラー: {e}")

        elif choice == '3':
            try:
                post_process_schedule(pulp_out, final_out)
            except Exception as e: print(f"❌ エラー: {e}")

        elif choice == '4':
            print("\n🚀 全自動プロセスを開始します...")
            try:
                fetch_and_split_sheets()
                solve_schedule_with_pulp(target_file, prev_file)
                post_process_schedule(pulp_out, final_out)
            except Exception as e: print(f"\n❌ エラー: {e}")

        elif choice == '0':
            print("👋 プログラムを終了します。お疲れ様でした！")
            break
        else:
            print("⚠️ 無効な入力です。0〜4の番号を入力してください。")

if __name__ == '__main__':
    main()