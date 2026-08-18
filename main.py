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
    """実行環境のベースパスを取得（PyInstallerのビルド環境に対応）"""
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
    print(f"❌ エラー: 設定ファイル [{config_path}] が見つかりません。")
    sys.exit(1)

FILES = CONFIG.get('FILE_SETTINGS', {})
SPREADSHEET_ID = FILES.get('SPREADSHEET_ID', '')
DOWNLOAD_TASKS = [
    (FILES.get('TARGET_MONTH_SHEET', ''), FILES.get('TARGET_LOCAL_FILE', 'Book1.xlsx')),
    (FILES.get('PREV_MONTH_SHEET', ''), FILES.get('PREV_LOCAL_FILE', 'BookLM.xlsx'))
]

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
CREDENTIALS_FILE = os.path.join(get_base_path(), 'credentials.json')
TOKEN_FILE = os.path.join(get_base_path(), 'token_drive_exporter.json')

# ================= 1. Google Drive API 連携 =================
def get_drive_service():
    """Google Drive API 認証サービスの初期化"""
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(f'❌ [{CREDENTIALS_FILE}] が見つかりません。')

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    return build('drive', 'v3', credentials=creds)

def fetch_and_split_sheets():
    """クラウドからスプレッドシートを取得し特定シートに分割・保存"""
    service = get_drive_service()
    temp_excel_path = os.path.join(get_base_path(), '_temp_full_export.xlsx')
    mime_type_xlsx = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    request = service.files().export_media(fileId=SPREADSHEET_ID, mimeType=mime_type_xlsx)

    with open(temp_excel_path, 'wb') as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()

    for target_sheet, output_filename in DOWNLOAD_TASKS:
        if not target_sheet: continue
        out_path = os.path.join(get_base_path(), output_filename)
        wb = openpyxl.load_workbook(temp_excel_path)
        if target_sheet not in wb.sheetnames: continue
        for sheet_name in wb.sheetnames:
            if sheet_name != target_sheet: del wb[sheet_name]
        wb.save(out_path)

    if os.path.exists(temp_excel_path): os.remove(temp_excel_path)

# ================= 2. PuLP 数理最適化ソルバー =================
def solve_schedule_with_pulp(file_path, prev_month_file_path=None):
    """PuLP/HiGHSソルバーによるシフト最適化計算"""
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

    # ヘッダー位置解析
    name_col, team_col, day_start_col, header_row = None, None, None, None
    for r in range(1, 10):
        for c in range(1, 20):
            val = ws.cell(row=r, column=c).value
            if val == '名前': name_col = c; header_row = r
            elif val == 'チーム': team_col = c
            elif str(val) == '1' and header_row and r == header_row: day_start_col = c

    target_rows, employees, teams, prefilled, red_border_cells = [], [], [], {}, set()

    def check_red_border(cell):
        """赤枠（希望休）の判定"""
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

    # 前月データの引き継ぎ解析
    prev_month_aug31_night, prev_month_aug31_ake, prev_month_last5_work = set(), set(), {}

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

    def is_red_or_pink_color(hex_color):
        """祝日・赤日のセル色判定（紫色の除外）"""
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
        """土日および赤日判定"""
        c = day_start_col + d
        if str(ws.cell(row=header_row + 1, column=c).value or '').strip() in ['土', '日', '土曜', '日曜', 'Sat', 'Sun']: return True
        for r_check in range(1, header_row + 3):
            fill = ws.cell(row=r_check, column=c).fill
            if fill and fill.start_color and fill.start_color.rgb and is_red_or_pink_color(str(fill.start_color.rgb).upper()): return True
        return False

    num_emp, num_days = len(employees), 30
    holiday_dates = [d for d in range(num_days) if is_holiday_check(d)]
    target_public_rests = len(holiday_dates) if CONFIG.get('USE_DYNAMIC_PUBLIC_REST') else CONFIG.get('FIXED_PUBLIC_REST')

    shifts = [0, 1, 2, 3, 4]  # 0:休, 1:D, 2:N, 3:明, 4:年
    trainees = CONFIG.get('TRAINEES', [])
    primary_mentors = CONFIG.get('PRIMARY_MENTORS', [])
    night_partner_map = CONFIG.get('NIGHT_PARTNER_MAP', {})

    noc_seniors = [e for e in range(num_emp) if teams[e] == 'NOC' and employees[e] not in trainees]
    csc_seniors = [e for e in range(num_emp) if teams[e] == 'CSC' and employees[e] not in trainees]
    all_seniors = noc_seniors + csc_seniors
    trainee_indices = [e for e in range(num_emp) if employees[e] in trainees]
    mentor_indices = [e for e in range(num_emp) if employees[e] in primary_mentors]

    def build_model(diagnose_mode=False):
        """数理最適化モデルの構築"""
        prob = pulp.LpProblem('Shift_Scheduling', pulp.LpMaximize)
        x = pulp.LpVariable.dicts('x', (range(num_emp), range(num_days), shifts), cat=pulp.LpBinary)
        y_d2 = pulp.LpVariable.dicts('y_d2', (mentor_indices, range(num_days)), cat=pulp.LpBinary) if mentor_indices else {}
        penalties, diag_vars = [], {}

        # 1. 1日1シフト制約
        for e in range(num_emp):
            for d in range(num_days): prob += pulp.lpSum([x[e][d][s] for s in shifts]) == 1

        # 2. 事前入力シフトの固定
        shift_map = {'休': 0, 'D': 1, 'D2': 1, 'D3': 1, 'メ': 1, '保全': 1, 'CPN\n設置作業日': 1, 'CPN設置作業日': 1, 'N': 2, 'N設置作業': 2, '明': 3, '年': 4, 'PM年\nor早上がり': 4}
        for e in range(num_emp):
            for d in range(num_days):
                if (e, d) in prefilled and prefilled[(e, d)] in shift_map:
                    prob += x[e][d][shift_map[prefilled[(e, d)]]] == 1
                elif (e, d) not in prefilled:
                    prob += x[e][d][4] == 0

        # 3. 公休日数制約
        for e in range(num_emp):
            prob += pulp.lpSum([x[e][d][0] for d in range(num_days)]) == target_public_rests

        # 4. 夜勤(N) -> 明け(明) -> 休日 の連続性制約
        for e in range(num_emp):
            if e in prev_month_aug31_night: prob += x[e][0][3] == 1
            else:
                if (e, 0) not in prefilled or prefilled[(e, 0)] != '明': prob += x[e][0][3] == 0
            for d in range(num_days - 1): prob += x[e][d + 1][3] == x[e][d][2]

            if e in prev_month_aug31_ake and (e, 0) not in prefilled: prob += x[e][0][0] + x[e][0][4] == 1
            for d in range(num_days - 1):
                if not ((e, d + 1) in prefilled): prob += x[e][d + 1][0] + x[e][d + 1][4] >= x[e][d][3]

        # 5. 6連勤禁止制約
        for e in range(num_emp):
            last5_history = prev_month_last5_work.get(e, [0] * 5)
            for m in range(1, 6):
                prob += sum(last5_history[5 - (6 - m) : 5]) + pulp.lpSum([x[e][d_i][1] + x[e][d_i][2] + x[e][d_i][3] for d_i in range(m)]) <= 5
            for d in range(num_days - 5):
                prob += pulp.lpSum([x[e][d + i][1] + x[e][d + i][2] + x[e][d + i][3] for i in range(6)]) <= 5

        # 6. 必要人員数制約（毎日夜勤2名、明け2名）
        for d in range(num_days):
            prob += pulp.lpSum([x[e][d][2] for e in noc_seniors]) == 1
            prob += pulp.lpSum([x[e][d][2] for e in csc_seniors]) == 1
            if d > 0: prob += pulp.lpSum([x[e][d][3] for e in all_seniors]) == 2

        # 7. 月間夜勤・日勤回数の上下限
        for e in noc_seniors:
            prob += pulp.lpSum([x[e][d][2] for d in range(num_days)]) >= 3
            prob += pulp.lpSum([x[e][d][2] for d in range(num_days)]) <= 4
            prob += pulp.lpSum([x[e][d][1] for d in range(num_days)]) >= 4

        for e in csc_seniors:
            prob += pulp.lpSum([x[e][d][2] for d in range(num_days)]) >= 3
            prob += pulp.lpSum([x[e][d][2] for d in range(num_days)]) <= 5

        for e in trainee_indices:
            prob += pulp.lpSum([x[e][d][2] for d in range(num_days)]) >= 3
            prob += pulp.lpSum([x[e][d][2] for d in range(num_days)]) <= 5

        # 8a. D2 メンター割り当て制約
        special_task_words = {'CPN\n設置作業日', 'CPN設置作業日', '保全', 'メ'}
        for d in range(num_days):
            is_hol = d in holiday_dates
            if mentor_indices:
                for m in mentor_indices:
                    prob += y_d2[m][d] <= x[m][d][1]
                    if (m, d) in prefilled:
                        m_val = str(prefilled[(m, d)]).strip()
                        if m_val in special_task_words or (m_val != '' and m_val != 'D'): prob += y_d2[m][d] == 0

                d2_sum = pulp.lpSum([y_d2[m][d] for m in mentor_indices])
                if is_hol: prob += d2_sum == 0
                else: prob += d2_sum == 1

        # 8b & 8c. D2 担当回数の平準化と密集緩和ペナルティ
        if mentor_indices:
            d2_max = pulp.LpVariable('d2_max', lowBound=0, cat=pulp.LpInteger)
            d2_min = pulp.LpVariable('d2_min', lowBound=0, cat=pulp.LpInteger)
            for m in mentor_indices:
                total_d2_m = pulp.lpSum([y_d2[m][d] for d in range(num_days)])
                prob += total_d2_m <= d2_max
                prob += total_d2_m >= d2_min
            penalties.append(CONFIG.get('PENALTY_D2_RANGE', 20000) * (d2_max - d2_min))

            for m in mentor_indices:
                for d in range(num_days - 2):
                    s_d2_dense = pulp.LpVariable(f'slack_d2_dense_{m}_{d}', cat=pulp.LpBinary)
                    prob += y_d2[m][d] + y_d2[m][d + 1] + y_d2[m][d + 2] - s_d2_dense <= 1
                    penalties.append(CONFIG.get('PENALTY_D2_DENSE', 15000) * s_d2_dense)

        # 8d. 実習生夜勤時の専任メンター同行（マンツーマン）ペナルティ
        for tr_name, mentor_name in night_partner_map.items():
            if tr_name in employees and mentor_name in employees:
                tr_idx, m_idx = employees.index(tr_name), employees.index(mentor_name)
                for d in range(num_days):
                    s_np = pulp.LpVariable(f'slack_np_{tr_idx}_{d}', cat=pulp.LpBinary)
                    prob += x[tr_idx][d][2] - x[m_idx][d][2] <= s_np
                    penalties.append(CONFIG.get('PENALTY_NIGHT_PARTNER_MISSING', 50000) * s_np)

        # 8f. 夜勤間隔の最適化ペナルティ
        for e in range(num_emp):
            for d in range(num_days):
                for gap in range(1, 6):
                    if d + gap < num_days:
                        s_night_gap = pulp.LpVariable(f'slack_night_gap_{e}_{d}_{gap}', cat=pulp.LpBinary)
                        prob += x[e][d][2] + x[e][d + gap][2] - 1 <= s_night_gap
                        p_weight = CONFIG.get('PENALTY_NIGHT_GAP_TIGHT', 50000) if gap <= 2 else CONFIG.get('PENALTY_NIGHT_GAP_MODERATE', 30000)
                        penalties.append(p_weight * s_night_gap)

        # 8g. チーム別日勤人数の調整
        noc_mentors = [m for m in mentor_indices if m in noc_seniors]
        for d in range(num_days):
            is_hol = d in holiday_dates
            noc_sp_count = sum(1 for e in noc_seniors if (e, d) in prefilled and str(prefilled[(e, d)]).strip() in special_task_words)
            noc_pure_d_sum = pulp.lpSum([x[e][d][1] for e in noc_seniors]) - (pulp.lpSum([y_d2[m][d] for m in noc_mentors]) if noc_mentors else 0) - noc_sp_count

            if is_hol: prob += noc_pure_d_sum == 1
            else: prob += noc_pure_d_sum >= 2

            s_noc_over4 = pulp.LpVariable(f'slack_noc_over4_{d}', lowBound=0, cat=pulp.LpInteger)
            prob += noc_pure_d_sum - s_noc_over4 <= 3
            penalties.append(CONFIG.get('PENALTY_NOC_PURE_D_OVER3', 20000) * s_noc_over4)

            if is_hol: prob += pulp.lpSum([x[e][d][1] for e in csc_seniors]) >= 1
            else:
                prob += pulp.lpSum([x[e][d][1] for e in csc_seniors]) >= 1
                s_csc_d2 = pulp.LpVariable(f'slack_csc_d2_{d}', lowBound=0, cat=pulp.LpInteger)
                prob += pulp.lpSum([x[e][d][1] for e in csc_seniors]) + s_csc_d2 >= 2
                penalties.append(CONFIG.get('PENALTY_CSC_WEEKDAY_D_SHORT', 50000) * s_csc_d2)

            if is_hol: prob += pulp.lpSum([x[e][d][2] for e in trainee_indices]) == 0
            else: prob += pulp.lpSum([x[e][d][2] for e in trainee_indices]) <= 1

        # 直近7日間の夜勤上限緩和ペナルティ
        for e in range(num_emp):
            for d in range(num_days):
                start_7 = max(0, d - 6)
                s_7_night = pulp.LpVariable(f'slack_7_night_{e}_{d}', lowBound=0, cat=pulp.LpInteger)
                prob += pulp.lpSum([x[e][d_i][2] for d_i in range(start_7, d + 1)]) - s_7_night <= 2
                penalties.append(CONFIG.get('PENALTY_ROLLING_7DAY_NIGHT_OVER2', 1000) * s_7_night)

        # 実習生の日勤（D3）人数制御
        for d in range(num_days):
            s_d3_over = pulp.LpVariable(f'slack_d3_over_{d}', lowBound=0, cat=pulp.LpInteger)
            prob += pulp.lpSum([x[e][d][1] for e in trainee_indices]) - s_d3_over <= 2
            penalties.append(CONFIG.get('PENALTY_TRAINEE_D3_OVER2', 18000) * s_d3_over)

        for d in holiday_dates:
            for tr in trainee_indices:
                penalties.append(CONFIG.get('PENALTY_TRAINEE_HOLIDAY_D3', 30000) * x[tr][d][1])

        # 9. 目的関数（連休報酬および休日休みの優先評価）
        objective_terms = []
        def is_rest(e_idx, day_idx):
            return (x[e_idx][day_idx][0] + x[e_idx][day_idx][4]) if 0 <= day_idx < num_days else None

        for e in range(num_emp):
            # 2連休評価
            for d in range(num_days - 1):
                b2 = pulp.LpVariable(f'block_2_{e}_{d}', cat=pulp.LpBinary)
                conds = [is_rest(e, d), is_rest(e, d + 1)]
                r_p, r_n = is_rest(e, d - 1), is_rest(e, d + 2)
                if r_p is not None: conds.append(1 - r_p)
                if r_n is not None: conds.append(1 - r_n)
                for c in conds: prob += b2 <= c
                prob += b2 >= pulp.lpSum(conds) - (len(conds) - 1)
                objective_terms.append(CONFIG.get('REWARD_REST_BLOCK_2', 30) * b2)

            # 3連休評価
            for d in range(num_days - 2):
                b3 = pulp.LpVariable(f'block_3_{e}_{d}', cat=pulp.LpBinary)
                conds = [is_rest(e, d), is_rest(e, d + 1), is_rest(e, d + 2)]
                r_p, r_n = is_rest(e, d - 1), is_rest(e, d + 3)
                if r_p is not None: conds.append(1 - r_p)
                if r_n is not None: conds.append(1 - r_n)
                for c in conds: prob += b3 <= c
                prob += b3 >= pulp.lpSum(conds) - (len(conds) - 1)
                objective_terms.append(CONFIG.get('REWARD_REST_BLOCK_3', 40) * b3)

            # 4連休以上評価
            for d in range(num_days - 3):
                b4 = pulp.LpVariable(f'block_4plus_{e}_{d}', cat=pulp.LpBinary)
                conds = [is_rest(e, d), is_rest(e, d + 1), is_rest(e, d + 2), is_rest(e, d + 3)]
                for c in conds: prob += b4 <= c
                prob += b4 >= pulp.lpSum(conds) - (len(conds) - 1)
                objective_terms.append(CONFIG.get('REWARD_REST_BLOCK_4PLUS', 45) * b4)

            # 6連休以上ペナルティ
            for d in range(num_days - 5):
                b6 = pulp.LpVariable(f'block_6plus_{e}_{d}', cat=pulp.LpBinary)
                conds6 = [is_rest(e, d + i) for i in range(6)]
                for c in conds6: prob += b6 <= c
                prob += b6 >= pulp.lpSum(conds6) - (len(conds6) - 1)
                penalties.append(CONFIG.get('PENALTY_REST_BLOCK_6PLUS', 1000) * b6)

        # 祝日休暇の割り当て優遇
        for e in range(num_emp):
            for d in holiday_dates:
                objective_terms.append(CONFIG.get('REWARD_HOLIDAY_REST', 60) * (x[e][d][0] + x[e][d][4]))

        prob += pulp.lpSum(objective_terms) - pulp.lpSum(penalties)
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
        print('❌ 解が見つかりません：条件が厳しすぎます。')
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
            if m in y_d2 and pulp.value(y_d2[m][d]) is not None and pulp.value(y_d2[m][d]) > 0.5:
                result_schedule[m][d] = 'D2'

    wb_out = openpyxl.load_workbook(file_path)
    ws_out = wb_out[sheet_name]

    # スタイル定義
    red_font = Font(color='FF0000')
    black_font = Font(color='000000')
    purple_fill = PatternFill(start_color='E6E6FA', end_color='E6E6FA', fill_type='solid')
    green_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')

    for e in range(num_emp):
        for d in range(num_days):
            if (e, d) not in prefilled or (e, d) in red_border_cells:
                val = result_schedule[e][d]
                cell = ws_out.cell(row=target_rows[e], column=day_start_col + d)
                cell.value = val

                if val == '休': cell.font = red_font
                elif val in ['N', '明']: cell.fill = purple_fill; cell.font = black_font
                elif val in ['D2', 'D2(メ)']: cell.fill = green_fill; cell.font = black_font
                else: cell.font = black_font

    wb_out.save(out_file)
    print(f'📁 中間シフト表の出力が完了しました: {out_file}\n')

# ================= 3. 役割後処理エンジン（装飾・公平性評価対応） =================
def post_process_schedule(input_file, output_file):
    """中間シフト表に対する具体的な役割割り当て（メ/保全/タスク/在メ）の後処理"""
    input_file = os.path.join(get_base_path(), input_file)
    output_file = os.path.join(get_base_path(), output_file)

    print("🚀 起動【排班表角色後処理エンジン (詳細解析版)】...\n")
    if not os.path.exists(input_file):
        print(f"❌ 找不到入力ファイル: [{input_file}]")
        return

    wb = openpyxl.load_workbook(input_file)
    ws = wb[wb.sheetnames[0]]

    name_col, team_col, day_start_col, header_row = None, None, None, None
    for r in range(1, 10):
        for c in range(1, 20):
            val = ws.cell(row=r, column=c).value
            if val == '名前': name_col = c; header_row = r
            elif val == 'チーム': team_col = c
            elif str(val) == '1' and header_row and r == header_row: day_start_col = c

    def is_holiday_check(d):
        c = day_start_col + d
        sample_cell = ws.cell(row=header_row + 1, column=c)
        if sample_cell.fill and sample_cell.fill.start_color and sample_cell.fill.start_color.rgb:
            if 'F4CC' in str(sample_cell.fill.start_color.rgb).upper(): return True
        return (d % 7) in [4, 5]

    num_days = 30
    holiday_dates = set(d for d in range(num_days) if is_holiday_check(d))
    trainees_set = set(CONFIG.get('TRAINEES', []))

    noc_seniors, noc_trainees, csc_employees, all_target_employees = [], [], [], []
    seen_employees = set()

    for r in range(header_row + 2, ws.max_row + 1):
        team = ws.cell(row=r, column=team_col).value
        raw_name = ws.cell(row=r, column=name_col).value
        name = str(raw_name or '').strip().replace(' ', '').replace('\u3000', '').replace('\n', '')

        if not name: continue
        if name in seen_employees: break

        if team in ['NOC', 'CSC']:
            seen_employees.add(name)
            all_target_employees.append((name, r))
            if team == 'NOC':
                if any(t in name for t in trainees_set) or name in trainees_set:
                    noc_trainees.append((name, r))
                else:
                    noc_seniors.append((name, r))
            elif team == 'CSC':
                csc_employees.append((name, r))

    print(f"📊 人員データ読み込み完了:")
    print(f"   ├─ NOC正社員 ({len(noc_seniors)}名): {[n for n, _ in noc_seniors]}")
    print(f"   ├─ NOC実習生 ({len(noc_trainees)}名): {[n for n, _ in noc_trainees]}")
    print(f"   └─ CSC社員   ({len(csc_employees)}名): {[n for n, _ in csc_employees]}\n")

    def get_consecutive_d_days(matrix, emp_name, current_day):
        """連続出勤日数のカウント"""
        consecutive, day = 0, current_day
        while day >= 0:
            val = str(matrix[emp_name][day]).upper()
            if any(k in val for k in ['D', 'D2', 'D3']):
                consecutive += 1; day -= 1
            else: break
        return consecutive

    # 1. NOC 正式社員の役割割り当て（メ/保全/タスク）
    noc_counts = {name: {'メ': 0, '保全': 0, 'メ/保全': 0, 'total': 0} for name, _ in noc_seniors}
    noc_last_day = {name: {'メ': -99, '保全': -99, 'メ/保全': -99, 'total': -99} for name, _ in noc_seniors}
    noc_matrix = {name: [str(ws.cell(row=r, column=day_start_col + d).value or '').strip() for d in range(num_days)] for name, r in noc_seniors}

    for d in range(num_days):
        d2_working_names = [name for name, _ in noc_seniors if 'D2' in noc_matrix[name][d]]
        pure_d_working_names = [name for name, _ in noc_seniors if noc_matrix[name][d] == 'D']

        for name in d2_working_names: noc_matrix[name][d] = 'D2'
        d_count = len(pure_d_working_names)

        def calc_noc_priority(emp_name):
            cons_d = get_consecutive_d_days(noc_matrix, emp_name, d)
            total_used = noc_counts[emp_name]['total']
            bonus = 3.0 if cons_d >= 3 else (1.0 if cons_d == 2 else 0.0)
            return -(bonus - (total_used * 1.2))

        if d_count == 1:
            noc_matrix[pure_d_working_names[0]][d] = 'D'
        elif d_count == 2:
            candidates = [(name, r) for name, r in noc_seniors if name in pure_d_working_names]
            candidates.sort(key=lambda x: (calc_noc_priority(x[0]), noc_counts[x[0]]['total'], noc_counts[x[0]]['メ/保全'], -(d - noc_last_day[x[0]]['total'])))
            combo_name, pure_d_name = candidates.pop(0)[0], candidates.pop(0)[0]
            noc_matrix[combo_name][d] = 'メ/保全'
            noc_matrix[pure_d_name][d] = 'D'
            noc_counts[combo_name]['メ/保全'] += 1; noc_counts[combo_name]['total'] += 1
            noc_last_day[combo_name]['メ/保全'] = d; noc_last_day[combo_name]['total'] = d
        elif d_count >= 3:
            candidates = [(name, r) for name, r in noc_seniors if name in pure_d_working_names]
            candidates.sort(key=lambda x: (calc_noc_priority(x[0]), noc_counts[x[0]]['メ'], -(d - noc_last_day[x[0]]['total'])))
            me_name = candidates.pop(0)[0]
            noc_matrix[me_name][d] = 'メ'
            noc_counts[me_name]['メ'] += 1; noc_counts[me_name]['total'] += 1
            noc_last_day[me_name]['メ'] = d; noc_last_day[me_name]['total'] = d

            candidates.sort(key=lambda x: (calc_noc_priority(x[0]), noc_counts[x[0]]['保全'], -(d - noc_last_day[x[0]]['total'])))
            ho_name = candidates.pop(0)[0]
            noc_matrix[ho_name][d] = '保全'
            noc_counts[ho_name]['保全'] += 1; noc_counts[ho_name]['total'] += 1
            noc_last_day[ho_name]['保全'] = d; noc_last_day[ho_name]['total'] = d

            candidates.sort(key=lambda x: (get_consecutive_d_days(noc_matrix, x[0], d), -calc_noc_priority(x[0])))
            primary_d_name = candidates.pop(0)[0]
            noc_matrix[primary_d_name][d] = 'D'

            for name, _ in candidates: noc_matrix[name][d] = 'タスク'

    for name, r in noc_seniors:
        for d in range(num_days): ws.cell(row=r, column=day_start_col + d).value = noc_matrix[name][d]

    # 2. NOC 実習生（トレーニー）の「在メ」自動変換
    trainee_matrix = {name: [str(ws.cell(row=r, column=day_start_col + d).value or '').strip() for d in range(num_days)] for name, r in noc_trainees}
    trainee_za_me_counts = {name: 0 for name, _ in noc_trainees}
    total_converted_count = 0

    for d in range(num_days):
        working_trainees = [name for name, _ in noc_trainees if any(k in str(trainee_matrix[name][d]).upper() for k in ['D3', 'D'])]
        if len(working_trainees) > 2:
            overflow_count = len(working_trainees) - 2
            working_trainees.sort(key=lambda x: (trainee_za_me_counts[x], -get_consecutive_d_days(trainee_matrix, x, d)))
            for name in working_trainees[:overflow_count]:
                print(f"   ⚡ 第 {d+1:2d} 日: 実習生出勤 {len(working_trainees)} 名 -> [{name}] を [在メ] に変換")
                trainee_matrix[name][d] = '在メ'
                trainee_za_me_counts[name] += 1
                total_converted_count += 1

    for name, r in noc_trainees:
        for d in range(num_days): ws.cell(row=r, column=day_start_col + d).value = trainee_matrix[name][d]

    # 3. CSC 社員の役割割り当て（メ/勤）
    csc_me_counts = {name: 0 for name, _ in csc_employees}
    csc_last_me_day = {name: -99 for name, _ in csc_employees}
    csc_matrix = {name: [str(ws.cell(row=r, column=day_start_col + d).value or '').strip() for d in range(num_days)] for name, r in csc_employees}

    for d in range(num_days):
        if d in holiday_dates: continue
        assigned_me = [name for name, _ in csc_employees if csc_matrix[name][d] == 'メ']
        d_candidates = [(name, r) for name, r in csc_employees if csc_matrix[name][d] == 'D']

        if not assigned_me and len(d_candidates) >= 2:
            def calc_csc_priority(emp_name):
                cons_d = get_consecutive_d_days(csc_matrix, emp_name, d)
                return -((2.5 if cons_d >= 3 else (1.0 if cons_d == 2 else 0.0)) - (csc_me_counts[emp_name] * 1.0))

            d_candidates.sort(key=lambda x: (calc_csc_priority(x[0]), -(d - csc_last_me_day[x[0]])))
            me_name = d_candidates[0][0]
            csc_matrix[me_name][d] = 'メ'
            csc_me_counts[me_name] += 1
            csc_last_me_day[me_name] = d
            d_candidates = [c for c in d_candidates if c[0] != me_name]

        if len(d_candidates) >= 2:
            d_candidates.sort(key=lambda x: -get_consecutive_d_days(csc_matrix, x[0], d))
            for name, r in d_candidates[:-1]: csc_matrix[name][d] = '勤'

    for name, r in csc_employees:
        for d in range(num_days): ws.cell(row=r, column=day_start_col + d).value = csc_matrix[name][d]

    # 4. 休日制御と最終スタイル（フォント・背景色）の反映
    red_font = Font(color='FF0000')
    black_font = Font(color='000000')
    purple_fill = PatternFill(start_color='E6E6FA', end_color='E6E6FA', fill_type='solid')
    green_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')

    for name, r in all_target_employees:
        for d in range(num_days):
            c = day_start_col + d
            cell = ws.cell(row=r, column=c)
            val = str(cell.value or '').strip()

            # 休日・祝日の D2 クリア
            if d in holiday_dates and val in ['D2', 'D2(メ)']:
                val = 'D'
                cell.value = 'D'

            # 役割に応じたセルの着色
            if val == '休':
                cell.font = red_font
            elif val in ['N', '明']:
                cell.fill = purple_fill
                cell.font = black_font
            elif val in ['D2', 'D2(メ)']:
                cell.fill = green_fill
                cell.font = black_font
            else:
                cell.font = black_font

    # レポート出力
    print(f"\n📌 変換完了統計:")
    print(f"   └─ 全月で【在メ】自動変換を {total_converted_count} 回実行しました。")

    if noc_trainees:
        print("\n📈 【NOC 実習生 在メ 割り当て公平性レポート】:")
        for name, count in trainee_za_me_counts.items():
            print(f"   - {name:8s}: 在メ = {count} 回")

    wb.save(output_file)
    print("\n" + "=" * 65)
    print(f"🎉 最終ファイルの保存に成功しました: [{output_file}]")
    print("=" * 65)

# ================= 4. メインメニュー（実行コントロール） =================
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
            try: fetch_and_split_sheets()
            except Exception as e: print(f"❌ エラー: {e}")

        elif choice == '2':
            try: solve_schedule_with_pulp(target_file, prev_file)
            except Exception as e: print(f"❌ エラー: {e}")

        elif choice == '3':
            try: post_process_schedule(pulp_out, final_out)
            except Exception as e: print(f"❌ エラー: {e}")

        elif choice == '4':
            try:
                fetch_and_split_sheets()
                solve_schedule_with_pulp(target_file, prev_file)
                post_process_schedule(pulp_out, final_out)
            except Exception as e: print(f"❌ エラー: {e}")

        elif choice == '0':
            print("👋 プログラムを終了します。")
            break

if __name__ == '__main__':
    main()