import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill
from ortools.sat.python import cp_model


def solve_schedule_with_cpsat(file_path, prev_month_file_path=None):
    print(
        "🚀【Google OR-Tools CP-SAT 数理最適化ソルバー"
        " (純色祝日判定/夜勤サポート対応/休日単独不可1名制限・平日出勤優先版)】を起動中...\n"
    )

    # =========================================================================
    # 🎯 0. 集中的設定センター：メンバーリスト、OJT師弟関係および重み付け設定
    # =========================================================================
    CONFIG = {
        # ------------------- 公休数（休日数）設定 -------------------
        'USE_DYNAMIC_PUBLIC_REST': True,  # True: 当月の赤日/土日を自動計算して公休数とする, False: FIXED_PUBLIC_RESTを使用
        'FIXED_PUBLIC_REST': 11,          # USE_DYNAMIC_PUBLIC_REST = False の場合の固定公休日数

        # ------------------- メンバーおよびOJT関係設定 -------------------
        'TRAINEES': [],

        'PRIMARY_MENTORS': [
            '森川亘浩', '今村孝', '高田康平', '檜垣優介', '森下優', '小林朔也',
            '駱中威', '佐藤俊一郎', '大原和士', '五十嵐康之', 'ﾏﾝｻﾞﾆﾘｱﾘｶﾙﾄﾞ'
        ],

        # ✨ 夜勤単独不可（サポート・同伴必須）のメンバー
        'NEEDS_NIGHT_SUPPORT': ['駱中威', '佐藤俊一郎', '大原和士'],

        # ✨ 夜勤単独不可メンバーのパートナー（同伴者）プール
        'NIGHT_SUPPORT_POOL': {
            '駱中威': ['森下優', '小林朔也'],
            '佐藤俊一郎': ['森下優', '小林朔也'],
            '大原和士': ['森川亘浩', '今村孝', '高田康平', '檜垣優介', '森下優', '小林朔也', '五十嵐康之', 'ﾏﾝｻﾞﾆﾘｱﾘｶﾙﾄﾞ']
        },

        # =========================================================
        # ✨ 夜勤・日勤の回数制限コントロールパネル
        # =========================================================
        'LIMIT_NOC_NIGHT_MAX': 4,       # NOC独立先輩の夜勤上限
        'LIMIT_CSC_NIGHT_MAX': 6,       # CSC先輩の夜勤上限
        'SUPPORT_NIGHT_TARGET': 3,      # 単独夜勤不可メンバーの夜勤回数

        # ------------------- 報酬 (Objectives / Rewards) -------------------
        'REWARD_HOLIDAY_REST': 60,
        'REWARD_SUPPORT_HOLIDAY_REST': 20000, # ✨ 単独不可メンバーの休日休み報酬（高額）
        'REWARD_REST_BLOCK_2': 30,
        'REWARD_REST_BLOCK_3': 40,
        'REWARD_REST_BLOCK_4PLUS': 45,

        # ------------------- 軟的制約ペナルティ (Soft Penalties) -------------------
        'PENALTY_REST_5_CONSECUTIVE': 100,
        'PENALTY_ISOLATED_REST': 10000,   # ✨ 10000 に引き上げ（単発休み・孤立休を強力防止）
        'PENALTY_ISOLATED_WORK': 10000,   # ✨ 10000 に引き上げ（1日のみ出勤・孤立出勤を強力防止）

        'PENALTY_NIGHT_VARIANCE': 50000,          # 独立先輩チーム内の夜勤格差最小化
        'PENALTY_NIGHT_PARTNER_MISSING': 50000,
        'PENALTY_NIGHT_GAP_TIGHT': 50000,
        'PENALTY_CSC_WEEKDAY_D_SHORT': 50000,
        'PENALTY_NIGHT_GAP_MODERATE': 30000,
        'PENALTY_TRAINEE_HOLIDAY_D3': 30000,
        'PENALTY_D2_RANGE': 20000,
        'PENALTY_NOC_PURE_D_OVER3': 20000,
        'PENALTY_TRAINEE_D3_OVER2': 18000,
        'PENALTY_D2_DENSE': 15000,
        'PENALTY_ROLLING_7DAY_NIGHT_OVER2': 1000,
    }

    if not os.path.exists(file_path):
        print(f"❌ 指定されたファイルが見つかりません: [{file_path}]")
        return

    # ================= 1. Excel データの読み込みと解析 =================
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    name_col, team_col, day_start_col, header_row = None, None, None, None
    for r in range(1, 10):
        for c in range(1, 20):
            val = ws.cell(row=r, column=c).value
            if val == '名前':
                name_col = c
                header_row = r
            elif val == 'チーム':
                team_col = c
            elif str(val) == '1' and header_row and r == header_row:
                day_start_col = c

    num_days = 0
    if day_start_col and header_row:
        while True:
            v = ws.cell(row=header_row, column=day_start_col + num_days).value
            if isinstance(v, (int, float)) or (isinstance(v, str) and str(v).isdigit()):
                num_days += 1
            else:
                break
    if num_days == 0:
        num_days = 31

    print(f"📅 当月の日数を自動判定しました: {num_days} 日間")

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
            for d in range(num_days):
                c = day_start_col + d
                cell = ws.cell(row=r, column=c)
                val_str = str(cell.value).strip() if cell.value is not None else ''
                has_red = check_red_border(cell)
                if has_red:
                    red_border_cells.add((emp_idx, d))
                    prefilled[(emp_idx, d)] = val_str if val_str else '休'
                elif val_str != '':
                    prefilled[(emp_idx, d)] = val_str

    # ================= 1b. 前月 (BookLM.xlsx) 引き継ぎデータの読み込み =================
    prev_month_aug31_night, prev_month_aug31_ake, prev_month_last5_work = set(), set(), {}

    def is_work_shift_name(val):
        val = str(val or '').strip()
        return False if val in ['', '休', '年'] else True

    if prev_month_file_path and os.path.exists(prev_month_file_path):
        wb_lm = openpyxl.load_workbook(prev_month_file_path, data_only=True)
        ws_lm = wb_lm[wb_lm.sheetnames[0]]

        name_col_l, day_start_col_l, header_row_l = None, None, None
        for r in range(1, 10):
            for c in range(1, 20):
                val = ws_lm.cell(row=r, column=c).value
                if val == '名前':
                    name_col_l = c
                    header_row_l = r
                elif str(val) == '1' and header_row_l and r == header_row_l:
                    day_start_col_l = c

        prev_num_days = 0
        if day_start_col_l and header_row_l:
            while True:
                v = ws_lm.cell(row=header_row_l, column=day_start_col_l + prev_num_days).value
                if isinstance(v, (int, float)) or (isinstance(v, str) and str(v).isdigit()):
                    prev_num_days += 1
                else:
                    break

        lm_shifts_map = {}
        for r in range(header_row_l + 2, ws_lm.max_row + 1):
            raw_name = ws_lm.cell(row=r, column=name_col_l).value
            name = str(raw_name or '').strip().replace(' ', '').replace('\u3000', '')
            if name and name not in lm_shifts_map:
                shifts_aug = []
                for d_aug in range(prev_num_days):
                    val = ws_lm.cell(row=r, column=day_start_col_l + d_aug).value
                    shifts_aug.append(str(val).strip() if val is not None else '')
                lm_shifts_map[name] = shifts_aug

        for e, name in enumerate(employees):
            aug_shifts = lm_shifts_map.get(name, [''] * prev_num_days)
            aug_last_val = aug_shifts[-1] if len(aug_shifts) > 0 else ''

            if aug_last_val in ['N', 'N設置作業']:
                prev_month_aug31_night.add(e)
            elif aug_last_val == '明':
                prev_month_aug31_ake.add(e)

            last5 = aug_shifts[-5:] if len(aug_shifts) >= 5 else [''] * 5
            prev_month_last5_work[e] = [1 if is_work_shift_name(s) else 0 for s in last5]

        print(f'📊 前月データの読み込み完了: 前月最終日の夜勤(N) {len(prev_month_aug31_night)} 名、明け(明) {len(prev_month_aug31_ake)} 名。')

    # ================= 100% 純色判定エンジン (赤日・ピンク色) =================
    def is_red_or_pink_color(cell):
        fill = cell.fill
        if not fill or not fill.fill_type or fill.fill_type == 'none':
            return False
        
        color_str = None
        if fill.start_color and fill.start_color.rgb:
            color_str = str(fill.start_color.rgb).upper()
        elif fill.fgColor and fill.fgColor.rgb:
            color_str = str(fill.fgColor.rgb).upper()
            
        if color_str:
            hex_rgb = color_str[-6:]
            if hex_rgb in ['FFFFFF', '000000', '00FFFFFF']:
                return False
            try:
                r = int(hex_rgb[0:2], 16)
                g = int(hex_rgb[2:4], 16)
                b = int(hex_rgb[4:6], 16)
                if r > 180 and (r - g > 15) and (r - b > 15):
                    return True
                red_keywords = ['F4CC', 'C0CB', 'D9D9', 'ECEC', 'FFC0', 'FFD', 'FFA', 'FFC']
                if any(kw in hex_rgb for kw in red_keywords):
                    return True
            except ValueError:
                pass

        if fill.start_color and fill.start_color.theme is not None:
            if fill.start_color.theme in [5, 6, 7, 8, 9]:
                return True

        return False

    def is_holiday_check(d):
        c = day_start_col + d
        for r_check in range(1, header_row + 3):
            cell = ws.cell(row=r_check, column=c)
            if is_red_or_pink_color(cell):
                return True
        return False

    num_emp = len(employees)
    holiday_dates = [d for d in range(num_days) if is_holiday_check(d)]

    if CONFIG['USE_DYNAMIC_PUBLIC_REST']:
        target_public_rests = len(holiday_dates)
    else:
        target_public_rests = CONFIG['FIXED_PUBLIC_REST']

    print(f'📅 判定された【赤日・祝日・休日】日数: {len(holiday_dates)} 日')
    print(f'📊 当月の公休割当配額: {target_public_rests} 日\n')

    # チーム編成と各種インデックスの取得
    trainees = CONFIG['TRAINEES']
    primary_mentors = CONFIG['PRIMARY_MENTORS']

    noc_seniors = [e for e in range(num_emp) if teams[e] == 'NOC' and employees[e] not in trainees]
    csc_seniors = [e for e in range(num_emp) if teams[e] == 'CSC' and employees[e] not in trainees]
    all_seniors = noc_seniors + csc_seniors
    trainee_indices = [e for e in range(num_emp) if employees[e] in trainees]
    mentor_indices = [e for e in range(num_emp) if employees[e] in primary_mentors]

    # =========================================================================
    # ⚡️ 【免ソルバー】超高速 Python ロジックプレチェック
    # =========================================================================
    print("🔍 超高速 Python 事前デッドロック検知（プレチェック）を実行中...")
    fatal_errors = []

    for e in range(num_emp):
        rest_count = sum(1 for d in range(num_days) if (e, d) in prefilled and prefilled[(e, d)] in ['休', '年'])
        if rest_count > target_public_rests:
            fatal_errors.append(f"❌ 【公休超過エラー】{employees[e]}: 事前入力された休みが {rest_count} 日あり、上限({target_public_rests}日)を超過しています！")

    for d in range(num_days):
        is_hol = d in holiday_dates
        noc_avail = sum(1 for e in noc_seniors if not ((e, d) in prefilled and prefilled[(e, d)] in ['休', '年']))
        csc_avail = sum(1 for e in csc_seniors if not ((e, d) in prefilled and prefilled[(e, d)] in ['休', '年']))
        
        noc_req = 2 if is_hol else 3 
        csc_req = 1 if is_hol else 2 
        
        if noc_avail < noc_req:
            fatal_errors.append(f"❌ 【出勤人数不足エラー】第 {d+1} 日: NOCチームの出勤可能人数が {noc_avail} 名しかいません（最低 {noc_req} 名必要）。")
        if csc_avail < csc_req:
            fatal_errors.append(f"❌ 【出勤人数不足エラー】第 {d+1} 日: CSCチームの出勤可能人数が {csc_avail} 名しかいません（最低 {csc_req} 名必要）。")

    for e in range(num_emp):
        if e in prev_month_aug31_night and (e, 0) in prefilled and prefilled[(e, 0)] != '明':
            fatal_errors.append(f"❌ 【シフト遷移エラー】{employees[e]}: 前月最終日が夜勤(N)ですが、当月1日が「明」以外で固定されています！")
        if e in prev_month_aug31_ake and (e, 0) in prefilled and prefilled[(e, 0)] not in ['休', '年']:
            fatal_errors.append(f"❌ 【シフト遷移エラー】{employees[e]}: 前月最終日が明け(明)ですが、当月1日が「休/年」以外で固定されています！")
        
        last5_hist = prev_month_last5_work.get(e, [0]*5)
        for m in range(1, 6):
            if sum(last5_hist[5-(6-m):5]) + sum(1 for di in range(m) if (e, di) in prefilled and is_work_shift_name(prefilled[(e, di)])) > 5:
                fatal_errors.append(f"❌ 【6連勤エラー】{employees[e]}: 前月からの連続出勤により、月初に6連勤が発生する事前入力があります！")

        for d in range(num_days - 1):
            if (e, d) in prefilled and prefilled[(e, d)] in ['N', 'N設置作業']:
                if (e, d+1) in prefilled and prefilled[(e, d+1)] != '明':
                    fatal_errors.append(f"❌ 【シフト遷移エラー】{employees[e]}: 第 {d+1} 日が夜勤(N)ですが、翌日が「明」以外で固定されています！")
            if (e, d) in prefilled and prefilled[(e, d)] == '明':
                if (e, d+1) in prefilled and prefilled[(e, d+1)] not in ['休', '年']:
                    fatal_errors.append(f"❌ 【シフト遷移エラー】{employees[e]}: 第 {d+1} 日が明け(明)ですが、翌日が休み以外で固定されています！")

    if fatal_errors:
        print("\n" + "="*65)
        print("🚨 プレチェック未通過！以下の絶対不可避な矛盾が検出されました:")
        for err in set(fatal_errors):
            print(err)
        print("="*65 + "\n")
        return

    print("✅ プレチェック通過！深刻なデッドロックは見つかりませんでした。CP-SAT モデル構築を開始します...\n")

    # =========================================================================
    # 2. Google OR-Tools CP-SAT モデルの構築
    # =========================================================================
    model = cp_model.CpModel()
    
    # シフト定義: 0:休, 1:D, 2:N, 3:明, 4:年
    shifts = [0, 1, 2, 3, 4]
    
    # 変数定義 (x[e, d, s] BoolVar)
    x = {}
    for e in range(num_emp):
        for d in range(num_days):
            for s in shifts:
                x[e, d, s] = model.NewBoolVar(f'x_{e}_{d}_{s}')

    # D2（メンター指導）フラグ変数
    y_d2 = {}
    for m in mentor_indices:
        for d in range(num_days):
            y_d2[m, d] = model.NewBoolVar(f'y_d2_{m}_{d}')

    # 目的関数の項（得点とペナルティ）
    objective_terms = []

    shift_map = {
        '休': 0, 'D': 1, 'D2': 1, 'D3': 1, 'メ': 1, '保全': 1,
        'CPN\n設置作業日': 1, 'CPN設置作業日': 1, 'N': 2, 'N設置作業': 2,
        '明': 3, '年': 4, 'PM年\nor早上がり': 4,
    }

    # 1. 1日1シフト制約 ＆ 事前入力固定
    for e in range(num_emp):
        for d in range(num_days):
            model.AddExactlyOne([x[e, d, s] for s in shifts])
            if (e, d) in prefilled:
                val = prefilled[(e, d)]
                if val in shift_map:
                    model.Add(x[e, d, shift_map[val]] == 1)
                else:
                    if (e, d) in red_border_cells:
                        model.Add(x[e, d, 0] == 1)
                    else:
                        model.Add(x[e, d, 1] == 1)
            else:
                model.Add(x[e, d, 4] == 0)

    # 2. 公休日数割当 (絶対的ハード制約)
    for e in range(num_emp):
        model.Add(sum(x[e, d, 0] for d in range(num_days)) == target_public_rests)

    # 3. Shift Transition: N -> 明 -> 休/年
    for e in range(num_emp):
        if e in prev_month_aug31_night:
            model.Add(x[e, 0, 3] == 1)
        elif (e, 0) not in prefilled or prefilled[(e, 0)] != '明':
            model.Add(x[e, 0, 3] == 0)

        if e in prev_month_aug31_ake and (e, 0) not in prefilled:
            model.Add(x[e, 0, 0] + x[e, 0, 4] == 1)

        for d in range(num_days - 1):
            model.Add(x[e, d + 1, 3] == x[e, d, 2])
            if (e, d + 1) not in prefilled:
                model.Add(x[e, d + 1, 0] + x[e, d + 1, 4] >= x[e, d, 3])

    # 4. 6連勤厳格禁止
    for e in range(num_emp):
        last5_history = prev_month_last5_work.get(e, [0] * 5)
        for m in range(1, 6):
            aug_work_sum = sum(last5_history[5 - (6 - m) : 5])
            model.Add(aug_work_sum + sum(x[e, d_i, 1] + x[e, d_i, 2] + x[e, d_i, 3] for d_i in range(m)) <= 5)

        for d in range(num_days - 5):
            model.Add(sum(x[e, d + i, 1] + x[e, d + i, 2] + x[e, d + i, 3] for i in range(6)) <= 5)

    # 5. 夜勤基礎配置（独立可能メンバー与サポート需要メンバー）
    support_needs_names = set(CONFIG.get('NEEDS_NIGHT_SUPPORT', []))
    support_needs_indices = {e for e in range(num_emp) if employees[e] in support_needs_names}
    
    noc_base_pool = [e for e in noc_seniors if e not in support_needs_indices]
    csc_base_pool = [e for e in csc_seniors if e not in support_needs_indices]

    for d in range(num_days):
        model.Add(sum(x[e, d, 2] for e in noc_base_pool) == 1)
        model.Add(sum(x[e, d, 2] for e in csc_base_pool) == 1)
        model.Add(sum(x[e, d, 2] for e in noc_seniors) <= 2)

    # 5b. 夜勤パートナーシップ制約 (Soft Penalty)
    support_pool = CONFIG.get('NIGHT_SUPPORT_POOL', {})
    for target_idx in support_needs_indices:
        target_name = employees[target_idx]
        pool_names = support_pool.get(target_name, [])
        valid_pool_indices = [employees.index(name) for name in pool_names if name in employees]
        
        if valid_pool_indices:
            for d in range(num_days):
                s_miss = model.NewBoolVar(f'slack_support_miss_{target_idx}_{d}')
                model.Add(s_miss >= x[target_idx, d, 2] - sum(x[p_idx, d, 2] for p_idx in valid_pool_indices))
                objective_terms.append(-CONFIG['PENALTY_NIGHT_PARTNER_MISSING'] * s_miss)

    # 6. 月間夜勤/日勤回数制限
    for e in noc_seniors:
        model.Add(sum(x[e, d, 1] for d in range(num_days)) >= 4)
        if e in support_needs_indices:
            model.Add(sum(x[e, d, 2] for d in range(num_days)) == CONFIG['SUPPORT_NIGHT_TARGET'])
        else:
            model.Add(sum(x[e, d, 2] for d in range(num_days)) >= 2)
            model.Add(sum(x[e, d, 2] for d in range(num_days)) <= CONFIG['LIMIT_NOC_NIGHT_MAX'])

    for e in csc_seniors:
        if e in support_needs_indices:
            model.Add(sum(x[e, d, 2] for d in range(num_days)) == CONFIG['SUPPORT_NIGHT_TARGET'])
        else:
            model.Add(sum(x[e, d, 2] for d in range(num_days)) >= 2)
            model.Add(sum(x[e, d, 2] for d in range(num_days)) <= CONFIG['LIMIT_CSC_NIGHT_MAX'])

    for e in trainee_indices:
        model.Add(sum(x[e, d, 2] for d in range(num_days)) >= 2)
        model.Add(sum(x[e, d, 2] for d in range(num_days)) <= 5)

    # 7. 夜勤回数格差最小化 (Min-Max Range Penalty)
    noc_n_max = model.NewIntVar(0, num_days, 'noc_n_max')
    noc_n_min = model.NewIntVar(0, num_days, 'noc_n_min')
    for e in noc_base_pool:
        total_n_e = sum(x[e, d, 2] for d in range(num_days))
        model.Add(total_n_e <= noc_n_max)
        model.Add(total_n_e >= noc_n_min)
    objective_terms.append(-CONFIG['PENALTY_NIGHT_VARIANCE'] * (noc_n_max - noc_n_min))

    csc_n_max = model.NewIntVar(0, num_days, 'csc_n_max')
    csc_n_min = model.NewIntVar(0, num_days, 'csc_n_min')
    for e in csc_base_pool:
        total_n_e = sum(x[e, d, 2] for d in range(num_days))
        model.Add(total_n_e <= csc_n_max)
        model.Add(total_n_e >= csc_n_min)
    objective_terms.append(-CONFIG['PENALTY_NIGHT_VARIANCE'] * (csc_n_max - csc_n_min))

    # 8. 孤立休/孤立出勤与 5連休抑制
    r = {}
    for e in range(num_emp):
        for d in range(num_days):
            r[e, d] = model.NewBoolVar(f'is_rest_{e}_{d}')
            model.Add(r[e, d] == x[e, d, 0] + x[e, d, 4])

    for e in range(num_emp):
        for d in range(1, num_days - 1):
            rest_prev = r[e, d - 1]
            rest_curr = r[e, d]
            rest_next = r[e, d + 1]

            s_iso_rest = model.NewBoolVar(f'iso_rest_{e}_{d}')
            model.Add(s_iso_rest >= rest_curr - rest_prev - rest_next)
            objective_terms.append(-CONFIG['PENALTY_ISOLATED_REST'] * s_iso_rest)

            s_iso_work = model.NewBoolVar(f'iso_work_{e}_{d}')
            model.Add(s_iso_work >= rest_prev + rest_next - rest_curr - 1)
            objective_terms.append(-CONFIG['PENALTY_ISOLATED_WORK'] * s_iso_work)

        for d in range(num_days - 4):
            prefilled_rests = sum(1 for i in range(5) if (e, d+i) in prefilled and prefilled[(e, d+i)] in ['休', '年'])
            if prefilled_rests < 5:
                s_5_rest = model.NewBoolVar(f'slack_5_rest_{e}_{d}')
                model.Add(sum(r[e, d + i] for i in range(5)) <= 4 + s_5_rest)
                objective_terms.append(-CONFIG['PENALTY_REST_5_CONSECUTIVE'] * s_5_rest)

    # 9. 特殊任務与日勤基礎人数防御
    special_task_words = {'CPN\n設置作業日', 'CPN設置作業日', '保全', 'メ'}

    for d in range(num_days):
        is_hol = d in holiday_dates

        for m in mentor_indices:
            model.Add(y_d2[m, d] <= x[m, d, 1])
            if (m, d) in prefilled and str(prefilled[(m, d)]).strip() in special_task_words:
                model.Add(y_d2[m, d] == 0)

        d2_sum = sum(y_d2[m, d] for m in mentor_indices)
        model.Add(d2_sum == (0 if is_hol else 1))

        noc_mentors = [m for m in mentor_indices if m in noc_seniors]
        noc_special_task_count = sum(1 for e in noc_seniors if (e, d) in prefilled and str(prefilled[(e, d)]).strip() in special_task_words)
        
        noc_pure_d_sum = sum(x[e, d, 1] for e in noc_seniors) - sum(y_d2[m, d] for m in noc_mentors) - noc_special_task_count
        noc_support_d = sum(x[e, d, 1] for e in noc_seniors if e in support_needs_indices)

        if is_hol:
            model.Add(noc_pure_d_sum == 1 + noc_support_d)
            model.Add(noc_support_d <= 1)
        else:
            model.Add(noc_pure_d_sum >= 2)
            for e_sup in support_needs_indices:
                if teams[e_sup] == 'NOC':
                    objective_terms.append(50 * x[e_sup, d, 1])

        s_noc_over3 = model.NewIntVar(0, num_emp, f's_noc_over3_{d}')
        model.Add(noc_pure_d_sum - 3 <= s_noc_over3)
        objective_terms.append(-CONFIG['PENALTY_NOC_PURE_D_OVER3'] * s_noc_over3)

        csc_support_d = sum(x[e, d, 1] for e in csc_seniors if e in support_needs_indices)
        if is_hol:
            model.Add(sum(x[e, d, 1] for e in csc_seniors) == 1 + csc_support_d)
            model.Add(csc_support_d <= 1)
        else:
            model.Add(sum(x[e, d, 1] for e in csc_seniors) >= 1)
            s_csc_d2 = model.NewIntVar(0, num_emp, f's_csc_d2_{d}')
            model.Add(sum(x[e, d, 1] for e in csc_seniors) + s_csc_d2 >= 2)
            objective_terms.append(-CONFIG['PENALTY_CSC_WEEKDAY_D_SHORT'] * s_csc_d2)
            
            for e_sup in support_needs_indices:
                if teams[e_sup] == 'CSC':
                    objective_terms.append(50 * x[e_sup, d, 1])

        if len(trainee_indices) > 0:
            if is_hol:
                model.Add(sum(x[e, d, 2] for e in trainee_indices) == 0)
            else:
                model.Add(sum(x[e, d, 2] for e in trainee_indices) <= 1)

    d2_max = model.NewIntVar(0, num_days, 'd2_max')
    d2_min = model.NewIntVar(0, num_days, 'd2_min')
    for m in mentor_indices:
        total_d2_m = sum(y_d2[m, d] for d in range(num_days))
        model.Add(total_d2_m <= d2_max)
        model.Add(total_d2_m >= d2_min)
    objective_terms.append(-CONFIG['PENALTY_D2_RANGE'] * (d2_max - d2_min))

    for m in mentor_indices:
        for d in range(num_days - 2):
            s_d2_dense = model.NewBoolVar(f'slack_d2_dense_{m}_{d}')
            model.Add(s_d2_dense >= y_d2[m, d] + y_d2[m, d + 1] + y_d2[m, d + 2] - 1)
            objective_terms.append(-CONFIG['PENALTY_D2_DENSE'] * s_d2_dense)

    for e in range(num_emp):
        for d in range(num_days):
            for gap in range(1, 6):
                if d + gap < num_days:
                    s_night_gap = model.NewBoolVar(f'slack_night_gap_{e}_{d}_{gap}')
                    model.Add(s_night_gap >= x[e, d, 2] + x[e, d + gap, 2] - 1)
                    pen_w = CONFIG['PENALTY_NIGHT_GAP_TIGHT'] if gap <= 2 else CONFIG['PENALTY_NIGHT_GAP_MODERATE']
                    objective_terms.append(-pen_w * s_night_gap)
                    
            start_7 = max(0, d - 6)
            s_7_night = model.NewIntVar(0, 7, f'slack_7_night_{e}_{d}')
            model.Add(sum(x[e, d_i, 2] for d_i in range(start_7, d + 1)) - 2 <= s_7_night)
            objective_terms.append(-CONFIG['PENALTY_ROLLING_7DAY_NIGHT_OVER2'] * s_7_night)

    # 10. 连休奖励与休日休息奖励 (Block Rewards)
    for e in range(num_emp):
        for d in range(num_days - 1):
            b2 = model.NewBoolVar(f'block_2_{e}_{d}')
            model.AddBoolAnd([r[e, d], r[e, d + 1]]).OnlyEnforceIf(b2)
            objective_terms.append(CONFIG['REWARD_REST_BLOCK_2'] * b2)

        for d in range(num_days - 2):
            b3 = model.NewBoolVar(f'block_3_{e}_{d}')
            model.AddBoolAnd([r[e, d], r[e, d + 1], r[e, d + 2]]).OnlyEnforceIf(b3)
            objective_terms.append(CONFIG['REWARD_REST_BLOCK_3'] * b3)

        for d in range(num_days - 3):
            b4 = model.NewBoolVar(f'block_4plus_{e}_{d}')
            model.AddBoolAnd([r[e, d], r[e, d + 1], r[e, d + 2], r[e, d + 3]]).OnlyEnforceIf(b4)
            objective_terms.append(CONFIG['REWARD_REST_BLOCK_4PLUS'] * b4)

        # 単独不可メンバーには休日休息の高額インセンティブ（20,000点）を付与
        for d in holiday_dates:
            if e in support_needs_indices:
                objective_terms.append(CONFIG['REWARD_SUPPORT_HOLIDAY_REST'] * r[e, d])
            else:
                objective_terms.append(CONFIG['REWARD_HOLIDAY_REST'] * r[e, d])

    # 设置最大化目标函数
    model.Maximize(sum(objective_terms))

    # ================= 3. 最速求解 (CP-SAT Custom Callback) =================
    solver = cp_model.CpSolver()
    
    threads_count = os.cpu_count() or 8
    solver.parameters.num_workers = threads_count
    solver.parameters.max_time_in_seconds = 600.0
    solver.parameters.relative_gap_limit = 0.01  # Gap <= 1% (99% 最優) で自動停止
    
    # デフォルトの雑多な内部ログを非表示
    solver.parameters.log_search_progress = False

    # 直感的な日本語ログ出力用のコールバッククラス
    class CleanJapaneseLogger(cp_model.CpSolverSolutionCallback):
        def __init__(self):
            cp_model.CpSolverSolutionCallback.__init__(self)
            self.__solution_count = 0

        def on_solution_callback(self):
            self.__solution_count += 1
            time_spent = self.WallTime()
            obj_val = self.ObjectiveValue()
            best_bound = self.BestObjectiveBound()
            
            # Gap (%) の計算
            if best_bound != 0:
                gap = abs(obj_val - best_bound) / abs(obj_val) * 100
            else:
                gap = 0.0
                
            print(f"⏱️ [{time_spent:5.1f}秒] 発見した解 #{self.__solution_count:<3} | スコア: {int(obj_val):>10} 点 | 残りGap: {gap:5.2f}%")

    print(f"⚡ 高性能 Google CP-SAT ソルバー実行中 (スレッド: {threads_count}, 目標Gap <= 1.00%)...\n")
    
    # カスタムコールバックを指定して実行
    status = solver.Solve(model, CleanJapaneseLogger())

    # ================= 4. 結果解析与 Excel 出力 =================
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print('\n' + '=' * 65)
        print('🎉 解の導出に成功しました！(目標精度を達成したため探索を終了)')
        print('=' * 65 + '\n')
    else:
        print('\n' + '=' * 65)
        print('❌ プレチェックは通過しましたが、制約条件を満たす解が見つかりませんでした。')
        print('=' * 65)
        return

    result_schedule = [['' for _ in range(num_days)] for _ in range(num_emp)]
    val_map = {0: '休', 1: 'D', 2: 'N', 3: '明', 4: '年'}

    for e in range(num_emp):
        for d in range(num_days):
            if (e, d) in prefilled and (e, d) not in red_border_cells:
                result_schedule[e][d] = prefilled[(e, d)]
            else:
                for s in shifts:
                    if solver.Value(x[e, d, s]) == 1:
                        result_schedule[e][d] = val_map[s]

    for d in range(num_days):
        for tr_i in trainee_indices:
            if result_schedule[tr_i][d] == 'D': result_schedule[tr_i][d] = 'D3'
        for m in mentor_indices:
            if solver.Value(y_d2[m, d]) == 1:
                result_schedule[m][d] = 'D2'

    red_font = Font(color='FF0000')
    black_font = Font(color='000000')
    purple_fill = PatternFill(start_color='E6E6FA', end_color='E6E6FA', fill_type='solid')
    green_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')

    wb_out = openpyxl.load_workbook(file_path)
    ws_out = wb_out[sheet_name]

    for e in range(num_emp):
        r_idx = target_rows[e]
        for d in range(num_days):
            c = day_start_col + d
            val = result_schedule[e][d]
            cell = ws_out.cell(row=r_idx, column=c)

            if (e, d) not in prefilled or (e, d) in red_border_cells:
                cell.value = val
                if val == '休': cell.font = red_font
                elif val in ['N', '明']: cell.fill = purple_fill; cell.font = black_font
                elif val == 'D2': cell.fill = green_fill; cell.font = black_font
                else: cell.font = black_font

    out_file = 'Book1_出力結果.xlsx'
    wb_out.save(out_file)
    print(f'📁 シフト表の出力が完了しました: {out_file}')


if __name__ == '__main__':
    solve_schedule_with_cpsat('Book1.xlsx', 'BookLM.xlsx')