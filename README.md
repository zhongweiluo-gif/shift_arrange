# 📅 shift_arrange (スマートシフト自動作成システム)

`shift_arrange` は、Google Drive API と数理最適化ソルバー（PuLP / HiGHS）を活用し、複雑な勤務条件や制約を考慮したシフト表を全自動で計算・生成する Python アプリケーションです。

---

## ✨ 主な機能 (Key Features)

* **📥 クラウドデータ連携**
  Google Drive API を利用し、クラウド上の Google スプレッドシート（前月および当月データ）をそのままの書式（セルカラーやレイアウト）を保持したまま自動取得・分割します。
* **🧠 高度な数理最適化 (PuLP + HiGHS)**
  `highspy` エンジンを活用した高速な HiGHS ソルバーにより、以下の複雑な制約を満たす最適なシフトを算出します：
  * **OJT 師弟関係・夜勤サポート体制の遵守**: 実習生（TRAINEES）および単独夜勤不可メンバー（NEEDS_NIGHT_SUPPORT）に対し、必ずサポート可能な同伴パートナーを同一夜勤シフトに配置。
  * **夜勤・日勤回数の柔軟な上下限制御**: メンバーの属性（独立先輩 / サポート対象 / CSC）ごとに夜勤回数の上限・下限・固定ターゲット数を JSON で制御。
  * **勤務間隔・連続出勤の制御**: 夜勤後の明け（明）および休日の確保、6連勤以上の厳格な防止。
  * **公休・祝日管理**: 当月の赤日（土日・祝日）を自動判定し、優先的な休暇割り当てを実施。
  * **連休の評価**: 2連休〜4連休以上のまとまった休暇に対して報酬スコアを付与。
* **🎨 役割後処理 & Excel 自動スタイリング (削峰填谷・D残高保護)**
  `openpyxl` を使用し、最適化されたシフト結果を Excel ファイルへ自動書き込み。純D（日勤）の残高を保護しながら役割（メ/保全/タスク/勤）を平坦化・自動装飾。
* **📦 GitHub Actions によるクロスプラットフォームビルド**
  Windows（`.exe`）および macOS 用の実行可能ファイルを GitHub Actions 上で自動ビルドします。

---

## 🛠 技術スタック (Tech Stack)

| カテゴリ | 使用技術 / ライブラリ |
| :--- | :--- |
| **言語** | Python 3.10 |
| **数理最適化** | PuLP, HiGHS (`highspy`) |
| **API / クラウド** | Google Drive API v3, Google OAuth 2.0 |
| **データ操作** | openpyxl |
| **CI/CD / ビルド** | GitHub Actions, PyInstaller |

---

## ⚙️ 設定ファイル (`config.json`) の変数仕様

システムの設定やパラメータはすべて `config.json` で管理されます。各変数の意味は以下の通りです：

### 1. ファイル・クラウド設定 (`FILE_SETTINGS`)
* **`SPREADSHEET_ID`**: Google スプレッドシートの固有ID。
* **`TARGET_MONTH_SHEET`**: 当月のシート名（例: `"202609"`）。
* **`PREV_MONTH_SHEET`**: 前月のシート名（例: `"202608"`）。
* **`TARGET_LOCAL_FILE` / `PREV_LOCAL_FILE`**: ローカル保存用 Excel ファイル名。
* **`PULP_OUTPUT_FILE` / `FINAL_OUTPUT_FILE`**: 中間出力および最終出力ファイル名。

### 2. メンバーおよび夜勤サポート設定 (【新機能】)
* **`TRAINEES`**: OJT対象の実習生リスト（※いない場合は `[]`）。
* **`PRIMARY_MENTORS`**: メンター（先輩）のリスト。
* **`NEEDS_NIGHT_SUPPORT`**: **✨【新規】** 単独での夜勤が不可で、必ず同伴パートナーが必要なメンバーのリスト。
* **`NIGHT_SUPPORT_POOL`**: **✨【新規】** 単独夜勤不可メンバーごとに、同伴可能なパートナー（先輩）のプールを定義するオブジェクト。

### 3. 勤務回数コントロールパネル (【新機能】)
* **`LIMIT_NOC_NIGHT_MAX`**: NOC独立先輩メンバーの当月夜勤上限回数（例: `4`）。
* **`LIMIT_CSC_NIGHT_MAX`**: CSC先輩メンバーの当月夜勤上限回数（例: `6`）。
* **`SUPPORT_NIGHT_TARGET`**: 単独夜勤不可メンバーの夜勤固定回数（例: `3`）。

---

## 💡 設定例 (Config Example)

### 実習生（TRAINEES）がおらず、夜勤サポート対象者がいる場合の設定例：

```json
{
  "USE_DYNAMIC_PUBLIC_REST": true,
  "FIXED_PUBLIC_REST": 11,
  
  "TRAINEES": [],
  "PRIMARY_MENTORS": [
    "森川亘浩", "今村孝", "高田康平", "檜垣優介", "森下優", "小林朔也",
    "駱中威", "佐藤俊一郎", "大原和士", "五十嵐康之", "ﾏﾝｻﾞﾆﾘｱﾘｶﾙﾄﾞ"
  ],

  "NEEDS_NIGHT_SUPPORT": ["駱中威", "佐藤俊一郎", "大原和士"],
  "NIGHT_SUPPORT_POOL": {
    "駱中威": ["森下優", "小林朔也"],
    "佐藤俊一郎": ["森下優", "小林朔也"],
    "大原和士": ["森川亘浩", "今村孝", "高田康平", "檜垣優介", "森下優", "小林朔也", "五十嵐康之", "ﾏﾝｻﾞﾆﾘｱﾘｶﾙﾄﾞ"]
  },

  "LIMIT_NOC_NIGHT_MAX": 4,
  "LIMIT_CSC_NIGHT_MAX": 6,
  "SUPPORT_NIGHT_TARGET": 3
}