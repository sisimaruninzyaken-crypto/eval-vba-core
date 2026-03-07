# VBA Module Dependency Map

## 1. 目的と適用範囲
- **目的**: `frmEval` を起点に、主要モジュール間の呼び出し方向を可視化し、修正時の影響範囲追跡を容易にする。
- **適用範囲**: 現行エクスポート済み VBA ファイル（`.bas / .cls / .frm`）。
- **補足**: 依頼文では `_vba_export` 配下指定だが、本リポジトリでは同等ファイルがリポジトリ直下に存在するため、その実体を対象に整理した。
- **記法**: 本書の「依存」は **実コード上で確認できる直接呼び出し** を優先し、推測のみで断定しない。

---

## 2. 主要モジュール一覧（優先確認対象）
- `frmEval.frm`（UI 起点）
- `modEvalEntry.bas`（補助エントリ／検証ユーティリティ）
- `modEvalIOEntry.bas`（保存・読込ハブ）
- `modPlanGen.bas`（計画生成入力整形）
- `modKinrenPlanBasicCore.bas`（基本プラン組立ロジック）
- `modOpenAIResponses.bas`（OpenAI Responses API 呼び出し）
- `modPhysEval.bas`（身体機能 UI 構築）
- `modROMIO.bas`（ROM I/O）
- `modPainIO.bas`（Pain I/O）
- `modEvalPrintPack.bas`（印刷パック本体）
- `modEvalPrintPackLatest.bas`（最新行印刷エントリ）

---

## 3. 依存関係マップ（テキストツリー）

### 3.1 frmEval 起点（主要フロー）
```text
frmEval
├─ modEvalIOEntry
│  ├─ (動的呼び出し) SaveROMToSheet / LoadROMFromSheet 〔modROMIO〕
│  ├─ SavePainToSheet / LoadPainFromSheet 〔modPainIO〕
│  ├─ (動的呼び出し) SaveMMTToSheet / LoadSensoryFromSheet ほか
│  └─ modToneReflexIO
├─ modPhysEval
├─ modOpenAIResponses
└─（印刷ボタン経由）clsPrintBtnHook
   └─ modEvalPrintPackLatest
      └─ modEvalPrintPack
```

### 3.2 計画生成系（独立寄り）
```text
modPlanGen
├─ modEvalIOEntry（EVAL_SHEET_NAME, FindLatestRowByName）
├─ modHeaderMap（ReadStr_Compat）
└─ modPainIO（IO_GetVal）

modKinrenPlanBasicCore
└─ frmEval（GetLowerMMTMap_FromFrmEval で UI 参照）
```

### 3.3 ROM/Pain 系
```text
modROMIO
├─ frmEval（型参照・デバッグ参照）
└─ modHeaderMap 等（ReadStr_Compat / HeaderCol_Compat 経由）

modPainIO
├─ frmEval（Pain UI 参照）
├─ modEvalIOEntry（FindLatestRowByName）
└─ modHeaderMap（ReadStr_Compat）
```

---

## 4. 主要な呼び出し方向

### 4.1 UI起点
- `frmEval` → `modEvalIOEntry.LoadEvaluation_ByName_From(Me)`（読込トリガ）
- `frmEval` → `SaveEvaluation_Append_From Me`（保存トリガ）
- `frmEval` → `modPhysEval.EnsurePhysicalFunctionTabs_Under`（身体機能 UI 構築）
- `frmEval` → `OpenAI_BuildDraft`（AI 下書き生成）
- `frmEval` → `clsPrintBtnHook`（印刷ボタンイベント）

### 4.2 I/O系
- `modEvalIOEntry` が Save/Load のハブ。
- `modEvalIOEntry` から ROM・Pain・歩行・ADL・感覚など各 I/O へ分配。
- `modEvalIOEntry` は `Application.Run`（`IO_SafeRunSave/Load`）で一部機能を**文字列指定の動的呼び出し**。

### 4.3 判定系
- `modPlanGen` が EvalData の最新行を参照し、痛み・ADL・MMT などの要約バンドを生成。
- `modKinrenPlanBasicCore` が目標文構築ロジックを提供（下肢 MMT 取得時のみ `frmEval` を参照）。

### 4.4 AI系
- `frmEval` から `modOpenAIResponses.OpenAI_BuildDraft` を呼び、Responses API へ POST。

### 4.5 出力系
- `frmEval` の印刷ボタン（`clsPrintBtnHook`）→ `modEvalPrintPackLatest.Run_PrintPack_LatestRow`。
- `modEvalPrintPackLatest` → `modEvalPrintPack.Build_TestEval_PrintPack`。

---

## 5. 循環参照の疑いがある箇所
- **`frmEval` ⇄ `modEvalIOEntry`**
  - `frmEval` は `modEvalIOEntry` の保存・読込 API を呼ぶ。
  - `modEvalIOEntry` 側も `EnsureFormLoaded` で `frmEval` を直接参照する。
  - 設計上の相互依存（密結合）があるため、UI 名称変更・読込手順変更で双方に影響が出やすい。

- **`frmEval` ⇄ `modPainIO`（間接）**
  - `modPainIO` が `frmEval.Controls(...)` を直接参照する箇所があり、Pain UI 変更の影響を受けやすい。
  - `frmEval` 側の保存・読込は `modEvalIOEntry` 経由で `modPainIO` に到達する。

> 上記は「コード上の参照関係」に基づく循環の疑いであり、実行時に必ず無限ループになるという意味ではない。

---

## 6. 修正時に影響が広がりやすいモジュール
- **`modEvalIOEntry`（最重要）**
  - Save/Load の集約点で、複数セクションへ扇状に依存が広がる。
  - `Application.Run` の動的呼び出しを含むため、プロシージャ名変更の影響が検知しづらい。

- **`frmEval`**
  - 画面イベント、AI 呼び出し、印刷起点、I/O 起点が集中。
  - コントロール名変更が `modEvalIOEntry` / `modPainIO` / `modKinrenPlanBasicCore` へ波及しうる。

- **`modPainIO` / `modROMIO`**
  - UI コントロール名と EvalData ヘッダ名双方に依存し、列見出し変更・UI 変更の影響を受けやすい。

- **`modEvalPrintPackLatest` / `modEvalPrintPack`**
  - 印刷導線で連結されており、最新行選択ロジック・帳票生成ロジックの変更が連鎖しやすい。

---

## 7. 一文要約
`frmEval` を起点に、**`modEvalIOEntry` を中核ハブとして I/O が放射状に接続**し、AI（`modOpenAIResponses`）と印刷（`modEvalPrintPackLatest`→`modEvalPrintPack`）が別系統でぶら下がる構造で、特に `frmEval` と `modEvalIOEntry` の相互依存が影響拡大ポイントである。
