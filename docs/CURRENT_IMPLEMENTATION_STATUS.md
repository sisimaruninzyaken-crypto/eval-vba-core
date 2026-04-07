# 現行実装到達点監査（2026-04-06）

## 調査方針
- 対象は **現行コード（`.frm/.bas/.cls`）のみ**。
- 旧ドキュメント・過去メモは一次情報として扱わない。
- 推測で埋めず、未確認は未確認として明示する。

---

## 1. 現在の全体構造サマリ

### 1-1. 主要モジュール（全体）
- フォーム: `frmEval.frm`（実運用UIの主本体）、`UserForm1.frm`（定義のみ・実装なし）。
- IOハブ: `modEvalIOEntry.bas`（保存/読込の統合ハブ、履歴シート解決、日次ログIO）。
- Basic生成パイプライン: `modBasicPipeline.bas` / `modExtractLayer.bas` / `modNormalizeLayer.bas` / `modJudgeLayer.bas` / `modKinrenPlanBasicCore.bas` / `modOpenAIResponses.bas`。
- 出力: `Module1.bas`（計画書・統合出力）、`modEvalPlanSheetOutput.bas`（計画書への転記）、`modLifeFuncCheckSheetOutput.bas`（生活機能チェックシート出力）、`modEvalPrintPack*.bas`（印刷系）。
- 物理機能UI/IO: `modPhysEval.bas`, `MMT.bas`, `modROMIO.bas`, `modParalysisIO.bas`, `modSenseReflexIO.bas`, `modToneReflexIO.bas`, `modPostureIO.bas`, `modPainIO.bas`。
- Utility/診断: `modCommonUtil.bas`, `ModUtil.bas`, `modSchema.bas`, `modImmediate.bas`, `modEvalDataArchive.bas`, `modUiInspect.bas`, `modUFDumpSafe.bas` ほか。

### 1-2. レイヤー別の実装状況

#### UI
- 実質的な中心は `frmEval`。初期化中に Pain/歩行/認知/日次ログ/ヘッダ/印刷ボタン等を動的構築し、各ボタンイベントを接続している。
- 保存・読込・計画生成・日次保存・日次抽出のユーザー操作入口は `frmEval` に集約。

#### IO
- `modEvalIOEntry` が保存・読込の統合ルート。
- 保存は `SaveEvaluation_Append_From -> SaveAllSectionsToSheet` で各セクション保存へ委譲。
- 読込は `LoadEvaluation_ByName_From -> LoadAllSectionsFromSheet` で各セクション読込へ委譲。
- 履歴シート解決は `ResolveUserHistorySheet` が担当（ID/氏名/かなの分岐含む）。
- 日次ログは同モジュールで `data\logs\DailyLog_yyyy.xlsx` を別管理。

#### Logic
- Basic生成の構造化判定は `Extract -> Normalize -> Judge -> BuildBasicPlanStructureFromJudge`。
- 旧系の `modPlanGen` も残存するが、主パイプラインは `modBasicPipeline` 経由。

#### AI
- `modOpenAIResponses` が Responses API (`/v1/responses`) を直接呼び出し。
- APIキーはワークブック名前定義 `OPENAI_API_KEY` を必須参照。
- AI呼び出しは Monitoring/目標6項目/プログラム/自主トレ/留意点/変化課題文に分割実行。

#### Output
- 計画書出力: `ExportEvalPlanSheet`（単体）/`ExportUnifiedPlanAndLifeFuncWorkbook`（統合）。
- 生活機能チェックシート出力: `modLifeFuncCheckSheetOutput`。
- 印刷: `frmEval`の印刷ボタン -> `Run_PrintPack_LatestRow` -> `Build_TestEval_PrintPack`。

#### Utility
- UI探索、ヘッダ解決、診断ダンプ、即時実行用のメンテ系マクロ群が多数存在。

### 1-3. 現在の中心ハブ
- **UIハブ**: `frmEval`。
- **データI/Oハブ**: `modEvalIOEntry`。
- **Basic生成ハブ**: `modBasicPipeline`。
- **出力ハブ**: `Module1` + `modEvalPlanSheetOutput` + `modLifeFuncCheckSheetOutput`。

---

## 2. 実際に到達している主要機能一覧

### 2-1. 保存（評価本体）
- 到達入口: `frmEval.btnSaveCtl_Click`（グローバル保存ボタンイベントもここへ集約）。
- 実処理: `SaveEvaluation_Append_From` でユーザー履歴シートを解決し、append行へ保存。
- 保存対象: Basic / 麻痺 / ROM / 姿勢 / MMT / 感覚 / トーン反射 / 疼痛 / TestEval / 歩行 / ADL / 認知。

### 2-2. 読込
- 到達入口: `frmEval.btnLoadPrevCtl_Click` / `cmdHdrLoadPrev_Click`。
- 実処理: `LoadEvaluation_ByName_From` が履歴シート解決→最新行特定→`LoadAllSectionsFromSheet`。
- 実際の読込対象はコード上で一部コメントアウトがあり、`LoadParalysisFromSheet` / `LoadROMFromSheet` / `LoadPostureFromSheet` / `LoadMMTFromSheet` が直接呼ばれない構成。

### 2-3. 評価入力
- `UserForm_Initialize` で評価関連UI（Pain/歩行/認知/日次等）を動的構築。
- 入力整形（IME設定、必須/範囲チェック、年齢同期等）は `frmEval` 内に集中。

### 2-4. 日次記録
- UIボタン生成: `BuildDailyLog_SaveButton` / `BuildDailyLog_ExtractButton`。
- 日次保存: `mDailySave_Click -> SaveDailyLog_Append`。
- 日次最新読込: `Load_DailyLog_Latest_FromForm`（現コード上、評価読込連動呼出はコメントアウト）。
- 月次モニタリング草案: `mDailyExtract_Click` 内で OpenAI呼出しあり、Excel出力へ連携。

### 2-5. Basic生成
- 生成ボタン: `btnGeneratePlanCtl_Click -> modBasicPipeline.ExportAllSheets`。
- 生成フロー: 抽出→正規化→判定→構造化→AI草案→計画書+生活機能チェックシートの統合出力。

### 2-6. AI草案生成
- `GenerateBasicPlanNarrative` で用途別に複数回 `OpenAI_BuildDraft` 実行。
- 継時比較（2回目以降想定）として `GenerateChangeAndIssue` も呼び出し可能。

### 2-7. 帳票/印刷/シート出力
- 計画書/統合ブック: `Module1.ExportEvalPlanSheet` / `ExportUnifiedPlanAndLifeFuncWorkbook`。
- 生活機能チェックシート単体: `modLifeFuncCheckSheetOutput.ExportLifeFuncCheckSheet`。
- 印刷: `AddPrintButton_TestEval` でフックし、`Run_PrintPack_LatestRow` 経由で印刷パック実行。


### 2-8. SaveAllSectionsToSheet と LoadAllSectionsFromSheet の対応表（コメントアウト読込の分類）

| 保存対象 | 保存側（到達） | 読込側（現状） | 状態 | 分類 | 根拠 |
|---|---|---|---|---|---|
| 麻痺 | `SaveAllSectionsToSheet -> SaveParalysisToSheet` | `LoadAllSectionsFromSheet` 内の直接呼出はコメントアウト。実際は `LoadBasicInfoFromSheet_FromMe` 内で `chkLoadParalysis` 条件付き実行 | 直接呼出停止・集約読込あり | 仕様上意図（読込集約） | `LoadAllSectionsFromSheet` のコメントと `LoadBasicInfoFromSheet_FromMe` 実呼出で確認 |
| ROM | `SaveAllSectionsToSheet -> SaveROMToSheet` | `LoadAllSectionsFromSheet` 内の直接呼出はコメントアウト。実際は `LoadBasicInfoFromSheet_FromMe` 内で `chkLoadROM` 条件付き実行 | 直接呼出停止・集約読込あり | 仕様上意図（読込集約） | 同上 |
| 姿勢 | `SaveAllSectionsToSheet -> SavePostureToSheet` | `LoadAllSectionsFromSheet` 内の直接呼出はコメントアウト。実際は `LoadBasicInfoFromSheet_FromMe` 内で `chkLoadPosture` 条件付き実行 | 直接呼出停止・集約読込あり | 仕様上意図（読込集約） | 同上 |
| MMT | `SaveAllSectionsToSheet -> SaveMMTToSheet` | `LoadAllSectionsFromSheet` 内の直接呼出はコメントアウト。実際は `LoadBasicInfoFromSheet_FromMe` 内で `QueueMMTLoadAfterUI` / `MMT.LoadMMTFromSheet` | 直接呼出停止・集約読込あり | 仕様上意図（UI構築後読込） | `LoadBasicInfoFromSheet_FromMe` の MMT 分岐で確認 |
| 日次ログ | 保存は評価本体とは別系統（`SaveDailyLog_Append`） | `LoadAllSectionsFromSheet` 内の `Load_DailyLog_Latest_FromForm owner` がコメントアウト | 評価読込と未統合 | 未実装（統合未完） | コメントアウト行が残り、別ボタン動作に依存 |

> 補足: 保存側と読込側の“呼出位置”は非対称だが、麻痺/ROM/姿勢/MMTは未接続ではなく **読込集約先が `LoadBasicInfoFromSheet_FromMe` に移動済み**。一方、日次ログのみは現時点で評価読込フローへの統合が未完。

---

## 3. 未完成 / 未統合 / 暫定 / 死に枝の整理

## 3-1. 実装途中・接続途中
- `LoadEvaluation_CurrentRow` は「廃止メッセージのみ」で実読込処理なし。
- `LoadAllSectionsFromSheet` 内コメントアウトのうち、麻痺/ROM/姿勢/MMTは `LoadBasicInfoFromSheet_FromMe` への集約で代替されている一方、日次ログは評価読込に未統合。
- `clsDailyLogList` は `WithEvents lb` を持つが、`frmEval` 側で `Set mDailyList.lb = ...` がコメントアウトされ、イベント未接続。

## 3-2. 仮置き/互換ラッパ
- `frmEval` に `BuildBliadlControls` / `BuildBIPage` / `BuildIADLPage` が no-op として残置。
- `modEvalPrintPackLatest.Build_TestEval_PrintPack_Forced` は実質 `Build_TestEval_PrintPack` 呼び出しラッパ。

## 3-3. 旧案残骸
- `ArchivePainIO_legacy_20251017.bas` が同系実装として残存。
- `UserForm1.frm` は実運用コードなし。

## 3-4. 呼ばれていない処理（repo内参照なしを確認できたもの）
- 例: `modBasicPipeline.RunBasicPlan`, `modEvalIOEntry.SaveEvaluation_Append`, `modEvalSupplementIO.AddHeaderArchiveDeleteButton`, `modCommonUtil.App_Main`, `modSchema.EnsureEvalDataSchema`, `modUiQuickInspect.ListFrmEvalControls` ほか多数。
- ただしVBAは手動マクロ実行やリボン/ボタン割当の外部経路があり得るため、**「現コード参照なし = 完全未使用」とは断定しない**。

## 3-5. 名前だけある処理
- 明示的 no-op 手続き（上記）
- メンテ/診断用途のDebug系サブルーチン群（実運用パスに未接続）

---

## 4. 現在地の総合判定

### 完成済み（コード上で一連到達が確認できる）
- `frmEval` 起点の評価UI起動と主要操作（保存・前回読込・計画生成・日次保存・印刷ボタン）。
- 保存時の統合書込み（複数セクション横断）。
- Basic生成～AI草案～計画書/生活機能シート出力。

### 運用可能
- 新規保存/追記保存（ユーザー履歴シート分岐あり）。
- 計画書と生活機能チェックシートのファイル生成。
- 日次ログの年別ファイル管理・追記。

### 未完成
- 読込ルートは「集約済み領域（麻痺/ROM/姿勢/MMT）」と「未統合領域（日次ログ）」が混在。
- 日次最新読込が評価本体読込に自動統合されていない（呼出コメントアウト）。
- イベントクラス（`clsDailyLogList`）の配線未完。

### ボトルネック
1. `frmEval` と `modEvalIOEntry` への責務集中（巨大化・局所修正の影響範囲が大きい）。
2. 保存対象と読込対象の非対称。
3. Legacy/診断コードと運用コードの混在により、現行有効経路の識別コストが高い。
4. AI呼び出しが同期HTTP依存で、失敗時の劣化運転設計が限定的。

### 次の実装着手点（優先順）
1. 保存/読込の対称性回復（コメントアウト部の扱いを仕様化して整理）。
2. 日次ログ連携（読込タイミング・イベント配線）を運用フローに統合。
3. `frmEval` / `modEvalIOEntry` の分割（UI組立、イベント、IO解決、バリデーションを分離）。
4. Legacy/診断モジュールを「運用」「保守」「退役候補」にタグ付け。

---

## 5. 旧資料との差分観点で重要な点（コード基準）

> ここでは「現コードから確実に言える変化のみ」を列挙。

1. **EvalData単一前提から、ユーザー履歴シート（`EV_*`）解決型へ実装重心が移動**。
2. **Basic生成は単独出力ではなく、計画書+生活機能チェックシートの統合出力が主経路**。
3. **日次ログは評価本体シートとは別の年別ファイル管理に分離**。
4. **読込系に未接続部分があり、保存された全項目が常に同等に復元される状態ではない**。
5. **旧/診断/検証マクロが多数残り、名称だけでは現行経路を誤認しやすい**。

---

## 6. 根拠となる主要モジュール・関数・呼び出し関係（要約）

```text
frmEval.btnSaveCtl_Click
  -> modEvalIOEntry.SaveEvaluation_Append_From
      -> ResolveUserHistorySheet
      -> SaveAllSectionsToSheet
      -> Save_CognitionMental_AtRow

frmEval.btnLoadPrevCtl_Click
  -> modEvalIOEntry.LoadEvaluation_ByName_From
      -> ResolveUserHistorySheet
      -> FindLatestValidEvalRowByIdentity / FindLatestRowByName
      -> LoadAllSectionsFromSheet

frmEval.btnGeneratePlanCtl_Click
  -> modBasicPipeline.ExportAllSheets
      -> GenerateBasicPlan
         -> ExtractBasicSourceData
         -> NormalizeBasicSourceData
         -> JudgeBasicPlanInputs
         -> BuildBasicPlanStructureFromJudge
         -> GenerateBasicPlanNarrative (OpenAI)
      -> GenerateChangeAndIssue (OpenAI, 条件付き)
      -> ExportUnifiedPlanAndLifeFuncWorkbook
         -> WriteEvalPlanSheet
         -> WriteLifeFuncCheckSheet

frmEval.mDailySave_Click
  -> modEvalIOEntry.SaveDailyLog_Append

frmEval.mDailyExtract_Click
  -> BuildMonthlyDraft_FromDailyLog
  -> OpenAI_BuildDraft
  -> ExportMonitoring_ToMonthlyWorkbook

frmEval AddPrintButton_TestEval (hook)
  -> clsPrintBtnHook.btn_Click
      -> modEvalPrintPackLatest.Run_PrintPack_LatestRow
         -> Build_TestEval_PrintPack
```

---

## 未確認事項（明示）
- 外部（Excelリボン、シート上ActiveX、個別ブックのマクロ割当）から呼ばれる可能性は本監査では未確認。
- 実行時データ（`EvalIndex` / `EV_*` / `DailyLog_*.xlsx`）の実態整合性は、コード読解のみでは確定不可。
- 文字化け混在のため、一部日本語ラベルは表示上の判読が困難。ロジック判定は制御名・呼出関係を優先した。

---

## 7. `SavePainToSheet` 現行有効定義・到達経路 監査（一次情報ベース）

### 7-1. 定義一覧（全件）
| 定義 | 種別 | ファイル | モジュール名 | 行 |
|---|---|---|---|---|
| `Public Sub SavePainToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)` | Public Sub | `ArchivePainIO.bas` | `ArchivePainIO` | 26 |
| `Public Sub SavePainToSheet(ByVal ws As Worksheet, ByVal r As Long, ByVal owner As Object)` | Public Sub | `ArchivePainIO_legacy_20251017.bas` | `ArchivePainIO_legacy_20251017` | 26 |

- `Function SavePainToSheet` / `Private Sub SavePainToSheet` は現行コード内で未検出。

### 7-2. 呼び出し元（逆引き）
| 呼び出し元 | 呼び出し形 | 行 |
|---|---|---|
| `modEvalIOEntry.SaveAllSectionsToSheet` | `Call SavePainToSheet(ws, r, owner)`（非修飾呼び出し） | `modEvalIOEntry.bas:147` |

- `frmEval` からの直接呼び出しは未検出。
- `SaveAllSectionsToSheet` 経由の統合保存ルートに組み込まれている。

### 7-3. `frmEval` 起点からの実到達経路（最短）
```text
frmEval.btnSaveCtl_Click
  -> modEvalIOEntry.SaveEvaluation_Append_From
    -> modEvalIOEntry.SaveAllSectionsToSheet
      -> SavePainToSheet  (非修飾呼び出し)
```

補助到達:
```text
frmEval.mGlobalSave_Clicked
  -> frmEval.btnSaveCtl_Click
  -> (上記と同一経路)
```

### 7-4. 現行有効定義の特定結果
- **到達先として確実に言えるのは「`SaveAllSectionsToSheet` の非修飾 `SavePainToSheet` 呼び出し」まで**。
- 同名 `Public Sub SavePainToSheet` が2モジュールに存在し、かつ呼び出しがモジュール修飾されていないため、
  **`.bas` テキスト監査のみでは最終バインド先（どちらのモジュールか）を一意確定できない**。
- ただしファイル命名上は `ArchivePainIO_legacy_20251017.bas` が legacy であることを示しているため、
  実運用意図としては `ArchivePainIO.bas` を現行本体とみなす設計意図が強い（＝意図推定）。

### 7-5. 疼痛保存の責務境界（`SavePainToSheet` 本体）
- 対象ページ特定: `ResolvePainPage(owner)` で pain UI ページを探索。
- 値取得: ComboBox / ListBox / CheckBox / VAS / 期間テキストを収集。
- 整形: `key:value` 形式の文字列へシリアライズ（区切り `|`, `:`, `,`）。
- 書込: `EnsureHeaderCol(ws, "IO_Pain")` で列を確保し、該当行へ保存。
- 担当外: 保存トランザクション（どの行へ書くか、履歴シート解決）は `modEvalIOEntry` 側。

### 7-6. 読込側との対応
- 疼痛読込側は `modPainIO.LoadPainFromSheet(ws, r, owner)`。
- `LoadAllSectionsFromSheet` から `IO_SafeRunLoad "LoadPainFromSheet"` で呼ばれる。
- よって Pain は「保存（`SavePainToSheet`）」「読込（`LoadPainFromSheet`）」の対が存在し、
  日次ログのように評価読込から切り離されている構造ではない。

### 7-7. 未確認事項
- 同名定義が2件ある状態で、実行時にどちらへバインドされるかは **VBEプロジェクト実体/コンパイル状態確認なしには確定不可**。
- `ArchivePainIO_legacy_20251017.bas` が実際にVBAプロジェクトへ読み込まれているかは、
  リポジトリ静的読解のみでは未確認。

---

## 8. `SavePainToSheet` 実行時バインド先確定のための「VBE実体確認」

### 8-1. 実体確認の実施結果
- 本リポジトリ/周辺ワークスペース内に、VBAプロジェクトを保持する `.xlsm/.xlsb/.xlam/.xls` を検出できなかった。
- 実行環境は Linux であり、Excel COM / VBE へ接続する実行基盤がない。
- したがって、**「VBEに実際にロードされているモジュール一覧」そのものは取得不能**。

### 8-2. 取得できた一次情報（コード実体）
- `import_to_excel.ps1` は `*.bas/*.cls/*.frm` をスクリプトディレクトリから全件収集し、同名既存コンポーネントを削除後に再インポートする実装。
- このスクリプトをそのまま使う運用なら、`ArchivePainIO.bas` と `ArchivePainIO_legacy_20251017.bas` はどちらもインポート対象に含まれる。
- ただし、これは **スクリプト仕様** であり、実際にどのブックへ実行されたか（実行履歴）は静的コードだけでは確定不可。

### 8-3. 要求項目への回答（実体確認観点）
1. VBEロード済みモジュール一覧: **未取得（環境制約）**。
2. `ArchivePainIO` / `ArchivePainIO_legacy_20251017` のVBE存在: **未確認**。
3. コンパイル対象含有: **未確認**（VBEプロジェクト実体が必要）。
4. `Option Private Module` の有無: 両モジュールとも定義あり（コード上確認）。
5. 同名 Public Sub 衝突の実コンパイル可否: **未確認**（実VBEでのCompile結果が必要）。
6. 実行時バインド先の最終確定: **未確定**（上記未確認のため）。

### 8-4. 結論（推測なし）
- 現時点の一次情報だけで確定できるのは、
  **`SaveAllSectionsToSheet` が非修飾で `SavePainToSheet` を呼ぶ事実** まで。
- 実行時バインド先を確定するには、
  1) 対象 `.xlsm` を特定し、
  2) 当該ブックの `VBProject.VBComponents` 一覧取得、
  3) `Debug->Compile VBAProject` 実行結果確認、
  4) 必要なら `ArchivePainIO.SavePainToSheet` へモジュール修飾して衝突排除、
  が必須。
