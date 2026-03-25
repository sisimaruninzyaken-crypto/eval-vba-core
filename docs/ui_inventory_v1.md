# frmEval UI台帳（4分類・正式版）

作成基準:
- `docs/frmEval_control_reference_audit.md`（UI台帳 v1.1 / A-B-C整理）
- コード上の直接参照・生成処理（`frmEval.frm`, `modPhysEval.bas`, `MMT.bas`, `modROMIO.bas`, `modEvalIOEntry.bas`, `modADL_Probe.bas`, `modLayoutHeader.bas`, `modUFDumpSafe.bas`）
- 推測は使用しない

---

## 1. UI台帳（A〜D分類）

### A. 実測固定UI

- 実行時に必ず存在する（または存在前提で直接参照される）
- 名前・親子経路が実測系資料または初期化経路で確認可能

| ControlName | ControlType | Parent | TopBlock | 根拠（実測 or 初期化経路） |
|---|---|---|---|---|
| frHeader | Frame | frmEval | Header / Global | v1.1でA分類。`modLayoutHeader.Align_LoadPrevButton_NextToHdrKana` が `frHeader` を直接参照。 |
| txtHdrKana | TextBox | frHeader | Header / Global | v1.1でA分類。`modLayoutHeader` / `modHdrProbe` が直接参照。 |
| cmdHdrLoadPrev | CommandButton | frHeader | Header / Global | v1.1でA分類。`modLayoutHeader` / `modHdrProbe` が直接参照。 |
| MultiPage1 | MultiPage | frmEval | 全体ルート | `MMT.GetMMTPage_FromPhys` / `modEvalDataArchive.Dump_AllLayout_Snapshot` で直接参照。 |
| Frame3 | Frame | MultiPage1.Pages(2) | Physical | `MMT.GetMMTPage_FromPhys` で `Pages(2).Controls("Frame3")` を直接参照。 |
| mpPhys | MultiPage | Frame3 | Physical | `MMT.GetMMTPage_FromPhys` および `modPhysEval.EnsurePhysMulti` で名称固定。 |
| Frame4 | Frame | MultiPage1.Pages(3) | ADL | `modEvalDataArchive.Dump_AllLayout_Snapshot` で `Pages(3).Controls("Frame4")`。 |
| mpADL | MultiPage | Frame4 | ADL | `modADL_Probe` / `modEvalDataArchive` で `mpADL` 直接参照。 |
| Frame6 | Frame | MultiPage1.Pages(5) | Walk-Host | `modEvalDataArchive` と `frmEval`（DailyLog系）で `Frame6` 直接参照。 |
| MultiPage2 | MultiPage | Frame6 | Walk | v1.1でB分類だが、現行コードでは `modEvalDataArchive` が固定経路として直接参照。 |
| Frame26 | Frame | MultiPage2.Pages(1) | Walk-Nested | `modEvalDataArchive` が `MultiPage2.Pages(1).Controls("Frame26")` を直接参照。 |
| MultiPage3 | MultiPage | Frame26 | Walk-Nested | `modEvalDataArchive` が `Frame26.Controls("MultiPage3")` を直接参照。 |
| mpCogMental | MultiPage | Cognitiveホスト | Cognitive | v1.1でA分類。`modEvalIOEntry.Load/Save_CognitionMental` が直接参照。 |
| pgCognition | Page | mpCogMental | Cognitive | v1.1でA分類。`frmEval.BuildCogMentalTabs` が `Pages(0).Name="pgCognition"` を設定。 |
| pgMental | Page | mpCogMental | Cognitive | v1.1でA分類。`frmEval.BuildCogMentalTabs` が `Pages(1).Name="pgMental"` を設定。 |

> 注記: Aは「現行コードが存在前提で直接依存しているUI」を採用。

---

### B. 規則生成UI

- 命名規則で生成
- 個数は配列・定数・ページ構成に依存

| 命名規則 | ControlType | 対象ブロック | 個数確定可否 | 根拠 |
|---|---|---|---|---|
| `txtROM_<Layer>_<Joint>_<Motion>_<R/L>` | TextBox | Physical（ROM） | 条件付きで可（定義配列に依存） | `modROMIO.SaveROMToSheet/LoadROMFromSheet` が同命名規則で参照。 |
| `txtROM_<Layer>_Memo` | TextBox | Physical（ROM） | 可（Layer数依存） | `modROMIO` が `txtROM_*_Memo` を規則参照。 |
| `cboR_<筋名>` / `cboL_<筋名>` / `lbl_<筋名>` | ComboBox / Label | Physical（MMT） | 可（`BuildMMTPage` の items配列依存） | `MMT.BuildMMTPage` が `MakeCbo/MakeLbl` で生成。 |
| `lblHdr*`（MMT見出し） | Label | Physical（MMT） | 可（固定3件） | `MMT.BuildMMTPage` で固定生成。 |
| `cmbBI_<0..9>` | ComboBox | ADL（BI） | 可（10固定） | `modADL_Probe` が `For i=0 To 9` で参照。 |
| `cmbIADL_<0..8>` | ComboBox | ADL（IADL） | 可（9固定） | `modADL_Probe` が `For i=0 To 8` で参照。 |

---

### C. 再生成UI

- 初期化・再描画で削除→再作成、または同等の作り直しを行うUI

| 対象コントロール群 | 対象ブロック | 再生成トリガ | 根拠（該当プロシージャ） |
|---|---|---|---|
| `MMTGEN` tag の MMT子コントロール（`lbl_*`, `cboR_*`, `cboL_*` 等） | Physical（MMT） | MMT構築時 / 再構築時 | `MMT.MMT_ClearGen` で削除後、`MMT.BuildMMTPage` で再生成。 |
| `mpMMTChildGen`（MMT子MultiPage） | Physical（MMT） | MMT構築時 | `MMT.GetMMTChildTabs` が未存在時にAdd、`MMT_BuildChildTabs_Direct` が内容再構築。 |
| 旧MMT残骸（legacy名一致コントロール） | Physical（MMT） | MMT移行処理時 | `MMT.RemoveLegacyMMTControlsFromPage` が Remove/Hide 実行。 |
| `lstDailyLogList` | DailyLog | 履歴リスト再構成 | `frmEval.BuildDailyLog_HistoryList` が既存 `lstDailyLogList` を remove して add。 |
| `fraDailyLog` 配下の一部レイアウト要素 | DailyLog | 日次ログタブ初期化/調整 | `frmEval.BuildDailyLogTab` / `BuildDailyLogLayout` / `BuildDailyLog_StaffAndNote` 呼出し系列。 |

---

### D. 注意UI（運用上リスクあり）

| ControlName（または説明） | 問題種別 | リスク内容 | 根拠 |
|---|---|---|---|
| `Frame31` 固定参照（旧） | 旧導線/旧参照 | 現行レイアウト不一致時に参照破綻 | v1.1（C分類）で「旧レイアウト残骸」扱い。 |
| `btnLoadPrevCtl`（旧イベント導線名） | 旧導線ボタン | 直接参照が残るとヘッダ再配置フォールバック依存が増える | v1.1（C分類）記載。 |
| `Label100`（旧アンカー名） | 命名/責務不一致 | 存在しない場合に代替推定が必要になる | v1.1（B分類注記）に「旧アンカー名」記載。 |
| ADL「起居」内の無名Combo（立ち上がり/立位保持） | 無名コントロール | 名前参照不可のため位置・Caption近傍探索に依存 | `modADL_Probe.ResolveKyoUnnamedCombos` がラベル右側探索で取得。 |
| `MultiPage2`（過去は可変扱い） | 分類揺れ | v1.1では可変、現行コードは固定経路で直接参照しており設計判定の不一致リスク | v1.1 B分類 + `modEvalDataArchive` 直接参照の併存。 |

---

## 2. ブロック別サマリ

> 固定UI数は「コードで直接名称参照される主要コンテナ/主要入力」を基準にした概算。完全件数（783内訳）は `frmEval_tree_SAFE_LAST.txt` 実体照合が必要。

| ブロック | 固定UI数（概算） | 規則生成UI | 再生成UI | 注意UI |
|---|---:|---|---|---|
| Header / Global | 3+ | なし | なし | あり（旧導線名の残存注意） |
| Basic | 10+ | あり（`cmbBI_*`） | 低 | 低 |
| Posture | 1+（主要Frame） | 不明（コード上で規則名断定不可） | 不明 | 低 |
| Physical（ROM/MMT/Sensory/Reflex/Pain/Paralysis） | 6+（`Frame3`,`mpPhys`等） | あり（`txtROM_*`,`cboR_*`,`cboL_*`） | あり（MMT再生成） | あり（legacy除去系） |
| ADL | 3+（`Frame4`,`mpADL`等） | あり（`cmbIADL_*`） | 低 | あり（無名Combo） |
| Test | 1+（Walk境界外のTest固定名） | 不明 | 不明 | 低 |
| Walk | 4+（`MultiPage2`,`Frame26`,`MultiPage3`,`Frame6`） | 一部あり（配下生成） | 中 | 低 |
| Cognitive | 3+（`mpCogMental`,`pgCognition`,`pgMental`） | 低 | 低 | 低 |
| DailyLog | 5+（`fraDailyLog`,`txtDaily*`,`lstDailyLogList`） | 低 | あり（履歴List再生成） | 低 |

---

## 3. 設計上の取り扱いルール

1. 固定UI（A）
   - 名称変更・親変更を禁止（互換レイヤを伴わない変更不可）。
   - 保存/読込/イベント導線はAを基準キーとする。

2. 規則生成UI（B）
   - 台帳では「命名規則 + 生成元プロシージャ + 配列/定数」を管理単位とし、個別列挙を原則しない。
   - 変更時は規則名と参照側（Save/Load両方）を同時更新する。

3. 再生成UI（C）
   - 参照時は「生成前提」で扱い、未生成タイミングへの直接アクセスを禁止。
   - 保存対象抽出は再生成後の安定状態で実施する。

4. 注意UI（D）
   - 旧導線・無名・命名不整合は新規依存を追加しない（触らない原則）。
   - 触る場合は先に実測ダンプと呼出経路を再確認し、分類見直しを先行する。

---

## 4. 未確定事項

1. `frmEval_tree_SAFE_LAST.txt` 実体未添付のため、以下は完全断定不可。
   - 783件の全件内訳（各ブロック件数）
   - 現時点で存在する無名コントロールの総数

2. 操作依存で出現するUIの完全性。
   - タブ切替・初期化順で生成されるMMT/DailyLog配下の最終状態
   - Walk/Cognitive配下の一部ネストMultiPage配下要素

3. 再生成タイミング依存。
   - `MMT_BuildChildTabs_Direct` 実行前後でのコントロール集合差分
   - `BuildDailyLog_HistoryList` 実行前後でのリスト関連UI差分


