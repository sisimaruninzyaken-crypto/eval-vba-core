\# frmEval コントロール参照整理（A/B/C分類）



\## 対象

\- `frmEval.frm`

\- `modEvalIOEntry.bas`

\- `modLayoutHeader.bas`

\- `MMT.bas`



\## 抽出した固定参照（今回の整理対象）

\- `Controls("Frame31")`

\- `Controls("Label100")`

\- `Controls("btnLoadPrevCtl")`

\- `Controls("mpCogMental")`

\- `Pages("pgCognition")`, `Pages("pgMental")`

\- `Controls("MultiPage2")`, `Controls("mpPhys")`, `Controls("Frame12")`



\## 分類

\### A: 必須コントロール（現行で必須）

\- `frHeader`

\- `txtHdrKana`

\- `cmdHdrLoadPrev`

\- `mpCogMental`（認知/精神タブ構築後）

\- `pgCognition`, `pgMental`（認知/精神タブ構築後）



\### B: 可変コントロール（存在しない可能性がある）

\- `MultiPage2`

\- `mpPhys`

\- `Frame12`

\- `Label100`（旧アンカー名。残っていれば使用、なければ代替Top推定）



\### C: 旧レイアウト残骸（参照削除/非依存化）

\- `Frame31` への固定参照

\- `btnLoadPrevCtl` へのヘッダー配置時のフォールバック依存



\## 実施ポリシー

\- A: 既存参照を維持

\- B: `SafeGetControl` / `SafeGetPage` で参照

\- C: 固定参照を削除し、現行ルート探索（`GetCogRootFrame` 等）へ統一



