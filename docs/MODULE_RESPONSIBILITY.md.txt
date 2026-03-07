# MODULE_RESPONSIBILITY
評価フォームAIシステム 主要モジュール責務定義

## 1. 目的
本ドキュメントは、主要モジュールの責務と責務外を定義し、
実装時の越境を防ぐための基準書である。

---

## 2. 基本原則
各モジュールは以下の原則に従う。

- UIは入力起点に集中する
- I/Oは保存と読込に集中する
- 判定は判定ロジックに集中する
- AIは文章生成に集中する
- 出力は整形と印刷に集中する

---

## 3. 主要モジュール一覧

### 3.1 frmEval
#### 責務
- 評価入力
- UI表示
- イベント起点
- 入力値保持

#### 責務外
- 主因判定
- 計画構造判定
- AIロジック判断
- 帳票意味決定

#### 備考
現状では一部AI呼出や出力起点を持つが、
思想上は入力ハブに寄せるのが原則。

---

### 3.2 modEvalIOEntry
#### 責務
- EvalData保存
- EvalData読込
- I/Oハブ
- ヘッダ整備
- 保存対象整理

#### 責務外
- 活動目標判定
- 主因判定
- AI文章生成

#### 備考
I/Oの中心モジュール。

---

### 3.3 modPlanGen
#### 責務
- Basic生成用抽出補助
- 判定ロジック
- Band化
- Tag化
- Basic入力生成

#### 責務外
- UI表示制御
- API接続
- 印刷整形

#### 備考
現状の最重要ロジック中核。
将来的に抽出・正規化・判定で分離候補。

---

### 3.4 modKinrenPlanBasicCore
#### 責務
- 計画構造生成
- Activity/Function/Participation骨格化
- 長期/短期構造化
- MMT対象筋選定

#### 責務外
- UI制御
- API通信
- 印刷出力
- 直接保存

#### 備考
文章化前の意味構造を決める層。

---

### 3.5 modOpenAIResponses
#### 責務
- OpenAI API接続
- AI応答取得
- 文章生成呼出

#### 責務外
- 主因判定
- 活動候補決定
- 保存処理
- UI処理

#### 備考
AI I/F専用層として扱う。

---

### 3.6 modEvalPrintPack 系
#### 責務
- 帳票整形
- 印刷処理
- 出力実行

#### 責務外
- 主因判定
- 活動判定
- AI判断
- 入力値決定

#### 備考
出力終端層。

---

### 3.7 modROMIO
#### 責務
- ROM保存
- ROM読込
- ROMデータI/O

#### 責務外
- ROM意味判定
- 主因判定
- 文章生成

---

### 3.8 modPhysEval
#### 責務
- 身体機能タブUI構築
- 身体機能入力基盤整備

#### 責務外
- 計画判定
- AI生成
- 帳票出力

---

### 3.9 clsPrintBtnHook
#### 責務
- UIと印刷処理の接着
- ボタンイベント中継

#### 責務外
- 判定
- 文章生成
- 保存処理

---

## 4. 越境の典型例
以下は越境として扱う。

- frmEval が主因を決める
- modPlanGen が印刷整形を行う
- modOpenAIResponses がロジック判断する
- 出力層が文章意味を変更する
- I/O層が活動目標を決定する

---

## 5. 改修時の寄せ先ルール
- UI変更 → frmEval / UI系
- 保存読込変更 → modEvalIOEntry / I/O系
- 判定変更 → modPlanGen
- 計画構造変更 → modKinrenPlanBasicCore
- AI接続変更 → modOpenAIResponses
- 印刷変更 → modEvalPrintPack系

---

## 6. 現時点の注意点
- frmEval は一部で責務が広い
- modPlanGen は密結合が強い
- 正規化専用モジュールは未分離

---

## 7. 一文要約
入力は frmEval、保存は modEvalIOEntry、判定は modPlanGen、構造化は modKinrenPlanBasicCore、AIは modOpenAIResponses、出力は modEvalPrintPack 系が担う。