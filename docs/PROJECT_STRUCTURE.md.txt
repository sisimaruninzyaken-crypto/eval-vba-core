PROJECT_STRUCTURE

評価フォームAIシステム プロジェクト構造定義

1. 目的

本ドキュメントは、評価フォームAIシステムの プロジェクト全体構造 を定義するものである。

目的は以下の3点。

プロジェクトの構造を明確化する

新規開発・修正時の迷子を防ぐ

将来の保守・拡張・引継ぎを容易にする

本プロジェクトは Excel/VBA を基盤とした評価システム に、
AI文章生成機能を追加した小規模SaaS型ツールとして設計されている。

2. システム全体構造

本システムは次の 3層構造 を持つ。

① 評価基盤
② AI生成基盤
③ 運用基盤
3. 評価基盤（Evaluation Layer）

評価基盤は データ入力と評価情報の管理を担う基盤である。

中心となるコンポーネントは以下。

frmEval
EvalData
評価タブ群
主な役割

利用者評価入力

評価データ保存

LIFE用CSV出力

評価グラフ生成

帳票基盤

AI生成の入力データ供給

設計原則
frmEval = UIではなく評価スキーマ入力ハブ

評価入力はすべて frmEval を入口とする。

4. AI生成基盤（AI Generation Layer）

AI生成基盤は、評価情報から 計画書およびモニタリング文章を生成するロジック層である。

Basic生成パイプライン

処理順は以下に固定する。

frmEval
 ↓
抽出
 ↓
正規化
 ↓
判定
 ↓
計画構造
 ↓
AI生成
 ↓
出力
主要モジュール
modEvalIOEntry
    I/Oハブ
    保存・読込・抽出

modPlanGen
    判定ロジック
    Basic入力生成

modKinrenPlanBasicCore
    計画構造生成

modOpenAIResponses
    OpenAI API接続

modEvalPrintPack*
    帳票出力
AIの役割

AIは以下のみ担当する。

計画書生成
モニタリング生成

AIは判断主体ではなく 文章生成エンジンとして使用する。

5. 運用基盤（Operation Layer）

運用基盤は、本システムを 複数施設で安全に運用するための仕組みである。

対象規模

初期：10施設
最大：100施設
主な機能
ライセンス管理
利用回数管理
認証
強制停止
ログ管理
エラーコード
コスト監視
更新通知
FAQ
導入ページ
管理画面
自動バックアップ
復旧手順

運用基盤は プロダクトとして成立させるための装備である。

6. docsディレクトリ構造

設計書はすべて docs フォルダで管理する。

docs/

  PROJECT_STRUCTURE.md
      プロジェクト全体構造

  architecture/
      system_architecture_overview_v1.md

  pipeline/
      basic_generation_pipeline_v1.md

  modules/
      basic_module_responsibility_v1.md

  product/
      ai_system_design_summary.md

  operation/
      operation_todo_list.md
7. 設計原則

本プロジェクトでは以下の設計原則を採用する。

Excel本体主義
Excel = 本体
AI = 補助

AIはExcelシステムを置き換えない。

責務分離
入力 → 評価基盤
判定 → AI生成基盤
文章生成 → AI
運用管理 → 運用基盤
AI限定利用

AIの用途は以下のみ。

計画書生成
モニタリング生成

評価入力や保存処理には使用しない。

ローカル継続性

AI停止時でも以下が可能である。

評価入力
保存
閲覧
帳票出力
8. 現在の設計ボトルネック

現時点で確認されているボトルネックは以下。

modPlanGen
抽出
判定
文字列生成

が密結合。

将来の拡張時に整理対象となる可能性あり。

frmEval

現状

AI呼出
帳票出力呼出

の起点も持つ。

将来

UI起点

に寄せる整理の可能性あり。

運用基盤

認証・ログ・バックアップ等は
今後の実装対象。

9. 一文要約

本システムは

Excel評価基盤
    ↓
AI生成基盤
    ↓
運用基盤

の 三層構造で構成される。

10. 本ドキュメントの役割

この文書は

設計基準
引継ぎ資料
開発判断基準

として使用する。