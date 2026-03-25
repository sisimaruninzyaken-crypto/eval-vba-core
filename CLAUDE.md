# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **Japanese-language Excel/VBA-based physiotherapy evaluation form system** with AI-assisted care plan generation. It is designed as a small-scale SaaS tool supporting 10–100 facilities.

All UI, documentation, and domain logic are in Japanese.

## No Build System

This is a pure VBA project — there are no build commands, package managers, or test runners. Files are `.bas` (modules), `.cls` (class modules), and `.frm`/`.frx` (UserForms). These are exported from an Excel workbook (`.xlsm`) for version control. Changes must be re-imported into Excel to take effect.

## Development Governance (GOVERNANCE.md)

This project follows a strict human-in-the-loop workflow:

- **AI agents** (Codex/Claude): generate code diffs only — no pushing, no remote operations, no branch decisions
- **Human**: applies diffs locally, commits via PowerShell, pushes to GitHub
- **Single-task principle**: only one task is assigned to an AI agent at a time
- **"Done" definition**: GitHub push complete + working tree clean + origin/main synchronized
- Changes without a git commit are considered non-existent

## Three-Layer Architecture

```
① 評価基盤 (Evaluation Layer)      — frmEval + EvalData sheet
② AI生成基盤 (AI Generation Layer) — modPlanGen → modKinrenPlanBasicCore → modOpenAIResponses
③ 運用基盤 (Operation Layer)       — licensing/auth/logging (not yet implemented)
```

## AI Generation Pipeline (固定順序)

```
frmEval → 抽出 → 正規化 → 活動候補判定 → 主因判定 → 機能候補判定
       → modKinrenPlanBasicCore (計画構造生成)
       → modOpenAIResponses (AI文章生成: gpt-4.1-mini via OpenAI Responses API)
       → modEvalPrintPack (帳票出力)
```

**Key principle**: AI only converts structured data to natural text. All clinical judgment (activity goals, primary cause, function targets) is decided by system logic before the AI is called.

## Module Responsibility Map

| Module | Responsibility | Out of Scope |
|--------|---------------|--------------|
| `frmEval` | UI input hub, event origin | Clinical judgment, AI logic |
| `modEvalIOEntry` | Save/load to EvalData sheet (I/O hub) | Activity/cause judgment |
| `modPlanGen` | Extraction, Band/Tag normalization, judgment logic | UI, API, printing |
| `modKinrenPlanBasicCore` | ICF plan structure (Activity/Function/Participation, long/short-term) | UI, API, printing |
| `modOpenAIResponses` | OpenAI API calls only | Any judgment or logic |
| `modEvalPrintPack` | Report formatting and printing | Any meaning changes |
| `modROMIO` / `modPainIO` | Domain-specific data I/O | Judgment |
| `modPhysEval` | Physical evaluation tab UI construction | Planning, AI, printing |

**Boundary violations to avoid** (越境禁止):
- `frmEval` deciding clinical causes
- `modPlanGen` doing print formatting
- `modOpenAIResponses` making logic decisions
- Output layer changing meaning of text
- I/O layer determining activity goals

When modifying, direct changes to the correct module per the table above.

## Known Design Constraints

- **`modPlanGen` is tightly coupled**: extraction, normalization, and judgment are currently mixed — a known refactor candidate
- **`frmEval` scope creep**: currently also triggers AI calls and print output; the design intent is to restrict it to input hub only
- **Normalization layer is not yet separated**: normalization logic is spread across `modEvalIOEntry` and `modPlanGen`
- **Dynamic control names**: `frmEval` has 783+ controls; UI control name changes propagate widely across I/O modules
- **`Application.Run` usage**: some modules use string-based dynamic dispatch, making call chains harder to trace

## Data Storage

- **EvalData sheet**: primary evaluation data (row-per-evaluation)
- **EvalIndex sheet**: client master index
- Column headers managed by `modHeaderMap`
- Left/right limb data stored in comparable format
