# AGENTS.md

## Project

CivilHub is a desktop application for Japanese construction management workflows.

The main purpose is to support:

- 施工体制ツリーの作成
- 施工体制台帳Excelの作成・取込・出力
- 再下請負通知書Excelの作成・取込・出力
- 施工体系図の出力
- 協力会社情報の管理
- 現場ごとの帳票・添付資料管理

The primary user is a real construction site manager, not a developer.

Prioritize practical usability, stable Excel output, and safe incremental changes.

## Must-read files

Before making changes, read these files if they exist:

- `AGENTS.md`
- `CODEX.md`
- `README.md`
- `docs/`
- `handoff/`

`AGENTS.md` is the operational entry point for AI agents.  
`CODEX.md` contains the detailed CivilHub-specific development rules.

## Core rules

Always follow these rules:

- Do not break existing features.
- Do not perform large refactors unless explicitly requested.
- Make small, reviewable changes.
- Check existing implementation before editing.
- Do not guess Excel cell mappings.
- Do not change Excel layout, merged cells, print areas, borders, or row/column structure without explicit instruction.
- Do not convert completed sample Excel files into common master data.
- Do not overwrite user-created files.
- Do not delete handoff files unless explicitly instructed.
- Do not invent test commands.
- Do not claim tests passed unless they were actually run.
- Do not touch unrelated files.

## Excel rules

Excel output is critical.

For construction forms, visual fidelity and cell position are more important than internal code elegance.

When editing Excel-related logic:

- Preserve existing template layout.
- Preserve cell mappings.
- Preserve print settings.
- Preserve merged cells and borders.
- Confirm output files when possible.
- Keep internal-only fields out of Excel unless the form requires them.

## Sample and master data rules

Completed sample Excel files, especially the 日本道路 completed construction ledger sample, are only for:

- format understanding
- cell mapping
- import testing
- completed-output reference

They are not common master data.

Do not treat filled-in site-specific data as reusable company master data unless the user explicitly approves it item by item.

## Data separation

Do not mix project-specific data with common master data.

Project-specific data includes:

- 工事名
- 発注者
- 元請情報
- 施工体制ツリー
- 現場ごとの下請関係
- 現場固有の施工範囲
- 取込済み帳票
- 添付資料

Common master candidates include:

- 会社名
- 所在地
- 代表者名
- 建設業許可
- 業種
- 電話番号

Even for common master candidates, do not automatically import from completed samples.

## UI rules

The UI is for construction site staff.

Prioritize:

- clear Japanese labels
- easy-to-understand buttons
- visible main actions
- simple workflows
- safe delete confirmations
- minimal developer jargon

Avoid exposing internal implementation details to normal users.

## Git workflow

Before editing:

```bash
git status
```

After editing:

```bash
git status
git diff
```

Do not include unrelated formatting changes.
Do not touch unrelated files.
Do not commit generated cache files.

If there are existing unrelated uncommitted changes, leave them untouched and clearly report that they already existed.

## Testing

Run existing tests if they are documented.

If no test command is documented, say so clearly.
Do not invent commands.

At minimum, when relevant, check:

- app startup
- affected screen opens
- Excel import still works
- Excel output still works
- existing data is not broken
- errors are shown clearly to the user

## Model guidance

Use stronger reasoning models for:

- database changes
- Excel form output
- import/export logic
- UI restructuring
- multi-file refactors
- construction ledger logic

Lightweight models are acceptable for:

- documentation
- wording
- small labels
- comments
- simple README edits

Current expected model choices are:

- 5.5 for heavy changes
- 5.4 for normal changes
- 5.4 mini for light edits

Do not assume 5.3 is available.

## Working style

When given a task:

1. Inspect the existing files.
2. Identify the smallest safe change.
3. Avoid broad refactors.
4. Make the change.
5. Run available checks.
6. Summarize the diff.
7. Report anything not verified.

If the task is ambiguous, prefer a conservative implementation and clearly state assumptions.

## Relationship to CODEX.md

`AGENTS.md` is the operational entry point.

`CODEX.md` contains detailed CivilHub-specific rules.

If there is a conflict:

1. User's latest explicit instruction wins.
2. `AGENTS.md` operational safety rules come next.
3. `CODEX.md` detailed project rules come next.
4. Existing code behavior should be preserved unless instructed otherwise.
