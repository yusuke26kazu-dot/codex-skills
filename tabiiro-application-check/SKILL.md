---
name: 申込書チェック
description: Use when the user provides Tabiiro/旅色 application review materials such as header paper photos or PDFs, application form photos or PDFs, and internal management system screenshots, and wants Codex to compare them and detect deficiencies, mismatches, missing flags, incorrect amounts, dates, payment settings, customer information, plan information, gross profit rules, companion ratio rules, or correction-seal requirements.
---

# 申込書チェック

Use this skill to inspect a submitted application package by comparing:

- Header paper photo/PDF
- Application form photo/PDF
- Internal management system screenshots
- Reference rule workbooks listed below

## Reference files

- N1/gross profit/incentive rules: `C:\Users\NX023066\Downloads\BM電子雑誌・AJ営業部／【公式】N1ルール・営業部粗利・インセン各種ルール.xlsx`
- Application form correction/manual rules: `C:\Users\NX023066\Downloads\BM電子雑誌サービス／【公式】申込書マニュアル.xlsx`

When the review involves同行比率,営業部粗利,N1,or修正印要否, open the relevant workbook and inspect the matching sheet/rule before concluding.

## Inputs to request

Header paper, application form, and internal management system materials are usually provided as attached screenshots/photos/PDFs, so do not require the user to type those paths in the prompt. Ask for missing materials only when the check cannot proceed reasonably. Expected prompt format:

```text
申込書チェック
新規／更新：
法人／個人：
備考：
```

Multiple screenshots are expected. Do not inspect screenshots below 応対履歴.

## Review workflow

1. Identify all plans shown across the header, application form, and system screenshots. Treat each system record as plan-specific: 旅色本誌, 旅色新着, 台湾本誌, 台湾新着, HP, PR, 入稿代行, etc. may each appear as separate system entries.
2. Extract visible facts from every material. Use exact values where visible; mark unreadable areas as `判読不可`.
3. Compare application-form amounts税込 against system amounts税抜 by converting with tax. System monthly split amounts may be rounded to the nearest yen.
4. Check that no overwritten text, double writing, or writing mistake appears.重ね書き・書き損じは不備.
5. Apply every checklist item in `references/checklist.md`.
6. If a mismatch requires修正印判断, consult the application form manual workbook before stating whether修正印 is required.
7. Return findings first. Group by severity: `要修正`, `確認必要`, `問題なし`.

## Output format

Use concise Japanese. For each finding include:

- 対象: plan/material/field
- 不備内容
- 根拠: what was seen in申込書,ヘッダー,orシステム
- 対応: what to fix and whether修正印確認が必要

If no issue is found, say so clearly and list any unreadable or missing materials as residual risk.
