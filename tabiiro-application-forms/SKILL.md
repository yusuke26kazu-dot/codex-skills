---
name: 申込書作成
description: Use when the user provides Tabiiro/旅色 acquisition plan details and wants the saved application form Excel generated. The skill fills the 旅色申込書 and, only when options are explicitly requested, the オプションサービス申込書 by referencing the saved plan list Excel workbook and the official Excel templates.
metadata:
  short-description: 旅色申込書をプラン情報から作成
---

# 旅色申込書作成

Use this skill when the user gives acquisition-plan details such as 掲載月, 獲得金額, 本誌プラン名, 新着有無, PR記事, HP, 入稿代行, 台湾 and asks for the Excel application form.

## Source files

Default source paths:

- Plan list workbook: `C:\Users\NX023066\Downloads\【BM】電子雑誌・AJ営業部／【公式】プラン一覧.xlsx`
- Normal form: `C:\Users\NX023066\Downloads\【A3横】旅色／申込書（ver.004）_260303.xlsx`
- Option form: `C:\Users\NX023066\Downloads\【A3横】旅色／オプションサービス申込書（ver.013）_260303.xlsx`

## Inputs to collect

Minimum useful format:

```text
掲載月：
獲得金額（税込）：
本誌プラン名：
新着プラン名：あり／なし
支払方法：
支払い開始月：
PR記事：
台湾：
入稿代行：
HP：
備考：
```

If PR記事 is present, also collect the PR service name/type and the Tuesday start date.

## Core rules

- 本誌欄には、プラン一覧の「新着なし」の税込金額を入れる。
- `本誌プラン名` means the normal form's `「旅色」情報掲載サービス` block.
- 新着は本誌とは別枠で記載する。新着の税込定価は原則 `220,000`。
- `台湾` means the normal form's `「旅色」多言語独自ページ制作サービス` block, not the normal 本誌 block.
- 台湾本誌はプラン一覧の `twA` を参照する。別名が必要な場合は `台湾プラン名` / `多言語プラン名` で指定する。
- 台湾本誌の掲載期間はプラン名末尾で判断する。`A` は12か月、`B` は24か月、`C` は36か月。`twA` は12か月。
- 台湾にも新着がある場合は、`「旅色」多言語独自ページ 新着特集掲載サービス` block にも反映する。
- HPが `あり` の場合は必ず `簡易HP_S_ssl` を指す。`HPS_ssl` / `簡易HPS_ssl` も同じ扱い。
- どのプランでも、弊社使用欄の担当名は `渡邊裕介`。
- 通常申込書の弊社使用欄には、本誌のプラン名を入れる。
- 申込日は、ユーザーが明示しない限り入力しない。
- 特記事項・備考は、ユーザーが明示しない限り入力しない。
- オプション申込書は、オプションについて明示がない場合は作成しない。
- 指定されていないプラン枠の支払額・月額に残っている `0` は削除する。オプション用紙も同じ。
- Excel出力時にシート保護・ブック保護をかけない。テンプレート側に保護がある場合も外して保存する。

## Amount and payment rules

For each specified service, fill:

- 定価税込
- 貴社向け特別値引きフラグ
- 値引き金額（税込）
- 値引き後の金額
- 支払い回数
- 1回ごとの支払額

Allocation rule:

- If 獲得金額（税込）が `0円`, specified services are full-discounted.
- If 新着あり and acquisition is at least `220,000円`, allocate paid amount to 新着 up to its max/list price, then overflow to 本誌.
- If 新着あり and acquisition is below `220,000円`, allocate to 新着 first unless the user says otherwise.
- If 新着なし, allocate to 本誌.
- 旅色の新着の掲載期間は1か月だが、支払回数は本誌の契約年数・支払回数に準じる。本誌が36回なら新着も36回。
- 0円のプランは支払方法欄・支払い開始月欄・支払回数欄には何も入力しない。1回あたり支払額欄は `0円` にする。テンプレートにもともと入っている `23` などの日付数字は残してよい。
- 口座振替で支払いが発生しているプランは、支払方法欄の口座振替チェックを入れる。
- `支払い開始月` が入力された場合、支払いが発生しているプランの支払方法欄にその開始年月を入れる。
- 支払開始日は特に指定がなければ `23日`。ライフペイメントの場合はユーザーが別途指定する。

## Date rules

- 本誌: 掲載月の25日。土日祝の場合は翌営業日。
- 新着: 掲載月の最終営業日。
- 台湾（多言語独自ページ）: 掲載月の最終営業日の前営業日。
- 台湾新着（多言語独自ページ新着）: 通常の新着と同じ。
- 入稿代行: 掲載月の1日から。
- HP: 掲載月の翌月25日。土日祝の場合は翌営業日。
- 新着・入稿代行の掲載期間は1か月。
- 本誌の掲載期間はプラン名の `A/B/C` とプラン一覧から判断する。
- PR記事は、指定された火曜日から3週間の掲載期間にする。

## Workflow

1. Convert the user's plan details to JSON.
2. Run `scripts/create_tabiiro_forms.py` with the bundled Python or available Python environment.
3. Verify the produced `.xlsx` opens in Excel when possible.
4. Return links only to the generated workbook(s). Do not produce option form unless explicitly requested by PR/HP/other option input.

Example:

```powershell
$json = @'
{
  "掲載月": "2026年7月",
  "獲得金額税込": 0,
  "本誌プラン名": "TG5A",
  "新着": "なし",
  "支払方法": "支払い無し",
  "支払い開始月": "",
  "台湾": false,
  "入稿代行": false,
  "HP": false,
  "PR記事": null
}
'@
$json | python C:\Users\NX023066\.codex\skills\tabiiro-application-forms\scripts\create_tabiiro_forms.py
```

Read `references/cell-map.md` only when modifying or troubleshooting workbook cell mappings.
