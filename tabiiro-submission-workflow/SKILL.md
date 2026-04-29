---
name: tabiiro-submission-workflow
description: Use when the user asks to prepare Tabiiro/旅色入稿 from an instruction through 入稿シート creation or update marking, image selection, Box Drive storage, and 入稿メール drafting for TG, TO, TL, TY, or HPS workflows.
metadata:
  short-description: 旅色入稿をシート作成からBox格納・メール案まで一括処理
---

# 旅色入稿ワークフロー

Use this skill when the user gives a Tabiiro/旅色入稿 instruction and expects the full workflow: create the correct 入稿シート, prepare approved images, store everything in Box Drive, and draft the入稿メール.

## Required Inputs

Collect or infer these items before producing final files:

- 案件名: 店舗名、施設名、商品名など。
- プラン: `TG5A`, `TG5C`, `TL5`, `TO3`, `HPS` など。本文と案件フォルダ名ではABCを残し、Boxのプラン階層と出力ファイル名ではABCを外す。
- 区分: `新規`, `更新`, `季節更新`。新規制作依頼なら `新規`。契約更新制作依頼なら `更新`。
- エリア: `京都`, `大阪` のように都道府県を付けない表記。
- 公開月または公開日: 月だけならその月の25日。25日が土日祝の場合は翌営業日。
- HPS同時入稿の有無: 有の場合は本誌公開日の翌月25日をHP掲載開始日にし、土日祝なら翌営業日。
- 指定画像、指定文章、申込用紙、補足条件がある場合はそれを優先する。

If a required item cannot be inferred safely, ask only for that missing item and then continue.

## Source Templates

Default local template paths:

- TG6: `C:\Users\NX023066\Downloads\入稿シート_TG6_【店舗名】_2025.1～.xlsx`
- TG5: `C:\Users\NX023066\Downloads\入稿シート_TG5_【店舗名】_2021.4～ (1).xlsx`
- TG4: `C:\Users\NX023066\Downloads\入稿シート_TG4_【店舗名】_2021.4～ (1).xlsx`
- TG3: `C:\Users\NX023066\Downloads\入稿シート_TG3_【店舗名】_2021.4～.xlsx`
- TG2: `C:\Users\NX023066\Downloads\入稿シート_TG2_【店舗名】_2021.4～.xlsx`
- TO: `C:\Users\NX023066\Downloads\新入稿シート_TO●_【施設名】_23.04～.xlsx`
- TL6: `C:\Users\NX023066\Downloads\レジャーTL6入稿シート_【施設名】202007改.xlsx`
- TL5: `C:\Users\NX023066\Downloads\レジャーTL5入稿シート_【施設名】202007改.xlsx`
- TL4: `C:\Users\NX023066\Downloads\レジャーTL4入稿シート_【施設名】202007改.xlsx`
- TL3: `C:\Users\NX023066\Downloads\レジャーTL3入稿シート_【施設名】202306改.xlsx`
- TL2: `C:\Users\NX023066\Downloads\レジャーTL2入稿シート_【施設名】.xlsx`
- HPS: `C:\Users\NX023066\Downloads\入稿シート／自社簡易HP_S＋ドメイン【施設名】_250107 (1).xlsx`

Use the matching template for the requested plan. For TO, select the appropriate plan sheet inside the same workbook.

## Excel Preservation Rules

The final workbook must preserve the template's dropdowns, fonts, cell sizes, row heights, merged cells, formulas, and formatting.

- Treat the original template as the source of truth.
- Prefer direct OOXML patching of cell values and formula cached values for existing Tabiiro templates.
- Do not round-trip the final workbook through tools that rebuild workbook XML if they change styles or data validations.
- Avoid saving final Tabiiro templates with `openpyxl` unless you have verified it preserves the needed validations. It can remove x14 dropdown validations.
- Keep existing label text such as `TEL：`, `FAX：`, `緯度：`, `経度：`, `昼：`, `夜：`, and write values after the colon in the same cell.
- Do not allow `選択してください。センタク` or any similar phonetic helper text to appear. Dropdown placeholders should remain only `選択してください`.
- Leave LP用キーワード blank unless the user explicitly asks otherwise.

Before delivery, compare the output with the template:

- `xl/styles.xml` should be unchanged unless there is an intentional style edit.
- Data validation counts, including x14 validations, should match the template.
- Cell dimensions and visible text should not be accidentally altered.

## TG Fill Rules

For TG2 through TG6, fill basic information consistently:

- 店舗名: use the instructed案件名.
- 店舗名ふりがな: hiragana.
- TEL/FAX: research online or use a prior申込用紙 if available. Enter after `TEL：` and `FAX：`.
- 掲載エリアガイド: choose from the existing dropdown based on address.
- 東京旅グルメジャンル: leave unchanged unless instructed.
- プレミアブック: choose `有` only for Sapporo, Kanagawa, Nagoya, Osaka, Kyoto, or Fukuoka restaurants when web sources indicate average spend over 10,000 yen. Otherwise leave untouched.
- 緯度/経度: use the map pin for the address and enter after `緯度：` and `経度：`.
- HP fields: own-domain official site only. If HPS is simultaneous, choose `無` and replace the URL cell with `簡易HP同時入稿のため`.
- SNS: enter active Twitter/X, Facebook, and Instagram URLs. If dormant, consult the user.
- 営業時間: use `10:00～14:00` style after `昼：` and `夜：`. Set lunch flag `有` or `無`.
- 席数: enter number of seats.
- 平均予算: preserve `昼：` and `夜：`; use the most appropriate form such as `1,000円～2,000円` or `2,000円`.
- アクセス: preserve the template format. For train access, use Google Maps walking time. If no nearby train, use bus stop and route. For car access, use `〇〇高速道路〇〇ICより約〇分`.
- 駐車場: choose `有` or `無`. If no parking but coin parking exists, write `近隣にコインパーキングあり` in the count/detail cell.
- 業態: choose the most relevant main genre from the dropdown. Select subgenres only when they genuinely fit.
- カード利用 and other cashless: choose from dropdowns.
- 店舗情報 flags: choose best values for 個室, 禁煙, 飲み放題, 座敷, 駅徒歩, 貸切, 食べ放題, テイクアウト.
- 衛生情報 and hygiene rows such as 手指 are not filled; leave existing values unchanged.

## 更新入稿 Rules

Use these rules when 区分 is `更新` or the user asks for 更新制作/契約更新制作. 更新入稿 assumes the public page already exists, so only apply the user-instructed corrections. Do not newly fill unrelated fields from web research unless the user asks.

- Edit only the instructed cells or image references.
- Mark every corrected cell with yellow fill.
- For basic information corrections and image filename cells, overwrite the cell with the new value, set the cell fill to yellow, and set the cell text to red.
- For image changes, the newly supplied image file name must match the image name being overwritten in the sheet. Rename the supplied image file if needed, then copy it into the `画像` folder.
- For text cells with character counts, copy the original text first, then change only the instructed part. Apply rich text formatting so only the changed words are red.
- If the instruction is to delete words rather than replace them, keep the deleted words in place, make only those words red, and add strikethrough.
- Preserve unchanged surrounding text, formulas, character-count cells, validations, and formatting.
- If rich text is required, patch the sheet XML with inline rich text runs or use an Excel automation method that preserves existing workbook formatting. Verify that unmodified text remains unmarked.

## Text And Image Rules

- For fields with stated character counts, write to roughly the stated count plus about 10 Japanese characters.
- User-provided text and images override web research.
- If images are not provided, source from official or reliable web pages and record the source.
- Do not use images with visible text overlays, black-and-white or clearly processed styling, people looking into the camera, or suspicious edits.
- Use only images that satisfy the workbook's required size and aspect ratio.
- Assign selected images numbered filenames such as `画像_01.jpg`, `画像_02.jpg`, or the workbook's requested image names, then enter those names in the relevant cells.
- 店舗情報画像_01 and 店舗情報画像_02 must match their nearby copy. For example, do not pair an interior description with a food-only image.
- Put only the selected/used images into the delivery image folder.

## 簡易HPS New Rules

Use these rules when the user asks for 自社簡易HP_S/簡易HPS入稿. 新規旅色入稿 and HPS template-change requests are different workflows, so do not combine their prompt templates.

- Preserve the HPS Excel template's formatting, dropdowns, row heights, fonts, sizes, formulas, and validations exactly.
- Follow stated character counts as closely as possible.
- When HPS is submitted with a new Tabiiro listing, use the same researched facts, images, and positioning, but write the HPS copy from the shop/facility's own point of view.
- Base color is specified by the user. If the color is white or black, choose it from the dropdown. For other colors, write the provided color code in the cell to the right of the theme-color label.
- If no logo is specified, get the logo image from the web, save it with the image name `ロゴ`, and enter the same name in the sheet.
- For メイン画像1 through メイン画像10, use specified images first. If unspecified, source suitable images from the web, respecting the size and aspect-ratio requirements. Crop if needed. It is not necessary to fill all 10 slots. If a clean compliant image is unavailable, leave that slot unused.
- For each メイン画像 slot, set the dropdown cell to the left of the image-name cell to `有` when used and `無` when unused.
- NEWS1 and TOPIX should announce that the website has opened. Use the logo or another suitable image.
- For レコメンド, unless specified, pick one store/facility strength or commitment similar to a new Tabiiro submission and write it. The image must match the content.
- 背景画像 should use an interior or exterior image.
- 電話ボタン defaults: background color black, button color orange, text color white, unless specified.
- SEO keywords, title, and description: unless specified, choose keywords based on likelihood of ranking in positions 1 through 10, not search volume alone. Avoid keywords where the top results are dominated by corporate sites and large media sites. Then write the title and description to match the selected keywords.
- SNS and external links: add links when accounts or relevant external pages exist.
- 取得希望ドメイン: if unspecified, choose the most appropriate domain candidate.
- For HPS submitted with a new Tabiiro listing, use the same case folder and include HPS in the new-production email. Shared images may use the same filenames, but image folders must be separated into `旅色用` and `HP用`.

## 簡易HPS Template Change Rules

Use these rules when the user asks for 自社簡易HPテンプレ変更, 簡易HPC→簡易HPS, or テンプレ変更制作依頼.

- The input rules are mostly the same as new HPS production, but do not fill the SEO section.
- Text and images are likely to be specified by the user. Use the specified content first.
- Do not use the new Tabiiro production prompt for template changes. Use the dedicated template-change prompt and email format.
- If the template-change direction differs from `簡易HPC→簡易HPS`, use the direction specified by the user in the email body.

## Output File Naming

The generated入稿シート filename must show the plan without ABC and the案件名:

```text
入稿シート_<プランABCなし>_【<案件名>】.xlsx
```

Example:

```text
入稿シート_TG5_【丹波茶屋ゆらり】.xlsx
```

## Box Storage

Default Box Drive base path:

```text
C:\Users\NX023066\Box\01_【公式】契約クライアント\★入稿フォルダ【全支店】\05_大阪
```

Plan category map:

- TG: `002_飲食`
- TO: `001_お取り寄せ`
- TY: `003_宿`
- TL: `004_レジャー`

Folder flow:

1. Go to the category folder.
2. Use or create `yyyymmdd_エリア`.
3. Use or create `新規`, `更新`, or `季節更新`.
4. Use or create the plan folder without ABC, such as `TG5` or `TL5`.
5. Inside it, create `mmdd_<プランABCあり>_<案件名>`.
6. Put the generated Excel file in that folder.
7. Create `画像` inside it and put the selected/used image files there.

Example:

```text
...\002_飲食\20260625_京都\新規\TG5\0625_TG5A_丹波茶屋ゆらり
```

For 更新, use the `更新` folder at the same level as `新規`. Create it if missing.

For the入稿メール, use a Web Box URL such as `https://app.box.com/folder/<folder-id>`, not the local Box Drive path. When possible, find the folder ID from Box Drive metadata. If the Web URL cannot be discovered, report the local path and ask the user to copy the Box web link.

## 入稿メール

Draft the email whenever an入稿指示 is handled.

To:

```text
入稿 <nyukou@brangista.com>
```

Cc:

```text
升本光典 <mitsunori_masumoto@brangista.com>, 寺内功次 <koji_terauchi@brangista.com>, 加藤安耶 <aya_kato@brangista.com>, 長杉菜月 <natsuki_nagasugi@brangista.com>, 長尾茜 <akane_nagao@brangista.com>, 大歳悠乃 <yuno_otoshi@brangista.com>, 谷内理彩 <risa_taniuchi@brangista.com>, 山口菜美 <nami_yamaguchi@brangista.com>, 岡田芽生 <mei_okada@brangista.com>
```

For 新規, use this subject format:

```text
mmdd【エリア_<プランABCあり>_新規制作依頼】案件名
```

If HPS is included:

```text
mmdd【エリア_<プランABCあり>_HPS_新規制作依頼】案件名
```

Do not put the HP掲載開始日 in the subject.

For 新規, use this body format:

```text
お疲れ様です。

表題案件の新規制作をお願いいたします。

クライアント名：<案件名>

掲載プラン：<プランABCあり>

掲載開始日：<mmdd>

ファイル保管場所：<Box Web URL>

以上、よろしくお願いいたします。
```

If HPS is included, use slash-separated plan and dates:

```text
掲載プラン：TG5A/HPS
掲載開始日：0625/0727
```

For 簡易HPSテンプレ変更, use this subject format:

```text
mmdd【自社簡易HP_テンプレ変更制作依頼】案件名
```

テンプレ変更 To:

```text
入稿 <nyukou@brangista.com>
```

テンプレ変更 Cc:

```text
BM開発制作課（WEBフォロー） <s_support@brangista.com>, 升本光典 <mitsunori_masumoto@brangista.com>, 加藤安耶 <aya_kato@brangista.com>
```

テンプレ変更 body format:

```text
お疲れ様です。

表題案件について、簡易HPC→簡易HPS のテンプレ変更をお願いします。

・クライアント：<案件名>

・掲載開始日：<M月D日>

・ファイル保管場所：<Box Web URL>
```

For 更新, use this subject format:

```text
mmdd【エリア_<プランABCあり>_In無_契約更新制作依頼】案件名
```

If the user specifies `In有` or another inclusion flag, use that value instead of `In無`. Keep the plan with ABC in the subject.

For 更新, list the corrected item names in the body under `【修正箇所】`. Use the item names supplied by the user, such as `LP用画像2～4` or `おすすめポイント①画像`.

更新 body format:

```text
お疲れ様です。
下記の案件の更新制作をお願いします。

【修正箇所】
・<修正箇所1>
・<修正箇所2>
-----------------------------------------------------------------------------------
・クライアント：<案件名>
・掲載プラン:<プランABCあり>_In無
・掲載日：<mmdd>
・掲載エリア：<エリア>
・ファイル保管場所：<Box Web URL>

宜しくお願いします。
```

## Date Rules

- Main publication date: requested month/day. If the user only gives a month, use the 25th.
- If that date falls on Saturday, Sunday, or a Japanese national holiday, use the next business day.
- HPS publication date: the next month's 25th after the main publication date, also adjusted to the next business day.
- Format date folders as `yyyymmdd`, and email/case folder dates as `mmdd`.

Use a reliable holiday source or known Japanese holiday calendar for the relevant year before finalizing dates.

## Validation Checklist

Before responding final:

- Correct template and plan were used.
- Output filename uses plan without ABC and案件名.
- Email body and case folder use plan with ABC.
- Box plan folder uses plan without ABC.
- 更新 uses the `更新` Box folder, not `新規`.
- 更新 email uses `契約更新制作依頼` and includes `【修正箇所】`.
- 更新 sheet marks corrected cells yellow and corrected text red; rich text cells mark only changed words, and deletions are red with strikethrough.
- HPS sheets preserve the template and separate `旅色用` and `HP用` image folders when submitted with a new Tabiiro listing.
- HPS template-change requests do not fill SEO and use the dedicated template-change email.
- Publication dates follow the 25th/next-business-day rule.
- Workbook styles, dimensions, dropdowns, and validations are preserved.
- `選択してください` placeholders are not polluted with `センタク`.
- Required web-researched fields are filled or clearly noted as unresolved.
- Images are compliant, numbered, copied into `画像`, and referenced in the workbook.
- Box Web URL is included in the email draft when discoverable.

## User Prompt Templates

Full prompt:

```text
$tabiiro-submission-workflow を使って、以下の入稿を一括作成してください。
案件名：
プラン：
区分：新規
エリア：
公開月または公開日：
HPS入稿：あり／なし
HPS入稿がある場合：
・HP掲載開始日：
・ベース色：
・ロゴ指定：
・メイン画像指定：
・NEWS/TOPIX指定：
・レコメンド指定：
・背景画像指定：
・電話ボタン色指定：
・SEO指定：
・SNS/外部リンク指定：
・取得希望ドメイン指定：
指定画像：
指定文章：
補足：
```

Short prompt:

```text
$tabiiro-submission-workflow：案件名「丹波茶屋ゆらり」、プランTG5C、エリア京都、公開6月、新規、HPS入稿なし。指定画像は添付優先で、入稿シート作成・Box格納・入稿メール案までお願いします。
```

HPS同時入稿 prompt:

```text
$tabiiro-submission-workflow：案件名「丹波茶屋ゆらり」、プランTG5C、エリア京都、公開6月、新規、HPS入稿あり。HPSのベース色は黒、ロゴ・メイン画像・SEO・希望ドメインは指定なしなので最適に選定してください。旅色用とHP用の画像フォルダを分け、入稿シート作成・Box格納・新規制作依頼メール案までお願いします。
```

HPSテンプレ変更 prompt:

```text
$tabiiro-submission-workflow：案件名「Over the Over」、掲載開始日7月27日、簡易HPC→簡易HPS のテンプレ変更制作依頼です。SEOは記載不要。指定文章・指定画像を優先し、Box格納とテンプレ変更依頼メール案までお願いします。
```

更新 prompt:

```text
$tabiiro-submission-workflow：案件名「鮨ふみ」、プランTG3A、エリア大阪、掲載日0525、区分更新、In無。修正箇所は「LP用画像2～4」「おすすめポイント①画像」です。指定画像は上書き対象の画像名に合わせ、修正セルは黄色、修正文字は赤で、Box格納と更新入稿メール案までお願いします。
```
