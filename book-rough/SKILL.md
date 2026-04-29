---
name: book-rough
description: Create BOOKラフ proposal PowerPoint slides by filling a supplied BOOKラフ .ppt/.pptx template with a restaurant name and three researched store images. Use when the user says BOOKラフ, BOOK rough, ブックラフ, asks to complete the attached PowerPoint with 店名 and photos, or asks to make a Tabiiro/旅色-style BOOKラフ proposal slide from a restaurant/shop name.
---

# BOOKラフ

## Overview

Use this skill to complete a supplied BOOKラフ PowerPoint template for one or more restaurants. Preserve the reference slide format: one large right-side photo behind the template, two overlapping upper-left photos above the template, and a transparent white restaurant-name text box.

If a PowerPoint file is involved, also follow the available presentation/PPT workflow for rendering and QA.

## Required Inputs

- A BOOKラフ `.ppt` or `.pptx` template. Prefer the reference deck that contains several completed examples.
- Restaurant/shop name(s).
- Optional: a user-specified output folder or preferred source pages. If omitted, create a nearby working folder and research sources yourself.

## Research Rules

Search the web for each store name. Prefer official or semi-official sources in this order:

1. Official website or official store page
2. Official Instagram
3. HotPepper, Tabelog, Gurunavi, Retty, Kiss PRESS, or similar listing pages

Select three real store images:

- Right large image: the most visually strong image showing the store's main appeal or craft. Usually a signature dish, course spread, meat/seafood dish, or visually rich plate.
- Upper-left back image: a different genre, usually exterior, signage, interior, or atmosphere.
- Upper-left front image: another different genre, usually drink, counter/interior detail, dish close-up, or sake/alcohol.

Avoid choosing three images of the same category. If only one source has usable photos, still vary the subject as much as possible. In the final response, mention the source pages used.

## Placement Standard

Use the reference-example layout from the corrected `v2` format:

| Element | Layer | Left | Top | Width | Height | Notes |
|---|---:|---:|---:|---:|---:|---|
| Large right image | bottom | 287.3 | 64.8 | 398.6 | 385.5 | Send behind the template group. |
| Template/group | above right image | keep existing | keep existing | keep existing | keep existing | Do not flatten or rebuild. |
| Upper-left back image | above template | 66.1 | 88.3 | 162.3 | 91.0 | Usually exterior/sign/interior. |
| Upper-left front image | top of the two | 152.8 | 156.4 | 130.6 | 73.3 | Must overlap and sit above the back image. |
| Store name | top | 66.1 | 263.0 | auto | auto | Transparent textbox, white text. |

Title styling:

- Text: exact store name supplied by the user.
- Fill/background: transparent. Do not add a black rectangle behind the name.
- Color: white.
- Font: preserve the reference if editable; otherwise use `HG明朝E`.
- Size: use the reference size, normally 28 pt.
- Do not cover the existing dummy title with a filled rectangle. Instead use a reference slide whose grouped title field is empty, or place a transparent text box in the correct position.

## PowerPoint Workflow

1. Inspect the supplied deck by exporting PNGs and listing shapes.
2. Choose the best base slide:
   - If the deck has completed examples, use the final/reference slide with the empty grouped title field if available.
   - Keep the original template group.
   - Remove example foreground photos and example title text.
3. Download and crop the three selected images to fit the target aspect ratios.
4. Place the large right photo first and send it to the bottom layer.
5. Place the upper-left back photo, then the upper-left front photo so the front photo sits above it.
6. Add the transparent white store-name text.
7. Save as `.ppt` when the source is an old `.ppt`; avoid forced `.pptx` conversion unless it reopens cleanly.
8. Reopen the saved file and export a PNG preview from the saved presentation itself.
9. Check:
   - Right image is behind the template and partly covered by the dark/gradient book area.
   - Upper-left images match the reference size and overlap correctly.
   - Store name has no black background and remains white.
   - Font and size resemble the reference examples.
   - The saved file reopens without corruption.

## Helper Script

Use `scripts/build-book-rough.ps1` after selecting/cropping images. It assumes the supplied deck contains a suitable reference slide and will keep only the selected base slide.

Example:

```powershell
$scriptPath = "$env:USERPROFILE\.codex\skills\book-rough\scripts\build-book-rough.ps1"
$buildBookRough = [scriptblock]::Create((Get-Content -LiteralPath $scriptPath -Encoding UTF8 -Raw))
& $buildBookRough `
  -TemplatePath "C:\path\BOOKラフ.ppt" `
  -OutputPath "C:\path\BOOKラフ_店名_完成.ppt" `
  -ShopName "飯酒 ゆき常" `
  -RightImagePath "C:\path\right.jpg" `
  -LeftTopImagePath "C:\path\left_top.jpg" `
  -LeftLowerImagePath "C:\path\left_lower.jpg" `
  -BaseSlideIndex 5 `
  -RenderDir "C:\path\preview"
```

Use the scriptblock form above on managed Windows machines where direct `.ps1` execution is blocked.

If the source template is only a flattened screenshot with no editable group, inspect it manually and prefer requesting or using the reference-example deck. A flattened one-slide template cannot truly place the right image below the template layer.

## Final Response

Keep the response short and include:

- Completed PowerPoint path.
- Preview PNG path.
- Whether the saved file was reopened and visually checked.
- Image source pages used.
