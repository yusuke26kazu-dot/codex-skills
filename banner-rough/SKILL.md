---
name: banner-rough
description: Create Japanese website banner rough PowerPoint mockups by inserting a Tabiiro/旅色 actress banner into a real website screenshot. Use when the user says バナーラフ, 女優バナー, asks to place a banner on a URL screenshot, or wants a PowerPoint rough showing a banner naturally inserted into an existing HP/site page.
---

# バナーラフ

## Purpose

Create a PowerPoint rough that shows how a supplied or known `旅色で紹介されました / Cover Woman` banner would look inside an actual website page.

If a PowerPoint file is involved, also follow the available presentation/PPT workflow for rendering and QA.

## Required Inputs

- Website URL.
- Banner reference, if not already known.
  - Prefer a supplied reference PowerPoint whose third slide contains the banner image.
  - If no banner is supplied, use the local reference banner when available:
    `C:\Users\NX023066\Documents\New project 3\actress_banner_reference\extracted_banners\image8.jpg`
  - If no local reference exists, ask for the banner PPT/image.

Optional:
- User-specified insertion area, spacing preference, or output folder.

## Banner Selection

- For normal website insertion, use the horizontal banner.
- When the reference PPT contains multiple banners, use the horizontal banner from slide 3 unless the user explicitly requests another size.
- Keep the banner as a finished image. Do not rebuild its text in PowerPoint.

## Screenshot Rules

1. Capture the instructed URL as a real browser screenshot.
2. Scroll through the page before capture when lazy-loaded content may be blank.
3. Prefer a full-page or stitched screenshot that includes content below the intended banner location.
4. Do not stop at a viewport that has only a blank/white/beige lower area.
5. If a PDF/print capture distorts the site, use browser screenshots instead.

## Insertion Layout

The rough must look like the banner has been inserted into the real HP, not pasted onto a blank area.

Use this structure:

1. Upper HP screenshot content.
2. Banner.
3. Lower HP screenshot content continuing below the banner.

Practical method:

- Cut the full screenshot into a top part and bottom part.
- Remove excessive blank space between the top section and the next real section.
- Add only enough slot height for the banner and comfortable breathing room.
- Place the bottom screenshot immediately after the banner slot so the original HP content continues.
- Keep the top screenshot, banner, and bottom screenshot as separate PowerPoint image parts whenever possible so placement can be adjusted later.

## Spacing Rules

Be especially careful with vertical whitespace.

- HP content should generally remain full-width unless the user asks otherwise.
- Do not create a large beige/white gap between the previous section's fade-out and the banner.
- Do not create a large beige/white gap between the banner and the next HP heading/content.
- If the site has a hero image that fades out, place the banner soon after the fade-out point.
- Place the next real HP section close enough below the banner that it feels like a natural in-page insertion.
- Avoid ending the visible slide area with only blank background below the banner.

For the current reference style, the accepted visual direction is:

- Full-width HP screenshot.
- A large horizontal banner centered below the hero/fade-out.
- The next heading, such as `ゆう菜について`, appearing soon below the banner.
- Real lower HP content visible underneath.

## Output

- Create a portrait PowerPoint rough matching the reference deck's long vertical proposal feel.
- The first slide should be the finished banner rough.
- Save a PNG preview rendered from the saved PPTX, not just an intermediate composite image.

## QA Checklist

Before final response, visually check:

- The HP screenshot width is not unintentionally narrowed.
- The banner is not just floating in a blank lower area.
- The lower HP content appears below the banner.
- The vertical beige/white spacing above and below the banner is not excessive.
- The banner is centered and not covering important page content.
- The saved PPTX opens as a valid package and has the intended visible slide count.

## Call Prompt

Use this form:

```text
バナーラフを作ってください。
URL：https://example.com/
バナー：既定の女優バナー
余白指定：なし
```

If a different banner is needed, attach the reference PPT/image and replace `既定の女優バナー` with `添付バナー`.

## Final Response

Keep the final response short. Include:

- Completed PowerPoint path.
- Preview PNG path.
- Whether the PPTX package and preview were checked.
