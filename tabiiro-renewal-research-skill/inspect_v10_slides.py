import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import win32com.client
import os

ppt = win32com.client.Dispatch("PowerPoint.Application")
file_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\epice_エピス_TG提案書_v10.pptx")
pres = ppt.Presentations.Open(file_path, WithWindow=False)

print(f"Total slides in epice v10 presentation: {pres.Slides.Count}")

for i in range(1, pres.Slides.Count + 1):
    slide = pres.Slides(i)
    
    texts = []
    def extract_text(shapes):
        for shape in shapes:
            try:
                if shape.Type == 6:
                    extract_text(shape.GroupItems)
                elif shape.HasTextFrame and shape.TextFrame.HasText:
                    texts.append(shape.TextFrame.TextRange.Text)
            except Exception:
                pass
    extract_text(slide.Shapes)
    
    slide_text = "\n".join(texts).strip()
    first_line = slide_text.split('\n')[0][:80] if slide_text else "(no text)"
    print(f"Slide {i:2d}: {first_line}")
    # Print lines that contain theme, ranking, pet, wine, or xmas
    for line in slide_text.split('\n'):
        line_s = line.strip()
        if any(w in line_s for w in ["ペット", "クリスマス", "ワイン", "女子会", "french", "kominka", "pet", "xmas", "wine", "josikai", "jyoshikai"]):
            print(f"    -> {line_s}")

pres.Close()
ppt.Quit()
