import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import win32com.client
import os

ppt = win32com.client.Dispatch("PowerPoint.Application")
file_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\epice_エピス_TG提案書_v10.pptx")
pres = ppt.Presentations.Open(file_path, WithWindow=False)

print(f"Total slides: {pres.Slides.Count}")

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
    # Print slide number and full text if it's a theme/ranking/SEO slide (1 to 14)
    if i <= 14:
        print(f"--- Slide {i:2d} ---")
        print(slide_text)

pres.Close()
ppt.Quit()
