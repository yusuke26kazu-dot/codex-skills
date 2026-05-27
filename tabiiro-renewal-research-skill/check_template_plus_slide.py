import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import win32com.client
import os

ppt = win32com.client.Dispatch("PowerPoint.Application")
file_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\施設専用資料テンプレ（TG）260316.pptx")
pres = ppt.Presentations.Open(file_path, WithWindow=False)

for i in range(1, pres.Slides.Count + 1):
    slide = pres.Slides(i)
    text_content = ""
    def extract_text(shapes):
        global text_content
        for shape in shapes:
            try:
                if shape.Type == 6:
                    extract_text(shape.GroupItems)
                elif shape.HasTextFrame and shape.TextFrame.HasText:
                    text_content += shape.TextFrame.TextRange.Text + "\n"
            except Exception:
                pass
    extract_text(slide.Shapes)
    if "プラス" in text_content or "plus" in text_content.lower():
        print(f"Slide {i}: Contains 'プラス'")
        print(f"  Text excerpt: {text_content[:300]}")

pres.Close()
ppt.Quit()
