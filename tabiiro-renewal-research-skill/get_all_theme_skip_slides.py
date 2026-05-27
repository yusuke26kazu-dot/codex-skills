import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import win32com.client
import os

ppt = win32com.client.Dispatch("PowerPoint.Application")
file_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\施設専用資料テンプレ（TG）260316.pptx")
pres = ppt.Presentations.Open(file_path, WithWindow=False)

for i in range(7, 15):
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
    print(f"--- Slide {i} ---")
    print(slide_text)

pres.Close()
ppt.Quit()
