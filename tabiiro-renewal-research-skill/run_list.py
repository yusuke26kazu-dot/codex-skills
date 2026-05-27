import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import win32com.client
import os

ppt = win32com.client.Dispatch("PowerPoint.Application")
file_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\asakura_祇園あさくら_TG提案書_v11.pptx")
pres = ppt.Presentations.Open(file_path, WithWindow=False)

print(f"Total slides: {pres.Slides.Count}")

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
    first_line = text_content.strip().split('\n')[0][:60] if text_content.strip() else "(no text)"
    print(f"  Slide {i:2d}: {first_line}")

pres.Close()
ppt.Quit()
