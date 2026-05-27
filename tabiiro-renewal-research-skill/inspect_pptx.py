import win32com.client
import os

def walk_shapes(shapes, slide_num, prefix=""):
    for j, shape in enumerate(shapes):
        try:
            if shape.Type == 6: # msoGroup
                walk_shapes(shape.GroupItems, slide_num, prefix + f"Group {shape.Name} > ")
            elif shape.HasTextFrame:
                if shape.TextFrame.HasText:
                    text = shape.TextFrame.TextRange.Text
                    if "〇〇〇〇〇〇〇〇様" in text:
                        print(f"Slide {slide_num}: {prefix}Shape '{shape.Name}' contains '〇〇〇〇〇〇〇〇様'")
                    elif "画像をコピペしてください" in text:
                        print(f"Slide {slide_num}: {prefix}Shape '{shape.Name}' contains '画像をコピペしてください'")
            elif shape.Type == 14: # msoPlaceholder
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    text = shape.TextFrame.TextRange.Text
                    if "〇〇〇〇〇〇〇〇様" in text:
                        print(f"Slide {slide_num}: {prefix}Shape '{shape.Name}' contains '〇〇〇〇〇〇〇〇様'")
        except Exception as e:
            pass

def inspect():
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    file_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\施設専用資料テンプレ（TG）260316.pptx")
    
    print(f"Opening: {file_path}")
    pres = ppt.Presentations.Open(file_path, WithWindow=False)
    
    for i, slide in enumerate(pres.Slides):
        walk_shapes(slide.Shapes, i+1)
            
    pres.Close()
    ppt.Quit()

if __name__ == "__main__":
    import sys
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    inspect()
