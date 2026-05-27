import win32com.client
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def walk_shapes(shapes, slide_num, prefix=""):
    for j, shape in enumerate(shapes):
        try:
            print(f"Slide {slide_num}: {prefix}Shape '{shape.Name}' (Type {shape.Type})")
            if shape.Type == 6: # msoGroup
                walk_shapes(shape.GroupItems, slide_num, prefix + f"Group {shape.Name} > ")
        except Exception as e:
            pass

def inspect_all():
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    file_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\施設専用資料テンプレ（TG）260316.pptx")
    
    print(f"Opening: {file_path}")
    pres = ppt.Presentations.Open(file_path, WithWindow=False)
    
    for i in range(3, 6): # Slides 3 to 5
        try:
            walk_shapes(pres.Slides(i).Shapes, i)
        except Exception as e:
            pass
            
    pres.Close()
    ppt.Quit()

if __name__ == "__main__":
    inspect_all()
