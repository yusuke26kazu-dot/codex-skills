import win32com.client
import os
import sys

def verify_pres(pres, out_path, stage_name):
    # Save a copy
    try:
        if os.path.exists(out_path):
            os.remove(out_path)
        pres.SaveCopyAs(out_path)
        # Check if we can open it
        ppt2 = win32com.client.Dispatch("PowerPoint.Application")
        try:
            pres2 = ppt2.Presentations.Open(out_path, WithWindow=False)
            slide_count = pres2.Slides.Count
            pres2.Close()
            print(f"[OK] Stage: {stage_name} (Slides count: {slide_count})")
            return True
        except Exception as e:
            print(f"[FAIL] Stage: {stage_name} - Corrupted when attempting to reopen! Error: {e}")
            return False
    except Exception as e:
        print(f"[FAIL] Stage: {stage_name} - SaveCopyAs failed! Error: {e}")
        return False

def test():
    template_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\施設専用資料テンプレ（TG）260316.pptx")
    out_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\debug_test.pptx")
    
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    pres = ppt.Presentations.Open(template_path, WithWindow=False)
    
    # 1. Text replace
    for shape in pres.Slides(1).Shapes:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            if "〇〇〇〇〇〇〇〇様" in shape.TextFrame.TextRange.Text:
                shape.TextFrame.TextRange.Replace("〇〇〇〇〇〇〇〇様", "epice エピス様")
    if not verify_pres(pres, out_path, "Text replace Slide 1"):
        pres.Close()
        ppt.Quit()
        return
        
    # Deleting "※印刷しない" slides
    print("Testing delete loops...")
    for i in range(pres.Slides.Count, 0, -1):
        slide = pres.Slides(i)
        text_content = ""
        def extract_text(shapes):
            nonlocal text_content
            for shape in shapes:
                if shape.Type == 6:
                    extract_text(shape.GroupItems)
                elif shape.HasTextFrame and shape.TextFrame.HasText:
                    text_content += shape.TextFrame.TextRange.Text + "\n"
        extract_text(slide.Shapes)
        
        if "※印刷しない" in text_content:
            slide.Delete()
            if not verify_pres(pres, out_path, f"Delete ※印刷しない at slide {i}"):
                pres.Close()
                ppt.Quit()
                return
            continue

        if "御社公式ホームページ" in text_content:
            slide.Delete()
            if not verify_pres(pres, out_path, f"Delete 御社公式ホームページ at slide {i}"):
                pres.Close()
                ppt.Quit()
                return
            continue
            
        if "女優バナー" in text_content or "旅色女優バナー" in text_content:
            slide.Delete()
            if not verify_pres(pres, out_path, f"Delete 旅色女優バナー at slide {i}"):
                pres.Close()
                ppt.Quit()
                return
            continue

        if "繁体字版旅色" in text_content and "Facebook" not in text_content:
            slide.Delete()
            if not verify_pres(pres, out_path, f"Delete 繁体字版旅色 at slide {i}"):
                pres.Close()
                ppt.Quit()
                return
            continue

    pres.Close()
    ppt.Quit()
    print("Test finished successfully!")

if __name__ == '__main__':
    test()
