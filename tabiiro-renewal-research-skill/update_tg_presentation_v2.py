import win32com.client
import os
import urllib.request
import re
from PIL import Image

def download_og_image(url, save_path):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        html = urllib.request.urlopen(req).read().decode('utf-8')
        match = re.search(r'<meta property="og:image" content="([^"]+)"', html)
        if match:
            img_url = match.group(1)
            print(f"Downloading {img_url} to {save_path}")
            req_img = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
            img_data = urllib.request.urlopen(req_img).read()
            with open(save_path, 'wb') as f:
                f.write(img_data)
            return True
    except Exception as e:
        print(f"Failed to fetch {url}: {e}")
    return False

def download_fv_image(url, save_path):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        html = urllib.request.urlopen(req).read().decode('utf-8')
        match = re.search(r'src="([^"]+_fv\.jpg[^"]*)"', html)
        if match:
            img_url = match.group(1)
            img_url = img_url.split('?')[0] + "?w=1600&h=900&mode=crop"
            print(f"Downloading FV {img_url} to {save_path}")
            req_img = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
            img_data = urllib.request.urlopen(req_img).read()
            with open(save_path, 'wb') as f:
                f.write(img_data)
            return True
    except Exception as e:
        print(f"Failed to fetch FV for {url}: {e}")
    return False

def replace_text_in_shapes(shapes, old_text, new_text):
    for shape in shapes:
        if shape.Type == 6: # msoGroup
            replace_text_in_shapes(shape.GroupItems, old_text, new_text)
        elif shape.HasTextFrame and shape.TextFrame.HasText:
            text = shape.TextFrame.TextRange.Text
            if old_text in text:
                shape.TextFrame.TextRange.Replace(old_text, new_text)

def find_gray_rectangle_group(shapes):
    """Finds the group containing '画像をコピペしてください'"""
    for shape in shapes:
        if shape.Type == 6: # msoGroup
            if find_gray_rectangle_group(shape.GroupItems):
                return shape
        elif shape.HasTextFrame and shape.TextFrame.HasText:
            text = shape.TextFrame.TextRange.Text
            if "画像をコピペしてください" in text:
                return shape
    return None

def insert_image_centered(slide, img_path, target_group, width_cm, height_cm):
    # Get the bounding box of the target group before deleting
    center_x = target_group.Left + target_group.Width / 2
    center_y = target_group.Top + target_group.Height / 2
    
    # Delete the entire group
    target_group.Delete()
    
    # 1 cm = 28.346 points
    target_width = width_cm * 28.346
    target_height = height_cm * 28.346
    
    # Add picture
    # MsoTriState: msoFalse = 0, msoTrue = -1
    # Adding picture with default size first
    pic = slide.Shapes.AddPicture(img_path, False, True, 0, 0, -1, -1)
    
    # Unlock aspect ratio so we can force exact dimensions
    pic.LockAspectRatio = 0
    pic.Width = target_width
    pic.Height = target_height
    
    # Center it at the original group's center
    pic.Left = center_x - pic.Width / 2
    pic.Top = center_y - pic.Height / 2
    
    return pic

def process_presentation():
    super_themes = [
        ("京都のフレンチランキング", "https://tabiiro.jp/gourmet/theme/french/ranking/kinki/kyoto/"),
        ("京都の古民家レストラン・カフェランキング", "https://tabiiro.jp/gourmet/theme/kominka-cafe/ranking/kinki/kyoto/"),
        ("近畿の高級店ランキング", "https://tabiiro.jp/gourmet/theme/high-class-restaurant/ranking/kinki/")
    ]
    seo_articles = [
        ("京都 ランチ おしゃれ", "https://tabiiro.jp/gourmet/article/kyoto-lunch-oshare/", "1位", "8389"),
        ("銀閣寺 周辺 グルメ", "https://tabiiro.jp/gourmet/article/ginkakuji-shuhen-gourmet/", "3位", "1308"),
        ("GW 京都 穴場 グルメ", "https://tabiiro.jp/gourmet/article/kyotoshinai-GW-anaba/", "1位", "284")
    ]
    
    img_dir = os.path.abspath("images")
    if not os.path.exists(img_dir):
        os.makedirs(img_dir)
        
    for i, (name, url) in enumerate(super_themes):
        path = os.path.join(img_dir, f"st_{i}.jpg")
        if download_og_image(url, path):
            try:
                img = Image.open(path)
                # Crop to center 1200x630
                top = (img.height - 630) / 2
                bottom = top + 630
                cropped = img.crop((0, top, img.width, bottom))
                cropped.convert('RGB').save(path)
            except Exception as e:
                print(f"Failed to crop {path}: {e}")
        
    for i, (kw, url, rank, views) in enumerate(seo_articles):
        download_fv_image(url, os.path.join(img_dir, f"seo_{i}.jpg"))
        
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    file_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\施設専用資料テンプレ（TG）260316.pptx")
    pres = ppt.Presentations.Open(file_path, WithWindow=False)

    replace_text_in_shapes(pres.Slides(1).Shapes, "〇〇〇〇〇〇〇〇様", "epice エピス様")
    
    base_slide2_index = 2
    for i in reversed(range(len(super_themes))):
        name, url = super_themes[i]
        new_slide = pres.Slides(base_slide2_index).Duplicate().Item(1)
        
        target_group = find_gray_rectangle_group(new_slide.Shapes)
        if target_group:
            try:
                if target_group.Type != 6 and hasattr(target_group, 'ParentGroup') and target_group.ParentGroup:
                    target_group = target_group.ParentGroup
            except Exception:
                pass
            
            img_path = os.path.join(img_dir, f"st_{i}.jpg")
            if os.path.exists(img_path):
                # Super theme: 10.77 cm x 19.15 cm
                insert_image_centered(new_slide, img_path, target_group, width_cm=19.15, height_cm=10.77)

    pres.Slides(base_slide2_index).Delete()
    
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
            continue
            
        if "テーマ特集への掲載のご案内" in text_content and "スーパー" not in text_content:
            slide.Delete()
            continue
            
        if "Google検索にて" in text_content and "〇〇〇 〇〇〇" in text_content:
            for kw, url, rank, views in reversed(seo_articles):
                new_slide = slide.Duplicate().Item(1)
                
                replace_text_in_shapes(new_slide.Shapes, "〇〇〇 〇〇〇", kw)
                replace_text_in_shapes(new_slide.Shapes, "●位", rank)
                replace_text_in_shapes(new_slide.Shapes, "●●●●回", views)
                
                target_group = find_gray_rectangle_group(new_slide.Shapes)
                if target_group:
                    try:
                        if target_group.Type != 6 and hasattr(target_group, 'ParentGroup') and target_group.ParentGroup:
                            target_group = target_group.ParentGroup
                    except Exception:
                        pass
                        
                    img_path = os.path.join(img_dir, f"seo_{seo_articles.index((kw, url, rank, views))}.jpg")
                    if os.path.exists(img_path):
                        # SEO article: 11.31 cm x 20.11 cm
                        insert_image_centered(new_slide, img_path, target_group, width_cm=20.11, height_cm=11.31)
            slide.Delete()
            continue
        
    out_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\epice_エピス_TG提案書_v3.pptx")
    print(f"Saving to {out_path}")
    pres.SaveAs(out_path)
    pres.Close()
    ppt.Quit()

if __name__ == "__main__":
    process_presentation()
