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

def replace_text_in_shapes(shapes, old_text, new_text):
    for shape in shapes:
        if shape.Type == 6: # msoGroup
            replace_text_in_shapes(shape.GroupItems, old_text, new_text)
        elif shape.HasTextFrame and shape.TextFrame.HasText:
            text = shape.TextFrame.TextRange.Text
            if old_text in text:
                shape.TextFrame.TextRange.Replace(old_text, new_text)

def insert_image_preserve_aspect(slide, img_path, target_shape):
    left = target_shape.Left
    top = target_shape.Top
    target_width = target_shape.Width
    target_shape.Delete()
    
    # Add picture at original size
    pic = slide.Shapes.AddPicture(img_path, False, True, left, top, -1, -1)
    
    # Scale to fit the original width, preserving aspect ratio
    scale = target_width / pic.Width
    pic.Width = pic.Width * scale
    pic.Height = pic.Height * scale
    
    return pic

def find_gray_rectangle(shapes):
    for shape in shapes:
        if shape.Type == 6: # msoGroup
            found = find_gray_rectangle(shape.GroupItems)
            if found: return found
        elif shape.HasTextFrame and shape.TextFrame.HasText:
            text = shape.TextFrame.TextRange.Text
            if "画像をコピペしてください" in text:
                return shape
    return None

def process_presentation():
    # 1. Download images
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
            # Crop top and bottom borders (green/black margins)
            # OGP is 1200x1200, center image is usually 1200x630
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
        download_og_image(url, os.path.join(img_dir, f"seo_{i}.jpg"))
        
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    file_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\施設専用資料テンプレ（TG）260316.pptx")
    pres = ppt.Presentations.Open(file_path, WithWindow=False)

    # 2. Slide 1: replace text
    replace_text_in_shapes(pres.Slides(1).Shapes, "〇〇〇〇〇〇〇〇様", "epice エピス様")
    
    # 3. Super Theme Features (Slide 2)
    # We will duplicate Slide 2 for each super theme, starting from the back to keep indices sane.
    base_slide2_index = 2
    for i in reversed(range(len(super_themes))):
        name, url = super_themes[i]
        new_slide = pres.Slides(base_slide2_index).Duplicate()
        # The duplicate is placed right after the original slide.
        # But wait, Duplicate() returns a SlideRange. The slide is newly added.
        # It's better to duplicate and then process the duplicated slide.
        slide_obj = new_slide.Item(1)
        gray_shape = find_gray_rectangle(slide_obj.Shapes)
        if gray_shape:
            # Add image without stretching and delete gray rectangle
            img_path = os.path.join(img_dir, f"st_{i}.jpg")
            if os.path.exists(img_path):
                insert_image_preserve_aspect(slide_obj, img_path, gray_shape)

    # Delete original Slide 2
    pres.Slides(base_slide2_index).Delete()
    
    # 4. Process SEO slides and Delete "※印刷しない" / Theme Feature slides
    # We iterate in reverse to avoid index shifting problems.
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
        
        # Mark for deletion if DO NOT PRINT
        if "※印刷しない" in text_content:
            slide.Delete()
            continue
            
        # Delete Theme feature slide (it has "テーマ特集" but maybe not "スーパーテーマ特集")
        if "テーマ特集への掲載のご案内" in text_content and "スーパー" not in text_content:
            slide.Delete()
            continue
            
        # SEO slide processing
        if "Google検索にて" in text_content and "〇〇〇 〇〇〇" in text_content:
            for kw, url, rank, views in reversed(seo_articles):
                new_slide = slide.Duplicate().Item(1)
                
                replace_text_in_shapes(new_slide.Shapes, "〇〇〇 〇〇〇", kw)
                replace_text_in_shapes(new_slide.Shapes, "●位", rank)
                replace_text_in_shapes(new_slide.Shapes, "●●●●回", views)
                
                gray_shape = find_gray_rectangle(new_slide.Shapes)
                if gray_shape:
                    img_path = os.path.join(img_dir, f"seo_{seo_articles.index((kw, url, rank, views))}.jpg")
                    if os.path.exists(img_path):
                        insert_image_preserve_aspect(new_slide, img_path, gray_shape)
            slide.Delete()
            continue
        
    out_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\epice_エピス_TG提案書.pptx")
    print(f"Saving to {out_path}")
    pres.SaveAs(out_path)
    pres.Close()
    ppt.Quit()

if __name__ == "__main__":
    process_presentation()
