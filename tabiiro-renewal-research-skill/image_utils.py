import win32com.client
import os

def insert_image_preserve_aspect(slide, img_path, target_shape):
    # Delete the target shape (the gray rectangle)
    left = target_shape.Left
    top = target_shape.Top
    width = target_shape.Width
    height = target_shape.Height
    target_shape.Delete()
    
    # Add picture at original size
    pic = slide.Shapes.AddPicture(img_path, False, True, left, top, -1, -1)
    
    # Calculate scale to fit within the width/height of the original gray rectangle
    # We want it to be as wide as the gray rectangle, preserving aspect ratio
    # If height exceeds, maybe scale to fit height instead
    scale_w = width / pic.Width
    scale_h = height / pic.Height
    
    # Usually we want it to fill the width or fit inside. The user said: "縦横比は元のサムネイル画像のままでいいので貼り付けてください"
    # This implies we can just make it fit the width and let the height be whatever, or fit inside the box.
    # Let's fit it inside the bounding box (aspect fit).
    scale = min(scale_w, scale_h)
    
    pic.Width = pic.Width * scale
    pic.Height = pic.Height * scale
    
    # Center it in the original bounding box
    pic.Left = left + (width - pic.Width) / 2
    pic.Top = top + (height - pic.Height) / 2
    return pic

if __name__ == "__main__":
    pass
