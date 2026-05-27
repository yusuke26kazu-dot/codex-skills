import struct
import pathlib

def get_image_info(filepath):
    """Extract JPEG width, height and basic metadata without external libraries."""
    with open(filepath, 'rb') as f:
        data = f.read()
    
    # Check JPEG signature
    if data[:2] != b'\xff\xd8':
        return None, None, "Not a valid JPEG"
    
    # Parse segments to find SOF (Start of Frame)
    size = len(data)
    idx = 2
    width, height = None, None
    while idx < size:
        if data[idx] == 0xff:
            marker = data[idx+1]
            if marker in (0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf):
                # SOF marker found
                length = struct.unpack('>H', data[idx+2:idx+4])[0]
                # SOF format: [marker, length(2), precision(1), height(2), width(2), components(1)]
                height, width = struct.unpack('>HH', data[idx+5:idx+9])
                break
            else:
                idx += 2
                # Skip length of this segment
                length = struct.unpack('>H', data[idx:idx+2])[0]
                idx += length
        else:
            idx += 1
            
    # Try to extract any printable strings in the first 2000 bytes (common place for comments/titles)
    printable_strings = []
    chunk = data[:4000]
    # Simple search for ASCII words
    word = []
    for b in chunk:
        if 32 <= b <= 126 or b in (10, 13):
            word.append(chr(b))
        else:
            if len(word) > 10:
                printable_strings.append(''.join(word).strip())
            word = []
    if len(word) > 10:
        printable_strings.append(''.join(word).strip())

    return width, height, printable_strings

src_dir = pathlib.Path(r'C:\Users\NX023066\Downloads\新しいフォルダー')
image_files = sorted(list(src_dir.glob('*.jpg')))

print("=== Image Analysis (Pure Python) ===")
for f in image_files:
    w, h, text = get_image_info(f)
    orientation = "Landscape (横)" if w and h and w > h else "Portrait (縦)" if w and h and h > w else "Square (正方形)"
    print(f"\nFile: {f.name}")
    print(f"Dimensions: {w} x {h} ({orientation})")
    # Clean up printed strings
    clean_texts = [t for t in text if any(keyword in t.lower() for keyword in ['exif', 'adobe', 'photoshop', 'nikon', 'canon', 'sony', 'iphone', 'icc', 'desc', 'title'])]
    if clean_texts:
        print(f"Metadata Snippets: {clean_texts[:3]}")
