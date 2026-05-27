import sys
content = open('facebook_history.py', 'r', encoding='utf-8').read()
content = content.replace('Search Tabiiro SNS sales-support Instagram workbooks', 'Search Tabiiro SNS sales-support Facebook workbooks')

old_include = """SHEET_INCLUDE_HINTS = (
    "sns",
    "ig",
    "instagram",
    "旅色ig",
    "旅色instagram",
    "お取り寄せ",
    "台湾",
)"""
new_include = """SHEET_INCLUDE_HINTS = (
    "fb",
    "facebook",
    "x・fb",
    "台湾sns",
    "sns",
    "台湾",
)"""
content = content.replace(old_include, new_include)

old_exclude = """SHEET_EXCLUDE_HINTS = (
    "fb",
    "tw",
    "line",
    "施策整理",
    "テーマ募集",
    "ルール",
    "参照",
    "フィードバック",
)"""
new_exclude = """SHEET_EXCLUDE_HINTS = (
    "ig",
    "instagram",
    "旅色ig",
    "line",
    "施策整理",
    "テーマ募集",
    "ルール",
    "参照",
    "フィードバック",
)"""
content = content.replace(old_exclude, new_exclude)

old_guess = """def account_guess(sheet_title: str) -> str:
    title_norm = norm(sheet_title)
    if "お取り寄せ" in title_norm:
        return "@tabiiro.otoriyose"
    if "台湾" in title_norm:
        return "@tabiiro_tw"
    if "近畿" in title_norm or "kinki" in title_norm:
        return "@tabiiro.kinki"
    if "ig" in title_norm or "instagram" in title_norm or "sns" in title_norm:
        return "@tabiiro"
    return "unknown\""""

new_guess = """def account_guess(sheet_title: str) -> str:
    title_norm = norm(sheet_title)
    if "台湾" in title_norm:
        return "@tabiiro_tw_fb"
    if "fb" in title_norm or "facebook" in title_norm:
        return "@tabiiro_fb"
    return "unknown_fb\""""
content = content.replace(old_guess, new_guess)

open('facebook_history.py', 'w', encoding='utf-8').write(content)
print('Replaced')
