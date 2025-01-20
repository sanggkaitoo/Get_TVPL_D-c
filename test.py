import re
import feedparser
import datetime

print(datetime.now().strftime("%B"))


# url = "https://thuvienphapluat.vn/rss.xml"
# docs = feedparser.parse(url)

# print(docs.entries)

# def extract_numbers(title):
#     pattern = r'\b\d{1,7}/[\w-]+(?:/[\w-]+)?'
#     matches = re.findall(pattern, title)
#     return matches[0] if matches else ""

# print(extract_numbers("Số: 2229/QĐ-BTTTT ngày 19/12/2024 của Bộ Thông tin và Truyền thông"))

row = ["Temp", "BTTTT"]

acronym_name = {
    "NĐ"   : "Nghị định",
    "TT"   : "Thông tư",
    "CĐ"   : "Công điện",
    "QĐ"   : "Quyết định",
    "BTTTT": "Bộ Thông tin và Truyền thông",
    "TTg"  : "Thủ tướng Chính phủ",
    "CP"   : "Chính phủ",
    "BCT"  : "Bộ Công Thương"
}


print(acronym_name[row[1]])