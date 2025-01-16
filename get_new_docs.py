import feedparser
import re
import openpyxl
from datetime import datetime

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

ministry_name = {
    "VPQH" : "Văn phòng Quốc hội",
    "VPCTN" : "Văn phòng Chủ tịch nước",
    "VPCP" : "Văn phòng Chính phủ",
    "TANDTC" : "Tòa án nhân dân tối cao",
    "VKSNDTC" : "Viện Kiểm sát ND tối cao",
    "BNG" : "Bộ Ngoại giao",
    "BTP" : "Bộ Tư pháp",
    "KTNN" : "Kiểm toán Nhà nước",
    "BKHĐT" : "Bộ Kế hoạch và Đầu tư",
    "TTCP" : "Thanh tra Chính phủ",
    "BTTTT" : "Bộ Thông tin và truyền thông",
    "HLHPNVN" : "Hội LH Phụ nữ Việt Nam",
    "ĐTNCSHCM" : "TW Đoàn TN CS HCM",
    "MTTQ" : "UB TW MTTQ Việt Nam",
    "LMHTX" : "Liên minh HTX Việt Nam",
    "HND" : "Hội Nông dân Việt Nam",
    "HCCB" : "Hội Cựu chiến binh Việt Nam",
    "BNV" : "Bộ Nội vụ",
    "BTC" : "Bộ Tài chính",
    "BVHTTDL" : "Bộ Văn hóa- Thể thao- Du lịch",
    "BGDĐT" : "Bộ Giáo dục và Đào tạo",
    "QGHN" : "Đại học Quốc gia Hà nội",
    "QGHCM" : "Đại học Quốc gia TP. HCM",
    "BKHCN" : "Bộ Khoa học và Công nghệ",
    "TTXVN" : "Thông tấn xã Việt nam",
    "THVN" : "Đài Truyền hình Việt nam",
    "TNVN" : "Đài Tiếng nói Việt nam",
    "KHCNVN" : "Viện Khoa học và c.nghệ VN",
    "KHXHVN" : "Viện Khoa học xã hội ViệtNam",
    "BQLHL" : "BQL khu c.nghệ cao Hoà Lạc",
    "BQLVHDL" : "Ban QL làng văn hoá Du lịch",
    "BYT" : "Bộ Y tế",
    "LĐLĐVN" : "Tổng Liên đoàn LĐ Việt Nam",
    "BLĐTBXH" : "Bộ Lao động - TB&XH",
    "BHXHVN" : "Bảo hiểm xã hội Việt nam",
    "UBDT" : "Uỷ ban Dân tộc",
    "BNN" : "Bộ Nông nghiệp và PTNT",
    "BCT" : "Bộ Công thương",
    "BTNMT" : "Bộ Tài nguyên - Môi trường",
    "BXD" : "Bộ Xây dựng",
    "BGTVT" : "Bộ Giao thông vận tải",
    "UBSMC" : "Uỷ ban Sông Mê Kông"
}

include_word = [
    "Công nghệ thông tin",
    "Truyền thông",
    "Cổng thông tin điện tử",
    "Chuyển đổi số",
    "An toàn thông tin",
    "Mạng",
    "Ứng dụng công nghệ thông tin",
    "Hệ thống thông tin",
    "Cơ sở dữ liệu",
    "Dữ liệu",
    "Điện tử",
    "Bưu chính",
    "Kết nối và chia sẻ dữ liệu",
    "Dịch vụ Internet",
    "Viễn thông",
    "Mạng truyền số liệu chuyên dùng",
    "Báo",
    "Báo chí",
    "Xuất bản",
    "Thông tin và Truyền thông",
    "Di động",
    "Công nghệ",
    "Internet",
    "Phần mềm",
    "Thông tin"
]


exclude_word = [
    "/QĐ-UBND"
]

exclude_type = [
    "Quy chuẩn"
]

excel = "./Template.xlsx"
url = "https://thuvienphapluat.vn/rss.xml"
docs = feedparser.parse(url)

# Step 1: Filter data
def filter_docs(title):
    key_word = []

    if "BTTTT" in title or acronym_name["BTTTT"] in title:
        return True
    if any(word in title for word in exclude_word):
        return False
    for word in include_word:
        if word in title:
            key_word.append(word)
    if key_word:
        return True

    return False

def filter_type(type_name):
    if any(word in type_name for word in exclude_type):
        return False
    
    return True

def date_format(date):
    datetime_object = datetime.strptime(date, '%a, %d %b %Y %H:%M:%S %Z')
    formatted_date = datetime_object.strftime('%d/%m/%Y')

    return formatted_date


def extract_numbers(title):
    pattern = r'\b\d{1,7}/[\w-]+(?:/[\w-]+)?'
    matches = re.findall(pattern, title)
    return matches[0] if matches else ""


def ministry_name(doc_number):
    extracted_string = extract_numbers(doc_number)
    if '-' in extracted_string:
        return extracted_string.split('-')[-1]
    else:
        return ""


# Step 2: Add excel

def main():
    # 1. Filter data
    docs_list = []
    for doc in docs.entries:
        doc_filtered = []
        if filter_docs(doc.title) and filter_type(doc.tags[0].term[9:-3]):
            type_name = doc.tags[0].term[9:-3]
            doc_filtered.append(type_name)
            doc_filtered.append(extract_numbers(doc.title))
            doc_filtered.append(date_format(doc.published))
            doc_filtered.append(ministry_name(extract_numbers(doc.title)))
            doc_filtered.append(doc.title)
        else:
            continue
        docs_list.append(doc_filtered)

    docs_list = sorted(docs_list, key=lambda x: datetime.strptime(x[2], "%d/%m/%Y"))
    # Reverse order sort
    # docs_list = docs_list[::-1]
    print(docs_list)


    # 2. Add to excel
    # wb = openpyxl.load_workbook(excel)
    # ws = wb.active

    # thin_border = openpyxl.styles.Border(
    #     left = openpyxl.styles.Side(style="thin"),
    #     right = openpyxl.styles.Side(style="thin"),
    #     top = openpyxl.styles.Side(style="thin"),
    #     bottom = openpyxl.styles.Side(style="thin")
    # )

    # for row_index, row_value in enumerate(docs_list, start = 2):
    #     cell = ws.cell(row=row_index, column=1, value=row_index - 1)
    #     cell = ws.cell(row=row_index, column=2, value=docs_list[0])
    #     cell = ws.cell(row=row_index, column=3, value="Số: " + docs_list[1] + " ngày " + docs_list[2] + " của " + )
    #     cell = 
    #     cell = ws.cell()

    #     # Add border for a row
    #     for col in range



if __name__ == "__main__":
    main()


# print(docs.entries[0].title)
# print(docs.entries[0].link)
# print(date_format(docs.entries[0].published))
# print(docs.entries[0].tags[0].term[9:-3])