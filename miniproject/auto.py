# 엑셀 파일 A1,B1은 비워두기 내용 저장 안됩니다.
# 전화번호 입력할 때 (-) 빠지면 안됩니다.
# 여기서 이제 사이트에 엑셀 파일 업로드하면 바로 vrf파일 형태로 변환해주는 툴 만들기 

import openpyxl
from vobject import vCard

def create_vcard(full_name, phone_number):
    card = vCard()
    card.add('fn').value = full_name
    card.add('tel').value = phone_number
    return card

def excel_to_vcards(excel_file, output_file):
    # 엑셀 파일 열기
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active

    # vCard 생성 및 파일에 쓰기
    with open(output_file, 'w', encoding='utf-8') as f:
        for row in sheet.iter_rows(min_row=2, values_only=True):  # 첫 번째 행은 헤더로 가정
            full_name, phone_number = row[0], row[1]  # A열: 성명, B열: 전화번호
            vcard = create_vcard(full_name, phone_number)
            f.write(vcard.serialize())

    print(f"연락처가 {output_file}에 성공적으로 저장되었습니다.")

# 사용 예시
excel_file = "연락처_목록.xlsx"
output_file = "연락처_목록.vcf"

excel_to_vcards(excel_file, output_file)