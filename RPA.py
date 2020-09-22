import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
# 엑셀파일 열기
sample = openpyxl.load_workbook('desktop/rpa/sample.xlsx') #로 데이터 시트 오픈
form = openpyxl.load_workbook('desktop/rpa/form.xlsx') #폼 데이터 시트 오픈

# 현재 Active Sheet 얻기
sample_work_sheet = sample.active #로 데이터 시트
form_work_sheet = form.active #폼 시트


# 샘플 데이터 범위 받아옴 셀에서 추출
sample_cell_1 = sample_work_sheet['A3':'l3']
data_list = []
for row in sample_cell_1:
    for cell in row:
        data_list.append(cell.value)

print(data_list)

# 샘플 데이터 입력단
form_work_sheet['C2'] = data_list[0]
form_work_sheet['E2'] = data_list[1]#ggg
form_work_sheet['C3'] = data_list[2]
form_work_sheet['E3'] = data_list[3]
form_work_sheet['C4'] = data_list[4]
form_work_sheet['E4'] = data_list[5]
form_work_sheet['C5'] = data_list[6]
form_work_sheet['E5'] = data_list[7]
form_work_sheet['C6'] = data_list[8]


sample_img = openpyxl.drawing.image.Image('desktop/rpa/image/capture.png')
#sample_img.anchor(form_work_sheet.cell('B6'))
form_work_sheet.add_image(sample_img,'B6')

form.save('desktop/rpa/output/output.xlsx')
form.close()
sample.close()
#work_sheet_new = work_book.create_sheet('new sheet')

# work_sheet.rows는 해당 쉬트의 모든 행을 객체로 가지고 있음
# for each_row in work_sheet.rows:
#     # cell(row=행 번호, column=열 번호).value = 해당 세로셀/가로셀에 어떤 값을 넣어주세요
#     work_sheet_new.cell(row=each_row[0].row, column=1).value = each_row[2].value
    
# 엑셀 파일 저장
# work_book.save("desktop/rpa/output.xlsx")
# work_book.close()