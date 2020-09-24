import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Alignment
import win32com.client
import time

## iteration 변수
iter = 4

## 건물 양식 입력 자동화 python file

# 엑셀파일 열기
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False

data = openpyxl.load_workbook('desktop/rpa/data.xlsx') #로 데이터 시트 오픈
form = openpyxl.load_workbook('desktop/rpa/RPA_form.xlsx') #폼 데이터 시트 오픈

# 로 데이터 시트와 폼 시트의 fopen
data_sheet = data.get_sheet_by_name('land') #로 데이터 시트 객체 파일(land,시트명 : land)
form_sheet_page_1 = form.get_sheet_by_name('land_1') #폼 시트 페이지 1(building, 시트명 : land_1)
form_sheet_page_2 = form.get_sheet_by_name('land_2') #폼 시트 페이지 2(building, 시트명 : land_2)
form_sheet_page_3 = form.get_sheet_by_name('land_3') #폼 시트 페이지 3(building, 시트명 : land_3)

# 로 데이터 개별 객체를 데이터 시트에서 추출
for i in range(iter): 
    sheet_start_cell = "A" + str(i+iter) 
    sheet_end_cell = "AP" + str(i+iter)
    
    land_object = data_sheet[sheet_start_cell:sheet_end_cell] # 반복시 변경해야 됩니다
    data_list = [] ## 개별 토지 객체의 정보가 담기는 리스트 생성
    for row in land_object: 
        for cell in row:
            data_list.append(cell.value)

    print(data_list)
    # (예정) for 문 삽입부분
    # 추출한 데이터를 양식에 매핑시켜주는 부분 :  land_1 sheet
    form_sheet_page_1['B1'] = data_list[0] ##관리번호
    form_sheet_page_1['D6'] = data_list[1] ##고유번호
    form_sheet_page_1['J6'] = data_list[2] ## 재산번호
    form_sheet_page_1['D7'] = data_list[3] ## 소재지
    form_sheet_page_1['J7'] = data_list[4] ## 재산명칭
    form_sheet_page_1['D8'] = data_list[5] ## 공유자수
    form_sheet_page_1['G8'] = data_list[6] ## 공유지분
    form_sheet_page_1['J8'] = data_list[7] ## 관리관
    form_sheet_page_1['D9'] = data_list[8] ## 재산구분
    form_sheet_page_1['J9'] = data_list[9] ## 위임기관

    form_sheet_page_1['D10'] = data_list[10] ## 대장면적
    form_sheet_page_1['J10'] = data_list[11] ## 대장지목
    form_sheet_page_1['D11'] = data_list[12] ## 이용면적
    form_sheet_page_1['J11'] = data_list[13] ## 이용지목
    form_sheet_page_1['D12'] = data_list[14] ## 공시지가
    form_sheet_page_1['J12'] = data_list[15] ## 점유형태
    form_sheet_page_1['D13'] = data_list[16] ## 취득일자
    form_sheet_page_1['J13'] = data_list[17] ## 점유필지
    form_sheet_page_1['D14'] = data_list[18] ## 사용형태
    form_sheet_page_1['D15'] = data_list[19] ## 조사현황
    form_sheet_page_1['C16'] = data_list[20] ## 토지이용계획
  


    #추출한 데이터를 매핑시켜주는 부분 :  building_2 sheet

    form_sheet_page_2['B16'] = data_list[21] ## 기호
    form_sheet_page_2['C16'] = data_list[22] ## 사용형태(계약방법)
    form_sheet_page_2['D16'] = data_list[23] ## 이용상황
    form_sheet_page_2['F16'] = data_list[24] ## 점유면적
    form_sheet_page_2['H16'] = data_list[25] ## 계약일
    form_sheet_page_2['J16'] = data_list[26] ## 계약기간
    form_sheet_page_2['L16'] = data_list[27] ## 대부료
    form_sheet_page_2['M16'] = data_list[28] ## 비고
    form_sheet_page_2['D22'] = data_list[29] ## 접근성
    form_sheet_page_2['D23'] = data_list[30] ## 주요시설

    form_sheet_page_2['D27'] = data_list[31] ## 1차분류=임시
    form_sheet_page_2['F27'] = data_list[32] ## 1차분류=임시
    form_sheet_page_2['H27'] = data_list[33] ## 1차분류=임시
    form_sheet_page_2['J27'] = data_list[34] ## 1차분류=임시
    form_sheet_page_2['L27'] = data_list[35] ## 1차분류=임시
    form_sheet_page_2['D28'] = data_list[36] ## 2차분류=임시
    form_sheet_page_2['F28'] = data_list[37] ## 2차분류=임시
    form_sheet_page_2['H28'] = data_list[38] ## 2차분류=임시
    form_sheet_page_2['J28'] = data_list[39] ## 2차분류=임시
    form_sheet_page_2['L28'] = data_list[40] ## 2차분류=임시
    form_sheet_page_2['D29'] = data_list[41] ## 활용의견
    
    # 이미지 입력부분 : 리사이즈 , 이미지 파일명 : 관리번호-1( ),관리번호-2( ), 관리번호-3( ),관리번호-4( ), 관리번호-5( ),
    # (향후 변경) 1 : 지적도 , 2 : 국토정보기본도, 3 : 현황사진, 4 :사용허가및무단점유현황, 5 : 토지이용계획확인서 
    for j in range(5):
        img_file_name = 'desktop/rpa/image/' + str(data_list[0]) +'-' +str(j+1) +'.png' # 첫번째 페이지 이미지 이름 지정
        print(img_file_name)
        img = openpyxl.drawing.image.Image(img_file_name)
        if(j==0): # 지적도
            img.width=430 # 이미지 리사이징, 가로.픽셀 단위입니다.
            img.height=286 # 이미지 리사이징 세로.픽셀 단위입니다.
            form_sheet_page_1.add_image(img,'B20')
        elif(j==1): # 국토정보기본도
            img.width=430 # 이미지 리사이징, 가로.픽셀 단위입니다.
            img.height=286 # 이미지 리사이징 세로.픽셀 단위입니다.
            form_sheet_page_1.add_image(img,'H20')
        elif(j==2): # 현황사진
            img.width=430 # 이미지 리사이징, 가로.픽셀 단위입니다.
            img.height=286 # 이미지 리사이징 세로.픽셀 단위입니다.
            form_sheet_page_1.add_image(img,'E34')
            cell = form_sheet_page_1.cell(row=12,column=13)
            cell.alignment = Alignment(horizontal='center')
        elif(j==3): # 사용허가및무단점유현황
            img.width=864 # 이미지 리사이징, 가로.픽셀 단위입니다.
            img.height=400 # 이미지 리사이징 세로.픽셀 단위입니다.
            form_sheet_page_2.add_image(img,'B3')
        elif(j==4):
            img.width=864 # 이미지 리사이징, 가로.픽셀 단위입니다.
            img.height=816 # 이미지 리사이징 세로.픽셀 단위입니다.
            form_sheet_page_3.add_image(img,'B3')

    #output file 저장 부분
    output_file_name = 'desktop/rpa/output/land' + str(data_list[0]) +'.xlsx'
    form.save(output_file_name)
    time.sleep(1)
    #output file PDF 저장
    output_file_root = "C:\\Users\\user\\Desktop\\RPA\\output\\land" + str(data_list[0]) + '.xlsx' 
    wb = excel.WorkBooks.Open(output_file_root)
    ws_chart = wb.WorkSheets(['land_1','land_2','land_3'])
    ws_chart.Select()
    pdf_save_path = "C:\\Users\\user\\Desktop\\RPA\\output\\" + str(data_list[0]) + '.pdf'
    wb.ActiveSheet.ExportAsFixedFormat(0,pdf_save_path)
    wb.Close(True)
    
    # sample_img.anchor(form_work_sheet.cell('B6'))
    # form_work_sheet.add_image(sample_img,'B6')

#form.save('desktop/rpa/output/output.xlsx')

form.close()
data.close()
excel.Quit()

#work_sheet_new = work_book.create_sheet('new sheet')

# work_sheet.rows는 해당 쉬트의 모든 행을 객체로 가지고 있음
# for each_row in work_sheet.rows:
#     # cell(row=행 번호, column=열 번호).value = 해당 세로셀/가로셀에 어떤 값을 넣어주세요
#     work_sheet_new.cell(row=each_row[0].row, column=1).value = each_row[2].value
    
# 엑셀 파일 저장
# work_book.save("desktop/rpa/output.xlsx")
# work_book.close()