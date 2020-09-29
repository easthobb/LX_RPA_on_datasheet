import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Alignment
import win32com.client
import PIL


## 토지 양식 입력 자동화 python file

###출력파일 및 excel PDF 파일 절대경로 설정
filepath ="C:\\Users\\user\\Desktop\\RPA\\"
input_image_format = '.png'
saved_name = '실태조사_'

### data sheet에서 필요한 설정요소 ### 
data = openpyxl.load_workbook(filepath+'resultDB.xlsx') #로 데이터 시트 오픈(변경가능)
data_sheet_name = '시계외' ## 데이터 파일 시트 
start_column = 'A' ## 데이터 파일 시트에서 값을 가져올 시작 열
end_column = 'Q' ## 데이터 파일 시트에서 값을 가져올 끝 열
iter = 5 ## 문서 생성을 할 데이터의 갯수
data_sheet = data.get_sheet_by_name(data_sheet_name) #로 데이터 파일(resultDB.xlxs,시트명 : '시계외' or '시내')





## 엑셀 PDF 연결, background 실행
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.ScreenUpdating = False
excel.DisplayAlerts = False
excel.EnableEvents = False

# 로 데이터 개별 객체를 데이터 시트에서 추출
for i in range(iter):

    ### form sheet 에서 필요한 설정요소
    form = openpyxl.load_workbook(filepath+'RPA_form.xlsx') #폼 데이터 시트 경로 지정(변경가능)
    form_sheet_list = ['simple'] ## 접근할 시트(추가 및 삭제 가능)
    form_sheet_page_1 = form.get_sheet_by_name(form_sheet_list[0]) #폼 시트 페이지 1(simple, 시트명 : simple)
    
    ### 시작셀과 마지막 셀 설정 
    sheet_start_cell = start_column + str(i+4) ## 하드코딩 시작 데이터 셀 ex) A4 -> i = 4 , A7 -> i = 7
    sheet_end_cell = end_column + str(i+4)
    
    land_object = data_sheet[sheet_start_cell:sheet_end_cell] # 반복시 변경해야 됩니다
    data_list = [] ## 개별 토지 객체의 정보가 담기는 리스트 생성
    for row in land_object: 
        for cell in row:
            data_list.append(cell.value)

    print(data_list) # 각 행의 값들을 콘솔에 표시
    key_value = str(data_list[10]) #넘버링을 위한 key value 엑셀에 ****수임번호****에 해당합니다. , 예상 정수값
    
    # (예정) for 문 삽입부분
    # 추출한 데이터를 양식에 매핑시켜주는 부분 :  land_1 sheet
    # form_sheet_page_1['B7'] = data_list[0] ## 조사연번
    form_sheet_page_1['C2'] = "   "+ data_list[1] ## 소재지
    form_sheet_page_1['D7'] = data_list[2] ## 대장지목
    form_sheet_page_1['J7'] = data_list[3] ## 대장면적
    form_sheet_page_1['D8'] = data_list[4] ## 공시지가(20년)
    form_sheet_page_1['J8'] = data_list[5] ## 취득일자
    #form_sheet_page_1['D9'] = data_list[6] ## 공유지분여부 - form에 없음
    form_sheet_page_1['D9'] = data_list[7] ## 공유자수(공유자수 시 포함)
    form_sheet_page_1['J9'] = data_list[8] ## 서울시 공유지분
    form_sheet_page_1['D10'] = data_list[9] ## PNU
    form_sheet_page_1['B2'] = data_list[10] ## 수임번호 - 라벨
    form_sheet_page_1['J10'] = data_list[10] ## 수임번호 - 시트
    form_sheet_page_1['D11'] = data_list[11] ## 수탁일자
    form_sheet_page_1['J11'] = data_list[12] ## 재산관리관
    form_sheet_page_1['D12'] = data_list[13] ## 사용형태
    form_sheet_page_1['D13'] = data_list[14] ## 조사일자
    form_sheet_page_1['D14'] = data_list[15] ## 조사내용
    #form_sheet_page_1['D13'] = data_list[16] ## 토지이용계획 - form에 없음
    
    
    # 이미지 입력부분 : 리사이즈 , 이미지 파일명 : 수임번호_1( ), 수임번호-2( ), 수임번호_3( )
    # (향후 변경) 1 : 지적도 , 2 : 국토정보기본도, 3 : 현황사진
    for j in range(3):
        img_file_name = filepath + '\\image\\' + '\\simple\\' + key_value +'_' +str(j+1) + input_image_format # 이미지경로 
        print(img_file_name)
        try: ##이미지 파일 삽입시도.
            img = openpyxl.drawing.image.Image(img_file_name)
        except: ##해당하는 이미지 파일이 없을경우, form overriding 방지
            print('수임번호'+ key_value +'에 해당하는 이미지 파일이 존재하지 않습니다.')
        else:
            if(j==0): # 지적도
                img.width=432 # 이미지 리사이징, 가로.픽셀 단위입니다.
                img.height=286 # 이미지 리사이징 세로.픽셀 단위입니다.
                form_sheet_page_1.add_image(img,'B18') ## 이미지가 들어갈 셀
            elif(j==1): # 국토정보기본도
                img.width=432 # 이미지 리사이징, 가로.픽셀 단위입니다.
                img.height=286 # 이미지 리사이징 세로.픽셀 단위입니다.
                form_sheet_page_1.add_image(img,'H18') ## 이미지가 들어갈 셀
            elif(j==2): # 현황사진
                img.width=432 # 이미지 리사이징, 가로.픽셀 단위입니다.
                img.height=286 # 이미지 리사이징 세로.픽셀 단위입니다.
                form_sheet_page_1.add_image(img,'E32') ## 이미지가 들어갈 셀
        

    #output file 저장 부분
    output_file_name = filepath + '\\output\\' + saved_name + key_value +'.xlsx' ## data_list[10] 수임번호
    form.save(output_file_name)
    form.close()

    #output file PDF 저장
    excel.Visible = False
    wb = excel.WorkBooks.Open(output_file_name)
    excel.Visible = False
    ws_chart = wb.WorkSheets(form_sheet_list)
    ws_chart.Select()
    pdf_save_path = filepath + "\\output\\" +saved_name + key_value + '.pdf'
    wb.ActiveSheet.ExportAsFixedFormat(0,pdf_save_path)
    wb.Close(True)

data.close()
excel.Quit()