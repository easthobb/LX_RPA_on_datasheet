@echo 엑셀 입력 자동화를 위한 필요 프로그램 설치입니다.
@echo 파이썬 설치파일 실행입니다.
@echo 파이썬 설치를 진행해주세요!
@echo 설치시 add Python 3.8 to path 옵션을 선택해주세요!!
@echo 작성일 : 2020/10/07 , ver 1.1

python-3.8.6-amd64.exe  
dir
tree
@echo 윈도우 제어 파이썬 패키지를 설치 중입니다.
pip3 install pywin32
@echo 엑셀 제어 파이썬 패키지를 설치 중입니다.
pip3 install openpyxl
@echo 이미지 제어 파이썬 패키지를 설치 중입니다.
pip3 install pillow
pause