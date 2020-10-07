# LX_RPA_on_datasheet

this project is for excel RPA datasheet for LX seoul HQ using OpenpyXl
엑셀 데이터 -> 문서(excel,PDF) 양식 입력 자동화를 위한 파이썬 배치파일
2020/09/28 - ver 1.0 rel.
2020/10/07 - ver 1.1 + 파이썬&의존성 설치 배치파일 작성 및 이미지 resize 디버그
---
## 기능 요구사항
-  엑셀파일의 데이터(1k건 단위)를 엑셀 form 의 문서로 변환 기능
-  특정 폴더에 있는 이미지를 resize를 거쳐 form 에 입력 기능 
-  excel 문서와 엑셀을 용이하게 출력하기 위한 PDF 변환 기능  

---
## 기술 요구사항
- No backend
- 조작과 매핑을 위한 간단한 UI

---
## 동작환경
- 일반적인 사무용 컴퓨터
- python3 dependency

---
## 가이드
- Notion 으로 배포

---
## dependency
- openpyxl
- pywin32
- pillow