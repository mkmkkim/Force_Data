![header](https://capsule-render.vercel.app/api?type=venom&color=auto&height=200&section=header&text=Force%20Data&fontSize=40&)</br>
* 활동: 전남대학교 플랜트공학전공 교수님의 연구 과제 참여하여 엑셀 도식화 코드를 작성해 업무 시간 70% 이상 단축</br>
* 주제: 피험자 100명에게 신체 15부위 별로 힘을 가하고 어느 정도 힘에서 통증을 느끼는지 기록한 데이터(이하 Force Data)를 분석하여 가상의 피험자 1,000명의 Force Data를 생성</br>
* 프로젝트 기간: 2020.09~12</br>
* 인원: 1명</br>
* 담당: 체중, 키, 나이 등 다양한 기준으로 피험자 데이터 분석 및 도식화</br>
* 사용 기술 및 도구:  Python3.6(matplotlib, pandas, openxl), Anaconda, Jupyter Notebook</br>
![image](https://github.com/mkmkkim/Force_Data/assets/74914390/c4d495bc-9310-404b-a130-29e328cf41a5)</br>
### 코드</br>
```
# 라인2 파일명을 바꿔가며 코드 실행 시 해당 엑셀 파일에 대한 차트를 그릴 수 있음
# 차트 종류는 라인12에서 변경 가능
Import pandas as pd 
Df=pd.read_excel(‘./age.xlsx’, headers=0, index_col=0)
From openpyxl import Workbook
From openpyxl.chart import BarChart, Reference
From openpyxl.utils.dataframe import dataframe_to_rows

Wb=Workbook(write_only=True)
Ws=wb.create_sheet() For row in dataframe_to_rows(Df, index=True, header=True):
	if len(row)>1:
	ws.append(row)
	print(row)
chart.=BarChart()	
Chart.type=“col”	
Chart.style=10	
Chart.title=“연령”	
Chart.y_axis.title=“force(N)”
Chart.x_axis.title=“location”
Data=Reference(ws, min_col=2, max_col=6, min_row=1, max_row=16)
Cats=Reference(ws, min_col=1, min_row=2, max_row=16)

Chart.add_data(data, titles_from_data=True) Chart.set_categories(cats)
Chart.shape=4 
Ws.add_chart(chart, “A19”)
Wb.save(“./age barchart.xlsx”)
```
