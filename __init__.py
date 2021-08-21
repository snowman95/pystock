#!/usr/bin/env python
#-*- coding: utf-8 -*-
# 
# 
import sys, datetime, tickers, pystocktool
sys.path.append(r'C:\Users\hoon\SDS\pyexcel\my_package')
import pyexceltool

time_now = datetime.datetime.today()
time_today = time_now.date()

excel_path = 'C:\\Users\\hoon\\SDS\\stock_report\\'
file_name = 'Portfolio.xlsx'
sheet_title = 'Portfolio'
report_file_name = str(time_today) + '_' + 'Stock_Report.xlsx'
new_report_file_name = f'{str(time_today)} {time_now.hour}-{time_now.minute}-{time_now.second}_Stock_Report.xlsx'
report_sheet_name = 'Report'



def __init__():	
	# Portfolio Excel 파일 로드
	workbook = pyexceltool.load_workbook_with_path(excel_path+file_name)
	df = pyexceltool.convert_worksheet_to_df(workbook, sheet_name=[sheet_title], include_index=False, include_column=True)
	
	# 데이터 수집
	for index, row in df.iterrows():
		print(row['ticker'])
		df.loc[index] = pystocktool.finpie_test.get_stock_data(row)

	# 레포트 Excel 파일 생성
	report_wb = pyexceltool.create_new_workbook(sheet_name=[report_sheet_name])
	pyexceltool.save_df_to_excel(df, report_wb, sheet_name=[report_sheet_name], file_path=excel_path+report_file_name, include_index=False, include_column=True)
