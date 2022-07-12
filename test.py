import imp
from openpyxl import Workbook

# 엑셀 파일 만들기
wb = Workbook()
ws = wb.active
ws.title = 'NadoSheet'
wb.save('nadoCoding_sample.xlsx')
wb.close()

