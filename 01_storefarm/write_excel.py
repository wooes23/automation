# -*- coding:utf-8 -*-

from pyexcel_xls import get_data
import xlsxwriter
import sys
from datetime import datetime

reload(sys)
sys.setdefaultencoding('utf-8')
date_time= datetime.now()


# 스토어팜 주문내역의 원본을 읽어옵니다.
data = get_data('20161122.xls')


# xlsx data 생성
# !! 생성한 날짜 기준으로 파일 이름 설정 해야함


workbook = xlsxwriter.Workbook('{}_Firex.xlsx'.format(date_time.strftime('%Y%m%d')))
worksheet = workbook.add_worksheet()


#style
bold = workbook.add_format({'bold':True, 'font_size':12})

form = workbook.add_format({'bold':True,
                            'font_color':'white',
                            'font_size':11,
                            'bg_color':'#3468bc',
                            'align':'center'})

data_form = workbook.add_format({'align': 'center'})

date_format = workbook.add_format({'num_format':'yyyy년 mm월 dd일 주문내역',
                                    'bold':True,
                                    'align':'center',
                                    'font_size':12})



#set column size

worksheet.set_column('A:A', 2)
worksheet.set_column('B:B', 10)
worksheet.set_column('C:C', 14)
worksheet.set_column('D:D', 5)
worksheet.set_column('E:E', 7)
worksheet.set_column('F:F', 8.5)
worksheet.set_column('G:G', 85)
worksheet.set_column('H:H', 16)
worksheet.set_column('I:I', 50)


# write xlsx
for tit, rec in data.items():
    for i in range(len(rec)):
        # print i, rec[i][12]
        #마지막 숫자를 조절하여 열을 조절합니다
        worksheet.write('A{n}'.format(n=i+2), i, data_form)   #no
        worksheet.write('B{n}'.format(n=i+2), rec[i][0][:8], data_form)    #주문일 pass
        worksheet.write('C{n}'.format(n=i+2), rec[i][12])   #품목
        worksheet.write('D{n}'.format(n=i+2), rec[i][16])   #수량
        worksheet.write('E{n}'.format(n=i+2), rec[i][7], data_form)   #수령인
        worksheet.write('F{n}'.format(n=i+2), rec[i][23], data_form)   #우편번호
        worksheet.write('G{n}'.format(n=i+2), rec[i][24])   #주소
        worksheet.write('H{n}'.format(n=i+2), rec[i][22], data_form)   #연락처
        worksheet.write('I{n}'.format(n=i+2), rec[i][25])   #배송메시지


#title area
#총합
worksheet.write('D1', '=SUM(D3:D100)', bold)
#주문 날짜
worksheet.write('G1', date_time, date_format)
#title 영역
worksheet.write('A2',' ', form)
worksheet.write('B2','주문일', form)
worksheet.write('C2','품목', form)
worksheet.write('D2','수량', form)
worksheet.write('E2','수령인', form)
worksheet.write('F2','우편번호', form)
worksheet.write('G2','주소', form)
worksheet.write('H2','연락처', form)
worksheet.write('I2','배송메시지', form)


workbook.close()