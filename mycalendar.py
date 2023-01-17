from datetime import date
import calendar
from lunarcalendar import Converter, Solar, Lunar, DateNotExist
import xlsxwriter


#year need to be updated 
yy = 2023

weeks = {1:'Mon',2:'Tue',3:'Wed',4:'Thur',5:'Fri',6:'Sat',7:'Sun'}
months = {1:'Jan',2:'Feb',3:'Mar',4:'April',5:'May',6:'June',7:'July',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('SolarLunarCalendar'+str(yy)+'.xlsx')

cell_format_weekend = workbook.add_format()
cell_format_weekend.set_bg_color('green')

cell_format_1n15 = workbook.add_format({'bold': True, 'font_color': 'red'})


# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

for mm in range(1,13):
    worksheet = workbook.add_worksheet(months[mm])
    row = 0
    worksheet.write(row, col,"    Name")
    worksheet.write(row , col + 5,"    Note")
    row += 1
    list = calendar.monthcalendar(yy, mm)
    for i in list:
        for k in i:
            if k != 0:
                nameOfDate = str(weeks[date.isoweekday(date(yy, mm, k))])
                solarDate = str(k)
                ludarDay = Converter.Solar2Lunar(Solar(yy,mm,k)).day
                lunarDate = str(Converter.Solar2Lunar(Solar(yy,mm,k)).day) + "/" + str(Converter.Solar2Lunar(Solar(yy,mm,k)).month)
                if nameOfDate == "Sat" or nameOfDate == "Sun":
                    worksheet.write(row , col ,nameOfDate + "    " + solarDate + "    " + lunarDate, cell_format_weekend)
                elif ludarDay == 1 or ludarDay == 15:
                    worksheet.write(row , col ,nameOfDate + "    " + solarDate + "    " + lunarDate, cell_format_1n15)
                else:
                    worksheet.write(row , col ,nameOfDate + "    " + solarDate + "    " + lunarDate)
                row += 1

workbook.close()