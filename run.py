# coding=utf-8
import fdb
import xlwt
import time
import calendar
import datetime
import os
import locale
from settings import *
import pytz
import smtplib
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders
from email.mime.multipart import MIMEMultipart


locale.setlocale(locale.LC_ALL, "es_CU")

# conexion to firebird database 
conex = fdb.connect(HOST, database=PATH_DB), user=USER, password=PASS)
cx = con.cursor()

def main():
   
      
    fecha = datetime.datetime.now()
    fecha_init = datetime.datetime(fecha.year,fecha.month - 1, 01)
    fecha_fin = datetime.datetime(fecha.year, fecha.month, 1)
    # start and finish in milliseconds
    mt_fin = int(time.mktime(fecha_fin.timetuple())) * 1000
    mt_init = int(time.mktime(fecha_init.timetuple())) * 1000

    # query to grt DB
    sql = "Select p.pchid, e.empid, e.empbadgenumber, (e.emplastname || ' ' || e.empfirstname || '' ) as Nombre , p.pchactualdate, p.pchcalcdate, E.schid AS HORARIO  \
           From punch P \
           Left Join employee E on P.empid = E.empid \
           where E.PCLID=41309  AND p.pchcalcdate between "+str(mt_init)+" And "+str(mt_fin)+" \
           ORDER BY e.empid, p.pchcalcdate asc"           
    cx.execute(sql)
    filas = cx.fetchall()


    # create a book with xlwt lib
    book = xlwt.Workbook(encoding='latin-1')
    sheet1 = book.add_sheet("Reporte Mensual GRT", cell_overwrite_ok=True)
    sheet1.col(0).width = 1000
    sheet1.col(1).width = 8000
    sheet1.col(2).width = 4000
    sheet1.col(3).width = 4000
    sheet1.col(4).width = 4000
    sheet1.col(5).width = 4000
    sheet1.col(6).width = 4000
    sheet1.col(7).width = 3000
    sheet1.col(8).width = 3000
    f = 1
    c = 0
    format = xlwt.Style.easyxf(
        """font: name Arial;borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour gray25; align: horiz center;""", )
    format_v = xlwt.Style.easyxf(
        """font: name Arial;borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour sea_green; align: horiz center;""", )
    format_o = xlwt.Style.easyxf(
        """font: name Arial;borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour light_orange; align: horiz center;""", )
    format_r = xlwt.Style.easyxf(
        """font: name Arial;borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour red; align: horiz center;""", )
    format_w = xlwt.Style.easyxf(
        """font: name Arial;borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour white; align: horiz center;""", )
    punch = {}
    j = 1
    for num in filas:
        row = sheet1.row(f - 1)
        row.write(0, "#", format)
        row.write(1, "NOMBRE", format)
        row.write(2, "PUNTUALIDAD", format_v)
        row.write(3, "RETARDO", format_o)
        row.write(4, "TARDE", format_r)       
        row = sheet1.row(f)
        row.write(0, j, format_w)
        row.write(1, num[3], format_w)
        row.write(2, retardo_puntualidad(num[1], filas)[1], format_w)
        row.write(3, retardo_puntualidad(num[1], filas)[0], format_w)      
        row.write(4, retardo_puntualidad(num[1], filas)[2], format_w)
        f += 2
        j += 1
        row = sheet1.row(f)
        row.write(1, "DIA", format)
        row.write(2, "FECHA", format)
        row.write(3, "PUNCH1", format)
        row.write(4, "PUNCH2", format)
        row.write(5, "PUNCH3", format)
        row.write(6, "PUNCH4", format)
        row.write(7, "COMIDA", format)
        row.write(8, "TOTAL", format)
        punch = ordeby_date(num[1], filas).items()
        punch.sort()
        for p in punch:
            row = sheet1.row(f + 1)
            row.write(1, datetime.datetime.strptime(p[0], '%Y-%m-%d').strftime('%A'), format_w)
            row.write(2, p[0], format_w)
            r = 2
            total_h = []
            for i in range(4):
                r += 1
                try:
                    if i==0:
                        time_llegada = datetime.time(9, 0, 0)
                        time_retardo = datetime.time(9, 15, 0)
                        time_p = datetime.datetime.fromtimestamp(int(p[1][i]) / 1000).strftime('%H:%M:%S')
                        hours_s = time_p.split(':')
                        hours = datetime.time(int(hours_s[0]), int(hours_s[1]), int(hours_s[2]))
                        if (hours > time_llegada and hours <= time_retardo):
                            format2 = format_o
                        elif (hours < time_llegada):
                            format2 = format_v
                        elif (hours > time_retardo):
                            format2 = format_r
                        row.write(r, datetime.datetime.fromtimestamp(int(p[1][i]) / 1000, tz).strftime('%H:%M:%S'),
                              format2)
                    else:
                        row.write(r, datetime.datetime.fromtimestamp(int(p[1][i]) / 1000, tz).strftime('%H:%M:%S'),
                              format_w)
                    total_h.append(int(p[1][i]) / 1000)

                except IndexError:
                    row.write(r,'',format_r)
            if len(total_h) == 4:
                pch1 = datetime.datetime.fromtimestamp(total_h[0], tz)
                pch2 = datetime.datetime.fromtimestamp(total_h[1], tz)
                pch3 = datetime.datetime.fromtimestamp(total_h[2], tz)
                pch4 = datetime.datetime.fromtimestamp(total_h[3], tz)
                alm = pch3 - pch2
                ht = pch4 - pch1
                row.write(7, str(alm),format_w)
                row.write(8, str(ht-alm), format_w)
            else:
                row.write(7,'', format_r)
                row.write(8, '', format_r)
            f += 1
        f += 3
        delemploy(num[1], filas)   
    book.save(PATH_BOOK)
    con.commit()
    con.close()
    mail()

def delemploy(epid, employee):
    for em in employee[:]:
        if (em[1] == epid):
            employee.remove(em)

def ordeby_date(epid, employee):
    result = {}
    for p in employee:
        if (p[1] == epid):
            punch = []
            fecha = datetime.datetime.fromtimestamp(int(p[4]) / 1000).strftime('%Y-%m-%d')
            if fecha not in result:
                punch.append(p[4])
                result[fecha] = punch
            else:
                result[fecha].append(p[4])
    return result

def tipo_horario(employee):
    if (employee == 248250):
        return "AM"
    elif (employee == 248252):
        return "PM"
    return ""

def retardo_puntualidad(epid, employee):
    tarde = 0
    retardo = 0
    puntualidad = 0
    result = []
    date = ""
    for p in employee:
        if (p[1] == epid):
            fecha = datetime.datetime.fromtimestamp(int(p[4]) / 1000).strftime('%Y-%m-%d')
            if (fecha != date):
                date = fecha
                punch = datetime.datetime.fromtimestamp(int(p[4]) / 1000).strftime('%H:%M:%S')
                hours_s = punch.split(':')
                hours = datetime.time(int(hours_s[0]), int(hours_s[1]), int(hours_s[2]))
                time_llegada = datetime.time(9, 0, 0)
                time_retardo = datetime.time(9, 15, 0)
                if (hours > time_llegada and hours <= time_retardo ):
                    retardo += 1
                elif (hours < time_llegada):
                    puntualidad += 1
                elif (hours > time_retardo):
                    tarde +=1

    result.append(str(retardo))
    result.append(str(puntualidad))
    result.append(str(tarde))
    return result
    

    
if __name__ == "__main__":
    main()
