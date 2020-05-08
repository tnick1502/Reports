import numpy as np
import matplotlib
import os

matplotlib.use("TkAgg")
matplotlib.rcParams['agg.path.chunksize'] = 100000
from matplotlib.figure import Figure
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from reportlab.platypus import Table
from reportlab.graphics import renderPDF

from openpyxl import load_workbook
from svglib.svglib import svg2rlg  # Эта


def frange(start, stop, step):   # Создает массив с нецелым шагом
    i = start
    while i < stop:
        yield i
        i += step
def str1(s):  # Проверяет строку из Exel и делает ее str. Если она пустая, то возвращает -
    if str(s) == "None":
        return '-'
    else:
        return str(s)
def objlen(s, leng):  # Перемещает часть строки на другую строчку. Возвращает 2 строки. s- то, что разбивается, leng - длина разбивки
    if len(s) > leng:
        for i in range(leng, 0, -1):
            if s[i] == " ":
                break
        s1 = s[0:i]
        s2 = s[(i + 1):(len(s))]
    else:
        s1 = s
        s2 = " "
    return s1, s2
def scr(canvas, shrift, x, y, Nadp1, Nadp2, Nadp3, Nadp4, st):  # Делает надпись с подстрокой. размер шрифта,    координата x,y,  надпись основная,надпись вторая,надпись степень,    1-верх,-1-низ
    canvas.setFont('Times', shrift)
    canvas.drawString(x * mm, y * mm, Nadp1)
    if Nadp2 != 0:
        canvas.drawString((x + 3 + (len(Nadp1) - 1) * 1.4) * mm, y * mm, Nadp2)
    canvas.setFont('Times', shrift - 3)
    canvas.drawString((x + 1.8 + (len(Nadp1) - 1) * 1.4) * mm, (y + st * 1) * mm, Nadp3)
    if Nadp4 != 0:
        canvas.drawString((x + 2.8 + (len(Nadp1) - 1) * 1.4 + (len(Nadp2) - 1) * 1.65) * mm, (y + 1) * mm,
                          Nadp4)
    canvas.setFont('Times', shrift)
def zap(s, m):  # Количство знаков после запятой. s - число в str, m - число знаков
    if s != "-":

        try:

            i = s.index(",")

            if len(s) - i > m:
                s = s[0:i + m + 1]
            elif len(s) - i <= m:
                for i in range(m - len(s) + i + 1):
                    s += "0"

        except ValueError:
            s += ","
            for i in range(m):
                s += "0"

        return s
    else:
        return s


def Frame(canvas, chislo, path, p1, Akred, List, Res, Cod):  # Рамка, шапка, Защитный код, Исполнители и аккредитация из файла

            def SaveCode(p1):  # Создает защитный код и записывает его в файл
                Buk = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S',
                       'T',
                       'U',
                       'W', 'Q', 'V', 'Z']
                Chis = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
                code = str(np.random.choice(Chis)) + str(np.random.choice(Buk)) + str(np.random.choice(Buk)) + str(
                    np.random.choice(Chis)) + str(np.random.choice(Chis)) + '-' + str(np.random.choice(Buk)) + str(
                    np.random.choice(Chis)) + str(np.random.choice(Chis)) + str(np.random.choice(Chis))
                with open(os.path.join(p1, "Результаты.txt"), "w") as file:
                    for i in range(len(Res)):
                        file.write(str1(Res[i]) + '\t')
                    file.write('\n')
                    file.write(code)
                file.close()
                return code


            canvas.setLineWidth(0.7 * mm)
            canvas.rect(20 * mm, 5 * mm, 185 * mm, 287 * mm)   # Основная рамка



            # Вертикальные линии в нижней таблице
            for i in range(37,125,17):
                canvas.line(i * mm, 5 * mm, i * mm, 20 * mm)
            canvas.line(188 * mm, 5 * mm, 188 * mm, 20 * mm)



            # Горизонтальные линии в нижней таблице
            canvas.line(20 * mm, 10 * mm, 122 * mm, 10 * mm)
            canvas.line(20 * mm, 15 * mm, 122 * mm, 15 * mm)
            canvas.line(20 * mm, 20 * mm, 205 * mm, 20 * mm)
            canvas.line(188 * mm, 15 * mm, 205 * mm, 15 * mm)



            # Данные внизу таблицы
            dat1 = [["", "", "", "", "", str(chislo[8] + chislo[9] + "." + chislo[5] + chislo[6] + "." + chislo[2] + chislo[3])],
                        ["Изм.", "Кол. уч.", "Лист", "№ док.", "Подпись", "Дата"]]
            t1 = Table(dat1, colWidths=17 * mm, rowHeights=5 * mm)
            t1.setStyle([("FONTNAME", (0, 0), (-1, -1), 'Times'),
                        ("FONTSIZE", (0, 0), (-1, -1), 10),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER")])
            t1.wrapOn(canvas, 0, 0)
            t1.drawOn(canvas, 20 * mm, 5 * mm)

            dat2 = [["Лист"], [List], [""]]
            t2 = Table(dat2, colWidths=17 * mm, rowHeights=5 * mm)
            t2.setStyle([("FONTNAME", (0, 0), (-1, -1), 'Times'),
                         ("FONTSIZE", (0, 0), (-1, -1), 10),
                         ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                         ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                         ('SPAN', (0, 1), (-1, -1))])
            t2.wrapOn(canvas, 0, 0)
            t2.drawOn(canvas, 188 * mm, 5 * mm)

            canvas.setFont('Times', 10)
            #canvas.drawString(192.5 * mm, 16.5 * mm, "Лист")
            #canvas.drawString(194.5 * mm, 8.5 * mm, "1/1")



            # Шапка
            xmoveV = 0
            ymoveV = 0

            canvas.line((50 + xmoveV) * mm, (280 + ymoveV) * mm, (198 + xmoveV) * mm, (280 + ymoveV) * mm) # Линия аккредитации
            canvas.drawImage(path + "Python(data)/Logo2.jpg", 23 * mm, 270 * mm,
                             width=21 * mm, height=21 * mm)  # логотип
            canvas.setFont('TimesDj', 20)
            canvas.drawString((50 + xmoveV) * mm, (282 + ymoveV) * mm, "МОСТДОРГЕОТРЕСТ")
            canvas.setFont('TimesDj', 12)
            canvas.drawString((130 + xmoveV) * mm, (284.8 + ymoveV) * mm, "испытательная лаборатория")
            canvas.setFont('Times', 10)
            canvas.drawString((130 + xmoveV) * mm, (281 + ymoveV) * mm, "129344, г. Москва, ул. Искры, д.31, к1")



            # Исполнители и Аккредитация
            xmoveN = 0
            ymoveN = 0

            A = []  # аккредитация и низ
            fi = open(path + "Python(data)/Data(НЕ УДАЛЯТЬ).txt")
            line = fi.readline().strip()
            while line:
                p = line.split('\t')
                A.append(p)
                line = fi.readline().strip()
            fi.close()

            if Akred == "AS" or Akred == "AN":
                s = 0
            elif Akred == "OS" or Akred == "ON":
                s = 5
            dat3 = [[A[0 + s][0], A[0 + s][1]],
                    ['', A[0 + s][2]],
                    [A[1 + s][0], A[1 + s][1]],
                    [A[2 + s][0], A[2 + s][1]],
                    [A[3 + s][0], A[3 + s][1]]]
            t3 = Table(dat3, colWidths=90 * mm, rowHeights=5.5 * mm)
            t3.setStyle([("FONTNAME", (0, 0), (-1, -1), 'Times'),
                        ("FONTSIZE", (0, 0), (-1, -1), 10),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (0, -1), 0),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                        ("ALIGN", (0, 0), (-1, -1), "LEFT"), ])

            t3.wrapOn(canvas, 0, 0)
            t3.drawOn(canvas, 25 * mm, 24 * mm)

            if Akred == "OS":
                dat4 = [[A[9][1]], [A[9][2]]]
            elif Akred == "ON":
                dat4 = [[A[10][1]], [A[10][2]]]
            elif Akred == "AS":
                dat4 = [[A[11][1]], [A[11][2]]]
            elif Akred == "AN":
                dat4 = [[A[12][1]], [A[12][2]]]

            t4 = Table(dat4, colWidths=145 * mm, rowHeights=4 * mm)
            t4.setStyle([("FONTNAME", (0, 0), (-1, -1), 'Times'),
                        ("FONTSIZE", (0, 0), (-1, -1), 8),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (0, -1), 0),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"), ])

            t4.wrapOn(canvas, 0, 0)
            t4.drawOn(canvas, (50 + xmoveN) * mm, (271.5 + ymoveN) * mm)



            # Защитный код
            if Cod==0:
                code = SaveCode(p1)
            else:
                code = Cod
            canvas.setFont('Times', 6)
            canvas.rotate(90)
            canvas.drawString(5 * mm, -18.5 * mm, code)
            canvas.rotate(-90)
            return code

def TopDate(canvas, wb, Nop):  # Верхняя таблица данных

            xmove = 0
            ymove = 0



            # Верхняя надпись
            if str1(wb["Лист1"]['IG' + str1(6 + Nop)].value) == "-":
                LabNom = str1(wb["Лист1"]['A' + str1(6 + Nop)].value)
            else:
                LabNom = str1(wb["Лист1"]['IG' + str1(6 + Nop)].value)

            ff=str(LabNom + "/" + str1(wb["Лист1"]['AI1'].value + "/РК"))
            dat1 = [["Протокол испытаний №:", ff],
                    ["ОПРЕДЕЛЕНИЕ СЕЙСМИЧЕСКОЙ РАЗЖИЖАЕМОСТИ ГРУНТОВ МЕТОДОМ", ""],
                    ["ЦИКЛИЧЕСКИХ ТРЁХОСНЫХ СЖАТИЙ С РЕГУЛИРУЕМОЙ НАГРУЗКОЙ (ASTM D5311-11)", ""]]
            t1 = Table(dat1, colWidths=100 * mm, rowHeights=5 * mm)

            t1.setStyle([("FONTNAME", (0, 0), (0, 0), 'TimesDj'),
                         ("FONTNAME", (1, 0), (1, 0), 'Times'),
                         ("FONTNAME", (0, 1), (-1, -1), 'TimesDj'),
                         ("FONTSIZE", (0, 0), (-1, -1), 10),
                         ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                         ("ALIGN", (0, 1), (-1, -1), "CENTER"),
                         ("LEFTPADDING", (0, 0), (0, 0), 62 * mm),
                         ("LEFTPADDING", (1, 0), (1, 0), 3 * mm),
                         ('SPAN', (0, 1), (1, 1)),
                         ('SPAN', (0, 2), (1, 2))])
            t1.wrapOn(canvas, 0, 0)
            t1.drawOn(canvas, (20 + xmove) * mm, (254 + ymove) * mm)



            # Таблица верхней части. создаем массив данных 4х9. Ширину отступов задаем в стиле
            s = str1(wb["Лист1"]['A2'].value)  # название объекта
            s1, s2 = objlen(s, 100)
            s2, s3 = objlen(s2, 100)

            NameGround = str1(wb["Лист1"]['D' + str(6 + Nop)].value)  # название грунта

            if str1(wb["Лист1"]['FV' + str(6 + Nop)].value) != "-":
                pref = str(round(float((wb["Лист1"]['FV' + str(6 + Nop)].value)), 3))  # раньше HN
            else:
                pref = "-"


            dat2 = [['Заказчик:', str1(wb["Лист1"]["A1"].value), '', ''],
                       ['Объект:', s1, '', ''],
                       ['', s2, '', ''],
                       ['', s3, '', ''],
                       ['Лабораторный номер №:', LabNom, 'ИГЭ:',
                        str1(wb["Лист1"]['ES' + str1(6 + Nop)].value)],
                       ['Наименование выработки:', str1(wb["Лист1"]['B' + str1(6 + Nop)].value), 'Глибина отбора, м:',
                        str1(wb["Лист1"]['C' + str1(6 + Nop)].value)],
                       ['Наименование грунта:', NameGround, '', ''],
                       ['Режим испытания:', 'Анизотропная реконсолидация, девиаторное циклическое нагружение', '',
                        ''],
                       ['Опорное давление p   , МПа:', pref, '', ''],
                       ['Диаметр образца, мм: 38', 'Высота образца, мм: 76',
                        'Оборудование: Wille Geotechnik 13-HG/020:001', '']]
            t2 = Table(dat2, colWidths=19 * mm, rowHeights=5.5 * mm)  # 4 для вилли
            t2.setStyle([("FONTNAME", (0, 0), (-1, -1), 'Times'),
                        ("FONTSIZE", (0, 0), (-1, -1), 10),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (0, -1), 0),  # весь первый столбец    (столбец,строка)
                        ("LEFTPADDING", (1, 0), (1, 0), 27 * mm),  # второй столбец
                        ("LEFTPADDING", (1, 0), (1, 0), -1 * mm),  # заказчик
                        ("LEFTPADDING", (1, 1), (1, 3), -1 * mm),  # объект
                        ("LEFTPADDING", (1, 4), (1, -1), 27 * mm),  # второй столбец
                        ("LEFTPADDING", (2, 0), (2, -1), 55 * mm),  # третий столбец
                        ("LEFTPADDING", (3, 0), (3, 0), 87 * mm),  # четвертый столбец
                        ("LEFTPADDING", (3, 1), (3, -1), 67 * mm),  # четвертый  столбец
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                        ("ALIGN", (0, 0), (-1, -1), "LEFT"), ])

            t2.wrapOn(canvas, 0, 0)
            t2.drawOn(canvas, (25 + xmove) * mm, (196 + ymove) * mm)
            canvas.setFont('Times', 6)
            canvas.drawString((55.5 + xmove) * mm, (200.5 + ymove + 3.1) * mm, "ref")

def Parameter(canvas, wb, Nop):   # Таблица характеристик

            xmove = 0
            ymove = 0

            canvas.setFont('TimesDj', 10)
            canvas.drawString((86.5 + xmove) * mm, (189.5 + ymove) * mm, "ХАРАКТЕРИСТИКИ ГРУНТА")
            canvas.drawString((87 + xmove) * mm, (168 + ymove) * mm, "РЕЗУЛЬТАТЫ ИСПЫТАНИЯ")

            # Характеристики
            canvas.setLineWidth(0.4 * mm)
            for i in frange(25, 202.5, 17.9):  # рисуем набор прямоугольников
                canvas.rect(i * mm, (176 + ymove) * mm, 13.9 * mm, 10 * mm)
                canvas.line(i * mm, (181 + ymove) * mm, (13.9 + i) * mm, (181 + ymove) * mm)

            canvas.setFont('Times', 10)


            h = 182.5 + ymove

            scr(canvas, 10, 26, h, 'ρ', ', г/см', 's', '3', -1)
            scr(canvas, 10, 43.9, h, 'ρ', ', г/см', '', '3', -1)
            scr(canvas, 10, 61.8, h, 'ρ', ', г/см', 'd', '3', -1)
            canvas.drawString(83 * mm, h * mm, "n, %")
            canvas.drawString(100 * mm, h * mm, "e, ед.")
            canvas.drawString(118 * mm, h * mm, "W, %")
            scr(canvas, 10, 135, h, 'S', ', д.е.', 'r', 0, -1)
            # canvas.drawString(135 * mm, h * mm, "Sr, д.е.")
            scr(canvas, 10, 153.5, h, 'I', ' , %', 'P  ', 0, -1)
            scr(canvas, 10, 170.5, h, 'I', ', д.е.', 'L  ', 0, -1)
            scr(canvas, 10, 189, h, 'I', ' , %', 'r  ', 0, -1)

            datHa = [[zap(str1(wb["Лист1"]['P' + str(6 + Nop)].value).replace(".", ","), 2),
                      zap(str1(wb["Лист1"]['Q' + str(6 + Nop)].value).replace(".", ","), 2),
                      zap(str1(wb["Лист1"]['R' + str(6 + Nop)].value).replace(".", ","), 2),
                      zap(str1(wb["Лист1"]['S' + str(6 + Nop)].value).replace(".", ","), 1),
                      zap(str1(wb["Лист1"]['T' + str(6 + Nop)].value).replace(".", ","), 2),
                      zap(str1(wb["Лист1"]['U' + str(6 + Nop)].value).replace(".", ","), 1),
                      zap(str1(wb["Лист1"]['V' + str(6 + Nop)].value).replace(".", ","), 2),
                      zap(str1(wb["Лист1"]['Y' + str(6 + Nop)].value).replace(".", ","), 1),
                      zap(str1(wb["Лист1"]['Z' + str(6 + Nop)].value).replace(".", ","), 2),
                      zap(str1(wb["Лист1"]['AE' + str(6 + Nop)].value).replace(".", ","), 1)]]
            t = Table(datHa, colWidths=17.9 * mm, rowHeights=5 * mm)
            t.setStyle([("FONTNAME", (0, 0), (-1, -1), 'Times'),
                        ("FONTSIZE", (0, 0), (-1, -1), 10),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER")])
            t.wrapOn(canvas, 0, 0)
            t.drawOn(canvas, 23 * mm, (176 + ymove) * mm)



def RCReport(p1, p2, Gam, G, y, A, Chastota, Nop, Res, Akred, path):  # p1 - папка сохранения отчета, p2-путь к файлу XL, Nop - номер опыта

    # Res подается как обычный массив чисел. Состоит из [G0, Gam07]
    # p1 - папка сохранения отчета
    # p2-путь к файлу XL
    # Nop - номер опыта
    # path - папка с файлами для отчета
    # Akred может быть "AN" - ОАО Новая, "AS" - ОАО Старая, "ON" - ООО Новая, "OS" - ООО Старая

    try:

        def Pictures(p1):  # Строим графики и сохраняем картинки с них
            fig1 = Figure()
            fig2 = Figure()
            ax_1 = fig1.add_subplot(1, 1, 1)
            ax_2 = fig2.add_subplot(1, 1, 1)

            ax_2.set_xscale('log')
            ax_2.grid(axis='both', linewidth='0.4')
            ax_2.set_ylim(0.8 * min(G), 1.05 * max(G))
            ax_2.set_xlim(0.0000005, 0.0015)
            ax_2.set_xlabel("Деформация сдвига γ, д.е.", fontfamily='Times New Roman', fontsize=14)  # ось абсцисс
            ax_2.set_ylabel("Модуль сдвига G, МПа", fontfamily='Times New Roman', fontsize=14)  # ось ординат9
            ax_2.scatter(Gam, G,  label='Опытные данные')
            ax_2.plot(Gam, y, color='darkorange', label='Кривая Гардина-Дрневича')

            ax_2.legend()
            ax_1.grid(axis='both', linewidth='0.4')
            ax_1.set_xlabel("Частота f, Гц", fontfamily='Times New Roman', fontsize=14)
            ax_1.set_ylabel("Деформация сдвига γ, д.е.", fontfamily='Times New Roman', fontsize=14)

            for i in range(len(A)):
                ax_1.plot(Chastota[i], A[i])
                ax_1.scatter(Chastota[i], A[i], s=5)

            fig1.savefig(os.path.join(p1, "1.svg"), transparent=True)
            fig2.savefig(os.path.join(p1, "2.svg"), transparent=True)
            pPic1 = os.path.join(p1, "1.svg")
            pPic2 = os.path.join(p1, "2.svg")

            return pPic1, pPic2

        def Results(Res):

            hr = 87
            for i in frange(25, 125, 87.5):
                canvas.rect(i * mm, (hr - 6.5) * mm, 87.5 * mm, 10 * mm)
                canvas.line(i * mm, (hr - 1.5) * mm, (87.5 + i) * mm, (hr - 1.5) * mm)

            canvas.drawString(26 * mm, (hr) * mm, "Модуль сдвига при сверхмалых деформациях")
            canvas.drawString(26 * mm, (hr - 5) * mm, "Пороговое значение сдвиговой деформации")
            scr(canvas,10, 96.5, hr, "G", ' , МПа:', ' 0', '', -1)
            scr(canvas,10, 96.5, (hr - 5), "γ", '', '0.7', '', -1)
            canvas.drawString(102 * mm, (hr - 5) * mm, ", д.е.:")
            scr(canvas, 10, 156, (hr - 5.5), "*10", '', ' -4', '', 1)
            rrr = [["", zap(str(round(Res[0],1)).replace(".",","),1)], ["", zap(str(round(Res[1],2)).replace(".",","),2)]]

            t = Table(rrr, colWidths=87.5 * mm, rowHeights=5 * mm)
            t.setStyle([("FONTNAME", (0, 0), (-1, -1), 'Times'),
                        ("FONTSIZE", (0, 0), (-1, -1), 10),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (1, 1), (1, 1), -4.5 * mm),  # второй столбец
                        ("ALIGN", (0, 0), (-1, -1), "CENTER")])
            t.wrapOn(canvas, 0, 0)
            t.drawOn(canvas, 25 * mm, (hr - 6.5) * mm)



        # Подгружаем шрифты
        pdfmetrics.registerFont(TTFont('Times', path + 'Python(data)/Times.ttf'))
        pdfmetrics.registerFont(TTFont('TimesK', path + 'Python(data)/TimesK.ttf'))
        pdfmetrics.registerFont(TTFont('TimesDj', path + 'Python(data)/TimesDj.ttf'))



        # Загружаем документ эксель, проверяем изменялось ли имя документа и создаем отчет
        wb = load_workbook(p2)

        if str1(wb["Лист1"]['IG' + str1(6 + Nop)].value) == "-":
            Vr = str1(wb["Лист1"]['A' + str1(6 + Nop)].value)
        else:
            Vr = str1(wb["Лист1"]['IG' + str1(6 + Nop)].value)

        if Vr != "-":
            Name = "Отчет " + Vr + "-РК" + ".pdf"
        else:
            Name = "Отчет.pdf"

        canvas = Canvas(os.path.join(p1, Name), pagesize=A4)



        # Заполняем лист
        chislo = str(wb["Лист1"]["Q1"].value)                       # Дата создания отчета
        List="1/1"                                                  # Номер листа
        pPic1, pPic2 = Pictures(p1)                                 # Строим графики и сохраняем картинки с них
        c = Frame(canvas, chislo, path, p1, Akred, List, Res, 0)    # Рамка, шапка, Защитный код, Исполнители и аккредитация из файла
        TopDate(canvas, wb, Nop)                                    # Верхняя таблица данных
        Parameter(canvas, wb, Nop)                                  # Таблица характеристик
        Results(Res)                                                # Таблица результатов



        # Подгружаем картинки в отчет
        drawing1 = svg2rlg(pPic1)
        drawing1.scale(0.48, 0.48)
        renderPDF.draw(drawing1, canvas, 25 * mm, 94 * mm)

        drawing2 = svg2rlg(pPic2)
        drawing2.scale(0.48, 0.48)
        renderPDF.draw(drawing2, canvas, 112 * mm, 94 * mm)



        # Сохраняем документ
        canvas.showPage()
        canvas.save()



        # Удаляем картинки
        os.remove(os.path.join(p1, "1.svg"))
        os.remove(os.path.join(p1, "2.svg"))


    except ValueError:
        pass

def WillieReport(p1, p2, x, PPRf, Epsf, NormOf, TanOf, N, Nop, Res, Akred, path):

    # Res подается как обычный массив чисел. Состоит из [sigma3, sigma1, ta, PPRMax, MaxEps, Nc, N, I, M, MSR, rd]
    # p1 - папка сохранения отчета
    # p2-путь к файлу XL
    # Nop - номер опыта
    # path - папка с файлами для отчета
    # Akred может быть "AN" - ОАО Новая, "AS" - ОАО Старая, "ON" - ООО Новая, "OS" - ООО Старая

    try:

        def Pictures(p1):  # Строим графики и сохраняем картинки с них

            fig1 = Figure()
            fig2 = Figure()
            fig3 = Figure()
            ax_1 = fig1.add_subplot(1, 1, 1)
            ax_2 = fig2.add_subplot(1, 1, 1)
            ax_3 = fig3.add_subplot(1, 1, 1)

            ax_1.set_xlabel('Количество циклов нагружения', fontfamily='Times New Roman', fontsize=16)
            ax_1.set_ylabel('Относительное поровое давление', fontfamily='Times New Roman', fontsize=16)

            ax_2.set_xlabel('Количество циклов нагружения', fontfamily='Times New Roman', fontsize=16)
            ax_2.set_ylabel('Относительная деформация, ε', fontfamily='Times New Roman', fontsize=16)

            ax_3.set_xlabel("Тангенциальное октаэдрическое напряжение, р' (кПа)", fontfamily='Times New Roman',fontsize=12)
            ax_3.set_ylabel("Нормальное октаэдрическое напряжение, q' (кПа)", fontfamily='Times New Roman', fontsize=12)


            ax_1.plot(x, PPRf)  # Деформация от девиатора
            ax_1.plot([0, N], [1, 1], linestyle='--', color='red')
            ax_2.plot(x, Epsf)  # Объемная деформация
            ax_3.plot(NormOf, TanOf)

            fig1.savefig(os.path.join(p1, "1.svg"), transparent=True)
            fig2.savefig(os.path.join(p1, "2.svg"), transparent=True)
            fig3.savefig(os.path.join(p1, "3.svg"), transparent=True)

            pPic1 = os.path.join(p1, "1.svg")
            pPic2 = os.path.join(p1, "2.svg")
            pPic3 = os.path.join(p1, "3.svg")


            return pPic1, pPic2, pPic3

        def Results(Res):

            Resrep=[[str(round(Res[0])), str(round(Res[1])), str(round(Res[2])),
                     zap(str(round(Res[3],3)).replace(".",","),3), zap(str(round(Res[4],3)).replace(".",","),3),
                     str(round(Res[5])),
                     zap(str(round(Res[6],1)).replace(".",","),1), zap(str(round(Res[7],1)).replace(".",","),1), zap(str(round(Res[8],1)).replace(".",","),1),
                     zap(str(round(Res[9],2)).replace(".",","),2), zap(str(round(Res[10],2)).replace(".",","),2)]]

            hr = 154
            for i in frange(25, 190, 16.175):
                canvas.rect(i * mm, hr * mm, 13.175 * mm, 10 * mm)
                canvas.line(i * mm, (hr+5) * mm, (13.175 + i) * mm, (hr+5) * mm)

            scr(canvas, 10, 26, (hr+6.5), "σ", ', кПа', '3', '', -1)
            canvas.drawString(28 * mm, (hr+6.5) * mm, "'")
            scr(canvas, 10, 42.675, (hr+6.5), "σ", ', кПа', '1', '', -1)
            canvas.drawString(44.675 * mm, (hr+6.5) * mm, "'")
            scr(canvas, 10, 58.6, (hr+6.5), "τ", ', кПа', 'a', '', -1)
            scr(canvas, 10, 75, (hr+6.5), "PPR", '', '  max', '', -1)
            scr(canvas, 10, 93.5, (hr+6.5), "ε", '', 'max', '', -1)
            canvas.drawString(110.5 * mm, (hr+6.5) * mm, "Nc")
            canvas.drawString(125 * mm, (hr+6.5) * mm, "μ, Гц")
            canvas.drawString(140 * mm, (hr+6.5) * mm, "I, балл")
            canvas.drawString(156.5 * mm, (hr+6.5) * mm, "M, ед.")
            canvas.drawString(171 * mm, (hr+6.5) * mm, "MSF, ед.")
            canvas.drawString(189.5 * mm, (hr+6.5) * mm, "rd, ед.")

            t = Table(Resrep, colWidths=16.175 * mm, rowHeights=5 * mm)
            t.setStyle([("FONTNAME", (0, 0), (-1, -1), 'Times'),
                        ("FONTSIZE", (0, 0), (-1, -1), 10),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER")])
            t.wrapOn(canvas, 0, 0)
            t.drawOn(canvas, 23.5 * mm, hr * mm)



        # Подгружаем шрифты
        pdfmetrics.registerFont(TTFont('Times', path + 'Python(data)/Times.ttf'))
        pdfmetrics.registerFont(TTFont('TimesK', path + 'Python(data)/TimesK.ttf'))
        pdfmetrics.registerFont(TTFont('TimesDj', path + 'Python(data)/TimesDj.ttf'))



        # Загружаем документ эксель, проверяем изменялось ли имя документа и создаем отчет
        wb = load_workbook(p2)

        if str1(wb["Лист1"]['IG' + str1(6 + Nop)].value) == "-":
            Vr = str1(wb["Лист1"]['A' + str1(6 + Nop)].value)
        else:
            Vr = str1(wb["Лист1"]['IG' + str1(6 + Nop)].value)

        if Vr != "-":
            Name = "Отчет " + Vr + "-РК" + ".pdf"
        else:
            Name = "Отчет.pdf"

        canvas = Canvas(os.path.join(p1, Name), pagesize=A4)



        # Заполняем лист
        chislo = str(wb["Лист1"]["Q1"].value)                      # Дата создания отчета
        List="1/2"                                                 # Номер листа
        pPic1, pPic2, pPic3 = Pictures(p1)                         # Строим графики и сохраняем картинки с них
        c = Frame(canvas, chislo, path, p1, Akred, List, Res, 0)   # Рамка, шапка, Защитный код, Исполнители и аккредитация из файла
        TopDate(canvas, wb, Nop)                                   # Верхняя таблица данных
        Parameter(canvas, wb, Nop)                                 # Таблица характеристик
        Results(Res)                                               # Таблица результатов


        # Подгружаем картинки в отчет
        drawing1 = svg2rlg(pPic1)
        drawing1.scale(0.48, 0.48)
        renderPDF.draw(drawing1, canvas, 21 * mm, 80 * mm)
        drawing2 = svg2rlg(pPic2)
        drawing2.scale(0.48, 0.48)
        renderPDF.draw(drawing2, canvas, 112 * mm, 80 * mm)



        # Вторая страница
        canvas.showPage()
        Page = 1
        List="2/2"                                                  # Номер листа
        c = Frame(canvas, chislo, path, p1, Akred, List, Res, c)    # Рамка, шапка, Защитный код, Исполнители и аккредитация из файла
        TopDate(canvas, wb, Nop)                                    # Верхняя таблица данных
        Parameter(canvas, wb, Nop)                                  # Таблица характеристик
        Results(Res)                                                # Таблица результатов



        # Подгружаем картинки в отчет
        drawing3 = svg2rlg(pPic3)
        drawing3.scale(0.6, 0.6)
        renderPDF.draw(drawing3, canvas, 55 * mm, 65 * mm)


        canvas.save()



        # Удаляем картинки
        os.remove(os.path.join(p1, "1.svg"))
        os.remove(os.path.join(p1, "2.svg"))
        os.remove(os.path.join(p1, "3.svg"))


    except ValueError:
        pass


if __name__ == '__main__':
    path1 = "C:/Users/Пользователь/Desktop/"
    path2 = "//192.168.0.1/files/Прикладные программы/"
    path3 = "Z:/files/Прикладные программы/"
    path = path1

    Akred = "AN"
    p1 = "C:/Users/Пользователь/Desktop/Новая папка (2)/Новая папка (2)"
    p2 = "C:/Users/Пользователь/Desktop/Новая папка (2)/810-19 Нансена - волновое.xlsx"

    x = [0, 0.05, 0.1, 0.15, 0.2]
    N = 3
    Epsf = [0, 0.00865341, 0.01664779, 0.02730129, 0.03229356 ]
    PPRf = [0, 0.00865341, 0.01664779, 0.02730129, 0.03229356 ]
    NormOf = [2.76610427, 2.93620221, 3.16389129, 3.33484985, 3.44312753 ]
    TanOf = [1.21395641,  1.4945028,   1.85950154, 2.14720953,  2.32427936]

    Nop = 2
    #Res = [[sigma3Re.get().replace(".", ","), sigma1Re.get().replace(".", ","), tRe.get().replace(".", ","),
                #PPRMaxRe.get().replace(".", ","), MaxEpsRe.get().replace(".", ","), NcRe.get().replace(".", ","),
                #NuRe.get().replace(".", ","), I.replace(".", ","), M, MSR,
                #rd.replace(".", ",")]]
    Res = [23.7, 23.7, 23.7, 23.7, 23.7345, 23.7, 23.7, 23, 23.7345, 23.7, 23.7345]

    WillieReport(p1, p2, x, PPRf, Epsf, NormOf, TanOf, N, Nop, Res, Akred, path)
