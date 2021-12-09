#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
VERSION: 1.1 of 2021-12-09
AUTHOR: Rafferty River. 
LICENSE: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007. 
This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY.

DESCRIPTION & USAGE:
(tested on Scribus 1.5.7 with python 3 on Windows 10).
This Scribus script generates a 12-months calendar on one page with following options:
1) You can choose between more than 20 languages (default is English). 
You may add, change or delete languages in the localization list in this script.
Please respect the Python syntax.
2) You can choose your font from the list of fonts available on your system. 
Please check if all special characters for your language are available in the chosen font! 
You can change fonts of many items afterwards in the Styles menu (Edit - Styles).
3) Calendar year, start month and week starting day are to be given. Saturdays and Sundays will 
be printed in separate colors (many colors can be changed afterwards with Edit - Colors and Fills).
4) Option to show week numbers with (or without) a week numbers heading in 
your local language. Calendar week numbers will be printed in dark gray.
5) Option to import holidays, special days and vacation dates from a 'holidays.txt' file. 
The 'holidays.txt' file from MonthlyCalendar script can be used here. See the example
holiday.txt-file for the layout. 
Automatic calculation of the holiday dates for each calendar year. 
6) Number of months per row determines the layout of the 12-month calendar.
7) You can position the 12-month calendar within your document.
8) Option to draw an empty image frame within the top and / or left 'offset' area and to get an 
'inner' margin between this frame and the calendar.
9) Option to generate a text frame at the bottom with the holiday dates and texts.
10) You can easily change the text styles, colors and fills of month title, weekday names, 
week numbers, weekends, holidays, normal dates, special dates, vacation and grids.
Many build-in controls.

Parts of this script are taken from the MonthlyCalendar script for Scribus.
"""
######################################################
# imports
from __future__ import division # overrules Python 2 integer division
import sys
import locale
import calendar
import datetime
from datetime import date, timedelta
import csv
import platform

try:
    from scribus import *
except ImportError:
    print("This Python script is written for the Scribus \
      scripting interface.")
    print("It can only be run from within Scribus.")
    sys.exit(1)

os = platform.system()
if os != "Windows" and os != "Linux":
    print("Your Operating System is not supported by this script.")
    messageBox("Script failed",
        "Your Operating System is not supported by this script.",
        ICON_CRITICAL)	
    sys.exit(1)

python_version = platform.python_version()
if python_version[0:1] != "3":
    print("This script runs only with Python 3.")
    messageBox("Script failed",
        "This script runs only with Python 3.",
        ICON_CRITICAL)	
    sys.exit(1)

try:
    from tkinter import * # python 3
    from tkinter import messagebox, filedialog, font
except ImportError:
    print("This script requires Python Tkinter properly installed.")
    messageBox('Script failed',
               'This script requires Python Tkinter properly installed.',
               ICON_CRITICAL)
    sys.exit(1)

######################################################
# you can insert additional languages and unicode pages in the 'localization'-list below:
localization = [['Bulgarian', 'CP1251', 'bg_BG.UTF8'], 
    ['Croatian', 'CP1250', 'hr_HR.UTF8'], 
    ['Czech', 'CP1250', 'cs_CZ.UTF8'], 
    ['Danish', 'CP1252','da_DK.UTF8'],
    ['Dutch', 'CP1252', 'nl_NL.UTF8'], 
    ['English', 'CP1252', 'en_US.UTF8'], 
    ['Estonian', 'CP1257', 'et_EE.UTF8'], 
    ['Finnish', 'CP1252', 'fi_FI.UTF8'], 
    ['French', 'CP1252', 'fr_FR.UTF8'], 
    ['German', 'CP1252', 'de_DE.UTF8'], 
    ['German_Austria', 'CP1252', 'de_AT.UTF8'], 
    ['Greek', 'CP1253', 'el_GR.UTF8'], 
    ['Hungarian', 'CP1250', 'hu_HU.UTF8'] ,
    ['Italian', 'CP1252', 'it_IT.UTF8'],
    ['Lithuanian', 'CP1257', 'lt_LT.UTF8'], 
    ['Latvian', 'CP1257', 'lv_LV.UTF8'],
    ['Norwegian', 'CP1252', 'nn_NO.UTF8'],
    ['Polish', 'CP1250', 'pl_PL.UTF8'], 
    ['Portuguese', 'CP1252', 'pt_PT.UTF8'],
    ['Romanian', 'CP1250', 'ro_RO.UTF8'],
    ['Russian', 'CP1251', 'ru_RU.UTF8'], 
    ['Slovak', 'CP1250', 'sk_SK.UTF8'],
    ['Slovenian', 'CP1250', 'sl_SI.UTF8'],
    ['Spanish', 'CP1252', 'es_ES.UTF8'], 
    ['Swedish', 'CP1252', 'sv_SE.UTF8']]

######################################################
class ScYearCalendar:
    """ Calendar matrix creator itself. """

    def __init__(self, year, months = [], nrHmonths = 0, firstDay = calendar.SUNDAY,
                weekNr=True, weekNrHd="Wk", offsetX=0.0, marginX=0.0, offsetY=0.0,
                marginY=0.0, drawImg=True,  drawLegend=True, cFont='Symbola Regular',
                lang='English', holidaysList = list()):
        """ Setup basic things """
        # params
        self.year = year
        self.months = months
        self.nrHmonths = nrHmonths     # number of horizontal months on page
        self.nrVmonths = ((12 // self.nrHmonths)+ (12 % self.nrHmonths > 0))
            # number of vertical months: if there is no remainder, then it stays the same integer,
            #       but if there is a remainder it adds 1
        self.weekNr = weekNr
        self.weekNrHd = weekNrHd #week numbers heading
        self.offsetX = offsetX
        self.offsetY = offsetY
        self.marginX = marginX
        self.marginY = marginY
        self.drawImg = drawImg # draw placeholder for image or logo (between margins and offsetX / offsetY)
        self.drawLegend = drawLegend # create text frame with holiday texts at bottom or at right side
        self.holidaysList = holidaysList #imported and converted from '*holidays.txt' (or empty list)
        if len(self.holidaysList) != 0:
            self.drawHolidays = True
        else:
            self.drawHolidays = False
        self.cFont = cFont
        self.lang = lang
        ix = [[x[0] for x in localization].index(self.lang)]
        if os == "Windows":
            self.calUniCode = (localization[ix[0]][1]) # get unicode page for the selected language
        else: # Linux
            self.calUniCode = "UTF-8"
        self.dayOrder=[] # first letter of weekday names in local language
        for i in range (0,7):
            try:            
                self.dayOrder.append((calendar.day_abbr[i][:1]).upper())
            except UnicodeError: # for Greek, Russian, etc.
                self.dayOrder.append((calendar.day_abbr[i][:2]).upper())
        if firstDay == calendar.SUNDAY:
            dl = self.dayOrder[:6]
            dl.insert(0, self.dayOrder[6])
            self.dayOrder = dl
        self.mycal = calendar.Calendar(firstDay)
        # layers
        self.layerCal = 'Calendar'
        # character styles
        self.cStylMonthHeading = "char_style_MonthHeading"
        self.cStylDayNames = "char_style_DayNames"
        self.cStylWeekNo = "char_style_WeekNo"
        self.cStylHolidays = "char_style_Holidays"
        self.cStylDate = "char_style_Date"
        self.cStylLegend = "char_style_Legend"
        # paragraph styles
        self.pStyleMonthHeading = "par_style_MonthHeading"
        self.pStyleDayNames = "par_style_DayNames"
        self.pStyleWeekNo = "par_style_WeekNo"
        self.pStyleHolidays = "par_style_Holidays"
        self.pStyleDate = "par_style_Date"
        self.pStyleLegend = "par_style_Legend"
        # line styles
        self.gridLineStyle = "grid_Line_Style"
        self.gridLineStyleDayNames = "grid_DayNames_Style"
        self.gridLineStyleWeekNo = "grid_WeekNo_Style"
        self.gridLineStyleMonthHeading = "grid_MonthHeading_Style"
        # other settings
        calendar.setfirstweekday(firstDay)
        progressTotal(len(months))

    def createCalendar(self):
        """ Walk through months """
        if not newDocDialog():
            return 'Create a new document'
        originalUnit = getUnit()
        setUnit(UNIT_POINTS)
        self.setupDocVariables()
        setActiveLayer(self.layerCal)
        run = 0
        nextYr = False
        self.nrHmthsCnt = self.nrHmonths # counter for number of horizontal months
        nrVmthsCnt = 0 # counter for number of vertical months
        for i in self.months: # loop for creating the months
            run += 1
            progressSet(run)
            j = self.months[0] + run
            if (j > 13) and (nextYr == False):
                self.year = self.year + 1
                nextYr = True # self.year is meanwhile next year
            if self.nrHmthsCnt == self.nrHmonths:
                 self.rowCnt = nrVmthsCnt * 9
                 nrVmthsCnt += 1
                 self.nrHmthsCnt = 0
            else:
                 self.rowCnt = (nrVmthsCnt - 1) * 9
            self.colCnt = self.nrHmthsCnt * (self.mthcols + 1)
            cal = self.mycal.monthdatescalendar(self.year, i)
            self.createMonthCalendar(i, cal, self.rowCnt, self.colCnt)
            self.nrHmthsCnt += 1
        if self.drawLegend:
            self.createLegend()
        setUnit(originalUnit)
        return None

    def setupDocVariables(self):
        """ Compute base metrics here. Page layout is bordered by margins
            and empty image frame(s). """
        # page dimensions
        page = getPageSize()
        self.pageX = page[0]
        self.pageY = page[1]
        marg = getPageMargins()
        self.marginT = marg[0]
        self.marginL = marg[1]
        self.marginR = marg[2]
        self.marginB = marg[3]
        self.width = self.pageX - self.marginL - self.marginR
        self.height = self.pageY - self.marginT - self.marginB
        # month cell rows and cols
        self.rows = 8 # month heading + weekday names +  6 weeks per month
        self.rows = (self.rows * self.nrVmonths) + (self.nrVmonths - 1)
            # add 1 row space between the month per column
        if self.drawLegend:  # create text frame with holiday texts at the bottom
            y = 0
            for x in range(len(self.holidaysList)):
                if len(self.holidaysList[x][3]) > 0:  # if there is a text
                    y += 1
            if y / self.nrHmonths == y // self.nrHmonths:
                y = y // self.nrHmonths
            else:
                y = y // self.nrHmonths + 1
            self.rows = self.rows + (y+2) * 0.6
        self.rowSize = (self.height - self.offsetY) / self.rows
        if self.weekNr:
            self.mthcols = 8 # weekNr column + 7 weekdays per month
        else:
            self.mthcols = 7 # 7 weekdays columns per month
        self.cols = (self.mthcols * self.nrHmonths) + (self.nrHmonths - 1)
            # add 1 column space between the months per row
        self.colSize = (self.width - self.offsetX) / self.cols
        baseLine = self.rowSize
        h = (self.marginT + self.offsetY)
        x =  h/baseLine - h//baseLine
        y = x * baseLine + baseLine * 0.75
        setBaseLine(baseLine, y) # for correct aligment of weekdays names
                                                      #  with ascender and descender characters
        # default calendar colors
        defineColorCMYK("Black", 0, 0, 0, 255)
        defineColorCMYK("White", 0, 0, 0, 0)
        defineColorCMYK("fillMonthHeading", 0, 0, 0, 0) # default is White
        defineColorCMYK("txtMonthHeading", 0, 0, 0, 255) # default is Black
        defineColorCMYK("fillDayNames", 0, 0, 0, 200) # default is Dark Grey
        defineColorCMYK("txtDayNames", 0, 0, 0, 0) # default is White
        defineColorCMYK("fillWeekNo", 0, 0, 0, 200) # default is Dark Grey
        defineColorCMYK("txtWeekNo", 0, 0, 0, 0) # default is White
        defineColorCMYK("fillDate", 0, 0, 0, 0) # default is White
        defineColorCMYK("txtDate", 0, 0, 0, 255) # default is Black
        defineColorCMYK("fillWeekend", 0, 0, 0, 25) # default is Light Grey
        defineColorCMYK("fillWeekend2", 0, 0, 0, 25) # default is Light Grey
        defineColorCMYK("txtWeekend", 0, 0, 0, 200) # default is Dark Grey
        defineColorCMYK("fillHoliday", 0, 0, 0, 25) # default is Light Grey
        defineColorCMYK("txtHoliday", 0, 234, 246, 0) # default is Red
        defineColorCMYK("fillSpecialDate", 0, 0, 0, 0) # default is White
        defineColorCMYK("txtSpecialDate", 0, 0, 0, 128) # default is Middle Grey
        defineColorCMYK("fillVacation", 0, 0, 0, 25) # default is Light Grey
        defineColorCMYK("txtVacation", 0, 0, 0, 255) # default is Black
        defineColorCMYK("gridColor", 0, 0, 0, 128) # default is Middle Grey
        defineColorCMYK("gridMonthHeading", 0, 0, 0, 0) # default is White
        defineColorCMYK("gridDayNames", 0, 0, 0, 128) # default is Middle Grey
        defineColorCMYK("gridWeekNo", 0, 0, 0, 128) # default is Middle Grey
        # styles
        scribus.createCharStyle(name=self.cStylMonthHeading, font=self.cFont,
            fontsize=(self.rowSize // 1.5), fillcolor="txtMonthHeading")
        scribus.createCharStyle(name=self.cStylDayNames, font=self.cFont,
            fontsize=(self.rowSize // 2), fillcolor="txtDayNames")
        scribus.createCharStyle(name=self.cStylWeekNo, font=self.cFont,
            fontsize=(self.rowSize // 2), fillcolor="txtWeekNo")
        scribus.createCharStyle(name=self.cStylHolidays, font=self.cFont,
            fontsize=(self.rowSize // 2), fillcolor="txtHoliday")
        scribus.createCharStyle(name=self.cStylDate, font=self.cFont,
            fontsize=(self.rowSize // 2), fillcolor="txtDate")
        scribus.createCharStyle(name=self.cStylLegend, font=self.cFont,
            fontsize=(self.rowSize // 2), fillcolor="txtDate")
        scribus.createParagraphStyle(name=self.pStyleMonthHeading, linespacingmode=2,
            alignment=ALIGN_CENTERED, charstyle=self.cStylMonthHeading)
        scribus.createParagraphStyle(name=self.pStyleDayNames, linespacingmode=2,
            alignment=ALIGN_CENTERED, charstyle=self.cStylDayNames)
        scribus.createParagraphStyle(name=self.pStyleWeekNo,  linespacingmode=2,
            alignment=ALIGN_CENTERED, charstyle=self.cStylWeekNo)
        scribus.createParagraphStyle(name=self.pStyleHolidays, linespacingmode=2,
            alignment=ALIGN_CENTERED, charstyle=self.cStylHolidays)
        scribus.createParagraphStyle(name=self.pStyleDate, linespacingmode=2,
            alignment=ALIGN_CENTERED, charstyle=self.cStylDate)
        scribus.createParagraphStyle(name=self.pStyleLegend,  linespacingmode=0,
            linespacing=(self.rowSize *0.6), alignment=ALIGN_LEFT, 
            charstyle=self.cStylLegend)
        scribus.createCustomLineStyle(self.gridLineStyle, [
            {
                'Color': "gridColor",
                'Width': 0.25
            }
        ]);
        scribus.createCustomLineStyle(self.gridLineStyleMonthHeading, [
            {
                'Color': "gridMonthHeading",
                'Width': 0.25
            }
        ]);
        scribus.createCustomLineStyle(self.gridLineStyleDayNames, [
            {
                'Color': "gridDayNames",
                'Width': 0.25
            }
        ]);
        scribus.createCustomLineStyle(self.gridLineStyleWeekNo, [
            {
                'Color': "gridWeekNo",
                'Width': 0.25
            }
        ]);
        # layers
        createLayer(self.layerCal)
        if self.drawImg:
            self.createImg()

    def createImg(self):
        """ Create Image frame(s). """
        if self.offsetX != 0:
            createImage(self.marginL, self.marginT, self.offsetX - self.marginX, self.height)
        if self.offsetY != 0: # if top AND left frame -> top frame does not overlap with left frame
            createImage(self.marginL + self.offsetX, self.marginT,
                self.width-self.offsetX, self.offsetY - self.marginY)

    def createLegend(self):
        """ Create text frame at the bottom of the page. """
        self.rowCnt += 2
        x = 1
        if self.weekNr and self.months[0]==1:  # indent legend texts if no year printed
             self.offsetX = self.offsetX + (self.colSize)
             x = 2
        cel = createText(self.marginL + self.offsetX,
                                 self.marginT + self.offsetY + self.rowCnt * self.rowSize, 
                                 self.width - self.offsetX, 
                                 self.height -self.offsetY - self.rowCnt * self.rowSize)
        setColumns(self.nrHmonths, cel)
        setColumnGap(self.colSize * x, cel)
        deselectAll()
        selectObject(cel)
        for x in range(len(self.holidaysList)):
            if len(self.holidaysList[x][3]) > 0:  # if there is a text
                txtHoliday = (("0" if len(self.holidaysList[x][2]) == 1 else "") + self.holidaysList[x][2]
                + "/" + ("0" if len(self.holidaysList[x][1]) == 1 else "") + self.holidaysList[x][1]
                + (("/" + str(self.holidaysList[x][0])) if self.months[0] != 1 else "")
                + " " + self.holidaysList[x][3] + "\n")
                insertText(txtHoliday, -1, cel)
            setParagraphStyle(self.pStyleLegend, cel)
    def createMonthCalendar(self, month, cal, rowCnt, colCnt):
        """ Draw one month calendar """
        self.rowCnt = rowCnt
        self.colCnt = colCnt
        self.createMonthHeader(calendar.month_name[month], self.rowCnt, self.colCnt)
        for week in cal:
            if self.weekNr:
                cel = createText(self.marginL + self.offsetX + self.colCnt * self.colSize,
                                 self.marginT + self.offsetY + self.rowCnt * self.rowSize, 
                                 self.colSize, self.rowSize)
                yr, mt, dt = str((week[0])).split("-")
                setText(str(datetime.date(int(yr), int(mt), int(dt)).isocalendar()[1]), cel)
                setFillColor("fillWeekNo", cel)
                setCustomLineStyle(self.gridLineStyleWeekNo, cel)
                deselectAll()
                selectObject(cel)
                setParagraphStyle(self.pStyleWeekNo, cel)
                setTextVerticalAlignment(ALIGNV_TOP,cel)
                self.colCnt += 1
            for day in week:
                cel = createText(self.marginL + self.offsetX + self.colCnt * self.colSize,
                    self.marginT + self.offsetY + self.rowSize * self.rowCnt,
                    self.colSize, self.rowSize)
                setFillColor("fillDate", cel)
                setCustomLineStyle(self.gridLineStyle, cel)
                if day.month == month:
                    setText(str(day.day), cel)
                    deselectAll()
                    selectObject(cel)
                    setParagraphStyle(self.pStyleDate, cel)
                    setTextVerticalAlignment(ALIGNV_TOP,cel)
                    if calendar.firstweekday() == 6: # weekend day
                        x = 1
                    else:
                        x = 6
                    y = 8 - self.mthcols
                    if (int(self.colCnt - self.nrHmthsCnt * (self.mthcols + 1) + y) == x or 
                        int(self.colCnt - self.nrHmthsCnt * (self.mthcols + 1) + y) == 7):
                            setTextColor("txtWeekend", cel)
                            setFillColor("fillWeekend", cel)
                    for x in range(len(self.holidaysList)): # holiday
                        if (self.holidaysList[x][0] == (day.year) and
                                self.holidaysList[x][1] == str(day.month) and
                                self.holidaysList[x][2] == str(day.day)):
                            if self.holidaysList[x][4] == "":
                                if getFillColor(cel) != "fillWeekend":
                                    setTextColor("txtVacation", cel)
                                    setFillColor("fillVacation", cel)
                            elif self.holidaysList[x][4] == '0':
                                setTextColor("txtSpecialDate", cel)
                                if getFillColor(cel) == "fillWeekend":
                                    setFillColor("fillWeekend", cel)
                                elif getFillColor(cel) == "fillVacation":
                                    setFillColor("fillVacation", cel)
                                else:
                                    setFillColor("fillSpecialDate", cel)
                            else:
                                setParagraphStyle(self.pStyleHolidays, cel)
                                setTextColor("txtHoliday", cel)
                                setFillColor("fillHoliday", cel)
                else:  # fill previous or next month weekend cells
                    if calendar.firstweekday() == 6: # weekend day
                        x = 1
                    else:
                        x = 6
                    y = 8 - self.mthcols
                    if (int(self.colCnt - self.nrHmthsCnt * (self.mthcols + 1) + y) == x or 
                        int(self.colCnt - self.nrHmthsCnt * (self.mthcols + 1) + y) == 7):
                            setFillColor("fillWeekend2", cel)
                self.colCnt += 1
            self.colCnt = self.colCnt - self.mthcols
            self.rowCnt += 1
        return

    def createMonthHeader(self, monthName, rowCnt, colCnt):
        """ Draw month calendars header """
        self.rowCnt = rowCnt
        self.colCnt = colCnt
        cel = createText(self.marginL + self.offsetX + self.colCnt * self.colSize,
                self.marginT + self.offsetY + self.rowCnt * self.rowSize,
                self.colSize * self.mthcols, self.rowSize)
        mtHd = monthName
        setText(mtHd.upper() + " " + str(self.year), cel)
        setFillColor("fillMonthHeading", cel)
        setCustomLineStyle(self.gridLineStyleMonthHeading, cel)
        deselectAll()
        selectObject(cel)
        setParagraphStyle(self.pStyleMonthHeading, cel)
        setTextVerticalAlignment(ALIGNV_TOP, cel)
        self.rowCnt += 1
        if self.weekNr:
            cel = createText(self.marginL + self.offsetX+ self.colCnt * self.colSize,
                self.marginT + self.offsetY + self.rowSize * self.rowCnt, self.colSize,
                self.rowSize)
            try:
                setText(self.weekNrHd, cel)
            except:
                pass
            setFillColor("fillWeekNo", cel)
            setCustomLineStyle(self.gridLineStyleWeekNo, cel)
            deselectAll()
            selectObject(cel)
            setParagraphStyle(self.pStyleWeekNo, cel)
            setTextVerticalAlignment(ALIGNV_TOP, cel)
            self.colCnt += 1
        for j in self.dayOrder: # day names
            cel = createText(self.marginL + self.offsetX + self.colCnt * self.colSize,
                self.marginT + self.offsetY + self.rowSize * self.rowCnt,
                self.colSize, self.rowSize)
            setText(j, cel)
            setFillColor("fillDayNames", cel)
            setCustomLineStyle(self.gridLineStyleDayNames, cel)
            deselectAll()
            selectObject(cel)
            setParagraphStyle(self.pStyleDayNames, cel)
            setTextVerticalAlignment(ALIGNV_TOP, cel)
            self.colCnt+= 1
        self.rowCnt += 1
        self.colCnt = self.colCnt - self.mthcols

######################################################
class calcHolidays:
    """ Import local holidays from '*holidays.txt'-file and convert the variable
    holidays into dates for the given year."""

    def __init__(self, year):
        self.year = year

    def calcEaster(self):
        """ Calculate Easter date for the calendar Year using Butcher's Algorithm. 
        Works for any date in the Gregorian calendar (1583 and onward)."""
        a = self.year % 19
        b = self.year // 100
        c = self.year % 100
        d = (19 * a + b - b // 4 - ((b - (b + 8) // 25 + 1) // 3) + 15) % 30
        e = (32 + 2 * (b % 4) + 2 * (c // 4) - d - (c % 4)) % 7
        f = d + e - 7 * ((a + 11 * d + 22 * e) // 451) + 114	
        easter = datetime.date(int(self.year), int(f//31), int(f % 31 + 1))
        return easter

    def calcEasterO(self):
        """ Calculate Easter date for the calendar Year using Meeus Julian Algorithm. 
        Works for any date in the Gregorian calendar between 1900 and 2099."""
        d = (self.year % 19 * 19 + 15) % 30
        e = (self.year % 4 * 2 + self.year % 7 * 4 - d + 34) % 7 + d + 127
        m = e / 31
        a = e % 31 + 1 + (m>4)
        if a>30: a, m=1,5
        easter = datetime.date(int(self.year), int(m), int(a))
        return easter

    def calcVarHoliday(self, base, delta):
        """ Calculate variable Christian holidays dates for the calendar Year. 
        'base' is Easter and 'delta' the days from Easter."""
        holiday = base + timedelta(days=int(delta))
        return holiday

    def calcNthWeekdayOfMonth(self, n, weekday, month, year):
        """ Returns (month, day) tuple that represents nth weekday of month in year.
        If n==0, returns last weekday of month. Weekdays: Monday=0."""
        if not 0 <= n <= 5:
            raise IndexError("Nth day of month must be 0-5. Received: {}".format(n))
        if not 0 <= weekday <= 6:
            raise IndexError("Weekday must be 0-6")
        firstday, daysinmonth = calendar.monthrange(year, month)
        # Get first WEEKDAY of month
        first_weekday_of_kind = 1 + (weekday - firstday) % 7
        if n == 0:
        # find last weekday of kind, which is 5 if these conditions are met, else 4
            if first_weekday_of_kind in [1, 2, 3] and first_weekday_of_kind + 28 <= daysinmonth:
                n = 5
            else:
                n = 4
        day = first_weekday_of_kind + ((n - 1) * 7)
        if day > daysinmonth:
            raise IndexError("No {}th day of month {}".format(n, month))
        return (year, month, day)

    def importHolidays(self):
        """ Import local holidays from '*holidays.txt'-file."""
        holidaysFile = filedialog.askopenfilename(title="Open the \
'holidays.txt'-file or cancel")
        holidaysList=list()
        try:
            csvfile = open(holidaysFile, mode="rt",  encoding="utf8")
        except:
            print("Holidays wil NOT be shown.")
            messageBox("Warning:",
               "Holidays wil NOT be shown.",
               ICON_CRITICAL)
            return holidaysList # returns an empty holidays list
        csvReader = csv.reader(csvfile, delimiter=",")
        for row in csvReader:
            try:
                if row[0] == "fixed":
                    holidaysList.append((self.year, row[1], row[2], row[4], row[5]))
                    holidaysList.append((self.year + 1, row[1], row[2], row[4], row[5]))
                elif row[0] == "nWDOM": # nth WeekDay Of Month
                    dt=self.calcNthWeekdayOfMonth(int(row[3]), int(row[2]), int(row[1]), int(self.year))
                    holidaysList.append((self.year, str(dt[1]), str(dt[2]), row[4], row[5]))
                    dt=self.calcNthWeekdayOfMonth(int(row[3]), int(row[2]), int(row[1]), int(self.year + 1))
                    holidaysList.append((self.year + 1, str(dt[1]), str(dt[2]), row[4], row[5]))
                elif row[0] == "variable":
                    if row[1] == "easter" :
                        base=self.calcEaster()
                        dt=self.calcVarHoliday(base, int(row[2]))
                        holidaysList.append(((dt.year), str(dt.month), str(dt.day), row[4], row[5]))
                        self.year =self. year + 1
                        base=self.calcEaster()
                        dt=self.calcVarHoliday(base, int(row[2]))
                        holidaysList.append(((dt.year), str(dt.month), str(dt.day), row[4], row[5]))
                        self.year = self.year - 1
                    elif row[1] == "easterO" :
                        base=self.calcEasterO()
                        dt=self.calcVarHoliday(base, int(row[2]))
                        holidaysList.append(((dt.year), str(dt.month), str(dt.day), row[4], row[5]))
                        self.year =self. year + 1
                        base=self.calcEaster()
                        dt=self.calcVarHoliday(base, int(row[2]))
                        holidaysList.append(((dt.year), str(dt.month), str(dt.day), row[4], row[5]))
                        self.year = self.year - 1
                else:
                    pass #do nothing
            except:
                print("Not a valid Holidays file.\nHolidays wil NOT be shown.")
                messageBox("Warning:",
                    "Not a valid Holidays file.\nHolidays wil NOT be shown.",
                    ICON_CRITICAL)
                break
        csvfile.close()
        return holidaysList

######################################################
class TkCalendar(Frame):
    """ GUI interface for Scribus calendar wizard with tkinter"""

    def __init__(self, master=None):
        """ Setup the dialog """
        Frame.__init__(self, master)
        self.grid()
        self.master.resizable(0, 0)
        self.master.title('Scribus Year Calendar')

        #define widgets
        self.statusVar = StringVar()
        self.statusLabel = Label(self, fg="red", textvariable=self.statusVar)
        self.statusVar.set('Select Options and Values')

        # languages (reference to the localization dictionary)
        langX = 'English'
        self.langLabel = Label(self, text='Select language:')
        self.langFrame = Frame(self)
        self.langFrame.grid()
        self.langScrollbar = Scrollbar(self.langFrame, orient=VERTICAL)
        self.langScrollbar.grid(row=0, column=1, sticky=N+S)
        self.langListbox = Listbox(self.langFrame, selectmode=SINGLE, height=12,
            yscrollcommand=self.langScrollbar.set)
        self.langListbox.grid(row=0, column=0, sticky=N+S+E+W)
        self.langScrollbar.config(command=self.langListbox.yview)
        for i in range(len(localization)):
            self.langListbox.insert(END, localization[i][0])
        self.langButton = Button(self, text='Change language',
            command=self.languageChange)
        
        # choose font
        self.fontLabel = Label(self, text='Change font:')
        self.fontFrame = Frame(self)
        self.fontScrollbar = Scrollbar(self.fontFrame, orient=VERTICAL)
        self.fontListbox = Listbox(self.fontFrame, selectmode=SINGLE, height=12, 
            yscrollcommand=self.fontScrollbar.set)
        self.fontScrollbar.config(command=self.fontListbox.yview)
        fonts = getFontNames()
        fonts.sort()
        for i in fonts:
            self.fontListbox.insert(END, i)
        self.font = 'Symbola Regular'
        self.fontButton = Button(self, text='Apply selected font', command=self.fontApply)

        # start year
        self.startyrLabel = Label(self, text='Start year:')
        self.startyrVar = StringVar()
        self.startyrEntry = Entry(self, textvariable=self.startyrVar, width=4)

        # start month
        self.startmthLabel = Label(self, text='Start month:')
        self.startmthVar = StringVar()
        self.startmthEntry = Entry(self, textvariable=self.startmthVar, width=2)

        #number of months per row
        self.nrHmthsLabel = Label(self, text='Number of months horizontal:')
        self.nrHmthsVar = StringVar()
        self.nrHmthsEntry = Entry(self, textvariable=self.nrHmthsVar, width=2)

        # start of week
        self.weekStartsLabel = Label(self, text='Week begins with:')
        self.weekVar = IntVar()
        self.weekMondayRadio = Radiobutton(self, text='Mon', variable=self.weekVar,
            value=calendar.MONDAY)
        self.weekSundayRadio = Radiobutton(self, text='Sun', variable=self.weekVar,
            value=calendar.SUNDAY)

       # include weeknumber
        self.weekNrLabel = Label(self, text='Show week numbers:')
        self.weekNrVar = IntVar()
        self.weekNrCheck = Checkbutton(self, variable=self.weekNrVar)
        self.weekNrHdLabel = Label(self, text='Week numbers heading:')
        self.weekNrHdVar = StringVar()
        self.weekNrHdEntry = Entry(self, textvariable=self.weekNrHdVar, width=6)
        
        # offsetX, offsetY and inner margins
        self.offsetXLabel = Label(self, text='Calendar offset \nfrom left margin (pt):')
        self.offsetXVar = DoubleVar()
        self.offsetXEntry = Entry(self, textvariable=self.offsetXVar, width=7)
        self.offsetYLabel = Label(self, text='Calendar offset \nfrom top margin (pt):')
        self.offsetYVar = DoubleVar()
        self.offsetYEntry = Entry(self, textvariable=self.offsetYVar, width=7)
        self.marginXLabel = Label(self, text='Inner vertical margin (pt):')
        self.marginXVar = DoubleVar()
        self.marginXEntry = Entry(self, textvariable=self.marginXVar, width=7)
        self.marginYLabel = Label(self, text='Inner horizontal margin (pt):')
        self.marginYVar = DoubleVar()
        self.marginYEntry = Entry(self, textvariable=self.marginYVar, width=7)

        # draw image frame
        self.imageLabel = Label(self, text='Draw Image Frame:')
        self.imageVar = IntVar()
        self.imageCheck = Checkbutton(self, variable=self.imageVar)

        # holidays
        self.holidaysLabel = Label(self, text='Show holidays:')
        self.holidaysVar = IntVar()
        self.holidaysCheck = Checkbutton(self, variable=self.holidaysVar)

        # legend
        self.legendLabel = Label(self, text='Show holiday texts:')
        self.legendVar = IntVar()
        self.legendCheck = Checkbutton(self, variable=self.legendVar)

        # closing/running
        self.okButton = Button(self, text="OK", width=6, command=self.okButton_pressed)
        self.cancelButton = Button(self, text="Cancel", command=self.quit)

        # setup values
        self.startyrVar.set(str(datetime.date(1, 1, 1).today().year+1)) # +1 for next year
        self.startmthVar.set("1")
        self.nrHmthsVar.set("3")
        self.weekMondayRadio.select()
        self.weekNrCheck.select()
        self.weekNrHdVar.set("wk")
        self.offsetXVar.set("0.0")
        self.offsetYVar.set("0.0")
        self.marginXVar.set("0.0")
        self.marginYVar.set("0.0")
        #self.imageCheck.select()
        self.holidaysCheck.select()
        self.legendCheck.select()

        # make layout
        self.columnconfigure(0, pad=6)
        currRow = 0
        self.statusLabel.grid(column=0, row=currRow, columnspan=4)
        currRow += 1
        self.langLabel.grid(column=0, row=currRow, sticky=W)
        self.fontLabel.grid(column=1, row=currRow, sticky=W) 
        currRow += 1
        self.langFrame.grid(column=0, row=currRow, rowspan=6, sticky=N)
        self.fontFrame.grid(column=1, row=currRow, sticky=N)
        self.fontScrollbar.grid(column=1, row=currRow, sticky=N+S+E)
        self.fontListbox.grid(column=0, row=currRow, sticky=N+S+W)
        currRow += 2
        self.langButton.grid(column=0, row=currRow)
        self.fontButton.grid(column=1, row=currRow)
        currRow += 1
        self.startyrLabel.grid(column=0, row=currRow, sticky=S+E)
        self.startyrEntry.grid(column=1, row=currRow, sticky=S+W)
        self.nrHmthsLabel.grid(column=2, row=currRow, sticky=S+E)
        self.nrHmthsEntry.grid(column=3, row=currRow, sticky=S+W)
        currRow += 1
        self.startmthLabel.grid(column=0, row=currRow, sticky=S+E)
        self.startmthEntry.grid(column=1, row=currRow, sticky=S+W)
        self.offsetYLabel.grid(column=2, row=currRow, sticky=S+E)
        self.offsetYEntry.grid(column=3, row=currRow, sticky=S+W)
        currRow += 1
        self.weekStartsLabel.grid(column=0, row=currRow, sticky=S+E)
        self.weekMondayRadio.grid(column=1, row=currRow, sticky=S+W)
        self.marginYLabel.grid(column=2, row=currRow, sticky=N+E)
        self.marginYEntry.grid(column=3, row=currRow, sticky=W)
        currRow += 1
        self.weekSundayRadio.grid(column=1, row=currRow, sticky=N+W)
        self.offsetXLabel.grid(column=2, row=currRow, sticky=S+E)
        self.offsetXEntry.grid(column=3, row=currRow, sticky=S+W)
        currRow += 1
        self.weekNrLabel.grid(column=0, row=currRow, sticky=N+E)
        self.weekNrCheck.grid(column=1, row=currRow, sticky=N+W)
        self.marginXLabel.grid(column=2, row=currRow, sticky=N+E)
        self.marginXEntry.grid(column=3, row=currRow, sticky=W)
        currRow += 1
        self.weekNrHdLabel.grid(column=0, row=currRow, sticky=N+E)
        self.weekNrHdEntry.grid(column=1, row=currRow, sticky=N+W)
        self.imageLabel.grid(column=2, row=currRow, sticky=N+E)
        self.imageCheck.grid(column=3, row=currRow, sticky=N+W)
        currRow += 1
        self.holidaysLabel.grid(column=0, row=currRow, sticky=N+E)
        self.holidaysCheck.grid(column=1, row=currRow, sticky=N+W)
        self.legendLabel.grid(column=2, row=currRow, sticky=N+E)
        self.legendCheck.grid(column=3, row=currRow, sticky=N+W)
        currRow += 1
        self.rowconfigure(currRow, pad=6)
        self.okButton.grid(column=1, row=currRow, sticky=E)
        self.cancelButton.grid(column=2, row=currRow, sticky=W)

        # fill the months values
        self.realLangChange()

    def languageChange(self):
        """ Called by Change button. Get language list value and
            call real re-filling. """
        ix = self.langListbox.curselection()
        if len(ix) == 0:
            self.statusVar.set('Select a language, please')
            return
        langX = self.langListbox.get(ix[0])
        self.lang = langX
        if os == "Windows":
            x = langX
        else: # Linux
            iy = [[x[0] for x in localization].index(self.lang)]
            x = (localization[iy[0]][2])
        try:
            locale.setlocale(locale.LC_CTYPE, x)
            locale.setlocale(locale.LC_TIME, x)
        except locale.Error:
            print("Language " + x + " is not installed on your operating system.")
            self.statusVar.set("Language '" + x + "' is not installed on your operating system")
            return
        self.realLangChange(langX)

    def realLangChange(self, langX='English'):
        """ Real widget setup. It takes values from localization list.
        [0] = months, [1] Days """
        self.lang = langX
        if os == "Windows":
            ix = [[x[0] for x in localization].index(self.lang)]
            self.calUniCode = (localization[ix[0]][1]) # get unicode page for the selected language
        else: # Linux
            self.calUniCode = "UTF-8"

    def fontApply(self, chosenFont = 'Symbola Regular'):
        """ Font selection. Called by "Apply selected font" button click. """
        ix = self.fontListbox.curselection()
        if len(ix) == 0:
            self.statusVar.set('Please select a font.')
            return
        self.font = self.fontListbox.get(ix[0])

    def okButton_pressed(self):
        """ User variables testing and preparing """
        # start year
        try:
            year = self.startyrVar.get().strip()
            if len(year) != 4:
                raise ValueError
            year = int(year, 10)
        except ValueError:
            self.statusVar.set('Year must be in the "YYYY" format e.g. 2020.')
            return
        # start month
        try:
            stmonth = self.startmthVar.get().strip()
            if (int(stmonth,10) < 1 or int(stmonth,10) > 12):
                raise ValueError
            stmonth = int(stmonth, 10)
            months = []
            for i in range (0, 12):
                 j = stmonth + i
                 if j > 12:     # Start month is not 1
                     j = j - 12
                 months.append(int(j))
        except ValueError:
            self.statusVar.set('Start month must be between 1 and 12.')
            return
        # number of months per row
        try:
            nrHmonths = self.nrHmthsVar.get().strip()
            if (int(nrHmonths,10) < 1 or int(nrHmonths,10) > 12):
                raise ValueError
            nrHmonths = int(nrHmonths, 10)
        except ValueError:
            self.statusVar.set('Number of months per row must be between 1 and 12.')
            return
        # inner margins
        if ((float(self.offsetXVar.get()) - float(self.marginXVar.get())) < 0 
            or (float(self.offsetYVar.get()) - float(self.marginYVar.get())) < 0):
            self.statusVar.set('Inner margins must be less than offsets.')
            return
        # fonts
        fonts = getFontNames()
        if self.font not in fonts:
            self.statusVar.set('Please select a font.')
            return
        # week numbers
        if self.weekNrVar.get() == 0:
            weekNr = False
        else:
            weekNr = True
        # draw images
        if self.imageVar.get() == 0:
            drawImg = False
        else:
            drawImg = True
        # holidays
        if self.holidaysVar.get() == 0: 
            holidaysList = list()
        else:
            hol = calcHolidays(year)                      
            holidaysList = hol.importHolidays()

            x = 0
            y = len(holidaysList)
            while x < y:
                if ((holidaysList[x][0] == year and int(holidaysList[x][1]) < stmonth) or
                    (holidaysList[x][0] == year+1 and int(holidaysList[x][1]) >= stmonth)):
                    del holidaysList[x] # delete holidays not needed and reset counters
                    x = x - 1
                    y = y - 1
                x = x + 1
            holidaysList.sort(key = lambda i: int(i[2])) # sort on day
            holidaysList.sort(key = lambda i: int(i[1])) # sort on month
            holidaysList.sort(key = lambda i: i[0]) # sort on year
        # draw legend (holiday texts)
        if self.legendVar.get() == 0:
            drawLegend = False
        else:
            drawLegend = True
        # create calendar (finally)
        cal = ScYearCalendar(year, months, nrHmonths, self.weekVar.get(), weekNr, 
                self.weekNrHdVar.get(), float(self.offsetXVar.get()), 
                float(self.marginXVar.get()), float(self.offsetYVar.get()),
                float(self.marginYVar.get()), drawImg, drawLegend,
                self.font, self.lang, holidaysList)
        self.master.withdraw()
        err = cal.createCalendar()
        if err != None:
            self.master.deiconify()
            self.statusVar.set(err)
        else:
            self.quit()

    def quit(self):
        self.master.destroy()

######################################################
def main():
    """ Application/Dialog loop with Scribus sauce around """
    try:
        statusMessage('Running script...')
        progressReset()
        original_locale1=locale.getlocale(locale.LC_CTYPE)
        original_locale2=locale.getlocale(locale.LC_TIME)
        root = Tk()
        app = TkCalendar(root)
        root.mainloop()
        locale.setlocale(locale.LC_CTYPE, original_locale1)
        locale.setlocale(locale.LC_TIME, original_locale2)
    finally:
        if haveDoc() > 0:
            redrawAll()
        statusMessage('Done.')
        progressReset()

if __name__ == '__main__':
    main()

