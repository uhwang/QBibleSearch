# -*- coding: utf-8 -*-
"""
	Author
	-------------------------------
	Uisang Hwang
	
	History
	-------------------------------
	Oct  7  2013   	Simple search(console app)
	Oct  8  2013   	Plot Hit (matplotlib)
	Oct  9  2013   	Plot Sorted Hit
	Oct 10  2013   	Export hit list Excel 2007 w/ chart (xlswriter)
	Nov  4  2013   	Make a corection of unorddered books w/ hit under unsort option
	Feb  7  2014   	Add option for setting a plot title in English or Korean 
	Nov  3  2014   	Add GUI (PyQt4)
	Nov  4  2014   	Save verse list
	Nov  5  2014   	Windows executible tested on cx_freeze 
					Error --icon=bth-0.png
					cxfreeze QBibSearch.py --base-name=Win32GUI --target-dir QBS --exclude-modules=scipy,numpy,matplotlib
	Nov  5  2014   	Fix: Bible Datafile
						NASB(engNasb3.bdf, Line# 1539)
							13 대상 1:1 
								"?Adam, Seth, Enos" -> "Adam, Seth, Enos"
	Nov  8  2014   	Convert bth-0.png to bth-0.xpm. Insert xpm string to QPixmap.
	Nov 11  2014   	HTML5 + Google chart
	Nov 12  2014   	Google chart: save chart as an image file (PNG)
					Add verse list to html 
	Nov 16  2014   	Add higlighting a keyword in html file
	Nov 18  2014   	Regular expression for searching a keyword
	Nov 25  2014   	Save search result as a HTML file with two columns of different Bible version
	Jan  2  2015   	Replace Google chart as AmCharts (http://www.amcharts.com/)
	Jan 18  2015   	Add menu bar
	Jan 19  2015   	Runtime error was found at different computers due to msvcr100.dll. 
					Microsoft Visual C++ 2010 Redistributable Package (x86)
					http://www.microsoft.com/en-gb/download/details.aspx?id=5555
	Jan 22  2015   	Fix: search option bug (case menu)
					Add lookbehind in the regular expression pattern to avoid '-' preceding a keyword
	Feb 3	2015	Fix: Bible Datafile
						NASB(engNASB4.bdf, Line# 2050)
						19시 68:17 
						"Yousands upon Yousands;" -> "thousands upon thousands;"
	Feb 12	2015	Add chart size menu
	Mar 01  2015    Add multiple pixmap for application icons
					Free online icon editor:
					http://www.xiconeditor.com/
					
					How to show icon on Windows task bar:
					http://stackoverflow.com/questions/12432637/pyqt4-set-windows-taskbar-icon
					http://stackoverflow.com/questions/1551605/how-to-set-applications-taskbar-icon-in-windows-7/1552105#1552105
	Mar 08  2015    Add Hebrew Bible database (THE WESTMINSTER LENINGRAD CODEX)
					http://tanach.us/Tanach.xml
	Mar 09  2015    Add SBL Greek New Testament 
					http://sblgnt.com/download/
	Mar 14  2015	Fix mismatching of Hebrew Bible and other Bible versions
					http://www.bibleworks.com/bw9help/bwh29_Setup.htm
						Editing Verse Maps
	Apr 09  2015    Chart only menu added				
	Jul 08  2015	Get BW stat data from clipboard and export plot
	Dec 22  2017    Fix: Bible Datafile
						NASB(engNASB3.bdf, Line# 2466) 
						1 Chronicles 29:16 "... all is Your" --> "... all is Yours"
	Apr 17  2018    Fix: Bible Datafile
						NIV(engNIV4.bdf, Line# 5726)
						Isaiah 46:3 "... conceived, an have ..." --> "... conceived, and have ..."
	Jun 17  2019    Fix: Bible Datafile
						개역개정-국한문(kchNKRV5.bdf, Line# 423)
						Jeremiah 17:11 "... TC새가 ..." --> "... 자고새가 ..."
	Feb 15  2020    Fix: Bible Datafile
						NASB (engNASB6.bdf, Line# 1268)
						Mark 6:3, "... James and Joses ..." --> "James and Joseph"
						
	Required Python Packages
	-------------------------------
	numpy (not required)
	xlsxwriter
	matplotlib (not required)
	
	About
	-------------------------------
	This program search a keyword from Bethlehem Bible datafiles and
	creates a vls file (BibleWorks Verse List) and a Microsoft Excel file. 
	
	Note
	-------------------------------
	Add choosing the Directory of Bathelehem program
	Add Save the path as a file
	Add Kor/Eng 
	Search Options
		Korean	Hangul Revised Version 	(KorHrv1-7.bdf)
		English	King James Versin		(EngKjv1-7.bdf)	
		
	Check if a string has Hangul (ask.python.kr/question/55116/문자열이 한글인지 체크하려면/)
   
    Nov 22 2017, HTML Table column width correction for Hebrew Text
	<caption> .... </caption>
		<colgroup>
			<col width="5%" />
			<col width="5%" />
			<col width="27%" />
			<col width="4%" />
			<col width="27%" />
			<col width="4%" />
			<col width="45%" />
		</colgroup>	
		
	Usage
	-------------------------------
	1. Run BibleWorks
	2. Goto main menus
		/ Tools
			/ Analyzing the text
				/ Verse List Manager (VLM)
	3. Goto VLM menu
		/ Import
			/ From File
	4. Choose res.vls
	5. Goto VLM menu
		/ Export
			/ Export to search window
"""
import re
import os
import sys
import string
import ctypes
import sqlite3 as db
from PyQt4 import QtCore, QtGui
import webbrowser
import wttmap as wmap
import qbsicon
import win32clipboard as cb
import win32con

#import qbs_icon_64
#import qbs_icon_48
#import qbs_icon_32
#import qbs_icon_16

def is_hangul(s):
	#from unicodedata import name
	#for c in unicode(s):
	for c in s:
		if '\uac00' <= c <= '\ud7a3' or 'ㄱ' <= c <= 'ㅎ':
			return True
	return False

# KEY  KBIB BW-INTERNAL KBIB             EBIB               BW-EXPORT 
#      ABBR    NAME     NAME             NAME               NAME 
book_table = {
'01': ['창'  , 'Gen' , '창세기'        , 'Genesis'        , 'Gen.'    ],
'02': ['출'  , 'Exo' , '출애굽기'      , 'Exodus'         , 'Exod.'   ],
'03': ['레'  , 'Lev' , '레위기'        , 'Leviticus'      , 'Lev.'    ],
'04': ['민'  , 'Num' , '민수기'        , 'Numbers'        , 'Num.'    ],
'05': ['신'  , 'Deu' , '신명기'        , 'Deuteronomy'    , 'Deut.'   ],
'06': ['수'  , 'Jos' , '여호수아'      , 'Joshua'         , 'Jos.'    ],
'07': ['삿'  , 'Jdg' , '사사기'        , 'Judges'         , 'Jdg.'    ],
'08': ['룻'  , 'Rut' , '룻기'          , 'Ruth'           , 'Ruth'    ],
'09': ['삼상', '1Sa' , '사무엘상'      , '1 Samuel'       , '1 Sam.'  ],
'10': ['삼하', '2Sa' , '사무엘하'      , '2 Samuel'       , '2 Sam.'  ],
'11': ['왕상', '1Ki' , '열왕기상'      , '1 Kings'        , '1 Ki.'   ],
'12': ['왕하', '2Ki' , '열왕기하'      , '2 Kings'        , '2 Ki.'   ],
'13': ['대상', '1Ch' , '역대상'        , '1 Chronicles'   , '1 Chr.'  ],
'14': ['대하', '2Ch' , '역대하'        , '2 Chronicles'   , '2 Chr.'  ],
'15': ['스'  , 'Ezr' , '에스라'        , 'Ezra'           , 'Ezr.'    ],
'16': ['느'  , 'Neh' , '느헤미야'      , 'Nehemiah'       , 'Neh.'    ],
'17': ['에'  , 'Est' , '에스더'        , 'Esther'         , 'Est.'    ],
'18': ['욥'  , 'Job' , '욥기'          , 'Job'            , 'Job'     ],
'19': ['시'  , 'Psa' , '시편'          , 'Psalms'         , 'Ps.'     ],
'20': ['잠'  , 'Pro' , '잠언'          , 'Proverbs'       , 'Prov.'   ],
'21': ['전'  , 'Ecc' , '전도서'        , 'Ecclesiastes'   , 'Eccl.'   ],
'22': ['아'  , 'Sol' , '아가'          , 'Song of Solomon', 'Cant.'   ],
'23': ['사'  , 'Isa' , '이사야'        , 'Isaiah'         , 'Isa.'    ],
'24': ['렘'  , 'Jer' , '예레미야'      , 'Jeremiah'       , 'Jer.'    ],
'25': ['애'  , 'Lam' , '예레미야애가'  , 'Lamentations'   , 'Lam.'    ],
'26': ['겔'  , 'Eze' , '에스겔'        , 'Ezekiel'        , 'Ezek.'   ],
'27': ['단'  , 'Dan' , '다니엘'        , 'Daniel'         , 'Dan.'    ],
'28': ['호'  , 'Hos' , '호세아'        , 'Hosea'          , 'Hos.'    ],
'29': ['욜'  , 'Joe' , '요엘'          , 'Joel'           , 'Joel'    ],
'30': ['암'  , 'Amo' , '아모스'        , 'Amos'           , 'Amos'    ],
'31': ['옵'  , 'Oba' , '오바댜'        , 'Obadiah'        , 'Obad.'   ],
'32': ['욘'  , 'Jon' , '요나'          , 'Jonah'          , 'Jon.'    ],
'33': ['미'  , 'Mic' , '미가'          , 'Micah'          , 'Mic.'    ],
'34': ['나'  , 'Nah' , '나훔'          , 'Nahum'          , 'Nah.'    ],
'35': ['합'  , 'Hab' , '하박국'        , 'Habakkuk'       , 'Hab.'    ],
'36': ['습'  , 'Zep' , '스바냐'        , 'Zephaniah'      , 'Zeph.'   ],
'37': ['학'  , 'Hag' , '학개'          , 'Haggai'         , 'Hag.'    ],
'38': ['슥'  , 'Zec' , '스가랴'        , 'Zechariah'      , 'Zech.'   ],
'39': ['말'  , 'Mal' , '말라기'        , 'Malachi'        , 'Mal.'    ],
'40': ['마'  , 'Mat' , '마태복음'      , 'Matthew'        , 'Matt.'   ],
'41': ['막'  , 'Mar' , '마가복음'      , 'Mark'           , 'Mk.'     ],
'42': ['누'  , 'Luk' , '누가복음'      , 'Luke'           , 'Lk.'     ],
'43': ['요'  , 'Joh' , '요한복음'      , 'John'           , 'Jn.'     ],
'44': ['행'  , 'Act' , '사도행전'      , 'Acts'           , 'Acts'    ],
'45': ['롬'  , 'Rom' , '로마서'        , 'Romans'         , 'Rom.'    ],
'46': ['고전', '1Co' , '고린도전서'    , '1Corinthians'   , '1 Co.'   ],
'47': ['고후', '2Co' , '고린도후서'    , '2Corinthians'   , '2 Co.'   ],
'48': ['갈'  , 'Gal' , '갈라디아서'    , 'Galatians'      , 'Gal.'    ],
'49': ['엡'  , 'Eph' , '에베소서'      , 'Ephesians'      , 'Eph.'    ],
'50': ['빌'  , 'Phi' , '빌립보서'      , 'Philippians'    , 'Phil.'   ],
'51': ['골'  , 'Col' , '골로새서'      , 'Colossians'     , 'Col.'    ],
'52': ['살전', '1Th' , '데살로니가전서', '1Thessalonians' , '1 Thess.'],
'53': ['살후', '2Th' , '데살로니가후서', '2Thessalonians' , '2 Thess.'],
'54': ['딤전', '1Ti' , '디모데전서'    , '1Timothy'       , '1 Tim.'  ],
'55': ['딤후', '2Ti' , '디모데후서'    , '2Timothy'       , '2 Tim.'  ],
'56': ['딛'  , 'Tit' , '디도서'        , 'Titus'          , 'Tit.'    ],
'57': ['몬'  , 'Phl' , '빌레몬서'      , 'Philemon'       , 'Phlm.'   ],
'58': ['히'  , 'Heb' , '히브리서'      , 'Hebrews'        , 'Heb.'    ],
'59': ['약'  , 'Jam' , '야고보서'      , 'James'          , 'Jas.'    ],
'60': ['벧전', '1Pe', '베드로전서'     , '1Peter'         , '1 Pet.'  ],
'61': ['벧후', '2Pe', '베드로후서'     , '2Peter'         , '2 Pet.'  ],
'62': ['요일', '1Jo', '요한1서'        , '1John'          , '1 Jn.'   ],
'63': ['요이', '2Jo', '요한2서'        , '2John'          , '2 Jn.'   ],
'64': ['요삼', '3Jo', '요한3서'        , '3John'          , '3 Jn.'   ],
'65': ['유'  , 'Jud', '유다서'         , 'Jude'           , 'Jude'    ],
'66': ['계'  , 'Rev', '요한계시록'     , 'Revelation'     , 'Rev.'    ]
}

hebrew_bible_db_table_name = 'HebBible'
greek_bible_db_table_name = 'GrkBible'
greek_bible_db = 'sblgnt.db'
hebrew_bible_db = 'wlc.db'
bdf_ext = ".bdf"
qbs_datafile = "qbs.dat"
sql_ext = ".db"
max_bdf_number = 7
number_of_ot_books = 39
number_of_nt_books = 27
number_of_bible_books = 66
	
number_of_korean_bible = 14
number_of_english_bible = 12
qbs_barchart_color_table = []

default_chart_width = 800
default_chart_height = 400
default_label_size = 11

korean_bible_name = [
"개역한글판", "개역개정판"    , "킹제임스흠정역", "한글킹제임스", 
"새번역"    , "공동번역개정판", "바른성경"      , "가톨릭성경"  , 
"우리말성경", "쉬운성경"      , "현대인의 성경" , "현대어성경"  ,
"개역한글-국한문", 
"개역개정-국한문"
#"바른성경-국한문"
]

korean_bible_prefix = [
"korHRV" , "korNKRV" , "korHKJV", "korKKJV",
"korNRSV", "korNKCB" , "korKTV" , "korCath",
"korDOB" , "korEASY" , "korKLB" , "korTKV" ,
"kchHRV",
"kchNKRV"
#"kchKTV" 
]

english_bible_name = [
"NASB", "ESV" , "KJV" , "GNT",
"HCSB", "ISV" , "MSG" , "NIV",
"NKJV", "NLT" , "NRSV", "TNIV" ]

english_bible_prefix = [
"engNASB", "engESV", "ENGKJV" , "engGNT",
"engHCSB", "engISV", "engMSG" , "ENGNIV",
"Engnkjv", "engnlt", "Engnrsv", "engTNIV"  ]

_find_bdfinfo = re.compile('([\d]{2}).*(?<!\d)([\d]*):([\d]*)')
_find_bwverse = re.compile('(.*) (\d*):(.*)')

class RGB:
	def __init__(self):
		self.r = 0
		self.g = 0
		self.b = 0
	
	def __init__(self, r_, g_, b_):
		self.r = r_
		self.g = g_
		self.b = b_

class SearchOption:
	def __init__(self):
		self.case_sensitive = False
		self.whole_word = False
		self.pattern_option = re.UNICODE
		
class Hit_Index:
	def __init__(self):
		self.file_index = []
		self.file_pointer = []
		self.number = []
		
class Book_Hit:
	def __init__(self, name, hit, id):
		self.book_name = name
		self.book_hit = hit
		self.id = id
		self.verse = Hit_Index()

class MultipleSearchOption:
	def __init__(self):
		self.gbib = 0 # Greek Bible
		self.hbib = 0 # Hebrew Bible
		self.kbib = 0 # Korean Bible
		self.ebib = 0 # English Bible
		self.select = 0

class ExportOption:
	def __init__(self):
		self.chart_only = 0
		
class ChartOption:
	def __init__(self):
		self.width = default_chart_width
		self.height = default_chart_height
		self.xlabel_size = default_label_size
		self.ylabel_size = default_label_size

'''
class QChartSizeEdit(QtGui.QLineEdit, QtCore.QObject):
	def __init__(self):
		super(QChartSizeEdit, self).__init__()
		self.textChanged.connect(self.on_text_changed)
		
	def on_text_changed(self, string):
		self.size = int(string)
'''
class QLabelSizeEditDlg(QtGui.QDialog):
	def __init__(self, chartOption):
		super(QLabelSizeEditDlg, self).__init__()
		self.initUI(chartOption)
		
	def initUI(self, chartOption):
		form_layout = QtGui.QGridLayout()
		edit_xlabel = QtGui.QLabel('XLabel Size')
		self.xlabel_edit  = QtGui.QLineEdit(self)
		self.xlabel_edit.setText('{}'.format(chartOption.xlabel_size))

		edit_ylabel = QtGui.QLabel('YLabel Size')
		self.ylabel_edit  = QtGui.QLineEdit(self)
		self.ylabel_edit.setText('{}'.format(chartOption.ylabel_size))
		
		self.ok = QtGui.QPushButton('OK', self)
		self.ok.clicked.connect(self.closeOption)
		
		form_layout.addWidget(edit_xlabel, 0, 0)
		form_layout.addWidget(self.xlabel_edit, 0, 1)
		form_layout.addWidget(edit_ylabel, 1, 0)
		form_layout.addWidget(self.ylabel_edit, 1, 1)
		form_layout.addWidget(self.ok, 2, 0)
		self.setLayout(form_layout)
		self.setWindowTitle('Label Size')

	def closeOption(self):
		self.done(1)
		
	def getXLabelSize(self):
		return int(self.xlabel_edit.text())		
	def getYLabelSize(self):
		return int(self.ylabel_edit.text())		
		
class QChartSizeEditDlg(QtGui.QDialog):
	def __init__(self, chartSize):
		super(QChartSizeEditDlg, self).__init__()
		self.initUI(chartSize)
		
	def initUI(self, chartSize):
		form_layout = QtGui.QGridLayout()
		width_label = QtGui.QLabel('Width')
		self.width_edit  = QtGui.QLineEdit(self)
		height_label = QtGui.QLabel('Height')
		self.height_edit  = QtGui.QLineEdit(self)
		
		#self.width_edit.textChanged.connect(self.OnWidthChanged)
		#self.theight_edit.extChanged.connect(self.OnHeightChanged)
		
		form_layout.addWidget(width_label, 0, 0)
		form_layout.addWidget(self.width_edit, 0, 1)
		
		form_layout.addWidget(height_label, 1, 0)
		form_layout.addWidget(self.height_edit, 1, 1)
		
		self.setLayout(form_layout)
		
		self.width_edit.setText('{}'.format(chartSize.width))
		self.height_edit.setText('{}'.format(chartSize.height))
		
		self.ok = QtGui.QPushButton('OK', self)
		self.ok.clicked.connect(self.closeOption)
		form_layout.addWidget(self.ok, 2, 0)

		self.setWindowTitle('Chart Size')
		
	def closeOption(self):
		self.done(1)
		
	def getChartSize(self):
		return int(self.width_edit.text()), int(self.height_edit.text())		
		
class CrossRefSearchOption:
	def __init__(self):
		korHRV  = False, 
		korNKRV = False, 
		korHKJV = False, 
		korKKJV = False,
		korNRSV = False, 
		korNKCB = False, 
		korKTV  = False, 
		korCATH = False,
		korDOB  = False, 
		korEASY = False, 
		korKLB  = False, 
		korTKV  = False,
		kchHRV  = False,
		kchNKRV = False,
		
		engNASB = False, 
		engESV  = False, 
		engKJV  = False, 
		engGNT  = False,
		engHCSB = False, 
		engISV  = False, 
		engMSG  = False, 
		engNIV  = False,
		engNKJV = False, 
		engNLT  = False, 
		engNRSV = False, 
		engTNIV = False
	
class CrossRefSearchDialog(QtGui.QDialog):
	def __init__(self, option):
		super(CrossRefSearchDialog, self).__init__()
		self.initUI(option)

	def initUI(self, option):
	
		layout = QtGui.QFormLayout()
		kbib_group = QtGui.QGroupBox('Korean Bible')
		ebib_group = QtGui.QGroupBox('English Bible')
		
		kbib_layout = QtGui.QGridLayout()
		ebib_layout = QtGui.QGridLayout()
		
		self.korHRV  = QtGui.QCheckBox("개역한글판", self)
		self.korNKRV = QtGui.QCheckBox("개역개정판", self) 
		self.korHKJV = QtGui.QCheckBox("킹제임스흠정역", self)
		self.korKKJV = QtGui.QCheckBox("한글킹제임스", self)
		self.korNRSV = QtGui.QCheckBox("새번역", self)
		self.korNKCB = QtGui.QCheckBox("공동번역개정판", self)
		self.korKTV  = QtGui.QCheckBox("바른성경", self)
		self.korCATH = QtGui.QCheckBox("가톨릭성경", self)
		self.korDOB  = QtGui.QCheckBox("우리말성경", self)
		self.korEASY = QtGui.QCheckBox("쉬운성경", self)
		self.korKLB  = QtGui.QCheckBox("현대인의 성경", self)
		self.korTKV  = QtGui.QCheckBox("현대어성경", self)
		self.kchHRV  = QtGui.QCheckBox("개역한글-국한문", self)
		self.kchNKRV = QtGui.QCheckBox("개역개정-국한문", self)
		
		kbib_layout.addWidget(self.korHRV  , 0, 0)
		kbib_layout.addWidget(self.korNKRV , 0, 1)
		kbib_layout.addWidget(self.korHKJV , 0, 2)
		kbib_layout.addWidget(self.korKKJV , 0, 3)
		kbib_layout.addWidget(self.korNRSV , 1, 0)
		kbib_layout.addWidget(self.korNKCB , 1, 1)
		kbib_layout.addWidget(self.korKTV  , 1, 2)
		kbib_layout.addWidget(self.korCATH , 1, 3)
		kbib_layout.addWidget(self.korDOB  , 2, 0)
		kbib_layout.addWidget(self.korEASY , 2, 1)
		kbib_layout.addWidget(self.korKLB  , 2, 2)
		kbib_layout.addWidget(self.korTKV  , 2, 3)
		kbib_layout.addWidget(self.kchHRV  , 3, 0)
		kbib_layout.addWidget(self.kchNKRV , 3, 1)
		
		self.engNASB  = QtGui.QCheckBox("NASB", self)
		self.engESV   = QtGui.QCheckBox("ESV" , self)
		self.engKJV   = QtGui.QCheckBox("KJV" , self)
		self.engGNT   = QtGui.QCheckBox("GNT" , self)
		self.engHCSB  = QtGui.QCheckBox("HCSB", self)
		self.engISV   = QtGui.QCheckBox("ISV" , self)
		self.engMSG   = QtGui.QCheckBox("MSG" , self)
		self.engNIV   = QtGui.QCheckBox("NIV" , self)
		self.engNKJV  = QtGui.QCheckBox("NKJV", self)
		self.engNLT   = QtGui.QCheckBox("NLT" , self)
		self.engNRSV  = QtGui.QCheckBox("NRSV", self)
		self.engTNIV  = QtGui.QCheckBox("TNIV", self)
		
		ebib_layout.addWidget(self.engNASB , 0, 0)
		ebib_layout.addWidget(self.engESV  , 0, 1)
		ebib_layout.addWidget(self.engKJV  , 0, 2)
		ebib_layout.addWidget(self.engGNT  , 0, 3)
		ebib_layout.addWidget(self.engHCSB , 1, 0)
		ebib_layout.addWidget(self.engISV  , 1, 1)
		ebib_layout.addWidget(self.engMSG  , 1, 2)
		ebib_layout.addWidget(self.engNIV  , 1, 3)
		ebib_layout.addWidget(self.engNKJV , 2, 0)
		ebib_layout.addWidget(self.engNLT  , 2, 1)
		ebib_layout.addWidget(self.engNRSV , 2, 2)
		ebib_layout.addWidget(self.engTNIV , 2, 3)
		
		kbib_group.setLayout(kbib_layout)
		ebib_group.setLayout(ebib_layout)
		
		layout.addRow(kbib_group)
		layout.addRow(ebib_group)
		
		self.setLayout(layout)
		self.setWindowTitle('Cross Ref Search')
		
class MultipleSearchDialog(QtGui.QDialog):
	def __init__(self, option):
		super(MultipleSearchDialog, self).__init__()
		self.initUI(option)
		
	def initUI(self, option):
		form_layout = QtGui.QGridLayout()
		
		self.sblgnt_bible = QtGui.QCheckBox('SBL Greek NT Bible', self)
		self.hebrew_bible = QtGui.QCheckBox('Hebrew Bible(BHS 1983)', self)
		self.korean_bible = QtGui.QComboBox(self)
		self.english_bible = QtGui.QComboBox(self)

		self.kor_check= QtGui.QRadioButton('Kor', self)
		self.eng_check= QtGui.QRadioButton('Eng', self)
		
		form_layout.addWidget(self.hebrew_bible, 0, 0)
		form_layout.addWidget(self.sblgnt_bible, 1, 0)
		form_layout.addWidget(self.kor_check, 2, 0)
		form_layout.addWidget(self.eng_check, 2, 1)

		self.korean_bible.addItem('None')
		for x in range(number_of_korean_bible):
			self.korean_bible.addItem(korean_bible_name[x])

		self.english_bible.addItem('None')
		for x in range(number_of_english_bible):
			self.english_bible.addItem(english_bible_name[x])

		form_layout.addWidget(self.korean_bible, 3, 0)
		form_layout.addWidget(self.english_bible, 3, 1)
			
		self.kor_check.toggled.connect(self.kor_clicked)
		self.eng_check.toggled.connect(self.eng_clicked)
		#self.kor_check.setChecked(True)
		if option.kbib > 0: 
			self.kor_check.setChecked(True)
			self.korean_bible.setCurrentIndex(option.select)
		else: 
			self.eng_check.setChecked(True)
			self.english_bible.setCurrentIndex(option.select)
			
		self.sblgnt_bible.setChecked(option.gbib)
		self.hebrew_bible.setChecked(option.hbib)
		
		self.ok = QtGui.QPushButton('OK', self)
		self.ok.clicked.connect(self.closeOption)
		form_layout.addWidget(self.ok, 4, 0)
		self.setLayout(form_layout)
		
		self.setWindowTitle('Multicolumn')
		
	def getSelectedItem(self):
		if self.kor_check.isChecked():
			return 1, 0, self.korean_bible.currentIndex(), self.sblgnt_bible.isChecked(), self.hebrew_bible.isChecked()
		elif self.eng_check.isChecked():
			return 0, 1, self.english_bible.currentIndex(), self.sblgnt_bible.isChecked(), self.hebrew_bible.isChecked()

	def closeOption(self):
		self.done(1)
		
	def kor_clicked(self):
		self.korean_bible.setEnabled(True)
		self.english_bible.setEnabled(False)
		
	def eng_clicked(self):
		self.korean_bible.setEnabled(False)
		self.english_bible.setEnabled(True)

#class QbsFindVerse(QtGui.QDialog):

		
def TG_HSV_To_RGB(H, S, V):
	import math
	I = 0.0
	F = 0.0
	P = 0.0
	Q = 0.0
	T = 0.0
	R1 = 0.0
	G1 = 0.0
	B1 = 0.0
	
	if S == 0:
		if H <= 0 or H > 360 :
			return (V,V,V)

	if H==360: H=0

	H = H/60;
	I = math.floor(H)
	F = H-I
	P = V*(1-S)
	Q = V*(1-S*F)
	T = V*(1-S*(1-F))
	
	int_I = int(I)
	
	if   int_I == 0: 
		R1 = V
		G1 = T
		B1 = P
	elif int_I == 1: 
		R1 = Q
		G1 = V
		B1 = P
	elif int_I == 2: 
		R1 = P
		G1 = V
		B1 = T
	elif int_I == 3: 
		R1 = P
		G1 = Q
		B1 = V
	elif int_I == 4: 
		R1 = T
		G1 = P
		B1 = V
	else           : 
		R1 = V
		G1 = P
		B1 = Q
	
	R = int(R1 * 255)
	G = int(G1 * 255)
	B = int(B1 * 255)

	return (R,G,B)

def TG_CreateHSVColorTable(H1, H2, S, V, order):
	
	if H1 > H2: 
		tempH = H2
		H2 = H1 
		H1 = tempH

	dH = (H2 - H1)/order
	tempH = H1

	for i in range(order):
		r,g,b = TG_HSV_To_RGB(tempH, S, V)
		qbs_barchart_color_table.append(RGB(r,g,b))
		tempH = tempH+dH

class QBibSearch(QtGui.QWidget):
	def __init__(self):
		super(QBibSearch, self).__init__()
		self.multisearch_option = MultipleSearchOption()
		self.chartOption = ChartOption()
		self.search_option = SearchOption()
		self.export_option = ExportOption()
		self.crossrefsearch_option = CrossRefSearchOption()
		self.initUI()
		self.readPathFile()
		self.sorted = False
		self.call_browser = True
		#self.createDBList()
		#self.checkDB()
		
	def setSearchOptionCase(self):
		self.search_option.case_sensitive = not self.search_option.case_sensitive
		self.setRegularExpressionCaseOption()
	
	def setSearchOptionWholeword(self):
		self.search_option.whole_word = not self.search_option.whole_word
		self.wholeword.setChecked(self.search_option.whole_word)
		
	def setChartSize(self):
		chart_size_dlg = QChartSizeEditDlg(self.chartOption)
		chart_size_dlg.exec_()
		self.chartOption.width, self.chartOption.height = chart_size_dlg.getChartSize()
	
	def setLabelSize(self):
		label_size_dlg = QLabelSizeEditDlg(self.chartOption)
		label_size_dlg.exec_()
		self.chartOption.xlabel_size = label_size_dlg.getXLabelSize()
		self.chartOption.ylabel_size = label_size_dlg.getYLabelSize()
		
	def setChartOnly(self):
		self.export_option.chart_only = not self.export_option.chart_only
		
	def setCallBrowser(self):
		self.call_browser = not self.call_browser
		
	def findVerse(self):
		do_nothing = ""
		
	def create_actions(self):
		# how to make qmenu item checkable pyqt4 python
		# http://stackoverflow.com/questions/10368947/how-to-make-qmenu-item-checkable-pyqt4-python
		self.searchOptionCaseMenu = QtGui.QAction('Case', self, checkable=True, triggered=self.setSearchOptionCase)
		self.chartSizeMenu = QtGui.QAction('Chart Size', self, triggered=self.setChartSize)
		self.chartOnlyMenu = QtGui.QAction('Chart Only', self, checkable=True, triggered=self.setChartOnly)
		self.labelSizeMenu = QtGui.QAction('Label Size', self, triggered=self.setLabelSize)
		self.utilMenu      = QtGui.QAction('Find verse', self, triggered=self.findVerse)
		self.clipboardBWMenu = QtGui.QAction('Clipboard BW', self, triggered=self.processClipboardBWPlotData)
		self.createBibleDB   = QtGui.QAction('Create Bible DB', self, triggered=self.processCreateBibleDB)
		self.bibleworks_export = QtGui.QAction('BW Export', self, triggered=self.process_bibleworks_exported_verlist)
		self.callBrowserMenu = QtGui.QAction('Call Browser', self, checkable=True, triggered=self.setCallBrowser)
					
		self.searchOptionCaseMenu.setChecked(True)
		self.chartOnlyMenu.setChecked(False)
		self.callBrowserMenu.setChecked(True)
		
	def readPathFile(self):
		self.message.appendPlainText('... Read {0}'.format(qbs_datafile))
		
		for x in range(number_of_korean_bible):
			self.korean_bible.addItem(korean_bible_name[x])
		
		for x in range(number_of_english_bible):
			self.english_bible.addItem(english_bible_name[x])
	
		try:
			reader = open(qbs_datafile)
			temp = reader.readline()
			self.destDir = temp.replace('\n', "")
			self.directory_path.setText(self.destDir)
			temp = reader.readline()
			self.search_keyword.setText(temp.replace('\n', ""))
		except IOError as err:
			errno, strerror = err.args
			self.message.appendPlainText("... I/O error({0}): {1}".format(errno, strerror))
			self.chooseDirectory()
			self.message.appendPlainText('... Write {0} to {1}'.format(self.destDir,qbs_datafile))
			writer = open("qbs.dat",'w')
			writer.write(self.destDir)
			writer.close()
	
	def pop_export_menu(self):
		aMenu = QtGui.QMenu(self)
		aMenu.addAction(self.chartSizeMenu)
		aMenu.addAction(self.chartOnlyMenu)
		aMenu.addAction(self.labelSizeMenu)
		return aMenu
	
	def pop_searchoption_menu(self):
		aMenu = QtGui.QMenu(self)
		aMenu.addAction(self.searchOptionCaseMenu)
		#aMenu.addAction(self.searchOptionWholewordMenu)
		return aMenu
		
	def pop_util_menu(self):
		aMenu = QtGui.QMenu(self)
		aMenu.addAction(self.utilMenu)
		aMenu.addAction(self.clipboardBWMenu)
		aMenu.addAction(self.bibleworks_export)
		aMenu.addAction(self.createBibleDB)
		return aMenu

	def pop_option_menu(self):
		aMenu = QtGui.QMenu(self)
		aMenu.addAction(self.callBrowserMenu)
		#aMenu.addAction(self.chartOnlyMenu)
		return aMenu
		
	def initUI(self):
	
		# http://qt-project.org/forums/viewthread/4199
		# Disable Windows Close Icon
		self.setWindowFlags((self.windowFlags() | QtCore.Qt.CustomizeWindowHint) & ~QtCore.Qt.WindowCloseButtonHint)
		self.create_actions()
		
		form_layout = QtGui.QFormLayout()
		# Jan 17 2015
		# PyQt MenuBar Outside MainWindow
		# https://acaciaecho.wordpress.com/2011/06/29/pyqt-menubar-outside-mainwindow/
		mainLayout = QtGui.QVBoxLayout()
		toolBar = QtGui.QToolBar()
		
		searchMenuButton = QtGui.QToolButton()
		searchMenuButton.setText('Search')
		searchMenuButton.setPopupMode(QtGui.QToolButton.MenuButtonPopup)
		searchMenuButton.setMenu(self.pop_searchoption_menu())
		
		chartMenuButton = QtGui.QToolButton()
		chartMenuButton.setText('Export')
		chartMenuButton.setPopupMode(QtGui.QToolButton.MenuButtonPopup)
		chartMenuButton.setMenu(self.pop_export_menu())
		
		utilMenuButton = QtGui.QToolButton()
		utilMenuButton.setText('Util')
		utilMenuButton.setPopupMode(QtGui.QToolButton.MenuButtonPopup)
		utilMenuButton.setMenu(self.pop_util_menu())
		
		optionMenuButton = QtGui.QToolButton()
		optionMenuButton.setText('Option')
		optionMenuButton.setPopupMode(QtGui.QToolButton.MenuButtonPopup)
		optionMenuButton.setMenu(self.pop_option_menu())
		
		toolBar.addWidget(searchMenuButton)
		toolBar.addWidget(chartMenuButton)
		toolBar.addWidget(utilMenuButton)
		toolBar.addWidget(optionMenuButton)

		mainLayout.addWidget(toolBar)
		form_layout.addRow(mainLayout)
		
		directory_label = QtGui.QLabel('Path')
		self.directory_path  = QtGui.QLineEdit(self)
		directory_button= QtGui.QPushButton('Directory', self)
		directory_button.clicked.connect(self.chooseDirectory)
		directory_layout = QtGui.QGridLayout()
		directory_layout.setSpacing(10)
		directory_layout.addWidget(directory_label    , 1, 0)
		directory_layout.addWidget(self.directory_path, 1, 1)
		directory_layout.addWidget(directory_button   , 1, 2)
		
		form_layout.addRow(directory_layout)
		
		search_label    = QtGui.QLabel('Keyword')
		self.search_keyword  = QtGui.QLineEdit(self)
		self.multiple_search = QtGui.QPushButton('Multi')
		self.multiple_search.clicked.connect(self.setMultiSearchOption)
		self.crossref_search = QtGui.QPushButton('Cross')
		self.crossref_search.clicked.connect(self.setCrossrefSearchOption)
		
		search_layout   = QtGui.QGridLayout()
		search_layout.addWidget(search_label  , 1, 0)
		search_layout.addWidget(self.search_keyword, 1, 1)
		search_layout.addWidget(self.multiple_search, 1, 2)
		search_layout.addWidget(self.crossref_search, 1, 3)
		
		form_layout.addRow(search_layout)
		
		self.kor_check= QtGui.QRadioButton('Kor', self)
		self.eng_check= QtGui.QRadioButton('Eng', self)
		lang_layout = QtGui.QHBoxLayout()
		lang_layout.addWidget(self.kor_check)
		lang_layout.addWidget(self.eng_check)
		
		form_layout.addRow("Keyword Language", lang_layout)

		self.korean_bible = QtGui.QComboBox(self)
		self.english_bible = QtGui.QComboBox(self)
		bible_layout = QtGui.QHBoxLayout()
		bible_layout.addWidget(self.korean_bible)
		bible_layout.addWidget(self.english_bible)
		form_layout.addRow("Choose Bible", bible_layout)
		
		self.check_ot      = QtGui.QCheckBox('OT', self)
		self.check_nt      = QtGui.QCheckBox('NT', self)
		self.wholeword     = QtGui.QCheckBox('Whole word', self)
		self.fullbook_name = QtGui.QCheckBox('FullBook Name', self)
		self.engbook       = QtGui.QCheckBox('EngBook Name', self)
		self.sort          = QtGui.QPushButton('Sort', self)
		option_layout      = QtGui.QHBoxLayout()
		self.sort.clicked.connect(self.sortList)
		self.fullbook_name.stateChanged.connect(self.toggleFullbookName)
		self.engbook      .stateChanged.connect(self.toggleEnglishbookName)
		option_layout.addWidget(self.check_ot)
		option_layout.addWidget(self.check_nt)
		option_layout.addWidget(self.wholeword)
		option_layout.addWidget(self.fullbook_name)
		option_layout.addWidget(self.engbook)
		option_layout.addWidget(self.sort)
		form_layout.addRow(option_layout)
	
		exit_button   = QtGui.QPushButton('Exit', self)
		view_button   = QtGui.QPushButton('View', self)
		excel_button  = QtGui.QPushButton('Excel', self)
		save_button   = QtGui.QPushButton('Save', self)
		clear_button   = QtGui.QPushButton('Clear', self)
		self.search_button = QtGui.QPushButton('Search', self)
		exit_button.clicked.connect(self.ExitProgram)
		self.search_button.clicked.connect(self.searchKeyword)
		excel_button.clicked.connect(self.saveExcel)
		view_button.clicked.connect(self.viewList)
		save_button.clicked.connect(self.saveVerseListAsHtmlAndJavascriptChart)
		clear_button.clicked.connect(self.clearMessageWindow)
		run_layout    = QtGui.QHBoxLayout()
		run_layout.addWidget(self.search_button)
		run_layout.addWidget(view_button)
		run_layout.addWidget(excel_button)
		run_layout.addWidget(save_button)
		run_layout.addWidget(clear_button)
		run_layout.addWidget(exit_button)
		form_layout.addRow(run_layout)

		# http://stackoverflow.com/questions/15561608/detecting-enter-on-a-qlineedit-or-qpushbutton
		self.search_keyword.returnPressed.connect(self.search_button.click)
		
		self.message = QtGui.QPlainTextEdit()
		# Plain Editor resize
		#http://stackoverflow.com/questions/13416000/qt-formlayout-not-expanding-qplaintextedit-vertically
		policy = self.sizePolicy()
		policy.setVerticalStretch(1)
		self.message.setSizePolicy(policy)
		form_layout.addRow(self.message)

		self.kor_check.setChecked(True)
		self.check_ot.setChecked(True)
		self.setWindowTitle("베들레헴 성경 단어 검색/엑셀그래프 출력")
		#self.setWindowIcon(QtGui.QIcon('bth-0.png'))
		self.setWindowIcon(QtGui.QIcon(QtGui.QPixmap(qbsicon.qbib_icon_table)))
		
		# Mar 01 2015
		# http://stackoverflow.com/questions/12432637/pyqt4-set-windows-taskbar-icon
		'''
		app_icon = QtGui.QIcon()
		app_icon.addPixmap(QtGui.QPixmap(qbs_icon_16.qbs_icon_16x16))
		app_icon.addPixmap(QtGui.QPixmap(qbs_icon_32.qbs_icon_32x32))
		app_icon.addPixmap(QtGui.QPixmap(qbs_icon_48.qbs_icon_48x48))
		app_icon.addPixmap(QtGui.QPixmap(qbs_icon_64.qbs_icon_64x64))
		self.setWindowIcon(app_icon)
		'''
		self.setLayout(form_layout)
		self.show()
				
	def setMultiSearchOption(self):
		option = MultipleSearchDialog(self.multisearch_option)
		option.exec_()
		
		self.multisearch_option.kbib,\
		self.multisearch_option.ebib,\
		self.multisearch_option.select,\
		self.multisearch_option.gbib,\
		self.multisearch_option.hbib = option.getSelectedItem()
		
		self.message.appendPlainText("... Hebrew Bible (BHS): {}".format(self.multisearch_option.hbib))
		self.message.appendPlainText("... SBL Greek New Testament: {}".format(self.multisearch_option.gbib))
		
	def setCrossrefSearchOption(self):
		option = CrossRefSearchDialog(self.crossrefsearch_option)
		option.exec_()
		return
		
	def toggleFullbookName(self):
		if self.fullbook_name.isChecked():
			self.message.appendPlainText("... Full Bible Book Name : ON")
		else:
			self.message.appendPlainText("... Full Bible Book Name : OFF")
		
	def toggleEnglishbookName(self):
		if self.engbook.isChecked():
			self.message.appendPlainText("... English Bible Book Name : ON")
		else:
			self.message.appendPlainText("... English Bible Book Name : OFF")
	
	def ExitProgram(self):
		writer = open(qbs_datafile,'w')
		form = "{0}\n{1}\n".format(self.destDir, self.search_keyword.text())
		writer.write(form)
		writer.close()
		self.close()
		
	def chooseDirectory(self):
		startingDir = os.getcwd() 
		self.destDir = QtGui.QFileDialog.getExistingDirectory(None, '베들레헴성경 폴더를 선택해주세요.', startingDir, QtGui.QFileDialog.ShowDirsOnly)
		if not self.destDir: return
		self.directory_path.setText(self.destDir)
		self.message.appendPlainText("... Folder path : {0}".format(self.destDir))

	def createPath(self):
		if self.kor_check.isChecked():
			self.output_infix = '-'+korean_bible_name[self.korean_bible.currentIndex()]			
			# Jan 24, 2015
			# http://stackoverflow.com/questions/16010992/how-to-use-directory-separator-in-both-linux-and-windows
			self.bdf_path = os.path.join(self.directory_path.text(),korean_bible_prefix[self.korean_bible.currentIndex()])
		elif self.eng_check.isChecked():
			self.output_infix = '-'+english_bible_name[self.english_bible.currentIndex()]
			self.bdf_path = os.path.join(self.directory_path.text(),english_bible_prefix[self.english_bible.currentIndex()])
		else:
			self.message.appendPlainText("... Error choose Kor/Eng")
			return
		self.message.appendPlainText("... Create path : {0}".format(self.bdf_path))
		
	def createFileList(self):
		self.fileList = []
		for i in range(max_bdf_number):
			self.fileList.append("{0}{1}{2}".format(self.bdf_path,i+1,bdf_ext))

	def searchKeyword(self):
		self.createPath()
		self.createFileList()
		
		ot = 0
		nt = 0
		self.eng_bookname = 0
		sort_hit = 0
		
		if self.check_ot.isChecked():
			ot = 1
		
		if self.check_nt.isChecked():
			nt = 1
			
		self.keyword = self.search_keyword.text()
		self.keyword.strip()
		self.SearchBibleKeyword(self.fileList, self.eng_bookname, self.keyword, ot, nt)
	
	def setRegularExpressionCaseOption(self):
		if self.search_option.case_sensitive:
			self.search_option.pattern_option = self.search_option.pattern_option ^ re.IGNORECASE
		else:
			self.search_option.pattern_option = self.search_option.pattern_option | re.IGNORECASE
			
	def SearchBibleKeyword(self, path, eng_bookname, keyword, ot, nt):
		nt_num = 40
		self.total_hit = 0
		nt_file_start = 5
		ot_file_end = 5
		
		self.hit_plot = dict()
		self.book_table_keys = list(book_table.keys())
		self.book_table_keys.sort()
		#book_names = list()
		self.setBookNameIndex()
			
		for i in range(0, 66):
			self.hit_plot[self.book_table_keys[i]] = Book_Hit(book_table[self.book_table_keys[i]],0,i)
		
		if self.eng_check.isChecked():
			if self.wholeword.isChecked():
				self.pattern = r'(?<!-)\b{0}\b(?!-)'.format(self.keyword)
			else:
				#self.pattern = r'\w*{0}\w*'.format(self.keyword)
				self.pattern = r'{}'.format(self.keyword)
		else:
			#if self.wholeword.isChecked():
			#	self.pattern = u'\b{0}\b'.format(self.keyword)
			#else:
			self.pattern = u'{}'.format(self.keyword)
		
		#writer = open(keyword+self.output_infix+'.vls', 'w')
		#writer.write("BWRL 1\n\n")
		
		if ot ==1 and nt == 0:
			self.message.appendPlainText('... Search keyword in OT')
			self.title = '검색어(구약): {0}'.format(keyword)
			start_file = 0
			end_file = ot_file_end
		elif ot == 0 and nt == 1:
			self.message.appendPlainText('... Search keyword in NT')
			self.title = '검색어(신약): {0}'.format(keyword)
			start_file = nt_file_start
			end_file = max_bdf_number
		else:
			self.message.appendPlainText('... Search keyword in OT & NT')
			self.title = '검색어(구약/신약): {0}'.format(keyword)
			start_file = 0
			end_file = max_bdf_number
		
		for i in range(start_file, end_file):
			reader = open(path[i], 'r')
			if not reader:
				err_msg = 'Cannot open {}'.format(path[i])
				QtGui.QMessageBox.question(self, 'Error', err_msg, QtGui.QMessageBox.Yes)
				return
			file_index = i
			file_pointer = 0
			for line in reader: 
				file_pointer += 1
				count = 0
				for count, match in enumerate(re.finditer(self.pattern, line, self.search_option.pattern_option)):
					count = count + 1
				if(count == 0): continue
				match = _find_bdfinfo.search(line)
				
				#key=line[:2]
				#book = book_table[key]
				#p1 = line.find(' ')
				#p2 = line.find(' ', p1+1)
				#chap = line[p1:p2]
				
				key = match.group(1)
				chap = match.group(2)
				vers = match.group(3)
				
				#writer.write("NAU {0} {1}\n".format(book[1], chap))
				#writer.write("{0} {1} {2}\n".format(book[self.index_bible_name], chap, count))
				self.hit_plot[key].book_hit += count
				self.hit_plot[key].verse.file_index.append(file_index) 
				self.hit_plot[key].verse.file_pointer.append(file_pointer)
				self.hit_plot[key].verse.number.append((key,chap,vers))
				
				"""
				Bethelem probram searches the keywrod from bdf files in such a way that
				If Bethelem finds any match first time, the number of occurence increases.
				It doesn't matter how many times the keyword appears in one verse.
				"""
				self.total_hit += count
			reader.close()
		#writer.close()
		self.message.appendPlainText("... Total hit ({0}) : {1}".format(keyword, self.total_hit))

		if self.total_hit == 0: return
		self.sorted = True
		self.sortList()
		self.preplot()
		
	def sortList(self):
		if not self.sorted :
			self.sorted = True
			self.message.appendPlainText("... Sorted Hit")
			self.sorted_hit = sorted(self.hit_plot.values(), key=lambda x:x.book_hit, reverse=True)
		else:
			self.message.appendPlainText("... Unsorted hit")
			self.sorted_hit = sorted(self.hit_plot.values(), key=lambda x:x.id)
			self.sorted = False

		self.valid_book_len=0
		self.valid_book_num=list()
		
		for i in range(0,66):
			if self.sorted_hit[i].book_hit > 0:
				self.valid_book_num.append(i)
				self.valid_book_len += 1		
			
		#self.preplot()
		
	def setBookNameIndex(self):
		if self.engbook.isChecked():
			if self.fullbook_name.isChecked(): self.index_bible_name = 3
			else: self.index_bible_name =  1
		else:
			if self.fullbook_name.isChecked(): self.index_bible_name = 2
			else: self.index_bible_name = 0

	def preplot(self):
		self.setBookNameIndex()
		self.x_pos = [0 for x in range(self.valid_book_len)]
		self.y_pos = [0 for x in range(self.valid_book_len)]
		self.excel_book_name = []

		for i in range(0,self.valid_book_len):
			book_obj = self.sorted_hit[self.valid_book_num[i]]
			self.excel_book_name.append("%s (%d)" % (book_obj.book_name[self.index_bible_name], book_obj.book_hit))
			self.x_pos[i] = i
			self.y_pos[i] = book_obj.book_hit
		
	def saveExcel(self):
		import xlsxwriter
		#print('... Create an Excel file and add a worksheet')
		self.preplot()
		
		try:
			excel_filename = self.keyword+self.output_infix+'.xlsx'
			workbook = xlsxwriter.Workbook(excel_filename)
			worksheet = workbook.add_worksheet()
			
			worksheet.write_column('A1', self.excel_book_name)
			worksheet.write_column('B1', self.y_pos)
			
			chart = workbook.add_chart({'type': 'column'})
			chart.set_title({'name': self.title+self.output_infix})
			chart.add_series(
				{
					'categories': '=Sheet1!$A${0}:$A${1}'.format(1, self.valid_book_len),
					'values': '=Sheet1!$B${0}:$B${1}'.format(1, self.valid_book_len)
				}
			)
			worksheet.insert_chart('D5', chart)
			workbook.close()
		except IOError as err:
			errno, strerror = err.args
			err_msg = "I/O error({0}): {1}\nFile is already open!".format(errno, strerror)
			QtGui.QMessageBox.question(self, 'Error', err_msg, QtGui.QMessageBox.Yes)
			self.message.appendPlainText('... Fail to create an Excel file and add a worksheet\n    {0}'.format(excel_filename))
			return
		self.message.appendPlainText('... Create an Excel file and add a worksheet\n    {0}'.format(excel_filename))

	def plotList(self):
		'''
		import numpy as np
		import matplotlib.pyplot as plt
		import matplotlib.font_manager as fm
		print('... Create a plot w/ Matplotlib')
		self.preplot()
		fp=fm.FontProperties(fname="NanumGothic.ttf", size=25)
		fig = plt.figure()	
		fig.suptitle(self.title, fontproperties=fp, fontweight='bold')
		
		x_pos_plot = np.asarray(self.x_pos)
		y_pos_plot = np.asarray(self.y_pos)
		
		a = fig.gca()
		a.set_xticklabels(a.get_xticks(), fontproperties=fp)
		plt.xticks(x_pos_plot, self.excel_book_name, rotation=90, size='small')
		plt.bar(x_pos_plot, y_pos_plot, align='center', color='b', alpha=0.4)
		#plt.bar(x_pos, y_pos, align='center', color='b', alpha=0.4)
		plt.show()
		'''
	def clearMessageWindow(self):
		self.message.clear()
		
	def viewList(self):
		self.setBookNameIndex()
		self.message.appendPlainText("... ========================")
		self.message.appendPlainText("    {0}\n    ========================".format(self.title))
		for i in range(0,self.valid_book_len):
			book_obj = self.sorted_hit[self.valid_book_num[i]]
			#if self.engbook.isChecked():
			self.message.appendPlainText("    {0} \t\t({1})".format(book_obj.book_name[self.index_bible_name], book_obj.book_hit))
			#else:
			#	self.message.appendPlainText("    {0} \t({1})".format(book_obj.book_name[0], book_obj.book_hit))
		self.message.appendPlainText("    ========================\n    \t\t{0}".format(self.total_hit))

	def saveVerseListAsHtmlAndJavascriptChart(self):

		ot_db_con = None
		ot_db_cur = None
		nt_db_con = None
		nt_db_cur = None
		hebrew_db_ok = False
		greek_db_ok = False

		bar_thk = 50
		width = 800
		height = 400
		haxis_label_fontsize = 10
		keyword_color = "red"

		hebrew_font_size = 20
		greek_font_size = 16
		
		if not self.export_option.chart_only:
			if self.multisearch_option.hbib:
				if not os.path.isfile(hebrew_bible_db):
					err_msg = 'Hebrew Bible Database({}) not exist !'.format(hebrew_bible_db)
					QtGui.QMessageBox.question(self, 'Error', err_msg, QtGui.QMessageBox.Yes)
				else: 
					hebrew_db_ok = True
					self.message.appendPlainText('... Open {}'.format(hebrew_bible_db))
					ot_db_con = db.connect(hebrew_bible_db)
					ot_db_cur = ot_db_con.cursor()
					self.message.appendPlainText('... Success')
			
			if self.multisearch_option.gbib:
				if not os.path.isfile(greek_bible_db):
					err_msg = 'SBL Greek NT Bible Database({}) not exist !'.format(hebrew_bible_db)
					QtGui.QMessageBox.question(self, 'Error', err_msg, QtGui.QMessageBox.Yes)
				else: 
					greek_db_ok = True
					self.message.appendPlainText('... Open {}'.format(greek_bible_db))
					nt_db_con = db.connect(greek_bible_db)
					nt_db_cur = nt_db_con.cursor()
					self.message.appendPlainText('... Success')
		
		list_fname = self.search_keyword.text()+self.output_infix
		print(list_fname)
		fw = open(list_fname+'.html', mode='w', encoding='utf-8')
		
		fw.write("<!DOCTYPE html>\n")
		fw.write("<html>\n")
		fw.write("<head>\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\n")
		
		fw.write("<style type=\"text/css\">\n")
		
		fw.write("span.hebrewtext { font-size:")
		fw.write("{}pt; line-height:150%; font-family:\"SBL Hebrew\";".format(hebrew_font_size))
		fw.write("color:black; direction:rtl; text-align:right; }\n")
		
		fw.write("span.greektext { font-size:")
		fw.write("{}pt; line-height:150%; font-family:\"SBL Greek\";".format(greek_font_size))
		fw.write("color:black; direction:rtl; text-align:right; }\n")
		
		#fw.write("span.engtext {font-family: \"굴림\"; font-size: 10pt}\n")
		
		fw.write("</style>\n")
		
		html_title = self.title+self.output_infix
		fw.write("<style> td { padding-top:10; padding-bottom:10;} </style>\n")
		fw.write("<title>{0}</title>\n".format(html_title))
		self.setBookNameIndex()
		qbs_barchart_color_table = []
		
		H2 = 240
		H1 = 0
		S = 0.8
		V = 1
		dH = (H2 - H1)/self.valid_book_len
		tempH = H1

		qbs_barchart_color_table = []
		for i in range(self.valid_book_len):
			r,g,b = TG_HSV_To_RGB(tempH, S, V)
			qbs_barchart_color_table.append(RGB(r,g,b))
			tempH = tempH+dH
						
		fw.write("<script src=\"amcharts/amcharts.js\" type=\"text/javascript\"></script>\n")
		fw.write("<script src=\"amcharts/serial.js\" type=\"text/javascript\"></script>\n")
		fw.write("<script src=\"amcharts/exporting/amexport.js\" type=\"text/javascript\"></script>\n")
		fw.write("<script src=\"amcharts/exporting/canvg.js\" type=\"text/javascript\"></script>\n")
		fw.write("<script src=\"amcharts/exporting/rgbcolor.js\" type=\"text/javascript\"></script>\n")
		fw.write("<script src=\"amcharts/exporting/filesaver.js\" type=\"text/javascript\"></script>\n")
		fw.write("<script type=\"text/javascript\">\n")
		fw.write("\tvar chart;\n\n")
		fw.write("\tvar chartData = [\n")
		
		for i in range(0,self.valid_book_len):
			book_obj = self.sorted_hit[self.valid_book_num[i]]
			#print(book_obj.book_name[0], book_obj.book_hit)
			fw.write("\t\t\t { \"book\": \""+book_obj.book_name[self.index_bible_name]+"({0})\", ".format(book_obj.book_hit))
			fw.write("forceShow:true, \"hits\": "+str(book_obj.book_hit)+",") 
			color = qbs_barchart_color_table[i]
			fw.write("\"color\" : \"rgb({0},{1},{2})\", ".format(color.r, color.g, color.b))
			fw.write("\"balloonValue\": "+str(book_obj.book_hit)+"},\n")

		fw.write("\t\t\t];\n\n")
		fw.write("\tvar chart = AmCharts.makeChart(\"chartdiv\", {\n")
		fw.write("\t\ttheme: \"none\",\n")
		fw.write("\t\ttype: \"serial\",\n")
		fw.write("\t\tdataProvider: chartData,\n")
		fw.write("\t\tcategoryField: \"book\",\n")
		fw.write("\t\tstartDuration : 1,\n")
		fw.write("\t\t\"titles\": [{\n")
		fw.write("\t\t\"text\": \"{0}\\nTotal Hit = {1}\",\n".format(self.title+self.output_infix, self.total_hit))
		fw.write("\t\t\t\"size\": 15\n\t\t\t}],\n\n")
		
		fw.write("\tcategoryAxis : {\n")
		fw.write("\t\tlabelRotation : 90,\n")
		fw.write("\t\tgidAlpha : 0,\n")
		fw.write("\t\tfillAlpha : 1,\n")
		fw.write("\t\tfontSize : {},\n".format(self.chartOption.xlabel_size))
		fw.write("\t\tparseDates : false,\n")
		fw.write("\t\tforceShowField : \"forceShow\",\n")
		fw.write("\t\tfillColor : \"#FAFAFA\",\n")
		fw.write("\t\tgridPosition : \"start\"\t\t\n\t\t},\n\n")

		fw.write("\tvalueAxis: [{\n")
		fw.write("\t\tdashLength : 5,\n")
		fw.write("\t\ttitle : \"Keyword Hits\",\n")
		fw.write("\t\tintegersOnly:true,\n")
		fw.write("\t\taxisAlpha : 0,\n")
		fw.write("\t}],\n\n")

		fw.write("\tgraphs: [{\n")
		fw.write("\t\tvalueField : \"hits\",\n")
		fw.write("\t\tcolorField : \"color\",\n")
		fw.write("\t\tballoonText : \"<b>[[category]]</b>\",\n")
		fw.write("\t\tlabelText : \"[[balloonValue]]\",\n")
		fw.write("\t\tlabelPosition : \"top\",\n")
		fw.write("\t\tfontSize : {},\n".format(self.chartOption.ylabel_size))
		fw.write("\t\ttype : \"column\",\n")
		fw.write("\t\tlineAlpha : 0,\n")
		fw.write("\t\tfillAlphas : 1\n")
		fw.write("\t}],\n\n")

		fw.write("\tchartCursor: {\n")
		fw.write("\t\tcursorAlpha : 0,\n")
		fw.write("\t\tzoomable : false,\n")
		fw.write("\t\tcategoryBalloonEnabled : false\n")
		fw.write("\t},\n")
		
		fw.write("\tpathToImages: \"amcharts/images/\",\n")
		fw.write("\tamExport: {\n")
		fw.write("\t\timageFileName: \"{0}\",\n".format(list_fname))
		fw.write("\t\ttop: 0,\n")
		fw.write("\t\tright: 0,\n")
		fw.write("\t\texportJPG: true,\n")
		fw.write("\t\texportPNG: true,\n")
		fw.write("\t\texportSVG: true\n\t}\n")
		fw.write("});\n")
		
		fw.write("</script>\n")
		fw.write("</head>\n")
		
		fw.write("<body>\n")
		fw.write("<div id=\"chartdiv\" style=\"width: {0}px; height:"\
		         " {1}px;\"></div>\n".format(self.chartOption.width, self.chartOption.height))
		fw.write("</body>\n")
		
		if self.export_option.chart_only:
			fw.write("</html>\n")
			fw.close()
			self.message.appendPlainText("... Saving Plot Only")
			return

		fw.write("<span style=\"font-family: 굴림; font-size: 10\">\n")
		fw.write("<style>table { table-layout: fixed  }</style>\n")
		fw.write("<table border = \"0\" cellpadding=5>\n") # style=\"border-collapse: collapse\">\n")
		fw.write("<caption>성경 구절</caption>\n")
		fw.write("<tr>\n")
		
		if self.eng_check.isChecked():
			fw.write("<tr><td>&nbsp</td><td>&nbsp</td><td align=center>{0}</td>".\
			format(english_bible_name[self.english_bible.currentIndex()]))
		else:
			fw.write("<tr><td>&nbsp</td><td>&nbsp</td><td align=center>{0}</td>".\
			format(korean_bible_name[self.english_bible.currentIndex()]))
			
		if self.multisearch_option.kbib > 0 and self.multisearch_option.select > 0:
			fw.write("<td></td><td align=center>{0}</td>".\
			format(korean_bible_name[self.multisearch_option.select-1]))
		elif self.multisearch_option.ebib > 0 and self.multisearch_option.select > 0:
			fw.write("<td></td><td align=center>{0}</td>".\
			format(english_bible_name[self.multisearch_option.select-1]))
			
		if self.multisearch_option.hbib or self.multisearch_option.gbib:
			fw.write("<td></td><td align=center>BHS(1983)/SBLGNT</td>")
			
		fw.write("</tr>")
		
		for i in range(0,self.valid_book_len):
			book_obj = self.sorted_hit[self.valid_book_num[i]]
			length = len(book_obj.verse.file_index)
			jj = 0
			index1 = 0
			
			while jj < length:
				findex1 = book_obj.verse.file_index[jj]
				kk = jj+1
	    
				if length == 1:
					index2 = 1
					jj += 1
				else:
					while kk < length:
						findex2 = book_obj.verse.file_index[kk]
						if findex1 == findex2: 
							jj += 1
							kk += 1
							if kk == length:
								index2 = kk
								jj = kk
								break
							else: continue
						else:
							jj = kk
							index2 = kk
							break

				try:
					fname = self.fileList[book_obj.verse.file_index[index1]]
					reader = open(fname)
				except IOError as err:
					errno, strerror = err.args
					self.message.appendPlainText("... I/O error({0}): {1}".format(errno, strerror))
					fw.close()
					return 			

				ext_reader = None
				
				if self.multisearch_option.kbib > 0 and self.multisearch_option.select > 0:
					msearch_path = os.path.join(self.directory_path.text(),\
					korean_bible_prefix[self.multisearch_option.select-1])
					fname = "{0}{1}{2}".format(msearch_path,book_obj.verse.file_index[index1]+1,bdf_ext)
					ext_title = korean_bible_name[self.multisearch_option.select-1]
					ext_reader = open(fname)
				elif self.multisearch_option.ebib > 0 and self.multisearch_option.select:
					msearch_path = os.path.join(self.directory_path.text(),\
					english_bible_prefix[self.multisearch_option.select-1])
					fname = "{0}{1}{2}".format(msearch_path,book_obj.verse.file_index[index1]+1,bdf_ext)
					ext_title = english_bible_name[self.multisearch_option.select-1]
					ext_reader = open(fname)

				file_pointer = -1
				fp1 = 0
				for l in range(index1, index2):
					fp2 = book_obj.verse.file_pointer[l]
					dfp = fp2-fp1-1
					fp3 = 0
					while fp3 < dfp: 
						temp = reader.readline()
						fp3 += 1
					fp1 = fp2
					line = reader.readline()
					key=line[:2]
					book = book_table[key]
					item = re.search(r'(\d+):(\d+)', line)
					chap = item.group(1)
					vers = item.group(2)
					
					fw.write("<tr><td nowrap>{0}</td><td>{1}:{2}</td>".format(book[self.index_bible_name], chap, vers))
					fw.write("<td style=\"border-left:0px; border-right:0px;"\
					         " border-top:1px dotted; border-bottom:0px none\" bordercolor=#CCCCCC\">")
							 
					keyword = self.search_keyword.text()
					pos1 = item.end(2)
					match_list = re.finditer(self.pattern, line, self.search_option.pattern_option)
					
					for match_item in match_list:
						pos2 = match_item.start()
						pos3 = match_item.end()
						if pos2 == pos1:
							fw.write("<b><font color=\"{0}\">{1}</font></b>".format(keyword_color, line[pos2:pos3]))
						elif pos2 > pos1:
							fw.write(line[pos1:pos2])
							fw.write("<b><font color=\"{0}\">{1}</font></b>".format(keyword_color, line[pos2:pos3]))
						pos1 = pos3
					fw.write(line[pos1:])
					fw.write("</td>")

					if ext_reader:
						for temp in ext_reader:
							ext_key=temp[:2]
							item = re.search(r'(\d+):(\d+)', temp)
							if not item: continue
							ext_chap = item.group(1)
							ext_vers = item.group(2)
							if key == ext_key and chap == ext_chap and vers == ext_vers:
								fw.write("<td>&nbsp&nbsp</td><td style=\"border-left:0px;"\
								" border-right:0px; border-top:1px dotted; border-bottom:0px none"\
								"\" bordercolor=#CCCCCC\">{}</td>".format(temp[item.end(2):]))
								#"\" bordercolor=#CCCCCC\"><span class=\"engtext\">{}</span></td>".format(temp[item.end(2):]))
								break
										
					if self.multisearch_option.hbib and hebrew_db_ok and int(key) <= number_of_ot_books:
						map_key = '{}:{}:{}'.format(key, chap, vers)
						map_str = ''
						if map_key in wmap.wtt_table:
							map_info = wmap.wtt_table[map_key]
							map_chap = map_info[0]
							map_vers = map_info[1]
							map_str  = '(BHS {}:{})'.format(map_chap, map_vers)
						else:
							map_chap = int(chap)
							map_vers = int(vers)
						
						sql = 'SELECT vtext from {} where book = {} and chap = {} and verse = {}'.\
                        format(hebrew_bible_db_table_name, int(key), map_chap, map_vers)
						for vtext in ot_db_cur.execute(sql):
							fw.write("<td>&nbsp&nbsp</td><td style=\"border-left:0px;"\
							         "border-right:0px; border-top:1px dotted; border-bottom:0px none"\
									 "\"bordercolor=#CCCCCC\"><span class=\"hebrewtext\">{}</span> {}</td>".\
									 format(vtext[0], map_str))

					if self.multisearch_option.gbib and greek_db_ok and\
					   (number_of_ot_books < int(key) <= number_of_bible_books):
						sql = 'SELECT vtext from {} where book = {} and chap = {} and verse = {}'.\
                        format(greek_bible_db_table_name, int(key), int(chap), int(vers))
						for vtext in nt_db_cur.execute(sql):
							fw.write("<td>&nbsp&nbsp</td><td style=\"border-left:0px;"\
							         "border-right:0px; border-top:1px dotted; border-bottom:0px none"\
									 "\" bordercolor=#CCCCCC\"><span class=\"greektext\">{}</span></td>".format(vtext[0]))
					fw.write("</tr>\n")
				index1 = index2
		fw.write("</tr>\n</table>")
		fw.write("</html>\n")
		fw.close()

		if hebrew_db_ok: 
			self.message.appendPlainText('... Close {}'.format(hebrew_bible_db))
			ot_db_con.close()

		if greek_db_ok: 
			self.message.appendPlainText('... Close {}'.format(greek_bible_db))
			nt_db_con.close()
		
		self.message.appendPlainText("... Saving Plot (Javascript Chart) & Verse list")
		self.message.appendPlainText("    Verst list is saved in {}.html".format(list_fname))
		#self.message.appendPlainText("\n    TOTAL HIT = {0}".format(count))
		
		#if self.call_browser:
		#	webbrowser.get('google-chrome').open(list_fname)
			
	def saveList(self):
		self.setBookNameIndex()
		findex1 = 0
		findex2 = 0
		index1 = 0
		index2 = 0 
		list_fname = self.search_keyword.text()+self.output_infix+'.txt'
		writer = open(list_fname, 'w')
		
		for i in range(0,self.valid_book_len):
			book_obj = self.sorted_hit[self.valid_book_num[i]]
			length = len(book_obj.verse.file_index)
			jj = 0
			index1 = 0
			
			while jj < length:
				findex1 = book_obj.verse.file_index[jj]
				kk = jj+1
	    
				if length == 1:
					index2 = 1
					jj += 1
				else:
					while kk < length:
						findex2 = book_obj.verse.file_index[kk]
						if findex1 == findex2: 
							jj += 1
							kk += 1
							if kk == length:
								index2 = kk
								jj = kk
								break
							else: continue
						else:
							jj = kk
							index2 = kk
							break

				try:
					fname = self.fileList[book_obj.verse.file_index[index1]]
					#self.message.appendPlainText("... Open {0})".format(fname))
					reader = open(fname)
				except IOError as err:
					errno, strerror = err.args
					self.message.appendPlainText("... I/O error({0}): {1}".format(errno, strerror))
					return 			
					
				file_pointer = -1
				fp1 = 0
				for l in range(index1, index2):
					fp2 = book_obj.verse.file_pointer[l]
					#print("l = {0} fp2 = {1}".format( l, fp2))
					dfp = fp2-fp1-1
					fp3 = 0
					while fp3 < dfp: 
						temp = reader.readline()
						fp3 += 1
					fp1 = fp2
					#self.message.appendPlainText(reader.readline())
					writer.write(reader.readline()[2:])
										
				index1 = index2
		self.message.appendPlainText("... Verst list is saved in {}.html".format(list_fname))

	def createDBList(self):
		self.kbib_db = []
		self.ebib_db = []
		
		for i in range(len(korean_bible_prefix)):
			sql_file = "{0}{1}".format(korean_bible_prefix[i],sql_ext)
			self.kbib_db.append(sql_file)

		for i in range(len(english_bible_prefix)):
			sql_file = "{0}{1}".format(english_bible_prefix[i],sql_ext)
			self.ebib_db.append(sql_file)
				
	def checkDB(self):
		if not os.path.isfile("{0}{1}".format(korean_bible_prefix[0],sql_ext)):
			self.convertBDFToSQL()
		'''
		found = False
		
		for i in range(len(korean_bible_prefix)):
			if not os.path.exists(self.kbib_db[i]):
				found = True
				break
				
		for i in range(len(english_bible_prefix)):
		'''
	def bdfTosql(self, fbdf, fsql):
		import sqlite3 as db
		
		db_con = db.connect(fsql)
		db_cur = db_con.cursor()
		db_cur.execute("CREATE TABLE bible(book INTEGER, chap INTEGER, verse INTEGER, vtext TEXT)")
		
		for i in range(len(fbdf)):
			self.message.appendPlainText("... Convert {0} to SQL".format(fbdf[i]))
			#print("... Convert {0} to SQL".format(fbdf[i]))
			reader = open(fbdf[i])
			line_count = 0
			
			while 1:
				line = reader.readline()
				line_count = line_count+1
				if not line: break
				if not line.find(':'): continue
				line = line.replace('\n', "")
				book=line[:2]
				book_tmp = book
				book.strip('0')
				p1 = line.find(' ')
				p2 = line.find(' ', p1+1)
				pos = line[p1+1:p2]
				dummy = pos.split(':')
				chap = dummy[0]
				verse = dummy[-1]
				try:
					db_cur.execute("INSERT INTO bible VALUES({0}, {1}, {2}, \"{3}\")".\
					format(int(book),int(chap), int(verse), line[p2+1:]));
				except ValueError:
					print("... Value ERROR at {0} {1} {2} {3} of {4}".\
					format(book_table[book_tmp][2], chap, verse, line_count, fbdf[i]))
					reader.close()
					return
				except db.OperationalError as e:
					#self.message.appendPlainText("... ERROR at {0} {1} {2} {3}".format(book_table[book_tmp][2], chap, verse, line_count))
					print("... ERROR at {0} {1} {2} {3} of {4}\n    {5}".\
					format(book_table[book_tmp][2], chap, verse, line_count, fbdf[i], e.args[0]))
					reader.close()
					return
			reader.close()
		db_con.commit()
		db_con.close()
		
		self.message.appendPlainText("... CREATE {0}: success".format(fsql))
		
	def convertBDFToSQL(self):
		# convert Korean Bible

		self.message.appendPlainText("... Convert Korean Bible data file to SQL")
		for i in range(len(korean_bible_prefix)):
			bdf_file = []
			for j in range(max_bdf_number):
				bdf_path = self.directory_path.text()+'\\'+korean_bible_prefix[i]
				bdf_file.append("{0}{1}{2}".format(bdf_path,j+1,bdf_ext))
			self.bdfTosql(bdf_file, self.kbib_db[i])
			
		#convert English Bible
		self.message.appendPlainText("... Convert English Bible data file to SQL")
		for i in range(len(english_bible_prefix)):
			bdf_file = []
			for j in range(max_bdf_number):
				bdf_path = self.directory_path.text()+'\\'+english_bible_prefix[i]
				bdf_file.append("{0}{1}{2}".format(bdf_path,j+1,bdf_ext))
			self.bdfTosql(bdf_file, self.ebib_db[i])

	# Sep 21 2016
	def processCreateBibleDB(self):
		bible = [korean_bible_prefix, english_bible_prefix]
		
		for i in range(len(bible)):
			bible_prefix = bible[i]
			for j in range(len(bible_prefix)):
				db_prefix = bible_prefix[j]
				db_name = db_prefix+'.db'
				if os.path.isfile(db_name): os.remove(db_name)
				#self.message.appendPlainText("... Open {}".format(db_name))
				db_con = db.connect(db_name)
				db_cur = db_con.cursor()
				db_cur.execute("CREATE TABLE {}(book INT, chap INT, verse INT, vtext TEXT)".format(db_prefix))
				#self.message.appendPlainText("... Create {} Table".format(db_prefix))
			
				for k in range(0, max_bdf_number):
					bdf_file = os.path.join(self.directory_path.text(), bible_prefix[k])+'{}{}'.format(k+1,bdf_ext)
					reader = open(bdf_file, 'r')
					if not reader:
						self.message.appendPlainText("... Error can't open {}".format(self.fileList[k]))
						return
						
					for line in reader: 
						match = _find_bdfinfo.search(line)
						if match:
							#book = book_table[key]
							book = match.group(1)
							chap = match.group(2)
							vers = match.group(3)
							tpos = match.end(3)
							text = line[tpos:].replace('\n','')
							text = text.replace('\"','\'')
							#print('{} {}:{}'.format(book, chap, vers))
							try:
								db_cur.execute("INSERT INTO {} VALUES('{}', '{}', '{}', \"{}\")".format(db_prefix, book, chap, vers, text))
							except db.OperationalError as err:
								#err1, err2, err3 = err.args
								#print("... SQlite3 error --> ({0}): {1}".format(errno, strerror))
								print("{} --> {} {}:{} {}".format(bible_prefix[k], book, chap, vers, text))
								db_con.commit()
								db_con.close()
								#break
								return
				db_con.commit()
				db_con.close()
				#self.message.appendPlainText("... Close {}".format(db_name))
	
	# 10/17/2017
	def process_bibleworks_exported_verlist(self):
		cb_data = ""
		cb.OpenClipboard()
		cb_data = cb.GetClipboardData(win32con.CF_TEXT)
		cb.CloseClipboard()
		verselist = cb_data.decode('utf-8').split('\n')
		if not len(verselist):
			QtGui.QMessageBox.question(None,'Error', 'No Clipboard Data')
			return
		
		self.hit_plot = dict()
		book_key_dict = dict()
		book_table_keys = list(book_table.keys())
		book_table_keys.sort()
		
		for i in range(0, 66):
			value = book_table[book_table_keys[i]]
			book_key_dict[value[4]] = book_table_keys[i]
				
		for i in range(0, 66):
			self.hit_plot[book_table_keys[i]] = Book_Hit(book_table[book_table_keys[i]],0,i)
	
		self.total_hit = 0
		for v in verselist:
			match = _find_bwverse.search(v)
			if match:
				#print(v)
				book = match.group(1)
				chap = match.group(2)
				verse = match.group(3)[:-1]
				vlist = verse.split(',')
				key = book_key_dict[book]
				#ifile = find_fileindex(key)
				self.hit_plot[key].book_hit += len(vlist)
				self.total_hit += len(vlist)
				#hit_plot[key].verse.file_index.append(ifile) 
				#hit_plot[key].verse.file_pointer.append(vlist)
				#hit_plot[key].verse.number.append((key,chap,verse))
				self.hit_plot[key].verse.number.append([chap, vlist])
				#print('{}, {}, {}, {} '.format(book, key, chap, vlist))
			else: print('no more data')

		temp_sort = self.sorted
		self.sorted = True #not self.sorted
		self.sortList()
		
		self.sorted = temp_sort
		temp_option = self.export_option.chart_only
		temp_keyword = self.search_keyword.text()
		self.output_infix = ''
		self.title = ''
		self.export_option.chart_only = True
		self.search_keyword.setText('aaaaaaa')
		self.saveVerseListAsHtmlAndJavascriptChart()
		self.search_keyword.setText(temp_keyword)
		self.export_option.chart_only = temp_option		

	# July 08 2015
	def processClipboardBWPlotData(self):
		cb_data = ""
		cb.OpenClipboard()
		cb_data = cb.GetClipboardData(win32con.CF_TEXT)
		cb.CloseClipboard()

		match_list = re.finditer(r'\"(.*)\"\s(\d+)\s(\d+)', cb_data.decode('cp949'))
		if not match_list:
			err_string = "Error: No Clipboard Data"
			self.message.appendPlainText(err_string)
			QtGui.QMessageBox.question(None,'Error', err_string)
		
		book_to_key = dict()
		self.hit_plot = dict()
		self.book_table_keys = list(book_table.keys())
		self.book_table_keys.sort()
		self.setBookNameIndex()
			
		for i in range(0, 66):
			one_record_of_book_table = book_table[self.book_table_keys[i]]
			self.hit_plot[self.book_table_keys[i]] = Book_Hit(book_table[self.book_table_keys[i]],0,i)
			book_to_key[one_record_of_book_table[1]] = self.book_table_keys[i]

		self.total_hit = 0
		for match_item in match_list:
			book = match_item.group(1)
			hits = match_item.group(2)
			self.total_hit = self.total_hit + int(hits)
			try:
				self.hit_plot[book_to_key[book]].book_hit = int(hits)
			except KeyError:
				self.message.appendPlainText("Warning: {} is one of Apocryphal Books".format(book))

		if self.total_hit is 0:
			QtGui.QMessageBox.question(None,'Error', 'Error: check if clipboard data is valid')
			return
			
		temp_sort = self.sorted
		self.sorted = True #not self.sorted
		self.sortList()
		self.sorted = temp_sort
		temp_option = self.export_option.chart_only
		temp_keyword = self.search_keyword.text()
		self.output_infix = ''
		self.title = ''
		self.export_option.chart_only = True
		self.search_keyword.setText('clipboardBW')
		self.saveVerseListAsHtmlAndJavascriptChart()
		self.search_keyword.setText(temp_keyword)
		self.export_option.chart_only = temp_option		

def main():
	# Mar 01 2015
	# http://stackoverflow.com/questions/1551605/how-to-set-applications-taskbar-icon-in-windows-7/1552105#1552105
	#myappid = 'QBibSearch.V1.0' # arbitrary string
	#ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

	app = QtGui.QApplication(sys.argv)
	# Mar 01 2015
	# http://stackoverflow.com/questions/12432637/pyqt4-set-windows-taskbar-icon
	#app_icon = QtGui.QIcon()
	#app_icon.addPixmap(QtGui.QPixmap(qbs_icon_16.qbs_icon_16x16))
	#app_icon.addPixmap(QtGui.QPixmap(qbs_icon_32.qbs_icon_32x32))
	#app_icon.addPixmap(QtGui.QPixmap(qbs_icon_48.qbs_icon_48x48))
	#app_icon.addPixmap(QtGui.QPixmap(qbs_icon_64.qbs_icon_64x64))
	#app.setWindowIcon(app_icon)
	#QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'Plastique'))
	QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'CDE'))
	#QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'Motif'))
	#QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'Cleanlooks'))
	bib_search = QBibSearch()
	sys.exit(app.exec_())
	
if __name__ == '__main__':
	main()
		
