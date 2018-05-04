#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper
import calendar
from com.sun.star.datatransfer import XTransferable
from com.sun.star.datatransfer import DataFlavor  # Struct
from com.sun.star.datatransfer import UnsupportedFlavorException
from com.sun.star.lang import Locale  # Struct
from com.sun.star.sheet import CellFlags  # 定数
COLORS = {\
# 		"lime": 0x00FF00,\
		"magenta3": 0xFF00FF,\
		"black": 0x000000,\
		"blue3": 0x0000FF,\
		"skyblue": 0x00CCFF,\
		"silver": 0xC0C0C0,\
		"red3": 0xFF0000,\
		"violet": 0x9999FF,\
		"cyan10": 0xCCFFFF}  # 色の16進数。	
HOLIDAYS = {\
		"2018":((1,2,3,8),(11,12),(21,),(29,30),(3,4,5),(),(16,),(11,),(17,23,24),(8,),(3,23),(23,24,28,29,30,31)),\
		"2019":((1,2,3,14),(11,),(21,),(29,),(3,4,5,6),(),(15,),(11,12),(16,23),(14,),(3,4,23),(23,28,29,30,31)),\
		"2020":((1,2,3,13),(11,),(20,),(29,),(3,4,5,6),(),(20,),(11,),(21,22),(12,),(3,23),(23,28,29,30,31)),\
		"2021":((1,2,3,11),(11,),(20,),(29,),(3,4,5,),(),(19,),(11,),(20,23),(11,),(3,23),(23,28,29,30,31)),\
		"2022":((1,2,3,10),(11,),(21,),(29,),(3,4,5),(),(18,),(11,),(19,23),(10,),(3,23),(23,28,29,30,31)),\
		"2023":((1,2,3,9),(11,),(21,),(29,),(3,4,5),(),(17,),(11,),(18,23),(9,),(3,23),(23,28,29,30,31)),\
		"2024":((1,2,3,8),(11,12),(20,),(29,),(3,4,5,6),(),(15,),(11,12),(16,22,23),(14,),(3,4,23),(23,28,29,30,31)),\
		"2025":((1,2,3,13),(11,),(20,),(29,),(3,4,5,6),(),(21,),(11,),(15,23),(13,),(3,23,24),(23,28,29,30,31)),\
		"2026":((1,2,3,12),(11,),(20,),(29,),(3,4,5,6),(),(20,),(11,),(21,22,23),(12,),(3,23),(23,28,29,30,31)),\
		"2027":((1,2,3,11),(11,),(21,22),(29,),(3,4,5),(),(19,),(11,),(20,23),(11,),(3,23),(23,28,29,30,31)),\
		"2028":((1,2,3,10),(11,),(20,),(29,),(3,4,5),(),(17,),(11,),(18,22),(9,),(3,23),(23,28,29,30,31)),\
		"2029":((1,2,3,8),(11,12),(20,),(29,30),(3,4,5),(),(16,),(11,),(17,23,24),(8,),(3,23),(23,24,28,29,30,31)),\
		"2030":((1,2,3,14),(11,),(20,),(29,),(3,4,5,6),(),(15,),(11,12),(16,23),(14,),(3,4,23),(23,28,29,30,31))}  # 祝日JSON
class TextTransferable(unohelper.Base, XTransferable):
	def __init__(self, txt):  # クリップボードに渡す文字列を受け取る。
		self.txt = txt
		self.unicode_content_type = "text/plain;charset=utf-16"
	def getTransferData(self, flavor):
		if flavor.MimeType.lower()!=self.unicode_content_type:
			raise UnsupportedFlavorException()
		return self.txt
	def getTransferDataFlavors(self):
		return DataFlavor(MimeType=self.unicode_content_type, HumanPresentableName="Unicode Text"),  # DataTypeの設定方法は不明。
	def isDataFlavorSupported(self, flavor):
		return flavor.MimeType.lower()==self.unicode_content_type
def formatkeyCreator(doc):  # ドキュメントを引数にする。
	def createFormatKey(formatstring):  # formatstringの書式はLocalによって異なる。 
		numberformats = doc.getNumberFormats()  # ドキュメントのフォーマット一覧を取得。デフォルトのフォーマット一覧はCalcの書式→セル→数値でみれる。
		locale = Locale(Language="ja", Country="JP")  # フォーマット一覧をくくる言語と国を設定。インストールしていないUIの言語でもよい。。 
		formatkey = numberformats.queryKey(formatstring, locale, True)  # formatstringが既存のフォーマット一覧にあるか調べて取得。第3引数のブーリアンは意味はないはず。 
		if formatkey == -1:  # デフォルトのフォーマットにformatstringがないとき。
			formatkey = numberformats.addNew(formatstring, locale)  # フォーマット一覧に追加する。保存はドキュメントごと。 
		return formatkey
	return createFormatKey
def createKeikaSheet(doc, sheet, ids, dateserial):  # 経過シートの作成。
	createFormatKey = formatkeyCreator(doc)	
	sheet["F2"].setString(" ".join(ids))  # ID漢字名ｶﾅ名を入力。
	daycount = 100  # 経過シートに入力する日数。
	celladdress = sheet["I2"].getCellAddress()  # 経過シートの日付の開始セルのセルアドレスを取得。
	r, c = celladdress.Row, celladdress.Column
	sheet[:r+1, c:].clearContents(CellFlags.VALUE+CellFlags.DATETIME+CellFlags.STRING+CellFlags.ANNOTATION+CellFlags.FORMULA+CellFlags.HARDATTR+CellFlags.STYLES)  # セルの内容を削除。
	endcolumn = c + daycount + 1
	
	
	sheet[r, c:endcolumn].setDataArray(([i for i in range(dateserial, dateserial+daycount+1)],))  # 日時シリアル値を経過シートに入力。
	
	
	
	sheet[r, c:endcolumn].setPropertyValue("NumberFormat", createFormatKey('YYYY/M/D'))  # 日時シリアルから年月日の取得のため一時的に2018/5/4の形式に変換する。
	y, m, d = sheet[r, c].getString().split("/")  # 年、月、日を文字列で取得。
	weekday, days = calendar.monthrange(y, m)  # 日曜日が曜日番号0。1日の曜日と一月の日数のタプルが返る。
	weekday = (weekday+(d-1)%7)%7  # dの曜日番号を取得。1日からの7の余りと1日の余りを加えた7の余りがdの曜日番号。
	n = 0  # 日曜日の曜日番号。
	sundayranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # 日曜日のセル範囲コレクション。
	[sundayranges.addRangeAddress(sheet[r, i].getRangeAddress()) for i in range(c+(n-weekday)%7, endcolumn, 7)]  # 曜日番号nの列番号だけについて。
	n = 6  # 土曜日の曜日番号。
	saturdayranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # 土曜日のセル範囲コレクション。
	[saturdayranges.addRangeAddress(sheet[r, i].getRangeAddress()) for i in range(c+(n-weekday)%7, endcolumn, 7)]  # 曜日番号nの列番号だけについて。
	holidayranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # 祝日のセル範囲コレクション。
	holidays = HOLIDAYS  # 祝日の辞書を取得。
	days = days - d + 1  # 翌月1日までの日数を取得。
	mr = r - 1  # 月を代入する行のインデックス。
	mc = c  # 1日を表示する列のインデックス。
	if y in holidays:  # 祝日一覧のキーがある時。
		[holidayranges.addRangeAddress(sheet[r, mc+i-1].getRangeAddress()) for i in holidays[y][m] if not i<d]
	while True:
		sheet[mr, mc].setString("{}月".format(m))  # 月を入力。
		mc += days  # 次月1日の列に進める。
		if mc<endcolumn:  # 日時シリアル値が入力されている列の時。
			ymd = sheet[r, mc].getString()  # 1日の年/月/日を取得。
			y, m = ymd.split("/")[:2]  # 年と月を取得。
			if y in holidays:  # 祝日一覧のキーがある時。。
				[holidayranges.addRangeAddress(sheet[r, mc+i-1].getRangeAddress()) for i in holidays[y][m] if mc+i-1<endcolumn]
			weekday, days = calendar.monthrange(y, m)  # 1日の曜日と月の日数を取得。
		else:
			break
	sheet[r, c:endcolumn].setPropertyValue("NumberFormat", createFormatKey('D'))  # 経過シートの日付の書式を日だけにする。
	holidayranges.setPropertyValue("CellBackColor", COLORS["red3"])  # 祝日の背景色を変更。
	sundayranges.setPropertyValue("CharColor", COLORS["red3"])  # 日曜日の文字色を変更。
	saturdayranges.setPropertyValue("CharColor", COLORS["skyblue"])  # 土曜日の文字色を変更。	
	return sheet

	