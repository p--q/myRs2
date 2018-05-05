#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper
# import calendar
from com.sun.star.datatransfer import XTransferable
from com.sun.star.datatransfer import DataFlavor  # Struct
from com.sun.star.datatransfer import UnsupportedFlavorException
from com.sun.star.lang import Locale  # Struct
# from com.sun.star.sheet import CellFlags  # 定数
# from com.sun.star.table.CellHoriJustify import CENTER  # enum
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
		2018:((1,2,3,8),(11,12),(21,),(29,30),(3,4,5),(),(16,),(11,),(17,23,24),(8,),(3,23),(23,24,28,29,30,31)),\
		2019:((1,2,3,14),(11,),(21,),(29,),(3,4,5,6),(),(15,),(11,12),(16,23),(14,),(3,4,23),(23,28,29,30,31)),\
		2020:((1,2,3,13),(11,),(20,),(29,),(3,4,5,6),(),(20,),(11,),(21,22),(12,),(3,23),(23,28,29,30,31)),\
		2021:((1,2,3,11),(11,),(20,),(29,),(3,4,5,),(),(19,),(11,),(20,23),(11,),(3,23),(23,28,29,30,31)),\
		2022:((1,2,3,10),(11,),(21,),(29,),(3,4,5),(),(18,),(11,),(19,23),(10,),(3,23),(23,28,29,30,31)),\
		2023:((1,2,3,9),(11,),(21,),(29,),(3,4,5),(),(17,),(11,),(18,23),(9,),(3,23),(23,28,29,30,31)),\
		2024:((1,2,3,8),(11,12),(20,),(29,),(3,4,5,6),(),(15,),(11,12),(16,22,23),(14,),(3,4,23),(23,28,29,30,31)),\
		2025:((1,2,3,13),(11,),(20,),(29,),(3,4,5,6),(),(21,),(11,),(15,23),(13,),(3,23,24),(23,28,29,30,31)),\
		2026:((1,2,3,12),(11,),(20,),(29,),(3,4,5,6),(),(20,),(11,),(21,22,23),(12,),(3,23),(23,28,29,30,31)),\
		2027:((1,2,3,11),(11,),(21,22),(29,),(3,4,5),(),(19,),(11,),(20,23),(11,),(3,23),(23,28,29,30,31)),\
		2028:((1,2,3,10),(11,),(20,),(29,),(3,4,5),(),(17,),(11,),(18,22),(9,),(3,23),(23,28,29,30,31)),\
		2029:((1,2,3,8),(11,12),(20,),(29,30),(3,4,5),(),(16,),(11,),(17,23,24),(8,),(3,23),(23,24,28,29,30,31)),\
		2030:((1,2,3,14),(11,),(20,),(29,),(3,4,5,6),(),(15,),(11,12),(16,23),(14,),(3,4,23),(23,28,29,30,31))}  # 祝日JSON。HOLIDAYS[年][月-1]で祝日の日のタプルが返る。日曜日の祝日も含まれる。
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
