#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.datatransfer import XTransferable
from com.sun.star.datatransfer import DataFlavor  # Struct
from com.sun.star.datatransfer import UnsupportedFlavorException
from com.sun.star.lang import Locale  # Struct
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
