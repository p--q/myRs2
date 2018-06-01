#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
from indoc import commons
# from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.sheet import CellFlags  # 定数


class Ent():  # シート固有の定数設定。
	def __init__(self, sheet):
		self.menurow  = 0  # メニュー行インデックス。
		self.splittedrow = 1  # 分割行インデックス。
		self.idcolumn = 0  # ID列インデックス。	
		self.kanjicolumn = 1  # 漢字列インデックス。
		self.kanacolumn = 2  # カナ列インデックス。	
		self.datecolumn = 3  # 入院日列インデックス。
		self.cleardatecolumn = 4  # リスト消去日列インデックス。
		cellranges = sheet[:, self.idcolumn].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # ID列の文字列が入っているセルに限定して抽出。数値の時もありうる。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # ID列の最終行インデックス+1を取得。








# def notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, contextmenuname):			
# 	if contextmenuname=="cell":  # セルのとき
# 		selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
# 		del contextmenu[:]  # contextmenu.clear()は不可。
# 		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
# 			addMenuentry("ActionTrigger", {"Text": "To blue", "CommandURL": baseurl.format("entry1")})  # listeners.pyの関数名を指定する。
# 		elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 連続した複数セルの時。
# 			addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format("entry2")})  # listeners.pyの関数名を指定する。
# 		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
# 	elif contextmenuname=="rowheader":  # 行ヘッダーのとき。
# 		del contextmenu[:]  # contextmenu.clear()は不可。
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
# 		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsBefore"})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteRows"}) 
# 	elif contextmenuname=="colheader":  # 列ヘッダーの時。
# 		pass  # contextmenuを操作しないとすべての項目が表示されない。
# 	elif contextmenuname=="sheettab":  # シートタブの時。
# 		del contextmenu[:]  # contextmenu.clear()は不可。
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
# def contextMenuEntries(target, entrynum):  # コンテクストメニュー番号の処理を振り分ける。
# 	colors = commons.COLORS
# 	if entrynum==1:
# 		target.setPropertyValue("CellBackColor", colors["ao"])  # 背景を青色にする。
# 	elif entrynum==2:
# 		target.setPropertyValue("CellBackColor", colors["aka"]) 
