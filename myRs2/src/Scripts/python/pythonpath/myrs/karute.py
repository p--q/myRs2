#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# カルテシートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from myrs import commons
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.awt import MouseButton  # MessageBoxButtons, MessageBoxResults # 定数
class Karute():  # シート固有の定数設定。
	pass
def getSectionName(controller, sheet, cell):  # 区画名を取得。
	"""
	A  ||  B
	===========  # 行の固定の境界。||は列の固定の境界。境界の行と列はそれぞれ下、右に含む。
	C  ||  D
	-----------  # Date列の文字列があるセルの背景色が青3の行。
	E  ||  F
	-----------  # Date列の文字列があるセルの背景色がスカイブルーの行。
	G  ||  H
	-----------  # Date列の文字列があるセルの背景色が赤3の行。
	I
	"""
	datecolumn = 2  # Date列インデックス。
	subcontollerrange = controller[0].getVisibleRange()
	startrow = subcontollerrange.EndRow + 1  # スクロールする枠の最初の行インデックス。
	startcolumn = subcontollerrange.EndColumn + 1  # スクロールする枠の最初の列インデックス。
	cellranges = sheet[startrow:, datecolumn].queryContentCells(CellFlags.STRING)  # Date列の文字列が入っているセルに限定して抽出。
	backcolors = commons.COLORS["blue3"], commons.COLORS["skyblue"], commons.COLORS["red3"]  # ジェネレーターに使うので順番が重要。
	gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in backcolors)
	bluerow = next(gene)
	skybluerow = next(gene)
	redrow = next(gene)
	karute = Karute()  # クラスをインスタンス化。	
	rangeaddress = cell.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	if len(sheet[:startrow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "A"
	elif len(sheet[:startrow, startcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "B"
	elif len(sheet[:bluerow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "C"
		karute.rng = sheet[startrow:bluerow, :startcolumn]
	elif len(sheet[:bluerow, startcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "D"
	elif len(sheet[:skybluerow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "E"
		karute.rng = sheet[bluerow+1:skybluerow, :startcolumn]
	elif len(sheet[:skybluerow, startcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "F"	
	elif len(sheet[:redrow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "G"
		karute.rng = sheet[skybluerow+1:redrow, :startcolumn]
	elif len(sheet[:redrow, startcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "H"	
	else:
		sectionname = "I"	
		cellranges = sheet[:, 6].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # 記事列の文字列が入っているセルに限定して抽出。空列は不可。数値の時もありうる。
		emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 記事列の最終行インデックス+1を取得。
		karute.rng = sheet[redrow+1:emptyrow, :startcolumn]
	karute.sectionname = sectionname   # 区画名
	return karute  
def selectionChanged(controller, sheet, args):  # 矢印キーでセル移動した時も発火する。
	pass
# 	borders = args	
# 	selection = controller.getSelection()
# 	if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択範囲がセルの時。矢印キーでセルを移動した時。マウスクリックハンドラから呼ばれると何回も発火するのでその対応。
# 		currenttableborder2 = selection.getPropertyValue("TableBorder2")  # 選択セルの枠線を取得。
# 		if all((currenttableborder2.TopLine.Color==currenttableborder2.LeftLine.Color==commons.COLORS["violet"],\
# 				currenttableborder2.RightLine.Color==currenttableborder2.BottomLine.Color==commons.COLORS["magenta3"])):  # 枠線の色を確認。
# 			return  # すでに枠線が書いてあったら何もしない。
# 	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
# 		drowBorders(controller, sheet, selection, borders)	
def activeSpreadsheetChanged(sheet):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet["C1"].setString("リストへ")
	sheet["E1"].setString("経過ｼｰﾄへ")
	sheet["I1"].setString("COPY")
def mousePressed(enhancedmouseevent, controller, sheet, target, args):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	borders, systemclipboard, transliteration = args
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				drowBorders(controller, sheet, target, borders)
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				ichiran = getSectionName(controller, sheet, target)
				section, startrow, emptyrow, sumi_retu, dstart = ichiran.sectionname, ichiran.startrow, ichiran.emptyrow, ichiran.sumi_retu, ichiran.dstart
				celladdress = target.getCellAddress()
				r, c = celladdress.Row, celladdress.Column  # targetの行と列のインデックスを取得。		
				txt = target.getString()  # クリックしたセルの文字列を取得。		
				
				
				
				
				
	return True  # セル編集モードにする。
def drowBorders(controller, sheet, cellrange, borders):  # ターゲットを交点とする行列全体の外枠線を描く。
	cell = cellrange[0, 0]  # セル範囲の左上端のセルで判断する。
	karute = getSectionName(controller, sheet, cell)
	sectionname = karute.sectionname
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = borders	
	if sectionname in ("C", "G"):
		sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
		
		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		
		
		
		
		cellranges = karute.rng[:, 1:-3].queryContentCells(CellFlags.STRING)  # #列からSubjct列について文字列の入っているセル範囲コレクションを取得。
		upper = None  # 枠線を上に引くセル。
		p = None  # forループの一つ前のセル。
		for i in cellranges.getCells():
			if i.getString().startswith("#"):  # '#'から始まるセルの時。
				if upper: # 枠線を上にセルがすでに取得出来ている時。
					bottom = i  # 枠線を下に引くセルの一つ下の行のセルとして取得。
					break
				else:
					tmp = i
			elif i==cell:
				upper = tmp	
			p = i
		else:
			bottom = i
		sheet[upper.getCellAddress().Row:bottom.getCellAddress().Row+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。	
	elif sectionname in ("D", "F", "H"):
		if cell.getPropertyValue("CellBackColor") in (-1, ):  # 背景色がない時。
			rangeaddress = cellrange.getRangeAddress()  # セル範囲アドレスを取得。
			sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。			
			sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。	
			cellrange.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	
	elif sectionname=="I":
		sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
def notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, contextmenuname):			
	if contextmenuname=="cell":  # セルのとき
		selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
		del contextmenu[:]  # contextmenu.clear()は不可。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
			addMenuentry("ActionTrigger", {"Text": "To Green", "CommandURL": baseurl.format("entry1")})  # listeners.pyの関数名を指定する。
		elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 連続した複数セルの時。
			addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format("entry2")})  # listeners.pyの関数名を指定する。
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
	elif contextmenuname=="rowheader":  # 行ヘッダーのとき。
		del contextmenu[:]  # contextmenu.clear()は不可。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsBefore"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteRows"}) 
	elif contextmenuname=="colheader":  # 列ヘッダーの時。
		pass  # contextmenuを操作しないとすべての項目が表示されない。
	elif contextmenuname=="sheettab":  # シートタブの時。
		del contextmenu[:]  # contextmenu.clear()は不可。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
def contextMenuEntries(target, entrynum):  # コンテクストメニュー番号の処理を振り分ける。
	colors = commons.COLORS
	if entrynum==1:
		target.setPropertyValue("CellBackColor", colors["midori"])  # 背景を青色にする。
	elif entrynum==2:
		target.setPropertyValue("CellBackColor", colors["aka"]) 
