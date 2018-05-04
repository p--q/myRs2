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
		sectionname = "I"  # 記事行がないときはrngアトリビュートはない。
		cellranges = sheet[redrow+1:, 6].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # 記事列の文字列が入っているセルに限定して抽出。空列は不可。数値の時もありうる。
		if len(cellranges):  # 区域Iにデータ行がある時のみ。
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
	sheet["C1"].setString("一覧へ")
	sheet["E1"].setString("経過へ")
	sheet["I1"].setString("COPY")
def mousePressed(enhancedmouseevent, controller, sheet, target, args):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	borders, systemclipboard, transliteration = args
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				drowBorders(controller, sheet, target, borders)
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				karute = getSectionName(controller, sheet, target)  # セル固有の定数を取得。
				sectionname = karute.sectionname  # クリックしたセルの区画名を取得。
				if sectionname=="A":
					txt = target.getString()  # クリックしたセルの文字列を取得。	
					doc = controller.getModel()
					sheets = doc.getSheets()  # シートコレクションを取得。
					if txt=="一覧へ":
						controller.setActiveSheet(sheets["一覧"])  # 一覧シートをアクティブにする。
					elif txt=="経過へ":
						newsheetname = "".join([sheet.getName(), "経"])  # 経過シート名を取得。
						if newsheetname in sheets:  # 経過シート名がある時。
							controller.setActiveSheet(sheets[newsheetname])  # 経過シートをアクティブにする。
						else:
							pass
								
				
				
				
				
# 				ichiran = getSectionName(controller, sheet, target)
# 				section, startrow, emptyrow, sumi_retu, dstart = ichiran.sectionname, ichiran.startrow, ichiran.emptyrow, ichiran.sumi_retu, ichiran.dstart
# 				celladdress = target.getCellAddress()
# 				r, c = celladdress.Row, celladdress.Column  # targetの行と列のインデックスを取得。		
# 				
				
				
				
				
				
	return True  # セル編集モードにする。
def drowBorders(controller, sheet, cellrange, borders):  # ターゲットを交点とする行列全体の外枠線を描く。
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = borders  # 枠線を取得。	
	cell = cellrange[0, 0]  # セル範囲の左上端のセルで判断する。
	rangeaddress = cell.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	karute = getSectionName(controller, sheet, cell)  # セル固有の定数を取得。
	sectionname = karute.sectionname  # クリックしたセルの区画名を取得。
	if sectionname in ("C", "G"):  # 同一プロブレムの上下に枠線を引く。
		if cell.getPropertyValue("CellBackColor") in (-1, ):  # 背景色がない時。
			sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
			doc = controller.getModel()  # ドキュメントモデルを取得。
			cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
			datarange = karute.rng  # 選択セルの区画のセル範囲を取得。但し背景色のある行を除く。
			datarows = datarange[:, 1:5].getDataArray()  # #列からSubject列までの行のタプルを取得。
			startrow = 0  # プロブレムの開始行の相対インデックス。
			for i, datarow in enumerate(datarows):  # 相対インデックスと行のタプルを列挙。
				if "#" in "{}{}{}{}".format(*datarow):  # #列からSubject列まで結合して#がある時。。日付は数値なので文字列への変換が必要なのでjoin()は使えない。
					if i>startrow:  # 開始行インデックスより大きい時。
						cellranges.addRangeAddress(datarange[startrow:i, :].getRangeAddress(), False)  # セル範囲コレクションにプロブレムのセル範囲を追加する。セル範囲は結合しない。
						startrow = i
			cellranges.addRangeAddress(datarange[startrow:i+1, :].getRangeAddress(), False)  # 最後のプロブレムのセル範囲を追加。
			for i in cellranges:  # 各セル範囲について。
				if len(i.queryIntersection(rangeaddress)):  # 選択したセルが含むセル範囲の時。
					cursor = sheet.createCursorByRange(i)  # セルカーサーを作成
					cursor.expandToEntireRows()  # セル範囲を行全体に拡大。
					cursor.setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。
					break  # 1つのプロブレムにしか対応しない。
	elif sectionname in ("D", "F", "H"):
		if cell.getPropertyValue("CellBackColor") in (-1, ):  # 背景色がない時。
			sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
			sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。			
			sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。	
			cellrange.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	
	elif sectionname in ("I", "E"):
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
