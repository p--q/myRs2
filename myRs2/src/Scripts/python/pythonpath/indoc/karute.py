#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# カルテシートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import commons
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.awt import MouseButton  # MessageBoxButtons, MessageBoxResults # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
class Karute():  # シート固有の定数設定。
	pass
def getSectionName(controller, sheet, target):  # 区画名を取得。
	"""
	A  ||  B
	===========  # 行の固定の境界。||は列の固定の境界。境界の行と列はそれぞれ下、右に含む。
	C  ||  D
	-----------  # Date列の文字列があるセルの背景色が青3の行。
	E  ||  F
	-----------  # Date列の文字列があるセルの背景色がスカイブルーの行。
	G  ||  H
	-----------  # Date列の文字列があるセルの背景色が赤3の行。
	I  ||  J
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
	rangeaddress = target.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	if len(sheet[:startrow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "A"
	elif len(sheet[:startrow, startcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "B"
	elif len(sheet[:bluerow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "C"
	elif len(sheet[:bluerow, startcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "D"
	elif len(sheet[:skybluerow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "E"
	elif len(sheet[:skybluerow, startcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "F"	
	elif len(sheet[:redrow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "G"
	elif len(sheet[:redrow, startcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "H"	
	elif len(sheet[redrow:, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "I"  
	else:
		sectionname = "J" 
	karute.sectionname = sectionname   # 区画名	
	karute.startrow = startrow  # スクロール枠の開始行インデックス。
	karute.startcolumn = startcolumn  # スクロール枠の開始列インデックス。
	karute.bluerow = bluerow  # 青3行インデックス。
	karute.skybluerow = skybluerow  # スカイブルー行インデックス。
	karute.redrow = redrow  # 赤3行インデックス。
	return karute  
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	sheet["C1"].setString("一覧へ")
	sheet["E1"].setString("経過へ")
	sheet["I1"].setString("COPY")
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	target = enhancedmouseevent.Target  # ターゲットのセルを取得。
	sheet = target.getSpreadsheet()
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				drowBorders(controller, sheet, target, commons.createBorders())  # 枠線の作成。
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				karute = getSectionName(controller, sheet, target)  # セル固有の定数を取得。
				sectionname = karute.sectionname  # クリックしたセルの区画名を取得。
				if sectionname=="A":
					txt = target.getString()  # クリックしたセルの文字列を取得。	
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
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	pass
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):		
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
	del contextmenu[:]  # contextmenu.clear()は不可。
	target = controller.getSelection()  # 現在選択しているセル範囲を取得。
	if contextmenuname=="cell":  # セルのとき
		karute = getSectionName(controller, sheet, target)  # セル固有の定数を取得。
		sectionname = karute.sectionname  # クリックしたセルの区画名を取得。			
		if sectionname in ("A", "B"):  # 固定行より上の時はコンテクストメニューを表示しない。
			return EXECUTE_MODIFIED
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Delete"})	
# 		if target.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
# 			addMenuentry("ActionTrigger", {"Text": "To Green", "CommandURL": baseurl.format("entry1")})  # listeners.pyの関数名を指定する。
# 		elif target.supportsService("com.sun.star.sheet.SheetCellRange"):  # 連続した複数セルの時。
# 			addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format("entry2")})  # listeners.pyの関数名を指定する。				
	elif contextmenuname=="rowheader":  # 行ヘッダーのとき。
		karute = getSectionName(controller, sheet, target[0, 0])  # 選択範囲の最初のセルの定数を取得。
		sectionname = karute.sectionname  # クリックしたセルの区画名を取得。			
		if sectionname in ("A",) or target[0, 0].getPropertyValue("CellBackColor")!=-1:  # 背景色のあるときは表示しない。
			return EXECUTE_MODIFIED
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsBefore"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsAfter"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteRows"}) 
		if sectionname in ("C",):
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			addMenuentry("ActionTrigger", {"Text": "最下行へ", "CommandURL": baseurl.format("entry1")})  # listeners.pyの関数名を指定する。
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			addMenuentry("ActionTrigger", {"Text": "過去ﾘｽﾄへ移動", "CommandURL": baseurl.format("entry2")})  # listeners.pyの関数名を指定する。
			addMenuentry("ActionTrigger", {"Text": "過去ﾘｽﾄにｺﾋﾟｰ", "CommandURL": baseurl.format("entry3")})  # listeners.pyの関数名を指定する。
		elif sectionname in ("G",):
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			addMenuentry("ActionTrigger", {"Text": "現ﾘｽﾄへ移動", "CommandURL": baseurl.format("entry4")})  # listeners.pyの関数名を指定する。
			addMenuentry("ActionTrigger", {"Text": "現ﾘｽﾄにｺﾋﾟｰ", "CommandURL": baseurl.format("entry5")})  # listeners.pyの関数名を指定する。		
	elif contextmenuname=="colheader":  # 列ヘッダーの時。
		pass
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。	
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	# 行ヘッダー
	if entrynum==1:  # 最下行へ
		pass

	elif entrynum==2:  # 過去ﾘｽﾄへ移動
		pass
	
	elif entrynum==3:  # 過去ﾘｽﾄにｺﾋﾟｰ
		pass

	elif entrynum==4:  # 現ﾘｽﾄへ移動
		pass

	elif entrynum==5:  # 現ﾘｽﾄにｺﾋﾟｰ
		pass
def drowBorders(controller, sheet, cellrange, borders):  # ターゲットを交点とする行列全体の外枠線を描く。
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = borders  # 枠線を取得。	
	cell = cellrange[0, 0]  # セル範囲の左上端のセルで判断する。
	rangeaddress = cell.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	karute = getSectionName(controller, sheet, cell)  # セル固有の定数を取得。
	sectionname = karute.sectionname  # クリックしたセルの区画名を取得。
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
	if sectionname in ("A", "B", "E", "I"):  # 枠線を消すだけ。
		return
	if sectionname in ("C", "G"):  # 同一プロブレムの上下に枠線を引く。
		datarange = sheet[karute.startrow:karute.bluerow, 1:7] if sectionname=="C" else sheet[karute.skybluerow+1:karute.redrow, 1:7]  # タイトル行を除く。
		doc = controller.getModel()  # ドキュメントモデルを取得。
		rstartrow = 0 # プロブレムの開始行の相対インデックス。
		datarows = datarange.getDataArray()  # #列からSubject列までの行のタプルを取得。
		ranges = []  # プロブレムリストのセル範囲のリスト。
		for i, datarow in enumerate(datarows):  # 相対インデックスと行のタプルを列挙。
			if "#" in "{}{}{}{}".format(*datarow[:4]):  # #列からSubject列まで結合して#がある時。。日付は数値なので文字列への変換が必要なのでjoin()は使えない。
				if i>rstartrow:  # 開始行インデックスより大きい時。
					ranges.append(datarange[rstartrow:i, :])
					rstartrow = i
		if ranges:  # すでにプロブレムがあるときのみ。一つも取得できていないときは一つもプロブレムがないので取得しない。
			ranges.append(datarange[rstartrow:, :])  # 最後のプロブレムのセル範囲を追加。
			cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
			cellranges.addRangeAddresses([i.getRangeAddress() for i in ranges], False)  # セル範囲コレクションにプロブレムのセル範囲を追加する。セル範囲は結合しない。
			for i in cellranges:  # 各セル範囲について。
				if len(i.queryIntersection(rangeaddress)):  # 選択したセルが含むセル範囲の時。
					cursor = sheet.createCursorByRange(i)  # セルカーサーを作成
					cursor.expandToEntireRows()  # セル範囲を行全体に拡大。
					cursor.setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。
					break  # 1つのプロブレムしか枠線を引かない。
	elif sectionname in ("D", "F", "H"):
		sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。			
		sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。	
		cellrange.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	