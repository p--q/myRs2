#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 経過シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import calendar
from itertools import chain
from indoc import commons, staticdialog
from com.sun.star.awt import MouseButton, MessageBoxButtons, Key  # 定数
from com.sun.star.awt import KeyEvent  # Struct
from com.sun.star.awt.MessageBoxType import ERRORBOX  # enum
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.sheet.CellInsertMode import ROWS as insert_rows  # enum
from com.sun.star.sheet.CellDeleteMode import ROWS as delete_rows  # enum
from com.sun.star.table.CellHoriJustify import CENTER, LEFT  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
class Keika():  # シート固有の定数設定。
	def __init__(self):
		self.daterow = 1  # 日付行インデックス。
		self.splittedrow = 4  # 分割行インデックス。
		self.yakucolumn = 5  # 薬名列インデックス。
		self.splittedcolumn = 9  # 分割列インcccfewfweデックス。
	def setSheet(self, sheet):
		self.sheet = sheet
		cellranges = sheet[:, self.yakucolumn].queryContentCells(CellFlags.STRING)  # 薬名列の文字列が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 薬名列の最終行インデックス+1を取得。
		gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in (commons.COLORS["black"],))
		self.blackrow = next(gene)  # 黒行インデックスを取得。	
VARS = Keika()		
# def getConsts(sheet, selection=None):  # 区画名を取得。
# 	"""
# 	A  ||  B
# 	===========  # 行の固定の境界。||は列の固定の境界。境界の行と列はそれぞれ下、右に含む。
# 	C  ||  D
# 	I-----------  # 黒行。この行は含まない。
# 	E  ||  F
# 	-----------  # 薬品列の最下行の一つ下の行。
# 	G  ||  H
# 	
# 	# J: C,E,Gの薬名列より左。
# 	
# 	"""
# 	consts = Keika(sheet)
# 	if selection is not None:
# 		splittedrow = consts.splittedrow
# 		splittedcolumn = consts.splittedcolumn
# 		blackrow = consts.blackrow
# 		emptyrow = consts.emptyrow
# 		rangeaddress = selection[0, 0].getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
# 		sectionname = ""
# 		if splittedrow<blackrow:
# 			if len(sheet[splittedrow:blackrow, :splittedcolumn].queryIntersection(rangeaddress)): 
# 				sectionname = "C"			
# 			elif len(sheet[splittedrow:blackrow, splittedcolumn:].queryIntersection(rangeaddress)): 
# 				sectionname = "D"			
# 		elif blackrow+1<emptyrow:
# 			if len(sheet[blackrow+1:emptyrow, :splittedcolumn].queryIntersection(rangeaddress)): 
# 				sectionname = "E"				
# 			elif len(sheet[blackrow+1:emptyrow, splittedcolumn:].queryIntersection(rangeaddress)): 
# 				sectionname = "F"	
# 		if not sectionname:		
# 			if len(sheet[:splittedrow, :splittedcolumn].queryIntersection(rangeaddress)): 
# 				sectionname = "A"	
# 			elif len(sheet[:splittedrow, splittedcolumn:].queryIntersection(rangeaddress)): 
# 				sectionname = "B"					
# 			elif len(sheet[emptyrow:, :splittedcolumn].queryIntersection(rangeaddress)): 
# 				sectionname = "G"					
# 			elif len(sheet[emptyrow:, splittedcolumn:].queryIntersection(rangeaddress)): 
# 				sectionname = "H"
# 			else:
# 				sectionname = "I"
# 		if sectionname in ("C", "E", "G") and len(sheet[splittedrow:, :consts.yakucolumn].queryIntersection(rangeaddress)):
# 			sectionname = "J"
# 		consts.sectionname = sectionname  # 区画名
# 	return consts  
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	sheet["F1:G1"].setDataArray((("一覧へ", "ｶﾙﾃへ"),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet["F3:F4"].setDataArray((("薬品整理",), ("薬品名抽出",)))
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。			
	daterow = VARS.daterow
	splittedcolumn = VARS.splittedcolumn
	startdatevalue = int(sheet[daterow, splittedcolumn].getValue())  # 日付行の最初のセルから日付のシリアル値の取得。
	todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。
	sheet[daterow-1, splittedcolumn:].setPropertyValue("CellBackColor", -1)  # r-1行目の背景色をクリア。
	c = splittedcolumn + (todayvalue - startdatevalue)  # 今日の日付の列インデックスを取得。
	if c<1024:
		sheet[daterow-1, c].setPropertyValue("CellBackColor", commons.COLORS["violet"])  # 日付行の上のセルの今日の背景色を設定。
	sheet[daterow+2:, splittedcolumn:].setPropertyValue("HoriJustify", LEFT)  # 分割列以降、日付行2行下以降すべて左詰めにする。
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。		
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			VARS.setSheet(selection.getSpreadsheet())
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				drowBorders(selection)  # 枠線の作成。
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				celladdress = selection.getCellAddress()
				r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。	
				if r<VARS.splittedrow:  # 分割行より上、の時。
					if c<VARS.splittedcolumn:  # 分割列より左、の時。
						return wClickMenu(enhancedmouseevent, xscriptcontext)
					else: 
						return wClickUpperRight(enhancedmouseevent, xscriptcontext)
				elif r!=VARS.blackrow:  # 黒行でない時。
					if r>VARS.splittedcolumn-1:  # 分割行含む右列。
						return wClickBottomRight(enhancedmouseevent, xscriptcontext)
					elif r==VARS.yakucolumn:  # 薬名列の時。
						return True  # セル編集モードにする。
					else:	
						return wClickBottomLeft(enhancedmouseevent, xscriptcontext)
				return False  # セル編集モードにしない。
def wClickMenu(enhancedmouseevent, xscriptcontext):
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	sheets = doc.getSheets()  # シートコレクションを取得。	
	sheet = VARS.sheet	
	controller = doc.getCurrentController()  # コントローラの取得。	
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	txt = selection.getString()	
	if txt=="一覧へ":
		controller.setActiveSheet(sheets["一覧"])  # 一覧シートをアクティブにする。
	elif txt=="ｶﾙﾃへ":  # カルテシートをアクティブにする、なければ作成する。
		datarow = sheet[1, VARS.yakucolumn:VARS.splittedcolumn+1].getDataArray()[0]  # IDセルから最初の日付セルまで取得。
		idcelltxts = datarow[0].split(" ")  # 半角スペースで分割。
		idtxt = idcelltxts[0]  # 最初の要素を取得。
		if idtxt.isdigit():  # IDが数値のみの時。					
			if idtxt in sheets:  # ID名のシートがあるとき。
				controller.setActiveSheet(sheets[idtxt])  # カルテシートをアクティブにする。
			else:
				if len(idcelltxts)==5:  # ID、漢字姓・名、カタカナ姓・名、の5つに分割できていた時。
					kanjitxt, kanatxt = " ".join(idcelltxts[1:3]), " ".join(idcelltxts[3:])
					datevalue = datarow[-1]
					karutesheet = commons.getKaruteSheet(commons.formatkeyCreator(doc), sheets, idtxt, kanjitxt, kanatxt, datevalue)
					controller.setActiveSheet(karutesheet)  # カルテシートをアクティブにする。
				else:
					commons.showErrorMessageBox(controller, "「ID(数値のみ) 漢字姓 名 カナ姓 名」の形式になっていません。")
		else:
			commons.showErrorMessageBox(controller, "IDが取得できませんでした。")	
	elif txt=="薬品整理":  # クリックするたびに終了順、昇順に並び替える。黒行の上のみ。
		if VARS.splittedrow>VARS.blackrow:  # 分割行から黒行より上に行がある時のみ。
			datarange = sheet[VARS.splittedrow:VARS.blackrow, :]  # 黒行より上の行のセル範囲を取得。
			controller.select(datarange)  # ソートするセル範囲を取得。
			if selection.getPropertyValue("CellBackColor")==-1:  # ボタンの背景色がない時、薬名列の昇順でソート。
				selection.setPropertyValue("CellBackColor", commons.COLORS["lime"])  # ボタンの背景色を付ける。				
				props = PropertyValue(Name="Col1", Value=VARS.yakucolumn+1),  # Col1の番号は優先順位。Valueはインデックス+1。 			
			else:  # ボタンの背景色がある時、終了順でソート。終了列インデックスを先頭列に代入しておく。
				datarows = []  # 終了行インデックスを入れる行のリスト。
				for i in range(VARS.blackrow-VARS.splittedrow):  # 分割行インデックスから、黒行の上までの相対インデックスを取得。
					cellranges = datarange[i, VARS.splittedcolumn:].queryContentCells(CellFlags.STRING)  # 文字列のあるセル範囲コレクションを取得。
					if len(cellranges):  # セル範囲が取得出来た時。
						datarows.append((cellranges.getRangeAddresses()[-1].EndColumn,))  # 最終列インデックスを取得。
					else:
						datarows.append((1,))  # 色セルがない行は1にして上に持ってくる。0にするとFalseになってしまう。
				datarange[:, 0].setDataArray(datarows)  # 開始列インデックスをシートに代入。
				datarange[:, 0].setPropertyValue("CharColor", commons.COLORS["white"])  # 先頭列の文字色を白色にする。
				selection.setPropertyValue("CellBackColor", -1)  # ボタンの背景色を消す。		
				props = PropertyValue(Name="Col1", Value=1),  # Col1の番号は優先順位。Valueはインデックス+1。 
			dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
			dispatcher.executeDispatch(controller.getFrame(), ".uno:DataSort", "", 0, props)  # ディスパッチコマンドでソート。sort()メソッドは挙動がおかしくて使えない。								
			controller.select(selection)  # ボタンを選択し直す。	
	elif txt=="薬品名抽出":
		firstrow = max(sheet[:, i].queryContentCells(CellFlags.STRING).getRangeAddresses()[-1].EndRow for i in (VARS.yakucolumn+1, VARS.yakucolumn+2)) + 1  # 用法列か回数列の最終行インデックスの下の行インデックスを取得。
		if firstrow<VARS.emptyrow:
			transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
			transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))  # 全角文字を半角にする。
			datarows = sheet[firstrow:VARS.emptyrow, VARS.yakucolumn].getDataArray()  # 用法設定していない薬品列の各行のタプルを取得。
			sep = "*sep*"  # 区切り文字。
			concatenetedtxt = sep.join(chain.from_iterable(datarows))  # 区切り文字で全行を結合。
			transliteration.transliterate(concatenetedtxt, 0, len(concatenetedtxt), [])[0]  # 半角に変換。
			rowtxts = concatenetedtxt.split(sep)  # 区切り文字で分割。
			rowtxtslength = len(rowtxts)
			newdatarows = []
			for i, rowtxt in enumerate(rowtxts):  # 行の相対インデックスとともにイテレートする。
				if rowtxt.endswith(("錠", "袋", "g", "本", "瓶", "管", "包", "枚", "個", "ｶﾌﾟｾﾙ", "ｷｯﾄ")):  # 特定の文字列で終わっている時は追加する。
					if rowtxt in ("ﾍﾟﾝﾆｰﾄﾞﾙ", "ﾋﾞﾀﾒｼﾞﾝ", "ﾌﾞﾄﾞｳ糖注50%PL", "生理食塩水PL", "CV主管", "CV副管"):  # 特定の文字列が含まれている時は追加しない。
						continue									
					for j in range(i+1, i+4):  # 3行下の行まで。
						if j<rowtxtslength:  # j行が存在する時。
							if "1日間" in rowtxts[j]:  # j行に"1日間"がある時。
								if j+1<rowtxtslength:  # j+1行が存在する時。
									if not "日間" in rowtxts[j+1]:  # j+1行に"日間"がない時。
										break  
								else:  # "1日間"で終わっている時。
									break	
						else:
							break
					else:  # breakされなかった時。
						if not rowtxt in newdatarows:  # まだ追加していない要素の時のみ。
							newdatarows.append((rowtxt,))  # その行を取得。
			sheets[firstrow:VARS.emptyrow, VARS.yakucolumn:VARS.splittedcolumn].clearContents(CellFlags.STRING+CellFlags.VALUE)  # 整理前のセルの文字列と数値をクリア。		
			sheets[firstrow:firstrow+len(newdatarows), VARS.yakucolumn].setDataArray(newdatarows)  # 整理した薬品名をシートに代入。		
	elif txt[:8].isdigit():  # 最初8文字が数値の時。						
		systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
		systemclipboard.setContents(commons.TextTransferable(txt[:8]), None)  # クリップボードにIDをコピーする。							
	return False  # セル編集モードにしない。	
def wClickUpperRight(enhancedmouseevent, xscriptcontext):
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	celladdress = selection.getCellAddress()
	r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。		
	items = []
	if r==VARS.daterow-1:  # 行インデックス0の時。月を入力。
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
		functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。							
		datevalue = int(VARS.sheet[VARS.daterow, c].getValue())
		m = int(functionaccess.callFunction("MONTH", (datevalue,)))  # 月、を取得。
		selection.setString("{}月".format(m))
	elif r==VARS.daterow+1:  # 日付行の下の時。
		staticdialog.createDialog(enhancedmouseevent, xscriptcontext, VARS.sheet[r, VARS.yakucolumn+1].getString())
	

		
		
		
		
	elif r==2:	
		items = ["", "○", "尿"]			
		horijustify	= CENTER
	elif r==3:
		items = ["", "胸Xp", "腹ｴ", "心ｴ"]
		horijustify	= LEFT
	if items:
		if txt in items:  # セルの内容にある時。
			txt = txtCycle(items, txt)	
			selection.setString(txt)
	if txt:  # 文字がある時。
		selection.setPropertyValues(("CellBackColor", "HoriJustify"), (commons.COLORS["skyblue"], horijustify))  # 背景をスカイブルーにする。		
	else:
		selection.setPropertyValue("CellBackColor", -1)  # 背景色を消す。
	return False  # セル編集モードにしない。
def wClickBottomLeft(enhancedmouseevent, xscriptcontext):
	
	
	pass
def wClickBottomRight(enhancedmouseevent, xscriptcontext):
	
	
	pass





# 				elif sectionname=="B":
# 					items = []
# 					if r==0:  # 月を入力。
# 						ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
# 						smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
# 						functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。							
# 						datevalue = int(sheet[consts.daterow, c].getValue())
# 						m = int(functionaccess.callFunction("MONTH", (datevalue,)))  # 月、を取得。
# 						selection.setString(txtCycle(["", "{}月".format(m)], txt))
# 						return False  # セル編集モードにしない。
# 					elif r==1:  # 日付行の時。
# 						return False  # セル編集モードにしない。
# 					elif r==2:	
# 						items = ["", "○", "尿"]			
# 						horijustify	= CENTER
# 					elif r==3:
# 						items = ["", "胸Xp", "腹ｴ", "心ｴ"]
# 						horijustify	= LEFT
# 					if items:
# 						if txt in items:  # セルの内容にある時。
# 							txt = txtCycle(items, txt)	
# 							selection.setString(txt)
# 					if txt:  # 文字がある時。
# 						selection.setPropertyValues(("CellBackColor", "HoriJustify"), (commons.COLORS["skyblue"], horijustify))  # 背景をスカイブルーにする。		
# 					else:
# 						selection.setPropertyValue("CellBackColor", -1)  # 背景色を消す。
# 					return False  # セル編集モードにしない。	
# 				elif sectionname in ("C", "E", "G"):
# 					if c==yakucolumn:  # 薬名列の時。
# 						return True  # セル編集モードにする。
# 					elif c==yakucolumn+1:  # 用法列の時。
# 						if txt:
# # 							dialogs.createDialog(xscriptcontext)
# 
# 							pass
# 						else:
# 							selection.setString("分3")
# 					elif c==yakucolumn+2:  # 回数列の時。
# 						if txt:
# 							
# 							
# 							pass
# 						else:
# 							selection.setString("持続")			
# 					return False  # セル編集モードにしない。	
# 				elif sectionname in ("D", "F"):
# 					if txt:
# 						
# 						
# 						pass
# 					else:
# 						selection.setString("止")								
# 					return False  # セル編集モードにしない。	
# 				elif sectionname in ("J",):
# 					header = sheet[1, c].getString()  # 行インデックス1のセルの文字列を取得。
# 					controller.select(sheet[r, yakucolumn])  # 薬名列のセルを選択。
# 					historydialogyaku.createDialog(xscriptcontext, enhancedmouseevent, header)  # 履歴ダイアログを表示。クリックした位置の下に表示。入力するとシートを下にスクロールする。		
# 					return False  # セル編集モードにしない。	
# 	return True  # セル編集モードにする。
# def txtCycle(items, txt):  # items: 循環させる文字列のリスト、txt: 現在の文字列。
# 	items.append(items[0])  # 最初の要素を最後の要素に追加する。
# 	dic = {items[i]: items[i+1] for i in range(len(items)-1)}  # 順繰り辞書の作成。
# 	return dic[txt]
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	selection = eventobject.Source.getSelection()
	if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択範囲がセルの時。矢印キーでセルを移動した時。マウスクリックハンドラから呼ばれると何回も発火するのでその対応。
		currenttableborder2 = selection.getPropertyValue("TableBorder2")  # 選択セルの枠線を取得。
		if all((currenttableborder2.TopLine.Color==currenttableborder2.LeftLine.Color==commons.COLORS["violet"],\
				currenttableborder2.RightLine.Color==currenttableborder2.BottomLine.Color==commons.COLORS["magenta3"])):  # 枠線の色を確認。
			return  # すでに枠線が書いてあったら何もしない。
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
		drowBorders(selection)  # 枠線の作成。
def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。		
	changes = changesevent.Changes	
	for change in changes:
		if change.Accessor=="cell-change":  # セルの値が変化した時。
			selection = change.ReplacedElement  # 値を変更したセルを取得。	
			celladdress = selection.getCellAddress()
			r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。				
			
		
			sheet = selection.getSpreadsheet()
			
			
			
			
			consts = getConsts(sheet, selection)  # 経過シート固有の定数を取得。
			sectionname = consts.sectionname  # クリックしたセルの区画名を取得。
			if not sectionname in ("A",):  # 領域A以外の時。
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
				transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
				transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))					
				txt = selection.getString()  # セルの文字列を取得。	
				txt = transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換。
				selection.setString(txt)
			break
# def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。				
# 	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
# 	sheet = controller.getActiveSheet()  # アクティブシートを取得。
# 	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
# 	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
# 	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
# 	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
# 	del contextmenu[:]  # contextmenu.clear()は不可。
# 	selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
# 	consts = getConsts(sheet, selection)  # セル固有の定数を取得。
# 	sectionname = consts.sectionname  # クリックしたセルの区画名を取得。		
# 	if contextmenuname=="cell":  # セルのとき		
# 		if sectionname in ("A",):
# 			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
# 				txt = selection.getString()
# 				if txt=="薬品整理":
# 					addMenuentry("ActionTrigger", {"Text": "同薬品抽出", "CommandURL": baseurl.format("entry1")}) 
# 					addMenuentry("ActionTrigger", {"Text": "同薬品結合", "CommandURL": baseurl.format("entry2")}) 
# 			return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
# 		elif sectionname in ("B",):
# 			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
# 				celladdress = selection.getCellAddress()  # 経過シートの日付の開始セルのセルアドレスを取得。
# 				r = celladdress.Row
# 				if r==consts.daterow:  # 日付行の時。
# 					if selection.getValue():  # セルに値があるとき。
# 						addMenuentry("ActionTrigger", {"Text": "日付追加", "CommandURL": baseurl.format("entry5")}) 
# 				elif r==consts.daterow+2:  # 処置行の時。
# 					commons.cutcopypasteMenuEntries(addMenuentry)
# 			return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
# 		elif sectionname in ("D", "F"):
# 			addMenuentry("ActionTrigger", {"Text": "処方", "CommandURL": baseurl.format("entry10")})
# 			
# 			addMenuentry("ActionTrigger", {"Text": "翌火へ", "CommandURL": baseurl.format("entry11")})
# 			addMenuentry("ActionTrigger", {"Text": "翌金へ", "CommandURL": baseurl.format("entry12")})
# 			
# 			addMenuentry("ActionTrigger", {"Text": "翌月まで", "CommandURL": baseurl.format("entry13")})  # 	回数列が空欄の時は金まで、それ以外は火まで。
# 			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
# 			commons.cutcopypasteMenuEntries(addMenuentry)
# 		elif sectionname in ("C", "E", "G", "H"):	
# 			commons.cutcopypasteMenuEntries(addMenuentry)
# 	elif contextmenuname=="rowheader" and len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 行ヘッダーのとき、かつ、選択範囲の列数がシートの列数が一致している時。	
# 		if sectionname in ("I",):
# 			return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
# 		elif sectionname in ("C",):
# 			addMenuentry("ActionTrigger", {"Text": "使用中最上行へ", "CommandURL": baseurl.format("entry15")})  
# 			addMenuentry("ActionTrigger", {"Text": "使用中最下行へ", "CommandURL": baseurl.format("entry16")}) 
# 			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 		elif sectionname in ("E",):	
# 			addMenuentry("ActionTrigger", {"Text": "黒行上へ", "CommandURL": baseurl.format("entry17")})  
# 			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 			addMenuentry("ActionTrigger", {"Text": "使用中最上行へ", "CommandURL": baseurl.format("entry18")})  
# 			addMenuentry("ActionTrigger", {"Text": "使用中最下行へ", "CommandURL": baseurl.format("entry19")}) 
# 			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 		commons.cutcopypasteMenuEntries(addMenuentry)
# 		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 		commons.rowMenuEntries(addMenuentry)
# 	elif contextmenuname=="colheader" and len(selection[:, 0].getColumns())==len(sheet[:, 0].getColumns()):  # 列ヘッダーのとき、かつ、選択範囲の行数がシートの行数が一致している時。	
# 		if sectionname in ("B",):
# 			if len(selection[0, :].getColumns())==1 and selection[0, 0].getCellAddress().Column>consts.splittedcolumn:  # 選択列数が1つだけ、かつ、固定列より右の時。
# 				if selection[consts.blackrow, 0].getPropertyValue("CellBackColor")==commons.COLORS["black"]:  # 選択範囲の黒行のセルの背景色が黒色の時。
# 					addMenuentry("ActionTrigger", {"Text": "退院翌日", "CommandURL": baseurl.format("entry20")}) 
# 				else:
# 					addMenuentry("ActionTrigger", {"Text": "退院取消", "CommandURL": baseurl.format("entry21")})
# 	elif contextmenuname=="sheettab":  # シートタブの時。
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
# 	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
# def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。	
# 	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
# 	controller = doc.getCurrentController()  # コントローラの取得。
# 	sheet = controller.getActiveSheet()  # アクティブシートを取得。
# 	selection = controller.getSelection()
# 	consts = getConsts(sheet, selection)	
# 	if entrynum==1:  # 同薬品抽出
# 		
# 		
# 		pass
# 	elif entrynum==2:  # 同薬品結合
# 		
# 		
# 		
# 		pass
# 	elif entrynum==5:  # 日付追加。selectionは単一セル。	
# 		setDates(doc, sheet, selection, int(selection.getValue()))  # 経過シートの日付を設定。
# 		if int(selection.getString())!=1:  # 日付が１日でない時。
# 			celladdress = selection.getCellAddress()  # 選択セルアドレスを取得。
# 			r, c = celladdress.Row, celladdress.Column
# 			if c!=consts.splittedcolumn:  # 固定列でないとき。
# 				sheet[r-1, c].setString("")  # 選択セルの上のセルの文字列を消す。
# 	elif entrynum==10:  # 処方。selectionは単一セルか複数セル。
# 		
# 		
# 		pass		
# 	elif entrynum==11:  # 翌月まで。selectionは単一セルか複数セル。
# 		pass		
# 	
# 	
# 	elif 14<entrynum<20:
# 		rangeaddress = selection.getRangeAddress()  # 選択範囲のアドレスを取得。
# 		blackrow = consts.blackrow
# 		emptyrow = consts.emptyrow
# 		if entrynum==15:  # 使用中最上行へ
# 			commons.toOtherEntry(sheet, rangeaddress, blackrow, blackrow+1)
# 		elif entrynum==16:  # 使用中最下行へ
# 			commons.toNewEntry(sheet, rangeaddress, blackrow, emptyrow) 
# 		elif entrynum==17:  # 黒行上へ
# 			commons.toOtherEntry(sheet, rangeaddress, emptyrow, blackrow)  
# 		elif entrynum==18:  # 使用中最上行へ
# 			commons.toOtherEntry(sheet, rangeaddress, emptyrow, blackrow+1)
# 		elif entrynum==19:  # 使用中最下行へ
# 			commons.toNewEntry(sheet, rangeaddress, blackrow, emptyrow) 
# 	elif entrynum in (20, 21):	
# 		if entrynum==20:  # 退院翌日
# 			selection[consts.splittedrow:, :].setPropertyValue("CellBackColor", commons.COLORS["skyblue"])  # 固定行より下すべてに色を付ける。
# 		elif entrynum==21:  # 退院取消
# 			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
# 			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
# 			dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
# 			docframe = controller.getFrame()
# 			c = selection[0, 0].getCellAddress().Column  # 選択セル範囲の一番上のセルの列インデックスを取得。
# 			controller.select(sheet[consts.splittedrow:, c-1])  # 選択列の左の列を選択。
# 			dispatcher.executeDispatch(docframe, ".uno:Copy", "", 0, ())  # コピー。
# 			controller.select(sheet[consts.splittedrow:, c])  # 元の列を選択し直す。
# 			nvs = ("Flags", "T"),\
# 				("FormulaCommand", 0),\
# 				("SkipEmptyCells", False),\
# 				("Transpose", False),\
# 				("AsLink", False),\
# 				("MoveMode", 4)
# 			props = [PropertyValue(Name=n, Value=v) for n, v in nvs]
# 			dispatcher.executeDispatch(docframe, ".uno:InsertContents", "", 0, props)  # 書式のみをペースト。ソースのセル範囲の枠が動く破線のままになるのでEscキーをシミュレートする必要がある。
# 			componentwindow	= controller.ComponentWindow  # コンポーネントウィンドウを取得。
# 			keyevent = KeyEvent(KeyCode=Key.ESCAPE, KeyChar=chr(0x1b), Modifiers=0, KeyFunc=0, Source=componentwindow)  # EscキーのKeyEventを取得。
# 			toolkit = componentwindow.getToolkit()  # ツールキットを取得。
# 			toolkit.keyPress(keyevent)  # キーを押す、をシミュレート。
# 			toolkit.keyRelease(keyevent)  # キーを離す、をシミュレート。
def setDates(doc, sheet, cell, datevalue):  # sheet:経過シート、cell: 日付開始セル、dateserial: 日付開始日のシリアル値。
	createFormatKey = commons.formatkeyCreator(doc)	
	colors = commons.COLORS
	holidays = commons.HOLIDAYS
	daycount = 100  # 経過シートに入力する日数。
	celladdress = cell.getCellAddress()  # 経過シートの日付の開始セルのセルアドレスを取得。
	r, c = celladdress.Row, celladdress.Column
	sheet[:r+1, c:].clearContents(511)  # 開始列より右の日付行の内容を削除。
	endcolumn = c + daycount + 1
	endcolumn = endcolumn if endcolumn<1024 else 1023  # 列インデックスの上限1023。
	sheet[r, c:endcolumn].setDataArray(([i for i in range(datevalue, datevalue+daycount+1)],))  # 日時シリアル値を経過シートに入力。
	sheet[r, c:endcolumn].setPropertyValue("NumberFormat", createFormatKey('YYYY/M/D'))  # 日時シリアルから年月日の取得のため一時的に2018/5/4の形式に変換する。
	y, m, d = map(int, sheet[r, c].getString().split("/"))  # 年、月、日を整数型で取得。
	weekday, days = calendar.monthrange(y, m)  # 月曜日が曜日番号0。1日の曜日と一月の日数のタプルが返る。
	weekday = (weekday+(d-1)%7) % 7  # dの曜日番号を取得。1日からの7の余りと1日の余りを加えた7の余りがdの曜日番号。
	n = 6  # 日曜日の曜日番号。
	sunsset = set(range(c+(n-weekday)%7, endcolumn, 7))  # 日曜日の列インデックスの集合。祝日と重ならないようにあとで使用する。
	setRangesProperty(doc, sheet, r, sunsset, ("CharColor", colors["red3"]))  # 日曜日の文字色を設定。
	n = 5  # 土曜日の曜日番号。
	setRangesProperty(doc, sheet, r, range(c+(n-weekday)%7, endcolumn, 7), ("CharColor", colors["skyblue"]))  # 土曜日の文字色を設定。	
	holidayset = set()  # 祝日の列インデックスを入れる集合。
	days = days - d + 1  # 翌月1日までの日数を取得。
	mr = r - 1  # 月を代入する行のインデックス。
	mc = c  # 1日を表示する列のインデックス。最初の月のみ開始日になる。
	if y in holidays:  # 祝日一覧のキーがある時。
		holidayset.update(mc+i-d for i in holidays[y][m-1] if i>=d)  # 祝日の列インデックスの集合を取得。
	while True:
		sheet[mr, mc].setString("{}月".format(m))  # 月を入力。
		mc += days  # 次月1日の列に進める。
		if mc<endcolumn:  # 日時シリアル値が入力されている列の時。
			ymd = sheet[r, mc].getString()  # 1日の年/月/日を取得。
			y, m = map(int, ymd.split("/")[:2])  # 年と月を取得。
			if y in holidays:  # 祝日一覧のキーがある時。
				holidayset.update(mc+i-1 for i in holidays[y][m-1] if mc+i-1<endcolumn)  # 祝日の列インデックスの集合を取得。
			weekday, days = calendar.monthrange(y, m)  # 1日の曜日と月の日数を取得。
		else:
			break	
	holidayset.difference_update(sunsset)  # 日曜日と重なっている祝日を除く。
	setRangesProperty(doc, sheet, r, holidayset, ("CellBackColor", colors["red3"]))  # 祝日の背景色を設定。	
	sheet[r, c:endcolumn].setPropertyValues(("NumberFormat", "HoriJustify"), (createFormatKey('D'), CENTER))  # 経過シートの日付の書式を設定。	
def setRangesProperty(doc, sheet, r, columnindexes, prop):  # r行のcolumnindexesの列のプロパティを変更。prop: プロパティ名とその値のリスト。
	sheetcellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
	sheetcellranges.addRangeAddresses([sheet[r, i].getRangeAddress() for i in columnindexes], False)  # セル範囲コレクションを取得。
	if len(sheetcellranges):  # sheetcellrangesに要素がないときはsetPropertyValue()でエラーになるので要素の有無を確認する。
		sheetcellranges.setPropertyValue(*prop)  # セル範囲コレクションのプロパティを変更。
def drowBorders(selection):  # ターゲットを交点とする行列全体の外枠線を描く。
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上端のセルアドレスを取得。
	r, c = celladdress.Row, celladdress.Column # selectionの行と列のインデックスを取得。		
	sheet = VARS.sheet
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = commons.createBorders()
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
	rangeaddress = selection.getRangeAddress() # 選択範囲のセル範囲アドレスを取得。
	if r<VARS.splittedrow:  # 分割行より上の時。
		if c<VARS.splittedcolumn:  # 分割列より左の時。
			return  # 線を消すだけ。
		else:  # 分割列含む右の時は縦線を引くだけ。
			sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。				
	else:  # 分割行以下の時。
		if r==VARS.blackrow:  # 黒行の時。
			return  # 線を消すだけ。
		sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。	
		if c!=VARS.yakucolumn:  # 薬名列でない時。縦線も引く。
			sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。			
		selection.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。
