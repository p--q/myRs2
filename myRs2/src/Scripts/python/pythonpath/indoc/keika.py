#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 経過シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import calendar
from itertools import chain
from indoc import commons, ichiran, dialogs
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.table.CellHoriJustify import CENTER, LEFT  # enum
from com.sun.star.awt import MouseButton, MessageBoxButtons  # 定数
from com.sun.star.awt.MessageBoxType import ERRORBOX  # enum
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.lang import Locale  # Struct
class Keika():  # シート固有の定数設定。
	def __init__(self, sheet):
		self.daterow = 1  # 日付行インデックス。
		self.splittedrow = 4  # 分割行インデックス。
		self.yakucolumn = 5  # 薬名列インデックス。
		self.splittedcolumn = 9  # 分割列インデックス。
		cellranges = sheet[:, self.yakucolumn].queryContentCells(CellFlags.STRING)  # 薬名列の文字列が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 薬名列の最終行インデックス+1を取得。
		gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in (commons.COLORS["black"],))
		self.blackrow = next(gene)  # 黒行インデックスを取得。			
def getSectionName(sheet, target):  # 区画名を取得。
	"""
	A  ||  B
	===========  # 行の固定の境界。||は列の固定の境界。境界の行と列はそれぞれ下、右に含む。
	C  ||  D
	I-----------  # 黒行。この行は含まない。
	E  ||  F
	-----------  # 薬品列の最下行の一つ下の行。
	G  ||  H
	
	"""
	keika = Keika(sheet)  # クラスをインスタンス化。	
	splittedrow = keika.splittedrow
	splittedcolumn = keika.splittedcolumn
	blackrow = keika.blackrow
	emptyrow = keika.emptyrow
	rangeaddress = target[0, 0].getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	sectionname = ""
	if splittedrow<blackrow:
		if len(sheet[splittedrow:blackrow, :splittedcolumn].queryIntersection(rangeaddress)): 
			sectionname = "C"			
		elif len(sheet[splittedrow:blackrow, splittedcolumn:].queryIntersection(rangeaddress)): 
			sectionname = "D"			
	elif blackrow+1<emptyrow:
		if len(sheet[blackrow+1:emptyrow, :splittedcolumn].queryIntersection(rangeaddress)): 
			sectionname = "E"				
		elif len(sheet[blackrow+1:emptyrow, splittedcolumn:].queryIntersection(rangeaddress)): 
			sectionname = "F"	
	if not sectionname:		
		if len(sheet[:splittedrow, :splittedcolumn].queryIntersection(rangeaddress)): 
			sectionname = "A"	
		elif len(sheet[:splittedrow, splittedcolumn:].queryIntersection(rangeaddress)): 
			sectionname = "B"					
		elif len(sheet[emptyrow:, :splittedcolumn].queryIntersection(rangeaddress)): 
			sectionname = "G"					
		elif len(sheet[emptyrow:, splittedcolumn:].queryIntersection(rangeaddress)): 
			sectionname = "H"
		else:
			sectionname = "I"
	keika.sectionname = sectionname  # 区画名
	return keika  
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	sheet["F1:G1"].setDataArray((("一覧へ", "ｶﾙﾃへ"),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet["F3:F4"].setDataArray((("薬品整理",), ("薬品名抽出",)))
	keika = Keika(sheet)
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。			
	daterow = keika.daterow
	splittedcolumn = keika.splittedcolumn
	startdatevalue = int(sheet[daterow, splittedcolumn].getValue())  # 日付行の最初のセルから日付のシリアル値の取得。
	todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。
	sheet[daterow-1, splittedcolumn:].setPropertyValue("CellBackColor", -1)  # r-1行目の背景色をクリア。
	c = splittedcolumn + (todayvalue - startdatevalue)  # 今日の日付の列インデックスを取得。
	if c<1024:
		sheet[daterow-1, c].setPropertyValue("CellBackColor", commons.COLORS["violet"])  # 日付行の上のセルの今日の背景色を設定。
	sheet[daterow+2:, splittedcolumn:].setPropertyValue("HoriJustify", LEFT)  # 分割列以降、日付行2行下以降すべて左詰めにする。
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。		
	target = enhancedmouseevent.Target  # ターゲットのセルを取得。
	sheet = target.getSpreadsheet()
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				drowBorders(sheet, target, commons.createBorders())  # 枠線の作成。
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
				controller = doc.getCurrentController()  # コントローラの取得。
				keika = getSectionName(sheet, target)  # セル固有の定数を取得。
				sectionname = keika.sectionname  # クリックしたセルの区画名を取得。
				yakucolumn = keika.yakucolumn
				txt = target.getString()  # クリックしたセルの文字列を取得。	
				if sectionname=="A":
					ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
					smgr = ctx.getServiceManager()  # サービスマネージャーの取得。						
					sheets = doc.getSheets()  # シートコレクションを取得。
					if txt=="一覧へ":
						controller.setActiveSheet(sheets["一覧"])  # 一覧シートをアクティブにする。
					elif txt=="ｶﾙﾃへ":  # カルテシートをアクティブにする、なければ作成する。
						datarow = sheet[1, yakucolumn:keika.splittedcolumn+1].getDataArray()[0]  # IDセルから最初の日付セルまで取得。
						idcelltxts = datarow[0].split(" ")  # 半角スペースで分割。
						idtxt = idcelltxts[0]  # 最初の要素を取得。
						if idtxt.isdigit():  # IDが数値のみの時。					
							sheets = doc.getSheets()
							if idtxt in sheets:  # ID名のシートがあるとき。
								controller.setActiveSheet(sheets[idtxt])  # カルテシートをアクティブにする。
							else:
								if len(idcelltxts)==5:  # ID、漢字姓・名、カタカナ姓・名、の5つに分割できていた時。
									kanjitxt, kanatxt = " ".join(idcelltxts[1:3]), " ".join(idcelltxts[3:])
									datevalue = datarow[-1]
									karutesheet = ichiran.getKaruteSheet(commons.formatkeyCreator(doc), sheets, idtxt, kanjitxt, kanatxt, datevalue)
									controller.setActiveSheet(karutesheet)  # カルテシートをアクティブにする。
								else:
									commons.showErrorMessageBox(controller, "「ID(数値のみ) 漢字姓 名 カナ姓 名」の形式になっていません。")
						else:
							commons.showErrorMessageBox(controller, "IDが取得できませんでした。")	
					elif txt=="薬品整理":  # クリックするたびに終了順、昇順に並び替える。黒行の上のみ。
						if keika.splittedrow>keika.blackrow:  # 分割行から黒行より上に行がある時のみ。
							datarange = sheet[keika.splittedrow:keika.blackrow, :]  # 黒行より上の行のセル範囲を取得。
							controller.select(datarange)  # ソートするセル範囲を取得。
							if target.getPropertyValue("CellBackColor")==-1:  # ボタンの背景色がない時、薬名列の昇順でソート。
								target.setPropertyValue("CellBackColor", commons.COLORS["lime"])  # ボタンの背景色を付ける。				
								props = PropertyValue(Name="Col1", Value=yakucolumn+1),  # Col1の番号は優先順位。Valueはインデックス+1。 			
							else:  # ボタンの背景色がある時、終了順でソート。終了列インデックスを先頭列に代入しておく。
								datarows = []  # 終了行インデックスを入れる行のリスト。
								for i in range(keika.blackrow-keika.splittedrow):  # 分割行インデックスから、黒行の上までの相対インデックスを取得。
									cellranges = datarange[i, keika.splittedcolumn:].queryContentCells(CellFlags.STRING)  # 文字列のあるセル範囲コレクションを取得。
									if len(cellranges):  # セル範囲が取得出来た時。
										datarows.append((cellranges.getRangeAddresses()[-1].EndColumn,))  # 最終列インデックスを取得。
									else:
										datarows.append((1,))  # 色セルがない行は1にして上に持ってくる。0にするとFalseになってしまう。
								datarange[:, 0].setDataArray(datarows)  # 開始列インデックスをシートに代入。
								datarange[:, 0].setPropertyValue("CharColor", commons.COLORS["white"])  # 先頭列の文字色を白色にする。
								target.setPropertyValue("CellBackColor", -1)  # ボタンの背景色を消す。		
								props = PropertyValue(Name="Col1", Value=1),  # Col1の番号は優先順位。Valueはインデックス+1。 
							dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
							dispatcher.executeDispatch(controller.getFrame(), ".uno:DataSort", "", 0, props)  # ディスパッチコマンドでソート。sort()メソッドは挙動がおかしくて使えない。								
							controller.select(target)  # ボタンを選択し直す。	
					elif txt=="薬品名抽出":
						firstrow = max(sheet[:, i].queryContentCells(CellFlags.STRING).getRangeAddresses()[-1].EndRow for i in (yakucolumn+1, yakucolumn+2)) + 1  # 用法列か回数列の最終行インデックスの下の行インデックスを取得。
						if firstrow<keika.emptyrow:
							transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
							transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))  # 全角文字を半角にする。
							datarows = sheet[firstrow:keika.emptyrow, yakucolumn].getDataArray()  # 用法設定していない薬品列の各行のタプルを取得。
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
							sheets[firstrow:keika.emptyrow, yakucolumn:keika.splittedcolumn].clearContents(CellFlags.STRING+CellFlags.VALUE)  # 整理前のセルの文字列と数値をクリア。		
							sheets[firstrow:firstrow+len(newdatarows), yakucolumn].setDataArray(newdatarows)  # 整理した薬品名をシートに代入。		
					elif txt[:8].isdigit():  # 最初8文字が数値の時。						
						systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
						systemclipboard.setContents(commons.TextTransferable(txt[:8]), None)  # クリップボードにIDをコピーする。							
					return False  # セル編集モードにしない。	
				elif sectionname=="B":
					celladdress = target.getCellAddress()
					r = celladdress.Row  # ダブルクリックしたセルの行インデックス、列インデックスを取得。
					items = []
					if r==2:	
						items = ["", "○", "尿"]			
						horijustify	= CENTER
					elif r==3:
						items = ["", "胸Xp", "腹ｴ", "心ｴ"]
						horijustify	= LEFT
					if items:
						if txt in items:  # セルの内容にある時。
							items.append(items[0])  # 最初の要素を最後の要素に追加する。
							dic = {items[i]: items[i+1] for i in range(len(items)-1)}  # 順繰り辞書の作成。
							txt = dic[txt]  # 次の要素を代入する。	
							target.setString(txt)
					if txt:  # 文字がある時。
						target.setPropertyValues(("CellBackColor", "HoriJustify"), (commons.COLORS["skyblue"], horijustify))  # 背景をスカイブルーにする。		
					else:
						target.setPropertyValue("CellBackColor", -1)  # 背景色を消す。
					return False  # セル編集モードにしない。	
				elif sectionname in ("C", "E", "G"):
					celladdress = target.getCellAddress()  # ターゲットのセルアドレスを取得。
					r, c = celladdress.Row, celladdress.Column
					if c==yakucolumn:  # 薬名列の時。
						return True  # セル編集モードにする。
					elif c==yakucolumn+1:  # 用法列の時。
						if txt:
# 							dialogs.createDialog(xscriptcontext)
							pass
							
						else:
							target.setString("分3")
					elif c==yakucolumn+2:  # 回数列の時。
						if txt:
							
							
							pass
						else:
							target.setString("持続")						
					return False  # セル編集モードにしない。	
				elif sectionname in ("D", "F"):
					if txt:
						
						
						pass
					else:
						target.setString("止")								
					return False  # セル編集モードにしない。		
	return True  # セル編集モードにする。
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	controller = eventobject.Source
	sheet = controller.getActiveSheet()
	selection = controller.getSelection()
	if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択範囲がセルの時。矢印キーでセルを移動した時。マウスクリックハンドラから呼ばれると何回も発火するのでその対応。
		currenttableborder2 = selection.getPropertyValue("TableBorder2")  # 選択セルの枠線を取得。
		if all((currenttableborder2.TopLine.Color==currenttableborder2.LeftLine.Color==commons.COLORS["violet"],\
				currenttableborder2.RightLine.Color==currenttableborder2.BottomLine.Color==commons.COLORS["magenta3"])):  # 枠線の色を確認。
			return  # すでに枠線が書いてあったら何もしない。
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
		drowBorders(sheet, selection, commons.createBorders())  # 枠線の作成。
def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。		
	changes = changesevent.Changes	
	for change in changes:
		if change.Accessor=="cell-change":  # セルの値が変化した時。
			target = change.ReplacedElement  # 値を変更したセルを取得。	
			sheet = target.getSpreadsheet()
			keika = getSectionName(sheet, target)  # 経過シート固有の定数を取得。
			sectionname = keika.sectionname  # クリックしたセルの区画名を取得。
			if not sectionname in ("A",):  # 領域A以外の時。
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
				transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
				transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))					
				txt = target.getString()  # セルの文字列を取得。	
				txt = transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換。
				target.setString(txt)
			break
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。				
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
	del contextmenu[:]  # contextmenu.clear()は不可。
	target = controller.getSelection()  # 現在選択しているセル範囲を取得。
	keika = getSectionName(sheet, target)  # セル固有の定数を取得。
	sectionname = keika.sectionname  # クリックしたセルの区画名を取得。		
	if contextmenuname=="cell":  # セルのとき		
		if sectionname in ("A",):
			if target.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
				txt = target.getString()
				if txt=="薬品整理":
					addMenuentry("ActionTrigger", {"Text": "同薬品抽出", "CommandURL": baseurl.format("entry1")}) 
					addMenuentry("ActionTrigger", {"Text": "同薬品結合", "CommandURL": baseurl.format("entry2")}) 
			return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
		elif sectionname in ("B",):
			if target.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
				celladdress = target.getCellAddress()  # 経過シートの日付の開始セルのセルアドレスを取得。
				r = celladdress.Row
				if r==keika.daterow:  # 日付行の時。
					if target.getValue():  # セルに値があるとき。
						addMenuentry("ActionTrigger", {"Text": "日付追加", "CommandURL": baseurl.format("entry5")}) 
				elif r==keika.daterow+2:  # 処置行の時。
					commons.cutcopypasteMenuEntries(addMenuentry)
			return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
		elif sectionname in ("D", "F"):
			addMenuentry("ActionTrigger", {"Text": "処方", "CommandURL": baseurl.format("entry10")})
			addMenuentry("ActionTrigger", {"Text": "翌月まで", "CommandURL": baseurl.format("entry11")})  # 	回数列が空欄の時は金まで、それ以外は火まで。
			
# 			if target.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
# 				pass
# 			elif target.supportsService("com.sun.star.sheet.SheetCellRange"):  # セル以外のセル範囲の時、つまり連続した複数セルの時。
# 				addMenuentry("ActionTrigger", {"Text": "翌月まで", "CommandURL": baseurl.format("entry11")})  # 	回数列が空欄の時は金まで、それ以外は火まで。
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
			commons.cutcopypasteMenuEntries(addMenuentry)
		elif sectionname in ("C", "E", "G", "H"):	
			commons.cutcopypasteMenuEntries(addMenuentry)
	elif contextmenuname=="rowheader":  # 行ヘッダーのとき。	
		if sectionname in ("I",):
			return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
		elif sectionname in ("C",):
			addMenuentry("ActionTrigger", {"Text": "使用中最上行へ", "CommandURL": baseurl.format("entry15")})  
			addMenuentry("ActionTrigger", {"Text": "使用中最下行へ", "CommandURL": baseurl.format("entry16")}) 
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		elif sectionname in ("E",):	
			addMenuentry("ActionTrigger", {"Text": "黒行上へ", "CommandURL": baseurl.format("entry17")})  
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			addMenuentry("ActionTrigger", {"Text": "使用中最上行へ", "CommandURL": baseurl.format("entry15")})  
			addMenuentry("ActionTrigger", {"Text": "使用中最下行へ", "CommandURL": baseurl.format("entry16")}) 
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		commons.rowMenuEntries(addMenuentry)
	elif contextmenuname=="colheader":  # 列ヘッダーの時。
		if sectionname in ("B",):
			addMenuentry("ActionTrigger", {"Text": "退院翌日", "CommandURL": baseurl.format("entry20")}) 
			addMenuentry("ActionTrigger", {"Text": "退院取消", "CommandURL": baseurl.format("entry21")})
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	selection = controller.getSelection()
# 	keika = getSectionName(sheet, selection)	
	keika = Keika(sheet)	
	
		
	
	
	if entrynum==1:  # 同薬品抽出
		pass
	elif entrynum==2:  # 同薬品結合
		pass
	elif entrynum==5:  # 日付追加。selectionは単一セル。
		setDates(doc, sheet, selection, int(selection.getValue()))  # 経過シートの日付を設定。
		if int(selection.getString())!=1:  # 日付が１日でない時。
			celladdress = selection.getCellAddress()  # 選択セルアドレスを取得。
			r, c = celladdress.Row, celladdress.Column
			if c!=keika.splittedcolumn:  # 固定列でないとき。
				sheet[r-1, c].setString("")  # 選択セルの上のセルの文字列を消す。
	elif entrynum==10:  # 処方。selectionは単一セルか複数セル。
		pass		
	elif entrynum==11:  # 翌月まで。selectionは複数セル。
		pass		
	
	
	elif entrynum in (15, 16, 17):
		if len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 列全体が選択されている場合もあるので行全体が選択されていることを確認する。
			if entrynum==15:  # 使用中最上行へ
				
				
				pass		
			elif entrynum==16:  # 使用中最下行へ
				pass		
			elif entrynum==17:  # 黒行上へ
				pass	
	
	elif entrynum in (20, 21):	
		if len(selection[:, 0].getRows())==len(sheet[:, 0].getRows()):  # 行全体が選択されている場合もあるので列全体が選択されていることを確認する。
			if entrynum==20:  # 退院翌日
				pass		
			elif entrynum==21:  # 退院取消
				pass
	
	
# def toNewEntry(sheet, rangeaddress, edgerow, dest_row):  # 新入院へ。新規行挿入は不要。
# 	startrow, endrowbelow = rangeaddress.StartRow, rangeaddress.EndRow+1  # 選択範囲の開始行と終了行の取得。
# 	if endrowbelow>edgerow:
# 		endrowbelow = edgerow
# 	sourcerangeaddress = sheet[startrow:endrowbelow, :].getRangeAddress()  # コピー元セル範囲アドレスを取得。
# 	sheet.moveRange(sheet[dest_row, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。	
# 	sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動したソース行を削除。
	
				
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
def drowBorders(sheet, cellrange, borders):  # ターゲットを交点とする行列全体の外枠線を描く。
	cell = cellrange[0, 0]  # セル範囲の左上端のセルで判断する。
	keika = getSectionName(sheet, cell)
	sectionname = keika.sectionname
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = borders	
	rangeaddress = cellrange.getRangeAddress()  # セル範囲アドレスを取得。
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
	if sectionname in ("A", "G", "H", "I"):  # 線を消すだけ。
		return
	if sectionname in ("D", "F"):  # 縦横線を引く。
		sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。			
		sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。		
	elif sectionname in ("B",):  # 縦線のみ引く。
		sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。				
	elif sectionname in ("C", "E"):  # 横線のみ引く。		
		sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。	
	cellrange.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	






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
