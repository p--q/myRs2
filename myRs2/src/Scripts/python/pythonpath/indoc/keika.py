#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import calendar
from indoc import commons, ichiran
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.table.CellHoriJustify import CENTER  # enum
from com.sun.star.awt import MouseButton, MessageBoxButtons  # 定数
from com.sun.star.table.CellHoriJustify import LEFT  # enum
from com.sun.star.awt.MessageBoxType import ERRORBOX  # enum
from com.sun.star.beans import PropertyValue  # Struct
# from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
class Keika():  # シート固有の定数設定。
	def __init__(self, sheet):
		self.daterow = 1  # 日付行インデックス。
		self.splittedrow = 4  # 分割行インデックス。
		self.yakucolumn = 5  # 薬名列インデックス。
		self.splittedcolumn = 8  # 分割列インデックス。
		cellranges = sheet[:, self.yakucolumn].queryContentCells(CellFlags.STRING)  # 薬名列の文字列が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 薬名列の最終行インデックス+1を取得。
		gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in (commons.COLORS["black"],))
		self.blackrow = next(gene)  # 黒行インデックスを取得。			
def getSectionName(sheet, target):  # 区画名を取得。
	"""
	A  ||  B
	===========  # 行の固定の境界。||は列の固定の境界。境界の行と列はそれぞれ下、右に含む。
	C  ||  D
	-----------  # 薬品列の最下行の一つ下の行。
	E  ||  F
	"""
	keika = Keika(sheet)  # クラスをインスタンス化。	
	splittedrow = keika.splittedrow
	splittedcolumn = keika.splittedcolumn
	emptyrow = keika.emptyrow
# 	cellranges = sheet[:, keika.yakucolumn].queryContentCells(CellFlags.STRING)  # 薬名列の文字列が入っているセルに限定して抽出。
# 	emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 薬名列の最終行インデックス+1を取得。
	rangeaddress = target.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	if len(sheet[splittedrow:emptyrow, splittedcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "D"	
	elif len(sheet[emptyrow:, :splittedcolumn].queryIntersection(rangeaddress)): 
		sectionname = "E"			
	elif len(sheet[:splittedrow, :splittedcolumn].queryIntersection(rangeaddress)): 
		sectionname = "A"			
	elif len(sheet[splittedrow:emptyrow, :splittedcolumn].queryIntersection(rangeaddress)): 
		sectionname = "C"	
	elif len(sheet[:splittedrow, splittedcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "B"			
	else:
		sectionname = "F"	
	keika.sectionname = sectionname  # 区画名
	return keika  
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	sheet["F1:G1"].setDataArray((("一覧へ", "ｶﾙﾃへ"),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet["F3:F4"].setDataArray((("薬品整理",), ("薬品名抽出",)))
	keika = Keika()
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
				txt = target.getString()  # クリックしたセルの文字列を取得。	
				if sectionname=="A":
					sheets = doc.getSheets()  # シートコレクションを取得。
					if txt=="一覧へ":
						controller.setActiveSheet(sheets["一覧"])  # 一覧シートをアクティブにする。
					elif txt=="ｶﾙﾃへ":  # カルテシートをアクティブにする、なければ作成する。
						datarow = sheet[1, keika.yakucolumn:keika.splittedcolumn+1].getDataArray()[0]  # IDセルから最初の日付セルまで取得。
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
					elif txt=="薬品整理":  # クリックするたびに初使用順、昇順に並び替える。黒行の上のみ。
						if keika.splittedrow>keika.blackrow:
							ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
							smgr = ctx.getServiceManager()  # サービスマネージャーの取得。								
							datarange = sheet[keika.splittedrow:keika.blackrow, :]  # 黒行より上の行のセル範囲を取得。
							datarange[:, 0].setDataArray([(i,) for i in range(keika.blackrow-keika.splittedrow)])  # 列インデックス0に行の順番を代入。
							datarows = list(datarange.getDataArray())  # 行をリストにして取得。
							sortkeycolumnindex = keika.yakucolumn  # 薬名列インデックスを取得。
							datarows.sort(key=lambda x:x[sortkeycolumnindex])  # 各行を薬名列インデックスでソート。
							
							
							
							controller.select(datarange)
							propertyvalue = PropertyValue(Name="Col1", Value=keika.yakucolumn)  # 薬名列インデックスでソートする。
							dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
							dispatcher.executeDispatch(controller.getFrame(), ".uno:DataSort", "", 0, (propertyvalue,))

							
# 							datarange.sort()
							
							
							
							datarows = list(map(list, datarange.getDataArray()))  # 各行をリストにして取得。
							orders = list(range(len(datarows)))  # 昇順の番号のリストを取得。
							for i in orders:
								datarows[i][0] = i  # 列インデックス0に行の順番を代入。
							sortkeycolumnindex = keika.yakucolumn  # 薬名列インデックスを取得。
							
							datarows.sort(key=lambda x:x[sortkeycolumnindex])  # 各行を薬名列インデックスでソート。
							
							if orders==[datarows[i][0] for i in orders]:  # 順番が入れ替わっていない時、初使用順にソートする。
								for i in range(keika.splittedrow, keika.blackrow-keika.splittedrow):  # 分割行インデックスから、黒行の前まで。
									for j in range(keika.splittedcolumn, 1024-keika.splittedcolumn):  # 開始日列インデックスから最終列まで。
										if sheet[i, j].getPropertyValue()!=-1:  # 背景色がある時。
											datarows[i][0] = j  # データ行の0列目に列インデックスを代入。
											break
										
								datarows.sort(key=lambda x:x[0])  # 各行を列インデックス0でソート。
								
							datarange.setDataArray(datarows)
							sheet[keika.splittedrow:keika.blackrow, 0].clearContents(511)  # 黒行より上の列インデックス0のセルをクリア。
					elif txt=="薬品名抽出":
						pass
							
							
							
					elif txt[:8].isdigit():  # 最初8文字が数値の時。
						ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
						smgr = ctx.getServiceManager()  # サービスマネージャーの取得。						
						systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
						systemclipboard.setContents(commons.TextTransferable(txt[:8]), None)  # クリップボードにIDをコピーする。							
					return False  # セル編集モードにしない。	
				elif sectionname=="B":
					celladdress = target.getCellAddress()
					r = celladdress.Row  # ダブルクリックしたセルの行インデックス、列インデックスを取得。
					items = []
					if r==2:	
						items = ["", "○", "尿"]						
					elif r==3:
						items = ["", "胸Xp", "腹ｴ", "心ｴ"]
					if items:
						if txt in items:  # セルの内容にある時。
							items.append(items[0])  # 最初の要素を最後の要素に追加する。
							dic = {items[i]: items[i+1] for i in range(len(items)-1)}  # 順繰り辞書の作成。
							txt = dic[txt]  # 次の要素を代入する。	
							target.setString(txt)
					if txt:  # 文字がある時。
						target.setPropertyValue("CellBackColor", commons.colors["skyblue"])  # 背景をスカイブルーにする。		
					else:
						target.setPropertyValue("CellBackColor", -1)  # 背景色を消す。
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
	pass
		
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。				
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
	del contextmenu[:]  # contextmenu.clear()は不可。
	target = controller.getSelection()  # 現在選択しているセル範囲を取得。	
	if contextmenuname=="cell":  # セルのとき
		keika = getSectionName(sheet, target)  # セル固有の定数を取得。
		sectionname = keika.sectionname  # クリックしたセルの区画名を取得。			
		if sectionname in ("A",):  # 固定行より上の時はコンテクストメニューを表示しない。
			return EXECUTE_MODIFIED
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Delete"})			
		
		
# 		del contextmenu[:]  # contextmenu.clear()は不可。
# 		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
# 		karute.rng	addMenuentry("ActionTrigger", {"Text": "To blue", "CommandURL": baseurl.format("entry1")})  # listeners.pyの関数名を指定する。
# 		elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 連続した複数セルの時。
# 			addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format("entry2")})  # listeners.pyの関数名を指定する。
# 		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})

# 		if target.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
# 			addMenuentry("ActionTrigger", {"Text": "To Green", "CommandURL": baseurl.format("entry1")}) 
# 		elif target.supportsService("com.sun.star.sheet.SheetCellRange"):  # 連続した複数セルの時。
# 			addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format("entry2")}) 


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
def contextMenuEntries(target, entrynum):  # コンテクストメニュー番号の処理を振り分ける。
	colors = commons.COLORS
	if entrynum==1:
		target.setPropertyValue("CellBackColor", colors["blue3"])  # 背景を青色にする。
	elif entrynum==2:
		target.setPropertyValue("CellBackColor", colors["red3"]) 		
		
		
def setDates(doc, sheet, cell, datevalue):  # sheet:経過シート、cell: 日付開始セル、dateserial: 日付開始日のシリアル値。。
	createFormatKey = commons.formatkeyCreator(doc)	
	colors = commons.COLORS
	holidays = commons.HOLIDAYS
	daycount = 100  # 経過シートに入力する日数。
	celladdress = cell.getCellAddress()  # 経過シートの日付の開始セルのセルアドレスを取得。
	r, c = celladdress.Row, celladdress.Column
	sheet[:r+1, c:].clearContents(CellFlags.VALUE+CellFlags.DATETIME+CellFlags.STRING+CellFlags.ANNOTATION+CellFlags.FORMULA+CellFlags.HARDATTR+CellFlags.STYLES)  # セルの内容を削除。
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
	if sectionname in ("A", "E", "F"):  # 線を消すだけ。
		return
	if sectionname in ("D",):  # 縦横線を引く。
		sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。			
		sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。		
	elif sectionname in ("B",):  # 縦線のみ引く。
		sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。				
	elif sectionname in ("C",):  # 横線のみ引く。		
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
