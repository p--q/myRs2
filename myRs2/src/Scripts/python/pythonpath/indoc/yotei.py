#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import commons
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.sheet.CellDeleteMode import COLUMNS as delete_columns  # enum
from com.sun.star.sheet.CellInsertMode import COLUMNS as insert_columns  # enum
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.table.CellHoriJustify import CENTER  # enum
class Schedule():  # シート固有の定数設定。
	def __init__(self):
		self.menurow = 0  # メニュー行。
		self.daterow = 2  # 日付行。
		self.datecolumn = 1  # 日付開始列。
		self.settingcolumn = 47  # 設定開始列。
	def setSheet(self, sheet):
		self.sheet = sheet
		cellranges = sheet[:, 0].queryContentCells(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME)  # 先頭列の文字列か数値が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 先頭列の最終行インデックス+1を取得。		
		cellranges = sheet[self.daterow+1, :].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # 曜日行の文字列か数値が入っているセルに限定して抽出。
		rangeaddresses = cellranges.getRangeAddresses()	
		self.firstemptycolumn = rangeaddresses[0].EndColumn+1  # 日付行の区切れの列インデックスを取得。
		self.templatestartcolumn = rangeaddresses[1].StartColumn
		self.templateendcolumnedge = rangeaddresses[1].EndColumn + 1
VARS = Schedule()
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	VARS.setSheet(sheet)
	sheet["A1"].setString("ﾘｽﾄに戻る")
	sheet["AF1"].setString("COPY")
	sheet["AK1"].setString("強有効")
	firstdatevalue = int(sheet[VARS.daterow, VARS.datecolumn].getValue())  # 最初の日付のシリアル値を整数で取得。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。
	todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。	
	diff = todayvalue - firstdatevalue
	if not diff>0:  # 先頭日付が今日より前でない時はここで終わる。
		return
	daycount = 31  # シートに表示する日数。
	if VARS.datecolumn+daycount>VARS.templatestartcolumn:
		daycount = VARS.templatestartcolumn - VARS.datecolumn  # 右上限はテンプレート列までにする。
	weekdays = "日", "月", "火", "水", "木", "金", "土"  # シートでは日=1であることに注意。
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。		
	weekdayidx = int(functionaccess.callFunction("WEEKDAY", (todayvalue,))) - 1   # 今日の曜日をインデックスで取得。
	templates = sheet[VARS.daterow:VARS.emptyrow, VARS.templatestartcolumn+1:VARS.templateendcolumnedge].getDataArray()  # テンプレートの値を日付行から取得。
	datarows = [[i for i in range(todayvalue, todayvalue+daycount)],\
					[weekdays[i%7] for i in range(weekdayidx, weekdayidx+daycount)]]  # 日付行と曜日行を作成。
	todaycolumn = VARS.datecolumn + diff  # 移動前の今日の日付列。
	if todaycolumn<VARS.firstemptycolumn:
		datarows.extend(sheet[VARS.daterow+2:VARS.emptyrow, todaycolumn:VARS.firstemptycolumn].getDataArray())  # 今日の日付の列以降の行を取得。
	else:
		datarows.extend([] for dummy in range(VARS.emptyrow-(VARS.daterow+2)))
	for i in range(len(datarows[0])):  # インデックスがVARS.firstemptycolumn-VARS.datecolumn+diff以降は2行目以降の要素はない。
		y = datarows[1][i]  # 曜日の文字列を取得。
		
	# 土曜日と日曜日の文字色を設定。
	# 祝日の背景色を設定。
			
		for k in range(len(templates[0])):  # テンプレートの列インデックスをイテレート。
			if y in templates[1][k]:  # 曜日が一致する時。
				w = templates[0][k]  # 週数行の値を取得。
				if w.endswith("w"):  # wで終わる時は週番号。
					d = int(functionaccess.callFunction("DAY", (datarows[0][i],)))  # 月の何日目から取得。
					if int(w[:-1])!=int(d/7)+1:  # 週番号が一致しない時。
						
						8/7と8/14がかかる。
						
						continue  # 週番号が一致しない時は次のループに行く。
				elif w.endswith("d"):  # dで終わる時は月のd日目。
					if int(w[:-1])!=int(functionaccess.callFunction("DAY", (datarows[0][i],))):  # d日目でない時。
						continue  # d日目ではない時は次のループに行く。
				elif isinstance(w, float):  # float型の時は日付シリアル値。
					if datarows[0][i]!=int(w):  # 日付が一致しない時は次の列に行く。
						continue
				for j in range(2, len(datarows)):  # 行インデックスをイテレート。
					if i<len(datarows[j]):  # 行にインデックスがある時。
						if datarows[j][i] in ("", "/", "x"):  # テンプレートを優先する文字列の時。
							datarows[j][i] = templates[j][k]  # テンプレートの文字列を採用。
					else:
						datarows[j].append(templates[j][k])	
						
					# datarows[j][i]をみてセルの背景色の設定。	
					
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)	
	
	endedgecolumn = VARS.datecolumn+daycount					
	sheet[VARS.daterow-1:VARS.emptyrow, VARS.datecolumn:endedgecolumn].clearContents(511)  # 内容を削除。
	sheet[VARS.daterow:VARS.emptyrow, VARS.datecolumn:endedgecolumn].setDataArray(datarows)
						
	# 土曜日と日曜日の文字色を設定。
	# 祝日の背景色を設定。
	
	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
	createFormatKey = commons.formatkeyCreator(doc)	
	sheet[VARS.daterow, VARS.datecolumn:endedgecolumn].setPropertyValue("NumberFormat", createFormatKey('D'))  
	sheet[VARS.daterow-1:VARS.emptyrow, VARS.datecolumn:endedgecolumn].setPropertyValue("HoriJustify", CENTER)  
	
	
	
	
# 	datevalues = [i for i in range(todayvalue, todayvalue+daycount)]
# 	sheet[VARS.daterow:VARS.daterow+2, VARS.datecolumn:endcolumnedge].setDataArray((datevalues,)*2)
# 	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
# 	createFormatKey = commons.formatkeyCreator(doc)	
# 	sheet[VARS.daterow, VARS.datecolumn:endcolumnedge].setPropertyValue("NumberFormat", createFormatKey('D'))  
# 	sheet[VARS.daterow+1, VARS.datecolumn:endcolumnedge].setPropertyValue("NumberFormat", createFormatKey('AAA'))  
# 	sheet[VARS.daterow-1:VARS.emptyrow, VARS.datecolumn:endcolumnedge].setPropertyValue("HoriJustify", CENTER)  		
# 	
# 	
# 	
# 	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。		
# 	weekdayval = int(functionaccess.callFunction("WEEKDAY", (todayvalue,)))  # 今日の曜日の番号を取得。日=1。
# 	weekdaydic = {1: "日", 2: "月", 3: "火", 4: "水", 5: "木", 6: "金", 7: "土"}
# 	
# 	
# 	templatedatarows = []
	
	
	
	
	
	
	
# 	todaycolumn = VARS.datecolumn + diff  # 移動前の今日の日付列。
# 	dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
# 	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
# 	controller = doc.getCurrentController()  # コントローラの取得。
# 	docframe = controller.getFrame()
# 	controller.select(sheet[VARS.daterow-1:VARS.emptyrow, todaycolumn:VARS.firstemptycolumn])  # 今日以降の列を選択。
# 	dispatcher.executeDispatch(docframe, ".uno:Cut", "", 0, ())  # 日付行を含めてカット。		
# 	controller.select(sheet[VARS.daterow-1, VARS.datecolumn])  # 日付列の先頭セルを選択。
# 	dispatcher.executeDispatch(docframe, ".uno:Paste", "", 0, ())  # ペースト。

	
# 	daycount = 31  # シートに表示する日数。
# 	sheet[VARS.daterow-1:VARS.daterow+2, VARS.datecolumn:VARS.settingstartcolumn].clearContents(511)  # 日付行の内容を削除。
# 	if VARS.datecolumn+daycount>VARS.settingstartcolumn:
# 		daycount = VARS.settingstartcolumn - VARS.datecolumn
# 	endcolumnedge = VARS.datecolumn + daycount
# 	datevalues = [i for i in range(todayvalue, todayvalue+daycount)]
# 	sheet[VARS.daterow:VARS.daterow+2, VARS.datecolumn:endcolumnedge].setDataArray((datevalues,)*2)
# 	createFormatKey = commons.formatkeyCreator(doc)	
# 	sheet[VARS.daterow, VARS.datecolumn:endcolumnedge].setPropertyValue("NumberFormat", createFormatKey('D'))  
# 	sheet[VARS.daterow+1, VARS.datecolumn:endcolumnedge].setPropertyValue("NumberFormat", createFormatKey('AAA'))  
# 	sheet[VARS.daterow-1:VARS.emptyrow, VARS.datecolumn:endcolumnedge].setPropertyValue("HoriJustify", CENTER)  		

	
	
# 	deccolumn = endcolumnedge - diff
# 	for i in range(deccolumn, endcolumnedge):
# 		weekday = sheet[VARS.daterow+1, i].getString()
		
		
			
		
		
	
	
# 	if diff:  # 最初の日付が過去の時。
# 		sheet.insertCells(sheet[0, VARS.emptycolumn].getRangeAddress(), insert_columns)  # 空列を挿入。	
# 		sheet.removeRange(sheet[0, VARS.datecolumn:VARS.datecolumn+diff].getRangeAddress(), delete_columns)
# def setDates(xscriptcontext, datevalue):  # sheet:経過シート、cell: 日付開始セル、dateserial: 日付開始日のシリアル値。
# 	sheet = VARS.sheet
# 	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
# 	createFormatKey = commons.formatkeyCreator(doc)	
# 	colors = commons.COLORS
# 	holidays = commons.HOLIDAYS
# 	daycount = 31  # シートに入力する日数。
# 	sheet[VARS.daterow-1:VARS.daterow+2, VARS.datecolumn:VARS.settingstartcolumn].clearContents(511)  # 日付行の内容を削除。
# 	if VARS.datecolumn+daycount>VARS.settingstartcolumn:
# 		daycount = VARS.settingstartcolumn - VARS.datecolumn
# 	endcolumnedge = VARS.datecolumn + daycount
# 	datevalues = [i for i in range(datevalue, datevalue+daycount)]
# 	sheet[VARS.daterow:VARS.daterow+2, VARS.datecolumn:endcolumnedge].setDataArray((datevalues,)*2)
# 	sheet[VARS.daterow, VARS.datecolumn:endcolumnedge].setPropertyValue("NumberFormat", createFormatKey('D'))  
# 	sheet[VARS.daterow+1, VARS.datecolumn:endcolumnedge].setPropertyValue("NumberFormat", createFormatKey('AAA'))  
# 	sheet[VARS.daterow-1:VARS.emptyrow, VARS.datecolumn:endcolumnedge].setPropertyValue("HoriJustify", CENTER)  
# 	
# 	
# 	
# 	
# 	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
# 	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
# 	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。		
# 	weekdayval = int(functionaccess.callFunction("WEEKDAY", (datevalue,)))
# 	
# 	
# 	
# 	setRangesProperty(doc, range(, endcolumnedge, 7), ("CharColor", colors["red3"]))  # 日曜日の文字色を設定。



# 	sheetcellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
# 	sheetcellranges.addRangeAddresses([sheet[VARS.daterow:VARS.daterow+2, i].getRangeAddress() for i in columnindexes], False)  # セル範囲コレクションを取得。
# 	if len(sheetcellranges):  # sheetcellrangesに要素がないときはsetPropertyValue()でエラーになるので要素の有無を確認する。
# 		sheetcellranges.setPropertyValue(*prop)  # セル範囲コレクションのプロパティを変更。

	
	
# 	celladdress = cell.getCellAddress()  # 経過シートの日付の開始セルのセルアドレスを取得。
# 	r, c = celladdress.Row, celladdress.Column
# 	sheet[:r+1, c:].clearContents(511)  # 開始列より右の日付行の内容を削除。
# 	endcolumn = c + daycount + 1
# 	endcolumn = endcolumn if endcolumn<1024 else 1023  # 列インデックスの上限1023。
# 	sheet[r, c:endcolumn].setDataArray(([i for i in range(datevalue, datevalue+daycount+1)],))  # 日時シリアル値を経過シートに入力。
# 	sheet[r, c:endcolumn].setPropertyValue("NumberFormat", createFormatKey('YYYY/M/D'))  # 日時シリアルから年月日の取得のため一時的に2018/5/4の形式に変換する。
# 	y, m, d = map(int, sheet[r, c].getString().split("/"))  # 年、月、日を整数型で取得。
# 	weekday, days = calendar.monthrange(y, m)  # 月曜日が曜日番号0。1日の曜日と一月の日数のタプルが返る。
# 	weekday = (weekday+(d-1)%7) % 7  # dの曜日番号を取得。1日からの7の余りと1日の余りを加えた7の余りがdの曜日番号。
# 	n = 6  # 日曜日の曜日番号。
# 	sunsset = set(range(c+(n-weekday)%7, endcolumn, 7))  # 日曜日の列インデックスの集合。祝日と重ならないようにあとで使用する。
# 	setRangesProperty(doc, sheet, r, sunsset, ("CharColor", colors["red3"]))  # 日曜日の文字色を設定。
# 	n = 5  # 土曜日の曜日番号。
# 	setRangesProperty(doc, sheet, r, range(c+(n-weekday)%7, endcolumn, 7), ("CharColor", colors["skyblue"]))  # 土曜日の文字色を設定。	
# 	holidayset = set()  # 祝日の列インデックスを入れる集合。
# 	days = days - d + 1  # 翌月1日までの日数を取得。
# 	mr = r - 1  # 月を代入する行のインデックス。
# 	mc = c  # 1日を表示する列のインデックス。最初の月のみ開始日になる。
# 	if y in holidays:  # 祝日一覧のキーがある時。
# 		holidayset.update(mc+i-d for i in holidays[y][m-1] if i>=d)  # 祝日の列インデックスの集合を取得。
# 	while True:
# 		sheet[mr, mc].setString("{}月".format(m))  # 月を入力。
# 		mc += days  # 次月1日の列に進める。
# 		if mc<endcolumn:  # 日時シリアル値が入力されている列の時。
# 			ymd = sheet[r, mc].getString()  # 1日の年/月/日を取得。
# 			y, m = map(int, ymd.split("/")[:2])  # 年と月を取得。
# 			if y in holidays:  # 祝日一覧のキーがある時。
# 				holidayset.update(mc+i-1 for i in holidays[y][m-1] if mc+i-1<endcolumn)  # 祝日の列インデックスの集合を取得。
# 			weekday, days = calendar.monthrange(y, m)  # 1日の曜日と月の日数を取得。
# 		else:
# 			break	
# 	holidayset.difference_update(sunsset)  # 日曜日と重なっている祝日を除く。
# 	setRangesProperty(doc, sheet, r, holidayset, ("CellBackColor", colors["red3"]))  # 祝日の背景色を設定。	
# 	sheet[r, c:endcolumn].setPropertyValues(("NumberFormat", "HoriJustify"), (createFormatKey('D'), CENTER))  # 経過シートの日付の書式を設定。	
# 
# 
# def setRangesProperty(doc, columnindexes, prop):  # r行のcolumnindexesの列のプロパティを変更。prop: プロパティ名とその値のリスト。
# 	sheetcellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
# 	sheetcellranges.addRangeAddresses([VARS.sheet[VARS.daterow:VARS.daterow+2, i].getRangeAddress() for i in columnindexes], False)  # セル範囲コレクションを取得。
# 	if len(sheetcellranges):  # sheetcellrangesに要素がないときはsetPropertyValue()でエラーになるので要素の有無を確認する。
# 		sheetcellranges.setPropertyValue(*prop)  # セル範囲コレクションのプロパティを変更。
def notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, contextmenuname):			
	if contextmenuname=="cell":  # セルのとき
		selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
		del contextmenu[:]  # contextmenu.clear()は不可。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
			addMenuentry("ActionTrigger", {"Text": "To blue", "CommandURL": baseurl.format("entry1")})  # listeners.pyの関数名を指定する。
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
		target.setPropertyValue("CellBackColor", colors["ao"])  # 背景を青色にする。
	elif entrynum==2:
		target.setPropertyValue("CellBackColor", colors["aka"]) 
