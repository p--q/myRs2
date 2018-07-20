#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import commons
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.sheet.CellDeleteMode import COLUMNS as delete_columns  # enum
from com.sun.star.sheet.CellInsertMode import COLUMNS as insert_columns  # enum
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.table.CellHoriJustify import CENTER  # enum
from com.sun.star.beans import PropertyValue # Struct
from com.sun.star.sheet import ConditionOperator2  # 定数
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
		daycount = VARS.templatestartcolumn - VARS.datecolumn  # daycountの上限はテンプレート列までにする。
	todaycolumn = VARS.datecolumn + diff # 移動前の今日の日付列。	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
	if todaycolumn<VARS.firstemptycolumn:  # 今日の日付列が表示されている範囲内にある時。今日の日付を先頭に移動させる。
		dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
		controller = doc.getCurrentController()  # コントローラの取得。
		controller.select(sheet[VARS.daterow-1:VARS.emptyrow, todaycolumn:VARS.templatestartcolumn])  #  移動前の今日の日付列以降テンプレート列左までを選択。
		docframe = controller.getFrame()
		dispatcher.executeDispatch(docframe, ".uno:Cut", "", 0, ())  # カット。	
		controller.select(sheet[VARS.daterow-1, VARS.datecolumn])  # ペーストする左上セルを選択。
		dispatcher.executeDispatch(docframe, ".uno:Paste", "", 0, ())  # ペースト。	
	weekday = int(functionaccess.callFunction("WEEKDAY", (todayvalue,)))   # 今日の曜日番号を取得。日=1。		
	weekdays = "日", "月", "火", "水", "木", "金", "土"  # シートでは日=1であることに注意。
	endedgecolumn = VARS.datecolumn + daycount
	datedatarows = [[i for i in range(todayvalue, todayvalue+daycount)],\
					[weekdays[i%7] for i in range(weekday-1, weekday-1+daycount)]]  # 日付行と曜日行を作成。
	scheduledatarows = sheet[VARS.daterow+2:VARS.emptyrow, VARS.datecolumn:endedgecolumn].getDataArray()	
	sheet[VARS.daterow-1:VARS.daterow+2, VARS.datecolumn:endedgecolumn].clearContents(511)  # 日付行を削除。

	templates = sheet[VARS.daterow-1:VARS.emptyrow, VARS.templatestartcolumn+1:VARS.templateendcolumnedge].getDataArray()  # テンプレートの値を日付行の上行から取得。
	templatecolumnlists = [[] for dummy in range(8)]
	for i, yobi in enumerate(templates[2], start=VARS.templatestartcolumn+1):  # 曜日行を列インデックスと共にイテレート。
		n = weekdays.index(yobi)+1 if yobi in weekdays else 0  # 曜日番号を取得。インデックス0は曜日以外のとき。
		templatecolumnlists[n].append(i)  # 日=1から始まる曜日番号をインデックスとしてその曜日のテンプレートの列インデックスを取得。

	y, m, d = [int(functionaccess.callFunction(i, (todayvalue,))) for i in ("YEAR", "MONTH", "DAY")]
	nextmdatevalue = todayvalue
	firstmidx = 0  # datedatarows[0]の中の月の初日の要素のインデックス。
	holidays = set()  # 祝日の列インデックスを入れる集合。

	ms, ds = {}, {}
	for i in templatecolumnlists[0]:  # 曜日指定のないテンプレートの列インデックスについて。月日か日のみが指定されているはず。
		tm = templates[0][i]  # テンプレートの月を取得。
		td = templates[1][i]
		if tm:  # 月の指定がある時。
			ms.setdefault(tm, []).append(i)
		elif td:  # 日の指定のみの時。
			ds[td] = i
	excludes = []  # 処理済列インデックスのリスト。
	while True:
		firstdaycolumn = VARS.datecolumn + firstmidx  # 月の初日の列インデックス。
		if not firstdaycolumn<endedgecolumn:  # 表示最終列を越えている時ループを抜ける。
			if m in ms:  # テンプレートに指定のある月の時。
				for i in ms[m]:  # テンプレートの列インデックスをイテレート。
					td = templates[1][i]  # 指定日を取得。
					if td<endedgecolumn-firstdaycolumn:  # 終了日より前の時。
						queryTemplateColumn(firstdaycolumn, td, i, scheduledatarows, templates, excludes)
			for td in ds.keys():
				if td<endedgecolumn-firstdaycolumn:  # 終了日より前の時。
					queryTemplateColumn(firstdaycolumn, td, ds[td], scheduledatarows, templates, excludes)			
			break		
		if m in ms:  # テンプレートに指定のある月の時。
			for i in ms[m]:  # テンプレートの列インデックスをイテレート。
				td = templates[1][i]  # 指定日を取得。
				if d-1<td:  # 開始日以降の時。
					queryTemplateColumn(firstdaycolumn, td, i, scheduledatarows, templates, excludes)
		for td in ds.keys():
			if d<=td:
				queryTemplateColumn(firstdaycolumn, td, ds[td], scheduledatarows, templates, excludes)
		if y in commons.HOLIDAYS:  # 年が祝日一覧のキーにある時。
			holidays.update(map(lambda x: firstdaycolumn+x, commons.HOLIDAYS[y][m-1]))  # 祝日の日付の列インデックスを取得。
		sheet[VARS.daterow-1, firstdaycolumn].setString("{}月".format(m))  # 月を代入。
		nextmdatevalue = int(functionaccess.callFunction("EOMONTH", (nextmdatevalue, 0))) + 1  # 翌月1日のシリアル値を取得。
		firstmidx = datedatarows[0].index(nextmdatevalue)  # 1日のインデックスを取得。
		if m>11:  # 12月の次は1月にする。
			m = 1  # 次月を取得。
			y += 1  # 年も更新する。
		else:
			m += 1  # 次月を取得。
		d = 1  # 日付を更新。
			
			
	for n in range(1, 8):  # 曜日番号をイテレート。
		templatecolumns = templatecolumnlists[n]  # 曜日のテンプレートの列インデックスのリストを取得。
		for templatecolumn in templatecolumns:
			range(VARS.datecolumn+(n-weekday)%7, endedgecolumn, 7)  			

		

		
		
			
	
	
	n = 7  # 土曜日の曜日番号。
	columnindexes = range(VARS.datecolumn+(n-weekday)%7, endedgecolumn, 7)   # 土曜日の列インデックスを取得。			
	setRangesProperty(doc, columnindexes, ("CharColor", commons.COLORS["skyblue"]))  # 土曜日の文字色を設定。	
	n = 1  # 日曜日の曜日番号。
	columnindexes = range(VARS.datecolumn+(n-weekday)%7, endedgecolumn, 7)   # 日曜日の列インデックスを取得。
	setRangesProperty(doc, columnindexes, ("CharColor", commons.COLORS["red3"]))  # 日曜日の文字色を設定。				
	holidays.difference_update(columnindexes)  # 日曜日と重なっている祝日を除く。	
	holidays = filter(lambda x: x<endedgecolumn, holidays)  # 上限を設定。
	setRangesProperty(doc, holidays, ("CellBackColor", commons.COLORS["red3"]))  # 祝日の背景色を設定。			
			
	
	

		
		
	
	

	
	
	
	
	for i in range(7):  # 最初の7列について。
		yobi = weekdays[i%7]  # 曜日文字を取得。
		
		
		
	for i in range(len(datarows[0])):  # インデックスがVARS.firstemptycolumn-VARS.datecolumn+diff以降は2行目以降の要素はない。
		yobi = datarows[1][i]  # 曜日の文字列を取得。
		for k in range(len(templates[0])):  # テンプレートの列インデックスをイテレート。
			if yobi in templates[1][k]:  # 曜日が一致する時。
				w = templates[0][k]  # 週数行の値を取得。
				if w.endswith("w"):  # wで終わる時は週番号。
					d = int(functionaccess.callFunction("DAY", (datarows[0][i],)))  # 月の何日目か取得。
					if int(w[:-1])!=-(-d//7):  # 週番号が一致しない時。-(-d//7)切り上げ。
						continue  # 週番号が一致しない時は次のループに行く。
				elif w.endswith("d"):  # dで終わる時は月のd日目。
					if int(w[:-1])!=int(functionaccess.callFunction("DAY", (datarows[0][i],))):  # d日目でない時。
						continue  # d日目ではない時は次のループに行く。
				elif isinstance(w, float):  # float型の時は日付シリアル値。
					if datarows[0][i]!=int(w):  # 日付が一致しない時は次の列に行く。
						continue
				for j in range(2, len(datarows)):  # 行インデックスを時間枠の先頭行からイテレート。
					if i<len(datarows[j]):  # 行にインデックスがある時。
						if datarows[j][i] in ("", "/", "x"):  # テンプレートを優先する文字列の時。
							datarows[j][i] = templates[j][k]  # テンプレートの文字列を採用。
					else:  # 行に要素がないインデックスの時は要素を追加する。
						datarows[j].append(templates[j][k])	
				break  # datarowsの次の列に行く。		
		
		
		
# 		newfirstdatecolumn = VARS.firstemptycolumn - diff
# 		newfistdatevalue = todayvalue + VARS.firstemptycolumn - todaycolumn  # 追加する最初の日付のシリアル値。
# 	else:  # 今日の日付列が表示されている範囲内にない時は今日の日付から始める。
# 		newfirstdatecolumn = VARS.datecolumn
# 		newfistdatevalue = todayvalue	
		
		

	

	
	
	

	
	
	
	
	
# 	todaycolumn = VARS.datecolumn + diff  # 移動前の今日の日付列。
# 	if todaycolumn<VARS.firstemptycolumn:
# 		datarows.extend(sheet[VARS.daterow+2:VARS.emptyrow, todaycolumn:VARS.firstemptycolumn].getDataArray())  # 今日の日付の列以降の行を取得。
# 	else:
# 		datarows.extend([] for dummy in range(VARS.emptyrow-(VARS.daterow+2)))  # コピーすべき列がない時は空行を追加する。


		
		
		
# 	for i in range(len(datarows[0])):  # インデックスがVARS.firstemptycolumn-VARS.datecolumn+diff以降は2行目以降の要素はない。
# 		yobi = datarows[1][i]  # 曜日の文字列を取得。
# 		for k in range(len(templates[0])):  # テンプレートの列インデックスをイテレート。
# 			if yobi in templates[1][k]:  # 曜日が一致する時。
# 				w = templates[0][k]  # 週数行の値を取得。
# 				if w.endswith("w"):  # wで終わる時は週番号。
# 					d = int(functionaccess.callFunction("DAY", (datarows[0][i],)))  # 月の何日目か取得。
# 					if int(w[:-1])!=-(-d//7):  # 週番号が一致しない時。-(-d//7)切り上げ。
# 						continue  # 週番号が一致しない時は次のループに行く。
# 				elif w.endswith("d"):  # dで終わる時は月のd日目。
# 					if int(w[:-1])!=int(functionaccess.callFunction("DAY", (datarows[0][i],))):  # d日目でない時。
# 						continue  # d日目ではない時は次のループに行く。
# 				elif isinstance(w, float):  # float型の時は日付シリアル値。
# 					if datarows[0][i]!=int(w):  # 日付が一致しない時は次の列に行く。
# 						continue
# 				for j in range(2, len(datarows)):  # 行インデックスを時間枠の先頭行からイテレート。
# 					if i<len(datarows[j]):  # 行にインデックスがある時。
# 						if datarows[j][i] in ("", "/", "x"):  # テンプレートを優先する文字列の時。
# 							datarows[j][i] = templates[j][k]  # テンプレートの文字列を採用。
# 					else:  # 行に要素がないインデックスの時は要素を追加する。
# 						datarows[j].append(templates[j][k])	
# 				break  # datarowsの次の列に行く。
			
			
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
	endedgecolumn = VARS.datecolumn + daycount					
	sheet[VARS.daterow-1:VARS.emptyrow, VARS.datecolumn:endedgecolumn].clearContents(511)  # 内容を削除。
	sheet[VARS.daterow:VARS.emptyrow, VARS.datecolumn:endedgecolumn].setDataArray(datarows)
	createFormatKey = commons.formatkeyCreator(doc)	
	sheet[VARS.daterow, VARS.datecolumn:endedgecolumn].setPropertyValue("NumberFormat", createFormatKey('D'))  
	sheet[VARS.daterow-1:VARS.emptyrow, VARS.datecolumn:endedgecolumn].setPropertyValue("HoriJustify", CENTER)  
	n = 1  # 日曜日の曜日番号。
	sunsset = set(range(VARS.datecolumn+(n-weekday)%7, endedgecolumn, 7))  # 日曜日の列インデックスの集合。祝日と重ならないようにあとで使用する。
	setRangesProperty(doc, sunsset, ("CharColor", commons.COLORS["red3"]))  # 日曜日の文字色を設定。	
	n = 7  # 土曜日の曜日番号。
	setRangesProperty(doc, range(VARS.datecolumn+(n-weekday)%7, endedgecolumn, 7), ("CharColor", commons.COLORS["skyblue"]))  # 土曜日の文字色を設定。	
	holidayset = set()  # 祝日の列インデックスを入れる集合。
	y, m, d = [int(functionaccess.callFunction(i, (todayvalue,))) for i in ("YEAR", "MONTH", "DAY")]
	sheet[VARS.daterow-1, VARS.datecolumn].setString("{}月".format(m))
	if y in commons.HOLIDAYS:  # 祝日一覧のキーがある時。
		holidayset.update(VARS.datecolumn+i-d for i in commons.HOLIDAYS[y][m-1] if i>=d)  # 祝日の列インデックスの集合を取得。		
	nextmdatevalue = todayvalue
	while True:
		nextmdatevalue = int(functionaccess.callFunction("EOMONTH", (nextmdatevalue, 0))) + 1  # 翌月1日のシリアル値を取得。
		nextmcolumn = VARS.datecolumn + nextmdatevalue - todayvalue
		if nextmcolumn>endedgecolumn-1:
			break
		y, m = [int(functionaccess.callFunction(i, (nextmdatevalue,))) for i in ("YEAR", "MONTH")]
		sheet[VARS.daterow-1, nextmcolumn].setString("{}月".format(m))
		if y in commons.HOLIDAYS:  # 祝日一覧のキーがある時。
			holidaycolumns = (nextmcolumn+i-1 for i in commons.HOLIDAYS[y][m-1])
			holidayset.update(i for i in holidaycolumns if i<endedgecolumn)  # 祝日の列インデックスの集合を取得。		
	holidayset.difference_update(sunsset)  # 日曜日と重なっている祝日を除く。
	setRangesProperty(doc, holidayset, ("CellBackColor", commons.COLORS["red3"]))  # 祝日の背景色を設定。	

# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
# 	
# 	ranges = sheet[VARS.daterow+2:VARS.emptyrow, VARS.datecolumn:VARS.firstemptycolumn],\
# 			sheet[VARS.daterow+2:VARS.emptyrow, VARS.templatestartcolumn:VARS.templateendcolumnedge]
# 	datarange = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
# 	datarange.addRangeAddresses((i.getRangeAddress() for i in ranges), False)	
# 	searchdescriptor = sheet.createSearchDescriptor()
# 	
# 
# 	searchdescriptor.setSearchString("x")  # 戻り値はない。
# 	cellranges = datarange.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
# 	if cellranges:
# 		cellranges.setPropertyValue("CellBackColor", commons.COLORS["gray7"])
# 	searchdescriptor.setSearchString("/")  # 戻り値はない。
# 	cellranges = datarange.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
# 	if cellranges:
# 		cellranges.setPropertyValue("CellBackColor", commons.COLORS["silver"])	
	
	
	
# 	ranges = sheet[VARS.daterow+2:VARS.emptyrow, VARS.datecolumn:VARS.firstemptycolumn],\
# 			sheet[VARS.daterow+2:VARS.emptyrow, VARS.templatestartcolumn:VARS.templateendcolumnedge]
# 	sheetcellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
# 	sheetcellranges.addRangeAddresses((i.getRangeAddress() for i in ranges), False)
# 	conditionalformat = sheetcellranges.getPropertyValue("ConditionalFormat")
# 	conditionalformat.clear()
# 	stylenames = "gray7", "silver", "magenta3"
# 	stylefamilies = doc.getStyleFamilies()
# 	cellstyles = stylefamilies["CellStyles"]
# 	for stylename in stylenames:
# 		if not stylename in cellstyles:
# 			newstyle = doc.createInstance("com.sun.star.style.CellStyle")
# 			cellstyles[stylename] = newstyle
# 			newstyle.setPropertyValue("CellBackColor", commons.COLORS[stylename])
# 			
# 			
# 			
# 	propertyvalues = PropertyValue(Name="Operator", Value=ConditionOperator2.EQUAL),\
# 					PropertyValue(Name="Formula1", Value="x"),\
# 					PropertyValue(Name="StyleName", Value="gray7")
# 	conditionalformat.addNew(propertyvalues)
# 
# 	propertyvalues = PropertyValue(Name="Operator", Value=ConditionOperator2.EQUAL),\
# 					PropertyValue(Name="Formula1", Value="/"),\
# 					PropertyValue(Name="StyleName", Value="silver")
# 	conditionalformat.addNew(propertyvalues)

# 	propertyvalues = PropertyValue(Name="Operator", Value=ConditionOperator2.EQUAL),\
# 					PropertyValue(Name="Formula1", Value="x"),\
# 					PropertyValue(Name="StyleName", Value="magenta3")
# 	conditionalformat.addNew(propertyvalues)

def queryTemplateColumn(firstdaycolumn, td, i, scheduledatarows, templates, excludes):
	c = firstdaycolumn + td - 1  # 列インデックスを取得。
	cellranges = VARS.sheet[VARS.daterow+2:VARS.emptyrow, c].queryRowDifferences(VARS.sheet[VARS.daterow, i].getCellAddress())  # テンプレートの列と異なる行のセル範囲を取得。
	j = c - VARS.datecolumn
	for cell in cellranges.getCells():
		k = cell.getCellAddress().Row - (VARS.daterow+2) 
		if scheduledatarows[k][j] in ("", "/", "x"):  # テンプレートを優先する文字列の時。
			scheduledatarows[k][j] = templates[k+3, i-(VARS.templatestartcolumn+1)]  # テンプレートを使う。
	excludes.append(c)	
def setRangesProperty(doc, columnindexes, prop):  # r行のcolumnindexesの列のプロパティを変更。prop: プロパティ名とその値のリスト。
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
	cellranges.addRangeAddresses((VARS.sheet[VARS.daterow:VARS.daterow+2, i].getRangeAddress() for i in columnindexes), False)  # セル範囲コレクションを取得。
	if len(cellranges):  # sheetcellrangesに要素がないときはsetPropertyValue()でエラーになるので要素の有無を確認する。
		cellranges.setPropertyValue(*prop)  # セル範囲コレクションのプロパティを変更。	
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
