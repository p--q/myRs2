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
		self.monthrow = 1  # 月行。
		self.dayrow = 2  # 日行。
		self.weekdayrow = 3
		self.datarow = 4  # データ開始行。
		self.datacolumn = 1  # データ開始列。
	def setSheet(self, sheet):
		self.sheet = sheet
		cellranges = sheet[:, 0].queryContentCells(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME)  # 先頭列の文字列か数値が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 先頭列の最終行インデックス+1を取得。		
		cellranges = sheet[self.dayrow+1, :].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # 曜日行の文字列か数値が入っているセルに限定して抽出。
		rangeaddresses = cellranges.getRangeAddresses()	
		self.firstemptycolumn = rangeaddresses[0].EndColumn + 1  # 日付行の区切れの列インデックスを取得。
		self.templatestartcolumn = rangeaddresses[1].StartColumn + 1  # テンプレートの設定開始列。
		self.templateendcolumnedge = rangeaddresses[1].EndColumn + 1  # テンプレートの終了列右。
VARS = Schedule()
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	VARS.setSheet(sheet)
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。
	sheet["A1"].setString("ﾘｽﾄに戻る")
	sheet["AF1"].setString("COPY")
	sheet["AK1"].setString("強有効")
	firstdatevalue = int(sheet[VARS.dayrow, VARS.datacolumn].getValue())  # 最初の日付のシリアル値を整数で取得。
	todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。	
	diff = todayvalue - firstdatevalue
	if not diff>0:  # 先頭日付が今日より前でない時はここで終わる。
		return
	daycount = 31  # シートに表示する日数。
	if VARS.datacolumn+daycount>VARS.templatestartcolumn:
		daycount = VARS.templatestartcolumn - VARS.datacolumn  # daycountの上限はテンプレート列までにする。
	todaycolumn = VARS.datacolumn + diff # 移動前の今日の日付列。	
	if todaycolumn<VARS.firstemptycolumn:  # 今日の日付列が表示されている範囲内にある時。今日の日付を先頭に移動させる。
		controller = doc.getCurrentController()  # コントローラの取得。
		docframe = controller.getFrame()
		dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
		controller.select(sheet[VARS.monthrow:VARS.emptyrow, todaycolumn:VARS.templatestartcolumn])  #  移動前の今日の日付列以降テンプレート列左までを選択。
		dispatcher.executeDispatch(docframe, ".uno:Cut", "", 0, ())  # カット。	
		controller.select(sheet[VARS.monthrow, VARS.datacolumn])  # ペーストする左上セルを選択。
		dispatcher.executeDispatch(docframe, ".uno:Paste", "", 0, ())  # ペースト。	
	weekday = int(functionaccess.callFunction("WEEKDAY", (todayvalue,)))   # 今日の曜日番号を取得。日=1。		
	weekdays = "日", "月", "火", "水", "木", "金", "土"  # シートでは日=1であることに注意。
	endedgecolumn = VARS.datacolumn + daycount  # データの右端列の右。
	datarows = [["" for dummy in range(daycount)],\
			[i for i in range(todayvalue, todayvalue+daycount)],\
			[weekdays[i%7] for i in range(weekday-1, weekday-1+daycount)]]  # 月行、日行と曜日行を作成。
	sheet[VARS.monthrow:VARS.datarow, VARS.datacolumn:endedgecolumn].clearContents(511)  # シートの日付行を削除。
	datarows.extend(sheet[VARS.datarow:VARS.emptyrow, VARS.datacolumn:endedgecolumn].getDataArray())  # シートのデータ部分を取得。	

	templates = sheet[VARS.monthrow:VARS.emptyrow, VARS.templatestartcolumn:VARS.templateendcolumnedge].getDataArray()  # テンプレートの値を月行から取得。
	templatecolumnlists = [[] for dummy in range(8)]  # 曜日別の列インデックスのリストのリスト。
	ms, ds = {}, {}  # 月指定がある時は月をキー、日指定の時は日をキー、値はテンプレートデータ行のインデックス、の辞書。
	for c, yobi in enumerate(templates[2], start=VARS.templatestartcolumn):  # 曜日行を列インデックスと共にイテレート。
		if yobi in weekdays:  # 曜日が一致している時。  
			n = weekdays.index(yobi) + 1  # 曜日番号を取得。日=1にしている。
			templatecolumnlists[n].append(c)  # 日=1から始まる曜日番号をインデックスとしてその曜日のテンプレートの列インデックスを取得。
		else:  # 曜日の指定がない時は日指定か月日指定とする。
			i = c - VARS.templatestartcolumn  # 列インデックスをデータ行の相対インデックスに変換。
			tm = templates[0][i]  # テンプレート列の月を取得。
			td = templates[1][i]  # テンプレート列の日を取得。
			if tm:  # 月の指定がある時。
				ms.setdefault(tm, []).append(i)  # 同じ月の列インデックスをリストで取得。
			elif td:  # 日の指定のみの時。
				ds[td] = i  # 日指定の列インデックスを取得。			


	holidays = set()  # 祝日の列インデックスを入れる集合。
	excludes = set()  # 処理済列インデックスの集合。
	y, m, d = [int(functionaccess.callFunction(i, (todayvalue,))) for i in ("YEAR", "MONTH", "DAY")]
	nextmdatevalue = todayvalue  # 次月の初日のシリアル値を入れる変数。
	firstmidx = 0  # datacolumnからの月の初日の相対インデックス。
	firstdaycolumn = VARS.datacolumn  # 今月の初日の列インデックス。
	queryTemplateColumn = createQueryTemplateColumn(datarows, templates, excludes)
	while True:
		if daycount-1<firstmidx:  # 次月の初日が表示最終列を越えている時、月の途中で終わっている。
			rdays = endedgecolumn - firstdaycolumn  # 月の最終日。firstdaycolumnはまだ次月のものに更新していない。
			if m in ms:  # テンプレートに指定のある月の時。
				for i in ms[m]:  # テンプレートの相対列インデックスをイテレート。
					td = templates[1][i]  # 指定日を取得。
					if td<rdays:  # 終了日より前の時。
						queryTemplateColumn(firstdaycolumn, td, i)
			for td in ds.keys():
				if td<rdays:  # 終了日より前の時。
					queryTemplateColumn(firstdaycolumn, td, ds[td])			
			break  # ループを抜ける。
		firstdaycolumn = VARS.datacolumn + firstmidx  # 今月の初日の列インデックスを取得。
		if m in ms:  # テンプレートに指定のある月の時。
			for i in ms[m]:  # テンプレートの相対列インデックスをイテレート。
				td = templates[1][i]  # 指定日を取得。
				if d-1<td:  # 開始日以降の時。
					queryTemplateColumn(firstdaycolumn, td, i)
		for td in ds.keys():
			if d-1<td:
				queryTemplateColumn(firstdaycolumn, td, ds[td])
		if y in commons.HOLIDAYS:  # 年が祝日一覧のキーにある時。
			holidays.update(map(lambda x: firstdaycolumn+x, commons.HOLIDAYS[y][m-1]))  # 祝日の日付の列インデックスを取得。
		datarows[0][firstmidx] = "{}月".format(m)  # 月を代入。
		nextmdatevalue = int(functionaccess.callFunction("EOMONTH", (nextmdatevalue, 0))) + 1  # 翌月1日のシリアル値を取得。
		firstmidx = nextmdatevalue - todayvalue  # 次月の初日のdatacolumnからの相対インデックスを取得。
		if m>11:  # 12月の次は1月にする。
			m = 1  # 次月を取得。
			y += 1  # 年も更新する。
		else:
			m += 1  # 次月を取得。
		d = 1  # 日付を更新。
			
	queryWeekdayColumn = createQueryWeekdayColumn(datarows, templates)		
	for n in range(1, 8):  # 曜日番号をn=1からイテレート。
		templatecolumns = templatecolumnlists[n]  # 同じ曜日のテンプレートの列インデックスのリストを取得。
		for c in range(VARS.datacolumn+(n-weekday)%7, endedgecolumn, 7):  # 同じ曜日の列インデックスを取得。
			if not c in excludes:  # 処理済の列インデックス以外の時。
				if templatecolumns>1:  # 複数列がある時は週番号指定(2wなど)列を含む。
					j = c - VARS.datacolumn  # 相対インデックスを取得。
					for tc in templatecolumns:
						w = templates[1][tc-VARS.templatestartcolumn]  # 週数の行の値を取得。
						if w.endswith("w"):  # wで終わる時は週番号。		
							d = int(functionaccess.callFunction("DAY", (datarows[1][j],)))  # 月の何日目か取得。		
							if int(w[:-1])==-(-d//7):  # 週番号が一致する時。-(-d//7)切り上げ。	
								queryWeekdayColumn(c, tc)	
						elif not w:  # 空セルのときは曜日のみ指定。
							queryWeekdayColumn(c, tc)				
				else:
					queryWeekdayColumn(c, tc)	
	n = 7  # 土曜日の曜日番号。
	columnindexes = range(VARS.datacolumn+(n-weekday)%7, endedgecolumn, 7)   # 土曜日の列インデックスを取得。			
	setRangesProperty(doc, columnindexes, ("CharColor", commons.COLORS["skyblue"]))  # 土曜日の文字色を設定。	
	n = 1  # 日曜日の曜日番号。
	columnindexes = range(VARS.datacolumn+(n-weekday)%7, endedgecolumn, 7)   # 日曜日の列インデックスを取得。
	setRangesProperty(doc, columnindexes, ("CharColor", commons.COLORS["red3"]))  # 日曜日の文字色を設定。				
	holidays.difference_update(columnindexes)  # 日曜日と重なっている祝日を除く。	
	holidays = filter(lambda x: x<endedgecolumn, holidays)  # 上限を設定。
	setRangesProperty(doc, holidays, ("CellBackColor", commons.COLORS["red3"]))  # 祝日の背景色を設定。
	for c in holidays:
		sheet[VARS.daterow, c].setDataArray(("x",)*(VARS.emptyrow-VARS.datarow))	
	createFormatKey = commons.formatkeyCreator(doc)	
	sheet[VARS.dayrow, VARS.datacolumn:endedgecolumn].setPropertyValue("NumberFormat", createFormatKey('D'))  
	
	sheet[VARS.monthrow:VARS.emptyrow, VARS.datacolumn:endedgecolumn].setPropertyValue("HoriJustify", CENTER)  			
	
	ranges = sheet[VARS.monthrow:VARS.emptyrow, VARS.datacolumn:endedgecolumn],\
			sheet[VARS.datarow:VARS.emptyrow, VARS.templatestartcolumn:VARS.templateendcolumnedge]
	datarange = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
	datarange.addRangeAddresses((i.getRangeAddress() for i in ranges), False)		
	datarange.setPropertyValue("CellBackColor", -1)  # 背景色をクリア。
	searchdescriptor = sheet.createSearchDescriptor()
	searchdescriptor.setSearchString("x")  # 戻り値はない。
	cellranges = datarange.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
	if cellranges:
		cellranges.setPropertyValue("CellBackColor", commons.COLORS["gray7"])
	searchdescriptor.setSearchString("/")  # 戻り値はない。
	cellranges = datarange.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
	if cellranges:
		cellranges.setPropertyValue("CellBackColor", commons.COLORS["silver"])	
	searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
	searchdescriptor.setSearchString("[^x/]")  # 戻り値はない。	
	if cellranges:
		cellranges.setPropertyValue("CellBackColor", commons.COLORS["magenta3"])		

		
		
		
# 	for i in range(len(datarows[0])):  # インデックスがVARS.firstemptycolumn-VARS.datacolumn+diff以降は2行目以降の要素はない。
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
		
		
		
# 		newfirstdatecolumn = VARS.firstemptycolumn - diff
# 		newfistdatevalue = todayvalue + VARS.firstemptycolumn - todaycolumn  # 追加する最初の日付のシリアル値。
# 	else:  # 今日の日付列が表示されている範囲内にない時は今日の日付から始める。
# 		newfirstdatecolumn = VARS.datacolumn
# 		newfistdatevalue = todayvalue	
		
		

	

	
	
	

	
	
	
	
	
# 	todaycolumn = VARS.datacolumn + diff  # 移動前の今日の日付列。
# 	if todaycolumn<VARS.firstemptycolumn:
# 		datarows.extend(sheet[VARS.datarow:VARS.emptyrow, todaycolumn:VARS.firstemptycolumn].getDataArray())  # 今日の日付の列以降の行を取得。
# 	else:
# 		datarows.extend([] for dummy in range(VARS.emptyrow-(VARS.datarow)))  # コピーすべき列がない時は空行を追加する。


		
		
		
# 	for i in range(len(datarows[0])):  # インデックスがVARS.firstemptycolumn-VARS.datacolumn+diff以降は2行目以降の要素はない。
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
			
			
# 	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
# 	endedgecolumn = VARS.datacolumn + daycount					
# 	sheet[VARS.monthrow:VARS.emptyrow, VARS.datacolumn:endedgecolumn].clearContents(511)  # 内容を削除。
# 	sheet[VARS.daterow:VARS.emptyrow, VARS.datacolumn:endedgecolumn].setDataArray(datarows)
# 	createFormatKey = commons.formatkeyCreator(doc)	
# 	sheet[VARS.daterow, VARS.datacolumn:endedgecolumn].setPropertyValue("NumberFormat", createFormatKey('D'))  
# 	sheet[VARS.monthrow:VARS.emptyrow, VARS.datacolumn:endedgecolumn].setPropertyValue("HoriJustify", CENTER)  
# 	n = 1  # 日曜日の曜日番号。
# 	sunsset = set(range(VARS.datacolumn+(n-weekday)%7, endedgecolumn, 7))  # 日曜日の列インデックスの集合。祝日と重ならないようにあとで使用する。
# 	setRangesProperty(doc, sunsset, ("CharColor", commons.COLORS["red3"]))  # 日曜日の文字色を設定。	
# 	n = 7  # 土曜日の曜日番号。
# 	setRangesProperty(doc, range(VARS.datacolumn+(n-weekday)%7, endedgecolumn, 7), ("CharColor", commons.COLORS["skyblue"]))  # 土曜日の文字色を設定。	
# 	holidayset = set()  # 祝日の列インデックスを入れる集合。
# 	y, m, d = [int(functionaccess.callFunction(i, (todayvalue,))) for i in ("YEAR", "MONTH", "DAY")]
# 	sheet[VARS.monthrow, VARS.datacolumn].setString("{}月".format(m))
# 	if y in commons.HOLIDAYS:  # 祝日一覧のキーがある時。
# 		holidayset.update(VARS.datacolumn+i-d for i in commons.HOLIDAYS[y][m-1] if i>=d)  # 祝日の列インデックスの集合を取得。		
# 	nextmdatevalue = todayvalue
# 	while True:
# 		nextmdatevalue = int(functionaccess.callFunction("EOMONTH", (nextmdatevalue, 0))) + 1  # 翌月1日のシリアル値を取得。
# 		nextmcolumn = VARS.datacolumn + nextmdatevalue - todayvalue
# 		if nextmcolumn>endedgecolumn-1:
# 			break
# 		y, m = [int(functionaccess.callFunction(i, (nextmdatevalue,))) for i in ("YEAR", "MONTH")]
# 		sheet[VARS.monthrow, nextmcolumn].setString("{}月".format(m))
# 		if y in commons.HOLIDAYS:  # 祝日一覧のキーがある時。
# 			holidaycolumns = (nextmcolumn+i-1 for i in commons.HOLIDAYS[y][m-1])
# 			holidayset.update(i for i in holidaycolumns if i<endedgecolumn)  # 祝日の列インデックスの集合を取得。		
# 	holidayset.difference_update(sunsset)  # 日曜日と重なっている祝日を除く。
# 	setRangesProperty(doc, holidayset, ("CellBackColor", commons.COLORS["red3"]))  # 祝日の背景色を設定。	

# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
# 	
# 	ranges = sheet[VARS.datarow:VARS.emptyrow, VARS.datacolumn:VARS.firstemptycolumn],\
# 			sheet[VARS.datarow:VARS.emptyrow, VARS.templatestartcolumn:VARS.templateendcolumnedge]
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
	
	
	
# 	ranges = sheet[VARS.datarow:VARS.emptyrow, VARS.datacolumn:VARS.firstemptycolumn],\
# 			sheet[VARS.datarow:VARS.emptyrow, VARS.templatestartcolumn:VARS.templateendcolumnedge]
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
def createQueryWeekdayColumn(datarows, templates):
	def queryWeekdayColumn(c, tc):
		j = c - VARS.datacolumn  # 相対インデックスを取得。
		cellranges = VARS.sheet[VARS.datarow:VARS.emptyrow, c].queryRowDifferences(VARS.sheet[VARS.daterow, tc].getCellAddress())  # テンプレートの列と異なる行のセル範囲を取得。
		for cell in cellranges.getCells():
			k = cell.getCellAddress().Row - VARS.monthrow 
			if datarows[k][j] in ("", "/", "x"):  # テンプレートを優先する文字列の時。
				datarows[k][j] = templates[k, tc-VARS.templatestartcolumn]  # テンプレートの値を使う。		
	return queryWeekdayColumn
def createQueryTemplateColumn(datarows, templates, excludes):
	def queryTemplateColumn(firstdaycolumn, td, i):  # 
		c = firstdaycolumn + td - 1  # 列インデックスを取得。
		cellranges = VARS.sheet[VARS.datarow:VARS.emptyrow, c].queryRowDifferences(VARS.sheet[VARS.daterow, VARS.templatestartcolumn+i].getCellAddress())  # テンプレートの列と異なる行のセル範囲を取得。
		j = c - VARS.datacolumn  # 相対インデックスを取得。
		for cell in cellranges.getCells():
			k = cell.getCellAddress().Row - VARS.monthrow 
			if datarows[k][j] in ("", "/", "x"):  # テンプレートを優先する文字列の時。
				datarows[k][j] = templates[k, i]  # テンプレートの値を使う。
		excludes.append(c)	
	return queryTemplateColumn
def setRangesProperty(doc, columnindexes, prop):  # r行のcolumnindexesの列のプロパティを変更。prop: プロパティ名とその値のリスト。
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
	cellranges.addRangeAddresses((VARS.sheet[VARS.daterow:VARS.datarow, i].getRangeAddress() for i in columnindexes), False)  # セル範囲コレクションを取得。
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
