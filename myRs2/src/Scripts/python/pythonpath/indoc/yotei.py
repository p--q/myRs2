#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import commons
from calendar import monthrange
from datetime import date, timedelta
from itertools import chain
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
	
	
	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	sheet["A1"].setString("ﾘｽﾄに戻る")
	sheet["AF1"].setString("COPY")
	sheet["AK1"].setString("強有効")
	VARS.setSheet(sheet)
	
	daycount = 31  # シートに表示する日数。
	if VARS.datacolumn+daycount>VARS.templatestartcolumn:  # daycountの上限はテンプレート列までにする。
		daycount = VARS.templatestartcolumn - VARS.datacolumn		
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。		
	endedgecolumn = VARS.datacolumn + daycount  # 更新後のデータの右端列の右。
	todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。	
	firstdatevalue = int(sheet[VARS.dayrow, VARS.datacolumn].getValue())  # 先頭の日付のシリアル値を整数で取得。空セルの時は0.0が返る。	
	if firstdatevalue>0:  # シリアル値が取得できた時。	
		diff = todayvalue - firstdatevalue  # 今日の日付と先頭の日付との差を取得。
		if not diff>0:  # 先頭日付が今日より前でない時はここで終わる。
			return
		todaycolumn = VARS.datacolumn + diff # 移動前の今日の日付列インデックスを取得。	
		if todaycolumn<VARS.firstemptycolumn:  # 今日の日付列が表示されている範囲内にある時。今日の日付を先頭に移動させる。
			controller = doc.getCurrentController()  # コントローラの取得。
			docframe = controller.getFrame()
			dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
			controller.select(sheet[VARS.monthrow:VARS.emptyrow, todaycolumn:VARS.templatestartcolumn-1])  #  移動前の今日の日付列以降テンプレート列左までを選択。
			dispatcher.executeDispatch(docframe, ".uno:Cut", "", 0, ())  # 選択範囲をカット。	
			controller.select(sheet[VARS.monthrow, VARS.datacolumn])  # ペーストする左上セルを選択。
			dispatcher.executeDispatch(docframe, ".uno:Paste", "", 0, ())  # ペースト。	
		else:
			sheet[VARS.monthrow:VARS.emptyrow, VARS.datacolumn:endedgecolumn].clearContents(511)  # シートのデータ部分を全部クリア。	
	else:
		sheet[VARS.monthrow:VARS.emptyrow, VARS.datacolumn:endedgecolumn].clearContents(511)  # シートのデータ部分を全部クリア。	
	todaydate = date.today()  # 今日のdateオブジェクトを取得。
	weekday = todaydate.weekday()  # 月=0が返る。
	weekdays = "月", "火", "水", "木", "金", "土", "日", "祝"  # シートでは日=1であることに注意。最後に祝日も追加している。		
	datarows = [["" for dummy in range(daycount)],\
			[i for i in range(todayvalue, todayvalue+daycount)],\
			[weekdays[i%7] for i in range(weekday, weekday+daycount)]]  # 月行、日行と曜日行を作成。日付はシリアル値で入力しないといけない。
	datarows.extend(list(i) for i in sheet[VARS.datarow:VARS.emptyrow, VARS.datacolumn:endedgecolumn].getDataArray())  # シートのデータ部分を取得。タプルをリストにして取得。			
	dates = [todaydate+timedelta(days=i) for i in range(daycount)]  # 表示する日数をdateオブジェクトで取得。
	templatedic = {}  # キー: テンプレート列インデックス、値: 日付列インデックスのリスト。
	templates = sheet[VARS.monthrow:VARS.emptyrow, VARS.templatestartcolumn:VARS.templateendcolumnedge].getDataArray()  # テンプレートの値を月行から取得。
	excludes = set()  # 処理済列インデックスの集合。
	
	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	
	
	for ti in range(len(templates[0]))[::-1]:  # テンプレートの列の相対インデックスをイテレート。優先度を付けるため後ろからイテレート。
		tm = templates[1][ti]  # 空文字、週数の文字列、月のfloat、のいずれかが返る。
		td = templates[2][ti]  # 曜日or日の要素を取得。
		if not td in weekdays:  # weekdaysの要素にない時は日指定。
			td = convertToInteger(td)  # 日を整数に変換して取得。
			if tm:  # 月日指定がある時。
				tm = convertToInteger(tm)  # 月を整数に変換して取得。
				for y in range(dates[0].year, dates[-1].year+1):  # 表示期間の年をイテレート。
					d = date(y, tm, td)
					if d in dates:
						c = VARS.datacolumn + dates.index(d)  # 列インデックスを取得。
						if not c in excludes:
							templatedic.setdefault(VARS.templatestartcolumn+ti, []).append(c)
							excludes.add(c)
			else:  # 日指定のみの時。
				d = dates[0].replace(day=td)  # 開始日と同じ月の日を取得。
				while d<=dates[-1]:
					if d in dates:
						c = VARS.datacolumn + dates.index(d)  # 列インデックスを取得。
						if not c in excludes:
							templatedic.setdefault(VARS.templatestartcolumn+ti, []).append(c)
							excludes.add(c)
					d += timedelta(days=monthrange(d.year, d.month)[1])  # 翌月の同じ日を取得。
		elif td=="祝":  # 祝日の時。
			holidays = commons.HOLIDAYS	
			for y in range(dates[0].year, dates[-1].year+1):  # 表示期間の年をイテレート。
				for m, ds in enumerate(holidays[y], start=1):  # 祝日のリストを月ごとにイテレート。
					for hd in ds:
						d = date(y, m, hd)
						if d in dates:
							c = VARS.datacolumn + dates.index(d)  # 列インデックスを取得。
							if not c in excludes:
								templatedic.setdefault(VARS.templatestartcolumn+ti, []).append(c)	
								excludes.add(c)
		else:  # 曜日指定のある時。
			n = weekdays.index(td)  # 月=0の曜日番号を取得。
			ws = range((n-weekday)%7, daycount, 7)  # 同じ曜日の相対インデックスを取得。
			if tm:  # 週数or月の指定がある時。
				if tm.endswith("w"):  # wで終わっている時、週数と曜日指定の時。
					w = convertToInteger(tm[:-1])  # 週数を整数に変換して取得。
					for i in ws:  # 同じ曜日の相対インデックスを取得。
						if w==-(-dates[i].day//7):  # 週番号が一致する時。商を切り上げ。	
							c = VARS.datacolumn + i
							if not c in excludes:
								templatedic.setdefault(VARS.templatestartcolumn+ti, []).append(c)
								excludes.add(c)
				else:  # 月と曜日指定の時。
					m = convertToInteger(m)  # 月を整数に変換して取得。
					for i in ws:  # 同じ曜日の相対インデックスを取得。
						if m==dates[i].month:
							c = VARS.datacolumn + i
							if not c in excludes:
								templatedic.setdefault(VARS.templatestartcolumn+ti, []).append(c)
								excludes.add(c)
			else:  # 曜日のみの指定の時。	
				for i in ws:  # 同じ曜日の相対インデックスを取得。
					c = VARS.datacolumn + i
					if not c in excludes:
						templatedic.setdefault(VARS.templatestartcolumn+ti, []).append(c)

	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	
	for tc, cs in templatedic.items():  # tc: テンプレートの列インデックス、cs:  書き換える列インデックスのリスト。
		celladdress = sheet[VARS.datarow, tc].getCellAddress()
		for c in cs:
			cellranges = sheet[VARS.datarow:VARS.emptyrow, c].queryRowDifferences(celladdress)  # テンプレートの列と異なる行のセル範囲を取得。セル範囲コレクションに対しては動かない。
			if len(cellranges):
				j = c - VARS.datacolumn  # 相対インデックスを取得。
				rowindexes = (range(i.StartRow-VARS.monthrow, i.EndRow+1-VARS.monthrow) for i in cellranges.getRangeAddresses())  # 相対インデックスをイテレートするイテレーター。getCells()ではなぜか何もイテレートされない。
				for k in chain.from_iterable(rowindexes):
					if datarows[k][j] in ("", "/", "x"):  # テンプレートを優先する文字列の時。
						datarows[k][j] = templates[k][i]  # テンプレートの値を使う。		
	sheet[VARS.monthrow:VARS.emptyrow, VARS.datacolumn:endedgecolumn].setDataArray(datarows)
		


			
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
# 	
# 	print("")		
			
# 祝日>月と日指定>日指定のみ>週数と曜日指定>月と曜日指定>曜日指定、の優先度。			


# 						
# 	holidays = set()  # 祝日の列インデックスを入れる集合。
# 	excludes = set()  # 処理済列インデックスの集合。
# 	
# 	
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)	
# 	
# 	
# 	y, m, d = [int(functionaccess.callFunction(i, (todayvalue,))) for i in ("YEAR", "MONTH", "DAY")]  # 今日の年月日を整数で取得。
# 	queryTemplateColumn = createQueryTemplateColumn(datarows, templates, excludes)
# 	firstdatevalue = todayvalue  # 次月の初日のシリアル値を入れる変数。最初は今日のシリアル値を入れておく。
# 	firstdaycolumn = VARS.datacolumn  # 今月の初日の列インデックス。
# 	for td in ds.keys():
# 		if d-1<td:  # 開始日以降の時。
# 			queryTemplateColumn(firstdaycolumn-d+td, ds[td])
# 	if m in ms:  # テンプレートに指定のある月の時。
# 		for i in ms[m]:  # テンプレートの相対列インデックスをイテレート。
# 			td = templates[1][i]  # 指定日を取得。
# 			if td:
# 				td = convertToInteger(td)
# 				if d-1<td:  # 開始日以降の時。
# 					queryTemplateColumn(firstdaycolumn-d+td, i)


# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)					
					
# 	if y in commons.HOLIDAYS:  # 年が祝日一覧のキーにある時。
# 		holidays.update(firstdaycolumn+i for i in commons.HOLIDAYS[y][m-1] if d-1<i)  # 祝日の日付の列インデックスを取得。									
# 	if m>11:  # 12月の次は1月にする。
# 		m = 1  # 次月を取得。
# 		y += 1  # 年も更新する。
# 	else:
# 		m += 1  # 次月を取得。
# 	firstdatevalue = int(functionaccess.callFunction("EOMONTH", (firstdatevalue, 0))) + 1  # 次月の初日のシリアル値を取得。
# 	firstdaycolumn += firstdatevalue - todayvalue  # 次月の初日の列インデックスを取得。
# 	while True:  # 月の全部が含まれている時。
# 		nextfirstdatevalue = int(functionaccess.callFunction("EOMONTH", (firstdatevalue, 0))) + 1  # 次月の初日のシリアル値を取得。
# 		nextfirstdaycolumn = firstdaycolumn + nextfirstdatevalue - firstdatevalue  # 次月の初日の列インデックスを取得。
# 		if nextfirstdaycolumn-1<endedgecolumn:  # 月の途中で終わっている時はwhile文を抜ける。
# 			break
# 		for td in ds.keys():
# 			queryTemplateColumn(firstdaycolumn+td-1, ds[td])
# 		if m in ms:  # テンプレートに指定のある月の時。
# 			for i in ms[m]:  # テンプレートの相対列インデックスをイテレート。
# 				td = templates[1][i]  # 指定日を取得。
# 				queryTemplateColumn(firstdaycolumn+td-1, i)			
# 		if y in commons.HOLIDAYS:  # 年が祝日一覧のキーにある時。
# 			holidays.update(firstdaycolumn+i for i in commons.HOLIDAYS[y][m-1])  # 祝日の日付の列インデックスを取得。	
# 		firstdatevalue = nextfirstdatevalue			
# 		firstdaycolumn = nextfirstdaycolumn	
# 		if m>11:  # 12月の次は1月にする。
# 			m = 1  # 次月を取得。
# 			y += 1  # 年も更新する。
# 		else:
# 			m += 1  # 次月を取得。		
# 	d = endedgecolumn - firstdaycolumn  # 月の最終日を取得。
# 	for td in ds.keys():  # 指定日についてイテレート。
# 		if td<d+1:  # 終了日より前の時。
# 			queryTemplateColumn(firstdaycolumn+td-1, ds[td])					
# 	if m in ms:  # テンプレートに指定のある月の時。
# 		for i in ms[m]:  # テンプレートの相対インデックスをイテレート。
# 			td = templates[1][i]  # 指定日を取得。
# 			if td<d+1:  # 終了日より前の時。
# 				queryTemplateColumn(firstdaycolumn+td-1, i)							
# 	if y in commons.HOLIDAYS:  # 年が祝日一覧のキーにある時。
# 		holidays.update(firstdaycolumn+i for i in commons.HOLIDAYS[y][m-1] if i<d+1)  # 祝日の日付の列インデックスを取得。	
# 		
# 
# 	queryWeekdayColumn = createQueryWeekdayColumn(datarows, templates)		
# 	for n in range(1, 8):  # 曜日番号をn=1からイテレート。
# 		templatecolumns = templatecolumnlists[n]  # 同じ曜日のテンプレートの列インデックスのリストを取得。
# 		for c in range(VARS.datacolumn+(n-weekday)%7, endedgecolumn, 7):  # 同じ曜日の列インデックスを取得。
# 			if not c in excludes:  # 処理済の列インデックス以外の時。
# 				if templatecolumns>1:  # 複数列がある時は週番号指定(2wなど)列を含む。
# 					j = c - VARS.datacolumn  # 相対インデックスを取得。
# 					for tc in templatecolumns:
# 						w = templates[1][tc-VARS.templatestartcolumn]  # 週数の行の値を取得。
# 						if w.endswith("w"):  # wで終わる時は週番号。		
# 							d = int(functionaccess.callFunction("DAY", (datarows[1][j],)))  # 月の何日目か取得。		
# 							if int(w[:-1])==-(-d//7):  # 週番号が一致する時。-(-d//7)切り上げ。	
# 								queryWeekdayColumn(c, tc)	
# 						elif not w:  # 空セルのときは曜日のみ指定。
# 							queryWeekdayColumn(c, tc)				
# 				else:
# 					queryWeekdayColumn(c, tc)	
# 	n = 7  # 土曜日の曜日番号。
# 	columnindexes = range(VARS.datacolumn+(n-weekday)%7, endedgecolumn, 7)   # 土曜日の列インデックスを取得。			
# 	setRangesProperty(doc, columnindexes, ("CharColor", commons.COLORS["skyblue"]))  # 土曜日の文字色を設定。	
# 	n = 1  # 日曜日の曜日番号。
# 	columnindexes = range(VARS.datacolumn+(n-weekday)%7, endedgecolumn, 7)   # 日曜日の列インデックスを取得。
# 	setRangesProperty(doc, columnindexes, ("CharColor", commons.COLORS["red3"]))  # 日曜日の文字色を設定。				
# 	holidays.difference_update(columnindexes)  # 日曜日と重なっている祝日を除く。	
# 	holidays = filter(lambda x: x<endedgecolumn, holidays)  # 上限を設定。
# 	setRangesProperty(doc, holidays, ("CellBackColor", commons.COLORS["red3"]))  # 祝日の背景色を設定。
# 	for c in holidays:
# 		sheet[VARS.dayrow, c].setDataArray(("x",)*(VARS.emptyrow-VARS.datarow))	
# 	createFormatKey = commons.formatkeyCreator(doc)	
# 	sheet[VARS.dayrow, VARS.datacolumn:endedgecolumn].setPropertyValue("NumberFormat", createFormatKey('D'))  
# 	
# 	sheet[VARS.monthrow:VARS.emptyrow, VARS.datacolumn:endedgecolumn].setPropertyValue("HoriJustify", CENTER)  			
# 	
# 	ranges = sheet[VARS.monthrow:VARS.emptyrow, VARS.datacolumn:endedgecolumn],\
# 			sheet[VARS.datarow:VARS.emptyrow, VARS.templatestartcolumn:VARS.templateendcolumnedge]
# 	datarange = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
# 	datarange.addRangeAddresses((i.getRangeAddress() for i in ranges), False)		
# 	datarange.setPropertyValue("CellBackColor", -1)  # 背景色をクリア。
# 	searchdescriptor = sheet.createSearchDescriptor()
# 	searchdescriptor.setSearchString("x")  # 戻り値はない。
# 	cellranges = datarange.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
# 	if cellranges:
# 		cellranges.setPropertyValue("CellBackColor", commons.COLORS["gray7"])
# 	searchdescriptor.setSearchString("/")  # 戻り値はない。
# 	cellranges = datarange.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
# 	if cellranges:
# 		cellranges.setPropertyValue("CellBackColor", commons.COLORS["silver"])	
# 	searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
# 	searchdescriptor.setSearchString("[^x/]")  # 戻り値はない。	
# 	if cellranges:
# 		cellranges.setPropertyValue("CellBackColor", commons.COLORS["magenta3"])		

		
		
		
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
# 	sheet[VARS.dayrow:VARS.emptyrow, VARS.datacolumn:endedgecolumn].setDataArray(datarows)
# 	createFormatKey = commons.formatkeyCreator(doc)	
# 	sheet[VARS.dayrow, VARS.datacolumn:endedgecolumn].setPropertyValue("NumberFormat", createFormatKey('D'))  
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
def convertToInteger(s):  # s: floatか文字列。
	if isinstance(s, float):  # floatの時。
		return int(s)
	elif s.isdigit():  # 数字のみの文字列の時。
		return int(s)	
# def createQueryWeekdayColumn(datarows, templates):
# 	def queryWeekdayColumn(c, tc):
# 		j = c - VARS.datacolumn  # 相対インデックスを取得。
# 		cellranges = VARS.sheet[VARS.datarow:VARS.emptyrow, c].queryRowDifferences(VARS.sheet[VARS.dayrow, tc].getCellAddress())  # テンプレートの列と異なる行のセル範囲を取得。
# 		for cell in cellranges.getCells():
# 			k = cell.getCellAddress().Row - VARS.monthrow 
# 			if datarows[k][j] in ("", "/", "x"):  # テンプレートを優先する文字列の時。
# 				datarows[k][j] = templates[k, tc-VARS.templatestartcolumn]  # テンプレートの値を使う。		
# 	return queryWeekdayColumn
def createQueryTemplateColumn(datarows, templates, excludes):
	def queryTemplateColumn(c, i):  #  c: 更新する列インデックス、i: テンプレートの相対インデックス。
		
		
		cellranges = VARS.sheet[VARS.datarow:VARS.emptyrow, c].queryRowDifferences(VARS.sheet[VARS.dayrow, VARS.templatestartcolumn+i].getCellAddress())  # テンプレートの列と異なる行のセル範囲を取得。
		j = c - VARS.datacolumn  # 相対インデックスを取得。
		rowindexes = (range(i.StartRow-VARS.monthrow, i.EndRow+1-VARS.monthrow) for i in cellranges.getRangeAddresses())  # 相対インデックスをイテレートするイテレーター。getCells()ではなぜか何もイテレートされない。
		for k in chain.from_iterable(rowindexes):
			if datarows[k][j] in ("", "/", "x"):  # テンプレートを優先する文字列の時。
				datarows[k][j] = templates[k][i]  # テンプレートの値を使う。		
		excludes.append(c)	
	return queryTemplateColumn
def setRangesProperty(doc, columnindexes, prop):  # r行のcolumnindexesの列のプロパティを変更。prop: プロパティ名とその値のリスト。
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
	cellranges.addRangeAddresses((VARS.sheet[VARS.dayrow:VARS.datarow, i].getRangeAddress() for i in columnindexes), False)  # セル範囲コレクションを取得。
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
