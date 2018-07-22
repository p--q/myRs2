#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import commons
from calendar import monthrange
from datetime import date, timedelta  # 日付計算はシート関数では複雑になりすぎてロジックが組めないのでこれを使う。
from com.sun.star.awt import Key  # 定数
from com.sun.star.awt import KeyEvent  # Struct
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.table.CellHoriJustify import CENTER  # enum
class Schedule():  # シート固有の定数設定。
	def __init__(self):
		self.menurow = 0  # メニュー行。
		self.monthrow = 1  # 月行。
		self.dayrow = 2  # 日行。
		self.weekdayrow = 3  # 曜日行。
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
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	sheet["A1"].setString("ﾘｽﾄに戻る")
	sheet["AF1"].setString("COPY")
	sheet["AK1"].setString("強有効")
	VARS.setSheet(sheet)
	daycount = 31  # シートに表示する日数。
	monthrow = VARS.monthrow
	datarow = VARS.datarow
	emptyrow = VARS.emptyrow
	datacolumn = VARS.datacolumn
	templatestartcolumn = VARS.templatestartcolumn
	if datacolumn+daycount>templatestartcolumn:  # daycountの上限はテンプレート列までにする。
		daycount = templatestartcolumn - datacolumn		
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。		
	endedgecolumn = datacolumn + daycount  # 更新後のデータの右端列の右。
	todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。	
	firstdatevalue = int(sheet[VARS.dayrow, datacolumn].getValue())  # 先頭の日付のシリアル値を整数で取得。空セルの時は0.0が返る。	
	if firstdatevalue>0:  # シリアル値が取得できた時。	
		diff = todayvalue - firstdatevalue  # 今日の日付と先頭の日付との差を取得。
		if diff<0:  # 先頭日付が未来の時はここで終わる。同日の時は更新する。
			return
		todaycolumn = datacolumn + diff # 移動前の今日の日付列インデックスを取得。	
		if diff and todaycolumn<VARS.firstemptycolumn:  # 今日の日付列が表示されている範囲内にある時。今日の日付を先頭に移動させる。先頭が今日でない時は移動させない。
			controller = doc.getCurrentController()  # コントローラの取得。
			docframe = controller.getFrame()
			dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
			controller.select(sheet[monthrow:emptyrow, todaycolumn:templatestartcolumn-1])  #  移動前の今日の日付列以降テンプレート列左までを選択。
			dispatcher.executeDispatch(docframe, ".uno:Cut", "", 0, ())  # 選択範囲をカット。	
			controller.select(sheet[monthrow, datacolumn])  # ペーストする左上セルを選択。
			dispatcher.executeDispatch(docframe, ".uno:Paste", "", 0, ())  # ペースト。	
			componentwindow	= controller.ComponentWindow  # コンポーネントウィンドウを取得。
			keyevent = KeyEvent(KeyCode=Key.ESCAPE, KeyChar=chr(0x1b), Modifiers=0, KeyFunc=0, Source=componentwindow)  # EscキーのKeyEventを取得。
			toolkit = componentwindow.getToolkit()  # ツールキットを取得。
			toolkit.keyPress(keyevent)  # キーを押す、をシミュレート。
			toolkit.keyRelease(keyevent)  # キーを離す、をシミュレート。
			controller.select(sheet[datarow, datacolumn])
		else:
			sheet[monthrow:emptyrow, datacolumn:endedgecolumn].clearContents(511)  # シートのデータ部分を全部クリア。	
	else:
		sheet[monthrow:emptyrow, datacolumn:endedgecolumn].clearContents(511)  # シートのデータ部分を全部クリア。	
	todaydate = date.today()  # 今日のdateオブジェクトを取得。
	weekday = todaydate.weekday()  # 月=0が返る。
	weekdays = "月", "火", "水", "木", "金", "土", "日", "祝"  # シートでは日=1であることに注意。最後に祝日も追加している。		
	datarows = [["" for dummy in range(daycount)],\
			[i for i in range(todayvalue, todayvalue+daycount)],\
			[weekdays[i%7] for i in range(weekday, weekday+daycount)]]  # 月行、日行と曜日行を作成。日付はシリアル値で入力しないといけない。
	datarows.extend(list(i) for i in sheet[datarow:emptyrow, datacolumn:endedgecolumn].getDataArray())  # シートのデータ部分を取得。タプルをリストにして取得。			
	dates = [todaydate+timedelta(days=i) for i in range(daycount)]  # 表示する日数をdateオブジェクトで取得。
	datarows[0][0] = "{}月".format(todaydate.month)
	nextidx = monthrange(todaydate.year, todaydate.month)[1] - todaydate.day + 1  # 次月の初日のインデックス。
	while nextidx<len(dates):
		d = dates[nextidx]
		datarows[0][nextidx] = "{}月".format(d.month)
		nextidx += monthrange(d.year, d.month)[1] + 1
	templatedic = {}  # キー: テンプレート列インデックス、値: 日付列インデックスのリスト。
	templates = sheet[monthrow:emptyrow, templatestartcolumn:VARS.templateendcolumnedge].getDataArray()  # テンプレートの値を月行から取得。
	excludes = set()  # 処理済列インデックスの集合。
	holidayset = set()  # 祝日列インデックスの集合。
	for ti in range(len(templates[0]))[::-1]:  # テンプレートの列の相対インデックスをイテレート。優先度を付けるため後ろからイテレート。
		tm = templates[1][ti]  # 空文字、週数の文字列、月のfloat、のいずれかが返る。
		td = templates[2][ti]  # 曜日or日の要素を取得。
		tc = templatestartcolumn + ti
		if not td in weekdays:  # weekdaysの要素にない時は日指定。
			td = convertToInteger(td)  # 日を整数に変換して取得。
			if tm:  # 月日指定がある時。
				tm = convertToInteger(tm)  # 月を整数に変換して取得。
				for y in range(dates[0].year, dates[-1].year+1):  # 表示期間の年をイテレート。
					d = date(y, tm, td)
					if d in dates:
						c = datacolumn + dates.index(d)  # 列インデックスを取得。
						if not c in excludes:
							templatedic.setdefault(tc, []).append(c)
							excludes.add(c)
			else:  # 日指定のみの時。
				d = dates[0].replace(day=td)  # 開始日と同じ月の日を取得。
				while d<=dates[-1]:
					if d in dates:
						c = datacolumn + dates.index(d)  # 列インデックスを取得。
						if not c in excludes:
							templatedic.setdefault(tc, []).append(c)
							excludes.add(c)
					d += timedelta(days=monthrange(d.year, d.month)[1])  # 翌月の同じ日を取得。
		elif td=="祝":  # 祝日の時。
			holidayset.add(tc)  # テンプレートの祝列インデックスを取得。
			holidays = commons.HOLIDAYS	
			for y in range(dates[0].year, dates[-1].year+1):  # 表示期間の年をイテレート。
				for m, ds in enumerate(holidays[y], start=1):  # 祝日のリストを月ごとにイテレート。
					for hd in ds:
						d = date(y, m, hd)
						if d in dates:
							c = datacolumn + dates.index(d)  # 列インデックスを取得。
							if not c in excludes:
								templatedic.setdefault(tc, []).append(c)	
								excludes.add(c)
								holidayset.add(c)
		else:  # 曜日指定のある時。
			n = weekdays.index(td)  # 月=0の曜日番号を取得。
			ws = range((n-weekday)%7, daycount, 7)  # 同じ曜日の相対インデックスを取得。
			if tm:  # 週数or月の指定がある時。
				if tm.endswith("w"):  # wで終わっている時、週数と曜日指定の時。
					w = convertToInteger(tm[:-1])  # 週数を整数に変換して取得。
					for i in ws:  # 同じ曜日の相対インデックスを取得。
						if w==-(-dates[i].day//7):  # 週番号が一致する時。商を切り上げ。	
							c = datacolumn + i
							if not c in excludes:
								templatedic.setdefault(tc, []).append(c)
								excludes.add(c)
				else:  # 月と曜日指定の時。
					m = convertToInteger(m)  # 月を整数に変換して取得。
					for i in ws:  # 同じ曜日の相対インデックスを取得。
						if m==dates[i].month:
							c = datacolumn + i
							if not c in excludes:
								templatedic.setdefault(tc, []).append(c)
								excludes.add(c)
			else:  # 曜日のみの指定の時。	
				for i in ws:  # 同じ曜日の相対インデックスを取得。
					c = datacolumn + i
					if not c in excludes:
						templatedic.setdefault(tc, []).append(c)
	for tc, cs in templatedic.items():  # tc: テンプレートの列インデックス、cs:  書き換える列インデックスのリスト。
		dataranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
		dataranges.addRangeAddresses((sheet[datarow:emptyrow, i].getRangeAddress() for i in cs), False)  # セル範囲コレクションを取得。		
		cellranges = dataranges.queryRowDifferences(sheet[datarow, tc].getCellAddress())  # テンプレートの列と異なる行のセル範囲を取得。		
		ti = tc - templatestartcolumn  # テンプレートの列の相対インデックスを取得。
		for rangeaddress in cellranges.getRangeAddresses():  # getCells()ではなぜか何もイテレートされない。
			for j in range(rangeaddress.StartColumn-datacolumn, rangeaddress.EndColumn+1-datacolumn):
				for k in range(rangeaddress.StartRow-monthrow, rangeaddress.EndRow+1-monthrow):
					if datarows[k][j] in ("", "/", "x"):  # テンプレートを優先する文字列の時。
						datarows[k][j] = templates[k][ti]  # テンプレートの値を使う。											
	sheet[monthrow:emptyrow, datacolumn:endedgecolumn].setDataArray(datarows)
	colors = commons.COLORS
	n = 5  # 土曜日の曜日番号。
	columnindexes = range(datacolumn+(n-weekday)%7, endedgecolumn, 7)   # 土曜日の列インデックスを取得。			
	setRangesProperty(doc, columnindexes, ("CharColor", colors["skyblue"]))  # 土曜日の文字色を設定。	
	n = 6  # 日曜日の曜日番号。
	columnindexes = range(datacolumn+(n-weekday)%7, endedgecolumn, 7)   # 日曜日の列インデックスを取得。
	setRangesProperty(doc, columnindexes, ("CharColor", colors["red3"]))  # 日曜日の文字色を設定。				
	holidayset.difference_update(columnindexes)  # 日曜日と重なっている祝日を除く。	
	setRangesProperty(doc, holidayset, ("CellBackColor", colors["red3"]))  # 祝日の背景色を設定。
	createFormatKey = commons.formatkeyCreator(doc)	
	sheet[VARS.dayrow, datacolumn:endedgecolumn].setPropertyValue("NumberFormat", createFormatKey('D'))  
	dataranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。	
	ranges = sheet[datarow:emptyrow, datacolumn:endedgecolumn],\
			sheet[datarow:emptyrow, templatestartcolumn:VARS.templateendcolumnedge]  # テンプレートを含めたデータ範囲。
	dataranges.addRangeAddresses((i.getRangeAddress() for i in ranges), False)		
	dataranges.setPropertyValue("CellBackColor", -1)  # 背景色をクリア。
	setPropSearchedCells = createSetPropSearchedCells(dataranges)
	setPropSearchedCells("x", ("CellBackColor", colors["gray7"]))
	setPropSearchedCells("/", ("CellBackColor", colors["silver"]))
	searchdescriptor = sheet.createSearchDescriptor()
	searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
	searchdescriptor.setSearchString("[^x/]")  # 戻り値はない。	
	cellranges = dataranges.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
	if cellranges:
		cellranges.setPropertyValue("CellBackColor", colors["magenta3"])	
	temlatedaterange = sheet[VARS.weekdayrow, templatestartcolumn:VARS.templateendcolumnedge]  # テンプレートの日付範囲のみ。
	setPropSearchedCells = createSetPropSearchedCells(temlatedaterange)
	setPropSearchedCells("土", ("CharColor", colors["skyblue"]))
	setPropSearchedCells("日", ("CharColor", colors["red3"]))
	ranges = sheet[monthrow:datarow, datacolumn:endedgecolumn], temlatedaterange
	dataranges.addRangeAddresses((i.getRangeAddress() for i in ranges), False)			
	dataranges.setPropertyValue("HoriJustify", CENTER) 
def createSetPropSearchedCells(cellrange):	
	searchdescriptor = VARS.sheet.createSearchDescriptor()
	def setPropSearchedCells(txt, prop):		
		searchdescriptor.setSearchString(txt)  # 戻り値はない。
		cellranges = cellrange.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
		if cellranges:
			cellranges.setPropertyValue(*prop)		
	return setPropSearchedCells
def convertToInteger(s):  # s: floatか文字列。
	if isinstance(s, float):  # floatの時。
		return int(s)
	elif s.isdigit():  # 数字のみの文字列の時。
		return int(s)	
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
