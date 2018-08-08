#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import commons, staticdialog, ichiran, transientdialog
from calendar import monthrange
from datetime import date, datetime, time, timedelta  # 日付計算はシート関数では遅いし複雑になりすぎてロジックが組めないのでこれを使う。
from com.sun.star.awt import MessageBoxButtons, MessageBoxResults, MouseButton, Key  # 定数
from com.sun.star.awt import KeyEvent  # Struct
from com.sun.star.awt.MessageBoxType import ERRORBOX, QUERYBOX  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.table.CellHoriJustify import CENTER, LEFT  # enum
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
class Schedule():  # シート固有の定数設定。
	def __init__(self):
		self.menurow = 0  # メニュー行。
		self.monthrow = 1  # 月行。
		self.dayrow = 2  # 日行。
		self.weekdayrow = 3  # 曜日行。
		self.datarow = 4  # データ開始行。
		self.datacolumn = 1  # データ開始列。
		self.weekdays = "月", "火", "水", "木", "金", "土", "日", "祝"  # シートでは日=1であることに注意。最後に祝日も追加している。		
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
	sheet["A1"].setString("一覧へ")
	sheet["C1"].setString("COPY")
	sheet["I1"].setString("強有効")
	sheet["O1"].setString("3wCOPY")
	sheet["AM1"].setString("休日更新")
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
		if diff>0:  # 先頭日付が過去の時。
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
				controller.select(sheet[emptyrow, datacolumn])			
		elif diff<0:  # 先頭日付が未来の時はここで終わる。
			return
	else:
		sheet[monthrow:emptyrow, datacolumn:endedgecolumn].clearContents(511)  # シートのデータ部分を全部クリア。	
	todaydate = date.today()  # 今日のdateオブジェクトを取得。
	weekday = todaydate.weekday()  # 月=0が返る。
	weekdays = VARS.weekdays	
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
	annotations = sheet.getAnnotations()  # コメントコレクションを取得。					
	comments = [(i.getPosition(), i.getString()) for i in annotations]  # setDataArray()でコメントがクリアされるのでここでセルアドレスとコメントの文字列をタプルで取得しておく。											
	sheet[monthrow:emptyrow, datacolumn:endedgecolumn].setDataArray(datarows)  # コメントが消されてしまう。
	[annotations.insertNew(*i) for i in comments]  # コメントを再挿入。
	colors = commons.COLORS
	n = 5  # 土曜日の曜日番号。
	columnindexes = range(datacolumn+(n-weekday)%7, endedgecolumn, 7)   # 土曜日の列インデックスを取得。			
	setRangesProperty(doc, columnindexes, ("CharColor", colors["skyblue"]))  # 土曜日の文字色を設定。	
	n = 6  # 日曜日の曜日番号。
	columnindexes = range(datacolumn+(n-weekday)%7, endedgecolumn, 7)   # 日曜日の列インデックスを取得。
	setRangesProperty(doc, columnindexes, ("CharColor", colors["red3"]))  # 日曜日の文字色を設定。
	
		
	# 休日の背景色をsilverにする。
	
	
	
	
				
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
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	selection = eventobject.Source.getSelection()
	if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択範囲がセルの時。矢印キーでセルを移動した時。マウスクリックハンドラから呼ばれると何回も発火するのでその対応。
		currenttableborder2 = selection.getPropertyValue("TableBorder2")  # 選択セルの枠線を取得。
		if all((currenttableborder2.TopLine.Color==currenttableborder2.LeftLine.Color==commons.COLORS["violet"],\
				currenttableborder2.RightLine.Color==currenttableborder2.BottomLine.Color==commons.COLORS["magenta3"])):  # 枠線の色を確認。
			return  # すでに枠線が書いてあったら何もしない。
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
		VARS.setSheet(selection.getSpreadsheet())  # シートを切り替えた時点でselectionChanged()メソッドが発火するためここで渡しておかないといけない。
		drowBorders(selection)  # 枠線の作成。
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
				if r==VARS.menurow:
					return wClickMenu(enhancedmouseevent, xscriptcontext)
				elif VARS.datarow-1<r<VARS.emptyrow:
					if VARS.datacolumn-1<c<VARS.firstemptycolumn or VARS.templatestartcolumn-1<c<VARS.templateendcolumnedge:
						return wClickCell(enhancedmouseevent, xscriptcontext)
	return True  # セル編集モードにする。	
def wClickMenu(enhancedmouseevent, xscriptcontext):
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	txt = selection.getString()	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。		
	if txt=="一覧へ":
		sheets = doc.getSheets()  # シートコレクションを取得。		
		controller.setActiveSheet(sheets["一覧"])  # 一覧シートをアクティブにする。
		return False  # セル編集モードにしない。	
	sheet = VARS.sheet
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。			
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。		
	systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
	startdatevalue = sheet[VARS.dayrow, VARS.datacolumn].getValue()
	startdate = date(*[int(functionaccess.callFunction(i, (startdatevalue,))) for i in ("YEAR", "MONTH", "DAY")])
	starttimevalue = sheet[VARS.datarow, 0].getValue()
	starttime = time(*[int(functionaccess.callFunction(i, (starttimevalue,))) for i in ("HOUR", "MINUTE")])
	starttime = datetime.combine(startdate, starttime)  # timeオブジェクトではtimedelta()で加減算できないのでdatetimeオブジェクトに変換する。
	timegen = (starttime+timedelta(minutes=30*i) for i in range(VARS.emptyrow-VARS.datarow))  # 30分毎に枠を取得。
	times = ["{}:{:0>2}".format(i.hour, i.minute) for i in timegen]
	outputs = [sheet[VARS.menurow, VARS.templatestartcolumn].getString()]  # 最初の文をセルから取得。
	if txt=="COPY":
		createScheduleToClip(systemclipboard, times, startdate, outputs)(14)					
	elif txt=="強有効":
		n = 14
		searchdescriptor = sheet.createSearchDescriptor()
		searchdescriptor.setSearchString("強")  # 戻り値はない。	
		dategene = (startdate+timedelta(days=i) for i in range(n))
		weekdays = VARS.weekdays
		dates = ["{}/{}({})".format(i.month, i.day, weekdays[i.weekday()]) for i in dategene]	
		for i in range(VARS.datacolumn, VARS.datacolumn+n):  # 列インデックスをイテレート。
			datarange = sheet[VARS.datarow:VARS.emptyrow, i]
			cellranges = datarange.queryEmptyCells()  # 空セルのセル範囲コレクションを取得。
			searchedcellranges = datarange.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
			if searchedcellranges:			
				cellranges.addRangeAddresses(searchedcellranges.getRangeAddresses(), True)	# Falseにするとセル範囲を取り出す時に追加順にある。	
			fs = [" ".join([times[j], "○"]) for i in cellranges.getRangeAddresses() for j in range(i.StartRow-VARS.datarow, i.EndRow+1-VARS.datarow)]
			if fs:
				outputs.extend(["", dates[i-VARS.datacolumn]])
				outputs.extend(fs)	
		systemclipboard.setContents(commons.TextTransferable("\n".join(outputs)), None)  # クリップボードにコピーする。	
	elif txt=="3wCOPY":
		createScheduleToClip(systemclipboard, times, startdate, outputs)(21)
	elif txt=="休日更新":  # 祝日も更新する。
		msg = "全経過シートの休日も更新します。\n祝日も含みます。"
		componentwindow = controller.ComponentWindow
		msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_OK, "myRs", msg)
		if msgbox.execute()==MessageBoxResults.OK:			
			searchdescriptor = sheet.createSearchDescriptor()
			searchdescriptor.setSearchString("休日設定")  # 戻り値はない。
			c = VARS.templatestartcolumn - 1
			searchedcell = sheet[VARS.emptyrow:, c].findFirst(searchdescriptor)  # 見つからなかった時はNoneが返る。
			if searchedcell:
				startrow = searchedcell.getCellAddress().Row + 2
				cellranges = sheet[startrow:, c].queryContentCells(CellFlags.STRING+CellFlags.DATETIME)  # 休日設定列の文字列か日付が入っているセルに限定して抽出。
				emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 最終行インデックス+1を取得。		
				datarows = sheet[startrow:emptyrow, c].getDataArray()	
				weekdays = VARS.weekdays
				offweekdays = []
				offdays = []
				for datarow in datarows:	
					d = datarow[0]
					if isinstance(d, float):  # floatの時は日付シリアル値と考える。
						offdays.append(d)
					else:  # 文字列の時。
						offweekdays.append(weekdays.index(i) for i in d.replace("曜日", ""))
						
				holidays = commons.HOLIDAYS		
				offdayrangeaddresses = []		
				holidayrangeaddresses = []		
				
				dayrow = sheet[VARS.dayrow, :VARS.firstemptycolumn].getDataArray()[0]		
				startyear, startmonth, startday = [int(functionaccess.callFunction(i, (dayrow[VARS.datacolumn],))) for i in ("YEAR", "MONTH", "DAY")]
				endyear, endmonth, endday = [int(functionaccess.callFunction(i, (dayrow[-1],))) for i in ("YEAR", "MONTH", "DAY")]
				holidaycolumns = set()
				if startyear in holidays:
					for d in holidays[startyear][startmonth-1:]:
						if d>=startday:
							holidaycolumns.add(VARS.datacolumn+d-startday)
				newyear = startyear + 1
				newstartcoumn = int(functionaccess.callFunction("EOMONTH", (dayrow[VARS.datacolumn], 0)))
				while newyear<endyear:
					for d in holidays[newyear]:	
						holidaycolumns.add(newstartcoumn+d)
					newyear += 1	
					newstartcoumn = int(functionaccess.callFunction("EOMONTH", (dayrow[newstartcoumn], 0)))
				if endyear in holidays:
					for d in holidays[endyear][:endmonth]:
						if d<=endday:
							holidaycolumns.add(newstartcoumn+d)	
				columnindexes = set()
				columnindexes.update(dayrow.index(i) for i in offdays)
				startweekdaytxt = sheet[VARS.weekdayrow, VARS.datacolumn].getString()
				startweekday = weekdays.index(startweekdaytxt)
				columnindexes.update(range(VARS.datacolumn+(i-startweekday)%7, VARS.firstemptycolumn, 7) for i in offweekdays)
				
				# 祝日から日曜日を除く。
				
				columnindexes.difference_update(holidaycolumns)
				offdayrangeaddresses.extend(sheet[VARS.dayrow:VARS.dayrow+2, i].getRangeAddress() for i in columnindexes)
				
		
				
				
				weekdayrow = datarows[1]	
				
				
				columnindexes.extend(weekdayrow.index(i) for i in offweekdays)
						
					
						
						
						startdatevalue = sheet[VARS.dayrow, VARS.datacolumn].getValue()
						startdate = date(*[int(functionaccess.callFunction(i, (startdatevalue,))) for i in ("YEAR", "MONTH", "DAY")])



						colors = commons.COLORS
						n = 5  # 土曜日の曜日番号。
						columnindexes = range(datacolumn+(n-weekday)%7, endedgecolumn, 7)   # 土曜日の列インデックスを取得。	
	
							
						weekday = weekdays.index(d)
						weekdays[i%7] for i in range(weekday, weekday+daycount)
		
						n = weekdays.index(td)  # 月=0の曜日番号を取得。
						ws = range((n-weekday)%7, daycount, 7)  # 同じ曜日の相対インデックスを取得。		
		
		
		
		
	return False  # セル編集モードにしない。	
def createScheduleToClip(systemclipboard, times, startdate, outputs):  # times: 時間枠のリスト、startdate: 開始日のdateオブジェクト、outputs: 出力行のリスト。
	def scheduleToClip(n):  # n: 取得する日数。
		dategene = (startdate+timedelta(days=i) for i in range(n))
		weekdays = VARS.weekdays
		dates = ["{}/{}({})".format(i.month, i.day, weekdays[i.weekday()]) for i in dategene]
		for i in range(VARS.datacolumn, VARS.datacolumn+n):  # 列インデックスをイテレート。
			cellranges = VARS.sheet[VARS.datarow:VARS.emptyrow, i].queryEmptyCells()  # 空セルのセル範囲コレクションを取得。
			fs = [" ".join([times[j], "○"]) for i in cellranges.getRangeAddresses() for j in range(i.StartRow-VARS.datarow, i.EndRow+1-VARS.datarow)]
			if fs:
				outputs.extend(["", dates[i-VARS.datacolumn]])
				outputs.extend(fs)	
		systemclipboard.setContents(commons.TextTransferable("\r\n".join(outputs)), None)  # クリップボードにコピーする。	\rはWindowsのメモ帳で開業するため。
	return scheduleToClip
def wClickCell(enhancedmouseevent, xscriptcontext):
	defaultrows = "2F", "3F", "強", "新", "閉", "外", "会", "手", "ｸﾘｱ", "x", "/"
	staticdialog.createDialog(enhancedmouseevent, xscriptcontext, "予定", defaultrows, callback=callback_wClickCell)	
	return False  # セル編集モードにしない。	
def callback_wClickCell(mouseevent, xscriptcontext):	
	selection = xscriptcontext.getDocument().getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
	setCellProp(selection)
def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。	
	changes = changesevent.Changes	
	for change in changes:
		if change.Accessor=="cell-change":  # セルの値が変化した時。マクロで変更したときはセル範囲が入ってくる時がある。
			selection = change.ReplacedElement  # 値を変更したセルを取得。	
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
				celladdress = selection.getCellAddress()
				r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。		
				if VARS.datarow-1<r<VARS.emptyrow:
					if VARS.datacolumn-1<c<VARS.firstemptycolumn or VARS.templatestartcolumn-1<c<VARS.templateendcolumnedge:
						setCellProp(selection)
			break
def setCellProp(selection):		
	txt = selection.getString()	
	if txt:  # セルに文字列がある時。
		horijustify	= LEFT if len(txt)>2 else CENTER  # 文字数が2個までの時は中央揃えにする。
		selection.setPropertyValue("HoriJustify", horijustify)  
		color = "magenta3"
		if txt=="x":
			color = "gray7"
		elif txt=="/":
			color = "silver"
		selection.setPropertyValue("CellBackColor", commons.COLORS[color])
		if txt=="ｸﾘｱ":
			selection.clearContents(511)
	else:
		selection.setPropertyValues(("CellBackColor", "HoriJustify"), (-1, LEFT))		
def drowBorders(selection):  # ターゲットを交点とする行列全体の外枠線を描く。
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上端のセルアドレスを取得。
	r, c = celladdress.Row, celladdress.Column # selectionの行と列のインデックスを取得。		
	sheet = VARS.sheet
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = commons.createBorders()
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
	rangeaddress = selection.getRangeAddress() # 選択範囲のセル範囲アドレスを取得。
	if VARS.datarow-1<r<VARS.emptyrow:
		if VARS.datacolumn-1<c<VARS.firstemptycolumn:
			sheet[VARS.monthrow:VARS.emptyrow, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。	
			sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, VARS.datacolumn:VARS.firstemptycolumn].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。	
			selection.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	
		if VARS.templatestartcolumn-1<c<VARS.templateendcolumnedge:
			sheet[VARS.monthrow:VARS.emptyrow, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。	
			sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, VARS.templatestartcolumn:VARS.templateendcolumnedge].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。		
			selection.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。				
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	VARS.setSheet(sheet)
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
	del contextmenu[:]  # contextmenu.clear()は不可。
	selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上角のセルのアドレスを取得。
	r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。		
	if contextmenuname=="cell":  # セルのとき		
		if VARS.datarow-1<r<VARS.emptyrow:
			if VARS.datacolumn-1<c<VARS.firstemptycolumn or VARS.templatestartcolumn-1<c<VARS.templateendcolumnedge:
				if selection.supportsService("com.sun.star.sheet.SheetCell") and selection.getString() in ("", "強"):  # ターゲットがセル、かつ、空セルか強セルの時。
					addMenuentry("ActionTrigger", {"Text": "患者一覧", "CommandURL": baseurl.format("entry2")}) 
					addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				commons.cutcopypasteMenuEntries(addMenuentry)					
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				addMenuentry("ActionTrigger", {"Text": "クリア", "CommandURL": baseurl.format("entry1")}) 				
	elif contextmenuname=="rowheader" and len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 行ヘッダーのとき、かつ、選択範囲の列数がシートの列数が一致している時。	
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		commons.rowMenuEntries(addMenuentry)
	elif contextmenuname=="colheader" and len(selection[:, 0].getRows())==len(sheet[:, 0].getRows()):  # 列ヘッダーのとき、かつ、選択範囲の行数がシートの行数が一致している時。	
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertColumnsBefore"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertColumnsAfter"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteColumns"})
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。		
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	VARS.setSheet(sheet)
	selection = controller.getSelection()
	ichiransheet = doc.getSheets()["一覧"]
	if entrynum==1:  # クリア。	
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
			annotation = selection.getAnnotation()
			cells = getMendanCell(annotation.getString().split(" ")[0], ichiransheet),  # コメントが消える前にIDを取得して一覧シート上の面談セルを取得。
		else:  # 複数セルの時。
			commentcellgene = (i for i in sheet.getAnnotations() if len(selection.queryIntersection(i.getParent().getRangeAddress())))  # 選択セル範囲にあるコメントのあるセルのジェネレーター。	
			cells = (getMendanCell(i.getString().split(" ")[0], ichiransheet) for i in commentcellgene)	
		[i.clearContents(CellFlags.ANNOTATION) for i in cells if i is not None]  # 一覧シートのコメントを削除する。cellsにはNoneが入ってくるのでそれを除外する。
		selection.clearContents(511)  # 予定シートの選択範囲をすべてクリアする。コメントも消える。
	elif entrynum==2:  # 患者一覧。
		ichiranvars = ichiran.VARS		
		ichiranvars.setSheet(ichiransheet)
		ichirandatarows = ichiransheet[ichiranvars.splittedrow:ichiranvars.emptyrow, ichiranvars.idcolumn:ichiranvars.kanacolumn+1].getDataArray()
		ichirandatarows = sorted(ichirandatarows, key=lambda x: x[2])[3:]  # カナ列でソート。タイトル行は空欄なので先頭に来るのでインデックス3以降のみ取得。
		defaultrows = [" ".join(i) for i in ichirandatarows]
		transientdialog.createDialog(xscriptcontext, "患者一覧", defaultrows, fixedtxt="面", callback=callback_wClickGrid)
def callback_wClickGrid(mouseevent, xscriptcontext, gridcelldata):  # gridcelldata: グリッドコントロールのダブルクリックしたセルのデータ。
	idtxt = gridcelldata.split(" ")[0]  # グリッドコントロールのセルからIDを取得。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	sheet = doc.getCurrentController().getActiveSheet()
	selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
	annotations = sheet.getAnnotations()
	for i in annotations:  # すべてのコメントについて。
		if i.getString().startswith(idtxt):  # すでに同じIDのコメントが存在する時。
			msg = "{}にすでに面談予定がありますがそれを取り消しますか?".format(getCelldatetime(xscriptcontext, i.getPosition()))
			componentwindow = doc.getCurrentController().ComponentWindow
			msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_CANCEL, "myRs", msg)
			if msgbox.execute()==MessageBoxResults.YES:	
				cell = i.getParent()
				cell.clearContents(511)
				setCellProp(cell)		
			elif msgbox.execute()==MessageBoxResults.CANCEL:
				selection.setString("")  # 選択セルの文字列をクリア。
				setCellProp(selection)
				return
	setCellProp(selection)	
	celladdress = selection.getCellAddress()
	annotations.insertNew(celladdress, gridcelldata)  # gridcelldataをセル注釈を挿入。
	ichiransheet = doc.getSheets()["一覧"]
	cell = getMendanCell(idtxt, ichiransheet)  # 一覧シートのそのIDの面談列のセルを取得。
	if cell:	
		ichiransheet.getAnnotations().insertNew(cell.getCellAddress(), "{} 面談".format(getCelldatetime(xscriptcontext, celladdress))) 
		cell.setString("")  # 面談列の文字列をクリア。
	else:
		msg = "IDが一覧に見つかりません。"	
		componentwindow = doc.getCurrentController().ComponentWindow
		componentwindow.getToolkit().createMessageBox(componentwindow, ERRORBOX, MessageBoxButtons.BUTTONS_OK, "myRs", msg)	
def getCelldatetime(xscriptcontext, celladdress):
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。			
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。		
	datevalue = VARS.sheet[VARS.dayrow, celladdress.Column].getValue()
	md = [int(functionaccess.callFunction(i, (datevalue,))) for i in ("MONTH", "DAY")]
	timevalue = VARS.sheet[celladdress.Row, 0].getValue()
	hm = [int(functionaccess.callFunction(i, (timevalue,))) for i in ("HOUR", "MINUTE")]	
	return "{}/{} {}:{:0>2}".format(*md, *hm)
def getMendanCell(idtxt, ichiransheet):
	ichiranvars = ichiran.VARS		
	ichiranvars.setSheet(ichiransheet)
	searchdescriptor = ichiransheet.createSearchDescriptor()
	searchdescriptor.setSearchString(idtxt)  # 戻り値はない。	
	searchedcell = ichiransheet[ichiranvars.splittedrow:ichiranvars.emptyrow, ichiranvars.idcolumn].findFirst(searchdescriptor)  # 見つからなかった時はNoneが返る。
	if searchedcell:
		r = searchedcell.getCellAddress().Row
		searchdescriptor.setSearchString("面談")  # 戻り値はない。	
		searchedcell = ichiransheet[ichiranvars.menurow, ichiranvars.checkstartcolumn:ichiranvars.memostartcolumn].findFirst(searchdescriptor)  # 見つからなかった時はNoneが返る。
		if searchedcell:
			c = searchedcell.getCellAddress().Column
			return ichiransheet[r, c]	
