#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 経過シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import platform
from itertools import chain
from . import commons, historydialog, staticdialog, yotei
from com.sun.star.awt import MouseButton, MessageBoxButtons, MessageBoxResults, Key  # 定数
from com.sun.star.awt.MessageBoxType import QUERYBOX, WARNINGBOX  # enum
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.sheet.CellDeleteMode import ROWS as delete_rows  # enum
from com.sun.star.table.CellHoriJustify import CENTER, LEFT  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
class Keika():  # シート固有の定数設定。
	def __init__(self):
		self.dayrow = 1  # 日付行インデックス。
		self.splittedrow = 4  # 分割行インデックス。
		self.yakucolumn = 5  # 薬名列インデックス。
		self.splittedcolumn = 9  # 分割列インデックス。
	def setSheet(self, sheet):
		self.sheet = sheet
		cellranges = sheet[:, self.yakucolumn].queryContentCells(CellFlags.STRING)  # 薬名列の文字列が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 薬名列の最終行インデックス+1を取得。
		gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in (commons.COLORS["black"],))
		headers = next(gene, None), 
		if None in headers:  # Noneがある時。
			rownames = "黒",
			raise RuntimeError("{0}行が取得できません。\n{0}色の背景色のID列に何らかの文字列をいれてください。".format(rownames[headers.index(None)]))
		self.blackrow, = headers  # 黒行インデックスを取得。	
VARS = Keika()		
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet = activationevent.ActiveSheet
	sheet["F1:G1"].setDataArray((("一覧へ", "ｶﾙﾃへ"),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet["F3:F4"].setDataArray((("薬品順",), ("クリップボードから薬品名抽出",)))
	sheet["I3"].setString("透析")
	highlightToday(sheet, xscriptcontext)
def highlightToday(sheet, xscriptcontext):	# 今日の日付の上のセルをハイライトする。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。			
	dayrow = VARS.dayrow
	splittedcolumn = VARS.splittedcolumn
	startdatevalue = int(sheet[dayrow, splittedcolumn].getValue())  # 日付行の最初のセルから日付のシリアル値の取得。
	todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。
	sheet[dayrow-1, splittedcolumn:].setPropertyValue("CellBackColor", -1)  # r-1行目の背景色をクリア。
	c = splittedcolumn + (todayvalue - startdatevalue)  # 今日の日付の列インデックスを取得。
	if splittedcolumn<=c<1024:
		sheet[dayrow-1, c].setPropertyValue("CellBackColor", commons.COLORS["cyan10"])  # 日付行の上のセルの今日の背景色を設定。
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。		
	if enhancedmouseevent.ClickCount==2 and enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ダブルクリックの時。まずselectionChanged()が発火している。シングルクリックの時はselectionChanged()メソッドで事足りる。
		selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			celladdress = selection.getCellAddress()
			r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。	
			if r<VARS.splittedrow:  # 分割行より上、の時。
				if c<VARS.splittedcolumn:  # 分割列より左、の時。
					return wClickMenu(enhancedmouseevent, xscriptcontext)
				else: 
					return wClickUpperRight(enhancedmouseevent, xscriptcontext)
			elif r!=VARS.blackrow:  # 黒行でない時。
				if c>VARS.splittedcolumn-1:  # 分割行含む右列。
					return wClickBottomRight(enhancedmouseevent, xscriptcontext)
				elif c==VARS.yakucolumn:  # 薬名列の時。
					return True  # セル編集モードにする。
				else:	
					return wClickBottomLeft(enhancedmouseevent, xscriptcontext)
	return True  # セル編集モードにする。				
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
		idcelltxts = datarow[0].replace("　", " ").split(" ")  # 半角スペースで分割。
		idtxt = idcelltxts[0]  # 最初の要素を取得。
		if idtxt.isdigit():  # IDが数値のみの時。					
			if idtxt in sheets:  # ID名のシートがあるとき。
				controller.setActiveSheet(sheets[idtxt])  # カルテシートをアクティブにする。
			else:
				if len(idcelltxts)==5:  # ID、漢字姓・名、カタカナ姓・名、の5つに分割できていた時。
					kanjitxt, kanatxt = " ".join(idcelltxts[1:3]), " ".join(idcelltxts[3:])
					datevalue = datarow[-1]
					karutesheet = commons.getKaruteSheet(doc, idtxt, kanjitxt, kanatxt, datevalue)
					controller.setActiveSheet(karutesheet)  # カルテシートをアクティブにする。
				else:
					commons.showErrorMessageBox(controller, "「ID(数値のみ) 漢字姓 名 カナ姓 名」の形式になっていません。")
		else:
			commons.showErrorMessageBox(controller, "IDが取得できませんでした。")	
	elif txt=="薬品順":  # クリックするたびに終了順、昇順に並び替える。黒行の上のみ。
		if VARS.splittedrow<VARS.blackrow:  # 分割行から黒行より上に行がある時のみ。
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
	elif txt=="クリップボードから薬品名抽出":									
		systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。			
		clipboardtxt = commons.getClipboardtxt(systemclipboard)
		if not clipboardtxt:  # クリップボードの文字列が取得出来なかった時。メッセージボックスを出して終わる。
			msg = "クリップボードから文字列を取得できませんでした。"
			commons.showErrorMessageBox(controller, msg)
			return False  # セル編集モードにしない。		
		transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
		transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))			
		clipboardtxt = transliteration.transliterate(clipboardtxt, 0, len(clipboardtxt), [])[0]  # 半角に変換。数字すら全角のときがあるので。
		newdatarows = []
		yoho = ""
		sep = "\r\n" if platform.system()=="Windows" else "\n"  # Windowsでは\r\nで分割しないとうまく処理できない。
		for rowtxt in clipboardtxt.split(sep)[::-1]:  # 改行文字で分割して下の行からイテレート。
			if not rowtxt:
				continue
			elif rowtxt.startswith("点滴"):
				yoho = ""
				continue
			elif rowtxt.startswith(("20", "[", "CV", "ﾍﾟﾝﾆｰﾄﾞﾙ", "ﾋﾞﾀﾒｼﾞﾝ", "ﾌﾞﾄﾞｳ糖注50%PL", "生理食塩水PL")):
				continue
			elif "本人" in rowtxt:
				continue
			elif "家族" in rowtxt:
				continue				
			elif rowtxt.endswith(("日分",)):
				continue				
			elif rowtxt.endswith(("錠", "袋", "g", "本", "瓶", "管", "包", "枚", "個", "ｶﾌﾟｾﾙ", "ｷｯﾄ")):  # 特定の文字列で終わっている時は追加する。
				if not yoho:  # 点滴の時
					rowtxt = rowtxt.replace("1袋", "").replace(" ", "").replace("号輸液", " ")
				datarow = rowtxt.replace("  ", ""), yoho, ""
				if not datarow in newdatarows:
					newdatarows.append(datarow)
			elif rowtxt.endswith("単位"):
				if yoho:
					datarow = rowtxt.split(" ")[0], yoho, ""
				else:
					datarow = rowtxt.split(" ")[0], "", "混注"
				if not datarow in newdatarows:
					newdatarows.append(datarow)						
			else:			
				yoho = ""
				if rowtxt.startswith(("外用",)):
					yoho = "外用"
					if "吸入" in rowtxt:
						yoho = "吸入"
				elif rowtxt.startswith("分3"):
					yoho = "分3"
				elif rowtxt.startswith("分2"):
					yoho = "分2"					
				elif rowtxt.startswith("分1"):
					if "朝" in rowtxt:
						yoho = "朝"
					elif "昼" in rowtxt:
						yoho = "昼"
					elif "夕" in rowtxt:
						yoho = "夕"					
					elif "寝" in rowtxt:
						yoho = "寝"					
					elif "起床時" in rowtxt:
						yoho = "起床時"					
		if not newdatarows:	
			msg = "薬品名に変換できませんでした。"
			commons.showErrorMessageBox(controller, msg)			
			return False  # セル編集モードにしない。				
		newdatarows.reverse()  # 行を昇順に戻す。
		edgerow = VARS.emptyrow+len(newdatarows)		
		sheet[VARS.emptyrow:edgerow, VARS.yakucolumn:VARS.yakucolumn+3].setDataArray(newdatarows)  # 整理した薬品名をシートに代入。		
		sheet[VARS.emptyrow:edgerow, VARS.yakucolumn+1].setPropertyValue("HoriJustify", CENTER)			
	elif txt=="透析":
		celladdress = selection.getCellAddress()
		if selection.getPropertyValue("CharColor")==commons.COLORS["silver"]:
			selection.setPropertyValue("CharColor", commons.COLORS["black"])
			tosekibicell = sheet[celladdress.Row+1, celladdress.Column]
			tosekibicell.setString("月水金")
			tosekibicell.setPropertyValue("CharColor", commons.COLORS["black"])
		else:
			selection.setPropertyValue("CharColor", commons.COLORS["silver"])
			sheet[celladdress.Row+1, celladdress.Column].setPropertyValue("CharColor", commons.COLORS["white"])
	elif txt=="月水金":
		changeDialysDate(selection, "火木土")
	elif txt=="火木土":
		changeDialysDate(selection, "月水金")
	elif txt[:8].isdigit():  # 最初8文字が数値の時。						
		systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
		systemclipboard.setContents(commons.TextTransferable(txt[:8]), None)  # クリップボードにIDをコピーする。							
	return False  # セル編集モードにしない。	
def changeDialysDate(cell, daystxt):
	if cell.getPropertyValue("CharColor")==-1:
		cell.setPropertyValue("CharColor", commons.COLORS["red3"])
	else:
		cell.setPropertyValue("CharColor", -1)
		cell.setString(daystxt)	
def wClickUpperRight(enhancedmouseevent, xscriptcontext):
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	celladdress = selection.getCellAddress()
	r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。		
	if r==VARS.dayrow-1:  # 日付行の直上の時。月を入力。
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
		functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。							
		datevalue = int(VARS.sheet[VARS.dayrow, c].getValue())
		m = int(functionaccess.callFunction("MONTH", (datevalue,)))  # 月、を取得。
		selection.setString("{}月".format(m))
	elif r==VARS.dayrow+1:
		defaultrows = "○", "尿"
		staticdialog.createDialog(enhancedmouseevent, xscriptcontext, VARS.sheet[r, VARS.yakucolumn+1].getString(), defaultrows, callback=callback_wClickUpperRightCreator(xscriptcontext))  # 行タイトル毎に定型句ダイアログを作成。
	elif r==VARS.dayrow+2:
		defaultrows = chain(commons.GAZOs, commons.GAZOd, commons.SHOCHIs, commons.ECHOs)
		staticdialog.createDialog(enhancedmouseevent, xscriptcontext, VARS.sheet[r, VARS.yakucolumn+1].getString(), defaultrows, callback=callback_wClickUpperRightCreator(xscriptcontext))  # 行タイトル毎に定型句ダイアログを作成。
	return False  # セル編集モードにしない。	
def callback_wClickUpperRightCreator(xscriptcontext):	
	def callback_wClickUpperRight(gridcelltxt):	
		selection = xscriptcontext.getDocument().getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
		if gridcelltxt:  # セルに文字列がある時。
			horijustify	= LEFT if len(gridcelltxt)>1 else CENTER  # 文字数が1個の時は中央揃えにする。
			selection.setPropertyValues(("CellBackColor", "HoriJustify"), (commons.COLORS["skyblue"], horijustify))  # 背景をスカイブルーにする。		
		else:
			selection.setPropertyValue("CellBackColor", -1)  # 背景色を消す。	
	return callback_wClickUpperRight
def wClickBottomLeft(enhancedmouseevent, xscriptcontext):
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	celladdress = selection.getCellAddress()
	c = celladdress.Column  # selectionの列のインデックスを取得。		
	sheet = VARS.sheet
	defaultrows = None
	if c<VARS.yakucolumn:
		headertxt = sheet[1, c].getString()
		if headertxt=="検査値":
			defaultrows = "APTT 目標1.5倍:検査値",\
				"PT-INR 目標2.0-2.5:検査値",\
				"ｶﾙﾊﾞﾏｾﾞﾋﾟﾝ(ﾃｸﾞﾚﾄｰﾙ内服)3-10ug/ml10日で定常:検査値",\
				"ｸﾛﾅｾﾞﾊﾟﾑ(ﾗﾝﾄﾞｾﾝ内服)25-75ug/mlT1/2=27hr:検査値",\
				"ｼﾞｺﾞｷｼﾝ(ﾗﾆﾗﾋﾟｯﾄﾞ内服)0.7-1ng/mlT1/2=24hr:検査値",\
				"ｿﾞﾆｻﾐﾄﾞ(ｴｸｾｸﾞﾗﾝ内服)20ug/ml発汗障害注意1wで定常:検査値",\
				"ﾃｵﾌｨﾘﾝ(ﾈｵﾌｨﾘﾝ注)5-15ug/mlT1/2=9hr:検査値",\
				"ﾊﾞﾙﾌﾟﾛ酸(ﾃﾞﾊﾟｹﾝ内服)40-120ug/mlT1/2=10hrRは1wで定常:検査値",\
				"ﾊﾛﾍﾟﾘﾄﾞｰﾙ(ｾﾚﾈｰｽ注)3-10ng/mlT1/2=14hr:検査値",\
				"ﾌｪﾆﾄｲﾝ(ｱﾚﾋﾞｱﾁﾝ注)10-20ug/mlT1/2=10hr:検査値",\
				"ﾌｪﾉﾊﾞﾙﾋﾞﾀｰﾙ(ﾌｪﾉﾊﾞｰﾙ内服)10-25ug/ml2-3wで定常:検査値"
		elif headertxt=="その他":
			defaultrows = "包括ｹｱ:病棟", "廃用:ﾘﾊﾋﾞﾘ", "運動器:ﾘﾊﾋﾞﾘ", "呼吸器:ﾘﾊﾋﾞﾘ", "運動器:ﾘﾊﾋﾞﾘ"
		historydialog.createDialog(enhancedmouseevent, xscriptcontext, "{}_履歴".format(headertxt), defaultrows, VARS.yakucolumn, callback=callback_wClickBottomLeft0Creator(xscriptcontext))		
	else:
		r = celladdress.Row
		defaultrows = []
		if c==VARS.yakucolumn+1:  # 用法列。
			defaultrows = "分3", "分2", "朝", "昼", "夕", "寝", "朝寝", "分2朝寝", "分2朝昼", "吸入", "外用", "皮下注", "頓"
			staticdialog.createDialog(enhancedmouseevent, xscriptcontext, sheet[1, c].getString(), defaultrows, callback=callback_wClickBottomLeftCreator(xscriptcontext))	
		elif c==VARS.yakucolumn+2:  # 回数列。
			yoho = sheet[r, VARS.yakucolumn+1].getString()
			if yoho:
				if yoho=="吸入":
					defaultrows = "1吸入1日1回", "2吸入1日1回", "1吸入1日2回", "2吸入1日2回"
				elif yoho=="外用":
					defaultrows = "1日1回", "1日2回", "1日3回", "1日4回"
				elif yoho=="皮下注":
					defaultrows = "毎食前", "朝前", "夕前", "眠前"
				staticdialog.createDialog(enhancedmouseevent, xscriptcontext, yoho, defaultrows, callback=callback_wClickBottomLeftCreator(xscriptcontext))
			else:
				defaultrows = "持続", "1回", "2回", "3回"
				staticdialog.createDialog(enhancedmouseevent, xscriptcontext, sheet[1, c].getString(), defaultrows, callback=callback_wClickBottomLeftCreator(xscriptcontext))	
		elif c==VARS.yakucolumn+3:  # 限定列。
			dialogtitle = sheet[1, c].getString()
			weekdays = "月火水木金土日"
			defaultrows = ["2日に1回", "3日に1回", "7日に1回", "月木", "火金"]
			if sheet[2, VARS.yakucolumn+3].getPropertyValue("CharColor")==commons.COLORS["black"]:  # 透析患者の時。
				tosekibi = sheet[3, VARS.yakucolumn+3].getString()  # 透析日を取得。
				table = str.maketrans(tosekibi, " "*len(tosekibi))  # 透析日を半角スペースに置換するテーブル。
				nontosekibi = weekdays.translate(table).replace(" ", "")  # 透析日以外。
				nontosekibizenjitu = "{}{}".format(tosekibi, "土" if tosekibi.startswith("月") else "日")  # 透析日前日以外
				defaultrows.extend(["{}(透析日のみ)".format(tosekibi), "{}(透析日以外)".format(nontosekibi), "{}(透析日前日以外)".format(nontosekibizenjitu)])
				dialogtitle = "{}透析日".format(tosekibi)
			defaultrows.extend(weekdays)
			staticdialog.createDialog(enhancedmouseevent, xscriptcontext, dialogtitle, defaultrows, callback=callback_wClickBottomLeftCreator(xscriptcontext))
	return False  # セル編集モードにしない。
def callback_wClickBottomLeft0Creator(xscriptcontext):
	def callback_wClickBottomLeft0(gridcelltxt):
		selection = xscriptcontext.getDocument().getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
		r = selection.getCellAddress().Row
		sheet = VARS.sheet
		if ":" in gridcelltxt:
			txts = gridcelltxt.split(":"),
			columnlength = len(txts[0])
			if columnlength<VARS.splittedcolumn-VARS.yakucolumn+1:
				sheet[r, VARS.yakucolumn:VARS.yakucolumn+columnlength].setDataArray(txts)
		if gridcelltxt.endswith(":検査値"):
			sheet[selection.getCellAddress().Row, VARS.splittedcolumn:].setPropertyValue("NumberFormat", commons.formatkeyCreator(xscriptcontext.getDocument())('@'))  # 書式を設定。 
	return callback_wClickBottomLeft0
def callback_wClickBottomLeftCreator(xscriptcontext):
	def callback_wClickBottomLeft(gridcelltxt):
		selection = xscriptcontext.getDocument().getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
		if gridcelltxt:  # セルに文字列がある時。
			horijustify	= LEFT if len(gridcelltxt)>2 else CENTER  # 文字数が2個までの時は中央揃えにする。
			selection.setPropertyValue("HoriJustify", horijustify)  
			if gridcelltxt=="皮下注":
				VARS.sheet[selection.getCellAddress().Row, VARS.splittedcolumn:].setPropertyValue("NumberFormat", commons.formatkeyCreator(xscriptcontext.getDocument())('@'))  # 書式を設定。 
	return callback_wClickBottomLeft
def wClickBottomRight(enhancedmouseevent, xscriptcontext):
	r = enhancedmouseevent.Target.getCellAddress().Row
	yoho = VARS.sheet[r, VARS.yakucolumn+1].getString()
	callback_wClickBottomRight = callback_wClickBottomRightCreator(xscriptcontext)
	if yoho:
		if yoho in ("吸入"):
			defaultrows = "朝", "昼", "夕", "寝", "処方"
			staticdialog.createDialog(enhancedmouseevent, xscriptcontext, yoho, defaultrows, callback=callback_wClickBottomRight)
		elif yoho in ("皮下注"):	
			defaultrows = "処方", "4-4-4", "4"
			staticdialog.createDialog(enhancedmouseevent, xscriptcontext, yoho, defaultrows, callback=callback_wClickBottomRight)
		else:
			defaultrows = "朝", "昼", "夕", "寝"
			staticdialog.createDialog(enhancedmouseevent, xscriptcontext, "処方", defaultrows, callback=callback_wClickBottomRight)
	else:  # 用法列が空セルの時は点滴とする。
		defaultrows = "朝", "昼", "夕", "1A", "2A", "3A", "4A", "5ml/hr"
		staticdialog.createDialog(enhancedmouseevent, xscriptcontext, "点滴", defaultrows, callback=callback_wClickBottomRight)
	return False  # セル編集モードにしない。
def callback_wClickBottomRightCreator(xscriptcontext):
	def callback_wClickBottomRight(gridcelltxt):	
		sheet = VARS.sheet
		selection = xscriptcontext.getDocument().getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
		if gridcelltxt in ("止", "変"):  # 代入したセルの背景色を消し、それより右を全て消し黒行より下なら、黒行の上に移動する。
			endPrescription(selection, gridcelltxt)
		elif gridcelltxt=="処方":
			selection.setString("")
			selection.setPropertyValue("CellBackColor", commons.COLORS["magenta3"])
		elif gridcelltxt:  # 上記以外の文字列の時。
			horijustify	= LEFT if len(gridcelltxt)>1 else CENTER  # 文字数が1個の時は中央揃えにする。
			if selection.getPropertyValue("CellBackColor")==-1:  # 背景色がまだない時。
				color = "lime" if sheet[selection.getCellAddress().Row, VARS.yakucolumn+1].getString() else "magenta3"  # 用法列に文字列がなければ点滴とする。
				selection.setPropertyValues(("CellBackColor", "HoriJustify"), (commons.COLORS[color], horijustify))  
			else:	
				selection.setPropertyValue("HoriJustify", horijustify)
		else:  # 文字列がない時。
			selection.setPropertyValue("CellBackColor", -1)  # 背景色を消す。	
	return callback_wClickBottomRight	
def endPrescription(cell, txt):
	cell.setString(txt)
	celladdress = cell.getCellAddress()
	r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。	
	cell.setPropertyValues(("CellBackColor", "HoriJustify"), (-1, CENTER))  # 背景を消して中央揃えにする。		
	VARS.sheet[r, c+1:].clearContents(511)
	if r>VARS.blackrow:  # 黒行より下の時。
		rangeaddress = cell.getRangeAddress()  # 選択範囲のアドレスを取得。
		commons.toOtherEntry(VARS.sheet, rangeaddress, VARS.emptyrow, VARS.blackrow)  # 黒行の上へ移動。
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	selection = eventobject.Source.getSelection()  # 必ずしもセル範囲とは限らない。
	if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
		VARS.setSheet(selection.getSpreadsheet())		
		detectDuplicates(selection, xscriptcontext)  # 薬名の重複をチェック。drowBorders()はこの中で実行。選択範囲が変わるのでdrowBorders()が実行される。
	elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。	
		VARS.setSheet(selection.getSpreadsheet())	
		drowBorders(selection)  # 枠線の作成。
def detectDuplicates(selection, xscriptcontext):  # 薬名の重複をチェック。	
	celladdress = selection.getCellAddress()	
	emptyrow = VARS.emptyrow  # selectionChanged()はselect()を使うと発火して最終行が更新されるのでここで取得しておく。	
	if celladdress.Row>=emptyrow and celladdress.Column==VARS.yakucolumn:   # 空行以下、かつ、薬列の時。
		controller = xscriptcontext.getDocument().getCurrentController()		
		indicator = controller.getStatusIndicator()  # メッセージボックスではマウス選択状態のままになってしまうのでステータスインジケーターを使う。		
		indicator.start("重複行をチェック。", 0)  # ステータスバーに表示する文字列とプログレスバーの目盛りを設定。setValue()が1回ではバーが表示されない模様。	
		sheet = VARS.sheet
		splittedrow = VARS.splittedrow
		datarange = sheet[splittedrow:emptyrow, VARS.yakucolumn:VARS.splittedcolumn]
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。			
		transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
		transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))		
		datarows = []  # 半角に変換したデータ行を入れるタプル。
		for datarow in datarange.getDataArray():  # 薬品列から固定列の左列まで行をイテレート。
			txt = transliteration.transliterate(datarow[0], 0, len(datarow[0]), [])[0]  # 薬名を半角に変換	
			newdatarow = txt, *datarow[1:]	
			datarows.append(newdatarow)
		datarange.setDataArray(datarows)  # シートに書き戻す。
		idxes = []  # 削除する行インデックスのリスト。
		datarowlength = len(datarows)
		newemptyrow = emptyrow
		for i in set(datarows):  # ユニークなデータ行をイテレート。
			if datarows.count(i)>1:  # 重複データ行がある時。
				j = 0  # 重複行インデックスを初期化。
				while j<datarowlength and i in datarows[j:]:  # インデックスj以降に行iが存在する間。
					j += datarows[j:].index(i)  # 重複行インデックスを取得。
					sourceidx = splittedrow + j
					sheet.moveRange(sheet[newemptyrow, 0].getCellAddress(), sheet[sourceidx, :].getRangeAddress())  # 行の内容を最下行に移動。	
					idxes.append(sourceidx)  # 移動元インデックスを、削除する行インデックスのリストに追加。
					newemptyrow += 1  # 最下行の空行を更新。
					j += 1  # 検索開始行インデックスを更新。
		if idxes:		
			rowc = len(idxes)	
			idxes.sort(reverse=True)  # 削除する行インデックスを降順にソート。昇順に並んでいるわけではないので[::-1]は不可。
			for i in idxes:
				sheet.removeRange(sheet[i, :].getRangeAddress(), delete_rows)  # 移動したソース行を削除。
			controller.select(sheet[emptyrow-rowc:emptyrow, VARS.yakucolumn:VARS.splittedcolumn])  # 移動させた行を選択状態にする。
		indicator.end()  # reset()の前にend()しておかないと元に戻らない。
		indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。		
		if idxes:
			return  # 選択範囲が変わってdrowBorders()が実行されているので、ここで終わる。
	drowBorders(selection)  # 枠線の作成。
def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。	
	selection = None
	for change in changesevent.Changes:
		if change.Accessor=="cell-change":  # セルの値が変化した時。
			selection = change.ReplacedElement  # 値を変更したセルを取得。セル範囲が返るときもある。
			break
	if selection:  # 背景色をペーストしても発火するのでセル範囲が膨大になるときがある。		
		cellranges = selection.queryContentCells(CellFlags.STRING+CellFlags.DATETIME+CellFlags.VALUE+CellFlags.FORMULA)  # 内容のあるセルのみのセル範囲コレクションを取得。
		if cellranges:
			skybluecells = []  # 背景色をスカイブルーにするセルのリスト。
			colorlesscells = []  # 背景色を無色にするセルのリスト。
			leftcells = []  # 左寄せにするセルのリスト。
			centercells = []  # 中央寄せにするセルのリスト。	
			sheet = selection.getSpreadsheet()
			dayrow = VARS.dayrow
			splittedrow = VARS.splittedrow
			splittedcolumn = VARS.splittedcolumn					
			for rangeaddress in cellranges.getRangeAddresses():
				for r in range(rangeaddress.StartRow, rangeaddress.EndRow+1):  # 行インデックスについてイテレート。				
					for c in range(rangeaddress.StartColumn, rangeaddress.EndColumn+1):  # 列インデックスについてイテレート。			
						if r<=dayrow or (c<splittedcolumn and r<splittedrow):  # 日付行を含む上行、または、左上枠、の時は何もしない。
							continue
						cell = sheet[r, c]  # セルを取得。
						txt = cell.getString()
						if txt:  # セルに文字列がある時のみ。	
							stringlength = 2 if c<splittedcolumn else 1  # 分割列より左の時は文字数を2個、それ以外は1個に設定。
							if len(txt)>stringlength:
								leftcells.append(cell)  # 左揃えにする。				
							else:
								centercells.append(cell)  # 中央揃えにする。				
							if r<splittedrow:  # 右上枠で日付行より下、かつ、文字がある、時。
								skybluecells.append(cell)  # 背景色を設定。
						elif r<splittedrow:  # 右上枠で日付行より下、かつ、文字がない、時。
							colorlesscells.append(cell)  # 背景色を消す。			
			doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
			setRangeProp(doc, skybluecells, "CellBackColor", commons.COLORS["skyblue"])
			setRangeProp(doc, colorlesscells, "CellBackColor", -1)
			setRangeProp(doc, leftcells, "HoriJustify", LEFT)
			setRangeProp(doc, centercells, "HoriJustify", CENTER)
def setRangeProp(doc, ranges, propname, propvalue):  # datarangeは問題リストの#を検索するセル範囲。
	if ranges:
		cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
		cellranges.addRangeAddresses([i.getRangeAddress() for i in ranges], False)  # rangesの要素がないとエラーになる。
		cellranges.setPropertyValue(propname, propvalue)						
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。	
	contextmenuname, addMenuentry, baseurl, selection = commons.contextmenuHelper(VARS, contextmenuexecuteevent, xscriptcontext)
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上角のセルのアドレスを取得。
	r, c = celladdress.Row, celladdress.Column  # selectionの左上角セルの行と列のインデックスを取得。		
	if contextmenuname=="cell":  # セルのとき。複数セルも含まれる。	
		if r<VARS.splittedrow:  # 分割行より上の時。
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 単一セルの時。
				if c>VARS.splittedcolumn-1:  # 分割列含む右の時。
					if r==VARS.dayrow:  # 日付行の時。
						if selection.getValue():  # セルに値があるとき。
							addMenuentry("ActionTrigger", {"Text": "日付追加", "CommandURL": baseurl.format("entry3")}) 
			if r>VARS.dayrow:  # 日付行より下の時
				commons.cutcopypasteMenuEntries(addMenuentry)					
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				addMenuentry("ActionTrigger", {"Text": "クリア", "CommandURL": baseurl.format("entry4")}) 
		elif r!=VARS.blackrow:  # 黒行以外の時。
			if c>VARS.splittedcolumn-1:  # 分割列を含む右列の時。
				sheetcell = selection.supportsService("com.sun.star.sheet.SheetCell")  # 単一セルの時True。
				yoho = VARS.sheet[r, VARS.yakucolumn+1].getString()
				if sheetcell and yoho in ("ﾘﾊﾋﾞﾘ", "病棟"):  # 単一セルかつ用法列がリハビリまたは病棟の時。
					addMenuentry("ActionTrigger", {"Text": "開始", "CommandURL": baseurl.format("entry24")})
				elif sheetcell and yoho in ("検査値",):  # 単一セルかつ用法列が検査値の時。
					addMenuentry("ActionTrigger", {"Text": "未", "CommandURL": baseurl.format("entry27")})
				else:
					if VARS.sheet[r, c-1].getPropertyValue("CellBackColor")>0 or selection[0, 0].getString() or c==VARS.splittedcolumn:  # 選択セル範囲の左上セルの左のセルに背景色がある、または、左上セルに文字列がある、または、開始列、の時。
						addMenuentry("ActionTrigger", {"Text": "止", "CommandURL": baseurl.format("entry5")})
						addMenuentry("ActionTrigger", {"Text": "変", "CommandURL": baseurl.format("entry6")})
						addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
					addMenuentry("ActionTrigger", {"Text": "処方", "CommandURL": baseurl.format("entry7")})
					if sheetcell and yoho=="皮下注":  # 単一セルかつ用法列が皮下注の時。
						addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
						addMenuentry("ActionTrigger", {"Text": "4-4-4", "CommandURL": baseurl.format("entry11")})		
						addMenuentry("ActionTrigger", {"Text": "インスリン残計算", "CommandURL": baseurl.format("entry22")})					
					elif sheetcell and yoho=="吸入":  # 単一セルかつ用法列が吸入の時。
						addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。	
						addMenuentry("ActionTrigger", {"Text": "7日間", "CommandURL": baseurl.format("entry8")})
						addMenuentry("ActionTrigger", {"Text": "30日間", "CommandURL": baseurl.format("entry12")})
					else:
						addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。	
						addMenuentry("ActionTrigger", {"Text": "7日間", "CommandURL": baseurl.format("entry8")})
						addMenuentry("ActionTrigger", {"Text": "翌週まで", "CommandURL": baseurl.format("entry9")})
						addMenuentry("ActionTrigger", {"Text": "翌々週まで", "CommandURL": baseurl.format("entry2")})
						addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
						addMenuentry("ActionTrigger", {"Text": "翌月まで", "CommandURL": baseurl.format("entry10")})
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				addMenuentry("ActionTrigger", {"Text": "値のみクリア", "CommandURL": baseurl.format("entry25")}) 			
				addMenuentry("ActionTrigger", {"Text": "以後消去", "CommandURL": baseurl.format("entry14")})
				addMenuentry("ActionTrigger", {"Text": "クリア", "CommandURL": baseurl.format("entry4")}) 
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。	
				commons.cutcopypasteMenuEntries(addMenuentry)	
			else:
				if c==VARS.yakucolumn:  # 薬名列の時。
					if selection.getPropertyValue("CellBackColor")==-1:
						addMenuentry("ActionTrigger", {"Text": "他科処方", "CommandURL": baseurl.format("entry13")}) 
					else:
						addMenuentry("ActionTrigger", {"Text": "他科処方クリア", "CommandURL": baseurl.format("entry26")}) 
					addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				commons.cutcopypasteMenuEntries(addMenuentry)					
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				addMenuentry("ActionTrigger", {"Text": "クリア", "CommandURL": baseurl.format("entry4")}) 		
	elif contextmenuname=="rowheader" and len(selection[0, :].getColumns())==len(VARS.sheet[0, :].getColumns()):  # 行ヘッダーのとき、かつ、選択範囲の列数がシートの列数が一致している時。	
		if r>VARS.splittedrow-1:
			if r<VARS.blackrow:
				addMenuentry("ActionTrigger", {"Text": "黒行下へ", "CommandURL": baseurl.format("entry15")})  # 黒行上から使用中最上行へ
				addMenuentry("ActionTrigger", {"Text": "最下行へ", "CommandURL": baseurl.format("entry16")})  # 黒行上から使用中最下行へ
			elif r>VARS.blackrow:  # 黒行以外の時。
				addMenuentry("ActionTrigger", {"Text": "黒行上へ", "CommandURL": baseurl.format("entry17")})  
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
				addMenuentry("ActionTrigger", {"Text": "黒行下へ", "CommandURL": baseurl.format("entry18")})  # 使用中から使用中最上行へ  
				addMenuentry("ActionTrigger", {"Text": "最下行へ", "CommandURL": baseurl.format("entry19")})  # 使用中から使用中最下行へ		
			if r!=VARS.blackrow:
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
				commons.cutcopypasteMenuEntries(addMenuentry)
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
				commons.rowMenuEntries(addMenuentry)
	elif contextmenuname=="colheader" and len(selection[:, 0].getRows())==len(VARS.sheet[:, 0].getRows()):  # 列ヘッダーのとき、かつ、選択範囲の行数がシートの行数が一致している時。	
		if c>VARS.splittedcolumn and len(selection[0, :].getColumns())==1:  # 分割列を含まない右列、かつ、選択列数が1つの時。
			if VARS.sheet[VARS.blackrow, c].getPropertyValue("CellBackColor")==commons.COLORS["black"]:  # 黒行のセルの背景色が黒の時。
				addMenuentry("ActionTrigger", {"Text": "退院翌日", "CommandURL": baseurl.format("entry20")}) 
			else:
				addMenuentry("ActionTrigger", {"Text": "退院取消", "CommandURL": baseurl.format("entry21")})
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Remove"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:RenameTable"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	selection = controller.getSelection()
	if entrynum==2:  # 翌々週まで。
		colorizeSelectionRange(xscriptcontext, selection, "2w")
	elif entrynum==3:  # 日付追加。selectionは単一セル。	
		setDates(xscriptcontext, doc, sheet, selection, int(selection.getValue()))  # 経過シートの日付を設定。
		if int(selection.getString())!=1:  # 日付が１日でない時。
			celladdress = selection.getCellAddress()  # 選択セルアドレスを取得。
			r, c = celladdress.Row, celladdress.Column
			if c!=VARS.splittedcolumn:  # 固定列でないとき。
				sheet[r-1, c].setString("")  # 選択セルの上のセルの文字列を消す。
	elif entrynum==4:  # クリア。書式設定とオブジェクト以外を消去。
		rangeaddress = selection.getRangeAddress()  # 選択範囲のアドレスを取得。
		titlerows = VARS.blackrow,
		for i in range(rangeaddress.StartRow, rangeaddress.EndRow+1):  # 選択範囲の行インデックスをイテレート。
			for j in range(rangeaddress.StartColumn, rangeaddress.EndColumn+1):  # 選択範囲の列インデックスをイテレート。
				if i in titlerows or i<=VARS.dayrow:  # 黒行、または、日付行以上、の時。
					continue
				elif j<VARS.splittedcolumn and i<VARS.splittedrow:  # 左上枠の時
					continue
				else:
					sheet[i, j].clearContents(511)  # 範囲をすべてクリアする。		
	elif entrynum in (5, 6):  # 止 変
		txt = "止" if entrynum==5 else "変"
		rangeaddress = selection.getRangeAddress()
		c = rangeaddress.StartColumn
		for r in range(rangeaddress.StartRow, rangeaddress.EndRow+1):
			endPrescription(sheet[r, c], txt)
	elif entrynum==7:  # 処方。selectionは単一セルか複数セル。
		colorizeSelectionRange(xscriptcontext, selection)
	elif entrynum==8:  # 7日間。selectionは単一セルか複数セル。
		rangeaddress = selection.getRangeAddress()
		colorizeSelectionRange(xscriptcontext, sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, rangeaddress.StartColumn:rangeaddress.StartColumn+7])		
	elif entrynum==9:  # 翌週まで。selectionは単一セルか複数セル。
		colorizeSelectionRange(xscriptcontext, selection, "w")
	elif entrynum==10:  # 翌月まで。selectionは単一セルか複数セル。
		colorizeSelectionRange(xscriptcontext, selection, "m")
	elif entrynum==11:  # 4-4-4。	
		selection.setString("4-4-4")
		selection.setPropertyValue("CellBackColor", commons.COLORS["magenta3"])
	elif entrynum==12:  # 30日観。
		rangeaddress = selection.getRangeAddress()
		colorizeSelectionRange(xscriptcontext, sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, rangeaddress.StartColumn:rangeaddress.StartColumn+30])	
	elif entrynum==13:  # 他科処方。
		selection.setPropertyValue("CellBackColor", commons.COLORS["cyan10"])		
	elif entrynum==14:  # 以後消去。selectionは単一セルか複数セル。		
		msg = "選択セルから右をすべてクリアします。"
		componentwindow = controller.ComponentWindow
		msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_OK, "myRs", msg)
		if msgbox.execute()==MessageBoxResults.OK:		
			rangeaddress = selection.getRangeAddress()
			sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, rangeaddress.StartColumn:].clearContents(511)
	elif 14<entrynum<20:
		rangeaddress = selection.getRangeAddress()  # 選択範囲のアドレスを取得。
		if entrynum==15:  # 黒行上から使用中最上行へ
			commons.toOtherEntry(sheet, rangeaddress, VARS.blackrow, VARS.blackrow+1)
		elif entrynum==16:  # 黒行上から使用中最下行へ
			commons.toNewEntry(sheet, rangeaddress, VARS.blackrow, VARS.emptyrow) 
		elif entrynum==17:  # 黒行上へ
			commons.toOtherEntry(sheet, rangeaddress, VARS.emptyrow, VARS.blackrow)  
		elif entrynum==18:  # 使用中から使用中最上行へ 
			commons.toOtherEntry(sheet, rangeaddress, VARS.emptyrow, VARS.blackrow+1)
		elif entrynum==19:  # 使用中から使用中最下行へ		
			commons.toNewEntry(sheet, rangeaddress, VARS.emptyrow, VARS.emptyrow) 
	elif entrynum==20:  # 退院翌日
		selection[VARS.dayrow+1:VARS.emptyrow+100, :].setPropertyValue("CellBackColor", commons.COLORS["skyblue"])  # 固定行より下すべてに色を付ける(時間がかるので最終行下100行までにする)。
	elif entrynum==21:  # 退院取消
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
		dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
		docframe = controller.getFrame()
		c = selection[0, 0].getCellAddress().Column  # 選択セル範囲の一番上のセルの列インデックスを取得。
		controller.select(sheet[VARS.dayrow+1:, c-1])  # 選択列の左の列を選択。
		dispatcher.executeDispatch(docframe, ".uno:Copy", "", 0, ())  # コピー。
		controller.select(sheet[VARS.dayrow+1:, c])  # 元の列を選択し直す。
		nvs = ("Flags", "T"),\
			("FormulaCommand", 0),\
			("SkipEmptyCells", False),\
			("Transpose", False),\
			("AsLink", False),\
			("MoveMode", 4)
		props = [PropertyValue(Name=n, Value=v) for n, v in nvs]
		dispatcher.executeDispatch(docframe, ".uno:InsertContents", "", 0, props)  # 書式のみをペースト。changesOccurred()が発火する。ソースのセル範囲の枠が動く破線のままになるのでEscキーをシミュレートする必要がある。
		commons.simulateKey(controller, Key.ESCAPE, chr(0x1b))  # Escキーをシミュレート。
	elif entrynum==22:  # インスリン残計算。選択セルは単一。
		u = 300  # 1本単位。
		e = 2  # 空打ち単位。
		celladdress = selection[0, 0].getCellAddress()
		r, c = celladdress.Row, celladdress.Column
		color = commons.COLORS["magenta3"]  # インスリン開始セルの背景色。
		for i in range(VARS.splittedcolumn, c+1)[::-1]:  # 選択セルから左に列インデックスをイテレート。
			if sheet[r, i].getPropertyValue("CellBackColor")==color:  # 背景色が開始セルの時。
				startindex = i  # 選択セルを含む左の最初のインスリン開始セルの列インデックスを取得。
				break
		else:  # 開始セルが取得出来なかった時。
			msg = "開始セルが取得できませんでした。"		
			commons.showErrorMessageBox(controller, msg)
			return
		cellranges = sheet[r, VARS.splittedcolumn:c+u//e].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # 空打ちだけの最大列インデックスまでの範囲で文字列または値のあるセル範囲コレクションを取得。	
		unitgene = (i for rangeaddress in cellranges.getRangeAddresses() for i in range(rangeaddress.StartColumn, rangeaddress.EndColumn+1))  # 文字列または値のある列インデックスを返すジェネレーター。
		j = 0  # 開始列インデックスを越える前の列インデックスを入れる変数。
		for i in unitgene:
			if i>startindex:  # 開始列より右の時。 開始列インデックスかその前にあるインスリン量の記載のあるセルの列インデックスになる。
				break
			j = i
		if j>0:	
			startunits = sheet[r, j].getString()  # その前の列インデックスにある文字列を取得。これが開始時のインスリン量のはず。4-4-4とか4とかで返ってくるのを-で分割する。
			dayu = sum([int(i)+e for i in startunits.split("-") if i.isdigit() and int(i)>0])  # インスリンの1日消費量を取得。数字、かつ、0より大きい時、空打ち量を加えて合計する。	
			if dayu>0:  # 開始時のインスリン量が取得出来た時。
				edgecolumn = startindex + u//dayu  # インスリンがなくなる日の右列インデックスを取得。
				unitgene = (i for rangeaddress in cellranges.getRangeAddresses() for i in range(rangeaddress.StartColumn, rangeaddress.EndColumn+1))  # 文字列のある列インデックスを返すジェネレーターを再作。
				for i in (i for i in unitgene if i>startindex):  # 開始時の次のインスリン量について。
					newdayu = sum([int(i)+e for i in sheet[r, i].getString().split("-") if int(i)>0])  # 新たな1日消費量を取得。
					edgecolumn = i + dayu*(edgecolumn-i)//newdayu  # 残日数から残インスリン量を取得して新しい1日消費量で残日数を再計算。
					dayu = newdayu  # 1日消費量を更新。
				sheet[r, edgecolumn:].setPropertyValue("CellBackColor", -1) 
				sheet[r, startindex+1:edgecolumn].setPropertyValue("CellBackColor", commons.COLORS["lime"]) 	
				return
		msg = "開始時のインスリン量が取得できません。"		
		commons.showErrorMessageBox(controller, msg)
	elif entrynum==24:  # 開始。	
		celladdress = selection[0, 0].getCellAddress()
		r, c = celladdress.Row, celladdress.Column
		yoho = sheet[r, VARS.yakucolumn+1].getString()
		if yoho=="病棟":
			sheet[r, VARS.splittedcolumn:].clearContents(511)
			sheet[r, c:c+60].setPropertyValue("CellBackColor", commons.COLORS["skyblue"]) 	
			datevalue = int(sheet[VARS.dayrow, c].getValue())
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
			functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。
			datevalue += 59
			enddate = "/".join([str(int(functionaccess.callFunction(i, (datevalue,)))) for i in ("YEAR", "MONTH", "DAY")])
			selection.setString("{} 終了".format(enddate))
			commons.toOtherEntry(sheet, selection.getRangeAddress(), VARS.emptyrow, VARS.blackrow+1)  # 黒行下へ移動。
		elif yoho=="ﾘﾊﾋﾞﾘ":
			sheet[r, c:].clearContents(511)
			sheet[r, c:c+30].setPropertyValue("CellBackColor", commons.COLORS["skyblue"]) 
			commons.toOtherEntry(sheet, selection.getRangeAddress(), VARS.emptyrow, VARS.blackrow+1)  # 黒行下へ移動。
	elif entrynum==25:  # 値のみクリア。書式設定とオブジェクト以外を消去。
		selection.clearContents(CellFlags.VALUE+CellFlags.DATETIME+CellFlags.STRING+CellFlags.ANNOTATION+CellFlags.FORMULA)		
	elif entrynum==26:  # 他科処方クリア。
		selection.setPropertyValue("CellBackColor", -1)	
	elif entrynum==27:  # 未。
		selection.setString("未")
def colorizeSelectionRange(xscriptcontext, selection, end=None):  # endが与えられている時はselectionは選択行だけが意味を持つ。
	rangeaddress = selection.getRangeAddress()
	startc = rangeaddress.StartColumn
	endc = rangeaddress.EndColumn
	sheet = VARS.sheet
	celladdress = selection[0, 0].getCellAddress()  # 選択セル左上端セルのアドレスを取得。
	r, c = celladdress.Row, celladdress.Column		
	datevalue = int(sheet[VARS.dayrow, c].getValue())
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。		
	weekdayval = int(functionaccess.callFunction("WEEKDAY", (datevalue,)))  # 選択範囲の最初の日付のシリアル値から曜日の数字を取得。日曜日=1。
	yakurows = sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, VARS.yakucolumn:VARS.splittedcolumn].getDataArray()  # (薬名、用法、回数、限定)のタプル。
	naifukurangeaddress = []
	tentekirangeaddress = []
	table = str.maketrans("日月火水木金土", "1234567")  # 曜日をシート関数WEEKDAY()の戻り値の数字に変換するテーブル。
	if end is not None:  # 終了日が指定されている時。
		n = 6  # 内服の終了曜日。6:金曜日。
		t = 3  # 点駅の終了曜日。3:火曜日。
		if end=="w":  # 翌週の時。翌週の指定曜日まで。
			nendc, tendc = getNewEndC(startc, weekdayval, n, t, 0)
		elif end=="2w":  # 翌々週の時。翌々週の指定曜日まで。
			nendc, tendc = getNewEndC(startc, weekdayval, n, t, 7)
		elif end=="m":  # 翌月の時。翌月の指定曜日まで。
			newdatevalue = int(functionaccess.callFunction("EOMONTH", (datevalue, 0))) + 8  # 翌月1日の1週間後の日付シリアル値を取得。
			newweekdayval = int(functionaccess.callFunction("WEEKDAY", (newdatevalue,)))  # 日付のシリアル値から曜日の数字を取得。日曜日=1。
			newendc = startc + newdatevalue - datevalue
			ndiff = n - newweekdayval
			if ndiff<0:  # 負数なら1週間繰り越す・
				ndiff += 7
			nendc = newendc + ndiff  # 内服用。
			tdiff = t - newweekdayval
			if tdiff<0:  # 負数なら1週間繰り越す・
				tdiff += 7			
			tendc = newendc + tdiff  # 点滴用。
	yohos = "頓", "屯", "吸入", "外用", "皮下注"  # 開始セルだけ点滴と同じ色にする用法。			
	for i, yakurow in enumerate(yakurows, start=r):  # 各行について
		yaku, yoho, dummy, gentei = yakurow
		if yaku and i!=VARS.blackrow:
			if end is not None:
				endc = nendc if yoho else tendc
			sheet[i, startc:endc+1].setPropertyValue("CellBackColor", -1)  # 描く予定のセル範囲の背景色をクリア。
			if yoho in ("検査値",):  # 色を付けない用法。
				continue	
			if gentei:  # 限定条件がある時。
				gentei = gentei.split("(", 1)[0]  # (から前のみを取得。
				genteidigit = gentei.translate(table)  # 曜日を数字に変換する。
				cols = []  # 背景色する列インデックスのリスト。
				if genteidigit.isdigit():  # 全て数字に変換できたときは、日月火水木金土しかない時。			
					cols = [j for j in range(startc, endc+1) if str((weekdayval-1+j-startc)%7+1) in genteidigit]  # 日曜日=1から始まる。
				elif gentei.endswith("日に1回"):
					k = int(gentei.replace("日に1回", ""))
					cols = range(startc, endc+1, k)  # range()はジェネレーターではなくシークエンス型を返す。
				if yoho in yohos:  # 開始セルだけ点滴と同じ色にする用法。	
					if cols:  # 背景色をつける列インデックスがある時。
						tentekirangeaddress.append(sheet[i, cols[0]].getRangeAddress())	
						if len(cols)>1:
							naifukurangeaddress.extend(sheet[i, j].getRangeAddress() for j in cols[1:])	
				elif yoho:	
					naifukurangeaddress.extend(sheet[i, j].getRangeAddress() for j in cols)	
				else:  # 用法列がない時は点滴と考える。
					tentekirangeaddress.extend(sheet[i, j].getRangeAddress() for j in cols)
			else:
				if yoho in yohos:  # 開始セルだけ点滴と同じ色にする用法。
					tentekirangeaddress.append(sheet[i, startc].getRangeAddress())	
					if startc<endc:
						naifukurangeaddress.append(sheet[i, startc+1:endc+1].getRangeAddress())	
				elif yoho:
					naifukurangeaddress.append(sheet[i, startc:endc+1].getRangeAddress())
				else:  # 用法列がない時は点滴と考える。
					tentekirangeaddress.append(sheet[i, startc:endc+1].getRangeAddress())	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	if naifukurangeaddress:
		sheetcellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。			
		sheetcellranges.addRangeAddresses(naifukurangeaddress, False)
		sheetcellranges.setPropertyValue("CellBackColor", commons.COLORS["lime"])
	if tentekirangeaddress:
		sheetcellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。			
		sheetcellranges.addRangeAddresses(tentekirangeaddress, False)
		sheetcellranges.setPropertyValue("CellBackColor", commons.COLORS["magenta3"])	
def getNewEndC(startc, weekdayval, n, t, w):		
	if weekdayval==1:
		weekdayval += 7  # 日曜日のときは翌週にまたがないように8にする。
	newendc = startc + 7 - weekdayval
	nendc = newendc + n + w  # 内服用。
	tendc = newendc + t + w  # 点滴用。		
	return nendc, tendc
def setDates(xscriptcontext, doc, sheet, cell, datevalue, *, daycount=100):  # sheet:経過シート、cell: 日付開始セル、dateserial: 日付開始日のシリアル値。daycount: 経過シートに入力する日数。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。			
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
	createFormatKey = commons.formatkeyCreator(doc)	
	colors = commons.COLORS
	celladdress = cell.getCellAddress()  # 経過シートの日付の開始セルのセルアドレスを取得。
	r, c = celladdress.Row, celladdress.Column  # cは開始列インデックスになる。
	sheet[:r+1, c:].clearContents(511)  # 開始列より右の日付行の内容を削除。
	endcolumn = c + daycount
	if not endcolumn<1024:
		endcolumn = 1023  # 列インデックスの上限1023。	
		daycount = endcolumn - c
	datevalues = [i for i in range(datevalue, datevalue+daycount)]  # 日付シリアル値を取得。
	sheet[r, c:endcolumn].setDataArray((datevalues,))  # 日時シリアル値を経過シートに入力。
	sheet[r, c:endcolumn].setPropertyValues(("NumberFormat", "HoriJustify"), (createFormatKey('D'), CENTER))  # 経過シートの日付の書式を設定。
	holidaycolumns = getHolidaycolumns(functionaccess, datevalues, c)
	startweekday = int(functionaccess.callFunction("WEEKDAY", (datevalues[0], 3)))  # 開始日の曜日を取得。月=0。
	offdaycolumns = getOffdaycolumns(doc, datevalues, startweekday, c, endcolumn)  # 予定シートの休日設定を取得して合致する列インデックスを取得する。
	offdaycolumns.difference_update(holidaycolumns)  # 休日インデックスから祝日インデックスを除く。
	n = 6  # 日曜日の曜日番号。
	sunindexes = set(range(c+(n-startweekday)%7, endcolumn, 7))  # 日曜日の列インデックスの集合。祝日と重ならないようにあとで使用する。	
	holidaycolumns.difference_update(sunindexes)  # 祝日インデックスから日曜日インデックスを除く。
	n = 5  # 土曜日の曜日番号。
	satindexes = set(range(c+(n-startweekday)%7, endcolumn, 7))  # 土曜日の列インデックスの集合。
	setRangesProperty = createSetRangesProperty(doc, sheet, r)
	setRangesProperty(holidaycolumns, ("CellBackColor", colors["red3"]))
	setRangesProperty(offdaycolumns, ("CellBackColor", colors["silver"]))
	setRangesProperty(sunindexes, ("CharColor", colors["red3"]))
	setRangesProperty(satindexes, ("CharColor", colors["skyblue"]))
	month = int(functionaccess.callFunction("MONTH", (datevalues[0],)))  # 開始月を取得。
	if c==VARS.splittedcolumn:  # 日付行の先頭列の時。
		sheet[r-1, c].setString("{}月".format(month))
	startmonthindex = 0
	while True:
		startmonthindex = int(functionaccess.callFunction("EOMONTH", (datevalues[startmonthindex], 0))) - datevalue + 1  # 次月の1日のdatevaluesでのインデックスを取得。
		month += 1
		if month>12:
			month = 1
		if startmonthindex<daycount:	
			sheet[r-1, c+startmonthindex].setString("{}月".format(month))
		else:
			break
	highlightToday(sheet, xscriptcontext)
def getHolidaycolumns(functionaccess, datevalues, c): # 祝日になる列インデックスを返す。datevalues: 日付シリアル値のタプル。c: 開始列インデックス。
	holidaycolumns = set()  # 祝日インデックスの集合。
	holidays = commons.HOLIDAYS	
	startyear, startmonth = [int(functionaccess.callFunction(i, (datevalues[0],))) for i in ("YEAR", "MONTH")]  # 開始年月日を取得。
	endyear, endmonth = [int(functionaccess.callFunction(i, (datevalues[-1],))) for i in ("YEAR", "MONTH")]  # 終了年月日を取得。
	if startyear in holidays:  # 開始年の祝日がある時。
		for m, days in enumerate(holidays[startyear][startmonth-1:], start=startmonth):  # 開始月以降の祝日のタプルを取得。
			for d in days:
				datevalue = int(functionaccess.callFunction("DATE", (startyear, m, d)))
				if datevalue in datevalues:
					holidaycolumns.add(c+datevalues.index(datevalue))
				elif m>startmonth:  # 開始月より後はもう日付列は終了しているので関数を出る。
					return holidaycolumns
	newyear = startyear + 1
	while newyear<endyear:  # 最終年ではない間。
		if newyear in holidays:
			for m, days in enumerate(holidays[newyear], start=1):
				for d in days:	
					datevalue = int(functionaccess.callFunction("DATE", (newyear, m, d)))
					holidaycolumns.add(c+datevalues.index(datevalue))
		newyear += 1	
	if newyear==endyear:  # 最終年の時。
		if endyear in holidays:
			for m, days in enumerate(holidays[endyear][:endmonth], start=1):
				for d in days:
					datevalue = int(functionaccess.callFunction("DATE", (endyear, m, d)))
					if datevalue in datevalues:
						holidaycolumns.add(c+datevalues.index(datevalue))
	return holidaycolumns
def getOffdaycolumns(doc, datevalues, startweekday, c, endcolumn):  # 予定シートの休日設定を取得して合致する列インデックスを取得する。
	offdays, offweekdays = getOffdays(doc)  # 予定シートの休日設定を取得。offdays; 休日シリアル値、offweeks: 休日にする曜日番号。
	offdaycolumns = set()  # 休日インデックスの集合。
	offdaycolumns.update(c+datevalues.index(i) for i in offdays if i in datevalues)  # 休日のシリアル値のインデックスを取得。
	offdaycolumns.update(j for i in offweekdays for j in range(c+(i-startweekday)%7, endcolumn, 7))  # 曜日のインデックスを取得。
	return offdaycolumns
def getOffdays(doc):  # 予定シートの休日を取得。
	sheets = doc.getSheets()  # シートコレクションを取得。
	sheet = sheets["予定"]
	yoteivars = yotei.VARS
	yoteivars.setSheet(sheet)	
	weekdays = yoteivars.weekdays  # 曜日のタプルを取得。
	searchdescriptor = sheet.createSearchDescriptor()
	searchdescriptor.setSearchString("休日設定")  # 戻り値はない。
	c = yoteivars.templatestartcolumn - 1  # 休日設定のある列インデックスを取得。
	searchedcell = sheet[yoteivars.emptyrow:, c].findFirst(searchdescriptor)  # 休日設定の開始セルを取得。見つからなかった時はNoneが返る。
	if searchedcell:  # 休日設定の開始セルがある時。
		startrow = searchedcell.getCellAddress().Row + 2  # 休日設定の開始行を取得。
		cellranges = sheet[startrow:, c].queryContentCells(CellFlags.STRING+CellFlags.DATETIME)  # 休日設定列の文字列か日付が入っているセルに限定して抽出。
		emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 最終行インデックス+1を取得。		
		offweekdays = []  # 休日の曜日のリスト。
		offdays = []  # 休日のシリアル値のリスト。
		if startrow<emptyrow:  # 休日設定行がある時。
			for datarow in sheet[startrow:emptyrow, c].getDataArray():	# 休日設定の各行について。
				d = datarow[0]
				if isinstance(d, float):  # floatの時は日付シリアル値と考える。
					offdays.append(int(d))
				else:  # 文字列の時。
					offweekdays.extend(weekdays.index(i) for i in d.replace("曜日", ""))  # 曜日は曜日番号で取得する。金土などの書き方も処理する。
	return offdays, offweekdays
def createSetRangesProperty(doc, sheet, r): 
	def setRangesProperty(columnindexes, prop):  # r行のcolumnindexesの列のプロパティを変更。prop: プロパティ名とその値のリスト。
		cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
		if columnindexes:  
			cellranges.addRangeAddresses([sheet[r, i].getRangeAddress() for i in columnindexes], False)  # セル範囲コレクションを取得。columnindexesが空要素だとエラーになる。
			if len(cellranges):  # sheetcellrangesに要素がないときはsetPropertyValue()でエラーになるので要素の有無を確認する。
				cellranges.setPropertyValue(*prop)  # セル範囲コレクションのプロパティを変更。
	return setRangesProperty
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
