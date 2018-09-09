#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import glob, os, unohelper
from indoc import commons, ichiran
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
class Ent():  # シート固有の定数設定。
	def __init__(self):
		self.menurow  = 0  # メニュー行インデックス。
		self.splittedrow = 1  # 分割行インデックス。
		self.idcolumn = 0  # ID列インデックス。	
		self.kanjicolumn = 1  # 漢字列インデックス。
		self.kanacolumn = 2  # カナ列インデックス。	
		self.datecolumn = 3  # 入院日列インデックス。
		self.keikacolumn = 5  # 経過列インデックス。
	def setSheet(self, sheet):
		self.sheet = sheet
		cellranges = sheet[:, self.idcolumn].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # ID列の文字列が入っているセルに限定して抽出。数値の時もありうる。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # ID列の最終行インデックス+1を取得。
VARS = Ent()
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	sheet["A1:G1"].setDataArray((("ID", "漢字名", "ｶﾅ名", "入院日", "ﾘｽﾄ消去日", "経過", "ﾘｽﾄに戻る"),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	if enhancedmouseevent.ClickCount==2 and enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ダブルクリックの時。まずselectionChanged()が発火している。
		selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			celladdress = selection.getCellAddress()
			r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。
			if r<VARS.splittedrow:
				return mousePressedWSectionM(enhancedmouseevent, xscriptcontext)			
			elif r<VARS.emptyrow:
				return mousePressedWSectionB(enhancedmouseevent, xscriptcontext)
			else:  # ID列が空欄の時。キーボードからの入力は想定しない。
				sortRows(c)  # 昇順にソート。
				return False  # セル編集モードにしない。	
	return True  # セル編集モードにする。	
def mousePressedWSectionM(enhancedmouseevent, xscriptcontext):
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	c = selection.getCellAddress().Column
	if c>VARS.keikacolumn:  # 経過列より右の時。
		txt = selection.getString()
		if txt=="ﾘｽﾄに戻る":
			controller = doc.getCurrentController()  # コントローラの取得。
			sheets = doc.getSheets()
			controller.setActiveSheet(sheets["一覧"])  # 一覧シートをアクティブにする。
		elif txt=="改行削除":
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。									
			systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。			
			clipboardtxt = commons.getClipboardtxt(systemclipboard)
			if clipboardtxt:
				outputs = []
				buffer = []
				for txt in clipboardtxt.split("\n"):
					txt = txt.strip()
					if txt.startswith("****"):
						continue
					elif txt.startswith("#"):
						if buffer and outputs:
							outputs[-1] = "".join([outputs[-1], *buffer])
						outputs.append(txt)
						buffer = []
					else:
						buffer.append(txt)	
				if buffer and outputs:
					outputs[-1] = "".join([outputs[-1], *buffer])
				systemclipboard.setContents(commons.TextTransferable("\r\n".join(outputs)), None)  # クリップボードにコピーする。\rはWindowsのメモ帳でも改行するため。
	elif c<VARS.keikacolumn:  # 経過列より左のときはその項目で逆順にする。
		sortRows(c, reverse=True)  # 逆順にソート。
	return False  # セル編集モードにしない。		
def sortRows(c, *, reverse=None):
	if VARS.splittedrow<VARS.emptyrow:
		datarange = VARS.sheet[VARS.splittedrow:VARS.emptyrow, :VARS.keikacolumn+1]
		datarows = list(datarange.getDataArray())  # 行をリストで取得。要素はタプル。
		datarows.sort(key=lambda x:x[c], reverse=reverse)  # 各行を列インデックスcでソート。
		datarange.setDataArray(datarows)  # シートに代入する。	
def mousePressedWSectionB(enhancedmouseevent, xscriptcontext):
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
	systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。	
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。		
	celladdress = selection.getCellAddress()
	r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。	
	if c==VARS.idcolumn:  # ID列の時。
		systemclipboard.setContents(commons.TextTransferable(selection.getString()), None)  # クリップボードにIDをコピーする。
	elif c==VARS.kanacolumn:  # カナ名列の時。
		idtxt, dummy, kanatxt = VARS.sheet[r, :VARS.kanacolumn+1].getDataArray()[0]
		kanatxt = commons.convertKanaFULLWIDTH(transliteration, kanatxt)  # カナ名を半角からスペースを削除して全角にする。
		systemclipboard.setContents(commons.TextTransferable("".join((kanatxt, idtxt))), None)  # クリップボードにカナ名+IDをコピーする。	
	elif c==VARS.keikacolumn+1:  # リスト消去日列の時。
		datarows = VARS.sheet[r, VARS.idcolumn:VARS.datecolumn].getDataArray()  # ID、漢字名、カナ名を取得。
		sheets = doc.getSheets()
		ichiransheet = sheets["一覧"]
		ichiranvars = ichiran.VARS
		ichiranvars.setSheet(ichiransheet)
		datarange = ichiransheet[ichiranvars.emptyrow, ichiranvars.idcolumn:ichiranvars.datecolumn]
		datarange.setDataArray(datarows)
		controller = doc.getCurrentController()  # コントローラの取得。
		controller.setActiveSheet(ichiransheet)
	return False  # セル編集モードにしない。			
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	selection = eventobject.Source.getSelection()
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
		VARS.setSheet(selection.getSpreadsheet())	
		drowBorders(selection)  # 枠線の作成。
def drowBorders(selection):  # ターゲットを交点とする行列全体の外枠線を描く。
	noneline, dummy, topbottomtableborder, dummy = commons.createBorders()
	VARS.sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
	r = selection[0, 0].getCellAddress().Row
	if VARS.splittedrow<=r<VARS.emptyrow:
		rangeaddress = selection.getRangeAddress()
		VARS.sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。					
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。	
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
	del contextmenu[:]  # contextmenu.clear()は不可。
	selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
	if selection[0, 0].getCellAddress().Row<VARS.splittedrow:  # 分割行より上の時はコンテクストメニューを表示しない。
		return EXECUTE_MODIFIED
	rangeaddress = selection.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。
	startrow = rangeaddress.StartRow
	if contextmenuname=="cell":  # セルのとき。セル範囲も含む。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
			startcolumn = rangeaddress.StartColumn
			idtxt, dummy, kanatxt = sheet[startrow, VARS.idcolumn:VARS.datecolumn].getDataArray()[0]
			filename = ""
			if startcolumn in (VARS.datecolumn,):  # 入院日列の時。
				filename = "{}{}_*入院.ods"  # カルテシートファイル名。
			elif startcolumn in (VARS.keikacolumn,):  # 経過列の時。
				filename = "{}{}経_*開始.ods"  # 経過シートファイル名。
			if filename:  # ファイル名が取得出来ている時。		
				doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。				
				transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
				for i, systempath in enumerate(glob.iglob(commons.createKeikaPathname(doc, transliteration, idtxt, kanatxt, filename), recursive=True)):  # アーカイブフォルダ内の経過ファイルリストを取得する。
					addMenuentry("ActionTrigger", {"Text": os.path.basename(systempath), "CommandURL": baseurl.format("entry{}".format(21+i))}) 
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。		
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"Text": "クリア", "CommandURL": baseurl.format("entry1")}) 
	elif contextmenuname=="rowheader":  # 行ヘッダーのとき。				
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		commons.rowMenuEntries(addMenuentry)
	elif contextmenuname=="colheader":  # 列ヘッダーの時。
		pass
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。	
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	selection = controller.getSelection()  # 選択範囲を取得。
	if entrynum==1:  # クリア。
		rangeaddress = selection.getRangeAddress()  # 選択範囲のアドレスを取得。
		sheet = VARS.sheet
		splittedrow = VARS.splittedrow
		keikacolumn = VARS.keikacolumn
		cellflags = CellFlags.VALUE + CellFlags.DATETIME + CellFlags.STRING + CellFlags.ANNOTATION + CellFlags.FORMULA
		for i in range(rangeaddress.StartRow, rangeaddress.EndRow+1):  # 選択範囲の行インデックスをイテレート。
			for j in range(rangeaddress.StartColumn, rangeaddress.EndColumn+1):  # 選択範囲の列インデックスをイテレート。
				if i<splittedrow:  # 固定行より上の時。
					continue
				elif j<=keikacolumn:
					sheet[i, j].clearContents(cellflags)
				else:  # それ以外の時。
					sheet[i, j].clearContents(511)  # 範囲をすべてクリアする。		
	elif entrynum>20:  # startentrynum以上の数値の時はアーカイブファイルを開く。
		startentrynum = 21
		c = entrynum - startentrynum  # コンテクストメニューからファイルリストのインデックスを取得。
		sheet = controller.getActiveSheet()  # アクティブシートを取得。
		rangeaddress = selection.getRangeAddress()  # 選択範囲のアドレスを取得。
		idtxt, dummy, kanatxt = sheet[rangeaddress.StartRow, VARS.idcolumn:VARS.datecolumn].getDataArray()[0]
		startcolumn = rangeaddress.StartColumn
		if startcolumn in (VARS.datecolumn,):  # 入院日列の時。
			filename = "{}{}_*入院.ods"  # カルテシートファイル名。
		elif startcolumn in (VARS.keikacolumn,):  # 経過列の時。
			filename = "{}{}経_*開始.ods"  # 経過シートファイル名。	
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
		transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。					
		for i, systempath in enumerate(glob.iglob(commons.createKeikaPathname(doc, transliteration, idtxt, kanatxt, filename), recursive=True)):  # アーカイブフォルダ内の経過ファイルリストを取得する。
			if i==c:  # インデックスが一致する時。
				desktop = xscriptcontext.getDesktop()
				desktop.loadComponentFromURL(unohelper.systemPathToFileUrl(systempath), "_blank", 0, ())  # ドキュメントを開く。
				break
