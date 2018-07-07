#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 一覧シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import os, unohelper, glob
from itertools import chain
from indoc import commons, keika, ent, datedialog
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.awt import MouseButton, MessageBoxButtons, MessageBoxResults # 定数
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.awt.MessageBoxType import QUERYBOX, ERRORBOX  # enum
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH, HIRAGANA_KATAKANA  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.table.CellHoriJustify import LEFT  # enum
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.sheet.CellDeleteMode import ROWS as delete_rows  # enum
from com.sun.star.beans import PropertyValue  # Struct
class Ichiran():  # シート固有の定数設定。
	def __init__(self, sheet):
		self.menurow  = 0  # メニュー行インデックス。
		self.splittedrow = 2  # 分割行インデックス。
		self.sumicolumn = 0  # 済列インデックス。
		self.yocolumn = 1  # 予列インデックス。
		self.idcolumn = 2  # ID列インデックス。	
		self.kanacolumn = 4  # カナ列インデックス。	
		self.datecolumn = 5  # 入院日列インデックス。
		self.checkstartcolumn = 8  # チェック列開始列インデックス。
		self.memostartcolumn = 22  # メモ列開始列インデックス。
		cellranges = sheet[self.splittedrow:, self.idcolumn].queryContentCells(CellFlags.STRING)  # ID列の文字列が入っているセルに限定して抽出。
		backcolors = commons.COLORS["blue3"], commons.COLORS["skyblue"], commons.COLORS["red3"]  # ジェネレーターに使うので順番が重要。
		gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in backcolors)
		self.bluerow = next(gene)  # 青3行インデックス。
		self.skybluerow = next(gene)  # スカイブルー行インデックス。
		self.redrow = next(gene)  # 赤3行インデックス。	
		cellranges = sheet[:, self.idcolumn].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # ID列の文字列が入っているセルに限定して抽出。数値の時もありうる。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # ID列の最終行インデックス+1を取得。
def getConsts(sheet, selection=None):  # 区画名を取得。
	"""
	M  |
	---
	C
	===========  # 行の固定の境界
	B  |D|E
	   | |
	-----------
	A  # ID列が空欄の行。
	
	M: メニュー行。
	C: メニュー行以外のスクロールしない部分。
	B: スクロールする部分のうちチェック列より左の部分。
	D: スクロールする部分のうちチェック列。
	E: スクロールする部分のうちチェック列より右の部分。
	A: ID列の最初の空行から下の部分。
	"""
	consts = Ichiran(sheet)  # クラスをインスタンス化。	
	if selection is not None:
		splittedrow = consts.splittedrow
		checkstartcolumn = consts.checkstartcolumn
		memostartcolumn = consts.memostartcolumn
		rangeaddress = selection[0, 0].getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
		emptyrow = consts.emptyrow
		if len(sheet[consts.menurow, :checkstartcolumn].queryIntersection(rangeaddress)):  # メニューセルの時。
			sectionname = "M"
		elif len(sheet[splittedrow:emptyrow, :checkstartcolumn].queryIntersection(rangeaddress)):  # Dの左。
			sectionname = "B"	
		elif len(sheet[splittedrow:emptyrow, checkstartcolumn:memostartcolumn].queryIntersection(rangeaddress)):  # チェック列の時。
			sectionname = "D"		
		elif len(sheet[splittedrow:emptyrow, memostartcolumn:].queryIntersection(rangeaddress)):  # Dの右。
			sectionname = "E"		
		elif len(sheet[emptyrow:, :].queryIntersection(rangeaddress)):  # まだデータのない行の時。
			sectionname = "A"	
		else:
			sectionname = "C"  
		consts.sectionname = sectionname   # 区画名
	return consts
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	sheet["C1:G1"].setDataArray((("済をﾘｾｯﾄ", "検予を反映", "予をﾘｾｯﾄ", "入力支援", "退院ﾘｽﾄ"),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	sheet = selection.getSpreadsheet()
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				drowBorders(sheet, selection, commons.createBorders())  # 枠線の作成。
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				consts = getConsts(sheet, selection)
				sectionname	= consts.sectionname	
				if sectionname=="M":
					doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
					ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
					smgr = ctx.getServiceManager()  # サービスマネージャーの取得。					
					functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
					return mousePressedWSectionM(doc, sheet, functionaccess, consts, selection)			
				elif not selection.getPropertyValue("CellBackColor") in (-1, commons.COLORS["cyan10"]):  # 背景色がないか薄緑色でない時以外何もしない。
					return False  # セル編集モードにしない。
				elif sectionname=="B":
					return mousePressedWSectionB(xscriptcontext, enhancedmouseevent, consts)
				elif sectionname=="D":
					return mousePressedWSectionD(sheet, consts, selection)
				elif sectionname=="A":
# 					celladdress = selection.getCellAddress()
# 					r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。	
# 					if c==consts.kanacolumn:  # カナ列の時。

						
					pass  # 漢字名からｶﾅを取得する。つまりふりがなを降る。

	return True  # セル編集モードにする。	
def mousePressedWSectionM(doc, sheet, functionaccess, consts, selection):
	controller = doc.getCurrentController()  # コントローラの取得。
	sheets = doc.getSheets()  # シートコレクションを取得。
	txt = selection.getString()  # クリックしたセルの文字列を取得。	
	if txt=="検予を反映":  # 経過シートから本日の検予を取得。
		splittedrow, checkstartcolumn, memostartcolumn = consts.splittedrow, consts.checkstartcolumn, consts.memostartcolumn
		cellranges = sheet[splittedrow:, consts.idcolumn].queryContentCells(CellFlags.STRING)  # ID列に文字列が入っているセルを取得。
		headerrow = sheet[consts.menurow, checkstartcolumn:memostartcolumn].getDataArray()[0]  # チェック列のヘッダーのタプルを取得。
		eketsucol, dokueicol, ketuekicol, gazocol, shochicol, echocol, ecgcol\
			= [headerrow.index(i) for i in ("ｴ結", "読影", "血液", "画像", "処置", "ｴｺ", "ECG")]  # headerrowタプルでのインデックスを取得。
		todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。	
		keikaconsts = None
		if len(cellranges)>0:  # ID列のセル範囲が取得出来ている時。
			iddatarows = cellranges[0].getDataArray()  # ID列のデータ行のタプルを取得。空行がないとする。
			checkrange = sheet[splittedrow:splittedrow+len(iddatarows), checkstartcolumn:memostartcolumn]  # チェック列範囲を取得。
			datarows = list(map(list, checkrange.getDataArray()))  # 各行をリストにして取得。
			for r, idtxt in enumerate(chain.from_iterable(iddatarows)):  # 各ID列について。rは相対インデックス。
				if idtxt.isdigit():  # IDがすべて数字の時。
					sheetname = "{}経".format(idtxt)  # 経過シート名を作成。
					if not sheetname in sheets:  # 経過シートがない時は次のループに行く。
						continue
					keikasheet = sheets[sheetname]  # 経過シートを取得。
					if keikaconsts is None:
						keikaconsts = keika.getConsts(keikasheet)  # 経過シートの定数を取得。
						daterow = keikaconsts.daterow  # 経過シートの日付行インデックスを取得。
						splittedcolumn = keikaconsts.splittedcolumn  # 日付列の最初の列インデックスを取得。
						c = splittedcolumn + todayvalue  # 分割列と今日の日付のシリアル値の和。
					startdatevalue = int(keikasheet[daterow, splittedcolumn].getValue())  # 日付行の最初のセルから日付のシリアル値の取得。
					keikadatarows = keikasheet[daterow+1:daterow+3, c-startdatevalue].getDataArray()  # 今日の日付列のセル範囲の値を取得。
					datarows[r][ketuekicol] = keikadatarows[0][0]  # 血液。
					s = keikadatarows[1][0]  # 2行目を取得。
					for i in commons.GAZOs:  # 読影のない画像。
						if i in s:
							if not i in datarows[r][gazocol]:  # すでにない時のみ。
								datarows[r][gazocol] += i
					for i in commons.GAZOd:  # 読影のある画像。
						if i in s:
							if not i in datarows[r][gazocol]:  # すでにない時のみ。
								datarows[r][gazocol] += i											
							datarows[r][dokueicol] = "○"
					for i in commons.ECHOs:  # エコー。
						if i in s:
							if not i in datarows[r][echocol]:  # すでにない時のみ。
								datarows[r][echocol] += i		
							datarows[r][eketsucol] = "○"	
					for i in commons.SHOCHIs:  # 処置。
						if i in s:
							if not i in datarows[r][shochicol]:  # すでにない時のみ。
								datarows[r][shochicol] += i			
					if "ECG" in s:  # ECG。
						if not "E" in datarows[r][ecgcol]:  # すでにない時のみ。
							datarows[r][ecgcol] = "E"							
			checkrange.setDataArray(datarows)  # シートに書き戻す。
	elif txt=="済をﾘｾｯﾄ":
		msg = "済列をリセットしますか？"
		componentwindow = controller.ComponentWindow
		msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_OK, "myRs", msg)
		if msgbox.execute()==MessageBoxResults.OK:
			emptyrow = consts.emptyrow
			sheet[splittedrow:emptyrow, :].setPropertyValue("CharColor", commons.COLORS["black"])  # 文字色を黒色にする。
			sheet[splittedrow:emptyrow, consts.sumicolumn].setDataArray([("未",)]*(emptyrow-splittedrow))  # 済列をリセット。
			searchdescriptor = sheet.createSearchDescriptor()
			searchdescriptor.setSearchString("済")
			cellranges = sheet[splittedrow:emptyrow, checkstartcolumn:memostartcolumn].findAll(searchdescriptor)  # チェック列の「済」が入っているセル範囲コレクションを取得。
			cellranges.setPropertyValue("CharColor", commons.COLORS["silver"])
	elif txt=="予をﾘｾｯﾄ":
		sheet[splittedrow:emptyrow, consts.sumicolumn+1].clearContents(CellFlags.STRING)  # 予列をリセット。
	elif txt=="入力支援":
		
		pass  # 入力支援odsを開く。
	
	elif txt=="退院ﾘｽﾄ":
		controller.setActiveSheet(sheets["退院"])
	return False  # セル編集モードにしない。	
def mousePressedWSectionB(xscriptcontext, enhancedmouseevent, consts):
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	sheet = selection.getSpreadsheet()
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
	createFormatKey = commons.formatkeyCreator(doc)
	controller = doc.getCurrentController()  # コントローラの取得。
	sheets = doc.getSheets()  # シートコレクションを取得。	
	celladdress = selection.getCellAddress()
	r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。		
	sumitxt, yotxt, idtxt, kanjitxt, kanatxt, datevalue, keikatxt = sheet[r, :consts.checkstartcolumn-1].getDataArray()[0]  # 日付はfloatで返ってくる。	
	datevalue = datevalue and int(datevalue)  # 計算しにくいのでdatevalueがあるときはfloatを整数にしておく。	
	if keikatxt and c==0:  # 経過列があり、かつ、済列の時。
		items = [("待", "skyblue"), ("済", "silver"), ("未", "black")]
		items.append(items[0])  # 最初の要素を最後の要素に追加する。
		dic = {items[i][0]: items[i+1] for i in range(len(items)-1)}  # 順繰り辞書の作成。								
		selection.setString(dic[sumitxt][0])
		sheet[r, :].setPropertyValue("CharColor", commons.COLORS[dic[sumitxt][1]])						
		refreshCounts(sheet, consts)  # カウントを更新する。
	elif keikatxt and c==consts.yocolumn:  # 経過列があり、かつ、予列の時。
		if yotxt:
			selection.clearContents(CellFlags.STRING)  # 予をクリア。
		else:  # セルの文字列が空の時。
			selection.setString("予")
	elif c==consts.idcolumn:  # ID列の時。
		if keikatxt:  # 経過列がある時。
			systemclipboard.setContents(commons.TextTransferable(idtxt), None)  # クリップボードにIDをコピーする。
		else:
			return True  # セル編集モードにする。		
	elif c==consts.idcolumn+1:  # 漢字列の時。カルテシートをアクティブにする、なければ作成する。カルトシート名はIDと一致。	
		if keikatxt and idtxt in sheets:  # 経過列があり、かつ、ID名のシートが存在する時。
			controller.setActiveSheet(sheets[idtxt])  # カルテシートをアクティブにする。
		else:  # 在院日数列が空欄の時、または、カルテシートがない時。
			if all((idtxt, kanjitxt, kanatxt, datevalue)):  # ID、漢字名、カナ名、入院日、すべてが揃っている時。	
				fillColumns(transliteration, createFormatKey, sheet, r, consts, idtxt, kanjitxt, kanatxt, datevalue)
				karutesheet = commons.getKaruteSheet(createFormatKey, sheets, idtxt, kanjitxt, kanatxt, datevalue)  # カルテシートを取得。
				controller.setActiveSheet(karutesheet)  # カルテシートをアクティブにする。	
			else:
				return True  # セル編集モードにする。		
	elif c==consts.kanacolumn:  # カナ名列の時。
		if keikatxt:  # 経過列がすでにある時。
			kanatxt = commons.convertKanaFULLWIDTH(transliteration, kanatxt)  # カナ名を半角からスペースを削除して全角にする。
			systemclipboard.setContents(commons.TextTransferable("".join((kanatxt, idtxt))), None)  # クリップボードにカナ名+IDをコピーする。	
		else:
			return True  # セル編集モードにする。		
	elif c==consts.datecolumn:  # 入院日列の時。
		datedialog.createDialog(xscriptcontext, enhancedmouseevent, "入院日", "YYYY/MM/DD")		
	elif c==consts.datecolumn+1:  # 経過列のボタンはカルテシートの作成時に作成されるのでカルテシート作成後のみ有効。			
		newsheetname = "".join([idtxt, "経"])  # 経過シート名を取得。
		if keikatxt and newsheetname in sheets:  # 経過列がすでにあり、かつ、経過シートがある時。。		
			controller.setActiveSheet(sheets[newsheetname])  # 経過シートをアクティブにする。
		else:  # 経過シートがなければ作成する。
			if all((idtxt, kanjitxt, kanatxt, datevalue)):  # ID、漢字名、カナ名、入院日、すべてが揃っている時。									
				fillColumns(transliteration, createFormatKey, sheet, r, consts, idtxt, kanjitxt, kanatxt, datevalue)
				keikasheet =  commons.getKeikaSheet(doc, createFormatKey, sheets, idtxt, kanjitxt, kanatxt, datevalue)  # 経過シートを取得。
				controller.setActiveSheet(keikasheet)  # 経過シートをアクティブにする。						
	return False  # セル編集モードにしない。		
def mousePressedWSectionD(sheet, consts, selection):
	txt = selection.getString()  # クリックしたセルの文字列を取得。	
	celladdress = selection.getCellAddress()
	c = celladdress.Column  # selectionの行と列のインデックスを取得。		
	dic = {\
		"4F": ["", "待", "○", "包"],\
		"ｴ結": ["", "ｴ", "済"],\
		"読影": ["", "読", "済", "無"],\
		"退処": ["", "済", "△", "待"],\
		"血液": ["", "尿", "○", "済"],\
		"ECG": ["", "E", "済"],\
		"糖尿": ["", "糖"],\
		"熱発": ["", "熱"],\
		"計書": ["", "済", "未"],\
		"面談": ["", "面"],\
		"便指": ["", "済", "少", "無"]\
	}
	header = sheet[consts.menurow, c].getString()  # 行インデックス0のセルの文字列を取得。
	newtxt = txt
	if header in dic:	
		items = dic[header]	 # 順繰りのリストを取得。			
		if txt in items:  # セルの内容にある時。
			items.append(items[0])  # 最初の要素を最後の要素に追加する。
			dic = {items[i]: items[i+1] for i in range(len(items)-1)}  # 順繰り辞書の作成。
			newtxt = dic[txt]  # 次の要素を代入する。
	else:
		if txt.endswith("済"):
			newtxt = txt.rstrip("済")
		elif txt:
			newtxt = "{}済".format(txt)
	selection.setString(newtxt)
	color = commons.COLORS["silver"] if "済" in newtxt else -1
	selection.setPropertyValue("CharColor", color)			
	return False  # セル編集モードにしない。
def fillColumns(transliteration, createFormatKey, sheet, r, consts, idtxt, kanjitxt, kanatxt, datevalue):		
	locale = Locale(Language = "ja", Country = "JP")
	transliteration.loadModuleNew((HIRAGANA_KATAKANA,), locale)  # 変換モジュールをロード。	
	kanatxt = transliteration.transliterate(kanatxt, 0, len(kanatxt), [])[0]  # ひらがなをカタカナに変換		
	transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), locale)
	kanatxt = transliteration.transliterate(kanatxt, 0, len(kanatxt), [])[0]  # 半角に変換								
	cellstringaddress = sheet[r, consts.datecolumn].getPropertyValue("AbsoluteName").split(".")[-1].replace("$", "")  # 入院日セルの文字列アドレスを取得。
	cell = sheet[r, consts.checkstartcolumn-1]
	cell.setFormula("=TODAY()+1-{}".format(cellstringaddress))  #  在院日数列に式を代入。	
	cell.setPropertyValue("NumberFormat", createFormatKey('0" ";[RED]-0" "'))  # 在院日数列の書式を設定。 
	datarow = "未", "", idtxt, kanjitxt.strip().replace("　", " "), kanatxt, datevalue, "経過"  # 他の列を追加。								
	sheet[r, :consts.checkstartcolumn-1].setDataArray((datarow,))
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
def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。マクロで変更した時は発火しない模様。	
	changes = changesevent.Changes	
	for change in changes:
		if change.Accessor=="cell-change":  # セルの値が変化した時。
			selection = change.ReplacedElement  # 値を変更したセルを取得。	
			sheet = selection.getSpreadsheet()
			consts = getConsts(sheet)  # 一覧シート固有の定数を取得。
			celladdress = selection.getCellAddress()
			r, c = celladdress.Row, celladdress.Column
			if r>consts.splittedrow-1:  # 分割行以降の時。
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。					
				transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
				transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))	
				txt = selection.getString()  # セルの文字列を取得。			
				if c==consts.idcolumn:  # ID列の時。
					txt = transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換。
					if txt.isdigit():  # 数値の時のみ。空文字の時0で埋まってしまう。
						selection.setString("{:0>8}".format(txt))  # 数値を8桁にして文字列として代入し直す。
				elif c==consts.kanacolumn:  # カナ列の時。
					transliteration2 = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
					transliteration2.loadModuleNew((HIRAGANA_KATAKANA,), Locale(Language = "ja", Country = "JP"))  # 変換モジュールをロード。
					txt = transliteration2.transliterate(txt, 0, len(txt), [])[0]  # ひらがなをカタカナに変換。
					txt = transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換
					if all(map(lambda x: "ｱ"<=x<="ﾝ", txt.replace(" ", ""))):  # すべて半角カタカナであることを確認。スペースは除去して評価する。
						selection.setString(transliteration.transliterate(txt, 0, len(txt), [])[0])  # 半角に変換してセルに代入。
					else:
						msg = "ｶﾅ名列にはカタカナかひらながのみ入力してください。"
						doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
						controller = doc.getCurrentController()  # コントローラの取得。						
						componentwindow = controller.ComponentWindow
						msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, ERRORBOX, MessageBoxButtons.BUTTONS_OK, "myRs", msg)
						msgbox.execute()							
						controller.select(selection)  # 元のセルに戻る。セル編集モードにするとおかしくなる。
				elif c==consts.datecolumn:  # 日付列の時。
					doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
					selection.setPropertyValues(("NumberFormat", "HoriJustify"), (commons.formatkeyCreator(doc)('YYYY/MM/DD'), LEFT))  # カルテシートの入院日の書式設定。左寄せにする。
			break
def refreshCounts(sheet, consts):  # カウントを更新する。
	datarows = [["総数", 0, "済", 0], ["未", 0, "待", 0]]
	datarange = sheet[consts.splittedrow:consts.emptyrow, consts.sumicolumn]
	searchdescriptor = sheet.createSearchDescriptor()
	counts = []
	for txt in ("済", "待"):  # 未はタイトル行にも入っているので正しく計算できない。
		searchdescriptor.setSearchString(txt)  # 戻り値はない。
		cellranges = datarange.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
		c = len([i for i in cellranges.getCells()]) if cellranges else 0  # セルで数えないといけない。
		counts.append(c)
	counts.append(consts.emptyrow-consts.splittedrow-3-sum(counts))  # 済、待、未、の順に数が入る。
	datarows[0][1] = sum(counts)
	datarows[0][3] = counts[0]
	datarows[1][1] = counts[2]
	datarows[1][3] = counts[1]
	sheet[:2, consts.memostartcolumn:consts.memostartcolumn+4].setDataArray(datarows)	
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。	
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
	del contextmenu[:]  # contextmenu.clear()は不可。
	selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
	consts = getConsts(sheet, selection)  # セル固有の定数を取得。
	sectionname = consts.sectionname  # クリックしたセルの区画名を取得。		
	if sectionname in ("M", "C"):  # 固定行より上の時はコンテクストメニューを表示しない。
		return EXECUTE_MODIFIED
	rangeaddress = selection.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。
	startrow = rangeaddress.StartRow
	if startrow in (consts.bluerow, consts.skybluerow, consts.redrow):  # タイトル行の時。
		return EXECUTE_MODIFIED
	if contextmenuname=="cell":  # セルのとき。セル範囲も含む。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
			if rangeaddress.StartColumn in (consts.yocolumn,):  # 予列の時。
				addMenuentry("ActionTrigger", {"Text": "退院ﾘｽﾄへ", "CommandURL": baseurl.format("entry1")}) 	
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。	
			elif rangeaddress.StartColumn in (consts.datecolumn+1,):  # 経過列の時。
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
				transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
				doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。
				idtxt, dummy, kanatxt = sheet[startrow, consts.idcolumn:consts.datecolumn].getDataArray()[0]			
				addMenuentry("ActionTrigger", {"Text": "経過ｼｰﾄをArchiveへ", "CommandURL": baseurl.format("entry2")}) 
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。	
				for i, systempath in enumerate(glob.iglob(commons.createKeikaPathname(doc, transliteration, idtxt, kanatxt, "{}{}経_*開始.ods"), recursive=True)):  # アーカイブフォルダ内の経過ファイルリストを取得する。
					addMenuentry("ActionTrigger", {"Text": os.path.basename(systempath), "CommandURL": baseurl.format("entry{}".format(21+i))}) 
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。	
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Delete"})	
	elif contextmenuname=="rowheader" and len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 行ヘッダーのとき、かつ、選択範囲の列数がシートの列数が一致している時。	
		if sectionname in ("A",):
			commons.cutcopypasteMenuEntries(addMenuentry)
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			commons.rowMenuEntries(addMenuentry)
			return EXECUTE_MODIFIED
		if startrow<consts.bluerow:  # 未入院
			addMenuentry("ActionTrigger", {"Text": "新入院へ", "CommandURL": baseurl.format("entry3")})  
		elif startrow<consts.skybluerow:  # Stable
			addMenuentry("ActionTrigger", {"Text": "Unstableへ", "CommandURL": baseurl.format("entry4")})
			addMenuentry("ActionTrigger", {"Text": "新入院へ", "CommandURL": baseurl.format("entry5")})	
		elif startrow<consts.redrow:  # Unstable
			addMenuentry("ActionTrigger", {"Text": "Stableへ", "CommandURL": baseurl.format("entry6")})
			addMenuentry("ActionTrigger", {"Text": "新入院へ", "CommandURL": baseurl.format("entry7")}) 		
		else:  # 新入院
			addMenuentry("ActionTrigger", {"Text": "未入院へ", "CommandURL": baseurl.format("entry8")}) 		
			addMenuentry("ActionTrigger", {"Text": "Stableへ", "CommandURL": baseurl.format("entry9")})
			addMenuentry("ActionTrigger", {"Text": "Unstableへ", "CommandURL": baseurl.format("entry10")}) 				
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})		
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		commons.rowMenuEntries(addMenuentry)
	elif contextmenuname=="colheader":  # 列ヘッダーの時。
		pass
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。	
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。	
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	desktop = xscriptcontext.getDesktop()
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	selection = controller.getSelection()  # 選択範囲を取得。
	rangeaddress = selection.getRangeAddress()  # 選択範囲のアドレスを取得。
	r = rangeaddress.StartRow
	consts = getConsts(sheet)  # シート固有の値を取得。
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
	if entrynum<3:  # セルのコンテクストメニュー。
		functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
		sheets = doc.getSheets()
		datarow = sheet[r, consts.idcolumn:consts.datecolumn+1].getDataArray()[0]   # ダブルクリックした行をID列からｶﾅ名列までのタプルを取得。
		idtxt, dummy, kanatxt, datevalue = datarow
		kanatxt = commons.convertKanaFULLWIDTH(transliteration, kanatxt)  # カナ名を半角からスペースを削除して全角にする。
		datetxt = "-".join([str(int(functionaccess.callFunction(i, (datevalue,)))) for i in ("YEAR", "MONTH", "DAY")])  # シリアル値をシート関数で年-月-日の文字列にする。
		k = kanatxt[0]  # 最初のカナ文字を取得。カタカナであることは入力時にチェック済。
		kana = "ア", "カ", "サ", "タ", "ナ", "ハ", "マ", "ヤ", "ラ", "ワ"
		for i in range(1, len(kana)):
			if kanatxt[0]<kana[i]:
				k = kana[i-1]
				break
		else:
			k = kana[i]
		dirpath = os.path.dirname(unohelper.fileUrlToSystemPath(doc.getURL()))  # このドキュメントのあるディレクトリのフルパスを取得。
		kanadirpath = os.path.join(dirpath, k)  # 最初のカナ文字のフォルダへのパス。
		if not os.path.exists(kanadirpath):  # カタカナフォルダがないとき。
			os.mkdir(kanadirpath)  # カタカナフォルダを作成。 
		detachSheet = createDetachSheet(desktop, controller, doc, sheets, kanadirpath)
		if entrynum==1:  # 退院リストへ。
			flgs = []
			newsheetname = "{}{}_{}入院".format(kanatxt, idtxt, datetxt)  # 新しいシート名を取得。
			flgs.append(detachSheet(idtxt, newsheetname))
			newsheetname = "{}{}経_{}開始".format(kanatxt, idtxt, datetxt)  # 新しいシート名を取得。
			flgs.append(detachSheet("".join([idtxt, "経"]), newsheetname))
			if not all(flgs):
				msg = "{}{}をリストシートから退院シートに移動させますか？".format(kanatxt, idtxt)
				componentwindow = controller.ComponentWindow
				msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_OK, "myRs", msg)
				if msgbox.execute()!=MessageBoxResults.OK:  # OKでない時は何もしない。
					return			
			datarow = list(datarow)
			todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。
			datarow.extend((todayvalue, "経過"))
			entsheet = sheets["退院"]  # 退院シートを取得。
			entconsts = ent.getConsts(entsheet)  # 退院シートの定数を取得。			
			entsheet[entconsts.emptyrow, entconsts.idcolumn:entconsts.idcolumn+len(datarow)].setDataArray((datarow,))  # 退院シートにデータを代入。
			entsheet[entconsts.splittedrow:entconsts.emptyrow+1, entconsts.datecolumn:entconsts.cleardatecolumn+1].setPropertyValue("NumberFormat", commons.formatkeyCreator(doc)('YYYY/MM/DD'))  # 日付書式を設定。
			searchdescriptor = sheet.createSearchDescriptor()
			searchdescriptor.setSearchString(idtxt)  # 戻り値はない。
			cellranges = entsheet[entconsts.splittedrow:entconsts.emptyrow, entconsts.idcolumn].findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
			if cellranges:  # ID列に同じIDがすでにある時。
				[entsheet.removeRange(i, delete_rows) for i in cellranges.getRangeAddresses()]  # 同じIDの行を削除。
			sheet.removeRange(rangeaddress, delete_rows)  # 移動したソース行を削除。
		elif entrynum==2:  # 経過ｼｰﾄをArchiveへ。
			newsheetname = "{}{}経_{}開始".format(kanatxt, idtxt, datetxt)  # 新しいシート名を取得。
			detachSheet("".join([idtxt, "経"]), newsheetname)  # 切り出したシートのfileurlを取得。
	elif entrynum>20:  # startentrynum以上の数値の時はアーカイブファイルを開く。
		startentrynum = 21
		c = entrynum - startentrynum  # コンテクストメニューからファイルリストのインデックスを取得。
		idtxt, dummy, kanatxt = sheet[r, consts.idcolumn:consts.datecolumn].getDataArray()[0]
		for i, systempath in enumerate(glob.iglob(commons.createKeikaPathname(doc, transliteration, idtxt, kanatxt, "{}{}経_*開始.ods"), recursive=True)):  # アーカイブフォルダ内の経過ファイルリストを取得する。
			if i==c:  # インデックスが一致する時。
				desktop.loadComponentFromURL(unohelper.systemPathToFileUrl(systempath), "_blank", 0, ())  # ドキュメントを開く。
				break
	elif entrynum==3:  # 未入院から新入院に移動。
		commons.toNewEntry(sheet, rangeaddress, consts.bluerow, consts.emptyrow)
	elif entrynum==4:  # StableからUnstableへ移動。
		commons.toOtherEntry(sheet, rangeaddress, consts.skybluerow, consts.redrow)
	elif entrynum==5:  # Stableから新入院へ移動。 
		commons.toNewEntry(sheet, rangeaddress, consts.skybluerow, consts.emptyrow)
	elif entrynum==6:  # UnstableからStableへ移動。
		commons.toOtherEntry(sheet, rangeaddress, consts.redrow, consts.skybluerow)
	elif entrynum==7:  # Unstableから新入院へ移動。
		commons.toNewEntry(sheet, rangeaddress, consts.redrow, consts.emptyrow)
	elif entrynum==8:  # 新入院から未入院へ移動。
		commons.toOtherEntry(sheet, rangeaddress, consts.emptyrow, consts.bluerow)
	elif entrynum==9:  # 新入院からStableへ移動。
		commons.toOtherEntry(sheet, rangeaddress, consts.emptyrow, consts.skybluerow)
	elif entrynum==10:  # 新入院からUnstableへ移動。
		commons.toOtherEntry(sheet, rangeaddress, consts.emptyrow, consts.redbluerow)
def createDetachSheet(desktop, controller, doc, sheets, kanadirpath):
	propertyvalues = PropertyValue(Name="Hidden", Value=True),  # 新しいドキュメントのプロパティ。
	def detachSheet(sheetname, newsheetname):
		if sheetname in sheets:  # シートがある時。
			existingsheet = sheets[sheetname]  # カルテシートを取得。
			existingsheet.setName(newsheetname)  # カルテシート名を変更。
			newdoc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, propertyvalues)  # 新規ドキュメントの取得。
			newsheets = newdoc.getSheets()  # 新規ドキュメントのシートコレクションを取得。
			newsheets.importSheet(doc, newsheetname, 0)  # 新規ドキュメントにシートをコピー。
			del newsheets["Sheet1"]  # 新規ドキュメントのデフォルトシートを削除する。 
			del sheets[newsheetname]  # 切り出したカルテシートを削除する。 
			systempath = os.path.join(kanadirpath, "{}.ods".format(newsheetname))
			if os.path.exists(systempath):  # すでにファイルが存在する時。
				msg = "シート{}はすでにバックアップ済です。\n上書きしますか？"
				componentwindow = controller.ComponentWindow
				msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_OK, "myRs", msg)
				if msgbox.execute()!=MessageBoxResults.OK:			
					return True  # 上書きしない時は、切り出したものとして返す。
			fileurl = unohelper.systemPathToFileUrl(systempath)
			newdoc.storeToURL(fileurl, ())  
			newdoc.close(True)		
			return True
		else:
			msg = "シート「{}」が存在しません。".format(sheetname)	
			componentwindow = controller.ComponentWindow
			msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, ERRORBOX, MessageBoxButtons.BUTTONS_OK, "myRs", msg)
			msgbox.execute()	
			return False
	return detachSheet
def drowBorders(sheet, cellrange, borders):  # ターゲットを交点とする行列全体の外枠線を描く。
	cell = cellrange[0, 0]  # セル範囲の左上端のセルで判断する。
	consts = getConsts(sheet, cell)
	sectionname = consts.sectionname	
	if sectionname in ("M", ):
		return	
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = borders	
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
	rangeaddress = cellrange.getRangeAddress()  # セル範囲アドレスを取得。
	if cell.getPropertyValue("CellBackColor") in (-1, commons.COLORS["cyan10"]):  # 背景色がないか薄緑色の時。
		if sectionname in ("A", "B", "E"):
			sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。		
		elif sectionname in ("D", ):
			sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。			
			sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。	
		elif sectionname in ("C", ):		
			sheet[1:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。				
		cellrange.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。
