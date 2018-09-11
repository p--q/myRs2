#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 一覧シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import glob, os, unohelper 
from itertools import chain
from indoc import commons, datedialog, ent, keika, karute
from com.sun.star.accessibility import AccessibleRole  # 定数
from com.sun.star.awt import MouseButton, MessageBoxButtons, MessageBoxResults, ScrollBarOrientation # 定数
from com.sun.star.awt.MessageBoxType import QUERYBOX  # enum
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH, HIRAGANA_KATAKANA  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.sheet.CellDeleteMode import LEFT as delete_left, ROWS as delete_rows  # enum
from com.sun.star.table.CellHoriJustify import LEFT  # enum
from com.sun.star.table import CellVertJustify2  # 定数
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
class Ichiran():  # シート固有の値。
	def __init__(self):
		self.menurow  = 0  # メニュー行インデックス。
		self.splittedrow = 2  # 分割行インデックス。
		self.sumicolumn = 0  # 済列インデックス。
		self.yocolumn = 1  # 予列インデックス。
		self.idcolumn = 2  # ID列インデックス。	
		self.kanjicolumn = 3  # 漢字列インデックス。
		self.kanacolumn = 4  # カナ列インデックス。	
		self.datecolumn = 5  # 入院日列インデックス。
		self.hospdayscolumn = 6  # 在院日数列インデックス。
		self.checkstartcolumn = 7  # チェック列開始列インデックス。
		self.memostartcolumn = 21  # メモ列開始列インデックス。
	def setSheet(self, sheet):  # 逐次変化する値。
		self.sheet = sheet
		cellranges = sheet[self.splittedrow:, self.idcolumn].queryContentCells(CellFlags.STRING)  # ID列の文字列が入っているセルに限定して抽出。
		backcolors = commons.COLORS["blue3"], commons.COLORS["skyblue"], commons.COLORS["red3"]  # ジェネレーターに使うので順番が重要。
		gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in backcolors)
		headers = next(gene, None), next(gene, None), next(gene, None)
		if None in headers:  # Noneがある時。
			rownames = "青", "スカイブルー", "赤"
			raise RuntimeError("{0}行が取得できません。\n{0}色の背景色のID列に何らかの文字列をいれてください。".format(rownames[headers.index(None)]))
		self.bluerow, self.skybluerow, self.redrow = headers
		cellranges = sheet[:, self.idcolumn].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # ID列の文字列が入っているセルに限定して抽出。数値の時もありうる。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # ID列の最終行インデックス+1を取得。
VARS = Ichiran()
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない(リスナー追加後でも)。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	initSheet(activationevent.ActiveSheet, xscriptcontext)
def initSheet(sheet, xscriptcontext):  # documentevent.pyから呼び出す。 
	sheet["C1:G1"].setDataArray((("済をﾘｾｯﾄ", "検予を反映", "予をﾘｾｯﾄ", "入力支援", "退院ﾘｽﾄ"),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	annotations = sheet.getAnnotations()
	doc = xscriptcontext.getDocument()
	yoteisheet = doc.getSheets()["予定"]
	yoteiids = [i.getString().split(" ")[0] for i in yoteisheet.getAnnotations()]  # 予定シートのコメントにあるIDをすべて取得。
	for i in annotations:  # すべてのコメントについて。予定シートにない予定を削除する。
		if i.getString().endswith("面談"):
			if not sheet[i.getPosition().Row, VARS.idcolumn].getString() in yoteiids:  # 予定シートにないIDの時。
				i.getParent().clearContents(CellFlags.ANNOTATION)
	sheet[VARS.splittedrow:, VARS.checkstartcolumn:VARS.memostartcolumn].setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.CENTER))  # チェック列固定行より下、全て左寄せ、上下中央揃えにする。
	refreshCounts()  # 一覧シートのカウントを更新する。
	sheet["Y1:Z1"].setPropertyValue("CharColor", commons.COLORS["silver"])  # カウントの文字色を設定。
	sheet["Y2:Z2"].setPropertyValue("CharColor", commons.COLORS["skyblue"])  # カウントの文字色を設定。	
	accessiblecontext = doc.getCurrentController().ComponentWindow.getAccessibleContext()  # コントローラーのアトリビュートからコンポーネントウィンドウを取得。
	for i in range(accessiblecontext.getAccessibleChildCount()): 
		childaccessiblecontext = accessiblecontext.getAccessibleChild(i).getAccessibleContext()
		if childaccessiblecontext.getAccessibleRole()==AccessibleRole.SCROLL_PANE:
			for j in range(childaccessiblecontext.getAccessibleChildCount()): 
				child2 = childaccessiblecontext.getAccessibleChild(j)
				childaccessiblecontext2 = child2.getAccessibleContext()
				if childaccessiblecontext2.getAccessibleRole()==AccessibleRole.SCROLL_BAR:  # スクロールバーの時。
					if child2.getOrientation()==ScrollBarOrientation.VERTICAL:  # 縦スクロールバーの時。
						if childaccessiblecontext2.getBounds().Height>0:  # 右上枠の縦スクロールバーのHeghtが0になっている。
							child2.setValue(0)  # 縦スクロールバーを一番上にする。
							return  # breakだと二重ループは抜けれない。
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	if enhancedmouseevent.Buttons==MouseButton.LEFT:
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			celladdress = selection.getCellAddress()
			r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。				
			if enhancedmouseevent.ClickCount==1:  # 左シングルクリックの時。
				if r>=VARS.splittedrow and r not in (VARS.bluerow, VARS.skybluerow, VARS.redrow):
					if c<VARS.memostartcolumn or c not in (VARS.kanjicolumn, VARS.datecolumn, VARS.hospdayscolumn):
						txt = selection.getString()
						if txt:
							ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
							smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
							systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。						
							if c==VARS.idcolumn:  # ID列の時。
								systemclipboard.setContents(commons.TextTransferable(txt), None)  # クリップボードにIDをコピーする。
							elif c==VARS.kanacolumn:  # カナ名列の時。
								idtxt = VARS.sheet[r, VARS.idcolumn].getString()
								if idtxt:
									transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
									kanatxt = commons.convertKanaFULLWIDTH(transliteration, txt)  # カナ名を半角からスペースを削除して全角にする。
									systemclipboard.setContents(commons.TextTransferable("".join((kanatxt, idtxt))), None)  # クリップボードにカナ名+IDをコピーする。	
						if r<VARS.emptyrow:  # ID列が入力されている時のみ実行するもの。
							if c==VARS.sumicolumn:  # 済列の時。
								items = [("待", "skyblue"), ("済", "silver"), ("未", "black")]
								items.append(items[0])  # 最初の要素を最後の要素に追加する。
								dic = {items[i][0]: items[i+1] for i in range(len(items)-1)}  # 順繰り辞書の作成。				
								newtxt = dic[txt][0]							
								selection.setString(newtxt)
								VARS.sheet[r, :].setPropertyValue("CharColor", commons.COLORS[dic[txt][1]])						
								refreshCounts()  # カウントを更新する。	
							elif c==VARS.yocolumn:  # 予列の時。
								if txt:
									selection.clearContents(CellFlags.STRING)  # 予をクリア。
								else:  # セルの文字列が空の時。
									selection.setString("予")
							elif c>=VARS.checkstartcolumn:
								sClickCheckCol(enhancedmouseevent, xscriptcontext)
			elif enhancedmouseevent.ClickCount==2:  # 左ダブルクリックの時。まずselectionChanged()が発火している。
				if r==VARS.menurow and c<VARS.checkstartcolumn:  # メニューセルの時。:
					return wClickMenu(enhancedmouseevent, xscriptcontext)
				elif r<VARS.splittedrow or r in (VARS.bluerow, VARS.skybluerow, VARS.redrow):  # 分割行より上または区切り行の時。
					return False # 何もしない。
				elif c<VARS.checkstartcolumn:  # チェック列より左の時。
					return wClickIDCol(enhancedmouseevent, xscriptcontext)
				return False  # セル編集モードにしない。
	return True  # セル編集モードにする。	
def wClickMenu(enhancedmouseevent, xscriptcontext):
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
	sheet = VARS.sheet
	controller = doc.getCurrentController()  # コントローラの取得。
	sheets = doc.getSheets()  # シートコレクションを取得。
	txt = selection.getString()  # クリックしたセルの文字列を取得。	
	if txt=="検予を反映":  # 経過シートから本日の検予を取得。
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
		functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。
		cellranges = sheet[VARS.splittedrow:, VARS.idcolumn].queryContentCells(CellFlags.STRING)  # ID列に文字列が入っているセルを取得。
		headerrow = sheet[VARS.menurow, VARS.checkstartcolumn:VARS.memostartcolumn].getDataArray()[0]  # チェック列のヘッダーのタプルを取得。
		eketsucol, dokueicol, ketuekicol, gazocol, shochicol, echocol, ecgcol, wardcol\
			= [headerrow.index(i) for i in ("ｴ結", "読影", "血液", "画像", "処置", "ｴｺ", "ECG", "病棟")]  # headerrowタプルでのインデックスを取得。
		todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。	
		keikavars = keika.VARS
		dayrow = keikavars.dayrow
		splittedcolumn = keikavars.splittedcolumn
		if len(cellranges)>0:  # ID列のセル範囲が取得出来ている時。
			iddatarows = cellranges[0].getDataArray()  # ID列のデータ行のタプルを取得。空行がないとする。
			checkrange = sheet[VARS.splittedrow:VARS.splittedrow+len(iddatarows), VARS.checkstartcolumn:VARS.memostartcolumn]  # チェック列範囲を取得。	
			datarows = list(map(list, checkrange.getDataArray())) # 各行をリストにして取得。既に済がある時は書き換えない。
			for r, idtxt in enumerate(chain.from_iterable(iddatarows)):  # 各ID列について。rは相対インデックス。
				if idtxt.isdigit():  # IDがすべて数字の時。
					sheetname = "{}経".format(idtxt)  # 経過シート名を作成。
					if not sheetname in sheets:  # 経過シートがない時は次のループに行く。
						continue
					keikasheet = sheets[sheetname]  # 経過シートを取得。
					startdatevalue = int(keikasheet[dayrow, splittedcolumn].getValue())  # 日付行の最初のセルから日付のシリアル値の取得。
					keikadatarows = keikasheet[dayrow+1:dayrow+3, splittedcolumn+todayvalue-startdatevalue].getDataArray()  # 今日の日付列のセル範囲の値を取得。
					if not "済" in datarows[r][ketuekicol]:  # 済があるときは書き換えない。
						datarows[r][ketuekicol] = keikadatarows[0][0]  # 血液。
					s = keikadatarows[1][0]  # 2行目を取得。
					if not "済" in datarows[r][gazocol]:  # 済があるときは書き換えない。
						for i in commons.GAZOs:  # 読影のない画像。
							if i in s:  # 経過列に文字列がある時。
								if not i in datarows[r][gazocol]:  # まだ文字列iが追加されていない時。
									datarows[r][gazocol] += i
							else:  # 経過列に文字列ないときはiを削除する。
								datarows[r][gazocol] = datarows[r][gazocol].replace(i, "")
						for i in commons.GAZOd:  # 読影のある画像。
							if i in s:
								if not i in datarows[r][gazocol]:  # まだ文字列iが追加されていない時。
									datarows[r][gazocol] += i			
								if datarows[r][wardcol] not in ("療", "包"):					
									datarows[r][dokueicol] = "読"
							else:  # 経過列に文字列ないときはiを削除する。
								datarows[r][gazocol] = datarows[r][gazocol].replace(i, "")
					if not "済" in datarows[r][echocol]:  # 済があるときは書き換えない。				
						for i in commons.ECHOs:  # エコー。
							if i in s:
								if not i in datarows[r][echocol]:  # まだ文字列iが追加されていない時。
									datarows[r][echocol] += i		
								datarows[r][eketsucol] = "ｴ"	
							else:  # 経過列に文字列ないときはiを削除する。
								datarows[r][eketsucol] = datarows[r][eketsucol].replace(i, "")
					if not "済" in datarows[r][shochicol]:  # 済があるときは書き換えない。						
						for i in commons.SHOCHIs:  # 処置。
							if i in s:
								if not i in datarows[r][shochicol]:  # まだ文字列iが追加されていない時。
									datarows[r][shochicol] += i		
							else:  # 経過列に文字列ないときはiを削除する。
								datarows[r][shochicol] = datarows[r][shochicol].replace(i, "")	
					if not "済" in datarows[r][ecgcol]:  # 済があるときは書き換えない。							
						if "ECG" in s:  # ECG。
							if not "E" in datarows[r][ecgcol]:  # まだ文字列iが追加されていない時。
								datarows[r][ecgcol] = "E"			
							else:  # 経過列に文字列ないときはiを削除する。
								datarows[r][ecgcol] = datarows[r][ecgcol].replace("E", "")
			annotations = sheet.getAnnotations()  # コメントコレクションを取得。
			comments = [(i.getPosition(), i.getString()) for i in annotations]  # setDataArray()でコメントがクリアされるのでここでセルアドレスとコメントの文字列をタプルで取得しておく。					
			checkrange.setDataArray(datarows)  # シートに書き戻す。
			[annotations.insertNew(*i) for i in comments]  # コメントを再挿入。
	elif txt=="済をﾘｾｯﾄ":
		msg = "済列をリセットします。"
		componentwindow = controller.ComponentWindow
		msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_OK, "myRs", msg)
		if msgbox.execute()==MessageBoxResults.OK:
			sheet[VARS.splittedrow:VARS.emptyrow, :].setPropertyValue("CharColor", commons.COLORS["black"])  # 文字色を黒色にする。
			sheet[VARS.splittedrow:VARS.emptyrow, VARS.sumicolumn].setDataArray([("未",)]*(VARS.emptyrow-VARS.splittedrow))  # 済列をリセット。
			searchdescriptor = sheet.createSearchDescriptor()
			searchdescriptor.setSearchString("済")
			cellranges = sheet[VARS.splittedrow:VARS.emptyrow, VARS.checkstartcolumn:VARS.memostartcolumn].findAll(searchdescriptor)  # チェック列の「済」が入っているセル範囲コレクションを取得。
			cellranges.setPropertyValue("CharColor", commons.COLORS["silver"])
			refreshCounts()  # 一覧シートのカウントを更新する。
	elif txt=="予をﾘｾｯﾄ":
		sheet[VARS.splittedrow:VARS.emptyrow, VARS.sumicolumn+1].clearContents(CellFlags.STRING)  # 予列をリセット。
	elif txt=="入力支援":
		
		pass  # 入力支援odsを開く。
	
	elif txt=="退院ﾘｽﾄ":
		controller.setActiveSheet(sheets["退院"])
	return False  # セル編集モードにしない。	
def wClickIDCol(enhancedmouseevent, xscriptcontext):
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	sheet = VARS.sheet
	celladdress = selection.getCellAddress()
	r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。		
	idtxt, kanjitxt, kanatxt, datevalue, hospdays = sheet[r, VARS.idcolumn:VARS.checkstartcolumn].getDataArray()[0]  # 日付はfloatで返ってくる。	
	if isinstance(datevalue , float):  # 入院日列がfloatの時。つまり日付シリアル値が取得出来た時。
		datevalue = int(datevalue)  # 計算しにくいのでdatevalueがあるときはfloatを整数にしておく。	
	if c==VARS.idcolumn:  # ID列の時。
		return True  # セル編集モードにする。		
	elif c==VARS.kanjicolumn:  # 漢字列の時。カルテシートをアクティブにする、なければ作成する。カルトシート名はIDと一致。	
		doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
		sheets = doc.getSheets()  # シートコレクションを取得。			
		if hospdays and idtxt in sheets:  # 経過列があり、かつ、ID名のシートが存在する時。
			doc.getCurrentController().setActiveSheet(sheets[idtxt])  # カルテシートをアクティブにする。
		else:  # 在院日数列が空欄の時、または、カルテシートがない時。
			if all((idtxt, kanjitxt, kanatxt, datevalue)):  # ID、漢字名、カナ名、入院日、すべてが揃っている時。	
				kanjitxt, kanatxt = fillColumns(enhancedmouseevent, xscriptcontext, idtxt, kanjitxt, kanatxt, datevalue)
				karutesheet = commons.getKaruteSheet(doc, idtxt, kanjitxt, kanatxt, datevalue)  # カルテシートを取得。
				doc.getCurrentController().setActiveSheet(karutesheet)  # カルテシートをアクティブにする。	
			else:
				return True  # セル編集モードにする。		
	elif c==VARS.kanacolumn:  # カナ名列の時。
		return True  # セル編集モードにする。		
	elif c==VARS.datecolumn:  # 入院日列の時。
		datedialog.createDialog(enhancedmouseevent, xscriptcontext, "入院日", "YYYY-MM-DD")		
	elif c==VARS.hospdayscolumn:  
		newsheetname = "".join([idtxt, "経"])  # 経過シート名を取得。
		doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
		sheets = doc.getSheets()  # シートコレクションを取得。	
		if hospdays and newsheetname in sheets:  # 経過列がすでにあり、かつ、経過シートがある時。
			doc.getCurrentController().setActiveSheet(sheets[newsheetname])  # 経過シートをアクティブにする。
		else:  # 経過シートがなければ作成する。
			if all((idtxt, kanjitxt, kanatxt, datevalue)):  # ID、漢字名、カナ名、入院日、すべてが揃っている時。								
				kanjitxt, kanatxt = fillColumns(enhancedmouseevent, xscriptcontext, idtxt, kanjitxt, kanatxt, datevalue)
				keikasheet = commons.getKeikaSheet(xscriptcontext, doc, idtxt, kanjitxt, kanatxt, datevalue)  # 経過シートを取得。
				doc.getCurrentController().setActiveSheet(keikasheet)  # 経過シートをアクティブにする。						
	return False  # セル編集モードにしない。		
def sClickCheckCol(enhancedmouseevent, xscriptcontext):
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	txt = selection.getString()  # クリックしたセルの文字列を取得。	
	c = selection.getCellAddress().Column  # selectionの行と列のインデックスを取得。		
	dic = {\
		"病棟": ["", "待", "療", "包", "共", "生"],\
		"ｴ結": ["", "ｴ", "済"],\
		"読影": ["", "未", "読", "済"],\
		"退処": ["", "済", "△", "待"],\
		"血液": ["", "尿", "○", "済"],\
		"ECG": ["", "E", "済"],\
		"糖尿": ["", "糖"],\
		"熱発": ["", "熱"],\
		"計書": ["", "済", "共", "未"],\
		"面談": ["", "面"],\
		"便指": ["", "済", "少", "無", "共"]\
	}
	header = VARS.sheet[VARS.menurow, c].getString()  # 行インデックス0のセルの文字列を取得。
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
def fillColumns(enhancedmouseevent, xscriptcontext, idtxt, kanjitxt, kanatxt, datevalue):  # kanjitxtとkanatxtは半角にして返す。
	sheet = VARS.sheet
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
	localestruct = Locale(Language = "ja", Country = "JP")
	transliteration.loadModuleNew((HIRAGANA_KATAKANA,), localestruct)  # 変換モジュールをロード。	
	kanatxt = transliteration.transliterate(kanatxt, 0, len(kanatxt), [])[0]  # ひらがなをカタカナに変換		
	transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), localestruct)
	kanatxt = transliteration.transliterate(kanatxt, 0, len(kanatxt), [])[0]  # 半角に変換
	r = enhancedmouseevent.Target.getCellAddress().Row				
	cellstringaddress = sheet[r, VARS.datecolumn].getPropertyValue("AbsoluteName").split(".")[-1].replace("$", "")  # 入院日セルの文字列アドレスを取得。
	cell = sheet[r, VARS.hospdayscolumn]
	cell.setFormula("=TODAY()+1-{}".format(cellstringaddress))  #  在院日数列に式を代入。	
	createFormatKey = commons.formatkeyCreator(xscriptcontext.getDocument())
	cell.setPropertyValue("NumberFormat", createFormatKey('0" ";[RED]-0" "'))  # 在院日数列の書式を設定。 
	kanjitxt = kanjitxt.strip().replace("　", " ")  # 前後のスペースを削除して、文字列間の全角スペースを半角スペースに変換する。
	datarow = "未", "", idtxt, kanjitxt, kanatxt, datevalue  # 他の列を追加。								
	sheet[r, :VARS.hospdayscolumn].setDataArray((datarow,))
	return kanjitxt, kanatxt
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	selection = eventobject.Source.getSelection()
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。	
		VARS.setSheet(selection.getSpreadsheet()) 	
		drowBorders(selection)  # 枠線の作成。
def drowBorders(selection):  # ターゲットを交点とする行列全体の外枠線を描く。
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上端のセルアドレスを取得。
	r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。	
	sheet = VARS.sheet
	if r==VARS.menurow and c<VARS.checkstartcolumn:  # メニューセルの時。
		return  # 何もしない。
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = commons.createBorders()
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
	rangeaddress = selection.getRangeAddress()  # 選択範囲のセル範囲アドレスを取得。
	if r in (VARS.bluerow, VARS.skybluerow, VARS.redrow):  # 区切り行の時。
		return  # 罫線を引き直さない。
	if r>VARS.splittedrow-1:  # 分割行以下の時。
		sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く
	if VARS.checkstartcolumn-1<c<VARS.memostartcolumn:  # チェック列の時。
		sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。
	selection.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。
def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。マクロで変更した時は発火しない模様。	
	selection = None
	for change in changesevent.Changes:
		if change.Accessor=="cell-change":  # セルの値が変化した時。
			selection = change.ReplacedElement  # 値を変更したセルを取得。セル範囲が返るときもある。
			break
	if selection:  # 背景色をペーストしても発火するのでセル範囲が膨大になるときがある。		
		cellranges = selection.queryContentCells(CellFlags.STRING+CellFlags.DATETIME+CellFlags.VALUE+CellFlags.FORMULA)  # 内容のあるセルのみのセル範囲コレクションを取得。
		if cellranges:	
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。					
			transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
			transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))		
			transliteration2 = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
			transliteration2.loadModuleNew((HIRAGANA_KATAKANA,), Locale(Language = "ja", Country = "JP"))  # 変換モジュールをロード。		
			doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
			datecolumncells = []	
			nonkanacells = []
			sheet = selection.getSpreadsheet()
			rangeaddress = selection.getRangeAddress()	
			flg = False
			titlerows = VARS.bluerow, VARS.skybluerow, VARS.redrow
			splittedrow = VARS.splittedrow
			idcolumn = VARS.idcolumn
			kanacolumn = VARS.kanacolumn
			datecolumn = VARS.datecolumn
			hospdayscolumn = VARS.hospdayscolumn
			for rangeaddress in cellranges.getRangeAddresses():
				for r in range(rangeaddress.StartRow, rangeaddress.EndRow+1):  # 行インデックスについてイテレート。				
					for c in range(rangeaddress.StartColumn, rangeaddress.EndColumn+1):  # 列インデックスについてイテレート。			
						if r>splittedrow-1 and r not in titlerows:  # 分割行を含めてその下、かつ、タイトル行でない、時。
							cell = sheet[r, c]
							txt = cell.getString()  # セルの文字列を取得。			
							if c==idcolumn:  # ID列の時。
								txt = transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換。
								if txt.isdigit():  # 数値の時のみ。空文字の時0で埋まってしまう。
									cell.setString("{:0>8}".format(txt))  # 数値を8桁にして文字列として代入し直す。
							elif c==kanacolumn:  # カナ列の時。
								txt = transliteration2.transliterate(txt, 0, len(txt), [])[0]  # ひらがなをカタカナに変換。
								txt = transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換
								if all(map(lambda x: chr(0xFF61)<=x<=chr(0xFF9F), txt.replace(" ", ""))):  # すべて半角カタカナであることを確認。スペースは除去して評価する。
									cell.setString(txt)
								else:
									nonkanacells.append(cell)
							elif c==datecolumn:  # 日付列の時。
								datecolumncells.append(cell)
							if idcolumn<=c<hospdayscolumn:  # ID列から入院日列までのセルが含まれている時。
								flg = True
			cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
			if datecolumncells:  # 日付列のセルがある時。
				cellranges.addRangeAddresses([i.getRangeAddress() for i in datecolumncells], False)
				cellranges.setPropertyValues(("NumberFormat", "HoriJustify"), (commons.formatkeyCreator(doc)('YYYY-MM-DD'), LEFT))  # カルテシートの入院日の書式設定。左寄せにする。
			if flg:  # ID列から入院日列までのセルが含まれている時背景色を再設定。
				ranges = [sheet[titlerows[2]+1:, idcolumn:hospdayscolumn]]  # 赤行より下のセル範囲。
				if splittedrow<titlerows[0]:  # 固定行から青行の上行までのセル範囲。
					ranges.append(sheet[splittedrow:titlerows[0], idcolumn:hospdayscolumn])
				if titlerows[0]+1<titlerows[1]:  # 青行からスカイブルー行の上行までのセル範囲。
					ranges.append(sheet[titlerows[0]+1:titlerows[1], idcolumn:hospdayscolumn])
				if titlerows[1]+1<titlerows[2]:  # スカイブルー行から赤行の上行までのセル範囲。
					ranges.append(sheet[titlerows[1]+1:titlerows[2], idcolumn:hospdayscolumn])
				cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
				cellranges.addRangeAddresses([i.getRangeAddress() for i in ranges], False)
				cellranges.setPropertyValue("CellBackColor", commons.COLORS["cyan10"])
			if nonkanacells:
				msg = "ｶﾅ名列にはカタカナかひらながのみ入力してください。"
				commons.showErrorMessageBox(doc.getCurrentController(), msg)	
def refreshCounts():  # カウントを更新する。
	sheet = VARS.sheet
	datarows = [["総数", 0, "済", 0], ["未", 0, "待", 0]]
	datarange = sheet[VARS.splittedrow:VARS.emptyrow, VARS.sumicolumn]
	searchdescriptor = sheet.createSearchDescriptor()
	counts = []
	for txt in ("済", "待"):  # 未はタイトル行にも入っているので正しく計算できない。
		searchdescriptor.setSearchString(txt)  # 戻り値はない。
		cellranges = datarange.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
		c = len([i for i in cellranges.getCells()]) if cellranges else 0  # セルで数えないといけない。
		counts.append(c)
	counts.append(VARS.emptyrow-VARS.splittedrow-3-sum(counts))  # 済、待、未、の順に数が入る。
	datarows[0][1] = sum(counts)
	datarows[0][3] = counts[0]
	datarows[1][1] = counts[2]
	datarows[1][3] = counts[1]
	sheet[:2, VARS.memostartcolumn:VARS.memostartcolumn+4].setDataArray(datarows)	
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。	
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
	selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
	del contextmenu[:]  # contextmenu.clear()は不可。
	rangeaddress = selection.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。
	startrow, startcolumn = rangeaddress.StartRow, rangeaddress.StartColumn  # 選択範囲の左上セルだけで判断する。
	if startrow<VARS.splittedrow:  # 固定行より上の時。
		if contextmenuname=="cell" and selection.supportsService("com.sun.star.sheet.SheetCell"):
			txt = selection.getString()  # 分割行より上、かつ、セルを右クリック、かつ、単一セル
			if txt=="ｶﾅ名":  # ｶﾅ名、のセルの時。
				addMenuentry("ActionTrigger", {"Text": "ﾌﾘｶﾞﾅ辞書設定", "CommandURL": baseurl.format("entry12")}) 
			elif txt=="読影":
				addMenuentry("ActionTrigger", {"Text": "済をリセット", "CommandURL": baseurl.format("entry14")}) 	
			return EXECUTE_MODIFIED
	elif startrow in (VARS.bluerow, VARS.skybluerow, VARS.redrow):  # タイトル行の時はコンテクストメニューを表示しない。
		return EXECUTE_MODIFIED
	elif contextmenuname=="cell":  # セルのとき。セル範囲も含む。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
			if startcolumn in (VARS.yocolumn,):  # 予列の時。
				addMenuentry("ActionTrigger", {"Text": "退院ﾘｽﾄへ", "CommandURL": baseurl.format("entry1")}) 	
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。	
			elif startcolumn in (VARS.hospdayscolumn,):  # 経過列の時。
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
				transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
				doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。
				idtxt, dummy, kanatxt = sheet[startrow, VARS.idcolumn:VARS.datecolumn].getDataArray()[0]			
				addMenuentry("ActionTrigger", {"Text": "経過ｼｰﾄをArchiveへ", "CommandURL": baseurl.format("entry2")}) 
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。	
				for i, systempath in enumerate(glob.iglob(commons.createKeikaPathname(doc, transliteration, idtxt, kanatxt, "{}{}経_*開始.ods"), recursive=True)):  # アーカイブフォルダ内の経過ファイルリストを取得する。
					addMenuentry("ActionTrigger", {"Text": os.path.basename(systempath), "CommandURL": baseurl.format("entry{}".format(21+i))}) 
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。	
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"Text": "クリア", "CommandURL": baseurl.format("entry11")}) 
	elif contextmenuname=="rowheader" and len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 行ヘッダーのとき、かつ、選択範囲の列数がシートの列数が一致している時。	
		if startrow>VARS.emptyrow-1:
			commons.cutcopypasteMenuEntries(addMenuentry)
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			commons.rowMenuEntries(addMenuentry)
			return EXECUTE_MODIFIED
		elif startrow<VARS.bluerow:  # 未入院
			addMenuentry("ActionTrigger", {"Text": "新入院へ", "CommandURL": baseurl.format("entry3")})  
		elif startrow<VARS.skybluerow:  # Stable
			addMenuentry("ActionTrigger", {"Text": "Unstableへ", "CommandURL": baseurl.format("entry4")})
			addMenuentry("ActionTrigger", {"Text": "新入院へ", "CommandURL": baseurl.format("entry5")})	
		elif startrow<VARS.redrow:  # Unstable
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
	desktop = xscriptcontext.getDesktop()
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	selection = controller.getSelection()  # 選択範囲を取得。
	rangeaddress = selection.getRangeAddress()  # 選択範囲のアドレスを取得。
	r = rangeaddress.StartRow
	if entrynum<3:  # セルのコンテクストメニュー。
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。			
		transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
		functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
		sheets = doc.getSheets()
		datarow = sheet[r, VARS.idcolumn:VARS.hospdayscolumn].getDataArray()[0]   # ダブルクリックした行をID列からｶﾅ名列までのタプルを取得。
		idtxt, kanjitxt, kanatxt, datevalue = datarow
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
		componentwindow = controller.ComponentWindow
		if entrynum==1:  # 退院リストへ。
			msg = "{} {}のシートをファイルに切り出します。".format(kanjitxt, kanatxt)
			msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_CANCEL, "myRs", msg)			
			if msgbox.execute()==MessageBoxResults.OK:
				flgs = []
				newsheetname = "{}{}_{}入院".format(kanatxt, idtxt, datetxt)  # 新しいシート名を取得。
				flgs.append(detachSheet(idtxt, newsheetname))  # 切り出したシートのfileurlを取得。
				newsheetname = "{}{}経_{}開始".format(kanatxt, idtxt, datetxt)  # 新しいシート名を取得。
				flgs.append(detachSheet("".join([idtxt, "経"]), newsheetname))
				if not all(flgs):
					msg = "{} {}を退院シートに登録しますか？".format(kanjitxt, kanatxt)
					msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_NO, "myRs", msg)
					if msgbox.execute()!=MessageBoxResults.YES:  # YESでない時はここで終わる。
						sheet.removeRange(rangeaddress, delete_rows)  # ソース行を削除。
						return			
				datarow = list(datarow)
				todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。
				datarow.extend((todayvalue, "経過"))
				entsheet = sheets["退院"]  # 退院シートを取得。
				entvars = ent.VARS  # 退院シートの定数を取得。		
				entvars.setSheet(entsheet)
				entsheet[entvars.emptyrow, entvars.idcolumn:entvars.idcolumn+len(datarow)].setDataArray((datarow,))  # 退院シートにデータを代入。
				entsheet[entvars.emptyrow, entvars.datecolumn:entvars.keikacolumn].setPropertyValue("NumberFormat", commons.formatkeyCreator(doc)('YYYY-MM-DD'))  # 追加した行の日付書式を設定。
				if entvars.splittedrow<entvars.emptyrow:
					searchdescriptor = sheet.createSearchDescriptor()
					searchdescriptor.setSearchString(idtxt)  # 戻り値はない。
					cellranges = entsheet[entvars.splittedrow:entvars.emptyrow, entvars.idcolumn].findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
					if cellranges:  # ID列に同じIDがすでにある時。
						[entsheet.removeRange(i, delete_rows) for i in cellranges.getRangeAddresses()]  # 同じIDの行を削除。
				sheet.removeRange(rangeaddress, delete_rows)  # 移動したソース行を削除。
		elif entrynum==2:  # 経過ｼｰﾄをArchiveへ。
			msg = "{}{}の経過シートをファイルに切り出します。".format(kanatxt, idtxt)
			msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_CANCEL, "myRs", msg)					
			if msgbox.execute()==MessageBoxResults.OK:
				newsheetname = "{}{}経_{}開始".format(kanatxt, idtxt, datetxt)  # 新しいシート名を取得。
				detachSheet("".join([idtxt, "経"]), newsheetname)
	elif entrynum>20:  # startentrynum以上の数値の時はアーカイブファイルを開く。
		startentrynum = 21
		c = entrynum - startentrynum  # コンテクストメニューからファイルリストのインデックスを取得。
		idtxt, dummy, kanatxt = sheet[r, VARS.idcolumn:VARS.datecolumn].getDataArray()[0]
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
		transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
		for i, systempath in enumerate(glob.iglob(commons.createKeikaPathname(doc, transliteration, idtxt, kanatxt, "{}{}経_*開始.ods"), recursive=True)):  # アーカイブフォルダ内の経過ファイルリストを取得する。
			if i==c:  # インデックスが一致する時。
				desktop.loadComponentFromURL(unohelper.systemPathToFileUrl(systempath), "_blank", 0, ())  # ドキュメントを開く。
				break
	elif entrynum==3:  # 未入院から新入院に移動。
		commons.toNewEntry(sheet, rangeaddress, VARS.bluerow, VARS.emptyrow)
	elif entrynum==4:  # StableからUnstableへ移動。
		commons.toOtherEntry(sheet, rangeaddress, VARS.skybluerow, VARS.redrow)
	elif entrynum==5:  # Stableから新入院へ移動。 
		commons.toNewEntry(sheet, rangeaddress, VARS.skybluerow, VARS.emptyrow)
	elif entrynum==6:  # UnstableからStableへ移動。
		commons.toOtherEntry(sheet, rangeaddress, VARS.redrow, VARS.skybluerow)
	elif entrynum==7:  # Unstableから新入院へ移動。
		commons.toNewEntry(sheet, rangeaddress, VARS.redrow, VARS.emptyrow)
	elif entrynum==8:  # 新入院から未入院へ移動。
		commons.toOtherEntry(sheet, rangeaddress, VARS.emptyrow, VARS.bluerow)
	elif entrynum==9:  # 新入院からStableへ移動。
		commons.toOtherEntry(sheet, rangeaddress, VARS.emptyrow, VARS.skybluerow)
	elif entrynum==10:  # 新入院からUnstableへ移動。
		commons.toOtherEntry(sheet, rangeaddress, VARS.emptyrow, VARS.redrow)
	elif entrynum==11:  # クリア。
		splittedrow = VARS.splittedrow
		edgerows = VARS.bluerow, VARS.skybluerow, VARS.redrow
		idcolumn = VARS.idcolumn
		datecolumn = VARS.datecolumn
		memostartcolumn = VARS.memostartcolumn
		cellflags = CellFlags.VALUE + CellFlags.DATETIME + CellFlags.STRING + CellFlags.ANNOTATION + CellFlags.FORMULA
		for i in range(rangeaddress.StartRow, rangeaddress.EndRow+1):  # 選択範囲の行インデックスをイテレート。
			for j in range(rangeaddress.StartColumn, rangeaddress.EndColumn+1):  # 選択範囲の列インデックスをイテレート。
				if i<splittedrow or i in edgerows:  # 固定行より上、または、タイトル行の時。
					continue
				elif idcolumn<=j<=datecolumn or sheet[0, j].getPropertyValue("CellBackColor")>0:  # ID列、漢字名列、ｶﾅ名列、入院日列、または、１行目に背景色があるとき、は背景色を消さない。
					sheet[i, j].clearContents(cellflags)
				elif j>=memostartcolumn:  # メモ列開始列含む右列の時。
					sheet.removeRange(sheet[i, j].getRangeAddress(), delete_left)  # セルを削除して左にずらす。
				else:  # それ以外の時。
					sheet[i, j].clearContents(511)  # 範囲をすべてクリアする。
	elif entrynum==12:  # ﾌﾘｶﾞﾅ辞書設定。
		
		pass
	
	elif entrynum==14:  # 読影列の済をリセット。読影列の済を消去し、4F列が○の時未にする。
		headerrow = sheet[VARS.menurow, VARS.checkstartcolumn:VARS.memostartcolumn].getDataArray()[0]  # チェック列のヘッダーのタプルを取得。
		wardcol, = [headerrow.index(i) for i in ("病棟",)]  # headerrowタプルでのインデックスを取得。
		searchdescriptor = sheet.createSearchDescriptor()
		searchdescriptor.setSearchString("療")  # 戻り値はない。	
		splittedrow = VARS.splittedrow
		cellranges = sheet[splittedrow:VARS.emptyrow, VARS.checkstartcolumn+wardcol].findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
		if cellranges:  # 病棟列に寮が入っているセルがある時。
			c = selection.getCellAddress().Column  # 選択セルの列インデックスを取得。
			datarange = sheet[splittedrow:VARS.emptyrow, c]  
			datarows = list(datarange.getDataArray())  # 選択列の行のタプルをリストにして取得。
			for i in cellranges.getCells():
				j = i.getCellAddress().Row - splittedrow  # 病棟列に療が入っているインデックスを取得。				
				datarows[j] = ("未",)  # 行ごと入れ替える。
			datarange.setDataArray(datarows)  # シートに戻す。				
def createDetachSheet(desktop, controller, doc, sheets, kanadirpath):
	def detachSheet(sheetname, newsheetname):
		if sheetname in sheets:  # シートがある時。
			existingsheet = sheets[sheetname]  # カルテシートを取得。
			existingsheet.setName(newsheetname)  # カルテシート名を変更。
			newdoc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())  # 新規ドキュメントの取得。シートの行と列の固定のためにhiddenでは開かない。
			newsheets = newdoc.getSheets()  # 新規ドキュメントのシートコレクションを取得。
			newsheets.importSheet(doc, newsheetname, 0)  # 新規ドキュメントにシートをコピー。
			del newsheets["Sheet1"]  # 新規ドキュメントのデフォルトシートを削除する。 
			del sheets[newsheetname]  # 切り出したカルテシートを削除する。 
			systempath = os.path.join(kanadirpath, "{}.ods".format(newsheetname))
			if os.path.exists(systempath):  # すでにファイルが存在する時。
				msg = "{}はすでにバックアップ済です。\n上書きしますか？".format(newsheetname)
				componentwindow = controller.ComponentWindow
				msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "myRs", msg)
				if msgbox.execute()!=MessageBoxResults.YES:			
					return True  # 上書きしない時は、切り出したものとして返す。
			frozenrow, frozencol = 0, 0
			if sheetname.isdigit():  # シート名が数字のみの時カルテシート。
				karutevars = karute.VARS
				frozenrow, frozencol = karutevars.splittedrow, karutevars.splittedcolumn
			elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
				keikavars = keika.VARS
				frozenrow, frozencol = keikavars.splittedrow, keikavars.splittedcolumn				
			newdoccontroller = newdoc.getCurrentController()  # コピー先のドキュメントのコントローラを取得。	
			newdoccontroller.setActiveSheet(newsheets[newsheetname])  # コピーしたシートをアクティブにする。
			newdoccontroller.freezeAtPosition(frozencol, frozenrow)  # 行と列の固定をする。切り出したシートは固定がなくなっているので。			
			fileurl = unohelper.systemPathToFileUrl(systempath)
			newdoc.storeToURL(fileurl, ())  
			newdoc.close(True)		
			return True
		else:
			msg = "シート「{}」が存在しません。".format(sheetname)	
			commons.showErrorMessageBox(controller, msg)	
			return False
	return detachSheet
