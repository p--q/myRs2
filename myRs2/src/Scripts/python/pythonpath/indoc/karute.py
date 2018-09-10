#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# カルテシートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import re
from datetime import date, datetime
from indoc import commons, datedialog, historydialog, staticdialog
from com.sun.star.awt import Key, MouseButton  # 定数
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.sheet.CellDeleteMode import ROWS as delete_rows  # enum
from com.sun.star.sheet.CellInsertMode import ROWS as insert_rows  # enum
from com.sun.star.table import CellVertJustify2  # 定数
from com.sun.star.table.CellHoriJustify import CENTER, LEFT, RIGHT  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
class Karute():  # シート固有の定数設定。
	def __init__(self):
		self.splittedrow = 2  # 分割行インデックス。
		self.sharpcolumn = 1  # 行区切列インデックス。
		self.datecolumn = 2  # 日付列インデックス。
		self.problemcolumn = 3  # プロブレム列インデックス。
		self.phrasecolumn = 4  # 定型句列インデックス。
		self.articlecolumn = 5  # 記事列インデックス。
		self.historycolumn = 6  # 履歴インデックス。
		self.insertdatecolumn = 7  # 日付挿入インデックス。
		self.replacedatecolumn = 8  # 日付前へ列インデックス。
		self.splittedcolumn = 9  # 分割列インデックス。コントローラーから動的取得が正しく出来ない。
		self.stringlength = 125  # 1セルあたりの文字数。
		self.dateformat = "%Y-%m-%d %H:%M:%S Copied"  # 記事をコピーした日時の書式。
	def setSheet(self, sheet): # 逐次変化する値。
		self.sheet = sheet	
		cellranges = sheet[self.splittedrow:, self.datecolumn].queryContentCells(CellFlags.STRING)  # Date列の文字列が入っているセルに限定して抽出。
		backcolors = commons.COLORS["blue3"], commons.COLORS["skyblue"], commons.COLORS["red3"]  # ジェネレーターに使うので順番が重要。
		gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in backcolors)
		headers = next(gene, None), next(gene, None), next(gene, None)
		if None in headers:  # Noneがある時。
			rownames = "青", "スカイブルー", "赤"
			raise RuntimeError("{0}行が取得できません。\n{0}色の背景色のID列に何らかの文字列をいれてください。".format(rownames[headers.index(None)]))
		self.bluerow, self.skybluerow, self.redrow = headers
VARS = Karute()		
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。
	doc = xscriptcontext.getDocument()
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	VARS.setSheet(sheet)
	cellrange = sheet["A1:L1"]  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	datarow = list(cellrange.getDataArray()[0])  # 行をリストで取得。
	datarow[VARS.datecolumn] = "一覧へ"
	datarow[VARS.problemcolumn] = "経過へ"
	datarow[VARS.splittedcolumn] = "COPY"
	datarow[VARS.splittedcolumn+1] = "退院ｻﾏﾘ"
	datarow[VARS.splittedcolumn+2] = "#分離"
	sheet[0, VARS.splittedcolumn+1].setPropertyValue("CellBackColor", -1)  # 退院ｻﾏﾘボタンの背景色をクリアする。
	cellrange.setDataArray((datarow,))  # 行をシートに戻す。
	# コピー日時セルの色を設定。
	copieddatecell = sheet[0, VARS.articlecolumn]  # コピー日時セルを取得。
	copieddatetxt = copieddatecell.getString()  # コピー日時セルの文字列を取得。
	try:
		copieddatetime = datetime.strptime(copieddatetxt, VARS.dateformat)  # コピーした日時を取得。曜日の文字列が含まれるととOSによってValueErrorになる。
	except:
		copieddatetime = None
	if copieddatetime:  # datetimeオブジェクトが取得出来た時。
		now = datetime.now()  # 現在の日時を取得。
		if copieddatetime.date()<now.date():  # 今日はまだコピーしていない時。
			copieddatecell.setPropertyValues(("CharColor", "CellBackColor"), (-1, commons.COLORS["magenta3"]))  # 文字色をリセットして背景色をマゼンダにする。
		elif now.hour>12 and copieddatetime.hour<12:  # 今日はコピーしていても、午後になって午前にしかコピーしていない時。
			copieddatecell.setPropertyValue("CharColor", commons.COLORS["red3"])  # 文字色を赤色にする。背景色はコピーした時にすでにライムになっているはず。
	# 本日の記事を過去の記事に移動させる。
	daterange = sheet[VARS.bluerow, VARS.articlecolumn]  # 本日の記事の日付セルを取得。
	articledatetxt = daterange.getString()  # 本日の記事の日付セルの文字列を取得。
	try:
		articledate = datetime.strptime(articledatetxt.split("(")[0], "****%Y年%m月%d日").date()  # 記事列の日付を取得。strptime()は0埋めは関係ないが曜日文字列はOS依存なので曜日は削除する。
	except:
		articledate = None
	todaydate = date.today()  # 今日のdateオブジェクトを取得。
	weekdays = "月", "火", "水", "木", "金", "土", "日"
	todaytxt = "****{}年{}月{}日({})****".format(todaydate.year, todaydate.month, todaydate.day, weekdays[todaydate.weekday()])
	if not articledate:  # 日付が取得出来なかった時。
		daterange.setString(todaytxt)  # 今日の日付を本日の記事欄に入力。	
		articledate = todaydate
	if articledate!=todaydate:  # 今日の日付でない時。
		todayarticle = sheet[VARS.bluerow+1:VARS.skybluerow, :]  # 青行とスカイブルー行の間の行のセル範囲。
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
		functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。			
		datarows = todayarticle[:, VARS.sharpcolumn:VARS.articlecolumn+1].getDataArray()  # 本日の記事欄のセルをすべて取得。
		txt = "".join(getRowTxt(functionaccess, i) for i in datarows)  # 1行の列を連結して文字列を返す。日付シリアル値を文字列に変換する。
		cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
		if txt:  # 記事の文字列があるときのみ。
			newdatarows = [(articledatetxt,)]  # 先頭行に日付を入れる。
			stringlength = VARS.stringlength  # 1セルあたりの文字数。
			newdatarows.extend((txt[i:i+stringlength],) for i in range(0, len(txt), stringlength))  # 過去記事欄へ代入するデータ。
			dest_start_ridx = VARS.redrow + 1  # 移動先の開始行インデックス。
			dest_endbelow_ridx = dest_start_ridx + len(newdatarows)  # 移動先の最終行の下行の行インデックス。
			dest_rangeaddress = sheet[dest_start_ridx:dest_endbelow_ridx, VARS.articlecolumn].getRangeAddress()  # 挿入前にセル範囲アドレスを取得しておく。
			sheet.insertCells(dest_rangeaddress, insert_rows)  # 赤行の下に空行を挿入。	
			sheet[dest_start_ridx:dest_endbelow_ridx, :].clearContents(511)  # 挿入した行の内容をすべて削除。挿入セルは挿入した行の上のプロパティを引き継いでいるのでリセットしないといけない。
			dest_range = sheet.queryIntersection(dest_rangeaddress)[0]  # 赤行の下の挿入行のセル範囲を取得。セル挿入後はアドレスから取得し直さないといけない。
			dest_range.setDataArray(newdatarows)  # 過去の記事に挿入する。
			cellranges.addRangeAddress(dest_range.getRangeAddress(), False)  # あとでプロパティを設定するセル範囲コレクションに追加する。
			todayarticle.clearContents(511)  # 本日の記事欄をクリア。
		cellranges.addRangeAddresses([todayarticle[:, i].getRangeAddress() for i in (VARS.datecolumn, VARS.problemcolumn, VARS.articlecolumn)], False)  # 本日の記事のDate列、プロブレム列、記事列のセル範囲コレクションを取得。
		cellranges.setPropertyValue("IsTextWrapped", True)  # セルの内容を折り返す。	
		daterange.setString(todaytxt)  # 今日の日付を本日の記事欄に入力。	
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	if enhancedmouseevent.Buttons==MouseButton.LEFT:
		selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if selection.supportsService("com.sun.star.sheet.SheetCell") and enhancedmouseevent.ClickCount==2:  # ターゲットがセル、かつ、左ダブルクリック、の時。まずselectionChanged()が発火している。
			r = selection.getCellAddress().Row
			if r<VARS.splittedrow or r==VARS.redrow:  # 分割行より上、または、赤行の時。
				return wClickMenu(enhancedmouseevent, xscriptcontext)
			elif r<VARS.redrow and r not in (VARS.bluerow, VARS.skybluerow):  # 分割行以下赤行より上、かつ、タイトル行でない時。
				return wClickCol(enhancedmouseevent, xscriptcontext)
	return True  # セル編集モードにする。
def wClickMenu(enhancedmouseevent, xscriptcontext):  # メニューセル。
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	txt = selection.getString()  # クリックしたセルの文字列を取得。	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
	sheets = doc.getSheets()  # シートコレクションを取得。	
	controller = doc.getCurrentController()
	sheet = VARS.sheet
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
	getCopyDataRows, formatArticleColumn, formatProblemList, copyCells, fullwidth_halfwidth = createCopyFuncs(xscriptcontext, functionaccess)
	if txt=="一覧へ":
		controller.setActiveSheet(sheets["一覧"])  # 一覧シートをアクティブにする。
	elif txt=="経過へ":
		newsheetname = "".join([sheet.getName(), "経"])  # 経過シート名を取得。
		if newsheetname in sheets:  # 経過シート名がある時。
			controller.setActiveSheet(sheets[newsheetname])  # 経過シートをアクティブにする。
		else:  # 経過シートの作成。
			idcelltxts = sheet[VARS.splittedrow-1, VARS.articlecolumn].getString().replace("　", " ").split(" ")  # 半角スペースで分割。
			idtxt = idcelltxts[0]  # 最初の要素を取得。
			if idtxt.isdigit():  # IDが数値のみの時。					
				if idtxt in sheets:  # ID名のシートがあるとき。
					controller.setActiveSheet(sheets[idtxt])  # カルテシートをアクティブにする。
				else:
					if len(idcelltxts)==5:  # ID、漢字姓・名、カタカナ姓・名、の5つに分割できていた時。
						kanjitxt, kanatxt = " ".join(idcelltxts[1:3]), " ".join(idcelltxts[3:])
						datevalue = sheet[VARS.splittedrow, VARS.datecolumn].getValue()
						keikasheet =  commons.getKeikaSheet(xscriptcontext, doc, idtxt, kanjitxt, kanatxt, datevalue)  # 経過シートを取得。
						controller.setActiveSheet(keikasheet)  # 経過シートをアクティブにする。
					else:
						commons.showErrorMessageBox(controller, "「ID(数値のみ) 漢字姓 名 カナ姓 名」の形式になっていません。")
			else:
				commons.showErrorMessageBox(controller, "IDが取得できませんでした。")		
	elif txt=="COPY":
		c = formatArticleColumn(sheet[VARS.bluerow+1:VARS.skybluerow, VARS.sharpcolumn:VARS.articlecolumn+1])  # 本日の記事欄の記事列を整形。追加した行数が返る。
		datarows = sheet[VARS.bluerow:VARS.skybluerow+c, VARS.sharpcolumn:VARS.articlecolumn+1].getDataArray()  # 文字数制限後の行のタプルを取得。
		copydatarows = [(datarows[0][4],)]  # 本日の記事の日付を取得。
		deletedrowcount = getCopyDataRows(copydatarows, datarows[1:], VARS.bluerow+1)  # クリップボードに取得する行の取得と空行の削除。削除された行数を返す。
		if deletedrowcount>0:  # 削除した行があるとき。
			startrow = VARS.skybluerow - deletedrowcount
			newrangeaddress = sheet[startrow:startrow+deletedrowcount, :].getRangeAddress()  # 挿入するセル範囲アドレスを取得。
			sheet.insertCells(newrangeaddress, insert_rows)  # 空行を挿入。	
			sheet.queryIntersection(newrangeaddress).clearContents(511)  # 追加行の内容をクリア。セル範囲アドレスから取得しないと行挿入後のセル範囲が異なってしまう。
		newdatarows = formatProblemList(VARS.splittedrow, VARS.bluerow, "****ｻﾏﾘ****")  # プロブレム欄を整形。
		startendedgerowpairs = (VARS.splittedrow, VARS.bluerow), (VARS.bluerow+1, VARS.skybluerow), (VARS.skybluerow+1, VARS.redrow)  # 赤行より上の色行以外の開始行と終了行下の行インデックスのペアのタプル。
		pcellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges") 
		rangeaddresses = [sheet[i[0]:i[1], VARS.sharpcolumn:VARS.problemcolumn+1].getRangeAddress() for i in startendedgerowpairs if i[0]<i[1]]
		if rangeaddresses:
			pcellranges.addRangeAddresses(rangeaddresses, False)  # #列からプロブレム列までのセル範囲コレクションを取得。
			pcellranges.setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.CENTER))  # 左寄せ、縦位置を中央にする。
			searchdescriptor = sheet.createSearchDescriptor()
			searchdescriptor.setSearchString("#")  # 戻り値はない。	
			cellranges = pcellranges.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
			if cellranges:  # 日付列とプロブレム列の#が入っているセル範囲コレクションがある時。
				cellranges.setPropertyValue("HoriJustify", RIGHT)  # #のセルのみ右寄せにする。
			rangeaddresses = [sheet[i[0]:i[1], VARS.articlecolumn].getRangeAddress() for i in startendedgerowpairs if i[0]<i[1]]
			if rangeaddresses:
				cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges") 
				cellranges.addRangeAddresses(rangeaddresses, False)  # 記事列のみ取得。
				cellranges.setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.TOP))  # 左寄せ、縦位置を上寄せにする。
				pcellranges.addRangeAddresses(cellranges.getRangeAddresses(), False)  # 記事列のセル範囲コレクションと合体する。
			pcellranges.setPropertyValue("IsTextWrapped", True)  # セルの内容を折り返す。
			sheet[VARS.splittedrow:VARS.redrow, 0].getRows().setPropertyValue("OptimalHeight", True)  # 内容を折り返した後の行の高さを調整。
			newdatarows.extend(copydatarows)  # 本日の記事欄をプロブレム欄の下に追加。
			copyCells(newdatarows)  # クリップボードにコピーする。
			now = datetime.now()
			datetxt = "{}-{}-{} {}:{}:{} Copied".format(now.year, now.month, now.day, now.hour, now.minute, now.second)  # コピーボタンを押した日付を入力。
			copieddatecell = sheet[0, VARS.articlecolumn]  # コピー日時セルを取得。	
			copieddatecell.setString(datetxt)
			copieddatecell.setPropertyValues(("CellBackColor", "CharColor"), (commons.COLORS["lime"], -1))  # コピー日時セルの背景色を変更。文字色をリセット。
	elif txt=="退院ｻﾏﾘ":
		newdatarows = formatProblemList(VARS.splittedrow, VARS.bluerow, "****退院ｻﾏﾘ****")  # プロブレム欄を整形。
		copieddatecell = sheet[0, VARS.articlecolumn]  # コピー日時セルを取得。	
		copyCells(newdatarows)
		selection.setPropertyValue("CellBackColor", commons.COLORS["lime"])  # 退院ｻﾏﾘボタンの背景色を変更。
	elif txt=="#分離":  # 記事列のセルの内容を#、日付、プロブレムに分解してセルに代入し直す。
		datarange = sheet[VARS.splittedrow:VARS.bluerow, :VARS.articlecolumn+1]
		datarows = datarange.getDataArray()  # 行のタプルをリストで取得。
		separateDS(doc, functionaccess, fullwidth_halfwidth, datarows)
	elif txt=="問題ﾘｽﾄへ変換":
		cellranges = sheet[VARS.redrow+1:, VARS.articlecolumn].queryContentCells(CellFlags.STRING)  # Article列の文字列が入っているセルに限定して抽出。
		if len(cellranges):  # セル範囲が取得出来た時。
			emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 最終行インデックス+1を取得。
			datarange = sheet[VARS.redrow+1:emptyrow, :VARS.articlecolumn+1]  # 赤行より下の文字のセル範囲を取得。
			datarows = datarange.getDataArray()  # ソース行を取得。次の行でシート上のデータはクリアする。			
			datarange.clearContents(CellFlags.STRING)  # コピー元の文字列をクリア。	
			problemrange = sheet[VARS.splittedrow:VARS.bluerow, VARS.sharpcolumn:VARS.articlecolumn+1]
			cellranges = problemrange.queryContentCells(CellFlags.STRING)
			emptyrow = max(i.EndRow for i in cellranges.getRangeAddresses()) + 1 if len(cellranges) else VARS.splittedrow  # 青行おり上の範囲の最下行の下行を取得。
			endrowbelow = emptyrow + len(datarows)  # 挿入後の最下行の下行インデックス。	
			sheet.insertCells(sheet[emptyrow:endrowbelow, :].getRangeAddress(), insert_rows)  # 空行を挿入。	
			sheet[emptyrow:endrowbelow, :].setPropertyValues(("CellBackColor", "CharColor"), (-1, -1))  # 追加行の背景色と文字色をクリア。	
			separateDS(doc, functionaccess, fullwidth_halfwidth, datarows, emptyrow)
	elif txt[:8].isdigit():  # 最初8文字が数値の時。
		systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
		systemclipboard.setContents(commons.TextTransferable(txt[:8]), None)  # クリップボードにIDをコピーする。							
	return False  # セルを編集モードにしない。
def separateDS(doc, functionaccess, fullwidth_halfwidth, datarows, startrow=VARS.splittedrow):
	newdatarows = []
	handleDS = createHandleDS(functionaccess)
	for datarow in datarows:
		articletxt = datarow[VARS.articlecolumn].strip()
		if articletxt:
			articletxt = fullwidth_halfwidth(articletxt)  # Article列の文字列を半角にして取得。
			if articletxt.startswith("#"):  # 記事列が#から始まっているセルの時。
				datetxt, problemtxt, newarticletxt = handleDS(articletxt.lstrip("#"))
				if datetxt:
					datarow = "", "#", datetxt, problemtxt, "", newarticletxt
				else:
					datarow = "", "", "#", problemtxt, "", newarticletxt
		if not articletxt.startswith("****"):
			newdatarows.append(datarow)
	sheet = VARS.sheet
	sheet[startrow:startrow+len(newdatarows), :VARS.articlecolumn+1].setDataArray(newdatarows)
	sheet[startrow:VARS.bluerow, VARS.sharpcolumn].setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.CENTER))  # #列の書式設定。左寄せにする。
	sheet[startrow:VARS.bluerow, VARS.datecolumn].setPropertyValues(("NumberFormat", "HoriJustify", "VertJustify"), (commons.formatkeyCreator(doc)('YYYY-M-D'), LEFT, CellVertJustify2.CENTER))  # カルテシートの日付列の書式設定。左寄せにする。
	sheet[startrow:VARS.bluerow, VARS.problemcolumn].setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.CENTER))  # Problem列の書式設定。左寄せにする。
	searchdescriptor = sheet.createSearchDescriptor()
	searchdescriptor.setSearchString("#")  # 戻り値はない。	
	cellranges = sheet[startrow:VARS.bluerow, VARS.sharpcolumn:VARS.problemcolumn+1].findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
	if cellranges:  # #列からプロブレム列の#が入っているセル範囲コレクションがある時。
		cellranges.setPropertyValues(("HoriJustify", "VertJustify"), (RIGHT, CellVertJustify2.CENTER))  # #のセルは右寄せにする。
def wClickCol(enhancedmouseevent, xscriptcontext):  # 列によって変える処理。
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	celladdress = selection.getCellAddress()
	r, c = celladdress.Row, celladdress.Column  # ダブルクリックしたセルの行インデックス、列インデックスを取得。
	sheet = VARS.sheet
	if c==0:  # 行挿入列の時。
		sheet.insertCells(sheet[r+1, :].getRangeAddress(), insert_rows)  # ダブルクリックした行の下に空行を挿入。	
		sheet[r+1, :].setPropertyValues(("CellBackColor", "CharColor"), (-1, -1))  # 追加行の背景色と文字色をクリア。						
	elif c==VARS.sharpcolumn:  # 区切列の時。
		txt = selection.getString()  # クリックしたセルの文字列を取得。
		if txt:
			selection.clearContents(CellFlags.STRING)
		else:
			selection.setString("#")
			selection.setPropertyValues(("HoriJustify", "VertJustify"), (CENTER, CellVertJustify2.CENTER))
	elif c==VARS.datecolumn:  # 日付列の時。
		datedialog.createDialog(enhancedmouseevent, xscriptcontext, "日付入力", "YYYY-M-D")	
		selection.setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.CENTER))  # 左寄せにする。
	elif c in (VARS.problemcolumn,):  # プロブレム列の時。
		if selection.getString()=="#":
			selection.setString("")
		selection.setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.CENTER))  # 左寄せにする。
		return True  # セル編集モードにする。
	elif c in (VARS.articlecolumn,):  # 記事列の時。
		return True  # セル編集モードにする。
	elif c==VARS.phrasecolumn:  # 定型句列インデックスの時。
		datarow = sheet[r, VARS.sharpcolumn:VARS.historycolumn].getDataArray()[0]   # クリックした行のデータを取得。
		if any(datarow):  # クリックした行のセルに値がある時。
			sheet.insertCells(sheet[r+1, c].getRangeAddress(), insert_rows)  # 下に空行を挿入。
			xscriptcontext.getDocument().getCurrentController().select(sheet[r+1, c])  # 新規行のセルを選択し直す。
		staticdialog.createDialog(enhancedmouseevent, xscriptcontext, "ﾌﾟﾛﾌﾞﾚﾑ", outputcolumn=VARS.problemcolumn, callback=callback_phrasecolumnCreator(xscriptcontext))
	elif c==VARS.insertdatecolumn:  # 日付挿入列の時。
		selection.setString("")  # 日付挿入列の文字列をクリア。
		datedialog.createDialog(enhancedmouseevent, xscriptcontext, "日付挿入", "YYYY-M-D", callback=callback_insertdatecolumnCreator(xscriptcontext))  # ダイアログの戻り値は取得できず、入力も待たず次のコードにいってしまう。
		selection.setPropertyValue("CharColor", commons.COLORS["white"])  # 日付挿入列の文字色を白色にする。
	elif c==VARS.replacedatecolumn:  # 日付入替列の時。
		datetxt = VARS.sheet[r, VARS.insertdatecolumn].getString()  # 日付挿入列の文字列を取得。
		if datetxt:  # 日付文字列が取得出来た時。
			articlecell = VARS.sheet[r, VARS.articlecolumn]  # 記事セルを取得。
			articletxt = articlecell.getString()  # 記事セルの文字列を取得。
			if articletxt:  # 記事列のセルに文字列がある時。
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。				
				transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
				transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))  # 全角文字を半角にするモジュールをロード。				
				articletxt = transliteration.transliterate(articletxt, 0, len(articletxt), [])[0]  # 半角に変換。
				if articletxt.endswith(datetxt):  # 記事列の最後が日付挿入列の日付で終わっている時。
					articletxt = articletxt[:-len(datetxt)].rstrip()  # すでにある日付を削って、後ろの空白を削る。				
					txts = articletxt.rsplit("｡", 1)  # 右から｡で1回分割。	
					if len(txts)>1:  # "｡"がある時。
						if txts[-1]:  # 日付の直前が｡でない時。
							articletxt = "".join((txts[0], "｡", datetxt, " ", txts[1]))  # ｡の後ろに日付を移動させる。
						else:  # 日付の直前が｡の時。txts[-1]は空文字になる。
							txts2 = txts[0].rsplit("｡", 1)  # 右から｡で再分割。	
							if len(txts2)>1:  # ｡の後ろに日付を移動させる。
								articletxt = "".join((txts2[0], "｡", datetxt, " ", txts2[1], "｡"))
							else:
								articletxt = "".join((datetxt, " ", txts2[0], "｡"))  # 先頭に日付を移動させる。
					else:  # "｡"がない時。
						articletxt = "".join((datetxt, " ", txts[0]))  # 先頭に日付を移動させる。
					articlecell.setString(articletxt)	
				controller = xscriptcontext.getDocument().getCurrentController()		
				controller.select(articlecell)
				commons.simulateKey(controller, Key.F2, 0)  # 選択セルをセル編集モードにする。	
	elif c==VARS.historycolumn:  # 履歴列の時。
		problemtxt = VARS.sheet[r, VARS.problemcolumn].getString()
		if not problemtxt:
			problemtxt = "履歴"
		historydialog.createDialog(enhancedmouseevent, xscriptcontext, problemtxt, None, VARS.articlecolumn)
	return False  # セルを編集モードにしない。
def callback_phrasecolumnCreator(xscriptcontext):	
	def callback_phrasecolumn(gridcelltxt):  # プロブレム列に、#today 心エコー:LV wall function normal、とあるのを処理する。
		selection = xscriptcontext.getDocument().getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
		sharptxt, todayvalue, problemtxt, articletxt = "", "", "", ""
		if gridcelltxt.startswith("#"):  # #から始まっている時。
			sharptxt = "#"
			gridcelltxt = gridcelltxt[1:].lstrip()  # 先頭文字を削って、先頭スペースも削る。
		if gridcelltxt.startswith("today"):  # todayで始まっている時。
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
			functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
			todayvalue = int(functionaccess.callFunction("TODAY", ()))  # シリアル値を整数で取得。floatで返る。シリアル値で入れないとsetDataArray()で日付にできない。
			gridcelltxt = gridcelltxt[len("today"):]  # todayを削る。
		if ":" in gridcelltxt:
			problemtxt, articletxt = gridcelltxt.split(":", 1)
		else:
			articletxt = gridcelltxt
		datarow = sharptxt, todayvalue, problemtxt.strip(), "", articletxt.strip()
		VARS.sheet[selection.getCellAddress().Row, VARS.sharpcolumn:VARS.articlecolumn+1].setDataArray((datarow,))
	return callback_phrasecolumn
def callback_insertdatecolumnCreator(xscriptcontext):
	def callback_insertdatecolumn(datetxt):  # 日付挿入列をダブルクリックした時に日付入力ダイアログに渡すコールバック関数。
		doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
		selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。	
		articlecell = VARS.sheet[selection.getCellAddress().Row, VARS.articlecolumn]  # 記事セルを取得。		
		articlecell.setString("".join([articlecell.getString(), datetxt]))  # 新規日付を代入。
		controller = doc.getCurrentController()
		controller.select(articlecell)
		commons.simulateKey(controller, Key.F2, 0)  # 選択セルをセル編集モードにする。
	return callback_insertdatecolumn
def createHandleDS(functionaccess):
	rgpat = r"^((([HS][0-3]?|20\d)\d[\.\-\/][01]?\d[\.\-\/][0-3]?\d)|(([HS][0-3]?|20\d)\d[\.\-\/][01]?\d)|(([HS][0-3]?|20\d)\d))[^\.\d]"  # 日付を取得する正規表現パターン。数字とピリオド以外が続く時のみ取得。
	rgx = re.compile(rgpat)	
	seprgx = re.compile(r"[\.\-\/]")  # 日付区切り。
	def _convertYear(y):
		if y.startswith("H"):
			return str(1988+int(y[1:]))  # 平成を西暦に変換して文字列で取得。
		elif y.startswith("S"):
			return str(1925+int(y[1:]))  # 昭和を西暦に変換して文字列で取得。
		return y
	def handleDS(articletxt):  # articletxtは#を除いた記事列。
		"""処理できるパターン、問題か記事かは後ろにコロンがついているかで判別する。年月日の区切り文字は.か-か/。
		記事
		問題:
		問題:記事
		H28記事
		H28問題:
		H28問題:記事	
		H28.8記事
		H28.8問題:
		H28.8問題:記事		
		H28.8.10記事
		H28.8.10問題:
		H28.8.10問題:記事		
		2018記事
		2018問題:
		2018問題:記事	
		2018.8記事
		2018.8問題:
		2018.8問題:記事		
		2018.8.10記事
		2018.8.10問題:
		2018.8.10問題:記事	
		"""
		datetxt, problemtxt, newarticletxt = "", "", ""
		m = rgx.match(articletxt)  # 日付だけのときはNoneが返る。
		ms = []
		if m:  # 日付から始まっている時。
			ms = m.groups()  # 0:いずれかの年月日パターン、1: 年月日、3: 年月、5: 年
			if ms[1]:  # 年月日の時は日付シリアル値に変換する。
				y, m, d = seprgx.split(ms[1])
				datetxt = int(functionaccess.callFunction("DATEVALUE", ("-".join([_convertYear(y), m, d]),)))  # シリアル値を整数で取得。floatで返る。シリアル値で入れないとsetDataArray()で日付にできない。区切り文字は/.-のいずれもOK。
			elif ms[3]:  # 年月の時。
				y, m = seprgx.split(ms[3])
				datetxt = "-".join([_convertYear(y), m])
			elif ms[5]:  # 年の時。
				datetxt = _convertYear(ms[5])
			else:
				datetxt = ms[0]	
			articletxt = articletxt[len(ms[0]):]  # 日付を切除。
		if ":" in articletxt:
			problemtxt, newarticletxt = [i.strip() for i in articletxt.split(":", 1)]	
		else:
			newarticletxt = articletxt
		return datetxt, problemtxt, newarticletxt
	return handleDS
def addDataRow(stringlength, sharpcell, datecell, subjectcell, articletxts, newdatarows):
	articletxt = "".join(articletxts).lstrip().replace("\n", "")  # 先頭の空白とセル内の改行文字も除去する。
	articlecells = [articletxt[i:i+stringlength] for i in range(0, len(articletxt), stringlength)]  # 文字列を制限したArticle列のジェネレーター。
	datarow = sharpcell, datecell, "", subjectcell.strip(), "", articlecells[0]  # プロブレムの1行目を取得。
	newdatarows.append(datarow)  # プロブレムの1行目を追加。
	if len(articlecells)>1:  # 複数行ある時。
		newdatarows.extend(("", "", "", "", "", i) for i in articlecells[1:])	 # 2行目以降について処理。
def getRowTxt(functionaccess, datarow):  # 1行の列を連結して文字列を返す。日付シリアル値を文字列に変換する。
	sharpcol, datecol, subjectcol, articlecol = datarow[0], datarow[1], datarow[2], datarow[4]  # #列、日付列、Problem列、記事列を取得。
	if datecol and isinstance(datecol, float):  # 日付列がfloat型のとき。
		datecol = "{} ".format("/".join([str(int(functionaccess.callFunction(i, (datecol,)))) for i in ("YEAR", "MONTH", "DAY")]))  # シリアル値をシート関数で年/月/日の文字列にする。引数のdatecolはfloatのままでよい。
	if subjectcol and subjectcol!="#":   # Subject列、かつ、#でないとき
		subjectcol = "{}: ".format(subjectcol)	# コロンを連結する。
	return "{}{}{}{}".format(sharpcol, datecol, subjectcol, articlecol)  # #列、Date列、Subject列、記事列を結合。		
def createCopyFuncs(xscriptcontext, functionaccess):  # コピーのための関数を返す関数。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	stringlength = VARS.stringlength  # 1セルあたりの文字数。		
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
	transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))  # 全角文字を半角にする。
	def fullwidth_halfwidth(txt):	# 全角を半角に変換する得。
		if txt and isinstance(txt, str):  # 空文字でなくかつ文字列の時。日付のときはfloatが返ってくるので。
			txt = txt.strip()  # 前後の空白文字を削除する。
			return transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換。
		else:  # 空文字または文字列でないときはそのまま返す。
			return txt	
	def getCopyDataRows(copydatarows, datarows, startrow):  # コピー用シートにコピーする行のリストを取得。
		sheet = VARS.sheet
		rowindexes = []  # 空行の行インデックスのリスト。
		for i, datarow in enumerate(datarows):
			rowtxt = getRowTxt(functionaccess, datarow)  # #列、日付列、Problem列、記事列を結合した文字列を取得。
			if rowtxt:  # 空行でない時。
				copydatarows.append((rowtxt,))  # クリップボードに取得する行のリストに取得。
			else:  # 空行の時。
				rowindexes.append(startrow+i)  # 空行の行インデックスを取得。
		if rowindexes:  # 空行がある時。	
			cellranges = xscriptcontext.getDocument().createInstance("com.sun.star.sheet.SheetCellRanges") 
			cellranges.addRangeAddresses([sheet[i, 0].getRangeAddress() for i in rowindexes], True)  # 削除する行のセル範囲を結合してセル範囲コレクションに取得。
			rangeaddresses = list(cellranges.getRangeAddresses())  # 合成したセル範囲アドレスをリストで取得。
			rangeaddresses.sort(key=lambda x: x.StartRow, reverse=True)  # StartRowで降順にソートする。
			for i in rangeaddresses:
				sheet.removeRange(i, delete_rows)  # 後ろからでないとすべて削除できない。
		return len(rowindexes)  # 削除した空行数を返す。
	def formatArticleColumn(datarange):  # 記事列の文字列を制限して整形する。
		c = 0  # 合計追加行数。	
		datarangestartrow = datarange.getRangeAddress().StartRow  # datarangeの開始行インデックスを取得。
		datarows = [[fullwidth_halfwidth(j) for j in i] for i in datarange.getDataArray()]  # 取得した行のタプルを半角にする。floatが混じってくるので結合してから半角にする方法は使えない。
		datarange.setDataArray(datarows)  # datarangeに代入し直す。
		articlecells = [str(i[4]) for i in datarows]  # 記事列の行を文字列にして1次元リストで取得。
		newarticlerows = []  # 記事列代入するための行のリスト。
		cellranges = getCellRanges(xscriptcontext.getDocument(), datarange, datarows)  # #ごとのセル範囲コレクションを取得。
		sheet = VARS.sheet
		for cellrange in cellranges:  # #ごとのセル範囲について。
			rangeaddress = cellrange.getRangeAddress()  # cellrangeのセル範囲アドレスを取得。
			articletxts = articlecells[rangeaddress.StartRow-datarangestartrow-c:rangeaddress.EndRow-datarangestartrow+1-c]  # セル範囲の記事列のセルの文字列をリストで取得。
			articletxt = "".join(articletxts)  # cellrangeの記事列を結合。
			if articletxt:  # 文字列が取得出来た時。
				newdatarows = [(articletxt[i:i+stringlength],) for i in range(0, len(articletxt), stringlength)]  # 文字列をセルあたりの文字数で分割。
				diff = len(newdatarows) - len(articletxts)  # 追加行数を取得。
				if diff>0:  # 追加行がある時。
					endrowbelow = cellrange.getRangeAddress().EndRow + 1  # #ごとの終了行下の行インデックス。
					newrangeaddress = sheet[endrowbelow:endrowbelow+diff, :].getRangeAddress()  # 追加する行のセル範囲アドレス。
					sheet.insertCells(newrangeaddress, insert_rows)  # 空行をプロブレムごとに挿入。	
					sheet.queryIntersection(newrangeaddress).clearContents(511)  # 追加行の内容をクリア。セル範囲アドレスでは行がずれるので不可。
					c += diff  # 合計追加行数に追加。
				elif diff<0:  # 行数が減る時。
					newdatarows.extend([("",)]*-diff)  # 減った分の空行を追加。	
			else:
				newdatarows = [("",)]*len(articletxts)  # 文字列が取得できなかった時は空行を追加。
			newarticlerows.extend(newdatarows)	# 新しい記事列に行を追加。	
		if newarticlerows:  # 新しい行があるとき。空行だけのときはエラーになるので。
			newrange = sheet[datarangestartrow:datarangestartrow+len(newarticlerows), VARS.articlecolumn]  # 記事列のセル範囲を取得。
			newrange.clearContents(CellFlags.STRING+CellFlags.VALUE)  # 記事列の文字列と数値をクリア。
			newrange.setDataArray(newarticlerows)  # 記事列に代入。
		return c  # 追加した行数を返す。
	def formatProblemList(startrow, endrow, title):  # プロブレム欄を整形。
		c = formatArticleColumn(VARS.sheet[startrow:endrow, VARS.sharpcolumn:VARS.articlecolumn+1])  # プロブレム欄の記事列を整形。追加した行数が返る。
		datarows = VARS.sheet[startrow:endrow+c, VARS.sharpcolumn:VARS.articlecolumn+1].getDataArray()  # 文字数制限後の行のタプルを取得。
		newdatarows = [(title,)]  # タイトルを取得。	
		getCopyDataRows(newdatarows, datarows, startrow)  # プロブレム欄の記事列を整形。
		return newdatarows
	def copyCells(newdatarows):
		systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
		txt = "\r\n".join(i[0] for i in newdatarows)  # Windowsのために\rも付ける。
		systemclipboard.setContents(commons.TextTransferable(txt), None)  # クリップボードにコピーする。シートのコピーからだとペーストできないアプリがある。クリップボードが開けないと言われる。			
	return getCopyDataRows, formatArticleColumn, formatProblemList, copyCells, fullwidth_halfwidth
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	selection = eventobject.Source.getSelection()
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択オブジェクトがセル範囲であることを確認する。シート削除したときにエラーになるので。	
		VARS.setSheet(selection.getSpreadsheet())			
		drowBorders(xscriptcontext, selection)  # 枠線の作成。
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):		
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
	del contextmenu[:]  # contextmenu.clear()は不可。
	selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上隅のセルアドレスを取得。
	r, c = celladdress.Row, celladdress.Column
	if contextmenuname=="cell":  # セルのとき
		if r>VARS.splittedrow-1:  # 分割行以下の時。
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
				if c in (VARS.datecolumn, VARS.problemcolumn):  # 日付列またはプロブレム列の時。
					if c==VARS.datecolumn and selection.getValue()>0:  # 日付列、かつ、文字列でない、の時。文字列の時は0.0が返る。
						addMenuentry("ActionTrigger", {"Text": "年-月", "CommandURL": baseurl.format("entry8")}) 	
						addMenuentry("ActionTrigger", {"Text": "年", "CommandURL": baseurl.format("entry9")}) 		
						addMenuentry("ActionTrigger", {"Text": "年-月-日", "CommandURL": baseurl.format("entry10")}) 	
						addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。	
					if selection.getString()!="#":  # セルの文字列が#ではない時のみ。
						addMenuentry("ActionTrigger", {"Text": "#", "CommandURL": baseurl.format("entry11")})		
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。			
			commons.cutcopypasteMenuEntries(addMenuentry)
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
			addMenuentry("ActionTrigger", {"Text": "クリア", "CommandURL": baseurl.format("entry6")}) 
	elif contextmenuname=="rowheader" and len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 行ヘッダーのとき、かつ、選択範囲の列数がシートの列数が一致している時。	
		if r<VARS.splittedrow:  # 分割行より上の時。
			return EXECUTE_MODIFIED  # コンテクストメニューを表示しない。
		elif r<VARS.bluerow:  # 青行より上の時。
			addMenuentry("ActionTrigger", {"Text": "過去へ移動", "CommandURL": baseurl.format("entry2")})  
			addMenuentry("ActionTrigger", {"Text": "過去にｺﾋﾟｰ", "CommandURL": baseurl.format("entry3")}) 
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			addMenuentry("ActionTrigger", {"Text": "最下行へ", "CommandURL": baseurl.format("entry1")})   
		elif VARS.skybluerow<r<VARS.redrow:  # スカイブルー行より下、かつ、赤行より上の時。
			addMenuentry("ActionTrigger", {"Text": "現へ移動", "CommandURL": baseurl.format("entry4")})  
			addMenuentry("ActionTrigger", {"Text": "現にｺﾋﾟｰ", "CommandURL": baseurl.format("entry5")})  
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})	
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		commons.rowMenuEntries(addMenuentry)
	elif contextmenuname=="colheader":  # 列ヘッダーの時。
		pass
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
	if entrynum in (1, 2, 3):
		problemranges = getProblemRanges(doc, sheet[VARS.splittedrow:VARS.bluerow, VARS.sharpcolumn:VARS.articlecolumn+1], selection)  # 現問題ごとのセル範囲コレクションを取得。
		if entrynum==1:  # 現リストの最下行へ。青行の上に移動する。セクションC。
			dest_start_ridx = VARS.bluerow  # 移動先開始行インデックス。	
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。			
		elif entrynum==2:  # 過去ﾘｽﾄへ移動。スカイブルー行の下に移動する。セクションC。
			dest_start_ridx = VARS.skybluerow + 1  # 移動先開始行インデックス。
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。			
		elif entrynum==3:  # 過去ﾘｽﾄにｺﾋﾟｰ。スカイブルー行の下にコピーする。
			dest_start_ridx = VARS.skybluerow + 1  # 移動先開始行インデックス。
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.copyRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
	elif entrynum in (4, 5):
		problemranges = getProblemRanges(doc, sheet[VARS.skybluerow+1:VARS.redrow, VARS.sharpcolumn:VARS.articlecolumn+1], selection)  # 過去問題ごとのセル範囲コレクションを取得。
		if entrynum==4:  # 現ﾘｽﾄへ移動。青行の上に移動する。
			dest_start_ridx = VARS.bluerow  # 移動先開始行インデックス。	
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。			
		elif entrynum==5:  # 現ﾘｽﾄにｺﾋﾟｰ。青行の上にコピーする。
			dest_start_ridx = VARS.bluerow  # 移動先開始行インデックス。	
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.copyRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
	elif entrynum==6:  # クリア。
		rangeaddress = selection.getRangeAddress()  # 選択範囲のアドレスを取得。
		splittedrow = VARS.splittedrow
		edgerows = VARS.bluerow, VARS.skybluerow, VARS.redrow
		for i in range(rangeaddress.StartRow, rangeaddress.EndRow+1):  # 選択範囲の行インデックスをイテレート。
			for j in range(rangeaddress.StartColumn, rangeaddress.EndColumn+1):  # 選択範囲の列インデックスをイテレート。
				if i<splittedrow or i in edgerows:  
					continue
				else:  # それ以外の時。
					sheet[i, j].clearContents(511)  # 範囲をすべてクリアする。
	elif entrynum in (8, 9, 10):
		formatkey = ""
		if entrynum==8:  # 年-月、書式にする。
			formatkey = 'YYYY-M'
		elif entrynum==9:  # 月、書式にする。
			formatkey = 'YYYY'
		elif entrynum==10:  # 年-月-日、書式にする。
			formatkey = 'YYYY-M-D'
		if formatkey:
			selection.setPropertyValues(("NumberFormat", "HoriJustify", "VertJustify"), (commons.formatkeyCreator(doc)(formatkey), LEFT, CellVertJustify2.CENTER)) 
	elif entrynum==11:  # #を代入。
		selection.setString("#")
		selection.setPropertyValues(("HoriJustify", "VertJustify"), (RIGHT, CellVertJustify2.CENTER)) 
def moveProblems(sheet, problemrange, dest_start_ridx):  # problemrange; 問題リストの塊。dest_start_ridx: 移動先開始行インデックス。
	dest_endbelow_ridx = dest_start_ridx + len(problemrange.getRows())  # 移動先最終行の次の行インデックス。
	dest_rangeaddress = sheet[dest_start_ridx:dest_endbelow_ridx, :].getRangeAddress()  # 挿入前にセル範囲アドレスを取得しておく。
	sheet.insertCells(dest_rangeaddress, insert_rows)  # 空行を挿入。	
	sheet.queryIntersection(dest_rangeaddress).clearContents(511)  # 挿入した行の内容をすべてを削除。挿入セルは挿入した行の上のプロパティを引き継いでいるのでリセットしないといけない。
	cursor = sheet.createCursorByRange(problemrange)  # セルカーサーを作成		
	cursor.expandToEntireRows()  # セル範囲を行全体に拡大。		
	return cursor.getRangeAddress()  # 移動前に移動元の問題リストのセル範囲アドレスを取得しておく。
def drowBorders(xscriptcontext, selection):  # selectionを交点とする行列全体の外枠線を描く。
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上端のセルアドレスを取得。
	r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。	
	sheet = VARS.sheet
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = commons.createBorders()
	sheet[:, :].setPropertyValue("TopBorder2", noneline) # 1辺をNONEにするだけですべての枠線が消える。
	if r<VARS.splittedrow or r>VARS.redrow:  # 分割行より上、または、赤行より下の時。
		return  # 枠線を消すだけ。
	elif r<VARS.bluerow or VARS.skybluerow-1<r<VARS.redrow:  # 青行より上、または、スカイブルー行と赤行の間の時。
		if c<VARS.articlecolumn+1:  # 記事列とその左の列の時。
			datarange = sheet[VARS.splittedrow:VARS.bluerow, VARS.sharpcolumn:VARS.articlecolumn+1] if r<VARS.bluerow else sheet[VARS.skybluerow+1:VARS.redrow, VARS.sharpcolumn:VARS.articlecolumn+1]  # タイトル行を除く。
			doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
			problemranges = getProblemRanges(doc, datarange, selection)  # 問題ごとのセル範囲コレクションを取得。
			for i in problemranges:
				cursor = sheet.createCursorByRange(i)  # セルカーサーを作成		
				cursor.expandToEntireRows()  # セル範囲を行全体に拡大。	
				cursor.setPropertyValue("TableBorder2", topbottomtableborder) # プロブレムの上下に枠線を引く。
		elif c<VARS.splittedcolumn:  # 分割列より左列の時。
			rangeaddress = selection.getRangeAddress() # 選択範囲のセル範囲アドレスを取得。
			sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder) # 行の上下に枠線を引く
			sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。
			selection.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	
def getProblemRanges(doc, datarange, selection):  # datarangeは問題リストの#を検索するセル範囲。
	cellranges = getCellRanges(doc, datarange)  # 問題リストの各セル範囲を取得。
	ranges = set(i for i in cellranges if len(i.queryIntersection(selection.getRangeAddress())))  # 選択したセル範囲と交差するプロブレムのセル範囲を集合で取得。
	problemranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
	problemranges.addRangeAddresses([i.getRangeAddress() for i in ranges], True)  # セル範囲コレクションにプロブレムのセル範囲を追加する。セル範囲を結合する。rangesの要素がなくてもエラーにならない。		
	return problemranges  # 問題ごとのセル範囲コレクションを返す。		
def getCellRanges(doc, datarange, datarows=None):  # 各プロブレムの行をまとめたセル範囲コレクションを返す。列はdatarangeと同じ。
	cellranges = []
	if datarows is None:
		datarows = datarange.getDataArray()  # #列からSubject列までの行のタプルを取得。
	ranges = []  # プロブレムリストのセル範囲のリスト。
	rstartrow = 0  # プロブレム開始行の相対インデックス。
	for i, datarow in enumerate(datarows):  # 相対インデックスと行のタプルを列挙。
		if "#" in "{}{}{}".format(datarow[0], datarow[1], datarow[3]):  # #列、Date列、Subject列を結合して#がある時。日付は数値なので文字列への変換が必要なのでjoin()は使えない。
			if i>rstartrow:  # 開始行相対インデックスより大きい時。
				ranges.append(datarange[rstartrow:i, :])  # プロブレムリストの開始行から終了行までのセル範囲を取得。
				rstartrow = i
	else:
		ranges.append(datarange[rstartrow:, :])  # 最後のプロブレムのセル範囲を追加。
	if ranges:
		cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
		cellranges.addRangeAddresses([i.getRangeAddress() for i in ranges], False)  # セル範囲コレクションにプロブレムのセル範囲を追加する。セル範囲は結合しない。
	return cellranges  # 列はdatarangeと同じ。
