#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# カルテシートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from datetime import date, datetime, timedelta
from itertools import chain
from indoc import commons, datedialog, historydialog, staticdialog
from com.sun.star.awt import MouseButton  # MessageBoxButtons, MessageBoxResults # 定数
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.sheet.CellDeleteMode import ROWS as delete_rows  # enum
from com.sun.star.sheet.CellInsertMode import ROWS as insert_rows  # enum
from com.sun.star.table import CellVertJustify2  # 定数
from com.sun.star.table.CellHoriJustify import LEFT, CENTER  # enum
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
		self.dateformat = "%Y/%m/%d %H:%M:%S Copied"  # 記事をコピーした日時の書式。
	def setSheet(self, sheet): # 逐次変化する値。
		self.sheet = sheet	
		cellranges = sheet[self.splittedrow:, self.datecolumn].queryContentCells(CellFlags.STRING)  # Date列の文字列が入っているセルに限定して抽出。
		backcolors = commons.COLORS["blue3"], commons.COLORS["skyblue"], commons.COLORS["red3"]  # ジェネレーターに使うので順番が重要。
		gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in backcolors)
		self.bluerow = next(gene)  # 青3行インデックス。
		self.skybluerow = next(gene)  # スカイブルー行インデックス。
		self.redrow = next(gene)  # 赤3行インデックス。
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
	if copieddatetxt:
		copieddatetime = datetime.strptime(copieddatetxt, VARS.dateformat)  # コピーした日時を取得。
		now = datetime.now()  # 現在の日時を取得。
		if copieddatetime.date()<now.date():  # 今日はまだコピーしていない時。
			copieddatecell.setPropertyValues(("CharColor", "CellBackColor"), (-1, commons.COLORS["magenta3"]))  # 文字色をリセットして背景色をマゼンダにする。
		elif now.hour>12 and copieddatetime.hour<12:  # 今日はコピーしていても、午後になって午前にしかコピーしていない時。
			copieddatecell.setPropertyValue("CharColor", commons.COLORS["magenta3"])  # 文字色をマゼンダにする。背景色はコピーした時にすでにライムになっているはず。
	# 本日の記事を過去の記事に移動させる。
	dateformat = "****%Y年%m月%d日(%a)****"
	daterange = sheet[VARS.bluerow, VARS.articlecolumn]  # 本日の記事の日付セルを取得。
	articledatetxt = daterange.getString()  # 本日の記事の日付セルの文字列を取得。
	articledate = datetime.strptime(articledatetxt, dateformat)  # 記事列の日付を取得。
	todaydate = date.today()  # 今日のdateオブジェクトを取得。
	if articledate!=todaydate:  # 今日の日付でない時。
		todayarticle = sheet[VARS.bluerow+1:VARS.skybluerow, :]  # 青行とスカイブルー行の間の行のセル範囲。
		datarows = todayarticle[:, VARS.sharpcolumn:VARS.articlecolumn+1].getDataArray()  # 本日の記事欄のセルをすべて取得。
		txt = "".join(map(str, chain.from_iterable(datarows)))  # 本日の記事欄を文字列にしてすべて結合。
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
		daterange.setString(todaydate.strftime(dateformat))  # 今日の日付を本日の記事欄に入力。
		cellranges.addRangeAddresses([todayarticle[:, i].getRangeAddress() for i in (VARS.datecolumn, VARS.subjectcolumn, VARS.articlecolumn)], False)  # 本日の記事のDate列、Subject列、記事列のセル範囲コレクションを取得。
		cellranges.setPropertyValue("IsTextWrapped", True)  # セルの内容を折り返す。	
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			VARS.setSheet(selection.getSpreadsheet())
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				drowBorders(xscriptcontext, selection)  # 枠線の作成。
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				r = selection.getCellAddress().Row  # 選択セルの行インデックスを取得。	
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
	if txt=="一覧へ":
		controller.setActiveSheet(sheets["一覧"])  # 一覧シートをアクティブにする。
	elif txt=="経過へ":
		newsheetname = "".join([sheet.getName(), "経"])  # 経過シート名を取得。
		if newsheetname in sheets:  # 経過シート名がある時。
			controller.setActiveSheet(sheets[newsheetname])  # 経過シートをアクティブにする。
		else:  # 経過シートの作成。
			idcelltxts = sheet[VARS.splittedrow-1, VARS.articlecolumn].getString().split(" ")  # 半角スペースで分割。
			idtxt = idcelltxts[0]  # 最初の要素を取得。
			if idtxt.isdigit():  # IDが数値のみの時。					
				if idtxt in sheets:  # ID名のシートがあるとき。
					controller.setActiveSheet(sheets[idtxt])  # カルテシートをアクティブにする。
				else:
					if len(idcelltxts)==5:  # ID、漢字姓・名、カタカナ姓・名、の5つに分割できていた時。
						kanjitxt, kanatxt = " ".join(idcelltxts[1:3]), " ".join(idcelltxts[3:])
						datevalue = sheet[VARS.splittedrow, VARS.datecolumn].getValue()
						keikasheet =  commons.getKeikaSheet(doc, idtxt, kanjitxt, kanatxt, datevalue)  # 経過シートを取得。
						controller.setActiveSheet(keikasheet)  # 経過シートをアクティブにする。
					else:
						commons.showErrorMessageBox(controller, "「ID(数値のみ) 漢字姓 名 カナ姓 名」の形式になっていません。")
			else:
				commons.showErrorMessageBox(controller, "IDが取得できませんでした。")		
	elif txt=="COPY":
		getCopyDataRows, formatArticleColumn, formatProblemList, copyCells = createCopyFuncs(xscriptcontext)
		c = formatArticleColumn(sheet[VARS.bluerow+1:VARS.skybluerow, VARS.sharpcolumn:VARS.articlecolumn+1])  # 本日の記事欄の記事列を整形。追加した行数が返る。
		datarows = sheet[VARS.bluerow:VARS.skybluerow+c, VARS.sharpcolumn:VARS.articlecolumn+1].getDataArray()  # 文字数制限後の行のタプルを取得。
		copydatarows = [(datarows[0][5],)]  # 本日の記事の日付を取得。
		deletedrowcount = getCopyDataRows(copydatarows, datarows[1:], VARS.bluerow+1)  # 削除された行数。
		if deletedrowcount>0:  # 削除した行があるとき。
			startrow = VARS.skybluerow - deletedrowcount
			newrangeaddress = sheet[startrow:startrow+deletedrowcount, :].getRangeAddress()  # 挿入するセル範囲アドレスを取得。
			sheet.insertCells(newrangeaddress, insert_rows)  # 空行を挿入。	
			sheet.queryIntersection(newrangeaddress).clearContents(511)  # 追加行の内容をクリア。セル範囲アドレスから取得しないと行挿入後のセル範囲が異なってしまう。
		newdatarows = formatProblemList(VARS.splittedrow, VARS.bluerow, "****ｻﾏﾘ****")  # プロブレム欄を整形。
		for i in (VARS.problemcolumn, VARS.articlecolumn):  # Subject列と記事列について。
			newrange = sheet[VARS.splittedrow:, i]
			newrange.setPropertyValue("IsTextWrapped", True)  # セルの内容を折り返す。
			newrange.getRows().setPropertyValue("OptimalHeight", True)  # 内容を折り返した後の行の高さを調整。
		cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges") 
		cellranges.addRangeAddresses([i.getRangeAddress() for i in (sheet[VARS.splittedrow:VARS.bluerow, VARS.sharpcolumn:VARS.problemcolumn+1], sheet[VARS.bluerow+1:VARS.skybluerow, VARS.sharpcolumn:VARS.problemcolumn+1])], False)  # プロブレム欄、本日の記事欄をセル範囲を取得。
		cellranges.setPropertyValue("VertJustify", CellVertJustify2.CENTER)  # 縦位置を中央にする。
		newdatarows.extend(copydatarows)  # 本日の記事欄をプロブレム欄の下に追加。
		copieddatecell = sheet[0, VARS.articlecolumn]  # コピー日時セルを取得。	
		copyCells(controller, copieddatecell, newdatarows)
		copieddatecell.setString(datetime.now().strftime(VARS.dateformat))  # コピーボタンを押した日付を入力。
		copieddatecell.setPropertyValues(("CellBackColor", "CharColor"), (commons.COLORS["lime"], -1))  # コピー日時セルの背景色を変更。文字色をリセット。
	elif txt=="退院ｻﾏﾘ":
		dummy, dummy, formatProblemList, copyCells = createCopyFuncs(xscriptcontext)
		newdatarows = formatProblemList(VARS.splittedrow, VARS.bluerow, "****退院ｻﾏﾘ****")  # プロブレム欄を整形。
		copieddatecell = sheet[0, VARS.articlecolumn]  # コピー日時セルを取得。	
		copyCells(controller, copieddatecell, newdatarows)
		selection.setPropertyValue("CellBackColor", commons.COLORS["lime"])  # 退院ｻﾏﾘボタンの背景色を変更。
	elif txt=="#分離":
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
		functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。
		datarange = sheet[VARS.splittedrow:VARS.bluerow, :VARS.articlecolumn+1]
		datarows = datarange.getDataArray()
		newdatarows = []
		for datarow in datarows:
			datatxt = datarow[VARS.articlecolumn]  # Article列の文字列を取得。
			if datatxt.startswith("#"):  # #がある時。
				datecell, subjectcell = "", ""
				if len(datatxt)>1:  # #以外の文字もある時。
					datatxt = datatxt[1:]  # #を除く。
				if ":" in datatxt:  # コロンがある時。
					ds, datatxt = datatxt.split(":", 1)  # 最初のコロンで分割。
					datecell, subjectcell = handleDS(functionaccess, ds, datecell, subjectcell)
				datarow = "", "#", datecell, "", subjectcell, "", datatxt
			newdatarows.append(datarow)
		datarange.setDataArray(newdatarows)
		sheet[VARS.splittedrow:VARS.bluerow, VARS.sharpcolumn].setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.CENTER))  # #列の書式設定。左寄せにする。
		createFormatKey = commons.formatkeyCreator(doc)		
		sheet[VARS.splittedrow:VARS.bluerow, VARS.datecolumn].setPropertyValues(("NumberFormat", "HoriJustify", "VertJustify"), (createFormatKey('YYYY/MM/DD'), LEFT, CellVertJustify2.CENTER))  # カルテシートの入院日の書式設定。左寄せにする。
		sheet[VARS.splittedrow:VARS.bluerow, VARS.problemcolumn].setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.CENTER))  # Subject列の書式設定。左寄せにする。
	elif txt=="問題ﾘｽﾄへ変換":
		cellranges = sheet[VARS.redrow+1:, VARS.articlecolumn].queryContentCells(CellFlags.STRING)  # Article列の文字列が入っているセルに限定して抽出。
		if len(cellranges):  # セル範囲が取得出来た時。
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
			functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。		
			transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
			transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))  # 全角文字を半角にする。
			newdatarows = [] 
			emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 最終行インデックス+1を取得。
			datarange = sheet[VARS.redrow+1:emptyrow, VARS.articlecolumn]
			datarows = datarange.getDataArray()
			stringlength = VARS.stringlength  # 1セルあたりの文字数。
			sharpcell, datecell, subjectcell, articletxts = "", "", "", []
			for datatxt in map(str, chain.from_iterable(datarows)):
				if datatxt:  # 空文字でない時。
					datatxt = transliteration.transliterate(datatxt, 0, len(datatxt), [])[0]  # 半角に変換。
					if datatxt.startswith("#"):  # #がある時。
						if articletxts:  # すでに取得したArticle列の行がある時。
							addDataRow(stringlength, sharpcell, datecell, subjectcell, articletxts, newdatarows)  # 新しいデータ行に追加する。	
							sharpcell, datecell, subjectcell, articletxts = "", "", "", []	# 変数をリセットする。					
						sharpcell = "#"	# #を取得。
						if ":" in datatxt:  # コロンがある時。
							ds, articletxt = datatxt[1:].split(":", 1)  # 最初のコロンで1回分割。
							articletxt and articletxts.append(articletxt)  # コロンの後ろがある時articletxtsに追加。
							datecell, subjectcell = handleDS(functionaccess, ds, datecell, subjectcell)
						else:  # コロンがない時。		
							articletxts.append(datatxt[1:])  # #を除いてArticle列の文字列のリストに取得。	
					else:  # #がない時。
						if not datatxt.startswith("****"):  # ****から始まっていない時。
							articletxts.append(datatxt)  # Article列の文字列のリストに追加。
			if articletxts:  # すでに取得したArticle列の行がある時。
				addDataRow(stringlength, sharpcell, datecell, subjectcell, articletxts, newdatarows)  # 最後のプロブレムを処理。
			problemrange = sheet[VARS.splittedrow:VARS.bluerow, VARS.sharpcolumn:VARS.articlecolumn+1]
			cellranges = problemrange.queryContentCells(CellFlags.STRING)
			emptyrow = max(i.EndRow for i in cellranges.getRangeAddresses()) + 1 if len(cellranges) else VARS.splittedrow
			endrowbelow = emptyrow + len(newdatarows)	
			sheet.insertCells(sheet[emptyrow:endrowbelow, :].getRangeAddress(), insert_rows)  # ダブルクリックした行の下に空行を挿入。	
			sheet[emptyrow:endrowbelow, :].setPropertyValues(("CellBackColor", "CharColor"), (-1, -1))  # 追加行の背景色と文字色をクリア。	
			sheet[emptyrow:endrowbelow, VARS.sharpcolumn:VARS.sharpcolumn+len(newdatarows[0])].setDataArray(newdatarows)
			createFormatKey = commons.formatkeyCreator(doc)		
			sheet[emptyrow:endrowbelow, VARS.datecolumn].setPropertyValues(("NumberFormat", "HoriJustify", "VertJustify"), (createFormatKey('YYYY/MM/DD'), LEFT, CellVertJustify2.CENTER))  # カルテシートの入院日の書式設定。左寄せにする。
			datarange.clearContents(CellFlags.STRING)  # コピー元の文字列をクリア。	
	elif txt[:8].isdigit():  # 最初8文字が数値の時。
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
		systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
		systemclipboard.setContents(commons.TextTransferable(txt[:8]), None)  # クリップボードにIDをコピーする。							
	return False  # セルを編集モードにしない。
def wClickCol(enhancedmouseevent, xscriptcontext):  # 列によって変える処理。
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	celladdress = selection.getCellAddress()
	r, c = celladdress.Row, celladdress.Column  # ダブルクリックしたセルの行インデックス、列インデックスを取得。
	if c==0:  # 行挿入列の時。
		sheet = VARS.sheet
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
		datedialog.createDialog(enhancedmouseevent, xscriptcontext, "日付入力", "YYYY/MM/DD")	
	elif c in (VARS.problemcolumn, VARS.articlecolumn):  # プロブレム列または記事列の時。
		return True  # セル編集モードにする。
	elif c==VARS.phrasecolumn:  # 定型句列インデックスの時。
		staticdialog.createDialog(enhancedmouseevent, xscriptcontext, "ﾌﾟﾛﾌﾞﾚﾑ", outputcolumn=VARS.problemcolumn, callback=callback_phrasecolumn)
		selection.setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.CENTER))
	elif c==VARS.insertdatecolumn:  # 日付挿入列の時。
		selection.setString("")  # 日付挿入列の文字列をクリア。
		datedialog.createDialog(enhancedmouseevent, xscriptcontext, "日付挿入", "YYYY/M/D", callback=callback_insertdatecolumn)  # ダイアログの戻り値は取得できず、入力も待たず次のコードにいってしまう。
		selection.setPropertyValue("CharColor", commons.COLORS["white"])  # 日付挿入列の文字色を白色にする。
	elif c==VARS.replacedatecolumn:  # 日付入替列の時。
		sheet = VARS.sheet
		datetxt = sheet[r, VARS.insertdatecolumn].getString()  # 日付挿入列の文字列を取得。
		if datetxt:  # 日付文字列が取得出来た時。
			articlecell = sheet[r, VARS.articlecolumn]  # 記事セルを取得。
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
					if len(txts)>1:  # "｡"がない時は何もしない。
						if txts[-1]:  # 日付の直前が｡でない時。
							articletxt = "".join((txts[0], "｡", datetxt, " ", txts[1]))  # ｡の後ろに日付を移動させる。
						else:  # 日付の直前が｡の時。txts[-1]は空文字になる。
							txts2 = txts[0].rsplit("｡", 1)  # 右から｡で再分割。	
							if len(txts2)>1:  # ｡の後ろに日付を移動させる。
								articletxt = "".join((txts2[0], "｡", datetxt, " ", txts2[1], "｡"))
						articlecell.setString(articletxt)
	elif c==VARS.historycolumn:  # 履歴列の時。
		
		
		
		pass
				
				
	return False  # セルを編集モードにしない。	
def callback_phrasecolumn(mouseevent, xscriptcontext):  # プロブレム列に、#2018/5/7 心エコー:LV wall function normal、とあるのを処理する。
	sheet = VARS.sheet
	selection = xscriptcontext.getDocument().getCurrentSelection()  # シート上で選択しているオブジェクトを取得。	
	datarow = sheet[selection.getCellAddress().Row, :VARS.articlecolumn+1].getDataArray()[0]
	problemtxt = datarow[VARS.problemcolumn]
	if problemtxt.startswith("#"):
		datarow[VARS.sharpcolumn] = "#"
		problemtxt = problemtxt.replace("#", "")
	
	
	

def callback_insertdatecolumn(mouseevent, xscriptcontext):  # 日付挿入列をダブルクリックした時に日付入力ダイアログに渡すコールバック関数。
	selection = xscriptcontext.getDocument().getCurrentSelection()  # シート上で選択しているオブジェクトを取得。	
	articlecell = VARS.sheet[selection.getCellAddress().Row, VARS.articlecolumn]  # 記事セルを取得。		
	articlecell.setString("".join([articlecell.getString(), selection.getString()]))  # 新規日付を代入。
def handleDS(functionaccess, ds, datecell, subjectcell):
	if " " in ds:  # スペースがある時。
		datetxt, subjectcell = ds.split(" ", 1)  # 最初のスペースで1回分割。
	else:  # スペースがない時とりあえず日付として処理する。CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME+CellFlags.FORMULA
		datetxt = ds
	if len(datetxt)>4 and datetxt[:4].isdigit():  # 最初の4文字がすべて数値の時。年月から始まっていると判断する。
		datecell = int(functionaccess.callFunction("DATEVALUE", (datetxt.replace(datetxt[4], "/"),)))  # シリアル値を整数で取得。floatで返る。シリアル値で入れないとsetDataArray()で日付にできない。
	else:
		subjectcell = ds  # スペースで分割した時の最初の要素が年月でない時はすべてSubject。
	return datecell, subjectcell
def addDataRow(stringlength, sharpcell, datecell, subjectcell, articletxts, newdatarows):
	articletxt = "".join(articletxts).lstrip().replace("\n", "")  # 先頭の空白とセル内の改行文字も除去する。
	articlecells = [articletxt[i:i+stringlength] for i in range(0, len(articletxt), stringlength)]  # 文字列を制限したArticle列のジェネレーター。
	datarow = sharpcell, datecell, "", subjectcell.strip(), "", articlecells[0]  # プロブレムの1行目を取得。
	newdatarows.append(datarow)  # プロブレムの1行目を追加。
	if len(articlecells)>1:  # 複数行ある時。
		newdatarows.extend(("", "", "", "", "", i) for i in articlecells[1:])	 # 2行目以降について処理。
def createCopyFuncs(xscriptcontext):  # コピーのための関数を返す関数。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	stringlength = VARS.stringlength  # 1セルあたりの文字数。
	functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。			
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
	transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))  # 全角文字を半角にする。
	def _getRowTxt(datarow):  # 1行の列を連結して文字列を返す。
		sharpcol, datecol, subjectcol, articlecol = datarow[0], datarow[1], datarow[3], datarow[5]  # #列、日付列、Subject列、記事列を取得。
		if datecol and isinstance(datecol, float):  # 日付列がfloat型のとき。
			datecol = "{} ".format("/".join([str(int(functionaccess.callFunction(i, (datecol,)))) for i in ("YEAR", "MONTH", "DAY")]))  # シリアル値をシート関数で年/月/日の文字列にする。引数のdatecolはfloatのままでよい。
		if subjectcol and subjectcol!="#":   # Subject列、かつ、#でないとき
			subjectcol = "{}: ".format(subjectcol)	# コロンを連結する。
		return "{}{}{}{}".format(sharpcol, datecol, subjectcol, articlecol)  # #列、Date列、Subject列、記事列を結合。
	def _fullwidth_halfwidth(txt):	# 全角を半角に変換する得。
		if txt and isinstance(txt, str):  # 空文字でなくかつ文字列の時。
			return transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換。
		else:  # 空文字または文字列でないときはそのまま返す。
			return txt	
	def getCopyDataRows(copydatarows, datarows, startrow):  # コピー用シートにコピーする行のリストを取得。
		sheet = VARS.sheet
		deletedrowcount = 0  # 空行の数。
		ranges = []  # 空行のセル範囲のリスト。
		for i, datarow in enumerate(datarows):
			rowtxt = _getRowTxt(datarow)  # #列、日付列、Subject列、記事列を結合した文字列を取得。
			if rowtxt:  # 空行でない時。
				copydatarows.append((rowtxt,))  # ｺﾋﾟｰ用シートにコピーする行のリストに取得。
			else:  # 空行の時。
				ranges.append(sheet[startrow+i, :])  # 削除行のセル範囲を取得。アドレスで取得するときは下から削除する必要がある。
				deletedrowcount += 1	
		cellranges = xscriptcontext.getDocument().createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
		cellranges.addRangeAddresses([i.getRangeAddress() for i in ranges], True)  # セル範囲を結合する。
		[sheet.removeRange(i.getRangeAddress(), delete_rows) for i in cellranges]  # cellrangesにあるセル範囲の行を削除する。 		
		return deletedrowcount  # 削除した空行数を返す。
	def formatArticleColumn(datarange):  # 記事列の文字列を制限して整形する。
		c = 0  # 合計追加行数。	
		datarangestartrow = datarange.getRangeAddress().StartRow  # datarangeの開始行インデックスを取得。
		datarows = datarange.getDataArray()  # datarangeの行のタプルを取得。
		datarows = [[_fullwidth_halfwidth(j) for j in i] for i in datarows]  # 取得した行のタプルを半角にする。		
		datarange.setDataArray(datarows)  # datarangeに代入し直す。
		articlecells = [str(i[5]) for i in datarows]  # 記事列の行を文字列にして1次元リストで取得。
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
				diff = 0 if diff<0 else diff  # 負の数の時は0にする。
				c += diff  # 合計追加行数に追加。
				endrowbelow = cellrange.getRangeAddress().EndRow + 1  # #ごとの終了行下の行インデックス。
				if diff>0:  # 追加行がある時。
					newrangeaddress = sheet[endrowbelow:endrowbelow+diff, :].getRangeAddress()  # 追加する行のセル範囲アドレス。
					sheet.insertCells(newrangeaddress, insert_rows)  # 空行をプロブレムごとに挿入。	
					sheet.queryIntersection(newrangeaddress).clearContents(511)  # セル範囲アドレスでは行がずれるので不可。
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
	def copyCells(controller, copieddatecell, newdatarows):
		sheetname = "ｺﾋﾟｰ用"
		doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
		sheets = doc.getSheets()  # シートコレクションを取得。	
		if not sheetname in sheets:  # コピー用シートがない時。
			sheets.insertNewByName(sheetname, len(sheets))  # コピー用シートを挿入。
		copysheet = xscriptcontext.getDocument().getSheets()[sheetname]  # コピー用シートを取得。
		copysheet.clearContents(511)  # シート内容をクリア。
		pasterange = copysheet[:len(newdatarows), :len(newdatarows[0])]  # コピー用シートのペーストするセル範囲を取得。。
		pasterange.setDataArray(newdatarows)  # コピー用シートにペーストする。
		pasterange.getColumns().setPropertyValue("Width", copieddatecell.getColumns().getPropertyValue("Width"))  # 単位は1/100mm
		pasterange.setPropertyValue("IsTextWrapped", True)  # ペーストしたセルの内容を折り返す。	
		pasterange.getRows().setPropertyValue("OptimalHeight", True)  # ペーストした行の高さを調整。
		dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
		controller.select(pasterange)  # ペーストしたセル範囲を取得。シートが切り替わってしまう。
		dispatcher.executeDispatch(controller.getFrame(), ".uno:Copy", "", 0, ())  # ペーストしたセルをコピー。
		controller.setActiveSheet(VARS.sheet)  # カルテシートに戻る.
	return getCopyDataRows, formatArticleColumn, formatProblemList, copyCells
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	controller = eventobject.Source
	selection = controller.getSelection()
	sheet = controller.getActiveSheet()
	VARS.setSheet(sheet)
	drowBorders(xscriptcontext, selection)  # 枠線の作成。
# def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):		
# 	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
# 	sheet = controller.getActiveSheet()  # アクティブシートを取得。
# 	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
# 	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
# 	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
# 	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
# 	del contextmenu[:]  # contextmenu.clear()は不可。
# 	selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
# 	if contextmenuname=="cell":  # セルのとき
# 		consts = getConsts(sheet, selection)  # セル固有の定数を取得。
# 		sectionname = consts.sectionname  # クリックしたセルの区画名を取得。			
# 		if sectionname in ("A", "B"):  # 固定行より上の時はコンテクストメニューを表示しない。
# 			return EXECUTE_MODIFIED
# 		commons.cutcopypasteMenuEntries(addMenuentry)
# 		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
# 		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Delete"})	
# # 		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
# # 			addMenuentry("ActionTrigger", {"Text": "To Green", "CommandURL": baseurl.format("entry1")}) 
# # 		elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 連続した複数セルの時。
# # 			addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format("entry2")}) 
# 	elif contextmenuname=="rowheader" and len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 行ヘッダーのとき、かつ、選択範囲の列数がシートの列数が一致している時。	
# 		consts = getConsts(sheet, selection)  # 選択範囲の最初のセルの定数を取得。
# 		sectionname = consts.sectionname  # クリックしたセルの区画名を取得。			
# 		if sectionname in ("A",) or selection[0, 0].getPropertyValue("CellBackColor")!=-1:  # 背景色のあるときは表示しない。
# 			return EXECUTE_MODIFIED
# 		if sectionname in ("C",):
# 			addMenuentry("ActionTrigger", {"Text": "過去ﾘｽﾄへ移動", "CommandURL": baseurl.format("entry2")})  
# 			addMenuentry("ActionTrigger", {"Text": "過去ﾘｽﾄにｺﾋﾟｰ", "CommandURL": baseurl.format("entry3")}) 
# 			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 			addMenuentry("ActionTrigger", {"Text": "最下行へ", "CommandURL": baseurl.format("entry1")})   
# 		elif sectionname in ("G",):
# 			addMenuentry("ActionTrigger", {"Text": "現ﾘｽﾄへ移動", "CommandURL": baseurl.format("entry4")})  
# 			addMenuentry("ActionTrigger", {"Text": "現ﾘｽﾄにｺﾋﾟｰ", "CommandURL": baseurl.format("entry5")})  
# 		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})	
# 		commons.cutcopypasteMenuEntries(addMenuentry)
# 		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 		commons.rowMenuEntries(addMenuentry)
# 	elif contextmenuname=="colheader":  # 列ヘッダーの時。
# 		pass
# 	elif contextmenuname=="sheettab":  # シートタブの時。
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
# 	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
# def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。	
# 	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
# 	controller = doc.getCurrentController()  # コントローラの取得。
# 	sheet = controller.getActiveSheet()  # アクティブシートを取得。
# 	selection = controller.getSelection()
# 	if len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 列全体が選択されている場合もあるので行全体が選択されていることを確認する。
# 		consts = getConsts(sheet, selection)
# 		splittedrow = consts.splittedrow
# 		bluerow = consts.bluerow
# 		skybluerow = consts.skybluerow
# 		redrow = consts.redrow
# 		sharpcolumn = consts.sharpcolumn
# 		articlecolumn = consts.articlecolumn
# 		if entrynum==1:  # 現リストの最下行へ。青行の上に移動する。セクションC。
# 			dest_start_ridx = bluerow  # 移動先開始行インデックス。	
# 			problemranges = getProblemRanges(doc, sheet[splittedrow:bluerow, sharpcolumn:articlecolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。
# 			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
# 				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
# 				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
# 				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。			
# 		elif entrynum==2:  # 過去ﾘｽﾄへ移動。スカイブルー行の下に移動する。セクションC。
# 			dest_start_ridx = skybluerow + 1  # 移動先開始行インデックス。
# 			problemranges = getProblemRanges(doc, sheet[splittedrow:bluerow, sharpcolumn:articlecolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。		
# 			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
# 				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
# 				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
# 				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。					
# 		elif entrynum==3:  # 過去ﾘｽﾄにｺﾋﾟｰ。スカイブルー行の下にコピーする。
# 			dest_start_ridx = skybluerow + 1  # 移動先開始行インデックス。
# 			problemranges = getProblemRanges(doc, sheet[splittedrow:bluerow, sharpcolumn:articlecolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。		
# 			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
# 				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
# 				sheet.copyRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
# 		elif entrynum==4:  # 現ﾘｽﾄへ移動。青行の上に移動する。
# 			dest_start_ridx = bluerow  # 移動先開始行インデックス。	
# 			problemranges = getProblemRanges(doc, sheet[skybluerow+1:redrow, sharpcolumn:articlecolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。
# 			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
# 				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
# 				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
# 				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。			
# 		elif entrynum==5:  # 現ﾘｽﾄにｺﾋﾟｰ。青行の上にコピーする。
# 			dest_start_ridx = bluerow  # 移動先開始行インデックス。	
# 			problemranges = getProblemRanges(doc, sheet[skybluerow+1:redrow, sharpcolumn:articlecolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。
# 			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
# 				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
# 				sheet.copyRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
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
