#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# カルテシートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from datetime import date, datetime
from itertools import chain
from indoc import commons
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.awt import MouseButton  # MessageBoxButtons, MessageBoxResults # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.sheet.CellInsertMode import ROWS as insert_rows  # enum
from com.sun.star.sheet.CellDeleteMode import ROWS as delete_rows  # enum
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.table import CellVertJustify2  # 定数
class Karute():  # シート固有の定数設定。
	def __init__(self):
		self.sharpcolumn = 1  # #列インデックス。
		self.datecolumn = 2  # Date列インデックス。
		self.kijicolumn = 6  # Article列インデックス。
		self.stringlength = 125  # 1セルあたりの文字数。
def getSectionName(controller, sheet, target):  # 区画名を取得。
	"""
	A  ||  B
	===========  # 行の固定の境界。||は列の固定の境界。境界の行と列はそれぞれ下、右に含む。
	C  ||  D
	-----------  # Date列の文字列があるセルの背景色が青3の行。
	E  ||  F
	-----------  # Date列の文字列があるセルの背景色がスカイブルーの行。
	G  ||  H
	-----------  # Date列の文字列があるセルの背景色が赤3の行。
	I  ||  J
	"""
	karute = Karute()  # クラスをインスタンス化。	
	subcontollerrange = controller[0].getVisibleRange()
	splittedrow = subcontollerrange.EndRow + 1  # スクロールする枠の最初の行インデックス。
	startcolumn = subcontollerrange.EndColumn + 1  # スクロールする枠の最初の列インデックス。
	cellranges = sheet[splittedrow:, karute.datecolumn].queryContentCells(CellFlags.STRING)  # Date列の文字列が入っているセルに限定して抽出。
	backcolors = commons.COLORS["blue3"], commons.COLORS["skyblue"], commons.COLORS["red3"]  # ジェネレーターに使うので順番が重要。
	gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in backcolors)
	bluerow = next(gene)
	skybluerow = next(gene)
	redrow = next(gene)
	rangeaddress = target.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	if len(sheet[:splittedrow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "A"
	elif len(sheet[:splittedrow, startcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "B"
	elif len(sheet[:bluerow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "C"
	elif len(sheet[:bluerow, startcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "D"
	elif len(sheet[:skybluerow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "E"
	elif len(sheet[:skybluerow, startcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "F"	
	elif len(sheet[:redrow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "G"
	elif len(sheet[:redrow, startcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "H"	
	elif len(sheet[redrow:, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "I"  
	else:
		sectionname = "J" 
	karute.sectionname = sectionname   # 区画名	
	karute.splittedrow = splittedrow  # スクロール枠の開始行インデックス。
	karute.startcolumn = startcolumn  # スクロール枠の開始列インデックス。
	karute.bluerow = bluerow  # 青3行インデックス。
	karute.skybluerow = skybluerow  # スカイブルー行インデックス。
	karute.redrow = redrow  # 赤3行インデックス。
	return karute  
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	cellrange = sheet["C1:J1"]  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	datarow = list(cellrange.getDataArray()[0])  # 行をリストで取得。
	datarow[0] = "一覧へ"
	datarow[2] = "経過へ"
	datarow[6] = "COPY"
	datarow[7] = "退院ｻﾏﾘ"
	cellrange.setDataArray((datarow,))  # 行をシートに戻す。
	sheet["J1"].setPropertyValue("CellBackColor", -1)  # 退院ｻﾏﾘボタンの背景色をクリアする。
	controller = activationevent.Source
	if len(controller)>3:  # シートが4分割されている時。
		controller[3].setFirstVisibleRow(0)  # 縦スクロールをリセット。controller[0].getVisibleRange()ではなぜか列インデックスが正しく取得できない。EndRowが0、EndColumnが9になる。
		controller[3].setFirstVisibleColumn(0)  # 横スクロールをリセット。
	target = controller[1].getReferredCells()[0, 0]  # 左下枠のS履歴列のセルを取得。列インデックスは0から7までならなんでもいいはず。
	karute = getSectionName(controller, sheet, target)  # セル固有の定数を取得。
	
	
	# 本日の記事を過去の記事に移動させる。
	dateformat = "****%Y年%m月%d日(%a)****"
	daterange = sheet[karute.bluerow, karute.kijicolumn]  # 本日の記事の日付セルを取得。
	articledatetxt = daterange.getString()  # 本日の記事の日付セルの文字列を取得。
	articledate = datetime().strptime(articledatetxt, dateformat)  # 記事列の日付を取得。
	todaydate = date.today()  # 今日のdateオブジェクトを取得。
	if articledate!=todaydate:  # 今日の日付でない時。
		todayarticle = sheet[karute.bluerow+1:karute.skybluerow, :]  # 青行とスカイブルー行の間の行のセル範囲。
		datarows = todayarticle[:, karute.sharpcolumn:karute.kijicolumn+1].getDataArray()  # 本日の記事欄のセルをすべて取得。
		txt = "".join(map(str, chain.from_iterable(datarows)))  # 本日の記事欄を文字列にしてすべて結合。
		cellranges = controller.getModel().createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
		if txt:  # 記事の文字列があるときのみ。
			newdatarows = [(articledatetxt,)]  # 先頭行に日付を入れる。
			stringlength = karute.stringlength  # 1セルあたりの文字数。
			newdatarows.extend((txt[i:i+stringlength],) for i in range(0, len(txt), stringlength))  # 過去記事欄へ代入するデータ。
			dest_start_ridx = karute.redrow + 1  # 移動先の開始行インデックス。
			dest_endbelow_ridx = dest_start_ridx + len(newdatarows)  # 移動先の最終行の下行の行インデックス。
			dest_rangeaddress = sheet[dest_start_ridx:dest_endbelow_ridx, karute.kijicolumn].getRangeAddress()  # 挿入前にセル範囲アドレスを取得しておく。
			sheet.insertCells(dest_rangeaddress, insert_rows)  # 赤行の下に空行を挿入。	
			sheet[dest_start_ridx:dest_endbelow_ridx, :].clearContents(511)  # 挿入した行の内容をすべてを削除。挿入セルは挿入した行の上のプロパティを引き継いでいるのでリセットしないといけない。
			dest_range = sheet.queryIntersection(dest_rangeaddress)[0]  # 赤行の下の挿入行のセル範囲を取得。セル挿入後はアドレスから取得し直さないといけない。
			dest_range.setDataArray(newdatarows)  # 過去の記事に挿入する。
			cellranges.addRangeAddress(dest_range.getRangeAddress(), False)  # あとでプロパティを設定するセル範囲コレクションに追加する。
			todayarticle.clearContents(511)  # 本日の記事欄をクリア。
		daterange.setString(todaydate.strftime(dateformat))  # 今日の日付を本日の記事欄に入力。
		cellranges.addRangeAddresses([todayarticle[:, i].getRangeAddress() for i in (2,4,6)], False)  # 本日の記事のDate列、Subject列、記事列のセル範囲コレクションを取得。
		cellranges.setPropertyValue("IsTextWrapped", True)  # セルの内容を折り返す。	
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	target = enhancedmouseevent.Target  # ターゲットのセルを取得。
	sheet = target.getSpreadsheet()
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				drowBorders(controller, sheet, target, commons.createBorders())  # 枠線の作成。
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				karute = getSectionName(controller, sheet, target)  # セル固有の定数を取得。
				sectionname = karute.sectionname  # クリックしたセルの区画名を取得。
				txt = target.getString()  # クリックしたセルの文字列を取得。	
				if sectionname=="A":
					sheets = doc.getSheets()  # シートコレクションを取得。
					if txt=="一覧へ":
						controller.setActiveSheet(sheets["一覧"])  # 一覧シートをアクティブにする。
					elif txt=="経過へ":
						newsheetname = "".join([sheet.getName(), "経"])  # 経過シート名を取得。
						if newsheetname in sheets:  # 経過シート名がある時。
							controller.setActiveSheet(sheets[newsheetname])  # 経過シートをアクティブにする。
						else:
							# 経過シートの作成。
							
							pass
					return False  # セルを編集モードにしない。
				elif sectionname=="B":	
					ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
					smgr = ctx.getServiceManager()  # サービスマネージャーの取得。						
					if txt=="COPY":
						splittedrow, bluerow, skybluerow = karute.splittedrow, karute.bluerow, karute.skybluerow
						getCopyDataRows, formatArticleColumn, formatProblemList, copyCells = createCopyFuncs(ctx, smgr, doc, sheet)
						c = formatArticleColumn(sheet[bluerow+1:skybluerow, karute.sharpcolumn:7])  # 本日の記事欄の記事列を整形。追加した行数が返る。
						datarows = sheet[bluerow:skybluerow+c, karute.sharpcolumn:7].getDataArray()  # 文字数制限後の行のタプルを取得。
						copydatarows = [(datarows[0][5],)]  # 本日の記事の日付を取得。
						deletedrowcount = getCopyDataRows(copydatarows, datarows[1:], bluerow+1)  # 削除された行数。
						if deletedrowcount>0:  # 削除した行があるとき。
							startrow = skybluerow - deletedrowcount
							newrangeaddress = sheet[startrow:startrow+deletedrowcount, :].getRangeAddress()  # 挿入するセル範囲アドレスを取得。
							sheet.insertCells(newrangeaddress, insert_rows)  # 空行を挿入。	
							sheet.queryIntersection(newrangeaddress).clearContents(511)  # 追加行の内容をクリア。セル範囲アドレスから取得しないと行挿入後のセル範囲が異なってしまう。
						newdatarows = formatProblemList(splittedrow, bluerow, "****ｻﾏﾘ****")  # プロブレム欄を整形。
						for i in (4, 6):  # Subject列と記事列について。
							newrange = sheet[splittedrow:, i]
							newrange.setPropertyValue("IsTextWrapped", True)  # セルの内容を折り返す。
							newrange.getRows().setPropertyValue("OptimalHeight", True)  # 内容を折り返した後の行の高さを調整。
						cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges") 
						cellranges.addRangeAddresses([i.getRangeAddress() for i in (sheet[splittedrow:bluerow, 1:5], sheet[bluerow+1:skybluerow, 1:5])], False)  # プロブレム欄、本日の記事欄をセル範囲を取得。
						cellranges.setPropertyValue("VertJustify", CellVertJustify2.CENTER)  # 縦位置を中央にする。
						newdatarows.extend(copydatarows)  # 本日の記事欄をプロブレム欄の下に追加。
						copieddatecell = sheet[0, 6]  # コピー日時セルを取得。	
						copyCells(controller, copieddatecell, newdatarows)
						copieddatecell.setString(datetime.now().strftime("%Y/%m/%d %H:%M:%S Copied"))  # コピーボタンを押した日付を入力。
						copieddatecell.setPropertyValues(("CellBackColor", "CharColor"), (commons.COLORS["lime"], -1))  # コピー日時セルの背景色を変更。文字色をリセット。
					elif txt=="退院ｻﾏﾘ":
						dummy, dummy, formatProblemList, copyCells = createCopyFuncs(ctx, smgr, doc, sheet)
						newdatarows = formatProblemList(karute.splittedrow, karute.bluerow, "****退院ｻﾏﾘ****")  # プロブレム欄を整形。
						copieddatecell = sheet[0, 6]  # コピー日時セルを取得。	
						copyCells(controller, copieddatecell, newdatarows)
						target.setPropertyValue("CellBackColor", commons.COLORS["lime"])  # 退院ｻﾏﾘボタンの背景色を変更。
					return False  # セルを編集モードにしない。
	return True  # セル編集モードにする。
def createCopyFuncs(ctx, smgr, doc, sheet):  # コピーのための関数を返す関数。
	karute = Karute()  # クラスをインスタンス化。	
	stringlength = karute.stringlength  # 1セルあたりの文字数。
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
		deletedrowcount = 0  # 空行の数。
		ranges = []  # 空行のセル範囲のリスト。
		for i, datarow in enumerate(datarows):
			rowtxt = _getRowTxt(datarow)  # #列、日付列、Subject列、記事列を結合した文字列を取得。
			if rowtxt:  # 空行でない時。
				copydatarows.append((rowtxt,))  # ｺﾋﾟｰ用シートにコピーする行のリストに取得。
			else:  # 空行の時。
				ranges.append(sheet[startrow+i, :])  # 削除行のセル範囲を取得。アドレスで取得するときは下から削除する必要がある。
				deletedrowcount += 1	
		cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
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
		cellranges = getCellRanges(doc, datarange, datarows)  # #ごとのセル範囲コレクションを取得。
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
			newrange = sheet[datarangestartrow:datarangestartrow+len(newarticlerows), karute.kijicolumn]  # 記事列のセル範囲を取得。
			newrange.clearContents(CellFlags.STRING+CellFlags.VALUE)  # 記事列の文字列と数値をクリア。
			newrange.setDataArray(newarticlerows)  # 記事列に代入。
		return c  # 追加した行数を返す。
	def formatProblemList(startrow, endrow, title):  # プロブレム欄を整形。
		c = formatArticleColumn(sheet[startrow:endrow, karute.sharpcolumn:karute.kijicolumn+1])  # プロブレム欄の記事列を整形。追加した行数が返る。
		datarows = sheet[startrow:endrow+c, karute.sharpcolumn:karute.kijicolumn+1].getDataArray()  # 文字数制限後の行のタプルを取得。
		newdatarows = [(title,)]  # タイトルを取得。	
		getCopyDataRows(newdatarows, datarows, startrow)  # プロブレム欄の記事列を整形。
		return newdatarows
	def copyCells(controller, copieddatecell, newdatarows):
		sheetname = "ｺﾋﾟｰ用"
		sheets = doc.getSheets()
		if not sheetname in sheets:  # コピー用シートがない時。
			sheets.insertNewByName(sheetname, len(sheets))  # コピー用シートを挿入。
		copysheet = doc.getSheets()[sheetname]  # コピー用シートを取得。
		copysheet.clearContents(511)  # シート内容をクリア。
		pasterange = copysheet[:len(newdatarows), :len(newdatarows[0])]  # コピー用シートのペーストするセル範囲を取得。。
		pasterange.setDataArray(newdatarows)  # コピー用シートにペーストする。
		pasterange.getColumns().setPropertyValue("Width", copieddatecell.getColumns().getPropertyValue("Width"))  # 単位は1/100mm
		pasterange.setPropertyValue("IsTextWrapped", True)  # ペーストしたセルの内容を折り返す。	
		pasterange.getRows().setPropertyValue("OptimalHeight", True)  # ペーストした行の高さを調整。
		dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
		controller.select(pasterange)  # ペーストしたセル範囲を取得。シートが切り替わってしまう。
		dispatcher.executeDispatch(controller.getFrame(), ".uno:Copy", "", 0, ())  # ペーストしたセルをコピー。
		controller.setActiveSheet(sheet)  # カルテシートに戻る.
	return getCopyDataRows, formatArticleColumn, formatProblemList, copyCells
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	controller = eventobject.Source
	selection = controller.getSelection()
	sheet = controller.getActiveSheet()
	if len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 列全体が選択されている時。
		drowBorders(controller, sheet, selection, commons.createBorders())  # 枠線の作成。
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):		
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
	del contextmenu[:]  # contextmenu.clear()は不可。
	target = controller.getSelection()  # 現在選択しているセル範囲を取得。
	if contextmenuname=="cell":  # セルのとき
		karute = getSectionName(controller, sheet, target)  # セル固有の定数を取得。
		sectionname = karute.sectionname  # クリックしたセルの区画名を取得。			
		if sectionname in ("A", "B"):  # 固定行より上の時はコンテクストメニューを表示しない。
			return EXECUTE_MODIFIED
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Delete"})	
# 		if target.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
# 			addMenuentry("ActionTrigger", {"Text": "To Green", "CommandURL": baseurl.format("entry1")}) 
# 		elif target.supportsService("com.sun.star.sheet.SheetCellRange"):  # 連続した複数セルの時。
# 			addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format("entry2")}) 
	elif contextmenuname=="rowheader":  # 行ヘッダーのとき。
		karute = getSectionName(controller, sheet, target[0, 0])  # 選択範囲の最初のセルの定数を取得。
		sectionname = karute.sectionname  # クリックしたセルの区画名を取得。			
		if sectionname in ("A",) or target[0, 0].getPropertyValue("CellBackColor")!=-1:  # 背景色のあるときは表示しない。
			return EXECUTE_MODIFIED
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsBefore"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsAfter"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteRows"}) 
		if sectionname in ("C",):
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			addMenuentry("ActionTrigger", {"Text": "最下行へ", "CommandURL": baseurl.format("entry1")})  
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			addMenuentry("ActionTrigger", {"Text": "過去ﾘｽﾄへ移動", "CommandURL": baseurl.format("entry2")})  
			addMenuentry("ActionTrigger", {"Text": "過去ﾘｽﾄにｺﾋﾟｰ", "CommandURL": baseurl.format("entry3")})  
		elif sectionname in ("G",):
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			addMenuentry("ActionTrigger", {"Text": "現ﾘｽﾄへ移動", "CommandURL": baseurl.format("entry4")})  
			addMenuentry("ActionTrigger", {"Text": "現ﾘｽﾄにｺﾋﾟｰ", "CommandURL": baseurl.format("entry5")})  
	elif contextmenuname=="colheader":  # 列ヘッダーの時。
		pass
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	selection = controller.getSelection()
	if len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 列全体が選択されている場合もあるので行全体が選択されていることを確認する。
		karute = getSectionName(controller, sheet, selection[0, 0])
		splittedrow = karute.splittedrow
		bluerow = karute.bluerow
		skybluerow = karute.skybluerow
		redrow = karute.redrow
		sharpcolumn = karute.sharpcolumn
		kijicolumn = karute.kijicolumn
		if entrynum==1:  # 現リストの最下行へ。青行の上に移動する。セクションC。
			dest_start_ridx = bluerow  # 移動先開始行インデックス。	
			problemranges = getProblemRanges(doc, sheet[splittedrow:bluerow, sharpcolumn:kijicolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。			
		elif entrynum==2:  # 過去ﾘｽﾄへ移動。スカイブルー行の下に移動する。セクションC。
			dest_start_ridx = skybluerow + 1  # 移動先開始行インデックス。
			problemranges = getProblemRanges(doc, sheet[splittedrow:bluerow, sharpcolumn:kijicolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。		
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。					
		elif entrynum==3:  # 過去ﾘｽﾄにｺﾋﾟｰ。スカイブルー行の下にコピーする。
			dest_start_ridx = skybluerow + 1  # 移動先開始行インデックス。
			problemranges = getProblemRanges(doc, sheet[splittedrow:bluerow, sharpcolumn:kijicolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。		
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.copyRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
		elif entrynum==4:  # 現ﾘｽﾄへ移動。青行の上に移動する。
			dest_start_ridx = bluerow  # 移動先開始行インデックス。	
			problemranges = getProblemRanges(doc, sheet[skybluerow+1:redrow, sharpcolumn:kijicolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。			
		elif entrynum==5:  # 現ﾘｽﾄにｺﾋﾟｰ。青行の上にコピーする。
			dest_start_ridx = bluerow  # 移動先開始行インデックス。	
			problemranges = getProblemRanges(doc, sheet[skybluerow+1:redrow, sharpcolumn:kijicolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.copyRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
def moveProblems(sheet, problemrange, dest_start_ridx):  # problemrange; 問題リストの塊。dest_start_ridx: 移動先開始行インデックス。
	dest_endbelow_ridx = dest_start_ridx + len(problemrange.getRows())  # 移動先最終行の次の行インデックス。
	dest_rangeaddress = sheet[dest_start_ridx:dest_endbelow_ridx, :].getRangeAddress()  # 挿入前にセル範囲アドレスを取得しておく。
	sheet.insertCells(dest_rangeaddress, insert_rows)  # スカイブルー行の下に空行を挿入。	
	sheet.queryIntersection(dest_rangeaddress).clearContents(511)  # 挿入した行の内容をすべてを削除。挿入セルは挿入した行の上のプロパティを引き継いでいるのでリセットしないといけない。
	cursor = sheet.createCursorByRange(problemrange)  # セルカーサーを作成		
	cursor.expandToEntireRows()  # セル範囲を行全体に拡大。		
	return cursor.getRangeAddress()  # 移動前に移動元の問題リストのセル範囲アドレスを取得しておく。
def drowBorders(controller, sheet, cellrange, borders):  # cellrangeを交点とする行列全体の外枠線を描く。
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = borders  # 枠線を取得。	
	cell = cellrange[0, 0]  # セル範囲の左上端のセルで判断する。
	rangeaddress = cell.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	karute = getSectionName(controller, sheet, cell)  # セル固有の定数を取得。
	sectionname = karute.sectionname  # クリックしたセルの区画名を取得。
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
	if sectionname in ("A", "B", "E", "I"):  # 枠線を消すだけ。
		return
	if sectionname in ("C", "G"):  # 同一プロブレムの上下に枠線を引く。
		datarange = sheet[karute.splittedrow:karute.bluerow, karute.sharpcolumn:7] if sectionname=="C" else sheet[karute.skybluerow+1:karute.redrow, karute.sharpcolumn:7]  # タイトル行を除く。
		doc = controller.getModel()  # ドキュメントモデルを取得。
		problemranges = getProblemRanges(doc, datarange, cellrange)  # 問題ごとのセル範囲コレクションを取得。
		for i in problemranges:
			cursor = sheet.createCursorByRange(i)  # セルカーサーを作成		
			cursor.expandToEntireRows()  # セル範囲を行全体に拡大。	
			cursor.setPropertyValue("TableBorder2", topbottomtableborder) # プロブレムの上下に枠線を引く。
	elif sectionname in ("D", "F", "H"):
		sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。			
		sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。	
		cellrange.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	
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
