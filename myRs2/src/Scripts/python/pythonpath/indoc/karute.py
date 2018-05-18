#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# カルテシートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from datetime import date
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
class Karute():  # シート固有の定数設定。
	pass
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
	datecolumn = 2  # Date列インデックス。
	subcontollerrange = controller[0].getVisibleRange()
	startrow = subcontollerrange.EndRow + 1  # スクロールする枠の最初の行インデックス。
	startcolumn = subcontollerrange.EndColumn + 1  # スクロールする枠の最初の列インデックス。
	cellranges = sheet[startrow:, datecolumn].queryContentCells(CellFlags.STRING)  # Date列の文字列が入っているセルに限定して抽出。
	backcolors = commons.COLORS["blue3"], commons.COLORS["skyblue"], commons.COLORS["red3"]  # ジェネレーターに使うので順番が重要。
	gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in backcolors)
	bluerow = next(gene)
	skybluerow = next(gene)
	redrow = next(gene)
	karute = Karute()  # クラスをインスタンス化。	
	rangeaddress = target.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	if len(sheet[:startrow, :startcolumn].queryIntersection(rangeaddress)): 
		sectionname = "A"
	elif len(sheet[:startrow, startcolumn:].queryIntersection(rangeaddress)): 
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
	karute.startrow = startrow  # スクロール枠の開始行インデックス。
	karute.startcolumn = startcolumn  # スクロール枠の開始列インデックス。
	karute.bluerow = bluerow  # 青3行インデックス。
	karute.skybluerow = skybluerow  # スカイブルー行インデックス。
	karute.redrow = redrow  # 赤3行インデックス。
	return karute  
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	sheet["C1"].setString("一覧へ")
	sheet["E1"].setString("経過へ")
	sheet["I1"].setString("COPY")
	controller = activationevent.Source
	if len(controller)>3:  # シートが4分割されている時。
		controller[3].setFirstVisibleRow(0)  # 縦スクロールをリセット。controller[0].getVisibleRange()ではなぜか列インデックスが正しく取得できない。EndRowが0、EndColumnが9になる。
		controller[3].setFirstVisibleColumn(0)  # 横スクロールをリセット。
	target = controller[1].getReferredCells()[0, 0]  # 左下枠のS履歴列のセルを取得。列インデックスは0から7までならなんでもいいはず。
	karute = getSectionName(controller, sheet, target)  # セル固有の定数を取得。
	daterange = sheet[karute.bluerow, 5:7]  # 日付データが入っているセル範囲。
	articleday, articledayformated = daterange.getDataArray()[0]  # 青行のISO8601形式の日付とフォーマットされた日付を取得。
	todaydate = date.today()  # 今日のdateオブジェクトを取得。
	if articleday!=todaydate.isoformat():  # 青行の列インデックス5の文字列が今日の日付でない時。
		daterange[0, 0].setPropertyValue("CharColor", commons.COLORS["blue3"])  # ISO8601形式の日付は見せない。
		todayarticle = sheet[karute.bluerow+1:karute.skybluerow, :]  # 青行とスカイブルー行の間の行のセル範囲。
		datarows = todayarticle[:, 1:7].getDataArray()  # 本日の記事欄のセルをすべて取得。
		txt = "".join(map(str, chain.from_iterable(datarows)))  # 本日の記事欄を文字列にしてすべて結合。
		cellranges = controller.getModel().createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
		if txt:  # 記事の文字列があるときのみ。
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
			txt = createTransliteration(ctx, smgr)(txt)
			newdatarows = [(articledayformated,)]  # 先頭行に日付を入れる。
			stringlength = 125  # 1セルあたりの文字数。
			newdatarows.extend((txt[i:i+stringlength],) for i in range(0, len(txt), stringlength))  # 過去記事欄へ代入するデータ。
			dest_start_ridx = karute.redrow + 1  # 移動先の開始行インデックス。
			dest_endbelow_ridx = dest_start_ridx + len(newdatarows)  # 移動先の最終行の下行の行インデックス。
			dest_rangeaddress = sheet[dest_start_ridx:dest_endbelow_ridx, 6].getRangeAddress()  # 挿入前にセル範囲アドレスを取得しておく。
			sheet.insertCells(dest_rangeaddress, insert_rows)  # 赤行の下に空行を挿入。	
			sheet[dest_start_ridx:dest_endbelow_ridx, :].clearContents(511)  # 挿入した行の内容をすべてを削除。挿入セルは挿入した行の上のプロパティを引き継いでいるのでリセットしないといけない。
			dest_range = sheet.queryIntersection(dest_rangeaddress)[0]  # 赤行の下の挿入行のセル範囲を取得。セル挿入後はアドレスから取得し直さないといけない。
			dest_range.setDataArray(newdatarows)  # 過去の記事に挿入する。
			cellranges.addRangeAddress(dest_range.getRangeAddress(), False)  # あとでプロパティを設定するセル範囲コレクションに追加する。
			todayarticle.clearContents(511)  # 本日の記事欄をクリア。
		daterange.setDataArray(((todaydate.isoformat(), todaydate.strftime("****%Y年%m月%d日(%a)****")),))  # 今日の日付を青行に代入。
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
				elif sectionname=="B":			
					if txt=="COPY":
						startrow, bluerow, skybluerow = karute.startrow, karute.bluerow, karute.skybluerow
						ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
						smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
						functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。			
						datarows = sheet[bluerow:skybluerow, 1:7].getDataArray()  # 本日の記事領域の行のタプルを取得。
						newdatarows2 = [(datarows[0][-1],)]  # 青行の日付を取得。
						lines = []  # 取得行のリスト。
						for datarow in datarows[1:]:  # 青行の下の行から開始。
							rowtxt = getRowTxt(functionaccess, datarow)  # 行のデータを結合した文字列を取得。				
							if rowtxt:														
								if rowtxt.startswith("#"):  # 行頭が#の時。
									if lines:  # 行がすでに取得出来ている時。
										articlerows = getArticleRows(lines)
										if ":" in articlerows[0][0]:
											articlerows[0][0] = articlerows[0][0].split(":")[-1]
										elif "#" in articlerows[0][0]:
											articlerows[0][0] = articlerows[0][0][1:]
										
						
										
										newdatarows2.extend(articlerows)  # 文字数制限をして新しいデータ行に追加。
									lines = [rowtxt]  # 取得行をリセットする。
								else:
									lines.append(rowtxt)	
						else:
							if lines:  # まだ取得行が残っている時。
								newdatarows2.extend(getArticleRows(lines))  # 文字数制限をして新しいデータ行に追加。
								diff = len(newdatarows2[1:]) - (skybluerow - bluerow - 1)
								if diff>0:  # 行が増えた時行を追加する。
									newrange = sheet[skybluerow:skybluerow+diff, :]
									sheet.insertCells(newrange.getRangeAddress(), insert_rows)  # 空行を挿入。
						pstartrow = bluerow + 1
						newrange = sheet[pstartrow:pstartrow+len(newdatarows2[1:]), :]			
						newrange.clearContents(511)												
						newrange[:, 6].setDataArray(newdatarows2[1:])
						datarows = sheet[startrow:bluerow, 1:7].getDataArray()  # 問題リスト領域について。日付列はLibreOfficeのシリアル値で返ってくる。
						lines = []
						newdatarows1 = []
						
# 						import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
						
						for i, datarow in enumerate(reversed(datarows)):  # 下行からいく。
							rowtxt = getRowTxt(functionaccess, datarow)			
							if rowtxt:
								lines.append(rowtxt)  # 逆順に行を取得する。
								if rowtxt.startswith("#"):  # プロブレム初行の時。このときはすでにプロブレムの全行が取得されているはず。
									articlerows = getArticleRows(reversed(lines))  # 正順でセルあたりの文字数を制限。
									newdatarows1.extend(reversed(articlerows))  # 新しいデータ行に逆順に追加。
									if len(articlerows)>len(lines):  # 行が増えた時行を追加する。
										pstartrow = bluerow - 1 - i  # プロブレム行の開始行インデックス。
										newrange = sheet[pstartrow+len(lines):pstartrow+len(articlerows), :]
										sheet.insertCells(newrange.getRangeAddress(), insert_rows)  # 空行をプロブレムごとに挿入。	
										newrange.clearContents(511)
									lines = []  # 取得行をリセット。
						newrange = sheet[startrow:startrow+len(newdatarows1), 6]
						newrange.clearContents(511)	
						newdatarows1.reverse()  # 行を正順に戻す。
						newrange.setDataArray(newdatarows1)
						newdatarows1.extend(newdatarows2)
						newrange = sheet[startrow:startrow+len(newdatarows1), 6]
						newrange.setPropertyValue("IsTextWrapped", True)  # セルの内容を折り返す。	
						newrange.getRows().setPropertyValue("OptimalHeight", True)						
						newdatarows = [("****ｻﾏﾘ****",)]
						newdatarows.extend(newdatarows1)
						copysheet = doc.getSheets()["ｺﾋﾟｰ用"]
						copysheet.clearContents(511)  # シート内容をクリア。
						pasterange = copysheet[:len(newdatarows), :len(newdatarows[0])]
						pasterange.setDataArray(newdatarows)
						pasterange.setPropertyValue("IsTextWrapped", True)  # セルの内容を折り返す。	
						pasterange.getRows().setPropertyValue("OptimalHeight", True)
						dispatcher = ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
						controller.select(pasterange)  # シートが切り替わってしまう。
						dispatcher.executeDispatch(controller.getFrame(), ".uno:Copy", "", 0, ())
						controller.setActiveSheet(sheet)  # 元のシートに戻る。
						return False  # セルを編集モードにしない。

						

# 						fullwidth_halfwidth = createTransliteration(ctx, smgr)
# 						systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
# 						datarange = sheet[karute.startrow:karute.bluerow, 1:7]  # タイトル行を除く記事範囲を取得。
# 						newdatarows = [map(fullwidth_halfwidth, i) for i in datarange.getDataArray()]
# 						datarange.setDateArray(newdatarows)
# 						cellranges = getCellRanges(doc, datarange)  # 問題リストの各セル範囲を取得。
# 						for cellrange in cellranges:
# 							datarows = cellrange[:, -1].getDataArray()
# 							kijitxt = "".join(map(str, chain.from_iterable(datarows)))  # 本日の記事欄を文字列にしてすべて結合。
# 							stringlength = 254  # 1セルあたりの文字数。
# 							newdatarows = [(kijitxt[i:i+stringlength],) for i in range(0, len(kijitxt), stringlength)]  # 記事欄へ代入するデータ。
# 							diff = len(newdatarows) - len(cellrange)
# 							if diff>0:
# 								rangeaddress = cellrange.getRangeAddress()
# 								endrow = rangeaddress.EndRow + diff + 1
# 								diffrange = sheet[rangeaddress.EndRow+1:endrow, :]
# 								sheet.insertCells(diffrange.getRangeAddress(), insert_rows)  
# 								diffrange.clearContents(511) 
# 								sheet[rangeaddress.StartRow:endrow, 6].setDataArray(newdatarows)
# 						cellranges = controller.getModel().createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
# 						cellranges.addRangeAddresses([datarange[:, i].getRangeAddress() for i in (2,4,6)], False)  # 本日の記事のDate列、Subject列、記事列のセル範囲コレクションを取得。
# 						cellranges.setPropertyValue("IsTextWrapped", True)  # セルの内容を折り返す。	
						
				
	
				
				
				
				
	return True  # セル編集モードにする。
def getArticleRows(lines):
	stringlength = 125  # 1セルあたりの文字数。
	ptxt = "".join(lines)  # 取得した行をすべて結合。	
	return [(ptxt[i:i+stringlength],) for i in range(0, len(ptxt), stringlength)]  # 文字列をセルあたりの文字数で分割。
def getRowTxt(functionaccess, datarow):
	sharpcol, datecol, subjectcol, articlecol = datarow[0], datarow[1], datarow[3], datarow[5]  # #列、日付列、Subject列、記事列を取得。
	if datecol and isinstance(datecol, float):  # 日付列がfloat型のとき。
		datecol = "{} ".format("/".join([str(int(functionaccess.callFunction(i, (datecol,)))) for i in ("YEAR", "MONTH", "DAY")]))  # シリアル値をシート関数で年/月/日の文字列にする。引数のdatecolはfloatのままでよい。
	if subjectcol and subjectcol!="#":   # Subject列、かつ、#でないとき
		subjectcol = "{}: ".format(subjectcol)	
	return "{}{}{}{}".format(sharpcol, datecol, subjectcol, articlecol)  # #列、Date列、Subject列、記事列を結合。
def createTransliteration(ctx, smgr):
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
	transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))  # 全角文字を半角にする。
	def fullwidth_halfwidth(txt):
		if isinstance(txt, str):  # 文字列の時。
			return transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換。
		else:  # 文字列でないときはそのまま返す。
			return txt
	return fullwidth_halfwidth
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
		if entrynum==1:  # 現リストの最下行へ。青行の上に移動する。セクションC。
			dest_start_ridx = karute.bluerow  # 移動先開始行インデックス。	
			problemranges = getProblemRanges(doc, sheet[karute.startrow:karute.bluerow, 1:7], selection)  # 問題ごとのセル範囲コレクションを取得。
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。			
		elif entrynum==2:  # 過去ﾘｽﾄへ移動。スカイブルー行の下に移動する。セクションC。
			dest_start_ridx = karute.skybluerow + 1  # 移動先開始行インデックス。
			problemranges = getProblemRanges(doc, sheet[karute.startrow:karute.bluerow, 1:7], selection)  # 問題ごとのセル範囲コレクションを取得。		
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。					
		elif entrynum==3:  # 過去ﾘｽﾄにｺﾋﾟｰ。スカイブルー行の下にコピーする。
			dest_start_ridx = karute.skybluerow + 1  # 移動先開始行インデックス。
			problemranges = getProblemRanges(doc, sheet[karute.startrow:karute.bluerow, 1:7], selection)  # 問題ごとのセル範囲コレクションを取得。		
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.copyRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
		elif entrynum==4:  # 現ﾘｽﾄへ移動。青行の上に移動する。
			dest_start_ridx = karute.bluerow  # 移動先開始行インデックス。	
			problemranges = getProblemRanges(doc, sheet[karute.skybluerow+1:karute.redrow, 1:7], selection)  # 問題ごとのセル範囲コレクションを取得。
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。			
		elif entrynum==5:  # 現ﾘｽﾄにｺﾋﾟｰ。青行の上にコピーする。
			dest_start_ridx = karute.bluerow  # 移動先開始行インデックス。	
			problemranges = getProblemRanges(doc, sheet[karute.skybluerow+1:karute.redrow, 1:7], selection)  # 問題ごとのセル範囲コレクションを取得。
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.copyRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
def moveProblems(sheet, problemrange, dest_start_ridx):  # problemrange; 問題リストの塊。dest_start_ridx: 移動先開始行インデックス。
	dest_endbelow_ridx = dest_start_ridx + len(problemrange.getRows())  # 移動先最終行の次の行インデックス。
	dest_rangeaddress = sheet[dest_start_ridx:dest_endbelow_ridx, 0].getRangeAddress()  # 挿入前にセル範囲アドレスを取得しておく。
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
		datarange = sheet[karute.startrow:karute.bluerow, 1:7] if sectionname=="C" else sheet[karute.skybluerow+1:karute.redrow, 1:7]  # タイトル行を除く。
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
def getCellRanges(doc, datarange):  # 各プロブレムの行をまとめたセル範囲コレクションを返す。列はdatarangeと同じ。
	cellranges = []
	datarows = datarange.getDataArray()  # #列からSubject列までの行のタプルを取得。
	ranges = []  # プロブレムリストのセル範囲のリスト。
	rstartrow = 0  # プロブレム開始行の相対インデックス。
	for i, datarow in enumerate(datarows):  # 相対インデックスと行のタプルを列挙。
		if "#" in "{}{}{}".format(datarow[0], datarow[1], datarow[3]):  # #列、Date列、Subject列を結合して#がある時。日付は数値なので文字列への変換が必要なのでjoin()は使えない。
			if i>rstartrow:  # 開始行相対インデックスより大きい時。
				ranges.append(datarange[rstartrow:i, :])  # プロブレムリストの開始行から終了行までのセル範囲を取得。
				rstartrow = i
	if ranges:  # すでにプロブレムがあるときのみ。一つも取得できていないときは一つもプロブレムがないので取得しない。
		ranges.append(datarange[rstartrow:, :])  # 最後のプロブレムのセル範囲を追加。
		cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
		cellranges.addRangeAddresses([i.getRangeAddress() for i in ranges], False)  # セル範囲コレクションにプロブレムのセル範囲を追加する。セル範囲は結合しない。
	return cellranges  # 列はdatarangeと同じ。
