#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# カルテシートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from datetime import date, datetime, timedelta
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
from com.sun.star.table.CellHoriJustify import LEFT, RIGHT  # enum
class Karute():  # シート固有の定数設定。
	def __init__(self, sheet):
		self.splittedrow = 2  # 分割行インデックス。
		self.sharpcolumn = 1  # #列インデックス。
		self.datecolumn = 2  # Date列インデックス。
		self.subjectcolumn = 4  # Subject列インデックス。
		self.articlecolumn = 6  # Article列インデックス。
		self.splittedcolumn = 10  # 分割列インデックス。コントローラーから動的取得が正しく出来ない。
		self.stringlength = 125  # 1セルあたりの文字数。
		self.dateformat = "%Y/%m/%d %H:%M:%S Copied"  # 記事をコピーした日時の書式。
		cellranges = sheet[self.splittedrow:, self.datecolumn].queryContentCells(CellFlags.STRING)  # Date列の文字列が入っているセルに限定して抽出。
		backcolors = commons.COLORS["blue3"], commons.COLORS["skyblue"], commons.COLORS["red3"]  # ジェネレーターに使うので順番が重要。
		gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in backcolors)
		self.bluerow = next(gene)  # 青3行インデックス。
		self.skybluerow = next(gene)  # スカイブルー行インデックス。
		self.redrow = next(gene)  # 赤3行インデックス。		
def getSectionName(sheet, target):  # 区画名を取得。
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
	karute = Karute(sheet)  # クラスをインスタンス化。	
	splittedrow = karute.splittedrow
	splittedcolumn = karute.splittedcolumn
	bluerow = karute.bluerow
	skybluerow = karute.skybluerow
	redrow = karute.redrow
	rangeaddress = target.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	if len(sheet[:splittedrow, :splittedcolumn].queryIntersection(rangeaddress)): 
		sectionname = "A"
	elif len(sheet[:splittedrow, splittedcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "B"
	elif len(sheet[:bluerow, :splittedcolumn].queryIntersection(rangeaddress)): 
		sectionname = "C"
	elif len(sheet[:bluerow, splittedcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "D"
	elif len(sheet[:skybluerow, :splittedcolumn].queryIntersection(rangeaddress)): 
		sectionname = "E"
	elif len(sheet[:skybluerow, splittedcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "F"	
	elif len(sheet[:redrow, :splittedcolumn].queryIntersection(rangeaddress)): 
		sectionname = "G"
	elif len(sheet[:redrow, splittedcolumn:].queryIntersection(rangeaddress)): 
		sectionname = "H"	
	elif len(sheet[redrow:, :splittedcolumn].queryIntersection(rangeaddress)): 
		sectionname = "I"  
	else:
		sectionname = "J" 
	karute.sectionname = sectionname   # 区画名	
	return karute  
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。
	doc = xscriptcontext.getDocument()
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	karute = Karute(sheet)  # クラスをインスタンス化。	
	cellrange = sheet["A1:M1"]  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	datarow = list(cellrange.getDataArray()[0])  # 行をリストで取得。
	datarow[karute.datecolumn] = "一覧へ"
	datarow[karute.subjectcolumn] = "経過へ"
	datarow[karute.splittedcolumn] = "COPY"
	datarow[karute.splittedcolumn+1] = "退院ｻﾏﾘ"
	datarow[karute.splittedcolumn+2] = "#分離"
	sheet[0, karute.splittedcolumn+1].setPropertyValue("CellBackColor", -1)  # 退院ｻﾏﾘボタンの背景色をクリアする。
	cellrange.setDataArray((datarow,))  # 行をシートに戻す。
	# コピー日時セルの色を設定。
	copieddatecell = sheet[0, karute.articlecolumn]  # コピー日時セルを取得。
	copieddatetxt = copieddatecell.getString()  # コピー日時セルの文字列を取得。
	if copieddatetxt:
		copieddatetime = datetime.strptime(copieddatetxt, karute.dateformat)  # コピーした日時を取得。
		now = datetime.now()  # 現在の日時を取得。
		if copieddatetime.date()<now.date():  # 今日はまだコピーしていない時。
			copieddatecell.setPropertyValues(("CharColor", "CellBackColor"), (-1, commons.COLORS["magenta3"]))  # 文字色をリセットして背景色をマゼンダにする。
		elif now.hour>12 and copieddatetime.hour<12:  # 今日はコピーしていても、午後になって午前にしかコピーしていない時。
			copieddatecell.setPropertyValue("CharColor", commons.COLORS["magenta3"])  # 文字色をマゼンダにする。背景色はコピーした時にすでにライムになっているはず。
	# 本日の記事を過去の記事に移動させる。
	dateformat = "****%Y年%m月%d日(%a)****"
	daterange = sheet[karute.bluerow, karute.articlecolumn]  # 本日の記事の日付セルを取得。
	articledatetxt = daterange.getString()  # 本日の記事の日付セルの文字列を取得。
	articledate = datetime.strptime(articledatetxt, dateformat)  # 記事列の日付を取得。
	todaydate = date.today()  # 今日のdateオブジェクトを取得。
	if articledate!=todaydate:  # 今日の日付でない時。
		todayarticle = sheet[karute.bluerow+1:karute.skybluerow, :]  # 青行とスカイブルー行の間の行のセル範囲。
		datarows = todayarticle[:, karute.sharpcolumn:karute.articlecolumn+1].getDataArray()  # 本日の記事欄のセルをすべて取得。
		txt = "".join(map(str, chain.from_iterable(datarows)))  # 本日の記事欄を文字列にしてすべて結合。
		cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
		if txt:  # 記事の文字列があるときのみ。
			newdatarows = [(articledatetxt,)]  # 先頭行に日付を入れる。
			stringlength = karute.stringlength  # 1セルあたりの文字数。
			newdatarows.extend((txt[i:i+stringlength],) for i in range(0, len(txt), stringlength))  # 過去記事欄へ代入するデータ。
			dest_start_ridx = karute.redrow + 1  # 移動先の開始行インデックス。
			dest_endbelow_ridx = dest_start_ridx + len(newdatarows)  # 移動先の最終行の下行の行インデックス。
			dest_rangeaddress = sheet[dest_start_ridx:dest_endbelow_ridx, karute.articlecolumn].getRangeAddress()  # 挿入前にセル範囲アドレスを取得しておく。
			sheet.insertCells(dest_rangeaddress, insert_rows)  # 赤行の下に空行を挿入。	
			sheet[dest_start_ridx:dest_endbelow_ridx, :].clearContents(511)  # 挿入した行の内容をすべて削除。挿入セルは挿入した行の上のプロパティを引き継いでいるのでリセットしないといけない。
			dest_range = sheet.queryIntersection(dest_rangeaddress)[0]  # 赤行の下の挿入行のセル範囲を取得。セル挿入後はアドレスから取得し直さないといけない。
			dest_range.setDataArray(newdatarows)  # 過去の記事に挿入する。
			cellranges.addRangeAddress(dest_range.getRangeAddress(), False)  # あとでプロパティを設定するセル範囲コレクションに追加する。
			todayarticle.clearContents(511)  # 本日の記事欄をクリア。
		daterange.setString(todaydate.strftime(dateformat))  # 今日の日付を本日の記事欄に入力。
		cellranges.addRangeAddresses([todayarticle[:, i].getRangeAddress() for i in (karute.datecolumn, karute.subjectcolumn, karute.articlecolumn)], False)  # 本日の記事のDate列、Subject列、記事列のセル範囲コレクションを取得。
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
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
				karute = getSectionName(sheet, target)  # セル固有の定数を取得。
				sectionname = karute.sectionname  # クリックしたセルの区画名を取得。
				txt = target.getString()  # クリックしたセルの文字列を取得。	
				if sectionname in ("A",):
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
				elif sectionname in ("B",):						
					if txt=="COPY":
						splittedrow, bluerow, skybluerow = karute.splittedrow, karute.bluerow, karute.skybluerow
						sharpcolumn, articlecolumn, subjectcolumn = karute.sharpcolumn, karute.articlecolumn, karute.subjectcolumn
						getCopyDataRows, formatArticleColumn, formatProblemList, copyCells = createCopyFuncs(ctx, smgr, doc, sheet)
						c = formatArticleColumn(sheet[bluerow+1:skybluerow, sharpcolumn:articlecolumn+1])  # 本日の記事欄の記事列を整形。追加した行数が返る。
						datarows = sheet[bluerow:skybluerow+c, sharpcolumn:articlecolumn+1].getDataArray()  # 文字数制限後の行のタプルを取得。
						copydatarows = [(datarows[0][5],)]  # 本日の記事の日付を取得。
						deletedrowcount = getCopyDataRows(copydatarows, datarows[1:], bluerow+1)  # 削除された行数。
						if deletedrowcount>0:  # 削除した行があるとき。
							startrow = skybluerow - deletedrowcount
							newrangeaddress = sheet[startrow:startrow+deletedrowcount, :].getRangeAddress()  # 挿入するセル範囲アドレスを取得。
							sheet.insertCells(newrangeaddress, insert_rows)  # 空行を挿入。	
							sheet.queryIntersection(newrangeaddress).clearContents(511)  # 追加行の内容をクリア。セル範囲アドレスから取得しないと行挿入後のセル範囲が異なってしまう。
						newdatarows = formatProblemList(splittedrow, bluerow, "****ｻﾏﾘ****")  # プロブレム欄を整形。
						for i in (subjectcolumn, articlecolumn):  # Subject列と記事列について。
							newrange = sheet[splittedrow:, i]
							newrange.setPropertyValue("IsTextWrapped", True)  # セルの内容を折り返す。
							newrange.getRows().setPropertyValue("OptimalHeight", True)  # 内容を折り返した後の行の高さを調整。
						cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges") 
						cellranges.addRangeAddresses([i.getRangeAddress() for i in (sheet[splittedrow:bluerow, sharpcolumn:subjectcolumn+1], sheet[bluerow+1:skybluerow, sharpcolumn:subjectcolumn+1])], False)  # プロブレム欄、本日の記事欄をセル範囲を取得。
						cellranges.setPropertyValue("VertJustify", CellVertJustify2.CENTER)  # 縦位置を中央にする。
						newdatarows.extend(copydatarows)  # 本日の記事欄をプロブレム欄の下に追加。
						copieddatecell = sheet[0, articlecolumn]  # コピー日時セルを取得。	
						copyCells(controller, copieddatecell, newdatarows)
						copieddatecell.setString(datetime.now().strftime(karute.dateformat))  # コピーボタンを押した日付を入力。
						copieddatecell.setPropertyValues(("CellBackColor", "CharColor"), (commons.COLORS["lime"], -1))  # コピー日時セルの背景色を変更。文字色をリセット。
					elif txt=="退院ｻﾏﾘ":
						dummy, dummy, formatProblemList, copyCells = createCopyFuncs(ctx, smgr, doc, sheet)
						newdatarows = formatProblemList(karute.splittedrow, karute.bluerow, "****退院ｻﾏﾘ****")  # プロブレム欄を整形。
						copieddatecell = sheet[0, karute.articlecolumn]  # コピー日時セルを取得。	
						copyCells(controller, copieddatecell, newdatarows)
						target.setPropertyValue("CellBackColor", commons.COLORS["lime"])  # 退院ｻﾏﾘボタンの背景色を変更。
					elif txt=="#分離":
						functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。
						splittedrow = karute.splittedrow
						bluerow = karute.bluerow
						datarange = sheet[splittedrow:bluerow, :karute.articlecolumn+1]
						datarows = datarange.getDataArray()
						newdatarows = []
						for datarow in datarows:
							datatxt = datarow[karute.articlecolumn]  # Article列の文字列を取得。
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
						sheet[splittedrow:bluerow, karute.sharpcolumn].setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.CENTER))  # #列の書式設定。左寄せにする。
						createFormatKey = commons.formatkeyCreator(doc)	
						sheet[splittedrow:bluerow, karute.datecolumn].setPropertyValues(("NumberFormat", "HoriJustify", "VertJustify"), (createFormatKey('YYYY/MM/DD'), LEFT, CellVertJustify2.CENTER))  # カルテシートの入院日の書式設定。左寄せにする。
						sheet[splittedrow:bluerow, karute.subjectcolumn].setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.CENTER))  # Subject列の書式設定。左寄せにする。
					return False  # セルを編集モードにしない。
				elif sectionname in ("C", "E", "G", "I"):	
					functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。			
					celladdress = target.getCellAddress()
					r, c = celladdress.Row, celladdress.Column  # ダブルクリックしたセルの行インデックス、列インデックスを取得。
					if c==0:  # 行挿列の時。
						sheet.insertCells(sheet[r+1, :].getRangeAddress(), insert_rows)  # ダブルクリックした行の下に空行を挿入。	
						sheet[r+1, :].setPropertyValues(("CellBackColor", "CharColor"), (-1, -1))  # 追加行の背景色と文字色をクリア。						
						return False  # セルを編集モードにしない。
					elif r in (karute.bluerow, karute.skybluerow, karute.redrow):  # カラー行の時。
						if txt=="問題ﾘｽﾄへ変換":
							articlecolumn = karute.articlecolumn
							sharpcolumn = karute.sharpcolumn
							cellranges = sheet[karute.redrow+1:, articlecolumn].queryContentCells(CellFlags.STRING)  # Article列の文字列が入っているセルに限定して抽出。
							if len(cellranges):  # セル範囲が取得出来た時。
								transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
								transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))  # 全角文字を半角にする。
								newdatarows = [] 
								emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # ID列の最終行インデックス+1を取得。
								datarange = sheet[karute.redrow+1:emptyrow, articlecolumn]
								datarows = datarange.getDataArray()
								stringlength = karute.stringlength  # 1セルあたりの文字数。
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
								problemrange = sheet[karute.splittedrow:karute.bluerow, karute.sharpcolumn:articlecolumn+1]
								cellranges = problemrange.queryContentCells(CellFlags.STRING)
								emptyrow = max(i.EndRow for i in cellranges.getRangeAddresses()) + 1 if len(cellranges) else karute.splittedrow
								endrowbelow = emptyrow + len(newdatarows)	
								sheet.insertCells(sheet[emptyrow:endrowbelow, :].getRangeAddress(), insert_rows)  # ダブルクリックした行の下に空行を挿入。	
								sheet[emptyrow:endrowbelow, :].setPropertyValues(("CellBackColor", "CharColor"), (-1, -1))  # 追加行の背景色と文字色をクリア。	
								sheet[emptyrow:endrowbelow, sharpcolumn:sharpcolumn+len(newdatarows[0])].setDataArray(newdatarows)
								createFormatKey = commons.formatkeyCreator(doc)	
								sheet[emptyrow:endrowbelow, karute.datecolumn].setPropertyValues(("NumberFormat", "HoriJustify", "VertJustify"), (createFormatKey('YYYY/MM/DD'), LEFT, CellVertJustify2.CENTER))  # カルテシートの入院日の書式設定。左寄せにする。
								datarange.clearContents(CellFlags.STRING)  # コピー元の文字列をクリア。
						return False  # セルを編集モードにしない。
					elif c==karute.sharpcolumn:  # #列の時。
						if txt:
							target.clearContents(CellFlags.STRING)
						else:
							target.setString("#")
						return False  # セルを編集モードにしない。
					elif c==karute.datecolumn:  # Date列の時。
						if not txt:  # 空文字の時。
							sheet[r, c:c+2].setDataArray((("#", ""),))  # Date列と日付列に値を代入。
							target.setPropertyValues(("HoriJustify", "VertJustify"), (RIGHT, CellVertJustify2.CENTER))
							return False  # セルを編集モードにしない。
						elif txt=="#":
							datevalue = int(functionaccess.callFunction("DATEVALUE", (date.today().isoformat(),)))  # 今日のシリアル値を整数で取得。floatで返る。						
							sheet[r, c:c+2].setDataArray(([datevalue]*2,))  # Date列と日付列に今日のシリアル値を代入。
							createFormatKey = commons.formatkeyCreator(doc)	
							target.setPropertyValues(("NumberFormat", "HoriJustify", "VertJustify"), (createFormatKey('YYYY/MM/DD'), LEFT, CellVertJustify2.CENTER))  # カルテシートの入院日の書式設定。左寄せにする。
							sheet[r, c+1].setPropertyValue("CharColor", commons.COLORS["white"])
							return False  # セルを編集モードにしない。
					elif c==karute.datecolumn+1:  # 日付列の時。
						datevalue = int(functionaccess.callFunction("DATEVALUE", (date.today().isoformat(),)))  # 今日のシリアル値を整数で取得。floatで返る。
						if txt:  # 日付列は日付シリアル値しか入っていないはず。
							celldatevalue = int(target.getValue())  # セルに入っているシリアル値を整数で取得。
							if celldatevalue>datevalue-2:  # 2日前までは1日遡る。
								datevalue = celldatevalue - 1	
							else:
								datevalue = ""
						sheet[r, c-1:c+1].setDataArray(([datevalue]*2,))  # Date列と日付列に今日のシリアル値を代入。
						createFormatKey = commons.formatkeyCreator(doc)	
						sheet[r, c-1].setPropertyValues(("NumberFormat", "HoriJustify", "VertJustify"), (createFormatKey('YYYY/MM/DD'), LEFT, CellVertJustify2.CENTER))  # カルテシートの入院日の書式設定。左寄せにする。
						target.setPropertyValue("CharColor", commons.COLORS["white"])					
						return False  # セルを編集モードにしない。
					elif c==karute.subjectcolumn:  # Subject列の時。
						if not txt:  # 空文字の時。
							target.setString("#")
							target.setPropertyValues(("HoriJustify", "VertJustify"), (RIGHT, CellVertJustify2.CENTER))
							return False  # セルを編集モードにしない。		
						elif txt=="#":
							target.setString("")
							target.setPropertyValues(("HoriJustify", "VertJustify"), (LEFT, CellVertJustify2.CENTER))
					elif c==karute.subjectcolumn+1:  # S履歴列の時。
						
						
						return False  # セルを編集モードにしない。	
					elif c>karute.articlecolumn:  # Article列の右列の時。
						transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
						transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))  # 全角文字を半角にする。						
						dateformat = "%Y/%-m/%-d"  # Article列にいれる日付書式。0で埋めない。
						kijirange = sheet[r, karute.articlecolumn:karute.articlecolumn+2]  # Article列と過去日列のみ取得。
						articletxt, datetxt = kijirange.getDataArray()[0]  # Articleセルと挿入済日付セルの値を取得。
						articletxt = articletxt and transliteration.transliterate(articletxt, 0, len(articletxt), [])[0]  # 半角に変換。
						newdateobj = date.today()  # 今日の日付オブジェクトをまず取得。						
						if datetxt:  # 日付が挿入済の時。
							if articletxt.endswith(datetxt):  # Article列の最後がこのボタンで入れた日付で終わっている時。
								dateobj = datetime.strptime(datetxt, dateformat.replace("-", "")).date()  # 日時を取得。0で埋めない-があるとValueError: '-' is a bad directiveがでる。
								articletxt = articletxt[:-len(datetxt)]  # すでにある日付を削る。
								if c==karute.articlecolumn+1 and dateobj>newdateobj-timedelta(days=2):  # 過去日列かつ2日前までの時。
									newdateobj = dateobj - timedelta(days=1)  # 1日遡る。
								elif c==karute.articlecolumn+2 and dateobj<newdateobj+timedelta(days=2):  # 未来日列かつ2日後までの時。	
									newdateobj = dateobj + timedelta(days=1)  # 1日進める。
								elif c==karute.articlecolumn+3:  # 入替列の時。
									txts = articletxt.rsplit("｡", 1)  # 右から｡で分割。	
									if len(txts)>1:  # ｡の後ろに日付を移動させる。
										if txts[-1]:  # 日付の直前が｡でない時。
											articletxt = "".join((txts[0], "｡", datetxt, txts[1]))
											kijirange.setDataArray(((articletxt, ""),))
										else:  # 日付の直前が｡の時。txts[-1]は空文字になる。
											txts2 = txts[0].rsplit("｡", 1)  # 右から｡で再分割。	
											if len(txts2)>1:  # ｡の後ろに日付を移動させる。
												articletxt = "".join((txts2[0], "｡", datetxt, txts2[1], "｡"))
												kijirange.setDataArray(((articletxt, ""),))
									return False  # セルを編集モードにしない。		
								else:  # 日付を削除して空文字にする。
									kijirange.setDataArray(((articletxt, ""),))
									return False  # セルを編集モードにしない。		
						datetxt = newdateobj.strftime(dateformat)							
						articletxt += datetxt
						kijirange.setDataArray(((articletxt, datetxt),))						
						sheet[r, karute.articlecolumn+1].setPropertyValue("CharColor", commons.COLORS["white"])
						return False  # セルを編集モードにしない。
	return True  # セル編集モードにする。
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
def createCopyFuncs(ctx, smgr, doc, sheet):  # コピーのための関数を返す関数。
	karute = Karute(sheet)  # クラスをインスタンス化。	
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
			newrange = sheet[datarangestartrow:datarangestartrow+len(newarticlerows), karute.articlecolumn]  # 記事列のセル範囲を取得。
			newrange.clearContents(CellFlags.STRING+CellFlags.VALUE)  # 記事列の文字列と数値をクリア。
			newrange.setDataArray(newarticlerows)  # 記事列に代入。
		return c  # 追加した行数を返す。
	def formatProblemList(startrow, endrow, title):  # プロブレム欄を整形。
		c = formatArticleColumn(sheet[startrow:endrow, karute.sharpcolumn:karute.articlecolumn+1])  # プロブレム欄の記事列を整形。追加した行数が返る。
		datarows = sheet[startrow:endrow+c, karute.sharpcolumn:karute.articlecolumn+1].getDataArray()  # 文字数制限後の行のタプルを取得。
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
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
	del contextmenu[:]  # contextmenu.clear()は不可。
	target = controller.getSelection()  # 現在選択しているセル範囲を取得。
	if contextmenuname=="cell":  # セルのとき
		karute = getSectionName(sheet, target)  # セル固有の定数を取得。
		sectionname = karute.sectionname  # クリックしたセルの区画名を取得。			
		if sectionname in ("A", "B"):  # 固定行より上の時はコンテクストメニューを表示しない。
			return EXECUTE_MODIFIED
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Delete"})	
# 		if target.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
# 			addMenuentry("ActionTrigger", {"Text": "To Green", "CommandURL": baseurl.format("entry1")}) 
# 		elif target.supportsService("com.sun.star.sheet.SheetCellRange"):  # 連続した複数セルの時。
# 			addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format("entry2")}) 
	elif contextmenuname=="rowheader":  # 行ヘッダーのとき。
		karute = getSectionName(sheet, target[0, 0])  # 選択範囲の最初のセルの定数を取得。
		sectionname = karute.sectionname  # クリックしたセルの区画名を取得。			
		if sectionname in ("A",) or target[0, 0].getPropertyValue("CellBackColor")!=-1:  # 背景色のあるときは表示しない。
			return EXECUTE_MODIFIED
		if sectionname in ("C",):
			addMenuentry("ActionTrigger", {"Text": "過去ﾘｽﾄへ移動", "CommandURL": baseurl.format("entry2")})  
			addMenuentry("ActionTrigger", {"Text": "過去ﾘｽﾄにｺﾋﾟｰ", "CommandURL": baseurl.format("entry3")}) 
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			addMenuentry("ActionTrigger", {"Text": "最下行へ", "CommandURL": baseurl.format("entry1")})   
		elif sectionname in ("G",):
			addMenuentry("ActionTrigger", {"Text": "現ﾘｽﾄへ移動", "CommandURL": baseurl.format("entry4")})  
			addMenuentry("ActionTrigger", {"Text": "現ﾘｽﾄにｺﾋﾟｰ", "CommandURL": baseurl.format("entry5")})  
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
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	selection = controller.getSelection()
	if len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 列全体が選択されている場合もあるので行全体が選択されていることを確認する。
		karute = getSectionName(sheet, selection[0, 0])
		splittedrow = karute.splittedrow
		bluerow = karute.bluerow
		skybluerow = karute.skybluerow
		redrow = karute.redrow
		sharpcolumn = karute.sharpcolumn
		articlecolumn = karute.articlecolumn
		if entrynum==1:  # 現リストの最下行へ。青行の上に移動する。セクションC。
			dest_start_ridx = bluerow  # 移動先開始行インデックス。	
			problemranges = getProblemRanges(doc, sheet[splittedrow:bluerow, sharpcolumn:articlecolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。			
		elif entrynum==2:  # 過去ﾘｽﾄへ移動。スカイブルー行の下に移動する。セクションC。
			dest_start_ridx = skybluerow + 1  # 移動先開始行インデックス。
			problemranges = getProblemRanges(doc, sheet[splittedrow:bluerow, sharpcolumn:articlecolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。		
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。					
		elif entrynum==3:  # 過去ﾘｽﾄにｺﾋﾟｰ。スカイブルー行の下にコピーする。
			dest_start_ridx = skybluerow + 1  # 移動先開始行インデックス。
			problemranges = getProblemRanges(doc, sheet[splittedrow:bluerow, sharpcolumn:articlecolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。		
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.copyRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
		elif entrynum==4:  # 現ﾘｽﾄへ移動。青行の上に移動する。
			dest_start_ridx = bluerow  # 移動先開始行インデックス。	
			problemranges = getProblemRanges(doc, sheet[skybluerow+1:redrow, sharpcolumn:articlecolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.moveRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
				sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動した問題リストの行を削除。			
		elif entrynum==5:  # 現ﾘｽﾄにｺﾋﾟｰ。青行の上にコピーする。
			dest_start_ridx = bluerow  # 移動先開始行インデックス。	
			problemranges = getProblemRanges(doc, sheet[skybluerow+1:redrow, sharpcolumn:articlecolumn+1], selection)  # 問題ごとのセル範囲コレクションを取得。
			for i in problemranges:  # 各セル範囲について。移動や挿入したセル範囲は逐次インデックスで取得する。
				sourcerangeaddress = moveProblems(sheet, i, dest_start_ridx)  # 問題リストを移動させる。
				sheet.copyRange(sheet[dest_start_ridx, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。
def moveProblems(sheet, problemrange, dest_start_ridx):  # problemrange; 問題リストの塊。dest_start_ridx: 移動先開始行インデックス。
	dest_endbelow_ridx = dest_start_ridx + len(problemrange.getRows())  # 移動先最終行の次の行インデックス。
	dest_rangeaddress = sheet[dest_start_ridx:dest_endbelow_ridx, :].getRangeAddress()  # 挿入前にセル範囲アドレスを取得しておく。
	sheet.insertCells(dest_rangeaddress, insert_rows)  # 空行を挿入。	
	sheet.queryIntersection(dest_rangeaddress).clearContents(511)  # 挿入した行の内容をすべてを削除。挿入セルは挿入した行の上のプロパティを引き継いでいるのでリセットしないといけない。
	cursor = sheet.createCursorByRange(problemrange)  # セルカーサーを作成		
	cursor.expandToEntireRows()  # セル範囲を行全体に拡大。		
	return cursor.getRangeAddress()  # 移動前に移動元の問題リストのセル範囲アドレスを取得しておく。
def drowBorders(controller, sheet, cellrange, borders):  # cellrangeを交点とする行列全体の外枠線を描く。
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = borders  # 枠線を取得。	
	cell = cellrange[0, 0]  # セル範囲の左上端のセルで判断する。
	rangeaddress = cell.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	karute = getSectionName(sheet, cell)  # セル固有の定数を取得。
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
