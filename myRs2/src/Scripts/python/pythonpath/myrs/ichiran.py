#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 一覧シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
# import calendar
from myrs import commons
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.awt import MouseButton, MessageBoxButtons, MessageBoxResults # 定数
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.awt.MessageBoxType import QUERYBOX, ERRORBOX  # enum
from com.sun.star.i18n.TransliterationModulesNew import HALFWIDTH_FULLWIDTH, FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.table.CellHoriJustify import LEFT  # enum
class Ichiran():  # シート固有の定数設定。
	pass
def getSectionName(controller, sheet, cell):  # 区画名を取得。
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
	B: スクロールする部分のうちヘッダが結合セルである列より左の部分。
	D: スクロールする部分のうちヘッダが結合セルである部分。
	E: スクロールする部分のうちヘッダが結合セルである列より右の部分。
	A: ID列の最初の空行から下の部分。
	"""
	menurow  = 0  # メニュー行インデックス。
	idcolumn = 2  # ID列インデックス。
	startrow = controller[0].getVisibleRange().EndRow + 1  # スクロールする枠の最初の行インデックス。
	emptycellranges = sheet[startrow-1, :].queryEmptyCells()  # 上枠の最下行の空セルのセル範囲コレクションを取得。
	mergedheaders = emptycellranges[0].getRangeAddress()  # ヘッダの結合セルの範囲を取得。
	dstart, dend = mergedheaders.StartColumn+1, mergedheaders.EndColumn+1  # ヘッダ結合セルの左端列インデックスと右端列インデックス+1の取得。
	rangeaddress = cell.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	cellranges = sheet[:, idcolumn].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # ID列の文字列が入っているセルに限定して抽出。空列は不可。数値の時もありうる。
	emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # ID列の最終行インデックス+1を取得。
	sectionname = "C"  
	if len(sheet[menurow, :dstart-1].queryIntersection(rangeaddress)):  # メニューセルの時。
		sectionname = "M"
	elif len(sheet[startrow:emptyrow, :dstart].queryIntersection(rangeaddress)):  # Dの左。
		sectionname = "B"	
	elif len(sheet[startrow:emptyrow, dstart:dend].queryIntersection(rangeaddress)):  # チェック列の時。
		sectionname = "D"		
	elif len(sheet[startrow:emptyrow, dstart:].queryIntersection(rangeaddress)):  # Dの右。
		sectionname = "E"		
	elif len(sheet[emptyrow:, :].queryIntersection(rangeaddress)):  # まだデータのない行の時。
		sectionname = "A"	
	ichiran = Ichiran()  # クラスをインスタンス化。	
	ichiran.sectionname = sectionname   # 区画名
	ichiran.menurow = menurow  # メニュー行インデックス。
	ichiran.startrow = startrow  # 左上枠の最下行のインデックス。	
	ichiran.emptyrow = emptyrow  # 最終行インデックス+1を取得。
	ichiran.sumi_retu = 0  # 済列インデックス。
	ichiran.dstart = dstart  # D左端列。
	ichiran.dend = dend  # D右端列+1
	return ichiran
def selectionChanged(controller, sheet, args):  # 矢印キーでセル移動した時も発火する。
	borders = args	
	selection = controller.getSelection()
	if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択範囲がセルの時。矢印キーでセルを移動した時。マウスクリックハンドラから呼ばれると何回も発火するのでその対応。
		currenttableborder2 = selection.getPropertyValue("TableBorder2")  # 選択セルの枠線を取得。
		if all((currenttableborder2.TopLine.Color==currenttableborder2.LeftLine.Color==commons.COLORS["violet"],\
				currenttableborder2.RightLine.Color==currenttableborder2.BottomLine.Color==commons.COLORS["magenta3"])):  # 枠線の色を確認。
			return  # すでに枠線が書いてあったら何もしない。
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
		drowBorders(controller, sheet, selection, borders)	
def activeSpreadsheetChanged(sheet):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。
	sheet["C1:F1"].setDataArray((("済をﾘｾｯﾄ", "検予を反映", "予をﾘｾｯﾄ", "入力支援"),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
def mousePressed(enhancedmouseevent, controller, sheet, target, args):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	borders, systemclipboard, transliteration = args
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				drowBorders(controller, sheet, target, borders)
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				ichiran = getSectionName(controller, sheet, target)
				section, startrow, emptyrow, sumi_retu, dstart = ichiran.sectionname, ichiran.startrow, ichiran.emptyrow, ichiran.sumi_retu, ichiran.dstart
				celladdress = target.getCellAddress()
				r, c = celladdress.Row, celladdress.Column  # targetの行と列のインデックスを取得。		
				txt = target.getString()  # クリックしたセルの文字列を取得。		
				if section=="M":
					if txt=="検予を反映":
						
						pass  # 経過シートから本日の検予を取得。
					
					elif txt=="済をﾘｾｯﾄ":
						containerwindow = controller.getFrame().getContainerWindow()  # コンテナウィンドウを取得。
						toolkit = containerwindow.getToolkit() # ウィンドウピアオブジェクトからツールキットを取得。
						msgbox = toolkit.createMessageBox(containerwindow, QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_OK, "済列の変更", "済をリセットしますか？")
						if msgbox.execute()==MessageBoxResults.OK:
							sheet[startrow:emptyrow, :].setPropertyValue("CharColor", commons.COLORS["black"])  # 文字色をリセット。
							sheet[startrow:emptyrow, sumi_retu].setDataArray([("未",)]*(emptyrow-startrow))  # 済列をリセット。
							searchdescriptor = sheet.createSearchDescriptor()
							searchdescriptor.setSearchString("済")
							cellranges = sheet[startrow:emptyrow, dstart:ichiran.dend].findAll(searchdescriptor)  # チェック列の「済」が入っているセル範囲コレクションを取得。
							cellranges.setPropertyValue("CharColor", commons.COLORS["silver"])
					elif txt=="予をﾘｾｯﾄ":
						sheet[startrow:emptyrow, sumi_retu+1].clearContents(CellFlags.STRING)  # 予列をリセット。
					elif txt=="入力支援":
						
						pass  # 入力支援odsを開く。
					
					return False  # セル編集モードにしない。
				elif not target.getPropertyValue("CellBackColor") in (-1, commons.COLORS["cyan10"]):  # 背景色がないか薄緑色でない時。何もしない。
					return False  # セル編集モードにしない。
				elif section=="B":
					header = sheet[startrow-1, c].getString()  # 固定行の最下端のセルの文字列を取得。
					doc = controller.getModel()
					sheets = doc.getSheets()  # シートコレクションを取得。
					if header=="済":
						if txt=="未":
							target.setString("待")
							sheet[r, :].setPropertyValue("CharColor", commons.COLORS["skyblue"])
						elif txt=="待":
							target.setString("済")
							sheet[r, :].setPropertyValue("CharColor", commons.COLORS["silver"])
							controller.getModel().store()  # ドキュメントを保存する。
						elif txt=="済":
							target.setString("未")
							sheet[r, :].setPropertyValue("CharColor", commons.COLORS["black"])
					elif header=="予":
						if txt:
							target.clearContents(CellFlags.STRING)  # 予をクリア。
						else:  # セルの文字列が空の時。
							target.setString("予")
					elif header=="ID":
						systemclipboard.setContents(commons.TextTransferable(txt), None)  # クリップボードにIDをコピーする。
					elif header=="漢字名":  # カルテシートをアクティブにする、なければ作成する。カルトシート名はIDと一致。	
						ids = list(sheet[r, 2:dstart].getDataArray()[0])  # ダブルクリックした行を経過列までのタプルをリストにして取得。
						ids[0] = "{:0>8}".format(int(ids[0]))  # IDは常に8桁の数字の文字列にする。全角にはここでは対応しない。数値の時はfloatで返ってくる。
						createFormatKey = None
						if not ids[-1]:  # 在院日数列に値がないときは未設定行と判断する。式が入っていても値がなければNoneが返る。
							if all(ids[:4]):  # ID、漢字名、カナ名、入院日、すべてが揃っている時。
								sheet[r, :2].setDataArray((("未", ""),))  # 未列と予列を設定。
								sheet[r, 6].setString("経過")  # 経過列を設定。
								cellstringaddress = sheet[r, 5].getPropertyValue("AbsoluteName").split(".")[-1].replace("$", "")  # 入院日セルの文字列アドレスを取得。
								sheet[r, 7].setFormula("=TODAY()+1-{}".format(cellstringaddress))  #  在院日数列に式を代入。
								createFormatKey = commons.formatkeyCreator(doc)							
								sheet[r, 7].setPropertyValue("NumberFormat", createFormatKey('0" ";[RED]-0" "'))  # 在院日数列の書式を設定。 	
								transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))
								sheet[r, 2].setPropertyValue("NumberFormat", createFormatKey('@'))  # ID列の書式を文字列に設定。 	
								ids[0] = transliteration.transliterate(ids[0], 0, len(ids[0]), [])[0]  # IDを半角に変換。
								ids[2] = transliteration.transliterate(ids[2], 0, len(ids[2]), [])[0]  # ｶﾅ名の全角を半角に変換
								sheet[r, 2].setString(ids[0])  # 半角にしたIDを代入。
								sheet[r, 4].setString(ids[2])  # 半角にしたｶﾅ名を代入。
								sheet[r, 5].setPropertyValue("NumberFormat", createFormatKey('YY/MM/DD'))
							else:
								msg = "ID、漢字名、カナ名、入院日\nすべてを入力してください。"
								componentwindow = controller.ComponentWindow
								msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, ERRORBOX, MessageBoxButtons.BUTTONS_OK, "myRs", msg)
								msgbox.execute()	
								return
						if ids[0] in sheets:  # すでにカルテシートが存在するときはそれをアクティブにする。
							controller.setActiveSheet(sheets[ids[0]])
						else:  # カルテシートがない時。					
							sheets.copyByName("00000000", ids[0], len(sheets))  # テンプレートシートをコピーしてID名のシートにして最後に挿入。
							newsheet = sheets[ids[0]]  # カルテシートを取得。  
							if createFormatKey is None:
								createFormatKey = commons.formatkeyCreator(doc)									
							newsheet["C3"].setValue(ids[3])  # カルテシートに入院日を入力。
							newsheet["C3"].setPropertyValues(("NumberFormat", "HoriJustify"), (createFormatKey('YYYY/MM/DD'), LEFT))  # カルテシートの入院日の書式設定。左寄せにする。
# 							newsheet["C3"].getColumns().setPropertyValue("OptimalWidth", True)  # 日付列の列幅を最適化する。
							newsheet["G1"].setString("")  # カルテシートのコピー日時をクリア。
							newsheet["G2"].setString(" ".join(ids[:3]))  # カルテシートのID名前を入力。
							controller.setActiveSheet(newsheet)  # カルテシートをアクティブにする。
					elif header=="ｶﾅ名":
						ns = sheet[r, c-2:c+1].getDataArray()  # ID、漢字名、ｶﾅ名、を取得。
						transliteration.loadModuleNew((HALFWIDTH_FULLWIDTH,), Locale(Language = "ja", Country = "JP"))
						kana = ns[0][2].replace(" ", "")  # 半角空白を除去。
						zenkana = transliteration.transliterate(kana, 0, len(kana), [])[0]  # ｶﾅを全角に変換。
						systemclipboard.setContents(commons.TextTransferable("".join((zenkana, ns[0][0]))), None)  # クリップボードにカナ名+IDをコピーする。	
					elif header=="入院日":
						if txt:  # すでに入力されている時。
							return True  # セル編集モードにする。
						else:
# 							dialog, addControl = dialogCreator(ctx, smgr, {"PositionX": 102, "PositionY": 41, "Width": 380, "Height": 380, "Title": "LibreOffice", "Name": "MyTestDialog", "Step": 0, "Moveable": True})  # "TabIndex": 0

							
							
							pass  # カレンダーpicker
					
					
					elif txt=="経過":  # このボタンはカルテシートの作成時に作成されるのでカルテシート作成後のみ有効。

						import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
						
						

						ids = list(sheet[r, 2:5].getDataArray()[0])  # ダブルクリックした行をID列からｶﾅ名列までのタプルを取得。						
						newsheetname = "".join([ids[0], "経"])  # 経過シート名を取得。
						if newsheetname in sheets:  # 経過シートがなければ作成する。
							controller.setActiveSheet(sheets[newsheetname])  # 経過シートをアクティブにする。
						else:  # 経過シートがなければ作成する。
							dateserial = int(sheet[r, 5].getValue())  # 入院日の日時シリアル値を取得。		
							sheets.copyByName("00000000経", newsheetname, len(sheets))  # テンプレートシートをコピーしてID経名のシートにして最後に挿入。							
							newsheet = commons.createKeikaSheet(doc, sheets[newsheetname], ids, dateserial)
							controller.setActiveSheet(newsheet)  # 経過シートをアクティブにする。
							
							
# 							createFormatKey = commons.formatkeyCreator(doc)	
# 							newsheet["F2"].setString(" ".join(ids))  # ID漢字名ｶﾅ名を入力。
# 							daycount = 100  # 経過シートに入力する日数。
# 							celladdress = newsheet["I2"].getCellAddress()  # 経過シートの日付の開始セルのセルアドレスを取得。
# 							r, c = celladdress.Row, celladdress.Column
# 							sheet[:r+1, c:].clearContents(CellFlags.VALUE+CellFlags.DATETIME+CellFlags.STRING+CellFlags.ANNOTATION+CellFlags.FORMULA+CellFlags.HARDATTR+CellFlags.STYLES)  # セルの内容を削除。
# 							endcolumn = c + daycount + 1
# 							sheet[r, c:endcolumn].setDataArray(([i for i in range(dateserial, dateserial+1)],))  # 日時シリアル値を経過シートに入力。
# 							sheet[r, c:endcolumn].setPropertyValue("NumberFormat", createFormatKey('YYYY/M/D'))  # 日時シリアルから年月日の取得のため一時的に2018/5/4の形式に変換する。
# 							y, m, d = sheet[r, c].getString().split("/")  # 年、月、日を文字列で取得。
# 							weekday, days = calendar.monthrange(y, m)  # 日曜日が曜日番号0。1日の曜日と一月の日数のタプルが返る。
# 							weekday = (weekday+(d-1)%7)%7  # dの曜日番号を取得。1日からの7の余りと1日の余りを加えた7の余りがdの曜日番号。
# 							n = 0  # 日曜日の曜日番号。
# 							sundayranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # 日曜日のセル範囲コレクション。
# 							[sundayranges.addRangeAddress(sheet[r, i].getRangeAddress()) for i in range(c+(n-weekday)%7, endcolumn, 7)]  # 曜日番号nの列番号だけについて。
# 							n = 6  # 土曜日の曜日番号。
# 							saturdayranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # 土曜日のセル範囲コレクション。
# 							[saturdayranges.addRangeAddress(sheet[r, i].getRangeAddress()) for i in range(c+(n-weekday)%7, endcolumn, 7)]  # 曜日番号nの列番号だけについて。
# 							holidayranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # 祝日のセル範囲コレクション。
# 							holidays = commons.HOLIDAYS  # 祝日の辞書を取得。
# 							days = days - d + 1  # 翌月1日までの日数を取得。
# 							mr = r - 1  # 月を代入する行のインデックス。
# 							mc = c  # 1日を表示する列のインデックス。
# 							if y in holidays:  # 祝日一覧のキーがある時。
# 								[holidayranges.addRangeAddress(sheet[r, mc+i-1].getRangeAddress()) for i in holidays[y][m] if not i<d]
# 							while True:
# 								sheet[mr, mc].setString("{}月".format(m))  # 月を入力。
# 								mc += days  # 次月1日の列に進める。
# 								if mc<endcolumn:  # 日時シリアル値が入力されている列の時。
# 									ymd = sheet[r, mc].getString()  # 1日の年/月/日を取得。
# 									y, m = ymd.split("/")[:2]  # 年と月を取得。
# 									if y in holidays:  # 祝日一覧のキーがある時。。
# 										[holidayranges.addRangeAddress(sheet[r, mc+i-1].getRangeAddress()) for i in holidays[y][m] if mc+i-1<endcolumn]
# 									weekday, days = calendar.monthrange(y, m)  # 1日の曜日と月の日数を取得。
# 								else:
# 									break
# 							sheet[r, c:endcolumn].setPropertyValue("NumberFormat", createFormatKey('D'))  # 経過シートの日付の書式を日だけにする。
# 							colors = commons.COLORS
# 							holidayranges.setPropertyValue("CellBackColor", colors["red3"])  # 祝日の背景色を変更。
# 							sundayranges.setPropertyValue("CharColor", colors["red3"])  # 日曜日の文字色を変更。
# 							saturdayranges.setPropertyValue("CharColor", colors["skyblue"])  # 土曜日の文字色を変更。	
# 							controller.setActiveSheet(newsheet)  # 経過シートをアクティブにする。
							
							
					return False  # セル編集モードにしない。		
				elif section=="D":
					header = sheet[ichiran.menurow, c].getString()  # 行インデックス0のセルの文字列を取得。
					if header=="4F":
						pass
					elif header=="血液":
						pass						



					return False  # セル編集モードにしない。
				elif section=="A":
					if sheet[startrow-1, c].getString()=="ｶﾅ名":  # 固定行の最下端のセルの文字列を取得。
						
						pass  # 漢字名からｶﾅを取得する。

	return True  # セル編集モードにする。
def drowBorders(controller, sheet, cellrange, borders):  # ターゲットを交点とする行列全体の外枠線を描く。
	cell = cellrange[0, 0]  # セル範囲の左上端のセルで判断する。
	ichiran = getSectionName(controller, sheet, cell)
	sectionname = ichiran.sectionname
	if sectionname in ("A", "B", "D", "E"):
		noneline, tableborder2, topbottomtableborder, leftrighttableborder = borders	
		sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
		if cell.getPropertyValue("CellBackColor") in (-1, commons.COLORS["cyan10"]):  # 背景色がないか薄緑色の時。
			rangeaddress = cellrange.getRangeAddress()  # セル範囲アドレスを取得。
			if sectionname=="D":
				sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。			
			sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。	
			cellrange.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	
def notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, contextmenuname):  # 右クリックメニュー。			
	if contextmenuname=="cell":  # セルのとき
		selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
# 		del contextmenu[:]  # contextmenu.clear()は不可。
# 		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
# 		karute.rng	addMenuentry("ActionTrigger", {"Text": "To blue", "CommandURL": baseurl.format("entry1")})  # listeners.pyの関数名を指定する。
# 		elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 連続した複数セルの時。
# 			addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format("entry2")})  # listeners.pyの関数名を指定する。
# 		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
# 	elif contextmenuname=="rowheader":  # 行ヘッダーのとき。
# 		del contextmenu[:]  # contextmenu.clear()は不可。
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
# 		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsBefore"})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteRows"}) 
# 	elif contextmenuname=="colheader":  # 列ヘッダーの時。
# 		pass  # contextmenuを操作しないとすべての項目が表示されない。
# 	elif contextmenuname=="sheettab":  # シートタブの時。
# 		del contextmenu[:]  # contextmenu.clear()は不可。
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
def contextMenuEntries(target, entrynum):  # コンテクストメニュー番号の処理を振り分ける。
	colors = commons.COLORS
	if entrynum==1:
		target.setPropertyValue("CellBackColor", colors["blue3"])  # 背景を青色にする。
	elif entrynum==2:
		target.setPropertyValue("CellBackColor", colors["red3"]) 
