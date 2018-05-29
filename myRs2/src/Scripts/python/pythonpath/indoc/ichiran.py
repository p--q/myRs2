#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 一覧シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import commons, keika, karute
from itertools import chain, combinations
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.awt import MouseButton, MessageBoxButtons, MessageBoxResults # 定数
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.awt.MessageBoxType import QUERYBOX, ERRORBOX  # enum
from com.sun.star.i18n.TransliterationModulesNew import HALFWIDTH_FULLWIDTH, FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.table.CellHoriJustify import LEFT  # enum
class Ichiran():  # シート固有の定数設定。
	def __init__(self):
		self.menurow  = 0  # メニュー行インデックス。
		self.splittedrow = 2  # 分割行インデックス。
		self.sumicolumn = 0  # 済列インデックス。
		self.idcolumn = 2  # ID列インデックス。	
		self.kanacolumn = 4  # カナ列インデックス。	
		self.datecolumn = 5  # 入院日列インデックス。
		self.checkstartcolumn = 8  # チェック列開始列インデックス。
		self.memostartcolumn = 22  # メモ列開始列インデックス。
def getSectionName(sheet, target):  # 区画名を取得。
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
	ichiran = Ichiran()  # クラスをインスタンス化。	
	splittedrow = ichiran.splittedrow
	checkstartcolumn = ichiran.checkstartcolumn
	memostartcolumn = ichiran.memostartcolumn
	rangeaddress = target.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	cellranges = sheet[:, ichiran.idcolumn].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # ID列の文字列が入っているセルに限定して抽出。数値の時もありうる。
	emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # ID列の最終行インデックス+1を取得。
	if len(sheet[ichiran.menurow, :checkstartcolumn].queryIntersection(rangeaddress)):  # メニューセルの時。
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
	ichiran.sectionname = sectionname   # 区画名
	ichiran.emptyrow = emptyrow  # 最終行インデックス+1を取得。
	return ichiran
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	sheet["C1:F1"].setDataArray((("済をﾘｾｯﾄ", "検予を反映", "予をﾘｾｯﾄ", "入力支援"),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	target = enhancedmouseevent.Target  # ターゲットのセルを取得。
	sheet = target.getSpreadsheet()
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				drowBorders(sheet, target, commons.createBorders())  # 枠線の作成。
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
				systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
				ichiran = getSectionName(sheet, target)
				sectionname, splittedrow, emptyrow, sumicolumn, checkstartcolumn, memostartcolumn\
					= ichiran.sectionname, ichiran.splittedrow, ichiran.emptyrow, ichiran.sumicolumn, ichiran.checkstartcolumn, ichiran.memostartcolumn
				celladdress = target.getCellAddress()
				r, c = celladdress.Row, celladdress.Column  # targetの行と列のインデックスを取得。		
				txt = target.getString()  # クリックしたセルの文字列を取得。		
				if sectionname=="M":
					if txt=="検予を反映":  # 経過シートから本日の検予を取得。
						cellranges = sheet[splittedrow:, ichiran.idcolumn].queryContentCells(CellFlags.STRING)  # ID列に文字列が入っているセルを取得。
						headerrow = sheet[ichiran.menurow, checkstartcolumn:memostartcolumn].getDataArray()[0]  # チェック列のヘッダーのタプルを取得。
						eketsucol, dokueicol, ketuekicol, gazocol, shochicol, echocol, ecgcol\
							= [headerrow.index(i) for i in ("ｴ結", "読影", "血液", "画像", "処置", "ｴｺ", "ECG")]  # headerrowタプルでのインデックスを取得。
						functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
						keikaconsts = keika.Keika()  # 経過シートの定数を取得。
						daterow = keikaconsts.daterow  # 経過シートの日付行インデックスを取得。
						splittedcolumn = keikaconsts.splittedcolumn  # 日付列の最初の列インデックスを取得。
						todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。
						c = splittedcolumn + todayvalue  # 分割列と今日の日付のシリアル値の和。
						if len(cellranges)>0:  # ID列のセル範囲が取得出来ている時。
							sheets = doc.getSheets()  # シートコレクションを取得。
							iddatarows = cellranges[0].getDataArray()  # ID列のデータ行のタプルを取得。空行がないとする。
							checkrange = sheet[splittedrow:splittedrow+len(iddatarows), checkstartcolumn:memostartcolumn]  # チェック列範囲を取得。
							datarows = list(map(list, checkrange.getDataArray()))  # 各行をリストにして取得。
							for r, idtxt in enumerate(chain.from_iterable(iddatarows)):  # 各ID列について。rは相対インデックス。
								if idtxt.isdigit():  # IDがすべて数字の時。
									sheetname = "{}経".format(idtxt)  # 経過シート名を作成。
									if not sheetname in sheets:  # 経過シートがない時は次のループに行く。
										continue
									keikasheet = sheets[sheetname]  # 経過シートを取得。
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
						containerwindow = controller.getFrame().getContainerWindow()  # コンテナウィンドウを取得。
						toolkit = containerwindow.getToolkit() # ウィンドウピアオブジェクトからツールキットを取得。
						msgbox = toolkit.createMessageBox(containerwindow, QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_OK, "済列の変更", "済をリセットしますか？")
						if msgbox.execute()==MessageBoxResults.OK:
							sheet[splittedrow:emptyrow, :].setPropertyValue("CharColor", commons.COLORS["black"])  # 文字色をリセット。
							sheet[splittedrow:emptyrow, sumicolumn].setDataArray([("未",)]*(emptyrow-splittedrow))  # 済列をリセット。
							searchdescriptor = sheet.createSearchDescriptor()
							searchdescriptor.setSearchString("済")
							cellranges = sheet[splittedrow:emptyrow, checkstartcolumn:memostartcolumn].findAll(searchdescriptor)  # チェック列の「済」が入っているセル範囲コレクションを取得。
							cellranges.setPropertyValue("CharColor", commons.COLORS["silver"])
					elif txt=="予をﾘｾｯﾄ":
						sheet[splittedrow:emptyrow, sumicolumn+1].clearContents(CellFlags.STRING)  # 予列をリセット。
					elif txt=="入力支援":
						
						
						
						
						pass  # 入力支援odsを開く。
					
					return False  # セル編集モードにしない。
				elif not target.getPropertyValue("CellBackColor") in (-1, commons.COLORS["cyan10"]):  # 背景色がないか薄緑色でない時。何もしない。
					return False  # セル編集モードにしない。
				elif sectionname=="B":
					header = sheet[splittedrow-1, c].getString()  # 固定行の最下端のセルの文字列を取得。
					sheets = doc.getSheets()  # シートコレクションを取得。
					if header=="済":
						if txt=="未":
							target.setString("待")
							sheet[r, :].setPropertyValue("CharColor", commons.COLORS["skyblue"])
						elif txt=="待":
							target.setString("済")
							sheet[r, :].setPropertyValue("CharColor", commons.COLORS["silver"])
							doc.store()  # ドキュメントを保存する。
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
						datarange = sheet[r, :checkstartcolumn]
						datarow = list(datarange.getDataArray()[0])
						if not datarow[-1]:  # 在院日数列に値がないときは未設定行と判断する。式が入っていても値がなければNoneが返る。
							if all(datarow[ichiran.idcolumn:ichiran.datecolumn+1]):  # ID、漢字名、カナ名、入院日、すべてが揃っている時。
								datarow = "未", "", *datarow[ichiran.idcolumn:ichiran.datecolumn+1], "経過", ""
								datarange.setDataArray((datarow,))
								createFormatKey = commons.formatkeyCreator(doc)	
								sheet[r, ichiran.idcolumn].setPropertyValue("NumberFormat", createFormatKey('@'))  # ID列の書式を文字列に設定。 	
								sheet[r, ichiran.datecolumn].setPropertyValue("NumberFormat", createFormatKey('YY/MM/DD'))
								cellstringaddress = sheet[r, ichiran.datecolumn].getPropertyValue("AbsoluteName").split(".")[-1].replace("$", "")  # 入院日セルの文字列アドレスを取得。
								sheet[r, ichiran.checkstartcolumn-1].setFormula("=TODAY()+1-{}".format(cellstringaddress))  #  在院日数列に式を代入。			
								sheet[r, ichiran.checkstartcolumn-1].setPropertyValue("NumberFormat", createFormatKey('0" ";[RED]-0" "'))  # 在院日数列の書式を設定。 	
							else:
								msg = "ID、漢字名、カナ名、入院日\nすべてを入力してください。"
								componentwindow = controller.ComponentWindow
								msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, ERRORBOX, MessageBoxButtons.BUTTONS_OK, "myRs", msg)
								msgbox.execute()	
								return
						idtxt = datarow[ichiran.idcolumn]
						if idtxt in sheets:  # すでにカルテシートが存在するときはそれをアクティブにする。
							controller.setActiveSheet(sheets[idtxt])
						else:  # カルテシートがない時。					
							sheets.copyByName("00000000", idtxt, len(sheets))  # テンプレートシートをコピーしてID名のシートにして最後に挿入。
							newsheet = sheets[idtxt]  # カルテシートを取得。  
							if createFormatKey is None:
								createFormatKey = commons.formatkeyCreator(doc)	
							karuteconsts = karute.Karute(newsheet)	
							karutesplittedrow = karuteconsts.splittedrow
							newsheet[karutesplittedrow, karuteconsts.datecolumn].setValue(datarow[ichiran.datecolumn])  # カルテシートに入院日を入力。
							newsheet[karutesplittedrow, karuteconsts.datecolumn].setPropertyValues(("NumberFormat", "HoriJustify"), (createFormatKey('YYYY/MM/DD'), LEFT))  # カルテシートの入院日の書式設定。左寄せにする。
							newsheet[:karutesplittedrow, karuteconsts.articlecolumn].setDataArray(("",), (" ".join(datarow[ichiran.idcolumn:ichiran.kanacolumn+1]),))  # カルテシートのコピー日時をクリア。ID名前を入力。
							controller.setActiveSheet(newsheet)  # カルテシートをアクティブにする。
					elif header=="ｶﾅ名":
						datarow = sheet[r, ichiran.idcolumn:ichiran.datecolumn].getDataArray()[0]  # ID、漢字名、ｶﾅ名、を取得。
						transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。
						transliteration.loadModuleNew((HALFWIDTH_FULLWIDTH,), Locale(Language = "ja", Country = "JP"))
						kana = datarow[2].replace(" ", "")  # 半角空白を除去してカナ名を取得。
						zenkana = transliteration.transliterate(kana, 0, len(kana), [])[0]  # ｶﾅを全角に変換。
						systemclipboard.setContents(commons.TextTransferable("".join((zenkana, datarow[0]))), None)  # クリップボードにカナ名+IDをコピーする。	
					elif header=="入院日":
						if txt:  # すでに入力されている時。
							return True  # セル編集モードにする。
						else:  # まだ空欄の時。
							functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
							todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。
							target.setValue(todayvalue)
							target.setPropertyValue("NumberFormat", commons.formatkeyCreator(doc)('YY/MM/DD'))
					elif txt=="経過":  # このボタンはカルテシートの作成時に作成されるのでカルテシート作成後のみ有効。
						ids = list(sheet[r, ichiran.idcolumn:ichiran.datecolumn].getDataArray()[0])  # ダブルクリックした行をID列からｶﾅ名列までのタプルを取得。						
						newsheetname = "".join([ids[0], "経"])  # 経過シート名を取得。
						if newsheetname in sheets:  # 経過シートがなければ作成する。
							controller.setActiveSheet(sheets[newsheetname])  # 経過シートをアクティブにする。
						else:  # 経過シートがなければ作成する。
							dateserial = int(sheet[r, ichiran.datecolumn].getValue())  # 入院日の日時シリアル値を取得。		
							sheets.copyByName("00000000経", newsheetname, len(sheets))  # テンプレートシートをコピーしてID経名のシートにして最後に挿入。	
							keikasheet = sheets[newsheetname]  # 新規経過シートを取得。
							keikasheet["F2"].setString(" ".join(ids))  # ID漢字名ｶﾅ名を入力。					
							keika.setDates(doc, keikasheet, keikasheet["I2"], dateserial)  # 経過シートの日付を設定。
							controller.setActiveSheet(keikasheet)  # 経過シートをアクティブにする。						
						
						
						
# 						ids = list(sheet[r, ichiran.idcolumn:ichiran.datecolumn].getDataArray()[0])  # ダブルクリックした行をID列からｶﾅ名列までのタプルを取得。						
# 						newsheetname = "".join([ids[0], "経"])  # 経過シート名を取得。
# 						if newsheetname in sheets:  # 経過シートがなければ作成する。
# 							controller.setActiveSheet(sheets[newsheetname])  # 経過シートをアクティブにする。
# 						else:  # 経過シートがなければ作成する。
# 							dateserial = int(sheet[r, ichiran.datecolumn].getValue())  # 入院日の日時シリアル値を取得。		
# 							sheets.copyByName("00000000経", newsheetname, len(sheets))  # テンプレートシートをコピーしてID経名のシートにして最後に挿入。	
# 							keikasheet = sheets[newsheetname]  # 新規経過シートを取得。
# 							keikasheet["F2"].setString(" ".join(ids))  # ID漢字名ｶﾅ名を入力。					
# 							keika.setDates(doc, keikasheet, keikasheet["I2"], dateserial)  # 経過シートの日付を設定。
# 							controller.setActiveSheet(keikasheet)  # 経過シートをアクティブにする。
					return False  # セル編集モードにしない。		
				elif sectionname=="D":
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
					header = sheet[ichiran.menurow, c].getString()  # 行インデックス0のセルの文字列を取得。
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
					target.setString(newtxt)
					color = commons.COLORS["silver"] if "済" in newtxt else -1
					target.setPropertyValue("CharColor", color)			
					return False  # セル編集モードにしない。
				elif sectionname=="A":
					if sheet[splittedrow-1, c].getString()=="ｶﾅ名":  # 固定行の最下端のセルの文字列を取得。
						
						pass  # 漢字名からｶﾅを取得する。

	return True  # セル編集モードにする。	
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
def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。		
	changes = changesevent.Changes	
	for change in changes:
		if change.Accessor=="cell-change":  # セルの値が変化した時。
			cell = change.ReplacedElement  # 値を変更したセルを取得。		
			ichiran = Ichiran()  # 一覧シート固有の定数を取得。
			celladdress = cell.getCellAddress()
			r, c = celladdress.Row, celladdress.Column
			if r>ichiran.splittedrow-1:  # 分割行以降の時。
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。					
				transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
				transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))					
				if c==ichiran.idcolumn:  # ID列の時。
					txt = cell.getString()  # セルの文字列を取得。
					txt = transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換。
					if txt.isdigit():  # 数値の時のみ。空文字の時0で埋まってしまう。
						cell.setString("{:0>8}".format(txt))  # 数値を8桁にして文字列として代入し直す。
				elif c==ichiran.kanacolumn:  # カナ列の時。
					txt = cell.getString()  # セルの文字列を取得。
					cell.setString(transliteration.transliterate(txt, 0, len(txt), [])[0])  # 半角に変換。	
				elif c==ichiran.datecolumn:  # 日付列の時。
					doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
					cell.setPropertyValues(("NumberFormat", "HoriJustify"), (commons.formatkeyCreator(doc)('YYYY/MM/DD'), LEFT))  # カルテシートの入院日の書式設定。左寄せにする。
			break
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。				
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。	
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


def drowBorders(sheet, cellrange, borders):  # ターゲットを交点とする行列全体の外枠線を描く。
	cell = cellrange[0, 0]  # セル範囲の左上端のセルで判断する。
	ichiran = getSectionName(sheet, cell)
	sectionname = ichiran.sectionname	
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