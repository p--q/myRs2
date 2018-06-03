#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
from indoc import commons
# from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.awt import MouseButton

class Ent():  # シート固有の定数設定。
	def __init__(self, sheet):
		self.menurow  = 0  # メニュー行インデックス。
		self.splittedrow = 1  # 分割行インデックス。
		self.idcolumn = 0  # ID列インデックス。	
		self.kanjicolumn = 1  # 漢字列インデックス。
		self.kanacolumn = 2  # カナ列インデックス。	
		self.datecolumn = 3  # 入院日列インデックス。
		self.cleardatecolumn = 4  # リスト消去日列インデックス。
		cellranges = sheet[:, self.idcolumn].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # ID列の文字列が入っているセルに限定して抽出。数値の時もありうる。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # ID列の最終行インデックス+1を取得。
def getSectionName(sheet, target):  # 区画名を取得。
	"""
	M  
	===========  # 行の固定の境界
	B  
	-----------
	A  # ID列が空欄の行。
	
	M: メニュー行。
	B: スクロールする部分のうちID欄が空欄でない行。
	A: ID列の最初の空行から下の部分。
	"""
	ent = Ent(sheet)  # クラスをインスタンス化。	
	rangeaddress = target.getRangeAddress()  # ターゲットのセル範囲アドレスを取得。セルアドレスは不可。
	if len(sheet[ent.menurow, :].queryIntersection(rangeaddress)):  # メニューセルの時。
		sectionname = "M"
	elif len(sheet[ent.splittedrow:ent.emptyrow, :].queryIntersection(rangeaddress)):  # スクロールする部分のうちID欄が空欄でない行。
		sectionname = "B"	
	else:  # ID列の最初の空行から下の部分。
		sectionname = "A"  
	ent.sectionname = sectionname   # 区画名
	return ent
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	target = enhancedmouseevent.Target  # ターゲットのセルを取得。
	sheet = target.getSpreadsheet()
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				drowBorders(sheet, target, commons.createBorders())  # 枠線の作成。
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
				doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
				functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
				ent = getSectionName(sheet, target)
				sectionname	= ent.sectionname	
				if sectionname=="M":
					return mousePressedWSectionM(doc, sheet, functionaccess, ent, target)			
				elif sectionname=="B":
					systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
					transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
					return mousePressedWSectionB(doc, sheet, systemclipboard, functionaccess, transliteration, ichiran, target)
				elif sectionname=="A":
					pass
	return True  # セル編集モードにする。	
def mousePressedWSectionM(doc, sheet, functionaccess, ent, target):
	controller = doc.getCurrentController()  # コントローラの取得。
	sheets = doc.getSheets()  # シートコレクションを取得。
	txt = target.getString()  # クリックしたセルの文字列を取得。	

	if txt=="予をﾘｾｯﾄ":
		sheet[splittedrow:emptyrow, ichiran.sumicolumn+1].clearContents(CellFlags.STRING)  # 予列をリセット。
	elif txt=="入力支援":
		
		pass  # 入力支援odsを開く。
	
	elif txt=="退院ﾘｽﾄ":
		controller.setActiveSheet(sheets["退院"])
	return False  # セル編集モードにしない。	
def mousePressedWSectionB(doc, sheet, systemclipboard, functionaccess, transliteration, ent, target):
	createFormatKey = commons.formatkeyCreator(doc)
	controller = doc.getCurrentController()  # コントローラの取得。
	sheets = doc.getSheets()  # シートコレクションを取得。	
	celladdress = target.getCellAddress()
	r, c = celladdress.Row, celladdress.Column  # targetの行と列のインデックスを取得。		
	sumitxt, yotxt, idtxt, kanjitxt, kanatxt, datevalue, keikatxt = sheet[r, :ichiran.checkstartcolumn-1].getDataArray()[0]  # 日付はfloatで返ってくる。	
	datevalue = datevalue and int(datevalue)  # 計算しにくいのでdatevalueがあるときはfloatを整数にしておく。	
	if keikatxt and c==0:  # 経過列があり、かつ、済列の時。
		items = [("待", "skyblue"), ("済", "silver"), ("未", "black")]
		items.append(items[0])  # 最初の要素を最後の要素に追加する。
		dic = {items[i][0]: items[i+1] for i in range(len(items)-1)}  # 順繰り辞書の作成。								
		target.setString(dic[sumitxt][0])
		sheet[r, :].setPropertyValue("CharColor", commons.COLORS[dic[sumitxt][1]])						
		refreshCounts(sheet, ichiran)  # カウントを更新する。
	elif keikatxt and c==ichiran.yocolumn:  # 経過列があり、かつ、予列の時。
		if yotxt:
			target.clearContents(CellFlags.STRING)  # 予をクリア。
		else:  # セルの文字列が空の時。
			target.setString("予")
	elif c==ichiran.idcolumn:  # ID列の時。
		if keikatxt:  # 経過列がある時。
			systemclipboard.setContents(commons.TextTransferable(idtxt), None)  # クリップボードにIDをコピーする。
		else:
			return True  # セル編集モードにする。		
	elif c==ichiran.idcolumn+1:  # 漢字列の時。カルテシートをアクティブにする、なければ作成する。カルトシート名はIDと一致。	
		if keikatxt and idtxt in sheets:  # 経過列があり、かつ、ID名のシートが存在する時。
			controller.setActiveSheet(sheets[idtxt])  # カルテシートをアクティブにする。
		else:  # 在院日数列が空欄の時、または、カルテシートがない時。
			if all((idtxt, kanjitxt, kanatxt, datevalue)):  # ID、漢字名、カナ名、入院日、すべてが揃っている時。	
				fillColumns(transliteration, createFormatKey, sheet, r, ichiran, idtxt, kanjitxt, kanatxt, datevalue)
				newsheet = sheets[idtxt]  # カルテシートを取得。  
				karuteconsts = karute.Karute(newsheet)	
				karutedatecell = newsheet[karuteconsts.splittedrow, karuteconsts.datecolumn]
				karutedatecell.setValue(datevalue)  # カルテシートに入院日を入力。
				karutedatecell.setPropertyValues(("NumberFormat", "HoriJustify"), (createFormatKey('YYYY/MM/DD'), LEFT))  # カルテシートの入院日の書式設定。左寄せにする。
				newsheet[:karuteconsts.splittedrow, karuteconsts.articlecolumn].setDataArray(("",), (" ".join((idtxt, kanjitxt, kanatxt)),))  # カルテシートのコピー日時をクリア。ID名前を入力。
				controller.setActiveSheet(newsheet)  # カルテシートをアクティブにする。	
			else:
				return True  # セル編集モードにする。		
	elif c==ichiran.kanacolumn:  # カナ名列の時。
		if keikatxt:  # 経過列がすでにある時。
			kanatxt = convertKanaFULLWIDTH(transliteration, kanatxt)  # カナ名を半角からスペースを削除して全角にする。
			systemclipboard.setContents(commons.TextTransferable("".join((kanatxt, idtxt))), None)  # クリップボードにカナ名+IDをコピーする。	
		else:
			return True  # セル編集モードにする。		
	elif c==ichiran.datecolumn:  # 入院日列の時。
		if keikatxt:  # 経過列がすでにある時。
			return True  # セル編集モードにする。
		else:
			todaydatevalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。
			if datevalue:  # すでに日付が入っている時。
				items = [todaydatevalue-1, todaydatevalue-2, todaydatevalue, todaydatevalue+1]
				items.append(items[0])  # 最初の要素を最後の要素に追加する。
				dic = {items[i]: items[i+1] for i in range(len(items)-1)}  # 順繰り辞書の作成。								
				if datevalue in dic:
					datevalue = dic[datevalue]
				else:
					return True  # セル編集モードにする。
			else:
				datevalue = todaydatevalue
			target.setValue(datevalue)
			target.setPropertyValue("NumberFormat", createFormatKey('YYYY/MM/DD'))
	elif c==ichiran.datecolumn+1:  # 経過列のボタンはカルテシートの作成時に作成されるのでカルテシート作成後のみ有効。			
		newsheetname = "".join([idtxt, "経"])  # 経過シート名を取得。
		if keikatxt and newsheetname in sheets:  # 経過列がすでにあり、かつ、経過シートがある時。。		
			controller.setActiveSheet(sheets[newsheetname])  # 経過シートをアクティブにする。
		else:  # 経過シートがなければ作成する。
			if all((idtxt, kanjitxt, kanatxt, datevalue)):  # ID、漢字名、カナ名、入院日、すべてが揃っている時。									
				fillColumns(transliteration, createFormatKey, sheet, r, ichiran, idtxt, kanjitxt, kanatxt, datevalue)
				if newsheetname in sheets:  # すでに経過シートがある時。
					keikasheet = sheets[newsheetname]  # 新規経過シートを取得。
				else:	
					sheets.copyByName("00000000経", newsheetname, len(sheets))  # テンプレートシートをコピーしてID経名のシートにして最後に挿入。	
					keikasheet = sheets[newsheetname]  # 新規経過シートを取得。
					keikasheet["F2"].setString(" ".join((idtxt, kanjitxt, kanatxt)))  # ID漢字名ｶﾅ名を入力。					
					keika.setDates(doc, keikasheet, keikasheet["I2"], datevalue)  # 経過シートの日付を設定。
				controller.setActiveSheet(keikasheet)  # 経過シートをアクティブにする。						
	return False  # セル編集モードにしない。		
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
def drowBorders(sheet, cellrange, borders):  # ターゲットを交点とする行列全体の外枠線を描く。
	cell = cellrange[0, 0]  # セル範囲の左上端のセルで判断する。
	ent = getSectionName(sheet, cell)
	sectionname = ent.sectionname	
	if sectionname in ("M", ):
		return	
	noneline, dummy, topbottomtableborder, dummy = borders	
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
	rangeaddress = cellrange.getRangeAddress()  # セル範囲アドレスを取得。
	if sectionname in ("A", "B"):
		sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。					

