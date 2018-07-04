#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper, json, locale  # import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from datetime import date, datetime, timedelta
from com.sun.star.awt import XMouseListener, XTextListener
from com.sun.star.awt import MenuItemStyle, MouseButton, PopupMenuDirection, PosSize  # 定数
from com.sun.star.awt import Point, Rectangle  # Struct
from com.sun.star.beans import NamedValue  # Struct
from com.sun.star.frame import XFrameActionListener
from com.sun.star.frame.FrameAction import FRAME_UI_DEACTIVATING  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.style.VerticalAlignment import MIDDLE  # enum
from com.sun.star.util import XCloseListener
from com.sun.star.util import MeasureUnit  # 定数
from com.sun.star.view.SelectionType import SINGLE  # enum 
SHEETNAME = "config"  # データを保存するシート名。
def createDialog(xscriptcontext, enhancedmouseevent, dialogtitle, formatstring=None):  # dialogtitleはダイアログのデータ保存名に使うのでユニークでないといけない。formatstringは代入セルの書式。
	dateformat = "%Y/%m/%d(%a)"  # 日付書式。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	doc = xscriptcontext.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	dialogpoint = getDialogPoint(doc, enhancedmouseevent)  # クリックした位置のメニューバーの高さ分下の位置を取得。単位ピクセル。一部しか表示されていないセルのときはNoneが返る。
	if not dialogpoint:  # クリックした位置が取得出来なかった時は何もしない。
		return
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	containerwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	maTopx = createConverters(containerwindow)  # ma単位をピクセルに変換する関数を取得。
	m = 2  # コントロール間の間隔。
	h = 12  # コントロールの高さ
	gridprops = {"PositionX": 0, "PositionY": 0, "Width": 76, "Height": 73, "ShowRowHeader": False, "ShowColumnHeader": False, "VScroll": False, "SelectionModel": SINGLE}  # グリッドコントロールのプロパティ。
	controlcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(gridprops), "Height": YHeight(gridprops), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
	optioncontrolcontainerprops = controlcontainerprops.copy()
	controlcontainer, addControl = controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。		
	items = ("セル入力で閉じる", MenuItemStyle.CHECKABLE+MenuItemStyle.AUTOCHECK, {"checkItem": True}),  # グリッドコントロールのコンテクストメニュー。XMenuListenerのmenuevent.MenuIdでコードを実行する。
	gridpopupmenu = menuCreator(ctx, smgr)("PopupMenu", items)  # 右クリックでまず呼び出すポップアップメニュー。 
	args = xscriptcontext, dateformat, formatstring
	mouselistener = MouseListener(gridpopupmenu, args)
	gridcontrolwidth = gridprops["Width"]  # gridpropsは消費されるので、グリッドコントロールの幅を取得しておく。
	gridcontrol1 = addControl("Grid", gridprops, {"addMouseListener": mouselistener})  # グリッドコントロールの取得。
	gridcolumn = gridcontrol1.getModel().getPropertyValue("ColumnModel")  # DefaultGridColumnModel
	column0 = gridcolumn.createColumn() # 列の作成。
	column0.ColumnWidth = 25 # 列幅。
	gridcolumn.addColumn(column0)  # 1列目を追加。
	column1 = gridcolumn.createColumn() # 列の作成。
	column1.ColumnWidth = gridcontrolwidth - column0.ColumnWidth #  列幅。列の合計がグリッドコントロールの幅に一致するようにする。
	gridcolumn.addColumn(column1)  # 2列目を追加。
	numericfieldprops1 = {"PositionY": m, "Width": 24, "Height": h+2, "Spin": True, "StrictFormat": True, "Value": 0, "ValueStep": 1, "ShowThousandsSeparator": False, "DecimalAccuracy": 0}
	fixedtextprops1 = {"PositionY": m, "Width": 14, "Height": h, "Label": "週後", "VerticalAlign": MIDDLE}
	fixedtextprops1.update({"PositionX": gridcontrolwidth-fixedtextprops1["Width"]})
	numericfieldprops1.update({"PositionX": fixedtextprops1["PositionX"]-numericfieldprops1["Width"]})
	optioncontrolcontainerprops.update({"PositionY": YHeight(optioncontrolcontainerprops), "Height": YHeight(numericfieldprops1, m)})
	optioncontrolcontainer, optionaddControl = controlcontainerMaCreator(ctx, smgr, maTopx, optioncontrolcontainerprops)  # コントロールコンテナの作成。
	textlistener = TextListener(dateformat, gridcontrol1)
	numericfield1 = optionaddControl("NumericField", numericfieldprops1, {"addTextListener": textlistener})		
	optionaddControl("FixedText", fixedtextprops1)
	rectangle = controlcontainer.getPosSize()  # コントロールコンテナのRectangle Structを取得。px単位。
	rectangle.X, rectangle.Y = dialogpoint  # クリックした位置を取得。ウィンドウタイトルを含めない座標。
	rectangle.Height += optioncontrolcontainer.getSize().Height
	taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
	args = NamedValue("PosSize", rectangle), NamedValue("FrameName", "controldialog")  # , NamedValue("MakeVisible", True)  # TaskCreatorで作成するフレームのコンテナウィンドウのプロパティ。
	dialogframe = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。サイズ変更は想定しない。
	mouselistener.dialogframe = dialogframe
	dialogwindow = dialogframe.getContainerWindow()  # ダイアログのコンテナウィンドウを取得。
	dialogframe.setTitle(dialogtitle)  # フレームのタイトルを設定。
	docframe.getFrames().append(dialogframe) # 新しく作ったフレームを既存のフレームの階層に追加する。
	toolkit = dialogwindow.getToolkit()  # ピアからツールキットを取得。 	
	controlcontainer.createPeer(toolkit, dialogwindow) # ウィンドウにコントロールコンテナを描画。
	optioncontrolcontainer.createPeer(toolkit, dialogwindow) # ウィンドウにオプションコントロールコンテナを描画。Visibleにはしない。
	frameactionlistener = FrameActionListener()  # FrameActionListener。フレームがアクティブでなくなった時に閉じるため。
	dialogframe.addFrameActionListener(frameactionlistener)  # FrameActionListenerをダイアログフレームに追加。
	controlcontainer.setVisible(True)  # コントロールの表示。
	optioncontrolcontainer.setVisible(True)
	dialogwindow.setVisible(True) # ウィンドウの表示。これ以降WindowListenerが発火する。
	numericfield1.setFocus()
	todayindex = 7//2  # 今日の日付の位置を決定。切り下げ。
	col0 = [""]*7  # 全てに空文字を挿入。
	cellvalue = enhancedmouseevent.Target.getValue()  # セルの値を取得。
	centerday = None
	if cellvalue and isinstance(cellvalue, float):  # セルの値がfloat型のとき。datevalueと決めつける。
		functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
		if cellvalue!=functionaccess.callFunction("TODAY", ()):  # セルの数値が今日でない時。
			centerday = date(*[int(functionaccess.callFunction(i, (cellvalue,))) for i in ("YEAR", "MONTH", "DAY")])  # シリアル値をシート関数で年、月、日に変換してdateオブジェクトにする。
			col0[todayindex] = "基準日"
	if centerday is None:
		centerday = date.today()
		col0[todayindex-1:todayindex+2] = "昨日", "今日", "明日"  # 列インデックス0に入れる文字列を取得。
	addDays(gridcontrol1, dateformat, centerday, col0)  # グリッドコントロールに行を入れる。	
	dialogstate = getSavedData(doc, "dialogstate_{}".format(dialogtitle))  # 保存データを取得。optioncontrolcontainerの表示状態は常にFalseなので保存されていない。
	if dialogstate is not None:  # 保存してあるダイアログの状態がある時。
		for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
			itemtext = gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
			if itemtext.startswith("セル入力で閉じる"):
				closecheck = dialogstate.get("CloseCheck")  # セル入力で閉じる、のチェックがある時。
				if closecheck is not None:
					gridpopupmenu.checkItem(menuid, closecheck)	
	args = doc, mouselistener, controlcontainer
	dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。		
def addDays(gridcontrol, dateformat, centerday, col0, daycount=7):
	todayindex = 7//2  # 今日の日付の位置を決定。切り下げ。
	startday = centerday - timedelta(days=1)*todayindex  # 開始dateを取得。
	dategene = (startday+timedelta(days=i) for i in range(daycount))  # daycount分のdateオブジェクトのジェネレーターを取得。
	locale.setlocale(locale.LC_ALL, '')  # これがないとWindows10では曜日が英語になる。locale.setlocale(locale.LC_TIME, 'ja_JP.utf-8')では文字化けする。
	datarows = tuple(zip(col0, (i.strftime(dateformat) for i in dategene)))  # 列インデックス0に語句、列インデックス1に日付を入れる。
	griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModel
	griddatamodel.removeAllRows()  # グリッドコントロールの行を全削除。
	griddatamodel.addRows(("",)*len(datarows), datarows)  # グリッドに行を追加。	
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログのリスナー削除用。
	def __init__(self, args):
		self.args = args
	def queryClosing(self, eventobject, getsownership):  # ノンモダルダイアログを閉じる時に発火。
		dialogframe = eventobject.Source
		doc, mouselistener, controlcontainer = self.args
		gridpopupmenu = mouselistener.gridpopupmenu
		for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
			itemtext = gridpopupmenu.getItemText(menuid)
			if itemtext.startswith("セル入力で閉じる"):
				dialogstate = {"CloseCheck": gridpopupmenu.isItemChecked(menuid)}
		dialogtitle = dialogframe.getTitle()  # コンテナウィンドウタイトルを取得。データ保存のIDに使う。
		saveData(doc, "dialogstate_{}".format(dialogtitle), dialogstate)  # ダイアログの状態を保存。
		controlcontainer.getControl("Grid1").removeMouseListener(mouselistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		pass
class TextListener(unohelper.Base, XTextListener):
	def __init__(self, *args):
		self.args = args
		self.val = 0  # 変更前の値。
	def textChanged(self, textevent):
		numericfield = textevent.Source
		dateformat, gridcontrol = self.args
		todayindex = 7//2  # 本日と同じインデックスを取得。
		datetxt = gridcontrol.getModel().getPropertyValue("GridDataModel").getCellData(1, todayindex)  # 中央行の日付文字列を取得。
		centerday = datetime.strptime(datetxt, dateformat).date()  # 現在の最初のdateオブジェクトを取得。
		val = numericfield.getValue()  # 数値フィールドの値を取得。		
		diff = val - self.val  # 前値との差を取得。
		centerday += timedelta(days=7*diff)  # 週を移動。
		col0 = [""]*7
		if val==0:
			if centerday==date.today():
				col0[todayindex-1:todayindex+2] = "昨日", "今日", "明日"  # 列インデックス0に入れる文字列を取得。
			else:	
				col0[todayindex] = "基準日"
		else:
			txt = "{}週後" if val>0 else "{}週前" 
			col0[todayindex] = txt.format(int(abs(val)))  # valはfloatなので小数点が入ってくる。		
		addDays(gridcontrol, dateformat, centerday, col0)  # グリッドコントロールに行を入れる。
		self.val = val  # 変更後の値を前値として取得。
	def disposing(self, eventobject):
		pass
def saveData(doc, rangename, obj):	# configシートの名前rangenameにobjをJSONにして保存する。グローバル変数SHEETNAMEを使用。
	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。
	if rangename in namedranges:  # 名前がある時。
		referredcells = namedranges[rangename].getReferredCells()  # 名前の参照セル範囲を取得。
		if referredcells:  # 参照セル範囲がある時。
			referredcells.setString(json.dumps(obj,  ensure_ascii=False))  # rangenameという名前のセルに文字列でPythonオブジェクトを出力する。
			return
		else:  # 名前があるが参照セル範囲がない時。
			namedranges.removeByName(rangename)  # 名前は重複しているとエラーになるので削除する。	
	if not rangename in namedranges:  # 名前がない時。
		sheets = doc.getSheets()  # シートコレクションを取得。
		if not SHEETNAME in sheets:  # 保存シートがない時。
			sheets.insertNewByName(SHEETNAME, len(sheets))   # 履歴シートを挿入。同名のシートがあるとRuntimeExceptionがでる。
		sheet = sheets[SHEETNAME]  # 保存シートを取得。
		sheet.setPropertyValue("IsVisible", False)  # 非表示シートにする。
		emptyranges = sheet[:, 0].queryEmptyCells()  # 1列目の空セル範囲コレクションを取得。
		if len(emptyranges):  # セル範囲コレクションが取得出来た時。
			cellcursor = sheet.createCursorByRange(emptyranges[0][0, 0])  # 最初のセル範囲の最初のセルからセルカーサーを取得。
			cellcursor.collapseToSize(2, 1)  # 行数1、列数2に変更。(列、行)で指定。
			cellcursor.setDataArray(((rangename, json.dumps(obj, ensure_ascii=False)),))  # セル範囲名を1列目、JSONデータを2列目に代入。
			namedranges.addNewByName(rangename, cellcursor[0, 1].getPropertyValue("AbsoluteName"), cellcursor[0, 1].getCellAddress(), 0)  # 2列目のセルに名前を付ける。名前、式(相対アドレス)、原点となるセル、NamedRangeFlag
def getSavedData(doc, rangename):  # configシートのragenameからデータを取得する。	
	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。	
	if rangename in namedranges:  # 名前がある時。
		referredcells = namedranges[rangename].getReferredCells()  # 名前が参照しているセル範囲を取得。参照アドレスがエラーのときはNoneが返る。
		if referredcells:
			txt = referredcells.getString()  # 名前が参照しているセルから文字列を取得。
			if txt:
				try:
					return json.loads(txt)  # pyunoオブジェクトは変換できない。
				except json.JSONDecodeError:
					import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	return None  # 保存された行が取得できない時はNoneを返す。
class MouseListener(unohelper.Base, XMouseListener):  
	def __init__(self, gridpopupmenu, args): 	
		self.args = args
		self.gridpopupmenu = gridpopupmenu  # CloseListenerでも使用する。
		self.dialogframe = None
	def mousePressed(self, mouseevent):  # グリッドコントロールをクリックした時。コントロールモデルにはNameプロパティはない。
		xscriptcontext, dateformat, formatstring = self.args
		gridcontrol = mouseevent.Source  # グリッドコントロールを取得。
		if mouseevent.Buttons==MouseButton.LEFT:
			if mouseevent.ClickCount==2:  # ダブルクリックの時。
				doc = xscriptcontext.getDocument()
				selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
				if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
					rowindexes = getSelectedRowIndexes(gridcontrol)  # グリッドコントロールの選択行インデックスを返す。昇順で返す。負数のインデックスがある時は要素をクリアする。
					if rowindexes:
						datetxt = gridcontrol.getModel().getPropertyValue("GridDataModel").getCellData(1, rowindexes[0])  # 選択行の日付文字列を取得。
						selectedday = datetime.strptime(datetxt, dateformat)  # 現在の最初のdatetimeオブジェクトを取得。	
						selection.setFormula(selectedday.isoformat())  # セルに式として代入。			
						numberformats = doc.getNumberFormats()  # ドキュメントのフォーマット一覧を取得。デフォルトのフォーマット一覧はCalcの書式→セル→数値でみれる。
						if formatstring is not None:  # 書式が与えられている時。
							locale = Locale(Language="ja", Country="JP")  # フォーマット一覧をくくる言語と国を設定。インストールしていないUIの言語でもよい。。 
							formatkey = numberformats.queryKey(formatstring, locale, True)  # formatstringが既存のフォーマット一覧にあるか調べて取得。第3引数のブーリアンは意味はないはず。 
							if formatkey == -1:  # デフォルトのフォーマットにformatstringがないとき。
								formatkey = numberformats.addNew(formatstring, locale)  # フォーマット一覧に追加する。保存はドキュメントごと。
						selection.setPropertyValue("NumberFormat", formatkey)  # セルの書式を設定。 
				for menuid in range(1, self.gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
					itemtext = self.gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
					if itemtext.startswith("セル入力で閉じる"):
						if self.gridpopupmenu.isItemChecked(menuid):  # 選択項目にチェックが入っている時。
							self.dialogframe.close(True)
						else:
							controller = doc.getCurrentController()  # 現在のコントローラを取得。	
							sheet = controller.getActiveSheet()
							celladdress = selection.getCellAddress()
							nextcell = sheet[celladdress.Row+1, celladdress.Column]  # 下のセルを取得。
							controller.select(nextcell)  # 下のセルを選択。							
						break
		elif mouseevent.Buttons==MouseButton.RIGHT:  # 右ボタンクリックの時。mouseevent.PopupTriggerではサブジェクトによってはTrueにならないので使わない。
			pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
			self.gridpopupmenu.execute(gridcontrol.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向							
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		pass
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):
		pass
def getSelectedRowIndexes(gridcontrol):  # グリッドコントロールの選択行インデックスを返す。昇順で返す。負数のインデックスがある時は要素をクリアする。
	selectedrowindexes = list(gridcontrol.getSelectedRows())  # 選択行のインデックスをリストで取得。
	selectedrowindexes.sort()  # 選択順にインデックスが入っているので昇順にソートする。
	if selectedrowindexes and selectedrowindexes[0]<0:  # 負数のインデックスがある時(すべての行を削除した後に行を追加した時など)。
		gridcontrol.deselectAllRows()  # 選択状態を外す。
		selectedrowindexes.clear()  # 選択行インデックスをクリア。
	return selectedrowindexes
class FrameActionListener(unohelper.Base, XFrameActionListener):
	def frameAction(self, frameactionevent):
		if frameactionevent.Action==FRAME_UI_DEACTIVATING:  # フレームがアクティブでなくなった時。TopWindowListenerのwindowDeactivated()だとウィンドウタイトルバーをクリックしただけで発火してしまう。
			frameactionevent.Frame.removeFrameActionListener(self)  # フレームにつけたリスナーを除去。
			frameactionevent.Frame.close(True)
	def disposing(self, eventobject):
		pass
def XWidth(props, m=0):  # 左隣のコントロールからPositionXを取得。mは間隔。
	return props["PositionX"] + props["Width"] + m  	
def YHeight(props, m=0):  # 上隣のコントロールからPositionYを取得。mは間隔。
	return props["PositionY"] + props["Height"] + m
def getDialogPoint(doc, enhancedmouseevent):  # クリックした位置x yのタプルで返す。但し、一部しか見えてないセルの場合はNoneが返る。TaskCreatorのRectangleには画面の左角からの座標を渡すが、ウィンドウタイトルバーは含まれない。
	controller = doc.getCurrentController()  # 現在のコントローラを取得。
	docframe = controller.getFrame()  # フレームを取得。
	containerwindow = docframe.getContainerWindow()  # コンテナウィドウの取得。
	accessiblecontextparent = containerwindow.getAccessibleContext().getAccessibleParent()  # コンテナウィンドウの親AccessibleContextを取得する。フレームの子AccessibleContextになる。
	accessiblecontext = accessiblecontextparent.getAccessibleContext()  # AccessibleContextを取得。
	for i in range(accessiblecontext.getAccessibleChildCount()): 
		childaccessiblecontext = accessiblecontext.getAccessibleChild(i).getAccessibleContext()
		if childaccessiblecontext.getAccessibleRole()==49:  # ROOT_PANEの時。
			rootpanebounds = childaccessiblecontext.getBounds()  # Yアトリビュートがウィンドウタイトルバーの高さになる。
			break 
	else:
		return  # ウィンドウタイトルバーのAccessibleContextが取得できなかった時はNoneを返す。
	componentwindow = docframe.getComponentWindow()  # コンポーネントウィンドウを取得。
	border = controller.getBorder()  # 行ヘッダの幅と列ヘッダの高さの取得のため。
	accessiblecontext = componentwindow.getAccessibleContext()  # コンポーネントウィンドウのAccessibleContextを取得。
	for i in range(accessiblecontext.getAccessibleChildCount()):  # 子AccessibleContextについて。
		childaccessiblecontext = accessiblecontext.getAccessibleChild(i).getAccessibleContext()  # 子AccessibleContextのAccessibleContext。
		if childaccessiblecontext.getAccessibleRole()==51:  # SCROLL_PANEの時。
			for j in range(childaccessiblecontext.getAccessibleChildCount()):  # 孫AccessibleContextについて。 
				grandchildaccessiblecontext = childaccessiblecontext.getAccessibleChild(j).getAccessibleContext()  # 孫AccessibleContextのAccessibleContext。
				if grandchildaccessiblecontext.getAccessibleRole()==84:  # DOCUMENT_SPREADSHEETの時。これが枠。
					bounds = grandchildaccessiblecontext.getBounds()  # 枠の位置と大きさを取得(SCROLL_PANEの左上角が原点)。
					if bounds.X==border.Left and bounds.Y==border.Top:  # SCROLL_PANEに対する相対座標が行ヘッダと列ヘッダと一致する時は左上枠。
						for k, subcontroller in enumerate(controller):  # 各枠のコントローラについて。インデックスも取得する。
							cellrange = subcontroller.getReferredCells()  # 見えているセル範囲を取得。一部しかみえていないセルは含まれない。
							if len(cellrange.queryIntersection(enhancedmouseevent.Target.getRangeAddress())):  # ターゲットが含まれるセル範囲コレクションが返る時その枠がクリックした枠。「ウィンドウの分割」では正しいiは必ずしも取得できない。
								sourcepointonscreen =  grandchildaccessiblecontext.getLocationOnScreen()  # 左上枠の左上角の点を取得(画面の左上角が原点)。
								if k==1:  # 左下枠の時。
									sourcepointonscreen = Point(X=sourcepointonscreen.X, Y=sourcepointonscreen.Y+bounds.Height)
								elif k==2:  # 右上枠の時。
									sourcepointonscreen = Point(X=sourcepointonscreen.X+bounds.Width, Y=sourcepointonscreen.Y)
								elif k==3:  # 右下枠の時。
									sourcepointonscreen = Point(X=sourcepointonscreen.X+bounds.Width, Y=sourcepointonscreen.Y+bounds.Height)
								x = sourcepointonscreen.X + enhancedmouseevent.X  # クリックした位置の画面の左上角からのXの取得。
								y = sourcepointonscreen.Y + enhancedmouseevent.Y + rootpanebounds.Y  # クリックした位置からメニューバーの高さ分下の位置の画面の左上角からのYの取得									
								return x, y
def menuCreator(ctx, smgr):  #  メニューバーまたはポップアップメニューを作成する関数を返す。
	def createMenu(menutype, items, attr=None):  # menutypeはMenuBarまたはPopupMenu、itemsは各メニュー項目の項目名、スタイル、適用するメソッドのタプルのタプル、attrは各項目に適用する以外のメソッド。
		if attr is None:
			attr = {}
		menu = smgr.createInstanceWithContext("com.sun.star.awt.{}".format(menutype), ctx)
		for i, item in enumerate(items, start=1):  # 各メニュー項目について。
			if item:
				if len(item) > 2:  # タプルの要素が3以上のときは3番目の要素は適用するメソッドの辞書と考える。
					item = list(item)
					attr[i] = item.pop()  # メニュー項目のIDをキーとしてメソッド辞書に付け替える。
				menu.insertItem(i, *item, i-1)  # ItemId, Text, ItemSytle, ItemPos。ItemIdは1から始まり区切り線(欠番)は含まない。ItemPosは0から始まり区切り線を含む。
			else:  # 空のタプルの時は区切り線と考える。
				menu.insertSeparator(i-1)  # ItemPos
		if attr:  # メソッドの適用。
			for key, val in attr.items():  # keyはメソッド名あるいはメニュー項目のID。
				if isinstance(val, dict):  # valが辞書の時はkeyは項目ID。valはcreateMenu()の引数のitemsであり、itemsの３番目の要素にキーをメソッド名とする辞書が入っている。
					for method, arg in val.items():  # 辞書valのキーはメソッド名、値はメソッドの引数。
						if method in ("checkItem", "enableItem", "setCommand", "setHelpCommand", "setHelpText", "setTipHelpText"):  # 第1引数にIDを必要するメソッド。
							getattr(menu, method)(key, arg)
						else:
							getattr(menu, method)(arg)
				else:
					getattr(menu, key)(val)
		return menu
	return createMenu
def controlcontainerMaCreator(ctx, smgr, maTopx, containerprops):  # ma単位でコントロールコンテナと、それにコントロールを追加する関数を返す。まずコントロールコンテナモデルのプロパティを取得。UnoControlDialogElementサービスのプロパティは使えない。propsのキーにPosSize、値にPOSSIZEが必要。
	container = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)  # コントロールコンテナの生成。
	container.setPosSize(*maTopx(containerprops.pop("PositionX"), containerprops.pop("PositionY")), *maTopx(containerprops.pop("Width"), containerprops.pop("Height")), PosSize.POSSIZE)
	containermodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)  # コンテナモデルの生成。
	containermodel.setPropertyValues(tuple(containerprops.keys()), tuple(containerprops.values()))  # コンテナモデルのプロパティを設定。存在しないプロパティに設定してもエラーはでない。
	container.setModel(containermodel)  # コンテナにコンテナモデルを設定。
	container.setVisible(False)  # 描画中のものを表示しない。
	def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、キーにPosSize、値にPOSSIZEが必要。attr: コントロールの属性。
		name = props.pop("Name") if "Name" in props else _generateSequentialName(controltype) # サービスマネージャーからインスタンス化したコントロールはNameプロパティがないので、コントロールコンテナのaddControl()で名前を使うのみ。
		controlidl = "com.sun.star.awt.grid.UnoControl{}".format(controltype) if controltype=="Grid" else "com.sun.star.awt.UnoControl{}".format(controltype)  # グリッドコントロールだけモジュールが異なる。
		control = smgr.createInstanceWithContext(controlidl, ctx)  # コントロールを生成。
		control.setPosSize(*maTopx(props.pop("PositionX"), props.pop("PositionY")), *maTopx(props.pop("Width"), props.pop("Height")), PosSize.POSSIZE)  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
		controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
		control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
		container.addControl(name, control)  # コントロールをコントロールコンテナに追加。
		if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
			for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
				if val is None:
					getattr(control, key)()
				else:
					getattr(control, key)(val)
		return control  # 追加したコントロールを返す。
	def _createControlModel(controltype, props):  # コントロールモデルの生成。
		controlmodelidl = "com.sun.star.awt.grid.UnoControl{}Model".format(controltype) if controltype=="Grid" else "com.sun.star.awt.UnoControl{}Model".format(controltype)
		controlmodel = smgr.createInstanceWithContext(controlmodelidl, ctx) # コントロールモデルを生成。UnoControlDialogElementサービスはない。
		if props:
			values = props.values()  # プロパティの値がタプルの時にsetProperties()でエラーが出るのでその対応が必要。
			if any(map(isinstance, values, [tuple]*len(values))):
				[setattr(controlmodel, key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。XMultiPropertySetのsetPropertyValues()では[]anyと判断されてタプルも使えない。
			else:
				controlmodel.setPropertyValues(tuple(props.keys()), tuple(values))
		return controlmodel
	def _generateSequentialName(controltype):  # コントロールの連番名の作成。
		i = 1
		flg = True
		while flg:
			name = "{}{}".format(controltype, i)
			flg = container.getControl(name)  # 同名のコントロールの有無を判断。
			i += 1
		return name
	return container, addControl  # コントロールコンテナとそのコントロールコンテナにコントロールを追加する関数を返す。
def createConverters(window):  # ma単位をピクセルに変換する関数を返す。
	def maTopx(x, y):  # maをpxに変換する。
		point = window.convertPointToPixel(Point(X=x, Y=y), MeasureUnit.APPFONT)
		return point.X, point.Y
	return maTopx
