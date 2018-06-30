#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper, json  # import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from com.sun.star.awt import XActionListener, XMenuListener, XMouseListener, XWindowListener
from com.sun.star.awt import MenuItemStyle, MessageBoxButtons, MessageBoxResults, MouseButton, PopupMenuDirection, PosSize  # 定数
from com.sun.star.awt import MenuEvent, Point, Rectangle  # Struct
from com.sun.star.awt.MessageBoxType import QUERYBOX  # enum
from com.sun.star.beans import NamedValue  # Struct
from com.sun.star.frame import XFrameActionListener
from com.sun.star.frame.FrameAction import FRAME_UI_DEACTIVATING  # enum
from com.sun.star.util import XCloseListener
from com.sun.star.util import MeasureUnit  # 定数
from com.sun.star.view.SelectionType import MULTI  # enum 
SHEETNAME = "config"  # データを保存するシート名。
def createDialog(xscriptcontext, enhancedmouseevent, dialogtitle, defaultrows=None):  # dialogtitleはダイアログのデータ保存名に使うのでユニークでないといけない。defaultrowsはグリッドコントロールのデフォルトデータ。
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
	gridprops = {"PositionX": 0, "PositionY": 0, "Width": 50, "Height": 50, "ShowRowHeader": False, "ShowColumnHeader": False, "SelectionModel": MULTI}  # グリッドコントロールのプロパティ。
	controlcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(gridprops), "Height": YHeight(gridprops), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
	controlcontainer, addControl = controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。		
	menulistener = MenuListener()  # コンテクストメニューにつけるリスナー。
	items = ("セル入力で閉じる", MenuItemStyle.CHECKABLE+MenuItemStyle.AUTOCHECK, {"checkItem": False}),\
			("オプション表示", MenuItemStyle.CHECKABLE+MenuItemStyle.AUTOCHECK, {"checkItem": False})  # グリッドコントロールのコンテクストメニュー。XMenuListenerのmenuevent.MenuIdでコードを実行する。
	gridpopupmenu = menuCreator(ctx, smgr)("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。 
	mouselistener = MouseListener(xscriptcontext, gridpopupmenu)
	gridcontrol1 = addControl("Grid", gridprops, {"addMouseListener": mouselistener})  # グリッドコントロールの取得。
	gridmodel = gridcontrol1.getModel()  # グリッドコントロールモデルの取得。
	gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
	gridcolumn.addColumn(gridcolumn.createColumn())  # 列を追加。
	griddatamodel = gridmodel.getPropertyValue("GridDataModel")  # GridDataModel
	datarows = getSavedData(doc, "GridDatarows_{}".format(dialogtitle))  # グリッドコントロールの行をconfigシートのragenameから取得する。	
	if datarows is None and defaultrows is not None:  # 履歴がなくデフォルトdatarowsがあるときデフォルトデータを使用。
		datarows = [i if isinstance(i, (list, tuple)) else (i,) for i in defaultrows]  # defaultrowsの要素をリストかタプルでなければタプルに変換する。
	if datarows:  # 行のリストが取得出来た時。
		griddatamodel.addRows(("",)*len(datarows), datarows)  # グリッドに行を追加。	
	controlcontainerwindowlistener = ControlContainerWindowListener(controlcontainer)		
	controlcontainer.addWindowListener(controlcontainerwindowlistener)  # コントロールコンテナの大きさを変更するとグリッドコントロールの大きさも変更するようにする。
	textboxprops = {"PositionX": 0, "PositionY": m, "Height": h}  # テクストボックスコントロールのプロパティ。
	checkboxprops1 = {"PositionX": 0, "PositionY": YHeight(textboxprops, m), "Width": 46, "Height": h, "Label": "~セルに追記", "State": 0} # セルに追記はデフォルトでは無効。
	buttonprops1 = {"PositionX": XWidth(checkboxprops1), "PositionY": YHeight(textboxprops, m), "Width": 18, "Height": h+2, "Label": "上へ"}  # ボタンのプロパティ。PushButtonTypeの値はEnumではエラーになる。VerticalAlignではtextboxと高さが揃わない。
	buttonprops3 = {"PositionX": XWidth(buttonprops1, 2), "PositionY": YHeight(textboxprops, m), "Width": 26, "Height": h+2, "Label": "行挿入"}
	checkboxprops2 = {"PositionX": 0, "PositionY": YHeight(checkboxprops1, 4), "Width": 46, "Height": h, "Label": "~サイズ復元", "State": 1}  # サイズ復元はデフォルトでは有効。		
	buttonprops2 = {"PositionX": XWidth(checkboxprops1), "PositionY": YHeight(buttonprops1, m*2), "Width": 18, "Height": h+2, "Label": "下へ"}
	buttonprops4 = {"PositionX": XWidth(buttonprops1, m), "PositionY": YHeight(buttonprops1, m*2), "Width": 26, "Height": h+2, "Label": "行削除"}
	textboxprops.update({"Width": XWidth(buttonprops3, m)})  # 右端のコントロールから左の余白mを除いた幅を取得。
	optioncontrolcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(textboxprops), "Height": YHeight(buttonprops2, 2), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
	optioncontrolcontainer, optionaddControl = controlcontainerMaCreator(ctx, smgr, maTopx, optioncontrolcontainerprops)  # コントロールコンテナの作成。		
	optionaddControl("Edit", textboxprops)  
	checkboxcontrol1 = optionaddControl("CheckBox", checkboxprops1)
	checkboxcontrol2 = optionaddControl("CheckBox", checkboxprops2)  
	actionlistener = ActionListener(gridcontrol1, datarows)  # ボタンコントロールにつけるリスナー。
	optionaddControl("Button", buttonprops1, {"addActionListener": actionlistener, "setActionCommand": "up"})  
	optionaddControl("Button", buttonprops2, {"addActionListener": actionlistener, "setActionCommand": "down"})  
	optionaddControl("Button", buttonprops3, {"addActionListener": actionlistener, "setActionCommand": "insert"})  
	optionaddControl("Button", buttonprops4, {"addActionListener": actionlistener, "setActionCommand": "delete"})  
	optioncontrolcontainerwindowlistener = OptionControlContainerWindowListener(optioncontrolcontainer)		
	optioncontrolcontainer.addWindowListener(optioncontrolcontainerwindowlistener)  # コントロールコンテナの大きさを変更するとグリッドコントロールの大きさも変更するようにする。
	mouselistener.optioncontrolcontainer = optioncontrolcontainer
	rectangle = controlcontainer.getPosSize()  # コントロールコンテナのRectangle Structを取得。px単位。
	rectangle.X, rectangle.Y = dialogpoint  # クリックした位置を取得。ウィンドウタイトルを含めない座標。
	taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
	args = NamedValue("PosSize", rectangle), NamedValue("FrameName", "controldialog")  # , NamedValue("MakeVisible", True)  # TaskCreatorで作成するフレームのコンテナウィンドウのプロパティ。
	dialogframe = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。
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
	dialogwindow.setVisible(True) # ウィンドウの表示。これ以降WindowListenerが発火する。
	windowlistener = WindowListener(controlcontainer, optioncontrolcontainer) # コンテナウィンドウからコントロールコンテナを取得する方法はないはずなので、ここで渡す。WindowListenerはsetVisible(True)で呼び出される。
	dialogwindow.addWindowListener(windowlistener) # コンテナウィンドウにリスナーを追加する。
	menulistener.args = dialogwindow, windowlistener
	dialogstate = getSavedData(doc, "dialogstate_{}".format(dialogtitle))  # 保存データを取得。optioncontrolcontainerの表示状態は常にFalseなので保存されていない。
	if dialogstate is not None:  # 保存してあるダイアログの状態がある時。
		for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
			itemtext = gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
			if itemtext.startswith("セル入力で閉じる"):
				closecheck = dialogstate.get("CloseCheck")  # セル入力で閉じる、のチェックがある時。
				if closecheck is not None:
					gridpopupmenu.checkItem(menuid, closecheck)
			elif itemtext.startswith("オプション表示"):
				optioncheck = dialogstate.get("OptionCheck")  # オプション表示、のチェックがある時。
				if optioncheck is not None:
					gridpopupmenu.checkItem(menuid, optioncheck)  # ItemIDは1から始まる。これでMenuListenerは発火しない。
					if optioncheck:  # チェックが付いている時MenuListenerを発火させる。
						menulistener.itemSelected(MenuEvent(MenuId=menuid, Source=mouselistener.gridpopupmenu))
		checkbox1sate = dialogstate.get("CheckBox1sate")  # セルに追記、チェックボックス。キーがなければNoneが返る。	
		if checkbox1sate is not None:  # セルに追記、が保存されている時。
			checkboxcontrol1.setState(checkbox1sate)  # 状態を復元。
		checkbox2sate = dialogstate.get("CheckBox2sate")  # サイズ復元、チェックボックス。	
		if checkbox2sate is not None:  # サイズ復元、が保存されている時。
			checkboxcontrol2.setState(checkbox2sate)  # 状態を復元。	
			if checkbox2sate:  # サイズ復元がチェックされている時。
				dialogwindow.setPosSize(0, 0, dialogstate["Width"], dialogstate["Height"], PosSize.SIZE)  # ウィンドウサイズを復元。WindowListenerが発火する。
	args = doc, actionlistener, dialogwindow, windowlistener, mouselistener, menulistener, controlcontainerwindowlistener, optioncontrolcontainerwindowlistener
	dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。	
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログのリスナー削除用。
	def __init__(self, args):
		self.args = args
	def queryClosing(self, eventobject, getsownership):  # ノンモダルダイアログを閉じる時に発火。
		dialogframe = eventobject.Source
		doc, actionlistener, dialogwindow, windowlistener, mouselistener, menulistener, controlcontainerwindowlistener, optioncontrolcontainerwindowlistener = self.args
		controlcontainer, optioncontrolcontainer = windowlistener.args
		dialogwindowsize = dialogwindow.getSize()	
		dialogstate = {"CheckBox1sate": optioncontrolcontainer.getControl("CheckBox1").getState(),\
					"CheckBox2sate": optioncontrolcontainer.getControl("CheckBox2").getState(),\
					"Width": dialogwindowsize.Width,\
					"Height": dialogwindowsize.Height}  # チェックボックスコントロールの状態とコンテナウィンドウの大きさを取得。
		gridpopupmenu = mouselistener.gridpopupmenu
		for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
			itemtext = gridpopupmenu.getItemText(menuid)
			if itemtext.startswith("セル入力で閉じる"):
				dialogstate.update({"CloseCheck": gridpopupmenu.isItemChecked(menuid)})
			elif itemtext.startswith("オプション表示"):
				dialogstate.update({"OptionCheck": gridpopupmenu.isItemChecked(menuid)})
		dialogtitle = dialogframe.getTitle()  # コンテナウィンドウタイトルを取得。データ保存のIDに使う。
		saveData(doc, "dialogstate_{}".format(dialogtitle), dialogstate)  # ダイアログの状態を保存。
		saveData(doc, "GridDatarows_{}".format(dialogtitle), actionlistener.datarows)  # ダイアログのグリッドコントロールの行を保存。
		gridpopupmenu.removeMenuListener(menulistener)
		controlcontainer.getControl("Grid1").removeMouseListener(mouselistener)
		[controlcontainer.getControl(i).removeActionListener(actionlistener) for i in ("Button1", "Button2", "Button3", "Button4")]
		controlcontainer.removeWindowListener(controlcontainerwindowlistener)
		optioncontrolcontainer.removeWindowListener(optioncontrolcontainerwindowlistener)
		dialogwindow.removeWindowListener(windowlistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		pass
class ActionListener(unohelper.Base, XActionListener):
	def __init__(self, gridcontrol, datarows):
		self.gridcontrol = gridcontrol
		self.datarows = datarows
	def actionPerformed(self, actionevent):
		cmd = actionevent.ActionCommand
		griddatamodel = self.gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。	
		selectedrowindexes = getSelectedRowIndexes(self.gridcontrol)
		if cmd=="insert":  # 選択行の位置に追加する。複数行を選択している時はボタンを無効にしてある。選択行がない時は最下行に追加する。
			txt = actionevent.Source.getContext().getControl("Edit1").getText()  # テキストボックスコントロールの文字列を取得。
			if selectedrowindexes:  # 選択行がある時。
				i = selectedrowindexes[0]
				griddatamodel.insertRow(i, "", (txt,))
				self.datarows.insert(i, (txt,))
			else:  # 選択行がない時。
				griddatamodel.addRow("", (txt,))
				self.datarows.append((txt,))	
			return	
		if not selectedrowindexes:
			return  # 選択行がない時はここで終わる。
		if cmd=="up":  # 先頭行や連続していない複数行を選択している時はボタンを無効にしてある。
			j = selectedrowindexes[0]  # 選択行の先頭行インデックスを取得。
			datarowsToMove = self.datarows[j:selectedrowindexes[-1]+1]  # 移動させる行のリストを取得。
			del self.datarows[j:selectedrowindexes[-1]+1]  # 移動させる行を削除。
			self.datarows.insert(j-1, "dummy")  # 置換される要素を挿入。
			self.datarows[j-1:j] = datarowsToMove  # 移動させる行を挿入。
			griddatamodel.removeAllRows()  # グリッドコントロールの行を全削除。		
			griddatamodel.addRows(("",)*len(self.datarows), self.datarows)  # グリッドに行を追加。
			firstrow = j - 1  # 選択開始行を取得。
			[self.gridcontrol.selectRow(i) for i in range(firstrow, firstrow+len(selectedrowindexes))]
		elif cmd=="down":  # 最終行や連続していない複数行を選択している時はボタンを無効にしてある。
			j = selectedrowindexes[-1]  # 選択行の最終行インデックスを取得。
			datarowsToMove = self.datarows[selectedrowindexes[0]:j+1]  # 移動させる行のリストを取得。
			self.datarows.insert(j+2, "dummy")  # 置換される要素を挿入。
			self.datarows[j+2:j+3] = datarowsToMove  # 移動させる行を挿入。
			del self.datarows[selectedrowindexes[0]:j+1]  # 移動させる行を削除。
			griddatamodel.removeAllRows()  # グリッドコントロールの行を全削除。	
			griddatamodel.addRows(("",)*len(self.datarows), self.datarows)  # グリッドに行を追加。
			c = len(selectedrowindexes)
			firstrow = j + 2- c # 選択開始行を取得。
			[self.gridcontrol.selectRow(i) for i in range(firstrow, firstrow+c)]
		elif cmd=="delete":
			peer = self.gridcontrol.getPeer()  # ピアを取得。			
			msg = "選択行を削除しますか?"
			msgbox = peer.getToolkit().createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, SHEETNAME, msg)
			if msgbox.execute()==MessageBoxResults.YES:		
				for i in selectedrowindexes[::-1]:  # 選択した行インデックスを後ろから取得。
					self.datarows.pop(i)
					griddatamodel.removeRow(i)  # グリッドコントロールから選択行を削除。
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
	def __init__(self, xscriptcontext, gridpopupmenu): 	
		self.xscriptcontext = xscriptcontext
		self.gridpopupmenu = gridpopupmenu
		self.optioncontrolcontainer = None
		self.dialogframe = None
	def mousePressed(self, mouseevent):  # グリッドコントロールをクリックした時。コントロールモデルにはNameプロパティはない。
		gridcontrol = mouseevent.Source  # グリッドコントロールを取得。
		optioncontrolcontainer = self.optioncontrolcontainer
		if mouseevent.Buttons==MouseButton.LEFT:
			if mouseevent.ClickCount==1:  # シングルクリックの時。
				selectedrowindexes = getSelectedRowIndexes(gridcontrol)
				if not selectedrowindexes:  # 選択行がない時(選択行を削除した時)。
					return  # 何もしない		
				upbuttoncontrol = optioncontrolcontainer.getControl("Button1")
				downbuttoncontrol = optioncontrolcontainer.getControl("Button2")
				insertbuttoncontrol = optioncontrolcontainer.getControl("Button3")
				upbuttoncontrol.setEnable(True)  # まず全てのボタンを有効にする。
				downbuttoncontrol.setEnable(True)
				insertbuttoncontrol.setEnable(True)
				if selectedrowindexes[0]==0:  # 先頭行が選択されている時。
					upbuttoncontrol.setEnable(False)  # 上へボタンを無効にする。
				griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")	
				if selectedrowindexes[-1]==griddatamodel.RowCount-1:  # 最終行が選択されている時。
					downbuttoncontrol.setEnable(False)  # 下へボタンを無効にする。
				indexcount = len(selectedrowindexes)  # 選択行数を取得。
				if indexcount>1:  # 複数行を選択している時。
					insertbuttoncontrol.setEnable(False)  # 行挿入ボタンを無効にする。
					if indexcount!=selectedrowindexes[-1]-selectedrowindexes[0]+1:  # 連続した行でない時。
						upbuttoncontrol.setEnable(False)  # 上へボタンを無効にする。
						downbuttoncontrol.setEnable(False)  # 下へボタンを無効にする。
				rowdata = griddatamodel.getRowData(selectedrowindexes[0])  # 選択行の最初の行のデータを取得。
				optioncontrolcontainer.getControl("Edit1").setText(rowdata[0])  # テキストボックスに選択行の初行の文字列を代入。
				if griddatamodel.RowCount==1:  # 1行しかない時はまた発火できるように選択を外す。
					gridcontrol.deselectRow(0)  # 選択行の選択を外す。選択していない行を指定すると永遠ループになる。
			elif mouseevent.ClickCount==2:  # ダブルクリックの時。
				doc = self.xscriptcontext.getDocument()
				selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
				if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
					griddata = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
					rowdata = griddata.getRowData(gridcontrol.getCurrentRow())  # グリッドコントロールで選択している行のすべての列をタプルで取得。
					if optioncontrolcontainer.getControl("CheckBox1").getState():  # セルに追記、にチェックがある時。グリッドコントロールは1列と決めつけて処理する。
						selection.setString("".join([selection.getString(), rowdata[0]]))  # セルに追記する。
					else:
						selection.setString(rowdata[0])  # セルに代入。
						controller = doc.getCurrentController()  # 現在のコントローラを取得。			
						sheet = controller.getActiveSheet()
						celladdress = selection.getCellAddress()
						nextcell = sheet[celladdress.Row+1, celladdress.Column]  # 下のセルを取得。
						controller.select(nextcell)  # 下のセルを選択。	
				gridpopupmenu = self.gridpopupmenu		
				for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
					itemtext = gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
					if itemtext.startswith("セル入力で閉じる"):
						if gridpopupmenu.isItemChecked(menuid):  # 選択項目にチェックが入っている時。
							self.dialogframe.close(True)
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
class MenuListener(unohelper.Base, XMenuListener):
	def __init__(self):
		self.args = None
	def itemHighlighted(self, menuevent):
		pass
	def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。どこのコントロールのメニューかを知る方法はない。
		menuid = menuevent.MenuId  # メニューIDを取得。1から始まる。
		gridpopupmenu = menuevent.Source
		itemtext = gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
		if itemtext.startswith("オプション表示"):		
			dialogwindow, windowlistener = self.args
			dummy, optioncontrolcontainer = windowlistener.args
			dialogwindowsize = dialogwindow.getSize()
			optioncontrolcontainersize = optioncontrolcontainer.getSize()		
			if gridpopupmenu.isItemChecked(menuid):  # 選択項目にチェックが入った時。
				windowlistener.option = True  # オプションコントロールダイアログを表示させるフラグを立てる。
				diff_width = optioncontrolcontainersize.Width - dialogwindowsize.Width  # オプションコントロールコンテナ幅とコンテナウィンドウ幅の差。
				diff_width = 0 if diff_width<0 else diff_width  # オプションコントロールコンテナ幅よりコンテナウィンドウ幅が大きい時は幅の調整をしない。
				diff_height = optioncontrolcontainersize.Height  # オプションコントロールコンテナの高さを追加する。
				createApplyDiff(diff_width, diff_height)(dialogwindow, PosSize.SIZE)  # コンテナウィンドウの大きさを変更。
			else:
				windowlistener.option = False  # オプションコントロールダイアログを表示させるフラグを倒す。
				diff_height = -optioncontrolcontainersize.Height  # オプションコントロールコンテナの高さを減らす。
				createApplyDiff(0, diff_height)(dialogwindow, PosSize.HEIGHT)  # コンテナウィンドウの大きさを変更。	
	def itemActivated(self, menuevent):
		pass
	def itemDeactivated(self, menuevent):
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
class WindowListener(unohelper.Base, XWindowListener):
	def __init__(self, *args):
		self.args = args
		self.option = False  # optioncontrolcontainerを表示しているかのフラグ。
	def windowResized(self, windowevent):
		controlcontainer, optioncontrolcontainer = self.args
		if self.option:  # optioncontrolcontainerを表示している時。
			optioncontrolcontainer.setVisible(True)
			newwidth, newheight = windowevent.Width, windowevent.Height
			controlcontainerheight = newheight - optioncontrolcontainer.getSize().Height  # オプションコントロールコンテナの高さを除いた高さを取得。
			optioncontrolcontainer.setPosSize(0, controlcontainerheight, newwidth, 0, PosSize.Y+PosSize.WIDTH)
			controlcontainer.setPosSize(0, 0, newwidth, controlcontainerheight, PosSize.SIZE)
		else:
			optioncontrolcontainer.setVisible(False)
			controlcontainer.setPosSize(0, 0, windowevent.Width, windowevent.Height, PosSize.SIZE)
	def windowMoved(self, windowevent):
		pass
	def windowShown(self, eventobject):
		pass
	def windowHidden(self, eventobject):
		pass
	def disposing(self, eventobject):
		pass
class ControlContainerWindowListener(unohelper.Base, XWindowListener):
	def __init__(self, controlcontainer):
		size = controlcontainer.getSize()
		self.oldwidth, self.oldheight = size.Width, size.Height  # 次の変更前の値として取得。		
		self.controlcontainer = controlcontainer
	def windowResized(self, windowevent):
		newwidth, newheight = windowevent.Width, windowevent.Height
		gridcontrol1 = self.controlcontainer.getControl("Grid1")
		diff_width = newwidth - self.oldwidth  # 幅変化分
		diff_height = newheight - self.oldheight  # 高さ変化分		
		createApplyDiff(diff_width, diff_height)(gridcontrol1, PosSize.SIZE)  # コントロールの位置と大きさを変更		
		self.oldwidth, self.oldheight = newwidth, newheight  # 次の変更前の値として取得。
	def windowMoved(self, windowevent):
		pass
	def windowShown(self, eventobject):
		pass
	def windowHidden(self, eventobject):
		pass
	def disposing(self, eventobject):
		pass
def createApplyDiff(diff_width, diff_height):		
	def applyDiff(control, possize):  # 第2引数でウィンドウサイズの変化分のみ適用するPosSizeを指定。
		rectangle = control.getPosSize()  # 変更前のコントロールの位置大きさを取得。
		control.setPosSize(rectangle.X+diff_width, rectangle.Y+diff_height, rectangle.Width+diff_width, rectangle.Height+diff_height, possize)  # Flagsで変更する値のみ指定。変更しない値は0(でもなんでもよいはず)。
	return applyDiff	
class OptionControlContainerWindowListener(unohelper.Base, XWindowListener):
	def __init__(self, optioncontrolcontainer):
		self.oldwidth = optioncontrolcontainer.getSize().Width  # 次の変更前の値として取得。		
		self.optioncontrolcontainer = optioncontrolcontainer
	def windowResized(self, windowevent): # ウィンドウの大きさの変更に合わせてコントロールの位置と大きさを変更。Yと幅のみ変更。
		optioncontrolcontainer = self.optioncontrolcontainer
		newwidth = windowevent.Width
		checkboxcontrol1 = optioncontrolcontainer.getControl("CheckBox1")
		buttoncontrol1 = optioncontrolcontainer.getControl("Button1")
		buttoncontrol3 = optioncontrolcontainer.getControl("Button3")
		minwidth = checkboxcontrol1.getPosSize().Width + buttoncontrol1.getPosSize().Width + buttoncontrol3.getPosSize().Width + 6  # 最低幅を取得。
		if newwidth<minwidth:  # 変更後のコントロールコンテナの幅を取得。サイズ下限より小さい時は下限値とする。
			newwidth = minwidth
		diff_width = newwidth - self.oldwidth  # 幅変化分
		applyDiff = createApplyDiff(diff_width, 0)  # コントロールの位置と大きさを変更する関数を取得。
		applyDiff(optioncontrolcontainer.getControl("Edit1"), PosSize.WIDTH)	
		applyDiff(buttoncontrol3, PosSize.X)
		applyDiff(optioncontrolcontainer.getControl("Button4"), PosSize.X)
		self.oldwidth = newwidth  # 次の変更前の値として取得。
	def windowMoved(self, windowevent):
		pass
	def windowShown(self, eventobject):
		pass
	def windowHidden(self, eventobject):
		pass
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
