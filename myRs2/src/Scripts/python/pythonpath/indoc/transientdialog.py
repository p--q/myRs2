#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import dialogcommons, staticdialog  # staticdialogのオブジェクトを複数流用している。
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import MenuItemStyle, MouseButton, PopupMenuDirection, PosSize  # 定数
from com.sun.star.awt import MenuEvent, Rectangle  # Struct
from com.sun.star.beans import NamedValue  # Struct
from com.sun.star.util import XCloseListener
def createDialog(xscriptcontext, dialogtitle, defaultrows, outputcolumn=None, *, enhancedmouseevent=None, fixedtxt=None, callback=None):  # dialogtitleはダイアログのデータ保存名に使うのでユニークでないといけない。defaultrowsはグリッドコントロールのデフォルトデータ。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	doc = xscriptcontext.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。  
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	containerwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	maTopx = dialogcommons.createConverters(containerwindow)  # ma単位をピクセルに変換する関数を取得。
	m = 2  # コントロール間の間隔。
	h = 12  # コントロールの高さ
	XWidth, YHeight = dialogcommons.XWidth, dialogcommons.YHeight
	gridprops = {"PositionX": 0, "PositionY": 0, "Width": 50, "Height": 50, "ShowRowHeader": False, "ShowColumnHeader": False}  # グリッドコントロールのプロパティ。
	controlcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(gridprops), "Height": YHeight(gridprops), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
	controlcontainer, addControl = dialogcommons.controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。		
	menulistener = staticdialog.MenuListener()  # コンテクストメニューにつけるリスナー。
	items = ("セル入力で閉じる", MenuItemStyle.CHECKABLE+MenuItemStyle.AUTOCHECK, {"checkItem": False}),\
			("オプション表示", MenuItemStyle.CHECKABLE+MenuItemStyle.AUTOCHECK, {"checkItem": False})  # グリッドコントロールのコンテクストメニュー。XMenuListenerのmenuevent.MenuIdでコードを実行する。
	gridpopupmenu = dialogcommons.menuCreator(ctx, smgr)("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。 
	args = gridpopupmenu, xscriptcontext, outputcolumn, fixedtxt, callback  # gridpopupmenuは先頭でないといけない。
	mouselistener = MouseListener(args)
	gridcontrol1 = addControl("Grid", gridprops, {"addMouseListener": mouselistener})  # グリッドコントロールの取得。
	gridmodel = gridcontrol1.getModel()  # グリッドコントロールモデルの取得。
	gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
	gridcolumn.addColumn(gridcolumn.createColumn())  # 列を追加。
	griddatamodel = gridmodel.getPropertyValue("GridDataModel")  # GridDataModel
	if defaultrows is not None:  # デフォルトdatarowsがあるときデフォルトデータを使用。	
		datarows = [i if isinstance(i, (list, tuple)) else (i,) for i in defaultrows]  # defaultrowsの要素をリストかタプルでなければタプルに変換する。
		griddatamodel.addRows(("",)*len(datarows), datarows)  # グリッドに行を追加。	
	else:
		datarows = []  # Noneのままではあとで処理できないので空リストを入れる。
	controlcontainerwindowlistener = staticdialog.ControlContainerWindowListener(controlcontainer)		
	controlcontainer.addWindowListener(controlcontainerwindowlistener)  # コントロールコンテナの大きさを変更するとグリッドコントロールの大きさも変更するようにする。
	checkboxprops1 = {"PositionX": 0, "PositionY": m, "Width": 46, "Height": h, "Label": "~セルに追記", "State": 0} # セルに追記はデフォルトでは無効。
	checkboxprops2 = {"PositionX": 0, "PositionY": YHeight(checkboxprops1, 4), "Width": 46, "Height": h, "Label": "~サイズ復元", "State": 1}  # サイズ復元はデフォルトでは有効。		
	optioncontrolcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(checkboxprops2), "Height": YHeight(checkboxprops2, 2), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
	optioncontrolcontainer, optionaddControl = dialogcommons.controlcontainerMaCreator(ctx, smgr, maTopx, optioncontrolcontainerprops)  # コントロールコンテナの作成。		
	checkboxcontrol1 = optionaddControl("CheckBox", checkboxprops1)
	checkboxcontrol2 = optionaddControl("CheckBox", checkboxprops2)  
	mouselistener.optioncontrolcontainer = optioncontrolcontainer
	rectangle = controlcontainer.getPosSize()  # コントロールコンテナのRectangle Structを取得。px単位。
	controller = doc.getCurrentController()  # 現在のコントローラを取得。
	if enhancedmouseevent is None:
		visibleareaonscreen = controller.getPropertyValue("VisibleAreaOnScreen")
		rectangle.X, rectangle.Y = visibleareaonscreen.X, visibleareaonscreen.Y
	else:
		dialogpoint = dialogcommons.getDialogPoint(doc, enhancedmouseevent)  # クリックした位置のメニューバーの高さ分下の位置を取得。単位ピクセル。一部しか表示されていないセルのときはNoneが返る。
		if not dialogpoint:  # クリックした位置が取得出来なかった時は何もしない。
			return
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
	frameactionlistener = dialogcommons.FrameActionListener()  # FrameActionListener。フレームがアクティブでなくなった時に閉じるため。
	dialogframe.addFrameActionListener(frameactionlistener)  # FrameActionListenerをダイアログフレームに追加。
	controlcontainer.setVisible(True)  # コントロールの表示。
	dialogwindow.setVisible(True) # ウィンドウの表示。これ以降WindowListenerが発火する。
	windowlistener = staticdialog.WindowListener(controlcontainer, optioncontrolcontainer) # コンテナウィンドウからコントロールコンテナを取得する方法はないはずなので、ここで渡す。WindowListenerはsetVisible(True)で呼び出される。
	dialogwindow.addWindowListener(windowlistener) # コンテナウィンドウにリスナーを追加する。
	menulistener.args = dialogwindow, windowlistener
	dialogstate = dialogcommons.getSavedData(doc, "dialogstate_{}".format(dialogtitle))  # 保存データを取得。optioncontrolcontainerの表示状態は常にFalseなので保存されていない。
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
	args = doc, dialogwindow, windowlistener, mouselistener, menulistener, controlcontainerwindowlistener
	dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。	
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログのリスナー削除用。
	def __init__(self, args):
		self.args = args
	def queryClosing(self, eventobject, getsownership):  # ノンモダルダイアログを閉じる時に発火。
		dialogframe = eventobject.Source
		doc, dialogwindow, windowlistener, mouselistener, menulistener, controlcontainerwindowlistener = self.args
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
		dialogcommons.saveData(doc, "dialogstate_{}".format(dialogtitle), dialogstate)  # ダイアログの状態を保存。
		gridpopupmenu.removeMenuListener(menulistener)
		controlcontainer.getControl("Grid1").removeMouseListener(mouselistener)
		controlcontainer.removeWindowListener(controlcontainerwindowlistener)
		dialogwindow.removeWindowListener(windowlistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		pass
class MouseListener(unohelper.Base, XMouseListener):  
	def __init__(self, args): 	
		self.gridpopupmenu, *self.args = args  # gridpopupmenuはCloseListenerで使うので、別にする。
		self.optioncontrolcontainer = None
		self.dialogframe = None
	def mousePressed(self, mouseevent):  # グリッドコントロールをクリックした時。コントロールモデルにはNameプロパティはない。
		xscriptcontext, outputcolumn, fixedtxt, callback = self.args
		gridcontrol = mouseevent.Source  # グリッドコントロールを取得。
		optioncontrolcontainer = self.optioncontrolcontainer
		if mouseevent.Buttons==MouseButton.LEFT:
			if mouseevent.ClickCount==2:  # ダブルクリックの時。
				doc = xscriptcontext.getDocument()
				selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
				if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
					griddata = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
					j = gridcontrol.getCurrentRow()
					if j<0:  # 選択行がない時は-1が返る。
						return
					celladdress = selection.getCellAddress()
					r, c = celladdress.Row, celladdress.Column
					if outputcolumn is not None:  # 出力する列が指定されている時。
						c = outputcolumn  # 同じ行の指定された列のセルに入力するようにする。
					controller = doc.getCurrentController()  # 現在のコントローラを取得。	
					sheet = controller.getActiveSheet()
					rowdata = griddata.getRowData(j)  # グリッドコントロールで選択している行のすべての列をタプルで取得。
					if fixedtxt is None:
						fixedtxt = rowdata[0]
					if optioncontrolcontainer.getControl("CheckBox1").getState():  # セルに追記、にチェックがある時。グリッドコントロールは1列と決めつけて処理する。
						sheet[r, c].setString("".join([selection.getString(), fixedtxt]))  # セルに追記する。
					else:
						sheet[r, c].setString(fixedtxt)  # セルに代入。
						
# 					import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)		
						
						
					if callback is not None:  # コールバック関数が与えられている時。
						callback(mouseevent, xscriptcontext, fixedtxt)				
						
					
										
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
