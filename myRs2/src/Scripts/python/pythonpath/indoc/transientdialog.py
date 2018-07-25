#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import dialogcommons
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
# SHEETNAME = "config"  # データを保存するシート名。
def createDialog(xscriptcontext, dialogtitle, defaultrows, outputcolumn=None, *, callback=None):  # dialogtitleはダイアログのデータ保存名に使うのでユニークでないといけない。defaultrowsはグリッドコントロールのデフォルトデータ。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	doc = xscriptcontext.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。  
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	containerwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	maTopx = dialogcommons.createConverters(containerwindow)  # ma単位をピクセルに変換する関数を取得。
	m = 2  # コントロール間の間隔。
	h = 12  # コントロールの高さ
	XWidth, YHeight = dialogcommons.XWidth, dialogcommons.YHeight
	gridprops = {"PositionX": 0, "PositionY": 0, "Width": 50, "Height": 50, "ShowRowHeader": False, "ShowColumnHeader": False, "SelectionModel": MULTI}  # グリッドコントロールのプロパティ。
	controlcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(gridprops), "Height": YHeight(gridprops), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
	controlcontainer, addControl = dialogcommons.controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。		
	menulistener = MenuListener()  # コンテクストメニューにつけるリスナー。
	items = ("セル入力で閉じる", MenuItemStyle.CHECKABLE+MenuItemStyle.AUTOCHECK, {"checkItem": False}),\
			("オプション表示", MenuItemStyle.CHECKABLE+MenuItemStyle.AUTOCHECK, {"checkItem": False})  # グリッドコントロールのコンテクストメニュー。XMenuListenerのmenuevent.MenuIdでコードを実行する。
	gridpopupmenu = dialogcommons.menuCreator(ctx, smgr)("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。 
	mouselistener = MouseListener(xscriptcontext, gridpopupmenu)
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
	controlcontainerwindowlistener = ControlContainerWindowListener(controlcontainer)		
	controlcontainer.addWindowListener(controlcontainerwindowlistener)  # コントロールコンテナの大きさを変更するとグリッドコントロールの大きさも変更するようにする。
	checkboxprops1 = {"PositionX": 0, "PositionY": m, "Width": 46, "Height": h, "Label": "~セルに追記", "State": 0} # セルに追記はデフォルトでは無効。
	checkboxprops2 = {"PositionX": 0, "PositionY": YHeight(checkboxprops1, 4), "Width": 46, "Height": h, "Label": "~サイズ復元", "State": 1}  # サイズ復元はデフォルトでは有効。		
	optioncontrolcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(checkboxprops2), "Height": YHeight(checkboxprops2, 2), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
	optioncontrolcontainer, optionaddControl = dialogcommons.controlcontainerMaCreator(ctx, smgr, maTopx, optioncontrolcontainerprops)  # コントロールコンテナの作成。		
	checkboxcontrol1 = optionaddControl("CheckBox", checkboxprops1)
	checkboxcontrol2 = optionaddControl("CheckBox", checkboxprops2)  
# 	optioncontrolcontainerwindowlistener = OptionControlContainerWindowListener(optioncontrolcontainer)		
# 	optioncontrolcontainer.addWindowListener(optioncontrolcontainerwindowlistener)  # コントロールコンテナの大きさを変更するとグリッドコントロールの大きさも変更するようにする。
	mouselistener.optioncontrolcontainer = optioncontrolcontainer
	rectangle = controlcontainer.getPosSize()  # コントロールコンテナのRectangle Structを取得。px単位。
	controller = doc.getCurrentController()  # 現在のコントローラを取得。
	visibleareaonscreen = controller.getPropertyValue("VisibleAreaOnScreen")
	rectangle.X, rectangle.Y = visibleareaonscreen.X, visibleareaonscreen.Y  # 
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
	windowlistener = WindowListener(controlcontainer, optioncontrolcontainer) # コンテナウィンドウからコントロールコンテナを取得する方法はないはずなので、ここで渡す。WindowListenerはsetVisible(True)で呼び出される。
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
# 		optioncontrolcontainer.removeWindowListener(optioncontrolcontainerwindowlistener)
		dialogwindow.removeWindowListener(windowlistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		pass
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
# 			if mouseevent.ClickCount==1:  # シングルクリックの時。
# 				selectedrowindexes = dialogcommons.getSelectedRowIndexes(gridcontrol)
# 				if not selectedrowindexes:  # 選択行がない時(選択行を削除した時)。
# 					return  # 何もしない		
# 				upbuttoncontrol = optioncontrolcontainer.getControl("Button1")
# 				downbuttoncontrol = optioncontrolcontainer.getControl("Button2")
# 				insertbuttoncontrol = optioncontrolcontainer.getControl("Button3")
# 				upbuttoncontrol.setEnable(True)  # まず全てのボタンを有効にする。
# 				downbuttoncontrol.setEnable(True)
# 				insertbuttoncontrol.setEnable(True)
# 				if selectedrowindexes[0]==0:  # 先頭行が選択されている時。
# 					upbuttoncontrol.setEnable(False)  # 上へボタンを無効にする。
# 				griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")	
# 				if selectedrowindexes[-1]==griddatamodel.RowCount-1:  # 最終行が選択されている時。
# 					downbuttoncontrol.setEnable(False)  # 下へボタンを無効にする。
# 				indexcount = len(selectedrowindexes)  # 選択行数を取得。
# 				if indexcount>1:  # 複数行を選択している時。
# 					insertbuttoncontrol.setEnable(False)  # 行挿入ボタンを無効にする。
# 					if indexcount!=selectedrowindexes[-1]-selectedrowindexes[0]+1:  # 連続した行でない時。
# 						upbuttoncontrol.setEnable(False)  # 上へボタンを無効にする。
# 						downbuttoncontrol.setEnable(False)  # 下へボタンを無効にする。
# 				rowdata = griddatamodel.getRowData(selectedrowindexes[0])  # 選択行の最初の行のデータを取得。
# 				optioncontrolcontainer.getControl("Edit1").setText(rowdata[0])  # テキストボックスに選択行の初行の文字列を代入。
# 				if griddatamodel.RowCount==1:  # 1行しかない時はまた発火できるように選択を外す。
# 					gridcontrol.deselectRow(0)  # 選択行の選択を外す。選択していない行を指定すると永遠ループになる。
			if mouseevent.ClickCount==2:  # ダブルクリックの時。
				doc = self.xscriptcontext.getDocument()
				selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
				if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
					griddata = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
					j = gridcontrol.getCurrentRow()
					if j<0:  # 選択行がない時は-1が返る。
						return
					rowdata = griddata.getRowData(j)  # グリッドコントロールで選択している行のすべての列をタプルで取得。
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
				dialogcommons.createApplyDiff(diff_width, diff_height)(dialogwindow, PosSize.SIZE)  # コンテナウィンドウの大きさを変更。
			else:
				windowlistener.option = False  # オプションコントロールダイアログを表示させるフラグを倒す。
				diff_height = -optioncontrolcontainersize.Height  # オプションコントロールコンテナの高さを減らす。
				dialogcommons.createApplyDiff(0, diff_height)(dialogwindow, PosSize.HEIGHT)  # コンテナウィンドウの大きさを変更。	
	def itemActivated(self, menuevent):
		pass
	def itemDeactivated(self, menuevent):
		pass   
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
		dialogcommons.createApplyDiff(diff_width, diff_height)(gridcontrol1, PosSize.SIZE)  # コントロールの位置と大きさを変更		
		self.oldwidth, self.oldheight = newwidth, newheight  # 次の変更前の値として取得。
	def windowMoved(self, windowevent):
		pass
	def windowShown(self, eventobject):
		pass
	def windowHidden(self, eventobject):
		pass
	def disposing(self, eventobject):
		pass	
# class OptionControlContainerWindowListener(unohelper.Base, XWindowListener):
# 	def __init__(self, optioncontrolcontainer):
# 		self.oldwidth = optioncontrolcontainer.getSize().Width  # 次の変更前の値として取得。		
# 		self.optioncontrolcontainer = optioncontrolcontainer
# 	def windowResized(self, windowevent): # ウィンドウの大きさの変更に合わせてコントロールの位置と大きさを変更。Yと幅のみ変更。
# 		optioncontrolcontainer = self.optioncontrolcontainer
# 		newwidth = windowevent.Width
# 		checkboxcontrol1 = optioncontrolcontainer.getControl("CheckBox1")
# # 		buttoncontrol1 = optioncontrolcontainer.getControl("Button1")
# # 		buttoncontrol3 = optioncontrolcontainer.getControl("Button3")
# # 		minwidth = checkboxcontrol1.getPosSize().Width + buttoncontrol1.getPosSize().Width + buttoncontrol3.getPosSize().Width + 6  # 最低幅を取得。
# # 		minwidth = checkboxcontrol1.getPosSize().Width + 6  # 最低幅を取得。
# # 		if newwidth<minwidth:  # 変更後のコントロールコンテナの幅を取得。サイズ下限より小さい時は下限値とする。
# # 			newwidth = minwidth
# # 		diff_width = newwidth - self.oldwidth  # 幅変化分
# # 		applyDiff = dialogcommons.createApplyDiff(diff_width, 0)  # コントロールの位置と大きさを変更する関数を取得。
# # 		applyDiff(optioncontrolcontainer.getControl("Edit1"), PosSize.WIDTH)	
# # 		applyDiff(buttoncontrol3, PosSize.X)
# # 		applyDiff(optioncontrolcontainer.getControl("Button4"), PosSize.X)
# 		self.oldwidth = newwidth  # 次の変更前の値として取得。
# 	def windowMoved(self, windowevent):
# 		pass
# 	def windowShown(self, eventobject):
# 		pass
# 	def windowHidden(self, eventobject):
# 		pass
# 	def disposing(self, eventobject):
# 		pass
