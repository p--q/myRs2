#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import dialogcommons
from com.sun.star.accessibility import AccessibleRole  # 定数
from com.sun.star.awt import XActionListener, XMenuListener, XMouseListener, XWindowListener, XTextListener, XItemListener
from com.sun.star.awt import MessageBoxButtons, MessageBoxResults, MouseButton, PopupMenuDirection, PosSize, ScrollBarOrientation  # 定数
from com.sun.star.awt import Point, Rectangle, Selection  # Struct
from com.sun.star.awt.MessageBoxType import QUERYBOX  # enum
from com.sun.star.beans import NamedValue  # Struct
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.util import XCloseListener
from com.sun.star.view.SelectionType import MULTI  # enum 
from com.sun.star.lang import Locale  # Struct
DATAROWS = []  # グリッドコントロールのデータ行、タプルのタプルやリストのタプルやリストのリスト、の可能性がある。複数クラスからアクセスするのでグローバルにしないといけない。
def createDialog(xscriptcontext, enhancedmouseevent, dialogtitle, defaultrows=None):  # dialogtitleはダイアログのデータ保存名に使うのでユニークでないといけない。defaultrowsはグリッドコントロールのデフォルトデータ。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	doc = xscriptcontext.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	*dialogpoint, vscrollbar = getDialogPoint(doc, enhancedmouseevent)  # X座標は固定列の右に表示する。クリックした位置のメニューバーの高さ分下の位置を取得。単位ピクセル。一部しか表示されていないセルのときはNoneが返る。
	if not dialogpoint:  # クリックした位置が取得出来なかった時は何もしない。
		return
	txt = doc.getCurrentSelection().getString()  # 選択セルの文字列を取得。
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	containerwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	toolkit = containerwindow.getToolkit()  # ピアからツールキットを取得。  		
	m = 2  # コントロール間の間隔。
	h = 12  # コントロール間の高さ。
	XWidth, YHeight = dialogcommons.XWidth, dialogcommons.YHeight
	gridprops = {"PositionX": 0, "PositionY": 0, "Height": 50, "ShowRowHeader": False, "ShowColumnHeader": False, "SelectionModel": MULTI}  # グリッドコントロールのプロパティ。
	textboxprops = {"PositionX": 0, "PositionY": YHeight(gridprops, 2), "Height": h, "Text": txt}  # テクストボックスコントロールのプロパティ。
	checkboxprops = {"PositionY": YHeight(textboxprops, m), "Height": h, "Tabstop": False}  # チェックボックスコントロールのプロパティ。
	checkboxprops1, checkboxprops2 = [checkboxprops.copy() for dummy in range(2)]
	checkboxprops1.update({"PositionX": 0, "Width": 42, "Label": "~サイズ復元", "State": 1})  # サイズ復元はデフォルトでは有効。
	checkboxprops2.update({"PositionX": XWidth(checkboxprops1), "Width": 38, "Label": "~逐次検索", "State": 0})  # 逐次検索はデフォルトでは無効。
	buttonprops = {"PositionX": XWidth(checkboxprops2), "PositionY": YHeight(textboxprops, 4), "Width": 30, "Height": h+2, "Label": "Enter"}  # ボタンのプロパティ。PushButtonTypeの値はEnumではエラーになる。VerticalAlignではtextboxと高さが揃わない。
	gridprops["Width"] = textboxprops["Width"] = XWidth(buttonprops)
	controlcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(gridprops), "Height": YHeight(buttonprops, m), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
	maTopx = dialogcommons.createConverters(containerwindow)  # ma単位をピクセルに変換する関数を取得。
	controlcontainer, addControl = dialogcommons.controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。		
	mouselistener = MouseListener(xscriptcontext, vscrollbar)
	menulistener = MenuListener(controlcontainer)  # コンテクストメニューにつけるリスナー。
	actionlistener = ActionListener(xscriptcontext, vscrollbar)  # ボタンコントロールにつけるリスナー。
	items = ("選択行を削除", 0, {"setCommand": "delete"}),\
			("全行を削除", 0, {"setCommand": "deleteall"})  # グリッドコントロールにつける右クリックメニュー。
	mouselistener.gridpopupmenu = dialogcommons.menuCreator(ctx, smgr)("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。 
	gridcontrol1 = addControl("Grid", gridprops, {"addMouseListener": mouselistener})  # グリッドコントロールの取得。gridは他のコントロールの設定に使うのでコピーを渡す。
	gridmodel = gridcontrol1.getModel()  # グリッドコントロールモデルの取得。
	gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
	gridcolumn.addColumn(gridcolumn.createColumn())  # 列を追加。
	griddatamodel = gridmodel.getPropertyValue("GridDataModel")  # GridDataModel
	datarows = dialogcommons.getSavedData(doc, "GridDatarows_{}".format(dialogtitle))  # グリッドコントロールの行をconfigシートのragenameから取得する。	
	if datarows is None and defaultrows is not None:  # 履歴がなくデフォルトdatarowsがあるときデフォルトデータを使用。
		datarows = [i if isinstance(i, (list, tuple)) else (i,) for i in defaultrows]  # defaultrowsの要素をリストかタプルでなければタプルに変換する。
	if datarows:  # 行のリストが取得出来た時。
		griddatamodel.addRows(("",)*len(datarows), datarows)  # グリッドに行を追加。
		global DATAROWS
		DATAROWS = datarows  # グリッドのデータをDATAROWSに反映。
	textlistener = TextListener(xscriptcontext)	
	addControl("Edit", textboxprops, {"addTextListener": textlistener})  
	checkboxcontrol1 = addControl("CheckBox", checkboxprops1)  
	itemlistener = ItemListener(textlistener)
	checkboxcontrol2 = addControl("CheckBox", checkboxprops2, {"addItemListener": itemlistener}) 
	addControl("Button", buttonprops, {"addActionListener": actionlistener, "setActionCommand": "enter"})  
	rectangle = controlcontainer.getPosSize()  # コントロールコンテナのRectangle Structを取得。px単位。
	rectangle.X, rectangle.Y = dialogpoint  # クリックした位置を取得。ウィンドウタイトルを含めない座標。
	taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
	args = NamedValue("PosSize", rectangle), NamedValue("FrameName", "controldialog")  # , NamedValue("MakeVisible", True)  # TaskCreatorで作成するフレームのコンテナウィンドウのプロパティ。
	dialogframe = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。
	dialogwindow = dialogframe.getContainerWindow()  # ダイアログのコンテナウィンドウを取得。
	dialogframe.setTitle(dialogtitle)  # フレームのタイトルを設定。
	docframe.getFrames().append(dialogframe) # 新しく作ったフレームを既存のフレームの階層に追加する。		
	controlcontainer.createPeer(toolkit, dialogwindow) # ウィンドウにコントロールを描画。 
	dialogframe.addFrameActionListener(dialogcommons.FrameActionListener())  # FrameActionListenerをダイアログフレームに追加。リスナーはフレームを閉じる時に削除するようにしている。
	windowlistener = WindowListener(controlcontainer)
	dialogwindow.addWindowListener(windowlistener) # setVisible(True)でも呼び出される。
	args = doc, controlcontainer, actionlistener, dialogwindow, windowlistener, mouselistener, menulistener, textlistener, itemlistener
	dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。		
	controlcontainer.setVisible(True)  # コントロールの表示。
	dialogwindow.setVisible(True) # ウィンドウの表示。ここでウィンドウリスナーが発火する。
	dialogstate = dialogcommons.getSavedData(doc, "dialogstate_{}".format(dialogtitle))  # 保存データを取得。
	if dialogstate is not None:  # 保存してあるダイアログの状態がある時。
		checkbox1sate = dialogstate.get("CheckBox1sate")  # サイズ復元、チェックボックス。キーがなければNoneが返る。	
		if checkbox1sate is not None:  # サイズ復元、が保存されている時。
			if checkbox1sate:  # サイズ復元がチェックされている時。
				dialogwindow.setPosSize(0, 0, dialogstate["Width"], dialogstate["Height"], PosSize.SIZE)  # ウィンドウサイズを復元。
			checkboxcontrol1.setState(checkbox1sate)  # 状態を復元。	
		checkbox2sate = dialogstate.get("CheckBox2sate")  # 逐語検索、チェックボックス。			
		if checkbox2sate is not None:  # 逐語検索、が保存されている時。
			if checkbox2sate:  # チェックされている時逐次検索を有効にする。	
				refreshRows(gridcontrol1, [i for i in DATAROWS if i[0].startswith(txt)])  # txtで始まっている行だけに絞る。txtが空文字の時はすべてTrueになる。
			checkboxcontrol2.setState(checkbox2sate)  # itemlistenerは発火しない。			
	scrollDown(gridcontrol1)		
class ItemListener(unohelper.Base, XItemListener):
	def __init__(self, textlistener):
		self.textlistener = textlistener
	def itemStateChanged(self, itemevent):  
		checkboxcontrol2 = itemevent.Source
		gridcontrol1 = checkboxcontrol2.getContext().getControl("Grid1")
		if checkboxcontrol2.getState():
			txt = checkboxcontrol2.getContext().getControl("Edit1").getText()
			refreshRows(gridcontrol1, [i for i in DATAROWS if i[0].startswith(txt)])  # txtで始まっている行だけに絞る。txtが空文字の時はすべてTrueになる。
		else:
			refreshRows(gridcontrol1, DATAROWS)
			scrollDown(gridcontrol1)	
	def disposing(self, eventobject):
		pass
def refreshRows(gridcontrol, datarows):	
	griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。	
	griddatamodel.removeAllRows()  # グリッドコントロールの行を全削除。
	if datarows:  # データ行がある時。
		griddatamodel.addRows(("",)*len(datarows), datarows)  # グリッドに行を追加。
class TextListener(unohelper.Base, XTextListener):
	def __init__(self, xscriptcontext):
		self.transliteration = fullwidth_halfwidth(xscriptcontext)
		self.history = ""  # 前値を保存する。
	def textChanged(self, textevent):  # 複数回呼ばれるので前値との比較が必要。
		editcontrol1 = textevent.Source
		txt = editcontrol1.getText()
		if txt!=self.history:  # 前値から変化する時のみ。
			txt = self.transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換
			editcontrol1.removeTextListener(self)
			editcontrol1.setText(txt)  # 永久ループになるのでTextListenerを発火しないようにしておかないといけない。
			editcontrol1.addTextListener(self)
			controlcontainer = editcontrol1.getContext()
			if controlcontainer.getControl("CheckBox2").getState():  # 逐次検索が有効になっている時。
				gridcontrol1 = controlcontainer.getControl("Grid1")
				datarows = [i for i in DATAROWS if i[0].startswith(txt)]  # 逐語抽出した行のリスト。
				if len(datarows)==1:  # 行が一行だけになる時。	
					selectedrowindex = gridcontrol1.getCurrentRow()  # 選択行インデックスを取得。
					if selectedrowindex>0:  # 選択行インデックスが1以上の時。
						griddatamodel = gridcontrol1.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
						if selectedrowindex+1<griddatamodel.RowCount:  # 選択行より後に行がある時。
							[griddatamodel.removeRow(i) for i in range(selectedrowindex+1, griddatamodel.RowCount)[::-1]]  # 選択行の次から最後までを最後から削除する。
						return  # 選択行より前の行の削除は諦める。選択行より上の行を削除するとグリッドコントロール以外マウスに反応しなくなるので。ソートして一番上に持ってきてもダメ。
				refreshRows(gridcontrol1, datarows)
			self.history = txt	
	def disposing(self, eventobject):
		pass
def fullwidth_halfwidth(xscriptcontext):
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。					
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
	transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))  # 全角を半角に変換するモジュール。
	return transliteration
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログのリスナー削除用。
	def __init__(self, args):
		self.args = args
	def queryClosing(self, eventobject, getsownership):  # ノンモダルダイアログを閉じる時に発火。
		dialogframe = eventobject.Source
		doc, controlcontainer, actionlistener, dialogwindow, windowlistener, mouselistener, menulistener, textlistener, itemlistener = self.args
		size = controlcontainer.getSize()
		checkboxcontrol2 = controlcontainer.getControl("CheckBox2")
		checkboxcontrol2.removeItemListener(itemlistener)			
		dialogstate = {"CheckBox1sate": controlcontainer.getControl("CheckBox1").getState(),\
					"CheckBox2sate": checkboxcontrol2.getState(),\
					"Width": size.Width,\
					"Height": size.Height}  # チェックボックスの状態と大きさを取得。
		dialogtitle = dialogframe.getTitle()
		dialogcommons.saveData(doc, "dialogstate_{}".format(dialogtitle), dialogstate)
		gridcontrol1 = controlcontainer.getControl("Grid1")
		dialogcommons.saveData(doc, "GridDatarows_{}".format(dialogtitle), DATAROWS)
		mouselistener.gridpopupmenu.removeMenuListener(menulistener)
		gridcontrol1.removeMouseListener(mouselistener)
		controlcontainer.getControl("Button1").removeActionListener(actionlistener)
		controlcontainer.getControl("Edit1").removeTextListener(textlistener)
		dialogwindow.removeWindowListener(windowlistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		pass
class ActionListener(unohelper.Base, XActionListener):
	def __init__(self, xscriptcontext, vscrollbar):
		self.xscriptcontext = xscriptcontext
		self.transliteration = fullwidth_halfwidth(xscriptcontext)
		self.vscrollbar = vscrollbar
	def actionPerformed(self, actionevent):  
		cmd = actionevent.ActionCommand
		if cmd=="enter":
			doc = self.xscriptcontext.getDocument()  
			selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
				controlcontainer = actionevent.Source.getContext()
				edit1 = controlcontainer.getControl("Edit1")  # テキストボックスコントロールを取得。
				txt = edit1.getText()  # テキストボックスコントロールの文字列を取得。
				if txt:  # テキストボックスコントロールに文字列がある時。
					txt = self.transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換
					datarows = DATAROWS
					if datarows:  # すでにグリッドコントロールにデータがある時。
						lastindex = len(datarows) - 1  # 最終インデックスを取得。
						[datarows.pop(lastindex-i) for i, datarow in enumerate(datarows[::-1]) if txt in datarow]  # txtがある行は後ろから削除する。
					datarows.append((txt,))  # txtの行を追加。
					gridcontrol1 = controlcontainer.getControl("Grid1")
					refreshRows(gridcontrol1, datarows)
					scrollDown(gridcontrol1)  # グリッドコントロールを下までスクロール。
					selection.setString(txt)  # 選択セルに代入。
					global DATAROWS
					DATAROWS = datarows								
				nexttxt = toNextCell(doc, selection, self.vscrollbar)  # 下のセルを選択してその値を返す。シートも1行分スクロールする。	
				edit1.setText(nexttxt)  # テキストボックスコントロールにセルの内容を取得。
				edit1.setFocus()  # テキストボックスコントロールをフォーカスする。
				textlength = len(nexttxt)  # テキストボックスコントロール内の文字列の文字数を取得。
				edit1selection = Selection(Min=textlength, Max=textlength)  # カーソルの位置を最後にする。指定しないと先頭になる。
				edit1.setSelection(edit1selection)  # テクストボックスコントロールのカーソルの位置を変更。ピア作成後でないと反映されない。
	def disposing(self, eventobject):
		pass
class MouseListener(unohelper.Base, XMouseListener):  
	def __init__(self, xscriptcontext, vscrollbar): 	
		self.xscriptcontext = xscriptcontext
		self.gridpopupmenu = None
		self.vscrollbar = vscrollbar
	def mousePressed(self, mouseevent):  # グリッドコントロールをクリックした時。コントロールモデルにはNameプロパティはない。
		gridcontrol = mouseevent.Source  # グリッドコントロールを取得。
		griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
		if mouseevent.Buttons==MouseButton.LEFT:  # 左ボタンクリックの時。
			if mouseevent.ClickCount==1:  # シングルクリックの時。
				selectedrowindexes = dialogcommons.getSelectedRowIndexes(gridcontrol)	 # 選択行のリストを取得。
				if selectedrowindexes:  # 選択行がある時。
					rowdata = griddatamodel.getRowData(selectedrowindexes[0])  # 選択行の最初の行のデータを取得。
					txt = rowdata[0]
				else:
					txt = ""  # 選択行がない時は空文字にする。
				gridcontrol.getContext().getControl("Edit1").setText(txt)  # テキストボックスに選択行の初行の文字列を代入。
			elif mouseevent.ClickCount==2:  # ダブルクリックの時。
				doc = self.xscriptcontext.getDocument()
				selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
				if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
					j = gridcontrol.getCurrentRow()  # 選択行インデックス。負数が返ってくることがある。
					if j<0:  # 負数の時は何もしない。
						return
					rowdata = griddatamodel.getRowData(j)  # グリッドコントロールで選択している行のすべての列をタプルで取得。
					selection.setString(rowdata[0])  # グリッドコントロールは1列と決めつけて、その最初の要素をセルに代入。
					nexttxt = toNextCell(doc, selection, self.vscrollbar)  # 下のセルを選択してその値を返す。シートも1行分スクロールする。
					edit1 = gridcontrol.getContext().getControl("Edit1")  # テキストボックスコントロールを取得。				
					edit1.setText(nexttxt)  # テキストボックスコントロールにセルの内容を取得。			
		elif mouseevent.Buttons==MouseButton.RIGHT:  # 右ボタンクリックの時。mouseevent.PopupTriggerではサブジェクトによってはTrueにならないので使わない。
			rowindex = gridcontrol.getRowAtPoint(mouseevent.X, mouseevent.Y)  # クリックした位置の行インデックスを取得。該当行がない時は-1が返ってくる。
			if rowindex>-1:  # クリックした位置に行が存在する時。
				if not gridcontrol.isRowSelected(rowindex):  # クリックした位置の行が選択状態でない時。
					gridcontrol.deselectAllRows()  # 行の選択状態をすべて解除する。
					gridcontrol.selectRow(rowindex)  # 右クリックしたところの行を選択する。		
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
def toNextCell(doc, selection, vscrollbar):  # 下のセルを選択してその値を返す。シートも1行分スクロールする。
	controller = doc.getCurrentController()  # 現在のコントローラを取得。	
	sheet = controller.getActiveSheet()
	celladdress = selection.getCellAddress()
	nextcell = sheet[celladdress.Row+1, celladdress.Column]  # 下のセルを取得。
	controller.select(nextcell)  # 下のセルを選択。
	newvalue = vscrollbar.getCurrentValue() + 1
	vscrollbar.setCurrentValue(newvalue)  # シートを下にスクロール
	return nextcell.getString()  # 下のセルの文字列を返す。
class MenuListener(unohelper.Base, XMenuListener):
	def __init__(self, controlcontainer):
		self.controlcontainer = controlcontainer
	def itemHighlighted(self, menuevent):
		pass
	def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。どこのコントロールのメニューかを知る方法はない。
		cmd = menuevent.Source.getCommand(menuevent.MenuId)
		controlcontainer = self.controlcontainer
		datarows = list(DATAROWS)
		peer = controlcontainer.getPeer()  # ピアを取得。	
		gridcontrol = controlcontainer.getControl("Grid1")  # グリッドコントロールを取得。
		griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。		
		selectedrowindexes = dialogcommons.getSelectedRowIndexes(gridcontrol)	 # 選択行のリストを取得。
		if not selectedrowindexes:
			return  # 選択行がない時何もしない。
		if cmd=="delete":  # 選択行を削除する。  
			msg = "選択行を削除しますか?"
			msgbox = peer.getToolkit().createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "履歴", msg)
			if msgbox.execute()==MessageBoxResults.YES:		
				if griddatamodel.RowCount==len(datarows):  # グリッドコントロールとDATAROWSの行数が一致している時。
					[datarows.pop(i) for i in selectedrowindexes[::-1]]  # 後ろから選択行を削除。
					refreshRows(gridcontrol, datarows)
				else:
					for i in selectedrowindexes[::-1]:  # 選択した行インデックスを後ろから取得。逐次検索のときはグリッドコントロールとDATAROWSが一致しないので別に処理する。
						d = griddatamodel.getRowData(i)[0]  # タプルが返るのでその先頭の要素を取得。
						datarows = [j for j in datarows if not d in j]  # dが要素にある行を除いて取得。
						griddatamodel.removeRow(i)  # グリッドコントロールから選択行を削除。
		elif cmd=="deleteall":  # 全行を削除する。  	
			msg = "表示しているすべての行を削除しますか?"
			msgbox = peer.getToolkit().createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "履歴", msg)
			if msgbox.execute()==MessageBoxResults.YES:		
				msg = "本当に表示しているすべての行を削除しますか？\n削除したデータは取り戻せません。"
				msgbox = peer.getToolkit().createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "履歴", msg)				
				if msgbox.execute()==MessageBoxResults.YES:	
					if griddatamodel.RowCount==len(datarows):  # グリッドコントロールとDATAROWSの行数が一致している時。
						griddatamodel.removeAllRows()  # グリッドコントロールの行を全削除。
						datarows.clear()  # 全データ行をクリア。	
					else:
						gridcontrol.selectAllRows()  # すべての行を選択。
						for i in selectedrowindexes[::-1]:  # 選択した行インデックスを後ろから取得。逐次検索のときはグリッドコントロールとDATAROWSが一致しないので別に処理する。
							d = griddatamodel.getRowData(i)[0]  # タプルが返るのでその先頭の要素を取得。
							datarows = [j for j in datarows if not d in j]  # dが要素にある行を除いて取得。
						griddatamodel.removeAllRows()  # グリッドコントロールの行を全削除。							
		global DATAROWS
		DATAROWS = datarows			
	def itemActivated(self, menuevent):
		pass
	def itemDeactivated(self, menuevent):
		pass   
	def disposing(self, eventobject):
		pass
class WindowListener(unohelper.Base, XWindowListener):
	def __init__(self, controlcontainer):
		size = controlcontainer.getSize()
		self.oldwidth, self.oldheight = size.Width, size.Height  # 次の変更前の値として取得。			
		self.controlcontainer = controlcontainer
	def windowResized(self, windowevent):
		newwidth, newheight = windowevent.Width, windowevent.Height
		controlcontainer = self.controlcontainer
		gridcontrol1 = controlcontainer.getControl("Grid1")
		checkboxcontrol1 = controlcontainer.getControl("CheckBox1")
		checkboxcontrol2 = controlcontainer.getControl("CheckBox2")
		buttoncontrol1 = controlcontainer.getControl("Button1")
		gridcontrol1rect = gridcontrol1.getPosSize()  # コントロール間の間隔を幅はX、高さはYから取得。
		checkbox1rect = checkboxcontrol1.getPosSize()  # hをHeightから取得。
		minwidth = checkbox1rect.Width + checkboxcontrol2.getPosSize().Width + buttoncontrol1.getSize().Width + gridcontrol1rect.X*3  # 幅下限を取得。
		minheight = checkbox1rect.Height*3 + gridcontrol1rect.Y*4  # 高さ下限を取得。
		if newwidth<minwidth:  # 変更後のコントロールコンテナの幅を取得。サイズ下限より小さい時は下限値とする。
			newwidth = minwidth
		if newheight<minheight:  # 変更後のコントロールコンテナの高さを取得。サイズ下限より小さい時は下限値とする。
			newheight = minheight
		diff_width = newwidth - self.oldwidth  # 幅変化分
		diff_height = newheight - self.oldheight  # 高さ変化分		
		if diff_width or diff_height:  # いずれかが0出ない時。
			applyDiff = dialogcommons.createApplyDiff(diff_width, diff_height)  # コントロールの位置と大きさを変更する関数を取得。
			applyDiff(controlcontainer, PosSize.SIZE)  # コントロールコンテナの大きさを変更する。
			applyDiff(gridcontrol1, PosSize.SIZE)
			applyDiff(controlcontainer.getControl("Edit1"), PosSize.Y+PosSize.WIDTH)
			applyDiff(checkboxcontrol1, PosSize.Y)
			applyDiff(checkboxcontrol2, PosSize.Y)
			applyDiff(buttoncontrol1, PosSize.POS)		
			self.oldwidth, self.oldheight = newwidth, newheight  # 次の変更前の値として取得。
		scrollDown(gridcontrol1)
	def windowMoved(self, windowevent):
		pass
	def windowShown(self, eventobject):
		pass
	def windowHidden(self, eventobject):
		pass
	def disposing(self, eventobject):
		pass
def scrollDown(gridcontrol):  # グリッドコントロールを下までスクロールする。		
	accessiblecontext = gridcontrol.getAccessibleContext()  # グリッドコントロールのAccessibleContextを取得。
	for i in range(accessiblecontext.getAccessibleChildCount()):  # 子要素をのインデックスを走査する。
		child = accessiblecontext.getAccessibleChild(i)  # 子要素を取得。
		if child.getAccessibleContext().getAccessibleRole()==AccessibleRole.SCROLL_BAR:  # スクロールバーの時。
			if child.getOrientation()==ScrollBarOrientation.VERTICAL:  # 縦スクロールバーの時。
				child.setValue(0)  # 一旦0にしないといけない？
				child.setValue(child.getMaximum())  # 最大値にスクロールさせる。
				break				
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
			flgs = [False, False]  # ループを抜けるフラグ。
			for j in range(childaccessiblecontext.getAccessibleChildCount()):  # 孫AccessibleContextについて。 
				grandchildaccessiblecontext = childaccessiblecontext.getAccessibleChild(j).getAccessibleContext()  # 孫AccessibleContextのAccessibleContext。
				accessiblerole = grandchildaccessiblecontext.getAccessibleRole()
				if accessiblerole==84:  # DOCUMENT_SPREADSHEETの時。これが枠。
					bounds = grandchildaccessiblecontext.getBounds()  # 枠の位置と大きさを取得(SCROLL_PANEの左上角が原点)。
					if bounds.X==border.Left and bounds.Y==border.Top:  # SCROLL_PANEに対する相対座標が行ヘッダと列ヘッダと一致する時は左上枠。
						sourcepointonscreen =  grandchildaccessiblecontext.getLocationOnScreen()  # 左上枠の左上角の点を取得(画面の左上角が原点)。
						x = sourcepointonscreen.X + enhancedmouseevent.X  # クリックしたX座標を取得。		
						y = sourcepointonscreen.Y + bounds.Height + enhancedmouseevent.Y + rootpanebounds.Y + border.Top*3  # クリックした位置からシートのヘッダー3行下にダイアログを表示。							
						flgs[0] = True
				elif accessiblerole==50:  # スクロールバーの時。
					bounds = grandchildaccessiblecontext.getBounds()  # 位置と大きさを取得(SCROLL_PANEの左上角が原点)。
					if bounds.X!=0:  # X座標が0でないときは縦スクロールバー。
						vscrollbar = grandchildaccessiblecontext  # 縦スクロールバーを取得。
						flgs[1] = True
				if all(flgs):
					return x, y, vscrollbar  # 縦スクロールバーも返す。
