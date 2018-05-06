#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# embeddedmacro.pyから呼び出した関数ではXSCRIPTCONTEXTは使えない。デコレーターも使えない。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)でブレークする。
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.document import XDocumentEventListener
from com.sun.star.sheet import XActivationEventListener
# from com.sun.star.table import BorderLine2  # Struct
# from com.sun.star.table import BorderLineStyle  # 定数
# from com.sun.star.table import TableBorder2  # Struct
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui.ContextMenuInterceptorAction import IGNORED  # enum
from com.sun.star.util import XChangesListener
from com.sun.star.view import XSelectionChangeListener
from myrs import commons, ichiran, karute, keika, rireki, taiin, yotei  # 相対インポートは不可。
def myRs(tdocimport, modulefolderpath, xscriptcontext):  # 引数は文書のイベント駆動用。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
# 	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
# 	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	controller = doc.getCurrentController()  # コントローラの取得。
# 	systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
# 	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。半角カタカナへの変換に利用。
# 	borders = createBorders()  # 枠線の作成。
	changeslistener = ChangesListener(xscriptcontext)  # ChangesListener。セルの変化の感知に利用。列の挿入も感知。
	
# 	selectionchangelistener = SelectionChangeListener(xscriptcontext, borders)  # SelectionChangeListener。選択範囲の変更の感知に利用。
	selectionchangelistener = SelectionChangeListener(xscriptcontext)  # SelectionChangeListener。選択範囲の変更の感知に利用。

	activationeventlistener = ActivationEventListener(xscriptcontext)  # ActivationEventListener。シートの切替の感知に利用。
# 	enhancedmouseclickhandler = EnhancedMouseClickHandler(xscriptcontext, controller, borders, systemclipboard, transliteration)  # EnhancedMouseClickHandler。マウスの左クリックの感知に利用。enhancedmouseeventのSourceはNone。

# 	enhancedmouseclickhandler = EnhancedMouseClickHandler(xscriptcontext, borders)  # EnhancedMouseClickHandler。マウスの左クリックの感知に利用。enhancedmouseeventのSourceはNone。
	enhancedmouseclickhandler = EnhancedMouseClickHandler(xscriptcontext)  # EnhancedMouseClickHandler。マウスの左クリックの感知に利用。enhancedmouseeventのSourceはNone。

# 	contextmenuinterceptor = ContextMenuInterceptor(xscriptcontext, ctx, smgr, doc)  # ContextMenuInterceptor。右クリックメニューの変更に利用。
	contextmenuinterceptor = ContextMenuInterceptor(xscriptcontext)  # ContextMenuInterceptor。右クリックメニューの変更に利用。
	
	doc.addChangesListener(changeslistener)
	controller.addSelectionChangeListener(selectionchangelistener)
	controller.addActivationEventListener(activationeventlistener)
	controller.addEnhancedMouseClickHandler(enhancedmouseclickhandler)
	controller.registerContextMenuInterceptor(contextmenuinterceptor)
	listeners = changeslistener, selectionchangelistener, activationeventlistener, enhancedmouseclickhandler, contextmenuinterceptor
	doc.addDocumentEventListener(DocumentEventListener(tdocimport, modulefolderpath, controller, *listeners))  # DocumentEventListener。ドキュメントとコントローラに追加したリスナーの除去に利用。
def getModule(sheetname):
	if sheetname.startswith("00000000"):  # テンプレートの時は何もしない。
		pass
	elif sheetname.isdigit():  # シート名が数字のみの時カルテシート。
		return karute
	elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
		return keika
	elif sheetname=="一覧":
		return ichiran
	elif sheetname=="予定":
		return yotei
	elif sheetname=="退院":
		return taiin
	elif sheetname=="履歴":
		return rireki
	return None
# def createBorders():# 枠線の作成。
# 	noneline = BorderLine2(LineStyle=BorderLineStyle.NONE)
# 	firstline = BorderLine2(LineStyle=BorderLineStyle.DASHED, LineWidth=62, Color=commons.COLORS["violet"])
# 	secondline =  BorderLine2(LineStyle=BorderLineStyle.DASHED, LineWidth=62, Color=commons.COLORS["magenta3"])	
# 	tableborder2 = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)
# 	topbottomtableborder = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=False, IsRightLineValid=False)
# 	leftrighttableborder = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=False, IsBottomLineValid=False, IsLeftLineValid=True, IsRightLineValid=True)
# 	return noneline, tableborder2, topbottomtableborder, leftrighttableborder  # 作成した枠線をまとめたタプル。
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, *args):
		self.args = args
	def documentEventOccured(self, documentevent):
		try:
			eventname = documentevent.EventName
			if eventname=="OnUnload":  # ドキュメントを閉じる時。リスナーを除去する。
				tdocimport, modulefolderpath, controller, changeslistener, selectionchangelistener, activationeventlistener, enhancedmouseclickhandler, contextmenuinterceptor = self.args
				tdocimport.remove_meta(modulefolderpath)  # modulefolderpathをメタパスから除去する。
				documentevent.Source.removeChangesListener(changeslistener)
				controller.removeSelectionChangeListener(selectionchangelistener)
				controller.removeActivationEventListener(activationeventlistener)
				controller.removeEnhancedMouseClickHandler(enhancedmouseclickhandler)
				controller.releaseContextMenuInterceptor(contextmenuinterceptor)
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	def disposing(self, eventobject):  # ドキュメントを閉じるときに発火する。	
		try:
			eventobject.Source.removeDocumentEventListener(self)
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
class ActivationEventListener(unohelper.Base, XActivationEventListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def activeSpreadsheetChanged(self, activationevent):  # アクティブシートが変化した時に発火。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		try:
			m = getModule(activationevent.ActiveSheet.getName())
			if hasattr(m, "activeSpreadsheetChanged"):
				getattr(m, "activeSpreadsheetChanged")(activationevent, self.xscriptcontext)
			
			
# 			sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
# 			sheetname = sheet.getName()  # アクティブシート名を取得。
# 			if sheetname.startswith("00000000"):  # テンプレートの時は何もしない。
# 				pass
# 			elif sheetname.isdigit():  # シート名が数字のみの時カルテシート。
# # 				karute.activeSpreadsheetChanged(sheet)
# 				karute.activeSpreadsheetChanged(activationevent, self.args)
# 			elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
# # 				controller = activationevent.Source
# # 				keika.activeSpreadsheetChanged(controller, sheet)
# 				keika.activeSpreadsheetChanged(activationevent, self.args)
# 			elif sheetname=="一覧":
# # 				ichiran.activeSpreadsheetChanged(sheet)
# 				ichiran.activeSpreadsheetChanged(activationevent, self.args)
# 			elif sheetname=="予定":
# 				pass
# 			elif sheetname=="退院":
# 				pass
# 			elif sheetname=="履歴":
# 				pass
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	def disposing(self, eventobject):
		try:
			eventobject.Source.removeActivationEventListener(self)	
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):  # このリスナーのメソッドの引数からコントローラーを取得する方法がない。
# 	def __init__(self, xscriptcontext, controller, borders, systemclipboard, transliteration):
# 		self.xscriptcontext = xscriptcontext
# 		self.controller = controller
# 		self.args = borders, systemclipboard, transliteration
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def mousePressed(self, enhancedmouseevent):  # セルをクリックした時に発火する。固定行列の最初のクリックは同じ相対位置の固定していないセルが返ってくる(表示されている自由行の先頭行に背景色がる時のみ）。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		try:
			target = enhancedmouseevent.Target  # ターゲットのセルを取得。
			if target.supportsService("com.sun.star.sheet.SheetCellRange"):  # targetがチャートの時がありうる?
				m = getModule(enhancedmouseevent.getSpreadsheet().getName())
				if hasattr(m, "mousePressed"):
					return getattr(m, "mousePressed")(enhancedmouseevent, self.xscriptcontext)			
			return True
			
			
# 			target = enhancedmouseevent.Target  # ターゲットのセルを取得。
# 			if target.supportsService("com.sun.star.sheet.SheetCellRange"):  # targetがチャートの時がありうる?
# 				sheet = target.getSpreadsheet()
# 				sheetname = sheet.getName()
# 				if sheetname.startswith("00000000"):  # テンプレートの時は何もしない。
# 					pass
# 				elif sheetname.isdigit():  # シート名が数字のみの時カルテシート。
# 					return karute.mousePressed(enhancedmouseevent, self.args)
# 				elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
# 					return keika.mousePressed(enhancedmouseevent, self.args)
# 				elif sheetname=="一覧":
# 					return ichiran.mousePressed(enhancedmouseevent, self.args)
# 				elif sheetname=="予定":
# 					return True
# 				elif sheetname=="退院":
# 					return True
# 				elif sheetname=="履歴":
# 					return True
			return True  # Falseを返すと右クリックメニューがでてこなくなる。	
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。	
	def mouseReleased(self, enhancedmouseevent):
		return True  # シングルクリックでFalseを返すとセル選択範囲の決定の状態になってどうしようもなくなる。
	def disposing(self, eventobject):  # eventobject.SourceはNone。
		try:
			self.controller.removeEnhancedMouseClickHandler(self)	
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
class SelectionChangeListener(unohelper.Base, XSelectionChangeListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def selectionChanged(self, eventobject):  # マウスから呼び出した時の反応が遅い。このメソッドでエラーがでるとショートカットキーでの操作が必要。
	# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		try:
			m = getModule(eventobject.Source.getActiveSheet().getName())
			if hasattr(m, "selectionChanged"):
				getattr(m, "selectionChanged")(eventobject, self.xscriptcontext)			
			
			
			
# 			controller = eventobject.Source
# 			sheet = controller.getActiveSheet()
# 			sheetname = sheet.getName()  # アクティブシート名を取得。		
# 			if sheetname.startswith("00000000"):  # テンプレートの時は何もしない。
# 				pass
# 			elif sheetname.isdigit():  # シート名が数字のみの時カルテシート。
# 				karute.selectionChanged(eventobject, self.args)
# 			elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
# 				keika.selectionChanged(eventobject, self.args)
# 			elif sheetname=="一覧":
# 				ichiran.selectionChanged(eventobject, self.args)
# 			elif sheetname=="予定":
# 				pass
# 			elif sheetname=="退院":
# 				pass
# 			elif sheetname=="履歴":
# 				pass			
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	def disposing(self, eventobject):
		try:
			eventobject.Source.removeSelectionChangeListener(self)		
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
class ChangesListener(unohelper.Base, XChangesListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def changesOccurred(self, changesevent):  # Sourceにはドキュメントが入る。
		try:
			m = getModule(changesevent.Source.getCurrentController().getActiveSheet().getName())
			if hasattr(m, "changesOccurred"):
				getattr(m, "changesOccurred")(changesevent, self.xscriptcontext)
			
			
# 			doc = changesevent.Source
# 			controller = doc.getCurrentController()
# 			sheet = controller.getActiveSheet()
# 			sheetname = sheet.getName()  # アクティブシート名を取得。
# 			if sheetname.startswith("00000000"):  # テンプレートの時は何もしない。
# 				pass
# 			elif sheetname.isdigit():  # シート名が数字のみの時カルテシート。
# 				pass
# 			elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
# 				pass
# 			elif sheetname=="一覧":
# 				pass
# 			elif sheetname=="予定":
# 				pass
# 			elif sheetname=="退院":
# 				pass
# 			elif sheetname=="履歴":
# 				pass		
# 		changes = changesevent.Changes
# 		for change in changes:
# 			accessor = change.Accessor
# 			if accessor=="cell-change":  # セルの内容が変化した時。
# 				cell = change.ReplacedElement  # 変化したセルを取得。		
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。								
	def disposing(self, eventobject):
		try:
			eventobject.Source.removeChangesListener(self)			
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, xscriptcontext):
# 		self.xscriptcontext = xscriptcontext
# 		self.args = getBaseURL(ctx, smgr, doc)  # ScriptingURLのbaseurlを取得。
		self.xscriptcontext = xscriptcontext
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 右クリックで呼ばれる関数。contextmenuexecuteevent.ActionTriggerContainerを操作しないとコンテクストメニューが表示されない。
	# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		try:
			m = getModule(contextmenuexecuteevent.Selection.getActiveSheet().getName())
			if hasattr(m, "notifyContextMenuExecute"):
				return getattr(m, "notifyContextMenuExecute")(contextmenuexecuteevent, self.xscriptcontext)			
			return IGNORED  # コンテクストメニューのカスタマイズをしない。
			
			
			
# 			baseurl = getBaseURL(self.xscriptcontext)  # ScriptingURLのbaseurlを取得。
# 			
# 			baseurl = self.args 
# 			controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
# 			contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
# 			contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
# 			addMenuentry = menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
# 			sheet = controller.getActiveSheet()  # アクティブシートを取得。
# 			sheetname = sheet.getName()  # シート名を取得。
# 			if sheetname.startswith("00000000"):  # テンプレートの時は何もしない。
# 				pass
# 			elif sheetname.isdigit():  # シート名が数字のみの時カルテシート。
# # 				return karute.notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, sheet, contextmenuname)
# 				return karute.notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, sheet, contextmenuname)
# 			elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
# 				pass
# 			elif sheetname=="一覧":
# 				pass
# 			elif sheetname=="予定":
# 				pass
# 			elif sheetname=="退院":
# 				pass
# 			elif sheetname=="履歴":
# 				pass
# 			return IGNORED  # コンテクストメニューのカスタマイズをしない。
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
# ContextMenuInterceptorのnotifyContextMenuExecute()メソッドで設定したメニュー項目から呼び出される関数。関数名変更不可。動的生成も不可。
def entry1():
	invokeMenuEntry(1)
def entry2():
	invokeMenuEntry(2)	
def entry3():
	invokeMenuEntry(3)	
def entry4():
	invokeMenuEntry(4)
def entry5():
	invokeMenuEntry(5)
def entry6():
	invokeMenuEntry(6)
def entry7():
	invokeMenuEntry(7)
def entry8():
	invokeMenuEntry(8)
def entry9():
	invokeMenuEntry(9)	


def invokeMenuEntry(entrynum):  # コンテクストメニュー項目から呼び出された処理をシートごとに振り分ける。コンテクストメニューから呼び出しているこの関数ではXSCRIPTCONTEXTが使える。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	selection = doc.getCurrentSelection()  # セル(セル範囲)またはセル範囲、セル範囲コレクションが入るはず。
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # セル範囲コレクション以外の時。
		m = getModule(doc.getCurrentController().getActiveSheet().getName())
		if hasattr(m, "contextMenuEntries"):
			getattr(m, "contextMenuEntries")(entrynum, XSCRIPTCONTEXT)	
		
		
# 		sheet = selection.getSpreadsheet()  # シートを取得。
# 		sheetname = sheet.getName()  # シート名を取得。
# 		if sheetname.startswith("00000000"):  # テンプレートの時は何もしない。
# 			pass
# 		elif sheetname.isdigit():  # シート名が数字のみの時カルテシート。
# 			karute.contextMenuEntries(XSCRIPTCONTEXT, selection, entrynum)
# 		elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
# 			keika.contextMenuEntries(selection, entrynum)
# 		elif sheetname=="一覧":
# 			ichiran.contextMenuEntries(selection, entrynum)
# 		elif sheetname=="予定":
# 			yotei.contextMenuEntries(selection, entrynum)
# 		elif sheetname=="退院":
# 			taiin.contextMenuEntries(selection, entrynum)
# 		elif sheetname=="履歴":
# 			rireki.contextMenuEntries(selection, entrynum)
# def menuentryCreator(menucontainer):  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
# 	i = 0  # インデックスを初期化する。
# 	def addMenuentry(menutype, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
# 		nonlocal i
# 		menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
# 		[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
# 		menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
# 		i += 1  # インデックスを増やす。
# 	return addMenuentry
# def getBaseURL(xscriptcontext):	 # 埋め込みマクロのScriptingURLのbaseurlを返す。__file__はvnd.sun.star.tdoc:/4/Scripts/python/filename.pyというように返ってくる。
# 	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
# 	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
# 	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
# 	modulepath = __file__  # ScriptingURLにするマクロがあるモジュールのパスを取得。ファイルのパスで場合分け。sys.path[0]は__main__の位置が返るので不可。
# 	ucp = "vnd.sun.star.tdoc:"  # 埋め込みマクロのucp。
# 	filepath = modulepath.replace(ucp, "")  #  ucpを除去。
# 	transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
# 	transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(doc)
# 	contentidentifierstring = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # __file__の数値部分に該当。
# 	macrofolder = "{}/Scripts/python".format(contentidentifierstring.replace(ucp, ""))  #埋め込みマクロフォルダへのパス。	
# 	location = "document"  # マクロの場所。	
# 	relpath = os.path.relpath(filepath, start=macrofolder)  # マクロフォルダからの相対パスを取得。パス区切りがOS依存で返ってくる。
# 	return "vnd.sun.star.script:{}${}?language=Python&location={}".format(relpath.replace(os.sep, "|"), "{}", location)  # ScriptingURLのbaseurlを取得。Windowsのためにos.sepでパス区切りを置換。	
