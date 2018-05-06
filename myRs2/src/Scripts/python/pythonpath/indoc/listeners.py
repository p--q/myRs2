#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# embeddedmacro.pyから呼び出した関数ではXSCRIPTCONTEXTは使えない。デコレーターも使えない。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)でブレークする。
import unohelper  # オートメーションには必須(必須なのはuno)。
from indoc.commons import getModule # 相対インポートはできない。
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.document import XDocumentEventListener
from com.sun.star.sheet import XActivationEventListener
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui.ContextMenuInterceptorAction import IGNORED  # enum
from com.sun.star.util import XChangesListener
from com.sun.star.view import XSelectionChangeListener
def addLinsteners(tdocimport, modulefolderpath, xscriptcontext):  # 引数は文書のイベント駆動用。
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	try:
		doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
		controller = doc.getCurrentController()  # コントローラの取得。
		changeslistener = ChangesListener(xscriptcontext)  # ChangesListener。セルの変化の感知に利用。列の挿入も感知。
		selectionchangelistener = SelectionChangeListener(xscriptcontext)  # SelectionChangeListener。選択範囲の変更の感知に利用。
		activationeventlistener = ActivationEventListener(xscriptcontext)  # ActivationEventListener。シートの切替の感知に利用。
		enhancedmouseclickhandler = EnhancedMouseClickHandler(xscriptcontext)  # EnhancedMouseClickHandler。マウスの左クリックの感知に利用。enhancedmouseeventのSourceはNone。
		contextmenuinterceptor = ContextMenuInterceptor(xscriptcontext)  # ContextMenuInterceptor。右クリックメニューの変更に利用。
		doc.addChangesListener(changeslistener)
		controller.addSelectionChangeListener(selectionchangelistener)
		controller.addActivationEventListener(activationeventlistener)
		controller.addEnhancedMouseClickHandler(enhancedmouseclickhandler)
		controller.registerContextMenuInterceptor(contextmenuinterceptor)
		listeners = changeslistener, selectionchangelistener, activationeventlistener, enhancedmouseclickhandler, contextmenuinterceptor
		doc.addDocumentEventListener(DocumentEventListener(tdocimport, modulefolderpath, controller, *listeners))  # DocumentEventListener。ドキュメントとコントローラに追加したリスナーの除去に利用。
	except:
		import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。	
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, *args):
		self.args = args
	def documentEventOccured(self, documentevent):
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
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
		eventobject.Source.removeDocumentEventListener(self)
class ActivationEventListener(unohelper.Base, XActivationEventListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def activeSpreadsheetChanged(self, activationevent):  # アクティブシートが変化した時に発火。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		try:
			m = getModule(activationevent.ActiveSheet.getName())
			if hasattr(m, "activeSpreadsheetChanged"):
				getattr(m, "activeSpreadsheetChanged")(activationevent, self.xscriptcontext)
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	def disposing(self, eventobject):
		eventobject.Source.removeActivationEventListener(self)	
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):  # enhancedmouseeventのSourceはNoneなので、このリスナーのメソッドの引数からコントローラーを直接取得する方法はない。
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def mousePressed(self, enhancedmouseevent):  # セルをクリックした時に発火する。固定行列の最初のクリックは同じ相対位置の固定していないセルが返ってくる(表示されている自由行の先頭行に背景色がる時のみ）。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		try:
			target = enhancedmouseevent.Target  # ターゲットのセルを取得。
			if target.supportsService("com.sun.star.sheet.SheetCellRange"):  # targetがチャートの時がありうる?
				m = getModule(target.getSpreadsheet().getName())
				if hasattr(m, "mousePressed"):
					return getattr(m, "mousePressed")(enhancedmouseevent, self.xscriptcontext)			
			return True  # Falseを返すと右クリックメニューがでてこなくなる。	
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。	
	def mouseReleased(self, enhancedmouseevent):
		return True  # シングルクリックでFalseを返すとセル選択範囲の決定の状態になってどうしようもなくなる。
	def disposing(self, eventobject):  # eventobject.SourceはNone。
		self.xscriptcontext.getDocument().getCurrentController().removeEnhancedMouseClickHandler(self)
class SelectionChangeListener(unohelper.Base, XSelectionChangeListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def selectionChanged(self, eventobject):  # マウスから呼び出した時の反応が遅い。このメソッドでエラーがでるとショートカットキーでの操作が必要。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		try:
			m = getModule(eventobject.Source.getActiveSheet().getName())
			if hasattr(m, "selectionChanged"):
				getattr(m, "selectionChanged")(eventobject, self.xscriptcontext)			
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	def disposing(self, eventobject):
		eventobject.Source.removeSelectionChangeListener(self)		
class ChangesListener(unohelper.Base, XChangesListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def changesOccurred(self, changesevent):  # Sourceにはドキュメントが入る。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		try:
			m = getModule(changesevent.Source.getCurrentController().getActiveSheet().getName())
			if hasattr(m, "changesOccurred"):
				getattr(m, "changesOccurred")(changesevent, self.xscriptcontext)	
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。								
	def disposing(self, eventobject):
		eventobject.Source.removeChangesListener(self)			
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 右クリックで呼ばれる関数。contextmenuexecuteevent.ActionTriggerContainerを操作しないとコンテクストメニューが表示されない。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		try:
			m = getModule(contextmenuexecuteevent.Selection.getActiveSheet().getName())
			if hasattr(m, "notifyContextMenuExecute"):
				return getattr(m, "notifyContextMenuExecute")(contextmenuexecuteevent, self.xscriptcontext)			
			return IGNORED  # コンテクストメニューのカスタマイズをしない。
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
