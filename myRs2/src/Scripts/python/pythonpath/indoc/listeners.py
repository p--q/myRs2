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
def invokeModuleMethod(name, methodname, *args):  # commons.getModle()でモジュールを振り分けてそのモジュールのmethodnameのメソッドを引数argsで呼び出す。
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # ここでブレークするとすべてのイベントでブレークすることになる。
	try:
		m = getModule(name)  # モジュールを取得。
		if hasattr(m, methodname):  # モジュールにmethodnameの関数が存在する時。	
			return getattr(m, methodname)(*args)  # その関数を実行。
		return None  # メソッドが見つからなかった時はNoneを返す。ハンドラやインターセプターは戻り値の処理が必ず必要。
	except:
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		xscriptcontext = args[-1]
		import os, platform, subprocess, traceback
		from com.sun.star.awt import MessageBoxButtons, MessageBoxResults  # 定数
		from com.sun.star.awt.MessageBoxType import ERRORBOX, QUERYBOX  # enum
		from com.sun.star.util import URL  # Struct
		traceback.print_exc()  # PyDevのコンソールにトレースバックを表示。stderrToServer=Trueが必須。
		#  メッセージボックスにも表示する。raiseだとPythonの構文エラーはエラーダイアログがでてこないので。
		lines = traceback.format_exc().split("\n")  # トレースバックを改行で分割。
		fileurl, lineno = "", ""
		for i, line in enumerate(lines[::-1], start=1):  # トレースバックの行を後ろからインデックス1としてイテレート。
			if line.lstrip().startswith("File "):  # File から始まる行の時。
				fileurl = line.split('"')[1]
				lineno = line.split(',')[1].split(" ")[2]
				break			 
		msg = "\n".join(lines[-i:])  # 一番最後のFile から始まる行以降のみ表示する。メッセージボックスに表示できる文字に制限があるので。
		doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
		controller = doc.getCurrentController()  # コントローラの取得。		
		componentwindow = controller.ComponentWindow
		toolkit = componentwindow.getToolkit()
		msgbox = toolkit.createMessageBox(componentwindow, ERRORBOX, MessageBoxButtons.BUTTONS_OK, lines[0], msg)
		msgbox.execute()		
		# Geayでエラー箇所を開く。
		if all([fileurl, lineno]):  # ファイル名と行番号が取得出来ている時。
			flg = (platform.system()=="Windows")  # Windowsかのフラグ。
			if flg:  # Windowsの時
				geanypath = "C:\\Program Files (x86)\\Geany\\bin\\geany.exe"  # 64bitでのパス。パス区切りは\\にしないとエスケープ文字に反応してしまう。
				if not os.path.exists(geanypath):  # binフォルダはなぜかos.path.exists()は常にFalseになるので使えない。
					geanypath = "C:\\Program Files\\Geany\\bin\\geany.exe"  # 32bitでのパス。
					if not os.path.exists(geanypath):
						geanypath = ""
			else:  # Linuxの時。
				p = subprocess.run(["which", "geany"], universal_newlines=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)  # which geanyの結果をuniversal_newlines=Trueで文字列で取得。
				geanypath = p.stdout.strip()  # /usr/bin/geany が返る。
			if geanypath:  # geanyがインストールされている時。
				msg = "Geanyでソースのエラー箇所を一時ファイルで表示しますか?"
				msgbox = toolkit.createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_OK, "Geany", msg)
				if msgbox.execute()==MessageBoxResults.OK:	
					ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
					smgr = ctx.getServiceManager()  # サービスマネージャーの取得。 			
					simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)					
					tempfile = smgr.createInstanceWithContext("com.sun.star.io.TempFile", ctx)  # 一時ファイルを取得。一時フォルダを利用するため。
					urltransformer = smgr.createInstanceWithContext("com.sun.star.util.URLTransformer", ctx)
					dummy, tempfileURL = urltransformer.parseStrict(URL(Complete=tempfile.Uri))
					dummy, fileURL = urltransformer.parseStrict(URL(Complete=fileurl))
					destfileurl = "".join([tempfileURL.Protocol, tempfileURL.Path, fileURL.Name])
					simplefileaccess.copy(fileurl, destfileurl)  # マクロファイルを一時フォルダにコピー。
					filepath =  unohelper.fileUrlToSystemPath(destfileurl)  # 一時フォルダのシステムパスを取得。
					if flg:  # Windowsの時。Windowsではなぜか一時ファイルが残る。削除してもLibreOffice6.0.5を終了すると復活して残る。C:\Users\pq\AppData\Local\Temp\
						os.system('start "" "{}" "{}:{}"'.format(geanypath, filepath, lineno))  # バックグランドでGeanyでカーソルの行番号を指定して開く。第一引数の""はウィンドウタイトル。
					else:
						os.system("{} {}:{} &".format(geanypath, filepath, lineno))  # バックグランドでGeanyでカーソルの行番号を指定して開く。
def addLinsteners(tdocimport, modulefolderpath, xscriptcontext):  # 引数は文書のイベント駆動用。
	invokeModuleMethod(None, "documentOnLoad", xscriptcontext)  # ドキュメントを開いた時に実行するメソッド。
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
	doc.addDocumentEventListener(DocumentEventListener(xscriptcontext, tdocimport, modulefolderpath, controller, *listeners))  # DocumentEventListener。ドキュメントとコントローラに追加したリスナーの除去に利用。
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, xscriptcontext, *args):
		self.xscriptcontext = xscriptcontext
		self.args = args
	def documentEventOccured(self, documentevent):
		eventname = documentevent.EventName
		if eventname=="OnUnload":  # ドキュメントを閉じる時。リスナーを除去する。
			tdocimport, modulefolderpath, controller, changeslistener, selectionchangelistener, activationeventlistener, enhancedmouseclickhandler, contextmenuinterceptor = self.args
			tdocimport.remove_meta(modulefolderpath)  # modulefolderpathをメタパスから除去する。
			documentevent.Source.removeChangesListener(changeslistener)
			controller.removeSelectionChangeListener(selectionchangelistener)
			controller.removeActivationEventListener(activationeventlistener)
			controller.removeEnhancedMouseClickHandler(enhancedmouseclickhandler)
			controller.releaseContextMenuInterceptor(contextmenuinterceptor)
			invokeModuleMethod(None, "documentUnLoad", self.xscriptcontext)  # ドキュメントを閉じた時に実行するメソッド。
	def disposing(self, eventobject):  # ドキュメントを閉じるときに発火する。	
		eventobject.Source.removeDocumentEventListener(self)
class ActivationEventListener(unohelper.Base, XActivationEventListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def activeSpreadsheetChanged(self, activationevent):  # アクティブシートが変化した時に発火。
		invokeModuleMethod(activationevent.ActiveSheet.getName(), "activeSpreadsheetChanged", activationevent, self.xscriptcontext)
	def disposing(self, eventobject):
		eventobject.Source.removeActivationEventListener(self)	
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):  # enhancedmouseeventのSourceはNoneなので、このリスナーのメソッドの引数からコントローラーを直接取得する方法はない。
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def mousePressed(self, enhancedmouseevent):  # セルをクリックした時に発火する。固定行列の最初のクリックは同じ相対位置の固定していないセルが返ってくる(表示されている自由行の先頭行に背景色がる時のみ）。
		target = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if target.supportsService("com.sun.star.sheet.SheetCellRange"):  # targetがチャートの時がありうる?
			b = invokeModuleMethod(target.getSpreadsheet().getName(), "mousePressed", enhancedmouseevent, self.xscriptcontext)  # 正しく実行されれば、ブーリアンが返ってくるはず。
			if b is not None:
				return b
		return True  # Falseを返すと右クリックメニューがでてこなくなる。		
	def mouseReleased(self, enhancedmouseevent):
		return True  # シングルクリックでFalseを返すとセル選択範囲の決定の状態になってどうしようもなくなる。
	def disposing(self, eventobject):  # eventobject.SourceはNone。
		self.xscriptcontext.getDocument().getCurrentController().removeEnhancedMouseClickHandler(self)
class SelectionChangeListener(unohelper.Base, XSelectionChangeListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
		self.selectionrangeaddress = None  # selectionChanged()メソッドが何回も無駄に発火するので選択範囲アドレスのStructをキャッシュして比較する。
	def selectionChanged(self, eventobject):  # マウスから呼び出した時の反応が遅い。このメソッドでエラーがでるとショートカットキーでの操作が必要。
		selection = eventobject.Source.getSelection()
		if hasattr(selection, "getRangeAddress"):  # 選択範囲がセル範囲とは限らないのでgetRangeAddress()メソッドがあるか確認する。
			selectionrangeaddress = selection.getRangeAddress()
			if selectionrangeaddress==self.selectionrangeaddress:  # キャッシュのセル範囲アドレスと一致する時。Structで比較。セル範囲では比較できない。
				return  # 何もしない。
			else:  # キャッシュのセル範囲と一致しない時。
				self.selectionrangeaddress = selectionrangeaddress  # キャッシュを更新。
		invokeModuleMethod(eventobject.Source.getActiveSheet().getName(), "selectionChanged", eventobject, self.xscriptcontext)		
	def disposing(self, eventobject):
		eventobject.Source.removeSelectionChangeListener(self)		
class ChangesListener(unohelper.Base, XChangesListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def changesOccurred(self, changesevent):  # Sourceにはドキュメントが入る。
		invokeModuleMethod(changesevent.Source.getCurrentController().getActiveSheet().getName(), "changesOccurred", changesevent, self.xscriptcontext)							
	def disposing(self, eventobject):
		eventobject.Source.removeChangesListener(self)			
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 右クリックで呼ばれる関数。contextmenuexecuteevent.ActionTriggerContainerを操作しないとコンテクストメニューが表示されない。:
		contextmenuinterceptoraction = invokeModuleMethod(contextmenuexecuteevent.Selection.getActiveSheet().getName(), "notifyContextMenuExecute", contextmenuexecuteevent, self.xscriptcontext)  # 正しく実行されれば、enumのcom.sun.star.ui.ContextMenuInterceptorActionのいずれかが返るはず。	
		if contextmenuinterceptoraction is not None:
			return contextmenuinterceptoraction
		return IGNORED  # コンテクストメニューのカスタマイズをしない。
