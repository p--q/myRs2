#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import os, platform, subprocess, traceback, unohelper
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import MessageBoxButtons, MessageBoxResults, PosSize, SystemPointer  # 定数
from com.sun.star.awt import Point  # Struct
from com.sun.star.awt.MessageBoxType import ERRORBOX, QUERYBOX  # enum
from com.sun.star.util import URL  # Struct
from com.sun.star.util import MeasureUnit  # 定数
from com.sun.star.style.VerticalAlignment import MIDDLE  # enum
def createDialog(xscriptcontext):
	traceback.print_exc()  # PyDevのコンソールにトレースバックを表示。stderrToServer=Trueが必須。
	#  ダイアログに表示する。raiseだとPythonの構文エラーはエラーダイアログがでてこないので。
	lines = traceback.format_exc().split("\n")  # トレースバックを改行で分割。
	h = 20  # FixedTextコントロールの高さ。ma単位。2行分。	
	docwindow = xscriptcontext.getDocument().getCurrentController().getFrame().getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	dialogwidth = 380  # ウィンドウの幅。ma単位。
	dialog, addControl = dialogCreator(xscriptcontext, {"PositionX": 20, "PositionY": 120, "Width": dialogwidth, "Height": 10, "Title": lines[0], "Name": "exceptiondialog", "Moveable": True})  # Heightは後で設定し直す。
	dialog.createPeer(docwindow.getToolkit(), docwindow)  # ダイアログを描画。親ウィンドウを渡す。
	mouselistener = MouseListener(xscriptcontext)
	controlheight = 0  # コントロールの高さ。ma単位。
	for i in lines[1:]:  # 2行目以降イテレート。
		if i:  # 空行は除外。
			fixedtextprops = [{"PositionX": 0, "PositionY": controlheight, "Width": dialogwidth, "Height": h, "Label": i, "MultiLine": True, "NoLabel": True, "VerticalAlign": MIDDLE}]
			if i.lstrip().startswith("File "):  # File から始まる行の時。	
				fixedtextprops[0]["TextColor"] = 0x0000FF  # 文字色をblue3にする。
				fixedtextprops.append({"addMouseListener": mouselistener})
			elif not i.startswith(" "):  # スペース以外から始まる時。
				fixedtextprops[0]["TextColor"] = 0xFF0000  # 文字色をred3にする。
			fixedtextcontrol = addControl("FixedText", *fixedtextprops)
			controlheight += h
	controlrectangle = fixedtextcontrol.getPosSize()  # コントロール間の間隔を幅はX、高さはYから取得。
	dialog.setPosSize(0, 0, 0, controlrectangle.Y+controlrectangle.Height, PosSize.HEIGHT)  # 最後の行からダイアログの高さを再設定。
	dialog.execute()
	dialog.dispose()	
class MouseListener(unohelper.Base, XMouseListener):  # Editコントロールではうまく動かない。
	def __init__(self, xscriptcontext):
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。			
		self.pointer = smgr.createInstanceWithContext("com.sun.star.awt.Pointer", ctx)  # ポインタのインスタンスを取得。
		self.args = ctx, smgr, xscriptcontext.getDocument()
	def mousePressed(self, mouseevent):
		ctx, smgr, doc = self.args
		txt = mouseevent.Source.getText()
		fileurl = txt.split('"')[1]  # エラー箇所のfileurlを取得。
		lineno = txt.split(',')[1].split(" ")[2]  # エラー箇所の行番号を取得。		
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
			componentwindow = doc.getCurrentController().ComponentWindow
			toolkit = componentwindow.getToolkit()			
			if geanypath:  # geanyがインストールされている時。
				msg = "Geanyでソースのエラー箇所を一時ファイルで表示しますか?"
				msgbox = toolkit.createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "myRs", msg)
				if msgbox.execute()==MessageBoxResults.YES:			
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
			else:
				msg = "Geanyがインストールされていません。"
				msgbox = toolkit.createMessageBox(componentwindow, ERRORBOX, MessageBoxButtons.BUTTONS_OK, "myRs", msg)
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		self.pointer.setType(SystemPointer.REFHAND)  # マウスポインタの種類を設定。
		mouseevent.Source.getPeer().setPointer(self.pointer)  # マウスポインタを変更。コントロールからマウスがでるとポインタは元に戻る。
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):
		eventobject.Source.removeMouseListener(self)	
def dialogCreator(xscriptcontext, dialogprops):  # ダイアログと、それにコントロールを追加する関数を返す。まずダイアログモデルのプロパティを取得。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)  # ダイアログの生成。
	dialogmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)  # ダイアログモデルの生成。
	dialogmodel.setPropertyValues(tuple(dialogprops.keys()), tuple(dialogprops.values()))  # ダイアログモデルのプロパティを設定。
	dialog.setModel(dialogmodel)  # ダイアログにダイアログモデルを設定。
	dialog.setVisible(False)  # 描画中のものを表示しない。
	def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、attr: コントロールの属性。
		controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
		dialogmodel.insertByName(props["Name"], controlmodel)  # ダイアログモデルにモデルを追加するだけでコントロールも作成される。
		control = dialog.getControl(props["Name"])  # コントロールコンテナに追加された後のコントロールを取得。
		if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
			for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
				if val is None:
					getattr(control, key)()
				else:
					getattr(control, key)(val)
		return control  # 追加したコントロールを返す。
	def _createControlModel(controltype, props):  # コントロールモデルの生成。
		if not "Name" in props:
			props["Name"] = _generateSequentialName(controltype)  # Nameがpropsになければ通し番号名を生成。
		controlmodel = dialogmodel.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
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
			flg = dialog.getControl(name)  # 同名のコントロールの有無を判断。
			i += 1
		return name
	return dialog, addControl  # コントロールコンテナとそのコントロールコンテナにコントロールを追加する関数を返す。