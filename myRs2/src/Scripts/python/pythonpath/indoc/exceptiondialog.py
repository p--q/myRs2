#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import os, platform, subprocess, traceback, unohelper
from com.sun.star.awt import MessageBoxButtons, MessageBoxResults  # 定数
from com.sun.star.awt.MessageBoxType import ERRORBOX, QUERYBOX  # enum
from com.sun.star.util import URL  # Struct
def createDialog(xscriptcontext):
	traceback.print_exc()  # PyDevのコンソールにトレースバックを表示。stderrToServer=Trueが必須。
	#  メッセージボックスにも表示する。raiseだとPythonの構文エラーはエラーダイアログがでてこないので。
	lines = traceback.format_exc().split("\n")  # トレースバックを改行で分割。
	fileurl, lineno = "", ""
	for i, line in enumerate(lines[::-1], start=1):  # トレースバックの行を後ろからインデックス1としてイテレート。
		if line.lstrip().startswith("File "):  # File から始まる行の時。
			fileurl = line.split('"')[1]  # エラー箇所のfileurlを取得。
			lineno = line.split(',')[1].split(" ")[2]  # エラー箇所の行番号を取得。
			break			 
		
	
	# ダイアログからfileurlをクリックして開けるようにする。
	# pyで終わるfileurlのときのみ。
		
		
		
	msg = "\n".join(lines[-i:])  # 一番最後のFile から始まる行以降のみ表示する。メッセージボックスに表示できる文字に制限があるので。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。		
	componentwindow = controller.ComponentWindow
	toolkit = componentwindow.getToolkit()
	msgbox = toolkit.createMessageBox(componentwindow, ERRORBOX, MessageBoxButtons.BUTTONS_OK, lines[0], msg)
	msgbox.execute()	
	
#   File "/opt/eclipse4.6jee/plugins/org.python.pydev_5.6.0.201703221358/pysrc/_pydevd_bundle/pydevd_breakpoin...Traceback (most recent call last):
#   File "vnd.sun.star.tdoc:/1/Scripts/python/pythonpath/indoc/listeners.py", line 18, in invokeModuleMethod
#     return getattr(m, methodname)(*args)  # その関数を実行。
#   File "vnd.sun.star.tdoc:/1/Scripts/python/pythonpath/indoc/yotei.py", line 186, in activeSpreadsheetChanged
#     setRangesProperty(holidaycolumns, ("CellBackColor", colors["red3"]))
#   File "vnd.sun.star.tdoc:/1/Scripts/python/pythonpath/indoc/yotei.py", line 231, in setRangesProperty
#     cellranges.addRangeAddresses((VARS.sheet[VARS.dayrow:VARS.datarow, i].getRangeAddress() for i in columnindexes), False)  # セル範囲コレクションを取得。
# indoc.yotei.com.sun.star.script.CannotConvertException: conversion not possible!
	
		
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
					
def dialogCreator(ctx, smgr, dialogprops):  # ダイアログと、それにコントロールを追加する関数を返す。まずダイアログモデルのプロパティを取得。
	dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)  # ダイアログの生成。
	if "PosSize" in dialogprops:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
		dialog.setPosSize(dialogprops.pop("PositionX"), dialogprops.pop("PositionY"), dialogprops.pop("Width"), dialogprops.pop("Height"), dialogprops.pop("PosSize"))  # ダイアログモデルのプロパティで設定すると単位がMapAppになってしまうのでコントロールに設定。
	dialogmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)  # ダイアログモデルの生成。
	dialogmodel.setPropertyValues(tuple(dialogprops.keys()), tuple(dialogprops.values()))  # ダイアログモデルのプロパティを設定。
	dialog.setModel(dialogmodel)  # ダイアログにダイアログモデルを設定。
	dialog.setVisible(False)  # 描画中のものを表示しない。
	def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、attr: コントロールの属性。
		control = None
		items, currentitemid = None, None
		if controltype == "Roadmap":  # Roadmapコントロールのとき、Itemsはダイアログモデルに追加してから設定する。そのときはCurrentItemIDもあとで設定する。
			if "Items" in props:  # Itemsはダイアログモデルに追加されてから設定する。
				items = props.pop("Items")
				if "CurrentItemID" in props:  # CurrentItemIDはItemsを追加されてから設定する。
					currentitemid = props.pop("CurrentItemID")
		if "PosSize" in props:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
			control = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
			control.setPosSize(props.pop("PositionX"), props.pop("PositionY"), props.pop("Width"), props.pop("Height"), props.pop("PosSize"))  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
			controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
			control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
			dialog.addControl(props["Name"], control)  # コントロールをコントロールコンテナに追加。
		else:  # Map AppFont (ma)のときはダイアログモデルにモデルを追加しないと正しくピクセルに変換されない。
			controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
			dialogmodel.insertByName(props["Name"], controlmodel)  # ダイアログモデルにモデルを追加するだけでコントロールも作成される。
		if items is not None:  # コントロールに追加されたRoadmapモデルにしかRoadmapアイテムは追加できない。
			for i, j in enumerate(items):  # 各Roadmapアイテムについて
				item = controlmodel.createInstance()
				item.setPropertyValues(("Label", "Enabled"), j)
				controlmodel.insertByIndex(i, item)  # IDは0から整数が自動追加される
			if currentitemid is not None:  #Roadmapアイテムを追加するとそれがCurrentItemIDになるので、Roadmapアイテムを追加してからCurrentIDを設定する。
				controlmodel.setPropertyValue("CurrentItemID", currentitemid)
		if control is None:  # コントロールがまだインスタンス化されていないとき
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