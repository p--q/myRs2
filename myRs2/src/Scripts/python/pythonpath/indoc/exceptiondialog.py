#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import os, platform, subprocess, traceback, unohelper
from indoc import dialogcommons
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