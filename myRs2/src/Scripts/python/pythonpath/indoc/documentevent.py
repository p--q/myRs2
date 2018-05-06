#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# ドキュメントイベントについて。
def documentOnLoad(xscriptcontext):  # ドキュメントを開いた時。リスナー追加前。
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.setActiveSheet(doc.getSheets()["一覧"])  # 一覧シートをアクティブにする。
def documentUnLoad(xscriptcontext):  # ドキュメントを閉じた時。リスナー削除後。
	pass
