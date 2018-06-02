#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
from indoc import commons, ichiran
# ドキュメントイベントについて。
def documentOnLoad(xscriptcontext):  # ドキュメントを開いた時。リスナー追加前。
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	sheets = doc.getSheets()
	sheet = sheets["一覧"]  # 一覧シートを取得。
	ichiran.refreshCounts(sheet, ichiran.Ichiran(sheet))  # 一覧シートのカウントを更新する。
	sheet["Y1:Z1"].setPropertyValue("CharColor", commons.COLORS["silver"])  # カウントの文字色を設定。
	sheet["Y2:Z2"].setPropertyValue("CharColor", commons.COLORS["skyblue"])  # カウントの文字色を設定。	
	controller.setActiveSheet(sheet)  # 一覧シートをアクティブにする。
def documentUnLoad(xscriptcontext):  # ドキュメントを閉じた時。リスナー削除後。
	pass
