#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import platform
from . import commons, ichiran, yotei
from com.sun.star.sheet import CellFlags  # 定数
# ドキュメントイベントについて。
def documentOnLoad(xscriptcontext):  # ドキュメントを開いた時。リスナー追加後。
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	sheets = doc.getSheets()
	if platform.system()=="Windows":  # Windowsの時はすべてのシートのフォントをMS Pゴシックにする。
		[i.setPropertyValues(("CharFontName", "CharFontNameAsian"), ("ＭＳ Ｐゴシック", "ＭＳ Ｐゴシック")) for i in sheets]
	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。
	for i in namedranges.getElementNames():  # namedrangesをイテレートするとfor文中でnamedrangesを操作してはいけない。
		if not namedranges[i].getReferredCells():
			namedranges.removeByName(i)  # 参照範囲がエラーの名前を削除する。		
	yoteisheet = sheets["予定"]  # 予定シートを取得。
	yoteivars = yotei.VARS		
	startdatecell = yoteisheet[yoteivars.dayrow, yoteivars.datacolumn]
	startdatevalue = int(startdatecell.getValue())  # 先頭の日付のシリアル値を整数で取得。空セルの時は0.0が返る。				
	if startdatevalue>0:  # シリアル値が取得できた時。	
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
		functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。		
		todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。			
		diff = todayvalue - startdatevalue  # 今日の日付と先頭の日付との差を取得。
		if diff>0:  # 先頭日付が過去の時。
			todaycolumn = yoteivars.datacolumn + diff # 今日の日付列インデックスを取得。		
			yoteivars.setSheet(yoteisheet)	# シートの変化する値を取得。	
			yoteisheet[yoteivars.datarow:yoteivars.emptyrow, yoteivars.datacolumn:todaycolumn].clearContents(CellFlags.ANNOTATION)  # 過去の日付の列のコメントをクリア。
	ichiransheet = sheets["一覧"]  # 一覧シートを取得。
	ichiranvars = ichiran.VARS
	ichiranvars.setSheet(ichiransheet)
	controller.setActiveSheet(ichiransheet)  # 一覧シートをアクティブにする。		
	ichiran.initSheet(ichiransheet, xscriptcontext)  # 一覧シートのactiveSpreadsheetChanged()で呼ばれる関数を実行。予定シートからリンクされたコメントもここで消える。
def documentUnLoad(xscriptcontext):  # ドキュメントを閉じた時。リスナー削除後。
	pass
