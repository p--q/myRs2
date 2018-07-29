#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import commons
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.sheet import CellFlags  # 定数
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	sheet = selection.getSpreadsheet()
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				celladdress = selection.getCellAddress()
				if celladdress.Column>0:  # 列インデックス2を含む右列をダブルクリックしたときクリップボードの行のリストの改行を削除する。
					ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
					smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
					controller = xscriptcontext.getDocument().getCurrentController()  # コントローラの取得。
					dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)					
					controller.select(sheet["A1"])  # ペーストする左上セルを選択。
					sheet.clearContents(511)
					dispatcher.executeDispatch(controller.getFrame(), ".uno:Paste", "", 0, ())  # ペースト。	
					cellranges = sheet[:, 0].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # 列インデックス0の文字列が入っているセルに限定して抽出。数値の時もありうる。
					if not len(cellranges):
						return False  # セル編集モードにしない。	
					emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # ID列の最終行インデックス+1を取得。					
					datarange = sheet[:emptyrow, 0]
					outputs = []
					buffer = []
					for datarow in datarange.getDataArray():
						txt = datarow[0]
						if txt.startswith("****"):
							continue
						elif txt.startswith("#"):
							if buffer:
								outputs[-1] = "".join([outputs[-1], *buffer])
							outputs.append(txt)
							buffer = []
						else:
							buffer.append(txt)
					systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
					systemclipboard.setContents(commons.TextTransferable("\n".join(outputs)), None)  # クリップボードにコピーする。	
					return False  # セル編集モードにしない。	
	return True  # セル編集モードにする。	
