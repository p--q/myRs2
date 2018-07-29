#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
from com.sun.star.awt import MouseButton  # 定数



def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	sheet = selection.getSpreadsheet()
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				
				
				
				pass
# 				VARS.setSheet(sheet)
# 				celladdress = selection.getCellAddress()
# 				r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。
# 				if r<VARS.splittedrow:
# 					return mousePressedWSectionM(enhancedmouseevent, xscriptcontext)			
# 				elif r<VARS.emptyrow:
# 					return mousePressedWSectionB(enhancedmouseevent, xscriptcontext)
# 				else:  # ID列が空欄の時。キーボードからの入力は想定しない。
# 					sortRows(c)  # 昇順にソート。
# 					return False  # セル編集モードにしない。	
	return True  # セル編集モードにする。	

