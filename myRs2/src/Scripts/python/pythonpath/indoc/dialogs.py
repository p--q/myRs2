#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-


def createDialog(ctx, smgr, doc):	
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	containerwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	toolkit = containerwindow.getToolkit()  # ピアからツールキットを取得。  
	m = 6  # コントロール間の間隔
	grid = {"PositionX": m, "PositionY": m, "Width": 100, "Height": 50, "ShowRowHeader": False, "ShowColumnHeader": False, "SelectionModel": SINGLE, "VScroll": True, "PosSize": POSSIZE}  # グリッドコントロールの基本プロパティ。
	textbox = {"PositionX": m, "PositionY": YHeight(grid, m), "Height": 12, "PosSize": POSSIZE}  # テクストボックスコントロールの基本プロパティ。
	button = {"PositionY": textbox["PositionY"]-1, "Width": 23, "Height":textbox["Height"]+2, "PushButtonType": 2, "PosSize": POSSIZE}  # ボタンの基本プロパティ。PushButtonTypeの値はEnumではエラーになる。VerticalAlignではtextboxと高さが揃わない。
	controlcontainerprops =  {"PositionX": 100, "PositionY": 40, "Width": grid["PositionX"]+grid["Width"]+m, "Moveable": True, "PosSize": POSSIZE}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さは最後に設定する。
	maTopx = createConverters(containerwindow)  # ma単位をピクセルに変換する関数を取得。
	controlcontainer, addControl = controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。
	menulistener = MenuListener()  # コンテクストメニューのリスナー。
	mouselistener = MouseListener(doc, menulistener, menuCreator(ctx, smgr))
	gridselectionlistener = GridSelectionListener()
	gridcontrol1 = addControl("Grid", grid, {"addMouseListener": mouselistener, "addSelectionListener": gridselectionlistener})  # グリッドコントロールの取得。
	gridmodel = gridcontrol1.getModel()  # グリッドコントロールモデルの取得。
	gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
	column0 = gridcolumn.createColumn()  # 列の作成。
	column0.ColumnWidth = 50  # 列幅。
	gridcolumn.addColumn(column0)  # 列を追加。
	column1 = gridcolumn.createColumn()  # 列の作成。
	column1.ColumnWidth = grid["Width"] - column0.ColumnWidth  #  列幅。列の合計がグリッドコントロールの幅に一致するようにする。
	gridcolumn.addColumn(column1)  # 列を追加。	
	griddata = gridmodel.getPropertyValue("GridDataModel")  # GridDataModel
	datarows = getSavedGridRows(doc, "Grid1")  # グリッドコントロールの行をhistoryシートのragenameから取得する。	
	now = datetime.now()  # 現在の日時を取得。
	d = now.date().isoformat()
	t = now.time().isoformat().split(".")[0]	
	if datarows:  # 行のリストが取得出来た時。
		griddata.insertRows(0, ("",)*len(datarows), datarows)  # グリッドに行を挿入。
	else:
		griddata.addRow("", (t, d))  # 現在の行を入れる。
	textbox1, textbox2 = [textbox.copy() for dummy in range(2)]
	textbox1["Width"] = 34
	textbox1["Text"] = t
	textbox2["PositionX"] = XWidth(textbox1) 
	textbox2["Width"] = 42
	textbox2["Text"] = d
	addControl("Edit", textbox1)  
	addControl("Edit", textbox2, {"addMouseListener": mouselistener})  
	button["Label"] = "~Close"
	button["PositionX"] = XWidth(textbox2) 
	addControl("Button", button, {"addMouseListener": mouselistener})  
	
	#  setPosSizeで指定しないといけない。
	controlcontainer.getModel().setPropertyValue("Height", YHeight(button, m))  # コントロールダイアログの高さを設定。
	
	
	taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
	
	# controlcontainerの大きさからウィンドウの大きさを指定。
	args = NamedValue("PosSize", Rectangle(*maTopx(100, 40), *maTopx(100, 50))), NamedValue("FrameName", "GridExample")  # , NamedValue("MakeVisible", True)  # TaskCreatorで作成するフレームのコンテナウィンドウのプロパティ。

	
	dialogframe = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。
	dialogframe.setTitle("Grid Example")  # フレームのタイトルを設定。ダイアログウィンドウのタイトルになる。
	docframe.getFrames().append(dialogframe)  # 新しく作ったフレームを既存のフレームの階層に追加する。
	dialogwindow = dialogframe.getContainerWindow()  # ダイアログフレームからコンテナウィンドウを取得。
	controlcontainer.createPeer(toolkit, dialogwindow)  # ダイアログウィンドウにコントロールを描画。
	menulistener.setDialog(controlcontainer)
	controlcontainer.setVisible(True)  # コントロールコンテナの表示。
	dialogwindow.setVisible(True)  # ウィンドウの表示
	args = mouselistener, gridselectionlistener
	dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。


