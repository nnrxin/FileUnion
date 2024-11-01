;=======================================================================================================================
; by:nnrxin
; email:nnrxin@163.com
;=======================================================================================================================

;基础参数设置
#Requires AutoHotkey v2.1-alpha.14
#NoTrayIcon               ;无托盘图标
#SingleInstance Ignore    ;不能双开
KeyHistory 0
ListLines 0
SendMode "Input"
SetWinDelay 0
SetControlDelay 0
ProcessSetPriority "H"


;基础库
#Include <_BasicLibs_>
#Include <GUI\ProgressGui>
#Include <File\Path>
#Include <Class_Printers>
#Include <Class_ADO>

;专有库
#Include FU_Lib.ahk


; APP名称
APP_NAME      := "FU"
;@Ahk2Exe-Let U_NameShort = %A_PriorLine~U)(^.*")|(".*$)%
; APP全称
APP_NAME_FULL := "FileUnion"
;@Ahk2Exe-Let U_Name = %A_PriorLine~U)(^.*")|(".*$)%
; APP中文名称
APP_NAME_CN   := "文件合并FU"
;@Ahk2Exe-Let U_NameCN = %A_PriorLine~U)(^.*")|(".*$)%
; 当前版本
APP_VERSION   := "0.1.4"
;@Ahk2Exe-Let U_ProductVersion = %A_PriorLine~U)(^.*")|(".*$)%


;编译后文件名
;@Ahk2Exe-Obey U_bits, = %A_PtrSize% * 8
;@Ahk2Exe-ExeName %U_NameCN%(%U_bits%bit) %U_ProductVersion%
;编译后属性信息
;@Ahk2Exe-SetName %U_Name%
;@Ahk2Exe-SetProductVersion %U_ProductVersion%
;@Ahk2Exe-SetLanguage 0x0804
;@Ahk2Exe-SetCopyright Copyright (c) 2024 nnrxin
;编译后的图标(与脚本名同目录同名的ico文件,不存在时会报错)
;@Ahk2Exe-SetMainIcon %A_ScriptName~\.[^\.]+$~.ico%



;APP保存信息(存储在AppData)
APP_DATA_PATH := A_AppData "\" APP_NAME_FULL                          ;在系统AppData的保存位置
APP_DATA_CACHE_PATH := APP_DATA_PATH "\cache"                         ;缓存文件路径
DirCreate APP_DATA_CACHE_PATH                                         ;路径不存在时需要新建
APP_JSON := JsonConfigFile(APP_DATA_PATH "\" APP_NAME "_config.json") ;创建放在用户数据文件夹配置json类
;APP保存信息(存储在本地文件夹APP_NAME "_Data"内)
DirCreate DATA_PATH := A_ScriptDir "\" APP_NAME "_Data"               ;产生数据文件位置
LOCAL_JSON := JsonConfigFile(DATA_PATH "\" APP_NAME "_config.json")   ;创建放在本地文件夹配置json类


;安装本地文件
#Include DirInstallTo_LOCAL.ahk
if !DirInstallTo_LOCAL(DATA_PATH)    ;非覆盖安装
	MsgBox "本地文件安装错误!"


;全局参数
global G := {}

;修改系统短日期格式
sShortDate := "yyyy-MM-dd"
if sShortDate != RegRead("HKEY_CURRENT_USER\Control Panel\International", "sShortDate")
    RegWrite(sShortDate, "REG_SZ", "HKEY_CURRENT_USER\Control Panel\International", "sShortDate")


;=================================
;↓↓↓↓↓↓↓↓↓  MainGUI 构建 ↓↓↓↓↓↓↓↓↓
;=================================

;创建主GUI
MainGuiWidth := 1000, MainGuiHeight := 700
MainGui := Gui("+Resize +MinSize" MainGuiWidth "x" MainGuiHeight , APP_NAME_CN " v" APP_VERSION)   ;GUI可修改尺寸
MainGui.Show("hide w" MainGuiWidth " h" MainGuiHeight)
MainGui.MarginX := MainGui.MarginY := 0
MainGui.SetFont("s9", "微软雅黑")
;MainGui.BackColor := 0xCCE8CF   ;护眼蓝色

;增加Guitooltip
MainGui.Tips := GuiCtrlTips(MainGui)


;构建配置GUI
#Include FU_ConfigGui.ahk

/****************************************************************************************************
 **************************************************************************************************** 
 * 左半区(待合并)
 ****************************************************************************************************
 ****************************************************************************************************/
MainGui.SetFont("c0070DE bold", "微软雅黑")
FU_GBfiles_W := 270
FU_GBfiles := MainGui.Add("GroupBox", "xm+5 ym w" FU_GBfiles_W " h" MainGuiHeight-25 " AH", "待合并")
MainGui.SetFont("cDefault norm", "微软雅黑")

;当前配置
MainGui.Add("Text", "xm+15 ym+17 w55 h25 Section +0x200", "当前配置")
L_DDLconfig := MainGui.Add("DDL", "x+0 yp w145")
L_DDLconfig.lastText := LOCAL_JSON.Init("L_DDLconfig.Text", "")
L_DDLconfig.OnEvent("Change", L_DDLconfig_Change)
L_DDLconfig_Change(*) {
	if FileUnion.Configs.Has(L_DDLconfig.Text) {
		G.ActiveConfig := FileUnion.Configs.Switch(L_DDLconfig.Text)
		SB.SetText("当前配置: " L_DDLconfig.Text)
		L_BTUnion.Enabled := true
	} else {
		G.ActiveConfig := ""
		SB.SetText("配置为空")
		L_BTUnion.Enabled := false
	}
}
;配置设置
L_BTsetConfig := MainGui.Add("Button", "x+0 yp w50 h25", "设置")
L_BTsetConfig.OnEvent("Click", (*) {
	MainGui.Opt("+Disabled")
	ConfigGui.Update()
	ConfigGui.Show()
})

;文件列表
L_LVfiles := MainGui.Add("ListView", "xs y+5 w250 h" MainGuiHeight-180 " Grid AltSubmit -LV0x10 +LV0x10000 BackgroundFEFEFE AH", ["文件名","路径"])
;双击行打开文件
L_LVfiles.OnEvent("DoubleClick", (thisLV, rowI) {
	if !rowI || !FileExist(path := thisLV.GetText(rowI,2))
		return
	Run path
})
;列表选择项目变化时显示文件信息
L_LVfiles.OnEvent("ItemSelect", (thisLV, Item, Selected) {
	rowI := thisLV.GetNext()
	if rowI && Item != rowI
		return
	SB.SetText(thisLV.GetText(rowI,2))
})
;点击标题行,重新上色
L_LVfiles.OnEvent("ColClick", (*) {
	L_LVfiles.Opt("-Redraw")
	Loop L_LVfiles.GetCount()
		L_LVfiles.Colors.Row(A_Index, FileUnion.Files[L_LVfiles.GetText(A_Index,2)].color)
	L_LVfiles.Opt("+Redraw")
})
;单元格颜色设置
L_LVfiles.Colors := LV_Colors(L_LVfiles)
L_LVfiles.SetRowColorByPath := (thisLV, path, color) {
	Loop L_LVfiles.GetCount() {
		if (L_LVfiles.GetText(A_Index,2) = path) {
			L_LVfiles.Colors.Row(A_Index, color)
			L_LVfiles.Opt("+Redraw")
			break
		}
	}
}
;加载文件
LV_LoadFiles(pathArray) {
	NewFiles := FileUnion.Files.Load(pathArray)
	if NewFiles.Length = 0
		return
	L_LVfiles.Opt("-Redraw")
	for _, file in NewFiles {
		L_LVfiles.Add("Icon" L_LVfiles.LoadFileIcon(file.path), file.name, file.path)
		L_LVfiles.Colors.RowCount := i := L_LVfiles.GetCount()
		L_LVfiles.Colors.Row(i, file.color)
	}
	L_LVfiles.AdjustColumnsWidth()
	L_LVfiles.Opt("+Redraw")
	;EnableBottons(L_LVfiles.GetCount()) ; 控制按钮
	SB.SetText("文件总数: " FileUnion.files.Count)
	FU_GBfiles.Text := "待合并 ( 文件数 : " FileUnion.files.Count " )"
}




;按钮-添加文件
L_BTaddFiles := MainGui.Add("button", "xs+0 y+1 w82 h30 AY", "添加文件")
L_BTaddFiles.OnEvent("Click", (*) {
	MainGui.Opt("+OwnDialogs")    ;对话框出现时禁止操作主GUI
	PathArray := FileSelect("M", A_ScriptDir, "选择文件 - " A_ScriptName)
	if PathArray.Length > 0 {
		LV_LoadFiles(PathArray)
	}
})

;按钮-添加文件夹
L_BTaddFiles := MainGui.Add("button", "x+2 yp wp hp AYP", "添加文件夹")
L_BTaddFiles.OnEvent("Click", (*) {
	MainGui.Opt("+OwnDialogs")    ;对话框出现时禁止操作主GUI
	if Path := FileSelect("D", A_ScriptDir, "选择文件夹 - " A_ScriptName) {
		LV_LoadFiles([Path])
	}
})

;按钮-清空列表
L_BTclearFiles := MainGui.Add("button", "x+2 yp wp hp AYP", "清空列表")
L_BTclearFiles.OnEvent("Click", (*) {
	L_LVfiles.Opt("-Redraw")
	L_LVfiles.Delete()
	FileUnion.Files.Clear()
	L_LVfiles.Opt("+Redraw")
	;EnableBottons(LV.GetCount()) ; 控制按钮
	SB.SetText("清空文件列表")
	FU_GBfiles.Text := "待合并 ( 文件数 : 0 )"
})


;进度条窗口
ProgGui := ProgressGui(MainGui) 
;提取文件内容
L_BTUnion := MainGui.Add("Button", "xs y+6 w250 h60 AYP", "提取内容 >>")
L_BTUnion.OnEvent("Click", L_BTUnion_Click)
L_BTUnion_Click(thisCtrl, Info) {
	MainGui.Opt("+Disabled")

	SB.SetText("开始提取文件内容")
	FileUnion.Data.Clear()
	deepRules := G.ActiveConfig.ConvertToDeep()
	ProgGui.Start(FileUnion.Files.Count)
	for i, file in FileUnion.Files {
		ProgGui.StepStart(file.name)
		try {
			result := (file.type = "excel") ? FileUnion.LoadExcel(file, deepRules) : FileUnion.LoadWord(file, deepRules)
		} catch {
		    ProgGui.StepFinsih(0, "文件提取失败")
			file.color := "red"
		} else {
			if result {
				ProgGui.StepFinsih(1, "文件提取成功-" result)
				file.color := 0x92D050 ; 绿
			} else {
				ProgGui.StepFinsih(0, "文件无法匹配")
				file.color := "yellow"
			}	
		} 
		L_LVfiles.SetRowColorByPath(file.path, file.color)
	}
	FileUnion.Data.DeleteRepeatRow(G.ActiveConfig.GetNoRepeatFields()) ; 删除重复行
	R_LVresult.LoadRecordset()
	R_LVresult.AdjustColumnsWidth()
	ProgGui.Finsih()
	SB.SetText("总 " R_LVresult.GetCount() " 行")
	FU_GBresult.Text := "合并结果 ( 行数 : " R_LVresult.GetCount() " )"

	MainGui.Opt("-Disabled")
}




/****************************************************************************************************
 **************************************************************************************************** 
 * 右半区(合并结果)
 ****************************************************************************************************
 ****************************************************************************************************/
MainGui.SetFont("c0070DE bold", "微软雅黑")
FU_GBresult := MainGui.Add("GroupBox", "Section x" FU_GBfiles_W+10 " ym w" MainGuiWidth-FU_GBfiles_W-15 " h" MainGuiHeight-25 " AH AW", "合并结果")
MainGui.SetFont("cDefault norm", "微软雅黑")

;列表
R_LVresult := MainGui.Add("ListView", "xs+10 ys+20 w" MainGuiWidth-FU_GBfiles_W - 265 " h" MainGuiHeight - 55 " Section Count10000 Grid -LV0x10 +LV0x10000 BackgroundFEFEFE AW AH")
;从JSON文件加载参数到LV
R_LVresult.LoadRecordset := (thisLV, FormatStrs := "") {
	thisLV.Delete()
	while thisLV.GetCount("Column")
		thisLV.DeleteCol(1)
	for i, fieldName in FileUnion.Data.FieldNames {
		thisLV.InsertCol(,, fieldName)
	}
	for i, row in FileUnion.Data {
		thisLV.Add("", row*)
	}
}

;Group 筛选
MainGui.SetFont("c9382C9 bold", "微软雅黑")
MainGui.Add("GroupBox", "x+7 ym+12 Section w225 h50 AX", "筛选")
MainGui.SetFont("cDefault norm", "微软雅黑")




;Group 导出
MainGui.SetFont("c9382C9 bold", "微软雅黑")
MainGui.Add("GroupBox", "xs y+5 Section w225 h230 AXP", "导出")
MainGui.SetFont("cDefault norm", "微软雅黑")

MainGui.Add("Text", "xs+10 ys+20 w50 h25 +0x200 AXP", "导出模板")
R_EDtemplatePath := MainGui.Add("Edit", "x+0 yp w130 h25 AXP")
R_EDtemplatePath.Value := LOCAL_JSON.Init("R_EDtemplatePath.Value", "")
R_BTtemplatePath := MainGui.Add("button", "x+0 yp w25 h25 AXP", "...")
R_BTtemplatePath.OnEvent("Click", (*) {
	MainGui.Opt("+OwnDialogs")    ;对话框出现时禁止操作主GUI
	if newPath := FileSelect(1, Path_Dir(R_EDpath.Value, true, A_ScriptDir), "选择一个文件 - " A_ScriptName, "Excel/Access文件 (*.xlsx; *.xls; *.accdb; *.mdb)") {
		R_EDtemplatePath.Value := newPath
	}
})

MainGui.Add("Text", "xs+10 y+5 w50 h25 +0x200 AXP", "文件路径")
R_EDpath := MainGui.Add("Edit", "x+0 yp w130 h25 AXP")
R_EDpath.Value := LOCAL_JSON.Init("R_EDpath.Value", "")
R_BTdir := MainGui.Add("button", "x+0 yp w25 h25 AXP", "...")
R_BTdir.OnEvent("Click", (*) {
	MainGui.Opt("+OwnDialogs")    ;对话框出现时禁止操作主GUI
	if newPath := FileSelect("D", Path_Dir(R_EDpath.Value, true, A_ScriptDir), "选择一个文件夹 - " A_ScriptName) {
		R_EDpath.Value := newPath
	}
})


MainGui.Add("Text", "xs+10 y+5 w50 h25 +0x200 AXP", "文件名称")
R_EDfileName := MainGui.Add("Edit", "x+0 yp w155 h25 AXP")
R_EDfileName.Value := LOCAL_JSON.Init("R_EDfileName.Value", "FUexport")

MainGui.Add("Text", "xs+10 y+5 w50 h25 +0x200 AXP", "文件类型")
R_DDLfileExt := MainGui.Add("DDL", "x+0 yp w155 AXP", ["xlsx","xls","accdb"])
R_DDLfileExt.Value := LOCAL_JSON.Init("R_DDLfileExt.Value", 1)

R_CBaddTimestamp:= MainGui.Add("CheckBox", "xs+10 y+5 w200 h25 AXP", "生成的文件名后增加时间戳")
R_CBaddTimestamp.Value := LOCAL_JSON.Init("R_CBaddTimestamp.Value", 1)

;导出为Excel
R_BTexport := MainGui.Add("Button", "xs+7 y+5 w210 h50 AXP", "导出为文件")
R_BTexport.OnEvent("Click", R_BTexport_Click)
R_BTexport_Click(thisCtrl, Info) {
	MainGui.Opt("+Disabled")

	fileName := (R_CBaddTimestamp.Value ? R_EDfileName.Value "-" A_Now : R_EDfileName.Value) "." R_DDLfileExt.Text
	path := Path_Full((R_EDpath.Value && DirExist(R_EDpath.Value) ? R_EDpath.Value : A_ScriptDir) "\" fileName)
	if FileExist(R_EDtemplatePath.Value)
	    FileCopy(R_EDtemplatePath.Value, path)
	else if FileExist(path) ;删除已经存在的文件
		FileDelete(path)
	switch R_DDLfileExt.Text {
		case "xlsx", "xls":
			FileUnion.ExportToExcel(path)
			MsgBox "导出成功！`n`n" path
			Runwait('explorer.exe /select, "' path '"') ; 文件资源管理器中显示
		case "accdb", "mdb":
			FileUnion.ExportToAccess(path)
			MsgBox "导出成功！`n`n" path
			Runwait('explorer.exe /select, "' path '"') ; 文件资源管理器中显示
		default:
			MsgBox "目前暂不支持"
	}

	MainGui.Opt("-Disabled")
}




;状态栏
SB := MainGui.Add("StatusBar",, "")
SB.SetFont("bold italic") ; 粗体、斜体

;监听鼠标左键事件
OnMessage(0x0201, On_WM_LBUTTONDOWN)
On_WM_LBUTTONDOWN(wParam, lParam, msg, hwnd) {
	if CurrCtrl := GuiCtrlFromHwnd(hwnd) {
		if CurrCtrl.HasMethod("On_WM_LBUTTONDOWN_10")
			SetTimer () => CurrCtrl.On_WM_LBUTTONDOWN_10(), -10
	}
}

;GUI菜单
MainGui.OnEvent("ContextMenu", (GuiObj, GuiCtrlObj, Item, IsRightClick, X, Y) {
	;右键某控件上
	if IsRightClick and GuiCtrlObj and GuiCtrlObj.HasMethod("ContextMenu")
		GuiCtrlObj.ContextMenu(Item, X, Y)
})

;GUI文件拖放
MainGui.OnEvent("DropFiles", (GuiObj, GuiCtrlObj, FileArray, X, Y) {
	MainGui.Opt("+Disabled")
	;主界面上的拖动
	LV_LoadFiles(FileArray)
	;拖动到控件上:
	/*
	Switch GuiCtrlObj
	{
	case  L_LVfiles:
		LV_LoadFiles(FileArray)
	}
	*/
	MainGui.Opt("-Disabled")
})

;改变GUI尺寸时调整控件
MainGui.OnEvent("Size", (thisGui, MinMax, W, H) {
	R_LVresult.AdjustColumnsWidth()
})

;GUI关闭
MainGui.OnEvent("Close", (*) => ExitApp())

;退出APP前运行
OnExit (*) {
	MainGui.Hide()

	APP_JSON.Save()                     ; 配置保存到用户数据的json文件
	LOCAL_JSON.Save()                   ; 配置保存到本地数据的json文件

	;清空缓存文件夹
	DirDelete APP_DATA_CACHE_PATH, 1
}



;初始化
MainGui.Init := (thisGui) {
	thisGui.Opt("+Disabled")

	;拖拽文件夹启动
	LV_LoadFiles(Path_InArgs()) ; 拖拽文件到程序图标上启动
	;加载合并的配置
	FileUnion.Configs.Load(LOCAL_JSON.Init("UnionConfigs", Map())) ;从json对象加载配置
	;刷新配置DDL
	L_DDLconfig.Update(FileUnion.Configs.instances, L_DDLconfig.lastText)
	L_DDLconfig_Change()

	thisGui.Opt("-Disabled")
}
MainGui.Init()


;GUI显示
dpiRate := 96 / A_ScreenDPI
MainGui.Show("hide Center w" SysGet(16) * dpiRate " h" SysGet(17) * dpiRate)
guiSizeRate := 0.9 * dpiRate
MainGui.Show("Center w" SysGet(16) * guiSizeRate " h" SysGet(17) * guiSizeRate)
;MainGui.Show("Maximize Center w" SysGet(16) * guiSizeRate " h" SysGet(17) * guiSizeRate) ; 最大化启动
/*
dpiRate := 96 / A_ScreenDPI
guiSizeRate := 0.9 * dpiRate
MainGui.Show("Center w" SysGet(16) * guiSizeRate " h" SysGet(17) * guiSizeRate)
*/

;=========================
return    ;自动运行段结束 |
;=========================