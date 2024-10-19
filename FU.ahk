;=======================================================================================================================
; by:nnrxin
; email:nnrxin@163.com
;=======================================================================================================================

;基础参数设置
#Requires AutoHotkey v2.0
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
APP_VERSION   := "0.0.5"
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


/****************************************************************************************************
 **************************************************************************************************** 
 * 左半区(待合并)
 ****************************************************************************************************
 ****************************************************************************************************/
MainGui.SetFont("c0070DE bold", "微软雅黑")
FU_GBfiles_W := 500
FU_GBfiles := MainGui.Add("GroupBox", "xm+5 ym w" FU_GBfiles_W " h" MainGuiHeight-25 " AH", "待合并")
MainGui.SetFont("cDefault norm", "微软雅黑")


;Group 文件添加规则
MainGui.SetFont("c9382C9 bold", "微软雅黑")
L_GBsending := MainGui.Add("GroupBox", "xm+15 ym+17 Section w250 h73", "文件添加规则")
MainGui.SetFont("cDefault norm", "微软雅黑")

;设置忽略文重复件名
L_CBFliesNoRepeat := MainGui.Add("CheckBox", "xs+10 ys+17 w210 h25", "忽略完全相同的文件")
L_CBFliesNoRepeat.Value := LOCAL_JSON.Init("L_CBFliesNoRepeat.Value", "1")
L_CBFliesNoRepeat.ToolTip := "勾选后会忽略文件名和修改日期完全相同的文件"
L_CBFliesNoRepeat.OnEvent("Click", L_CBFliesNoRepeat_Click)
L_CBFliesNoRepeat_Click(thisCtrl, Info) {
	MainGui.Opt("+Disabled")
	LV_LoadDir(true)
	MainGui.Opt("-Disabled")
}

;设置忽略文件大小小于XX的附件
L_CBLimitFileSizeKB := MainGui.Add("CheckBox", "xp y+0 w140 h25", "忽略文件尺寸(KB)小于")
L_CBLimitFileSizeKB.Value := LOCAL_JSON.Init("L_CBLimitFileSizeKB.Value", "1")
L_CBLimitFileSizeKB.OnEvent("Click", L_CBLimitFileSizeKB_Click)
L_CBLimitFileSizeKB_Click(thisCtrl, Info) {
	MainGui.Opt("+Disabled")
	L_EDLimitFileSizeKB.Enabled := thisCtrl.Value = 1 ? true : false
	LV_LoadDir(true)
	MainGui.Opt("-Disabled")
}
L_EDLimitFileSizeKB := MainGui.Add("Edit", "x+0 yp w65 h25 Center Number Limit6")
L_EDLimitFileSizeKB.Value := LOCAL_JSON.Init("L_EDLimitFileSizeKB.Value", "1")
L_EDLimitFileSizeKB.OnEvent("Change", L_EDLimitFileSizeKB_Change)
L_EDLimitFileSizeKB_Change(thisCtrl, Info) {
	MainGui.Opt("+Disabled")
	LV_LoadDir(true)
	MainGui.Opt("-Disabled")
}


MainGui.Add("Text", "xs y+13 w60 h25 Center +0x200 Section", "文件夹路径")
L_EDpath := MainGui.Add("Edit", "x+0 yp w165 hp ReadOnly")
L_BTdir := MainGui.Add("button", "x+0 yp w25 h25", "...")
L_BTdir.OnEvent("Click", (*) {
	MainGui.Opt("+OwnDialogs")    ;对话框出现时禁止操作主GUI
	if newPath := FileSelect("D", Path_Dir(L_EDpath.Value, true, A_ScriptDir), "选择一个文件夹 - " A_ScriptName) {
		L_EDpath.Value := newPath
		LV_LoadDir(true)
	}
})

L_LVfiles := MainGui.Add("ListView", "xs y+2 w250 h" MainGuiHeight-160 " Grid AltSubmit BackgroundFEFEFE AH", ["序号", "文件名"])
;列表加载文件
LV_LoadDir(force := false) {
	dirPath := L_EDpath.Value
	noRepeat := L_CBFliesNoRepeat.Value = 1 ? true : false
	limitFileSizeKB := L_CBLimitFileSizeKB.Value = 1 and IsNumber(L_EDLimitFileSizeKB.Value) ? L_EDLimitFileSizeKB.Value : -1
	files := FileUnion.LoadDir(dirPath, force, noRepeat, limitFileSizeKB)
	;开始加载
	L_LVfiles.Opt("-Redraw")
	L_LVfiles.Delete()
	for i, flie in files
		L_LVfiles.Add("Icon" L_LVfiles.LoadFileIcon(flie.path), i, flie.name)
	L_LVfiles.AdjustColumnsWidth()
	L_LVfiles.Opt("+Redraw")
	SB.SetText("文件总数: " files.Length)
}



;Group 配置
MainGui.SetFont("c9382C9 bold", "微软雅黑")
L_GBsending := MainGui.Add("GroupBox", "x+7 ym+17 Section w225 h83", "配置")
MainGui.SetFont("cDefault norm", "微软雅黑")

;设置表格编号
MainGui.Add("Text", "xs+10 ys+20 w50 h25 +0x200", "当前配置")
L_CBconfig := MainGui.Add("ComboBox", "x+0 yp w155")
L_CBconfig.lastText := LOCAL_JSON.Init("L_CBconfig.Text", "")
L_CBconfig.OnEvent("Change", L_CBconfig_Change)
L_CBconfig_Change(*) {
	if FileUnion.Configs.Has(L_CBconfig.Text) {
		G.ActiveConfig := FileUnion.Configs.Switch(L_CBconfig.Text)
		SB.SetText("配置切换为: " L_CBconfig.Text)
		EnabledConfigButtons(true)
		EnabledGroupBoxRule(true)
	} else {
		G.ActiveConfig := ""
		EnabledConfigButtons(false)
		EnabledGroupBoxRule(false)
		if L_CBconfig.Text = ""
			L_BTAddConfig.Enabled := false
	}
	L_LVrule.SaveRule(L_CBconfig.lastText, L_DDLruleIndex.Value)
	L_DDLruleIndex.lastValue := L_DDLruleIndex.Value := 1
	L_LVrule.LoadRule()
	L_CBconfig.lastText := L_CBconfig.Text
}
;控制按钮状态
EnabledConfigButtons(configNameExists) {
	L_BTAddConfig.Enabled := configNameExists ? false : true
	L_BTRenameConfig.Enabled := L_BTCopyConfig.Enabled := L_BTDeleteConfig.Enabled := configNameExists ? true : false
}
;按钮-新建配置
L_BTAddConfig := MainGui.Add("Button", "xs+9 y+4 w51 h26", "新建")
L_BTAddConfig.OnEvent("Click", L_BTAddConfig_Click)
L_BTAddConfig_Click(thisCtrl, Info) {
	FileUnion.Configs.Add(L_CBconfig.Text)
	L_CBconfig.Update(FileUnion.Configs.instances, L_CBconfig.Text)
	L_CBconfig_Change()
	SB.SetText("新建配置: " L_CBconfig.Text)
}
;按钮-重命名
L_BTRenameConfig := MainGui.Add("Button", "x+1 yp wp hp", "重命名")
L_BTRenameConfig.OnEvent("Click", L_BTRenameConfig_Click)
L_BTRenameConfig_Click(thisCtrl, Info) {
	MainGui.Opt("+OwnDialogs")
	IB := InputBox("输入一个新的配置名称:", "配置重命名", "w250 h100", L_CBconfig.Text)
	if IB.Result = "OK" &&  IB.Value != "" && IB.Value != L_CBconfig.Text {
		if FileUnion.Configs.Has(IB.Value)
			return SB.SetText('重命名失败: 配置"' IB.Value '"已存在')
		if FileUnion.Configs.ReName(L_CBconfig.Text, IB.Value)
			return SB.SetText('重命名失败: 配置文件"' IB.Value '"重命名失败')
		L_CBconfig.Update(FileUnion.Configs.instances, L_CBconfig.Text := IB.Value)
		L_CBconfig_Change()
		SB.SetText("配置重命名为: " L_CBconfig.Text)
	}
}
;按钮-复制配置
L_BTCopyConfig := MainGui.Add("Button", "x+1 yp wp hp", "复制")
L_BTCopyConfig.OnEvent("Click", L_BTCopyConfig_Click)
L_BTCopyConfig_Click(thisCtrl, Info) {
	MainGui.Opt("+OwnDialogs")
	IB := InputBox("复制为新配置的名称:", "配置复制", "w250 h100", L_CBconfig.Text "_副本")
	if IB.Result = "OK" &&  IB.Value != "" && IB.Value != L_CBconfig.Text {
		if FileUnion.Configs.Has(IB.Value)
			return SB.SetText('复制失败: 配置"' IB.Value '"已存在')
		FileUnion.Configs.Clone(IB.Value, L_CBconfig.Text)
		L_CBconfig.Update(FileUnion.Configs.instances, L_CBconfig.Text := IB.Value)
		L_CBconfig_Change()
		SB.SetText("配置复制为: " L_CBconfig.Text)
	}
}
;按钮-删除配置
L_BTDeleteConfig := MainGui.Add("Button", "x+1 yp wp hp", "删除")
L_BTDeleteConfig.OnEvent("Click", L_BTDeleteConfig_Click)
L_BTDeleteConfig_Click(thisCtrl, Info) {
	FileUnion.Configs.Delete(OldText := L_CBconfig.Text)
	L_CBconfig.Update(FileUnion.Configs.instances)
	L_CBconfig_Change()
	SB.SetText("配置已删除: " OldText)
}





;Group 文件合并规则
MainGui.SetFont("c9382C9 bold", "微软雅黑")
L_GBsending := MainGui.Add("GroupBox", "xs y+10 Section w225 h480 AH", "文件合并规则")
MainGui.SetFont("cDefault norm", "微软雅黑")

;设置表格编号
MainGui.Add("Text", "xs+10 ys+20 w50 h25 +0x200", "规则编号")
L_DDLruleIndex := MainGui.Add("DDL", "x+0 yp w65", [1,2,3,4,5,6,7,8,9,10])
L_DDLruleIndex.lastValue := L_DDLruleIndex.Value := LOCAL_JSON.Init("L_DDLruleIndex.Value", 1)
L_DDLruleIndex.OnEvent("Change", (*) {
	L_LVrule.SaveRule(, L_DDLruleIndex.lastValue)
	L_LVrule.LoadRule()
	L_DDLruleIndex.lastValue := L_DDLruleIndex.Value
})
;按钮-模板
L_BTdefaultRule := MainGui.Add("Button", "x+10 yp w40 h25", "模板")
L_BTdefaultRule.OnEvent("Click", L_BTdefaultRule_Click)
L_BTdefaultRule_Click(thisCtrl, Info) {
	if !G.ActiveConfig
		return
	G.ActiveConfig[L_DDLruleIndex.Value] := FileUnion.Configs.GetDefaultRule()
	L_LVrule.LoadRule()
}
;按钮-清空
L_BTclearRule := MainGui.Add("Button", "x+0 yp wp hp", "清空")
L_BTclearRule.OnEvent("Click", L_BTclearRule_Click)
L_BTclearRule_Click(thisCtrl, Info) {
	if !G.ActiveConfig
		return
	G.ActiveConfig[L_DDLruleIndex.Value].Length := 0
	L_LVrule.LoadRule()
}

;位置信息
L_LVrule := MainGui.Add("ListView", "xs+8 y+5 w207 h390 Grid -ReadOnly BackgroundFEFEFE AH", ["键","值","附加"])
L_LVrule.ModifyCol(1, 70)
L_LVrule.ModifyCol(2, 70)
L_LVrule.ModifyCol(3, 63)
;LV单元格可编辑,编辑后执行函数
LV_InCellEditing(L_LVrule, (thisLV, Row, Col, OldText, NewText) {
	;有变化则进行动作
})
;点选项目触发动作
L_LVrule.OnEvent("Click", (thisLV, rowI) {
	SB.SetText(rowI ? (thisLV.GetText(rowI,1) "   " thisLV.GetText(rowI,2) "   " thisLV.GetText(rowI,3)) : "")
})
;从FileUnion.Configs加载参数到LV
L_LVrule.LoadRule := (thisLV) {
	thisLV.Delete()
	if !G.ActiveConfig
		return
	for _, arr in G.ActiveConfig[L_DDLruleIndex.Value]
		thisLV.Add(, arr[1], arr[2], arr[3])
}
;保存LV参数到FileUnion.Configs
L_LVrule.SaveRule := (thisLV, name?, i?) {
	if !FileUnion.Configs.Has(name ?? L_CBconfig.Text)
		return
	rule := FileUnion.Configs[name ?? L_CBconfig.Text][i ?? L_DDLruleIndex.Value]
	rule.length := 0
	Loop thisLV.GetCount()
		rule.push([thisLV.GetText(A_Index,1), thisLV.GetText(A_Index,2), thisLV.GetText(A_Index,3)])
}

;添加按钮
L_BTaddKey := MainGui.Add("Button", "xs+8 y+1 w103 h30 AY", "添加参数")
L_BTaddKey.OnEvent("Click", L_BTaddKey_Click)
L_BTaddKey_Click(thisCtrl, Info) {
	RowNumber := L_LVrule.GetNext(0) || L_LVrule.GetCount() + 1 ; 优先插入到选中行下方，否则插入到最后一行
	RowNumber := L_LVrule.Insert(RowNumber,, "key" RowNumber, "value" RowNumber)
	L_LVrule.Focus()
	L_LVrule.Modify(0, "-Select")          ;全部取消选中
	L_LVrule.Modify(RowNumber, "Select")   ;选中
	L_LVrule.Modify(RowNumber, "Vis")      ;可见
}
;删除按钮
L_BTdeleteKey := MainGui.Add("Button", "x+1 yp wp hp AYP", "删除参数")
L_BTdeleteKey.OnEvent("Click", L_BTdeleteKey_Click)
L_BTdeleteKey_Click(thisCtrl, Info) {
	selectRows := []
	RowNumber := 0  ; 这样使得首次循环从列表的顶部开始搜索.
	Loop {
		RowNumber := L_LVrule.GetNext(RowNumber)  ; 在前一次找到的位置后继续搜索.
		if not RowNumber  ; 上面返回零, 所以选择的行已经都找到了.
			break
		selectRows.Push(RowNumber)
	}
	for _, RowNumber in selectRows.Reverse()
		L_LVrule.Delete(RowNumber)
}

;控制文件合并规则控件状态
EnabledGroupBoxRule(s) {
	L_BTUnion.Enabled := L_DDLruleIndex.Enabled := L_BTdefaultRule.Enabled := L_BTclearRule.Enabled := L_LVrule.Enabled := L_BTaddKey.Enabled := L_BTdeleteKey.Enabled := s ? true : false
}



;进度条窗口
ProgGui := ProgressGui(MainGui) 
;提取文件内容
L_BTUnion := MainGui.Add("Button", "xs y+13 w225 h78 AYP", "提取内容 >>")
L_BTUnion.OnEvent("Click", L_BTUnion_Click)
L_BTUnion_Click(thisCtrl, Info) {
	MainGui.Opt("+Disabled")

	SB.SetText("开始提取文件内容", 2)
	L_LVrule.SaveRule()
	FileUnion.Data.Clear()
	deepRules := G.ActiveConfig.ConvertToDeep()
	ProgGui.Start(FileUnion.files.Length)
	for i, file in FileUnion.files {
		ProgGui.StepStart(file.name)
		result := (file.type = "excel") ? FileUnion.LoadExcel(file, deepRules) : FileUnion.LoadWord(file, deepRules)
		try {
			result := result
			;result := (file.type = "excel") ? FileUnion.LoadExcel(file, deepRules) : FileUnion.LoadWord(file, deepRules)
		} catch {
		    ProgGui.StepFinsih(0, "文件提取失败")
		} else {
			if result
				ProgGui.StepFinsih(1, "文件提取成功-" result)
			else
				ProgGui.StepFinsih(0, "文件无法匹配")
		}	
	}
	R_LVresult.LoadRecordset()
	R_LVresult.AdjustColumnsWidth()
	ProgGui.Finsih()
	SB.SetText("总 " R_LVresult.GetCount() " 行", 2)

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
R_LVresult := MainGui.Add("ListView", "xs+10 ys+20 w" MainGuiWidth-FU_GBfiles_W - 265 " h" MainGuiHeight - 55 " Section Count10000 Grid -ReadOnly BackgroundFEFEFE AW AH")
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

;Group 导出
MainGui.SetFont("c9382C9 bold", "微软雅黑")
L_GBsending := MainGui.Add("GroupBox", "x+7 ym+12 Section w225 h560 AX", "导出")
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
SB.SetFont("bold italic")
SB.SetParts(FU_GBfiles_W + 7) ; 按左右分区


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
	for i, path in FileArray {
		if DirExist(path) {
			dirPath := path
			break
		}
	}
	if !IsSet(dirPath) {
		SB.SetTextWithAutoEmpty("需要拖入文件夹!")
		return
	}
	;主界面上的拖动
	MainGui.Opt("+Disabled")
	Switch GuiCtrlObj
	{
	case L_EDpath, L_LVfiles:
		L_EDpath.Value := dirPath
		MainGui.Tips.SetTip(L_EDpath, L_EDpath.Value)
		LV_LoadDir(true)
	}
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

	L_LVrule.SaveRule()                 ; 保存当前界面的LV
	APP_JSON.Save()                     ; 配置保存到用户数据的json文件
	LOCAL_JSON.Save()                   ; 配置保存到本地数据的json文件

	;清空缓存文件夹
	DirDelete APP_DATA_CACHE_PATH, 1
}



;初始化
MainGui.Init := (thisGui) {
	thisGui.Opt("+Disabled")

	;拖拽文件夹启动
	for i, path in Path_InArgs() {    ;拖拽文件到程序图标上启动
		if DirExist(path) {
			L_EDpath.Value := path
			MainGui.Tips.SetTip(L_EDpath, L_EDpath.Value)
			LV_LoadDir(true)
			break
		}
	}
	;加载合并的配置
	FileUnion.Configs.Load(LOCAL_JSON.Init("UnionConfigs", Map())) ;从json对象加载配置
	;刷新配置CB               
	L_CBconfig.Update(FileUnion.Configs.instances, L_CBconfig.lastText)
	if L_CBconfig.Value { 
		G.ActiveConfig := FileUnion.Configs.Switch(L_CBconfig.Text)
		L_LVrule.LoadRule()
		EnabledConfigButtons(true)
		EnabledGroupBoxRule(true)
	} else {
		G.ActiveConfig := ""
		SB.SetText("未找到配置,请新建配置")
		EnabledConfigButtons(false)
		EnabledGroupBoxRule(false)
		if L_CBconfig.Text = ""
			L_BTAddConfig.Enabled := false
	}
	thisGui.Opt("-Disabled")
}
MainGui.Init()


;GUI显示
dpiRate := 96 / A_ScreenDPI
MainGui.Show("hide Center w" SysGet(16) * dpiRate " h" SysGet(17) * dpiRate)
guiSizeRate := 0.9 * dpiRate
MainGui.Show("Maximize Center w" SysGet(16) * guiSizeRate " h" SysGet(17) * guiSizeRate)
/*
dpiRate := 96 / A_ScreenDPI
guiSizeRate := 0.9 * dpiRate
MainGui.Show("Center w" SysGet(16) * guiSizeRate " h" SysGet(17) * guiSizeRate)
*/

;=========================
return    ;自动运行段结束 |
;=========================