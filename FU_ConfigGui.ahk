

;创建配置GUI
ConfigGuiWidth := 500, ConfigGuiHeight := 600
ConfigGui := Gui("-MaximizeBox -MinimizeBox", APP_NAME_CN "-配置")
;ConfigGui.Opt("+Owner")
ConfigGui.Show("hide w" ConfigGuiWidth " h" ConfigGuiHeight)
ConfigGui.MarginX := ConfigGui.MarginY := 0
ConfigGui.SetFont("s9", "微软雅黑")
;ConfigGui.BackColor := 0xCCE8CF   ;护眼蓝色
;GUI关闭
ConfigGui.OnEvent("Close", ConfigGui_Close)
ConfigGui_Close(*) {
	ConfigGui.Hide()
	;保存配置
	C1_LV.SaveRule()
	C2_LV.SaveRule()
	C3_LV.SaveRule()
	SaveAdvancedRules()
	C5_LV.SaveRule()
	; 调整主界面的插件并激活主界面
	L_DDLconfig_Change()
	;MainGui.Opt("-Disabled")
	WinActivate(MainGui.Hwnd)
}
;GUI界面刷新
ConfigGui.Update := (*) {
	C_LV.Update()
	RowI := 0
	Loop C_LV.GetCount() {
		if L_DDLconfig.Text = C_LV.GetText(A_Index) {
			C_LV.Modify(A_Index, "+Select +focus")
			RowI := A_Index
			break
		}	
	}
	C_LVconfigsUpdate(RowI)
}

;状态栏
ConfigGui_SB := ConfigGui.Add("StatusBar")
ConfigGui_SB.SetFont("bold italic") ; 粗体、斜体

;增加Guitooltip
ConfigGui.Tips := GuiCtrlTips(ConfigGui)



/************\
*            *
*    ****    *
*   ******   *
*   **  **   *
*   **  **   *
*   **  **   *
*   **  **   *
*   ******   *
*    ****    *
*            *
\************/

; 配置列表
C_LV := ConfigGui.Add("ListView", "xm+10 ym+10 w150 h" ConfigGuiHeight - 70 " Grid -Multi +LV0x10000 BackgroundFEFEFE", ["配置"])
;双击LV单元格可编辑,编辑后有变化则执行重命名函数
LV_InCellEditing(C_LV,, (this, Row, Col, OldText, NewText) {
	ConfigGui_SB.SetText("")
	if FileUnion.Configs.Has(NewText) {
		this.LV.Modify(Row, "col" Col, OldText)
		return ConfigGui_SB.SetText('重命名失败: 配置"' NewText '"已存在')
	}
	if FileUnion.Configs.ReName(OldText, NewText) {
		this.LV.Modify(Row, "col" Col, OldText)
		return ConfigGui_SB.SetText('重命名失败: 配置"' OldText '"重命名失败')
	}
	L_DDLconfig.Update(FileUnion.Configs.instances, NewText)
	ConfigGui_SB.SetText('配置"' OldText '"重命名为: "' NewText '"')
})
;列表选择项目变化
C_LV.SelectedRow := 0
C_LV.SelectedConfig := ""
C_LV.OnEvent("ItemSelect", (thisLV, Item, Selected) {
	rowI := thisLV.GetNext()
	if rowI && Item != rowI
		return
	C_LV.SelectedRow := rowI
	C_LVconfigsUpdate(rowI)
})
C_LVconfigsUpdate(rowI := 0) {
	;保存前一个配置的规则
	C1_LV.SaveRule()
	C2_LV.SaveRule(, C2_DDLruleIndex.Value)
	C3_LV.SaveRule()
	SaveAdvancedRules()
	C5_LV.SaveRule()
	;按行切换配置
	if rowI	{
		Text := C_LV.GetText(rowI)
		C_LV.SelectedConfig := FileUnion.Configs[Text]
		C2_DDLruleIndex.Update()
		C2_DDLruleIndex.Value := 1
		EnabledGroupBoxRule(true)
		ConfigGui_SB.SetText("配置: " Text)
	} else {
		C_LV.SelectedConfig := ""
		C2_DDLruleIndex.Update("")
		C2_DDLruleIndex.Value := 0
		EnabledGroupBoxRule(false)
		ConfigGui_SB.SetText("未选择配置")
	}
	;加载当前配置的规则
	C1_LV.LoadRule()
	C2_LV.LoadRule()
	C3_LV.LoadRule()
	LoadAdvancedRules()
	C5_LV.LoadRule()
}
;控制文件合并规则控件状态
EnabledGroupBoxRule(Enabled) {
	C2_EnabledButtons()
	C1_LV.Enabled := C1_BTreset.Enabled :=
	C2_DDLruleIndex.Enabled := C2_LV.Enabled := C2_BTaddKey.Enabled := C2_BTdeleteKey.Enabled :=
	C3_LV.Enabled := C3_BTaddKey.Enabled := C3_BTdeleteKey.Enabled := 
	C4_EDnoRepeatFields.Enabled := Enabled ? true : false
}
;右键某行弹出菜单
C_LV.OnEvent("ContextMenu", (thisLV, rowI, IsRightClick, X, Y) {
	;初次调用时创建菜单
	if !IsSet(MyMenu) {
		static MyMenu := Menu()
		MyMenu.rowI := 0
		MyMenu.Add("复制配置", MyMenu_Call)
		MyMenu.Add("删除配置", MyMenu_Call)
		MyMenu.Add()
		MyMenu.Add("新建配置", MyMenu_Call)
		;MyMenu.SetIcon("4&", "SHELL32.dll", 4) ; 获取文件夹图标

		MyMenu_Call(ItemName, ItemPos, MyMenu) {
			switch ItemPos {
				case 1: C_LV.CopyConfig(MyMenu.rowI)
				case 2: C_LV.DeleteConfig(MyMenu.rowI)
				case 4: C_LV.AddConfig()
			}
		}
	}
	;显示菜单
	if MyMenu.rowI := rowI {
		MyMenu.Enable("1&"), MyMenu.Enable("2&"), MyMenu.Enable("3&")
	} else {
		MyMenu.Disable("1&"), MyMenu.Disable("2&"), MyMenu.Disable("3&")
	}
	MyMenu.Show()
})
; 从FileUnion.Configs加载参数到LV
C_LV.Update := (thisLV) {
	thisLV.Delete()
	for name, config in FileUnion.Configs
		thisLV.Add(, name)
}
; 保存LV参数到FileUnion.Configs
C_LV.SaveRule := (thisLV, name?, i?) {
	/*
	if !FileUnion.Configs.Has(name ?? C_CBconfig.Text)
		return
	rule := FileUnion.Configs[name ?? C_CBconfig.Text][i ?? C2_DDLruleIndex.Value]
	rule.length := 0
	Loop thisLV.GetCount()
		rule.push([thisLV.GetText(A_Index,1), thisLV.GetText(A_Index,2), thisLV.GetText(A_Index,3)])
	*/
}
;新建配置
C_LV.AddConfig := (thisLV) {
	ConfigGui.Opt("+OwnDialogs")
	ConfigGui_SB.SetText("")
	IB := InputBox("输入一个新的配置名称:", "新建配置", "w250 h100", "")
	if IB.Result = "OK" &&  IB.Value != "" && IB.Value != "" {
		if FileUnion.Configs.Has(IB.Value)
			return ConfigGui_SB.SetText('配置创建失败: 配置"' IB.Value '"已存在')
		FileUnion.Configs.Add(IB.Value)
		L_DDLconfig.Update(FileUnion.Configs.instances, IB.Value)
		thisLV.Add(, IB.Value)
		thisLV.Modify(RowI := thisLV.GetCount(), "+Select +focus")
		C_LVconfigsUpdate(RowI)
		ConfigGui_SB.SetText("新建配置: " IB.Value)
	}
}
;删除配置
C_LV.DeleteConfig := (thisLV, RowNumber?) {
	ConfigGui_SB.SetText("")
	RowNumber := RowNumber ?? C_LV.SelectedRow
	if RowNumber = 0
		return
	FileUnion.Configs.Delete(OldText := thisLV.GetText(RowNumber))
	thisLV.Delete(RowNumber)
	L_DDLconfig.Update(FileUnion.Configs.instances)
	C_LVconfigsUpdate(0)
	ConfigGui_SB.SetText("配置已删除: " OldText)
}
;复制配置
C_LV.CopyConfig := (thisLV, RowNumber) {
	ConfigGui.Opt("+OwnDialogs")
	ConfigGui_SB.SetText("")
	text := thisLV.GetText(RowNumber)
	IB := InputBox("复制为新配置的名称:", "配置复制", "w250 h100", text "_副本")
	if IB.Result = "OK" &&  IB.Value != "" && IB.Value != text {
		if FileUnion.Configs.Has(IB.Value)
			return ConfigGui_SB.SetText('复制失败: 配置"' IB.Value '"已存在')
		FileUnion.Configs.Clone(IB.Value, text)
		L_DDLconfig.Update(FileUnion.Configs.instances, IB.Value)
		thisLV.Add(, IB.Value)
		thisLV.Modify(RowI := thisLV.GetCount(), "+Select +focus")
		C_LVconfigsUpdate(RowI)
		ConfigGui_SB.SetText("配置复制为: " IB.Value)
	}
}

;控制按钮状态
EnabledConfigButtons(configNameExist) {
	C_BTDeleteConfig.Enabled := configNameExist ? true : false 
}
;按钮-新建配置
C_BTAddConfig := ConfigGui.Add("Button", "xp y+4 w73 h26", "新建配置")
C_BTAddConfig.OnEvent("Click", (*) => C_LV.AddConfig())
;按钮-删除配置
C_BTDeleteConfig := ConfigGui.Add("Button", "x+1 yp wp hp", "删除配置")
C_BTDeleteConfig.OnEvent("Click", (*) => C_LV.DeleteConfig())







;Group 配置
ConfigGui.SetFont("c0070DE bold", "微软雅黑")
C_TABconfig := ConfigGui.Add("Tab3", "x+10 y10 Section w" ConfigGuiWidth - 180 " h" ConfigGuiHeight - 40, ["文件筛选","内容提取","正则替换","数据处理","数据导出"])
ConfigGui.SetFont("cDefault norm", "微软雅黑")
C_TABconfig.Value := C_TABconfig.lastValue := 1
/*
C_TABconfig.OnEvent("Change", (*) {
    switch C_TABconfig.lastValue {
		case 1:
			
		case 2:
			C2_LV.SaveRule()
		case 3:
			C3_LV.SaveRule()
	}
	C_TABconfig.lastValue := C_TABconfig.Value
})
*/


/***************\
*               *
*      ***      *
*     ****      *
*    *****      *
*      ***      *
*      ***      *
*      ***      *
*      ***      *
*      ***      *
*   *********   *
*   *********   *
*               *
\***************/

;文件筛选
C_TABconfig.UseTab(1)

;文件筛选LV
C1_LV := ConfigGui.Add("ListView", "xs+10 ys+30 w300 h" ConfigGuiHeight - 120 " Grid NoSort -LV0x10 +LV0x10000 BackgroundFEFEFE", ["","名称","数值"])
C1_LV.ModifyCol(1, 0)
C1_LV.ModifyCol(2, 136)
C1_LV.ModifyCol(3, 160)
;LV单元格可编辑,编辑后有变化执行函数
LV_InCellEditing(C1_LV, [3], (this, R, C, OldText, NewText) {
	ConfigGui_SB.SetText(R ? ("[" C1_LV.GetText(R,2) "] 设置为: " NewText) : "")
})
;点选项目触发动作
C1_LV.OnEvent("Click", (thisLV, R) {
	ConfigGui_SB.SetText(R ? ("[" thisLV.GetText(R,2) "] : " thisLV.GetText(R,3)) : "")
})
;从FileUnion.Configs加载参数到LV
C1_LV.LoadRule := (thisLV, config?) {
	thisLV.Delete()
	if !(config := config ?? C_LV.SelectedConfig)
		return
	for _, arr in config.FileFilter
		thisLV.Add(,, arr[1], arr[2])
}
;保存LV参数到FileUnion.Configs
C1_LV.SaveRule := (thisLV, config?) {
	if !(config := config ?? C_LV.SelectedConfig)
		return
	config.FileFilter.length := 0
	Loop thisLV.GetCount()
		config.FileFilter.push([thisLV.GetText(A_Index,2), thisLV.GetText(A_Index,3)])
}
;添加重置为默认
C1_BTreset := ConfigGui.Add("Button", "xs+10 y+5 w147 h35", "重置设置")
C1_BTreset.OnEvent("Click", (*) {
	if !C_LV.SelectedConfig
		return
	C_LV.SelectedConfig.ResetFileFilter()
	C1_LV.LoadRule()
})





/***************\
*               *
*     *****     *
*   *********   *
*   ***   ***   *
*        ***    *
*       ***     *
*      ***      *
*     ***       *
*    ***        *
*   *********   *
*   *********   *
*               *
\***************/

;提取规则 Extract
C_TABconfig.UseTab(2)

;设置表格编号
ConfigGui.Add("Text", "xs+10 ys+30 w50 h25 +0x200", "规则编号")
C2_DDLruleIndex := ConfigGui.Add("DDL", "x+0 yp w70", ["通用",1,2,3,4,5,6,7,8,9,10])
C2_DDLruleIndex.Value := C2_DDLruleIndex.lastValue := 1
C2_DDLruleIndex.OnEvent("Change", C2_DDLruleIndex_Change)
C2_DDLruleIndex_Change(param1 := "", *) {
	C2_LV.SaveRule(, C2_DDLruleIndex.lastValue)
	if param1 is Number
		C2_DDLruleIndex.Value := param1
	C2_LV.LoadRule()
	C2_EnabledButtons(C2_DDLruleIndex.Value)
}
C2_DDLruleIndex.Update := (thisDDL, config?) {
	thisDDL.Delete()
	if !(config := config ?? C_LV.SelectedConfig)
		return C2_DDLruleIndex.lastValue := 0
	thisDDL.Add(["通用"])
	loop config.Extract.Length - 1
		thisDDL.Add([A_Index])
	C2_DDLruleIndex.Value := (C2_DDLruleIndex.lastValue > config.Extract.Length) ? config.Extract.Length : C2_DDLruleIndex.lastValue
}

;控制按钮状态
C2_EnabledButtons(i?) {
	i := i ?? C2_DDLruleIndex.Value
	C2_BTaddRule.Enabled := i ? true : false
	C2_BTcopyRule.Enabled := (i >= 2) ? true : false
	C2_BTdeleteRule.Enabled := (i >= 2) ? true : false
	C2_BTresetRule.Enabled := i ? true : false
	C2_BTclearRule.Enabled := i ? true : false
}

; 按钮-新增规则
C2_BTaddRule := ConfigGui.Add("Button", "x+5 yp w35 h25", "新增")
C2_BTaddRule.OnEvent("Click", (*) {
	if !(config := C_LV.SelectedConfig)
		return
	config.AddExtractRule()
	C2_DDLruleIndex.Update()
	C2_DDLruleIndex_Change(config.Extract.Length)
})
; 按钮-复制规则
C2_BTcopyRule := ConfigGui.Add("Button", "x+0 yp wp hp", "复制")
C2_BTcopyRule.OnEvent("Click", (*) {
	if !(config := C_LV.SelectedConfig)
		return
	if C_LV.SelectedConfig.CopyExtractRule(C2_DDLruleIndex.Value) {
		C2_DDLruleIndex.Update()
		C2_DDLruleIndex_Change(config.Extract.Length)
	}
})
; 按钮-删除规则
C2_BTdeleteRule := ConfigGui.Add("Button", "x+0 yp wp hp", "删除")
C2_BTdeleteRule.OnEvent("Click", (*) {
	if !C_LV.SelectedConfig
		return
	if C_LV.SelectedConfig.RemoveExtractRule(C2_DDLruleIndex.Value) {
		C2_DDLruleIndex.Update()
		C2_LV.LoadRule()
		C2_EnabledButtons(C2_DDLruleIndex.Value)
	}
})
; 按钮-模板
C2_BTresetRule := ConfigGui.Add("Button", "x+0 yp wp hp", "模板")
C2_BTresetRule.OnEvent("Click", (*) {
	if !C_LV.SelectedConfig
		return
	C_LV.SelectedConfig.ResetExtractRule(C2_DDLruleIndex.Value)
	C2_LV.LoadRule()
})
; 按钮-清空
C2_BTclearRule := ConfigGui.Add("Button", "x+0 yp wp hp", "清空")
C2_BTclearRule.OnEvent("Click", (*) {
	if !C_LV.SelectedConfig
		return
	if C_LV.SelectedConfig.ClearExtractRule(C2_DDLruleIndex.Value)
		C2_LV.LoadRule()
})


;提取规则LV
C2_LV := ConfigGui.Add("ListView", "xs+10 y+5 w300 h" ConfigGuiHeight - 150 " Grid NoSort -LV0x10 +LV0x10000 BackgroundFEFEFE", ["","键","值","附加"])
C2_LV.ModifyCol(1, 0)
C2_LV.ModifyCol(2, 90)
C2_LV.ModifyCol(3, 130)
C2_LV.ModifyCol(4, 76)
;LV单元格可编辑,编辑后执行函数
LV_InCellEditing(C2_LV,, (thisLV, Row, Col, OldText, NewText) {
	;有变化则进行动作
})
;点选项目触发动作
C2_LV.OnEvent("Click", (thisLV, rowI) {
	ConfigGui_SB.SetText(rowI ? (thisLV.GetText(rowI,2) "   " thisLV.GetText(rowI,3) "   " thisLV.GetText(rowI,4)) : "")
})
;从FileUnion.Configs加载参数到LV
C2_LV.LoadRule := (thisLV, config?, i?) {
	thisLV.Delete()
	if !(config := config ?? C_LV.SelectedConfig)
		return C2_DDLruleIndex.lastValue := 0
	for _, arr in config.Extract[i ?? C2_DDLruleIndex.Value]
		thisLV.Add(,, arr[1], arr[2], arr[3])
	C2_DDLruleIndex.lastValue := C2_DDLruleIndex.Value
}
;保存LV参数到FileUnion.Configs
C2_LV.SaveRule := (thisLV, config?, i?) {
	if !(config := config ?? C_LV.SelectedConfig)
		return
	rule := config.Extract[i ?? C2_DDLruleIndex.Value]
	rule.length := 0
	Loop thisLV.GetCount()
		rule.push([thisLV.GetText(A_Index,2), thisLV.GetText(A_Index,3), thisLV.GetText(A_Index,4)])
}


; 按钮-添加键
C2_BTaddKey := ConfigGui.Add("Button", "xs+10 y+5 w147 h35", "添加参数")
C2_BTaddKey.OnEvent("Click", (*) => LVAddKey(C2_LV))
; 按钮-删除键
C2_BTdeleteKey := ConfigGui.Add("Button", "x+5 yp wp hp", "删除参数")
C2_BTdeleteKey.OnEvent("Click", (*) => LVDeleteKey(C2_LV))


;新增键
LVAddKey(thisLV, RowNumber?) {
	if !(IsSet(RowNumber) && IsInteger(RowNumber) && RowNumber > 0 && RowNumber <= thisLV.GetCount())
		RowNumber := thisLV.GetNext(0) || thisLV.GetCount() + 1 ; 优先插入到选中行下方，否则插入到最后一行
	RowNumber := thisLV.Insert(RowNumber,,, "key" RowNumber, "value" RowNumber)
	thisLV.Focus()
	thisLV.Modify(0, "-Select")          ;全部取消选中
	thisLV.Modify(RowNumber, "Select")   ;选中
	thisLV.Modify(RowNumber, "Vis")      ;可见
}
;删除键
LVDeleteKey(thisLV, Rows?) {
	if IsSet(Rows) && Rows is Array {
		for _, RowNumber in Rows.Sort((a, b) => b - a) ; 倒序
			if IsInteger(RowNumber) && RowNumber > 0 && RowNumber <= thisLV.GetCount()
				thisLV.Delete(RowNumber)
	} else {
		selectRows := []
		RowNumber := 0  ; 这样使得首次循环从列表的顶部开始搜索.
		Loop {
			RowNumber := thisLV.GetNext(RowNumber)  ; 在前一次找到的位置后继续搜索.
			if not RowNumber  ; 上面返回零, 所以选择的行已经都找到了.
				break
			selectRows.Push(RowNumber)
		}
		for _, RowNumber in selectRows.Reverse()
			thisLV.Delete(RowNumber)
	}
}


/***************\
*               *
*   *********   *
*   *********   *
*        ***    *
*       ***     *
*      ***      *
*        ***    *
*         ***   *
*   ***   ***   *
*   *********   *
*     *****     *
*               *
\***************/

;正则替换
C_TABconfig.UseTab(3)


;正则替换规则LV
C3_LV := ConfigGui.Add("ListView", "xs+10 ys+30 w300 h" ConfigGuiHeight - 120 " Grid NoSort -LV0x10 +LV0x10000 BackgroundFEFEFE", ["","字段","替换字符","替换为"])
C3_LV.ModifyCol(1, 0)
C3_LV.ModifyCol(2, 100)
C3_LV.ModifyCol(3, 100)
C3_LV.ModifyCol(4, 96)
;LV单元格可编辑,编辑后执行函数
LV_InCellEditing(C3_LV,, (thisLV, Row, Col, OldText, NewText) {
	;有变化则进行动作
})
;点选项目触发动作
C3_LV.OnEvent("Click", (thisLV, rowI) {
	ConfigGui_SB.SetText(rowI ? (thisLV.GetText(rowI,2) "   " thisLV.GetText(rowI,3) "   " thisLV.GetText(rowI,4)) : "")
})
;从FileUnion.Configs加载参数到LV
C3_LV.LoadRule := (thisLV, config?) {
	thisLV.Delete()
	if !(config := config ?? C_LV.SelectedConfig)
		return
	for _, arr in config.RegExReplaceRules
		thisLV.Add(,, arr[1], arr[2], arr[3])
}
;保存LV参数到FileUnion.Configs
C3_LV.SaveRule := (thisLV, config?) {
	if !(config := config ?? C_LV.SelectedConfig)
		return
	config.RegExReplaceRules.length := 0
	Loop thisLV.GetCount()
		config.RegExReplaceRules.push([thisLV.GetText(A_Index,2), thisLV.GetText(A_Index,3), thisLV.GetText(A_Index,4)])
}

;添加按钮
C3_BTaddKey := ConfigGui.Add("Button", "xs+10 y+5 w147 h35", "添加参数")
C3_BTaddKey.OnEvent("Click", (*) => LVAddKey(C3_LV))
;删除按钮
C3_BTdeleteKey := ConfigGui.Add("Button", "x+5 yp wp hp", "删除参数")
C3_BTdeleteKey.OnEvent("Click", (*) => LVDeleteKey(C3_LV))



/***************\
*               *
*         ***   *
*        ****   *
*       *****   *
*      ******   *
*     *** ***   *
*    ***  ***   *
*   *********   *
*   *********   *
*         ***   *
*         ***   *
*               *
\***************/

;数据处理
C_TABconfig.UseTab(4)

;加载当前配置的高级规则
LoadAdvancedRules(config?) {
	C4_EDnoRepeatFields.Value := ""

	if !(config := config ?? C_LV.SelectedConfig)
		return
	advanceRules := config.advanceRules

	C4_EDnoRepeatFields.Value := advanceRules.Has("noRepeatFields") ? advanceRules["noRepeatFields"] : ""
}
;保存高级规则到当前配置
SaveAdvancedRules(config?) {
	if !(config := config ?? C_LV.SelectedConfig)
		return
	advanceRules := config.advanceRules

	advanceRules["noRepeatFields"] := C4_EDnoRepeatFields.Value
}

;忽略重复行
ConfigGui.Add("Text", "xs+10 ys+33 w60 h25 +0x200", "忽略重复项")
C4_EDnoRepeatFields := ConfigGui.Add("Edit", "x+2 yp w240 hp")
ConfigGui.Tips.SetTip(C4_EDnoRepeatFields, "填写用[]括起来字段名,可以是多个的组合")
C4_EDnoRepeatFields.OnEvent("Change", (thisCtrl, Info) {
	ConfigGui.Opt("+Disabled")
	;LV_LoadDir(true)
	ConfigGui.Opt("-Disabled")
})



/***************\
*               *
*   *********   *
*   *********   *
*   ***         *
*   ********    *
*   *********   *
*         ***   *
*         ***   *
*   ***   ***   *
*   *********   *
*    *******    *
*               *
\***************/



;高级规则
C_TABconfig.UseTab(5)

;文件筛选LV
C5_LV := ConfigGui.Add("ListView", "xs+10 ys+30 w300 h" ConfigGuiHeight - 120 " Grid NoSort -LV0x10 +LV0x10000 BackgroundFEFEFE", ["","名称","数值"])
C5_LV.ModifyCol(1, 0)
C5_LV.ModifyCol(2, 106)
C5_LV.ModifyCol(3, 190)
;LV单元格可编辑,编辑后有变化执行函数
LV_InCellEditing(C5_LV, [3], (this, R, C, OldText, NewText) {
	switch C5_LV.GetText(R,2) {
		case "导出模板":
			if NewText = "" || InStr(FileGetAttrib(NewText), "A") && ["xls","xlsx"].IndexOf(Path_Ext(NewText))
				this.LV.Modify(R, "col" C, NewText := NewText)
			else {
				this.LV.Modify(R, "col" C, OldText)
				return ConfigGui_SB.SetText("[" C5_LV.GetText(R,2) "] 设置失败: " NewText " 不是有效路径" )
			}
		case "导出文件类型":
			if !["","xls","xlsx","accdb"].IndexOf(NewText) {
				this.LV.Modify(R, "col" C, OldText)
				return ConfigGui_SB.SetText("[" C5_LV.GetText(R,2) "] 设置失败: " NewText " 只能是xls,xlsx,accdb中的一个" )
			}
		case "导出文件路径":
		case "导出文件名":
		case "文件名后时间戳":
	}
	ConfigGui_SB.SetText(R ? ("[" C5_LV.GetText(R,2) "] 设置为: " NewText) : "")
})
;点选项目触发动作
C5_LV.OnEvent("Click", (thisLV, R) {
	ConfigGui_SB.SetText(R ? ("[" thisLV.GetText(R,2) "] : " thisLV.GetText(R,3)) : "")
})
;从FileUnion.Configs加载参数到LV
C5_LV.LoadRule := (thisLV, config?) {
	thisLV.Delete()
	if !(config := config ?? C_LV.SelectedConfig)
		return
	for _, arr in config.ExportRules
		thisLV.Add(,, arr[1], arr[2])
}
;保存LV参数到FileUnion.Configs
C5_LV.SaveRule := (thisLV, config?) {
	if !(config := config ?? C_LV.SelectedConfig)
		return
	config.ExportRules.length := 0
	Loop thisLV.GetCount()
		config.ExportRules.push([thisLV.GetText(A_Index,2), thisLV.GetText(A_Index,3)])
}
;添加重置为默认
C5_BTreset := ConfigGui.Add("Button", "xs+10 y+5 w147 h35", "重置设置")
C5_BTreset.OnEvent("Click", (*) {
	if !C_LV.SelectedConfig
		return
	C_LV.SelectedConfig.ResetExportRules()
	C5_LV.LoadRule()
})




/***************\
*****************
*****************
*****************
*****************
*****************
*****************
*****************
*****************
*****************
*****************
*****************
*****************
\***************/

;=========================
;热键
;=========================
#HotIf (WinActive("ahk_id " ConfigGui.Hwnd))
;插入
~Insert:: {
	Switch ConfigGui.FocusedCtrl
	{
	case C_LV: C_LV.AddConfig()
	;case C2_LV: LVAddKey(C2_LV)
	;case C3_LV: LVAddKey(C3_LV)
	}
}
;删除
~Delete:: {
	Switch ConfigGui.FocusedCtrl
	{
	case C_LV: C_LV.DeleteConfig()
	;case C2_LV: LVDeleteKey(C2_LV)
	;case C3_LV: LVDeleteKey(C3_LV)
	}
}