

;创建配置GUI
ConfigGuiWidth := 500, ConfigGuiHeight := 600
ConfigGui := Gui("-MaximizeBox -MinimizeBox", APP_NAME_CN "-配置")
;ConfigGui.Opt("+Owner")
ConfigGui.Show("hide w" ConfigGuiWidth " h" ConfigGuiHeight)
ConfigGui.MarginX := ConfigGui.MarginY := 0
ConfigGui.SetFont("s9", "微软雅黑")
;ConfigGui.BackColor := 0xCCE8CF   ;护眼蓝色
;GUI关闭
ConfigGui.OnEvent("Close", (*) {
	;保存配置
	C_LVrule.SaveRule()
	C_LVprocess.SaveRule()
	SaveAdvancedRules()
	; 调整主界面的插件并激活主界面
	L_DDLconfig_Change()
	MainGui.Opt("-Disabled")
	WinActivate(MainGui.Hwnd)
})
;GUI界面刷新
ConfigGui.Update := (*) {
	C_LVconfigs.Update()
	RowI := 0
	Loop C_LVconfigs.GetCount() {
		if L_DDLconfig.Text = C_LVconfigs.GetText(A_Index) {
			C_LVconfigs.Modify(A_Index, "+Select +focus")
			RowI := A_Index
			break
		}	
	}
	C_LVconfigsUpdate(RowI)
}

;状态栏
ConfigGui_SB := ConfigGui.Add("StatusBar")

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
C_LVconfigs := ConfigGui.Add("ListView", "xm+10 ym+10 w150 h" ConfigGuiHeight - 70 " Grid -Multi +LV0x10000 BackgroundFEFEFE", ["配置"])
;列表选择项目变化
C_LVconfigs.SelectedConfig := ""
C_LVconfigs.OnEvent("ItemSelect", (thisLV, Item, Selected) {
	rowI := thisLV.GetNext()
	if rowI && Item != rowI
		return
	C_LVconfigsUpdate(rowI)
})
C_LVconfigsUpdate(rowI := 0) {
	if rowI	{
		Text := C_LVconfigs.GetText(rowI)
		EnabledGroupBoxRule(true)
		ConfigGui_SB.SetText("配置: " Text)
	} else {
		Text := rowI
		EnabledGroupBoxRule(false)
		ConfigGui_SB.SetText("未选择配置")
	}
	;保存前一个配置的规则
	C_LVrule.SaveRule(, C_DDLruleIndex.Value)
	C_DDLruleIndex.lastValue := C_DDLruleIndex.Value := 1
	C_LVprocess.SaveRule()
	SaveAdvancedRules()
	;选择项变化
	C_LVconfigs.SelectedConfig := Text
	;加载当前配置的规则
	C_LVrule.LoadRule()
	C_LVprocess.LoadRule()
	LoadAdvancedRules()
}
;控制文件合并规则控件状态
EnabledGroupBoxRule(s) {
	C_EDnoRepeatFields.Enabled := C_LVprocess.Enabled := C_BTaddKey2.Enabled := C_BTdeleteKey2.Enabled
	:= C_DDLruleIndex.Enabled := C_BTaddRule.Enabled := C_BTdeleteRule.Enabled := C_BTdefaultRule.Enabled 
	:= C_BTclearRule.Enabled := C_LVrule.Enabled := C_BTaddKey.Enabled := C_BTdeleteKey.Enabled := s ? true : false
}
;右键某行弹出菜单
C_LVconfigs.OnEvent("ContextMenu", (thisLV, rowI, IsRightClick, X, Y) {
	;初次调用时创建菜单
	if !IsSet(MyMenu) {
		static MyMenu := Menu()
		MyMenu.rowI := 0
		MyMenu.Add("重命名", MyMenu_Call)
		;MyMenu.SetIcon("1&", "HICON:" hBitMaps[3])    ;word
		MyMenu.Add("复制配置", MyMenu_Call)
		MyMenu.Add("删除配置", MyMenu_Call)
		MyMenu.Add()
		MyMenu.Add("新建配置", MyMenu_Call)
		;MyMenu.SetIcon("4&", "SHELL32.dll", 4) ; 获取文件夹图标

		MyMenu_Call(ItemName, ItemPos, MyMenu) {
			switch ItemPos
			{
			case 1: 
				C_LVconfigs.RenameConfig(MyMenu.rowI)
			case 2:
				C_LVconfigs.CopyConfig(MyMenu.rowI)
			case 3:
				C_LVconfigs.DeleteConfig(MyMenu.rowI)
			case 5:
				C_LVconfigs.AddConfig()
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
C_LVconfigs.Update := (thisLV) {
	thisLV.Delete()
	for name, config in FileUnion.Configs
		thisLV.Add(, name)
}
; 保存LV参数到FileUnion.Configs
C_LVconfigs.SaveRule := (thisLV, name?, i?) {
	/*
	if !FileUnion.Configs.Has(name ?? C_CBconfig.Text)
		return
	rule := FileUnion.Configs[name ?? C_CBconfig.Text][i ?? C_DDLruleIndex.Value]
	rule.length := 0
	Loop thisLV.GetCount()
		rule.push([thisLV.GetText(A_Index,1), thisLV.GetText(A_Index,2), thisLV.GetText(A_Index,3)])
	*/
}
;新建配置
C_LVconfigs.AddConfig := (thisLV) {
	ConfigGui_SB.SetText("")
	ConfigGui.Opt("+OwnDialogs")
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
C_LVconfigs.DeleteConfig := (thisLV, RowNumber?) {
	ConfigGui_SB.SetText("")
	RowNumber := RowNumber ?? thisLV.GetNext()
	if RowNumber = 0
		return
	FileUnion.Configs.Delete(OldText := thisLV.GetText(RowNumber))
	thisLV.Delete(RowNumber)
	L_DDLconfig.Update(FileUnion.Configs.instances)
	C_LVconfigsUpdate(0)
	ConfigGui_SB.SetText("配置已删除: " OldText)
}
;重命名配置
C_LVconfigs.RenameConfig := (thisLV, RowNumber) {
	ConfigGui_SB.SetText("")
	text := thisLV.GetText(RowNumber)
	ConfigGui.Opt("+OwnDialogs")
	IB := InputBox("输入一个新的配置名称:", "配置重命名", "w250 h100", text)
	if IB.Result = "OK" &&  IB.Value != "" && IB.Value != text {
		if FileUnion.Configs.Has(IB.Value)
			return ConfigGui_SB.SetText('重命名失败: 配置"' IB.Value '"已存在')
		if FileUnion.Configs.ReName(text, IB.Value)
			return ConfigGui_SB.SetText('重命名失败: 配置文件"' IB.Value '"重命名失败')
		L_DDLconfig.Update(FileUnion.Configs.instances, IB.Value)
		thisLV.Modify(RowNumber,, IB.Value)
		ConfigGui_SB.SetText("配置重命名为: " IB.Value)
	}
}
;复制配置
C_LVconfigs.CopyConfig := (thisLV, RowNumber) {
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
C_BTAddConfig.OnEvent("Click", (*) {
	C_LVconfigs.AddConfig()
})
;按钮-删除配置
C_BTDeleteConfig := ConfigGui.Add("Button", "x+1 yp wp hp", "删除配置")
C_BTDeleteConfig.OnEvent("Click", (thisCtrl, Info) {
	C_LVconfigs.DeleteConfig()
})







;Group 配置
ConfigGui.SetFont("c0070DE bold", "微软雅黑")
C_TABconfig := ConfigGui.Add("Tab3", "x+10 y10 Section w" ConfigGuiWidth - 180 " h" ConfigGuiHeight - 40, ["文件添加","内容提取","内容处理","高级"])
ConfigGui.SetFont("cDefault norm", "微软雅黑")
C_TABconfig.Value := C_TABconfig.lastValue := 2
/*
C_TABconfig.OnEvent("Change", (*) {
    switch C_TABconfig.lastValue {
		case 1:
			
		case 2:
			C_LVrule.SaveRule()
		case 3:
			C_LVprocess.SaveRule()
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

;文件添加规则
;C_TABconfig.UseTab(1)

;设置忽略文重复件名
C_CBFliesNoRepeat := ConfigGui.Add("CheckBox", "xs+10 ys+30 w210 h25", "忽略完全相同的文件")
C_CBFliesNoRepeat.Value := LOCAL_JSON.Init("C_CBFliesNoRepeat.Value", "1")
C_CBFliesNoRepeat.ToolTip := "勾选后会忽略文件名和修改日期完全相同的文件"
C_CBFliesNoRepeat.OnEvent("Click", (thisCtrl, Info) {
	ConfigGui.Opt("+Disabled")
	;LV_LoadDir(true)
	ConfigGui.Opt("-Disabled")
})

;设置忽略文件大小小于XX的附件
C_CBLimitFileSizeKB := ConfigGui.Add("CheckBox", "xp y+0 w140 h25", "忽略文件尺寸(KB)小于")
C_CBLimitFileSizeKB.Value := LOCAL_JSON.Init("C_CBLimitFileSizeKB.Value", "1")
C_CBLimitFileSizeKB.OnEvent("Click", (thisCtrl, Info) {
	ConfigGui.Opt("+Disabled")
	C_EDLimitFileSizeKB.Enabled := thisCtrl.Value = 1 ? true : false
	;LV_LoadDir(true)
	ConfigGui.Opt("-Disabled")
})
C_EDLimitFileSizeKB := ConfigGui.Add("Edit", "x+0 yp w65 h25 Center Number Limit6")
C_EDLimitFileSizeKB.Value := LOCAL_JSON.Init("C_EDLimitFileSizeKB.Value", "1")
C_EDLimitFileSizeKB.OnEvent("Change", (thisCtrl, Info) {
	ConfigGui.Opt("+Disabled")
	;LV_LoadDir(true)
	ConfigGui.Opt("-Disabled")
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

;内容提取规则
C_TABconfig.UseTab(2)

;设置表格编号
ConfigGui.Add("Text", "xs+10 ys+30 w50 h25 +0x200", "规则编号")
C_DDLruleIndex := ConfigGui.Add("DDL", "x+0 yp w80", [1,2,3,4,5,6,7,8,9,10])
C_DDLruleIndex.Value := C_DDLruleIndex.lastValue := 1
C_DDLruleIndex.OnEvent("Change", (*) {
	C_LVrule.SaveRule(, C_DDLruleIndex.lastValue)
	C_LVrule.LoadRule()
	C_DDLruleIndex.lastValue := C_DDLruleIndex.Value
})
;按钮-新增
C_BTaddRule := ConfigGui.Add("Button", "x+10 yp w40 h25", "新增")
C_BTaddRule.OnEvent("Click", C_BTaddRule_Click)
C_BTaddRule_Click(thisCtrl, Info) {

}
;按钮-删除
C_BTdeleteRule := ConfigGui.Add("Button", "x+0 yp wp hp", "删除")
C_BTdeleteRule.OnEvent("Click", C_BTdeleteRule_Click)
C_BTdeleteRule_Click(thisCtrl, Info) {

}
;按钮-模板
C_BTdefaultRule := ConfigGui.Add("Button", "x+0 yp wp hp", "模板")
C_BTdefaultRule.OnEvent("Click", C_BTdefaultRule_Click)
C_BTdefaultRule_Click(thisCtrl, Info) {
	if !FileUnion.Configs.Has(C_LVconfigs.SelectedConfig)
		return
	FileUnion.Configs[C_LVconfigs.SelectedConfig][C_DDLruleIndex.Value] := FileUnion.Configs.GetDefaultRule()
	C_LVrule.LoadRule()
}
;按钮-清空
C_BTclearRule := ConfigGui.Add("Button", "x+0 yp wp hp", "清空")
C_BTclearRule.OnEvent("Click", C_BTclearRule_Click)
C_BTclearRule_Click(thisCtrl, Info) {
	if !FileUnion.Configs.Has(C_LVconfigs.SelectedConfig)
		return
	FileUnion.Configs[C_LVconfigs.SelectedConfig][C_DDLruleIndex.Value].Length := 0
	C_LVrule.LoadRule()
}



;提取规则LV
C_LVrule := ConfigGui.Add("ListView", "xs+10 y+5 w300 h" ConfigGuiHeight - 150 " Grid +LV0x10000 BackgroundFEFEFE", ["键","值","附加"])
C_LVrule.ModifyCol(1, 90)
C_LVrule.ModifyCol(2, 130)
C_LVrule.ModifyCol(3, 76)
;LV单元格可编辑,编辑后执行函数
LV_InCellEditing(C_LVrule, (thisLV, Row, Col, OldText, NewText) {
	;有变化则进行动作
})
;点选项目触发动作
C_LVrule.OnEvent("Click", (thisLV, rowI) {
	ConfigGui_SB.SetText(rowI ? (thisLV.GetText(rowI,1) "   " thisLV.GetText(rowI,2) "   " thisLV.GetText(rowI,3)) : "")
})
;从FileUnion.Configs加载参数到LV
C_LVrule.LoadRule := (thisLV, name?, i?) {
	thisLV.Delete()
	name := name ?? C_LVconfigs.SelectedConfig
	if !FileUnion.Configs.Has(name)
		return
	for _, arr in FileUnion.Configs[name][i ?? C_DDLruleIndex.Value]
		thisLV.Add(, arr[1], arr[2], arr[3])
}
;保存LV参数到FileUnion.Configs
C_LVrule.SaveRule := (thisLV, name?, i?) {
	name := name ?? C_LVconfigs.SelectedConfig
	if !FileUnion.Configs.Has(name)
		return
	rule := FileUnion.Configs[name][i ?? C_DDLruleIndex.Value]
	rule.length := 0
	Loop thisLV.GetCount()
		rule.push([thisLV.GetText(A_Index,1), thisLV.GetText(A_Index,2), thisLV.GetText(A_Index,3)])
}

;添加按钮
C_BTaddKey := ConfigGui.Add("Button", "xs+10 y+5 w147 h35", "添加参数")
C_BTaddKey.OnEvent("Click", (thisCtrl, Info) {
	RowNumber := C_LVrule.GetNext(0) || C_LVrule.GetCount() + 1 ; 优先插入到选中行下方，否则插入到最后一行
	RowNumber := C_LVrule.Insert(RowNumber,, "key" RowNumber, "value" RowNumber)
	C_LVrule.Focus()
	C_LVrule.Modify(0, "-Select")          ;全部取消选中
	C_LVrule.Modify(RowNumber, "Select")   ;选中
	C_LVrule.Modify(RowNumber, "Vis")      ;可见
})
;删除按钮
C_BTdeleteKey := ConfigGui.Add("Button", "x+5 yp wp hp", "删除参数")
C_BTdeleteKey.OnEvent("Click", (thisCtrl, Info) {
	selectRows := []
	RowNumber := 0  ; 这样使得首次循环从列表的顶部开始搜索.
	Loop {
		RowNumber := C_LVrule.GetNext(RowNumber)  ; 在前一次找到的位置后继续搜索.
		if not RowNumber  ; 上面返回零, 所以选择的行已经都找到了.
			break
		selectRows.Push(RowNumber)
	}
	for _, RowNumber in selectRows.Reverse()
		C_LVrule.Delete(RowNumber)
})



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

;内容处理规则
C_TABconfig.UseTab(3)


;处理规则LV
C_LVprocess := ConfigGui.Add("ListView", "xs+10 ys+30 w300 h" ConfigGuiHeight - 120 " Grid +LV0x10000 BackgroundFEFEFE", ["字段","替换字符","替换为"])
C_LVprocess.ModifyCol(1, 100)
C_LVprocess.ModifyCol(2, 100)
C_LVprocess.ModifyCol(3, 96)
;LV单元格可编辑,编辑后执行函数
LV_InCellEditing(C_LVprocess, (thisLV, Row, Col, OldText, NewText) {
	;有变化则进行动作
})
;点选项目触发动作
C_LVprocess.OnEvent("Click", (thisLV, rowI) {
	ConfigGui_SB.SetText(rowI ? (thisLV.GetText(rowI,1) "   " thisLV.GetText(rowI,2) "   " thisLV.GetText(rowI,3)) : "")
})
;从FileUnion.Configs加载参数到LV
C_LVprocess.LoadRule := (thisLV, name?) {
	thisLV.Delete()
	name := name ?? C_LVconfigs.SelectedConfig
	if !FileUnion.Configs.Has(name)
		return
	for _, arr in FileUnion.Configs[name].process
		thisLV.Add(, arr[1], arr[2], arr[3])
}
;保存LV参数到FileUnion.Configs
C_LVprocess.SaveRule := (thisLV, name?) {
	name := name ?? C_LVconfigs.SelectedConfig
	if !FileUnion.Configs.Has(name)
		return
	process := FileUnion.Configs[name].process
	process.length := 0
	Loop thisLV.GetCount()
		process.push([thisLV.GetText(A_Index,1), thisLV.GetText(A_Index,2), thisLV.GetText(A_Index,3)])
}

;添加按钮
C_BTaddKey2 := ConfigGui.Add("Button", "xs+10 y+5 w147 h35", "添加参数")
C_BTaddKey2.OnEvent("Click", (thisCtrl, Info) {
	RowNumber := C_LVprocess.GetNext(0) || C_LVprocess.GetCount() + 1 ; 优先插入到选中行下方，否则插入到最后一行
	RowNumber := C_LVprocess.Insert(RowNumber,, "key" RowNumber, "value" RowNumber)
	C_LVprocess.Focus()
	C_LVprocess.Modify(0, "-Select")          ;全部取消选中
	C_LVprocess.Modify(RowNumber, "Select")   ;选中
	C_LVprocess.Modify(RowNumber, "Vis")      ;可见
})
;删除按钮
C_BTdeleteKey2 := ConfigGui.Add("Button", "x+5 yp wp hp", "删除参数")
C_BTdeleteKey2.OnEvent("Click", (thisCtrl, Info) {
	selectRows := []
	RowNumber := 0  ; 这样使得首次循环从列表的顶部开始搜索.
	Loop {
		RowNumber := C_LVprocess.GetNext(RowNumber)  ; 在前一次找到的位置后继续搜索.
		if not RowNumber  ; 上面返回零, 所以选择的行已经都找到了.
			break
		selectRows.Push(RowNumber)
	}
	for _, RowNumber in selectRows.Reverse()
		C_LVprocess.Delete(RowNumber)
})



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

;高级规则
C_TABconfig.UseTab(4)

;加载当前配置的高级规则
LoadAdvancedRules(name?) {
	C_EDnoRepeatFields.Value := ""

	name := name ?? C_LVconfigs.SelectedConfig
	if !FileUnion.Configs.Has(name)
		return
	advanceRules := FileUnion.Configs[name].advanceRules

	C_EDnoRepeatFields.Value := advanceRules.Has("noRepeatFields") ? advanceRules["noRepeatFields"] : ""
}
;保存高级规则到当前配置
SaveAdvancedRules(name?) {
	name := name ?? C_LVconfigs.SelectedConfig
	if !FileUnion.Configs.Has(name)
		return
	advanceRules := FileUnion.Configs[name].advanceRules

	advanceRules["noRepeatFields"] := C_EDnoRepeatFields.Value
}

;忽略重复行
ConfigGui.Add("Text", "xs+10 ys+33 w60 h25 +0x200", "忽略重复项")
C_EDnoRepeatFields := ConfigGui.Add("Edit", "x+2 yp w240 hp")
ConfigGui.Tips.SetTip(C_EDnoRepeatFields, "填写用[]括起来字段名,可以是多个的组合")
C_EDnoRepeatFields.OnEvent("Change", (thisCtrl, Info) {
	ConfigGui.Opt("+Disabled")
	;LV_LoadDir(true)
	ConfigGui.Opt("-Disabled")
})


