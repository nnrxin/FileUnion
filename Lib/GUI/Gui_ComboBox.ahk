/*
	Name: Gui_ComboBox.ahk
	Version 0.1 (2024-10-10)
	Created: 2024-10-10
	Author: nnrxin

	Description:
	为原生组合框对象Gui.ComboBox增加了一些有用的方法

    Methods:
    CB.Update(items := [""], textChoosen := "")  => 更新ComboBox控件内容和当前选定
*/

/**
 * function : 更新ComboBox控件内容和当前选定
 * @param items ; 可以为Array或者Map()
 * @param textChoosen ; 选择一项
 * @returns
 */
Gui.ComboBox.Prototype.Update := (this, items := [""], textChoosen?) {
	this.Delete()
	if items is Array {
		for i, item in items {
			this.Add([item])
			if IsSet(textChoosen) && item = textChoosen
				this.Choose(i)
		}
	} else if items is Map {
		for key in items {
			this.Add([key])
			if IsSet(textChoosen) && key = textChoosen
				this.Choose(key)
		}
	}
    if this.Value = 0
        try this.Value := 1
}