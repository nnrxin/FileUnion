/************************************************************************
 * @description 创建一个带进度条等控件的子窗口
 * @author nnrxin
 * @date 2024/09/27
 * @version 0.0.0
 ***********************************************************************/

class ProgressGui {

	;新建
	__New(OwnerGui, title := "执行中...") {
		this.OwnerGui := OwnerGui
		this.title := title
		this.Gui := Gui("-SysMenu +Owner" OwnerGui.Hwnd, title)
		this.Gui.MarginX := this.Gui.MarginY := 15
		this.Gui.OnEvent("Close", (*) {
			this.Gui.Hide()
			this.Close()
		})
		this.Progress := this.Gui.Add("Progress", "xm+3  w700 h22")
		this.Button := this.Gui.Add("Button", "Default x+5 yp-5 w95 h32", "暂停")
		this.Button.OnEvent("Click", (*){
			if this.Button.Text = "暂停" {
				this.Pause()
				this.Button.Text := "继续"
				this.Gui.Opt("+SysMenu -MaximizeBox -MinimizeBox")
			} else {
				this.Resume()
				this.Button.Text := "暂停"
				this.Gui.Opt("-SysMenu")
			}
		})
		this.Text := this.Gui.Add("Text", "xm y+15 w797")
		this.Edit := this.Gui.Add("Edit", "xm y+15 w800 h150 ReadOnly")
	}

	;初始化并显示窗口
	Start(MaxValue := 100) {
		this.Text.Value := ""
		this.Edit.Value := ""
		this.Progress.Value := 0
		this.Progress.SuccValue := 0
		this.Progress.MaxValue := MaxValue
		this.Progress.Opt("Range0-" this.Progress.MaxValue)
		this.Button.Enabled := true
		this.Button.Text := "暂停"
		this.Gui.Opt("-SysMenu")
		this.Gui.Show("Center")
	}

	;一步开始
	StepStart(info) {
		this.Text.Value := this.Progress.Value + 1 " / " this.Progress.MaxValue " " info
		this.currentStepInfo := info
	}

	;一步结束
	StepFinsih(isSuccessed := 1, resultInfo := "成功") {
		this.Progress.Value++
		this.Progress.SuccValue += isSuccessed
		this.Edit.Value := this.Progress.Value "`t" this.currentStepInfo "`t" resultInfo "`n" this.Edit.Value
	}

	successDelay := 1     ; 成功关闭窗口的延迟
	failureDelay := 180   ; 失败的延迟
	;结束任务
	Finsih() {
		allSuccessed := (this.Progress.SuccValue = this.Progress.MaxValue)
		this.Text.Value := "总计: " this.Progress.MaxValue " , 成功: " this.Progress.SuccValue " , 失败: " this.Progress.MaxValue - this.Progress.SuccValue
		this.Button.Text := "结束"
		this.Button.Enabled := false
		this.Gui.Opt("+SysMenu -MinimizeBox -MaximizeBox")
		if (this.Progress.SuccValue = this.Progress.MaxValue)
			SetTimer(GuiHide, -1000 * this.successDelay)
		GuiHide() {
			this.Gui.Hide()
		}
	}

	Close() => 0
	Pause() => 1
	Resume() => 0
}