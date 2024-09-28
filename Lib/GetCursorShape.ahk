/**
 * 获取光标特征码
 * @returns {Integer | String} 光标特征码
 */
GetCursorShape() {	; 获取光标特征码 by nnrxin，FeiYue改进
	CursorInfo := Buffer(40, 0)	; 创建 光标信息 结构（申请空间大一点没关系）
	NumPut("int", 16 + A_PtrSize, CursorInfo, 0)	; 写入 结构 的大小cbSize（必须准确）
	DllCall("GetCursorInfo", "Ptr", CursorInfo)	; 获取光标信息填入结构
	bShow := NumGet(CursorInfo, 4, "int")	; 读取光标状态 flags字段
	hCursor := NumGet(CursorInfo, 8, "Ptr")	; 读取光标句柄
	if (!bShow)
		return 0
	IconInfo := Buffer(40, 0)	; 创建 图标信息 结构（申请空间大一点没关系）
	DllCall("GetIconInfo", "Ptr", hCursor, "Ptr", IconInfo)	;获取图标信息填入结构
	hBMMask := NumGet(IconInfo, 8 + A_PtrSize, "Ptr")
	hBMColor := NumGet(IconInfo, 8 + A_PtrSize * 2, "Ptr")
	MaskCode := ColorCode := 0, size := 32 * 32
	lpvMaskBits := Buffer(size // 8, 0)	; 创造 数组-掩码图信息，每个字节含8个掩码位
	DllCall("GetBitmapBits", "Ptr", hBMMask, "int", size // 8, "Ptr", lpvMaskBits)
	Loop size // 8
		MaskCode += NumGet(lpvMaskBits, A_Index - 1, "UChar")
	if (hBMColor) {
		lpvColorBits := Buffer(size * 4, 0)	; 创造 数组-彩色图信息，每个彩色位占4个字节
		DllCall("GetBitmapBits", "Ptr", hBMColor, "int", size * 4, "Ptr", lpvColorBits)
		Loop size	; 只读取彩色位中的绿色分量
			ColorCode += NumGet(lpvColorBits, A_Index * 4 - 3, "UChar")
		DllCall("DeleteObject", "Ptr", hBMColor)	; *清理彩色图
	}
	DllCall("DeleteObject", "Ptr", hBMMask)	; *清理掩码图
	return MaskCode . ColorCode	;输出特征码
}