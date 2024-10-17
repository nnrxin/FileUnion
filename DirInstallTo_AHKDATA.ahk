;安装文件函数
DirInstallTo_AHKDATA(targetPath, overwrite := 0)
{
	try
	{
		;创建文件夹
		DirCreate(targetPath "\XL\32bit")
		DirCreate(targetPath "\XL\64bit")
		;安装文件
		if overwrite or !FileExist(targetPath "\XL\32bit\libxl.dll")
			FileInstall("D:\Admin\OneDrive\ahk 2.0\9.自编软件\5.FileUnion文件合并\FileUnion\NeedInstall\AppData_AHKDATA\XL\32bit\libxl.dll", targetPath "\XL\32bit\libxl.dll", 1)
		if overwrite or !FileExist(targetPath "\XL\64bit\libxl.dll")
			FileInstall("D:\Admin\OneDrive\ahk 2.0\9.自编软件\5.FileUnion文件合并\FileUnion\NeedInstall\AppData_AHKDATA\XL\64bit\libxl.dll", targetPath "\XL\64bit\libxl.dll", 1)
	}
	return 1
}