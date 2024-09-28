;安装文件函数
AHK_DirInstallTo(targetPath, overwrite := 0)
{
	try
	{
		;创建文件夹
		DirCreate(targetPath "\XL\32bit")
		DirCreate(targetPath "\XL\64bit")
		;安装文件
		if overwrite or !FileExist(targetPath "\XL\32bit\libxl.dll")
			FileInstall("D:\Admin\OneDrive\ahk 2.0\9.自编软件\7.批量操作Excel工具合集\NeedInstall\XL\32bit\libxl.dll", targetPath "\XL\32bit\libxl.dll", 1)
		if overwrite or !FileExist(targetPath "\XL\64bit\libxl.dll")
			FileInstall("D:\Admin\OneDrive\ahk 2.0\9.自编软件\7.批量操作Excel工具合集\NeedInstall\XL\64bit\libxl.dll", targetPath "\XL\64bit\libxl.dll", 1)
	}
	return 1
}