;安装文件函数
DirInstallTo_LOCAL(targetPath, overwrite := 0)
{
	try
	{
		;创建文件夹
		;安装文件
		if overwrite or !FileExist(targetPath "\报验申请汇总.xlsx")
			FileInstall("D:\Admin\OneDrive\ahk 2.0\9.自编软件\5.FileUnion文件合并\FileUnion\NeedInstall\ScriptDir_FU_Data\报验申请汇总.xlsx", targetPath "\报验申请汇总.xlsx", 1)
	}
	return 1
}