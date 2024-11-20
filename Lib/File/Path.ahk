/************************************************************************
 * @description 处理文件路径的函数
 * @author nnrxin
 * @date 2024/11/12
 * @version 0.0.0
 ***********************************************************************/

/**
* 从文件名中获取文件所在文件夹路径,路径不存在时启用default
* @param {String}  filePath  文件路径
* @param {boolean} verify    是否验证
* @param {String}  default   默认返回值
* @returns {String}          所在的文件夹路径
*/
Path_Dir(filePath, verify := true, default := "") {
	SplitPath(filePath,, &OutDir)
	return verify && !DirExist(OutDir) ? default : OutDir
}

/**
* 返回含后缀的文件名
* @param {String} path 文件/文件夹的路径
* @returns {String}    含后缀的文件名
*/
Path_FileName(path) {
	SplitPath(Path, &OutFileName)
	return OutFileName
}

/** 
* 返回不含后缀的文件名
* @param {String} path 文件/文件夹的路径
* @returns {String}    不含后缀的文件名
*/
Path_FileNameNoExt(path) {
	SplitPath(path,,,, &OutNameNoExt)
	return OutNameNoExt
}

/** 
* 返回文件后缀
* @param {String} path 文件/文件夹的路径
* @returns {String}    返回文件后缀(类似 txt doc)
*/
Path_Ext(path) {
	SplitPath(path,,, &OutExtension)
	return OutExtension
}

/**
* 返回文件或文件夹的完整路径
* @param {String} path 文件/文件夹的路径
* @returns {String}    返回文件/文件夹的完整路径
*/
Path_Full(path) {
	Loop Files, path, "FD"  ; 包括文件和目录.
		return A_LoopFileFullPath
	return path
}

/**
 * 文件重命名后缀后返回新路径
 * @param {String}  path    文件/文件夹的路径
 * @param {String}  newExt  新后缀
 * @param {boolean} verify  是否验证路径
 * @returns {String}        返回新路径, 如果verify为true且路径不存在则返回原路径
 */
Path_RenameExt(path, newExt, verify := false) {
	SplitPath(path, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt)
	return (!verify || DirExist(path)) ? (OutDir && OutDir "\") . OutNameNoExt . (OutExtension ? "." newExt : "") : path
}

/**
* 文件重命名后返回新路径
* @param {String}  path          文件/文件夹的路径
* @param {String}  newNameNoExt  不带后缀的新名字
* @param {boolean} verify        是否验证路径
* @returns {String}              返回新路径, 如果verify为true且路径不存在则返回原路径
*/
Path_Rename(path, newNameNoExt, verify := false) {
	SplitPath(path, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt)
	return (!verify || DirExist(path)) ? (OutDir && OutDir "\") . newNameNoExt . (OutExtension && "." OutExtension) : path
}

/**
* 文件名后面增加字符
* @param {String} path      文件/文件夹的路径
* @param {String} appendStr 增加的字符
* @param {String} verify    是否验证路径
* @returns {String}         返回新路径
*/
Path_NameAppend(path, appendStr := "", verify := false) {
	SplitPath(path, &OutFileName, &OutDir, &OutExt, &OutNameNoExt)
	if !verify
		return (OutDir && OutDir "\") . OutNameNoExt . appendStr . (OutExt && "." OutExt)
	else if !FileExist(path)
		return ""
	else if (OutExt && DirExist(path))    ;目标为文件夹且最后文件夹名称中带 . 的特殊情况
		return (OutDir && OutDir "\") . OutFileName . appendStr
	else
		return (OutDir && OutDir "\") . OutNameNoExt . appendStr . (OutExt && "." OutExt)
}

/**
* 尝试获取简单的相对路径(不带..),不成功则返回原路径
* @param {String} path      文件/文件夹的路径
* @param {String} rootPath  上级目录
* @returns {String}         返回相对路径
*/
Path_Relative(path, rootPath) {
	return (InStr(path, rootPath) = 1) ? SubStr(path, StrLen(rootPath) + 2) : path
}


/**
* 获取A_Args中包含的路径 (拖放到脚本上的文件)
* @returns {array} 路径数组
* 局限:路径中不能含有连续两个以上空格
*/
Path_InArgs() {
	argPaths := []
	argPath := ""
	subPath := ""
	for i, arg in A_Args {  ; 对每个参数 (或拖放到脚本上的文件) 进行循环:
		if !FileExist(arg) {
			subPath .= subPath ? " " arg : arg
			if FileExist(subPath)
				argPath := subPath
			else
				continue
		} else
			argPath := arg
		subPath := ""
		argPaths.push(argPath)
	}
	return argPaths
}


/**
* 路径合法化(长度限制,非法字符去除)
* @param {String}  str   字符串
* @param {Integer} limit 长度限制
* @returns {String}      路径字符串
*/
Path_Legalize(str, limit := 255) {
	path := ""
	str := RegExReplace(str, '[/\*\?"<>\|`r`n]')   ;去除非法字符/*?"<>|   不包括\:
	SplitPath str, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt
	limit := OutExtension ? limit - StrLen(OutExtension) - 2 : limit - 1    ;过长路径将在末尾添加~
	len := 0
	Loop Parse, OutDir "\" OutNameNoExt {
		len += (Ord(A_LoopField) > 0xFF) ? 2 : 1
		if (len > limit) {
			path .= "~"
			break
		}
		path .= A_LoopField
	}
	return OutExtension ? path "." OutExtension : path
}