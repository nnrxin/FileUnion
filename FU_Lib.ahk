;安装XL\XL库文件
#Include DirInstallTo_AHKDATA.ahk
if !DirInstallTo_AHKDATA(AHK_DATA_DIR_PATH := A_AppData "\AHKDATA")    ;非覆盖安装
	MsgBox "XL\XL库文件安装错误!"
DllCall('LoadLibrary', 'str', AHK_DATA_DIR_PATH '\XL\' (A_PtrSize * 8) 'bit\libxl.dll', 'ptr')
#Include <XL\XL>
;其他必要库
#Include <File\Path>
#Include <DB\ADODB>


Class FileUnion {

	; 释放com对象
	static __Delete() {
		if this.HasProp("WordApp")
			this.WordApp.Quit()
	}

	/**
	 * 文件相关
	 */
	class Files {
		static exts := ["xls","xlsx","doc","docx"]           ; 文件扩展名
		static ignorefileNames := ["Thumbs.db", "thumbs.db"] ; 忽略的文件名称

		static items := Map() ; 实例集合
		static Has(path) => this.items.Has(path) ; 判断是否有该实例
		static Count => this.items.Count ; 实例数量
		static Delete(path) => this.items.Delete(path) ; 删除
		static Clear() => this.items.Clear() ; 清空
		static __Enum(NumberOfVars) => this.items.__Enum(NumberOfVars) ; 枚举实例
		static __Item[path] {
			get {
				if this.items.Has(path)
					return this.items[path]
			}
			set {
				this.items[path] := value
			}
		}
		;添加文件, 添加成功时返回文件
		static Add(path) {
			if this.Has(path)
				return
			SplitPath path, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt, &OutDrive
			if !this.exts.IndexOf(OutExtension)   ; 跳过后缀不匹配的文件
			or InStr(FileGetAttrib(path), "H")    ; 跳过隐藏文件
			or FileGetSize(path, "KB") < 1        ; 跳过文件太小的文件
			;or noRepeat and existsFileName.Has(A_LoopFileName "|" A_LoopFileTimeModified) ; 文件名相同且修改日期相同
				return
			return this[path] := this(path, OutFileName, OutExtension)
		}
		;加载一些文件, 返回新增的文件
		static Load(pathArray) {
			files := []
			for _, path in pathArray {
				if DirExist(path) {
					Loop Files, path "\*.*", "FR"
						if file := this.Add(A_LoopFileFullPath)
							files.Push(file)
					continue
				}
				if file := this.Add(path)
					files.Push(file)
			}
			return files
		}
		; 构造函数
		__New(path, name, ext) {
			this.path := path
			this.name := name
			this.ext := ext
			this.type := RegExMatch(ext, "i)^xlsx?") ? "excel" : "word"
		}
	}


	/**
	 * 配置相关
	 */
	class Configs {

		; 静态参数
		static defaultRule := [
			["表序号"     , "1"                          , ""      ],
			["表名称"     , ""                           , ""      ],
			["2,1"        , "(日期|date)"                , ""      ],
			["3,4"        , "(项目编号|item No)"         , ""      ],
			["3,7"        , "(检验项目|inspection item)" , ""      ],
			["3,8"        , "(检验员|QC|inspector)"      , ""      ],
			["起始行"     , "4"                          , ""      ],
			["非空列"     , "7"                          , ""      ],
			["中止检测列" , "1"                          , "5"     ],
			["[文件名]"   , "%flieName%"                 , ""      ],
			["[日期]"     , "2,2"                        , "{:d}"  ],
			["[船号]"     , "2"                          , ""      ],
			["[序号]"     , "1"                          , "{:d}"  ],
			["[项目编号]" , "4"                          , ""      ],
			["[检验项目]" , "7"                          , ""      ],
			["[检验员]"   , "8"                          , ""      ],
		]

		; 实例管理
		static instances := Map() ; 实例集合
		static activeConfig := "" ; 当前活动配置
		static Has(name) => this.instances.Has(name) ; 判断是否有该实例
		static Count => this.instances.Count ; 实例数量
		static __Enum(NumberOfVars) => this.instances.__Enum(NumberOfVars) ; 枚举实例
		static __Item[name] {
			get {
				if this.instances.Has(name)
					return this.instances[name]
				throw Error('Config [' name '] does not exist')
			}
			set {
				this.instances[name] := value
			}
		}
		static Load(JsonMap) { ;从json对象加载配置
			JsonMap := (JsonMap is Map) ? JsonMap : Map()
			for name, configMap in (this.instances := JsonMap) {
				config := this.Add(name, true)
				for k, v in configMap
					config.%k% := v
			}
		}
		static Add(name, force := false) {
			if force || !this.Has(name)
				this[name] := this(name)
			return this.activeConfig := this[name]
		}
		static Clone(name, cloneFromName) {
			if this.Has(name)
				throw Error('Config [' name '] already exist')
			else if !this.Has(cloneFromName)
				throw Error('Config [' cloneFromName '] does not exist')
			this[name] := this(name)
			this[name].rules := this[cloneFromName].rules.Clone()
			this[name].process := this[cloneFromName].process.Clone()
			return this.activeConfig := this[name]
		}
		static Switch(name) {
			if !this.Has(name)
				throw Error('Config [' name '] does not exist')
			return this.activeConfig := this[name]
		}
		static ReName(name, NewName) {
			if this.Has(NewName)
				throw Error('Config [' NewName '] already exist')
			this[name].name := NewName
			this.instances[NewName] := this.instances[name]
			this.instances.Delete(name) ; 待确定
		}
		static Delete(name) {
			if !this.Has(name)
				throw Error('Config [' name '] does not exist')
			if this.activeConfig && this.activeConfig.name = name
				this.activeConfig := ""
			this.instances.Delete(name)
		}
		static Clear() {
			this.activeConfig := ""
			this.instances.Clear()
		}
		; 获取规则模板
		static GetDefaultRule() => this.defaultRule.Clone()


		; 构造函数
		__New(name) {
			this.name := name
			this.rules := [[],[],[],[],[],[],[],[],[],[]] ; 预设10个提取规则
			this.process := [] ; 内容处理规则
		}
		; 删除
		__Delete() {
		}
		; Map方式调用
		__Item[i] {
			get {
				return this.rules[i]
			}
			set {
				this.rules[i] := value
			}
		}

		;转化成底层配置
		ConvertToDeep() {
			deepRules := []
			;配置提取规则
			for _, rule in this.rules {
				;跳过空规则
				if !rule.Length
					continue
				deepRule := {}
				deepRules.Push(deepRule)
				;确定各参数
				deepRule.match := []       ; 匹配信息
				deepRule.fields := []      ; 字段信息
				for i, arr in rule {
					k := arr[1], v := arr[2], v2 := arr[3]
					if !k && !v
						continue
					else if (k = "表序号") && v && IsDigit(v) && v > 0
						deepRule.tableName := Number(v)
					else if (k = "表名称") && v
						deepRule.tableName := v
					else if (k = "起始行") && v && IsDigit(v)
						deepRule.startRow := Number(v)
					else if (k = "非空列") && v && IsDigit(v)
						deepRule.nonemptyColumn := Number(v)
					else if (k = "中止检测列") && v && IsDigit(v) {
						deepRule.endCheckColumn := Number(v)
						deepRule.endCheckMaxCount := v2 && IsDigit(v2) ? Number(v2) : 0 ; 默认最大容忍次数为0
					} else if p := RegExMatch(k, "(?<=\d),(?=\d)")    ; 匹配信息
						deepRule.match.push({row:SubStr(k, 1, p-1), column:SubStr(k, p+1), value:v})
					else if RegExMatch(k, "^\[.+]$") {              ; 字段信息
						fieldName := SubStr(k, 2, -1)
						fields := deepRule.fields
						fieldsLength := fields.Length
						if p := RegExMatch(v, "(?<=\d),(?=\d)") ; 固定单元格
							fields.Push({name:fieldName, row:SubStr(v, 1, p-1), column:SubStr(v, p+1)})
						else if IsDigit(v) ; 固定列
							fields.Push({name:fieldName, column:v})
						else if RegExMatch(v, "^%.+%$") ; 参数
							fields.Push({name:fieldName, variable:SubStr(v, 2, -1)})
						else
							fields.Push({name:fieldName, value:v}) ; 原义字符串
						if v2 != "" && fields.Length > fieldsLength
							fields[fields.Length].FormatStr := v2 ; 添加格式字符串
					}
				}
				;补全必要参数
				deepRule.tableName := deepRule.HasProp("tableName") ? deepRule.tableName : 1
				deepRule.startRow := deepRule.HasProp("startRow") ? deepRule.startRow : 1
				deepRule.endCheckColumn := deepRule.HasProp("endCheckColumn") ? deepRule.endCheckColumn : 1
				deepRule.endCheckMaxCount := deepRule.HasProp("endCheckMaxCount") ? deepRule.endCheckMaxCount : 0
				;添加配置内容处理规则
				FieldIndex := Map()
				for i, field in deepRule.fields {
					FieldIndex[field.name] := i
					field.RegExReplaceOpts := []
					field.GetValue := GetValue ;绑定函数
				}
				for i, arr in this.process {
					FieldName := arr[1], NeedleRegEx := arr[2], Replacement := arr[3]
					if !FieldIndex.Has(FieldName) ; 字段不存在时跳过
						continue
					fields[FieldIndex[FieldName]].RegExReplaceOpts.Push([NeedleRegEx,Replacement])
				}
			}
			return deepRules

			; 内部函数: 数据处理函数
			GetValue(field, value) {
				;尝试格式化
				try value := Format(field.FormatStr, value)
				catch
					value := value
				;正则替换处理
				for _, rule in field.RegExReplaceOpts {
					value := RegExReplace(value, rule[1], rule[2])
				}
				return value
			}
		}
	}


	/**
	 * 数据文件类
	 */
	class Data {
		static __Call(R, C?) {
			if !IsSet(C)
				return this.Rows[R]
			else if IsInteger(C) 
				return this.Rows[R][C]
			else 
				return this.Rows[R][this.FieldIndex(C)]
		}
		static __Enum(NumberOfVars) => this.Rows.__Enum(NumberOfVars) ; 枚举实例
		static Rows := []
		static FieldNames := []
		static FormatStrs := []
		static FieldIndex(FieldName) => this.FieldNames.IndexOf(FieldName)   ; 获取字段序号
		static FieldName(i) => this.FieldNames[i]                            ; 获取字段名称
		static RowCount() => this.Rows.Length                                ; 获取行数
		static ColumnCount() => this.FieldNames.Length                       ; 获取列数

		; 新增数据
		static Add(value*) {
			Row := Array(value*)
			Row.Length := this.FieldNames.Length ; 截短数组长度
			this.Rows.push(Row)
			return Row
		}

		; 清空数据
		static Clear() {
			this.Rows.Length := 0
			this.FieldNames.Length := 0
			this.FormatStrs.Length := 0
		}

		; 增加字段名, 返回序号
		static AddField(name) {
			this.FieldNames.push(name)
			return this.FieldNames.Length
		}
	}

	/**
	 * 合并文件示例
	static UnionFiles() {
		this.Data.Clear()
		deepRules := this.Configs.activeConfig.ConvertToDeep()
		for i, file in this.files {
			if (file.type = "excel")
			    this.LoadExcel(file, deepRules)
			else
				this.LoadWord(file, deepRules)
		}
	}
	*/

	/**
	 * 提取Excel文件内容,识别成功返回匹配的配置序号,失败返回0,文件读取失败返回-1
	 */
	static LoadExcel(file, deepRules) {
		flieName := file.name ; 供参数调用
		book := XL.Load(file.path)
		for cfgI, deepRule in deepRules {
			;尝试获取表
			try sheet := book[IsInteger(deepRule.tableName) ? deepRule.tableName - 1 : deepRule.tableName] ; 数字则-1
			catch
				continue
			;匹配信息确认
			passMatch := true
			for _, match in deepRule.match {
				if !RegExMatch(sheet[match.row-1, match.column-1].value, "i)" match.value) {
					passMatch := false
					break
				}
			}
			if !passMatch
				continue
			;字段信息分类
			fixedFields := Map()  ; 固定
			loopedFields := Map() ; 循环
			for _, field in deepRule.fields {
				; 判断是否包含字段名,不包含时新建
				Index := this.Data.FieldIndex(field.name) || this.Data.AddField(field.name)
				; 预处理配置
				if field.HasProp("value") {
					fixedFields[Index] := field.value
				} else if field.HasProp("variable") {
					if !IsSet(v := %(field.variable)%) 
						continue
					fixedFields[Index] := field.GetValue(v)
				} else if field.HasProp("row") {
					fixedFields[Index] := field.GetValue(sheet[field.row-1,field.column-1].value)
				} else
					loopedFields[Index] := field
			}
			;循环获取各行数据(rowI起始行为1)
			endCheckCount := 0
			Loop sheet.lastRow()+1 - deepRule.startRow + 1 {
				rowI := deepRule.startRow + A_Index - 1
				;检查是否中止
				if sheet[rowI-1, deepRule.endCheckColumn-1].value = '' {
					if ++endCheckCount > deepRule.endCheckMaxCount
						break
				}
				;检查非空列状态
				if deepRule.HasProp("nonemptyColumn") and Trim(sheet[rowI-1, deepRule.nonemptyColumn-1].value, ' `t`r`n') = ""
					continue
				;开始添加一条信息
				row := []
				row.Length := this.Data.ColumnCount()
				for Index, value in fixedFields ; 固定
					row[Index] := value
				for Index, field in loopedFields {  ; 循环
					row[Index] := field.GetValue(sheet[rowI-1,field.column-1].value)
				}
				this.Data.Add(row*)
			}
			;跳过其他配置
			matched := cfgI
			break
		}
	    book := ''
		return matched ?? 0
	}

	/**
	 * 提取Word文件内容,识别成功返回匹配的配置序号,失败返回0
	 * 说明: Range.Text 最后两位分别是ASCII 13和ASCII 7
	 */
	static LoadWord(file, deepRules) {
		if !this.HasProp("WordApp") {
			this.WordApp := ComObject("Word.application")
			this.WordApp.Visible := False  ; 不可见
			this.WordApp.DisplayAlerts := 0 ; 警告和消息的处理的方式(不显示任何警告或消息框)
		}
		flieName := file.name ; 供参数调用
		document := this.WordApp.documents.Open(file.path,, true) ; 只读打开

		for cfgI, deepRule in deepRules {
			;尝试获取表
			try table := document.Tables.Item(IsInteger(deepRule.tableName) ? deepRule.tableName : 1) ; 非整数则转化为数字1
			catch
				continue
			;匹配信息确认
			passMatch := true
			for _, match in deepRule.match {
				if !RegExMatch(table.Cell(match.row, match.column).Range.Text, "i)" match.value) {
					passMatch := false
					break
				}
			}
			if !passMatch
				continue
			;字段信息分类
			fixedFields := Map()  ; 固定
			loopedFields := Map() ; 循环
			for _, field in deepRule.fields {
				; 判断是否包含字段名,不包含时新建
				Index := this.Data.FieldIndex(field.name) || this.Data.AddField(field.name)
				; 预处理配置
				if field.HasProp("value") {
					fixedFields[Index] := field.value
				} else if field.HasProp("variable") {
					if !IsSet(v := %(field.variable)%) 
						continue
					fixedFields[Index] := field.GetValue(v)
				} else if field.HasProp("row") {
					fixedFields[Index] := field.GetValue(TableText(table,field.row,field.column))
				} else
					loopedFields[Index] := field
			}
			;循环获取各行数据(rowI起始行为1)
			endCheckCount := 0
			Loop table.Rows.Count - deepRule.startRow + 1 {
				rowI := deepRule.startRow + A_Index - 1
				;检查是否中止
				if TableText(table,rowI,deepRule.endCheckColumn) = '' {
					if ++endCheckCount > deepRule.endCheckMaxCount
						break
				}
				;检查非空列状态
				if deepRule.HasProp("nonemptyColumn") and Trim(TableText(table,rowI,deepRule.nonemptyColumn), ' `t`r`n') = ""
					continue
				;开始添加一条信息
				row := []
				row.Length := this.Data.ColumnCount()
				for Index, value in fixedFields ; 固定
					row[Index] := value
				for Index, field in loopedFields {  ; 循环
					row[Index] := field.GetValue(TableText(table,rowI,field.column))
				}
				this.Data.Add(row*)
			}
			;跳过其他配置
			matched := cfgI
			break
		}
	    document.Close()
		document := ''
		return matched ?? 0

		;内部函数:处理word格内一些字符: 垂直制表符{11}和回车键{13}转换成换行键{10}, 去除末尾的{13}{7}
		static TableText(table, row, col) {
			return RegExReplace(SubStr(table.Cell(row,col).Range.Text,1,-2), "(" Chr(13) "|" Chr(11) ")" , "`n")
		}
	}



	/**
	 * 导出为Excel
	 */
	static ExportToExcel(path) {
		if FileExist(path) {
			book := XL.Load(path)
			sheet := book[0]
		} else {
			book := XL.New(Path_Extension(path))
			sheet := book.addSheet('Sheet1')
			for R, FieldName in this.Data.FieldNames
				sheet[0, R-1] := FieldName
		}
		for R, row in this.Data {
			for C, value in row {
				sheet[R, C-1] := IsNumber(value) ? Number(value) : value
				if R > 1
					sheet[R, C-1].format := sheet.cellFormat(R-1, C-1)
			}
			if R < this.Data.Rows.Length
				sheet.insertRow(R+1, R+1) ; 向下插入一行
		}
		book.save(path)
		book := ''
	}

	/**
	 * 导出为Access
	 */
	static ExportToAccess(path) {
		; 创建Access数据库
		acApp := ComObject("Access.Application")
		if !FileExist(path)
			acApp.NewCurrentDatabase(path, (Path_Extension(path) = "mdb") ? 10 : 0)
		acApp.Quit()
		acApp := ""
		; 连接到数据库并新建表
		ado := ADODB.Open(path)
		fields := []
		for _, FieldName in this.Data.FieldNames
			fields.Push({name: FieldName, datatype: "TEXT", constraints: ""})
		ado.SQL(SQL_CREATE_TABLE("result", fields))
		;批量插入数据
		rs := ado.Recordset("result", 1, 3)
		; 开始事务
		ado.conn.BeginTrans ; 开始事务

		for R, row in this.Data {
			rs.AddNew
			for C, value in row
				rs.Fields.Item(this.Data.FieldNames[C]).Value := value
			rs.Update
		}

		; 提交事务
		ado.conn.CommitTrans
		
		; 清理
		rs.Close
		ado.conn.Close
		ado.Quit()
		ado := ""
	}
}

;创造一个新表
SQL_CREATE_TABLE(tableName, fields) {
	str := "CREATE TABLE " tableName " ("
	for _, field in fields
		str .= "`n`t" field.name " " field.datatype " " field.constraints ","
	return RTrim(str, " ,") ");"
}








 

 






/*
`CREATE TABLE` 是SQL（Structured Query Language）中用来创建新表的语句。在关系型数据库中，表是存储数据的基础结构，它由行（记录）和列（字段）组成。下面是`CREATE TABLE`语句的基本语法：

```sql
CREATE TABLE table_name (
    column1 datatype constraints,
    column2 datatype constraints,
    column3 datatype constraints,
    ...
);
```

各部分详解如下：

1. **table_name**: 要创建的表的名称。

2. **column1, column2, column3, ...**: 表中的列（字段）名称。

3. **datatype**: 列的数据类型，例如 `INT`、`VARCHAR`、`DATE`、`DECIMAL` 等。

4. **constraints**: 列的约束条件，用于规定列的数据规则，如 `PRIMARY KEY`、`FOREIGN KEY`、`UNIQUE`、`NOT NULL` 等。

### 数据类型

数据类型定义了列中可以存储的数据的种类和格式。常见的数据类型包括：

- **INT**: 整数。
- **VARCHAR(n)**: 变长字符串，`n` 指定最大字符数。
- **TEXT**: 长文本字符串。
- **DATE**: 日期。
- **DECIMAL(m,n)**: 小数，`m` 指定精度（小数点两边的数字总数），`n` 指定小数点后的位数。
- **BOOLEAN**: 布尔值，通常为 TRUE 或 FALSE。

### 约束条件

约束条件用于限制列中的数据，确保数据的准确性和完整性。常见的约束条件包括：

- **PRIMARY KEY**: 主键，唯一标识表中的每一行。一个表可以有一个或多个列作为主键，但每个列的值必须是唯一的。
- **FOREIGN KEY**: 外键，用于建立两个表之间的关系。它指向另一个表的主键或唯一键，并确保引用完整性。
- **UNIQUE**: 唯一约束，确保列中的所有值都是唯一的。
- **NOT NULL**: 非空约束，确保列中的每个记录必须含有值，不能为NULL。
- **CHECK**: 检查约束，确保列中的值符合一定的条件。
- **DEFAULT**: 默认值约束，为列指定默认值。

### 示例

```sql
CREATE TABLE Employees (
    EmployeeID INT PRIMARY KEY,
    FirstName VARCHAR(50) NOT NULL,
    LastName VARCHAR(50) NOT NULL,
    BirthDate DATE,
    Department VARCHAR(50),
    Salary DECIMAL(10, 2) DEFAULT 50000
);
```

在这个例子中，我们创建了一个名为 `Employees` 的表，它有六个列：

- `EmployeeID`: 员工编号，整数类型，设置为主键。
- `FirstName`: 名字，字符串类型，最多50个字符，不能为空。
- `LastName`: 姓氏，字符串类型，最多50个字符，不能为空。
- `BirthDate`: 出生日期，日期类型。
- `Department`: 部门，字符串类型，最多50个字符。
- `Salary`: 薪资，小数类型，总共10位，小数点后2位，默认值为50000。

创建表时，你可以根据自己的需求来定义列的数据类型和约束条件。
*/