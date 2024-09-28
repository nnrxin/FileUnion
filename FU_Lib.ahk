;安装XL\XL库文件
#Include AHK_InstallFiles.ahk
if !AHK_DirInstallTo(AHK_DATA_DIR_PATH := A_AppData "\AHKDATA")    ;非覆盖安装
	MsgBox "XL\XL库文件安装错误!"
DllCall('LoadLibrary', 'str', AHK_DATA_DIR_PATH '\XL\' (A_PtrSize * 8) 'bit\libxl.dll', 'ptr')
#Include <XL\XL>
#Include <File\Path>
#Include <DB\ADODB>



Class FileUnion {

	/**
	 * 获取文件列表
	 */
	static files := []
	static LoadDir(dirPath, force := false, noRepeat := false, limitFileSizeKB := -1) {
		;重复路径优化
		static lastpath := ""
		if !dirPath || dirPath = lastpath && force = false
			return this.files
		lastpath := dirPath
		;相关参数
		static exts := ["xls","xlsx"]
		static ignorefileNames := Map("Thumbs.db",true, "thumbs.db",true) ; 忽略的文件名称
		;开始加载
		this.files.Length := 0
		existsFileName := Map()
		Loop Files, dirPath "\*.*", "FR"{
			if !exts.IndexOf(A_LoopFileExt)
			or InStr(A_LoopFileAttrib, "H") ; 跳过隐藏文件
			or ignorefileNames.Has(A_LoopFileName)
			or noRepeat and existsFileName.Has(A_LoopFileName "|" A_LoopFileTimeModified) ; 文件名相同且修改日期相同
			or A_LoopFileSizeKB < limitFileSizeKB
				continue
			existsFileName[A_LoopFileName "|" A_LoopFileTimeModified] := true
			this.files.Push({path: A_LoopFileFullPath, name: A_LoopFileName})
		}
		return this.files
	}


	/**
	 * 配置相关
	 */
	class Configs {

		; 静态参数
		static WorkingDir := A_ScriptDir
		static Encoding := "UTF-8"
		static defaultSubConfig := [
			["表序号"     , "1"                          , ""      ],
			["表名称"     , ""                           , ""      ],
			["2,1"        , "(日期|date)"                , ""      ],
			["3,4"        , "(项目编号|item No)"         , ""      ],
			["3,7"        , "(检验项目|inspection item)" , ""      ],
			["3,8"        , "(检验员|QC|inscpetor)"      , ""      ],
			["起始行"     , "4"                          , ""      ],
			["非空列"     , "7"                          , ""      ],
			["中止检测列" , "1"                          , "0"     ],
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
		static Has(key) => this.instances.Has(key) ; 判断是否有该实例
		static __Enum(NumberOfVars) => this.instances.__Enum(NumberOfVars) ; 枚举实例
		static __Item[key] {
			get {
				if this.instances.Has(key)
					return this.instances[key]
				throw Error('IMData ' key ' does not exist')
			}
			set {
				this.instances[key] := value
			}
		}
		static Add(name) {
			if !FileUnion.Configs.Has(name)
				FileUnion.Configs[name] := FileUnion.Configs(name)
			return FileUnion.Configs.activeConfig := FileUnion.Configs[name]
		}


		; 构造函数
		__New(name) {
			this.name := name
			this.WorkingDir := FileUnion.Configs.WorkingDir
			this.Encoding := FileUnion.Configs.Encoding
			this.FilePath := this.WorkingDir "\config-" name ".json"
			this.items := [[],[],[],[],[],[],[],[],[],[]] ; 预设10个配置
			this.deepConfigs := []
		}

		; 数组方式调用
		__Item[i] {
			get {
				return this.items[i]
			}
			set {
				this.items[i] := value
			}
		}
		; 枚举实例
		__Enum(NumberOfVars) => this.items.__Enum(NumberOfVars) 

		; 获取模板
		GetDefault() => FileUnion.Configs.defaultSubConfig.Clone()

		;从JSON文件加载参数,失败返回1
		LoadFromFile() {
			try this.items := JSON.parse(FileRead(this.FilePath, this.Encoding))
			catch
				return 1
		}
		;保存LV参数到JSON文件
		SaveToFile() {
			DirCreate Path_Dir(this.FilePath)
			try FileDelete(this.FilePath)
			FileAppend(JSON.stringify(this.items), this.FilePath, this.Encoding)
		}

		;转化成底层配置
		ConvertToDeep() {
			this.deepConfigs.Length := 0
			for _, SubConfig in this.items {
				;跳过空配置
				if !SubConfig.Length
					continue
				deepConfig := {}
				this.deepConfigs.Push(deepConfig)
				;确定各参数
				deepConfig.match := []       ; 匹配信息
				deepConfig.fields := []      ; 字段信息
				for i, arr in SubConfig {
					k := arr[1], v := arr[2], v2 := arr[3]
					if !k && !v
						continue
					else if (k = "表序号") && v && IsDigit(v)
						deepConfig.tableName := Number(v) - 1
					else if (k = "表名称") && v
						deepConfig.tableName := v
					else if (k = "起始行") && v && IsDigit(v)
						deepConfig.startRow := Number(v)
					else if (k = "非空列") && v && IsDigit(v)
						deepConfig.nonemptyColumn := Number(v)
					else if (k = "中止检测列") && v && IsDigit(v) {
						deepConfig.endCheckColumn := Number(v)
						deepConfig.endCheckMaxCount := v2 && IsDigit(v2) ? Number(v2) : 0 ; 默认最大容忍次数为0
					} else if p := RegExMatch(k, "(?<=\d),(?=\d)")    ; 匹配信息
						deepConfig.match.push({row:SubStr(k, 1, p-1), column:SubStr(k, p+1), value:v})
					else if RegExMatch(k, "^\[.+]$") {              ; 字段信息
						fieldName := SubStr(k, 2, -1)
						fields := deepConfig.fields
						fieldsLength := fields.Length
						if p := RegExMatch(v, "(?<=\d),(?=\d)") ; 固定单元格
							fields.Push({name:fieldName, row:SubStr(v, 1, p-1), column:SubStr(v, p+1)})
						else if IsDigit(v) ; 固定列
							fields.Push({name:fieldName, column:v})
						else if RegExMatch(v, "^%.+%$") ; 参数
							fields.Push({name:fieldName, variable:SubStr(v, 2, -1)})
						if v2 != "" && fields.Length > fieldsLength
							fields[fields.Length].FormatStr := v2 ; 添加格式字符串
					}
				}
				;补全必要参数
				deepConfig.tableName := deepConfig.HasProp("tableName") ? deepConfig.tableName : 0
				deepConfig.startRow := deepConfig.HasProp("startRow") ? deepConfig.startRow : 1
				deepConfig.endCheckColumn := deepConfig.HasProp("endCheckColumn") ? deepConfig.endCheckColumn : 1
				deepConfig.endCheckMaxCount := deepConfig.HasProp("endCheckMaxCount") ? deepConfig.endCheckMaxCount : 0
			}
			return this.deepConfigs
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
	 * 合并文件
	 */
	static UnionFiles() {
		this.Data.Clear()
		deepConfigs := this.Configs.activeConfig.ConvertToDeep()
		for i, file in this.files {
			this.LoadExcel(file, deepConfigs)
		}
	}

	/**
	 * 读取Excel文件,识别成功返回匹配的配置序号,失败返回0
	 */
	static LoadExcel(file, deepConfigs) {
		flieName := file.name ; 供参数调用
		book := XL.Load(file.path)
		for cfgI, deepConfig in deepConfigs {
			;尝试获取表
			try sheet := book[deepConfig.tableName]
			catch
				continue
			;匹配信息确认
			passMatch := true
			for _, match in deepConfig.match {
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
			for _, field in deepConfig.fields {
				; 判断是否包含字段名,不包含时新建
				Index := this.Data.FieldIndex(field.name) || this.Data.AddField(field.name)
				; 预处理配置              field.HasProp("FormatStr") 
				if field.HasProp("variable") {
					if !IsSet(v := %(field.variable)%) 
						continue
					try fixedFields[Index] := Format(field.FormatStr, v)
					catch 
						fixedFields[Index] := v
				} else if field.HasProp("row") {
					try fixedFields[Index] := Format(field.FormatStr, sheet[field.row-1,field.column-1].value)
					catch
						fixedFields[Index] := sheet[field.row-1,field.column-1].value
				} else
					loopedFields[Index] := field
			}
			;循环获取各行数据
			rowI := deepConfig.startRow - 1
			endCheckCount := 0
			Loop {
				rowI++
				;检查是否中止
				if sheet[rowI, deepConfig.endCheckColumn-1].value = '' {
					if ++endCheckCount > deepConfig.endCheckMaxCount
						break
				}
				;检查非空列状态
				if deepConfig.HasProp("nonemptyColumn") and sheet[rowI-1, deepConfig.nonemptyColumn-1].value = ""
					continue
				;开始添加一条信息
				row := []
				row.Length := this.Data.ColumnCount()
				for Index, value in fixedFields ; 固定
					row[Index] := value
				for Index, field in loopedFields {  ; 循环
					try row[Index] := Format(field.FormatStr, sheet[rowI-1,field.column-1].value)
					catch
						row[Index] := sheet[rowI-1,field.column-1].value
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
	 * 导出为Excel
	 */
	static ExportToExcel(path) {
		book := XL.New(Path_Extension(path))
		sheet := book.addSheet('Sheet1')
		for R, FieldName in this.Data.FieldNames
			sheet[0, R-1] := FieldName
		for R, row in this.Data {
			for C, value in row
				sheet[R, C-1] := value
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