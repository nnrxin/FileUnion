;安装XL\XL库文件
#Include DirInstallTo_AHKDATA.ahk
if !DirInstallTo_AHKDATA(AHK_DATA_DIR_PATH := A_AppData "\AHKDATA")    ;非覆盖安装
	MsgBox "XL\XL库文件安装错误!"
DllCall('LoadLibrary', 'str', AHK_DATA_DIR_PATH '\XL\' (A_PtrSize * 8) 'bit\libxl.dll', 'ptr')
#Include <XL\XL_Lib>
;其他必要库
#Include <File\Path>
#Include <DB\ADODB>


Class FileUnion {

	; 释放com对象
	static __Delete() {
		try {
			if this.HasProp("WordApp")
				this.WordApp.Quit()
		}
	}

	/**
	 * 文件相关
	 */
	class Files {

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
		static Add(path, rules?) {
			if this.Has(path)
				return
			SplitPath path, &OutFileName, &OutDir, &OutExt, &OutNameNoExt, &OutDrive
			FileSize := FileGetSize(path, "KB")
			ModifyTime := FileGetTime(path)
			Date := GetDateYYYYMMDD(OutFileName)
			if IsSet(rules) {
				if rules.exts && !rules.exts.IndexOf(OutExt)                                       ; 跳过后缀不匹配的
				|| rules.IgnoreHidden && InStr(FileGetAttrib(path), "H")                           ; 跳过隐藏
				|| rules.FilePathNeedleRegEx && !RegExMatch(path, "i)" rules.FilePathNeedleRegEx)  ; 跳过路径不匹配的
				|| FileSize < rules.FileSize1 || FileSize > rules.FileSize2                        ; 跳过文件太小(KB)不在范围内的
				|| ModifyTime < rules.ModifyTime1 || ModifyTime > rules.ModifyTime2                ; 跳过修改时间不在范围内的
				|| Date && rules.FileNameDate1 && Date < rules.FileNameDate1                       ; 跳过文件名日期太早的
				|| Date && rules.FileNameDate2 && Date > rules.FileNameDate2                       ; 跳过文件名日期太早的
					return
			}
			return this[path] := this(path, OutFileName, OutExt, FileSize, ModifyTime, Date)
		}
		;加载一些文件, 返回新增的文件
		static Load(pathArray, rules?) {
			files := []
			for _, path in pathArray {
				if DirExist(path) {
					Loop Files, path "\*.*", "FR"
						if file := this.Add(A_LoopFileFullPath, rules?)
							files.Push(file)
					continue
				}
				if file := this.Add(path, rules?)
					files.Push(file)
			}
			return files
		}
		; 构造函数
		__New(path, name, ext, sizeKB, ModifyTime, date) {
			this.path := path
			this.name := name
			this.ext := ext
			this.type := RegExMatch(ext, "i)^xlsx?") ? "excel" : "word"
			this.sizeKB := sizeKB
			this.ModifyTime := ModifyTime
			this.date := date
			this.color := 0xFEFEFE
		}
	}


	/**
	 * 配置相关
	 */
	class Configs {

		; 实例继承属性: 预设文件筛选规则
		static Prototype.defaultFileFilter := [
			["指定文件夹"     , ""                  ],
			["文件扩展名"     , "xls xlsx doc docx" ],
			["文件路径包含"   , ""                  ],
			["忽略隐藏文件"   , 1                   ],
			["文件大小(KB) ≥" , ""                  ],
			["文件大小(KB) ≤" , ""                  ],
			["文件修改时间 ≥" , ""                  ],
			["文件修改时间 ≤" , ""                  ],
			["文件名日期 ≥"   , ""                  ],
			["文件名日期 ≤"   , ""                  ],
		]

		; 实例继承属性: 预设文件提取通用规则
		static Prototype.defaultExtractCommonRule := [
			["[文件名]"     , "%flieName%"               , ""      ],
			["[日期]"       , ""                         , "{:d}"  ],
			["[船号]"       , ""                         , ""      ],
			["[序号]"       , ""                         , "{:d}"  ],
			["[项目编号]"   , ""                         , ""      ],
			["[检验项目]"   , ""                         , ""      ],
			["[检验员]"     , ""                         , ""      ],
		]

		; 实例继承属性: 预设文件提取特殊规则
		static Prototype.defaultExtractSpecRule := [
			["表序号"     , ""                           , ""      ],
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
		static active := "" ; 当前活动配置
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
		; 从json对象加载配置
		static Load(JsonMap) { 
			JsonMap := (JsonMap is Map) ? JsonMap : Map()
			for name, configMap in (this.instances := JsonMap) {
				config := this.Add(name, true)
				for k, v in configMap
					config.%k% := v
			}
		}
		; 新增配置
		static Add(name, force := false) {
			if force || !this.Has(name)
				this[name] := this(name)
			return this.active := this[name]
		}
		; 复制配置
		static Clone(name, cloneFromName) {
			if this.Has(name)
				throw Error('Config [' name '] already exist')
			else if !this.Has(cloneFromName)
				throw Error('Config [' cloneFromName '] does not exist')
			this[name] := this(name)
			this[name].FileFilter := this[cloneFromName].FileFilter.Clone()
			this[name].Extract := this[cloneFromName].Extract.Clone()
			this[name].Transform := this[cloneFromName].Transform.Clone()
			this[name].advanceRules := this[cloneFromName].advanceRules.Clone()
			return this.active := this[name]
		}
		; 切换配置
		static Switch(name) {
			if !this.Has(name)
				throw Error('Config [' name '] does not exist')
			return this.active := this[name]
		}
		; 重命名配置
		static ReName(name, NewName) {
			if this.Has(NewName)
				throw Error('Config [' NewName '] already exist')
			this[name].name := NewName
			this.instances[NewName] := this.instances[name]
			this.instances.Delete(name) ; 待确定
		}
		; 删除配置
		static Delete(name) {
			if !this.Has(name)
				throw Error('Config [' name '] does not exist')
			if this.active && this.active.name = name
				this.active := ""
			this.instances.Delete(name)
		}
		; 清空所有配置
		static Clear() {
			this.active := ""
			this.instances.Clear()
		}


		; 构造函数
		__New(name) {
			this.name := name
			this.FileFilter := this.defaultFileFilter.Clone() ; 文件筛选规则
			this.Extract := [this.defaultExtractCommonRule.Clone(), this.defaultExtractSpecRule.Clone()] ; 文件提取规则(1个通用规则,一个特殊规则)
			this.Transform := [] ; 内容转化规则
			this.advanceRules := Map(
				"noRepeatFields", "" ; 不允许重复的字段
			) ; 高级规则
		}
		; 删除
		__Delete() {
		}

		;重置FileFilter
		ResetFileFilter() => this.FileFilter := this.defaultFileFilter.Clone()

		;重置Extract的第i个规则
		ResetExtractRule(i) => this.Extract[i] := (i = 1) ? this.defaultExtractCommonRule.Clone() : this.defaultExtractSpecRule.Clone()

		;新增提取规则
		AddExtractRule(asDefault := true) => this.Extract.Push(asDefault ? this.defaultExtractSpecRule.Clone() : [])

		;复制提取规则, 成功返回1
		CopyExtractRule(i) {
			if i >= 2 && i <= this.Extract.Length {
				this.Extract.Push(this.Extract[i].Clone())
				return 1
			}
		}

		;移除提取规则, 成功返回被移除项
		RemoveExtractRule(i) {
			if i >= 2 && i <= this.Extract.Length && (i >= 3 || this.Extract.Length >= 3)
				return this.Extract.RemoveAt(i)
		}

		;移除提取规则, 成功返回被移除项
		ClearExtractRule(i) {
			if i >= 2 && i <= this.Extract.Length {
				this.Extract[i].Length = 0
				return 1
			}
		}

		;获取文件筛选规则
		GetFileFilterRules() {
			rules := {}
			mp := Map()
			for _, v in this.FileFilter
				mp[v[1]] := v[2]
			rules.path := mp["指定文件夹"] ?? ""
			rules.exts := []
			for _, v in StrSplit(mp["文件扩展名"], [A_Space, A_Tab, ",", ";", "/", "\", "，", "、"], " `t")
				if Trim(v) != ""
					rules.exts.Push(v)
			rules.FilePathNeedleRegEx := mp["文件路径包含"] ?? ""
			rules.IgnoreHidden := mp["忽略隐藏文件"] ?? 1
			rules.FileSize1 := mp["文件大小(KB) ≥"] ?? ""
			rules.FileSize2 := mp["文件大小(KB) ≤"] ?? ""
			rules.ModifyTime1 := mp["文件修改时间 ≥"] ?? ""
			rules.ModifyTime2 := mp["文件修改时间 ≤"] ?? ""
			rules.FileNameDate1 := mp["文件名日期 ≥"] ?? ""
			rules.FileNameDate2 := mp["文件名日期 ≤"] ?? ""
			rules.FileSize1 := (IsDigit(rules.FileSize1) && rules.FileSize1 != "") ? Number(rules.FileSize1) : 0
			rules.FileSize2 := (IsDigit(rules.FileSize2) && rules.FileSize2 != "") ? Number(rules.FileSize2) : 999999999
			rules.ModifyTime1 := (IsDigit(rules.ModifyTime1) && rules.ModifyTime1 != "") ? FormatTime(rules.ModifyTime1, "yyyyMMddHHmmss") : "00000000000000"
			rules.ModifyTime2 := (IsDigit(rules.ModifyTime2) && rules.ModifyTime2 != "") ? FormatTime(rules.ModifyTime2, "yyyyMMddHHmmss") : "99999999999999"
			rules.FileNameDate1 := (rules.FileNameDate1 != "") ? FormatTime(rules.FileNameDate1, "yyyyMMdd") : ""
			rules.FileNameDate2 := (rules.FileNameDate2 != "") ? FormatTime(rules.FileNameDate2, "yyyyMMdd") : ""
			return rules
		}

		;转化成底层配置
		GetDeepRule() {
			deepRules := []
			;通用提取规则
			CommonFields := []
			CommonFieldsIndex := Map()
			for _, arr in this.Extract[1] {
				k := arr[1], v := arr[2], v2 := arr[3]
				if !k && !v
					continue
				else if RegExMatch(k, "^\[(.+)]$", &m) {
					; 字段信息
					fieldName := m[1]
					if RegExMatch(v, "(\d+),(\d+)", &m) ; 固定单元格
						CommonFields.Push({name: fieldName, R: m[1], C: m[2]})
					else if IsDigit(v) ; 固定列
						CommonFields.Push({name: fieldName, C: v})
					else if RegExMatch(v, "^%(.+)%$", &m) ; 参数
						CommonFields.Push({name: fieldName, variable: m[1]})
					else if v != ""
						CommonFields.Push({name: fieldName, value: v}) ; 原义字符串
					else
						CommonFields.Push({name: fieldName})
					; 字段添加后
					CommonFieldsIndex[fieldName] := CommonFields.Length ; 字段的序号
					CommonField := CommonFields[CommonFields.Length]
					if v2 != ""
						CommonField.FormatStr := v2    ; 添加格式字符串
					CommonField.GetValue := GetValue   ; 绑定函数
					CommonField.RegExReplaceOpts := [] ; 正则替换选项
				}
			}
			;添加配置内容转化规则
			for _, arr in this.Transform {
				FieldName := arr[1], NeedleRegEx := arr[2], Replacement := arr[3]
				if !CommonFieldsIndex.has(fieldName) ; 不在通用字段中时跳过
					continue
				CommonField := CommonFields[CommonFieldsIndex[fieldName]]
				CommonField.RegExReplaceOpts.Push([NeedleRegEx,Replacement])
			}
			;特殊提取规则
			for i, rule in this.Extract {
				;跳过通用规则和空规则
				if i = 1 || !rule.Length
					continue
				;确定各参数
				deepRule := {
					matchs : [],                     ; 匹配信息
					fields : DeepClone(CommonFields) ; 字段信息克隆自通用规则
				}
				deepRules.Push(deepRule)
				for _, arr in rule {
					k := arr[1], v := arr[2], v2 := arr[3]
					if !k && !v
						continue
					else if (k = "表序号") && v && IsDigit(v) && v > 0 {
						deepRule.tableIndex := Number(v)
						deepRule.inculdeHidenTable := v2 ; 包含隐藏表
					} else if (k = "表名称") && v
						deepRule.tableName := v
					else if (k = "起始行") && v && IsDigit(v)
						deepRule.startRow := Number(v)
					else if (k = "非空列") && v && IsDigit(v)
						deepRule.nonemptyColumn := Number(v)
					else if (k = "中止检测列") && v && IsDigit(v) {
						deepRule.endCheckColumn := Number(v)
						deepRule.endCheckMaxCount := v2 && IsDigit(v2) ? Number(v2) : 0 ; 默认最大容忍次数为0
					} else if RegExMatch(k, "^(\d+),(\d+)$", &m)    ; 匹配信息
						deepRule.matchs.push({R: m[1], C: m[2], value: v})
					else if RegExMatch(k, "^\[(.+)]$", &m) && CommonFieldsIndex.has(m[1]) {
						field := deepRule.fields[CommonFieldsIndex[m[1]]]
						; v
						if RegExMatch(v, "(\d+),(\d+)", &m) ; 固定单元格
							field.R := m[1], field.C := m[2]
						else if IsDigit(v) ; 固定列
							field.C := v
						else if RegExMatch(v, "^%(.+)%$", &m) ; 参数
							field.variable := m[1]
						else ; 原义字符串
							field.value := v
						; v2
						if v2 != ""
							field.FormatStr := v2 ; 添加格式字符串
					}
				}
				;补全必要参数
				deepRule.tableIndex := deepRule.tableIndex ?? ""
				deepRule.inculdeHidenTable := deepRule.inculdeHidenTable ?? ""
				deepRule.startRow := deepRule.startRow ?? 1
				deepRule.endCheckColumn := deepRule.endCheckColumn ?? 1
				deepRule.endCheckMaxCount := deepRule.endCheckMaxCount ?? 0
			}
			/* 检测
			str := ""
			for i, deepRule in deepRules {
				str .= "`n`n" i ":`n "
				str .= "tableIndex: " deepRule.tableIndex "`tincludeHidenTable: " deepRule.inculdeHidenTable "`tstartRow: " deepRule.startRow "`tendCheckColumn: " deepRule.endCheckColumn "`tendCheckMaxCount: " deepRule.endCheckMaxCount "`n"
				for j, field in deepRule.fields {
					str .= "name: " field.name "`tR: " (field.R ?? " ") "`tC: " (field.C ?? " ") "`tvar: " (field.variable ?? " ") "`tv: " (field.value ?? " ") "`tF: " (field.FormatStr ?? " ") "`n"
				}
				for j, match in deepRule.matchs {
					str .= "R: " (match.R ?? " ") "`tC: " (match.C ?? " ") "`tvalue: " (match.value ?? " ") "`n"
				}
			}
			A_Clipboard := str
			*/
			return deepRules

			; 内部函数: 数据转化函数
			GetValue(field, value) {
				;尝试格式化
				if field.HasProp("FormatStr") {
					try value := Format(field.FormatStr, value)
					catch
						value := value
				}
				;正则替换转化
				for _, rule in field.RegExReplaceOpts
					value := RegExReplace(value, rule[1], rule[2])
				return value
			}
		}

		; 获取无重复字段数组
		GetNoRepeatFields() {
			noRepeatFields := []
			for _, v in RegExMatchAll(this.advanceRules["noRepeatFields"], "U)(?<=\[).*(?=\])")
				noRepeatFields.Push(v[0])
			return noRepeatFields
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
		static RowCount => this.Rows.Length                                  ; 获取行数
		static ColumnCount => this.FieldNames.Length                         ; 获取列数

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

		; 删除重复行,输入为字段名数组,返回删除的行数量
		static DeleteRepeatRow(noRepeatFields) {
			;确定不重复的字段序号
			noRepeatFieldIndexs := []
			for i, fieldName in this.FieldNames
				if noRepeatFields.indexOf(fieldName)
				    noRepeatFieldIndexs.Push(i)
			if noRepeatFieldIndexs.Length = 0
				return 0
			;开始倒着删除重复信息
			exists := Map()
			Loop CountBefore := i := this.RowCount {
				str := ""
				for _, index in noRepeatFieldIndexs
					str .= this.Rows[i][index] "@@"
				if exists.Has(str)
				    this.Rows.RemoveAt(i)
				else
					exists[str] := true
				i--
			}
			return CountBefore - this.RowCount
		}
	}

	/**
	 * 合并文件示例
	static UnionFiles() {
		this.Data.Clear()
		deepRules := this.Configs.active.GetDeepRule()
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
		;遍历全部表格,获取表名/隐藏表等信息
		hidenSheetIndexs := []
		sheetNames := []
		loop book.sheetCount() {
			sheet := book[A_Index]
			sheetNames.Push(sheet.name())
			if sheet.hidden() 
				hidenSheetIndexs.Push(A_Index)
		}
		;按规则获取文件内容
		for cfgI, deepRule in deepRules {
			;尝试获取表,优先使用姓名,序号为空或0时使用当前活动表,最后使用序号时会跳过
			if deepRule.HasProp("tableName") && sheetNames.IndexOf(deepRule.tableName)
				sheetName := deepRule.tableName
			else if !deepRule.tableIndex
				sheetName := book.activeSheet()
			else {
				sheetName := deepRule.tableIndex
				;跳过隐藏表
				if !deepRule.inculdeHidenTable {
					for hidenSheetIndex in hidenSheetIndexs
						if hidenSheetIndex <= sheetName
							sheetName++
						else
							break 
				}
			}
			try sheet := book[sheetName]
			catch
				continue
			;匹配信息确认
			passMatch := true
			for _, match in deepRule.matchs {
				if !RegExMatch(sheet[match.R, match.C].value, "i)" match.value) {
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
				} else if field.HasProp("R") {
					fixedFields[Index] := field.GetValue(sheet[field.R, field.C].value)
				} else
					loopedFields[Index] := field
			}
			;循环获取各行数据(rowI起始行为1)
			endCheckCount := 0
			Loop sheet.lastFilledRow() - deepRule.startRow + 1 {
				R := deepRule.startRow + A_Index - 1
				;检查是否中止
				if sheet[R, deepRule.endCheckColumn].value = '' {
					if ++endCheckCount > deepRule.endCheckMaxCount
						break
				}
				;检查非空列状态
				if deepRule.HasProp("nonemptyColumn") and Trim(sheet[R, deepRule.nonemptyColumn].value, ' `t`r`n') = ""
					continue
				;开始添加一条信息
				row := []
				row.Length := this.Data.ColumnCount
				for Index, value in fixedFields ; 固定
					row[Index] := value
				for Index, field in loopedFields {  ; 循环
					row[Index] := field.GetValue(sheet[R,field.C].value)
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
			try table := document.Tables.Item(deepRule.tableIndex)
			catch
				continue
			;匹配信息确认
			passMatch := true
			for _, match in deepRule.matchs {
				if !RegExMatch(table.Cell(match.R, match.C).Range.Text, "i)" match.value) {
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
				} else if field.HasProp("R") {
					fixedFields[Index] := field.GetValue(TableText(table,field.R,field.C))
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
				row.Length := this.Data.ColumnCount
				for Index, value in fixedFields ; 固定
					row[Index] := value
				for Index, field in loopedFields {  ; 循环
					row[Index] := field.GetValue(TableText(table,rowI,field.C))
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
		static TableText(table, R, col) {
			return RegExReplace(SubStr(table.Cell(R,col).Range.Text,1,-2), "(" Chr(13) "|" Chr(11) ")" , "`n")
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