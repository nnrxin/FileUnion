/************************************************************************
 * @file: XL_Lib.ahk
 * @description: High performance library for reading and writing Excel(xls,xlsx) files.
 * @author thqby
 * @date 2023/07/09
 * @version 1.1.3 (libxl 4.2.0)
 * @documentation https://www.libxl.com/documentation.html
 * 
 * @modified to one-based index 
 * @modifiedBy nnrxin
 * @modifiedVersion 1.0.1
 * @modifiedDate 2024/11/11
 * 
 * @enum var
 * Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
 * NumFormat {GENERAL = 0, NUMBER = 1, NUMBER_D2 = 2, NUMBER_SEP = 3, NUMBER_SEP_D2 = 4, CURRENCY_NEGBRA = 5, CURRENCY_NEGBRARED = 6, CURRENCY_D2_NEGBRA = 7, CURRENCY_D2_NEGBRARED = 8, PERCENT = 9, PERCENT_D2 = 10, SCIENTIFIC_D2 = 11, FRACTION_ONEDIG = 12, FRACTION_TWODIG = 13, DATE = 14, CUSTOM_D_MON_YY = 15, CUSTOM_D_MON = 16, CUSTOM_MON_YY = 17, CUSTOM_HMM_AM = 18, CUSTOM_HMMSS_AM = 19, CUSTOM_HMM = 20, CUSTOM_HMMSS = 21, CUSTOM_MDYYYY_HMM = 22, NUMBER_SEP_NEGBRA=37 = 23, NUMBER_SEP_NEGBRARED = 24, NUMBER_D2_SEP_NEGBRA = 25, NUMBER_D2_SEP_NEGBRARED = 26, ACCOUNT = 27, ACCOUNTCUR = 28, ACCOUNT_D2 = 29, ACCOUNT_D2_CUR = 30, CUSTOM_MMSS = 31, CUSTOM_H0MMSS = 32, CUSTOM_MMSS0 = 33, CUSTOM_000P0E_PLUS0 = 34, TEXT = 35}
 * AlignH {GENERAL = 0, LEFT = 1, CENTER = 2, RIGHT = 3, FILL = 4, JUSTIFY = 5, MERGE = 6, DISTRIBUTED = 7}
 * AlignV {TOP = 0, CENTER = 1, BOTTOM = 2, JUSTIFY = 2, DISTRIBUTED = 3}
 * BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
 * BorderDiagonal {NONE = 0, DOWN = 1, UP = 2, BOTH = 3}
 * FillPattern {NONE, SOLID, GRAY50, GRAY75, GRAY25, HORSTRIPE, VERSTRIPE, REVDIAGSTRIPE, DIAGSTRIPE, DIAGCROSSHATCH, THICKDIAGCROSSHATCH, THINHORSTRIPE, THINVERSTRIPE, THINREVDIAGSTRIPE, THINDIAGSTRIPE, THINHORCROSSHATCH, THINDIAGCROSSHATCH, GRAY12P5, GRAY6P25}
 * Script {NORMAL = 0, SUPER = 1, SUB = 2}
 * Underline {NONE = 0, SINGLE, DOUBLE, SINGLEACC = 0x21, DOUBLEACC = 0x22}
 * Paper {DEFAULT, LETTER, LETTERSMALL, TABLOID, LEDGER, LEGAL, STATEMENT, EXECUTIVE, A3, A4, A4SMALL, A5, B4, B5, FOLIO, QUATRO, 10x14, 10x17, NOTE, ENVELOPE_9, ENVELOPE_10, ENVELOPE_11, ENVELOPE_12, ENVELOPE_14, C_SIZE, D_SIZE, E_SIZE, ENVELOPE_DL, ENVELOPE_C5, ENVELOPE_C3, ENVELOPE_C4, ENVELOPE_C6, ENVELOPE_C65, ENVELOPE_B4, ENVELOPE_B5, ENVELOPE_B6, ENVELOPE, ENVELOPE_MONARCH, US_ENVELOPE, FANFOLD, GERMAN_STD_FANFOLD, GERMAN_LEGAL_FANFOLD, B4_ISO, JAPANESE_POSTCARD, 9x11, 10x11, 15x11, ENVELOPE_INVITE, US_LETTER_EXTRA = 50, US_LEGAL_EXTRA, US_TABLOID_EXTRA, A4_EXTRA, LETTER_TRANSVERSE, A4_TRANSVERSE, LETTER_EXTRA_TRANSVERSE, SUPERA, SUPERB, US_LETTER_PLUS, A4_PLUS, A5_TRANSVERSE, B5_TRANSVERSE, A3_EXTRA, A5_EXTRA, B5_EXTRA, A2, A3_TRANSVERSE, A3_EXTRA_TRANSVERSE, JAPANESE_DOUBLE_POSTCARD, A6, JAPANESE_ENVELOPE_KAKU2, JAPANESE_ENVELOPE_KAKU3, JAPANESE_ENVELOPE_CHOU3, JAPANESE_ENVELOPE_CHOU4, LETTER_ROTATED, A3_ROTATED, A4_ROTATED, A5_ROTATED, B4_ROTATED, B5_ROTATED, JAPANESE_POSTCARD_ROTATED, DOUBLE_JAPANESE_POSTCARD_ROTATED, A6_ROTATED, JAPANESE_ENVELOPE_KAKU2_ROTATED, JAPANESE_ENVELOPE_KAKU3_ROTATED, JAPANESE_ENVELOPE_CHOU3_ROTATED, JAPANESE_ENVELOPE_CHOU4_ROTATED, B6, B6_ROTATED, 12x11, JAPANESE_ENVELOPE_YOU4, JAPANESE_ENVELOPE_YOU4_ROTATED, PRC16K, PRC32K, PRC32K_BIG, PRC_ENVELOPE1, PRC_ENVELOPE2, PRC_ENVELOPE3, PRC_ENVELOPE4, PRC_ENVELOPE5, PRC_ENVELOPE6, PRC_ENVELOPE7, PRC_ENVELOPE8, PRC_ENVELOPE9, PRC_ENVELOPE10, PRC16K_ROTATED, PRC32K_ROTATED, PRC32KBIG_ROTATED, PRC_ENVELOPE1_ROTATED, PRC_ENVELOPE2_ROTATED, PRC_ENVELOPE3_ROTATED, PRC_ENVELOPE4_ROTATED, PRC_ENVELOPE5_ROTATED, PRC_ENVELOPE6_ROTATED, PRC_ENVELOPE7_ROTATED, PRC_ENVELOPE8_ROTATED, PRC_ENVELOPE9_ROTATED, PRC_ENVELOPE10_ROTATED}
 * SheetType {SHEET, CHART, UNKNOWN}
 * CellType {EMPTY, NUMBER, STRING, BOOLEAN, BLANK, ERROR, STRICTDATE}
 * ErrorType {NULL = 0x0, DIV_0 = 0x7, VALUE = 0x0F, REF = 0x17, NAME = 0x1D, NUM = 0x24, NA = 0x2A, NOERROR = 0xFF}
 * PictureType {PNG, JPEG, GIF, WMF, DIB, EMF, PICT, TIFF, ERROR = 0xFF}
 * SheetState {VISIBLE, HIDDEN, VERYHIDDEN}
 * Scope {UNDEFINED = -2, WORKBOOK = -1}
 * Position {MOVE_AND_SIZE, ONLY_MOVE, ABSOLUTE}
 * Operator {EQUAL, GREATER_THAN, GREATER_THAN_OR_EQUAL, LESS_THAN, LESS_THAN_OR_EQUAL, NOT_EQUAL}
 * Filter {VALUE, TOP10, CUSTOM, DYNAMIC, COLOR, ICON, EXT, NOT_SET}
 * IgnoredError {NO_ERROR = 0, EVAL_ERROR = 1, EMPTY_CELLREF = 2, NUMBER_STORED_AS_TEXT = 4, INCONSIST_RANGE = 8, INCONSIST_FMLA = 16, TWODIG_TEXTYEAR = 32, UNLOCK_FMLA = 64, DATA_VALIDATION = 128}
 * EnhancedProtection {DEFAULT = -1, ALL = 0, OBJECTS = 1, SCENARIOS = 2, FORMAT_CELLS = 4, FORMAT_COLUMNS = 8, FORMAT_ROWS = 16, INSERT_COLUMNS = 32, INSERT_ROWS = 64, INSERT_HYPERLINKS = 128, DELETE_COLUMNS = 256, DELETE_ROWS = 512, SEL_LOCKED_CELLS = 1024, SORT = 2048, AUTOFILTER = 4096, PIVOTTABLES = 8192, SEL_UNLOCKED_CELLS = 16384}
 * DataValidationType {TYPE_NONE, TYPE_WHOLE, TYPE_DECIMAL, TYPE_LIST, TYPE_DATE, TYPE_TIME, TYPE_TEXTLENGTH, TYPE_CUSTOM}
 * DataValidationOperator {OP_BETWEEN, OP_NOTBETWEEN, OP_EQUAL, OP_NOTEQUAL, OP_LESSTHAN, OP_LESSTHANOREQUAL, OP_GREATERTHAN, OP_GREATERTHANOREQUAL}
 * DataValidationErrorStyle {ERRSTYLE_STOP, ERRSTYLE_WARNING, ERRSTYLE_INFORMATION}
 * CalcModeType {MANUAL, AUTO, AUTONOTABLE}
 * CheckedType {CHECKEDTYPE_UNCHECKED, CHECKEDTYPE_CHECKED, CHECKEDTYPE_MIXED}
 * ObjectType {OBJECT_UNKNOWN, OBJECT_BUTTON, OBJECT_CHECKBOX, OBJECT_DROP, OBJECT_GBOX, OBJECT_LABEL, OBJECT_LIST, OBJECT_RADIO, OBJECT_SCROLL, OBJECT_SPIN, OBJECT_EDITBOX, OBJECT_DIALOG}
 * CFormatType {CFORMAT_BEGINWITH, CFORMAT_CONTAINSBLANKS, CFORMAT_CONTAINSERRORS, CFORMAT_CONTAINSTEXT, CFORMAT_DUPLICATEVALUES, CFORMAT_ENDSWITH, CFORMAT_EXPRESSION, CFORMAT_NOTCONTAINSBLANKS, CFORMAT_NOTCONTAINSERRORS, CFORMAT_NOTCONTAINSTEXT, CFORMAT_UNIQUEVALUES}
 * CFormatOperator {CFOPERATOR_LESSTHAN, CFOPERATOR_LESSTHANOREQUAL, CFOPERATOR_EQUAL, CFOPERATOR_NOTEQUAL, CFOPERATOR_GREATERTHANOREQUAL, CFOPERATOR_GREATERTHAN, CFOPERATOR_BETWEEN, CFOPERATOR_NOTBETWEEN, CFOPERATOR_CONTAINSTEXT, CFOPERATOR_NOTCONTAINS, CFOPERATOR_BEGINSWITH, CFOPERATOR_ENDSWITH}
 * CFormatTimePeriod {CFTP_LAST7DAYS, CFTP_LASTMONTH, CFTP_LASTWEEK, CFTP_NEXTMONTH, CFTP_NEXTWEEK, CFTP_THISMONTH, CFTP_THISWEEK, CFTP_TODAY, CFTP_TOMORROW, CFTP_YESTERDAY}
 * CFVOType {CFVO_MIN, CFVO_MAX, CFVO_FORMULA, CFVO_NUMBER, CFVO_PERCENT, CFVO_PERCENTILE}
 * CellStyle {CELLSTYLE_NORMAL, CELLSTYLE_BAD, CELLSTYLE_GOOD, CELLSTYLE_NEUTRAL, CELLSTYLE_CALC, CELLSTYLE_CHECKCELL, CELLSTYLE_EXPLANATORY, CELLSTYLE_INPUT, CELLSTYLE_OUTPUT, CELLSTYLE_HYPERLINK, CELLSTYLE_LINKEDCELL, CELLSTYLE_NOTE, CELLSTYLE_WARNING, CELLSTYLE_TITLE, CELLSTYLE_HEADING1, CELLSTYLE_HEADING2, CELLSTYLE_HEADING3, CELLSTYLE_HEADING4, CELLSTYLE_TOTAL, CELLSTYLE_20ACCENT1, CELLSTYLE_40ACCENT1, CELLSTYLE_60ACCENT1, CELLSTYLE_ACCENT1, CELLSTYLE_20ACCENT2, CELLSTYLE_40ACCENT2, CELLSTYLE_60ACCENT2, CELLSTYLE_ACCENT2, CELLSTYLE_20ACCENT3, CELLSTYLE_40ACCENT3, CELLSTYLE_60ACCENT3, CELLSTYLE_ACCENT3, CELLSTYLE_20ACCENT4, CELLSTYLE_40ACCENT4, CELLSTYLE_60ACCENT4, CELLSTYLE_ACCENT4, CELLSTYLE_20ACCENT5, CELLSTYLE_40ACCENT5, CELLSTYLE_60ACCENT5, CELLSTYLE_ACCENT5, CELLSTYLE_20ACCENT6, CELLSTYLE_40ACCENT6, CELLSTYLE_60ACCENT6, CELLSTYLE_ACCENT6, CELLSTYLE_COMMA, CELLSTYLE_COMMA0, CELLSTYLE_CURRENCY, CELLSTYLE_CURRENCY0, CELLSTYLE_PERCENT}
 ***********************************************************************/

class XL {
	static _ := DllCall('LoadLibrary', 'str', A_LineFile '\..\' (A_PtrSize * 8) 'bit\libxl.dll', 'ptr')
	static Load(path, as_xlsx?) {
		if !FileExist(path)
			throw Error('Excel file does not exist.')
		if IsSet(as_xlsx)
			ext := as_xlsx ? 'xlsx' : 'xls'
		else SplitPath(path, , , &ext)
		handle := ext = 'xlsx' ? DllCall('libxl\xlCreateXMLBook', 'cdecl ptr') : DllCall('libxl\xlCreateBook', 'cdecl ptr')
		book := XL.IBook(handle)
		book.setKey('libxl', 'windows-28232b0208c4ee0369ba6e68abv6v5i3')
		if (book.load(path))
			return book
		throw Error('Failed to load')
	}
	static New(ext := 'xlsx') {
		book := XL.IBook(ext = 'xlsx' ? DllCall('libxl\xlCreateXMLBook', 'cdecl ptr') : DllCall('libxl\xlCreateBook', 'cdecl ptr'))
		book.setKey('libxl', 'windows-28232b0208c4ee0369ba6e68abv6v5i3')
		return book
	}
	class IBase {
		ptr := 0, parent := 0
		__New(handle, parent := 0) => (this.parent := parent, this.ptr := handle)
	}
	class IAutoFilter extends XL.IBase {
		; 获取筛选范围
		getRef(&rowFirst, &rowLast, &colFirst, &colLast) => (res := DllCall('libxl\xlAutoFilterGetRef', 'ptr', this, 'int*', &rowFirst := 0, 'int*', &rowLast := 0, 'int*', &colFirst := 0, 'int*', &colLast := 0, 'cdecl'), rowFirst++, rowLast++, colFirst++, colLast++, res)
		
		; 设置筛选范围
		setRef(rowFirst, rowLast, colFirst, colLast) => DllCall('libxl\xlAutoFilterSetRef', 'ptr', this, 'int', rowFirst-1, 'int', rowLast-1, 'int', colFirst-1, 'int', colLast-1, 'cdecl')
		
		; 返回第colId个筛选列
		column(colId) => XL.IFilterColumn(DllCall('libxl\xlAutoFilterColumn', 'ptr', this, 'int', colId-1, 'cdecl ptr'))
		
		; 获取筛选列总数
		columnSize() => DllCall('libxl\xlAutoFilterColumnSize', 'ptr', this, 'cdecl')
		
		; 返回第index个有筛选信息的筛选列
		columnByIndex(index) => XL.IFilterColumn(DllCall('libxl\xlAutoFilterColumnByIndex', 'ptr', this, 'int', index, 'cdecl ptr'))
		
		; 获取排序范围(不包含header),未排序时全部返回0
		getSortRange(&rowFirst, &rowLast, &colFirst, &colLast) => (res := DllCall('libxl\xlAutoFilterGetSortRange', 'ptr', this, 'int*', &rowFirst := -1, 'int*', &rowLast := -1, 'int*', &colFirst := -1, 'int*', &colLast := -1, 'cdecl'), rowFirst++, rowLast++, colFirst++, colLast++, res)
		
		; 获取排序列在筛选中的"索引"和"是否是倒叙"
		getSort(&columnIndex, &descending) => (res := DllCall('libxl\xlAutoFilterGetSort', 'ptr', this, 'int*', &columnIndex := -1, 'int*', &descending := 0, 'cdecl'), columnIndex++, res)
		
		; 设置排序,默认升序,执行时会取消多级排序
		setSort(columnIndex, descending := false) => DllCall('libxl\xlAutoFilterSetSort', 'ptr', this, 'int', columnIndex-1, 'int', descending, 'cdecl')
		
		; 添加排序,默认升序,用于多级排序中,当排序已存在时会报错
		addSort(columnIndex, descending := false) => DllCall('libxl\xlAutoFilterAddSort', 'ptr', this, 'int', columnIndex-1, 'int', descending, 'cdecl')
	}
	class IBook extends XL.IBase {
		path := ''
		active => this.getSheet(this.activeSheet())
		__Item[it] {
			get {
				count := this.sheetCount()
				if IsNumber(it) {
					if (it < 1 && it > count)
						throw Error('Invalid index')
					return this.getSheet(it)
				}
				Loop count
					if (this.getSheetName(A_Index) = it)
						return this.getSheet(A_Index)
				throw Error('table ' it ' does not exist')
			}
		}
		; 返回活动表的索引
		activeSheet() => DllCall('libxl\xlBookActiveSheet', 'ptr', this, 'cdecl') + 1
		
		; 为工作簿添加了一个新的条件格式，用于与条件格式化规则一起使用（仅限于xlsx文件）。
		addConditionalFormat(customNumFormat) => XL.IConditionalFormat(DllCall('libxl\xlBookAddConditionalFormat', 'ptr', this, 'cdecl ptr'))
		
		; 向工作簿添加新的自定义数字格式。格式字符串customNumFormat指示如何格式化和呈现单元格的数值。返回自定义格式标识符。如果过程中出现错误，该函数将返回0
		addCustomNumFormat(customNumFormat) => DllCall('libxl\xlBookAddCustomNumFormat', 'ptr', this, 'str', customNumFormat, 'cdecl')
		
		; 在工作簿中添加一个新的字体，其初始参数可以从其他字体复制。如果发生错误，则返回NULL
		addFont(initFont := 0) => XL.IFont(DllCall('libxl\xlRichStringAddFont', 'ptr', this, 'ptr', initFont, 'cdecl ptr'))
		
		; 向工作簿添加新格式，初始参数可以从其他格式复制。如果发生错误，返回NULL。请注意，加载文件后，它将被删除。
		addFormat(initFormat := 0) => XL.IFormat(DllCall('libxl\xlBookAddFormat', 'ptr', this, 'ptr', initFormat, 'cdecl ptr'))
		
		; 从预定义样式向工作簿添加新格式。如果发生错误，返回NULL。请注意，加载文件后，它将被删除。
		; CellStyle {CELLSTYLE_NORMAL, CELLSTYLE_BAD, CELLSTYLE_GOOD, CELLSTYLE_NEUTRAL, CELLSTYLE_CALC, CELLSTYLE_CHECKCELL, CELLSTYLE_EXPLANATORY, CELLSTYLE_INPUT, CELLSTYLE_OUTPUT, CELLSTYLE_HYPERLINK, CELLSTYLE_LINKEDCELL, CELLSTYLE_NOTE, CELLSTYLE_WARNING, CELLSTYLE_TITLE, CELLSTYLE_HEADING1, CELLSTYLE_HEADING2, CELLSTYLE_HEADING3, CELLSTYLE_HEADING4, CELLSTYLE_TOTAL, CELLSTYLE_20ACCENT1, CELLSTYLE_40ACCENT1, CELLSTYLE_60ACCENT1, CELLSTYLE_ACCENT1, CELLSTYLE_20ACCENT2, CELLSTYLE_40ACCENT2, CELLSTYLE_60ACCENT2, CELLSTYLE_ACCENT2, CELLSTYLE_20ACCENT3, CELLSTYLE_40ACCENT3, CELLSTYLE_60ACCENT3, CELLSTYLE_ACCENT3, CELLSTYLE_20ACCENT4, CELLSTYLE_40ACCENT4, CELLSTYLE_60ACCENT4, CELLSTYLE_ACCENT4, CELLSTYLE_20ACCENT5, CELLSTYLE_40ACCENT5, CELLSTYLE_60ACCENT5, CELLSTYLE_ACCENT5, CELLSTYLE_20ACCENT6, CELLSTYLE_40ACCENT6, CELLSTYLE_60ACCENT6, CELLSTYLE_ACCENT6, CELLSTYLE_COMMA, CELLSTYLE_COMMA0, CELLSTYLE_CURRENCY, CELLSTYLE_CURRENCY0, CELLSTYLE_PERCENT}
		addFormatFromStyle(style) => XL.IFormat(DllCall('libxl\xlBookAddFormatFromStyle', 'ptr', this, 'int', style, 'cdecl ptr'))
		
		; 向工作簿中添加图片。返回图片标识符。支持BMP， DIB， PNG， JPG和WMF图片格式。使用Sheet.setPicture()的图片标识符(one-based)。如果发生错误，返回0。
		addPicture(filename) => DllCall('libxl\xlBookAddPicture', 'ptr', this, 'str', filename, 'cdecl') + 1
		
		; 将图片从内存缓冲区添加到工作簿中：data-指向图片数据缓冲区的指针(BMP， DIB， PNG， JPG或WMF格式); size-缓冲区中数据的大小。返回图片标识符(one-based)。使用Sheet.setPicture()的图片标识符。如果发生错误，返回0。
		addPicture2(data, size) => DllCall('libxl\xlBookAddPicture2', 'ptr', this, 'ptr', data, 'uint', size, 'cdecl') + 1
		
		; 将图片作为链接添加到工作簿中（仅适用于xlsx文件）：
		; Insert = false -只存储到文件的链接；
		; Insert = true -存储图片和文件链接。
		; 返回图片标识符。支持BMP， DIB， PNG， JPG和WMF图片格式。使用Sheet.setPicture()的图片标识符(one-based)。如果发生错误，返回0。
		addPictureAsLink(filename, insert := false) => DllCall('libxl\xlBookAddPictureAsLink', 'ptr', this, 'str', filename, 'int', insert, 'cdecl') + 1
		
		; 使用Sheet.writeRichStr()方法向工作簿中添加一个新的富字符串，用于在单个单元格中使用不同的字体。不要手动释放返回指针。
		addRichString() => XL.IRichString(DllCall('libxl\xlBookAddRichString', 'ptr', this, 'cdecl ptr'))
		
		; 向本簿添加新工作表，返回工作表。使用initSheet参数，如果你想复制一个现有的工作表。注意initSheet必须只来自同一个工作簿
		addSheet(name, initSheet := 1) => XL.ISheet(DllCall('libxl\xlBookAddSheet', 'ptr', this, 'str', name, 'ptr', initSheet-1, 'cdecl ptr'), this)
		
		; 返回二进制文件的BIFF版本。仅用于xls格式。
		biffVersion() => DllCall('libxl\xlBookBiffVersion', 'ptr', this, 'cdecl')
		
		; 返回工作簿的计算模式
		calcMode() => DllCall('libxl\xlBookCalcMode', 'ptr', this, 'cdecl')
		
		; 将红色、绿色和蓝色成分打包进颜色类中
		colorPack(red, green, blue) => DllCall('libxl\xlBookColorPack', 'ptr', this, 'int', red, 'int', green, 'int', blue, 'cdecl')
		
		; 将一种颜色类分解成红色、绿色和蓝色三个组成部分
		colorUnpack(color, &red, &green, &blue) => DllCall('libxl\xlBookColorUnpack', 'ptr', this, 'int', color, 'int*', &red := 0, 'int*', &green := 0, 'int*', &blue := 0, 'cdecl')
		
		; 向工作簿添加新的自定义数字格式。格式字符串customNumFormat指示如何格式化和呈现单元格的数值。请参阅自定义格式字符串指南。返回自定义格式标识符。它在Format.setNumFormat()中使用。如果发生错误，返回0。
		customNumFormat(fmt) => DllCall('libxl\xlBookCustomNumFormat', 'ptr', this, 'int', fmt, 'cdecl str')
		
		; 将日期和时间的信息打包进一个双精度浮点数
		datePack(year, month, day, hour := 0, min := 0, sec := 0, msec := 0) => DllCall('libxl\xlBookDatePack', 'ptr', this, 'int', year, 'int', month, 'int', day, 'int', hour, 'int', min, 'int', sec, 'int', msec, 'cdecl double')
		
		; 从双精度浮点数中提取日期和时间信息，如果提取过程中出现错误，则返回false
		dateUnpack(value, &year, &month, &day, &hour := 0, &min := 0, &sec := 0, &msec := 0) => DllCall('libxl\xlBookDateUnpack', 'ptr', this, 'double', value, 'int*', &year := 0, 'int*', &month := 0, 'int*', &day := 0, 'int*', &hour := 0, 'int*', &min := 0, 'int*', &sec := 0, 'int*', &msec := 0, 'cdecl')
		
		; 返回此工作簿的默认字体名称和大小。如果发生错误，返回0。
		defaultFont(&fontSize) => DllCall('libxl\xlBookDefaultFont', 'ptr', this, 'int*', &fontSize := 0, 'cdecl str')

		; 删除具有指定索引的工作表。如果发生错误，返回false。
		delSheet(index) => DllCall('libxl\xlBookDelSheet', 'ptr', this, 'int', index-1, 'cdecl')

		;返回最后一个错误消息 
		errorMessage() => DllCall('libxl\xlBookErrorMessage', 'ptr', this, 'cdecl astr')

		; 返回第index个字体。index必须小于等于fontSize()方法的返回值
		font(index) => XL.IFont(DllCall('libxl\xlBookFont', 'ptr', this, 'int', index-1, 'cdecl ptr'))

		; 返回此工作簿中字体数量
		fontSize() => DllCall('libxl\xlBookFontSize', 'ptr', this, 'cdecl')

		; 返回第index个格式。索引必须小于formatSize()方法的返回值
		format(index) => XL.IFormat(DllCall('libxl\xlBookFormat', 'ptr', this, 'int', index-1, 'cdecl ptr'))

		; 返回此工作簿中格式数量
		formatSize() => DllCall('libxl\xlBookFormatSize', 'ptr', this, 'cdecl')

		/** 
		 * 返回内存缓冲区中位置索引处的图片
		 * @param {Integer} index 在工作簿中的位置
		 * @param {Pointer} data  对缓冲区的引用
		 * @param {Pointer} size  对保存大小的引用
		 * @returns {Integer} 返回图片类型 {0:PNG, 1:JPEG, 2:GIF, 3:WMF, 4:DIB, 5:EMF, 6:PICT, 7:TIFF, ERROR = 0xFF}
		 */
		getPicture(index, &data, &size) => DllCall('libxl\xlBookGetPicture', 'ptr', this, 'int', index-1, 'ptr*', &data := 0, 'uint*', &size := 0, 'cdecl')
		
		; 返回第index个工作表
		getSheet(index) => XL.ISheet(DllCall('libxl\xlBookGetSheet', 'ptr', this, 'int', index-1, 'cdecl ptr'), this)

		; 返回第index个工作表的名称
		getSheetName(index) => DllCall('libxl\xlBookGetSheetName', 'ptr', this, 'int', index-1, 'cdecl str')

		; 插入一个工作表
		insertSheet(index, name, initSheet := 1) => XL.ISheet(DllCall('libxl\xlBookInsertSheet', 'ptr', this, 'int', index-1, 'str', name, 'ptr', initSheet-1, 'cdecl ptr'), this)
		
		; 返回1904年日期系统是否处于激活状态。
		isDate1904() => DllCall('libxl\xlBookIsDate1904', 'ptr', this, 'cdecl')
		
		; 返回工作簿是否为模板
		isTemplate() => DllCall('libxl\xlBookIsTemplate', 'ptr', this, 'cdecl')
		
		; 返回工作簿是否为只读
		isWriteProtected() => DllCall('libxl\xlBookIsWriteProtected', 'ptr', this, 'cdecl')
		
		; 将整个文件装入内存。指定一个临时文件以减少内存消耗。如果发生错误，返回false
		load(filename, tempFile := '') => (this.path := filename, tempFile ? DllCall('libxl\xlBookLoadUsingTempFile', 'ptr', this, 'str', filename, 'str', tempFile, 'cdecl') : DllCall('libxl\xlBookLoad', 'ptr', this, 'str', filename, 'cdecl'))
		
		; 只加载有关工作表的信息。之后，您可以调用Book.sheetCount()和Book.getSheetName()方法。如果发生错误，返回false
		loadInfo(filename) => DllCall('libxl\xlBookLoadInfo', 'ptr', this, 'str', filename, 'cdecl')
		
		; 仅将具有指定工作表索引和行范围的文件加载到内存中。指定一个临时文件以减少内存消耗。如果keepAllSheets为true，则所有未指定的工作表将被保存回来，但它减少了内存消耗，否则只保存指定的工作表（其他工作表将被删除）。如果发生错误，返回false。
		loadPartially(filename, sheetIndex, firstRow, lastRow, tempFile := '') => (tempFile ? DllCall('libxl\xlBookLoadPartiallyUsingTempFile', 'ptr', this, 'str', filename, 'int', sheetIndex-1, 'int', firstRow-1, 'int', lastRow-1, 'str', tempFile, 'cdecl') : DllCall('libxl\xlBookLoadPartially', 'ptr', this, 'str', filename, 'int', sheetIndex-1, 'int', firstRow-1, 'int', lastRow-1, 'cdecl'))
		
		/**
		 * 从用户的内存缓冲区加载文件
		 * @param {Pointer} data       指向缓冲区的数据指针
		 * @param {Pointer} size       缓冲区中数据的大小
		 * @param {Number}  sheetIndex 只加载具有指定表索引的文件, 0 加载所有表
		 * @param {Number}  firstRow   加载范围的第一行，0 加载到lastRow之前的所有行
		 * @param {Number}  lastRow    加载范围的最后一行，0 加载firstRow之后的所有行；
		 */
		loadRaw(data, size, sheetIndex := 0, firstRow := 0, lastRow := 0) => (sheetIndex = 0 ? DllCall('libxl\xlBookLoadRaw', 'ptr', this, 'ptr', data, 'uint', size, 'cdecl') : DllCall('libxl\xlBookLoadRawPartially', 'ptr', this, 'astr', data, 'uint', size, 'int', sheetIndex-1, 'int', firstRow-1, 'int', lastRow-1, 'cdecl'))
		
		; 仅将具有指定表索引的文件加载到内存中。指定一个临时文件以减少内存消耗。如果keepAllSheets为true，则所有未指定的工作表将被保存回来，但它减少了内存消耗，否则只保存指定的工作表（其他工作表将被删除）。如果发生错误，返回false。
		loadSheet(filename, sheetIndex, tempFile := '') => (this.load(filename, tempFile), this.setActiveSheet(sheetIndex))
		
		; 加载不带空单元格的文件，其中包含格式化信息，以减少内存消耗
		loadWithoutEmptyCells(filename) => DllCall('libxl\xlBookLoadWithoutEmptyCells', 'ptr', this, 'str', filename, 'cdecl')
		
		; 获取带有srcIndex的表，并将其插入带有dstIndex的表的前面。如果发生错误，返回false
		moveSheet(srcIndex, dstIndex) => DllCall('libxl\xlBookMoveSheet', 'ptr', this, 'int', srcIndex-1, 'int', dstIndex-1, 'cdecl')
		
		; 返回此工作簿中的图片数量
		pictureSize() => DllCall('libxl\xlSheetPictureSize', 'ptr', this, 'cdecl')
		
		; 返回R1C1引用模式是否激活
		refR1C1() => DllCall('libxl\xlBookRefR1C1', 'ptr', this, 'cdecl')

		; 删除该对象并释放资源
		release() => (this.ptr ? (DllCall('libxl\xlBookRelease', 'ptr', this, 'cdecl'), this.ptr := 0) : 0)
		
		; 返回RGB模式是否激活
		rgbMode() => DllCall('libxl\xlBookRgbMode', 'ptr', this, 'cdecl')
		
		; 将当前工作簿保存到文件中。使用临时文件来减少内存消耗。如果发生错误，返回false。
		save(filename := '', useTempFile := false) {
			filename := filename || this.path
			if !(useTempFile ? DllCall('libxl\xlBookSaveUsingTempFile', 'ptr', this, 'str', filename, 'int', useTempFile, 'cdecl') : DllCall('libxl\xlBookSave', 'ptr', this, 'str', filename, 'cdecl'))
				throw Error(this.errorMessage())
		}
		
		; 将文件保存到内部内存缓冲区。data-指向缓冲区的数据指针；size-指向保存大小的指针。如果发生错误，返回false。
		saveRaw(&data, &size) => DllCall('libxl\xlBookSaveRaw', 'ptr', this, 'ptr*', &data := 0, 'uint*', &size := 0, 'cdecl')
		
		; 在此工作簿中设置活动工作表索引
		setActiveSheet(index) => DllCall('libxl\xlBookSetActiveSheet', 'ptr', this, 'int', index-1, 'cdecl')
		
		; 设置工作簿的计算模式：CalcModeType {0:MANUAL, 1:AUTO, 2:AUTONOTABLE}
		setCalcMode(CalcMode) => DllCall('libxl\xlBookSetCalcMode', 'ptr', this, 'int', calcMode, 'cdecl')
		
		; 设置日期系统模式：
		; false - 1900日期系统(默认), 在1900年的数据库系统中，下限为1900年1月1日，其序列值为1。
		; true  - 1904日期系统,       在1904年的数据库系统中，下限为1904年1月1日，其序列值为0。
		setDate1904(date1904 := true) => DllCall('libxl\xlBookSetDate1904', 'ptr', this, 'int', date1904, 'cdecl')
		
		; 设置此工作簿的默认字体名称和大小
		setDefaultFont(fontName, fontSize) => DllCall('libxl\xlBookSetDefaultFont', 'ptr', this, 'str', fontName, 'int', fontSize, 'cdecl')
		
		; 设置客户的许可密钥
		setKey(name, key) => DllCall('libxl\xlBookSetKey', 'ptr', this, 'str', name, 'str', key, 'cdecl')
		
		; 设置此库的区域设置。locale参数与C runtime Library中的setlocale（）函数中的locale参数相同。
		; 例如，输入“en_US”。UTF-8”允许在Linux或Mac中使用非ascii字符。它接受在Windows和其他操作系统中使用UTF-8字符编码的特殊值“UTF-8”。它对具有宽字符串的unicode项目没有影响(带有_UNICODE预处理器变量)。如果给出了有效的语言环境参数，则返回true。
		setLocale(locale) => DllCall('libxl\xlBookSetLocale', 'ptr', this, 'astr', locale, 'cdecl')
		
		; 设置R1C1引用模式：true - R1C1，false - A1（默认）
		setRefR1C1(refR1C1 := true) => DllCall('libxl\xlBookSetRefR1C1', 'ptr', this, 'int', refR1C1, 'cdecl')
		
		; 设置RGB模式：true - RGB模式，false -索引模式(默认值)。在RGB模式下，使用colorPack()和colorUnpack()方法获取/设置颜色
		setRgbMode(rgbMode := true) => DllCall('libxl\xlBookSetRgbMode', 'ptr', this, 'int', rgbMode, 'cdecl')
		
		; 设置模板标志：true -工作簿是模板，false -工作簿不是模板（默认）。它允许将文件类型从模板文件（xlt和xltx）更改为常规文件（xls和xlsx），反之亦然。
		setTemplate(tmpl := true) => DllCall('libxl\xlBookSetTemplate', 'ptr', this, 'int', tmpl, 'cdecl')
		
		; 返回工作表数量
		sheetCount() => DllCall('libxl\xlBookSheetCount', 'ptr', this, 'cdecl')
		
		; 返回第index个工作表类型。SheetType {0:SHEET, 1:CHART, 2:UNKNOWN}
		sheetType(index) => DllCall('libxl\xlBookSheetType', 'ptr', this, 'int', index-1, 'cdecl')
		
		; 返回使用的LibXL库的十六进制版本
		version() => DllCall('libxl\xlBookVersion', 'ptr', this, 'cdecl')
		
		__Delete() => this.release()
	}
	class IConditionalFormat extends XL.IBase {
		; 返回此条件格式当前使用的字体的句柄。
		font() => XL.IFont(DllCall('libxl\xlConditionalFormatFont', 'ptr', this, 'cdecl ptr'))
		
		; 返回数字格式标识符。
		numFormat() => DllCall('libxl\xlConditionalFormatNumFormat', 'ptr', this, 'cdecl')
		
		; 设置数字格式标识符。标识符必须是有效的内置数字格式标识符：
		setNumFormat(numFormat) => DllCall('libxl\xlConditionalFormatSetNumFormat', 'ptr', this, 'int', numFormat, 'cdecl')
		
		; 返回自定义数字格式字符串。
		customNumFormat() => DllCall('libxl\xlConditionalFormatCustomNumFormat', 'ptr', this, 'cdecl str')
		
		; 设置自定义数字格式字符串。格式字符串customNumFormat指示如何格式化和呈现单元格的数值。
		setCustomNumFormat(customNumFormat) => DllCall('libxl\xlConditionalFormatSetCustomNumFormat', 'ptr', this, 'str', customNumFormat, 'cdecl')
		
		; 设置边框 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		setBorder(style := 1) => DllCall('libxl\xlConditionalFormatSetBorder', 'ptr', this, 'int', style, 'cdecl')
		
		; 设置边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setBorderColor(color) => DllCall('libxl\xlConditionalFormatSetBorderColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回左边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		borderLeft() => DllCall('libxl\xlConditionalFormatBorderLeft', 'ptr', this, 'cdecl')
		
		; 设置左边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		setBorderLeft(style := 1) => DllCall('libxl\xlConditionalFormatSetBorderLeft', 'ptr', this, 'int', style, 'cdecl')
		
		; 返回右边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		borderRight() => DllCall('libxl\xlConditionalFormatBorderRight', 'ptr', this, 'cdecl')
		
		; 设置右边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		setBorderRight(style := 1) => DllCall('libxl\xlConditionalFormatSetBorderRight', 'ptr', this, 'int', style, 'cdecl')
		
		; 返回顶边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		borderTop() => DllCall('libxl\xlConditionalFormatBorderTop', 'ptr', this, 'cdecl')
		
		; 设置顶边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		setBorderTop(style := 1) => DllCall('libxl\xlConditionalFormatSetBorderTop', 'ptr', this, 'int', style, 'cdecl')
		
		; 返回底边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		borderBottom() => DllCall('libxl\xlConditionalFormatBorderBottom', 'ptr', this, 'cdecl')
		
		; 设置底边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		setBorderBottom(style := 1) => DllCall('libxl\xlConditionalFormatSetBorderBottom', 'ptr', this, 'int', style, 'cdecl')
		
		; 返回左边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		borderLeftColor() => DllCall('libxl\xlConditionalFormatBorderLeftColor', 'ptr', this, 'cdecl')
		
		; 设置右边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setBorderLeftColor(color) => DllCall('libxl\xlConditionalFormatSetBorderLeftColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回右边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		borderRightColor() => DllCall('libxl\xlConditionalFormatBorderRightColor', 'ptr', this, 'cdecl')
		
		; 设置右边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setBorderRightColor(color) => DllCall('libxl\xlConditionalFormatSetBorderRightColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回顶边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		borderTopColor() => DllCall('libxl\xlConditionalFormatBorderTopColor', 'ptr', this, 'cdecl')
		
		; 设置顶边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setBorderTopColor(color) => DllCall('libxl\xlConditionalFormatSetBorderTopColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回底边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		borderBottomColor() => DllCall('libxl\xlConditionalFormatBorderBottomColor', 'ptr', this, 'cdecl')
		
		; 设置底边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setBorderBottomColor(color) => DllCall('libxl\xlConditionalFormatSetBorderBottomColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回填充模式 FillPattern {NONE, SOLID, GRAY50, GRAY75, GRAY25, HORSTRIPE, VERSTRIPE, REVDIAGSTRIPE, DIAGSTRIPE, DIAGCROSSHATCH, THICKDIAGCROSSHATCH, THINHORSTRIPE, THINVERSTRIPE, THINREVDIAGSTRIPE, THINDIAGSTRIPE, THINHORCROSSHATCH, THINDIAGCROSSHATCH, GRAY12P5, GRAY6P25}
		fillPattern() => DllCall('libxl\xlConditionalFormatFillPattern', 'ptr', this, 'cdecl')
		
		; 设置填充模式 FillPattern {NONE, SOLID, GRAY50, GRAY75, GRAY25, HORSTRIPE, VERSTRIPE, REVDIAGSTRIPE, DIAGSTRIPE, DIAGCROSSHATCH, THICKDIAGCROSSHATCH, THINHORSTRIPE, THINVERSTRIPE, THINREVDIAGSTRIPE, THINDIAGSTRIPE, THINHORCROSSHATCH, THINDIAGCROSSHATCH, GRAY12P5, GRAY6P25}
		setFillPattern(pattern) => DllCall('libxl\xlConditionalFormatSetFillPattern', 'ptr', this, 'int', pattern, 'cdecl')
		
		; 返回填充颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		patternForegroundColor() => DllCall('libxl\xlConditionalFormatPatternForegroundColor', 'ptr', this, 'cdecl')
		
		; 设置填充颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setPatternForegroundColor(color) => DllCall('libxl\xlConditionalFormatSetPatternForegroundColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回填充背景颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		patternBackgroundColor() => DllCall('libxl\xlConditionalFormatPatternBackgroundColor', 'ptr', this, 'cdecl')
		
		; 设置填充背景颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setPatternBackgroundColor(color) => DllCall('libxl\xlConditionalFormatSetPatternBackgroundColor', 'ptr', this, 'int', color, 'cdecl')
	}
	class IConditionalFormatting extends XL.IBase {
		; 向这些条件格式规则添加一个范围。
		addRange(rowFirst, rowLast, colFirst, colLast) => DllCall('libxl\xlConditionalFormattingAddRange', 'ptr', this, 'int', rowFirst-1, 'int', rowLast-1, 'int', colFirst-1, 'int', colLast-1, 'cdecl')
		
		/**
		 * 添加条件格式规则，该规则突出显示值与指定条件对应的单元格
		 * @param type                   条件格式规则类型 CFormatType {CFORMAT_BEGINWITH, CFORMAT_CONTAINSBLANKS, CFORMAT_CONTAINSERRORS, CFORMAT_CONTAINSTEXT, CFORMAT_DUPLICATEVALUES, CFORMAT_ENDSWITH, CFORMAT_EXPRESSION, CFORMAT_NOTCONTAINSBLANKS, CFORMAT_NOTCONTAINSERRORS, CFORMAT_NOTCONTAINSTEXT, CFORMAT_UNIQUEVALUES}
		 * @param cFormat                用于突出显示单元格的条件格式，使用Book.AddConditionalFormat()来添加条件格式
		 * @param value                  指定条件格式规则的标准
		 * @param {Integer} stopIfTrue   当该规则为true时，如果不为零，则没有优先级较低的规则可以应用于该规则
		 */
		addRule(type, cFormat, value?, stopIfTrue := false) => DllCall('libxl\xlConditionalFormattingAddRule', 'ptr', this, 'int', type, 'ptr', cFormat, IsSet(value) ? 'str' : 'ptr', value ?? 0, 'char', stopIfTrue, 'cdecl')
		
		/**
		 * 添加条件格式规则，突出显示值位于[top N]或[bottom N]括号内的单元格
		 * @param cFormat                用于突出显示单元格的条件格式，使用Book.AddConditionalFormat()来添加条件格式
		 * @param value                  指定[top N]或[bottom N]括号
		 * @param {Integer} bottom       对于[下N]规则为真，对于[上N]规则为假
		 * @param {Integer} percent      对于百分比顶/底规则为true
		 * @param {Integer} stopIfTrue   当此规则为true时，没有优先级较低的规则可以应用于此规则
		 */
		addTopRule(cFormat, value, bottom := false, percent := false, stopIfTrue := 0) => DllCall('libxl\xlConditionalFormattingAddTopRule', 'ptr', this, 'ptr', cFormat, 'int', value, 'char', bottom, 'char', percent, 'char', stopIfTrue, 'cdecl')
		
		/**
		 * 添加一个条件格式规则，该规则使用操作符突出显示其值与计算结果进行比较的单元格
		 * @param {Integer} op          条件格式规则中的操作符 CFormatOperator {CFOPERATOR_LESSTHAN, CFOPERATOR_LESSTHANOREQUAL, CFOPERATOR_EQUAL, CFOPERATOR_NOTEQUAL, CFOPERATOR_GREATERTHANOREQUAL, CFOPERATOR_GREATERTHAN, CFOPERATOR_BETWEEN, CFOPERATOR_NOTBETWEEN, CFOPERATOR_CONTAINSTEXT, CFOPERATOR_NOTCONTAINS, CFOPERATOR_BEGINSWITH, CFOPERATOR_ENDSWITH}
		 * @param cFormat               用于突出显示单元格的条件格式，使用Book.AddConditionalFormat()来添加条件格式
		 * @param value1                带有指定运算符的表达式的数值
		 * @param {Integer} value2      仅对某些操作符，表达式的第二个可选数值
		 * @param {Integer} stopIfTrue  当此规则为true时，没有优先级较低的规则可以应用于此规则
		 */
		addOpNumRule(op, cFormat, value1, value2 := 0, stopIfTrue := false) => DllCall('libxl\xlConditionalFormattingAddOpNumRule', 'ptr', this, 'int', op, 'ptr', cFormat, 'double', value1, 'double', value2, 'char', stopIfTrue, 'cdecl')

		/**
		 * 添加一个条件格式规则，该规则使用操作符突出显示其值与计算结果进行比较的单元格
		 * @param {Integer} op          条件格式规则中的操作符 CFormatOperator {CFOPERATOR_LESSTHAN, CFOPERATOR_LESSTHANOREQUAL, CFOPERATOR_EQUAL, CFOPERATOR_NOTEQUAL, CFOPERATOR_GREATERTHANOREQUAL, CFOPERATOR_GREATERTHAN, CFOPERATOR_BETWEEN, CFOPERATOR_NOTBETWEEN, CFOPERATOR_CONTAINSTEXT, CFOPERATOR_NOTCONTAINS, CFOPERATOR_BEGINSWITH, CFOPERATOR_ENDSWITH}
		 * @param cFormat               用于突出显示单元格的条件格式，使用Book.AddConditionalFormat()来添加条件格式
		 * @param {String} value1       带有指定运算符的表达式的字符串值
		 * @param value2                仅对某些操作符为表达式的字符串第二个可选值
		 * @param {Integer} stopIfTrue  当此规则为true时，没有优先级较低的规则可以应用于此规则
		 */
		addOpStrRule(op, cFormat, value1, value2?, stopIfTrue := false) => DllCall('libxl\xlConditionalFormattingAddOpStrRule', 'ptr', this, 'int', op, 'ptr', cFormat, 'str', value1, IsSet(value2) ? 'str' : 'ptr', value2 ?? 0, 'char', stopIfTrue, 'cdecl')
		
		/**
		 * 添加一个条件格式规则，该规则突出显示区域中所有值高于或低于平均值的单元格
		 * @param cFormat                   用于突出显示单元格的条件格式，使用Book.AddConditionalFormat()来添加条件格式
		 * @param {Integer} aboveAverage    高于平均水平时为true，低于平均水平时为false
		 * @param {Integer} equalAverage    包含平均值本身为true，不包含该值为false，仅对[高于平均值]规则有效
		 * @param {Integer} stdDev          包括高于或低于平均值的标准偏差数，仅适用于[高于平均值]规则
		 * @param {Integer} stopIfTrue      当此规则为true时，没有优先级较低的规则可以应用于此规则
		 */
		addAboveAverageRule(cFormat, aboveAverage := true, equalAverage := false, stdDev := 0, stopIfTrue := false) => DllCall('libxl\xlConditionalFormattingAddAboveAverageRule', 'ptr', this, 'ptr', cFormat, 'char', aboveAverage, 'char', equalAverage, 'int', stdDev, 'char', stopIfTrue, 'cdecl')
		
		/**
		 * 添加条件格式规则，该规则突出显示包含指定时间段内日期的单元格
		 * @param cFormat                 用于突出显示单元格的条件格式，使用Book.AddConditionalFormat()来添加条件格式
		 * @param {Integer} timePeriod    适用的时间段 CFormatTimePeriod {CFTP_LAST7DAYS, CFTP_LASTMONTH, CFTP_LASTWEEK, CFTP_NEXTMONTH, CFTP_NEXTWEEK, CFTP_THISMONTH, CFTP_THISWEEK, CFTP_TODAY, CFTP_TOMORROW, CFTP_YESTERDAY}
		 * @param {Integer} stopIfTrue    当此规则为true时，没有优先级较低的规则可以应用于此规则
		 */
		addTimePeriodRule(cFormat, timePeriod, stopIfTrue := false) => DllCall('libxl\xlConditionalFormattingAddTimePeriodRule', 'ptr', this, 'ptr', cFormat, 'int', timePeriod, 'char', stopIfTrue, 'cdecl')
		
		/**
		 * 添加条件格式规则，该规则可在单元格上创建渐变的2色比例
		 * @param minColor              最小值的颜色
		 * @param maxColor              最大值的颜色
		 * @param {Integer} minType     CFVOType {CFVO_MIN, CFVO_MAX, CFVO_FORMULA, CFVO_NUMBER, CFVO_PERCENT, CFVO_PERCENTILE}
		 * @param {Integer} minValue    最小值的数值
		 * @param {Integer} maxType     CFVOType {CFVO_MIN, CFVO_MAX, CFVO_FORMULA, CFVO_NUMBER, CFVO_PERCENT, CFVO_PERCENTILE}
		 * @param {Integer} maxValue    最大值的数值
		 * @param {Integer} stopIfTrue  当此规则为true时，没有优先级较低的规则可以应用于此规则
		 */
		add2ColorScaleRule(minColor, maxColor, minType := 0, minValue := 0, maxType := 1, maxValue := 0, stopIfTrue := false) => DllCall('libxl\xlConditionalFormattingAdd2ColorScaleRule', 'ptr', this, 'int', minColor, 'int', maxColor, 'int', minType, 'double', minValue, 'int', maxType, 'double', maxValue, 'char', stopIfTrue, 'cdecl')
		
		/**
		 * 添加条件格式规则，该规则可在单元格上创建渐变的2色比例
		 * @param minColor              最小值的颜色
		 * @param maxColor              最大值的颜色
		 * @param {Integer} minType     CFVOType {CFVO_MIN, CFVO_MAX, CFVO_FORMULA, CFVO_NUMBER, CFVO_PERCENT, CFVO_PERCENTILE}
		 * @param {Integer} minValue    最小值的数值
		 * @param {Integer} maxType     CFVOType {CFVO_MIN, CFVO_MAX, CFVO_FORMULA, CFVO_NUMBER, CFVO_PERCENT, CFVO_PERCENTILE}
		 * @param {Integer} maxValue    最大值的数值
		 * @param {Integer} stopIfTrue  当此规则为true时，没有优先级较低的规则可以应用于此规则
		 */
		add2ColorScaleFormulaRule(minColor, maxColor, minType := 2, minValue?, maxType := 2, maxValue?, stopIfTrue := false) => DllCall('libxl\xlConditionalFormattingAdd2ColorScaleFormulaRule', 'ptr', this, 'int', minColor, 'int', maxColor, 'int', minType, IsSet(minValue) ? 'str' : 'ptr', minValue ?? 0, 'int', maxType, IsSet(maxValue) ? 'str' : 'ptr', maxValue ?? 0, 'char', stopIfTrue, 'cdecl')

		/**
		 * 添加条件格式规则，该规则可在单元格上创建渐变的3色比例
		 * @param minColor              最小值的颜色
		 * @param midColor              中间值的颜色
		 * @param maxColor              最大值的颜色
		 * @param {Integer} minType     CFVOType {CFVO_MIN, CFVO_MAX, CFVO_FORMULA, CFVO_NUMBER, CFVO_PERCENT, CFVO_PERCENTILE}
		 * @param {Integer} minValue    最小值的数值
		 * @param {Integer} midType     CFVOType {CFVO_MIN, CFVO_MAX, CFVO_FORMULA, CFVO_NUMBER, CFVO_PERCENT, CFVO_PERCENTILE}
		 * @param {Integer} midValue    中间值的数值
		 * @param {Integer} maxType     CFVOType {CFVO_MIN, CFVO_MAX, CFVO_FORMULA, CFVO_NUMBER, CFVO_PERCENT, CFVO_PERCENTILE}
		 * @param {Integer} maxValue    最大值的数值
		 * @param {Integer} stopIfTrue  当此规则为true时，没有优先级较低的规则可以应用于此规则
		 */
		add3ColorScaleRule(minColor, midColor, maxColor, minType := 0, minValue := 0, midType := 5, midValue := 50, maxType := 1, maxValue := 0, stopIfTrue := false) => DllCall('libxl\xlConditionalFormattingAdd3ColorScaleRule', 'ptr', this, 'int', minColor, 'int', midColor, 'int', maxColor, 'int', minType, 'double', minValue, 'int', midType, 'double', midValue, 'int', maxType, 'double', maxValue, 'char', stopIfTrue, 'cdecl')
		
		/**
		 * 添加条件格式规则，该规则可在单元格上创建渐变的3色比例
		 * @param minColor              最小值的颜色
		 * @param midColor              中间值的颜色
		 * @param maxColor              最大值的颜色
		 * @param {Integer} minType     CFVOType {CFVO_MIN, CFVO_MAX, CFVO_FORMULA, CFVO_NUMBER, CFVO_PERCENT, CFVO_PERCENTILE}
		 * @param {Integer} minValue    最小值的数值
		 * @param {Integer} midType     CFVOType {CFVO_MIN, CFVO_MAX, CFVO_FORMULA, CFVO_NUMBER, CFVO_PERCENT, CFVO_PERCENTILE}
		 * @param {Integer} midValue    中间值的数值
		 * @param {Integer} maxType     CFVOType {CFVO_MIN, CFVO_MAX, CFVO_FORMULA, CFVO_NUMBER, CFVO_PERCENT, CFVO_PERCENTILE}
		 * @param {Integer} maxValue    最大值的数值
		 * @param {Integer} stopIfTrue  当此规则为true时，没有优先级较低的规则可以应用于此规则
		 */
		add3ColorScaleFormulaRule(minColor, midColor, maxColor, minType := 2, minValue?, midType := 2, midValue?, maxType := 2, maxValue?, stopIfTrue := false) => DllCall('libxl\xlConditionalFormattingAdd3ColorScaleFormulaRule', 'ptr', this, 'int', minColor, 'int', midColor, 'int', maxColor, 'int', minType, IsSet(minValue) ? 'str' : 'ptr', minValue ?? 0, 'int', midType, IsSet(midValue) ? 'str' : 'ptr', midValue ?? 0, 'int', maxType, IsSet(maxValue) ? 'str' : 'ptr', maxValue ?? 0, 'char', stopIfTrue, 'cdecl')
	}
	class IFilterColumn extends XL.IBase {
		; 返回在筛选列中的索引
		index() => DllCall('libxl\xlFilterColumnIndex', 'ptr', this, 'cdecl') + 1
		
		; 返回筛选列类型 Filter {0:VALUE, 1:TOP10, 2:CUSTOM, 3:DYNAMIC, 4:COLOR, 5:ICON, 6:EXT, 7:NOT_SET}
		filterType() => DllCall('libxl\xlFilterColumnFilterType', 'ptr', this, 'cdecl')
		
		; 返回自定义筛选里的条件数量
		filterSize() => DllCall('libxl\xlFilterColumnFilterSize', 'ptr', this, 'cdecl')
		
		; 返回自定义筛选里的第index个条件
		filter(index) => DllCall('libxl\xlFilterColumnFilter', 'ptr', this, 'int', index-1, 'cdecl str')
		
		; 在筛选列表里的value前面打勾,value不存在时不会报错
		addFilter(value) => DllCall('libxl\xlFilterColumnAddFilter', 'ptr', this, 'str', value, 'cdecl')
		
		; 返回top10设置状态, value 设置的数值(0时为未设置), top {1:最大值, 0:最小值}, percent {1:百分比, 0:绝对值}
		getTop10(&value, &top, &percent) => DllCall('libxl\xlFilterColumnGetTop10', 'ptr', this, 'double*', &value := 0, 'int*', &top := 0, 'int*', &percent := 0, 'cdecl')
		
		; 设置top10设置状态, value 设置的数值(0时无法生效), top {1:最大值, 0:最小值}, percent {1:百分比, 0:绝对值}
		setTop10(value, top := true, percent := false) => DllCall('libxl\xlFilterColumnSetTop10', 'ptr', this, 'double', value, 'int', top, 'int', percent, 'cdecl')
		
		; 返回自定义筛选里的条件 op1 op2 {0:EQUAL, 1:GREATER_THAN, 2:GREATER_THAN_OR_EQUAL, 3:LESS_THAN, 4:LESS_THAN_OR_EQUAL, 5:NOT_EQUAL}
		;                       AndOp {1:AND, 0:OR}
		getCustomFilter(&op1, &v1, &op2, &v2, &andOp) => DllCall('libxl\xlFilterColumnGetCustomFilter', 'ptr', this, 'int*', &op1 := 0, 'str*', &v1 := '', 'int*', &op2 := 0, 'str*', &v2 := '', 'int*', &andOp := 0, 'cdecl')
		
		; 设置自定义筛选里的条件 op1 op2 {0:EQUAL, 1:GREATER_THAN, 2:GREATER_THAN_OR_EQUAL, 3:LESS_THAN, 4:LESS_THAN_OR_EQUAL, 5:NOT_EQUAL}
		;                       AndOp {1:AND, 0:OR}
		; **有BUG,要先执行清空clear(),不然第一个条件设置不上
		setCustomFilter(op1, v1, op2 := 0, v2 := '', andOp := false) => DllCall('libxl\xlFilterColumnSetCustomFilterEx', 'ptr', this, 'int', op1, 'str', v1, 'int', op2, 'str', v2, 'int', andOp, 'cdecl')
		
		; 清除筛选
		clear() => DllCall('libxl\xlFilterColumnClear', 'ptr', this, 'cdecl')
	}
	class IFormControl extends XL.IBase {
		; 返回类型 ObjectType {OBJECT_UNKNOWN, OBJECT_BUTTON, OBJECT_CHECKBOX, OBJECT_DROP, OBJECT_GBOX, OBJECT_LABEL, OBJECT_LIST, OBJECT_RADIO, OBJECT_SCROLL, OBJECT_SPIN, OBJECT_EDITBOX, OBJECT_DIALOG}
		objectType() => DllCall('libxl\xlFormControlObjectType', 'ptr', this, 'cdecl')

		; 返回是选中了复选框还是选中了单选按钮。此属性仅适用于复选框和单选按钮表单控件。
		checked() => DllCall('libxl\xlFormControlChecked', 'ptr', this, 'cdecl')

		; 设置是选中复选框还是选中单选按钮。此属性仅适用于复选框和单选按钮表单控件。
		; CheckedType {CHECKEDTYPE_UNCHECKED, CHECKEDTYPE_CHECKED, CHECKEDTYPE_MIXED}
		setChecked(checked) => DllCall('libxl\xlFormControlSetChecked', 'ptr', this, 'int', checked, 'cdecl')

		; 返回链接到的组框中的单元格引用。仅适用于组框窗体控件。
		fmlaGroup() => DllCall('libxl\xlFormControlFmlaGroup', 'ptr', this, 'cdecl str')

		; 设置链接到的组框中的单元格引用。仅适用于组框窗体控件。
		setFmlaGroup(group) => DllCall('libxl\xlFormControlSetFmlaGroup', 'ptr', this, 'str', group, 'cdecl')

		; 返回链接到的单元格引用。仅适用于复选框、单选按钮、滚动条、旋转框、下拉框和列表框。
		fmlaLink() => DllCall('libxl\xlFormControlFmlaLink', 'ptr', this, 'cdecl str')
		
		; 设置链接到的单元格引用。仅适用于复选框、单选按钮、滚动条、旋转框、下拉框和列表框。
		setFmlaLink(link) => DllCall('libxl\xlFormControlSetFmlaLink', 'ptr', this, 'str', link, 'cdecl')
		
		; 返回具有源数据单元格范围的单元格引用。此属性仅适用于列表框和下拉表单控件。
		fmlaRange() => DllCall('libxl\xlFormControlFmlaRange', 'ptr', this, 'cdecl str')
		
		; 使用源数据单元格的范围设置单元格引用。此属性仅适用于列表框和下拉表单控件。
		setFmlaRange(range) => DllCall('libxl\xlFormControlSetFmlaRange', 'ptr', this, 'str', range, 'cdecl')
		
		; 返回包含表单控件对象的数据链接到的源数据的单元格引用。此属性仅适用于标签和编辑框表单控件。
		fmlaTxbx() => DllCall('libxl\xlFormControlFmlaTxbx', 'ptr', this, 'cdecl str')
		
		; 使用表单控件对象的数据链接到的源数据设置单元格引用。可以指定任何单元格范围，但只考虑该范围中的第一个单元格。此属性仅适用于标签和编辑框表单控件。
		setFmlaTxbx(txbx) => DllCall('libxl\xlFormControlSetFmlaTxbx', 'ptr', this, 'str', txbx, 'cdecl')
		
		; 返回此嵌入控件的名称。
		name() => DllCall('libxl\xlFormControlName', 'ptr', this, 'cdecl str')
		
		; 返回链接到控件值的工作表范围。
		linkedCell() => DllCall('libxl\xlFormControlLinkedCell', 'ptr', this, 'cdecl str')
		
		; 返回用于填充列表框的源数据单元格的范围。
		listFillRange() => DllCall('libxl\xlFormControlListFillRange', 'ptr', this, 'cdecl str')
		
		; 返回与对象关联的宏。
		macro() => DllCall('libxl\xlFormControlMacro', 'ptr', this, 'cdecl str')
		
		; 返回对象的替代文本。
		altText() => DllCall('libxl\xlFormControlAltText', 'ptr', this, 'cdecl str')
		
		; 返回在工作表被保护时对象是否被锁定。
		locked() => DllCall('libxl\xlFormControlLocked', 'ptr', this, 'cdecl')
		
		; 返回对象是否处于其默认大小。
		defaultSize() => DllCall('libxl\xlFormControlDefaultSize', 'ptr', this, 'cdecl')
		
		; 返回在打印文档时是否打印对象。
		print() => DllCall('libxl\xlFormControlPrint', 'ptr', this, 'cdecl')
		
		; 返回是否允许对象运行附加的宏。
		disabled() => DllCall('libxl\xlFormControlDisabled', 'ptr', this, 'cdecl')
		
		; 按索引从列表框或下拉窗体控件返回项。
		item(index) => DllCall('libxl\xlFormControlItem', 'ptr', this, 'int', index-1, 'cdecl str')
		
		; 返回列表框或下拉表单控件中的项目数量。
		itemSize() => DllCall('libxl\xlFormControlItemSize', 'ptr', this, 'cdecl')
		
		; 将项添加到列表框或下拉窗体控件。
		addItem(value) => DllCall('libxl\xlFormControlAddItem', 'ptr', this, 'str', value, 'cdecl')
		
		; 将项插入到列表框或下拉窗体控件的指定位置。
		insertItem(index, value) => DllCall('libxl\xlFormControlInsertItem', 'ptr', this, 'int', index-1, 'str', value, 'cdecl')
		
		; 清除列表框或下拉窗体控件中的所有项。
		clearItems() => DllCall('libxl\xlFormControlClearItems', 'ptr', this, 'cdecl')
		
		; 返回添加滚动条之前下拉列表中的行数。
		dropLines() => DllCall('libxl\xlFormControlDropLines', 'ptr', this, 'cdecl')
		
		; 设置添加滚动条之前下拉框中的行数。此属性仅适用于下拉表单控件。最小为0，最大为30000。
		setDropLines(lines) => DllCall('libxl\xlFormControlSetDropLines', 'ptr', this, 'int', lines, 'cdecl')
		
		; 以像素为单位返回滚动条的宽度。
		dx() => DllCall('libxl\xlFormControlDx', 'ptr', this, 'cdecl')
		
		; 以像素为单位设置滚动条的宽度。此属性仅适用于列表框、滚动条、旋转框和下拉框。
		setDx(dx) => DllCall('libxl\xlFormControlSetDx', 'ptr', this, 'int', dx, 'cdecl')
		
		; 返回对象是否是单选按钮集合中的第一个按钮。
		firstButton() => DllCall('libxl\xlFormControlFirstButton', 'ptr', this, 'cdecl')
		
		; 设置对象是否为单选按钮集中的第一个按钮。此属性仅适用于单选按钮表单控件。
		setFirstButton(firstButton) => DllCall('libxl\xlFormControlSetFirstButton', 'ptr', this, 'int', firstButton, 'cdecl')
		
		; 返回滚动条是否水平。
		horiz() => DllCall('libxl\xlFormControlHoriz', 'ptr', this, 'cdecl')
		
		; 设置滚动条是否水平。此属性仅适用于滚动条窗体控件。
		setHoriz(horiz) => DllCall('libxl\xlFormControlSetHoriz', 'ptr', this, 'int', horiz, 'cdecl')
		
		; 返回由于增量单击而对滚动条或自旋框窗体控件的当前值所做的更改。
		inc() => DllCall('libxl\xlFormControlInc', 'ptr', this, 'cdecl')
		
		; 设置由于增量单击而对滚动条或自旋框窗体控件的当前值所做的更改。最少为0，最多为30000。此属性仅适用于滚动条或旋转框窗体控件。
		setInc(inc) => DllCall('libxl\xlFormControlSetInc', 'ptr', this, 'int', inc, 'cdecl')
		
		; 返回滚动条或自旋框生成的最大值。
		getMax() => DllCall('libxl\xlFormControlGetMax', 'ptr', this, 'cdecl')
		
		; 设置滚动条或旋转框生成的最大值。最少为0，最多为30000。此属性仅适用于滚动条和旋转框。
		setMax(max) => DllCall('libxl\xlFormControlSetMax', 'ptr', this, 'int', max, 'cdecl')
		
		; 返回由滚动条或自旋框生成的最小值。
		getMin() => DllCall('libxl\xlFormControlGetMin', 'ptr', this, 'cdecl')
		
		; 设置滚动条或自旋框生成的最小值。最少为0，最多为30000。此属性仅适用于滚动条和旋转框。
		setMin(min) => DllCall('libxl\xlFormControlSetMin', 'ptr', this, 'int', min, 'cdecl')
		
		; 以逗号分隔的列表形式返回选定项的索引。列表索引是基于1的。只有当属性selType的值为“multi”时，此属性才有效。此属性仅适用于列表框窗体控件。
		multiSel() => DllCall('libxl\xlFormControlMultiSel', 'ptr', this, 'cdecl str')
		
		; 将选定项的索引设置为逗号分隔的列表。列表索引是基于1的。只有当属性selType的值为“multi”时，此属性才有效。此属性仅适用于列表框窗体控件。
		setMultiSel(value) => DllCall('libxl\xlFormControlSetMultiSel', 'ptr', this, 'str', value, 'cdecl')
		
		; 所选项的索引。该指数是以1为基础的。如果设置为0，则不选择任何项。此属性仅适用于列表框和下拉表单控件。
		sel() => DllCall('libxl\xlFormControlSel', 'ptr', this, 'cdecl')
		
		; 设置所选项的索引。该指数是以1为基础的。如果设置为0，则不选择任何项。此属性仅适用于列表框和下拉表单控件。
		setSel(sel) => DllCall('libxl\xlFormControlSetSel', 'ptr', this, 'int', sel, 'cdecl')
		
		/**
		 * 获取文档中控件左上角位置。
		 * @param {String} col     列号
		 * @param {String} colOff  在EMUs中与列的偏移量
		 * @param {String} row     行号
		 * @param {String} rowOff  在EMUs中与行的偏移量
		 * @returns {Boolean}      找到返回1，未找到返回0。
		 */
		fromAnchor(&col, &colOff, &row, &rowOff) => (res := DllCall('libxl\xlFormControlFromAnchor', 'ptr', this, 'int*', &col := -1, 'int*', &colOff := 0, 'int*', &row := -1, 'int*', &rowOff := 0, 'cdecl'), row++, col++, res)
		
		/**
		 * 获取文档中控件右下角位置。
		 * @param {String} col     列号
		 * @param {String} colOff  在EMUs中与列的偏移量
		 * @param {String} row     行号
		 * @param {String} rowOff  在EMUs中与行的偏移量
		 * @returns {Boolean}      找到返回1，未找到返回0。
		 */
		toAnchor(&col, &colOff, &row, &rowOff) => (res := DllCall('libxl\xlFormControlToAnchor', 'ptr', this, 'int*', &col := 0, 'int*', &colOff := 0, 'int*', &row := 0, 'int*', &rowOff := 0, 'cdecl'), rwo++, col++, res)
	}
	class IFont extends XL.IBase {
		; 返回字体大小
		size() => DllCall('libxl\xlFontSize', 'ptr', this, 'cdecl')

		; 设置字体大小
		setSize(size) => DllCall('libxl\xlFontSetSize', 'ptr', this, 'int', size, 'cdecl')

		; 是否是斜体
		italic() => DllCall('libxl\xlFontItalic', 'ptr', this, 'cdecl')

		; 设置斜体, true - 斜体, false - 正常
		setItalic(italic := true) => DllCall('libxl\xlFontSetItalic', 'ptr', this, 'int', italic, 'cdecl')

		; 是否有删除线
		strikeOut() => DllCall('libxl\xlFontStrikeOut', 'ptr', this, 'cdecl')

		; 设置删除线, true - 有删除线, false - 无删除线
		setStrikeOut(strikeOut := true) => DllCall('libxl\xlFontSetStrikeOut', 'ptr', this, 'int', strikeOut, 'cdecl')

		; 返回字体颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		color() => DllCall('libxl\xlFontColor', 'ptr', this, 'cdecl')

		; 设置字体颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setColor(Color) => DllCall('libxl\xlFontSetColor', 'ptr', this, 'int', color, 'cdecl')

		; 是否是粗体
		bold() => DllCall('libxl\xlFontBold', 'ptr', this, 'cdecl')

		; 设置粗体
		setBold(bold) => DllCall('libxl\xlFontSetBold', 'ptr', this, 'int', bold, 'cdecl')

		; 返回上下标样式 Script {NORMAL = 0, SUPER = 1, SUB = 2}
		script() => DllCall('libxl\xlFontScript', 'ptr', this, 'cdecl')

		; 设置上下标样式 Script {NORMAL = 0, SUPER = 1, SUB = 2}
		setScript(Script) => DllCall('libxl\xlFontSetScript', 'ptr', this, 'int', script, 'cdecl')

		; 返回下划线样式 Underline {NONE = 0, SINGLE, DOUBLE, SINGLEACC = 0x21, DOUBLEACC = 0x22}
		underline() => DllCall('libxl\xlFontUnderline', 'ptr', this, 'cdecl')

		; 设置下划线样式 Underline {NONE = 0, SINGLE, DOUBLE, SINGLEACC = 0x21, DOUBLEACC = 0x22}
		setUnderline(Underline) => DllCall('libxl\xlFontSetUnderline', 'ptr', this, 'int', underline, 'cdecl')

		; 返回字体名称, 默认名称为"Arial"
		name() => DllCall('libxl\xlFontName', 'ptr', this, 'cdecl str')

		; 设置字体名称
		setName(name) => DllCall('libxl\xlFontSetName', 'ptr', this, 'str', name, 'cdecl')
	}
	class IFormat extends XL.IBase {

		; 返回当前字体
		font() => XL.IFont(DllCall('libxl\xlFormatFont', 'ptr', this, 'cdecl ptr'))

		; 设置当前字体
		setFont(font) => DllCall('libxl\xlFormatSetFont', 'ptr', this, 'ptr', font, 'cdecl')

		; 返回数字格式标识符 {GENERAL = 0, NUMBER = 1, NUMBER_D2 = 2, NUMBER_SEP = 3, NUMBER_SEP_D2 = 4, CURRENCY_NEGBRA = 5, CURRENCY_NEGBRARED = 6, CURRENCY_D2_NEGBRA = 7, CURRENCY_D2_NEGBRARED = 8, PERCENT = 9, PERCENT_D2 = 10, SCIENTIFIC_D2 = 11, FRACTION_ONEDIG = 12, FRACTION_TWODIG = 13, DATE = 14, CUSTOM_D_MON_YY = 15, CUSTOM_D_MON = 16, CUSTOM_MON_YY = 17, CUSTOM_HMM_AM = 18, CUSTOM_HMMSS_AM = 19, CUSTOM_HMM = 20, CUSTOM_HMMSS = 21, CUSTOM_MDYYYY_HMM = 22, NUMBER_SEP_NEGBRA=37 = 23, NUMBER_SEP_NEGBRARED = 24, NUMBER_D2_SEP_NEGBRA = 25, NUMBER_D2_SEP_NEGBRARED = 26, ACCOUNT = 27, ACCOUNTCUR = 28, ACCOUNT_D2 = 29, ACCOUNT_D2_CUR = 30, CUSTOM_MMSS = 31, CUSTOM_H0MMSS = 32, CUSTOM_MMSS0 = 33, CUSTOM_000P0E_PLUS0 = 34, TEXT = 35}
		numFormat() => DllCall('libxl\xlFormatNumFormat', 'ptr', this, 'cdecl')

		; 设置为内置的数字格式标识符 {GENERAL = 0, NUMBER = 1, NUMBER_D2 = 2, NUMBER_SEP = 3, NUMBER_SEP_D2 = 4, CURRENCY_NEGBRA = 5, CURRENCY_NEGBRARED = 6, CURRENCY_D2_NEGBRA = 7, CURRENCY_D2_NEGBRARED = 8, PERCENT = 9, PERCENT_D2 = 10, SCIENTIFIC_D2 = 11, FRACTION_ONEDIG = 12, FRACTION_TWODIG = 13, DATE = 14, CUSTOM_D_MON_YY = 15, CUSTOM_D_MON = 16, CUSTOM_MON_YY = 17, CUSTOM_HMM_AM = 18, CUSTOM_HMMSS_AM = 19, CUSTOM_HMM = 20, CUSTOM_HMMSS = 21, CUSTOM_MDYYYY_HMM = 22, NUMBER_SEP_NEGBRA=37 = 23, NUMBER_SEP_NEGBRARED = 24, NUMBER_D2_SEP_NEGBRA = 25, NUMBER_D2_SEP_NEGBRARED = 26, ACCOUNT = 27, ACCOUNTCUR = 28, ACCOUNT_D2 = 29, ACCOUNT_D2_CUR = 30, CUSTOM_MMSS = 31, CUSTOM_H0MMSS = 32, CUSTOM_MMSS0 = 33, CUSTOM_000P0E_PLUS0 = 34, TEXT = 35}
		; 设置自定义号码格式标识符。使用Book.addCustomNumFormat()创建自定义格式。
		; https://www.libxl.com/format.html
		setNumFormat(numFormat) => DllCall('libxl\xlFormatSetNumFormat', 'ptr', this, 'int', numFormat, 'cdecl')
		
		; 返回水平对齐 AlignH {GENERAL = 0, LEFT = 1, CENTER = 2, RIGHT = 3, FILL = 4, JUSTIFY = 5, MERGE = 6, DISTRIBUTED = 7}
		alignH() => DllCall('libxl\xlFormatAlignH', 'ptr', this, 'cdecl')

		; 设置水平对齐 AlignH {GENERAL = 0, LEFT = 1, CENTER = 2, RIGHT = 3, FILL = 4, JUSTIFY = 5, MERGE = 6, DISTRIBUTED = 7}
		setAlignH(Align) => DllCall('libxl\xlFormatSetAlignH', 'ptr', this, 'int', align, 'cdecl')
		
		; 返回垂直对齐 AlignV {TOP = 0, CENTER = 1, BOTTOM = 2, JUSTIFY = 2, DISTRIBUTED = 3}
		alignV() => DllCall('libxl\xlFormatAlignV', 'ptr', this, 'cdecl')

		; 设置垂直对齐 AlignV {TOP = 0, CENTER = 1, BOTTOM = 2, JUSTIFY = 2, DISTRIBUTED = 3}
		setAlignV(Align) => DllCall('libxl\xlFormatSetAlignV', 'ptr', this, 'int', align, 'cdecl')

		; 返回是否自动换行
		wrap() => DllCall('libxl\xlFormatWrap', 'ptr', this, 'cdecl')

		; 设置自动换行
		setWrap(wrap := true) => DllCall('libxl\xlFormatSetWrap', 'ptr', this, 'int', wrap, 'cdecl')

		; 返回方向角度 0 - 90 : 文本逆时针旋转0到90度, 91 - 180 : 文本顺时针旋转1至90度, 255 : 垂直文本
		rotation() => DllCall('libxl\xlFormatRotation', 'ptr', this, 'cdecl')

		; 设置方向方向 0 - 90 : 文本逆时针旋转0到90度, 91 - 180 : 文本顺时针旋转1至90度, 255 : 垂直文本
		setRotation(rotation) => DllCall('libxl\xlFormatSetRotation', 'ptr', this, 'int', rotation, 'cdecl')

		; 返回文本的缩进级别。必须小于或等于15
		indent() => DllCall('libxl\xlFormatIndent', 'ptr', this, 'cdecl')

		; 设置文本的缩进级别。必须小于或等于15
		setIndent(indent) => DllCall('libxl\xlFormatSetIndent', 'ptr', this, 'int', indent, 'cdecl')

		; 返回是否设置了缩小字体填充
		shrinkToFit() => DllCall('libxl\xlFormatShrinkToFit', 'ptr', this, 'cdecl')

		; 设置缩小字体填充
		setShrinkToFit(shrinkToFit := true) => DllCall('libxl\xlFormatSetShrinkToFit', 'ptr', this, 'int', shrinkToFit, 'cdecl')

		; 设置边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		setBorder(Style := 1) => DllCall('libxl\xlFormatSetBorder', 'ptr', this, 'int', style, 'cdecl')
		
		; 设置边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setBorderColor(Color) => DllCall('libxl\xlFormatSetBorderColor', 'ptr', this, 'int', color, 'cdecl')

		; 返回左边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		borderLeft() => DllCall('libxl\xlFormatBorderLeft', 'ptr', this, 'cdecl')
		
		; 设置左边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		setBorderLeft(Style := 1) => DllCall('libxl\xlFormatSetBorderLeft', 'ptr', this, 'int', style, 'cdecl')
		
		; 返回右边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		borderRight() => DllCall('libxl\xlFormatBorderRight', 'ptr', this, 'cdecl')
		
		; 设置右边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		setBorderRight(style := 1) => DllCall('libxl\xlFormatSetBorderRight', 'ptr', this, 'int', style, 'cdecl')
		
		; 返回顶边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		borderTop() => DllCall('libxl\xlFormatBorderTop', 'ptr', this, 'cdecl')
		
		; 设置顶边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		setBorderTop(style := 1) => DllCall('libxl\xlFormatSetBorderTop', 'ptr', this, 'int', style, 'cdecl')
		
		; 返回底边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		borderBottom() => DllCall('libxl\xlFormatBorderBottom', 'ptr', this, 'cdecl')
		
		; 设置底边框样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		setBorderBottom(style := 1) => DllCall('libxl\xlFormatSetBorderBottom', 'ptr', this, 'int', style, 'cdecl')
		
		; 返回左边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		borderLeftColor() => DllCall('libxl\xlFormatBorderLeftColor', 'ptr', this, 'cdecl')
		
		; 设置左边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setBorderLeftColor(color) => DllCall('libxl\xlFormatSetBorderLeftColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回右边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		borderRightColor() => DllCall('libxl\xlFormatBorderRightColor', 'ptr', this, 'cdecl')
		
		; 设置右边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setBorderRightColor(color) => DllCall('libxl\xlFormatSetBorderRightColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回顶边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		borderTopColor() => DllCall('libxl\xlFormatBorderTopColor', 'ptr', this, 'cdecl')
		
		; 设置顶边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setBorderTopColor(color) => DllCall('libxl\xlFormatSetBorderTopColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回底边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		borderBottomColor() => DllCall('libxl\xlFormatBorderBottomColor', 'ptr', this, 'cdecl')
		
		; 设置底边框颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setBorderBottomColor(color) => DllCall('libxl\xlFormatSetBorderBottomColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回对角线方向 BorderDiagonal {NONE = 0, DOWN = 1, UP = 2, BOTH = 3}
		borderDiagonal() => DllCall('libxl\xlFormatBorderDiagonal', 'ptr', this, 'cdecl')

		; 设置对角线方向 BorderDiagonal {NONE = 0, DOWN = 1, UP = 2, BOTH = 3}
		setBorderDiagonal(Border) => DllCall('libxl\xlFormatSetBorderDiagonal', 'ptr', this, 'int', border, 'cdecl')
		
		; 返回对角线样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		borderDiagonalStyle() => DllCall('libxl\xlFormatBorderDiagonalStyle', 'ptr', this, 'cdecl')
		
		; 设置对角线样式 BorderStyle {NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUMDASHED, DASHDOT, MEDIUMDASHDOT, DASHDOTDOT, MEDIUMDASHDOTDOT, SLANTDASHDOT}
		setBorderDiagonalStyle(style) => DllCall('libxl\xlFormatSetBorderDiagonalStyle', 'ptr', this, 'int', style, 'cdecl')
		
		; 返回对角线颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		borderDiagonalColor() => DllCall('libxl\xlFormatBorderDiagonalColor', 'ptr', this, 'cdecl')
		
		; 设置对角线颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setBorderDiagonalColor(color) => DllCall('libxl\xlFormatSetBorderDiagonalColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回填充模式 FillPattern {NONE, SOLID, GRAY50, GRAY75, GRAY25, HORSTRIPE, VERSTRIPE, REVDIAGSTRIPE, DIAGSTRIPE, DIAGCROSSHATCH, THICKDIAGCROSSHATCH, THINHORSTRIPE, THINVERSTRIPE, THINREVDIAGSTRIPE, THINDIAGSTRIPE, THINHORCROSSHATCH, THINDIAGCROSSHATCH, GRAY12P5, GRAY6P25}
		fillPattern() => DllCall('libxl\xlFormatFillPattern', 'ptr', this, 'cdecl')
		
		; 设置填充模式 FillPattern {NONE, SOLID, GRAY50, GRAY75, GRAY25, HORSTRIPE, VERSTRIPE, REVDIAGSTRIPE, DIAGSTRIPE, DIAGCROSSHATCH, THICKDIAGCROSSHATCH, THINHORSTRIPE, THINVERSTRIPE, THINREVDIAGSTRIPE, THINDIAGSTRIPE, THINHORCROSSHATCH, THINDIAGCROSSHATCH, GRAY12P5, GRAY6P25}
		setFillPattern(Pattern) => DllCall('libxl\xlFormatSetFillPattern', 'ptr', this, 'int', pattern, 'cdecl')
		
		; 返回填充颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		patternForegroundColor() => DllCall('libxl\xlFormatPatternForegroundColor', 'ptr', this, 'cdecl')
		
		; 设置填充颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setPatternForegroundColor(color) => DllCall('libxl\xlFormatSetPatternForegroundColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回填充背景颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		patternBackgroundColor() => DllCall('libxl\xlFormatPatternBackgroundColor', 'ptr', this, 'cdecl')
		
		; 设置填充背景颜色 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setPatternBackgroundColor(color) => DllCall('libxl\xlFormatSetPatternBackgroundColor', 'ptr', this, 'int', color, 'cdecl')
		
		; 返回“锁定保护”属性。为真返回1；为假返回0。
		locked() => DllCall('libxl\xlFormatLocked', 'ptr', this, 'cdecl')
		
		; 设置“锁定保护”属性。locked := true锁定, locked := false解锁。
		setLocked(locked := true) => DllCall('libxl\xlFormatSetLocked', 'ptr', this, 'int', locked, 'cdecl')
		
		; 返回“隐藏保护属性”属性。为真返回1；为假返回0。
		hidden() => DllCall('libxl\xlFormatHidden', 'ptr', this, 'cdecl')
		
		; 返回“隐藏保护属性”属性。hidden := true隐藏, hidden := false显现。
		setHidden(hidden := true) => DllCall('libxl\xlFormatSetHidden', 'ptr', this, 'int', hidden, 'cdecl')
	}
	class IRichString extends XL.IBase {
		; 为富文本字符串 RichString 添加一个新的字体，以便使用 RichString.addText() 方法，初始参数可以从其他字体中复制。
		addFont(initFont := 0) => XL.IFont(DllCall('libxl\xlRichStringAddFont', 'ptr', this, 'ptr', initFont, 'cdecl ptr'))

		; 使用指定字体为富文本字符串 RichString 添加一段文本（run），以便在同一单元格中混合不同的字体。
		addText(text, font := 0) => DllCall('libxl\xlRichStringAddText', 'ptr', this, 'str', text, 'ptr', font, 'cdecl')

		; 返回富文本字符串 RichString 中指定索引的文本（run）和字体。
		getText(index, &font := 0) => DllCall('libxl\xlRichStringGetText', 'ptr', this, 'int', index-1, 'ptr*', &font := 0, 'cdecl str')

		; 返回富文本字符串 RichString 中的文本（run）的数量。
		textSize() => DllCall('libxl\xlRichStringTextSize', 'ptr', this, 'cdecl')
	}
	class ISheet extends XL.IBase {
		; 返回单元格的类型 CellType {0:EMPTY, 1:NUMBER, 2:STRING, 3:BOOLEAN, 4:BLANK, ERROR, STRICTDATE}
		cellType(row, col) => DllCall('libxl\xlSheetCellType', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl')

		; 返回单元格是否包含公式。包含公式返回1；不包含公式返回0。
		isFormula(row, col) => DllCall('libxl\xlSheetIsFormula', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl')

		; 返回单元格的格式。这个格式可以由用户更改。
		cellFormat(row, col) => XL.IFormat(DllCall('libxl\xlSheetCellFormat', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl ptr'))

		; 设置单元格的格式。
		setCellFormat(row, col, format) => DllCall('libxl\xlSheetSetCellFormat', 'ptr', this, 'int', row-1, 'int', col-1, 'ptr', format, 'cdecl')
		
		; 在单元格中读取字符串
		readStr(row, col, &format := 0) {
			ret := DllCall('libxl\xlSheetReadStr', 'ptr', this, 'int', row-1, 'int', col-1, 'ptr*', &format := 0, 'cdecl str')
			if (!format)
				throw Error(this.parent.errorMessage())
			return (format := XL.IFormat(format), ret)
		}

		; 在单元格中写入字符串
		writeStr(row, col, value, format := 0) => DllCall('libxl\xlSheetWriteStr', 'ptr', this, 'int', row-1, 'int', col-1, 'str', value, 'ptr', format, 'cdecl')
		
		; 在单元格中读取富文本字符串
		readRichStr(row, col, &format := 0) {
			ret := XL.IRichString(DllCall('libxl\xlSheetReadRichStr', 'ptr', this, 'int', row-1, 'int', col-1, 'ptr*', &format := 0, 'cdecl ptr'))
			if (!format)
				throw Error(this.parent.errorMessage())
			return (format := XL.IFormat(format), ret)
		}

		; 在单元格中写入富文本字符串
		writeRichStr(row, col, richString, format := 0) => DllCall('libxl\xlSheetWriteRichStr', 'ptr', this, 'int', row-1, 'int', col-1, 'ptr', richString, 'ptr', format, 'cdecl')
		
		; 在单元格中读取数字
		readNum(row, col, &format := 0) {
			ret := DllCall('libxl\xlSheetReadNum', 'ptr', this, 'int', row-1, 'int', col-1, 'ptr*', &format := 0, 'cdecl double')
			if (!format)
				throw Error(this.parent.errorMessage())
			return (format := XL.IFormat(format), ret)
		}

		; 在单元格中写入数字
		writeNum(row, col, value, format := 0) => DllCall('libxl\xlSheetWriteNum', 'ptr', this, 'int', row-1, 'int', col-1, 'double', value, 'ptr', format, 'cdecl')
		
		; 在单元格内读取布尔值
		readBool(row, col, &format := 0) {
			ret := DllCall('libxl\xlSheetReadBool', 'ptr', this, 'int', row-1, 'int', col-1, 'ptr*', &format := 0, 'cdecl')
			if (!format)
				throw Error(this.parent.errorMessage())
			return (format := XL.IFormat(format), ret)
		}

		; 在单元格中写入布尔值
		writeBool(row, col, value, format := 0) => DllCall('libxl\xlSheetWriteBool', 'ptr', this, 'int', row-1, 'int', col-1, 'int', value, 'ptr', format, 'cdecl')
		
		; 读取空白单元格
		readBlank(row, col, &format := 0) {
			ret := DllCall('libxl\xlSheetReadBlank', 'ptr', this, 'int', row-1, 'int', col-1, 'ptr*', &format := 0, 'cdecl')
			if (!format)
				throw Error(this.parent.errorMessage())
			return (format := XL.IFormat(format), ret)
		}

		; 写入空白单元格
		writeBlank(row, col, format) => DllCall('libxl\xlSheetWriteBlank', 'ptr', this, 'int', row-1, 'int', col-1, 'ptr', format, 'cdecl')

		; 在单元格内读取公式
		readFormula(row, col, &format := unset) {
			ret := DllCall('libxl\xlSheetReadFormula', 'ptr', this, 'int', row-1, 'int', col-1, 'ptr*', &format := 0, 'cdecl str')
			if (!format)
				throw Error(this.parent.errorMessage())
			return (format := XL.IFormat(format), ret)
		}

		; 将公式写入指定格式的单元格。如果format等于0，则忽略format。
		writeFormula(row, col, expr, format := 0) => DllCall('libxl\xlSheetWriteFormula', 'ptr', this, 'int', row-1, 'int', col-1, 'str', expr, 'ptr', format, 'cdecl')
		
		; 将具有预先计算的双精度值的公式表达式写入指定格式的单元格中。如果format等于0，则忽略format。如果发生错误，返回0。
		writeFormulaNum(row, col, expr, value, format := 0) => DllCall('libxl\xlSheetWriteFormulaNum', 'ptr', this, 'int', row-1, 'int', col-1, 'str', expr, 'double', value, 'ptr', format, 'cdecl')
		
		; 将具有预先计算的字符串值的公式表达式写入具有指定格式的单元格。如果format等于0，则忽略format。如果发生错误，返回0。
		writeFormulaStr(row, col, expr, value, format := 0) => DllCall('libxl\xlSheetWriteFormulaStr', 'ptr', this, 'int', row-1, 'int', col-1, 'str', expr, 'str', value, 'ptr', format, 'cdecl')
		
		; 将具有预先计算的bool值的公式表达式写入具有指定格式的单元格。如果format等于0，则忽略format。如果发生错误，返回0。
		writeFormulaBool(row, col, expr, value, format := 0) => DllCall('libxl\xlSheetWriteFormulaBool', 'ptr', this, 'int', row-1, 'int', col-1, 'str', expr, 'int', value, 'ptr', format, 'cdecl')
		
		; 从指定单元格读取注释（仅适用于xls格式）。
		readComment(row, col) => DllCall('libxl\xlSheetReadComment', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl str')

		/**
		 * 向单元格写入注释（仅适用于xls格式）。
		 * @param {Integer} row     行号
		 * @param {Integer} col     列号
		 * @param {String}  value   注释内容
		 * @param {Integer} author  作者
		 * @param {Integer} width   宽度（像素）
		 * @param {Integer} height  高度（像素）
		 * @returns {Float | Integer | String} 
		 */
		writeComment(row, col, value, author := 0, width := 129, height := 75) => DllCall('libxl\xlSheetWriteComment', 'ptr', this, 'int', row-1, 'int', col-1, 'str', value, 'str', author, 'int', width, 'int', height, 'cdecl')
		
		; 从单元格中删除注释（仅适用于xls格式）。
		removeComment(row, col) => DllCall('libxl\xlSheetRemoveComment', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl')
		
		; 检查单元格是否包含日期或时间值。如果返回值为真，使用Sheet.readNum（）方法读取它，并使用Book.dateUnpack（）方法解包它。
		isDate(row, col) => DllCall('libxl\xlSheetIsDate', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl')

		; 检查单元格是否包含具有多种字体的富字符串。如果返回值为真，用Sheet.readRichStr（）方法读取它。
		isRichStr(row, col) => DllCall('libxl\xlSheetIsRichStr', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl')

		; 从单元格读取错误。
		; ErrorType {NULL = 0x0, DIV_0 = 0x7, VALUE = 0x0F, REF = 0x17, NAME = 0x1D, NUM = 0x24, NA = 0x2A, NOERROR = 0xFF}
		readError(row, col) => DllCall('libxl\xlSheetReadError', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl')

		; 将错误写入指定格式的单元格。如果format等于0，则忽略format。
		; ErrorType {NULL = 0x0, DIV_0 = 0x7, VALUE = 0x0F, REF = 0x17, NAME = 0x1D, NUM = 0x24, NA = 0x2A, NOERROR = 0xFF}
		writeError(row, col, ErrorType, format := 0) => DllCall('libxl\xlSheetWriteError', 'ptr', this, 'int', row-1, 'int', col-1, 'int', ErrorType, 'ptr', format, 'cdecl')
		
		; 返回列宽度。列宽是以数字0、1、2、…的最大数字宽度的字符数来衡量的。以正常样式的字体呈现。
		colWidth(col) => DllCall('libxl\xlSheetColWidth', 'ptr', this, 'int', col-1, 'cdecl double')

		; 以排版点为单位返回行高度。点是1/72英寸。
		rowHeight(row) => DllCall('libxl\xlSheetRowHeight', 'ptr', this, 'int', row-1, 'cdecl double')

		; 返回以像素为单位的列宽度。
		colWidthPx(col) => DllCall('libxl\xlSheetColWidthPx', 'ptr', this, 'int', col-1, 'cdecl')

		; 以像素为单位返回行高度。
		rowHeightPx(row) => DllCall('libxl\xlSheetRowHeightPx', 'ptr', this, 'int', row-1, 'cdecl')

		; 设置从colFirst到colLast的所有列的列宽度和格式。列宽度测量为数字0、1、2、…的最大数字宽度的字符数。，以正常样式的字体呈现。值-1用于自动拟合列宽度。如果format等于0，则忽略format。列可能是隐藏的。如果发生错误，返回0。
		setCol(colFirst, colLast, width, format := 0, hidden := false) => DllCall('libxl\xlSheetSetCol', 'ptr', this, 'int', colFirst-1, 'int', colLast-1, 'double', width, 'ptr', format, 'int', hidden, 'cdecl')
		
		; 设置从colFirst到colLast所有列的列宽度（以像素为单位）和格式。值-1用于自动拟合列宽度。如果format等于0，则忽略format。列可能是隐藏的。如果发生错误，返回0。
		setColPx(colFirst, colLast, widthPx, format := 0, hidden := false) => DllCall('libxl\xlSheetSetColPx', 'ptr', this, 'int', colFirst-1, 'int', colLast-1, 'int', widthPx, 'ptr', format, 'int', hidden, 'cdecl')
		
		; 设置行高度和格式。以点大小测量的行高度。如果format等于0，则忽略format。行可以隐藏。如果发生错误，返回false。
		setRow(row, height, format := 0, hidden := false) => DllCall('libxl\xlSheetSetRow', 'ptr', this, 'int', row-1, 'double', height, 'ptr', format, 'int', hidden, 'cdecl')
		
		; 以像素为单位设置行高度。如果format等于0，则忽略format。行可以隐藏。如果发生错误，返回false。
		setRowPx(row, heightPx, format := 0, hidden := false) => DllCall('libxl\xlSheetSetRowPx', 'ptr', this, 'int', row-1, 'int', heightPx, 'ptr', format, 'int', hidden, 'cdecl')
		
		; 返回行是否隐藏。
		rowHidden(row) => DllCall('libxl\xlSheetRowHidden', 'ptr', this, 'int', row-1, 'cdecl')

		; 隐藏行
		setRowHidden(row, hidden := true) => DllCall('libxl\xlSheetSetRowHidden', 'ptr', this, 'int', row-1, 'int', hidden, 'cdecl')

		; 返回列是否隐藏。
		colHidden(col) => DllCall('libxl\xlSheetColHidden', 'ptr', this, 'int', col-1, 'cdecl')

		; 隐藏列
		setColHidden(col, hidden := true) => DllCall('libxl\xlSheetSetColHidden', 'ptr', this, 'int', col-1, 'int', hidden, 'cdecl')

		; 返回以点大小测量的默认行高。
		defaultRowHeight() => DllCall('libxl\xlSheetDefaultRowHeight', 'ptr', this, 'cdecl double')

		; 设置以点大小测量的默认行高。
		setDefaultRowHeight(height) => DllCall('libxl\xlSheetSetDefaultRowHeight', 'ptr', this, 'double', height, 'cdecl double')
		
		; 获取单元格的合并单元格范围。结果写入rowFirst， rowLast, colFirst, colLast。如果指定的单元格在合并区域中，则返回1，否则返回0。
		getMerge(row, col, &rowFirst := -1, &rowLast := -1, &colFirst := -1, &colLast := -1) => (res := DllCall('libxl\xlSheetGetMerge', 'ptr', this, 'int', row-1, 'int', col-1, 'int*', &rowFirst := -1, 'int*', &rowLast := -1, 'int*', &colFirst := -1, 'int*', &colLast := -1, 'cdecl'), rowFirst++, rowLast++, colFirst++, colLast++, res)
		
		; 合并范围内的单元格：rowFirst - rowLast， colFirst - colLast。如果发生错误，返回0。
		setMerge(rowFirst, rowLast, colFirst, colLast) => DllCall('libxl\xlSheetSetMerge', 'ptr', this, 'int', rowFirst-1, 'int', rowLast-1, 'int', colFirst-1, 'int', colLast-1, 'cdecl')
		
		; 删除合并的单元格。如果发生错误，返回0。
		delMerge(row, col) => DllCall('libxl\xlSheetDelMerge', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl')

		; 返回此工作表中合并的单元格的数量。
		mergeSize() => DllCall('libxl\xlSheetMergeSize', 'ptr', this, 'cdecl')

		; 按索引获取合并的单元格范围。
		merge(index, &rowFirst, &rowLast, &colFirst, &colLast) => (res := DllCall('libxl\xlSheetMerge', 'ptr', this, 'int', index-1, 'int*', &rowFirst := 0, 'int*', &rowLast := 0, 'int*', &colFirst := 0, 'int*', &colLast := 0, 'cdecl'), rowFirst++, rowLast++, colFirst++, colLast++, res)
		
		; 按索引删除合并的单元格。
		delMergeByIndex(index) => DllCall('libxl\xlSheetDelMergeByIndex', 'ptr', this, 'int', index-1, 'cdecl')

		; 返回此工作表中图片的数量。
		pictureSize() => DllCall('libxl\xlSheetPictureSize', 'ptr', this, 'cdecl')

		/**
		 * 获取工作表第index个图片的信息, 使用xlBookGetPicture()通过工作簿图片索引提取图片的二进制数据。如果发生错误，返回-1。
		 * @param index                 图片索引
		 * @param {Integer} rowTop      图片上边缘的行号
		 * @param {Integer} colLeft     图片左边缘的列号
		 * @param {Integer} rowBottom   图片下边缘的行号
		 * @param {Integer} colRight    图片右边缘的列号
		 * @param {Integer} width       图片的宽度(像素)
		 * @param {Integer} height      图片的高度(像素)
		 * @param {Integer} offset_x    图像水平偏移量(像素)
		 * @param {Integer} offset_y    图像垂直偏移量(像素)
		 * @returns {Integer}           如果发生错误，返回-1
		 */
		getPicture(index, &rowTop := 0, &colLeft := 0, &rowBottom := 0, &colRight := 0, &width := 0, &height := 0, &offset_x := 0, &offset_y := 0) => (res := DllCall('libxl\xlSheetGetPicture', 'ptr', this, 'int', index-1, 'int*', &rowTop := 0, 'int*', &colLeft := 0, 'int*', &rowBottom := 0, 'int*', &colRight := 0, 'int*', &width := 0, 'int*', &height := 0, 'int*', &offset_x := 0, 'int*', &offset_y := 0, 'cdecl'), rowTop++, colLeft++, rowBottom++, colRight++, res)
		
		; 按指定索引删除图片。如果发生错误，返回0。
		removePictureByIndex(index) => DllCall('libxl\xlSheetRemovePictureByIndex', 'ptr', this, 'int', index-1, 'cdecl')
		
		/**
		 * 设置具有pictureId标识符的图片，其位置为行和col，并具有比例因子和以像素为单位的偏移量。使用Book.addPicture()添加新图片并获取标识符。
		 * 图片可以对齐到单元格的左上角，单元格的中心或在单元格内拉伸：
		 * Scale > 0 - 图片以指定的缩放比例与单元格或合并区域的左上角对齐
		 * Scale = 0 - 图片在指定的单元格或合并区域内拉伸
		 * Scale < 0 - 图片以指定的缩放比例对齐单元格或合并区域的中心
		 * @param row                     行号
		 * @param col                     列号
		 * @param pictureId               图片标识符(one-based)
		 * @param {Float} scale           缩放比例
		 * @param {Integer} offset_x      偏移量（像素）
		 * @param {Integer} offset_y      偏移量（像素）
		 * @param {Integer} pos           设定在调整行和列的大小时如何移动或调整图片的大小 Position {MOVE_AND_SIZE, ONLY_MOVE, ABSOLUTE}
		 * @returns {Float | Integer | String} 
		 */
		setPicture(row, col, pictureId, scale := 1.0, offset_x := 0, offset_y := 0, pos := 0) => DllCall('libxl\xlSheetSetPicture', 'ptr', this, 'int', row-1, 'int', col-1, 'int', pictureId-1, 'double', scale, 'int', offset_x, 'int', offset_y, 'int', pos, 'cdecl')
		
		; 设置具有pictureId标识符的图片，位置为行和列，并以像素为单位自定义大小和偏移量。使用Book.addPicture()添加新图片并获取标识符。
		setPicture2(row, col, pictureId, width := -1, height := -1, offset_x := 0, offset_y := 0, pos := 0) => DllCall('libxl\xlSheetSetPicture2', 'ptr', this, 'int', row-1, 'int', col-1, 'int', pictureId-1, 'int', width, 'int', height, 'int', offset_x, 'int', offset_y, 'int', pos, 'cdecl')
		
		; 移除指定位置上的图片。如果发生错误，返回0。
		removePicture(row, col) => DllCall('libxl\xlSheetRemovePicture', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl')

		; 返回第index个有水平分页符的行。
		getHorPageBreak(index) => DllCall('libxl\xlSheetGetHorPageBreak', 'ptr', this, 'int', index-1, 'cdecl')

		; 返回有水平分页符的行的数量。
		getHorPageBreakSize() => DllCall('libxl\xlSheetGetHorPageBreakSize', 'ptr', this, 'cdecl')

		; 返回第index个有垂直分页符的列。
		getVerPageBreak(index) => DllCall('libxl\xlSheetGetVerPageBreak', 'ptr', this, 'int', index-1, 'cdecl')

		; 返回有垂直分页符的列的数量。
		getVerPageBreakSize() => DllCall('libxl\xlSheetGetVerPageBreakSize', 'ptr', this, 'cdecl')

		; 设置/删除一个水平分页符（如果pageBreak := 1设置，如果pageBreak := 0删除）。如果发生错误返回0。
		setHorPageBreak(row, pageBreak := true) => DllCall('libxl\xlSheetSetHorPageBreak', 'ptr', this, 'int', row-1, 'int', pageBreak, 'cdecl')
		
		; 设置/删除一个垂直分页符（如果pageBreak := 1设置，如果pageBreak := 0删除）。如果发生错误返回0。
		setVerPageBreak(col, pageBreak := true) => DllCall('libxl\xlSheetSetVerPageBreak', 'ptr', this, 'int', col-1, 'int', pageBreak, 'cdecl')
		
		; 冻结单元格左侧和上边的窗格。此函数允许在顶部位置冻结标题或在右侧冻结某些列。
		split(row, col) => DllCall('libxl\xlSheetSplit', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl')
		
		; 获取冻结窗格位置, 返回左上角未被冻结单元格的行和列。
		splitInfo(&row, &col) => (res := DllCall('libxl\xlSheetSplitInfo', 'ptr', this, 'int*', &row := 0, 'int*', &col := 0, 'cdecl'), row++, col++, res)
		
		; 分组从rowFirst到rowLast的行。如果发生错误，返回0。
		groupRows(rowFirst, rowLast, collapsed := true) => DllCall('libxl\xlSheetGroupRows', 'ptr', this, 'int', rowFirst-1, 'int', rowLast-1, 'int', collapsed, 'cdecl')
		
		; 分组从colFirst到colLast的列。如果发生错误，返回0。
		groupCols(colFirst, colLast, collapsed := true) => DllCall('libxl\xlSheetGroupCols', 'ptr', this, 'int', colFirst-1, 'int', colLast-1, 'int', collapsed, 'cdecl')
		
		; 返回分组行摘要是否在下面。如果摘要在下面则返回1，如果不在下面则返回0。
		groupSummaryBelow() => DllCall('libxl\xlSheetGroupSummaryBelow', 'ptr', this, 'cdecl')

		; 设置分组行摘要flag：1 -下面，0 -上面。
		setGroupSummaryBelow(below) => DllCall('libxl\xlSheetSetGroupSummaryBelow', 'ptr', this, 'int', below, 'cdecl')

		; 返回分组列摘要是否在右边。在右边返回1，在左边返回0。
		groupSummaryRight() => DllCall('libxl\xlSheetGroupSummaryRight', 'ptr', this, 'cdecl')

		; 设置分组列摘要的flag：1 -右，0 -左。
		setGroupSummaryRight(right) => DllCall('libxl\xlSheetSetGroupSummaryRight', 'ptr', this, 'int', right, 'cdecl')

		; 清除范围内的单元格。
		clear(rowFirst := 1, rowLast := 1048576, colFirst := 1, colLast := 16384) => DllCall('libxl\xlSheetClear', 'ptr', this, 'int', rowFirst-1, 'int', rowLast-1, 'int', colFirst-1, 'int', colLast-1, 'cdecl')
		
		; 从colFirst到rowLast插入列。更新现有的命名范围。如果发生错误，返回0。
		insertCol(colFirst, colLast, updateNamedRanges := true) => DllCall('libxl\xlSheetInsertCol', 'ptr', this, 'int', colFirst-1, 'int', colLast-1, 'cdecl')
		
		; 从rowFirst到rowLast插入行。更新现有的命名范围。如果发生错误，返回0。
		insertRow(rowFirst, rowLast, updateNamedRanges := true) => DllCall('libxl\xlSheetInsertRow', 'ptr', this, 'int', rowFirst-1, 'int', rowLast-1, 'cdecl')
		
		; 删除从colFirst到colLast的列。更新现有的命名范围。如果发生错误，返回0。
		removeCol(colFirst, colLast, updateNamedRanges := true) => DllCall('libxl\xlSheetRemoveCol', 'ptr', this, 'int', colFirst-1, 'int', colLast-1, 'cdecl')
		
		; 删除从rowFirst到rowLast的行。更新现有的命名范围。如果发生错误，返回0。
		removeRow(rowFirst, rowLast, updateNamedRanges := true) => DllCall('libxl\xlSheetRemoveRow', 'ptr', this, 'int', rowFirst-1, 'int', rowLast-1, 'cdecl')
		
		; 从colFirst到rowLast插入列。不更新现有的已命名范围。如果发生错误，返回0。
		insertColAndKeepRanges(colFirst, colLast) => DllCall('libxl\xlSheetInsertColAndKeepRanges', 'ptr', this, 'int', colFirst-1, 'int', colLast-1, 'cdecl')
		
		; 从rowFirst到rowLast插入行。不更新现有的已命名范围。如果发生错误，返回0。
		insertRowAndKeepRanges(rowFirst, rowLast) => DllCall('libxl\xlSheetInsertRowAndKeepRanges', 'ptr', this, 'int', rowFirst-1, 'int', rowLast-1, 'cdecl')
		
		; 删除从colFirst到colLast的列。不更新现有的已命名范围。如果发生错误，返回0。
		removeColAndKeepRanges(colFirst, colLast) => DllCall('libxl\xlSheetRemoveColAndKeepRanges', 'ptr', this, 'int', colFirst-1, 'int', colLast-1, 'cdecl')
		
		; 删除从rowFirst到rowLast的行。不更新现有的已命名范围。如果发生错误，返回0。
		removeRowAndKeepRanges(rowFirst, rowLast) => DllCall('libxl\xlSheetRemoveRowAndKeepRanges', 'ptr', this, 'int', rowFirst-1, 'int', rowLast-1, 'cdecl')
		
		; 将单元格(rowSrc, colSrc)复制到单元格(rowDst, colDst)。如果发生错误，返回0。
		copyCell(rowSrc, colSrc, rowDst, colDst) => DllCall('libxl\xlSheetCopyCell', 'ptr', this, 'int', rowSrc-1, 'int', colSrc-1, 'int', rowDst-1, 'int', colDst-1, 'cdecl')
		
		; 返回包含已使用单元格（包括仅具有格式的空白单元格）的工作表中第一行的索引(one-based)。
		firstRow() => DllCall('libxl\xlSheetFirstRow', 'ptr', this, 'cdecl') + 1
		
		; 返回包含已使用单元格（包括仅具有格式的空白单元格）的工作表中最后一行的索引(one-based)。
		lastRow() => DllCall('libxl\xlSheetLastRow', 'ptr', this, 'cdecl') + 1
		
		; 返回包含已使用单元格（包括仅具有格式的空白单元格）的工作表中第一列的索引(one-based)。
		firstCol() => DllCall('libxl\xlSheetFirstCol', 'ptr', this, 'cdecl') + 1
		
		; 返回包含已使用单元格（包括仅具有格式的空白单元格）的工作表中最后一列的索引(one-based)。
		lastCol() => DllCall('libxl\xlSheetLastCol', 'ptr', this, 'cdecl') + 1
		
		; 返回包含带值的单元格的工作表第一行的索引。忽略具有格式的空白单元格(one-based)。
		firstFilledRow() => DllCall('libxl\xlSheetFirstFilledRow', 'ptr', this, 'cdecl') + 1
		
		; 返回包含带值的单元格的最后一行的索引。忽略具有格式的空白单元格(one-based)。
		lastFilledRow() => DllCall('libxl\xlSheetLastFilledRow', 'ptr', this, 'cdecl') ; 这里不用+1, 原函数是返回最后一行后面一行
		
		; 返回包含带值的单元格的工作表第一列的索引。忽略具有格式的空白单元格(one-based)。
		firstFilledCol() => DllCall('libxl\xlSheetFirstFilledCol', 'ptr', this, 'cdecl') + 1
		
		; 返回包含带值的单元格的最后一列的索引。忽略具有格式的空白单元格(one-based)。
		lastFilledCol() => DllCall('libxl\xlSheetLastFilledCol', 'ptr', this, 'cdecl') ; 这里不用+1, 原函数是返回最后一行后面一行
		
		; 返回是否显示网格线。如果显示网格线则返回1，如果不显示则返回0。
		displayGridlines() => DllCall('libxl\xlSheetDisplayGridlines', 'ptr', this, 'cdecl')
		
		; 设置要显示的网格线，1表示显示网格线，0表示不显示网格线。
		setDisplayGridlines(show := true) => DllCall('libxl\xlSheetSetDisplayGridlines', 'ptr', this, 'int', show, 'cdecl')

		; 返回是否打印网格线。如果网格线被打印则返回1，如果没有打印则返回0。
		printGridlines() => DllCall('libxl\xlSheetPrintGridlines', 'ptr', this, 'cdecl')

		; 设置要打印的网格线，1 - 打印网格线，0 - 不打印网格线。
		setPrintGridlines(print := true) => DllCall('libxl\xlSheetSetPrintGridlines', 'ptr', this, 'int', print, 'cdecl')

		; 以百分比形式返回当前视图的缩放比例。标准为100
		zoom() => DllCall('libxl\xlSheetZoom', 'ptr', this, 'cdecl')

		; 设置当前视图的缩放比例。标准为100
		setZoom(zoom) => DllCall('libxl\xlSheetSetZoom', 'ptr', this, 'int', zoom, 'cdecl')

		; 以百分比形式返回打印的缩放比例。标准为100
		printZoom() => DllCall('libxl\xlSheetPrintZoom', 'ptr', this, 'cdecl')

		; 设置打印的缩放比例。标准为100
		setPrintZoom(zoom) => DllCall('libxl\xlSheetSetPrintZoom', 'ptr', this, 'int', zoom, 'cdecl')

		; 返回"页面布局-调整为合适大小"的选项。wPages-宽度适合的页数；页数-高度适合的页数。
		getPrintFit(&wPages, &hPages) => DllCall('libxl\xlSheetGetPrintFit', 'ptr', this, 'int*', &wPages := 0, 'int*', &hPages := 0, 'cdecl')
		
		; 设置"页面布局-调整为合适大小"的选项。wPages-宽度适合的页数；页数-高度适合的页数。
		setPrintFit(wPages := 1, hPages := 1) => DllCall('libxl\xlSheetSetPrintFit', 'ptr', this, 'int', wPages, 'int', hPages, 'cdecl')
		
		; 返回"页面布局-页面设置"的纸张方向，1 - 横向，0 - 纵向。
		landscape() => DllCall('libxl\xlSheetLandscape', 'ptr', this, 'cdecl')

		; 设置"页面布局-页面设置"的纸张方向，1 - 横向，0 - 纵向。
		setLandscape(landscape := true) => DllCall('libxl\xlSheetSetLandscape', 'ptr', this, 'int', landscape, 'cdecl')

		; 返回纸张大小。Paper {DEFAULT, LETTER, LETTERSMALL, TABLOID, LEDGER, LEGAL, STATEMENT, EXECUTIVE, A3, A4, A4SMALL, A5, B4, B5, FOLIO, QUATRO, 10x14, 10x17, NOTE, ENVELOPE_9, ENVELOPE_10, ENVELOPE_11, ENVELOPE_12, ENVELOPE_14, C_SIZE, D_SIZE, E_SIZE, ENVELOPE_DL, ENVELOPE_C5, ENVELOPE_C3, ENVELOPE_C4, ENVELOPE_C6, ENVELOPE_C65, ENVELOPE_B4, ENVELOPE_B5, ENVELOPE_B6, ENVELOPE, ENVELOPE_MONARCH, US_ENVELOPE, FANFOLD, GERMAN_STD_FANFOLD, GERMAN_LEGAL_FANFOLD, B4_ISO, JAPANESE_POSTCARD, 9x11, 10x11, 15x11, ENVELOPE_INVITE, US_LETTER_EXTRA = 50, US_LEGAL_EXTRA, US_TABLOID_EXTRA, A4_EXTRA, LETTER_TRANSVERSE, A4_TRANSVERSE, LETTER_EXTRA_TRANSVERSE, SUPERA, SUPERB, US_LETTER_PLUS, A4_PLUS, A5_TRANSVERSE, B5_TRANSVERSE, A3_EXTRA, A5_EXTRA, B5_EXTRA, A2, A3_TRANSVERSE, A3_EXTRA_TRANSVERSE, JAPANESE_DOUBLE_POSTCARD, A6, JAPANESE_ENVELOPE_KAKU2, JAPANESE_ENVELOPE_KAKU3, JAPANESE_ENVELOPE_CHOU3, JAPANESE_ENVELOPE_CHOU4, LETTER_ROTATED, A3_ROTATED, A4_ROTATED, A5_ROTATED, B4_ROTATED, B5_ROTATED, JAPANESE_POSTCARD_ROTATED, DOUBLE_JAPANESE_POSTCARD_ROTATED, A6_ROTATED, JAPANESE_ENVELOPE_KAKU2_ROTATED, JAPANESE_ENVELOPE_KAKU3_ROTATED, JAPANESE_ENVELOPE_CHOU3_ROTATED, JAPANESE_ENVELOPE_CHOU4_ROTATED, B6, B6_ROTATED, 12x11, JAPANESE_ENVELOPE_YOU4, JAPANESE_ENVELOPE_YOU4_ROTATED, PRC16K, PRC32K, PRC32K_BIG, PRC_ENVELOPE1, PRC_ENVELOPE2, PRC_ENVELOPE3, PRC_ENVELOPE4, PRC_ENVELOPE5, PRC_ENVELOPE6, PRC_ENVELOPE7, PRC_ENVELOPE8, PRC_ENVELOPE9, PRC_ENVELOPE10, PRC16K_ROTATED, PRC32K_ROTATED, PRC32KBIG_ROTATED, PRC_ENVELOPE1_ROTATED, PRC_ENVELOPE2_ROTATED, PRC_ENVELOPE3_ROTATED, PRC_ENVELOPE4_ROTATED, PRC_ENVELOPE5_ROTATED, PRC_ENVELOPE6_ROTATED, PRC_ENVELOPE7_ROTATED, PRC_ENVELOPE8_ROTATED, PRC_ENVELOPE9_ROTATED, PRC_ENVELOPE10_ROTATED}
		paper() => DllCall('libxl\xlSheetPaper', 'ptr', this, 'cdecl')

		; 设置纸张大小。Paper {DEFAULT, LETTER, LETTERSMALL, TABLOID, LEDGER, LEGAL, STATEMENT, EXECUTIVE, A3, A4, A4SMALL, A5, B4, B5, FOLIO, QUATRO, 10x14, 10x17, NOTE, ENVELOPE_9, ENVELOPE_10, ENVELOPE_11, ENVELOPE_12, ENVELOPE_14, C_SIZE, D_SIZE, E_SIZE, ENVELOPE_DL, ENVELOPE_C5, ENVELOPE_C3, ENVELOPE_C4, ENVELOPE_C6, ENVELOPE_C65, ENVELOPE_B4, ENVELOPE_B5, ENVELOPE_B6, ENVELOPE, ENVELOPE_MONARCH, US_ENVELOPE, FANFOLD, GERMAN_STD_FANFOLD, GERMAN_LEGAL_FANFOLD, B4_ISO, JAPANESE_POSTCARD, 9x11, 10x11, 15x11, ENVELOPE_INVITE, US_LETTER_EXTRA = 50, US_LEGAL_EXTRA, US_TABLOID_EXTRA, A4_EXTRA, LETTER_TRANSVERSE, A4_TRANSVERSE, LETTER_EXTRA_TRANSVERSE, SUPERA, SUPERB, US_LETTER_PLUS, A4_PLUS, A5_TRANSVERSE, B5_TRANSVERSE, A3_EXTRA, A5_EXTRA, B5_EXTRA, A2, A3_TRANSVERSE, A3_EXTRA_TRANSVERSE, JAPANESE_DOUBLE_POSTCARD, A6, JAPANESE_ENVELOPE_KAKU2, JAPANESE_ENVELOPE_KAKU3, JAPANESE_ENVELOPE_CHOU3, JAPANESE_ENVELOPE_CHOU4, LETTER_ROTATED, A3_ROTATED, A4_ROTATED, A5_ROTATED, B4_ROTATED, B5_ROTATED, JAPANESE_POSTCARD_ROTATED, DOUBLE_JAPANESE_POSTCARD_ROTATED, A6_ROTATED, JAPANESE_ENVELOPE_KAKU2_ROTATED, JAPANESE_ENVELOPE_KAKU3_ROTATED, JAPANESE_ENVELOPE_CHOU3_ROTATED, JAPANESE_ENVELOPE_CHOU4_ROTATED, B6, B6_ROTATED, 12x11, JAPANESE_ENVELOPE_YOU4, JAPANESE_ENVELOPE_YOU4_ROTATED, PRC16K, PRC32K, PRC32K_BIG, PRC_ENVELOPE1, PRC_ENVELOPE2, PRC_ENVELOPE3, PRC_ENVELOPE4, PRC_ENVELOPE5, PRC_ENVELOPE6, PRC_ENVELOPE7, PRC_ENVELOPE8, PRC_ENVELOPE9, PRC_ENVELOPE10, PRC16K_ROTATED, PRC32K_ROTATED, PRC32KBIG_ROTATED, PRC_ENVELOPE1_ROTATED, PRC_ENVELOPE2_ROTATED, PRC_ENVELOPE3_ROTATED, PRC_ENVELOPE4_ROTATED, PRC_ENVELOPE5_ROTATED, PRC_ENVELOPE6_ROTATED, PRC_ENVELOPE7_ROTATED, PRC_ENVELOPE8_ROTATED, PRC_ENVELOPE9_ROTATED, PRC_ENVELOPE10_ROTATED}
		setPaper(Paper := 0) => DllCall('libxl\xlSheetSetPaper', 'ptr', this, 'int', paper, 'cdecl')

		; 返回在打印时工作表的页眉文本。
		header() => DllCall('libxl\xlSheetHeader', 'ptr', this, 'cdecl str')
		
		; 设置打印时工作表的页眉文本。打印时，这些文字出现在每页的顶部。文本的长度必须小于或等于255。标题文本可以包含特殊的命令，例如页码、当前日期或文本格式属性的占位符。特殊命令用一个带“&”的字母表示。边距以英寸为单位指定。
		; https://www.libxl.com/spreadsheet.html
		setHeader(header, margin := 0.5) => DllCall('libxl\xlSheetSetHeader', 'ptr', this, 'str', header, 'double', margin, 'cdecl')
		
		; 返回页眉页边距，单位为英寸。
		headerMargin() => DllCall('libxl\xlSheetHeaderMargin', 'ptr', this, 'cdecl double')
		
		; 返回在打印时工作表的页脚文本。
		footer() => DllCall('libxl\xlSheetFooter', 'ptr', this, 'cdecl str')
		
		; 设置打印时工作表的页脚文本。打印时，这些文字出现在每页的底部。文本的长度必须小于或等于255。标题文本可以包含特殊的命令，例如页码、当前日期或文本格式属性的占位符。特殊命令用一个带“&”的字母表示。边距以英寸为单位指定。
		; https://www.libxl.com/spreadsheet.html
		setFooter(footer, margin := 0.5) => DllCall('libxl\xlSheetSetFooter', 'ptr', this, 'str', footer, 'double', margin, 'cdecl')
		
		; 返回页脚页边距，单位为英寸。
		footerMargin() => DllCall('libxl\xlSheetFooterMargin', 'ptr', this, 'cdecl double')
		
		; 返回打印时工作表是否水平居中：1 -是，0 -不是。
		hCenter() => DllCall('libxl\xlSheetHCenter', 'ptr', this, 'cdecl')
		
		; 设置打印时工作表是否水平居中：1 -是，0 -不是。
		setHCenter(hCenter := true) => DllCall('libxl\xlSheetSetHCenter', 'ptr', this, 'int', hCenter, 'cdecl')
		
		; 返回打印时工作表是否垂直居中：1 -是，0 -不是。
		vCenter() => DllCall('libxl\xlSheetVCenter', 'ptr', this, 'cdecl')
		
		; 设置打印时工作表是否垂直居中：1 -是，0 -不是。
		setVCenter(vCenter := true) => DllCall('libxl\xlSheetSetVCenter', 'ptr', this, 'int', vCenter, 'cdecl')
		
		; 返回工作表的左边距，以英寸为单位。
		marginLeft() => DllCall('libxl\xlSheetMarginLeft', 'ptr', this, 'cdecl double')
		
		; 设置工作表的左边距，以英寸为单位。
		setMarginLeft(margin) => DllCall('libxl\xlSheetSetMarginLeft', 'ptr', this, 'double', margin, 'cdecl')
		
		; 返回工作表的右边距，以英寸为单位。
		marginRight() => DllCall('libxl\xlSheetMarginRight', 'ptr', this, 'cdecl double')
		
		; 设置工作表的右边距，以英寸为单位。
		setMarginRight(margin) => DllCall('libxl\xlSheetSetMarginRight', 'ptr', this, 'double', margin, 'cdecl')
		
		; 返回工作表的顶边距，以英寸为单位。
		marginTop() => DllCall('libxl\xlSheetMarginTop', 'ptr', this, 'cdecl double')
		
		; 设置工作表的顶边距，以英寸为单位。
		setMarginTop(margin) => DllCall('libxl\xlSheetSetMarginTop', 'ptr', this, 'double', margin, 'cdecl')

		; 返回工作表的底边距，以英寸为单位。
		marginBottom() => DllCall('libxl\xlSheetMarginBottom', 'ptr', this, 'cdecl double')

		; 设置工作表的底边距，以英寸为单位。
		setMarginBottom(margin) => DllCall('libxl\xlSheetSetMarginBottom', 'ptr', this, 'double', margin, 'cdecl')

		; 返回是否打印行和列标题：1 -是，0 -否。
		printRowCol() => DllCall('libxl\xlSheetPrintRowCol', 'ptr', this, 'cdecl')

		; 设置是否打印行标题和列标题：1 -是，0 -否。
		setPrintRowCol(print := true) => DllCall('libxl\xlSheetSetPrintRowCol', 'ptr', this, 'int', print, 'cdecl')
		
		; 获取每页上从rowFirst到rowLast的重复行。如果没有找到重复的行，则返回0。
		printRepeatRows(&rowFirst, &rowLast) => (res := DllCall('libxl\xlSheetPrintRepeatRows', 'ptr', this, 'int*', &rowFirst := -1, 'int*', &rowLast := -1, 'cdecl'), rowFirst++, rowLast++, res)
		
		; 设置每页上从rowFirst到rowLast的重复行。
		setPrintRepeatRows(rowFirst, rowLast) => DllCall('libxl\xlSheetSetPrintRepeatRows', 'ptr', this, 'int', rowFirst-1, 'int', rowLast-1, 'cdecl')
		
		; 获取每页上从colFirst到colLast的重复列。如果没有找到重复的列，则返回0。
		printRepeatCols(&colFirst, &colLast) => (res := DllCall('libxl\xlSheetPrintRepeatCols', 'ptr', this, 'int*', &colFirst := 0, 'int*', &colLast := 0, 'cdecl'), colFirst++, colLast++, res)
		
		; 设置每页上从colFirst到colLast的重复列。
		setPrintRepeatCols(colFirst, colLast) => DllCall('libxl\xlSheetSetPrintRepeatCols', 'ptr', this, 'int', colFirst-1, 'int', colLast-1, 'cdecl')
		
		; 获取打印区域。如果没有找到打印区域，返回0。
		printArea(&rowFirst, &rowLast, &colFirst, &colLast) => (res := DllCall('libxl\xlSheetPrintArea', 'ptr', this, 'int*', &rowFirst := -1, 'int*', &rowLast := -1, 'int*', &colFirst := -1, 'int*', &colLast := -1, 'cdecl'), rowFirst++, rowLast++, colFirst++, colLast++, res)
		
		; 设置打印区域。
		setPrintArea(rowFirst, rowLast, colFirst, colLast) => DllCall('libxl\xlSheetSetPrintArea', 'ptr', this, 'int', rowFirst-1, 'int', rowLast-1, 'int', colFirst-1, 'int', colLast-1, 'cdecl')
		
		; 清除每页上重复的行和列。
		clearPrintRepeats() => DllCall('libxl\xlSheetClearPrintRepeats', 'ptr', this, 'cdecl')
		
		; 清除打印区域。
		clearPrintArea() => DllCall('libxl\xlSheetClearPrintArea', 'ptr', this, 'cdecl')

		/**
		 * 按名称获取已命名的范围坐标
		 * @param {String}  name      命名范围的名称
		 * @param {Integer} rowFirst  第一行行号
		 * @param {Integer} rowLast   最后一行行号
		 * @param {Integer} colFirst  第一列列号
		 * @param {Integer} colLast   最后一列列号
		 * @param {Number}  scopeId   局部命名范围的工作表索引或全局命名范围的SCOPE_WORKBOOK
		 * @param {Integer} hidden    是否隐藏，是为1，否为0
		 * @returns {Integer}         如果没有找到指定的命名范围或发生错误，则返回0
		 */
		getNamedRange(name, &rowFirst, &rowLast, &colFirst, &colLast, scopeId := -2, &hidden := 0) => (res := DllCall('libxl\xlSheetGetNamedRange', 'ptr', this, 'str', name, 'int*', &rowFirst := -1, 'int*', &rowLast := -1, 'int*', &colFirst := -1, 'int*', &colLast := -1, 'int', scopeId, 'int*', &hidden := 0, 'cdecl'), rowFirst++, rowLast++, colFirst++, colLast++, res)
		
		/**
		 * 设置命名范围
		 * @param {String}  name      命名范围的名称
		 * @param {Integer} rowFirst  第一行行号
		 * @param {Integer} rowLast   最后一行行号
		 * @param {Integer} colFirst  第一列列号
		 * @param {Integer} colLast   最后一列列号
		 * @param {Number}  scopeId   局部命名范围的工作表索引或全局命名范围的SCOPE_WORKBOOK
		 * @param {Integer} hidden    是否隐藏，是为1，否为0
		 * @returns {Integer}         发生错误时返回0
		 */
		setNamedRange(name, rowFirst, rowLast, colFirst, colLast, scopeId := -2, hidden := false) => DllCall('libxl\xlSheetSetNamedRange', 'ptr', this, 'str', name, 'int', rowFirst-1, 'int', rowLast-1, 'int', colFirst-1, 'int', colLast-1, 'int', scopeId, 'cdecl')
		
		; 按名称删除指定范围。scopeId - 局部命名范围的工作表索引或全局命名范围的SCOPE_WORKBOOK。如果发生错误，返回0。
		delNamedRange(name, scopeId := -2) => DllCall('libxl\xlSheetDelNamedRange', 'ptr', this, 'str', name, 'int', scopeId, 'cdecl')
		
		; 返回工作表中命名范围的数量。
		namedRangeSize() => DllCall('libxl\xlSheetNamedRangeSize', 'ptr', this, 'cdecl')
		
		; 按索引获取命名范围坐标。scopeId -局部命名范围的工作表索引或全局命名范围的SCOPE_WORKBOOK。如果命名范围是隐藏的，则为1，如果不是，则为0。
		namedRange(index, &rowFirst, &rowLast, &colFirst, &colLast, &scopeId := 0, &hidden := 0) => (res := DllCall('libxl\xlSheetNamedRange', 'ptr', this, 'int', index-1, 'int*', &rowFirst := -1, 'int*', &rowLast := -1, 'int*', &colFirst := -1, 'int*', &colLast := -1, 'int*', &scopeId := 0, 'int*', &hidden := 0, 'cdecl str'), rowFirst++, rowLast++, colFirst++, colLast++, res)
		
		; 按名称获取表参数。如果找到表，返回1。
		; headerRowCount — 显示在表顶部的标题行数。0表示不显示标题行。
		; totalsRowCount — 将显示在表底部的总行数。0表示不显示总计行。
		getTable(name, &rowFirst, &rowLast, &colFirst, &colLast, &headerRowCount, &totalsRowCount) => (res := DllCall('libxl\xlSheetGetTable', 'ptr', this, 'str', name, 'int*', &rowFirst := -1, 'int*', &rowLast := -1, 'int*', &colFirst := -1, 'int*', &colLast := -1, 'int*', &headerRowCount := 0, 'int*', &totalsRowCount := 0, 'cdecl str'), rowFirst++, rowLast++, colFirst++, colLast++, res)
		
		; 返回工作表中的表格数量。
		tableSize() => DllCall('libxl\xlSheetTableSize', 'ptr', this, 'cdecl')
		
		; 按索引获取表参数。返回表名称字符串。
		; headerRowCount—显示在表顶部的标题行数。0表示不显示标题行。
		; totalsRowCount—将显示在表底部的总行数。0表示不显示总计行。
		table(index, &rowFirst, &rowLast, &colFirst, &colLast, &headerRowCount, &totalsRowCount) => (res := DllCall('libxl\xlSheetTable', 'ptr', this, 'int', index-1, 'int*', &rowFirst := -1, 'int*', &rowLast := -1, 'int*', &colFirst := -1, 'int*', &colLast := -1, 'int*', &headerRowCount := 0, 'int*', &totalsRowCount := 0, 'cdecl str'), rowFirst++, rowLast++, colFirst++, colLast++, res)
		
		; 返回工作表中超链接的数量。
		hyperlinkSize() => DllCall('libxl\xlSheetHyperlinkSize', 'ptr', this, 'cdecl')
		
		; 按索引获取超链接及其范围。
		hyperlink(index, &rowFirst, &rowLast, &colFirst, &colLast) => (res := DllCall('libxl\xlSheetHyperlink', 'ptr', this, 'int', index-1, 'int*', &rowFirst := -1, 'int*', &rowLast := -1, 'int*', &colFirst := -1, 'int*', &colLast := -1, 'cdecl str'), rowFirst++, rowLast++, colFirst++, colLast++, res)
		
		; 删除索引为index的超链接
		delHyperlink(index) => DllCall('libxl\xlSheetDelHyperlink', 'ptr', this, 'int', index-1, 'cdecl')
		
		; 新增超链接
		addHyperlink(hyperlink, rowFirst, rowLast, colFirst, colLast) => DllCall('libxl\xlSheetAddHyperlink', 'ptr', this, 'str', hyperlink, 'int', rowFirst-1, 'int', rowLast-1, 'int', colFirst-1, 'int', colLast-1, 'cdecl')
		
		; 检查单元格是否包含超链接。如果存在，则返回超链接的索引，如果此单元格中没有超链接，则返回0。
		hyperlinkIndex(row, col) => DllCall('libxl\xlSheetHyperlinkIndex', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl') + 1
		
		; 如果筛选器已经存在（仅适用于xlsx文件），则返回true。不存在则返回false。
		isAutoFilter() => DllCall('libxl\xlSheetIsAutoFilter', 'ptr', this, 'cdecl')
		
		; 返回筛选器。如果不存在则创建它（仅用于xlsx文件）。
		autoFilter() => XL.IAutoFilter(DllCall('libxl\xlSheetAutoFilter', 'ptr', this, 'cdecl ptr'))
		
		; 应用筛选器的设定（仅适用于xlsx文件）。刷新表
		applyFilter() => DllCall('libxl\xlSheetApplyFilter', 'ptr', this, 'cdecl')
		
		; 删除筛选器（仅适用于xlsx文件）。
		removeFilter() => DllCall('libxl\xlSheetRemoveFilter', 'ptr', this, 'cdecl')
		
		; 返回表名称
		name() => DllCall('libxl\xlSheetName', 'ptr', this, 'cdecl str')
		
		; 设置表名称,不能包含 \ * ? / [ ] :
		setName(name) => DllCall('libxl\xlSheetSetName', 'ptr', this, 'str', name, 'cdecl')
		
		; 返回工作表是否受保护：1 -是，0 -否。
		protect() => DllCall('libxl\xlSheetProtect', 'ptr', this, 'cdecl')
		
		; 设置工作表保护: protect = 1-保护 0-不保护。使用下面的密码和增强参数保护工作表。可以将几个增强保护值与操作符|组合在一起。
		; EnhancedProtection {DEFAULT = -1, ALL = 0, OBJECTS = 1, SCENARIOS = 2, FORMAT_CELLS = 4, FORMAT_COLUMNS = 8, FORMAT_ROWS = 16, INSERT_COLUMNS = 32, INSERT_ROWS = 64, INSERT_HYPERLINKS = 128, DELETE_COLUMNS = 256, DELETE_ROWS = 512, SEL_LOCKED_CELLS = 1024, SORT = 2048, AUTOFILTER = 4096, PIVOTTABLES = 8192, SEL_UNLOCKED_CELLS = 16384}
		setProtect(protect := true, password := 0, enhancedProtection := -1) => DllCall('libxl\xlSheetSetProtectEx', 'ptr', this, 'int', protect, 'ptr', Type(password) = 'String' ? StrPtr(password) : password, 'int', enhancedProtection, 'cdecl')
		
		; 返回工作表是否隐藏：1 -是，0 -否。
		hidden() => DllCall('libxl\xlSheetHidden', 'ptr', this, 'cdecl')
		
		; 设置工作表隐藏状态：SheetState {0:VISIBLE, 1:HIDDEN, 2:VERYHIDDEN}
		setHidden(SheetState := 1) => DllCall('libxl\xlSheetSetHidden', 'ptr', this, 'int', SheetState, 'cdecl')
		
		; 获取工作表的第一个可见行和最左边可见列。
		getTopLeftView(&row, &col) => (res := DllCall('libxl\xlSheetGetTopLeftView', 'ptr', this, 'int*', &row := -1, 'int*', &col := -1, 'cdecl'), row++, col++, res)
		
		; 设置工作表的第一个可见行和最左边可见列。
		setTopLeftView(row, col) => DllCall('libxl\xlSheetSetTopLeftView', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl')
		
		; 返回文本是否以从右到左的模式显示：1 -是，0 -否。
		rightToLeft() => DllCall('libxl\xlSheetRightToLeft', 'ptr', this, 'cdecl')
		
		; 设置从右到左模式：
		; 1 -文本以从右到左的模式显示，
		; 0 -文本以从左到右的方式显示。
		setRightToLeft(rightToLeft := true) => DllCall('libxl\xlSheetSetRightToLeft', 'ptr', this, 'int', rightToLeft, 'cdecl')
		
		; 设置自动拟合列宽度特性的边框。宽度值为-1的函数Sheet.SetCol()将只影响指定的有限区域。
		setAutoFitArea(rowFirst := 1, colFirst := 1, rowLast := 0, colLast := 0) => DllCall('libxl\xlSheetSetAutoFitArea', 'ptr', this, 'int', rowFirst-1, 'int', colFirst-1, 'int', rowLast-1, 'int', colLast-1, 'cdecl')
		
		/**
		 * 将单元格引用转换为行号和列号
		 * @param {String}  addr          单元格引用，可以是相对的也可以是绝对的，例如C5或$C$5
		 * @param {Integer} row           从单元格引用中提取的行号
		 * @param {Integer} col           从单元格参考中提取的列号
		 * @param {Integer} rowRelative   如果行是相对的为true，如果行是绝对的为false
		 * @param {Integer} colRelative   如果列是相对的为true，如果列是绝对的为false
		 */
		addrToRowCol(addr, &row, &col, &rowRelative := 0, &colRelative := 0) => (res := DllCall('libxl\xlSheetAddrToRowCol', 'ptr', this, 'str', StrUpper(addr), 'int*', &row := 0, 'int*', &col := 0, 'int*', &rowRelative := 0, 'int*', &colRelative := 0, 'cdecl'), row++, col++, res)
		
		/**
		 * 将行号和列号转换为单元格引用
		 * @param {Integer} row          单元格引用的行号
		 * @param {Integer} col          单元格引用的列号
		 * @param {Integer} rowRelative  如果row应该是相对的为true, 如果row应该是绝对的为false
		 * @param {Integer} colRelative  如果列应该是相对的为true, 如果列应该是绝对的为false
		 * @returns {String}             返回单元格引用，可以是相对的也可以是绝对的，例如C5或$C$5。 
		 */
		rowColToAddr(row, col, rowRelative := true, colRelative := true) => DllCall('libxl\xlSheetRowColToAddr', 'ptr', this, 'int', row-1, 'int', col-1, 'int', rowRelative, 'int', colRelative, 'cdecl str')
		
		; 返回工作表选项卡的颜色。 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		tabColor(Color) => DllCall('libxl\xlSheetTabColor', 'ptr', this, 'cdecl')
		
		; 设置工作表选项卡的颜色。 Color {BLACK = 8, WHITE, RED, BRIGHTGREEN, BLUE, YELLOW, PINK, TURQUOISE, DARKRED, GREEN, DARKBLUE, DARKYELLOW, VIOLET, TEAL, GRAY25, GRAY50, PERIWINKLE_CF, PLUM_CF, IVORY_CF, LIGHTTURQUOISE_CF, DARKPURPLE_CF, CORAL_CF, OCEANBLUE_CF, ICEBLUE_CF, DARKBLUE_CL, PINK_CL, YELLOW_CL, TURQUOISE_CL, VIOLET_CL, DARKRED_CL, TEAL_CL, BLUE_CL, SKYBLUE, LIGHTTURQUOISE, LIGHTGREEN, LIGHTYELLOW, PALEBLUE, ROSE, LAVENDER, TAN, LIGHTBLUE, AQUA, LIME, GOLD, LIGHTORANGE, ORANGE, BLUEGRAY, GRAY40, DARKTEAL, SEAGREEN, DARKGREEN, OLIVEGREEN, BROWN, PLUM, INDIGO, GRAY80, DEFAULT_FOREGROUND = 0x40, DEFAULT_BACKGROUND = 0x41, TOOLTIP = 0x51, NONE = 0x7F, AUTO = 0x7FFF}
		; https://www.libxl.com/colors.html
		setTabColor(Color) => DllCall('libxl\xlSheetSetTabColor', 'ptr', this, 'int', Color, 'cdecl')
		
		; 返回工作表选项卡的RGB颜色。
		getTabRGBColor(&red, &green, &blue) => DllCall('libxl\xlSheetGetTabRgbColor', 'ptr', this, 'int*', &red := 0, 'int*', &green := 0, 'int*', &blue := 0, 'cdecl')
		
		; 设置工作表选项卡的RGB颜色。
		setTabRGBColor(red, green, blue) => DllCall('libxl\xlSheetSetTabRgbColor', 'ptr', this, 'int', red, 'int', green, 'int', blue, 'cdecl')
		
		; 添加指定范围的忽略错误。它允许隐藏单元格左侧的绿色三角形。例如，如果单元格的格式为文本，但包含数值，则认为这是一个潜在的错误，因为在计算中，该数字不会被视为数字。
		; 可以将几个IgnoredError值与操作符|组合在一起。如果发生错误，返回0。
		; IgnoredError {NO_ERROR = 0, EVAL_ERROR = 1, EMPTY_CELLREF = 2, NUMBER_STORED_AS_TEXT = 4, INCONSIST_RANGE = 8, INCONSIST_FMLA = 16, TWODIG_TEXTYEAR = 32, UNLOCK_FMLA = 64, DATA_VALIDATION = 128}
		addIgnoredError(rowFirst, colFirst, rowLast, colLast, IgnoredError) => DllCall('libxl\xlSheetAddIgnoredError', 'ptr', this, 'int', rowFirst-1, 'int', colFirst-1, 'int', rowLast-1, 'int', colLast-1, 'int', IgnoredError, 'cdecl')
		
		/**
		 * 为指定范围添加数据验证（仅适用于xlsx文件）
		 * @param {Integer} type      DataValidationType {TYPE_NONE, TYPE_WHOLE, TYPE_DECIMAL, TYPE_LIST, TYPE_DATE, TYPE_TIME, TYPE_TEXTLENGTH, TYPE_CUSTOM}
		 * @param {Integer} op        DataValidationOperator {OP_BETWEEN, OP_NOTBETWEEN, OP_EQUAL, OP_NOTEQUAL, OP_LESSTHAN, OP_LESSTHANOREQUAL, OP_GREATERTHAN, OP_GREATERTHANOREQUAL}
		 * @param {Integer} rowFirst  第一行行号
		 * @param {Integer} rowLast   最后一行行号
		 * @param {Integer} colFirst  第一列列号
		 * @param {Integer} colLast   最后一列列号
		 * @param {String}  value1    关系运算符的第一个值，如果要直接指定值列表，请使用双引号(例如"A,B,C")，如果你想用值指定对区域的引用，不要使用引号(例如A1:A6)；
		 * @param {String}  value2    VALIDATION_OP_BETWEEN 或 VALIDATION_OP_NOTBETWEEN 操作符的第二个值。
		 */
		addDataValidation(type, op, rowFirst, rowLast, colFirst, colLast, value1, value2) => DllCall('libxl\xlSheetAddDataValidation', 'ptr', this, 'int', type, 'int', op, 'int', rowFirst-1, 'int', rowLast-1, 'int', colFirst-1, 'int', colLast-1, 'str', value1, 'str', value2, 'cdecl')
		
		/**
		 * 为指定范围添加数据验证（仅适用于xlsx文件）
		 * @param {Integer} type      DataValidationType {TYPE_NONE, TYPE_WHOLE, TYPE_DECIMAL, TYPE_LIST, TYPE_DATE, TYPE_TIME, TYPE_TEXTLENGTH, TYPE_CUSTOM}
		 * @param {Integer} op        DataValidationOperator {OP_BETWEEN, OP_NOTBETWEEN, OP_EQUAL, OP_NOTEQUAL, OP_LESSTHAN, OP_LESSTHANOREQUAL, OP_GREATERTHAN, OP_GREATERTHANOREQUAL}
		 * @param {Integer} rowFirst  第一行行号
		 * @param {Integer} rowLast   最后一行行号
		 * @param {Integer} colFirst  第一列列号
		 * @param {Integer} colLast   最后一列列号
		 * @param {String}  value1    关系运算符的第一个值，如果要直接指定值列表，请使用双引号(例如"A,B,C")，如果你想用值指定对区域的引用，不要使用引号(例如A1:A6)；
		 * @param {String}  value2    VALIDATION_OP_BETWEEN 或 VALIDATION_OP_NOTBETWEEN 操作符的第二个值。
		 * @param {boolean} allowBlank        指示数据验证是否将空或空白条目视为有效，‘true’表示空条目是OK的，并且不违反验证约束
		 * @param {boolean} hideDropDown      指示是否为列表类型的数据验证（VALIDATION_TYPE_LIST）显示下拉组合框
		 * @param {boolean} showInputMessage  指示是否显示输入提示消息
		 * @param {boolean} showErrorMessage  指示是否根据指定的标准在输入无效值时显示错误警报消息
		 * @param {String}  promptTitle       输入提示符的标题栏文本
		 * @param {String} prompt             输入提示的消息文本
		 * @param {String} errorTitle         错误警告的标题栏文本
		 * @param {String} error              错误警报的消息文本
		 * @param {Integer} errorStyle        用于此数据验证的错误警报样式 DataValidationErrorStyle {0:ERRSTYLE_STOP, 1:ERRSTYLE_WARNING, 2:ERRSTYLE_INFORMATION}
		 */
		addDataValidationEx(type, op, rowFirst, rowLast, colFirst, colLast, value1, value2, allowBlank := true, hideDropDown := false, showInputMessage := true, showErrorMessage := true, promptTitle := 0, prompt := 0, errorTitle := 0, error := 0, errorStyle := 0) => DllCall('libxl\xlSheetAddDataValidationEx', 'ptr', this, 'int', type, 'int', op, 'int', rowFirst-1, 'int', rowLast-1, 'int', colFirst-1, 'int', colLast-1, 'str', value1, 'str', value2, 'int', allowBlank, 'int', hideDropDown, 'int', showInputMessage, 'int', showErrorMessage, 'str', promptTitle, 'str', prompt, 'str', errorTitle, 'str', error, 'int', errorStyle, 'cdecl')
		
		/**
		 * 使用关系运算符的双精度值或日期值为指定范围添加数据验证（仅适用于xlsx文件）。
		 * @param {Integer} type      DataValidationType {TYPE_NONE, TYPE_WHOLE, TYPE_DECIMAL, TYPE_LIST, TYPE_DATE, TYPE_TIME, TYPE_TEXTLENGTH, TYPE_CUSTOM}
		 * @param {Integer} op        DataValidationOperator {OP_BETWEEN, OP_NOTBETWEEN, OP_EQUAL, OP_NOTEQUAL, OP_LESSTHAN, OP_LESSTHANOREQUAL, OP_GREATERTHAN, OP_GREATERTHANOREQUAL}
		 * @param {Integer} rowFirst  第一行行号
		 * @param {Integer} rowLast   最后一行行号
		 * @param {Integer} colFirst  第一列列号
		 * @param {Integer} colLast   最后一列列号
		 * @param {String}  value1    关系运算符的第一个值，如果要直接指定值列表，请使用双引号(例如"A,B,C")，如果你想用值指定对区域的引用，不要使用引号(例如A1:A6)；
		 * @param {String}  value2    VALIDATION_OP_BETWEEN 或 VALIDATION_OP_NOTBETWEEN 操作符的第二个值。
		 */
		addDataValidationDouble(type, op, rowFirst, rowLast, colFirst, colLast, value1, value2) => DllCall('libxl\xlSheetAddDataValidationDouble', 'ptr', this, 'int', type, 'int', op, 'int', rowFirst-1, 'int', rowLast-1, 'int', colFirst-1, 'int', colLast-1, 'double', value1, 'double', value2, 'cdecl')
		
		/**
		 * 为带有扩展参数的关系运算符添加双值或日期值的指定范围的数据验证（仅适用于xlsx文件）。
		 * @param {Integer} type      DataValidationType {TYPE_NONE, TYPE_WHOLE, TYPE_DECIMAL, TYPE_LIST, TYPE_DATE, TYPE_TIME, TYPE_TEXTLENGTH, TYPE_CUSTOM}
		 * @param {Integer} op        DataValidationOperator {OP_BETWEEN, OP_NOTBETWEEN, OP_EQUAL, OP_NOTEQUAL, OP_LESSTHAN, OP_LESSTHANOREQUAL, OP_GREATERTHAN, OP_GREATERTHANOREQUAL}
		 * @param {Integer} rowFirst  第一行行号
		 * @param {Integer} rowLast   最后一行行号
		 * @param {Integer} colFirst  第一列列号
		 * @param {Integer} colLast   最后一列列号
		 * @param {String}  value1    关系运算符的第一个值，如果要直接指定值列表，请使用双引号(例如"A,B,C")，如果你想用值指定对区域的引用，不要使用引号(例如A1:A6)；
		 * @param {String}  value2    VALIDATION_OP_BETWEEN 或 VALIDATION_OP_NOTBETWEEN 操作符的第二个值。
		 * @param {boolean} allowBlank        指示数据验证是否将空或空白条目视为有效，‘true’表示空条目是OK的，并且不违反验证约束
		 * @param {boolean} hideDropDown      指示是否为列表类型的数据验证（VALIDATION_TYPE_LIST）显示下拉组合框
		 * @param {boolean} showInputMessage  指示是否显示输入提示消息
		 * @param {boolean} showErrorMessage  指示是否根据指定的标准在输入无效值时显示错误警报消息
		 * @param {String}  promptTitle       输入提示符的标题栏文本
		 * @param {String} prompt             输入提示的消息文本
		 * @param {String} errorTitle         错误警告的标题栏文本
		 * @param {String} error              错误警报的消息文本
		 * @param {Integer} errorStyle        用于此数据验证的错误警报样式 DataValidationErrorStyle {0:ERRSTYLE_STOP, 1:ERRSTYLE_WARNING, 2:ERRSTYLE_INFORMATION}
		 */
		addDataValidationDoubleEx(type, op, rowFirst, rowLast, colFirst, colLast, value1, value2, allowBlank := true, hideDropDown := false, showInputMessage := true, showErrorMessage := true, promptTitle := 0, prompt := 0, errorTitle := 0, error := 0, errorStyle := 0) => DllCall('libxl\xlSheetAddDataValidationDoubleEx', 'ptr', this, 'int', type, 'int', op, 'int', rowFirst-1, 'int', rowLast-1, 'int', colFirst-1, 'int', colLast-1, 'double', value1, 'double', value2, 'int', allowBlank, 'int', hideDropDown, 'int', showInputMessage, 'int', showErrorMessage, 'str', promptTitle, 'str', prompt, 'str', errorTitle, 'str', error, 'int', errorStyle, 'cdecl')
		
		; 删除工作表的所有数据验证（仅针对xlsx文件）。
		removeDataValidations() => DllCall('libxl\xlSheetRemoveDataValidations', 'ptr', this, 'cdecl')
		
		; 返回这个工作表中的表单控件数量（仅限于xlsx文件）。
		formControlSize() => DllCall('libxl\xlSheetFormControlSize', 'ptr', this, 'cdecl')
		
		; 返回具有指定索引的表单控件（仅适用于xlsx文件）。索引必须小于等于Sheet.FormControlSize()函数的返回值。
		; index待验证
		formControl(index) => XL.IFormControl(DllCall('libxl\xlSheetFormControl', 'ptr', this, 'int', index-1, 'cdecl ptr'))
		
		; 向工作表添加条件格式规则（仅适用于xlsx文件）。
		; index待验证
		addConditionalFormatting() => XL.IConditionalFormatting(DllCall('libxl\xlSheetAddConditionalFormatting', 'ptr', this, 'cdecl ptr'))
		
		; 获取工作表的活动单元格。如果找到活动单元格，则返回1，否则返回0。
		getActiveCell(&row, &col) => (res := DllCall('libxl\xlSheetGetActiveCell', 'ptr', this, 'int*', &row := -1, 'int*', &col := -1, 'cdecl'), row++, col++, res)
		
		; 设置工作表的活动单元格
		setActiveCell(row, col) => DllCall('libxl\xlSheetSetActiveCell', 'ptr', this, 'int', row-1, 'int', col-1, 'cdecl')
		
		; 返回所选内容的范围。
		selectionRange() => DllCall('libxl\xlSheetSelectionRange', 'ptr', this, 'cdecl str')
		
		; 向所选内容添加一个范围。
		addSelectionRange(sqref) => DllCall('libxl\xlSheetAddSelectionRange', 'ptr', this, 'str', sqref, 'cdecl')
		
		; 移除所有选择。
		removeSelection() => DllCall('libxl\xlSheetRemoveSelection', 'ptr', this, 'cdecl')
		



		/********************
		 * 自定义的表格函数 *
		 ********************/

		/**
		 * 复制某一行的到另一行
		 * @param {Integer}  srcRow  源行号
		 * @param {Integer}  dstRow  目标行号
		 */
		CopyRowTo(srcRow, dstRow) {
			Loop this.lastCol()
				this.copyCell(srcRow, A_Index, dstRow, A_Index)
		}

		/**
		 * 复制某一列的到另一列
		 * @param {Integer}  srcCol  源列号
		 * @param {Integer}  dstCol  目标列号
		 */
		CopyColTo(srcCol, dstCol) {
			Loop this.lastRow()  
				this.copyCell(A_Index, srcCol, A_Index, dstCol)
		}

		/**
		 * 复制某一行的格式到另一行
		 * @param {Integer}  srcRow   源行号
		 * @param {Integer}  dstRow   目标行号
		 * @param {XL.Sheet} dstSheet 目标工作表，默认为当前工作表
		 */
		CopyRowFormatTo(srcRow, dstRow, dstSheet?) {
			dstSheet := dstSheet ?? this
			Loop this.lastCol()	  
				if (srcFormat := this.cellFormat(srcRow, A_Index)) 
					dstSheet[dstRow, A_Index].format := srcFormat
		}

		/**
		 * 复制某一列的格式到另一列
		 * @param {Integer}  srcCol 
		 * @param {Integer}  dstCol 
		 * @param {XL.Sheet} dstSheet 目标工作表，默认为当前工作表
		 */
		CopyColFormatTo(srcCol, dstCol, dstSheet?) {
			dstSheet := dstSheet ?? this
			Loop this.lastRow()  
				if (srcFormat := this.cellFormat(A_Index, srcCol)) 
					dstSheet[A_Index, dstCol].format := srcFormat
		}

		/**
		 * 向某一行向上插入若干行,并带有该行的格式
		 * @param {Integer} row    行号
		 * @param {Integer} length 插入的行数
		 */
		InsertRowWithFormat(row, length := 1) {
			if length < 1
				return
			srcRow := row + length
			this.insertRow(row, srcRow - 1)
			loop length {
				R := srcRow - A_Index
				Loop this.lastCol() {
					C := A_Index			  
					if (srcFormat := this.cellFormat(srcRow, C)) 
						this[R, C].format := srcFormat
				}
			}
		}

		/**
		 * 向某一行向下插入若干行,并带有该行的格式
		 * @param {Integer} row    行号
		 * @param {Integer} length 插入的行数
		 */
		InsertRowBelowWithFormat(row, length := 1) {
			if length < 1
				return
			srcRow := row
			this.insertRow(srcRow + 1, srcRow + length)
			loop length {
				R := srcRow + A_Index
				Loop this.lastCol() {
					C := A_Index			  
					if (srcFormat := this.cellFormat(srcRow, C)) 
						this[R, C].format := srcFormat
				}
			}
		}




		__Delete() => (this.parent := '')
		__Item[row, col := ''] {
			get => (IsNumber(row) ? '' : this.addrToRowCol(row, &row, &col), XL.ISheet.ICell(row, col, this))
			set {
				if (ret := format := 0, bool := formula := '', !IsNumber(row))
					this.addrToRowCol(row, &row, &col)
				rechecktype:
				switch Type(value) {
				case 'Object':
					val := value, value := ''
					for k in val.OwnProps()
						switch StrLower(k) {
						case 'format':
							format := val.format
						case 'bool':
							value := val.bool, bool := true
						case 'exp', 'expr', 'formula':
							formula := val.%k%
						case 'int', 'integer':
							value := Integer(val.%k%)
						case 'num', 'float', 'number', 'double':
							value := Float(val.%k%)
						default:
							value := val.%k%
						}
					if (formula != '') {
						if (bool)
							ret := this.writeFormulaBool(row, col, formula, !!value, format)
						else if (value = '')
							ret := this.writeFormula(row, col, formula, format)
						else if ('String' = Type(value))
							ret := this.writeFormulaStr(row, col, formula, value, format)
						else ret := this.writeFormulaNum(row, col, formula, value, format)
					} else if (bool)
						ret := this.writeBool(row, col, !!value, format)
					else goto rechecktype
				case 'String':
					ret := this.writeStr(row, col, value, format)
				case 'Integer', 'Float':
					ret := this.writeNum(row, col, value, format)
				case 'XL.IRichString':
					ret := this.writeRichStr(row, col, value, format)
				default:
					throw Error('Wrong parameter type')
				}
				if (!ret && (msg := this.parent.errorMessage()) != 'ok')
					throw Error(msg)
			}
		}
		class ICell {
			__New(row, col, parent) {
				this.row := row, this.col := col, this.parent := parent
			}
			content {
				get {
					format := 0, ret := {value: '', type: '', format: 0}
					switch this.parent.cellType(row := this.row, col := this.col) {
					case 0:	; EMPTY
						ret.type := 'EMPTY'
						return ret
					case 1:	; NUMBER, DATE
						if (this.parent.isDate(row, col)) {
							year := month := day := hour := min := sec := msec := 0
							value := this.parent.readNum(row, col, &format)
							this.parent.parent.dateUnpack(value, &year, &month, &day, &hour, &min, &sec, &msec)
							ret := {year: year, month: month, day: day, hour: hour, min: min, sec: sec, msec: msec}
							ret.type := 'DATE', ret.value := value
						} else if (this.parent.isFormula(row, col)) {
							ret.formula := this.parent.readFormula(row, col, &format), ret.type := 'FORMULA'
							try ret.value := this.parent.readNum(row, col)
						} else ret.type := 'NUMBER', ret.value := this.parent.readNum(row, col, &format)
						ret.format := format
					case 2:	; STRING, FORMULA, RICHSTRING
						if (this.parent.isRichStr(row, col))
							ret.richstr := this.parent.readRichStr(row, col, &format), ret.type := 'RICHSTRING', ret.value := this.parent.readStr(row, col)
						else if (this.parent.isFormula(row, col)) {
							ret.formula := this.parent.readFormula(row, col, &format), ret.type := 'FORMULA'
							try ret.value := this.parent.readStr(row, col)
						} else ret.type := 'STRING', ret.value := this.parent.readStr(row, col, &format)
						ret.format := format
					case 3:	; BOOLEAN
						ret.value := this.parent.readBool(row, col, &format), ret.format := format, ret.type := 'BOOLEAN'
					case 4:	; BLANK
						this.parent.readBlank(row, col, &format), ret.format := format, ret.type := 'BLANK'
					default: ;ERROR
						ret.type := 'ERROR'
						switch ret.errcode := this.parent.readError(row, col) {
						case 0:
							ret.value := '#NULL!'
						case 0x7:
							ret.value := '#DIV/0!'
						case 0xF:
							ret.value := '#VALUE!'
						case 0x17:
							ret.value := '#REF!'
						case 0x1D:
							ret.value := '#NAME?'
						case 0x24:
							ret.value := '#NUM!'
						case 0x2A:
							ret.value := '#N/A'
						default:
							ret.value := 'no error'
						}
					}
					return ret
				}
			}
			; 数值
			value {
				get => this.content.value
				set => (this.parent[this.row, this.col] := {value: value, format: this.format})
			}
			; 格式
			format {
				get => this.parent.cellFormat(this.row, this.col)
				set => this.parent.setCellFormat(this.row, this.col, value)
			}
			; 注释
			comment {
				get => this.parent.readComment(this.row, this.col)
				set {
					if (value = '')
						return this.parent.removeComment(this.row, this.col)
					author := height := width := ''
					for k in ['author', 'width', 'height', 'value']
						%k% := value.HasOwnProp(k) ? value.%k% : ''
					return this.parent.writeComment(this.row, this.col, value, author, height || 129, width || 75)
				}
			}
			; 宽度
			width {
				get => this.parent.colWidth(this.col)
				set => this.parent.setCol(this.col, this.col, value, this.format, this.parent.colHidden(this.col))
			}
			; 高度
			height {
				get => this.parent.rowHeight(this.row)
				set => this.parent.setRow(this.row, this.row, value, this.format, this.parent.rowHidden(this.row))
			}
			; 是否隐藏
			hidden => this.parent.rowHidden(this.row) || this.parent.colHidden(this.col)
			; 复制到目标单元格
			CopyTo(rowDst, colDst) => this.parent.copyCell(this.row, this.col, rowDst, colDst)
		}
	}
}




/**
 * 复制工作表到目标工作簿, 返回新表
 * @param {XL.Sheet}   srcSheet 
 * @param {XL.Book}    dstBook 
 * @param {String}     dstSheetName 
 * @return {XL.Sheet}
 */
XL_CopySheetTo(srcSheet, dstBook, dstSheetName := "") {
	;判断是否是同工作表
	isSameBook := (dstBook = srcSheet.parent) ? true : false
	;确定目标工作表名
	if isSameBook
		dstSheetName := (dstSheetName = srcSheet.name() || dstSheetName = "") ? srcSheet.name() "-" A_Now : dstSheetName
	else
		dstSheetName := (dstSheetName = "") ? srcSheet.name() : dstSheetName
	;确定目标工作表
	try dstSheet := dstBook[dstSheetName]
	catch
		dstSheet := dstBook.addSheet(dstSheetName)
	; 设置列宽
	Loop srcSheet.lastCol()
		dstSheet.setCol(A_Index, A_Index, srcSheet.colWidth(A_Index),, srcSheet.colHidden(A_Index))
	; 复制合并的单元格
	Loop srcSheet.mergeSize()
		if srcSheet.merge(A_index, &rowFirst, &rowLast, &colFirst, &colLast)
			dstSheet.setMerge(rowFirst, rowLast, colFirst, colLast)
	formats := Map()
	Loop srcSheet.lastRow() {
		R := A_Index
		; 设置行高
		dstSheet.setRow(R, srcSheet.rowHeight(R),, srcSheet.rowHidden(R))
		Loop srcSheet.lastCol() {
			C := A_Index			  
			; 复制格式
			if !(srcFormat := srcSheet.cellFormat(R, C)) 
				continue
			; 检查格式
			if !formats.Has(srcFormat)
				formats[srcFormat] := isSameBook ? srcFormat : dstBook.addFormat(srcFormat) ; 不同工作簿时创建格式
			dstFormat := formats[srcFormat]
			; 复制单元格值
			switch srcSheet.cellType(R, C)
			{
				; 1: NUMBER, DATE
				case 1:
					value := srcSheet.readNum(R, C, &srcFormat)
					dstSheet.writeNum(R, C, value, dstFormat)
	
				; 3: BOOLEAN
				case 3:
					value := srcSheet.readBool(R, C, &srcFormat)
					dstSheet.writeBool(R, C, value, dstFormat)
	
				; 2: STRING, FORMULA, RICHSTRING
				case 2:
					value := srcSheet.readStr(R, C, &srcFormat)
					dstSheet.writeStr(R, C, value, dstFormat)
	
				; 4: BLANK
				case 4:
					srcSheet.readBlank(R, C, &srcFormat)
					dstSheet.writeBlank(R, C, dstFormat)
			}
		}
	}
	return dstSheet
}