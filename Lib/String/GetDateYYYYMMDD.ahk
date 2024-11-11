/**
 * 从字符串中提取出日期, 转化成特定格式的时间字符串, 未提取到时返回""
 * @param {String} str 含有特定格式的日期的字符串
 * @param {String} sep Format的格式字符串, 其中年为{1}, 月为{2}, 日为{3}
 * @returns {String}   指定格式的时间字符串
 */
GetDateYYYYMMDD(str, FormatStr := "{1:04}{2:02}{3:02}") {
	if RegExMatch(str, "(\d{2}|\d{4})[-/\.](\d{1,2})[-/\.](\d{1,2})", &SubPat)
        return Format(FormatStr, (StrLen(SubPat[1]) = 2 ? "20" SubPat[1] : SubPat[1]), SubPat[2], SubPat[3])
}




/*
	Function: DateParse
		Converts almost any date format to a YYYYMMDDHH24MISS value.
	Parameters:
		str - a date/time stamp as a string
	Returns:
		A valid YYYYMMDDHH24MISS value which can be used by FormatTime, EnvAdd and other time commands.
	Example:
> time := DateParse("2:35 PM, 27 November, 2007")
	License:
		- Version 1.05 <http://www.autohotkey.net/~polyethene/#dateparse>
		- Dedicated to the public domain (CC0 1.0) <http://creativecommons.org/publicdomain/zero/1.0/>

DateParse(str) {
	static e2 = "i)(?:(\d{1,2}+)[\s\.\-\/,]+)?(\d{1,2}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*)[\s\.\-\/,]+(\d{2,4})"
	str := RegExReplace(str, "((?:" . SubStr(e2, 42, 47) . ")\w*)(\s*)(\d{1,2})\b", "$3$2$1", "", 1)
	If RegExMatch(str, "i)^\s*(?:(\d{4})([\s\-:\/])(\d{1,2})\2(\d{1,2}))?"
		. "(?:\s*[T\s](\d{1,2})([\s\-:\/])(\d{1,2})(?:\6(\d{1,2})\s*(?:(Z)|(\+|\-)?"
		. "(\d{1,2})\6(\d{1,2})(?:\6(\d{1,2}))?)?)?)?\s*$", i)
		d3 := i1, d2 := i3, d1 := i4, t1 := i5, t2 := i7, t3 := i8
	Else If !RegExMatch(str, "^\W*(\d{1,2}+)(\d{2})\W*$", t)
		RegExMatch(str, "i)(\d{1,2})\s*:\s*(\d{1,2})(?:\s*(\d{1,2}))?(?:\s*([ap]m))?", t)
			, RegExMatch(str, e2, d)
	f = %A_FormatFloat%
	SetFormat, Float, 02.0
	d := (d3 ? (StrLen(d3) = 2 ? 20 : "") . d3 : A_YYYY)
		. ((d2 := d2 + 0 ? d2 : (InStr(e2, SubStr(d2, 1, 3)) - 40) // 4 + 1.0) > 0
			? d2 + 0.0 : A_MM) . ((d1 += 0.0) ? d1 : A_DD) . t1
			+ (t1 = 12 ? t4 = "am" ? -12.0 : 0.0 : t4 = "am" ? 0.0 : 12.0) . t2 + 0.0 . t3 + 0.0
	SetFormat, Float, %f%
	Return, d
}

*/