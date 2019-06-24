Attribute VB_Name = "modEncodeQuery"
'' Character Encode Query for Passolo
'' (c) 2014 - 2019 by wanfu (Last modified on 2019.09.25)

'' Command Line Format: Command <Source><-><Translation> <Switch>
'' Command: Name of this Macros file
''<Source>
'' String: The source string to be converted.
''<Translation>
'' String: The translation string to be converted.
''<->: This is the delimiter between the source string and the translation string.
''<Switch>
''Codepage:
'' -scp[N]: Codepage value of Source Text. N is Numeric value of codepage. Such as: 936,1200 et.
'' -tcp[N]: Codepage value of Translation Text. N is Numeric value of codepage. Such as: 936,1200 et.
''Sring escape:
'' -se: Escape text before convert text to code or after converted code to text. No this switch, do not escape text
''Escape type:
'' -et[N]: Escape type. N is 0 = Hex (Default), 1 = Hex escape, 2 = RUL Encode, 3 = HTML Encode, 4 = ISO-8829-1 Encode, 5 = Base4 Encode
''Multibyte only:
'' -ch: Every 2 Hex characters separated by spaces for Hex type
'' -ca: Convert ASCII characters to Hex escape or HTML escape. No this switch, do not convert ASCII characters
'' -ci: Convert illegal character to URL escape. No this switch, do not convert illegal character
''Convert mode:
'' -ac: Auto convert string to code or code to string after enter the data. No this switch, manually convert for UI mode only
''Convert type:
'' -cs: Convert code to String. No this switch, Convert string to code
'' -lb[N]: No this switch to use vbCrLf by default, Otherwise, Use the specified as: 0 = vbCrLf, 1 = vbCr, 2 = vbLf
''Fill type:
'' -fz: By source text length in bytes, padded with null bytes.
'' -fs: By source text length in bytes, padded with null characters.
''Both, if the translation byte is longer than the source byte, will be truncated to be the same as the source byte.
''UI mode:
'' -noui: Do not display a user interface, run silently
''Display Option:
'' -td: Frist display translation windows. No this switch, frist display source windows,
'' -lng[hex language code]: Display UI Language. Supports EngLish, Chinese Simplified and Chinese Traditional only. For sample: 0804,1004,0404,0C04,1404.

'' Return: None
'' For example: modEncodeQuery.bas This is strings.<->This is converted Hex code. -cp:1201 -se -sc -ac -et:1

Option Explicit

Private Const Version = "2018.05.24"
Private Const Build = "180524"
Private Const TextJoinStr = vbCrLf
Private Const ItemJoinStr = ";"

'SendMessage API 部分常数
Private Enum SendMsgValue
	EM_GETLIMITTEXT = &HD5			'0,0				获取一个编辑控件中文本的最大长度
	EM_LIMITTEXT = &HC5				'最大值,0			设置编辑控件中的最大文本长度
	WM_GETTEXT = &H0D				'字节数,字符串地址	获取窗口文本控件的文本
	WM_GETTEXTLENGTH = &H0E			'0,0				获取窗口文本控件的文本的长度(字节数)
	WM_SETTEXT = &H0C				'0,字符串地址		设置窗口文本控件的文本
	WM_VSCROLL = &H115				'控件句柄,滚动条类型,滚动条位置	设置垂着滚动条位置
	SB_BOTTOM = &H07				'控件句柄,滚动条类型,滚动条位置	设置垂着滚动条位置
End Enum

Private Enum KnownCodePage
	CP_UNKNOWN = -1
	CP_ACP = 0
	CP_OEMCP = 1
	CP_MACCP = 2
	CP_THREAD_ACP = 3
	CP_SYMBOL = 42
' ARABIC
	CP_AWIN = 101			'Bidi Windows codepage
	CP_709 = 102			'MS-DOS Arabic Support 	CP 709
	CP_720 = 103			'MS-DOS Arabic Support 	CP 720
	CP_A708 = 104			'ASMO 708
	CP_A449 = 105			'ASMO 449+
	CP_TARB = 106			'MS Transparent Arabic
	CP_NAE = 107			'Nafitha Enhanced Arabic Char Set
	CP_V4 = 108				'Nafitha v 4.0
	CP_MA2 = 109			'Mussaed Al Arabi (MA/2) 	CP 786
	CP_I864 = 110			'IBM Arabic Supplement 	CP 864
	CP_A437 = 111			'Ansi 437 codepage
	CP_AMAC = 112			'Macintosh Code Page
' HEBREW
	CP_HWIN = 201			'Bidi Windows codepage
	CP_862I = 202			'IBM Hebrew Supplement 	CP 862
	CP_7BIT = 203			'IBM Hebrew Supplement 	CP 862 Folded
	CP_ISO = 204			'ISO Hebrew 8859-8 Character Set
	CP_H437 = 205			'Ansi 437 codepage
	CP_HMAC = 206			'Macintosh Code Page
' OEM CODE PAGES
	CP_IBM437 = 437			'OEM United States
	CP_ASMO708 = 708		'Arabic (ASMO 708)
	CP_DOS720 = 720			'Arabic (Transparent ASMO); Arabic (DOS)
	CP_IBM737 = 737			'OEM Greek (formerly 437G); Greek (DOS)
	CP_IMB775 = 775			'OEM Baltic; Baltic (DOS)
	CP_IBM850 = 850			'OEM Multilingual Latin 1; Western European (DOS)
	CP_IBM852 = 852			'OEM Latin 2; Central European (DOS)
	CP_IBM855 = 855			'OEM Cyrillic (primarily Russian)
	CP_IBM857 = 857			'OEM Turkish; Turkish (DOS)
	CP_IBM00858 = 858		'OEM Multilingual Latin 1 + Euro symbol
	CP_IBM860 = 860			'OEM Portuguese; Portuguese (DOS)
	CP_IMB861 = 861			'OEM Icelandic; Icelandic (DOS)
	CP_DOS862 = 862			'OEM Hebrew; Hebrew (DOS)
	CP_IBM863 = 863			'OEM French Canadian; French Canadian (DOS)
	CP_IBM864 = 864			'OEM Arabic; Arabic (864)
	CP_IBM865 = 865			'OEM Nordic; Nordic (DOS)
	CP_CP866 = 866			'OEM Russian; Cyrillic (DOS)
	CP_IMB869 = 869			'OEM Modern Greek; Greek, Modern (DOS)
	CP_IMB870 = 870			'IBM EBCDIC Multilingual/ROECE (Latin 2); IBM EBCDIC Multilingual Latin 2
	CP_THAI = 874			'ANSI/OEM Thai (same as 28605, ISO 8859-15); Thai (Windows)
	CP_CP875 = 875			'IBM EBCDIC Greek Modern
	CP_JAPAN = 932			'ANSI/OEM Japanese; Japanese (Shift-JIS)
	CP_CHINA = 936			'ANSI/OEM Simplified Chinese (PRC, Singapore); Chinese Simplified (GBK)
	CP_KOREA = 949			'ANSI/OEM Korean (Unified Hangul Code)
	CP_TAIWAN = 950			'ANSI/OEM Traditional Chinese (Taiwan; Hong Kong SAR, PRC); Chinese Traditional (Big5)
' Windows UNICODE CODE PAGES
	CP_UNICODELITTLE = 1200	'Unicode UTF-16, little endian byte order (BMP of ISO 10646); available only to managed applications
	CP_UNICODEBIG = 1201	'Unicode UTF-16, big endian byte order; available only to managed applications
' Windows ANSI CODE PAGES
	CP_EASTEUROPE = 1250	'ANSI Central European; Central European (Windows)
	CP_RUSSIAN = 1251		'ANSI Cyrillic; Cyrillic (Windows)
	CP_WESTEUROPE = 1252	'ANSI Latin 1; Western European (Windows)
	CP_GREEK = 1253			'ANSI Greek; Greek (Windows)
	CP_TURKISH = 1254		'ANSI Turkish; Turkish (Windows)
	CP_HEBREW = 1255		'ANSI Hebrew; Hebrew (Windows)
	CP_ARABIC = 1256		'ANSI Arabic; Arabic (Windows)
	CP_BALTIC = 1257		'ANSI Baltic; Baltic (Windows)
	CP_VIETNAMESE = 1258	'ANSI/OEM Vietnamese; Vietnamese (Windows)
' KOREAN
	CP_JOHAB = 1361			'Korean (Johab)
' MAC
	CP_MAC_ROMAN = 10000	'MAC Roman; Western European (Mac)
	CP_MAC_JAPAN = 10001	'Japanese (Mac)
	CP_MAC_CHINESETRAD = 10002	'MAC Traditional Chinese (Big5); Chinese Traditional (Mac)
	CP_MAC_KOREAN = 10003	'Korean (Mac)
	CP_MAC_ARABIC = 10004	'Arabic (Mac)
	CP_MAC_HEBREW = 10005	'Hebrew (Mac)
	CP_MAC_GREEK = 10006	'Greek (Mac)
	CP_MAC_CYRILLIC = 10007	'Cyrillic (Mac)
	CP_MAC_CHINESESIMP = 10008	'MAC Simplified Chinese (GB 2312); Chinese Simplified (Mac)
	CP_MAC_ROMANIAN = 10010	'Romanian (Mac)
	CP_MAC_UKRAINIAN = 10017	'Ukrainian (Mac)
	CP_MAC_THAI = 10021		'Thai (Mac)
	CP_MAC_LATIN2 = 10029	'MAC Latin 2; Central European (Mac)
	CP_MAC_ICELANDIC = 10079	'Icelandic (Mac)
	CP_MAC_TURKISH = 10081	'Turkish (Mac)
	CP_MAC_CROATIAN = 10082	'Croatian (Mac)
' Windows UNICODE CODE PAGES
	CP_UTF_32LE = 12000  	'Unicode UTF-32, little endian byte order; available only to managed applications
	CP_UTF_32BE = 12001		'Unicode UTF-32, big endian byte order; available only to managed applications
' CODE PAGES
	CP_CHINESECNS = 20000	'CNS Taiwan; Chinese Traditional (CNS)
	CP_CHINESEETEN = 20002	'Eten Taiwan; Chinese Traditional (Eten)
	CP_IA5WEST = 20105		'Wang Taiwan
	CP_IA5GERMAN = 20106	'IA5 German (7-bit)
	CP_IA5SWEDISH = 20107	'IA5 Swedish (7-bit)
	CP_IA5NORWEGIAN = 20108	'IA5 Norwegian (7-bit)
	CP_ASCII = 20127		'US-ASCII (7-bit)
	CP_RUSSIANKOI8R = 20866	'Russian (KOI8-R); Cyrillic (KOI8-R)
	CP_RUSSIANKOI8U = 21866	'Ukrainian (KOI8-U); Cyrillic (KOI8-U)
	CP_ISOLATIN1 = 28591	'ISO 8859-1 Latin 1; Western European (ISO)  西欧语言
	CP_ISOEASTEUROPE = 28592	'ISO 8859-2 Central European; Central European (ISO)  中欧语言
	CP_ISOTURKISH = 28593	'ISO 8859-3 Latin 3  南欧语言。世界语也可用此字符集显示。
	CP_ISOBALTIC = 28594	'ISO 8859-4 Baltic	 北欧语言
	CP_ISORUSSIAN = 28595	'ISO 8859-5 Cyrillic   斯拉夫语言
	CP_ISOARABIC = 28596	'ISO 8859-6 Arabic  阿拉伯语
	CP_ISOGREEK = 28597		'ISO 8859-7 Greek 希腊语
	CP_ISOHEBREW = 28598	'ISO 8859-8 Hebrew; Hebrew (ISO-Visual) 希伯来语（视觉顺序）；ISO 8859-8-I是 希伯来语（逻辑顺序）
	CP_ISOTURKISH2 = 28599	'ISO 8859-9 Turkish 它把Latin-1的冰岛语字母换走，加入土耳其语字母
	CP_ISOESTONIAN = 28603	'ISO 8859-13 Estonian   波罗的语族
	CP_ISOLATIN9 = 28605	'ISO 8859-15 Latin 9 西欧语言，加入Latin-1欠缺的芬兰语字母和大写法语重音字母，以及欧元（�）符号。
	CP_HEBREWLOG = 38598	'ISO 8859-8 Hebrew; Hebrew (ISO-Logical)
	CP_USER = 50000
	CP_AUTOALL = 50001
	CP_JAPANNHK = 50220		'ISO 2022 Japanese with no halfwidth Katakana; Japanese (JIS)
	CP_JAPANESC = 50221		'ISO 2022 Japanese with halfwidth Katakana; Japanese (JIS-Allow 1 byte Kana)
	CP_JAPANISO = 50222		'ISO 2022 Japanese JIS X 0201-1989; Japanese (JIS-Allow 1 byte Kana - SO/SI)
	CP_KOREAISO = 50225		'ISO 2022 Korean
	CP_TAIWANISO = 50227	'ISO 2022 Simplified Chinese; Chinese Simplified (ISO 2022)
	CP_CHINAISO = 50229		'ISO 2022 Traditional Chinese
	CP_AUTOJAPAN = 50932
	CP_AUTOCHINA = 50936	'EBCDIC Simplified Chinese
	CP_AUTOKOREA = 50949
	CP_AUTOTAIWAN = 50950
	CP_AUTORUSSIAN = 51251
	CP_AUTOGREEK = 51253
	CP_AUTOARABIC = 51256
	CP_JAPANEUC = 51932		'EUC Japanese
	CP_CHINAEUC = 51936		'EUC Simplified Chinese; Chinese Simplified (EUC)
	CP_KOREAEUC = 51949		'EUC Korean
	CP_TAIWANEUC = 51950	'EUC Traditional Chinese
	CP_CHINAHZ = 52936		'HZ-GB2312 Simplified Chinese; Chinese Simplified (HZ)
	CP_GB18030 = 54936		'Windows XP and later: GB18030 Simplified Chinese (4 byte); Chinese Simplified (GB18030)
' Windows UNICODE CODE PAGES
	CP_UTF7 = 65000
	CP_UTF8 = 65001
	CP_UTF32LE = 65005  'Unicode (UTF-32 LE)
	CP_UTF32BE = 65006	'Unicode (UTF-32 Big-Endian)
End Enum

'代码页转换
Private Declare Function MultiByteToWideChar Lib "kernel32.dll" ( _
	ByVal CodePage As Long, _
	ByVal dwFlags As Long, _
	ByRef lpMultiByteStr As Any, _
	ByVal cchMultiByte As Long, _
	ByVal lpWideCharStr As Long, _
	ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" ( _
	ByVal CodePage As Long, _
	ByVal dwFlags As Long, _
	ByVal lpWideCharStr As Long, _
	ByVal cchWideChar As Long, _
	ByRef lpMultiByteStr As Any, _
	ByVal cchMultiByte As Long, _
	ByVal lpDefaultChar As Long, _
	ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GetACP Lib "kernel32.dll" () As Long
'内存复制和比较函数
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
	ByRef Destination As Any, _
	ByRef Source As Any, _
	ByVal Length As Long)
'用于文本框查找定位函数
Private Declare Function SendMessageLNG Lib "user32.dll" Alias "SendMessage" ( _
	ByVal hwnd As Long, _
	ByVal wMsg As Long, _
	ByVal wParam As Long, _
	ByVal lParam As Long) As Long
'用于返回焦点控件的句柄
Private Declare Function GetFocus Lib "user32.dll" () As Long
'用于返回控件ID的句柄
Private Declare Function GetDlgItem Lib "user32.dll" ( _
	ByVal hDlg As Long, _
	ByVal nIDDlgItem As Long) As Long

Private MsgList() As String,SetList() As String,RegExp As Object


' 主程序
Sub Main
	Dim Obj As Object,Temp As String
	'检测系统语言
	On Error Resume Next
	Set Obj = CreateObject("WScript.Shell")
	If Obj Is Nothing Then
		MsgBox Err.Description & " - " & "WScript.Shell",vbInformation
		Exit Sub
	End If
	Temp = "HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language\Default"
	Temp = Obj.RegRead(Temp)
	If Temp = "" Then
		Temp = "HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language\InstallLanguage"
		Temp = Obj.RegRead(Temp)
		If Err.Source = "WshShell.RegRead" Then
			MsgBox Err.Description,vbInformation
			Exit Sub
		End If
	End If
	Set Obj = Nothing
	'检测 VBScript.RegExp 是否存在
	Set RegExp = CreateObject("VBScript.RegExp")
	If RegExp Is Nothing Then
		MsgBox(Err.Description & " - " & "VBScript.RegExp",vbInformation)
		Exit Sub
	End If
	On Error GoTo SysErrorMsg
	SetList = SplitArgument(Command$,14)
	If SetList(13) <> "" Then
		If StrToLong(SetList(13),3) = 3 Then Temp = ReSplit(SetList(13),";")(0)
	End If
	If UCase(Temp) <> Right$("0" & Hex$(PSL.Option(pslOptionSystemLanguage)),4) Then
		Temp = Right$("0" & Hex$(PSL.Option(pslOptionSystemLanguage)),4)
	End If
	If GetMsgList(MsgList,Temp) = False Then GoTo SysErrorMsg
	If StrToLong(SetList(10)) = 0 Then
		ConvertByUI(SetList)
	Else
		ConvertByNOUI(SetList)
	End If
	Exit Sub
	SysErrorMsg:
	If Err.Source <> "ExitSub" Then Call sysErrorMassage(Err,0)
End Sub


'命令行方式，返回转换结果
Public Function ConvertByNOUI(SetList() As String) As String
	Dim i As Integer
	If SetList(0) = "" And SetList(1) = "" Then Exit Function
	If StrToLong(SetList(11)) = 0 Then
		If StrToLong(SetList(8)) = 0 Then
			If StrToLong(SetList(4)) = 0 Then
				ConvertByNOUI = SetList(0)
			Else
				Select Case StrToLong(SetList(12))
				Case 0
					ConvertByNOUI = Convert(SetList(0))
				Case 1
					ConvertByNOUI = Convert(Replace$(SetList(0),vbCrLf,"\r"))
				Case 2
					ConvertByNOUI = Convert(Replace$(SetList(0),vbCrLf,"\n"))
				End Select
			End If
			Select Case StrToLong(SetList(5))
			Case 0
				ConvertByNOUI = Str2Hex(ConvertByNOUI,StrToLong(SetList(2)),0,0)
				If StrToLong(SetList(6)) = 1 Then ConvertByNOUI = SeparatHex(ConvertByNOUI)
			Case 1
				ConvertByNOUI = Str2HexEsc(ConvertByNOUI,StrToLong(SetList(2)),StrToLong(SetList(6)))
			Case 2
				ConvertByNOUI = Str2URLEsc(ConvertByNOUI,StrToLong(SetList(2)),StrToLong(SetList(6)))
			Case 3
				ConvertByNOUI = Str2HTMLEsc(ConvertByNOUI,StrToLong(SetList(2)),StrToLong(SetList(6)))
			Case 4
				ConvertByNOUI = Str2ISOEsc(ConvertByNOUI,StrToLong(SetList(2)),StrToLong(SetList(6)))
			Case 5
				ConvertByNOUI = Base64Encode(ConvertByNOUI,StrToLong(SetList(2)))
			End Select
		Else
			If StrToLong(SetList(4)) = 0 Then
				ConvertByNOUI = Replace$(Replace$(SetList(0)," ",""),vbCrLf,"")
			Else
				ConvertByNOUI = SetList(0)
			End If
			If StrToLong(SetList(5)) > 0 Then i = StrToLong(SetList(6))
			Select Case CheckHex(ConvertByNOUI,StrToLong(SetList(2)),StrToLong(SetList(5)),i)
			Case 0
				Select Case StrToLong(SetList(5))
				Case 0
					ConvertByNOUI = ByteToString(HexStr2Bytes(ConvertByNOUI),StrToLong(SetList(2)))
				Case 1
					ConvertByNOUI = HexEsc2Str(ConvertByNOUI,StrToLong(SetList(2)))
				Case 2
					ConvertByNOUI = URLEsc2Str(ConvertByNOUI,StrToLong(SetList(2)))
				Case 3
					ConvertByNOUI = HTMLEsc2Str(ConvertByNOUI,StrToLong(SetList(2)))
				Case 4
					ConvertByNOUI = ISOEsc2Str(ConvertByNOUI,StrToLong(SetList(2)))
				Case 5
					ConvertByNOUI = Base64Decode(ConvertByNOUI,StrToLong(SetList(2)))
				End Select
				If StrToLong(SetList(4)) = 1 Then
					Select Case StrToLong(SetList(12))
					Case 0
						ConvertByNOUI = ReConvert(ConvertByNOUI)
					Case 1
						ConvertByNOUI = ReConvert(Replace$(ConvertByNOUI,vbCrLf,"\r"))
					Case 2
						ConvertByNOUI = ReConvert(Replace$(ConvertByNOUI,vbCrLf,"\n"))
					End Select
				End If
			Case 1
				PSL.Output MsgList(63)
			Case 2
				PSL.Output MsgList(64)
			Case 3
				PSL.Output MsgList(65)
			Case 4
				PSL.Output MsgList(66)
			End Select
		End If
	Else
		If StrToLong(SetList(8)) = 0 Then
			If StrToLong(SetList(4)) = 0 Then
				ConvertByNOUI = SetList(1)
			Else
				Select Case StrToLong(SetList(12))
				Case 0
					ConvertByNOUI = Convert(SetList(1))
				Case 1
					ConvertByNOUI = Convert(Replace$(SetList(1),vbCrLf,"\r"))
				Case 2
					ConvertByNOUI = Convert(Replace$(SetList(1),vbCrLf,"\n"))
				End Select
			End If
			Select Case StrToLong(SetList(5))
			Case 0
				i = StrHexLength(SetList(0),StrToLong(SetList(9)),StrToLong(SetList(4)))
				ConvertByNOUI = Str2Hex(ConvertByNOUI,StrToLong(SetList(3)),StrToLong(SetList(9)),i)
				If StrToLong(SetList(6)) = 1 Then ConvertByNOUI = SeparatHex(ConvertByNOUI)
			Case 1
				ConvertByNOUI = Str2HexEsc(ConvertByNOUI,StrToLong(SetList(3)),StrToLong(SetList(6)))
			Case 2
				ConvertByNOUI = Str2URLEsc(ConvertByNOUI,StrToLong(SetList(3)),StrToLong(SetList(6)))
			Case 3
				ConvertByNOUI = Str2HTMLEsc(ConvertByNOUI,StrToLong(SetList(3)),StrToLong(SetList(6)))
			Case 4
				ConvertByNOUI = Str2ISOEsc(ConvertByNOUI,StrToLong(SetList(3)),StrToLong(SetList(6)))
			Case 5
				ConvertByNOUI = Base64Encode(ConvertByNOUI,StrToLong(SetList(3)))
			End Select
		Else
			If StrToLong(SetList(4)) = 0 Then
				ConvertByNOUI = Replace$(Replace$(SetList(1)," ",""),vbCrLf,"")
			Else
				ConvertByNOUI = SetList(1)
			End If
			If StrToLong(SetList(5)) > 0 Then i = StrToLong(SetList(6))
			Select Case CheckHex(ConvertByNOUI,StrToLong(SetList(3)),StrToLong(SetList(5)),i)
			Case 0
				Select Case StrToLong(SetList(5))
				Case 0
					ConvertByNOUI = ByteToString(HexStr2Bytes(ConvertByNOUI),StrToLong(SetList(3)))
				Case 1
					ConvertByNOUI = HexEsc2Str(ConvertByNOUI,StrToLong(SetList(3)))
				Case 2
					ConvertByNOUI = URLEsc2Str(ConvertByNOUI,StrToLong(SetList(3)))
				Case 3
					ConvertByNOUI = HTMLEsc2Str(ConvertByNOUI,StrToLong(SetList(3)))
				Case 4
					ConvertByNOUI = ISOEsc2Str(ConvertByNOUI,StrToLong(SetList(3)))
				Case 5
					ConvertByNOUI = Base64Decode(ConvertByNOUI,StrToLong(SetList(3)))
				End Select
				If StrToLong(SetList(4)) = 1 Then
					Select Case StrToLong(SetList(12))
					Case 0
						ConvertByNOUI = ReConvert(ConvertByNOUI)
					Case 1
						ConvertByNOUI = ReConvert(Replace$(ConvertByNOUI,vbCrLf,"\r"))
					Case 2
						ConvertByNOUI = ReConvert(Replace$(ConvertByNOUI,vbCrLf,"\n"))
					End Select
				End If
			Case 1
				PSL.Output MsgList(63)
			Case 2
				PSL.Output MsgList(64)
			Case 3
				PSL.Output MsgList(65)
			Case 4
				PSL.Output MsgList(66)
			End Select
		End If
	End If
End Function


'对话框方式
Public Function ConvertByUI(SetList() As String) As String
	Begin Dialog UserDialog 790,378,Replace$(Replace$(MsgList(39),"%v",Version),"%b",Build),.ConvertByUIDlgFunc ' %GRID:10,7,1,1
		CheckBox 690,7,30,14,"",.CheckBox
		CheckBox 10,7,110,21,"",.sLineBreakBox,1
		CheckBox 10,7,110,21,"",.tLineBreakBox,1
		TextBox 0,0,0,21,.SuppValueBox
		PushButton 680,7,90,21,"",.TempButton

		OptionGroup .OptionGroup
			OptionButton 230,7,170,21,MsgList(40),.SourceButton
			OptionButton 460,7,170,21,MsgList(41),.TransButton

		Text 10,35,130,14,MsgList(42),.sStrText
		Text 10,35,130,14,MsgList(42),.tStrText
		DropListBox 150,31,150,21,MsgList(),.sModeList
		DropListBox 150,31,150,21,MsgList(),.tModeList
		CheckBox 320,35,180,14,MsgList(43),.sSeparatHexBox
		CheckBox 320,35,180,14,MsgList(43),.tSeparatHexBox
		CheckBox 320,35,350,14,MsgList(44),.sHEXASCIICharBox
		CheckBox 320,35,350,14,MsgList(44),.tHEXASCIICharBox
		CheckBox 320,35,350,14,MsgList(45),.sURLLegalCharBox
		CheckBox 320,35,350,14,MsgList(45),.tURLLegalCharBox
		CheckBox 320,35,350,14,MsgList(44),.sHTMLASCIICharBox
		CheckBox 320,35,350,14,MsgList(44),.tHTMLASCIICharBox
		CheckBox 320,35,350,14,MsgList(44),.sISOASCIICharBox
		CheckBox 320,35,350,14,MsgList(44),.tISOASCIICharBox
		DropListBox 510,31,160,21,MsgList(),.FillList
		TextBox 10,56,660,119,.sStrBox,1
		TextBox 10,56,660,119,.tStrBox,1
		PushButton 680,56,100,28,MsgList(46),.Str2CodeButton
		PushButton 680,84,100,28,MsgList(47),.StrCopyButton
		PushButton 680,112,100,28,MsgList(48),.StrPasteButton
		PushButton 680,140,100,28,MsgList(49),.StrCleanButton

		Text 10,182,130,14,MsgList(50),.sCodeText
		Text 10,182,130,14,MsgList(50),.tCodeText
		DropListBox 150,178,230,21,MsgList(),.sCPNameList
		DropListBox 150,178,230,21,MsgList(),.sCPValueList
		DropListBox 150,178,230,21,MsgList(),.tCPNameList
		DropListBox 150,178,230,21,MsgList(),.tCPValueList
		CheckBox 400,182,130,14,MsgList(51),.sConvertBox
		CheckBox 400,182,130,14,MsgList(51),.tConvertBox
		CheckBox 540,182,130,14,MsgList(52),.ModeBox
		TextBox 10,203,660,168,.sCodeBox,1
		TextBox 10,203,660,168,.tCodeBox,1
		PushButton 680,203,100,28,MsgList(53),.Code2StrButton
		PushButton 680,231,100,28,MsgList(54),.CodeCopyButton
		PushButton 680,259,100,28,MsgList(55),.CodePasteButton
		PushButton 680,287,100,28,MsgList(56),.CodeCleanButton
		PushButton 680,315,100,28,MsgList(72),.LangButton
		PushButton 680,343,100,28,MsgList(57),.AboutButton
		CancelButton 680,7,100,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then Exit Function
	If dlg.OptionGroup = 0 Then
		ConvertByUI = IIf(StrToLong(SetList(6)) = 0,dlg.sCodeBox,dlg.sStrBox)
	Else
		ConvertByUI = IIf(StrToLong(SetList(6)) = 0,dlg.tCodeBox,dlg.tStrBox)
	End If
End Function


'主程序对话框函数
Private Function ConvertByUIDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,Temp As String
	DlgValue "CheckBox",0
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		DlgVisible "TempButton",False
		DlgVisible "CancelButton",False
		DlgVisible "sCPValueList",False
		DlgVisible "tCPValueList",False
		DlgVisible "CheckBox",False
		DlgVisible "sLineBreakBox",False
		DlgVisible "tLineBreakBox",False
		DlgVisible "SuppValueBox",False
		DlgValue "CheckBox",0
		DlgText "SuppValueBox",CStr$(SuppValue)

		ReDim TempList(23) As String
		TempList(0) = CStr$(CP_OEMCP)			'系统默认
		TempList(1) = CStr$(CP_MACCP)
		TempList(2) = CStr$(CP_THREAD_ACP)
		TempList(3) = CStr$(CP_WESTEUROPE)		'拉丁文 1 (ANSI) = 1252
		TempList(4) = CStr$(CP_EASTEUROPE)		'拉丁文 2 (中欧) = 1250
		TempList(5) = CStr$(CP_RUSSIAN)			'西里尔文 (斯拉夫) = 1251
		TempList(6) = CStr$(CP_GREEK)			'希腊文 = 1253
		TempList(7) = CStr$(CP_TURKISH)			'拉丁文 5 (土耳其) = 1254
		TempList(8) = CStr$(CP_HEBREW)			'希伯来文 = 1255
		TempList(9) = CStr$(CP_ARABIC)			'阿拉伯文 = 1256
		TempList(10) = CStr$(CP_BALTIC)			'波罗的海文 = 1257
		TempList(11) = CStr$(CP_VIETNAMESE)		'越南文 = 1258
		TempList(12) = CStr$(CP_JAPAN)			'日文 = 932
		TempList(13) = CStr$(CP_CHINA)			'简体中文 = 936
		TempList(14) = CStr$(CP_GB18030)		'简体中文 = 54936
		TempList(15) = CStr$(CP_KOREA)			'韩文 = 949
		TempList(16) = CStr$(CP_TAIWAN) 		'繁体中文 = 950
		TempList(17) = CStr$(CP_THAI)			'泰文 = 874
		TempList(18) = CStr$(CP_UTF7)			'UTF-7 = 65000
		TempList(19) = CStr$(CP_UTF8)			'UTF-8 = 65001
		TempList(20) = CStr$(CP_UNICODELITTLE)	'UnicodeLE = 1200
		TempList(21) = CStr$(CP_UNICODEBIG)		'UnicodeBE = 1201
		TempList(22) = CStr$(CP_UTF32LE)		'UnicodeLE = 65005
		TempList(23) = CStr$(CP_UTF32BE)		'UnicodeBE = 65006
		DlgListBoxArray "sCPValueList",TempList()
		DlgListBoxArray "tCPValueList",TempList()
		For i = 0 To 23
			TempList(i) = TempList(i) & " - " & MsgList(i + 14)
		Next i
		DlgListBoxArray "sCPNameList",TempList()
		DlgListBoxArray "tCPNameList",TempList()
		TempList = ReSplit(MsgList(38),ItemJoinStr)
		DlgListBoxArray "sModeList",TempList()
		DlgListBoxArray "tModeList",TempList()
		ReDim TempList(4) As String
		For i = 0 To 4
			TempList(i) = MsgList(i + 58)
		Next i
		DlgListBoxArray "FillList",TempList()

		'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
		If StrToLong(SetList(8)) = 0 Then
			SetTextBoxLength GetDlgItem(SuppValue,DlgControlId("sStrBox")),Len(SetList(0)),False
			DlgText "sStrBox",SetList(0)
			SetTextBoxLength GetDlgItem(SuppValue,DlgControlId("tStrBox")),Len(SetList(1)),False
			DlgText "tStrBox",SetList(1)
		Else
			SetTextBoxLength GetDlgItem(SuppValue,DlgControlId("sCodeBox")),Len(SetList(0)),False
			DlgText "sCodeBox",SetList(0)
			SetTextBoxLength GetDlgItem(SuppValue,DlgControlId("tCodeBox")),Len(SetList(1)),False
			DlgText "tCodeBox",SetList(1)
		End If
		DlgText "sCPValueList",SetList(2)
		DlgText "tCPValueList",SetList(3)
		DlgValue "sConvertBox",IIf(Command = "",1,StrToLong(SetList(4)))
		DlgValue "tConvertBox",IIf(Command = "",1,StrToLong(SetList(4)))
		DlgValue "sModeList",StrToLong(SetList(5))
		DlgValue "tModeList",StrToLong(SetList(5))
		Select Case DlgValue("sModeList")
		Case 0
			DlgValue "sSeparatHexBox",IIf(Command = "",1,StrToLong(SetList(6)))
		Case 1
			DlgValue "sHEXASCIICharBox",StrToLong(SetList(6))
		Case 2
			DlgValue "sURLLegalCharBox",StrToLong(SetList(6))
		Case 3
			DlgValue "sHTMLASCIICharBox",StrToLong(SetList(6))
		Case 4
			DlgValue "sISOASCIICharBox",StrToLong(SetList(6))
		End Select
		Select Case DlgValue("tModeList")
		Case 0
			DlgValue "tSeparatHexBox",IIf(Command = "",1,StrToLong(SetList(6)))
		Case 1
			DlgValue "tHEXASCIICharBox",StrToLong(SetList(6))
		Case 2
			DlgValue "tURLLegalCharBox",StrToLong(SetList(6))
		Case 3
			DlgValue "tHTMLASCIICharBox",StrToLong(SetList(6))
		Case 4
			DlgValue "tISOASCIICharBox",StrToLong(SetList(6))
		End Select
		DlgValue "ModeBox",IIf(Command = "",1,StrToLong(SetList(7)))
		DlgValue "FillList",StrToLong(SetList(9))
		DlgValue "OptionGroup",StrToLong(SetList(11))
		DlgValue "sLineBreakBox",StrToLong(SetList(12))
		DlgValue "tLineBreakBox",StrToLong(SetList(12))
		TempList = ReSplit(MsgList(71),";")
		DlgText "sConvertBox",Replace$(MsgList(51),"%s",TempList(DlgValue("sLineBreakBox")))
		DlgText "tConvertBox",Replace$(MsgList(51),"%s",TempList(DlgValue("tLineBreakBox")))
		If DlgText("sCPValueList") = "" Then DlgText "sCPValueList",CStr$(GetACP)
		If DlgText("tCPValueList") = "" Then DlgText "tCPValueList",CStr$(GetACP)
		DlgValue "sCPNameList",DlgValue("sCPValueList")
		DlgValue "tCPNameList",DlgValue("tCPValueList")

		If DlgValue("OptionGroup") = 0 Then
			DlgVisible "tModeList",False
			DlgVisible "tSeparatHexBox",False
			DlgVisible "tConvertBox",False
			DlgVisible "tStrText",False
			DlgVisible "tCodeText",False
			DlgVisible "tStrBox",False
			DlgVisible "tCodeBox",False
			DlgVisible "tCPNameList",False
			DlgVisible "tSeparatHexBox",False
			DlgVisible "tHEXASCIICharBox",False
			DlgVisible "tURLLegalCharBox",False
			DlgVisible "tHTMLASCIICharBox",False
			DlgVisible "tISOASCIICharBox",False
			DlgVisible "sSeparatHexBox",IIf(DlgValue("sModeList") = 0,True,False)
			DlgVisible "sHEXASCIICharBox",IIf(DlgValue("sModeList") = 1,True,False)
			DlgVisible "sURLLegalCharBox",IIf(DlgValue("sModeList") = 2,True,False)
			DlgVisible "sHTMLASCIICharBox",IIf(DlgValue("sModeList") = 3,True,False)
			DlgVisible "sISOASCIICharBox",IIf(DlgValue("sModeList") = 4,True,False)
			DlgVisible "FillList",False
		Else
			DlgVisible "sModeList",False
			DlgVisible "sSeparatHexBox",False
			DlgVisible "sConvertBox",False
			DlgVisible "sStrText",False
			DlgVisible "sCodeText",False
			DlgVisible "sStrBox",False
			DlgVisible "sCodeBox",False
			DlgVisible "sCPNameList",False
			DlgVisible "sSeparatHexBox",False
			DlgVisible "sHEXASCIICharBox",False
			DlgVisible "sURLLegalCharBox",False
			DlgVisible "sHTMLASCIICharBox",False
			DlgVisible "sISOASCIICharBox",False
			DlgVisible "tSeparatHexBox",IIf(DlgValue("tModeList") = 0,True,False)
			DlgVisible "tHEXASCIICharBox",IIf(DlgValue("tModeList") = 1,True,False)
			DlgVisible "tURLLegalCharBox",IIf(DlgValue("tModeList") = 2,True,False)
			DlgVisible "tHTMLASCIICharBox",IIf(DlgValue("tModeList") = 3,True,False)
			DlgVisible "tISOASCIICharBox",IIf(DlgValue("tModeList") = 4,True,False)
			DlgVisible "FillList",IIf(DlgValue("tModeList") = 0,True,False)
		End If

		DlgText "sStrText",Replace$(Replace$(MsgList(42),"%s","0"),"%d","00")
		DlgText "sCodeText",Replace$(Replace$(MsgList(50),"%s","0"),"%d","00")
		DlgText "tStrText",Replace$(Replace$(MsgList(42),"%s","0"),"%d","00")
		DlgText "tCodeText",Replace$(Replace$(MsgList(50),"%s","0"),"%d","00")
		If DlgText("sStrBox") <> "" Then
			If DlgValue("sConvertBox") = 0 Then
				Temp = DlgText("sStrBox")
			Else
				Select Case DlgValue("sLineBreakBox")
				Case 0
					Temp = Convert(DlgText("sStrBox"))
				Case 1
					Temp = Convert(Replace$(DlgText("sStrBox"),vbCrLf,"\r"))
				Case 2
					Temp = Convert(Replace$(DlgText("sStrBox"),vbCrLf,"\n"))
				End Select
			End If
			'显示字符数和字节数
			i = Len(Temp)
			DlgText "sStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
			i = StrHexLength(Temp,CLng(Trim$(DlgText("sCPValueList"))),0)
			DlgText "sCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
			If DlgValue("ModeBox") = 1 Then
				Select Case DlgValue("sModeList")
				Case 0
					Temp = Str2Hex(Temp,CLng(Trim$(DlgText("sCPValueList"))),0,0)
					If DlgValue("sSeparatHexBox") = 1 Then Temp = SeparatHex(Temp)
				Case 1
					Temp = Str2HexEsc(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sHEXASCIICharBox"))
				Case 2
					Temp = Str2URLEsc(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sURLLegalCharBox"))
				Case 3
					Temp = Str2HTMLEsc(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sHTMLASCIICharBox"))
				Case 4
					Temp = Str2ISOEsc(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sISOASCIICharBox"))
				Case 5
					Temp = Base64Encode(Temp,CLng(Trim$(DlgText("sCPValueList"))))
				End Select
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(SuppValue,DlgControlId("sCodeBox")),Len(Temp),False
				DlgText "sCodeBox",Temp
			End If
		ElseIf DlgText("sCodeBox") <> "" Then
			If DlgValue("ModeBox") = 1 Then
				If DlgValue("sModeList") = 0 Then
					Temp = Replace$(Replace$(DlgText("sCodeBox")," ",""),vbCrLf,"")
				Else
					Temp = DlgText("sCodeBox")
				End If
				i = 0
				Select Case DlgValue("sModeList")
				Case 1
					i = DlgValue("sHEXASCIICharBox")
				Case 2
					i = DlgValue("sURLLegalCharBox")
				Case 3
					i = DlgValue("sHTMLASCIICharBox")
				Case 3
					i = DlgValue("sISOASCIICharBox")
				End Select
				Select Case CheckHex(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sModeList"),i)
				Case 0
					Select Case DlgValue("sModeList")
					Case 0
						Temp = ByteToString(HexStr2Bytes(Temp),CLng(Trim$(DlgText("sCPValueList"))))
					Case 1
						Temp = HexEsc2Str(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					Case 2
						Temp = URLEsc2Str(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					Case 3
						Temp = HTMLEsc2Str(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					Case 4
						Temp = ISOEsc2Str(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					Case 5
						Temp = Base64Decode(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					End Select
					'显示字符数和字节数
					i = Len(Temp)
					DlgText "sStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					i = StrHexLength(Temp,CLng(Trim$(DlgText("sCPValueList"))),0)
					DlgText "sCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					If DlgValue("sConvertBox") = 1 Then
						Select Case DlgValue("sLineBreakBox")
						Case 0
							Temp = ReConvert(Temp)
						Case 1
							Temp = ReConvert(Replace$(Temp,vbCrLf,"\r"))
						Case 2
							Temp = ReConvert(Replace$(Temp,vbCrLf,"\n"))
						End Select
					End If
					'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
					SetTextBoxLength GetDlgItem(SuppValue,DlgControlId("sStrBox")),Len(Temp),False
					DlgText "sStrBox",Temp
				Case 1
					MsgBox(MsgList(63),vbOkOnly+vbInformation,MsgList(0))
				Case 2
					MsgBox(MsgList(64),vbOkOnly+vbInformation,MsgList(0))
				Case 3
					MsgBox(MsgList(65),vbOkOnly+vbInformation,MsgList(0))
				Case 4
					MsgBox(MsgList(66),vbOkOnly+vbInformation,MsgList(0))
				End Select
			End If
		End If
		If DlgText("tStrBox") <> "" Then
			If DlgValue("tConvertBox") = 0 Then
				Temp = DlgText("tStrBox")
			Else
				Select Case DlgValue("tLineBreakBox")
				Case 0
					Temp = Convert(DlgText("tStrBox"))
				Case 1
					Temp = Convert(Replace$(DlgText("tStrBox"),vbCrLf,"\r"))
				Case 2
					Temp = Convert(Replace$(DlgText("tStrBox"),vbCrLf,"\n"))
				End Select
			End If
			'显示字符数和字节数
			i = Len(Temp)
			DlgText "tStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
			i = StrHexLength(Temp,CLng(Trim$(DlgText("tCPValueList"))),0)
			DlgText "tCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
			If DlgValue("ModeBox") = 1 Then
				Select Case DlgValue("tModeList")
				Case 0
					i = StrHexLength(DlgText("sStrBox"),CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sConvertBox"))
					Temp = Str2Hex(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("FillList"),i)
					'显示原文的字符数和字节数
					If DlgValue("FillList") <> 0 Then
						DlgText "tCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(Len(Temp) \ 2)),"%d",FormatHexStr(Hex$(i),2))
					End If
					If DlgValue("tSeparatHexBox") = 1 Then Temp = SeparatHex(Temp)
				Case 1
					Temp = Str2HexEsc(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tHEXASCIICharBox"))
				Case 2
					Temp = Str2URLEsc(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tURLLegalCharBox"))
				Case 3
					Temp = Str2HTMLEsc(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tHTMLASCIICharBox"))
				Case 4
					Temp = Str2ISOEsc(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tISOASCIICharBox"))
				Case 5
					Temp = Base64Encode(Temp,CLng(Trim$(DlgText("tCPValueList"))))
				End Select
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(SuppValue,DlgControlId("tCodeBox")),Len(Temp),False
				DlgText "tCodeBox",Temp
			End If
		ElseIf DlgText("tCodeBox") <> "" Then
			If DlgValue("ModeBox") = 1 Then
				If DlgValue("tModeList") = 0 Then
					Temp = Replace$(Replace$(DlgText("tCodeBox")," ",""),vbCrLf,"")
				Else
					Temp = DlgText("tCodeBox")
				End If
				i = 0
				Select Case DlgValue("tModeList")
				Case 1
					i = DlgValue("tHEXASCIICharBox")
				Case 2
					i = DlgValue("tURLLegalCharBox")
				Case 3
					i = DlgValue("tHTMLASCIICharBox")
				Case 3
					i = DlgValue("tISOASCIICharBox")
				End Select
				Select Case CheckHex(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tModeList"),i)
				Case 0
					Select Case DlgValue("tModeList")
					Case 0
						Temp = ByteToString(HexStr2Bytes(Temp),CLng(Trim$(DlgText("tCPValueList"))))
					Case 1
						Temp = HexEsc2Str(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					Case 2
						Temp = URLEsc2Str(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					Case 3
						Temp = HTMLEsc2Str(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					Case 4
						Temp = ISOEsc2Str(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					Case 5
						Temp = Base64Decode(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					End Select
					'显示字符数和字节数
					i = Len(Temp)
					DlgText "tStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					i = StrHexLength(Temp,CLng(Trim$(DlgText("tCPValueList"))),0)
					DlgText "tCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					If DlgValue("tConvertBox") = 1 Then
						Select Case DlgValue("tLineBreakBox")
						Case 0
							Temp = ReConvert(Temp)
						Case 1
							Temp = ReConvert(Replace$(Temp,vbCrLf,"\r"))
						Case 2
							Temp = ReConvert(Replace$(Temp,vbCrLf,"\n"))
						End Select
					End If
					'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
					SetTextBoxLength GetDlgItem(SuppValue,DlgControlId("tStrBox")),Len(Temp),False
					DlgText "tStrBox",Temp
				Case 1
					MsgBox(MsgList(63),vbOkOnly+vbInformation,MsgList(0))
				Case 2
					MsgBox(MsgList(64),vbOkOnly+vbInformation,MsgList(0))
				Case 3
					MsgBox(MsgList(65),vbOkOnly+vbInformation,MsgList(0))
				Case 4
					MsgBox(MsgList(66),vbOkOnly+vbInformation,MsgList(0))
				End Select
			End If
		End If
	Case 2 ' 数值更改或者按下了按钮
		If DlgItem$ = "CancelButton" Then Exit Function
		ConvertByUIDlgFunc = True ' 防止按下按钮关闭对话框窗口
		Select Case DlgItem$
		Case "OptionGroup"
			If DlgValue("OptionGroup") = 0 Then
				DlgVisible "sModeList",True
				DlgVisible "sSeparatHexBox",True
				DlgVisible "sConvertBox",True
				DlgVisible "sStrText",True
				DlgVisible "sCodeText",True
				DlgVisible "sStrBox",True
				DlgVisible "sCodeBox",True
				DlgVisible "sCPNameList",True
				DlgVisible "sSeparatHexBox",IIf(DlgValue("sModeList") = 0,True,False)
				DlgVisible "sHEXASCIICharBox",IIf(DlgValue("sModeList") = 1,True,False)
				DlgVisible "sURLLegalCharBox",IIf(DlgValue("sModeList") = 2,True,False)
				DlgVisible "sHTMLASCIICharBox",IIf(DlgValue("sModeList") = 3,True,False)
				DlgVisible "sISOASCIICharBox",IIf(DlgValue("sModeList") = 4,True,False)

				DlgVisible "tModeList",False
				DlgVisible "tSeparatHexBox",False
				DlgVisible "tConvertBox",False
				DlgVisible "tStrText",False
				DlgVisible "tCodeText",False
				DlgVisible "tStrBox",False
				DlgVisible "tCodeBox",False
				DlgVisible "tCPNameList",False
				DlgVisible "tSeparatHexBox",False
				DlgVisible "tHEXASCIICharBox",False
				DlgVisible "tURLLegalCharBox",False
				DlgVisible "tHTMLASCIICharBox",False
				DlgVisible "tISOASCIICharBox",False
				DlgVisible "FillList",False
			Else
				DlgVisible "sModeList",False
				DlgVisible "sSeparatHexBox",False
				DlgVisible "sConvertBox",False
				DlgVisible "sStrText",False
				DlgVisible "sCodeText",False
				DlgVisible "sStrBox",False
				DlgVisible "sCodeBox",False
				DlgVisible "sCPNameList",False
				DlgVisible "sSeparatHexBox",False
				DlgVisible "sHEXASCIICharBox",False
				DlgVisible "sURLLegalCharBox",False
				DlgVisible "sHTMLASCIICharBox",False
				DlgVisible "sISOASCIICharBox",False

				DlgVisible "tModeList",True
				DlgVisible "tSeparatHexBox",True
				DlgVisible "tConvertBox",True
				DlgVisible "tStrText",True
				DlgVisible "tCodeText",True
				DlgVisible "tStrBox",True
				DlgVisible "tCodeBox",True
				DlgVisible "tCPNameList",True
				DlgVisible "tSeparatHexBox",IIf(DlgValue("tModeList") = 0,True,False)
				DlgVisible "tHEXASCIICharBox",IIf(DlgValue("tModeList") = 1,True,False)
				DlgVisible "tURLLegalCharBox",IIf(DlgValue("tModeList") = 2,True,False)
				DlgVisible "tHTMLASCIICharBox",IIf(DlgValue("tModeList") = 3,True,False)
				DlgVisible "tISOASCIICharBox",IIf(DlgValue("tModeList") = 4,True,False)
				DlgVisible "FillList",IIf(DlgValue("tModeList") = 0,True,False)
			End If
		Case "AboutButton"
			MsgBox Replace$(Replace$(MsgList(12),"%v",Version),"%b",Build),vbOkOnly+vbInformation,MsgList(13)
		Case "StrCopyButton"
			If DlgValue("OptionGroup") = 0 Then
				If DlgText("sStrBox") = "" Then Exit Function
				Clipboard DlgText("sStrBox")
			Else
				If DlgText("tStrBox") = "" Then Exit Function
				Clipboard DlgText("tStrBox")
			End If
		Case "CodeCopyButton"
			If DlgValue("OptionGroup") = 0 Then
				If DlgText("sCodeBox") = "" Then Exit Function
				If DlgValue("sModeList") = 0 Then
					ReDim TempList(2) As String
					TempList(0) = MsgList(67)
					TempList(1) = MsgList(68)
					TempList(2) = MsgList(69)
				Else
					Clipboard DlgText("sCodeBox")
					Exit Function
				End If
			Else
				If DlgText("tCodeBox") = "" Then  Exit Function
				If DlgValue("tModeList") = 0 Then
					ReDim TempList(2) As String
					TempList(0) = MsgList(67)
					TempList(1) = MsgList(68)
					TempList(2) = MsgList(69)
				Else
					Clipboard DlgText("tCodeBox")
					Exit Function
				End If
			End If
			Select Case ShowPopupMenu(TempList,vbPopupUseRightButton)
			Case 0
				If DlgValue("OptionGroup") = 0 Then
					Clipboard DlgText("sCodeBox")
				Else
					Clipboard DlgText("tCodeBox")
				End If
			Case 1
				If DlgValue("OptionGroup") = 0 Then
					Select Case CLng(Trim$(DlgText("sCPValueList")))
					Case CP_UNICODELITTLE, CP_UNICODEBIG
						Clipboard DlgText("sCodeBox") & IIf(DlgValue("sSeparatHexBox") = 0,"0000"," 00 00")
					Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
						Clipboard DlgText("sCodeBox") & IIf(DlgValue("sSeparatHexBox") = 0,"00000000"," 00 00 00 00")
					Case Else
						Clipboard DlgText("sCodeBox") & IIf(DlgValue("sSeparatHexBox") = 0,"00"," 00")
					End Select
				Else
					Select Case CLng(Trim$(DlgText("tCPValueList")))
					Case CP_UNICODELITTLE, CP_UNICODEBIG
						Clipboard DlgText("tCodeBox") & IIf(DlgValue("tSeparatHexBox") = 0,"0000"," 00 00")
					Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
						Clipboard DlgText("tCodeBox") & IIf(DlgValue("tSeparatHexBox") = 0,"00000000"," 00 00 00 00")
					Case Else
						Clipboard DlgText("tCodeBox") & IIf(DlgValue("tSeparatHexBox") = 0,"00"," 00")
					End Select
				End If
			Case 2
				If DlgValue("OptionGroup") = 0 Then
					Select Case CLng(Trim$(DlgText("sCPValueList")))
					Case CP_UNICODELITTLE, CP_UNICODEBIG
						Clipboard IIf(DlgValue("sSeparatHexBox") = 0,"0000","00 00 ") & _
							DlgText("sCodeBox") & IIf(DlgValue("sSeparatHexBox") = 0,"0000"," 00 00")
					Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
						Clipboard IIf(DlgValue("sSeparatHexBox") = 0,"00000000","00 00 00 00 ") & _
							DlgText("sCodeBox") & IIf(DlgValue("sSeparatHexBox") = 0,"00000000"," 00 00 00 00")
					Case Else
						Clipboard IIf(DlgValue("sSeparatHexBox") = 0,"00","00 ") & _
							DlgText("sCodeBox") & IIf(DlgValue("sSeparatHexBox") = 0,"00"," 00")
					End Select
				Else
					Select Case CLng(Trim$(DlgText("tCPValueList")))
					Case CP_UNICODELITTLE, CP_UNICODEBIG
						Clipboard IIf(DlgValue("tSeparatHexBox") = 0,"0000","00 00 ") & _
							DlgText("tCodeBox") & IIf(DlgValue("tSeparatHexBox") = 0,"0000"," 00 00")
					Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
						Clipboard IIf(DlgValue("tSeparatHexBox") = 0,"00000000","00 00 00 00 ") & _
							DlgText("tCodeBox") & IIf(DlgValue("tSeparatHexBox") = 0,"00000000"," 00 00 00 00")
					Case Else
						Clipboard IIf(DlgValue("tSeparatHexBox") = 0,"00","00 ") & _
							DlgText("tCodeBox") & IIf(DlgValue("tSeparatHexBox") = 0,"00"," 00")
					End Select
				End If
			End Select
		Case "StrPasteButton"
			'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
			If DlgValue("OptionGroup") = 0 Then
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("sStrBox")),Len(Clipboard),True
				DlgText "sStrBox",Clipboard
				If DlgValue("ModeBox") = 0 Then Exit Function
				If DlgText("sStrBox") <> "" Then DlgItem$ = "Str2CodeButton"
			Else
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("tStrBox")),Len(Clipboard),True
				DlgText "tStrBox",Clipboard
				If DlgValue("ModeBox") = 0 Then Exit Function
				If DlgText("tStrBox") <> "" Then DlgItem$ = "Str2CodeButton"
			End If
		Case "CodePasteButton"
			'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
			If DlgValue("OptionGroup") = 0 Then
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("sCodeBox")),Len(Clipboard),True
				DlgText "sCodeBox",Clipboard
				If DlgValue("ModeBox") = 0 Then Exit Function
				If DlgText("sCodeBox") <> ""Then DlgItem$ = "Code2StrButton"
			Else
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("tCodeBox")),Len(Clipboard),True
				DlgText "tCodeBox",Clipboard
				If DlgValue("ModeBox") = 0 Then Exit Function
				If DlgText("tCodeBox") <> ""Then DlgItem$ = "Code2StrButton"
			End If
		Case "StrCleanButton"
			If DlgValue("OptionGroup") = 0 Then
				DlgText "sStrBox",""
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("sStrBox")),0,True
			Else
				DlgText "tStrBox",""
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("tStrBox")),0,True
			End If
		Case "CodeCleanButton"
			If DlgValue("OptionGroup") = 0 Then
				DlgText "sCodeBox",""
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("sCodeBox")),0,True
			Else
				DlgText "tCodeBox",""
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("tCodeBox")),0,True
			End If
		Case "sCPNameList"
			DlgValue "sCPValueList",DlgValue("sCPNameList")
			If DlgValue("ModeBox") = 0 Then Exit Function
			If DlgText("sStrBox") <> "" Then
				DlgItem$ = "Str2CodeButton"
			ElseIf DlgText("sCodeBox") <> ""Then
				DlgItem$ = "Code2StrButton"
			End If
		Case "tCPNameList"
			DlgValue "tCPValueList",DlgValue("tCPNameList")
			If DlgValue("ModeBox") = 0 Then Exit Function
			If DlgText("tStrBox") <> "" Then
				DlgItem$ = "Str2CodeButton"
			ElseIf DlgText("tCodeBox") <> ""Then
				DlgItem$ = "Code2StrButton"
			End If
		Case "sModeList"
			Select Case DlgValue("sModeList")
			Case 0
				DlgVisible "sSeparatHexBox",True
				DlgVisible "sHEXASCIICharBox",False
				DlgVisible "sURLLegalCharBox",False
				DlgVisible "sHTMLASCIICharBox",False
				DlgVisible "sISOASCIICharBox",False
			Case 1
				DlgVisible "sSeparatHexBox",False
				DlgVisible "sHEXASCIICharBox",True
				DlgVisible "sURLLegalCharBox",False
				DlgVisible "sHTMLASCIICharBox",False
				DlgVisible "sISOASCIICharBox",False
			Case 2
				DlgVisible "sSeparatHexBox",False
				DlgVisible "sHEXASCIICharBox",False
				DlgVisible "sURLLegalCharBox",True
				DlgVisible "sHTMLASCIICharBox",False
				DlgVisible "sISOASCIICharBox",False
			Case 3
				DlgVisible "sSeparatHexBox",False
				DlgVisible "sHEXASCIICharBox",False
				DlgVisible "sURLLegalCharBox",False
				DlgVisible "sHTMLASCIICharBox",True
				DlgVisible "sISOASCIICharBox",False
			Case 4
				DlgVisible "sSeparatHexBox",False
				DlgVisible "sHEXASCIICharBox",False
				DlgVisible "sURLLegalCharBox",False
				DlgVisible "sHTMLASCIICharBox",False
				DlgVisible "sISOASCIICharBox",True
			Case Else
				DlgVisible "sSeparatHexBox",False
				DlgVisible "sHEXASCIICharBox",False
				DlgVisible "sURLLegalCharBox",False
				DlgVisible "sHTMLASCIICharBox",False
				DlgVisible "sISOASCIICharBox",False
			End Select
			If DlgValue("ModeBox") = 0 Then Exit Function
			If DlgText("sStrBox") <> "" Then
				DlgItem$ = "Str2CodeButton"
			ElseIf DlgText("sCodeBox") <> ""Then
				DlgItem$ = "Code2StrButton"
			End If
		Case "tModeList"
			Select Case DlgValue("tModeList")
			Case 0
				DlgVisible "tSeparatHexBox",True
				DlgVisible "tHEXASCIICharBox",False
				DlgVisible "tURLLegalCharBox",False
				DlgVisible "tHTMLASCIICharBox",False
				DlgVisible "tISOASCIICharBox",False
				DlgVisible "FillList",IIf(DlgValue("OptionGroup") = 0,False,True)
			Case 1
				DlgVisible "tSeparatHexBox",False
				DlgVisible "tHEXASCIICharBox",True
				DlgVisible "tURLLegalCharBox",False
				DlgVisible "tHTMLASCIICharBox",False
				DlgVisible "tISOASCIICharBox",False
				DlgVisible "FillList",False
			Case 2
				DlgVisible "tSeparatHexBox",False
				DlgVisible "tHEXASCIICharBox",False
				DlgVisible "tURLLegalCharBox",True
				DlgVisible "tHTMLASCIICharBox",False
				DlgVisible "tISOASCIICharBox",False
				DlgVisible "FillList",False
			Case 3
				DlgVisible "tSeparatHexBox",False
				DlgVisible "tHEXASCIICharBox",False
				DlgVisible "tURLLegalCharBox",False
				DlgVisible "tHTMLASCIICharBox",True
				DlgVisible "tISOASCIICharBox",False
				DlgVisible "FillList",False
			Case 4
				DlgVisible "tSeparatHexBox",False
				DlgVisible "tHEXASCIICharBox",False
				DlgVisible "tURLLegalCharBox",False
				DlgVisible "tHTMLASCIICharBox",False
				DlgVisible "tISOASCIICharBox",True
				DlgVisible "FillList",False
			Case Else
				DlgVisible "tSeparatHexBox",False
				DlgVisible "tHEXASCIICharBox",False
				DlgVisible "tURLLegalCharBox",False
				DlgVisible "tHTMLASCIICharBox",False
				DlgVisible "tISOASCIICharBox",False
				DlgVisible "FillList",False
			End Select
			If DlgValue("ModeBox") = 0 Then Exit Function
			If DlgText("tStrBox") <> "" Then
				DlgItem$ = "Str2CodeButton"
			ElseIf DlgText("tCodeBox") <> ""Then
				DlgItem$ = "Code2StrButton"
			End If
		Case "sSeparatHexBox", "tSeparatHexBox"
			If DlgValue("OptionGroup") = 0 Then
				If DlgText("sCodeBox") = "" Then Exit Function
				If InStr(DlgText("sCodeBox")," ") Then
					DlgText "sCodeBox",StrListJoin(ReSplit(DlgText("sCodeBox")," "),"")
				Else
					DlgText "sCodeBox",SeparatHex(DlgText("sCodeBox"))
				End If
			Else
				If DlgText("tCodeBox") = "" Then Exit Function
				If InStr(DlgText("tCodeBox")," ") Then
					DlgText "tCodeBox",StrListJoin(ReSplit(DlgText("tCodeBox")," "),"")
				Else
					DlgText "tCodeBox",SeparatHex(DlgText("tCodeBox"))
				End If
			End If
			DlgValue "CheckBox",1
		Case "sHEXASCIICharBox","sURLLegalCharBox","sHTMLASCIICharBox","sISOASCIICharBox"
			If DlgValue("ModeBox") = 1 Then DlgItem$ = "Str2CodeButton"
		Case "tHEXASCIICharBox","tURLLegalCharBox","tHTMLASCIICharBox","tISOASCIICharBox"
			If DlgValue("ModeBox") = 1 Then DlgItem$ = "Str2CodeButton"
		Case "sConvertBox", "tConvertBox"
			If DlgValue("OptionGroup") = 0 Then
				If DlgText("sStrBox") = "" Then Exit Function
				If DlgValue(DlgItem$) = 0 Then
					Temp = DlgText("sStrBox")
				Else
					TempList = ReSplit(MsgList(70),";")
					i = ShowPopupMenu(TempList,vbPopupUseRightButton)
					If i > -1 Then
						DlgValue "sLineBreakBox",i
						TempList = ReSplit(MsgList(71),";")
						DlgText "sConvertBox",Replace$(MsgList(51),"%s",TempList(i))
					End If
					Select Case DlgValue("sLineBreakBox")
					Case 0
						Temp = Convert(DlgText("sStrBox"))
					Case 1
						Temp = Convert(Replace$(DlgText("sStrBox"),vbCrLf,vbCr))
					Case 2
						Temp = Convert(Replace$(DlgText("sStrBox"),vbCrLf,vbLf))
					End Select
				End If
				'显示字符数和字节数
				i = Len(Temp)
				DlgText "sStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
				i = StrHexLength(Temp,CLng(Trim$(DlgText("sCPValueList"))),0)
				DlgText "sCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
			Else
				If DlgText("tStrBox") = "" Then Exit Function
				If DlgValue(DlgItem$) = 0 Then
					Temp = DlgText("tStrBox")
				Else
					TempList = ReSplit(MsgList(70),";")
					i = ShowPopupMenu(TempList,vbPopupUseRightButton)
					If i > -1 Then
						DlgValue "tLineBreakBox",i
						TempList = ReSplit(MsgList(71),";")
						DlgText "tConvertBox",Replace$(MsgList(51),"%s",TempList(i))
					End If
					Select Case DlgValue("tLineBreakBox")
					Case 0
						Temp = Convert(DlgText("tStrBox"))
					Case 1
						Temp = Convert(Replace$(DlgText("tStrBox"),vbCrLf,vbCr))
					Case 2
						Temp = Convert(Replace$(DlgText("tStrBox"),vbCrLf,vbLf))
					End Select
				End If
				'显示字符数和字节数
				i = Len(Temp)
				DlgText "tStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
				i = StrHexLength(Temp,CLng(Trim$(DlgText("tCPValueList"))),0)
				DlgText "tCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
			End If
			If DlgValue("ModeBox") = 1 Then DlgItem$ = "Str2CodeButton"
		Case "ModeBox", "FillList"
			If DlgValue("ModeBox") = 0 Then Exit Function
			If DlgValue("OptionGroup") = 0 Then
				If DlgText("sStrBox") <> "" Then
					DlgItem$ = "Str2CodeButton"
				ElseIf DlgText("sCodeBox") <> ""Then
					DlgItem$ = "Code2StrButton"
				End If
			Else
				If DlgText("tStrBox") <> "" Then
					DlgItem$ = "Str2CodeButton"
				ElseIf DlgText("tCodeBox") <> ""Then
					DlgItem$ = "Code2StrButton"
				End If
			End If
		End Select
		Select Case DlgItem$
		Case "Str2CodeButton"
			If DlgValue("OptionGroup") = 0 Then
				If DlgText("sStrBox") = "" Then
					DlgText "sStrText",Replace$(Replace$(MsgList(42),"%s","0"),"%d","00")
					Exit Function
				End If
				If Temp = "" Then
					If DlgValue("sConvertBox") = 0 Then
						Temp = DlgText("sStrBox")
					Else
						Select Case DlgValue("sLineBreakBox")
						Case 0
							Temp = Convert(DlgText("sStrBox"))
						Case 1
							Temp = Convert(Replace$(DlgText("sStrBox"),vbCrLf,vbCr))
						Case 2
							Temp = Convert(Replace$(DlgText("sStrBox"),vbCrLf,vbLf))
						End Select
					End If
					'显示字符数和字节数
					i = Len(Temp)
					DlgText "sStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					i = StrHexLength(Temp,CLng(Trim$(DlgText("sCPValueList"))),0)
					DlgText "sCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
				End If
				Select Case DlgValue("sModeList")
				Case 0
					Temp = Str2Hex(Temp,CLng(Trim$(DlgText("sCPValueList"))),0,0)
					If DlgValue("sSeparatHexBox") = 1 Then Temp = SeparatHex(Temp)
				Case 1
					Temp = Str2HexEsc(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sHEXASCIICharBox"))
				Case 2
					Temp = Str2URLEsc(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sURLLegalCharBox"))
				Case 3
					Temp = Str2HTMLEsc(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sHTMLASCIICharBox"))
				Case 4
					Temp = Str2ISOEsc(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sISOASCIICharBox"))
				Case 5
					Temp = Base64Encode(Temp,CLng(Trim$(DlgText("sCPValueList"))))
				End Select
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("sCodeBox")),Len(Temp),True
				DlgText "sCodeBox",Temp
			Else
				If DlgText("tStrBox") = "" Then
					DlgText "tStrText",Replace$(Replace$(MsgList(42),"%s","0"),"%d","00")
					Exit Function
				End If
				If Temp = "" Then
					If DlgValue("tConvertBox") = 0 Then
						Temp = DlgText("tStrBox")
					Else
						Select Case DlgValue("tLineBreakBox")
						Case 0
							Temp = Convert(DlgText("tStrBox"))
						Case 1
							Temp = Convert(Replace$(DlgText("tStrBox"),vbCrLf,vbCr))
						Case 2
							Temp = Convert(Replace$(DlgText("tStrBox"),vbCrLf,vbLf))
						End Select
					End If
					'显示字符数和字节数
					i = Len(Temp)
					DlgText "tStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					i = StrHexLength(Temp,CLng(Trim$(DlgText("tCPValueList"))),0)
					DlgText "tCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
				End If
				Select Case DlgValue("tModeList")
				Case 0
					i = StrHexLength(DlgText("sStrBox"),CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sConvertBox"))
					Temp = Str2Hex(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("FillList"),i)
					'显示原文的字符数和字节数
					If DlgValue("FillList") <> 0 Then
						DlgText "tCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(Len(Temp) \ 2)),"%d",FormatHexStr(Hex$(i),2))
					End If
					If DlgValue("tSeparatHexBox") = 1 Then Temp = SeparatHex(Temp)
				Case 1
					Temp = Str2HexEsc(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tHEXASCIICharBox"))
				Case 2
					Temp = Str2URLEsc(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tURLLegalCharBox"))
				Case 3
					Temp = Str2HTMLEsc(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tHTMLASCIICharBox"))
				Case 4
					Temp = Str2ISOEsc(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tISOASCIICharBox"))
				Case 5
					Temp = Base64Encode(Temp,CLng(Trim$(DlgText("tCPValueList"))))
				End Select
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("tCodeBox")),Len(Temp),True
				DlgText "tCodeBox",Temp
			End If
			DlgValue "CheckBox",1
		Case "Code2StrButton"
			If DlgValue("OptionGroup") = 0 Then
				If DlgText("sCodeBox") = "" Then
					DlgText "sCodeText",Replace$(Replace$(MsgList(50),"%s","0"),"%d","00")
					Exit Function
				End If
				If DlgValue("sModeList") = 0 Then
					Temp = Replace$(Replace$(DlgText("sCodeBox")," ",""),vbCrLf,"")
				Else
					Temp = DlgText("sCodeBox")
				End If
				i = 0
				Select Case DlgValue("sModeList")
				Case 1
					i = DlgValue("sHEXASCIICharBox")
				Case 2
					i = DlgValue("sURLLegalCharBox")
				Case 3
					i = DlgValue("sHTMLASCIICharBox")
				Case 4
					i = DlgValue("sISOASCIICharBox")
				End Select
				Select Case CheckHex(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sModeList"),i)
				Case 0
					Select Case DlgValue("sModeList")
					Case 0
						Temp = ByteToString(HexStr2Bytes(Temp),CLng(Trim$(DlgText("sCPValueList"))))
					Case 1
						Temp = HexEsc2Str(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					Case 2
						Temp = URLEsc2Str(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					Case 3
						Temp = HTMLEsc2Str(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					Case 4
						Temp = ISOEsc2Str(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					Case 5
						Temp = Base64Decode(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					End Select
					'显示字符数和字节数
					i = Len(Temp)
					DlgText "sStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					i = StrHexLength(Temp,CLng(Trim$(DlgText("sCPValueList"))),0)
					DlgText "sCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					If DlgValue("sConvertBox") = 1 Then
						Select Case DlgValue("sLineBreakBox")
						Case 0
							Temp = ReConvert(Temp)
						Case 1
							Temp = ReConvert(Replace$(Temp,vbCrLf,vbCr))
						Case 2
							Temp = ReConvert(Replace$(Temp,vbCrLf,vbLf))
						End Select
					End If
					'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
					SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("sStrBox")),Len(Temp),True
					DlgText "sStrBox",Temp
					DlgValue "CheckBox",1
				Case 1
					MsgBox(MsgList(63),vbOkOnly+vbInformation,MsgList(0))
				Case 2
					MsgBox(MsgList(64),vbOkOnly+vbInformation,MsgList(0))
				Case 3
					MsgBox(MsgList(65),vbOkOnly+vbInformation,MsgList(0))
				Case 4
					MsgBox(MsgList(66),vbOkOnly+vbInformation,MsgList(0))
				End Select
			Else
				If DlgText("tCodeBox") = "" Then
					DlgText "tCodeText",Replace$(Replace$(MsgList(50),"%s","0"),"%d","00")
					Exit Function
				End If
				If DlgValue("tModeList") = 0 Then
					Temp = Replace$(Replace$(DlgText("tCodeBox")," ",""),vbCrLf,"")
				Else
					Temp = DlgText("tCodeBox")
				End If
				i = 0
				Select Case DlgValue("tModeList")
				Case 1
					i = DlgValue("tHEXASCIICharBox")
				Case 2
					i = DlgValue("tURLLegalCharBox")
				Case 3
					i = DlgValue("tHTMLASCIICharBox")
				Case 4
					i = DlgValue("tISOASCIICharBox")
				End Select
				Select Case CheckHex(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tModeList"),i)
				Case 0
					Select Case DlgValue("tModeList")
					Case 0
						Temp = ByteToString(HexStr2Bytes(Temp),CLng(Trim$(DlgText("tCPValueList"))))
					Case 1
						Temp = HexEsc2Str(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					Case 2
						Temp = URLEsc2Str(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					Case 3
						Temp = HTMLEsc2Str(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					Case 4
						Temp = ISOEsc2Str(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					Case 5
						Temp = Base64Decode(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					End Select
					'显示字符数和字节数
					i = Len(Temp)
					DlgText "tStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					i = StrHexLength(Temp,CLng(Trim$(DlgText("tCPValueList"))),0)
					DlgText "tCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					If DlgValue("tConvertBox") = 1 Then
						Select Case DlgValue("tLineBreakBox")
						Case 0
							Temp = ReConvert(Temp)
						Case 1
							Temp = ReConvert(Replace$(Temp,vbCrLf,vbCr))
						Case 2
							Temp = ReConvert(Replace$(Temp,vbCrLf,vbLf))
						End Select
					End If
					'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
					SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("tStrBox")),Len(Temp),True
					DlgText "tStrBox",Temp
					DlgValue "CheckBox",1
				Case 1
					MsgBox(MsgList(63),vbOkOnly+vbInformation,MsgList(0))
				Case 2
					MsgBox(MsgList(64),vbOkOnly+vbInformation,MsgList(0))
				Case 3
					MsgBox(MsgList(65),vbOkOnly+vbInformation,MsgList(0))
				Case 4
					MsgBox(MsgList(66),vbOkOnly+vbInformation,MsgList(0))
				End Select
			End If
		Case "LangButton"
			ReDim TempList(0) As String
			TempList = ReSplit(MsgList(73),";")
			i = ShowPopupMenu(TempList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			TempList = ReSplit(MsgList(74),";")
			If GetMsgList(MsgList,TempList(i)) = False Then Exit Function

			'更改文本框语言
			DlgText -1,Replace$(Replace$(MsgList(39),"%v",Version),"%b",Build)
			DlgText "SourceButton",MsgList(40)
			DlgText "TransButton",MsgList(41)

			'DlgText "sStrText",MsgList(42)
			'DlgText "tStrText",MsgList(42)
			'DlgListBoxArray "sModeList",MsgList()
			'DlgListBoxArray "tModeList",MsgList()
			DlgText "sSeparatHexBox",MsgList(43)
			DlgText "tSeparatHexBox",MsgList(43)
			DlgText "sHEXASCIICharBox",MsgList(44)
			DlgText "tHEXASCIICharBox",MsgList(44)
			DlgText "sURLLegalCharBox",MsgList(45)
			DlgText "tURLLegalCharBox",MsgList(45)
			DlgText "sHTMLASCIICharBox",MsgList(44)
			DlgText "tHTMLASCIICharBox",MsgList(44)
			DlgText "sISOASCIICharBox",MsgList(44)
			DlgText "tISOASCIICharBox",MsgList(44)
			'DlgListBoxArray "FillList",MsgList()
			DlgText "Str2CodeButton",MsgList(46)
			DlgText "StrCopyButton",MsgList(47)
			DlgText "StrPasteButton",MsgList(48)
			DlgText "StrCleanButton",MsgList(49)

			'DlgText "sCodeText",MsgList(50)
			'DlgText "tCodeText",MsgList(50)
			'DlgListBoxArray "sCPNameList",MsgList()
			'DlgListBoxArray "sCPValueList",MsgList()
			'DlgListBoxArray "tCPNameList",MsgList()
			'DlgListBoxArray "tCPValueList",MsgList()
			'DlgText "sConvertBox",MsgList(51)
			'DlgText "tConvertBox",MsgList(51)
			DlgText "ModeBox",MsgList(52)
			DlgText "Code2StrButton",MsgList(53)
			DlgText "CodeCopyButton",MsgList(54)
			DlgText "CodePasteButton",MsgList(55)
			DlgText "CodeCleanButton",MsgList(56)
			DlgText "LangButton",MsgList(72)
			DlgText "AboutButton",MsgList(57)

			'更改列表框语言
			ReDim TempList(23) As String
			TempList(0) = CStr$(CP_OEMCP)			'系统默认
			TempList(1) = CStr$(CP_MACCP)
			TempList(2) = CStr$(CP_THREAD_ACP)
			TempList(3) = CStr$(CP_WESTEUROPE)		'拉丁文 1 (ANSI) = 1252
			TempList(4) = CStr$(CP_EASTEUROPE)		'拉丁文 2 (中欧) = 1250
			TempList(5) = CStr$(CP_RUSSIAN)			'西里尔文 (斯拉夫) = 1251
			TempList(6) = CStr$(CP_GREEK)			'希腊文 = 1253
			TempList(7) = CStr$(CP_TURKISH)			'拉丁文 5 (土耳其) = 1254
			TempList(8) = CStr$(CP_HEBREW)			'希伯来文 = 1255
			TempList(9) = CStr$(CP_ARABIC)			'阿拉伯文 = 1256
			TempList(10) = CStr$(CP_BALTIC)			'波罗的海文 = 1257
			TempList(11) = CStr$(CP_VIETNAMESE)		'越南文 = 1258
			TempList(12) = CStr$(CP_JAPAN)			'日文 = 932
			TempList(13) = CStr$(CP_CHINA)			'简体中文 = 936
			TempList(14) = CStr$(CP_GB18030)		'简体中文 = 54936
			TempList(15) = CStr$(CP_KOREA)			'韩文 = 949
			TempList(16) = CStr$(CP_TAIWAN) 		'繁体中文 = 950
			TempList(17) = CStr$(CP_THAI)			'泰文 = 874
			TempList(18) = CStr$(CP_UTF7)			'UTF-7 = 65000
			TempList(19) = CStr$(CP_UTF8)			'UTF-8 = 65001
			TempList(20) = CStr$(CP_UNICODELITTLE)	'UnicodeLE = 1200
			TempList(21) = CStr$(CP_UNICODEBIG)		'UnicodeBE = 1201
			TempList(22) = CStr$(CP_UTF32LE)		'UnicodeLE = 65005
			TempList(23) = CStr$(CP_UTF32BE)		'UnicodeBE = 65006
			For i = 0 To 23
				TempList(i) = TempList(i) & " - " & MsgList(i + 14)
			Next i
			i = DlgValue("sCPValueList")
			DlgListBoxArray "sCPNameList",TempList()
			DlgValue "sCPNameList",i
			i = DlgValue("tCPValueList")
			DlgListBoxArray "tCPNameList",TempList()
			DlgValue "tCPNameList",i

			TempList = ReSplit(MsgList(38),ItemJoinStr)
			i = DlgValue("sModeList")
			DlgListBoxArray "sModeList",TempList()
			DlgValue "sModeList",i
			i = DlgValue("tModeList")
			DlgListBoxArray "tModeList",TempList()
			DlgValue "tModeList",i

			ReDim TempList(4) As String
			For i = 0 To 4
				TempList(i) = MsgList(i + 58)
			Next i
			i = DlgValue("FillList")
			DlgListBoxArray "FillList",TempList()
			DlgValue "FillList",i

			'更改转义选项语言
			TempList = ReSplit(MsgList(71),";")
			DlgText "sConvertBox",Replace$(MsgList(51),"%s",TempList(DlgValue("sLineBreakBox")))
			DlgText "tConvertBox",Replace$(MsgList(51),"%s",TempList(DlgValue("tLineBreakBox")))

			'更改字符字节数量显示语言
			Temp = Mid$(DlgText("sStrText"),InStr(DlgText("sStrText"),"("))
			DlgText "sStrText",Replace$(MsgList(42),"(%s/%d):",Temp)
			Temp = Mid$(DlgText("sCodeText"),InStr(DlgText("sCodeText"),"("))
			DlgText "sCodeText",Replace$(MsgList(50),"(%s/%d):",Temp)
			Temp = Mid$(DlgText("tStrText"),InStr(DlgText("tStrText"),"("))
			DlgText "tStrText",Replace$(MsgList(42),"(%s/%d):",Temp)
			Temp = Mid$(DlgText("tCodeText"),InStr(DlgText("tCodeText"),"("))
			DlgText "tCodeText",Replace$(MsgList(50),"(%s/%d):",Temp)
		End Select
	Case 3 ' 文本框或者组合框文本被更改
		If DlgValue("CheckBox") = 1 Then Exit Function
		Select Case DlgItem$
		Case "sStrBox", "tStrBox"
			If DlgText(DlgItem$) = "" Then
				DlgText Replace$(DlgItem$,"Box","Text"),Replace$(Replace$(MsgList(42),"%s","0"),"%d","00")
				Exit Function
			End If
			If DlgValue("OptionGroup") = 0 Then
				If DlgValue("sConvertBox") = 0 Then
					Temp = DlgText(DlgItem$)
				Else
					Select Case DlgValue(IIf(DlgItem$ = "sStrBox","sLineBreakBox","tLineBreakBox"))
					Case 0
						Temp = Convert(DlgText(DlgItem$))
					Case 1
						Temp = Convert(Replace$(DlgText(DlgItem$),vbCrLf,vbCr))
					Case 2
						Temp = Convert(Replace$(DlgText(DlgItem$),vbCrLf,vbLf))
					End Select
				End If
				'显示字符数和字节数
				i = Len(Temp)
				DlgText "sStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
				i = StrHexLength(Temp,CLng(Trim$(DlgText("sCPValueList"))),0)
				DlgText "sCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
				If DlgValue("ModeBox") = 0 Then Exit Function
				Select Case DlgValue("sModeList")
				Case 0
					Temp = Str2Hex(Temp,CLng(Trim$(DlgText("sCPValueList"))),0,0)
					If DlgValue("sSeparatHexBox") = 1 Then Temp = SeparatHex(Temp)
				Case 1
					Temp = Str2HexEsc(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sHEXASCIICharBox"))
				Case 2
					Temp = Str2URLEsc(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sURLLegalCharBox"))
				Case 3
					Temp = Str2HTMLEsc(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sHTMLASCIICharBox"))
				Case 4
					Temp = Str2ISOEsc(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sISOASCIICharBox"))
				Case 5
					Temp = Base64Encode(Temp,CLng(Trim$(DlgText("sCPValueList"))))
				End Select
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("sCodeBox")),Len(Temp),True
				DlgText "sCodeBox",Temp
			Else
				If DlgValue("tConvertBox") = 0 Then
					Temp = DlgText(DlgItem$)
				Else
					Select Case DlgValue(IIf(DlgItem$ = "sStrBox","sLineBreakBox","tLineBreakBox"))
					Case 0
						Temp = Convert(DlgText(DlgItem$))
					Case 1
						Temp = Convert(Replace$(DlgText(DlgItem$),vbCrLf,vbCr))
					Case 2
						Temp = Convert(Replace$(DlgText(DlgItem$),vbCrLf,vbLf))
					End Select
				End If
				'显示字符数和字节数
				i = Len(Temp)
				DlgText "tStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
				i = StrHexLength(Temp,CLng(Trim$(DlgText("tCPValueList"))),0)
				DlgText "tCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
				If DlgValue("ModeBox") = 0 Then Exit Function
				Select Case DlgValue("tModeList")
				Case 0
					i = StrHexLength(DlgText("sStrBox"),CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sConvertBox"))
					Temp = Str2Hex(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("FillList"),i)
					'显示原文的字符数和字节数
					If DlgValue("FillList") <> 0 Then
						DlgText "tCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(Len(Temp) \ 2)),"%d",FormatHexStr(Hex$(i),2))
					End If
					If DlgValue("tSeparatHexBox") = 1 Then Temp = SeparatHex(Temp)
				Case 1
					Temp = Str2HexEsc(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tHEXASCIICharBox"))
				Case 2
					Temp = Str2URLEsc(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tURLLegalCharBox"))
				Case 3
					Temp = Str2HTMLEsc(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tHTMLASCIICharBox"))
				Case 4
					Temp = Str2ISOEsc(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tISOASCIICharBox"))
				Case 5
					Temp = Base64Encode(Temp,CLng(Trim$(DlgText("tCPValueList"))))
				End Select
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("tCodeBox")),Len(Temp),True
				DlgText "tCodeBox",Temp
			End If
			DlgValue "CheckBox",1
		Case "sCodeBox", "tCodeBox"
			If DlgText(DlgItem$) = "" Then
				DlgText Replace$(DlgItem$,"Box","Text"),Replace$(Replace$(MsgList(50),"%s","0"),"%d","00")
				Exit Function
			End If
			If DlgValue("ModeBox") = 0 Then Exit Function
			If DlgValue("OptionGroup") = 0 Then
				If DlgValue("sModeList") = 0 Then
					Temp = Replace$(Replace$(DlgText(DlgItem$)," ",""),vbCrLf,"")
				Else
					Temp = DlgText(DlgItem$)
				End If
				i = 0
				Select Case DlgValue("sModeList")
				Case 1
					i = DlgValue("sHEXASCIICharBox")
				Case 2
					i = DlgValue("sURLLegalCharBox")
				Case 3
					i = DlgValue("sHTMLASCIICharBox")
				Case 4
					i = DlgValue("sISOASCIICharBox")
				End Select
				Select Case CheckHex(Temp,CLng(Trim$(DlgText("sCPValueList"))),DlgValue("sModeList"),i)
				Case 0
					Select Case DlgValue("sModeList")
					Case 0
						Temp = ByteToString(HexStr2Bytes(Temp),CLng(Trim$(DlgText("sCPValueList"))))
					Case 1
						Temp = HexEsc2Str(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					Case 2
						Temp = URLEsc2Str(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					Case 3
						Temp = HTMLEsc2Str(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					Case 4
						Temp = ISOEsc2Str(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					Case 5
						Temp = Base64Decode(Temp,CLng(Trim$(DlgText("sCPValueList"))))
					End Select
					'显示字符数和字节数
					i = Len(Temp)
					DlgText "sStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					i = StrHexLength(Temp,CLng(Trim$(DlgText("sCPValueList"))),0)
					DlgText "sCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					If DlgValue("sConvertBox") = 1 Then
						Select Case DlgValue("sLineBreakBox")
						Case 0
							Temp = ReConvert(Temp)
						Case 1
							Temp = ReConvert(Replace$(Temp,vbCrLf,vbCr))
						Case 2
							Temp = ReConvert(Replace$(Temp,vbCrLf,vbLf))
						End Select
					End If
					'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
					SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("sStrBox")),Len(Temp),True
					DlgText "sStrBox",Temp
					DlgValue "CheckBox",1
				Case 1
					MsgBox(MsgList(63),vbOkOnly+vbInformation,MsgList(0))
				Case 2
					MsgBox(MsgList(64),vbOkOnly+vbInformation,MsgList(0))
				Case 3
					MsgBox(MsgList(65),vbOkOnly+vbInformation,MsgList(0))
				Case 4
					MsgBox(MsgList(66),vbOkOnly+vbInformation,MsgList(0))
				End Select
			Else
				If DlgValue("tModeList") = 0 Then
					Temp = Replace$(Replace$(DlgText(DlgItem$)," ",""),vbCrLf,"")
				Else
					Temp = DlgText(DlgItem$)
				End If
				i = 0
				Select Case DlgValue("tModeList")
				Case 1
					i = DlgValue("tHEXASCIICharBox")
				Case 2
					i = DlgValue("tURLLegalCharBox")
				Case 3
					i = DlgValue("tHTMLASCIICharBox")
				Case 4
					i = DlgValue("tISOASCIICharBox")
				End Select
				Select Case CheckHex(Temp,CLng(Trim$(DlgText("tCPValueList"))),DlgValue("tModeList"),i)
				Case 0
					Select Case DlgValue("tModeList")
					Case 0
						Temp = ByteToString(HexStr2Bytes(Temp),CLng(Trim$(DlgText("tCPValueList"))))
					Case 1
						Temp = HexEsc2Str(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					Case 2
						Temp = URLEsc2Str(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					Case 3
						Temp = HTMLEsc2Str(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					Case 4
						Temp = ISOEsc2Str(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					Case 5
						Temp = Base64Decode(Temp,CLng(Trim$(DlgText("tCPValueList"))))
					End Select
					'显示字符数和字节数
					i = Len(Temp)
					DlgText "tStrText",Replace$(Replace$(MsgList(42),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					i = StrHexLength(Temp,CLng(Trim$(DlgText("tCPValueList"))),0)
					DlgText "tCodeText",Replace$(Replace$(MsgList(50),"%s",CStr$(i)),"%d",FormatHexStr(Hex$(i),2))
					If DlgValue("tConvertBox") = 1 Then
						Select Case DlgValue("tLineBreakBox")
						Case 0
							Temp = ReConvert(Temp)
						Case 1
							Temp = ReConvert(Replace$(Temp,vbCrLf,vbCr))
						Case 2
							Temp = ReConvert(Replace$(Temp,vbCrLf,vbLf))
						End Select
					End If
					'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
					SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("tStrBox")),Len(Temp),True
					DlgText "tStrBox",Temp
					DlgValue "CheckBox",1
				Case 1
					MsgBox(MsgList(63),vbOkOnly+vbInformation,MsgList(0))
				Case 2
					MsgBox(MsgList(64),vbOkOnly+vbInformation,MsgList(0))
				Case 3
					MsgBox(MsgList(65),vbOkOnly+vbInformation,MsgList(0))
				Case 4
					MsgBox(MsgList(66),vbOkOnly+vbInformation,MsgList(0))
				End Select
			End If
		End Select
	Case 4 ' 焦点被更改
		Select Case DlgItem$
		Case "sStrBox"
			If Len(Clipboard) < Len(DlgText("sStrBox")) Then Exit Function
			'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
			i = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("sStrBox"))
			SetTextBoxLength i,Len(DlgText("sStrBox")) + Len(Clipboard),False
		Case "tStrBox"
			If Len(Clipboard) < Len(DlgText("tStrBox")) Then Exit Function
			'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
			i = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("tStrBox"))
			SetTextBoxLength i,Len(DlgText("tStrBox")) + Len(Clipboard),False
		Case "sCodeBox"
			If Len(Clipboard) < Len(DlgText("sCodeBox")) Then Exit Function
			'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
			i = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("sCodeBox"))
			SetTextBoxLength i,Len(DlgText("sCodeBox")) + Len(Clipboard),False
		Case "tCodeBox"
			If Len(Clipboard) < Len(DlgText("tCodeBox")) Then Exit Function
			'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
			i = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("tCodeBox"))
			SetTextBoxLength i,Len(DlgText("tCodeBox")) + Len(Clipboard),False
		End Select
	End Select
End Function


'常用的代码页：UTF-8:65001；GB2312:936；GB18030：54936；UTF-7：65000
Private Function MultiByteToUTF16(UTF8() As Byte, ByVal CodePage As Long) As String
	On Error GoTo errHandle
	Dim bufSize As Long,lRet As Long
	'bufSize = MultiByteToWideChar(CodePage, 0&, VarPtr(UTF8(0)), UBound(UTF8) + 1, 0, 0)
	bufSize = MultiByteToWideChar(CodePage, 0&, UTF8(0), UBound(UTF8) + 1, 0, 0)
	'If CodePage = CP_UTF8 Then bufSize = (UBound(UTF8) + 1) * 2 Else bufSize = UBound(UTF8) + 1
	MultiByteToUTF16 = Space(bufSize)
	'lRet = MultiByteToWideChar(CodePage, 0&, VarPtr(UTF8(0)), UBound(UTF8) + 1, StrPtr(MultiByteToUTF16), bufSize)
	lRet = MultiByteToWideChar(CodePage, 0&, UTF8(0), UBound(UTF8) + 1, StrPtr(MultiByteToUTF16), bufSize)
	If lRet > 0 Then MultiByteToUTF16 = Left$(MultiByteToUTF16, lRet)
	Exit Function
	errHandle:
	MultiByteToUTF16 = ""
End Function


Private Function UTF16ToMultiByte(ByVal UTF16 As String, ByVal CodePage As Long) As Byte()
	Dim bufSize As Long,lRet As Long
	On Error GoTo errHandle
	ReDim arr(0) As Byte
	bufSize = WideCharToMultiByte(CodePage, 0&, StrPtr(UTF16), Len(UTF16), arr(0), 0, 0, 0)
	'If CodePage = CP_UTF8 Then bufSize = 2 * LenB(UTF16) - 1 Else bufSize = LenB(UTF16)
	If bufSize < 1 Then bufSize = 1
	ReDim arr(bufSize - 1) As Byte
	lRet = WideCharToMultiByte(CodePage, 0&, StrPtr(UTF16), Len(UTF16), arr(0), bufSize, 0, 0)
	If lRet > 0 Then ReDim Preserve arr(lRet - 1) As Byte
	UTF16ToMultiByte = arr
	Exit Function
	errHandle:
	ReDim arr(0) As Byte
	UTF16ToMultiByte = arr
End Function


'字节转 Hex 码
'StartPos <= EndPos 获取低位到高位的 Hex 代码，否则获取高位到低位的 Hex 代码
Private Function Byte2Hex(Bytes() As Byte,ByVal StartPos As Long,ByVal EndPos As Long) As String
	Dim i As Long,n As Long
	If StartPos < 0 Then StartPos = LBound(Bytes)
	If EndPos < 0 Then EndPos = UBound(Bytes)
	Byte2Hex = Space$((Abs(EndPos - StartPos) + 1) * 2)
	n = 1
	For i = StartPos To EndPos Step IIf(StartPos <= EndPos,1,-1)
		Mid$(Byte2Hex,n,2) = Right$("0" & Hex$(Bytes(i)),2)
		n = n + 2
	Next i
End Function


'字节转 Hex 转义码
'StartPos <= EndPos 获取低位到高位的 Hex 代码，否则获取高位到低位的 Hex 代码
Private Function Byte2HexEsc(Bytes() As Byte,ByVal StartPos As Long,ByVal EndPos As Long,ByVal CodePage As Long) As String
	Dim i As Long,n As Long
	If StartPos < 0 Then StartPos = LBound(Bytes)
	If EndPos < 0 Then EndPos = UBound(Bytes)
	Select Case CodePage
	Case CP_UNICODELITTLE
		Byte2HexEsc = Space$((Abs(EndPos - StartPos) + 1) * 3)
		n = 1
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,2,-2)
			Mid$(Byte2HexEsc,n,6) = "\u" & Right$("0" & Hex$(Bytes(i + 1)),2) & Right$("0" & Hex$(Bytes(i)),2)
			n = n + 6
		Next i
	Case CP_UNICODEBIG
		Byte2HexEsc = Space$((Abs(EndPos - StartPos) + 1) * 3)
		n = 1
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,2,-2)
			Mid$(Byte2HexEsc,n,6) = "\u" & Right$("0" & Hex$(Bytes(i)),2) & Right$("0" & Hex$(Bytes(i + 1)),2)
			n = n + 6
		Next i
	Case CP_UTF32LE, CP_UTF_32LE
		Byte2HexEsc = Space$((Abs(EndPos - StartPos) + 1) * 2.5)
		n = 1
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,4,-4)
			Mid$(Byte2HexEsc,n,10) = "\u" & Right$("0" & Hex$(Bytes(i + 3)),2) & Right$("0" & Hex$(Bytes(i + 2)),2) & _
									Right$("0" & Hex$(Bytes(i + 1)),2) & Right$("0" & Hex$(Bytes(i)),2)
			n = n + 10
		Next i
	Case CP_UTF32BE, CP_UTF_32BE
		Byte2HexEsc = Space$((Abs(EndPos - StartPos) + 1) * 2.5)
		n = 1
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,4,-4)
			Mid$(Byte2HexEsc,n,10) = "\u" & Right$("0" & Hex$(Bytes(i)),2) & Right$("0" & Hex$(Bytes(i + 1)),2) & _
									Right$("0" & Hex$(Bytes(i + 2)),2) & Right$("0" & Hex$(Bytes(i + 3)),2)
			n = n + 10
		Next i
	Case Else
		Byte2HexEsc = Space$((Abs(EndPos - StartPos) + 1) * 4)
		n = 1
		For i = StartPos To EndPos Step IIf(StartPos <= EndPos,1,-1)
			Mid$(Byte2HexEsc,n,4) = "\x" & Right$("0" & Hex$(Bytes(i)),2)
			n = n + 4
		Next i
	End Select
End Function


'字节转 RUL 转义符
'StartPos <= EndPos 获取低位到高位的 Hex 代码，否则获取高位到低位的 Hex 代码
Private Function Byte2URLEsc(Bytes() As Byte,ByVal StartPos As Long,ByVal EndPos As Long,ByVal CodePage As Long) As String
	Dim i As Long,n As Long
	If StartPos < 0 Then StartPos = LBound(Bytes)
	If EndPos < 0 Then EndPos = UBound(Bytes)
	Select Case CodePage
	Case CP_UNICODELITTLE
		Byte2URLEsc = Space$((Abs(EndPos - StartPos) + 1) * 3)
		n = 1
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,2,-2)
			Mid$(Byte2URLEsc,n,6) = "%u" & Right$("0" & Hex$(Bytes(i + 1)),2) & Right$("0" & Hex$(Bytes(i)),2)
			n = n + 6
		Next i
	Case CP_UNICODEBIG
		Byte2URLEsc = Space$((Abs(EndPos - StartPos) + 1) * 3)
		n = 1
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,2,-2)
			Mid$(Byte2URLEsc,n,6) = "%u" & Right$("0" & Hex$(Bytes(i)),2) & Right$("0" & Hex$(Bytes(i + 1)),2)
			n = n + 6
		Next i
	Case CP_UTF32LE, CP_UTF_32LE
		Byte2URLEsc = Space$((Abs(EndPos - StartPos) + 1) * 2.5)
		n = 1
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,4,-4)
			Mid$(Byte2URLEsc,n,10) = "%u" & Right$("0" & Hex$(Bytes(i + 3)),2) & Right$("0" & Hex$(Bytes(i + 2)),2) & _
									Right$("0" & Hex$(Bytes(i + 1)),2) & Right$("0" & Hex$(Bytes(i)),2)
			n = n + 10
		Next i
	Case CP_UTF32BE, CP_UTF_32BE
		Byte2URLEsc = Space$((Abs(EndPos - StartPos) + 1) * 2.5)
		n = 1
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,4,-4)
			Mid$(Byte2URLEsc,n,10) = "%u" & Right$("0" & Hex$(Bytes(i)),2) & Right$("0" & Hex$(Bytes(i + 1)),2) & _
									Right$("0" & Hex$(Bytes(i + 2)),2) & Right$("0" & Hex$(Bytes(i + 3)),2)
			n = n + 10
		Next i
	Case Else
		Byte2URLEsc = Space$((Abs(EndPos - StartPos) + 1) * 3)
		n = 1
		For i = StartPos To EndPos Step IIf(StartPos <= EndPos,1,-1)
			Mid$(Byte2URLEsc,n,3) = "%" & Right$("0" & Hex$(Bytes(i)),2)
			n = n + 3
		Next i
	End Select
End Function


'字节转 HTML 转义符(转入的字节数组必须都是 UnicodeLE)
'StartPos <= EndPos 获取低位到高位的 Hex 代码，否则获取高位到低位的 Hex 代码
Private Function Byte2HTMLEsc(Bytes() As Byte,ByVal StartPos As Long,ByVal EndPos As Long,ByVal CodePage As Long) As String
	Dim i As Long,n As Long,TempList() As String
	If StartPos < 0 Then StartPos = LBound(Bytes)
	If EndPos < 0 Then EndPos = UBound(Bytes)
	Select Case CodePage
 	Case CP_UNICODELITTLE
		ReDim TempList(Abs(EndPos - StartPos)) As String
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,2,-2)
			TempList(n) = "&#" & CStr$(Bytes(i) + 256& * Bytes(i + 1)) & ";"
			n = n + 1
		Next i
		Byte2HTMLEsc = StrListJoin(TempList,"")
	Case CP_UNICODEBIG
		ReDim TempList(Abs(EndPos - StartPos)) As String
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,2,-2)
			TempList(n) = "&#" & CStr$(Bytes(i + 1) + 256& * Bytes(i)) & ";"
			n = n + 1
		Next i
		Byte2HTMLEsc = StrListJoin(TempList,"")
	Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
		ReDim tmpBytes(0) As Byte
		tmpBytes = StringToByte(ByteToString(Bytes,CodePage),CP_UNICODELITTLE)	'转 CP_UNICODELITTLE 字节序
		If StartPos > 0 Then StartPos = StartPos \ 2
		If EndPos > 0 Then EndPos = EndPos \ 2
		ReDim TempList(Abs(EndPos - StartPos)) As String
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,2,-2)
			TempList(n) = "&#" & CStr$(tmpBytes(i) + 256& * tmpBytes(i + 1)) & ";"
			n = n + 1
		Next i
		Byte2HTMLEsc = StrListJoin(TempList,"")
	Case Else
		ReDim TempList(Abs(EndPos - StartPos)) As String
		For i = StartPos To EndPos Step IIf(StartPos <= EndPos,1,-1)
			TempList(n) = "&#" & CStr$(Bytes(i) + 0) & ";"
			n = n + 1
		Next i
		Byte2HTMLEsc = StrListJoin(TempList,"")
	End Select
End Function


'字节转 ISO88591 转义符(转入的字节数组必须都是 UnicodeLE)
'StartPos <= EndPos 获取低位到高位的 Hex 代码，否则获取高位到低位的 Hex 代码
Private Function Byte2ISOEsc(Bytes() As Byte,ByVal StartPos As Long,ByVal EndPos As Long,ByVal CodePage As Long) As String
	Dim i As Long,n As Long,TempList() As String
	If StartPos < 0 Then StartPos = LBound(Bytes)
	If EndPos < 0 Then EndPos = UBound(Bytes)
	Select Case CodePage
 	Case CP_UNICODELITTLE
		ReDim TempList(Abs(EndPos - StartPos)) As String
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,2,-2)
			TempList(n) = "#" & CStr$(Bytes(i) + 256& * Bytes(i + 1))
			n = n + 1
		Next i
		Byte2ISOEsc = StrListJoin(TempList,"")
	Case CP_UNICODEBIG
		ReDim TempList(Abs(EndPos - StartPos)) As String
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,2,-2)
			TempList(n) = "#" & CStr$(Bytes(i + 1) + 256& * Bytes(i))
			n = n + 1
		Next i
		Byte2ISOEsc = StrListJoin(TempList,"")
	Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
		ReDim tmpBytes(0) As Byte
		tmpBytes = StringToByte(ByteToString(Bytes,CodePage),CP_UNICODELITTLE)	'转 CP_UNICODELITTLE 字节序
		If StartPos > 0 Then StartPos = StartPos \ 2
		If EndPos > 0 Then EndPos = EndPos \ 2
		ReDim TempList(Abs(EndPos - StartPos)) As String
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,2,-2)
			TempList(n) = "#" & CStr$(tmpBytes(i) + 256& * tmpBytes(i + 1))
			n = n + 1
		Next i
		Byte2ISOEsc = StrListJoin(TempList,"")
	Case Else
		ReDim TempList(Abs(EndPos - StartPos)) As String
		For i = StartPos To EndPos Step IIf(StartPos <= EndPos,1,-1)
			TempList(n) = "#" & CStr$(Bytes(i) + 0)
			n = n + 1
		Next i
		Byte2ISOEsc = StrListJoin(TempList,"")
	End Select
End Function


'转换 HEX 代码为字节数组
Private Function HexStr2Bytes(ByVal HexStr As String) As Byte()
	Dim i As Long,n As Long,Length As Long
	Length = Len(HexStr)
	If Length > 1 Then
		ReDim tempByte(Length \ 2 - 1) As Byte
		For i = 1 To Length Step 2
			tempByte(n) = Val("&H" & Mid$(HexStr,i,2))
			n = n + 1
		Next i
	Else
		ReDim tempByte(0) As Byte
	End If
	HexStr2Bytes = tempByte
End Function


'字符串转字节数组
Private Function StringToByte(ByVal textStr As String,ByVal CodePage As Long) As Byte()
	Dim i As Long,n As Long,k As Long
	If textStr = "" Then
		ReDim StringToByte(0) As Byte
		Exit Function
	End If
	Select Case CodePage
	Case CP_UNICODELITTLE
		StringToByte = textStr
	Case CP_UNICODEBIG
		StringToByte = textStr
		StringToByte = LowByte2HighByte(StringToByte,2)
	Case CP_UTF32LE, CP_UTF_32LE
		ReDim Bytes(Len(textStr) * 4 - 1) As Byte
		For i = 1 To Len(textStr)
			'转换负值为正值
			k = AscW(Mid$(textStr,i,1)) And 65535	'开辟4个字节的内存空间
			CopyMemory Bytes(n), k, 4
			n = n + 4
		Next i
		StringToByte = Bytes
	Case CP_UTF32BE, CP_UTF_32BE
		ReDim Bytes(Len(textStr) * 4 - 1) As Byte
		For i = 1 To Len(textStr)
			'转换负值为正值
			k = AscW(Mid$(textStr,i,1)) And 65535	'开辟4个字节的内存空间
			CopyMemory Bytes(n), k, 4
			n = n + 4
		Next i
		StringToByte = LowByte2HighByte(Bytes,4)
	Case Else
		StringToByte = UTF16ToMultiByte(textStr,CodePage)	'按指定代码页转Unicode字符为ANSI组数
		'StringToByte = StrConv(textStr,vbFromUnicode)		'按本机代码页转Unicode字符为ANSI组数
	End Select
End Function


'字节数组转字符串
Private Function ByteToString(Bytes() As Byte,ByVal CodePage As Long) As String
	Dim i As Long,n As Long
	On Error Resume Next
	Select Case CodePage
	Case CP_UNICODELITTLE
		ByteToString = Bytes
	Case CP_UNICODEBIG
		ByteToString = LowByte2HighByte(Bytes,2)
	Case CP_UTF32LE, CP_UTF_32LE
		ByteToString = Space$((UBound(Bytes) + 1) \ 4)
		For i = 0 To UBound(Bytes) Step 4
			n = n + 1
			'检查是否符合 CP_UTF32LE, CP_UTF_32LE 编码标准
			If Bytes(i + 3) = 0 Then
				CopyMemory CodePage, Bytes(i), 4
				'转换大于&H7FFF的值为负值
				'ChrW函数的数值范围为 Integer 的取值范围：-32768 到 32767，大于 32767 时需要减去 65536
				If CodePage > &H7FFF Then CodePage = CodePage - 65536
				Mid$(ByteToString,n,1) = ChrW(CodePage)
			Else
				Mid$(ByteToString,n,1) = vbNullChar
			End If
		Next i
	Case CP_UTF32BE, CP_UTF_32BE
		ReDim tmpBytes(0) As Byte
		tmpBytes = LowByte2HighByte(Bytes,4)
		ByteToString = Space$((UBound(Bytes) + 1) \ 4)
		For i = 0 To UBound(Bytes) Step 4
			n = n + 1
			'检查是否符合 CP_UTF32BE, CP_UTF_32BE 编码标准
			If Bytes(i) = 0 Then
				CopyMemory CodePage, tmpBytes(i), 4
				'转换大于&H7FFF的值为负值
				'ChrW函数的数值范围为 Integer 的取值范围：-32768 到 32767，大于 32767 时需要减去 65536
				If CodePage > &H7FFF Then CodePage = CodePage - 65536
				Mid$(ByteToString,n,1) = ChrW(CodePage)
			Else
				Mid$(ByteToString,n,1) = vbNullChar
			End If
		Next i
	Case Else
		ByteToString = MultiByteToUTF16(Bytes,CodePage)
	End Select
End Function


'字符型字节数组的高字节和低字节互换
'适用于 UNICODE LITTLE 和 UNICODE BIG 字节数组的相互转换
Private Function LowByte2HighByte(Bytes() As Byte,ByVal Setp As Integer) As Byte()
	Dim i As Long,Temp() As Byte
	Temp = Bytes
	If Setp = 2 Then
		For i = LBound(Temp) To UBound(Temp) - 1 Step Setp
			Temp(i) = Bytes(i + 1)
			Temp(i + 1) = Bytes(i)
		Next i
	ElseIf Setp = 4 Then
		For i = LBound(Temp) To UBound(Temp) - 1 Step Setp
			Temp(i) = Bytes(i + 3)
			Temp(i + 1) = Bytes(i + 2)
			Temp(i + 2) = Bytes(i + 1)
			Temp(i + 3) = Bytes(i)
		Next i
	End If
	LowByte2HighByte = Temp
End Function


'反转字节数组，适用于数值型字节数组的高字节和低字节互换
Private Function ReverseValByte(Bytes() As Byte,ByVal StartPos As Long,ByVal EndPos As Long) As Byte()
	Dim i As Long,Temp() As Byte
	If StartPos < 0 Then StartPos = LBound(Bytes)
	If EndPos < 0 Then EndPos = UBound(Bytes)
	Temp = Bytes
	For i = StartPos To EndPos
		Temp(i) = Bytes(EndPos - i)
	Next i
	ReverseValByte = Temp
End Function


'检查输入的 Hex 字串是否符合要求
'返回 CheckHex = 0 合格, = 1 HEX 代码字符数不符, = 2 数字数不符, = 3 包含非法字符 = 4 没有要转换的代码(仅 Ignore = 1 的情况)
Private Function CheckHex(ByVal textStr As String,ByVal CodePage As Long,ByVal Mode As Integer,ByVal Ignore As Integer) As Long
	Dim i As Long,Temp As String,Matches As Object
	If textStr = "" Then Exit Function
	Select Case Mode
	Case 0	'Hex Mode
		If (Len(textStr) Mod 2) <> 0 Then
			CheckHex = 1
		ElseIf CheckStrRegExp(textStr,"[^0-9A-Fa-f]",0,2) = True Then
			CheckHex = 3
		End If
	Case 1	'HexDsc Mode
		Select Case CodePage
		Case CP_UNICODELITTLE, CP_UNICODEBIG
			If Ignore = 0 Then
				If (Len(textStr) Mod 6) <> 0 Then
					CheckHex = 1
				ElseIf CheckStrRegExp(textStr,"[^\\u0-9A-Fa-f]",0,2,True) = True Then
					CheckHex = 3
				End If
			ElseIf CheckStrRegExp(textStr,"\\u[0-9A-Fa-f]{4}",0,2,True) = False Then
				CheckHex = 4
			ElseIf CheckStrRegExp(textStr,"[^\x00-\x7F]+",0,2,True) = True Then
				CheckHex = 3
			End If
		Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
			If Ignore = 0 Then
				If (Len(textStr) Mod 10) <> 0 Then
					CheckHex = 1
				ElseIf CheckStrRegExp(textStr,"[^\\u0-9A-Fa-f]",0,2,True) = True Then
					CheckHex = 3
				End If
			ElseIf CheckStrRegExp(textStr,"\\u[0-9A-Fa-f]{8}",0,2,True) = False Then
				CheckHex = 4
			ElseIf CheckStrRegExp(textStr,"[^\x00-\x7F]+",0,2,True) = True Then
				CheckHex = 3
			End If
		Case Else
			If Ignore = 0 Then
				If (Len(textStr) Mod 4) <> 0 Then
					CheckHex = 1
				ElseIf CheckStrRegExp(textStr,"[^\\x0-9A-Fa-f]",0,2,True) = True Then
					CheckHex = 3
				End If
			ElseIf CheckStrRegExp(textStr,"\\x[0-9A-Fa-f]{2}",0,2,True) = False Then
				CheckHex = 4
			ElseIf CheckStrRegExp(textStr,"[^\x00-\x7F]+",0,2,True) = True Then
				CheckHex = 3
			End If
		End Select
	Case 2	'RULDec Mode
		Select Case CodePage
		Case CP_UNICODELITTLE, CP_UNICODEBIG
			If Ignore = 0 Then
				If (Len(textStr) Mod 6) <> 0 Then
					CheckHex = 1
				ElseIf CheckStrRegExp(textStr,"[^%u0-9A-Fa-f]",0,2,True) = True Then
					CheckHex = 3
				End If
			ElseIf CheckStrRegExp(textStr,"%u[0-9A-Fa-f]{4}",0,2,True) = False Then
				CheckHex = 4
			ElseIf CheckStrRegExp(textStr,"[^\x21\x23\x24\x26-\x3B\x3D\x3F-\x5A\x5F\x61-\x7A\x7E]+",0,2,True) = True Then
				CheckHex = 3
			End If
		Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
			If Ignore = 0 Then
				If (Len(textStr) Mod 10) <> 0 Then
					CheckHex = 1
				ElseIf CheckStrRegExp(textStr,"[^%u0-9A-Fa-f]",0,2,True) = True Then
					CheckHex = 3
				End If
			ElseIf CheckStrRegExp(textStr,"%u[0-9A-Fa-f]{8}",0,2,True) = False Then
				CheckHex = 4
			ElseIf CheckStrRegExp(textStr,"[^\x21\x23\x24\x26-\x3B\x3D\x3F-\x5A\x5F\x61-\x7A\x7E]+",0,2,True) = True Then
				CheckHex = 3
			End If
		Case Else
			If Ignore = 0 Then
				If (Len(textStr) Mod 3) <> 0 Then
					CheckHex = 1
				ElseIf CheckStrRegExp(textStr,"[^%0-9A-Fa-f]",0,2,True) = True Then
					CheckHex = 3
				End If
			ElseIf CheckStrRegExp(textStr,"%[0-9A-Fa-f]{2}",0,2,True) = False Then
				CheckHex = 4
			ElseIf CheckStrRegExp(textStr,"[^\x21\x23\x24\x26-\x3B\x3D\x3F-\x5A\x5F\x61-\x7A\x7E]+",0,2,True) = True Then
				CheckHex = 3
			End If
		End Select
	Case 3	'HTMLDec Mode
		If Ignore = 0 Then
			If CheckStrRegExp(textStr,"[^&#0-9;]",0,2,True) = True Then
				CheckHex = 3
			ElseIf CheckStrRegExp(textStr,"&#[0-9]{2,5};",0,5,True) = False Then
				CheckHex = 2
			End If
		ElseIf CheckStrRegExp(textStr,"&#[0-9]{2,5};",0,2,True) = False Then
			CheckHex = 4
		ElseIf CheckStrRegExp(textStr,"&#[0-9]{0,1};",0,2,True) = True Then
			CheckHex = 2
		ElseIf CheckStrRegExp(textStr,"&#[0-9]{6,};",0,2,True) = True Then
			CheckHex = 2
		ElseIf CheckStrRegExp(textStr,"[^\x00-\x7F]+",0,2,True) = True Then
			CheckHex = 3
		End If
	Case 4	'ISO-8859-1Dec Mode
		If Ignore = 0 Then
			If CheckStrRegExp(textStr,"[^#0-9;]",0,2,True) = True Then
				CheckHex = 3
			ElseIf CheckStrRegExp(textStr,"#[0-9]{2,5}",0,5,True) = False Then
				CheckHex = 2
			End If
		ElseIf CheckStrRegExp(textStr,"#[0-9]{2,5}",0,2,True) = False Then
			CheckHex = 4
		ElseIf CheckStrRegExp(textStr,"#[0-9]{0,1}#",0,2,True) = True Then
			CheckHex = 2
		ElseIf CheckStrRegExp(textStr,"#[0-9]{6,}#",0,2,True) = True Then
			CheckHex = 2
		ElseIf CheckStrRegExp(textStr,"[^\x00-\x7F]+",0,2,True) = True Then
			CheckHex = 3
		End If
	Case 5	'Base64 Mode
		If CheckStrRegExp(textStr,"[^A-Za-z0-9+/=]+",0,2,True) = True Then CheckHex = 3
	End Select
End Function


'字串常数正向转换
Private Function Convert(ByVal ConverString As String) As String
	Convert = ConverString
	If Convert = "" Then Exit Function
	If InStr(Convert,"\") = 0 Then Exit Function
	If InStr(Convert,"\\") Then Convert = Replace$(Convert,"\\","*a!N!d*")
	If InStr(Convert,"\r\n") Then Convert = Replace$(Convert,"\r\n",vbCrLf)
	If InStr(Convert,"\r\n") Then Convert = Replace$(Convert,"\r\n",vbNewLine)
	If InStr(Convert,"\r") Then Convert = Replace$(Convert,"\r",vbCr)
	If InStr(Convert,"\r") Then Convert = Replace$(Convert,"\r",vbNewLine)
	If InStr(Convert,"\n") Then Convert = Replace$(Convert,"\n",vbLf)
	If InStr(Convert,"\b") Then Convert = Replace$(Convert,"\b",vbBack)
	If InStr(Convert,"\f") Then Convert = Replace$(Convert,"\f",vbFormFeed)
	If InStr(Convert,"\v") Then Convert = Replace$(Convert,"\v",vbVerticalTab)
	If InStr(Convert,"\t") Then Convert = Replace$(Convert,"\t",vbTab)
	If InStr(Convert,"\'") Then Convert = Replace$(Convert,"\'","'")
	If InStr(Convert,"\""") Then Convert = Replace$(Convert,"\""","""")
	If InStr(Convert,"\?") Then Convert = Replace$(Convert,"\?","?")
	If InStr(Convert,"\") Then Convert = ConvertB(Convert)
	If InStr(Convert,"\0") Then Convert = Replace$(Convert,"\0",vbNullChar)
	If InStr(Convert,"*a!N!d*") Then Convert = Replace$(Convert,"*a!N!d*","\")
End Function


'转换八进制或十六进制转义符
Private Function ConvertB(ByVal ConverString As String) As String
	Dim i As Long,j As Long,EscStr As String
	ConvertB = ConverString
	i = InStr(ConvertB,"\")
	Do While i > 0
		EscStr = Mid$(ConvertB,i,2)
		Select Case EscStr
		Case "\x", "\X"
			ConverString = Mid$(ConvertB,i + 2,2)
			If CheckStrRegExp(ConverString,"[0-9A-Fa-f]",0,1,True) = True Then
				j = Val("&H" & ConverString)
				ConvertB = Replace$(ConvertB,EscStr & ConverString,Val2Bytes(j,2))
			End If
		Case "\u", "\U"
			ConverString = Mid$(ConvertB,i + 2,4)
			If CheckStrRegExp(UCase$(ConverString),"[0-9A-Fa-f]",0,1) = True Then
				j = Val("&H" & ConverString)
				ConvertB = Replace$(ConvertB,EscStr & ConverString,Val2Bytes(j,2))
			End If
		Case Is <> ""
			EscStr = "\"
			For j = 3 To 1 Step -1
				ConverString = Mid$(ConvertB,i + 1,j)
				If CheckStrRegExp(ConverString,"[0-7]",0,1) = True Then
					j = Val("&O" & ConverString)
					If j > 256 Then
						ConverString = Left$(ConverString,2)
						j = Val("&O" & ConverString)
					End If
					ConvertB = Replace$(ConvertB,EscStr & ConverString,Val2Bytes(j,2))
					Exit For
				End If
			Next j
		End Select
		i = InStr(i + 1,ConvertB,"\")
	Loop
End Function


'字串常数反向转换
'Mode = 0 按 PSL 版本的宏引擎不同分别转义控制字符和全部或部分拉丁文扩展字符
'Mode = 1 转义控制字符和所有拉丁文扩展字符
'Mode = 2 转义控制字符和操作系统不能显示的拉丁文扩展字符
'Mode = 3 仅转义控制字符
Private Function ReConvert(ByVal ConverString As String,Optional ByVal Mode As Integer) As String
	ReConvert = ConverString
	If ReConvert = "" Then Exit Function
	If InStr(ReConvert,"\") Then ReConvert = Replace$(ReConvert,"\","\\")
	If InStr(ReConvert,vbCrLf) Then ReConvert = Replace$(ReConvert,vbCrLf,"\r\n")
	If InStr(ReConvert,vbNewLine) Then ReConvert = Replace$(ReConvert,vbNewLine,"\r\n")
	If InStr(ReConvert,vbCr) Then ReConvert = Replace$(ReConvert,vbCr,"\r")
	If InStr(ReConvert,vbLf) Then ReConvert = Replace$(ReConvert,vbLf,"\n")
	If InStr(ReConvert,vbBack) Then ReConvert = Replace$(ReConvert,vbBack,"\b")
	If InStr(ReConvert,vbFormFeed) Then ReConvert = Replace$(ReConvert,vbFormFeed,"\f")
	If InStr(ReConvert,vbVerticalTab) Then ReConvert = Replace$(ReConvert,vbVerticalTab,"\v")
	If InStr(ReConvert,vbTab) Then ReConvert = Replace$(ReConvert,vbTab,"\t")
	If InStr(ReConvert,vbNullChar) Then ReConvert = Replace$(ReConvert,vbNullChar,"\0")
	ReConvert = ReConvertBRegExp(ReConvert,Mode)
End Function


'转换拉丁文扩展字符为十六进制转义符
'Mode = 0 按 PSL 版本的宏引擎不同分别转义控制字符和全部或部分拉丁文扩展字符
'Mode = 1 转义控制字符和所有拉丁文扩展字符
'Mode = 2 转义控制字符和操作系统不能显示的拉丁文扩展字符
'Mode = 3 仅转义控制字符
Private Function ReConvertB(ByVal ConverString As String,Optional ByVal Mode As Integer) As String
	Dim i As Long,Dec As Long,Temp As String
	ReConvertB = ConverString
	If Mode = 0 Then
		Mode = IIf(Int(PSL.Version / 100) < 15,1,2)
	End If
	Select Case Mode
	Case 1
		For i = 1 To Len(ConverString)
			Temp = Mid$(ConverString,i,1)
			If InStr(ReConvertB,Temp) Then
				Dec = AscW(Temp)
				Select Case Dec
				Case 0 To 31,127 To 255
					ReConvertB = Replace$(ReConvertB,Temp,"\x" & Right$("0" & Hex$(Dec),2))
				End Select
			End If
		Next i
	Case 2
		For i = 1 To Len(ConverString)
			Temp = Mid$(ConverString,i,1)
			If InStr(ReConvertB,Temp) Then
				Dec = AscW(Temp)
				Select Case Dec
				Case 0 To 31,127,129,141,143,144,157,173
					ReConvertB = Replace$(ReConvertB,Temp,"\x" & Right$("0" & Hex$(Dec),2))
				End Select
			End If
		Next i
	Case 3
		For i = 1 To Len(ConverString)
			Temp = Mid$(ConverString,i,1)
			If InStr(ReConvertB,Temp) Then
				Dec = AscW(Temp)
				Select Case Dec
				Case 0 To 31,127
					ReConvertB = Replace$(ReConvertB,Temp,"\x" & Right$("0" & Hex$(Dec),2))
				End Select
			End If
		Next i
	End Select
End Function


'转换拉丁文扩展字符为十六进制转义符
'Mode = 0 按 PSL 版本的宏引擎不同分别转义控制字符和全部或部分拉丁文扩展字符
'Mode = 1 转义控制字符和所有拉丁文扩展字符
'Mode = 2 转义控制字符和操作系统不能显示的拉丁文扩展字符
'Mode = 3 仅转义控制字符
Private Function ReConvertBRegExp(ByVal ConverString As String,Optional ByVal Mode As Integer) As String
	Dim i As Long,Matches As Object
	ReConvertBRegExp = ConverString
	If Mode = 0 Then
		Mode = IIf(Int(PSL.Version / 100) < 15,1,2)
	End If
	With RegExp
		.Global = True
		.IgnoreCase = True
		Select Case Mode
		Case 1
			.Pattern = "[\x01-\x1F\x7F-\xFF]"
		Case 2
			.Pattern = "[\x01-\x1F\x7F\x81\x8D\x8F\x90\x9D\xAD]"
		Case 3
			.Pattern = "[\x01-\x1F\x7F]"
		End Select
		Set Matches = .Execute(ConverString)
		If Matches.Count > 0 Then
			For i = 0 To Matches.Count - 1
				ConverString = Right$("0" & Hex$(AscW(Matches(i).Value)),2)
				ReConvertBRegExp = Replace$(ReConvertBRegExp,Matches(i).Value,"\x" & ConverString)
			Next i
		End If
	End With
End Function


'检查字串是否包含指定字符(正则表达式比较)
'Mode = 0 检查字串是否包含指定字符，并找出指定字符的位置
'Mode = 1 检查字串是否只包含指定字符
'Mode = 2 检查字串是否包含指定字符
'Mode = 3 检查字串是否只包含大小混写的指定字符，此时 IgnoreCase 参数无效
'Mode = 4 检查字串是否有连续相同的字符，StrNum 为最少重复字符个数
'Mode = 5 检查字串是否包含指定字串，并返回匹配的字串总长度 (适合字符组合查询)
'Patrn  为正则表达式模板
Private Function CheckStrRegExp(ByVal textStr As String,ByVal Patrn As String,Optional ByVal StrNum As Long, _
				Optional ByVal Mode As Long,Optional ByVal IgnoreCase As Boolean) As Long
	Dim n As Long,Matches As Object
	If Patrn = "" Then Exit Function
	If Trim$(textStr) = "" Then Exit Function
	With RegExp
		Select Case Mode
		Case 0
			.Global = True
			.IgnoreCase = IgnoreCase
			.Pattern = Patrn
			Set Matches = .Execute(textStr)
			If Matches.Count > 0 Then CheckStrRegExp = Matches(0).FirstIndex + 1
		Case 1
			.Global = True
			.IgnoreCase = IgnoreCase
			.Pattern = Patrn
			Set Matches = .Execute(textStr)
			If Matches.Count = Len(textStr) Then CheckStrRegExp = True
		Case 2
			.Global = False
			.IgnoreCase = IgnoreCase
			.Pattern = Patrn
			If .Test(textStr) Then CheckStrRegExp = True
		Case 3
			If InStr(textStr," ") Then Exit Function
			n = Len(textStr)
			If n < 2 Then Exit Function
			If LCase$(textStr) = textStr Then Exit Function
			If UCase$(textStr) = textStr Then Exit Function
			.Global = True
			.IgnoreCase = False
			.Pattern = Patrn
			Set Matches = .Execute(textStr)
			If Matches.Count <> n Then Exit Function
			textStr = Mid$(textStr,2)
			If LCase$(textStr) = textStr Then Exit Function
			CheckStrRegExp = True
		Case 4
			If StrNum < 2 Then Exit Function
			If Len(textStr) < StrNum Then Exit Function
			If InStr(textStr," ") Then Exit Function
			If StrNum = 3 Then
				If InStr(textStr,"://www.") Then Exit Function
			End If
			.Global = False
			.IgnoreCase = IgnoreCase
			.Pattern = "(" & Patrn & ")\1{" & CStr$(StrNum - 1) & ",}"
			If .Test(textStr) Then CheckStrRegExp = True
		Case 5
			.Global = True
			.IgnoreCase = IgnoreCase
			.Pattern = Patrn
			Set Matches = .Execute(textStr)
			If Matches.Count = 0 Then Exit Function
			For n = 0 To Matches.Count - 1
				CheckStrRegExp = CheckStrRegExp + Matches(n).Length
			Next n
			CheckStrRegExp = IIf(CheckStrRegExp = Len(textStr),True,False)
		End Select
	End With
End Function


'按字符编码计算字符的十六进制字节长度
'Mode = 1 转义, 否则不转义
Private Function StrHexLength(ByVal textStr As String,ByVal CodePage As Long,ByVal Mode As Long) As Long
	If textStr = "" Then Exit Function
	If Mode = 1 Then textStr = Convert(textStr)
	Select Case CodePage
	Case CP_UNICODELITTLE, CP_UNICODEBIG
		'Bin = textStr
		'StrHexLength = UBound(Bin) + 1
		StrHexLength = LenB(textStr)
	Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
		StrHexLength = LenB(textStr) * 2
	Case Else
		ReDim Bin(0) As Byte
		StrHexLength = WideCharToMultiByte(CodePage, 0&, StrPtr(textStr), Len(textStr), Bin(0), 0, 0, 0)
	End Select
End Function


'解析命令行参数
'' Command Line Format: Command <Source><-><Translation> <Switch>
'' Command: Name of this Macros file
''<Source>
'' String: The source string to be converted.
''<Translation>
'' String: The translation string to be converted.
''<->: This is the delimiter between the source string and the translation string.
''<Switch>
''Codepage:
'' -scp[N]: Codepage value of Source Text. N is Numeric value of codepage. Such as: 936,1200 et.
'' -tcp[N]: Codepage value of Translation Text. N is Numeric value of codepage. Such as: 936,1200 et.
''Sring escape:
'' -se: Escape text before convert text to code or after converted code to text. No this switch, do not escape text
''Escape type:
'' -et[N]: Escape type. N is 0 = Hex (Default), 1 = Hex escape, 2 = RUL Encode, 3 = HTML Encode, 4 = ISO-8829-1 Encode, 5 = Base4 Encode
''Multibyte only:
'' -ch: Every 2 Hex characters separated by spaces for Hex type
'' -ca: Convert ASCII characters to Hex escape or HTML escape. No this switch, do not convert ASCII characters
'' -ci: Convert illegal character to URL escape. No this switch, do not convert illegal character
''Convert mode:
'' -ac: Auto convert string to code or code to string after enter the data. No this switch, manually convert for UI mode only
''Convert type:
'' -cs: Convert code to String. No this switch, Convert string to code
''Line break type:
'' -lb[N]: No this switch to use vbCrLf by default, Otherwise, Use the specified as: 0 = vbCrLf, 1 = vbCr, 2 = vbLf
''Fill type:
'' -fz: By source text length in bytes, padded with null bytes.
'' -fs: By source text length in bytes, padded with null characters.
''Both, if the translation byte is longer than the source byte, will be truncated to be the same as the source byte.
''UI mode:
'' -noui: Do not display a user interface, run silently
''Display Option:
'' -td: Frist display translation windows. No this switch, frist display source windows,
'' -lng[hex language code]: Display UI Language. Supports EngLish, Chinese Simplified and Chinese Traditional only. For sample: 0804,1004,0404,0C04,1404.

'' Return: None
'' For example: modEncodeQuery.bas This is strings.<->This is converted Hex code. -cp:1201 -se -sc -ac -et:1

Private Function SplitArgument(ByVal Argument As String,ByVal MaxNum As Long) As String()
	Dim i As Long,j As Long,k As Long,TempList() As String
	ReDim ArgArray(MaxNum - 1) As String
	If Argument = "" Then
		SplitArgument = ArgArray
		Exit Function
	End If
	'从后向前查找数值项参数的最小索引
	TempList = ReSplit(Argument," ")
	k = UBound(TempList)
	MaxNum = IIf(k - MaxNum > 0,k - MaxNum,1)
	For i = k To MaxNum Step -1
		Argument = Trim$(TempList(i))
		If CheckStrRegExp(Argument,"-scp:[0-9]+",0,5,True) = True Then
			Argument = "-scp:"
		ElseIf CheckStrRegExp(Argument,"-tcp:[0-9]+",0,5,True) = True Then
			Argument = "-tcp:"
		ElseIf CheckStrRegExp(Argument,"-et:[0-3]",0,5,True) = True Then
			Argument = "-et:"
		ElseIf CheckStrRegExp(Argument,"-lng:[0-9a-f;]+",0,5,True) = True Then
			Argument = "-lng:"
		End If
		Select Case Argument
		Case "-scp:"
			If ArgArray(2) = "" Then
				ArgArray(2) = Mid$(Trim$(TempList(i)),6)
				j = j + 1
			End If
		Case "-tcp:"
			If ArgArray(3) = "" Then
				ArgArray(3) = Mid$(Trim$(TempList(i)),6)
				j = j + 1
			End If
		Case "-se"
			If ArgArray(4) = "" Then
				ArgArray(4) = "1"
				j = j + 1
			End If
		Case "-et:"
			If ArgArray(5) = "" Then
				ArgArray(5) = Mid$(Trim$(TempList(i)),5)
				j = j + 1
			End If
		Case "-ch","-ca","-ci"
			If ArgArray(6) = "" Then
				ArgArray(6) = "1"
				j = j + 1
			End If
		Case "-ac"
			If ArgArray(7) = "" Then
				ArgArray(7) = "1"
				j = j + 1
			End If
		Case "-cs"
			If ArgArray(8) = "" Then
				ArgArray(8) = "1"
				TempList(i) = ""
				j = j + 1
			End If
		Case "-fz"
			If ArgArray(9) = "" Then
				ArgArray(9) = "1"
				j = j + 1
			End If
		Case "-fs"
			If ArgArray(9) = "" Then
				ArgArray(9) = "2"
				j = j + 1
			End If
		Case "-noui"
			If ArgArray(10) = "" Then
				ArgArray(10) = "1"
				j = j + 1
			End If
		Case "-td"
			If ArgArray(11) = "" Then
				ArgArray(11) = "1"
				j = j + 1
			End If
		Case "-lb:"
			If ArgArray(12) = "" Then
				ArgArray(12) = Mid$(Trim$(TempList(i)),5)
				j = j + 1
			End If
		Case "-lng:"
			If ArgArray(13) = "" Then
				ArgArray(13) = Mid$(Trim$(TempList(i)),6)
				j = j + 1
			End If
		End Select
	Next i
	ReDim Preserve TempList(k - j) As String
	TempList = ReSplit(StrListJoin(TempList," "),"<->",2)
	Select Case UBound(TempList)
	Case 0
		ArgArray(0) = TempList(0)
	Case 1
		ArgArray(0) = TempList(0)
		ArgArray(1) = TempList(1)
	End Select
	SplitArgument = ArgArray
End Function


'转换字符为整数数值
Private Function StrToLong(ByVal mStr As String,Optional ByVal DefaultValue As Long) As Long
	On Error GoTo errHandle
	StrToLong = CLng(mStr)
	Exit Function
	errHandle:
	StrToLong = DefaultValue
End Function


'检查字串数组是否为空，非空返回 True
Private Function CheckArray(DataList() As String) As Boolean
	Dim i As Long
	On Error GoTo errHandle
	For i = LBound(DataList) To UBound(DataList)
		If DataList(i) <> "" Then
			CheckArray = True
			Exit For
		End If
	Next i
	errHandle:
End Function


'格式化 HEX 字串
Private Function FormatHexStr(ByVal textStr As String,ByVal Length As Integer) As String
	If textStr = "" Then Exit Function
	If (Len(textStr) Mod Length) = 0 Then
		FormatHexStr = textStr
		Exit Function
	End If
	FormatHexStr = "0" & textStr
End Function


'转换数值为字节数组(短于长度的高位截断)
Private Function Val2Bytes(ByVal Value As Long,ByVal Length As Integer,Optional ByVal ByteOrder As Boolean) As Byte()
	On Error GoTo errHandle
	ReDim Bytes(Length - 1) As Byte
	CopyMemory Bytes(0), Value, Length
	If ByteOrder = True Then
		Val2Bytes = ReverseValByte(Bytes,0,-1)
	Else
		Val2Bytes = Bytes
	End If
	Exit Function
	errHandle:
	ReDim Bytes(0) As Byte
	Val2Bytes = Bytes
End Function


'每二个字符空格分隔
Private Function SeparatHex(ByVal textStr As String) As String
	Dim i As Long,j As Long,n As Long
	j = Len(textStr)
	SeparatHex = Space$(j * 2)
	n = 1
	For i = 1 To j - 1 Step 2
		Mid$(SeparatHex,n,2) = Mid(textStr,i,2) & " "
		n = n + 3
	Next i
	SeparatHex = Trim$(SeparatHex)
End Function


'字符串转 Hex
'FillMode: 0 = 按实际所需字节转换
'          1 = 超长时后面截断，否则后面用零字节填满
'          2 = 超长时前面截断，否则前面用零字节填满
'          3 = 超长时后面截断，否则后面用空字节填满
'          4 = 超长时前面截断，否则前面用空字节填满
'ByteLength 为需要的字节长度，仅在截断时需要
Private Function Str2Hex(ByVal textStr As String,ByVal CodePage As Long,ByVal FillMode As Long,Optional ByVal ByteLength As Long) As String
	Select Case FillMode
	Case 0
		Str2Hex = Byte2Hex(StringToByte(textStr,CodePage),0,-1)
	Case 1
		Select Case ByteLength - StrHexLength(textStr,CodePage,0)
		Case 0
			Str2Hex = Byte2Hex(StringToByte(textStr,CodePage),0,-1)
		Case Is < 0
			Call IsDBCSLeadPos(textStr,CodePage,ByteLength,True,False)
			Str2Hex = Byte2Hex(StringToByte(textStr,CodePage),0,-1)
		Case Is > 0
			ReDim Bytes(0) As Byte
			Bytes = StringToByte(textStr,CodePage)
			ReDim Preserve Bytes(ByteLength - 1) As Byte
			Str2Hex = Byte2Hex(Bytes,0,-1)
		End Select
	Case 2
		Select Case ByteLength - StrHexLength(textStr,CodePage,0)
		Case 0
			Str2Hex = Byte2Hex(StringToByte(textStr,CodePage),0,-1)
		Case Is < 0
			Call IsDBCSLeadPos(textStr,CodePage,ByteLength,True,True)
			Str2Hex = Byte2Hex(StringToByte(textStr,CodePage),0,-1)
		Case Is > 0
			ReDim Bytes(ByteLength - 1) As Byte,TempB(0) As Byte
			TempB = StringToByte(textStr,CodePage)
			ByteLength = ByteLength - UBound(TempB) - 1
			CopyMemory Bytes(ByteLength), TempB(0), UBound(TempB) + 1
			Str2Hex = Byte2Hex(Bytes,0,-1)
		End Select
	Case 3
		Call FillStrWithSpape(textStr,CodePage,ByteLength,True,False)
		Str2Hex = Byte2Hex(StringToByte(textStr,CodePage),0,-1)
	Case 4
		Call FillStrWithSpape(textStr,CodePage,ByteLength,True,True)
		Str2Hex = Byte2Hex(StringToByte(textStr,CodePage),0,-1)
	End Select
End Function


'字符串转 Hex 转义符
Private Function Str2HexEsc(ByVal textStr As String,ByVal CodePage As Long,ByVal MultibyteOnly As Long) As String
	If textStr = "" Then Exit Function
	If MultibyteOnly = 0 Then
		Str2HexEsc = Byte2HexEsc(StringToByte(textStr,CodePage),0,-1,CodePage)
		Exit Function
	End If
	Dim i As Long,Matches As Object
	Str2HexEsc = textStr
	RegExp.Global = True
	RegExp.IgnoreCase = False
	RegExp.Pattern = "[^\x00-\x7F]+"
	Set Matches = RegExp.Execute(textStr)
	If Matches.Count = 0 Then Exit Function
	For i = Matches.Count - 1 To 0 Step -1
		With Matches(i)
			If .FirstIndex > 0 Then
				Str2HexEsc = Left$(Str2HexEsc,.FirstIndex) & _
							Byte2HexEsc(StringToByte(.Value,CodePage),0,-1,CodePage) & _
							Mid$(Str2HexEsc,.FirstIndex + .Length + 1)
			Else
				Str2HexEsc = Byte2HexEsc(StringToByte(.Value,CodePage),0,-1,CodePage) & _
							Mid$(Str2HexEsc,.Length + 1)
			End If
		End With
	Next i
End Function


'字符串转 URL 转义符
Private Function Str2URLEsc(ByVal textStr As String,ByVal CodePage As Long,ByVal MultibyteOnly As Long) As String
	If textStr = "" Then Exit Function
	If MultibyteOnly = 0 Then
		Str2URLEsc = Byte2URLEsc(StringToByte(textStr,CodePage),0,-1,CodePage)
		Exit Function
	End If
	Dim i As Long,Matches As Object
	Str2URLEsc = textStr
	RegExp.Global = True
	RegExp.IgnoreCase = False
	'RegExp.Pattern = "[\x00-x20\x22\x25\x3C\x3E\x5B-\x5E\x60\x7B-\x7D\x7F]+"
	RegExp.Pattern = "[^\x21\x23\x24\x26-\x3B\x3D\x3F-\x5A\x5F\x61-\x7A\x7E]+"
	Set Matches = RegExp.Execute(textStr)
	If Matches.Count = 0 Then Exit Function
	For i = Matches.Count - 1 To 0 Step -1
		With Matches(i)
			If .FirstIndex > 0 Then
				Str2URLEsc = Left$(Str2URLEsc,.FirstIndex) & _
							Byte2URLEsc(StringToByte(.Value,CodePage),0,-1,CodePage) & _
							Mid$(Str2URLEsc,.FirstIndex + .Length + 1)
			Else
				Str2URLEsc = Byte2URLEsc(StringToByte(.Value,CodePage),0,-1,CodePage) & _
							Mid$(Str2URLEsc,.Length + 1)
			End If
		End With
	Next i
End Function


'字符串转 HTML 转义符
Private Function Str2HTMLEsc(ByVal textStr As String,ByVal CodePage As Long,ByVal MultibyteOnly As Long) As String
	If textStr = "" Then Exit Function
	If MultibyteOnly = 0 Then
		Str2HTMLEsc = Byte2HTMLEsc(StringToByte(textStr,CodePage),0,-1,CodePage)
		Exit Function
	End If
	Dim i As Long,Matches As Object
	Str2HTMLEsc = textStr
	RegExp.Global = True
	RegExp.IgnoreCase = False
	RegExp.Pattern = "[^\x00-\x7F]+"
	Set Matches = RegExp.Execute(textStr)
	If Matches.Count = 0 Then Exit Function
	For i = Matches.Count - 1 To 0 Step -1
		With Matches(i)
			If .FirstIndex > 0 Then
				Str2HTMLEsc = Left$(Str2HTMLEsc,.FirstIndex) & _
							Byte2HTMLEsc(StringToByte(.Value,CodePage),0,-1,CodePage) & _
							Mid$(Str2HTMLEsc,.FirstIndex + .Length + 1)
			Else
				Str2HTMLEsc = Byte2HTMLEsc(StringToByte(.Value,CodePage),0,-1,CodePage) & _
							Mid$(Str2HTMLEsc,.Length + 1)
			End If
		End With
	Next i
End Function


'字符串转 ISO88591 编码
Private Function Str2ISOEsc(ByVal textStr As String,ByVal CodePage As Long,ByVal MultibyteOnly As Long) As String
	If textStr = "" Then Exit Function
	If MultibyteOnly = 0 Then
		Str2ISOEsc = Byte2ISOEsc(StringToByte(textStr,CodePage),0,-1,CodePage)
		Exit Function
	End If
	Dim i As Long,Matches As Object
	Str2ISOEsc = textStr
	RegExp.Global = True
	RegExp.IgnoreCase = False
	RegExp.Pattern = "[^\x00-\x7F]+"
	Set Matches = RegExp.Execute(textStr)
	If Matches.Count = 0 Then Exit Function
	For i = Matches.Count - 1 To 0 Step -1
		With Matches(i)
			If .FirstIndex > 0 Then
				Str2ISOEsc = Left$(Str2ISOEsc,.FirstIndex) & _
							Byte2ISOEsc(StringToByte(.Value,CodePage),0,-1,CodePage) & _
							Mid$(Str2ISOEsc,.FirstIndex + .Length + 1)
			Else
				Str2ISOEsc = Byte2ISOEsc(StringToByte(.Value,CodePage),0,-1,CodePage) & _
							Mid$(Str2ISOEsc,.Length + 1)
			End If
		End With
	Next i
End Function


'Hex 转义符转字符串
Private Function HexEsc2Str(ByVal textStr As String,ByVal CodePage As Long) As String
	Dim i As Long,Temp As String,Matches As Object
	If textStr = "" Then Exit Function
	HexEsc2Str = textStr
	RegExp.Global = True
	RegExp.IgnoreCase = True
	Select Case CodePage
	Case CP_UNICODELITTLE, CP_UNICODEBIG
		RegExp.Pattern = "(\\u[0-9A-Fa-f]{4})+"
		Set Matches = RegExp.Execute(textStr)
		If Matches.Count = 0 Then Exit Function
		CodePage = CP_UNICODELITTLE
		For i = 0 To Matches.Count - 1
			Temp = Replace$(LCase$(Matches(i).Value),"\u","")
			Temp = ByteToString(LowByte2HighByte(HexStr2Bytes(Temp),2),CodePage)
			HexEsc2Str = Replace$(HexEsc2Str,Matches(i).Value,Temp,,1)
		Next i
	Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
		RegExp.Pattern = "(\\u[0-9A-Fa-f]{8})+"
		Set Matches = RegExp.Execute(textStr)
		If Matches.Count = 0 Then Exit Function
		CodePage = CP_UTF32LE
		For i = 0 To Matches.Count - 1
			Temp = Replace$(LCase$(Matches(i).Value),"\u","")
			Temp = ByteToString(LowByte2HighByte(HexStr2Bytes(Temp),4),CodePage)
			HexEsc2Str = Replace$(HexEsc2Str,Matches(i).Value,Temp,,1)
		Next i
	Case Else
		RegExp.Pattern = "(\\x[0-9A-Fa-f]{2})+"
		Set Matches = RegExp.Execute(textStr)
		If Matches.Count = 0 Then Exit Function
		For i = 0 To Matches.Count - 1
			Temp = Replace$(LCase$(Matches(i).Value),"\x","")
			Temp = ByteToString(HexStr2Bytes(Temp),CodePage)
			HexEsc2Str = Replace$(HexEsc2Str,Matches(i).Value,Temp,,1)
		Next i
	End Select
End Function


'URL 转义符转字符串
Private Function URLEsc2Str(ByVal textStr As String,ByVal CodePage As Long) As String
	Dim i As Long,Temp As String,Matches As Object
	If textStr = "" Then Exit Function
	URLEsc2Str = textStr
	RegExp.Global = True
	RegExp.IgnoreCase = True
	Select Case CodePage
	Case CP_UNICODELITTLE, CP_UNICODEBIG
		RegExp.Pattern = "(%u[0-9A-Fa-f]{4})+"
		Set Matches = RegExp.Execute(textStr)
		If Matches.Count = 0 Then Exit Function
		CodePage = CP_UNICODELITTLE
		For i = 0 To Matches.Count - 1
			Temp = Replace$(LCase$(Matches(i).Value),"%u","")
			Temp = ByteToString(LowByte2HighByte(HexStr2Bytes(Temp),2),CodePage)
			URLEsc2Str = Replace$(URLEsc2Str,Matches(i).Value,Temp,,1)
		Next i
	Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
		RegExp.Pattern = "(%u[0-9A-Fa-f]{8})+"
		Set Matches = RegExp.Execute(textStr)
		If Matches.Count = 0 Then Exit Function
		CodePage = CP_UTF32LE
		For i = 0 To Matches.Count - 1
			Temp = Replace$(LCase$(Matches(i).Value),"%u","")
			Temp = ByteToString(LowByte2HighByte(HexStr2Bytes(Temp),4),CodePage)
			URLEsc2Str = Replace$(URLEsc2Str,Matches(i).Value,Temp,,1)
		Next i
	Case Else
		RegExp.Pattern = "(%[0-9A-Fa-f]{2})+"
		Set Matches = RegExp.Execute(textStr)
		If Matches.Count = 0 Then Exit Function
		For i = 0 To Matches.Count - 1
			Temp = Replace$(Matches(i).Value,"%","")
			Temp = ByteToString(HexStr2Bytes(Temp),CodePage)
			URLEsc2Str = Replace$(URLEsc2Str,Matches(i).Value,Temp,,1)
		Next i
	End Select
End Function


'HTML 转义符转字符串
Private Function HTMLEsc2Str(ByVal textStr As String,ByVal CodePage As Long) As String
	Dim i As Long,j As Long,TempList() As String,Matches As Object
	If textStr = "" Then Exit Function
	HTMLEsc2Str = textStr
	RegExp.Global = True
	RegExp.IgnoreCase = False
	RegExp.Pattern = "(&#[0-9]{2,5};)+"
	Set Matches = RegExp.Execute(textStr)
	If Matches.Count = 0 Then Exit Function
	Select Case CodePage
	Case CP_UNICODELITTLE, CP_UNICODEBIG, CP_UTF32LE, CP_UTF32BE
		CodePage = CP_UNICODELITTLE
		For i = 0 To Matches.Count - 1
			TempList = ReSplit(Replace$(Matches(i).Value,"&#",""),";")
			For j = 0 To UBound(TempList) - 1
				TempList(j) = Byte2Hex(Val2Bytes(CLng(TempList(j)),2),0,-1)
			Next j
			TempList(0) = ByteToString(HexStr2Bytes(StrListJoin(TempList,"")),CodePage)
			HTMLEsc2Str = Replace$(HTMLEsc2Str,Matches(i).Value,TempList(0),,1)
		Next i
	Case Else
		For i = 0 To Matches.Count - 1
			TempList = ReSplit(Replace$(Matches(i).Value,"&#",""),";")
			For j = 0 To UBound(TempList) - 1
				TempList(j) = Byte2Hex(Val2Bytes(CLng(TempList(j)),1),0,-1)
			Next j
			TempList(0) = ByteToString(HexStr2Bytes(StrListJoin(TempList,"")),CodePage)
			HTMLEsc2Str = Replace$(HTMLEsc2Str,Matches(i).Value,TempList(0),,1)
		Next i
	End Select
End Function


'ISO-88591-1 解码
Private Function ISOEsc2Str(ByVal textStr As String,ByVal CodePage As Long) As String
	Dim i As Long,j As Long,TempList() As String,Matches As Object
	If textStr = "" Then Exit Function
	ISOEsc2Str = textStr
	RegExp.Global = True
	RegExp.IgnoreCase = False
	RegExp.Pattern = "(#[0-9]{2,5})+"
	Set Matches = RegExp.Execute(textStr)
	If Matches.Count = 0 Then Exit Function
	Select Case CodePage
	Case CP_UNICODELITTLE, CP_UNICODEBIG, CP_UTF32LE, CP_UTF32BE
		CodePage = CP_UNICODELITTLE
		For i = 0 To Matches.Count - 1
			TempList = ReSplit(Matches(i).Value,"#")
			For j = 1 To UBound(TempList)
				TempList(j) = Byte2Hex(Val2Bytes(CLng(TempList(j)),2),0,-1)
			Next j
			TempList(0) = ByteToString(HexStr2Bytes(StrListJoin(TempList,"")),CodePage)
			ISOEsc2Str = Replace$(ISOEsc2Str,Matches(i).Value,TempList(0),,1)
		Next i
	Case Else
		For i = 0 To Matches.Count - 1
			TempList = ReSplit(Matches(i).Value,"#")
			For j = 1 To UBound(TempList)
				TempList(j) = Byte2Hex(Val2Bytes(CLng(TempList(j)),1),0,-1)
			Next j
			TempList(0) = ByteToString(HexStr2Bytes(StrListJoin(TempList,"")),CodePage)
			ISOEsc2Str = Replace$(ISOEsc2Str,Matches(i).Value,TempList(0),,1)
		Next i
	End Select
End Function


'Base64 编码
'例：对ABC进行BASE64 编码：
'1) 首先取ABC对应的ASCII码值。A(65)B(66)C(67)
'2) 再取二进制值A(01000001)B(01000010)C(01000011)
'3) 然后把这三个字节的二进制码接起来(010000010100001001000011)
'4) 再以6位为单位分成4个数据块,并在最高位填充两个0后形成4个字节的编码后的值，(00010000)(00010100)(00001001)(00000011)
'5) 再把这四个字节数据转化成10进制数得(16)(20)(9)(3)
'6) 最后根据BASE64给出的64个基本字符表，查出对应的ASCII码字符(Q)(U)(J)(D)，这里的值实际就是数据在字符表中的索引
'结果：ABC编码为QUJD
Private Function Base64Encode(ByVal strSource As String,ByVal CodePage As Long) As String
	Const cstBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
	Dim i As Long,Length As Long,Mods As Long,arrB() As Byte
	arrB = StringToByte(strSource,CodePage)
	'获取除以3的余数
	Length = UBound(arrB) + 1
	Mods = Length Mod 3
	Length = Length - Mods
	ReDim Buf(Length / 3 * 4 - 1 + IIf(Mods = 0,0,4)) As Byte
	For i = 0 To Length - 1 Step 3
		Buf(i / 3 * 4) = (arrB(i) And &HFC) / &H4
		Buf(i / 3 * 4 + 1) = (arrB(i) And &H3) * &H10 + (arrB(i + 1) And &HF0) / &H10
		Buf(i / 3 * 4 + 2) = (arrB(i + 1) And &HF) * &H4 + (arrB(i + 2) And &HC0) / &H40
		Buf(i / 3 * 4 + 3) = arrB(i + 2) And &H3F
	Next i
	'处理余数
	i = Length
	If Mods = 1 Then
		Buf(i / 3 * 4) = (arrB(i) And &HFC) / &H4
		Buf(i / 3 * 4 + 1) = (arrB(i) And &H3) * &H10
		Buf(i / 3 * 4 + 2) = 64
		Buf(i / 3 * 4 + 3) = 64
	ElseIf Mods = 2 Then
		Buf(i / 3 * 4) = (arrB(i) And &HFC) / &H4
		Buf(i / 3 * 4 + 1) = (arrB(i) And &H3) * &H10 + (arrB(i + 1) And &HF0) / &H10
		Buf(i / 3 * 4 + 2) = (arrB(i + 1) And &HF) * &H4
		Buf(i / 3 * 4 + 3) = 64
	End If
	'开辟内存空间
	Base64Encode = Space$(UBound(Buf) + 1)
	'应用Base64模板，转换为Base64码
	For i = 0 To UBound(Buf)
		Mid$(Base64Encode,i + 1,1) = Mid$(cstBase64,Buf(i) + 1,1)
	Next i
End Function


'Base64 解码
Private Function Base64Decode(strEncoded As String,ByVal CodePage As Long) As String
    Const cstBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    Dim i As Long,j As Long,Length As Long,Mods As Long,Buf(3) As Byte
    '获取Base64真实长度,除去补位的字符
    Length = InStr(strEncoded & "=", "=") - 1
    Mods = Length Mod 4
	Length = Length - Mods
    ReDim bRet(Length / 4 * 3 - 1 + IIf(Mods = 0,0,Mods - 1)) As Byte
    For i = 1 To Length Step 4
    	'根据字符的位置取得索引值
		For j = 0 To 3
			Buf(j) = InStr(1,cstBase64,Mid(strEncoded,i + j,1)) - 1
		Next j
		bRet((i - 1) / 4 * 3) = Buf(0) * &H4 + (Buf(1) And &H30) / &H10
		bRet((i - 1) / 4 * 3 + 1) = (Buf(1) And &HF) * &H10 + (Buf(2) And &H3C) / &H4
		bRet((i - 1) / 4 * 3 + 2) = (Buf(2) And &H3) * &H40 + Buf(3)
    Next i
    '处理余数
    i = Length
	If Mods = 2 Then
		For j = 1 To 2
			Buf(j) = InStr(1,cstBase64,Mid(strEncoded,i + j,1)) - 1
		Next j
		bRet(i / 4 * 3) = Buf(1) * &H4 + (Buf(2) And &H30) / &H10
	ElseIf Mods = 3 Then
		For j = 1 To 3
			Buf(j) = InStr(1,cstBase64,Mid(strEncoded,i + j,1)) - 1
		Next j
		bRet(i / 4 * 3) = Buf(1) * &H4 + (Buf(2) And &H30) / &H10
		bRet(i / 4 * 3 + 1) = (Buf(2) And &HF) * &H10 + (Buf(3) And &H3C) / &H4
	End If
	'转换为字串
    Base64Decode = ByteToString(bRet,CodePage)
End Function


'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
Private Function SetTextBoxLength(ByVal hwnd As Long,ByVal NewLength As Long,ByVal Mode As Boolean) As Long
	SetTextBoxLength = SendMessageLNG(hwnd,EM_GETLIMITTEXT,0&,0&)
	If NewLength < SetTextBoxLength Then
		If Mode = False Then Exit Function
		If NewLength < 30001 Then Exit Function
		If NewLength > (SetTextBoxLength \ 3) Then Exit Function
	End If
	SetTextBoxLength = NewLength * 2
	SendMessageLNG hwnd,EM_LIMITTEXT,SetTextBoxLength,0&
End Function


'获取字符串的截取位置，避免截取半个双字节字符
'返回 IsDBCSLeadPos = 截取位置(截取后的字节长度)
'Mode = False 不返回截取后的字串, 否则返回截取后的字串
'fType = False 后截取，否则前截取
'ByteLength 为需要的字节长度
Private Function IsDBCSLeadPos(textStr As String,ByVal CodePage As Long,ByVal ByteLength As Long,Optional ByVal Mode As Boolean,Optional ByVal fType As Boolean) As Long
	Dim i As Long,k As Long,l As Long
	IsDBCSLeadPos = StrHexLength(textStr,CodePage,0)
	If IsDBCSLeadPos <= ByteLength Then Exit Function
	l = IIf(CodePage = CP_UTF8 Or CodePage = CP_UTF7,2,1)
	i = Len(textStr) * ByteLength / IsDBCSLeadPos
	If fType = False Then
		Do While i > 0
			k = StrHexLength(Left$(textStr,i),CodePage,0)
			If k < ByteLength - l Then
				i = i + 1
			ElseIf k <= ByteLength Then
				IsDBCSLeadPos = k
				If Mode = True Then textStr = Left$(textStr,i)
				Exit Do
			Else
				i = i - 1
			End If
		Loop
	Else
		Do While i > 0
			k = StrHexLength(Right$(textStr,i),CodePage,0)
			If k < ByteLength - l Then
				i = i + 1
			ElseIf k <= ByteLength Then
				IsDBCSLeadPos = k
				If Mode = True Then textStr = Right$(textStr,i)
				Exit Do
			Else
				i = i - 1
			End If
		Loop
	End If
End Function


'获取字符串的空格补齐字节数
'返回 FillStrWithSpape = 补齐所需的空格字节数(fType = False 为正值，否则为负值)
'Mode = False 不返回补齐字符数后的字串, 否则返回补齐字符数后的字串
'fType = False 后端空格补齐, 否则前端空格补齐
'ByteLength 为需要的字节长度
Private Function FillStrWithSpape(textStr As String,ByVal CodePage As Long,ByVal ByteLength As Long,Optional ByVal Mode As Boolean,Optional ByVal fType As Boolean) As Long
	Dim i As Long,k As Long,l As Long
	l = IIf(CodePage = CP_UTF8 Or CodePage = CP_UTF7,2,1)
	i = ByteLength - StrHexLength(textStr,CodePage,0)
	If i < 0 Then
		IsDBCSLeadPos(textStr,CodePage,ByteLength,Mode,fType)
		Exit Function
	End If
	Do While i > 0
		k = StrHexLength(textStr & Space$(i),CodePage,0)
		If k < ByteLength - l Then
			i = i + 1
		ElseIf k <= ByteLength Then
			If fType = False Then
				FillStrWithSpape = StrHexLength(Space$(i),CodePage,0)
				If Mode = True Then textStr = textStr & Space$(i)
			Else
				FillStrWithSpape = -StrHexLength(Space$(i),CodePage,0)
				If Mode = True Then textStr = Space$(i) & textStr
			End If
			Exit Do
		Else
			i = i - 1
		End If
	Loop
End Function


'输出程序错误消息
Private Sub sysErrorMassage(sysError As ErrObject,ByVal fType As Long)
	Dim TempArray() As String
	Dim ErrorNumber As Long,ErrorSource As String,ErrorDescription As String
	Dim TitleMsg As String,ContinueMsg As String,Msg As String

	ErrorNumber = sysError.Number
	ErrorSource = sysError.Source
	ErrorDescription = sysError.Description

	TitleMsg = "Error"
	Select Case fType
	Case 0
		ContinueMsg = vbCrLf & vbCrLf & "The program cannot continue and will exit."
	Case 1
		ContinueMsg = vbCrLf & vbCrLf & "Do you want to continue?"
	Case 2
		ContinueMsg = vbCrLf & vbCrLf & "The program will continue to run."
	End Select

	If CheckArray(MsgList) = True Then
		TitleMsg = MsgList(0)
		Select Case fType
		Case 0
			ContinueMsg = MsgList(1)
		Case 1
			ContinueMsg = MsgList(2)
		Case 2
			ContinueMsg = MsgList(3)
		End Select

		Select Case ErrorSource
		Case ""
			If ErrorNumber = 10051 And PSL.Version >= 1500 Then
				Msg = Replace$(MsgList(4),"%s",ErrorSource)
			Else
				Msg = Replace$(Replace$(MsgList(5),"%d",CStr(ErrorNumber)),"%v",ErrorDescription)
			End If
		Case "NotSection"
			TempArray = ReSplit(ErrorDescription,TextJoinStr,-1)
			Msg = Replace$(Replace$(MsgList(6),"%s",TempArray(1)),"%d",TempArray(0))
		Case "NotValue"
			TempArray = ReSplit(ErrorDescription,TextJoinStr,-1)
			Msg = Replace$(Replace$(MsgList(7),"%s",TempArray(1)),"%d",TempArray(0))
		Case "NotReadFile"
			Msg = Replace$(ErrorDescription,TextJoinStr,vbCrLf)
		Case "NotWriteFile"
			Msg = Replace$(ErrorDescription,TextJoinStr,vbCrLf)
		Case "NotUnWriteFile"
			Msg = Replace$(ErrorDescription,TextJoinStr,vbCrLf)
		Case "NotOpenFile"
			Msg = Replace$(ErrorDescription,TextJoinStr,vbCrLf)
		Case "NotINIFile"
			Msg = Replace$(MsgList(8),"%s",ErrorDescription)
		Case "NotExitFile"
			Msg = Replace$(MsgList(9),"%s",ErrorDescription)
		Case "NotVersion"
			TempArray = ReSplit(ErrorDescription,TextJoinStr,-1)
			Msg = Replace$(MsgList(10),"%s",TempArray(0))
			Msg = Replace$(Replace$(Msg,"%d",TempArray(1)),"%v",TempArray(2))
		Case Else
			Msg = Replace$(MsgList(11),"%s",ErrorSource)
			Msg = Replace$(Replace$(Msg,"%d",CStr(ErrorNumber)),"%v",ErrorDescription)
		End Select
	Else
		Select Case ErrorSource
		Case ""
			If ErrorNumber = 10051 And PSL.Version >= 1500 Then
				Msg = "Unable to open the file. Please verify the file path and file name" & _
						"contains characters in Asian languages. " & vbCrLf & _
						"Note: Passolo 2015 Version of the macro engine does not recognize" & _
						"the file path and file name contains Asian language characters."
			Else
				Msg = "An Error occurred in the program design." & vbCrLf & "Error Code: %d, Content: %v" & _
						vbCrLf & "Please restart the Passolo try and please report to the software developer."
				Msg = Replace$(Replace$(Msg,"%d",CStr(ErrorNumber)),"%v",ErrorDescription)
			End If
		Case "NotSection"
			TempArray = ReSplit(ErrorDescription,TextJoinStr,-1)
			Msg = "The following file is missing [%s] section." & vbCrLf & "%d"
			Msg = Replace$(Replace$(Msg,"%s",TempArray(1)),"%d",TempArray(0))
		Case "NotValue"
			TempArray = ReSplit(ErrorDescription,TextJoinStr,-1)
			Msg = "The following file is missing [%s] Value." & vbCrLf & "%d"
			Msg = Replace$(Replace$(Msg,"%s",TempArray(1)),"%d",TempArray(0))
		Case "NotReadFile"
			Msg = Replace$(ErrorDescription,TextJoinStr,vbCrLf)
		Case "NotWriteFile"
			Msg = Replace$(ErrorDescription,TextJoinStr,vbCrLf)
		Case "NotUnWriteFile"
			Msg = Replace$(ErrorDescription,TextJoinStr,vbCrLf)
		Case "NotOpenFile"
			Msg = Replace$(ErrorDescription,TextJoinStr,vbCrLf)
		Case "NotINIFile"
			Msg = "The following contents of the file is not correct." & vbCrLf & "%s"
			Msg = Replace$(Msg,"%s",ErrorDescription)
		Case "NotExitFile"
			Msg = "The following file does not exist! Please check and try again." & vbCrLf & "%s"
			Msg = Replace$(Msg,"%s",ErrorDescription)
		Case "NotVersion"
			TempArray = ReSplit(ErrorDescription,TextJoinStr,-1)
			Msg = "The following file version is %d, requires version at least %v." & vbCrLf & "%s"
			Msg = Replace$(Msg,"%s",TempArray(0))
			Msg = Replace$(Replace$(Msg,"%d",TempArray(1)),"%v",TempArray(2))
		Case Else
			Msg = "Your system is missing %s server." & vbCrLf & "Error Code: %d, Content: %v"
			Msg = Replace$(Msg,"%s",ErrorSource)
			Msg = Replace$(Replace$(Msg,"%d",CStr(ErrorNumber)),"%v",ErrorDescription)
		End Select
	End If

	If Msg = "" Then Exit Sub
	Select Case fType
	Case 0
		MsgBox(Msg & ContinueMsg,vbOkOnly+vbInformation,TitleMsg)
		Exit All
	Case 1
		If MsgBox(Msg & ContinueMsg,vbYesNo+vbInformation,TitleMsg) = vbNo Then
			Exit All
		End If
	Case Else
		MsgBox(Msg & ContinueMsg,vbOkOnly+vbInformation,TitleMsg)
	End Select
End Sub


'修正 PSL 2015 及以上版本宏引擎的 Split 函数拆分空字符串时返回未初始化数组的错误
Public Function ReSplit(ByVal textStr As String,Optional ByVal Sep As String = " ",Optional ByVal Max As Integer = -1) As String()
	If textStr = "" Then
		ReDim TempList(0) As String
		ReSplit = TempList
	Else
		ReSplit = Split(textStr,Sep,Max)
	End If
End Function


'连接字串数组为字串，因为 Join 函数效率太低
'Mode = False 按 Join 函数方式连接，后面不带连接符，否则最后带连接符
Private Function StrListJoin(StrArray() As String,Optional ByVal Sep As String = " ",Optional ByVal Mode As Boolean) As String
	Dim i As Long,j As Long,sb As Object
	On Error GoTo ExitFunction
	Set sb = CreateObject("System.Text.StringBuilder")
	j = UBound(StrArray)
	If Mode = False Then
		For i = 0 To j
			If i < j Then
				sb.AppendFormat("{0}",StrArray(i) & Sep)
			Else
				sb.AppendFormat("{0}",StrArray(i))
			End If
		Next i
	Else
		For i = 0 To j
			sb.AppendFormat("{0}",StrArray(i) & Sep)
		Next i
	End If
	StrListJoin = sb.ToString()
	Set sb = Nothing
	Exit Function
	ExitFunction:
	Set sb = Nothing
	On Error GoTo errHandle
	StrListJoin = Join$(StrArray,Sep)
	errHandle:
End Function


'消息字符串
Private Function GetMsgList(MsgList() As String,ByVal Language As String) As Boolean
	Dim i As Integer
	ReDim MsgList(74) As String
	On Error GoTo errHandle
	Language = LCase$(Language)
	Select Case Language
	Case "chs","0804","1004"
		MsgList(0) = "错误"
		MsgList(1) = "\r\n\r\n程序无法继续运行，将退出。"
		MsgList(2) = "\r\n\r\n要继续运行程序吗？"
		MsgList(3) = "\r\n\r\n程序将继续运行。"
		MsgList(4) = "无法打开文件，请确认文件路径和文件名中是否包含亚洲语言的字符。\r\n" & _
					"注意：Passolo 2015 版本的宏引擎无法识别包含亚洲语言字符的文件路径和文件名。"
		MsgList(5) = "发生程序设计上的错误。\r\n错误代码: %d，错误描述: %v\r\n" & _
					"请重新启动 Passolo 再试，或报告给软件开发者。"
		MsgList(6) = "下列文件中缺少 [%s] 节。\r\n%d"
		MsgList(7) = "下列文件中缺少 [%s] 值。\r\n%d"
		MsgList(8) = "下列文件的内容不正确。\r\n%s"
		MsgList(9) = "下列文件不存在！请检查后再试。\r\n%s"
		MsgList(10) = "下列文件版本为 %d，需要的版本至少为 %v。\r\n%s"
		MsgList(11) = "您的系统缺少 ""%s"" 服务。\r\n错误代码: %d，错误描述: %v"

		MsgList(12) = "版本: %v (构建 %b)\r\n" & _
					"OS 版本: Windows XP/2000 或以上\r\n" & _
					"Passolo 版本: Passolo 5.0 或以上\r\n" & _
					"版权: 汉化新世纪\r\n授权: 免费软件\r\n" & _
					"网址: http://www.hanzify.org\r\n" & _
					"作者: 汉化新世纪成员 - wanfu (2014 - 2018)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(13) = "关于字符编码查询"

		MsgList(14) = "OEM"
		MsgList(15) = "MAC"
		MsgList(16) = "THREAD"
		MsgList(17) = "WEST EUROPE"
		MsgList(18) = "EAST EUROPE"
		MsgList(19) = "RUSSIAN"
		MsgList(20) = "GREEK"
		MsgList(21) = "TURKISH"
		MsgList(22) = "HEBREW"
		MsgList(23) = "ARABIC"
		MsgList(24) = "BALTIC"
		MsgList(25) = "VIETNAMESE"
		MsgList(26) = "JAPAN"
		MsgList(27) = "简体中文 GBK"
		MsgList(28) = "简体中文 GB18030"
		MsgList(29) = "KOREA"
		MsgList(30) = "繁体中文 BIG5"
		MsgList(31) = "THAI"
		MsgList(32) = "UTF-7"
		MsgList(33) = "UTF-8"
		MsgList(34) = "UTF-16LE (UniCode LE)"
		MsgList(35) = "UTF-16BE (Unicode BE)"
		MsgList(36) = "UTF-32LE"
		MsgList(37) = "UTF-32BE"

		MsgList(38) = "十六进制;十六进制转义;URL 转义;HTML 转义;ISO-8859-1;Base64"

		MsgList(39) = "字符编码查询 - 版本 %v (构建 %b)"
		MsgList(40) = "原文"
		MsgList(41) = "翻译"
		MsgList(42) = "文本 (%s/%d):"
		MsgList(43) = "字节间隔模式输出"
		MsgList(44) = "仅多字节字符"
		MsgList(45) = "仅不安全字符"
		MsgList(46) = "转为编码"
		MsgList(47) = "复制"
		MsgList(48) = "粘贴"
		MsgList(49) = "清空"
		MsgList(50) = "编码 (%s/%d):"
		MsgList(51) = "转义 (%s)"
		MsgList(52) = "自动转换"
		MsgList(53) = "转为文本"
		MsgList(54) = "复制"
		MsgList(55) = "粘贴"
		MsgList(56) = "清空"
		MsgList(57) = "关于"
		MsgList(58) = "实际长度"
		MsgList(59) = "剪切/填充零在后面"
		MsgList(60) = "剪切/填充零在前面"
		MsgList(61) = "剪切/填充空格在后面"
		MsgList(62) = "剪切/填充空格在前面"

		MsgList(63) = "十六进制编码位数错误。"
		MsgList(64) = "十进制编码位数错误。"
		MsgList(65) = "包含非法字符。"
		MsgList(66) = "没有要转换的编码。"

		MsgList(67) = "原始编码"
		MsgList(68) = "在编码后附加零"
		MsgList(69) = "在编码前后附加零"

		MsgList(70) = "换行符: \\r\\n;换行符: \\r;换行符: \\n"
		MsgList(71) = "\\r\\n;\\r;\\n"
		MsgList(72) = "语言"
		MsgList(73) = "英语;简体中文;繁体中文"
		MsgList(74) = "enu;chs;cht"
	Case "cht","0404","0c04","1404"
		MsgList(0) = "岿粇"
		MsgList(1) = "\r\n\r\n祘Α礚猭膥尿磅︽盢挡"
		MsgList(2) = "\r\n\r\n璶膥尿磅︽祘Α盾"
		MsgList(3) = "\r\n\r\n祘Α盢膥尿磅︽"
		MsgList(4) = "礚猭秨币郎叫絋粄郎隔畖㎝郎い琌ㄈ瑆粂ēじ\r\n" & _
					"猔種Passolo 2015 セエ栋ま篮礚猭侩醚ㄈ瑆粂ēじ郎隔畖㎝郎"
		MsgList(5) = "祇ネ祘Α砞璸岿粇\r\n岿粇絏: %d岿粇磞瓃: %v\r\n" & _
					"叫穝币笆 Passolo 刚┪厨倒硁砰秨祇"
		MsgList(6) = "郎いぶ [%s] 竊\r\n%d"
		MsgList(7) = "郎いぶ [%s] \r\n%d"
		MsgList(8) = "郎ず甧ぃタ絋\r\n%s"
		MsgList(9) = "郎ぃ叫浪琩刚\r\n%s"
		MsgList(10) = "郎セ %d惠璶セぶ %v\r\n%s"
		MsgList(11) = "眤╰参ぶ%s狝叭\r\n岿粇絏: %d岿粇磞瓃: %v"

		MsgList(12) = "セ: %v (篶 %b)\r\n" & _
					"OS セ: Windows XP/2000 ┪\r\n" & _
					"Passolo セ: Passolo 5.0 ┪\r\n" & _
					"舦: 簙て穝\r\n甭舦: 禣硁砰\r\n" & _
					"呼: http://www.hanzify.org\r\n" & _
					": 簙て穝Θ - wanfu (2014 - 2018)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(13) = "闽じ絪絏琩高"

		MsgList(14) = "OEM"
		MsgList(15) = "MAC"
		MsgList(16) = "THREAD"
		MsgList(17) = "WEST EUROPE"
		MsgList(18) = "EAST EUROPE"
		MsgList(19) = "RUSSIAN"
		MsgList(20) = "GREEK"
		MsgList(21) = "TURKISH"
		MsgList(22) = "HEBREW"
		MsgList(23) = "ARABIC"
		MsgList(24) = "BALTIC"
		MsgList(25) = "VIETNAMESE"
		MsgList(26) = "JAPAN"
		MsgList(27) = "虏砰いゅ GBK"
		MsgList(28) = "虏砰いゅ GB18030"
		MsgList(29) = "KOREA"
		MsgList(30) = "タ砰いゅ BIG5"
		MsgList(31) = "THAI"
		MsgList(32) = "UTF-7"
		MsgList(33) = "UTF-8"
		MsgList(34) = "UTF-16LE (UniCode LE)"
		MsgList(35) = "UTF-16BE (Unicode BE)"
		MsgList(36) = "UTF-32LE"
		MsgList(37) = "UTF-32BE"

		MsgList(38) = "せ秈;せ秈锣竡;URL 锣竡;HTML 锣竡;ISO-8859-1;Base64"

		MsgList(39) = "じ絪絏琩高 - セ %v (篶 %b)"
		MsgList(40) = "ゅ"
		MsgList(41) = "陆亩"
		MsgList(42) = "ゅ (%s/%d):"
		MsgList(43) = "じ舱丁筳家Α块"
		MsgList(44) = "度じ舱じ"
		MsgList(45) = "度ぃじ"
		MsgList(46) = "锣絪絏"
		MsgList(47) = "狡籹"
		MsgList(48) = "禟"
		MsgList(49) = "睲埃"
		MsgList(50) = "絪絏 (%s/%d):"
		MsgList(51) = "锣竡 (%s)"
		MsgList(52) = "笆锣传"
		MsgList(53) = "锣ゅ"
		MsgList(54) = "狡籹"
		MsgList(55) = "禟"
		MsgList(56) = "睲埃"
		MsgList(57) = "闽"
		MsgList(58) = "龟悔"
		MsgList(59) = "芭/恶箂"
		MsgList(60) = "芭/恶箂玡"
		MsgList(61) = "芭/恶"
		MsgList(62) = "芭/恶玡"

		MsgList(63) = "せ秈絪絏计岿粇"
		MsgList(64) = "秈絪絏计岿粇"
		MsgList(65) = "獶猭じ"
		MsgList(66) = "⊿Τ璶锣传絪絏"

		MsgList(67) = "﹍絪絏"
		MsgList(68) = "絪絏箂"
		MsgList(69) = "絪絏玡箂"

		MsgList(70) = "传︽才: \\r\\n;传︽才: \\r;传︽才: \\n"
		MsgList(71) = "\\r\\n;\\r;\\n"
		MsgList(72) = "粂ē"
		MsgList(73) = "璣粂;虏砰いゅ;タ砰いゅ"
		MsgList(74) = "enu;chs;cht"
	Case Else
		MsgList(0) = "Error"
		MsgList(1) = "\r\n\r\nThe program cannot continue and will exit."
		MsgList(2) = "\r\n\r\nDo you want to continue?"
		MsgList(3) = "\r\n\r\nThe program will continue to run."
		MsgList(4) = "Unable to open the file. Please verify the file path and file name" & _
					"contains characters in Asian languages.\r\n" & _
					"Note: Passolo 2015 Version of the macro engine does not recognize" & _
					"the file path and file name contains Asian language characters."
		MsgList(5) = "An Error occurred in the program design.\r\nError Code: %d, Content: %v\r\n" & _
					"Please restart the Passolo try and report this error to the software developer."
		MsgList(6) = "The following file is missing [%s] section.\r\n%d"
		MsgList(7) = "The following file is missing [%s] Value.\r\n%d"
		MsgList(8) = "The following contents of the file is not correct.\r\n%s"
		MsgList(9) = "The following file does not exist! Please check and try again.\r\n%s"
		MsgList(10) = "The following file version is %d, requires version at least %v.\r\n%s"
		MsgList(11) = "Your system is missing %s server.\r\nError Code: %d, Content: %v"

		MsgList(12) = "Version: %v (Build %b)\r\n" & _
					"OS Version: Windows XP/2000 or higher\r\n" & _
					"Passolo Version: Passolo 5.0 or higher\r\n" & _
					"Copyright: Hanzify\r\nLicense: Freeware\r\n" & _
					"HomePage: http://www.hanzify.org\r\n" & _
					"Author: Hanzify member - wanfu (2014 - 2018)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(13) = "About Character Encode Query"

		MsgList(14) = "OEM"
		MsgList(15) = "MAC"
		MsgList(16) = "THREAD"
		MsgList(17) = "WEST EUROPE"
		MsgList(18) = "EAST EUROPE"
		MsgList(19) = "RUSSIAN"
		MsgList(20) = "GREEK"
		MsgList(21) = "TURKISH"
		MsgList(22) = "HEBREW"
		MsgList(23) = "ARABIC"
		MsgList(24) = "BALTIC"
		MsgList(25) = "VIETNAMESE"
		MsgList(26) = "JAPAN"
		MsgList(27) = "CHINA GBK"
		MsgList(28) = "CHINA GB18030"
		MsgList(29) = "KOREA"
		MsgList(30) = "TAIWAN"
		MsgList(31) = "THAI"
		MsgList(32) = "UTF-7"
		MsgList(33) = "UTF-8"
		MsgList(34) = "UTF-16LE (UniCode LE)"
		MsgList(35) = "UTF-16BE (Unicode BE)"
		MsgList(36) = "UTF-32LE"
		MsgList(37) = "UTF-32BE"

		MsgList(38) = "Hex;Hex Escape;URL Escape;HTML Escape;ISO-8859-1;Base64"

		MsgList(39) = "Character Encode Query - Version %v (Build %b)"
		MsgList(40) = "Source"
		MsgList(41) = "Translation"
		MsgList(42) = "Text (%s/%d):"
		MsgList(43) = "Byte Mode Output"
		MsgList(44) = "Multibyte Characters Only"
		MsgList(45) = "Unsafe Characters Only"
		MsgList(46) = "To Code"
		MsgList(47) = "Copy"
		MsgList(48) = "Paste"
		MsgList(49) = "Clear"
		MsgList(50) = "Code (%s/%d):"
		MsgList(51) = "Escape (%s)"
		MsgList(52) = "Auto Convert"
		MsgList(53) = "To Text"
		MsgList(54) = "Copy"
		MsgList(55) = "Paste"
		MsgList(56) = "Clear"
		MsgList(57) = "About"
		MsgList(58) = "Real Length"
		MsgList(59) = "Cut/Fill Zero at End"
		MsgList(60) = "Cut/Fill Zero at Front"
		MsgList(61) = "Cut/Fill Space at End"
		MsgList(62) = "Cut/Fill Space at Front"

		MsgList(63) = "Number of Hex code errors."
		MsgList(64) = "Number of Dec code errors."
		MsgList(65) = "Contains illegal characters."
		MsgList(66) = "There is no code to convert."

		MsgList(67) = "Original codes"
		MsgList(68) = "Append zero at after codes"
		MsgList(69) = "Append zero at before and after codes"

		MsgList(70) = "Line Break: \\r\\n;Line Break: \\r;Line Break: \\n"
		MsgList(71) = "\\r\\n;\\r;\\n"
		MsgList(72) = "Language"
		MsgList(73) = "EngLish;Chinese Simplified;Chinese Traditional"
		MsgList(74) = "enu;chs;cht"
	End Select
	For i = 0 To UBound(MsgList)
		Select Case Language
		Case "chs","0804","1004"
			MsgList(i) = PSL.ConvertASCII2Unicode(MsgList(i),CP_CHINA)
		Case "cht","0404","0c04","1404"
			MsgList(i) = PSL.ConvertASCII2Unicode(MsgList(i),CP_TAIWAN)
		End Select
		If InStr(MsgList(i),"\\") Then MsgList(i) = Replace$(MsgList(i),"\\","*a!N!d*")
		If InStr(MsgList(i),"\r\n") Then MsgList(i) = Replace$(MsgList(i),"\r\n",vbCrLf)
		If InStr(MsgList(i),"\t") Then MsgList(i) = Replace$(MsgList(i),"\t",vbTab)
		If InStr(MsgList(i),"*a!N!d*") Then MsgList(i) = Replace$(MsgList(i),"*a!N!d*","\")
	Next i
	GetMsgList = True
	Exit Function
	errHandle:
	ReDim MsgList(0) As String
End Function
