Attribute VB_Name = "modCommon"
'' Common Module for PSlHardCodedString.bas
'' (c) 2010-2019 by wanfu (Last modified on 2019.06.23

Option Explicit

Public Const Version = "2019.06.23"
Public Const Build = "190623
Public Const ToUpdateDataVersion = "2012.10.25"
Public Const ToUpdateStrTypeVersion = "2015.02.13"
Public Const ToUpdateRuleVersion = "2015.08.18"
Public Const RegKey = "HKCU\Software\VB and VBA Program Settings\PSLHardCodedString\"
Public Const FilterFile = "HcsFilterList.dat"
Public Const ReserveFile = "HcsReserveList.dat"
Public Const JoinStr = vbFormFeed  'vbBack
Public Const TextJoinStr = vbCrLf
Public Const ValJoinStr = ","
Public Const RefJoinStr = "|"
Public Const ItemJoinStr = ";"
Public Const SkipStr = "!""#$%&'()+,-./0123456789:;<=>?@[\]^_`{|}~\x7F\u201A\u201E\u2026" & _
						"\u2020\u2021\u20AC\u2030\u2039\u2018\u2019\u201C\u201D\u2022\u2013" & _
						"\u2014\x98\u203A\x5C\xA0\xA6\xA7\u2122\xA9\xAE\xAB\xAC\xAD\xB0\xB1" & _
						"\u0491\x5C\xB7\u2116\xBB"
Public Const SkipHexStr = "\x01\x02\x03\x04\x05\x06\x07\x0E\x0F\x10\x11\x12\x13\x14\x15\x16" & _
						"\x17\x18\x19\x1A\x1B\x1C\x1D\x1E\x1F\xFF"
Public Const ConvertStrHexRange = "\x30-\x39,\x41-\x46;\x30-\x39,\x41-\x46;\x30-\x37"
			'0=用于十六进制转义符,1=用于 Unicode 转义符,2=用于八进制转义符
Public Const CheckStrHexRange = "\x00-\x07,\x0E-\x1F;" & _
								"\x00-\x40,\x5B-\x60,\x7B-\xBF;" & _
								"\x00-\x60,\x7B-\xBF;" & _
								"\x00-\x40,\x5B-\xBF;" & _
								"\x41-\x5A,\x61-\x7A;" & _
								"\x30-\x39,\x41-\x5A,\x61-\x7A;" & _
								"\x30-\x39,\x41-\x46"
			'0=控制字符,1=全为数字和符号,2=全为大写英文,3=全为小写英文,4=大小写混合英文,5=快捷键字符范围,6=十六进制字符
Public Const NoVowelsRegExp = "^[^aeiouAEIOU]+$"
Public Const FilePathRegExp = "^([A-Za-z]:|\.+)?([/\\]{1,2}([^/\\:*?""<>|]|(?!\\[nrt]))+)+?([/\\]+|\.([^/\\:*?""<>|]|(?!\\[nrt]))+)?$"
Public Const RegKeyRegExp = "^([^\f\n\r\t\v]|(?!\\[nrt]))+(\\{1,2}([^\\\f\n\r\t\v]|(?!\\[nrt]))+?)+?[\\]?$"
Public Const WebSiteRegExp = "^(ht|f)tps?://[-\w]+(\.[-\w]+)+([-\w\.,@?^=%&/~\+#:]*[-\w\@?^=%&/~\+#])?$"
Public Const ExcludeStr = "OK;at;to;in;is;or;on;cm;inches;kg;KG;km;KM;kb;KB"
Public Const EndCharOfString = "0;(\x00+),([^\x00]?),(\x00*),(\r+),(\n+),([\r\n]+),(\t+)"
Public Const AppName = "PSLHardCodedString"
Public Const FreeByteMinLength = 5&
Public Const NewPESecName = ".movehcs"
Public Const NewMacSecName = "__MOVEHCS"

'文件属性
Private Type HCS_FILE
	SourceFileName		As String	'数据文件的来源文件名
	SourceFileVersion	As String	'数据文件的来源文件版本
	SourceFileSize		As Long		'数据文件的来源文件大小
	SourceFileDateTime	As Date		'数据文件的来源文件修改日期
	SourceFileLangID	As String	'数据文件的来源文件语言ID
	AppVersion			As String	'处理数据文件所用的宏版本
	DateLastModified	As Date		'数据文件的修改日期
	FilePath			As String	'数据文件的路径 (含文件名)
	FileFormat			As Boolean	'数据文件的格式，Ture 为符合要求
End Type

'节定义
Private Type SUB_SECTION_PROPERTIE
	sName				As String	'区段名
	RWA					As Long		'数据记录所在物理偏移地址 (实际没有这个值，只是为了方便定位各个流的位置)
	lVirtualSize		As Long		'EXE文件中表示节的实际字节数
	lVirtualAddress		As Long		'本节的低位RVA
	lVirtualAddress1	As Long		'本节的高位RVA (64 位文件)
	lSizeOfRawData		As Long		'本节经文件对齐后的尺寸
	lPointerToRawData	As Long		'本节原始数据在文件中的位置
End Type

'节定义
Private Type SECTION_PROPERTIE
	sName				As String	'区段名
	RWA					As Long		'数据记录所在物理偏移地址 (实际没有这个值，只是为了方便定位各个流的位置)
	lVirtualSize		As Long		'EXE文件中表示节的实际字节数
	lVirtualAddress		As Long		'本节的低位RVA
	lVirtualAddress1	As Long		'本节的高位RVA (64 位文件)
	lSizeOfRawData		As Long		'本节经文件对齐后的尺寸
	lPointerToRawData	As Long		'本节原始数据在文件中的位置
	SubSecs				As Integer	'子节数
	SubSecList()		As SUB_SECTION_PROPERTIE
End Type

'文件属性
Public Type FILE_PROPERTIE
	CompanyName			As String	'开发公司名称
	FileDescription		As String	'文件描述
	FileVersion			As String	'文件版本
	InternalName		As String	'内部名称
	LegalCopyright		As String	'版权
	OrigionalFileName	As String	'原始文件名
	ProductName			As String	'产品名称
	ProductVersion		As String	'产品版本
	LanguageID			As String	'语言ID
	FileSize			As Long		'文件大小
	DateCreated			As Date		'文件创建日期
	DateLastModified	As Date		'文件最后修改日期
	CopePage			As Long		'文件代码页
	FileName			As String	'文件名 (不含路径)
	FilePath			As String	'文件路径 (含文件名)
	FileType			As Integer	'文件类型的开始地址，如：PE 文件开始处的 MZ

	Magic				As String	'文件类型，PE32,NET32,PE64,NET64,MAC32,MAC64，其他为零
	SecAlign			As Long		'内存对齐值
	FileAlign			As Long		'文件对齐值
	MinSecID 			As Integer	'最小偏移开始地址所在节索引号
	MaxSecID 			As Integer	'最大偏移开始地址所在节索引号
	MaxSecIndex 		As Integer	'文件节的最大索引号
	USStreamID 			As Integer	'.NET 字符串所在流索引号
	LangType			As Long		'文件的编写语言
	NumberOfSub			As Integer	'内嵌的子 PE 数
	ImageBase			As Variant	'载入文件的RVA基地址
	DataDirs			As Integer	'数据目录的数量
	NetStreams			As Integer	'.NET 文件流的数量
	SecList() 			As SECTION_PROPERTIE
	DataDirectory()		As SUB_SECTION_PROPERTIE
	CLRList()			As SUB_SECTION_PROPERTIE
	StreamList()		As SUB_SECTION_PROPERTIE
	hcsFile				As HCS_FILE
End Type

'自定义引用
Public Type REF_TYPE
	sName				As String	'算法名称
	Algorithm 			As String	'引用代码的算法公式
	ByteLength			As Integer	'引用的字节长度
	ByteOrder			As Integer	'字节序，0 = 小端，1 = 大端
	FileMagic			As String	'引用类型，NotPE32,NotPE64
	PrefixByte			As String	'前缀字节，按十六进制输入
	PrefixLength		As Integer	'前缀字节长度，按输入值自动计算
	StrAddAlgorithm 	As String	'从引用代码计算字串地址的算法公式
	Template			As String	'保存到字串数据文件的内容
End Type

'PE文件结构(Visual Basic版)部分代码一
'签名定义
Private Enum ImageSignatureTypes
	IMAGE_DOS_SIGNATURE = &H5A4D		'// MZ
	IMAGE_OS2_SIGNATURE = &H454E		'// NE
	IMAGE_OS2_SIGNATURE_LE = &H454C		'// LE
	IMAGE_VXD_SIGNATURE = &H454C		'// LE
	IMAGE_NT_SIGNATURE = &H4550			'// PE00
End Enum

Private Enum mac_header_magic
	MH_MAGIC_32			= &HFEEDFACE	' 32-bit mach Object file
	MH_MAGIC_64			= &HFEEDFACF	' 64-bit mach Object file
	MH_MAGIC_FAT		= &HCAFEBABE	' Universal Object file / FAT_MAGIC
	MH_MAGIC_FAT_CIGAM	= &HBEBAFECA
End Enum

'程序编写语言或平台定义
Public Enum PELangType
	DELPHI_FILE_SIGNATURE = &H50
	NET_FILE_SIGNATURE = &H424A5342
End Enum

Private Type IMAGE_DATA_DIRECTORY
	lVirtualAddress		As Long		'起始RVA地址
	lSize				As Long		'lVirtualAddress所指向数据结构的字节数
End Type

'空字节类型
Public Type FREE_BTYE_SPACE
	Address				As Long		'可用开始地址
	Length				As Long		'可用最大长度
	MaxAddress			As Long		'字符块结束地址
	inSectionID			As Integer	'字串所在节的索引号
	inSubSecID			As Integer	'字串所在子节的索引号
	lNumber				As Long		'可用地址索引号，<-1 = 非字串空间索引，>-1 = 字串空间ID号
	MoveType			As Long		'可用状态, -3 = 不可占用，-2 = 节尾空间，-1 = 未被占用，>-1 = 已被占用(字串ID号)
	IndexList()			As Long		'相同ID的索引列表
End Type

'进度消息类型
Public Type PROGRESS_MSG
	Total				As Long		'总处理数
	Passed				As Long		'已处理数
	Massage				As String	'基本消息
	hwnd				As Long		'显示消息的对话框控件句柄
End Type

'编辑工具类型
Public Type TOOLS_PROPERTIE
	sName				As String	'工具名称
	FilePath			As String	'工具文件路径(含文件名)
	Argument			As String	'运行参数
End Type

'代码页类型
Public Type CODEPAGE_DATA
	sName				As String	'代码名称
	CharSet				As String	'字符编码
End Type

'过滤器类型
Public Type FILTER_PROPERTIE
	Item				As String	'过滤项目
	Value				As String	'判断值
	Mode				As Integer	'判断类型
	Other				As String	'其他值
End Type

'语言文件
Public Type UI_FILE
	FilePath			As String	'语言文件完全路径
	AppName				As String	'程序名称
	Version				As String	'程序版本
	LangName			As String	'语言名称
	LangID				As String	'适用语言ID
	Encoding			As String	'字符编码
End Type

'INI 文件
Public Type INIFILE_DATA
	Title				As String	'主题
	Item()				As String	'项目
	Value()				As String	'字串值
End Type

'引用
Public Type REFERENCE_PROPERTIE
	sCode				As String	'引用代码
	lAddress			As Long		'引用地址列表
	inSecID				As Integer	'字串所在节的索引号
	StrType				As Integer	'跟随引用代码的字串类型
	Index               As Long		'引用的索引号，以 1 为开始，用于在子字串的翻译中附加索引号，方便引用拆分
End Type

'字串子类型
Public Type STRING_SUB_PROPERTIE
	lStartAddress		As Long		'字串的开始地址
	lEndAddress			As Long		'字串的结束地址
	lMaxAddress			As Long		'字符块结束地址
	lCharLength			As Long		'字串的字符长度
	lHexLength			As Long		'字串的十六进制长度
	lMaxHexLength		As Long		'字串的最大允许十六进制长度
	CodePage			As Long		'字串的代码页
	sString				As String	'字符串
	inSectionID			As Integer	'字串所在节的索引号
	inSubSecID			As Integer	'字串所在节的子节索引号
	StrTypeLength		As Integer	'字串类型的前置标识符字节数
	MoveLength			As Integer	'可变型字串长度标识符的字串类型长度更改值, 0=无更改, >0=增加值, <0=减少值
									'原始字串的该值 = 翻译字串 - 原始字串，翻译字串的该值 = 截断后 - 原始翻译
	iNullByteLength		As Integer	'字串后要预留的空字节长度(Unicode = 2，其他代码页 = 1，Pascal Short = 0)
	lReferenceNum		As Long		'引用次数
	GetRefState			As Integer	'获取字串引用列表的状态，0 = 未获取，1 = 已获取，0=导入
	NewStemp			As Integer	'字串的新增标记，用于附加导入时可以过滤显示新增的字串，原始字串：0=提取，1=导入，2=添加，翻译字串：1=导入
	Reference()			As REFERENCE_PROPERTIE
End Type

'字串类型
Public Type STRING_PROPERTIE
	ID						As Long		'字串编号(ID号), <0 引用地址(子字串)，>-1 字串开始地址(父字串)
	StrType					As Integer	'字串类型, <-4 Android，<0=.NET，0=默认，1=Pascal Unicode, 2=Pascal Wide, 3=Pascal Ansi, 4=Pascal Short, >4=自定义
	WriteType				As Integer	'写入类型
										'PE 文件: -2=丢失(已翻译)，-1=丢失(未翻译)，0=未翻译，1=原址完整写入，2=原址截断写入，3=全部移位写入，4=部分移位写入，5=原址超长写入
										'非 PE 文件: -2=丢失(已翻译)，-1=丢失(未翻译)，0=未翻译，1=原长完整写入，2=最长完整写入，3=超长完整写入，4=原址截断写入
	WriteState				As Integer	'写入状态, 0=未写入，1=完整写入，2=截断写入，3=字串写入失败，4=引用代码修改失败，5=标识符修改失败
	LockState				As Integer	'翻译字串的锁定状态，0=未锁定，1=已锁定
	ScapeIDBeMoved			As Long		'被占用的其他移位字串ID号: -1=没有被占用，>-1=占用该字串地址的其他移位字串的字串ID号
	ScapeIDForMove			As Long		'移位写入所使用的地址ID号: -1=没有移动，>-1=其他移位字串的字串索引号，<-1=字串列表以外的空余地址ID号
	SourceStringClearState	As Integer	'字串原址的清空状态，0=未清空，1=已清空
	EndByteLength			As Integer	'字串类型的后随标识符字节数
	GetLengthState			As Integer	'字串类型长度的获取状态，0 = 未获取，1 = 已获取
	MoveType				As Integer	'移动类型, 0=不移位, 1=字串空位(字串ID), 2=非字串空位(索引号), 3=节后原有空位(-2)
												  '4=字串所在节后扩展空位(-3), 5=最大节扩展空位(-开始地址), 6=新增节(-开始地址)
	MoveMode				As Integer	'移位模式：0=超长移位(未移位时引用地址为空), 1=强制移位(未移位时引用地址非空),
												  '2=原址移位(字串类型字节长度改变), 3=手动移位，4=不移位不扩展，5=原址扩展
	SplitState				As Long		'拆分状态, 0=未拆分, >0=已拆分的父串(数字表示子串数), <0=子串(数字表示其父串ID)
	TagType					As Integer	'标识符类型, 0=无标识符, 1=跟随字串, 2=跟随引用代码
	LengthModeID			As Integer	'检测到的跟随字串的自定义字串类型的字符串长度标识符的字串长度计算依据
	FillLength				As Integer	'短于原始长度的翻译字串补齐空格的空格字节长度, <0=前端空格增加值, 0=无补齐, >0=后端空格增加值
	iError					As Integer	'更新字串列表时的错误处理方法，PSL 文本解析器不支持创建在 16k 以上的字串，0=常规, 1=忽略, 2=删除
	Moveable				As Integer  '可移位属性(和有无引用无关), 0=可移位, 1=非 PE 字串, 2=.NET 可移位字串, 3=隐藏节字串, 4=丢失字串
	OverLengthWrite			As Integer  '超长写入属性 (可否利用原字串后的空位), 0=允许, 1=全局设置为禁止, 2=选定字串设置为禁止
	Missing					As Integer  '原始或目标文件中丢失的字串，0=未丢失，1=丢失
	Source					As STRING_SUB_PROPERTIE
	Trans					As STRING_SUB_PROPERTIE
End Type

'语言类型
Public Type LANG_PROPERTIE
	LangName					As String	'语言名称
	LangID						As Long		'语言ID
	CPName						As String	'代码页名称
	CodePage					As Long		'代码页代码值
	UniCodeRange				As String	'Unidoce的编码范围
	UniCodeRegExpPattern		As String	'Unicode的编码范围(RegExp模板)
	UniCodeByteRange			As String	'Unicode的字符范围
	FeatureCode					As String	'特征码范围
	FeatureCodeRegExpPattern	As String	'特征码范围(RegExp模板)
	FeatureCodeByteRange		As String	'特征码的字符范围
	FeatureCodeEnable			As Integer  '特征码启用项,True 为启用,默认不启用
	dwFlags						As Boolean	'语言标志, True = 用户添加的语言
End Type

'字节
'Private Type BYTE_PROPERTIE
'	NullValNum		As Long		'&H00字节数
'	CBLValNum		As Long		'&H00-&H7F字节数
'	NullValPos		As Long		'&H00字节所在位置
'	ByteType		As Long		'字节块类型
'End Type

Public Type CHECK_STRING_VALUE
	AscRange		As String
	Range			As String
End Type

'内存映射方式查找字串
'Private Type SAFEARRAYBOUND
'	cElements		As Long		'一维有多少个元素
'	lLbound			As Long		'索引开始值
'End Type

'内存映射方式查找字串
'Private Type SAFEARRAYID
'	cDims			As Integer 	'数组的维数
'	fFeatures		As Integer	'数组的特性
'	cbElements		As Long		'数组的每个元素大小
'	clocks			As Long		'数组被锁定的次数
'	pvData			As Long		'数组里的数据存放位置
'	rgsabound(0)	As SAFEARRAYBOUND
'End Type

'打开文件方式的结构体
Public Type FILE_IMAGE
	ModuleName		As String	'被加载文件的文件名
	hFile			As Long		'调用 Create 文件映射或 OpenFile 的句柄
	hMap			As Long		'调用 CreateFileMapping 文件映射的句柄
	MappedAddress	As Long		'文件映射到的内存地址
	SizeOfImage		As Long		'映射的 Image 或字节数组的大小
	SizeOfFile		As Long		'实际文件大小
	ImageByte()		As Byte		'文件的字节数组
End Type

'Delphi 字符串类型定义
'-----------------------------------------------------------------
'1 ShortString		可以容纳255个字符,主要为了老版本兼容
'2 AnsiString		可以容纳2的31次方个字符,D2009前默认的String类型
'3 UnicodeString	可以容纳2的30次方个字符,D2009及以后的默认String类型
'4 WideString		可以容纳2的30次方个字符,主要在COM中用的比较多
'-----------------------------------------------------------------
'ShortString
Public Type DELPHI_SHORT_STRING
	Length					As Byte		'字符串字节长度，1个字节
	'Strings()				As Byte		'字符串，字节长度为 Length * 1
	'EndChar(0)				As Byte		'不一定以&H00结束
End Type

'AnsiString
Public Type DELPHI_ANSI_STRING
	RefCount				As Long		'字符串引用次数，4个字节
	Length					As Long		'字符串字节长度，4个字节
	'Strings()				As Byte		'字符串，字节长度为 Length * 1
	'EndChar(0)				As Byte		'以&H00结束
End Type

'WideString
Public Type DELPHI_WIDE_STRING
	Length					As Long		'字符串字节长度，4个字节
	'Strings()				As Byte		'字符串，字节长度为 Length * 1
	'EndChar(1)				As Byte		'以&H0000结束
End Type

'UnicodeString
Public Type DELPHI_UNICODE_STRING
	CodePage				As Integer	'字符串代码页，2个字节, 支持 Unicode, UTF-8, ANSI
	elemSize				As Integer	'每个字符的字节数，2个字节，Unicode = 2, UTF-8 = 1 or 3, GB2312 = 2
	RefCount				As Long		'字符串引用次数，4个字节
	Length					As Long		'字符串的字符数，4个字节
	'Strings()				As Byte		'字符串，字节长度为 Length * elemSize
	'EndChar(1)				As Byte		'以&H0000结束
End Type

'自定义字符串类型定义
Public Type STRING_TYPE
	sName					As String	'字符串类型的名称
	CodeLoc					As Integer	'所有字符串标识符所在位置, 0 = 字串前, 1 = 引用地址前
	FristCodePos			As Integer	'第一个字符串标识符位于字串前的位置
	CPCodePos				As Integer	'字符串代码页标识符位于字串前的位置
	CPCodeSize				As Integer	'字符串代码页标识符的字节数
	CPCodeStartString		As String	'字符串代码页标识符开始标记(Hex文本)
	CPCodeStartLength		As Integer	'字符串代码页标识符开始标记(Hex文本)的字节长度
	CPCodeStartByte()		As Byte		'字符串代码页标识符开始标记(Hex文本)的字节数组
	LengthCodePos			As Integer	'字符串长度标识符位于字串前的位置
	LengthCodeSize			As Integer	'字符串长度标识符的字节数
	LengthMode				As Integer	'字符串长度标识符的字串长度计算依据
	LengthReviseVal			As Integer	'字符串长度标识符的字串长度调整值
	ByteLengthReviseVal		As Integer	'字符串长度标识符的字串长度调整值
	CharLengthReviseVal		As Integer	'字符串长度标识符的字串长度调整值
	LengthCodeStartString	As String	'字符串长度标识符开始标记(Hex文本)
	LengthCodeStartLength	As Integer	'字符串长度标识符开始标记(Hex文本)的字节数
	LengthCodeStartByte()	As Byte		'字符串长度标识符开始标记(Hex文本)的字节数组
	StartCodePos			As Integer	'字符串开始标识符位于字串前的位置
	StartCodeString			As String	'字符串开始标识符(Hex文本)
	StartCodeLength			As Integer	'字符串开始标识符的字节数
	StartCodeByte()			As Byte		'字符串开始标识符的字节数组
	EndCodeString			As String	'字符串结束标识符(Hex文本)
	EndCodeLength			As Integer	'字符串结束标识符的字节数
	EndCodeByte()			As Byte		'字符串结束标识符的字节数组
	RefCodeStartPos			As Integer	'引用代码开始位于引用代码前的位置
	RefCodeStartString		As String	'引用代码开始标记(Hex文本)
	RefCodeStartLength		As Integer	'引用代码开始标记(Hex文本)的字节数
	RefCodeStartByte()		As Byte		'引用代码开始标记(Hex文本)的字节数组
	RegExpPattern()			As String	'正则表达式模板
End Type

'字串长度判断值定义
Public Type STRING_TYPE_LENGTH
	Pattern					As String	'正则表达式模板
	Length1					As Long		'首个长度判断值
	Length2					As Long		'次个长度判断值
	Size					As Long		'长度大小判断值
	Bytes()					As Byte		'长度的字节数组
	ByteOrder				As Integer	'字节序, -1 = 大端在前, 0 = 小端在前, 1 = 未知
End Type

'文件名称及所在文件夹定义
Public Type FILE_LIST
	sName					As String	'文本解析器名称
	FilePath				As String	'解析器数据文件路径(含文件名)
End Type

Public Enum KnownCodePage
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
	CP_UTF32LE = 65005  	'Unicode (UTF-32 LE)
	CP_UTF32BE = 65006		'Unicode (UTF-32 Big-Endian)
End Enum

'SetLocaleInfo()中 LCTYPE values 的具体意义
Public Enum LOCALE
	LOCALE_ILANGUAGE = &H1			'语言ID
	LOCALE_SLANGUAGE = &H2			'语言区域名称，如: "English (United States)"
	LOCALE_SENGLANGUAGE = &H1001	'语言英语名称
	LOCALE_SABBREVLANGNAME = &H3	'语言名称缩写，如: "ENU"
	LOCALE_SNATIVELANGNAME = &H4	'当地语言名称，如: "English"
	LOCALE_ICOUNTRY = &H5			'国家代码
	LOCALE_SCOUNTRY = &H6			'国家本地名称
	LOCALE_SENGCOUNTRY = 4098		'国家英语名称
	LOCALE_SABBREVCTRYNAME = &H7	'国家名称缩写
	LOCALE_SNATIVECTRYNAME = &H8	'当地语言国家名称
	LOCALE_IDEFAULTLANGUAGE = &H9	'缺省语言ID
	LOCALE_IDEFAULTCOUNTRY = &HA	'缺省国家代码
	LOCALE_IDEFAULTCODEPAGE = &HB	'缺省的OEM代码
	LOCALE_IDEFAULTANSICODEPAGE = &H1004	'缺省的ASCII代码
	LOCALE_IDEFAULTMACCODEPAGE = &H1011		'缺省的MACINTOH代码
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
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" ( _
	ByVal LOCALE As Long, _
	ByVal LCType As Long, _
	ByVal lpLCData As String, _
	ByVal cchData As Long) As Long
	'cchData 为 lpLCData 缓冲区的长度；如设为零，表示获取必要的缓冲区长度

'内存复制和比较函数
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
	ByRef Destination As Any, _
	ByRef Source As Any, _
	ByVal Length As Long)
Public Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
	ByRef Dest As Any, _
	ByVal Source As Long, _
	ByVal Length As Long)
Private Declare Sub ReadMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
	ByVal Dest As Long, _
	ByVal Source As Long, _
	ByVal Length As Long)
Public Declare Sub WriteMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
	ByVal Dest As Long, _
	ByRef Source As Any, _
	ByVal Length As Long)
Private Declare Function CompMemory Lib "ntdll.dll" Alias "RtlCompareMemory" ( _
	ByRef Dest As Any, _
	ByRef Source As Any, _
	ByVal Length As Long) As Long
'Private Declare Function vbVarPtr Lib "msvbvm60.dll" Alias "VarPtr" (Ptr As Any) As Long

'光标坐标定义
Public Type POINTAPI
	x As Long
	y As Long
End Type

'SendMessage API 部分常数
Public Enum SendMsgValue
	'用于组合列表框常数
	CB_GETEDITSEL = &H140			'起点,终点		用于取得组合框所包含编辑框子控件中当前被选中的字符串的起止位置
	CB_SETEDITSEL = &h142			'0,0 or -1		用于选中组合框所包含编辑框子控件中的部分字符串,对应函数

	'用于文本框查找定位常数
	EM_GETSEL = &HB0				'0,变量			获取光标位置（以本机默认编码的字符数表示）
	EM_SETSEL = &HB1				'起点,终点		设置编辑控件中文本选定内容范围（或设置光标位置）起点和终点均为字符值
									'				当指定的起点等于0和终点等于-1时，文本全部被选中，此法常用在清空编辑控件
									'				当指定的起点等于-2和终点等于-1时，全文均不选，光标移至文本未端
	EM_GETLINECOUNT = &HBA			'0,0			获取编辑控件的总行数
	EM_LINEINDEX = &HBB				'行号,0			获取指定行(或:-1,0 表示光标所在行)首字符在文本中的位置（以字符数表示）
	EM_LINELENGTH = &HC1			'偏移值,0		获取指定位置所在行(或:-1,0 表示光标所在行）的文本长度（以字符数表示）
	EM_LINEFROMCHAR = &HC9			'偏移值,0		获取指定位置(或:-1,0 表示光标位置)所在的行号
	EM_GETLINE = &HC4				'行号,ByVal变量	获取编辑控件某一行的内容，变量须预先赋空格
	EM_SCROLLCARET = &HB7			'0,0 			把可见范围移至光标处
	EM_UNDO = &HC7					'0,0 			撤消前一次编辑操作，当重复发送本消息，控件将在撤消和恢复中来回切换
	EM_REPLACESEL = &HC2			'1(0),字符串	用指定字符串替换编辑控件中的当前选定内容
									'				如果第三个参数wParam为1，则本次操作允许撤消，0禁止撤消。字符串可用传值方式，也可用传址方式
									'				（例：SendMessage Text1.hwnd, EM_REPLACESEL, 0, Text2.Text '这是传值方式）
	EM_GETMODIFY = &HB8				'0,0			判断编辑控件的内容是否已发生变化，返回TRUE(1)则控件文本已被修改，返回FALSE(0)则未变
	EN_CHANGE = &H300 				'				编辑控件的内容发生改变。与EN_UPDATE不同，该消息是在编辑框显示的正文被刷新后才发出的
	EN_UPDATE = &H400				'				控件准备显示改变了的正文时发送该消息。它与EN_CHANGE通知消息相似，只是它发生于更新文本显示出来之前
	EM_GETLIMITTEXT = &HD5			'0,0			获取一个编辑控件中文本的最大长度
	EM_LIMITTEXT = &HC5				'最大值,0		设置编辑控件中的最大文本长度
	EM_GETFIRSTVISIBLEINE = &HCE	'0,0			获得文本控件中处于可见位置的最顶部的文本所在的行号
	EM_GETHANDLE = &HBD				'0,0			取得文本缓冲区
	'EM_SETCHARFORMAT = &H444		'颜色值,0		改变选定文本的颜色

	'用于列表框常数
	'LB_ADDFILE = &H0196			'0,文件名地址	增加文件名
	'LB_ADDSTRING = &H0180			'0,字符串地址	追加一个列表项返回索引。如果指定了LBS_SORT风格，表项将被重排序，否则将被追加在列表框的最后一项
	LB_DELETESTRING = &H0182		'列表项序号,0	删除指定的列表项返回列表框剩餘表項
	'LB_DIR = &H018D				'DDL_ARCHIVE,指向通配符地址	添加文件名列表，返回最後一個添加的文件名的索引
	'LB_FINDSTRING = &H018F			'开始表项序号,字符串地址	查找匹配字符串，忽略大小写，从指定开始表项序号开始查找，当查到某表项的文本字符串的前面包括指定的字符串则结束，找不到则转到列表框第一项继续查找，直到查完所有表项，如果wParam为-1则从列表框第一项开始查找，如果找到则返回表项序号，否则返回LB_ERR。如：表项字符串为"abc123"和指定字串"ABC"就算匹配
	'LB_FINDSTRINGEXACT = &H01A2	'开始表项序号,字符串地址	查找字符串，忽略大小写，与LB_FINDSTRING不同，本操作必须整个字符串相同。如果找到则返回表项序号，否则返回LB_ERR
	'LB_GETANCHORINDEX = &H019D		'0,0			返回鼠标最后选中的项的索引
	'LB_GETCARETINDEX = &H019F		'0,0			返回具有矩形焦点的项的索引
	LB_GETCOUNT = &H018B			'0,0			返回列表项的总项数，若出错则返回LB_ERR
	'LB_GETCURSEL = &H0188			'0,0			本操作仅适用于单选择列表框，用来返回当前被选择项的索引，如果没有列表项被选择或有错误发生，则返回LB_ERR
	LB_GETHORIZONTALEXTENT = &H0193	'0,0		返回列表框的可滚动的宽度（象素）
	'LB_GETITEMDATA = &H0199		'索引,0			每个列表项都有一个32位的附加数据．该函数返回指定列表项的附加数据。若出错则函数返回LB_ERR
	'LB_GETITEMHEIGHT = &H01A1		'索引,0			返回列表框中某一项的高度（象素）
	'LB_GETITEMRECT = &H0198		'索引,RECT结构地址	获得列表项的客户区的RECT
	'LB_GETLOCALE = &H01A6			'0,0			取列表项当前用于排序的语言代码，当用户使用LB_ADDSTRING向组合框中的列表框中添加记录并使用LBS_SORT风格进行重新排序时，必须使用该语言代码。返回值中高16位为国家代码
	LB_GETSEL = &H0187				'索引,0			返回指定列表项的状态。如果查询的列表项被选择了，函数返回一个正值，否则返回0，若出错则返回LB_ERR
	LB_GETSELCOUNT = &H0190			'0,0			本操作仅用于多重选择列表框，它返回选择项的数目，若出错函数返回LB_ERR
	LB_GETSELITEMS = &H0191			'数组的大小,缓冲区	本操作仅用于多重选择列表框，用来获得选中的项的数目及位置。参数lParam指向一个整型数数组缓冲区，用来存放选中的列表项的索引。wParam说明了数组缓冲区的大小。本操作返回放在缓冲区中的选择项的实际数目，若出错函数返回LB_ERR
	'LB_GETTEXT = &H0189			'索引,缓冲区 	用于获取指定列表项的字符串。参数lParam指向一个接收字符串的缓冲区．wParam则指定了接收字符串的列表项索引。返回获得的字符串的长度，若出错，则返回LB_ERR
	'LB_GETTEXTLEN = &H018A			'索引,0 		返回指定列表项的字符串的字节长度。wParam指定了列表项的索引．若出错则返回LB_ERR返回和給定項相關聯的字符串長度（單位：字符）
	LB_GETTOPINDEX = &H018E			'0,0			返回列表框中第一个可见项的索引，若出错则返回LB_ERR
	'LB_INITSTORAGE = &H01A8		'表项数,内存字节数	本操作只适用于Windows95版本，当你将要向列表框中加入很多表项或有很大的表项时，本操作将预先分配一块内存，以免在今后的操作中一次一次地分配内存，从而加快程序运行速度
	LB_INSERTSTRING = &H0181		'索引,字符串地址	在列表框中的指定位置插入字符串。wParam指定了列表项的索引，如果为-1，则字符串将被添加到列表的末尾。lParam指向要插入的字符串。本操作返回实际的插入位置，若发生错误，会返回LB_ERR或LB_ERRSPACE。与LB_ADDSTRING不同，本操作不会导致LBS_SORT风格的列表框重新排序。建议不要在具有LBS_SORT风格的列表框中使用本操作，以免破坏列表项的次序
	'LB_ITEMFROMPOINT = &H01A9		'0,位置			获得与指定点最近的项的索引，lParam指定在列表框客户区，低16位为X坐标，高16位为Y坐标
	LB_RESETCONTENT = &H0184		'0,0			清除所有列表项
	'LB_SELECTSTRING = &H018C		'开始表项序号,字符串地址	本操作仅适用于单选择列表框，设定与指定字符串相匹配的列表项为选中项。本操作会滚动列表框以使选择项可见。参数的意义及搜索的方法与LB_FINDSTRING类似。如果找到了匹配的项，返回该项的索引，如果没有匹配的项，返回LB_ERR并且当前的选中项不被改变
	'LB_SELITEMRANGE = &H019B		'TRUE或FALSE,范围	本操作仅用于多重选择列表框，用来使指定范围内的列表项选中或落选．参数lParam指定了列表项索引的范围，低16位为开始项高16位为结束项。如果参数wParam为TRUE，那么就选择这些列表项，否则就使它们落选。若出错函数返回LB_ERR
	'LB_SELITEMRANGEEX = &H0183		'起点,终点		仅用于多重选择列表框，若指定终点大于起点则设定该范围为选中，若指定起点大于终点则设定该范围为落选
	'LB_SETANCHORINDEX = &H019C		'索引,0			设置鼠标最后选中的表项成指定表项
	'LB_SETCARETINDEX = &H019E		'索引,TRUE或FALSE	设置键盘输入焦点到指定表项，若lParam为TRUE则滚动到指定项部份可见，若lParam为FALSE则滚动到指定项全部可见
	'LB_SETCOLUMNWIDTH = &H0195		'宽度(点),0		设置列的宽度（單位：象素）
	'LB_SETCOUNT = &H01A7			'项数,0			设置表项数目
	'LB_SETCURSEL = &H0186			'索引,0			仅适用于单选择列表框，设置指定的列表项为当前选择项，并自动滚动到可见区域。参数wParam指定了列表项的索引，若为-1，那么将清除列表框中的选择。若出错函数返回LB_ERR
	LB_SETHORIZONTALEXTENT = &H0194	'宽度(点),0 设置列表框的滚动宽度（單位：象素）
	'LB_SETITEMDATA = &H019A		'索引,数据值	更新指定列表项的32位附加数据。
	'LB_SETITEMHIEGHT = &H01A0		'索引,高度(点)	指定列表项显示高度，带有LBS_OWNERDRAWVARIABLE(自绘列表项)风格的控件，只设置由wParam指定项的高度，其它风格将更新所有的列表项的高度（單位：象素）
	'LB_SETLOCALE = &H01A5			'语言代码,0		取列表项当前用于排序的语言代码，当用户使用LB_ADDSTRING向组合框中的列表框中添加记录并使用LBS_SORT风格进行重新排序时，必须使用该语言代码。返回值中高16位为国家代码
	LB_SETSEL = &H0185				'TRUE或FALSE,索引	仅适用于多重选择列表框，它使指定的列表项选中或落选，并自动滚动到可见区域。参数lParam指定了列表项的索引，若为-1，则相当于指定了所有的项。参数wParam为TRUE时选中列表项，否则使之落选。若出错则返回LB_ERR
	'LB_SETTABSTOPS = &H0192		'站数,索引顺序表	设置列表框的光标(输入焦点)站数及索引顺序表
	LB_SETTOPINDEX = &H0197			'索引,0			用来将指定的列表项设置为列表框的第一个可见项，该函数会将列表框滚动到合适的位置。wParam指定了列表项的索引．若操作成功，返回0值，否则返回LB_ERR
	'LB_MULTIPLEADDSTRING = &H01B1
	'LB_GETLISTBOXINFO = &H01B2
	'LB_MSGMAX_501 = &H01B3
	'LB_MSGMAX_WCE4 = &H01B1
	'LB_MSGMAX_4 = &H01B0
	'LB_MSGMAX_PRE4 = &H01A8

	'用于文本控件常数
	WM_GETTEXT = &H0D				'字节数,字符串地址	获取窗口文本控件的文本
	WM_GETTEXTLENGTH = &H0E			'0,0				获取窗口文本控件的文本的长度（不包含空字符）(字符数)
	WM_SETTEXT = &H0C				'0,字符串地址		设置窗口文本控件的文本

	'用于对话框字体
	WM_SETFONT = &H30				'字体句柄,True		绘制文本时程序发送此消息获取控件要用的字体
	WM_GETFONT = &H31				'0,0 				获取当前控件绘制文本的字体句柄
	WM_FONTCHANGE = &H1D 			'0,0				当系统的字体资源库变化时发送此消息给所有顶级窗口
	WM_SETREDRAW = &H0B 			'Boolean,0			设置窗口是否能重画，False 禁止重画，True 允许重画
	'WM_CTLCOLORMSGBOX = &H132		'设备句柄,控件句柄	设置消息框颜色
	'WM_CTLCOLOREDIT = &H133		'设备句柄,控件句柄	设置编辑框颜色
	'WM_CTLCOLORLISTBOX = &H134		'设备句柄,控件句柄	设置列表框颜色
	'WM_CTLCOLORBTN = &H135			'设备句柄,控件句柄	设置按钮颜色
	'WM_CTLCOLORDLG = &H136			'设备句柄,控件句柄	设置对话框颜色
	'WM_CTLCOLORSCROLLBAR = &H137	'设备句柄,控件句柄	设置滚动条颜色
	'WM_CTLCOLORSTATIC = &H138		'设备句柄,控件句柄	设置状态栏颜色
	WM_SETFOCUS = &H7				'控件句柄,0,0		设置焦点

	'WM_KEYDOWN = &H100				'控件句柄,虚拟键,0	模拟按下按钮
	'WM_KEYUP = &H101				'控件句柄,虚拟键,0	模拟抬起按钮
	'WM_LBUTTONDOWN = &H201			'移动鼠标
	'WM_LBUTTONUP = &H202			'按下鼠标左键
	'WM_LBUTTONDBLCLK = &H203		'释放鼠标左键
	'WM_RBUTTONDOWN = &H204			'双击鼠标左键
	'WM_RBUTTONUP = &H205			'按下鼠标右键
	'WM_RBUTTONDBLCLK = &H206		'释放鼠标右键
	'WM_MBUTTONDOWN = &H207			'双击鼠标右键
	WM_MBUTTONUP = &H208			'按下鼠标中键
	'WM_MBUTTONDBLCLK = &H209		'释放鼠标中键
	'WM_MOUSEWHEEL = &H20A			'双击鼠标中键

	'WM_HSCROLL= &H114				'控件句柄,滚动条类型,滚动条位置	设置 SB_BOTTOM 指定的水平滚动条位置
	WM_VSCROLL = &H115				'控件句柄,滚动条类型,滚动条位置	设置 SB_BOTTOM 指定的垂直滚动条位置
	'SB_TOP = &H06					'滚动条位置, 设置垂直滚动条到顶部
	'SB_LEFT = &H06					'滚动条位置, 设置水平滚动条到右边
	SB_BOTTOM = &H07				'滚动条位置, 设置垂直滚动条到底部
	'SB_RIGHT = &H07				'滚动条位置, 设置水平滚动条到右边

	'按钮事件
	'BM_GETCHECK = &HF0				'控件句柄,0,0		获取单选按钮或复选框的选定状态
	'BM_SETCHECK = &HF1				'控件句柄,0,0		设置单选按钮或复选框的选定状态
	BM_GETSTATE = &HF2				'控件句柄,0,0		获取按钮是否被按下过
	BM_SETSTATE = &HF3				'控件句柄,按钮状态,0	设置按钮的状态，True 未按下状态，False 按下状态
	'BM_SETSTYLE = &HF4				'控件句柄,按钮样式,0	设置按钮的样式
	BM_CLICK = &HF5					'控件句柄,0,0		模拟点击按钮
	'BM_GETIMAGE = &HF6
	'BM_SETIMAGE = &HF7

	'BN_CLICKED = &H0				'控件句柄,0,0		用户单击了按钮

	'BST_UNCHECKED = &H0      		'设置单选框和复选框复选框为未选中状态
	'BST_CHECKED = &H1				'设置单选框和复选框复选框为已选中状态
	'BST_INDETERMINATE = &H2
	BST_PUSHED = &H4      			'设置按钮为按下状态
	'BST_FOCUS = &H8      			'设置焦点
End Enum

'用于文本框查找定位函数
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
	ByVal hwnd As Long, _
	ByVal wMsg As Long, _
	ByVal wParam As Long, _
	ByRef lParam As Any) As Long
Public Declare Function SendMessageOLD Lib "user32.dll" Alias "SendMessageA" ( _
	ByVal hwnd As Long, _
	ByVal wMsg As Long, _
	ByVal wParam As Long, _
	ByVal lParam As Long) As Long
Public Declare Function SendMessageLNG Lib "user32.dll" Alias "SendMessage" ( _
	ByVal hwnd As Long, _
	ByVal wMsg As Long, _
	ByVal wParam As Long, _
	ByVal lParam As Long) As Long

'用于控件的显示和隐藏
'ShowWindow 部分常数
Public Enum ShowWindowValue
	SW_SHOW = 5
	SW_HIDE = 0
End Enum
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function EnableWindow Lib "user32" Alias "EnableWindow" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
'返回，非零表示成功，零表示失败。会设置GetLastError
'hwnd，窗口或控件句柄
'fEnable Long，非零允许，零禁止

'用于设置焦点到控件，会出错，无法使用
'Private Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long)
'用于返回焦点控件的句柄
Public Declare Function GetFocus Lib "user32.dll" () As Long
'用于返回控件ID的句柄
Public Declare Function GetDlgItem Lib "user32.dll" ( _
	ByVal hDlg As Long, _
	ByVal nIDDlgItem As Long) As Long
'获取指定点窗口的句柄
Public Declare Function WindowFromPoint Lib "user32.dll" ( _
	ByVal xPoint As Long, _
	ByVal yPoint As Long) As Long
'获取指定点子窗口的句柄
'Private Declare Function ChildWindowFromPoint Lib "user32.dll" ( _
'	ByVal hwnd As Long, _
'	ByVal xPoint As Long, _
'	ByVal yPoint As Long) As Long

'获取和设置滚动条位置函数
'Private Declare Function GetScrollPos Lib "user32.dll" ( _
'	ByVal hwnd As Long, _
'	ByVal nBar As Long) As Long
'Private Declare Function SetScrollPos Lib "user32.dll" ( _
'	ByVal hwnd As Long, _
'	ByVal nBar As Long, _
'	ByVal nPos As Long, _
'	ByVal bRedraw As Long) As Long

'像素矩形坐标定义
Private Type RECT
	Left As Long
	Top As Long
	Right As Long
	Bottom As Long
End Type

'DrawText 的 wFormat 参数定义
Private Enum DrawTextConstants
	DT_BOTTOM = &H8				'将正文调整到矩形底部。此值必须和 DT_SINGLELINE 组合。
	DT_CALCRECT = &H400			'决定矩形的宽和高。如果正文有多行，DrawText使用lpRect定义的矩形的宽度，并扩展矩形的底训以容纳正文的最后一行，
								'如果正文只有一行，则DrawText改变矩形的右边界，以容纳下正文行的最后一个字符，上述任何一种情况，DrawText返回格式化正文的高度而不是写正文。
	DT_CENTER = &H1				'使正文在矩形中水平居中。
	DT_EXPANDTABS = &H40		'扩展制表符，每个制表符的缺省字符数是8
	DT_EXTERNALLEADING = &H200	'在行的高度里包含字体的外部标头，通常，外部标头不被包含在正文行的高度里。
	DT_INTERNAL = &H1000		'用系统字体来计算正文度量。
	DT_LEFT = &H0				'正文左对齐。
	DT_NOCLIP = &H100			'无裁剪绘制当DT_NOCLIP使用时DrawText的使用会有所加快。
	DT_NOPREFIX = &H800			'关闭前缀字符的处理，通常DrawText解释助记前缀字符，&为给其后的字符加下划线，解释&&为显示单个&。指定DT_NOPREFIX，这种处理被关闭。
	DT_RIGHT = &H2				'正文右对齐。
	DT_SINGLELINE = &H20		'显示正文的同一行，回车和换行符都不能折行。
	DT_TABSTOP = &H80			'设置制表，参数uFormat的15"C8位（低位字中的高位字节）指定每个制表符的字符数，每个制表符的缺省字符数是8。
	DT_TOP = &H0				'正文顶端对齐（仅对单行）。
	DT_VCENTER = &H4			'正文水平居中（仅对单行）。
	DT_WORDBREAK = &H10			'断开字。当一行中的字符将会延伸到由lpRect指定的矩形的边框时，此行自动地在字之间断开。一个回车一换行也能使行折断。
	DT_EDITCONTROL = &H2000&	'复制多行编辑控制的正文显示特性，特殊地，为编辑控制的平均字符宽度是以同样的方法计算的，此函数不显示只是部分可见的最后一行。
	DT_END_ELLIPSIS = &H8000&	'可以指定DT_END_ELLIPSIS来替换在字符串末尾的字符，或指定DT_PATH_ELLIPSIS来替换字符串中间的字符。
								'如果字符串里含有反斜扛，DT_PATH_ELLIPSIS尽可能地保留最后一个反斜杠之后的正文。
	DT_MODIFYSTRING = &H10000	'修改给定的字符串来匹配显示的正文，此标志必须和DT_END_ELLIPSIS或DT_PATH_ELLIPSIS同时使用。
	DT_PATH_ELLIPSIS = &H4000&
	DT_RTLREADING = &H20000		'当选择进设备环境的字体是Hebrew或Arabicf时，为双向正文安排从右到左的阅读顺序都是从左到右的。
	DT_WORD_ELLIPSIS = &H40000	'截短不符合矩形的正文，并增加椭圆。
	'注意：DT_CALCRECT, DT_EXTERNALLEADING, DT_INTERNAL, DT_NOCLIP, DT_NOPREFIX值不能和DT_TABSTOP值一起使用。
End Enum

'获取字符串的像素大小使用的函数
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" ( _
	ByVal hwnd As Long, _
	ByVal hDC As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" ( _
	ByVal hDC As Long, _
	ByVal lpStr As String, _
	ByVal nCount As Long, _
	ByRef lpRect As RECT, _
	ByVal wFormat As Long) As Long
'Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" ( _
'	ByVal hDC As Long, _
'	ByVal lpString As String, _
'	ByVal cbString As Long, _
'	ByRef lpSize As POINTAPI) As Long

'设置文本颜色
'Private Declare Function SetTextColor Lib "gdi32.dll" ( _
'	ByVal hDC As Long, _
'	ByVal crColor As Long) As Long

'鼠标 VK 键值定义，用于 GetAsyncKeyState 函数
Public Enum VK
	VK_LBUTTON = &H01	'鼠标左键
	VK_RBUTTON = &H02	'鼠标右键
	VK_MBUTTON = &H04	'鼠标中键
	'VK_END = &H23		'键盘 End 键
	'VK_HOME = &H24		'键盘 Home 键
	'VK_LEFT = &H25		'键盘向左键
	'VK_UP = &H26		'键盘向上键
	'VK_RIGHT = &H27	'键盘向右键
	'VK_DOWN = &H28		'键盘向下键
	VK_ESCAPE = &H1B	'Esc 键
End Enum

'获取鼠标按键状态
'Private Declare Function GetKeyState Lib "user32" Alias "GetKeyState" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
'GetAsyncKeyState 函数返回的是指定虚拟键瞬时的硬件中断状态值，它有四种返回值：
'0 键当前未处于按下状态，而且自上次调用GetAsyncKeyState后改键也未被按过
'1 键当前未处于按下状态，但在此之前（自上次调用GetAsyncKeyState后）键曾经被按过
'-32768（即16进制数&H8000）键当前处于按下状态，但在此之前（自上次调用GetAsyncKeyState后）键未被按过
'-32767（即16进制数&H8001）键当前处于按下状态，而且在此之前（自上次调用GetAsyncKeyState后）键也曾经被按过

'获取光标的屏幕位置
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
'移动光标到屏幕的指定位置
'Private Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
'转换屏幕光标坐标为客户区坐标
'Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'转换客户区光标坐标为屏幕坐标
'Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

' Logical Font
'Private Const LF_FACESIZE = 32
'Private Const LF_FULLFACESIZE = 64

'字体类型，用于 ChooseFont 函数
'Private Enum FONTTYPE
'	BOLD_FONTTYPE				'字体为粗体。此信息从 LOGFONT 结构成员 lfWeight 复制，并和FW_BOLD 等效。
'	ITALIC_FONTTYPE				'字体为斜体。此信息从 LOGFONT 结构成员 lfItalic 复制。
'	PRINTER_FONTTYPE			'字体为打印机字体。
'	REGULAR_FONTTYPE = &H400	'字体为标准。此信息是从在 LOGFONT 结构成员 lfWeight 复制，并和 FW_REGULAR 等效。
'	SCREEN_FONTTYPE				'字体为屏幕字体。
'	SIMULATED_FONTTYPE			'字体被图形设备接口 (GDI) 模拟
'End Enum

'字符集，用于 LOG_FONT 类型的 lfCharSet
'Private Enum CHARSET
'	ANSI_CHARSET = 0
'	DEFAULT_CHARSET = 1
'	SYMBOL_CHARSET = 2
'	MAC_CHARSET = 77
'	SHIFTJIS_CHARSET = 128
'	HANGEUL_CHARSET = 129
'	JOHAB_CHARSET = 130
'	GB2312_CHARSET = 134
'	CHINESEBIG5_CHARSET = 136
'	GREEK_CHARSET = 161
'	TURKISH_CHARSET = 162
'	HEBREW_CHARSET = 177
'	ARABIC_CHARSET = 178
'	BALTIC_CHARSET = 186
'	RUSSIAN_CHARSET = 204
'	THAI_CHARSET = 222
'	EASTEUROPE_CHARSET = 238
'	OEM_CHARSET = 255
'End Enum

'ChooseFont 类型的 flags 参数定义
Private Enum CF_VALUE
	CF_APPLY = &H200
	CF_ANSIONLY = &H400
	CF_TTONLY = &H40000
	CF_ENABLEHOOK = &H8
	CF_ENABLETEMPLATE = &H10
	CF_ENABLETEMPLATEHANDLE = &H20
	CF_FIXEDPITCHONLY = &H4000
	CF_NOVECTORFONTS = &H800
	CF_NOOEMFONTS = CF_NOVECTORFONTS
	CF_NOFACESEL = &H80000
	CF_NOSCRIPTSEL = CF_NOFACESEL
	CF_NOSTYLESEL = &H100000
	CF_NOSIZESEL = &H200000
	CF_NOSIMULATIONS = &H1000
	CF_NOVERTFONTS = &H1000000
	CF_SCALABLEONLY = &H20000
	CF_SCRIPTSONLY = CF_ANSIONLY
	CF_SELECTSCRIPT = &H400000
	CF_SHOWHELP = &H4
	CF_USESTYLE = &H80
	CF_WYSIWYG = &H8000			'must also have CF_SCREENFONTS CF_PRINTERFONTS
	CF_FORCEFONTEXIST = &H10000
	CF_INACTIVEFONTS = &H2000000
	CF_INITTOLOGFONTSTRUCT = &H40&
	CF_SCREENFONTS = &H1		'显示屏幕字体
	CF_PRINTERFONTS = &H2		'显示打印机字体
	CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)	'两者都显示
	CF_EFFECTS = &H100&			'添加字体效果
	CF_LIMITSIZE = &H2000&		'设置字体大小限制
End Enum

'字体类型
Public Type LOG_FONT
	lfHeight As Long			'字体大小
	lfWidth As Long				'字体宽度
	lfEscapement As Long		'字体显示角度
	lfOrientation As Long		'字体角度
	lfWeight As Long			'是否粗体
	lfItalic As Byte			'是否斜体
	lfUnderline As Byte			'是否下划线
	lfStrikeOut As Byte			'是否删除线
	lfCharSet As Byte			'字符集
	lfOutPrecision As Byte		'输出精度
	lfClipPrecision As Byte		'裁减精度
	lfQuality As Byte			'逻辑字体与输出设备实际字体之间的精度
	lfPitchAndFamily As Byte	'字体间距和字体集
	'lfFaceName As String * LF_FACESIZE	'字体名称(不能这样定义，创建字体时会出错)
	lfFaceName(31) As Byte		'字体名称
	lfColor As Long				'字体颜色
End Type

'字体对话框类型
Private Type CHOOSE_FONT
	lStructSize As Long			' size of CHOOSEFONT structure in byte
	hwndOwner As Long			' caller's window handle
	hDC As Long					' printer DC/IC or NULL
	lpLogFont As Long			' LogFont 结构地址
	iPointSize As Long			' 10 * size in points of selected font
	flags As CF_VALUE			' enum type flags
	rgbColors As Long			' returned text color
	lCustData As Long			' data passed to hook fn
	lpfnHook As Long			' ptr. to hook function
	lpTemplateName As String	' custom template name
	hInstance As Long			' instance handle of.EXE that contains cust. dlg. template
	lpszStyle As String			' return the style field here must be LF_FACESIZE or bigger
	nFontType As Integer		' same value reported to the EnumFonts call back with the extra FONTTYPE_ bits added
	MISSING_ALIGNMENT As Integer
	nSizeMin As Long			' minimum pt size allowed
	nSizeMax As Long			' max pt size allowed if CF_LIMITSIZE is used
End Type

Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSE_FONT) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOG_FONT) As Long
Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" ( _
	ByVal hObject As Long, _
	ByVal nCount As Long, _
	ByVal lpObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" ( _
	ByVal hDC As Long, _
	ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

'RedrawWindow 函数的 fuRedraw 参数定义
Private Enum RDW
	RDW_INVALIDATE = &H1		'禁用（屏蔽）重画区域
	RDW_INTERNALPAINT = &H2		'即使窗口并非无效，也向其投递一条WM_PAINT消息
	RDW_ERASE = &H4				'重画前，先清除重画区域的背景。也必须指定RDW_INVALIDATE
	RDW_VALIDATE = &H8			'检验重画区域
	RDW_NOINTERNALPAINT = &H10	'禁止内部生成或由这个函数生成的任何待决WM_PAINT消息。针对无效区域，仍会生成WM_PAINT消息
	RDW_NOERASE = &H20			'禁止删除重画区域的背景
	RDW_NOCHILDREN = &H40		'重画操作排除子窗口（前提是它们存在于重画区域）
	RDW_ALLCHILDREN = &H80		'重画操作包括子窗口（前提是它们存在于重画区域）
	RDW_UPDATENOW = &H100		'立即更新指定的重画区域
	RDW_ERASENOW = &H200		'立即删除指定的重画区域
	RDW_FRAME = &H400			'如非客户区包含在重画区域中，则对非客户区进行更新。也必须指定RDW_INVALIDATE
	RDW_NOFRAME = &H800			'禁止非客户区域重画（如果它是重画区域的一部分）。也必须指定RDW_VALIDATE
End Enum

'重画对话框函数
Private Declare Function RedrawWindow Lib "user32.dll" ( _
	ByVal hwnd As Long, _
	ByVal lprcUpdate As Long, _
	ByVal hrgnUpdate As Long, _
	ByVal fuRedraw As Long) As Long

'读取注册表值名函数
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" ( _
	ByVal hKey As Long, _
	ByVal lpSubKey As String, _
	phkResult As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" ( _
	ByVal hKey As Long, _
	ByVal dwIndex As Long, _
	ByVal lpValueName As String, _
	lpcbValueName As Long, _
	ByVal lpReserved As Long, _
	lpType As Long, _
	lpData As Byte, _
	lpcbData As Long) As Long

'获取 Windows 版本号
Private Type OSVERSIONINFO
	dwOSVersionInfoSize As Long		'在使用GetVersionEx之前要将此初始化为结构的大小
	dwMajorVersion As Long			'系统主版本号
	dwMinorVersion As Long			'系统次版本号
	dwBuildNumber As Long			'系统构建号
	dwPlatformId As Long			'系统支持的平台(详见附1)
	szCSDVersion As String * 128	'系统补丁包的名称
	wServicePackMajor As Integer	'系统补丁包的主版本
	wServicePackMinor As Integer	'系统补丁包的次版本
	wSuiteMask As Integer			'标识系统上的程序组(详见附2)
	wProductType As Byte			'标识系统类型(详见附3)
	wReserved As Byte				'保留,未使用
End Type
'获取版本信息
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'获取窗口标题或控件中的文本字符数
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLength" (ByVal hwnd As Long) As Long
'获取窗口标题或控件中的文本
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowText" (ByVal hwnd As Long, _
	ByVal lpString As String, _
	ByVal cch As Long) As Long
'设置窗口标题或控件中的文本
'不能改变在其他应用程序中的控件的文本内容，如果需要可以用 SengMessage 函数发送一条 WM_SETTEX 消息。
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowText" (ByVal hwnd As Long, _
	ByVal lpString As String) As Long

'内存映射文件定义
Private Enum FileMaping
	PAGE_READONLY = &H02
	PAGE_READWRITE = &H04
	PAGE_WRITECOPY = &H08
	PAGE_EXECUTE = &H10
	PAGE_EXECUTE_READ = &H20
	PAGE_EXECUTE_READWRITE = &H40
	PAGE_EXECUTE_WRITECOPY = &H80
	SEC_COMMIT = &H8000000
	SEC_IMAGE = &H1000000
	SEC_NOCACHE = &H10000000
	SEC_RESERVE = &H4000000
	SEC_IMAGE_NO_EXECUTE = &H11000000
	SEC_LARGE_PAGES = &H80000000
	SEC_WRITECOMBINE = &H40000000
	GENERIC_READ = &H80000000
	GENERIC_WRITE = &H40000000
	OPEN_EXISTING = 3
	OPEN_ALWAYS = 4
	FILE_MAP_COPY = &H01
	FILE_MAP_WRITE = &H02
	FILE_MAP_READ = &H04
	FILE_MAP_ALL_ACCESS = &H02 Or &H04
	FILE_MAP_EXECUTE = &H20
	FILE_SHARE_READ = &H01
	FILE_SHARE_WRITE = &H02
	FILE_ATTRIBUTE_NORMAL = &H80
	FILE_ATTRIBUTE_ARCHIVE = &H20
	FILE_ATTRIBUTE_READONLY = &H01
	FILE_ATTRIBUTE_HIDDEN = &H02
	FILE_ATTRIBUTE_SYSTEM = &H04
End Enum

'用于读写文件函数
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" ( _
	ByVal lpFileName As String, _
	ByVal dwDesiredAccess As Long, _
	ByVal dwShareMode As Long, _
	ByVal lpSecurityAttributes As Long, _
	ByVal dwCreationDisposition As Long, _
	ByVal dwFlagsAndAttributes As Long, _
	ByVal hTemplateFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" ( _
	ByVal lngFile As Long, _
	ByVal lDistanceToMove As Long, _
	lpDistanceToMoveHigh As Long, _
	ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" ( _
	ByVal lngFile As Long, _
	lpBuffer As Any, _
	ByVal nNumberOfBytesToRead As Long, _
	lpNumberOfBytesRead As Long, _
	ByVal lpOverlapped As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" ( _
	ByVal hFile As Long, _
	ByVal lpFileSizeHigh As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32.dll" (ByVal hFile As Long) As Long
'Private Declare Function WriteFile Lib "kernel32.dll" ( _
'	ByVal hFile As Long, _
'	ByVal lpBuffer As Long, _
'	ByVal nNumberOfBytesToWrite As Long, _
'	lpNumberOfBytesWritten As Long, _
'	ByVal lpOverlapped As Long) As Long

'用于文件映射函数
Private Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingA" ( _
	ByVal hFile As Long, _
	ByVal lpFileMappigAttributes As Long, _
	ByVal flProtect As Long, _
	ByVal dwMaximumSizeHigh As Long, _
	ByVal dwMaximumSizeLow As Long, _
	ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32.dll" ( _
	ByVal hFileMappingObject As Long, _
	ByVal dwDesiredAccess As Long, _
	ByVal dwFileOffsetHigh As Long, _
	ByVal dwFileOffsetLow As Long, _
	ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function MapViewOfFileEx Lib "kernel32.dll" ( _
	ByVal hFileMappingObject As Long, _
	ByVal dwDesiredAccess As Long, _
	ByVal dwFileOffsetHigh As Long, _
	ByVal dwFileOffsetLow As Long, _
	ByVal dwNumberOfBytesToMap As Long, _
	lpBaseAddress As Any) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByVal lpBaseAddress As Long) As Long
Private Declare Function FlushViewOfFile Lib "kernel32.dll" ( _
	ByVal lpBaseAddress As Long, _
	ByVal dwNumberOfBytesToFlush As Long) As Long
'Private Declare Function MapAndLoad Lib "imagehlp.dll" ( _
'	ByVal ImageName As String, _
'	ByVal DllPath As String, _
'	LoadedImage As LOADED_IMAGE, _
'	ByVal DotDll As Boolean, _
'	ByVal ReadOnly As Boolean) As Boolean
	'ImageName	载入的 PE 文件的文件名
	'DllPath 定位文件的路径。若传递 Null，则搜索 PATH 环境变量中的路径。
	'LoadedImage 结构体 LOADED_IMAGE 定义在 IMAGEHLP.H file
	'DotDll 若需要查找该文件而且没有指定扩展名，则使用 .exe 或 .dll 作扩展名
	'若 DotDll 标志设为 True，则使用 .dll 扩展名; 否则用 .exe 扩展名
	'ReadOnly 若设为 True，则文件被映射为 Read-only。
	'返回值 若成功，返回 True; 否则返回 False
'Private Declare Function UnMapAndLoad Lib "imagehlp.dll" (LoadedImage As LOADED_IMAGE) As Boolean
	'使用完映射文件后，应该调用 UnMapAndLoad() 函数。 此函数解除 PE 文件的映射并回收由 MapAndLoad() 分配的资源。
	'LoadedImage 指向 LOADED_IMAGE 结构体的指针，该指针就是前面调用 MapAndLoad() 函数返回的指针。
	'返回值 若成功，则返回 True; 否则返回 False

'错误消息定义
Private Enum FormatMSG
	FORMAT_MESSAGE_FROM_SYSTEM = &H1000
	FORMAT_MESSAGE_IGNORE_INSERTS = &H200
End Enum

'用于显示错误消息
Private Declare Function GetLastError Lib "kernel32.dll" () As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" ( _
	ByVal dwFlags As Long, _
	lpSource As Any, _
	ByVal dwMessageId As Long, _
	ByVal dwLanguageId As Long, _
	ByVal lpBuffer As String, _
	ByVal nSize As Long, _
	Arguments As Long) As Long

'用于修改文件的时间属性
Private Type FILETIME
	dwLowDateTime As Long
	dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
	wYear As Integer
	wMonth As Integer
	wDayOfWeek As Integer
	wDay As Integer
	wHour As Integer
	wMinute As Integer
	wSecond As Integer
	wMilliseconds As Integer
End Type

'用于修改文件的时间属性函数
Private Declare Function SetFileTime Lib "kernel32.dll" ( _
	ByVal hFile As Long, _
	lpCreationTime As FILETIME, _
	lpLastAccessTime As FILETIME, _
	lpLastWriteTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32.dll" ( _
	lpLocalFileTime As FILETIME, _
	lpFileTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32.dll" ( _
	lpSystemTime As SYSTEMTIME, _
	lpFileTime As FILETIME) As Long

'有关注册表导入导出常量
Private Enum REG_KEY_IMPORT_EXPORT
	REG_FORCE_RESTORE = 8&
	TOKEN_QUERY= &H8&
	TOKEN_ADJUST_PRIVILEGES = &H20&
	SE_PRIVILEGE_ENABLED = &H2
End Enum
Private Const SE_RESTORE_NAME As String = "SeRestorePrivilege"
Private Const SE_BACKUP_NAME As String = "SeBackupPrivilege"

'注册表关键字访问选项
Private Enum REG_KEY_ACCESS
	'READ_CONTROL = &H20000
	'KEY_QUERY_VALUE = &H1
	'KEY_SET_VALUE = &H2
	'KEY_CREATE_SUB_KEY = &H4
	'KEY_ENUMERATE_SUB_KEYS = &H8
	'KEY_NOTIFY = &H10
	'KEY_CREATE_LINK = &H20
	'KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
	'KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
	'KEY_EXECUTE = KEY_READ
	'KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + _
	'				KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
	'KEY_WOW64_32KEY = &H0200
	'KEY_WOW64_64KEY = &H0100
	'KEY_READ = &H20019
	'KEY_WRITE = &H20006
	'KEY_EXECUTE = &H20019
	KEY_ALL_ACCESS = &HF003F
End Enum

'注册表关键字根类型...
Public Enum REG_KEYROOT
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Private Type REG_LUID
    lowpart As Long
    highpart As Long
End Type

Private Type REG_LUID_AND_ATTRIBUTES
    pLuid As REG_LUID
    Attributes As Long
End Type

Private Type REG_TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges As REG_LUID_AND_ATTRIBUTES
End Type

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
	ByVal hKey As Long, _
	ByVal lpSubKey As String, _
	ByVal ulOptions As Long, _
	ByVal samDesired As Long, _
	ByRef phkResult As Long) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" ( _
	ByVal hKey As Long, _
	ByVal lpFile As String, _
	ByVal lpSecurityAttributes As Long) As Long
	'ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" ( _
	ByVal hKey As Long, _
	ByVal lpFile As String, _
	ByVal dwFlags As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" ( _
	ByVal ProcessHandle As Long, _
	ByVal DesiredAccess As Long, _
	ByRef TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" ( _
	ByVal lpSystemName As String, _
	ByVal lpName As String, _
	ByRef lpLuid As REG_LUID) As Long	'Returns a valid LUID which is important when making security changes in NT.
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" ( _
	ByVal TokenHandle As Long, _
	ByVal DisableAllPriv As Long, _
	ByRef NewState As REG_TOKEN_PRIVILEGES, _
	ByVal BufferLength As Long, _
	ByRef PreviousState As REG_TOKEN_PRIVILEGES, _
	ByRef ReturnLength As Long) As Long	'Used to adjust your program's security privileges, can't restore without it!

'自定义弹出菜单
Private Enum POP_MENU
	MF_ENABLED = &H0&
	MF_BYCOMMAND = &H0&
	MF_STRING = &H0&
	MF_GRAYED = &H1&
	MF_DISABLED = &H2&
	MF_CHECKED = &H8&
	MF_POPUP = &H10&
	MF_BYPOSITION = &H400&
	MF_SEPARATOR = &H800&

	TPM_RIGHTBUTTON = &H2&
	TPM_LEFTALIGN = &H0&
	TPM_NONOTIFY = &H80&
	TPM_RETURNCMD = &H100&
End Enum

Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenu" ( _
	ByVal hMenu As Long, _
	ByVal wFlags As Long, _
	ByVal wIDNewItem As Long, _
	ByVal sCaption As String) As Long
Private Declare Function TrackPopupMenu Lib "user32" ( _
	ByVal hMenu As Long, _
	ByVal wFlags As Long, _
	ByVal X As Long, _
	ByVal Y As Long, _
	ByVal nReserved As Long, _
	ByVal hwnd As Long, _
	ByRef nIgnored As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuString" ( _
	ByVal hMenu As Long, _
	ByVal wIDItem As Long, _
	ByVal lpString As String, _
	ByVal nMaxCount As Long, _
	ByVal wFlag As Long) As Long

'获取屏幕分辨率和设置窗口大小
'Private Const SM_CXSCREEN = 0
'Private Const SM_CYSCREEN = 1
'Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
'	ByVal hWndInsertAfter As Long, _
'	ByVal x As Long, _
'	ByVal y As Long, _
'	ByVal cx As Long, _
'	ByVal cy As Long, _
'	ByVal wFlags As Long) As Long
'Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
'	ByVal x As Long, _
'	ByVal y As Long, _
'	ByVal nWidth As Long, _
'	ByVal nHeight As Long, _
'	ByVal bRepaint As Long) As Long

Public OSLanguage As String,Selected() As String,UpdateSet() As String
Public UIDataList() As INIFILE_DATA,LangFile As String,LFList() As LOG_FONT
Public CheckHexStr() As CHECK_STRING_VALUE,CheckSkipStr() As String

Public SourceFile As FILE_PROPERTIE,TargetFile As FILE_PROPERTIE

Public UniLangList() As LANG_PROPERTIE,UseLangList() As LANG_PROPERTIE
Public LangIDIndexDic As Object

Public AllStrList() As String,UseStrList() As String,UseStrListBak() As String
Public Tools() As TOOLS_PROPERTIE,ToolsBak() As TOOLS_PROPERTIE
Public FileDataList() As String,CodeList() As CODEPAGE_DATA,RegExp As Object

Public ExtractSet() As String,StrTypeList() As STRING_TYPE,RefTypeList() As REF_TYPE


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


'检查文本的代码页
Function CheckCodePage(ByVal textStr As String,CPList() As LANG_PROPERTIE,Optional ByVal Mode As Integer) As Long
	Dim i As Long
	CheckCodePage = CP_UNKNOWN
	If textStr = "" Then Exit Function
	Select Case Mode
	Case 0
		For i = 0 To UBound(CPList)
			With CPList(i)
				If .CodePage <> CP_UNKNOWN Then
					If MultiByteToUTF16(StringToByte(textStr,.CodePage),.CodePage) = textStr Then
						CheckCodePage = .CodePage
						Exit For
					End If
				End If
			End With
		Next i
	Case 1
		For i = 0 To UBound(CPList)
			With CPList(i)
				If .CodePage <> CP_UNKNOWN Then
					If CheckStrRegExp(textStr,.UniCodeRegExpPattern,0,1) = True Then
						CheckCodePage = .CodePage
						Exit For
					End If
				End If
			End With
		Next i
	Case 2
		For i = 0 To UBound(CPList)
			With CPList(i)
				If .CodePage <> CP_UNKNOWN Then
					If CheckStrRegExp(textStr,.FeatureCodeRegExpPattern,0,2) = True Then
						If CheckStrRegExp(textStr,.UniCodeRegExpPattern,0,1) = True Then
							CheckCodePage = .CodePage
							Exit For
						End If
					End If
				End If
			End With
		Next i
	End Select
End Function


'自动监测字节数组的代码页
'fType = 0 不前补位，否则.CodePage < 1000、UniCode、UniCode BE 前补位
Public Function getCodePageRegExp(FN As FILE_IMAGE,strData As STRING_SUB_PROPERTIE,CPList() As LANG_PROPERTIE, _
			EncodeList() As Integer,ByVal MinLength As Integer,StrTypeEndChar() As String,EndChar() As String, _
			ByVal Mode As Long,ByVal fType As Integer,ByVal LengthFilterSet As Long,ByVal MaxLength As Long) As Long
	Dim i As Long,j As Long,k As Long,l As Long,m As Long,n As Long,sp As Long,ep As Long
	Dim Matches As Object,Data As STRING_SUB_PROPERTIE
	On Error GoTo ExitFunciton
	'初始化
	getCodePageRegExp = strData.lEndAddress
	strData.CodePage = CP_UNKNOWN
	With Data
	.lStartAddress = strData.lStartAddress
	.lEndAddress = strData.lEndAddress
	.lMaxHexLength = strData.lHexLength
	'检测代码页并获取字串地址
	sp = .lStartAddress: ep = .lEndAddress
	For i = 0 To UBound(CPList)
		For j = 0 To UBound(EncodeList)
			k = EncodeList(j)
			Select Case k
			Case 1
				.CodePage = CPList(i).CodePage
				m = 0: l = 0
			Case 2
				.CodePage = CP_UNICODELITTLE
				m = 0: l = IIf(fType = 0,0,1)
			Case 3
				.CodePage = CP_UTF8
				k = 1: m = 0: l = 0
			Case 4
				.CodePage = CP_UNICODEBIG
				k = 2: m = IIf(fType = 0,0,1): l = 0
			Case 5
				.CodePage = CP_UTF7
				k = 1: m = 0: l = 0
			Case 6
				.CodePage = CP_UTF32LE
				k = 4: m = 0: l = IIf(fType = 0,0,3)
			Case 7
				.CodePage = CP_UTF32BE
				k = 4: m = IIf(fType = 0,0,3): l = 0
			End Select
			If LengthFilterSet = 1 Then
				.sString = Replace$(Replace$(EndChar(1),")(",StrTypeEndChar(i,j),,1),")(","(",,1)
				If getEndByteRegExp(FN,sp,ep,Mode,.sString,.CodePage) - sp > MaxLength * k Then GoTo NextNo
			End If
			For n = m To l Step IIf(m = 0,1,-1)
				RegExp.Global = False
				RegExp.IgnoreCase = False
				RegExp.Pattern = EndChar(0) & CPList(i).UniCodeRegExpPattern & "{" & CStr$(MinLength) & ",}" & _
								Replace$(EndChar(1),")(",StrTypeEndChar(i,j),,1)
				Set Matches = RegExp.Execute(ByteToString(GetBytes(FN,.lMaxHexLength + k,sp - n,Mode),.CodePage) & vbNullChar)
				If Matches.Count > 0 Then
					.lHexLength = -1
					.lStartAddress = sp - n: .lMaxAddress = sp - n
					Call GetStrAddress(FN,Data,Matches(0),Mode)
					If .lHexLength > -1 Then
						If fType = 0 Then
							If .lStartAddress = sp Then
								strData.lStartAddress = .lStartAddress
								strData.lEndAddress = .lEndAddress
								strData.lMaxAddress = .lMaxAddress
								strData.CodePage = .CodePage
								strData.lHexLength = .lHexLength
								strData.sString = Matches(0).SubMatches(1)
								getCodePageRegExp = .lMaxAddress
								Exit Function
							End If
						Else
							strData.lStartAddress = .lStartAddress
							strData.lEndAddress = .lEndAddress
							strData.lMaxAddress = .lMaxAddress
							strData.CodePage = .CodePage
							strData.lHexLength = .lHexLength
							strData.sString = Matches(0).SubMatches(1)
							getCodePageRegExp = .lMaxAddress
							Exit Function
						End If
					End If
					Exit For
				End If
			Next n
			NextNo:
		Next j
	Next i
	End With
	Exit Function
	ExitFunciton:
	getCodePageRegExp = strData.lEndAddress
	strData.CodePage = CP_UNKNOWN
End Function


'获取正则表达式字串的地址
Public Sub GetStrAddress(FN As FILE_IMAGE,strData As STRING_SUB_PROPERTIE,Match As Object,ByVal Mode As Long)
	With strData
	Select Case .CodePage
	Case CP_WESTEUROPE
		.lCharLength = .lStartAddress + Match.FirstIndex
		.lStartAddress = .lCharLength + Len(Match.SubMatches(0))
		.lEndAddress = .lStartAddress + Len(Match.SubMatches(1)) - 1
		.lMaxAddress = .lEndAddress + Len(Match.SubMatches(2))
		.lHexLength = .lEndAddress - .lStartAddress + 1
	Case CP_UNICODELITTLE, CP_UNICODEBIG
		.lCharLength = .lStartAddress + Match.FirstIndex * 2
		.lStartAddress = .lCharLength + Len(Match.SubMatches(0)) * 2
		.lEndAddress = .lStartAddress + Len(Match.SubMatches(1)) * 2 - 1
		.lMaxAddress = .lEndAddress + Len(Match.SubMatches(2)) * 2
		.lHexLength = .lEndAddress - .lStartAddress + 1
	Case CP_UTF32LE, CP_UTF_32LE, CP_UTF32BE, CP_UTF_32BE
		.lCharLength = .lStartAddress + Match.FirstIndex * 4
		.lStartAddress = .lCharLength + Len(Match.SubMatches(0)) * 4
		.lEndAddress = .lStartAddress + Len(Match.SubMatches(1)) * 4 - 1
		.lMaxAddress = .lEndAddress + Len(Match.SubMatches(2)) * 4
		.lHexLength = .lEndAddress - .lStartAddress + 1
	Case Else
		Dim tempByte() As Byte
		If .lMaxAddress < .lStartAddress + Match.FirstIndex Then
			.lMaxAddress = .lStartAddress + Match.FirstIndex
		End If
		tempByte = StringToByte(Match.SubMatches(0) & Match.SubMatches(1) & Match.SubMatches(2),.CodePage)
		If Mode = 0 Then
			.lCharLength = InByteRegExp(FN.ImageByte,tempByte,.lMaxAddress,.lEndAddress + 1) - 1
		Else
			.lCharLength = .lMaxAddress + InByteRegExp(GetBytes(FN,.lEndAddress - .lMaxAddress + 1,.lMaxAddress,Mode),tempByte) - 1
		End If
		If .lCharLength >= .lMaxAddress Then
			.lStartAddress = .lCharLength + StrHexLength(Match.SubMatches(0),.CodePage,0)
			.lEndAddress = .lStartAddress + StrHexLength(Match.SubMatches(1),.CodePage,0) - 1
			.lMaxAddress = .lEndAddress + StrHexLength(Match.SubMatches(2),.CodePage,0)
			.lHexLength = .lEndAddress - .lStartAddress + 1
		End If
	End Select
	End With
End Sub


'反转 Hex 码
Public Function ReverseHexCode(ByVal HexStr As String,Optional ByVal Num As Long) As String
	Dim i As Long
	i = Len(HexStr)
	If Num = 0 Then Num = GetEvenPos(i)
	If i < Num Then HexStr = String$(Num - i,"0") & HexStr
	ReverseHexCode = HexStr
	For i = 1 To Num - 1 Step 2
		Mid$(ReverseHexCode,i,2) = Mid$(HexStr,Num - i,2)
	Next i
End Function


'字节转 Hex 码
'StartPos <= EndPos 获取低位到高位的 Hex 代码，否则获取高位到低位的 Hex 代码
Public Function Byte2Hex(Bytes As Variant,ByVal StartPos As Long,ByVal endPos As Long) As String
	Dim i As Long,n As Long
	If StartPos < 0 Then StartPos = LBound(Bytes)
	If endPos < 0 Then endPos = UBound(Bytes)
	Byte2Hex = Space$((Abs(endPos - StartPos) + 1) * 2)
	n = 1
	For i = StartPos To endPos Step IIf(StartPos <= endPos,1,-1)
		Mid$(Byte2Hex,n,2) = Right$("0" & Hex$(Bytes(i)),2)
		n = n + 2
	Next i
End Function


'字节转 Hex 转义码
'StartPos <= EndPos 获取低位到高位的 Hex 代码，否则获取高位到低位的 Hex 代码
Private Function Byte2HexEsc(Bytes() As Byte,ByVal StartPos As Long,ByVal endPos As Long,ByVal CodePage As Long) As String
	Dim i As Long,n As Long
	If StartPos < 0 Then StartPos = LBound(Bytes)
	If endPos < 0 Then endPos = UBound(Bytes)
	Select Case CodePage
	Case CP_UNICODELITTLE
		Byte2HexEsc = Space$((Abs(endPos - StartPos) + 1) * 3)
		n = 1
		For i = StartPos To endPos - 1 Step IIf(StartPos <= endPos,2,-2)
			Mid$(Byte2HexEsc,n,6) = "\u" & Right$("0" & Hex$(Bytes(i + 1)),2) & Right$("0" & Hex$(Bytes(i)),2)
			n = n + 6
		Next i
	Case CP_UNICODEBIG
		Byte2HexEsc = Space$((Abs(endPos - StartPos) + 1) * 3)
		n = 1
		For i = StartPos To endPos - 1 Step IIf(StartPos <= endPos,2,-2)
			Mid$(Byte2HexEsc,n,6) = "\u" & Right$("0" & Hex$(Bytes(i)),2) & Right$("0" & Hex$(Bytes(i + 1)),2)
			n = n + 6
		Next i
	Case CP_UTF32LE, CP_UTF_32LE
		Byte2HexEsc = Space$((Abs(endPos - StartPos) + 1) * 2.5)
		n = 1
		For i = StartPos To endPos - 1 Step IIf(StartPos <= endPos,4,-4)
			Mid$(Byte2HexEsc,n,10) = "\u" & Right$("0" & Hex$(Bytes(i + 3)),2) & Right$("0" & Hex$(Bytes(i + 2)),2) & _
									Right$("0" & Hex$(Bytes(i + 1)),2) & Right$("0" & Hex$(Bytes(i)),2)
			n = n + 10
		Next i
	Case CP_UTF32BE, CP_UTF_32BE
		Byte2HexEsc = Space$((Abs(endPos - StartPos) + 1) * 2.5)
		n = 1
		For i = StartPos To endPos - 1 Step IIf(StartPos <= endPos,4,-4)
			Mid$(Byte2HexEsc,n,10) = "\u" & Right$("0" & Hex$(Bytes(i)),2) & Right$("0" & Hex$(Bytes(i + 1)),2) & _
									Right$("0" & Hex$(Bytes(i + 2)),2) & Right$("0" & Hex$(Bytes(i + 3)),2)
			n = n + 10
		Next i
	Case Else
		Byte2HexEsc = Space$((Abs(endPos - StartPos) + 1) * 4)
		n = 1
		For i = StartPos To endPos Step IIf(StartPos <= endPos,1,-1)
			Mid$(Byte2HexEsc,n,4) = "\x" & Right$("0" & Hex$(Bytes(i)),2)
			n = n + 4
		Next i
	End Select
End Function


'检测数值的字节序
'CheckByteOrder 返回检测结果，-1 = 大端，0 = 小端，1 = 未知
Private Function CheckByteOrder(Bytes() As Byte,ByVal Value As Long,ByVal Length As Integer, _
				ByVal StartNo As Long,ByVal EndNo As Long) As Integer
	Dim tmpBytes() As Byte
	CheckByteOrder = 1
	If EndNo < 0 Then EndNo = UBound(Bytes)
	tmpBytes = Val2Bytes(Value,Length,False)
	If InByteRegExp(Bytes,tmpBytes,StartNo,EndNo) > 0 Then
		CheckByteOrder = False
	ElseIf InByteRegExp(Bytes,ReverseValByte(tmpBytes,0,-1),StartNo,EndNo) > 0 Then
		CheckByteOrder = True
	End If
End Function


'转换字节数组为数值
'ByteOrder = False 按高位在后转，否则按高位在前转
Public Function Bytes2Val(Bytes() As Byte,ByVal Length As Integer,Optional ByVal ByteOrder As Boolean) As Long
	On Error GoTo errHandle
	If UBound(Bytes) + 1 < Length Then Exit Function
	If ByteOrder = False Then
		CopyMemory Bytes2Val, Bytes(0), Length
	Else
		CopyMemory Bytes2Val, ReverseValByte(Bytes,0,-1)(0), Length
	End If
	errHandle:
End Function


'转换数值为字节数组(短于长度的高位截断)
Public Function Val2Bytes(ByVal Value As Long,ByVal Length As Integer,Optional ByVal ByteOrder As Boolean) As Byte()
	On Error GoTo errHandle
	ReDim Bytes(Length - 1) As Byte
	CopyMemory Bytes(0), Value, Length
	If ByteOrder = False Then
		Val2Bytes = Bytes
	Else
		Val2Bytes = ReverseValByte(Bytes,0,-1)
	End If
	Exit Function
	errHandle:
	ReDim Bytes(0) As Byte
	Val2Bytes = Bytes
End Function


'转换数值为字节数组(短于长度的低位截断)
Private Function Val2BytesRev(ByVal Value As Long,ByVal MaxLength As Integer,ByVal Length As Integer,Optional ByVal ByteOrder As Boolean) As Byte()
	On Error GoTo errHandle
	If GetEvenPos(Len(Hex$(Value))) \ 2 > MaxLength Then GoTo errHandle
	ReDim Bytes(MaxLength - 1) As Byte
	CopyMemory Bytes(0), Value, MaxLength
	If MaxLength > Length Then
		CopyMemory Bytes(0), Bytes(MaxLength - Length), Length
		ReDim Preserve Bytes(Length - 1) As Byte
	End If
	If ByteOrder = False Then
		Val2BytesRev = Bytes
	Else
		Val2BytesRev = ReverseValByte(Bytes,0,-1)
	End If
	Exit Function
	errHandle:
	ReDim Bytes(0) As Byte
	Val2BytesRev = Bytes
End Function


'转换 HEX 代码为字节数组
Public Function HexStr2Bytes(ByVal HexStr As String) As Byte()
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


'检查字串是否包含指定字符(文本和通配符比较)
'Mode = 0 检查字串是否包含指定字符，并找出指定字符的位置
'Mode = 1 检查字串是否只包含指定字符
'Mode = 2 检查字串是否包含指定字符
'Mode = 3 检查字串是否只包含大小混写的指定字符，此时 IgnoreCase 参数无效
'Mode = 4 检查字串中是否有连续相同的字符，StrNum 为检查的字符个数
'StrRange  定义字串检查范围 (可用 [Min - Max|Min - Max] 表示范围)
Private Function CheckStr(ByVal textStr As String,ByVal StrRange As String,Optional ByVal StrNum As Long, _
				Optional ByVal Mode As Long,Optional ByVal IgnoreCase As Boolean) As Long
	Dim i As Long,Temp As String
	If StrRange = "" Then Exit Function
	If Trim$(textStr) = "" Then Exit Function
	Select Case Mode
	Case 0
		If IgnoreCase = True Then
			textStr = LCase$(textStr)
			StrRange = LCase$(StrRange)
		End If
		StrRange = Replace$(StrRange,"]|[","")
		If (textStr Like "*" & StrRange & "*") = False Then Exit Function
		For i = 1 To Len(textStr)
			If (Mid$(textStr,i,1) Like StrRange) = True Then
				CheckStr = i
				Exit For
			End If
		Next i
	Case 1
		If IgnoreCase = True Then
			textStr = LCase$(textStr)
			StrRange = LCase$(StrRange)
		End If
		If (textStr Like "*[!" & Replace$(Replace$(StrRange,"]|[",""),"[","") & "*") = False Then
			CheckStr = True
		End If
	Case 2
		If IgnoreCase = True Then
			textStr = LCase$(textStr)
			StrRange = LCase$(StrRange)
		End If
		If (textStr Like "*[" & Replace$(Replace$(StrRange,"]|[",""),"[","") & "*") = True Then
			CheckStr = True
		End If
	Case 3
		If InStr(textStr," ") Then Exit Function
		If Len(textStr) < 2 Then Exit Function
		If LCase$(textStr) = textStr Then Exit Function
		If UCase$(textStr) = textStr Then Exit Function
		If (textStr Like "*[!" & Replace$(Replace$(StrRange,"]|[",""),"[","") & "*") = True Then Exit Function
		textStr = Mid$(textStr,2)
		If LCase$(textStr) = textStr Then Exit Function
		CheckStr = True
	Case 4
		If StrNum < 2 Then Exit Function
		If Len(textStr) < StrNum Then Exit Function
		If InStr(textStr," ") Then Exit Function
		If IsNumeric(textStr) = True Then Exit Function
		If IsDate(textStr) = True Then Exit Function
		If InStr(textStr,"...") Then Exit Function
		If InStr(textStr,"://www.") Then Exit Function
		If IgnoreCase = True Then
			textStr = LCase$(textStr)
			StrRange = LCase$(StrRange)
		End If
		If (textStr Like "*" & Replace$(StrRange,"]|[","") & "*") = False Then Exit Function
		For i = 1 To Len(textStr)
			Temp = Mid$(textStr,i,1)
			If IsNumeric(Temp) = False Then
				If Temp <> "." Then
					If InStr(textStr,String$(StrNum,Temp)) Then
						CheckStr = True
						Exit For
					End If
				End If
			End If
		Next i
	End Select
End Function


'检查字串是否包含指定字符(正则表达式比较)
'Mode = 0 检查字串是否包含指定字符，并找出指定字符的位置
'Mode = 1 检查字串是否只包含指定字符
'Mode = 2 检查字串是否包含指定字符
'Mode = 3 检查字串是否只包含大小混写的指定字符，此时 IgnoreCase 参数无效
'Mode = 4 检查字串是否有连续相同的字符，StrNum 为最少重复字符个数
'Mode = 5 检查字串是否包含指定字串，并返回匹配的字串总长度 (适合字符组合查询)
'Patrn  为正则表达式模板
Public Function CheckStrRegExp(ByVal textStr As String,ByVal Patrn As String,Optional ByVal StrNum As Long, _
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
			If InStr(textStr,"...") Then Exit Function
			If InStr(textStr,"://www.") Then Exit Function
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


'字串常数反向转换
'Mode = 0 按 PSL 版本的宏引擎不同分别转义控制字符和全部或部分拉丁文扩展字符
'Mode = 1 转义控制字符和所有拉丁文扩展字符
'Mode = 2 转义控制字符和操作系统不能显示的拉丁文扩展字符
'Mode = 3 仅转义控制字符
Public Function ReConvert(ByVal ConverString As String,Optional ByVal Mode As Integer) As String
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


'字串常数正向转换
Public Function Convert(ByVal ConverString As String) As String
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
	If InStr(Convert,"\") Then Convert = ConvertBRegExp(Convert)
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
			If CheckStr(UCase$(ConverString),"[0-9A-Fa-f]",0,1) = True Then
				j = Val("&H" & ConverString)
				ConvertB = Replace$(ConvertB,EscStr & ConverString,Val2Bytes(j,2))
			End If
		Case "\u", "\U"
			ConverString = Mid$(ConvertB,i + 2,4)
			If CheckStr(UCase$(ConverString),"[0-9A-Fa-f]",0,1) = True Then
				j = Val("&H" & ConverString)
				ConvertB = Replace$(ConvertB,EscStr & ConverString,Val2Bytes(j,2))
			End If
		Case Is <> ""
			EscStr = "\"
			For j = 3 To 1 Step -1
				ConverString = Mid$(ConvertB,i + 1,j)
				If CheckStr(ConverString,"[0-7]",0,1) = True Then
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


'转换八进制或十六进制转义符
Private Function ConvertBRegExp(ByVal ConverString As String) As String
	Dim i As Long,j As Long,CodeVal As Long,Matches As Object
	ConvertBRegExp = ConverString
	With RegExp
		.Global = True
		.IgnoreCase = True
		For i = 0 To UBound(CheckHexStr)
			.Pattern = CheckHexStr(i).Range
			Set Matches = .Execute(ConverString)
			If Matches.Count > 0 Then
				For j = 0 To Matches.Count - 1
					If i = 0 Then
						If Matches(j).Length = 4 Then
							CodeVal = Val("&H" & Mid$(Matches(j).Value,3))
							ConvertBRegExp = Replace$(ConvertBRegExp,Matches(j).Value,Val2Bytes(CodeVal,2))
						End If
					ElseIf i = 1 Then
						If Matches(j).Length = 6 Then
							CodeVal = Val("&H" & Mid$(Matches(j).Value,3))
							ConvertBRegExp = Replace$(ConvertBRegExp,Matches(j).Value,Val2Bytes(CodeVal,2))
						End If
					ElseIf Matches(j).Length > 1 And Matches(j).Length < 5 Then
						CodeVal = Val("&O" & Replace$(Matches(j).Value,"\",""))
						If CodeVal > 256 Then
							Matches(j).Value = Left$(Matches(j).Value,3)
							CodeVal = Val("&O" & Replace$(Matches(j).Value,"\",""))
						End If
						ConvertBRegExp = Replace$(ConvertBRegExp,Matches(j).Value,Val2Bytes(CodeVal,2))
					End If
				Next j
			End If
		Next i
	End With
End Function


'转换 Athena-A 转义符为标准格式
Public Function AthenaA(ByVal ConverString As String) As String
	AthenaA = ConverString
	If AthenaA = "" Then Exit Function
	If InStr(AthenaA,"\") = 0 Then Exit Function
	If InStr(AthenaA,"\") Then AthenaA = Replace$(AthenaA,"\","\\")
	If InStr(AthenaA,"[\\\\]") Then AthenaA = Replace$(AthenaA,"[\\\\]","\\")
	If InStr(AthenaA,"[\\r]") Then AthenaA = Replace$(AthenaA,"[\\r]","\r")
	If InStr(AthenaA,"[\\n]") Then AthenaA = Replace$(AthenaA,"[\\n]","\n")
	If InStr(AthenaA,"[\\b]") Then AthenaA = Replace$(AthenaA,"[\\b]","\b")
	If InStr(AthenaA,"[\\f]") Then AthenaA = Replace$(AthenaA,"[\\f]","\f")
	If InStr(AthenaA,"[\\v]") Then AthenaA = Replace$(AthenaA,"[\\v]","\v")
	If InStr(AthenaA,"[\\t]") Then AthenaA = Replace$(AthenaA,"[\\t]","\t")
	If InStr(AthenaA,"[\\']") Then AthenaA = Replace$(AthenaA,"[\\']","\'")
	If InStr(AthenaA,"[\\""]") Then AthenaA = Replace$(AthenaA,"[\\""]","\""")
	If InStr(AthenaA,"[\\?]") Then AthenaA = Replace$(AthenaA,"[\\?]","\?")
	If InStr(AthenaA,"[\\0]") Then AthenaA = Replace$(AthenaA,"[\\0]","\0")
	If InStr(AthenaA,"[\\") Then AthenaA = AthenaARegExp(AthenaA)
End Function


'转换 Athena-A 八进制或十六进制转义符为标准格式
Private Function AthenaARegExp(ByVal ConverString As String) As String
	Dim i As Long,j As Long,Temp As String,Matches As Object
	AthenaARegExp = ConverString
	With RegExp
		.Global = True
		.IgnoreCase = True
		For i = 0 To UBound(CheckHexStr)
			.Pattern = "\[\\" & CheckHexStr(i).Range & "]"
			Set Matches = .Execute(ConverString)
			If Matches.Count > 0 Then
				For j = 0 To Matches.Count - 1
					Temp = Replace$(Replace$(Matches(j).Value,"[\",""),"]","")
					AthenaARegExp = Replace$(AthenaARegExp,Matches(j).Value,Temp)
				Next j
			End If
		Next i
	End With
End Function


'转换字符为 Long 整数值
Public Function StrToLong(ByVal mStr As String,Optional ByVal DefaultValue As Long) As Long
	On Error GoTo errHandle
	StrToLong = CLng(mStr)
	Exit Function
	errHandle:
	StrToLong = DefaultValue
End Function


'除去字串前后指定的 PreStr 和 AppStr
'fType = -1 不去除字串前后的空格和所有指定的 PreStr 和 AppStr，但不去除字串内前后空格
'fType = 0 去除字串前后的空格和所有指定的 PreStr 和 AppStr，但不去除字串内前后空格
'fType = 1 去除字串前后的空格和所有指定的 PreStr 和 AppStr，并去除字串内前后空格
'fType = 2 去除字串前后的空格和指定的 PreStr 和 AppStr 1 次，但不去除字串内前后空格
'fType > 2 去除字串前后的空格和指定的 PreStr 和 AppStr 1 次，并去除字串内前后空格
Public Function RemoveBackslash(ByVal Path As String,ByVal PreStr As String,ByVal AppStr As String,ByVal fType As Long) As String
	Dim i As Long,a As Long,p As Long,Stemp As Boolean
	RemoveBackslash = Path
	If Path = "" Then Exit Function
	a = Len(AppStr)
	p = Len(PreStr)
	If fType > -1 Then RemoveBackslash = Trim(RemoveBackslash)
	Do
		Stemp = False
		If p <> 0 Then
			If Left$(RemoveBackslash,p) = PreStr Then
				RemoveBackslash = Mid$(RemoveBackslash,p + 1)
				Stemp = True
			End If
		End If
		If a <> 0 Then
			If Right$(RemoveBackslash,a) = AppStr Then
				RemoveBackslash = Left$(RemoveBackslash,Len(RemoveBackslash) - a)
				Stemp = True
			End If
		End If
		If fType = 1 Or fType > 2 Then RemoveBackslash = Trim$(RemoveBackslash)
		If Stemp = True Then
			If fType < 2 Then i = 0 Else i = 1
		Else
			i = 1
		End If
	Loop Until i = 1
End Function


'字串前后附加指定的 PreStr 和 AppStr
'fType = 0 不去除字串前后空格，但在字串前后附加指定的 PreStr 和 AppStr
'fType = 1 去除字串前后空格，并在字串前后附加指定的 PreStr 和 AppStr
Public Function AppendBackslash(ByVal Path As String,ByVal PreStr As String,ByVal AppStr As String,ByVal fType As Long) As String
	AppendBackslash = Path
	If fType = 1 Then AppendBackslash = Trim$(AppendBackslash)
	If AppendBackslash = "" And PreStr = AppStr Then
		AppendBackslash = PreStr & AppendBackslash & AppStr
	Else
		If PreStr <> "" Then
			If Left$(AppendBackslash,Len(PreStr)) <> PreStr Then
				AppendBackslash = PreStr & AppendBackslash
			End If
		End If
		If AppStr <> "" Then
			If Right$(AppendBackslash,Len(AppStr)) <> AppStr Then
				AppendBackslash = AppendBackslash & AppStr
			End If
		End If
	End If
End Function


'获取4个字节值 (32 值, 4个字节)
Private Function GetDWord(Source As Variant,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Double
	GetDWord# = GetWord(Source, Offset, Mode)
	GetDWord# = GetDWord# + 65536# * GetWord(Source, Offset + 2, Mode)
End Function


'获取2个字节值 (16 值, 2个字节)
Private Function GetWord(Source As Variant,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Long
	GetWord& = GetByte(Source, Offset, Mode)
	GetWord& = GetWord& + 256& * GetByte(Source, Offset + 1,Mode)
End Function


'获取8个字节值 (64 位值,8个字节)
Private Function GetDouble(Source As Variant,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Double
	On Error GoTo errHandle
	If Offset + 8 > Source.SizeOfImage Then GoTo errHandle
	Select Case Mode
	Case Is < 0
		If Offset < 0 Then
			Get #Source.hFile, , GetDouble
		Else
			Get #Source.hFile, Offset + 1, GetDouble
		End If
	Case 0
		If Offset < 0 Then GoTo errHandle
		CopyMemory GetDouble, Source.ImageByte(Offset), 8
	Case Else
		If Offset < 0 Then GoTo errHandle
		MoveMemory GetDouble, Source.MappedAddress + Offset, 8
	End Select
	errHandle:
End Function


'获取4个字节值 (32 位值,4个字节, -2,147,483,648 到 2,147,483,647)
Public Function GetLong(Source As Variant,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Long
	On Error GoTo errHandle
	If Offset + 4 > Source.SizeOfImage Then GoTo errHandle
	Select Case Mode
	Case Is < 0
		If Offset < 0 Then
			Get #Source.hFile, , GetLong
		Else
			Get #Source.hFile, Offset + 1, GetLong
		End If
	Case 0
		If Offset < 0 Then GoTo errHandle
		CopyMemory GetLong, Source.ImageByte(Offset), 4
	Case Else
		If Offset < 0 Then GoTo errHandle
		MoveMemory GetLong, Source.MappedAddress + Offset, 4
	End Select
	errHandle:
End Function


'获取2个字节值 (16 位值, 2个字节, -32,768 到 32,767)
Public Function GetInteger(Source As Variant,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Integer
	On Error GoTo errHandle
	If Offset + 2 > Source.SizeOfImage Then GoTo errHandle
	Select Case Mode
	Case Is < 0
		If Offset < 0 Then
			Get #Source.hFile, , GetInteger
		Else
			Get #Source.hFile, Offset + 1, GetInteger
		End If
	Case 0
		If Offset < 0 Then GoTo errHandle
		CopyMemory GetInteger, Source.ImageByte(Offset), 2
	Case Else
		If Offset < 0 Then GoTo errHandle
		MoveMemory GetInteger, Source.MappedAddress + Offset, 2
	End Select
	errHandle:
End Function


'按指定地址获取一个字节(8 位值, 1个字节)
Public Function GetByte(Source As Variant,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Byte
	On Error GoTo errHandle
	If Offset + 1 > Source.SizeOfImage Then GoTo errHandle
	Select Case Mode
	Case Is < 0
		If Offset < 0 Then
			Get #Source.hFile, , GetByte
		Else
			Get #Source.hFile, Offset + 1, GetByte
		End If
	Case 0
		If Offset < 0 Then GoTo errHandle
		GetByte = Source.ImageByte(Offset)
	Case Else
		If Offset < 0 Then GoTo errHandle
		MoveMemory GetByte, Source.MappedAddress + Offset, 1
	End Select
	errHandle:
End Function


'获取区间内的字节数组
Public Function GetBytes(Source As Variant,ByVal Length As Long,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Byte()
	On Error GoTo errHandle
	If Offset + Length > Source.SizeOfImage Then
		Length = Source.SizeOfImage - Offset
		If Length < 1 Then GoTo errHandle
	End If
	ReDim Bytes(Length - 1) As Byte
	Select Case Mode
	Case Is < 0
		If Offset < 0 Then
			Get #Source.hFile, , Bytes
		Else
			Get #Source.hFile, Offset + 1, Bytes
		End If
	Case 0
		If Offset < 0 Then GoTo errHandle
		CopyMemory Bytes(0), Source.ImageByte(Offset), Length
	Case Else
		If Offset < 0 Then GoTo errHandle
		MoveMemory Bytes(0), Source.MappedAddress + Offset, Length
	End Select
	GetBytes = Bytes
	Exit Function
	errHandle:
	ReDim Bytes(0) As Byte
	GetBytes = Bytes
End Function


'获取字节数组(内存映射方式)
'Private Function getBytesByMap(ByVal Source As Long,ByVal Length As Long) As Byte()
'	Dim ppSA As Long, pSA As Long
'	Dim tagNewSA As SAFEARRAYID, tagOldSA As SAFEARRAYID
'	ReDim Bytes(0) As Byte							'初始化数组
'	ppSA = VarPtr(Bytes(0))							'获得指向SAFEARRAY的指针的指针
'	MoveMemory pSA, ppSA, 4							'获得指向SAFEARRAY的指针
'	MoveMemory tagOldSA, pSA, Len(tagOldSA)			'保存原来的SAFEARRAY成员信息
'	CopyMemory tagNewSA, tagOldSA, Len(tagNewSA)	'复制SAFEARRAY成员信息
'	tagNewSA.rgsabound(0).cElements = Length		'修改数组元素个数
'	tagNewSA.pvData = Source						'修改数组数据地址
'	WriteMemory pSA, tagNewSA, Len(tagNewSA)		'将映射后的数据地址绑定至数组
'	getBytesByMap = Bytes
'	WriteMemory pSA, tagOldSA, Len(tagOldSA)		'恢复数组的SAFEARRAY结构成员信息
'End Function


'获取自定义类型值
Private Function GetTypeValue(Source As FILE_IMAGE,ByVal Offset As Long,Target As Variant,Optional ByVal Mode As Long = -1) As Boolean
	On Error GoTo errHandle
	If Offset < 0 Then GoTo errHandle
	If Offset +  Len(Target) > Source.SizeOfImage Then GoTo errHandle
	Select Case Mode
	Case Is < 0
		Get #Source.hFile, Offset + 1, Target
	Case 0
		CopyMemory Target, Source.ImageByte(Offset), Len(Target)
	Case Else
		MoveMemory Target, Source.MappedAddress + Offset, Len(Target)
	End Select
	GetTypeValue = True
	errHandle:
End Function


'获取自定义类型数组
Private Function GetTypeArray(Source As FILE_IMAGE,ByVal Offset As Long,Target As Variant,Optional ByVal Mode As Long = -1) As Boolean
	On Error GoTo errHandle
	If Offset < 0 Then GoTo errHandle
	If Offset + Len(Target(0)) * (UBound(Target) + 1) > Source.SizeOfImage Then GoTo errHandle
	Select Case Mode
	Case Is < 0
		Get #Source.hFile, Offset + 1, Target
	Case 0
		CopyMemory Target(0), Source.ImageByte(Offset), Len(Target(0)) * (UBound(Target) + 1)
	Case Else
		MoveMemory Target(0), Source.MappedAddress + Offset, Len(Target(0)) * (UBound(Target) + 1)
	End Select
	GetTypeArray = True
	errHandle:
End Function


'写入字节数组
Public Function PutBytes(Target As FILE_IMAGE,ByVal Offset As Long,Source() As Byte,ByVal Length As Long,Optional ByVal Mode As Long = -1) As Boolean
	'On Error GoTo errHandle
	Select Case Mode
	Case Is < 0
		If Offset < 0 Then
			Put #Target.hFile, , Source
		Else
			Put #Target.hFile, Offset + 1, Source
		End If
	Case 0
		If (Offset + Length) > Target.SizeOfFile Then
			Target.SizeOfFile = Offset + Length
			If Target.SizeOfFile > Target.SizeOfImage Then
				Target.SizeOfImage = Target.SizeOfFile + 2048
				ReDim Preserve Target.ImageByte(Target.SizeOfImage - 1) 'As Byte
			End If
		End If
		CopyMemory Target.ImageByte(Offset), Source(0), Length
	Case Else
		If (Offset + Length) > Target.SizeOfFile Then
			Target.SizeOfFile = Offset + Length
			If Target.SizeOfFile > Target.SizeOfImage Then
				If MapFile(Target.ModuleName,Target,Target.SizeOfFile + 2048,1,Target.SizeOfFile) = False Then
					Exit Function
				End If
			End If
		End If
		WriteMemory Target.MappedAddress + Offset, Source(0), Length
	End Select
	PutBytes = True
	errHandle:
End Function


'写入自定义类型值
Private Function PutTypeValue(Target As FILE_IMAGE,ByVal Offset As Long,Source As Variant,Optional ByVal Mode As Long = -1) As Boolean
	'On Error GoTo errHandle
	Select Case Mode
	Case Is < 0
		If Offset < 0 Then
			Put #Target.hFile, , Source
		Else
			Put #Target.hFile, Offset + 1, Source
		End If
	Case 0
		If (Offset + Len(Source)) > Target.SizeOfFile Then
			Target.SizeOfFile = Offset + Len(Source)
			If Target.SizeOfFile > Target.SizeOfImage Then
				Target.SizeOfImage = Target.SizeOfFile + 2048
				ReDim Preserve Target.ImageByte(Target.SizeOfImage - 1) 'As Byte
			End If
		End If
		CopyMemory Target.ImageByte(Offset), Source, Len(Source)
	Case Else
		If (Offset + Len(Source)) > Target.SizeOfFile Then
			Target.SizeOfFile = Offset + Len(Source)
			If Target.SizeOfFile > Target.SizeOfImage Then
				If MapFile(Target.ModuleName,Target,Target.SizeOfFile + 2048,1,Target.SizeOfFile) = False Then
					Exit Function
				End If
			End If
		End If
		WriteMemory Target.MappedAddress + Offset, Source, Len(Source)
	End Select
	PutTypeValue = True
	errHandle:
End Function


'写入自定义类型数组
Private Function PutTypeArray(Target As FILE_IMAGE,ByVal Offset As Long,Source As Variant,Optional ByVal Mode As Long = -1) As Boolean
	'On Error GoTo errHandle
	Select Case Mode
	Case Is < 0
		If Offset < 0 Then
			Put #Target.hFile, , Source
		Else
			Put #Target.hFile, Offset + 1, Source
		End If
	Case 0
		If (Offset + Len(Source(0)) * (UBound(Source) + 1)) > Target.SizeOfFile Then
			Target.SizeOfFile = Offset + Len(Source(0)) * (UBound(Source) + 1)
			If Target.SizeOfFile > Target.SizeOfImage Then
				Target.SizeOfImage = Target.SizeOfFile + 2048
				ReDim Preserve Target.ImageByte(Target.SizeOfImage - 1) 'As Byte
			End If
		End If
		CopyMemory Target.ImageByte(Offset), Source(0), Len(Source(0)) * (UBound(Source) + 1)
	Case Else
		If (Offset + Len(Source(0)) * (UBound(Source) + 1)) > Target.SizeOfFile Then
			Target.SizeOfFile = Offset + Len(Source(0)) * (UBound(Source) + 1)
			If Target.SizeOfFile > Target.SizeOfImage Then
				If MapFile(Target.ModuleName,Target,Target.SizeOfFile + 2048,1,Target.SizeOfFile) = False Then
					Exit Function
				End If
			End If
		End If
		WriteMemory Target.MappedAddress + Offset, Source(0), Len(Source(0)) * (UBound(Source) + 1)
	End Select
	PutTypeArray = True
	errHandle:
End Function


'获取变量字节长度
Public Function GetFileLength(Source As FILE_IMAGE,Optional ByVal Mode As Long = -1) As Long
	Select Case Mode
	Case Is < 0
		GetFileLength = LOF(Source.hFile)
	Case 0
		GetFileLength = Source.SizeOfFile	'UBound(Source.ImageByte) + 1
	Case Else
		GetFileLength = Source.SizeOfFile	'GetFileSize(Source.hFile, 0&)
	End Select
End Function


'创建子文件夹
Public Function MkSubDir(ByVal DirPath As String) As Boolean
	Dim i As Long,Temp As String
	On Error GoTo ErrorHandle
	DirPath = Trim$(DirPath)
	If DirPath = "" Then Exit Function
	If Right$(DirPath,1) <> "\" Then DirPath = DirPath & "\"
	If Dir$(DirPath & "*.*",vbDirectory) = "" Then
		i = InStr(DirPath,"\\")
		If i > 0 Then
			i = InStr(i + 2,DirPath,"\")
			If i > 0 Then i = InStr(i + 1,DirPath,"\")
		Else
			i = InStr(DirPath,"\")
		End If
		Do While i > 0
			Temp = Left$(DirPath,i)
			If Dir$(Temp & "*.*",vbDirectory) = "" Then
				MkDir Temp
			End If
			i = InStr(i + 1,DirPath,"\")
		Loop
	End If
	MkSubDir = True
	ErrorHandle:
End Function


'创建文本文件
Public Function CreateTXTFile(ByVal FilePath As String,ByVal Text As String,ByVal Code As String) As Boolean
	Dim i As Long
	i = InStrRev(FilePath,"\")
	If i > 0 Then
		If MkSubDir(Left$(FilePath,i)) = False Then Exit Function
		CreateTXTFile = WriteToFile(FilePath,Text,Code)
	End If
End Function


'获取当前文件夹中的文件列表
'Mode = True 无论是否存在 DefaultFile 都将作为存在添加为首个文件
Public Function GetFiles(List() As FILE_LIST,ByVal DirPath As String,Optional ByVal DefaultFile As String, _
		Optional ByVal FindFile As String,Optional ByVal Mode As Boolean) As String()
	Dim n As Long,m As Long,FileList() As String,File As String,TempList() As String
	DirPath = Trim$(DirPath)
	If DirPath = "" Or (DefaultFile = "" And FindFile = "") Then
		ReDim List(0) As FILE_LIST,FileList(0) As String
		GetFiles = FileList
		Exit Function
	End If
	If Right$(DirPath,1) <> "\" Then DirPath = DirPath & "\"
	If DefaultFile <> "" Then
		DefaultFile = Mid$(DefaultFile,InStrRev(DefaultFile,"\") + 1)
	End If
	m = 20
	ReDim List(m) As FILE_LIST,FileList(m) As String
	If Mode = True Then
		List(0).sName = DefaultFile
		List(0).FilePath = DirPath & DefaultFile
		FileList(0) = DefaultFile
		n = 1
	End If
	DefaultFile = LCase$(DefaultFile): FindFile = LCase$(FindFile)
	File = Dir$(DirPath & "*.*")
	Do While File <> ""
		If n > m Then
			m = m * 2
			ReDim Preserve List(m) As FILE_LIST,FileList(m) As String
		End If
		If DefaultFile <> "" Then
			If LCase$(File) Like DefaultFile Then
				If Mode = True Then GoTo NextNum
				List(n).sName = File
				List(n).FilePath = DirPath & File
				FileList(n) = File
				n = n + 1
				GoTo NextNum
			End If
		End If
		If FindFile <> "" Then
			If LCase$(File) Like FindFile Then
				List(n).sName = File
				List(n).FilePath = DirPath & File
				FileList(n) = File
				n = n + 1
				GoTo NextNum
			End If
		End If
		NextNum:
		File = Dir$()
	Loop
	If n > 0 Then n = n - 1
	ReDim Preserve List(n) As FILE_LIST,FileList(n) As String
	GetFiles = FileList
End Function


'获取当前文件夹的子文件夹中的文件列表
'FindFile = "" 时 .sName = 文件所在子目录名，否则 .sName = 子目录中的文件
Public Function getSubFiles(List() As FILE_LIST,ByVal DirPath As String,Optional ByVal DefaultFile As String,Optional ByVal FindFile As String) As String()
	Dim i As Long,j As Long,n As Long,m As Long
	Dim File As String,PathList() As String,FileList() As String,TempList() As String
	DirPath = Trim$(DirPath)
	If DirPath = "" Or (DefaultFile = "" And FindFile = "") Then
		ReDim List(0) As FILE_LIST,FileList(0) As String
		getSubFiles = FileList
		Exit Function
	End If
	If Right$(DirPath,1) <> "\" Then DirPath = DirPath & "\"
	If Dir$(DirPath & "*.*",vbDirectory) = "" Then
		ReDim List(0) As FILE_LIST,FileList(0) As String
		Exit Function
	End If
	If DefaultFile <> "" Then
		DefaultFile = LCase$(Mid$(DefaultFile,InStrRev(DefaultFile,"\") + 1))
	End If
	If FindFile <> "" Then FindFile = LCase$(FindFile)
	ReDim PathList(0) As String
	PathList(0) = DirPath
	m = 20
	ReDim List(m) As FILE_LIST,FileList(m) As String
	Do
		DirPath = PathList(i)
		File = Dir$(DirPath & "*.*",vbDirectory)
		While File <> ""
			If File <> "." And File <> ".." Then
				If GetAttr(DirPath & File) And vbDirectory Then
					j = j + 1
					ReDim Preserve PathList(j) As String
					PathList(j) = DirPath & File & "\"
				ElseIf DirPath <> PathList(0) Then
					If n > m Then
						m = m * 2
						ReDim Preserve List(m) As FILE_LIST,FileList(m) As String
					End If
					If DefaultFile <> "" Then
						If LCase$(File) Like DefaultFile Then
							If FindFile = "" Then
								TempList = ReSplit(DirPath,"\")
								List(n).sName = TempList(UBound(TempList) - 1)
							Else
								List(n).sName = File
							End If
							List(n).FilePath = DirPath & File
							FileList(n) = List(n).sName
							n = n + 1
							GoTo NextNum
						End If
					End If
					If FindFile <> "" Then
						If LCase$(File) Like FindFile Then
							List(n).sName = File
							List(n).FilePath = DirPath & File
							FileList(n) = List(n).sName
							n = n + 1
							GoTo NextNum
						End If
					End If
				End If
				NextNum:
			End If
			File = Dir$()
		Wend
		i = i + 1
	Loop Until i = j + 1
	If n > 0 Then n = n - 1
	ReDim Preserve List(n) As FILE_LIST,FileList(n) As String
	getSubFiles = FileList
End Function


'删除文件夹，不会删除子文件夹
Public Function DelDir(ByVal DirPath As String) As Boolean
	Dim File As String
	DirPath = Trim$(DirPath)
	If DirPath = "" Then Exit Function
	If Right$(DirPath,1) <> "\" Then DirPath = DirPath & "\"
	File = Dir$(DirPath & "*.*")
	On Error GoTo errHandle
	Do While File <> ""
		Kill DirPath & File
		File = Dir$()
	Loop
	If Dir$(DirPath & "*.*") = "" Then
		RmDir DirPath
		DelDir = True
	End If
	errHandle:
End Function


'删除文件夹，包括所有子文件夹
Public Function DelDirs(ByVal DirPath As String) As Boolean
	Dim i As Long,j As Long,File As String
	DirPath = Trim$(DirPath)
	If DirPath = "" Then Exit Function
	If Right$(DirPath,1) <> "\" Then DirPath = DirPath & "\"
	If Dir$(DirPath & "*.*",vbDirectory) = "" Then Exit Function
	On Error Resume Next
	ReDim PathList(0) As String
	PathList(0) = DirPath
	Do
		DirPath = PathList(i)
		File = Dir$(DirPath & "*.*",vbDirectory)
		While File <> ""
			If File <> "." And File <> ".." Then
				If GetAttr(DirPath & File) And vbDirectory Then
					j = j + 1
					ReDim Preserve PathList(j) As String
					PathList(j) = DirPath & File & "\"
				Else
					Kill DirPath & File
				End If
			End If
			File = Dir$()
		Wend
		RmDir PathList(j)
		i = i + 1
	Loop Until i = j + 1
	For i = j To 0 Step -1
		If Dir$(PathList(i) & "*.*") = "" Then RmDir PathList(i)
	Next i
	If Dir$(PathList(0) & "*.*",vbDirectory) = "" Then DelDirs = True
End Function


'字节数组转正则表达式使用的转义符模板
'Mode = 0 转为有 [] 形式，否则为无 [] 形式
Public Function Byte2RegExpPattern(Bytes() As Byte,Optional ByVal Mode As Long,Optional ByVal CodePage As Long) As String
	If Mode = 0 Then
		Byte2RegExpPattern = "[" & Byte2HexEsc(Bytes,0,-1,CodePage) & "]"
	Else
		Byte2RegExpPattern = Byte2HexEsc(Bytes,0,-1,CodePage)
	End If
End Function


'字符串转正则表达式使用的 Unicode 转义符模板
'Mode = 0 转为有 [] 形式，否则为无 [] 形式
Public Function Str2RegExpPattern(ByVal textStr As String,Optional ByVal Mode As Long) As String
	Dim i As Long,j As Long,n As Long,Bytes() As Byte
	If textStr = "" Then Exit Function
	Bytes = textStr
	j = UBound(Bytes)
	If Mode = 0 Then
		Str2RegExpPattern = "[" & Space$((j + 1) * 3) & "]"
		n = 2
	Else
		Str2RegExpPattern = Space$((j + 1) * 3)
		n = 1
	End If
	For i = 0 To j - 1 Step 2
		Mid$(Str2RegExpPattern,n,6) = "\u" & Right$("0" & Hex$(Bytes(i + 1)),2) & Right$("0" & Hex$(Bytes(i)),2)
		n = n + 6
	Next i
End Function


'Hex 字符串转正则表达式使用的 Hex 转义符模板
'Mode = 0 转为有 [] 形式，否则为无 [] 形式
Public Function HexStr2RegExpPattern(ByVal HexStr As String,Optional ByVal Mode As Long) As String
	Dim i As Long,j As Long,n As Long
	j = Len(HexStr)
	If j = 0 Then Exit Function
	If Mode = 0 Then
		HexStr2RegExpPattern = "[" & Space$(j * 2) & "]"
		n = 2
	Else
		HexStr2RegExpPattern = Space$(j * 2)
		n = 1
	End If
	For i = 1 To j Step 2
		Mid$(HexStr2RegExpPattern,n,4) = "\x" & Mid$(HexStr,i,2)
		n = n + 4
	Next i
End Function


'获取偶数位
'Mode = 0 奇数加 1 个字节，Mode = 1 奇数减 1 个字节
Public Function GetEvenPos(ByVal Pos As Long,Optional ByVal Mode As Long) As Long
	If Pos And 1 Then
		If Mode = 0 Then
			GetEvenPos = Pos + 1
		Else
			GetEvenPos = Pos - 1
		End If
	Else
		GetEvenPos = Pos
	End If
End Function


'字符串转字节数组
Public Function StringToByte(ByVal textStr As String,ByVal CodePage As Long) As Byte()
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
Public Function ByteToString(Bytes() As Byte,ByVal CodePage As Long) As String
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
Public Function LowByte2HighByte(Bytes() As Byte,ByVal Setp As Integer) As Byte()
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
Public Function ReverseValByte(Bytes() As Byte,ByVal StartPos As Long,ByVal endPos As Long) As Byte()
	Dim i As Long,Temp() As Byte
	If StartPos < 0 Then StartPos = LBound(Bytes)
	If endPos < 0 Then endPos = UBound(Bytes)
	Temp = Bytes
	For i = StartPos To endPos
		Temp(i) = Bytes(endPos - i)
	Next i
	ReverseValByte = Temp
End Function


'按字符编码计算字符的十六进制字节长度
'Mode = 1 转义, 否则不转义
Public Function StrHexLength(ByVal textStr As String,ByVal CodePage As Long,ByVal Mode As Long) As Long
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


'获取字符串的截取位置，避免截取半个双字节字符
'返回 IsDBCSLeadPos = 截取位置(截取后的字节长度)
'Mode = False 不返回截取后的字串, 否则返回截取后的字串
'fType = False 后截取，否则前截取
'ByteLength 为需要的字节长度
Public Function IsDBCSLeadPos(textStr As String,ByVal CodePage As Long,ByVal ByteLength As Long,Optional ByVal Mode As Boolean,Optional ByVal fType As Boolean) As Long
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
Public Function FillStrWithSpape(textStr As String,ByVal CodePage As Long,ByVal ByteLength As Long,Optional ByVal Mode As Boolean,Optional ByVal fType As Boolean) As Long
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


'顺序在字节数组中查找匹配数组
'Mode = 0，查找 StartPos 到 EndPos 之间的位置(包含 StartPos 和 EndPos)
'Mode <> 0，查找除 StartPos 到 EndPos 之间以外的位置(不包含 StartPos 和 EndPos)
'返回 InByte 值 = 0，没有找到，> 0 找到 (以 1 开始的地址)
'注意：StartPos、EndPos 和 MaxPos 均为以 0 开始的地址
Private Function InByte(Bytes() As Byte,Find() As Byte,Optional ByVal StartPos As Long,Optional ByVal EndPos As Long, _
		Optional ByVal MaxPos As Long,Optional Mode As Long) As Long
	Dim i As Long,j As Long,k As Long,Length As Long,tempByte() As Byte
	Length = UBound(Find) + 1
	tempByte = Find
	k = -1
	For i = 0 To Length - 1
		If Find(i) <> 0 Then
			If i <> 0 Then
				ReDim tempByte(Length - i - 1) As Byte
				CopyMemory tempByte(0), Find(i), Length - i
			End If
			k = i
			Exit For
		End If
	Next i
	If k = -1 Then
		If Mode = 0 Then
			InByte = NullInByteRegExp(Bytes,StartPos,endPos,Length)
		Else
			InByte = NullInByteRegExp(Bytes,0,StartPos - 1,Length)
			If InByte = 0 Then
				InByte = NullInByteRegExp(Bytes,EndPos + 1,MaxPos,Length)
			End If
		End If
		Exit Function
	End If
	If Mode = 0 Then
		i = InStrB(StartPos + k + 1,Bytes,tempByte)
	Else
		i = InStrB(k + 1,Bytes,tempByte)
	End If
	Do While i > 0
		If k > 0 Then
			For j = i - k - 1 To i - 2
				If Bytes(j) <> 0 Then GoTo NextNum
			Next j
			i = i - k
		End If
		If Mode = 0 Then
			If EndPos > 0 Then
				If i + Length - 2 > EndPos Then Exit Do
			End If
			InByte = i	'注意 InStrB 函数找到第一个数就返回"1"
			Exit Do
		ElseIf (i + Length - 2 < StartPos) Or (i + Length - 2 > EndPos) Then
			If MaxPos > 0 Then
				If i + Length - 2 > MaxPos Then Exit Do
			End If
			InByte = i	'注意 InStrB 函数找到第一个数就返回"1"
			Exit Do
		End If
		NextNum:
		i = InStrB(i + Length,Bytes,tempByte)
	Loop
End Function


'顺序在字节数组中查找匹配数组
'Mode = 0，查找 StartPos 到 EndPos 之间的位置(包含 StartPos 和 EndPos)
'Mode <> 0，查找除 StartPos 到 EndPos 之间以外的位置(不包含 StartPos 和 EndPos)
'返回 InByte 值 = 0，没有找到，> 0 找到 (以 1 开始的地址)
'注意：StartPos、EndPos 和 MaxPos 均为以 0 开始的地址
Public Function InByteRegExp(Bytes() As Byte,Find() As Byte,Optional ByVal StartPos As Long,Optional ByVal endPos As Long, _
		Optional ByVal MaxPos As Long,Optional Mode As Long) As Long
	Dim i As Long,j As Long,Length As Long,tempByte() As Byte,Matches As Object
	On Error GoTo ExitFunction
	Length = UBound(Find) + 1
	If Mode = 0 Then
		If StartPos <> 0 And endPos <> 0 Then
			ReDim tempByte(endPos - StartPos) As Byte
			CopyMemory tempByte(0), Bytes(StartPos), endPos - StartPos + 1
		ElseIf StartPos <> 0 Then
			ReDim tempByte(UBound(Bytes) - StartPos) As Byte
			CopyMemory tempByte(0), Bytes(StartPos), UBound(Bytes) - StartPos + 1
		Else
			tempByte = Bytes
			If endPos <> 0 Then ReDim Preserve tempByte(endPos) As Byte
		End If
	Else
		tempByte = Bytes
		If MaxPos <> 0 Then ReDim Preserve tempByte(MaxPos) As Byte
	End If
	RegExp.Global = IIf(Mode = 0,False,True)
	RegExp.IgnoreCase = False
	RegExp.Pattern = Byte2RegExpPattern(Find,1,CP_ISOLATIN1)
	Set Matches = RegExp.Execute(ByteToString(tempByte,CP_ISOLATIN1))
	If Matches.Count = 0 Then Exit Function
	For j = 0 To Matches.Count - 1
		i = StartPos + Matches(j).FirstIndex + 1
		If Mode = 0 Then
			InByteRegExp = i	'注意找到第一个数就返回"1"
			Exit For
		ElseIf (i + Length - 2 < StartPos) Or (i + Length - 2 > endPos) Then
			InByteRegExp = i	'注意找到第一个数就返回"1"
			Exit For
		End If
	Next j
	ExitFunction:
End Function


'逆序在字节数组中查找匹配数组
'Mode = 0，查找 StartPos 到 EndPos 之间的位置(包含 StartPos 和 EndPos)
'Mode <> 0，查找除 StartPos 到 EndPos 之间以外的位置(不包含 StartPos 和 EndPos)
'返回 InByteRev 值 = 0，没有找到，> 0 找到 (以 1 开始的地址)
'注意：StartPos、EndPos 和 MaxPos 均为以 0 开始的地址
Private Function InByteRev(Bytes() As Byte,Find() As Byte,Optional ByVal StartPos As Long,Optional ByVal endPos As Long,Optional ByVal MaxPos As Long,Optional Mode As Long) As Long
	Dim i As Long,j As Long,k As Long,Length As Long,tempByte() As Byte
	Length = UBound(Find) + 1
	tempByte = Find
	k = -1
	For i = 0 To Length - 1
		If Find(i) <> 0 Then
			If i <> 0 Then
				ReDim tempByte(Length - i - 1) As Byte
				CopyMemory tempByte(0), Find(i), Length - i
			End If
			k = i
			Exit For
		End If
	Next i
	If k = -1 Then
		If Mode = 0 Then
			InByteRev = NullInByteRegExp(Bytes,StartPos,endPos,Length,1)
		Else
			InByteRev = NullInByteRegExp(Bytes,0,StartPos - 1,Length,1)
			If InByteRev = 0 Then
				InByteRev = NullInByteRegExp(Bytes,EndPos + 1,MaxPos,Length,1)
			End If
		End If
		Exit Function
	End If
	If Mode = 0 Then
		i = InStrB(StartPos + k + 1,Bytes,tempByte)
	Else
		i = InStrB(k + 1,Bytes,tempByte)
	End If
	Do While i > 0
		If k > 0 Then
			For j = i - k - 1 To i - 2
				If Bytes(j) <> 0 Then GoTo NextNum
			Next j
			i = i - k
		End If
		If Mode = 0 Then
			If EndPos > 0 Then
				If i + Length - 2 > EndPos Then Exit Do
			End If
			InByteRev = i	'注意 InStrB 函数找到第一个数就返回"1"
		ElseIf (i + Length - 2 < StartPos) Or (i + Length - 2 > EndPos) Then
			If MaxPos > 0 Then
				If i + Length - 2 > MaxPos Then Exit Do
			End If
			InByteRev = i	'注意 InStrB 函数找到第一个数就返回"1"
		End If
		NextNum:
		i = InStrB(i + Length,Bytes,tempByte)
	Loop
End Function


'逆序在字节数组中查找匹配数组
'Mode = 0，查找 StartPos 到 EndPos 之间的位置(包含 StartPos 和 EndPos)
'Mode <> 0，查找除 StartPos 到 EndPos 之间以外的位置(不包含 StartPos 和 EndPos)
'返回 InByteRev 值 = 0，没有找到，> 0 找到 (以 1 开始的地址)
'注意：StartPos、EndPos 和 MaxPos 均为以 0 开始的地址
Public Function InByteRevRegExp(Bytes() As Byte,Find() As Byte,Optional ByVal StartPos As Long,Optional ByVal EndPos As Long, _
		Optional ByVal MaxPos As Long,Optional Mode As Long) As Long
	Dim i As Long,j As Long,Length As Long,tempByte() As Byte,Matches As Object
	On Error GoTo ExitFunction
	Length = UBound(Find) + 1
	If Mode = 0 Then
		If StartPos <> 0 And EndPos <> 0 Then
			ReDim tempByte(EndPos - StartPos) As Byte
			CopyMemory tempByte(0), Bytes(StartPos), EndPos - StartPos + 1
		ElseIf StartPos <> 0 Then
			ReDim tempByte(UBound(Bytes) - StartPos) As Byte
			CopyMemory tempByte(0), Bytes(StartPos), UBound(Bytes) - StartPos + 1
		Else
			tempByte = Bytes
			If EndPos <> 0 Then ReDim Preserve tempByte(EndPos) As Byte
		End If
	Else
		tempByte = Bytes
		If MaxPos <> 0 Then ReDim Preserve tempByte(MaxPos) As Byte
	End If
	RegExp.Global = True
	RegExp.IgnoreCase = False
	RegExp.Pattern = Byte2RegExpPattern(Find,1,CP_ISOLATIN1)
	Set Matches = RegExp.Execute(ByteToString(tempByte,CP_ISOLATIN1))
	If Matches.Count = 0 Then Exit Function
	For j = Matches.Count - 1 To 0 Step -1
		i = StartPos + Matches(j).FirstIndex + 1
		If Mode = 0 Then
			InByteRevRegExp = i	'注意找到第一个数就返回"1"
			Exit For
		ElseIf (i + Length - 2 < StartPos) Or (i + Length - 2 > EndPos) Then
			InByteRevRegExp = i	'注意找到第一个数就返回"1"
			Exit For
		End If
	Next j
	ExitFunction:
End Function


'跳到指定数量的空字节位置，并返回空字节开始位置 (以 1 开始的地址)
'Mode = 0 返回首个开始位置，= 1 返回最后一个开始位置，= 2 返回首个结束位置，>= 3 返回最后一个结束位置
'返回 NullInByteRegExp 值 = 0，没有找到，> 0 找到 (以 1 开始的地址)
'注意：StartPos、EndPos 均为以 0 开始的地址
Public Function NullInByteRegExp(Bytes() As Byte,ByVal StartPos As Long,ByVal EndPos As Long, _
				Optional ByVal MinLength As Long = 1,Optional ByVal Mode As Long,Optional fType As Boolean) As Long
	Dim Matches As Object
	If EndPos <= 0 Then EndPos = UBound(Bytes)
	If EndPos - StartPos + 1 < MinLength Then Exit Function
	With RegExp
		Select Case Mode
		Case 0,2
			.Global = False
		Case Else
			.Global = True
		End Select
		.IgnoreCase = False
		.Pattern = "\x00{" & CStr$(MinLength) & ",}"
		If fType = False Then
			ReDim tempByte(endPos - StartPos) As Byte
			CopyMemory tempByte(0), Bytes(StartPos), EndPos - StartPos + 1
			Set Matches = .Execute(ByteToString(tempByte,CP_ISOLATIN1))
		Else
			Set Matches = .Execute(ByteToString(Bytes,CP_ISOLATIN1))
		End If
		If Matches.Count > 0 Then
			Select Case Mode
			Case 0
				NullInByteRegExp = StartPos + Matches(0).FirstIndex + 1
			Case 1
				NullInByteRegExp = StartPos + Matches(Matches.Count - 1).FirstIndex + 1
			Case 2
				NullInByteRegExp = StartPos + Matches(0).FirstIndex + Matches(0).Length
			Case Else
				NullInByteRegExp = StartPos + Matches(Matches.Count - 1).FirstIndex + Matches(Matches.Count - 1).Length
			End Select
		End If
	End With
End Function


'合并字节数组
Public Function MergeBytes(Bytes1() As Byte,Bytes2() As Byte) As Byte()
	Dim Length1 As Long, Length2 As Long
	On Error GoTo errHandle
	Length1 = UBound(Bytes1) + 1
	Length2 = UBound(Bytes2) + 1
	ReDim Bytes(Length1 + Length2 - 1) As Byte
	CopyMemory Bytes(0), Bytes1(0), Length1
	CopyMemory Bytes(Length1), Bytes2(0), Length2
	MergeBytes = Bytes
	Exit Function
	errHandle:
	ReDim Bytes(0) As Byte
	MergeBytes = Bytes
End Function


'设置光标位置（起始行和起始列为0）
'LineNo = 行号(文本框头开始算)，起始行 = 0
'ColNo = 列号(当前行首开始算)，起始列 = 0
'Length = 光标的选定长度
'strText = ""， ColNo 和 Length 为字符数，用于 Win7 及以下版本
'strText <> ""，ColNo 和 Length 为字节数，用于 Win10 及以上版本
'strText 为文本框中的文本，可以用 GetTextBoxString(hwnd) 来获取
Private Sub SetCurPos(ByVal hwnd As Long,ByVal LineNo As Long,ByVal ColNo As Long,ByVal Length As Long,Optional ByVal strText As String)
	'获取指定行的首字符在文本中的字符数偏移
	LineNo = SendMessageLNG(hwnd, EM_LINEINDEX, LineNo, 0&)
	'Win10 下 EM_SETSEL 需要字节数，而返回的 LineNo 为字符数，故需要转为字节数偏移
	If strText <> "" Then
		'LineNo = StrHexLength(Left$(strText,LineNo),GetACP,0) '返回行首的字节数
		ReDim tmpByte(0) As Byte
		tmpByte = StrConv(Left$(strText,LineNo),vbFromUnicode)
		LineNo = UBound(tmpByte) + 1
	End If
	'选定指定文本的整个范围 (Win10 下的参数为字节数偏移)
	SendMessageLNG hwnd, EM_SETSEL, LineNo + ColNo, LineNo + ColNo + Length
	'将选定内容放到可视范围之内
	SendMessageLNG hwnd, EM_SCROLLCARET, 0&, 0&
End Sub


'获取光标位置（行号和列号，起如行和起始列均为0）
'Mode = False，GetCurPos.y 为行号，GetCurPos.x 为光标终点的列号(行首开始算的字符数)
'Mode = True， GetCurPos.y 为行号，GetCurPos.x 为光标始点的列号(行首开始算的字符数)
'strText = ""，不转换 GetCurPos.x 为字符数，用于 Win7 及以下版本
'strText <> ""， 转换 GetCurPos.x 为字符数，用于 Win10 及以上版本
'strText 为文本框中的文本，可以用 GetTextBoxString(hwnd) 来获取
Private Function GetCurPos(ByVal hwnd As Long,ByVal Mode As Boolean,Optional ByVal strText As String) As POINTAPI
	If Mode = False Then
		'获取光标所在最后位置在文本中的字节数偏移
		SendMessage hwnd, EM_GETSEL, 0&, GetCurPos
	Else
		'获取光标所在位置在文本中的字节数偏移(低16位 = 光标始点,高16位 = 光标终点)，可用于逆序查找
		GetCurPos.x = SendMessageLNG(hwnd, EM_GETSEL, 0&, 0&)	'高位、低位的最大返回值为65535，否则返回 -1
		If GetCurPos.x > -1 Then
			'Int hi = DWORD / 0x10000; Int low = DWORD And 0xffff;
			If Mode = False Then
				GetCurPos.x = GetCurPos.x / &H10000	'高位，光标终点
			Else
				GetCurPos.x = GetCurPos.x And 65535	'低位，光标始点
			End If
		Else
			SendMessage hwnd, EM_GETSEL, 0&, GetCurPos
		End If
	End If
	If strText <> "" Then
		'转换 GetCurPos.x 为字符数，因为Win10 下 EM_LINEFROMCHAR 按字符数偏移获取行号
		ReDim tmpByte(0) As Byte
		'tmpByte = StringToByte(strText,GetACP)
		tmpByte = StrConv(strText,vbFromUnicode)
		ReDim Preserve tmpByte(GetCurPos.x + 1) As Byte
		'GetCurPos.x = Len(ByteToString(tmpByte,,GetACP))
		GetCurPos.x = Len(StrConv$(tmpByte,vbUnicode))
	End If
	'获得光标所在行的行号
	GetCurPos.y = SendMessageLNG(hwnd, EM_LINEFROMCHAR, GetCurPos.x, 0&)
	'返回光标所在行的字符数
	GetCurPos.x = GetCurPos.x - SendMessageLNG(hwnd, EM_LINEINDEX, GetCurPos.y, 0&)
End Function


'获取光标所在行的整行文本
Private Function GetCurPosLine(ByVal hwnd As Long,ByVal LineNo As Long) As String
	Dim Length As Long
	'获取光标所在行的首字符在文本中的字符数偏移
	Length = SendMessageLNG(hwnd, EM_LINEINDEX, LineNo, 0&)
	'获取光标所在行的文本长度(字符数)
	Length = SendMessageLNG(hwnd, EM_LINELENGTH, Length, 0&)
	If Length < 1 Then Exit Function
	'预设可接收文本内容的字节数，须预先赋空格
	ReDim byteBuffer(Length * 2 + 1) As Byte
	'最大允许存放 1024 个字符
	byteBuffer(1) = 4
	'获取光标所在行的文本字节数组
	SendMessage hwnd, EM_GETLINE, LineNo, byteBuffer(0)
	'转换为文本，并清除空字符
	GetCurPosLine = Replace$(StrConv$(byteBuffer,vbUnicode),vbNullChar,"")
End Function


'获取查找字串的查找方式
'GetFindMode = 0 常规，= 1 通配符, = 2 正则表达式
Public Function GetFindMode(FindStr As String) As Long
	'不含通配符和正则表达式专用字符时
	If (FindStr Like "*[$()+.^{|*?#[\]*") = False Then
		Exit Function
	End If
	'不含正则表达式专用字符时
	If (FindStr Like "*[$()+.^{|\]*") = False Then
		If (FindStr Like "*\[*?#[]*") = False Then
			GetFindMode = 1
		End If
		Exit Function
	End If
	GetFindMode = 2
End Function


'查找字串并移动光标位置
'Mode = False 从头到尾，否则从尾到头
'FindCurPos > 0 已找到，= 0 未找到, = -1 相同位置(只找到一项), = -2 通配符语法错误, = -3 正则表达式语法错误
Private Function FindCurPos(ByVal hwnd As Long,ByVal FindStr As String,ByVal Mode As Boolean,Optional strText As String) As Long
	Dim i As Long,n As Long,Lines As Long,Stemp As Integer,sLength As Long,bLength As Long
	Dim ptPos As POINTAPI,bkPos As POINTAPI,Matches As Object
	On Error GoTo errHandle
	'获取光标位置及文本框内字串的行数
	Lines = SendMessageLNG(hwnd, EM_GETLINECOUNT, 0&, 0&) - 1 'Lines 以 1 为起点
	If Lines = -1 Then Exit Function
	'检测查找内容的查找方式
	ReDim TempList(0) As String,tmpByte(0) As Byte
	Stemp = GetFindMode(FindStr)
	Select Case Stemp
	Case 0
		If (FindStr Like "*\[*?#[]*") = True Then
			TempList = ReSplit("*,?,#,[",",",-1)
			For i = 0 To UBound(TempList)
				FindStr = Replace$(FindStr,"\" & TempList(i),TempList(i))
			Next i
		End If
		sLength = Len(FindStr)
		'bLength = StrHexLength(FindStr,GetACP,0)
		tmpByte = StrConv(FindStr,vbFromUnicode)
		bLength = UBound(tmpByte) + 1
	Case 1
		'转换通配符为正则表达式模板
		TempList = ReSplit("\*,\#,\?,\[",",",-1)
		For i = 0 To UBound(TempList)
			FindStr = Replace$(FindStr,TempList(i),CStr$(i) & vbNullChar & CStr$(i) & vbNullChar & CStr$(i))
		Next i
		FindStr = Replace$(FindStr,"?",".")
		FindStr = Replace$(FindStr,"*",".*")
		FindStr = Replace$(FindStr,"#","\d")
		FindStr = Replace$(FindStr,"[!","[^")
		For i = 0 To UBound(TempList)
			FindStr = Replace(FindStr,CStr$(i) & vbNullChar & CStr$(i) & vbNullChar & CStr$(i),TempList(i))
		Next i
		FindStr = Replace$(FindStr,"\#","#")
	Case 2
		If CheckStrRegExp(FindStr,"(\\\(.+\\\))|(\\\(.+\))|(\(\?.+\))|(\(.+\).*\\[1-9]\d?)",0,2) = False Then
			If (FindStr Like "*(.*)*") = True Then Stemp = 3
		End If
	End Select
	ReDim TempList(1) As String,tmpByte(0) As Byte
	'初始化正则表达式
	With RegExp
		.Global = Mode
		.IgnoreCase = False
		.Pattern = FindStr
	End With
	'Win10 下，需要用到编辑框中的所有字串
	'If StrToLong(GetWindowsVersion()) < 62 Then
		strText = ""	'Win10 以下版本不需要
	'ElseIf strText = "" Then
	'	strText = GetTextBoxString(hwnd)
	'End If
	'获取光标位置  .y 为行号 (起点 = 0)，.x 为所在行的字符数偏移(起点 = 0)
	ptPos = GetCurPos(hwnd,Mode,strText)
	'备份起始点，用于判断找到的位置是否就是光标起点位置
	bkPos = ptPos
	'查找字串
	With ptPos
		Do
			TempList(0) = GetCurPosLine(hwnd,ptPos.y)
			If TempList(0) = "" Then GoTo NextLine
			If Stemp = 0 Then
				If Mode = False Then
   		 			.x = InStr(.x + 1,TempList(0),FindStr)
   	 			Else
    				.x = InStrRev(TempList(0),FindStr,.x - 1)
    			End If
   		 	Else
    			If Mode = False Then
    				TempList(1) = Mid$(TempList(0),.x + 1)
    			ElseIf .x > 0 Then
					TempList(1) = Left$(TempList(0),.x - 1)
				Else
					TempList(1) = TempList(0)
				End If
				If TempList(1) = "" Then GoTo NextLine
   				Set Matches = RegExp.Execute(TempList(1))
				If Matches.Count = 0 Then GoTo NextLine
				If Mode = False Then
					If Stemp = 3 Then
						.x = .x + InStr(1,TempList(1),Matches(0).SubMatches(0))
						sLength = Len(Matches(0).SubMatches(0))
						'bLength = StrHexLength(Matches(0).SubMatches(0),GetACP,0)
						tmpByte = StrConv(Matches(0).SubMatches(0),vbFromUnicode)
						bLength = UBound(tmpByte) + 1
					Else
						.x = .x + Matches(0).FirstIndex + 1
						sLength = Matches(0).Length
						'bLength = StrHexLength(Matches(0).Value,GetACP,0)
						tmpByte = StrConv(Matches(0).Value,vbFromUnicode)
						bLength = UBound(tmpByte) + 1
					End If
				Else
					i = Matches.Count - 1
					If Stemp = 3 Then
						.x = InStrRev(TempList(1),Matches(i).SubMatches(0),-1)
						sLength = Len(Matches(i).SubMatches(0))
						'bLength = StrHexLength(Matches(i).SubMatches(0),GetACP,0)
						tmpByte = StrConv(Matches(i).SubMatches(0),vbFromUnicode)
						bLength = UBound(tmpByte) + 1
					Else
						.x = Matches(i).FirstIndex + 1
						sLength = Matches(i).Length
						'bLength = StrHexLength(Matches(i).Value,GetACP,0)
						tmpByte = StrConv(Matches(i).Value,vbFromUnicode)
						bLength = UBound(tmpByte) + 1
					End If
				End If
			End If
   		 	If .x > 0 Then
				.x = .x - 1
				If .y = bkPos.y And .x + sLength = bkPos.x Then
					FindCurPos = -1
				Else
					FindCurPos = .y + 1
				End If
				If strText = "" Then
					Call SetCurPos(hwnd,.y,.x,sLength,strText)
				Else
					'Call SetCurPos(hwnd,.y,StrHexLength(Left$(TempList(0),.x),GetACP,0),bLength,strText)
					tmpByte = StrConv(Left$(TempList(0),.x),vbFromUnicode)
					Call SetCurPos(hwnd,.y,UBound(tmpByte) + 1,bLength,strText)
				End If
				strText = Mid$(TempList(0),.x + 1,sLength)
				Exit Do
			End If
			NextLine:
			.x = 0
			If Mode = False Then
				.y = .y + 1
				If .y > Lines Then .y = 0
			Else
				.y = .y - 1
				If .y < 0 Then .y = Lines
			End If
			n = n + 1
		Loop Until n > Lines + 1
	End With
	Exit Function
	errHandle:
	If Stemp > 0 Then FindCurPos = IIf(Stemp < 2,-2,-3)
End Function


'过滤字串
'Mode = 0 常规，= 1 通配符, = 2 正则表达式
'FilterStr = 1 已找到，= 0 未找到, = -1 程序错误, = -2 通配符语法错误 = -3 正则表达式语法错误
Public Function FilterStr(ByVal txtStr As String,ByVal FindStr As String,ByVal Mode As Long,Optional ByVal IgnoreCase As Boolean) As Long
	Dim i As Long,TempList() As String
	On Error GoTo errHandle
	Select Case Mode
	Case 0
		If (FindStr Like "*\[*?#[]*") = True Then
			TempList = ReSplit("*,?,#,[",",",-1)
			For i = 0 To UBound(TempList)
				FindStr = Replace$(FindStr,"\" & TempList(i),TempList(i))
			Next i
		End If
		If IgnoreCase = True Then
			txtStr = LCase(txtStr)
			FindStr = LCase(FindStr)
		End If
		If InStr(txtStr,FindStr) Then FilterStr = 1
	Case 1
		If IgnoreCase = True Then
			txtStr = LCase(txtStr)
			FindStr = LCase(FindStr)
		End If
		If (txtStr Like FindStr) = True Then FilterStr = 1
	Case 2
		If CheckStrRegExp(txtStr,FindStr,0,2,IgnoreCase) = True Then FilterStr = 1
	End Select
	Exit Function
	errHandle:
	FilterStr = -(Mode + 1)
End Function


'插入文本到文本框光标所在开始处，返回 lpPoint 光标坐标
Public Function InsertStr(ByVal hwnd As Long,ByVal strText As String,ByVal InsertText As String) As String
	Dim lpPoint As POINTAPI
	With lpPoint
		lpPoint = GetCurPos(hwnd,False)
		'SendMessage(hwnd, CB_GETEDITSEL, .x, .y)
		InsertStr = Left$(strText,.x) & InsertText & Mid$(strText,.x + 1)
		.x = .x + Len(InsertText)
	End With
	Call SetCurPos(hwnd,lpPoint.x,lpPoint.y,0)
End Function


'检查正则表达式是否正确
Public Function CheckRegExp(ByVal RegEx As Object,ByVal Patrn As String) As Boolean
	If Patrn = "" Then Exit Function
	On Error GoTo ExitFunction
	With RegEx
		.Pattern = Patrn
		.Test("CheckRegExp")
	End With
	CheckRegExp = True
	ExitFunction:
End Function


'检查是否需要检测特征码
Public Function GetFeaturePattern(CPList() As LANG_PROPERTIE) As String
	Dim i As Integer,n As Integer
	ReDim TempList(UBound(CPList)) As String
	For i = 0 To UBound(CPList)
		With CPList(i)
			If .FeatureCodeEnable = 1 Then
				TempList(i) = .FeatureCodeRegExpPattern
				n = n + 1
			Else
				TempList(i) = .UniCodeRegExpPattern
			End If
		End With
	Next i
	GetFeaturePattern = IIf(n > 0,StrListJoin(TempList,"|"),"")
End Function


'转换字符编码的提取顺序
Public Function ConvertStrCPOrder(StrEnCodeOrder As String) As Integer()
	Dim i As Integer,j As Integer,n As Integer
	ReDim intList(Len(StrEnCodeOrder)) As Integer
	For i = 1 To Len(StrEnCodeOrder)
		j = StrToLong(Mid$(StrEnCodeOrder,i,1))
		If j > 0 Then
			intList(j - 1) = i
			n = n + 1
		End If
	Next i
	If n > 0 Then n = n - 1
	ReDim Preserve intList(n) As Integer
	ConvertStrCPOrder = intList
End Function


'判断是否需要检查跟随引用地址的自定义字符类型
Public Function CheckRefCustomStrType(TypeList() As STRING_TYPE,ByVal FileType As Long) As Integer
	If FileType = 0 Then Exit Function
	For FileType = 0 To UBound(TypeList)
		With TypeList(FileType)
			If .CodeLoc > 0 Then
				If .CPCodePos <> 0 Or .LengthCodePos <> 0 Then
					CheckRefCustomStrType = 2
					Exit For
				End If
			End If
		End With
	Next FileType
End Function


'获取字符串结束符参数
'Index = 1 合格检查，= 2 去除结束符名称，= 3 结束符不前置 "?"，用于字串搜索，= 0 或 > 3 转换为前置符和结束符正则表达式
Public Function StrEndChar2Pattern(ByVal UseEndChar As String,Optional ByVal Index As Integer) As String()
	On Error GoTo ExitFunction
	ReDim Pattern(1) As String
	Select Case Index
	Case 1
		UseEndChar = Mid$(UseEndChar,InStr(UseEndChar,"("))
		If (UseEndChar Like "(*)") = True Then Pattern(0) = UseEndChar
	Case 2
		If Trim$(UseEndChar) = "" Then UseEndChar = ConvertStrEndCharSet(EndCharOfString,False)
		Pattern(0) = ReSplit(UseEndChar,ItemJoinStr,2)(1)
		Pattern(0) = ReSplit(Pattern(0),JoinStr)(StrToLong(ReSplit(UseEndChar,ItemJoinStr,2)(0)))
		Pattern(0) = Mid$(Pattern(0),InStr(Pattern(0),"("))
		Pattern(0) = Left$(Pattern(0),InStrRev(Pattern(0),")"))
	Case 3
		If Trim$(UseEndChar) = "" Then UseEndChar = ConvertStrEndCharSet(EndCharOfString,False)
		Pattern(0) = ReSplit(UseEndChar,ItemJoinStr,2)(1)
		Index = StrToLong(ReSplit(UseEndChar,ItemJoinStr,2)(0))
		Select Case Index
		Case 0
			Pattern(0) = "(\x00*)("
			Pattern(1) = ")(\x00+)"
		Case 1
			Pattern(0) = "(\x00*)("
			Pattern(1) = ")([^\x00]?)"
		Case 2
			Pattern(0) = "(\x00*)("
			Pattern(1) = ")(\x00*)"
		Case 3
			Pattern(0) = "(\r*)("
			Pattern(1) = ")(\r+)"
		Case 4
			Pattern(0) = "(\n*)("
			Pattern(1) = ")(\n+)"
		Case 5
			Pattern(0) = "([\r\n]*)("
			Pattern(1) = ")([\r\n]+)"
		Case 6
			Pattern(0) = "(\t*)("
			Pattern(1) = ")(\t+)"
		Case Else
			Pattern(1) = ReSplit(Pattern(0),JoinStr)(Index)
			Pattern(1) = Mid$(Pattern(1),InStr(Pattern(1),"("))
			Pattern(1) = Left$(Pattern(1),InStrRev(Pattern(1),")"))
			Pattern(0) = Pattern(1) & "?("
			Pattern(1) = ")" & Pattern(1)
		End Select
	Case Else
		If Trim$(UseEndChar) = "" Then UseEndChar = ConvertStrEndCharSet(EndCharOfString,False)
		Pattern(0) = ReSplit(UseEndChar,ItemJoinStr,2)(1)
		Index = StrToLong(ReSplit(UseEndChar,ItemJoinStr,2)(0))
		Select Case Index
		Case 0
			Pattern(0) = "(\x00*)("
			Pattern(1) = ")(\x00+)"
		Case 1
			Pattern(0) = "(\x00*)("
			Pattern(1) = ")([^\x00]?)"
		Case 2
			Pattern(0) = "(\x00*)("
			Pattern(1) = ")(\x00*)"
		Case 3
			Pattern(0) = "(\r*)("
			Pattern(1) = "?)(\r+)"
		Case 4
			Pattern(0) = "(\n*)("
			Pattern(1) = "?)(\n+)"
		Case 5
			Pattern(0) = "([\r\n]*)("
			Pattern(1) = "?)([\r\n]+)"
		Case 6
			Pattern(0) = "(\t*)("
			Pattern(1) = "?)(\t+)"
		Case Else
			Pattern(1) = ReSplit(Pattern(0),JoinStr)(Index)
			Pattern(1) = Mid$(Pattern(1),InStr(Pattern(1),"("))
			Pattern(1) = Left$(Pattern(1),InStrRev(Pattern(1),")"))
			Pattern(0) = Pattern(1) & "?("
			Pattern(1) = "?)" & Pattern(1)
		End Select
	End Select
	StrEndChar2Pattern = Pattern
	Exit Function
	ExitFunction:
	If Index > 6 Then
		Pattern(0) = "(\x00*)("
		Pattern(1) = ")(\x00+)"
		StrEndChar2Pattern = Pattern
	End If
End Function


'合并字符串结束符参数
Public Function MergeStrEndCharSet(ByVal UseEndChar As String,Optional ByVal DefEndChar As String) As String
	Dim i As Integer,TempList() As String,TempArray() As String
	If Trim$(UseEndChar) = "" Then
		MergeStrEndCharSet = ConvertStrEndCharSet(EndCharOfString,False)
		Exit Function
	End If
	If Trim$(DefEndChar) = "" Then
		TempList = ReSplit(ReSplit(UseEndChar,ItemJoinStr,2)(1),JoinStr)
		TempArray = ReSplit(ReSplit(EndCharOfString,ItemJoinStr,2)(1),ValJoinStr)
		For i = 0 To UBound(TempArray)
			TempList(i) = TempArray(i)
		Next i
	Else
		Dim TempArry() As String
		TempList = ReSplit(ReSplit(UseEndChar,ItemJoinStr,2)(1),JoinStr)
		TempArry = ReSplit(ReSplit(EndCharOfString,ItemJoinStr,2)(1),ValJoinStr)
		TempArray = ReSplit(DefEndChar,ItemJoinStr)
		For i = 0 To UBound(TempArray)
			TempList(i) = TempArray(i) & " " & TempArry(i)
		Next i
	End If
	MergeStrEndCharSet = ReSplit(UseEndChar,ItemJoinStr,2)(0) & ItemJoinStr & StrListJoin(TempList,JoinStr)
End Function


'添加字符串结束符参数
Public Function AddStrEndCharSet(ByVal AllEndChar As String,ByVal UseEndChar As String) As String
	Dim i As Integer,TempList() As String
	AddStrEndCharSet = AllEndChar
	If InStr(UseEndChar,JoinStr) = 0 Then Exit Function
	AllEndChar = Trim$(ReSplit(UseEndChar,JoinStr,2)(1))
	If AllEndChar = "" Then Exit Function
	TempList = ReSplit(ReSplit(AddStrEndCharSet,ItemJoinStr,2)(1),JoinStr)
	For i = 0 To UBound(TempList)
		TempList(i) = Mid$(TempList(i),InStr(TempList(i),"("))
		TempList(i) = Left$(TempList(i),InStrRev(TempList(i),")"))
		If TempList(i) = AllEndChar Then
			i = -(i + 1)
			Exit For
		End If
	Next i
	If i > -1 Then
		i = i - 1
		AddStrEndCharSet = AddStrEndCharSet & JoinStr & Replace$(UseEndChar,JoinStr," ")
	End If
	AddStrEndCharSet = CStr$(Abs(i + 1)) & ItemJoinStr & ReSplit(AddStrEndCharSet,ItemJoinStr,2)(1)
End Function


'转换字符串结束符参数，以便输入空字节以外所有字符
Public Function ConvertStrEndCharSet(ByVal UseEndChar As String,ByVal Mode As Boolean) As String
	If Trim$(UseEndChar) = "" Then Exit Function
	If Mode = False Then
		ConvertStrEndCharSet = Replace$(UseEndChar,")" & ValJoinStr,")" & JoinStr)
	Else
		ConvertStrEndCharSet = Replace$(UseEndChar,")" & JoinStr,")" & ValJoinStr)
	End If
End Function


'转换字串列表为以列表值为 Key 的字典
'DelAccKey <> 0 删除快捷键并添加转义的字串到字典，否则不删除，返回的字串列表都不转义
'Mode = 0 转义不返回字串列表，Mode = 1 不转义不返回字串列表，Mode = 2 不转义并返回被转换为字典的 List 字串列表
Public Function StrList2StrDic(StrDic As Object,StrList() As String,ByVal DelKey As Integer,ByVal Mode As Integer) As Boolean
	Dim i As Long,n As Long,Temp As String
	StrDic.RemoveAll
	If CheckArray(StrList) = False Then
		ReDim StrList(0) As String
		Exit Function
	End If
	For i = 0 To UBound(StrList)
		StrList(i) = Trim$(StrList(i))
		If StrList(i) <> "" Then
			If Mode = 0 Then
				Temp = Convert(StrList(i))
			Else
				Temp = StrList(i)
			End If
			If DelKey = 0 Then
				If Not StrDic.Exists(Temp) Then
					StrDic.Add(Temp,n)
					If Mode = 2 Then StrList(n) = StrList(i)
					n = n + 1
				End If
			Else
				Temp = Trim$(DelAccKey(Temp))
				If Temp <> "" Then
					If Not StrDic.Exists(Temp) Then
						StrDic.Add(Temp,n)
						If Mode = 2 Then StrList(n) = StrList(i)
						n = n + 1
					End If
				End If
			End If
		End If
	Next i
	If n > 0 Then
		If Mode = 2 Then
			ReDim Preserve StrList(n - 1) As String
		End If
		StrList2StrDic = True
	ElseIf Mode = 2 Then
		ReDim StrList(0) As String
	End If
End Function


'删除快捷键 (包括带括号的亚洲语言的快捷键)
Public Function DelAccKey(ByVal strText As String) As String
	Dim i As Long,j As Long,TempList() As String,Matches As Object
	On Error GoTo errHandle
	DelAccKey = strText
	i = InStr(strText,"&")
	If i < 1 Then Exit Function
	'初始化正则表达式
	With RegExp
		.Global = True
		.IgnoreCase = True
		TempList = ReSplit("\u0028,\u0029,\u005B,\u005D,\u003C,\u003E,\uFF08,\uFF09,\uFF3B,\uFF3D,\uFF1C,\uFF1F,\u3008,\u3009",",")
		For i = 0 To UBound(TempList) - 1 Step 2
			.Pattern = TempList(i) & "((&amp;)|&)\S" & TempList(i + 1)
			Set Matches = .Execute(strText)
			If Matches.Count > 0 Then
				For j = 0 To Matches.Count - 1
					strText = Replace$(strText,Matches(j).Value,"")
					Exit For
				Next j
			End If
		Next i
	End With
	errHandle:
	DelAccKey = Replace$(Replace$(strText,"&amp;",""),"&","")
End Function


'用途：将十进制转化为二进制
'输入：Dec(十进制数)
'输出：DECtoBIN(二进制数)
'输入的最大数为2147483647,输出最大数为1111111111111111111111111111111(31个1)
Private Function DECtoBIN(Dec As Long) As String
	Do While Dec > 0
		DECtoBIN = Dec Mod 2 & DECtoBIN
		If Dec < 2 Then Exit Do
		Dec = Dec \ 2
	Loop
End Function


'位左移
Private Function SHL(nSource As Long, n As Byte) As Double
	On Error GoTo ExitFunction:
	SHL = nSource * 2 ^ n
	ExitFunction:
End Function


'位右移
Private Function SHR(nSource As Long, n As Byte) As Long
	On Error GoTo ExitFunction:
	SHR = nSource / 2 ^ n
	ExitFunction:
End Function


'获得指定的位
Private Function GetBits(nSource As Long, n As Byte) As Long
	On Error GoTo ExitFunction:
	GetBits = nSource And 2 ^ n
	ExitFunction:
End Function


'设置指定的位
Private Function SetBits(nSource As Long, n As Byte) As Long
	On Error GoTo ExitFunction:
	SetBits = nSource Or 2 ^ n
	ExitFunction:
End Function


'清除指定的位
Private Function ResetBits(nSource As Long, n As Byte) As Long
	On Error GoTo ExitFunction:
	ResetBits = nSource And Not 2 ^ n
	ExitFunction:
End Function


'翻转字节顺序 (16-bit)
Private Function UInt16ReverseBytes(ByVal Value As Integer) As Integer
	UInt16ReverseBytes = SHL((Value And &HFF),8) Or SHR((Value And &HFF00),8)
End Function


'翻转字节顺序 (32-bit)
Private Function UInt32ReverseBytes(ByVal Value As Long) As Long
	UInt32ReverseBytes = SHL((Value And &HFF),24) Or SHL((Value And &HFF00),8) Or _
         SHR((Value And &HFF0000),8) Or SHR((Value And &HFF000000),24)
End Function


'获取 N 个字符串的最长公共子串
Public Function MaxSubString(StrList() As String) As String
	Dim i As Integer,j As Integer,k As Long
	Dim Length As Long,Index As Integer,Max As Integer
	Dim SubStr As String,AllisExt As Boolean
	Max = UBound(StrList)
	For i = 0 To Max
		k = Len(StrList(i))
		If Length = 0 Then
			Length = k
		ElseIf k < Length Then
			Length = k: Index = i
		End If
	Next
	For i = Length To 1 Step -1
		For j = 1 To Length - i + 1
			SubStr = Mid$(StrList(Index),j,i)
			For k = 0 To Max
				If k <> Index Then
					AllisExt = InStr(StrList(k),SubStr)
					If Not AllisExt Then Exit For
				End If
			Next k
			If AllisExt Then
				MaxSubString = SubStr
				Exit Function
			End If
		Next j
	Next i
End Function


'计算并检查二个 Long 数值相加值是否溢出
'返回值: =0 溢出, >0 未溢出
Public Function CheckLongPlus(ByVal Long1 As Long,ByVal Long2 As Long) As Variant
	On Error GoTo errHandle
	CheckLongPlus = Long1 + Long2
	Exit Function
	errHandle:
	Err.Clear
	CheckLongPlus = 0
End Function


'检查数组是否已经初始化
'返回值:TRUE 已经初始化, FALSE 未初始化
Public Function CheckArrEmpty(ByRef MyArr As Variant) As Boolean
	On Error Resume Next
	If UBound(MyArr) >= 0 Then CheckArrEmpty = True
	Err.Clear
End Function


'检查数组的某个数组ID是否存在
'返回值:TRUE 已经初始化, FALSE 未初始化
Private Function CheckArrID(ByRef MyArr As Variant,ID As Long) As Boolean
	Dim i As Variant
	On Error Resume Next
	i = MyArr(ID)
	CheckArrID = (Err.Number = 0)
	Err.Clear
End Function


'检查字串数组是否为空，非空返回 True
Public Function CheckArray(DataList() As String) As Boolean
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


'检查 INI 数据数组是否为空，非空返回 True
Private Function CheckINIArray(DataList() As INIFILE_DATA) As Boolean
	Dim i As Long
	On Error GoTo errHandle
	For i = LBound(DataList) To UBound(DataList)
		If DataList(i).Title <> "" Then
			CheckINIArray = True
			Exit For
		End If
	Next i
	errHandle:
End Function


'检查 HCS 字串数据数组是否为空，非空返回 True
Public Function CheckDataArray(DataList() As STRING_PROPERTIE) As Boolean
	Dim i As Long
	On Error GoTo errHandle
	For i = LBound(DataList) To UBound(DataList)
		If DataList(i).Source.CodePage > 0 Then
			CheckDataArray = True
			Exit For
		End If
	Next i
	errHandle:
End Function


'检查自定义字串类型的数组是否为空，非空返回 True
Public Function CheckStrTypeArray(DataList() As STRING_TYPE,Optional ByVal SetID As Long = -1) As Boolean
	Dim i As Long,Max As Long
	CheckStrTypeArray = True
	On Error GoTo errHandle
	If SetID = -1 Then
		SetID = 0: Max = UBound(DataList)
	Else
		Max = SetID
	End If
	For i = SetID To Max
		If DataList(i).FristCodePos > 0 Then Exit Function
		If DataList(i).CPCodePos <> 0 Then Exit Function
		If DataList(i).CPCodeSize <> 0 Then Exit Function
		If DataList(i).LengthCodePos <> 0 Then Exit Function
		If DataList(i).LengthCodeSize <> 0 Then Exit Function
		If DataList(i).StartCodePos <> 0 Then Exit Function
		If DataList(i).StartCodeLength <> 0 Then Exit Function
		If DataList(i).EndCodeLength <> 0 Then Exit Function
		If DataList(i).RefCodeStartPos <> 0 Then Exit Function
		If DataList(i).RefCodeStartLength <> 0 Then Exit Function
	Next i
	errHandle:
	CheckStrTypeArray = False
End Function


'返回列表框所有项目数
'Mode = 0  所有项，= 1 选定项，= 2 未选定项
Public Function GetListBoxCount(ByVal hwnd As Long,ByVal Mode As Long) As Long
	On Error GoTo errHandle
	Select Case Mode
	Case 0
		GetListBoxCount = SendMessageLNG(hwnd, LB_GETCOUNT, 0&, 0&)
	Case 1
		GetListBoxCount = SendMessageLNG(hwnd, LB_GETSELCOUNT, 0&, 0&)
	Case 2
		GetListBoxCount = SendMessageLNG(hwnd, LB_GETCOUNT, 0&, 0&)
		GetListBoxCount = GetListBoxCount - SendMessageLNG(hwnd, LB_GETSELCOUNT, 0&, 0&)
	End Select
	Exit Function
	errHandle:
	GetListBoxCount = 0
End Function


'获取选定列表框项目的索引
Public Function GetListBoxIndexs(ByVal hwnd As Long) As Long()
	Dim i As Long
	ReDim Indexs(0) As Long
	On Error GoTo errHandle
	i = SendMessageLNG(hwnd, LB_GETSELCOUNT, 0&, 0&)
	If i < 1 Then GoTo errHandle
	ReDim Indexs(i - 1) As Long
	SendMessage(hwnd, LB_GETSELITEMS, i, Indexs(0))
	errHandle:
	GetListBoxIndexs = Indexs
End Function


'返回列表框中第一个可见项的索引
Public Function GetListBoxTopIndex(ByVal hwnd As Long) As Long
	On Error GoTo errHandle
	GetListBoxTopIndex = SendMessageLNG(hwnd, LB_GETTOPINDEX, 0&, 0&)
	Exit Function
	errHandle:
	GetListBoxTopIndex = 0
End Function


'设置列表框中第一个可见项的索引
Private Function SetListBoxTopIndex(ByVal hwnd As Long,ByVal TopItem As Long) As Boolean
	On Error GoTo errHandle
	SendMessageLNG(hwnd, LB_SETTOPINDEX, TopItem, 0&)
	SetListBoxTopIndex = True
	errHandle:
End Function


'选定指定的列表框项目
'Indexs = -1 全选，否则选择指定项
Public Function SetListBoxItems(ByVal hwnd As Long,ByVal Indexs As Variant,Optional ByVal TopItem As Long = -1) As Boolean
	Dim i As Long
	On Error GoTo errHandle
	i = SendMessageLNG(hwnd, LB_GETCOUNT, 0&, 0&)
	If i = 0 Then Exit Function
	If IsArray(Indexs) Then
		SendMessageLNG(hwnd, LB_SETSEL, False, -1)
		For i = 0 To UBound(Indexs)
			SendMessageLNG(hwnd, LB_SETSEL, True, Indexs(i))
		Next i
	Else
		If Indexs = -1 Then
			SendMessageLNG(hwnd, LB_SETSEL, True, Indexs)
		Else
			SendMessageLNG(hwnd, LB_SETSEL, False, -1)
			SendMessageLNG(hwnd, LB_SETSEL, True, Indexs)
		End If
	End If
	If TopItem > -1 Then
		SendMessageLNG(hwnd, LB_SETTOPINDEX, TopItem, 0&)
	End If
	SetListBoxItems = True
	Exit Function
	errHandle:
End Function


'插入或附加列表框项目
'InsPos = -1 附加到最后，否则插入到指定索引位置
Public Function AddListBoxItems(ByVal hwnd As Long,ByVal StrList As Variant,ByVal InsPos As Variant) As Boolean
	Dim i As Long
	On Error GoTo errHandle
	If Int(PSL.Version / 100) < 16 Then
		If IsArray(StrList) Then
			If IsArray(InsPos) Then
				For i = 0 To UBound(StrList)
					SendMessageOLD(hwnd, LB_INSERTSTRING, InsPos(i), StrPtr(StrConv$(StrList(i),vbFromUnicode)))
				Next i
			ElseIf InsPos > -1 Then
				For i = 0 To UBound(StrList)
					SendMessageOLD(hwnd, LB_INSERTSTRING, InsPos + i, StrPtr(StrConv$(StrList(i),vbFromUnicode)))
				Next i
			Else
				For i = 0 To UBound(StrList)
					SendMessageOLD(hwnd, LB_INSERTSTRING, -1, StrPtr(StrConv$(StrList(i),vbFromUnicode)))
				Next i
			End If
		Else
			If IsArray(InsPos) Then
				For i = 0 To UBound(InsPos)
					SendMessageOLD(hwnd, LB_INSERTSTRING, InsPos(i), StrPtr(StrConv$(StrList,vbFromUnicode)))
				Next i
			ElseIf InsPos > -1 Then
				SendMessageOLD(hwnd, LB_INSERTSTRING, InsPos, StrPtr(StrConv$(StrList,vbFromUnicode)))
			Else
				SendMessageOLD(hwnd, LB_INSERTSTRING, -1, StrPtr(StrConv$(StrList,vbFromUnicode)))
			End If
		End If
	Else
		If IsArray(StrList) Then
			If IsArray(InsPos) Then
				For i = 0 To UBound(StrList)
					SendMessageLNG(hwnd, LB_INSERTSTRING, InsPos(i), StrPtr(StrList(i)))
				Next i
			ElseIf InsPos > -1 Then
				For i = 0 To UBound(StrList)
					SendMessageLNG(hwnd, LB_INSERTSTRING, InsPos + i, StrPtr(StrList(i)))
				Next i
			Else
				For i = 0 To UBound(StrList)
					SendMessageLNG(hwnd, LB_INSERTSTRING, -1, StrPtr(StrList(i)))
				Next i
			End If
		Else
			If IsArray(InsPos) Then
				For i = 0 To UBound(InsPos)
					SendMessageLNG(hwnd, LB_INSERTSTRING, InsPos(i), StrPtr(StrList))
				Next i
			ElseIf InsPos > -1 Then
				SendMessageLNG(hwnd, LB_INSERTSTRING, InsPos, StrPtr(StrList))
			Else
				SendMessageLNG(hwnd, LB_INSERTSTRING, -1, StrPtr(StrList))
			End If
		End If
	End If
	AddListBoxItems = True
	Exit Function
	errHandle:
End Function


'删除列表框项目
'DelPos = -1 删除最大项，< -1 全部清空，否则删除指定索引号的项目
Public Function DelListBoxItems(ByVal hwnd As Long,ByVal DelPos As Variant) As Boolean
	Dim i As Long
	On Error GoTo errHandle
	If IsArray(DelPos) Then
		For i = UBound(DelPos) To 0 Step -1
			SendMessageLNG(hwnd, LB_DELETESTRING, DelPos(i), 0&)
		Next i
	ElseIf DelPos > -1 Then
		SendMessageLNG(hwnd, LB_DELETESTRING, DelPos, 0&)
	ElseIf DelPos = -1 Then
		SendMessageLNG(hwnd, LB_DELETESTRING, GetListBoxCount(hwnd, 0&) - 1, 0&)
	Else
		SendMessageLNG(hwnd, LB_RESETCONTENT, 0&, 0&)
		SendMessageLNG(hwnd, LB_SETHORIZONTALEXTENT, 0&, 0&)
	End If
	DelListBoxItems = True
	Exit Function
	errHandle:
End Function


'更改列表框项目
Public Function ChangeListBoxItems(ByVal hwnd As Long,ByVal StrList As Variant,ByVal DelPos As Variant,ByVal InsPos As Variant) As Boolean
	Dim i As Long
	On Error GoTo errHandle
	If IsArray(StrList) And IsArray(DelPos) And IsArray(InsPos) Then
		For i = 0 To UBound(DelPos)
			If DelListBoxItems(hwnd,DelPos(i)) = True Then
				ChangeListBoxItems = AddListBoxItems(hwnd,StrList(i),InsPos(i))
			End If
		Next i
	ElseIf DelListBoxItems(hwnd,DelPos) = True Then
		ChangeListBoxItems = AddListBoxItems(hwnd,StrList,InsPos)
	End If
	ChangeListBoxItems = True
	Exit Function
	errHandle:
End Function


'返回对话框某个控件中的字串
Public Function GetTextBoxString(ByVal hwnd As Long) As String
	Dim i As Long
	On Error GoTo errHandle
	'可以访问其他程序的窗口和控件
	'i = SendMessageLNG(hwnd,WM_GETTEXTLENGTH,0&,0&)
	'只能访问自身的窗口和控件
	i = GetWindowTextLength(hwnd)
	If i > 0 Then
		GetTextBoxString = String$(i + 1,0)
		'可以访问其他程序的窗口和控件，但速度很慢
		'SendMessageLNG hwnd, WM_GETTEXT, i + 1, StrPtr(GetTextBoxString)
		'GetTextBoxString = Replace$(StrConv$(GetTextBoxString,vbUnicode),vbNullChar,"")
		'只能访问自身的窗口和控件，但速度快好几倍
		GetWindowText hwnd, GetTextBoxString, i + 1
		GetTextBoxString = Replace$(GetTextBoxString,vbNullChar,"")
	End If
	Exit Function
	errHandle:
	GetTextBoxString = ""
End Function


'设置对话框某个控件中的字串
'Mode = False 替换显示，否则，追加显示
Public Function SetTextBoxString(ByVal hwnd As Long,ByVal StrText As String,Optional ByVal Mode As Boolean) As Boolean
	On Error GoTo errHandle
	If Mode = False Then
		'可以访问其他程序的窗口和控件
		'SetTextBoxString = SendMessageLNG(hwnd,WM_SETTEXT,0&,StrPtr(StrConv$(StrText,vbFromUnicode)))
		'只能访问自身的窗口和控件，但会保留并正确显示最前面的一部分字符
		SetTextBoxString = SetWindowText(hwnd,StrText)
	Else
		StrText = GetTextBoxString(hwnd) & vbCrLf & StrText
		'设置文本框的最大允许长度
		SetTextBoxLength hwnd, Len(StrText), True
		'可以访问其他程序的窗口和控件，但会丢失并乱码最前面的一部分字符
		'SetTextBoxString = SendMessageLNG(hwnd,WM_SETTEXT,0&,StrPtr(StrConv$(StrText,vbFromUnicode)))
		'只能访问自身的窗口和控件，但会保留并正确显示最前面的一部分字符
		SetTextBoxString = SetWindowText(hwnd,StrText)
		'设置滚动条到文本框底部
		SendMessageLNG hwnd, WM_VSCROLL, SB_BOTTOM, 0&
	End If
	errHandle:
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


'获取字串的像素大小
Public Function GetStrPixels(ByVal hwnd As Long,ByVal StrText As String) As Long
	Dim hDC As Long,rt As RECT	',PT As POINTAPI	'
	On Error GoTo errHandle
	hDC = GetDC(hwnd)
	If DrawText(hDC,StrText,-1,rt,DT_CALCRECT) = 0 Then GoTo errHandle
	GetStrPixels = rt.Right - rt.Left + 8
	'If GetTextExtentPoint32(hDC,StrText,Len(StrText) * 1.2,PT) = 0 Then GoTo errHandle
	'GetStrPixels = PT.X + 4
	ReleaseDC(hwnd, hDC)
	Exit Function
	errHandle:
	GetStrPixels = SendMessageLNG(hwnd, LB_GETHORIZONTALEXTENT, 0&, 0&)
End Function


'检查字体是否为空，非空返回 True
Public Function CheckFont(LF As LOG_FONT) As Boolean
	If ReSplit(StrConv$(LF.lfFaceName,vbUnicode),vbNullChar,2)(0) <> "" Then CheckFont = True
End Function


'获取字体名称和字号
Public Function GetFontText(ByVal hwnd As Long,LF As LOG_FONT) As String
	Dim LF2 As LOG_FONT
	LF2 = LF
	If CheckFont(LF2) = False Then
		If hwnd = 0 Then Exit Function
		GetObjectAPI(SendMessageLNG(hwnd,WM_GETFONT,0,0),Len(LF2),VarPtr(LF2))
	End If
	'GetFontText = ReSplit(StrConv$(LF2.lfFaceName,vbUnicode),vbNullChar,2)(0) & " " & CStr$(-LF2.lfHeight)
	GetFontText = ReSplit(MultiByteToUTF16(LF2.lfFaceName,GetACP),vbNullChar,2)(0) & " " & CStr$(-LF2.lfHeight)
End Function


'比较二个字体数组是否相同，不相同返回 True
Public Function FontComps(LF() As LOG_FONT,LF2() As LOG_FONT) As Boolean
	Dim i As Long
	If UBound(LF) <> UBound(LF2) Then
		FontComps = True
		Exit Function
	End If
	For i = LBound(LF) To UBound(LF)
		If FontComp(LF(i),LF2(i)) = True Then
			FontComps = True
			Exit For
		End If
	Next i
End Function


'比较二个字体是否相同，不相同返回 True
Public Function FontComp(LF As LOG_FONT,LF2 As LOG_FONT) As Boolean
	FontComp = True
	With LF
		If .lfCharSet <> LF2.lfCharSet Then Exit Function
		If .lfClipPrecision <> LF2.lfClipPrecision Then Exit Function
		If .lfEscapement <> LF2.lfEscapement Then Exit Function
		If .lfFaceName <> LF2.lfFaceName Then Exit Function
		If .lfHeight <> LF2.lfHeight Then Exit Function
		If .lfItalic <> LF2.lfItalic Then Exit Function
		If .lfOrientation <> LF2.lfOrientation Then Exit Function
		If .lfOutPrecision <> LF2.lfOutPrecision Then Exit Function
		If .lfPitchAndFamily <> LF2.lfPitchAndFamily Then Exit Function
		If .lfQuality <> LF2.lfQuality Then Exit Function
		If .lfStrikeOut <> LF2.lfStrikeOut Then Exit Function
		If .lfUnderline <> LF2.lfUnderline Then Exit Function
		If .lfWeight <> LF2.lfWeight Then Exit Function
		If .lfWidth <> LF2.lfWidth Then Exit Function
		If .lfColor <> LF2.lfColor Then Exit Function
	End With
	FontComp = False
End Function


'弹出系统字体对话框选择字体，确定时返回非零
Public Function SelectFont(ByVal hwnd As Long,LF As LOG_FONT) As Long
	Dim CF As CHOOSE_FONT,LF2 As LOG_FONT
	LF2 = LF
	If CheckFont(LF2) = False Then
		If hwnd <> 0 Then
			GetObjectAPI(SendMessageLNG(hwnd,WM_GETFONT,0,0),Len(LF2),VarPtr(LF2))
		End If
	End If
	With CF
		.lStructSize = Len(CF)			'size of structure
		.hwndOwner = hwnd				'window Form1 is opening this dialog box
		'.hDC = GetDC(hWnd)				'device context of default printer (using VB's mechanism)
		.lpLogFont = VarPtr(LF2)		'LogFont结构地址
		'.iPointSize = LF.lfHeight		'10 * size in points of selected font
		.flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE Or CF_INACTIVEFONTS
		.rgbColors = LF2.lfColor		'RGB(0,0,0)		'black
		'.lCustData = 0					' data passed to hook fn
		'.lpfnHook = 0					' ptr. to hook function
		'.lpTemplateName = ""			' custom template name
		'.hInstance = 0					' instance handle of.EXE that contains cust. dlg. template
		'.lpszStyle = LF2.lfFaceName	' return the style field here must be LF_FACESIZE or bigger
		.nFontType = LF2.lfWeight		'REGULAR_FONTTYPE	'regular font type i.e. not bold or anything
		.nSizeMin = 8 					'minimum point size
		.nSizeMax = 16 					'maximum point size
	End With
	SelectFont = ChooseFont(CF)
	If SelectFont = 0 Then Exit Function
	LF = LF2
	LF.lfColor = CF.rgbColors
End Function


'创建字体，返回字体句柄
Public Function CreateFont(ByVal hwnd As Long,LF As LOG_FONT) As Long
	Dim LF2 As LOG_FONT
	LF2 = LF
	If CheckFont(LF2) = False Then
		If hwnd = 0 Then Exit Function
		GetObjectAPI(SendMessageLNG(hwnd,WM_GETFONT,0,0),Len(LF2),VarPtr(LF2))
	End If
	CreateFont = CreateFontIndirect(LF2)
End Function


'重画整个对话框
Public Function DrawWindow(ByVal hwnd As Long,ByVal hFont As Long) As Long
	'Dim New_hFont As Long,hDC As Long
	'hDC = GetDC(hwnd)
	'New_hFont = SelectObject(hDC, hFont)
	'SendMessageLNG(hWnd,WM_SETREDRAW,True,0)
	SendMessageLNG(hwnd,WM_SETFONT,hFont,0)
	'SendMessageLNG(hWnd,WM_PAINT,0,0)
	DrawWindow = RedrawWindow(hwnd,0,0,RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW)
	'SelectObject hDC, New_hFont
	'DeleteObject(hFont)
	'ReleaseDC(hwnd, hDC)
End Function


'重画对话框文本
'Private Function DrawDlgText(ByVal hwnd As Long,hFont As Long,ByVal StrText As String,ByVal Color As Long) As Long
	'On Error Resume Next
'	Dim Old_hFont As Long, hDC As Long, rt As RECT
'	hDC = GetDC(hwnd)
	'hFont = SendMessageLNG(hwnd,WM_GETFONT,0,0)
'	Old_hFont = SelectObject(hDC, hFont)
'	SetTextColor(hDC,Color)
'	DrawDlgText = DrawText(hDC,StrConv$(StrText,vbFromUnicode),-1,rt,DT_SINGLELINE And DT_CALCRECT)
'	SelectObject hDC, Old_hFont
'	ReleaseDC(hwnd, hDC)
'End Function


'字串类型结束标识符转正则表达式模板
'Setting = "0" 时，StrTypeEndChar2Pattern 返回 2 维空数组
Public Function StrTypeEndChar2Pattern(TypeList() As STRING_TYPE,CPList() As LANG_PROPERTIE,EncodeList() As Integer,ByVal OptionSet As String) As String()
	Dim i As Long,j As Long,k As Long,n As Long,CP As Long,Bytes() As Byte
	ReDim TempArray(UBound(CPList),UBound(EncodeList)) As String
	If OptionSet = "1" Then
		If CheckStrTypeArray(TypeList) = False Then OptionSet = "0"
	End If
	For i = 0 To UBound(CPList)
		For j = 0 To UBound(EncodeList)
			Select Case EncodeList(j)
			Case 1
				CP = CPList(i).CodePage
			Case 2
				CP = CP_UNICODELITTLE
			Case 3
				CP = CP_UTF8
			Case 4
				CP = CP_UNICODEBIG
			Case 5
				CP = CP_UTF7
			Case 6
				CP = CP_UTF32LE
			Case 7
				CP = CP_UTF32BE
			End Select
			n = 0
			If OptionSet = "1" Then
				ReDim TempList(UBound(TypeList)) As String
				For k = 0 To UBound(TypeList)
					With TypeList(k)
						If .CodeLoc = 0 And .EndCodeLength > 0 Then
							If .EndCodeString = "00" Then
								TempList(n) = "[\x00\x01]"
							Else
								Bytes = .EndCodeByte
								Select Case CP
								Case CP_UNICODELITTLE, CP_UNICODEBIG, CP_UTF32LE, CP_UTF32BE
									If .EndCodeLength And 1 Then
										ReDim Preserve Bytes(.EndCodeLength) As Byte
									End If
								End Select
								TempList(n) = "(" & Str2RegExpPattern(ByteToString(Bytes,CP),1) & ")"
							End If
							n = n + 1
						End If
					End With
				Next k
				If n > 0 Then
					n = n - 1
					ReDim Preserve TempList(n) As String
					TempArray(i,j) = ")(" & StrListJoin(TempList,"|") & "?)("
				Else
					TempArray(i,j) = ")()("
				End If
			Else
				TempArray(i,j) = ")()("
			End If
		Next j
	Next i
	StrTypeEndChar2Pattern = TempArray
End Function


'获取 Blob 流长度压缩后所需长度
'返回 CorSigCompressLength = 签名长度
Public Function CorSigCompressLength(ByVal Length As Long) As Long
	Dim i As Integer
	If Length > &H3FFF Then
		CorSigCompressLength = 4
	ElseIf Length > &H7F Then
		CorSigCompressLength = 2
	Else
		CorSigCompressLength = 1
	End If
End Function


'Blob 流长度压缩
'返回 CorSigCompressByte = 压缩后的字节数组，Length = 字节长度
Public Function CorSigCompressByte(ByVal Length As Long) As Byte()
	If Length > &H3FFF Then
		CorSigCompressByte = Val2Bytes((&HC0000000 Or Length),4,True)
	ElseIf Length > &H7F Then
		CorSigCompressByte = Val2Bytes((&H8000 Or Length),2,True)
	Else
		CorSigCompressByte = Val2Bytes(Length,1,True)
	End If
End Function


'Blob 流长度压缩
'返回 CorSigCompressData = 签名长度，Length = 字节长度
Public Function CorSigCompressData(ByVal Length As Long,Bytes() As Byte) As Integer
	If Length > &H3FFF Then
		CorSigCompressData = 4
		Bytes = Val2Bytes((&HC0000000 Or Length),4,True)
	ElseIf Length > &H7F Then
		CorSigCompressData = 2
		Bytes = Val2Bytes((&H8000 Or Length),2,True)
	Else
		CorSigCompressData = 1
		Bytes = Val2Bytes(Length,1,True)
	End If
End Function


'Blob 流长度解压缩
'返回 CorSigUncompressData = 压缩长度，Length = 除长度标识符外的字节长度（包括是否包含 > &H7F 字符标识符）
'每个二进制数据块头，都有1个长度数据块，通过移位运算，计算出长度数据块的实际长度
'如果第一个字节最高位为0，则此数据块长度为1个字节
'如果第一个字节最高位为10，则此数据块长度为2个字节
'如果第一个字节最高位为110，则此数据块长度为4个字节
Public Function CorSigUncompressData(FN As Variant,ByVal Index As Long,Length As Long,ByVal Mode As Long) As Integer
	Dim Bytes() As Byte
	Length = 0
	Bytes = GetBytes(FN,4,Index,Mode)
	If (Bytes(0) And &H80) = 0 Then
		CorSigUncompressData = 1
		Length = Bytes(0)
	ElseIf (Bytes(0) And &HC0) = &H80 Then
		CorSigUncompressData = 2
		Length = SHL((Bytes(0) And &H3F),8) Or Bytes(1)
	ElseIf (Bytes(0) And &HE0) = &HC0 Then
		CorSigUncompressData = 4
		Length = SHL((Bytes(0) And &H1F),24) Or SHL(CLng(Bytes(1)),16) Or SHL(CLng(Bytes(2)),8) Or Bytes(3)
	End If
	Exit Function

    If Bytes(0) = 0 Then
		CorSigUncompressData = 1
		Length = 0
	ElseIf ((Bytes(0) And &HC0) = &HC0) And ((Bytes(0) And &H20) = 0) Then
		CorSigUncompressData = 4
		Length = SHL((Bytes(0) And &H1F),24) Or SHL(CLng(Bytes(1)),16) Or SHL(CLng(Bytes(2)),8) Or Bytes(3)
	ElseIf ((Bytes(0) And &H80) = &H80) And ((Bytes(0) And &H40) = 0) Then
		CorSigUncompressData = 2
		Length = SHL((Bytes(0) And &H3F),8) Or Bytes(1)
	Else
		CorSigUncompressData = 1
		Length = Bytes(0) And &H7F
	End If
End Function


'计算 PE 文件对齐
Public Function Alignment(ByVal orgValue As Long,ByVal AlignVal As Long,ByVal RoundVal As Long) As Long
	If AlignVal < 1 Then
		Alignment = orgValue
	Else
		Alignment = IIf(orgValue Mod AlignVal = 0,orgValue,AlignVal * ((orgValue \ AlignVal) + RoundVal))
	End If
End Function


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
Public Function StrListJoin(StrArray() As String,Optional ByVal Sep As String = " ",Optional ByVal Mode As Boolean) As String
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


'复制字典为副本
Public Function CopyDic(ByVal OldDic As Object,ByVal NewDic As Object,Optional ByVal ShowMsg As Long) As Boolean
	Dim i As Long,Msg As String
	If OldDic Is Nothing Then Exit Function
	If NewDic Is Nothing Then
		Set NewDic = CreateObject("Scripting.Dictionary")
	Else
		NewDic.RemoveAll
	End If
	If OldDic.Count = 0 Then Exit Function
	ReDim Keys(0) As Variant,Items(0) As Variant
	Keys = OldDic.Keys: Items = OldDic.Items
	Select Case ShowMsg
	Case 0
		For i = 0 To OldDic.Count - 1
			NewDic.Add(Keys(i),Items(i))
		Next i
	Case Is > 0
		Msg = GetTextBoxString(ShowMsg) & " "
		For i = 0 To OldDic.Count - 1
			NewDic.Add(Keys(i),Items(i))
			If i Mod 100 = 1 Then
				SetTextBoxString ShowMsg,Msg & Format$(i / OldDic.Count,"#%")
			End If
		Next i
		SetTextBoxString ShowMsg,Msg & "100%"
	Case Is < 0
		ReDim TempList(PSL.OutputWnd(0).LineCount - 1) As String
		For i = 1 To PSL.OutputWnd(0).LineCount
			TempList(i - 1) = PSL.OutputWnd(0).Text(i)
		Next i
		Msg = StrListJoin(TempList,vbCrLf) & " "
		For i = 0 To OldDic.Count - 1
			NewDic.Add(Keys(i),Items(i))
			If i Mod 100 = 1 Then
				PSL.OutputWnd(0).Clear
				PSL.Output Msg & Format$(i / OldDic.Count,"#%")
			End If
		Next i
		PSL.OutputWnd(0).Clear
		PSL.Output Msg & "100%"
	End Select
	CopyDic = True
End Function


'获取字典指定键的值
Public Function GetDicVal(ByVal Dic As Object,ByVal Key As Variant,Optional ByVal DefaultValue As Variant) As Variant
	If Dic.Exists(Key) Then
		GetDicVal = Dic.Item(Key)
	Else
		GetDicVal = DefaultValue
	End If
End Function


'获取本地化语言及代码页信息
Public Function GetLCIDInfo(LCID As Long,LCType As Long) As String
	Dim iRet As Long
	iRet = GetLocaleInfo(LCID, LCType, "", 0)
	If iRet = 0 Then Exit Function
	GetLCIDInfo = String$(iRet, 0)
    If GetLocaleInfo(LCID, LCType, GetLCIDInfo, iRet) = 0 Then Exit Function
    GetLCIDInfo = Replace$(GetLCIDInfo,vbNullChar,"")
End Function


'获取字串中的子字符串(正则表达式)
'Patrn 为正则表达式模板
Public Function GetSubStringRegExp(ByVal textStr As String,ByVal Patrn As String,Optional ByVal MatchID As Long, _
		Optional ByVal SubMatchID As Long = -1,Optional ByVal bGlobal As Boolean,Optional ByVal IgnoreCase As Boolean) As String
	Dim Matches As Object
	On Error GoTo ExitFunction
	If Patrn = "" Then Exit Function
	If Trim$(textStr) = "" Then Exit Function
	With RegExp
		.Global = bGlobal
		.IgnoreCase = IgnoreCase
		.Pattern = Patrn
		Set Matches = .Execute(textStr)
	End With
	If Matches.Count = 0 Then Exit Function
	If MatchID < 0 Then MatchID = Matches.Count - 1
	If SubMatchID < 0 Then
		GetSubStringRegExp = Matches(MatchID).Value
	Else
		GetSubStringRegExp = Matches(MatchID).SubMatches(SubMatchID)
	End If
	ExitFunction:
End Function


'合并二个区间关系但无法匹配的 \x 转义符正则表达式为 [\x-\x] 形式的表达式
Private Function MergeRegExpPattern(ByVal lPattern As String,ByVal uPattern As String) As String
	Dim i As Long,j As Long,n As Long
	Dim lArray() As String,uArray() As String
	j = Len(lPattern)
	If j <> Len(uPattern) Then Exit Function
	MergeRegExpPattern = Space$(j * 11)
	n = 1
	For i = 1 To j Step 4
		Mid$(MergeRegExpPattern,n,11) = "[" & Mid$(lPattern,i,4) & "-" & Mid$(uPattern,i,4) & "]"
 		n = n + 11
	Next i
End Function


'输出程序错误消息
Public Sub sysErrorMassage(sysError As ErrObject,ByVal fType As Long)
	Dim TempArray() As String,MsgList() As String
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

	If CheckINIArray(UIDataList) = True Then
		If getMsgList(UIDataList,MsgList,"sysErrorMassage",3) = False Then
			If getMsgList(UIDataList,MsgList,"Main",3) = False Then
				Msg = "The following file is missing [Main|sysErrorMassage] section." & vbCrLf & "%s"
				Msg = Replace$(Msg,"%s",LangFile)
			Else
				TitleMsg = MsgList(7)
				If fType <> 0 Then ContinueMsg = MsgList(9) Else ContinueMsg = MsgList(10)
				Msg = Replace$(Replace$(MsgList(8),"%s","sysErrorMassage"),"%d",LangFile)
			End If
		Else
			TitleMsg = MsgList(0)
			Select Case fType
			Case 0
				ContinueMsg = MsgList(12)
			Case 1
				ContinueMsg = MsgList(13)
			Case 2
				ContinueMsg = MsgList(14)
			End Select

			Select Case ErrorSource
			Case ""
				If ErrorNumber = 10051 And PSL.Version >= 1500 Then
					Msg = Replace$(MsgList(15),"%s",ErrorSource)
				Else
					Msg = Replace$(Replace$(MsgList(1),"%d",CStr(ErrorNumber)),"%v",ErrorDescription)
				End If
			Case "NotSection"
				TempArray = ReSplit(ErrorDescription,JoinStr,-1)
				Msg = Replace$(Replace$(MsgList(3),"%s",TempArray(1)),"%d",TempArray(0))
			Case "NotValue"
				TempArray = ReSplit(ErrorDescription,JoinStr,-1)
				Msg = Replace$(Replace$(MsgList(4),"%s",TempArray(1)),"%d",TempArray(0))
			Case "NotReadFile"
				TempArray = ReSplit(ErrorDescription,JoinStr,-1)
				Msg = Replace$(MsgList(5),"%s",TempArray(1))
			Case "NotWriteFile"
				TempArray = ReSplit(ErrorDescription,JoinStr,-1)
				Msg = Replace$(MsgList(6),"%s",TempArray(1))
			Case "NotUnWriteFile"
				TempArray = ReSplit(ErrorDescription,JoinStr,-1)
				Msg = Replace$(MsgList(7),"%s",TempArray(1))
			Case "NotOpenFile"
				TempArray = ReSplit(ErrorDescription,JoinStr,-1)
				Msg = Replace$(MsgList(8),"%s",TempArray(1))
			Case "NotINIFile"
				Msg = Replace$(MsgList(9),"%s",ErrorDescription)
			Case "NotExitFile"
				Msg = Replace$(MsgList(10),"%s",ErrorDescription)
			Case "NotVersion"
				TempArray = ReSplit(ErrorDescription,JoinStr,-1)
				Msg = Replace$(MsgList(11),"%s",TempArray(0))
				Msg = Replace$(Replace$(Msg,"%d",TempArray(1)),"%v",TempArray(2))
			Case Else
				Msg = Replace$(MsgList(2),"%s",ErrorSource)
				Msg = Replace$(Replace$(Msg,"%d",CStr(ErrorNumber)),"%v",ErrorDescription)
			End Select
		End If
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
			TempArray = ReSplit(ErrorDescription,JoinStr,-1)
			Msg = "The following file is missing [%s] section." & vbCrLf & "%d"
			Msg = Replace$(Replace$(Msg,"%s",TempArray(1)),"%d",TempArray(0))
		Case "NotValue"
			TempArray = ReSplit(ErrorDescription,JoinStr,-1)
			Msg = "The following file is missing [%s] Value." & vbCrLf & "%d"
			Msg = Replace$(Replace$(Msg,"%s",TempArray(1)),"%d",TempArray(0))
		Case "NotReadFile"
			Msg = Replace$(ErrorDescription,JoinStr,vbCrLf)
		Case "NotWriteFile"
			Msg = Replace$(ErrorDescription,JoinStr,vbCrLf)
		Case "NotUnWriteFile"
			Msg = Replace$(ErrorDescription,JoinStr,vbCrLf)
		Case "NotOpenFile"
			Msg = Replace$(ErrorDescription,JoinStr,vbCrLf)
		Case "NotINIFile"
			Msg = "The following contents of the file is not correct." & vbCrLf & "%s"
			Msg = Replace$(Msg,"%s",ErrorDescription)
		Case "NotExitFile"
			Msg = "The following file does not exist! Please check and try again." & vbCrLf & "%s"
			Msg = Replace$(Msg,"%s",ErrorDescription)
		Case "NotVersion"
			TempArray = ReSplit(ErrorDescription,JoinStr,-1)
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
		MsgBox Msg & ContinueMsg,vbOkOnly+vbInformation,TitleMsg
		Call ExitMacro(1)
	Case 1
		If MsgBox(Msg & ContinueMsg,vbYesNo+vbInformation,TitleMsg) = vbNo Then
			Call ExitMacro(1)
		End If
	Case Else
		MsgBox Msg & ContinueMsg,vbOkOnly+vbInformation,TitleMsg
	End Select
End Sub


'安全退出程序
Public Sub ExitMacro(ByVal Mode As Long)
	Dim Temp As String
	On Error Resume Next
	If Mid$(LCase$(TargetFile.FilePath),InStrRev(LCase$(TargetFile.FilePath),".")) = ".bak" Then
		Temp = Left$(TargetFile.FilePath,InStrRev(LCase$(TargetFile.FilePath),".bak") - 1)
	Else
		Temp = TargetFile.FilePath
	End If
	If Dir$(SourceFile.FilePath & ".tmp") <> "" Then Kill SourceFile.FilePath & ".tmp"
	If Dir$(SourceFile.FilePath & ".xls") <> "" Then Kill SourceFile.FilePath & ".xls"
	If Dir$(Temp & ".tmp") <> "" Then Kill Temp & ".tmp"
	If Dir$(Temp & ".xls") <> "" Then Kill Temp & ".xls"
	If Mode > 0 Then Exit All
End Sub


'获取在字节数组中找到的匹配数组的列表(正则表达式方式)
'注意：StartPos、EndPos 均为以 0 开始的地址
Public Function GetVAListRegExp(ByVal StrText As String,ByVal Patrn As String,ByVal StartPos As Long) As String()
	Dim i As Long,Matches As Object
	On Error GoTo ExitFunction
	RegExp.Global = True
	RegExp.IgnoreCase = False
	RegExp.Pattern = Patrn
	Set Matches = RegExp.Execute(StrText)
	If Matches.Count = 0 Then Exit Function
	ReDim TempList(Matches.Count - 1) As String
	For i = 0 To Matches.Count - 1
		TempList(i) = CStr$(StartPos + Matches(i).FirstIndex)
	Next i
	GetVAListRegExp = TempList
	ExitFunction:
End Function


'正向跳到非单空字节位置，并返回非单空字节开始位置
Public Function getNotNullByte(FN As FILE_IMAGE,ByVal Offset As Long,ByVal Max As Long,ByVal Mode As Long) As Long
	If Offset + 1 > Max Then
		getNotNullByte = Offset
		Exit Function
	End If
	Do
		getNotNullByte = GetByte(FN,Offset,Mode)
		Offset = Offset + 1
	Loop Until getNotNullByte <> 0 Or Offset > Max
	getNotNullByte = Offset - 1
End Function


'正向跳到非单空字节位置，并返回非单空字节开始位置
Private Function getNotNullByteRegExp(FN As FILE_IMAGE,ByVal Offset As Long,ByVal Max As Long,ByVal Mode As Long) As Long
	If Offset + 1 > Max Then
		getNotNullByteRegExp = Offset
		Exit Function
	End If
	Dim Matches As Object,endPos As Long
	With RegExp
		.Global = False
		.IgnoreCase = False
		.Pattern = "[^\x00]+?"
		Do
			EndPos = IIf(Offset + 512 < Max,Offset + 512,Max)
			Set Matches = .Execute(ByteToString(GetBytes(FN,EndPos - Offset + 1,Offset,Mode),CP_ISOLATIN1))
			If Matches.Count > 0 Then
				Offset = Offset + Matches(0).FirstIndex
				Exit Do
			End If
			Offset = endPos
		Loop Until EndPos >= Max
	End With
	getNotNullByteRegExp = Offset
End Function


'逆向跳到非单空字节位置，并返回最后一个非单空字节结束位置
Public Function getNotNullByteRev(FN As FILE_IMAGE,ByVal Offset As Long,ByVal Min As Long,ByVal Mode As Long) As Long
	If Offset - 1 < Min Then
		getNotNullByteRev = Offset
		Exit Function
	End If
	Do
		getNotNullByteRev = GetByte(FN,Offset,Mode)
		Offset = Offset - 1
	Loop Until getNotNullByteRev <> 0 Or Offset < Min
	getNotNullByteRev = Offset + 1
End Function


'逆向跳到非单空字节位置，并返回最后一个非单空字节结束位置
Private Function getNotNullByteRevRegExp(FN As FILE_IMAGE,ByVal Offset As Long,ByVal Min As Long,ByVal Mode As Long) As Long
	If Offset - 1 < Min Then
		getNotNullByteRevRegExp = Offset
		Exit Function
	End If
	Dim Matches As Object,StartPos As Long
	With RegExp
		.Global = True
		.IgnoreCase = False
		.Pattern = "[^\x00]+"
		Do
			StartPos = IIf(Offset - 512 > Min,Offset - 512,Min)
			Set Matches = .Execute(ByteToString(GetBytes(FN,Offset - StartPos + 1,StartPos,Mode),CP_ISOLATIN1))
			If Matches.Count > 0 Then
				Offset = StartPos + Matches(Matches.Count - 1).FirstIndex + Matches(Matches.Count - 1).Length - 1
				Exit Do
			End If
			Offset = StartPos
		Loop Until Offset <= Min
	End With
	getNotNullByteRevRegExp = Offset
End Function


'正向跳到指定数量的空字节位置，并返回空字节开始位置，Bit 为最小空字节数
Public Function getNullByte(FN As FILE_IMAGE,ByVal Offset As Long,ByVal Max As Long,ByVal Mode As Long,ByVal Bit As Integer) As Long
	If Offset + 1 > Max Then
		getNullByte = Offset
		Exit Function
	End If
  	Do
		If GetByte(FN,Offset,Mode) = 0 Then
			getNullByte = getNullByte + 1
		Else
			getNullByte = 0
		End If
		Offset = Offset + 1
	Loop Until Offset > Max Or getNullByte = Bit
	getNullByte = Offset - 1
End Function


'逆向跳到指定数量的空字节位置，并返回最后一个单空字节结束位置，Bit 为最小空字节数
Private Function getNullByteRev(FN As FILE_IMAGE,ByVal Offset As Long,ByVal Min As Long,ByVal Mode As Long,ByVal Bit As Integer) As Long
	If Offset - 1 < Min Then
		getNullByteRev = Offset
		Exit Function
	End If
  	Do
		If GetByte(FN,Offset,Mode) = 0 Then
			getNullByteRev = getNullByteRev + 1
		Else
			getNullByteRev = 0
		End If
		Offset = Offset - 1
	Loop Until Offset < Min Or getNullByteRev = Bit
	getNullByteRev = Offset + 1
End Function


'正向跳到指定数量的结束符位置，CodePage 为代码页
'getEndByteRegExp：返回结束符开始位置 (fType = False) 或结束位置 + 1 (fType = True)
Public Function getEndByteRegExp(FN As FILE_IMAGE,ByVal Offset As Long,ByVal Max As Long,ByVal Mode As Long, _
		ByVal Pattern As String,Optional ByVal CodePage As Long = CP_UNKNOWN,Optional ByVal fType As Boolean) As Long
	getEndByteRegExp = Offset
	If Offset + 1 > Max Then Exit Function
	Dim Matches As Object,endPos As Long,Temp As String
	With RegExp
		.Global = False
		.IgnoreCase = False
		.Pattern = Pattern
		If CodePage > 0 Then
			.Pattern = Pattern
		Else
			.Pattern = Pattern & "{" & IIf(-CodePage < 2,-CodePage,-CodePage \ 2) & ",}"
			CodePage = CP_UNICODELITTLE
		End If
		Do
			endPos = IIf(Offset + 512 < Max,Offset + 512,Max)
			Set Matches = .Execute(ByteToString(GetBytes(FN,endPos - Offset + 1,Offset,Mode),CodePage))
			If Matches.Count > 0 Then
				With Matches(0)
					If fType = False Then
						Select Case CodePage
						Case CP_WESTEUROPE, CP_ISOLATIN1
							Offset = Offset + .FirstIndex
						Case CP_UNICODELITTLE, CP_UNICODEBIG
							Offset = Offset + .FirstIndex * 2
						Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
							Offset = Offset + .FirstIndex * 4
						Case Else
							Offset = Offset + .FirstIndex
							If Mode = 0 Then
								Offset = InByteRegExp(FN.ImageByte,StringToByte(.Value,CodePage),Offset,endPos) - 1
							Else
								Offset = Offset + InByteRegExp(GetBytes(FN,endPos - Offset + 1,Offset,Mode),StringToByte(.Value,CodePage)) - 1
							End If
						End Select
						getEndByteRegExp = Offset
						Exit Function
					ElseIf .FirstIndex = 0 Then
						Select Case CodePage
						Case CP_WESTEUROPE, CP_ISOLATIN1
							Offset = Offset + (.FirstIndex + .Length)
						Case CP_UNICODELITTLE, CP_UNICODEBIG
							Offset = Offset + (.FirstIndex + .Length) * 2
						Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
							Offset = Offset + (.FirstIndex + .Length) * 4
						Case Else
							Offset = Offset + StrHexLength(.Value,CodePage,Mode)
						End Select
						getEndByteRegExp = Offset
						If Offset < endPos Then Exit Function
					Else
						Exit Function
					End If
				End With
			ElseIf fType = True Then
				Exit Function
			End If
			Offset = endPos
		Loop Until endPos >= Max
	End With
	getEndByteRegExp = Offset
End Function


'逆向跳到指定数量的空字节位置，CodePage 为代码页
'getEndByteRevRegExp：返回结束符结束位置
Public Function getEndByteRevRegExp(FN As FILE_IMAGE,ByVal Offset As Long,ByVal Min As Long,ByVal Mode As Long, _
		ByVal Pattern As String,Optional ByVal CodePage As Long = CP_UNKNOWN) As Long
	If Offset - 1 < Min Then
		getEndByteRevRegExp = Offset
		Exit Function
	End If
	Dim Matches As Object,StartPos As Long
	With RegExp
		.Global = True
		.IgnoreCase = False
		.Pattern = Pattern
		If CodePage > 0 Then
			.Pattern = Pattern
		Else
			.Pattern = Pattern & "{" & IIf(-CodePage < 2,-CodePage,-CodePage \ 2) & ",}"
			CodePage = CP_UNICODELITTLE
		End If
		Do
			StartPos = IIf(Offset - 512 > Min,Offset - 512,Min)
			Set Matches = .Execute(ByteToString(GetBytes(FN,Offset - StartPos + 1,StartPos,Mode),CodePage))
			If Matches.Count > 0 Then
				With Matches(Matches.Count - 1)
					Select Case CodePage
					Case CP_WESTEUROPE, CP_ISOLATIN1
						Offset = StartPos + (.FirstIndex + .Length) - 1
					Case CP_UNICODELITTLE, CP_UNICODEBIG
						Offset = StartPos + (.FirstIndex + .Length) * 2 - 1
					Case CP_UTF32LE, CP_UTF32BE, CP_UTF_32LE, CP_UTF_32BE
						Offset = StartPos + (.FirstIndex + .Length) * 4 - 1
					Case Else
						StartPos = StartPos + (.FirstIndex + .Length) - 1
						If Mode = 0 Then
							Offset = InByteRegExp(FN.ImageByte,StringToByte(.Value,CodePage),StartPos,Offset) - 1
						Else
							Offset = StartPos + InByteRegExp(GetBytes(FN,Offset - StartPos + 1,StartPos,Mode),StringToByte(.Value,CodePage)) - 1
						End If
					End Select
				End With
				Exit Do
			End If
			Offset = StartPos
		Loop Until Offset <= Min
	End With
	getEndByteRevRegExp = Offset
End Function


' 检查文件编码
' ----------------------------------------------------
' ANSI      无格式定义
' 2B2F 76[38|39|2B|2F] UTF-7
' EFBB BF   UTF-8
' FFFE      UTF-16LE/UCS-2, Little Endian with BOM
' FEFF      UTF-16BE/UCS-2, Big Endian with BOM
' XX00 XX00 UTF-16LE/UCS-2, Little Endian without BOM
' 00XX 00XX UTF-16BE/UCS-2, Big Endian without BOM
' FFFE 0000 UTF-32LE/UCS-4, Little Endian with BOM
' 0000 FEFF UTF-32BE/UCS-4, Big Endian with BOM
' XX00 0000 UTF-32LE/UCS-4, Little Endian without BOM
' 0000 00XX UTF-32BE/UCS-4, Big Endian without BOM
' 上述中的 XX 表示任意十六进制字符

Private Function CheckCode(ByVal FilePath As String) As String
	Dim i As Long,j As Long,objStream As Object,Bytes() As Byte,Temp As String

	If Dir$(FilePath) = "" Then Exit Function
	i = FileLen(FilePath)
	If i = 0 Then
		CheckCode = "ANSI"
		Exit Function
	End If
	On Error Resume Next
	Set objStream = CreateObject("Adodb.Stream")
	On Error GoTo 0
	If objStream Is Nothing Then
		CheckCode = "ANSI"
		Exit Function
	End If
	If i > 1 Then
		With objStream
			.Type = 1
			.Mode = 3
			.Open
			.Position = 0
			.LoadFromFile FilePath
			Bytes = .read(IIf(i > 3,4,i))
			.Close
		End With
		If i > 2 Then
			If Bytes(0) = 239 And Bytes(1) = 187 And Bytes(2) = 191 Then
				CheckCode = "utf-8EFBB"
			ElseIf Bytes(0) = &HFF And Bytes(1) = &HFE Then
				CheckCode = "unicodeFFFE"
			ElseIf Bytes(0) = &HFE And Bytes(1) = &HFF Then
				CheckCode = "unicodeFEFF"
			End If
		Else
			If Bytes(0) = &HFF And Bytes(1) = &HFE Then
				CheckCode = "unicodeFFFE"
			ElseIf Bytes(0) = &HFE And Bytes(1) = &HFF Then
				CheckCode = "unicodeFEFF"
			End If
		End If
		If i > 3 Then
			If Bytes(0) <> &H00 And Bytes(1) = &H00 And Bytes(2) <> &H00 And Bytes(3) = &H00 Then
				CheckCode = "utf-16"
			ElseIf Bytes(0) = &H00 And Bytes(1) <> &H00 And Bytes(2) = &H00 And Bytes(3) <> &H00 Then
				CheckCode = "utf-16BE"
			ElseIf Bytes(0) = &HFF And Bytes(1) = &HFE And Bytes(2) = &H00 And Bytes(3) = &H00 Then
				CheckCode = "unicode-32FFFE"
			ElseIf Bytes(0) = &H00 And Bytes(1) = &H00 And Bytes(2) = &HFE And Bytes(3) = &HFF Then
				CheckCode = "unicode-32FEFF"
			ElseIf Bytes(0) = &H00 And Bytes(1) = &H00 And Bytes(2) = &H00 And Bytes(3) <> &H00 Then
				CheckCode = "utf-32"
			ElseIf Bytes(0) <> &H00 And Bytes(1) = &H00 And Bytes(2) = &H00 And Bytes(3) = &H00 Then
				CheckCode = "utf-32BE"
			End If
		End If
		If CheckCode <> "" Then
			Set objStream = Nothing
			Exit Function
		End If
	End If
	On Error GoTo ExitFunction
	With objStream
		.Type = 2
		.Mode = 3
		.Open
		.CharSet = "_autodetect_all"
		.Position = 0
		.LoadFromFile FilePath
		Temp = .ReadText(IIf(i > 10240,10240,i))
		'On Error GoTo NextNum
		For j = 38 To 2 Step -1
			If j <> 9 And j <> 13 Then
				.Position = 0
				.CharSet = CodeList(j).CharSet
				If Temp = .ReadText(IIf(i > 10240,10240,i)) Then
					CheckCode = .CharSet
					.Close
					Set objStream = Nothing
					Exit Function
				End If
			End If
			'NextNum:
		Next j
		'On Error GoTo NextPos
		For j = 41 To 39 Step -1
			If j <> 40 Then
				.Position = 0
				.CharSet = CodeList(j).CharSet
				If Temp = .ReadText(IIf(i > 10240,10240,i)) Then
					CheckCode = .CharSet
					.Close
					Set objStream = Nothing
					Exit Function
				End If
			End If
			'NextPos:
		Next j
		.Close
	End With
	ExitFunction:
	If CheckCode = "" Then CheckCode = "ANSI"
	Set objStream = Nothing
End Function


' 读取文本文件
Public Function ReadTextFile(ByVal FilePath As String,CharSet As String) As String
	Dim i As Long,objStream As Object,Code As String,FN As Variant
	If Dir$(FilePath) = "" Then Exit Function
	i = FileLen(FilePath)
	If i = 0 Then Exit Function
	On Error Resume Next
	Set objStream = CreateObject("Adodb.Stream")
	On Error GoTo ErrorMsg
	Code = CharSet
	If Not objStream Is Nothing Then
		If Code = "" Then Code = CheckCode(FilePath)
		If Code = "" Then Code = "_autodetect_all"
		If Code <> "ANSI" Then
			With objStream
				.Type = 2
				.Mode = 3
				.CharSet = IIf(Code = "utf-8EFBB","utf-8",Code)
				.Open
				.LoadFromFile FilePath
				ReadTextFile = .ReadText
				.Close
			End With
		End If
	End If
	If objStream Is Nothing Or Code = "ANSI" Then
		ReDim Bytes(i - 1) As Byte
		FN = FreeFile
		Open FilePath For Binary Access Read Lock Write As #FN
		Get #FN,,Bytes
		Close #FN
		ReadTextFile = StrConv$(Bytes,vbUnicode)
	End If
	If CharSet = "" Then CharSet = Code
	Set objStream = Nothing
	Exit Function

	ErrorMsg:
	ReadTextFile = ""
	If objStream Is Nothing Or Code = "ANSI" Then Close #FN
	Set objStream = Nothing
	Err.Source = "NotReadFile"
	Err.Description = Err.Description & JoinStr & FilePath
	Call sysErrorMassage(Err,1)
End Function


' 写入文本文件
Public Function WriteToFile(ByVal FilePath As String,ByVal textStr As String,CharSet As String) As Boolean
	Dim objStream As Object,Code As String,Bytes() As Byte,FN As Variant
	If FilePath = "" Then Exit Function
	On Error Resume Next
	Set objStream = CreateObject("Adodb.Stream")
	On Error GoTo ErrorMsg
	Code = CharSet
	If Not objStream Is Nothing Then
		If Code = "" Then Code = CheckCode(FilePath)
		If LCase$(Code) = "_autodetect_all" Then Code = "ANSI"
		If Code <> "ANSI" Then
			With objStream
				.Type = 2
				.Mode = 3
				.CharSet = IIf(Code = "utf-8EFBB","utf-8",Code)
				.Open
				.WriteText textStr
				'去除不带 BOM 格式的 BOM
				If Code = "utf-16LE" Or Code = "utf-8" Then
					.Position = 0
					.Type = 1
					.Position = IIf(Code = "utf-16LE",2,3)
					Bytes = .Read(.Size - IIf(Code = "utf-16LE",2,3))
					.Position = 0
					.SetEOS
					.Write Bytes
				End If
				.SaveToFile FilePath,2
				.Close
			End With
			WriteToFile = True
		End If
	End If
	If objStream Is Nothing Or Code = "ANSI" Then
		Bytes = StrConv(textStr,vbFromUnicode)
		FN = FreeFile
		Open FilePath For Binary Access Write Lock Write As #FN
		Put #FN,,Bytes
		Close #FN
		WriteToFile = True
	End If
	If CharSet = "" Then CharSet = Code
	Set objStream = Nothing
	Exit Function

	ErrorMsg:
	If objStream Is Nothing Or Code = "ANSI" Then Close #FN
	Set objStream = Nothing
	Err.Source = "NotWriteFile"
	Err.Description = Err.Description & JoinStr & FilePath
	Call sysErrorMassage(Err,1)
End Function


'创建 Adodb.Stream 使用的代码页数组
Public Function getCodePageList(ByVal MinNum As Long,ByVal MaxNum As Long) As CODEPAGE_DATA()
	Dim i As Long,MsgList() As String,Code As String
	Dim CharSetList() As String,CodePage() As CODEPAGE_DATA

	If getMsgList(UIDataList,MsgList,"CodePageList",0) = False Then Exit Function
	Code = "ANSI|_autodetect_all|gb2312|hz-gb-2312|gb18030|big5|euc-jp|iso-2022-jp|shift_jis|" & _
			"_autodetect|ks_c_5601-1987|euc-kr|iso-2022-kr|_autodetect_kr|windows-874|" & _
			"windows-1258|iso-8859-4|windows-1257|ASMO-708|DOS-720|iso-8859-6|windows-1256|" & _
			"DOS-862|iso-8859-8-i|iso-8859-8|windows-1255|iso-8859-9|iso-8859-7|windows-1253|" & _
			"iso-8859-1|cp866|iso-8859-5|koi8-r|koi8-ru|windows-1251|ibm852|iso-8859-2|" & _
			"windows-1250|iso-8859-3|utf-7|utf-8EFBB|utf-8|unicodeFFFE|unicodeFEFF|utf-16|" & _
			"utf-16BE|unicode-32FFFE|unicode-32FEFF|utf-32|utf-32BE"
	CharSetList = ReSplit(Code,"|")

	i = UBound(MsgList)
	If MaxNum > i Then MaxNum = i
	ReDim CodePage(MaxNum - MinNum) As CODEPAGE_DATA
	For i = MinNum To MaxNum
		CodePage(i - MinNum).sName = MsgList(i)
		CodePage(i - MinNum).CharSet = CharSetList(i)
	Next i
	getCodePageList = CodePage
End Function


'从来源列表中获取指定项目的目标列表
'Mode = 0 发生错误直接退出程序，Mode = 1 发生错误给出是否退出程序提示
Public Function getMsgList(SourceList() As INIFILE_DATA,TargetList() As String,Titles As String,ByVal Mode As Long) As Boolean
	Dim i As Long,j As Long,n As Long,m As Long,Temp As String
	Temp = "|" & Titles & "|"
	n = -1
	For i = 0 To UBound(SourceList)
		If InStr(Temp,"|" & SourceList(i).Title & "|") > 0 Then
			If n = -1 Then
				TargetList = SourceList(i).Value
				n = UBound(TargetList)
			Else
				m = UBound(SourceList(i).Value)
				ReDim Preserve TargetList(n + m + 1) As String
				For j = 0 To m
					TargetList(n + 1) = SourceList(i).Value(j)
				Next j
			End If
			Temp = Replace$(Temp,"|" & SourceList(i).Title & "|","")
			If Temp = "" Then Exit For
		End If
	Next i
	If n > -1 Then
		Titles = RemoveBackslash(Temp,"|","|",1)
		getMsgList = True
	Else
		If Mode = 0 Then
			Err.Raise(1,"NotSection",LangFile & JoinStr & Titles)
		ElseIf Mode < 3 Then
			On Error GoTo ErrorMassage
			Err.Raise(1,"NotSection",LangFile & JoinStr & Titles)
			ErrorMassage:
			Call sysErrorMassage(Err,Mode)
		End If
	End If
End Function


'写入设置
Public Function WriteSettings(ByVal WriteType As String) As Boolean
	Dim i As Long,n As Long,Temp As String,TempArray() As String
	On Error GoTo ExitFunction
	SaveSetting(AppName,"Option","Version",Version)
	'保存界面语言设置
	If WriteType = "Sets" Or WriteType = "All" Then
		SaveSetting(AppName,"Option","UILanguageID",Selected(0))
		SaveSetting(AppName,"Option","OpenFileMode",Selected(1))
		SaveSetting(AppName,"Option","MoveToStrFreeByte",Selected(2))
		SaveSetting(AppName,"Option","MoveToNotStrFreeByte",Selected(3))
		SaveSetting(AppName,"Option","MoveToSectionEndFreeByte",Selected(4))
		SaveSetting(AppName,"Option","MoveToAddSectionEndFreeByte",Selected(5))
		SaveSetting(AppName,"Option","MoveToAddLastSectionEnd",Selected(6))
		SaveSetting(AppName,"Option","MoveToNewSection",Selected(7))
		SaveSetting(AppName,"Option","MoveOnOff",Selected(8))
		SaveSetting(AppName,"Option","NoMarkPEFreeByteMinLength",Selected(9))
		'SaveSetting(AppName,"Option","StringEndOriginalFreeByte",Selected(10))
		SaveSetting(AppName,"Option","WriteOrder",Selected(11))
		SaveSetting(AppName,"Option","OptimizateFileSize",Selected(12))
		SaveSetting(AppName,"Option","WriteLogFile",Selected(13))
		SaveSetting(AppName,"Option","OpenLogFile",Selected(14))
		SaveSetting(AppName,"Option","TranStrOnlyMsgShow",Selected(15))
		SaveSetting(AppName,"Option","StrAddDisplayFormat",Selected(16))
		SaveSetting(AppName,"Option","MsgOutputWindows",Selected(17))
		SaveSetting(AppName,"Option","StrFreeByteMatchLoops",Selected(18))
		SaveSetting(AppName,"Option","SpaceWithPointCheckBox",Selected(19))
		SaveSetting(AppName,"Option","FilterDisplayAtStartupBox",Selected(20))
		SaveSetting(AppName,"Option","GetStringVoiceCheckBox",Selected(21))
		SaveSetting(AppName,"Option","ImportSrcVoiceCheckBox",Selected(22))
		SaveSetting(AppName,"Option","ImportTrnVoiceCheckBox",Selected(23))
		SaveSetting(AppName,"Option","ParseStrDataVoiceCheckBox",Selected(24))
		SaveSetting(AppName,"Option","GetStrTypeVoiceCheckBox",Selected(25))
		SaveSetting(AppName,"Option","MoveStringVoiceCheckBox",Selected(26))
		SaveSetting(AppName,"Option","AddAndDelFilterVoiceCheckBox",Selected(27))
		SaveSetting(AppName,"Option","AddAndDelReserveVoiceCheckBox",Selected(28))
		SaveSetting(AppName,"Option","WriteStringVoiceCheckBox",Selected(29))
	End If
	'保存字串提取选项
	If WriteType = "GetString" Or WriteType = "All" Then
		SaveSetting(AppName,"GetString","ExtractModeOption",ExtractSet(0))
		'SaveSetting(AppName,"GetString","ANSICheckBox",ExtractSet(1))
		'SaveSetting(AppName,"GetString","UnicodeCheckBox",ExtractSet(2))
		'SaveSetting(AppName,"GetString","UTF8CheckBox",ExtractSet(3))
		SaveSetting(AppName,"GetString","CharEncodeOrder",ExtractSet(4))
		SaveSetting(AppName,"GetString","ReferenceCheckBox",ExtractSet(11))
		SaveSetting(AppName,"GetString","NullStrCheckBox",ExtractSet(12))
		SaveSetting(AppName,"GetString","ControlStrCheckBox",ExtractSet(13))
		SaveSetting(AppName,"GetString","AllSymbolCheckBox",ExtractSet(14))
		SaveSetting(AppName,"GetString","AllUCaseCheckBox",ExtractSet(15))
		SaveSetting(AppName,"GetString","AllLCaseCheckBox",ExtractSet(16))
		SaveSetting(AppName,"GetString","AllULCaseCheckBox",ExtractSet(17))
		SaveSetting(AppName,"GetString","ContinuousCheckBox",ExtractSet(18))
		SaveSetting(AppName,"GetString","OtherStrCheckBox",ExtractSet(19))
		SaveSetting(AppName,"GetString","IncludeStrCheckBox",ExtractSet(20))
		SaveSetting(AppName,"GetString","FilterStrCheckBox",ExtractSet(21))
		SaveSetting(AppName,"GetString","IgnoreFilterAccKeyCheckBox",ExtractSet(22))
		SaveSetting(AppName,"GetString","KeepStrCheckBox",ExtractSet(23))
		SaveSetting(AppName,"GetString","IgnoreKeepAccKeyCheckBox",ExtractSet(24))
		'SaveSetting(AppName,"GetString","ExcludeStrCheckBox",ExtractSet(23))
		'SaveSetting(AppName,"GetString","IgnoreAccKeyCheckBox",ExtractSet(24))
		SaveSetting(AppName,"GetString","OutputMsgCheckBox",ExtractSet(25))
		SaveSetting(AppName,"GetString","LogingMsgCheckBox",ExtractSet(26))
		SaveSetting(AppName,"GetString","MinStrLengthTextBox",ExtractSet(27))
		SaveSetting(AppName,"GetString","ContinuousTextBox",ExtractSet(28))
		SaveSetting(AppName,"GetString","OtherStrTextBox",ExtractSet(29))
		SaveSetting(AppName,"GetString","IncludeStrTextBox",ExtractSet(30))
		SaveSetting(AppName,"GetString","EndCharOfString",ConvertStrEndCharSet(MergeStrEndCharSet(ExtractSet(31)),True))
		SaveSetting(AppName,"GetString","OtherStrMatchMode",ExtractSet(42))
		SaveSetting(AppName,"GetString","OtherStrMatchCaseCheckBox",ExtractSet(43))
		SaveSetting(AppName,"GetString","IgnoreOtherAccKeyCheckBox",ExtractSet(44))
		SaveSetting(AppName,"GetString","OtherStrForWildcardTextBox",ExtractSet(45))
		SaveSetting(AppName,"GetString","OtherStrForRegExpTextBox",ExtractSet(46))
		SaveSetting(AppName,"GetString","RangeCheckBox",ExtractSet(49))
		SaveSetting(AppName,"GetString","LogingFilteredStrOnlyCheckBox",ExtractSet(51))
		SaveSetting(AppName,"GetString","MaxStrLengthCheckBox",ExtractSet(52))
		SaveSetting(AppName,"GetString","MaxStrLengthTextBox",ExtractSet(53))
	End If
	'保存自动更新设置
	If WriteType = "Update" Or WriteType = "Sets" Or WriteType = "All" Then
		If CheckArray(UpdateSet) = True Then
			On Error Resume Next
			DeleteSetting(AppName,"Update")
			On Error GoTo 0
			TempArray = ReSplit(UpdateSet(1),vbCrLf,-1)
			SaveSetting(AppName,"Update","UpdateMode",UpdateSet(0))
			n = 0
			For i = 0 To UBound(TempArray)
				If Trim$(TempArray(i)) <> "" Then
					SaveSetting(AppName,"Update",CStr(n),TempArray(i))
					n = n + 1
				End If
			Next i
			If n > 0 Then SaveSetting(AppName,"Update","Count",n - 1)
			SaveSetting(AppName,"Update","Path",UpdateSet(2))
			SaveSetting(AppName,"Update","Argument",UpdateSet(3))
			SaveSetting(AppName,"Update","UpdateCycle",UpdateSet(4))
			SaveSetting(AppName,"Update","UpdateDate",UpdateSet(5))
		End If
	End If
	'保存自定义字串类型
	If WriteType = "Sets" Or WriteType = "All" Then
		'删除原配置项
		Temp = GetSetting(AppName,"Option","StringTypeCount","")
		If Temp <> "" Then
			On Error Resume Next
			For i = 0 To StrToLong(Temp)
				DeleteSetting(AppName,"StringType_" & CStr$(i))
			Next i
			On Error GoTo 0
		End If
		'写入新配置项
		If CheckStrTypeArray(StrTypeList) = True Then
			For i = 0 To UBound(StrTypeList)
				Temp = "StringType_" & CStr$(i)
				With StrTypeList(i)
					SaveSetting(AppName,Temp,"Name",.sName)
					SaveSetting(AppName,Temp,"CodeLocation",.CodeLoc)
					SaveSetting(AppName,Temp,"FristCodePos",.FristCodePos)
					SaveSetting(AppName,Temp,"CodePageCodePos",.CPCodePos)
					SaveSetting(AppName,Temp,"CodePageCodeSize",.CPCodeSize)
					SaveSetting(AppName,Temp,"CodePageCodeStartString",.CPCodeStartString)
					SaveSetting(AppName,Temp,"LengthCodePos",.LengthCodePos)
					SaveSetting(AppName,Temp,"LengthCodeSize",.LengthCodeSize)
					SaveSetting(AppName,Temp,"LengthReviseVal",.LengthReviseVal)
					SaveSetting(AppName,Temp,"ByteLengthReviseVal",.ByteLengthReviseVal)
					SaveSetting(AppName,Temp,"CharLengthReviseVal",.CharLengthReviseVal)
					SaveSetting(AppName,Temp,"LengthMode",.LengthMode)
					SaveSetting(AppName,Temp,"LengthCodeStartString",.LengthCodeStartString)
					SaveSetting(AppName,Temp,"StartCodePos",.StartCodePos)
					SaveSetting(AppName,Temp,"StartCodeString",.StartCodeString)
					SaveSetting(AppName,Temp,"EndCodeString",.EndCodeString)
					SaveSetting(AppName,Temp,"RefCodeStartPos",.RefCodeStartPos)
					SaveSetting(AppName,Temp,"RefCodeStartString",.RefCodeStartString)
				End With
			Next i
			SaveSetting(AppName,"Option","StringTypeCount",UBound(StrTypeList))
		End If
	End If
	'保存自定义引用算法
	If WriteType = "Sets" Or WriteType = "All" Then
		'删除原配置项
		Temp = GetSetting(AppName,"Option","RefAlgorithmCount","")
		If Temp <> "" Then
			On Error Resume Next
			For i = 0 To StrToLong(Temp)
				DeleteSetting(AppName,"RefAlgorithm_" & CStr$(i))
			Next i
			On Error GoTo 0
		End If
		'写入新配置项
		If CheckRefTypeArray(RefTypeList) = True Then
			For i = 0 To UBound(RefTypeList)
				Temp = "RefAlgorithm_" & CStr$(i)
				With RefTypeList(i)
					SaveSetting(AppName,Temp,"Name",.sName)
					SaveSetting(AppName,Temp,"Algorithm",.Algorithm)
					SaveSetting(AppName,Temp,"ByteLength",.ByteLength)
					SaveSetting(AppName,Temp,"ByteOrder",.ByteOrder)
					SaveSetting(AppName,Temp,"PrefixByteRegExp",.PrefixByte)
					SaveSetting(AppName,Temp,"PrefixByteLength",.PrefixLength)
					SaveSetting(AppName,Temp,"StrAddAlgorithm",.StrAddAlgorithm)
				End With
			Next i
			SaveSetting(AppName,"Option","RefAlgorithmCount",UBound(RefTypeList))
		End If
	End If
	'保存自定义工具
	If WriteType = "Tools" Or WriteType = "All" Then
		On Error Resume Next
		DeleteSetting(AppName,"Tools")
		On Error GoTo 0
		If UBound(Tools) > 3 Then
			n = 0
			For i = 4 To UBound(Tools)
				If Tools(i).sName <> "" And Tools(i).FilePath <> "" Then
					SaveSetting(AppName,"Tools",CStr$(n) & "_Name",Tools(i).sName)
					SaveSetting(AppName,"Tools",CStr$(n) & "_Path",Tools(i).FilePath)
					SaveSetting(AppName,"Tools",CStr$(n) & "_Argument",Tools(i).Argument)
					n = n + 1
				End If
			Next i
			If n > 0 Then SaveSetting(AppName,"Tools","Count",n - 1)
		End If
	End If
	'保存自定义代码页
	If WriteType = "Sets" Or WriteType = "All" Then
		'删除原配置项
		On Error Resume Next
		DeleteSetting(AppName,"Languages")
		On Error GoTo 0
		'写入新配置项
		If UBound(UniLangList) > 0 Then
			n = 0
			For i = 0 To UBound(UniLangList)
				With UniLangList(i)
					If .UniCodeRange <> "" Then
						Temp = CStr$(n)
						If .dwFlags = True Then
							SaveSetting(AppName,"Languages",Temp & "_LanguageName",.LangName)
							SaveSetting(AppName,"Languages",Temp & "_CodePageName",.CPName)
						End If
						SaveSetting(AppName,"Languages",Temp & "_LanguageID",.LangID)
						SaveSetting(AppName,"Languages",Temp & "_CodePage",.CodePage)
						SaveSetting(AppName,"Languages",Temp & "_UniCodeRange",.UniCodeRange)
						SaveSetting(AppName,"Languages",Temp & "_FeatureCodeRange",.FeatureCode)
						SaveSetting(AppName,"Languages",Temp & "_FeatureCodeEnable",.FeatureCodeEnable)
						n = n + 1
					End If
				End With
			Next i
		End If
	End If
	'保存过滤字串列表设置
	If WriteType = "FilterStrDic" Or WriteType = "All" Then
		SaveSetting(AppName,"GetString","UseMatchFilterListFile",ExtractSet(34))
		SaveSetting(AppName,"GetString","FilterListFileSelect",ExtractSet(39))
		SaveSetting(AppName,"GetString","FilterListFile",Mid$(ExtractSet(35),InStrRev(ExtractSet(35),"\") + 1))
	End If
	'保存保留字串列表设置
	If WriteType = "ReserveStrDic" Or WriteType = "All" Then
		SaveSetting(AppName,"GetString","UseMatchReserveListFile",ExtractSet(36))
		SaveSetting(AppName,"GetString","ReserveListFileSelect",ExtractSet(40))
		SaveSetting(AppName,"GetString","ReserveListFile",Mid$(ExtractSet(37),InStrRev(ExtractSet(37),"\") + 1))
	End If
	'保存对话框字体设置
	If WriteType = "DlgFont" Or WriteType = "Sets" Then
		On Error Resume Next
		DeleteSetting(AppName,"DlgFonts")
		On Error GoTo 0
		For i = 0 To UBound(LFList)
			Select Case i
			Case 0
				Temp = "MainFont"
			Case 1
				Temp = "SrcStrFont"
			Case 2
				Temp = "TrnStrFont"
			End Select
			If CheckFont(LFList(i)) = True Then
				With LFList(i)
					SaveSetting(AppName,"DlgFonts",Temp & "_lfCharSet",CStr$(.lfCharSet))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfClipPrecision",CStr$(.lfClipPrecision))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfEscapement",CStr$(.lfEscapement))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfFaceName",ReSplit(StrConv$(.lfFaceName,vbUnicode),vbNullChar,2)(0))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfHeight",CStr$(.lfHeight))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfItalic",CStr$(.lfItalic))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfOrientation",CStr$(.lfOrientation))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfOutPrecision",CStr$(.lfOutPrecision))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfPitchAndFamily",CStr$(.lfPitchAndFamily))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfQuality",CStr$(.lfQuality))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfStrikeOut",CStr$(.lfStrikeOut))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfUnderline",CStr$(.lfUnderline))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfWeight",CStr$(.lfWeight))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfWidth",CStr$(.lfWidth))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfColor",CStr$(.lfColor))
				End With
			Else
				ReDim tmpLFList(0) As LOG_FONT
				LFList(i) = tmpLFList(0)
			End If
		Next i
	End If
	WriteSettings = True
	ExitFunction:
End Function


'清理字串数组中重复数据
'Mode = False 不清除空置项，否则清除空置项
Public Function ClearTextArray(MyArray() As String,Optional ByVal Mode As Boolean) As String()
	Dim i As Long,n As Long,Dic As Object
	ClearTextArray = MyArray
	If CheckArrEmpty(MyArray) = False Then Exit Function
	i = UBound(MyArray)
	If i = 0 Then Exit Function
	ReDim TempArray(i) As String
	Set Dic = CreateObject("Scripting.Dictionary")
	If Mode = False Then
		For i = LBound(MyArray) To UBound(MyArray)
			If Not Dic.Exists(MyArray(i)) Then
				Dic.Add(MyArray(i),"")
				TempArray(n) = MyArray(i)
				n = n + 1
			End If
		Next i
	Else
		For i = LBound(MyArray) To UBound(MyArray)
			If MyArray(i) <> "" Then
				If Not Dic.Exists(MyArray(i)) Then
					Dic.Add(MyArray(i),"")
					TempArray(n) = MyArray(i)
					n = n + 1
				End If
			End If
		Next i
	End If
	Set Dic = Nothing
	If n = 0 Then Exit Function
	ReDim Preserve TempArray(n - 1) As String
	ClearTextArray = TempArray
End Function


'清理数值数组中重复数据
'Mode = False 不清除为零项，否则清除为零项
Public Function ClearDecArray(MyArray() As Long,Optional ByVal Mode As Boolean) As Long()
	Dim i As Long,n As Long,Dic As Object
	ClearDecArray = MyArray
	If CheckArrEmpty(MyArray) = False Then Exit Function
	i = UBound(MyArray)
	If i = 0 Then Exit Function
	ReDim TempArray(i) As Long
	Set Dic = CreateObject("Scripting.Dictionary")
	If Mode = False Then
		For i = LBound(MyArray) To UBound(MyArray)
			If Not Dic.Exists(MyArray(i)) Then
				Dic.Add(MyArray(i),"")
				TempArray(n) = MyArray(i)
				n = n + 1
			End If
		Next i
	Else
		For i = LBound(MyArray) To UBound(MyArray)
			If MyArray(i) <> 0 Then
				If Not Dic.Exists(MyArray(i)) Then
					Dic.Add(MyArray(i),"")
					TempArray(n) = MyArray(i)
					n = n + 1
				End If
			End If
		Next i
	End If
	Set Dic = Nothing
	If n = 0 Then Exit Function
	ReDim Preserve TempArray(n - 1) As Long
	ClearDecArray = TempArray
End Function


'读取 INI 文件
'Mode = 0 删除项目值前后空格及双引号，并转义项目值
'Mode = 1 删除项目值前空格，不转义
'Mode = 2 删除项目值前后空格，不转义
'Mode > 2 不删除项目值前后空格，不转义
Public Function getINIFile(DataList() As INIFILE_DATA,ByVal File As String,Code As String,ByVal Mode As Long) As Boolean
	Dim i As Long,j As Long,m As Long,n As Long,iMax As Long,vMax As Long,TempArray() As String
	If Trim$(File) <> "" Then
		TempArray = ReSplit(ReadTextFile(File,Code),vbCrLf)
	ElseIf File = "" And Code <> "" Then
		TempArray = ReSplit(Code,vbCrLf)
	Else
		Exit Function
	End If
	If CheckArray(TempArray) = False Then Exit Function
	ReDim DataList(iMax) As INIFILE_DATA
	m = -1
	For i = 0 To UBound(TempArray)
		TempArray(i) = Trim$(TempArray(i))
		If TempArray(i) Like "[[]*]" Then
			m = m + 1
			If m >= iMax Then
				iMax = m * 50
				ReDim Preserve DataList(iMax) As INIFILE_DATA
			End If
			DataList(m).Title = Trim$(Mid$(TempArray(i),2,Len(TempArray(i)) - 2))
			If m > 0 Then
				If n > 0 Then n = n - 1
				ReDim Preserve DataList(m - 1).Item(n) 'As String
				ReDim Preserve DataList(m - 1).Value(n) 'As String
			End If
			n = 0: vMax = 0
		ElseIf DataList(IIf(m = -1,0,m)).Title <> "" Then
			j = InStr(TempArray(i),"=")
			If j > 0 Then
				If Trim$(Left$(TempArray(i),j - 1)) <> "" Then
					If n >= vMax Then
						vMax = n * 100
						ReDim Preserve DataList(m).Item(vMax) 'As String
						ReDim Preserve DataList(m).Value(vMax) 'As String
					End If
					DataList(m).Item(n) = Trim$(Left$(TempArray(i),j - 1))
					Select Case Mode
					Case 0
						DataList(m).Value(n) = Convert(RemoveBackslash(Mid$(TempArray(i),j + 1),"""","""",2))
					Case 1
						DataList(m).Value(n) = LTrim$(Mid$(TempArray(i),j + 1))
					Case 2
						DataList(m).Value(n) = Trim$(Mid$(TempArray(i),j + 1))
					Case Else
						DataList(m).Value(n) = Mid$(TempArray(i),j + 1)
					End Select
					n = n + 1
				End If
			End If
		End If
	Next i
	If m < 0 Then m = 0
	If n > 0 Then
		n = n - 1
		getINIFile = True
	End If
	ReDim Preserve DataList(m).Item(n) 'As String
	ReDim Preserve DataList(m).Value(n) 'As String
	ReDim Preserve DataList(m) As INIFILE_DATA
End Function


'分割语言数组的名称和值
'GetType = 0 获取语言名称列表
'GetType = 1 获取语言 ID 列表
'GetType = 2 获取代码页值列表 (相同的代码页被过滤)
'GetType = 3 获取代码页值 - 代码页名称格式的列表 (相同的代码页被过滤，默认代码页中没有的代码页将添加，用于显示)
'GetType = 4 获取代码页值列表 (相同的代码页被过滤，默认代码页中没有的代码页将添加，用于显示)
'GetType = 5 获取代码页名称列表 (相同的代码页被过滤，默认代码页中没有的代码页将添加，用于显示)
'GetType = 6 获取代码页名称 + vbNullChar + 代码页值的列表 (相同的代码页被过滤，默认代码页中没有的代码页将添加，用于显示)
'GetType = 7 获取代码页值列表 (相同的代码页被过滤，默认代码页中没有的代码页将添加，自动检测项去除，用于显示)
'GetType = 8 获取代码页名称列表 (相同的代码页被过滤，默认代码页中没有的代码页将添加，自动检测项去除，用于显示)
'Mode = False 过滤掉 Unicode 编码项目，否则不过滤
Public Function GetLangStrList(LangList() As LANG_PROPERTIE,ByVal GetType As Long,Optional ByVal Mode As Boolean) As String()
	Dim i As Long,n As Long,Dic As Object,MsgList() As String
	If GetType = 0 Then
		ReDim TempList(UBound(LangList)) As String
		For n = 0 To UBound(LangList)
			TempList(n) = LangList(n).LangName
		Next n
	ElseIf GetType = 1 Then
		ReDim TempList(UBound(LangList)) As String
		For n = 0 To UBound(LangList)
			TempList(n) = CStr$(LangList(n).LangID)
		Next n
	ElseIf GetType = 2 Then
		Set Dic = CreateObject("Scripting.Dictionary")
		ReDim TempList(UBound(LangList)) As String
		For i = 0 To UBound(LangList)
			If Not Dic.Exists(LangList(i).CodePage) Then
				Dic.Add(LangList(i).CodePage,"")
				TempList(n) = CStr$(LangList(i).CodePage)
				n = n + 1
			End If
		Next i
	ElseIf getMsgList(UIDataList,MsgList,"GetUniCodePageList",1) = True Then
		ReDim CodePage(20) As Long
		CodePage(0) = CP_UNKNOWN		'未知 (自动监测) = -1
		CodePage(1) = CP_WESTEUROPE		'拉丁文 1 (ANSI) = 1252
		CodePage(2) = CP_EASTEUROPE	    '拉丁文 2 (中欧) = 1250
		CodePage(3) = CP_RUSSIAN		'西里尔文 (斯拉夫) = 1251
		CodePage(4) = CP_GREEK			'希腊文 = 1253
		CodePage(5) = CP_TURKISH		'拉丁文 5 (土耳其) = 1254
		CodePage(6) = CP_HEBREW			'希伯来文 = 1255
		CodePage(7) = CP_ARABIC			'阿拉伯文 = 1256
		CodePage(8) = CP_BALTIC			'波罗的海文 = 1257
		CodePage(9) = CP_VIETNAMESE		'越南文 = 1258
		CodePage(10) = CP_JAPAN			'日文 = 932
		CodePage(11) = CP_CHINA			'简体中文 = 936
		CodePage(12) = CP_KOREA			'韩文 = 949
		CodePage(13) = CP_TAIWAN 		'繁体中文 = 950
		CodePage(14) = CP_THAI 			'泰文 = 874
		CodePage(15) = CP_UTF7			'UTF-7 = 65000
		CodePage(16) = CP_UTF8			'UTF-8 = 65001
		CodePage(17) = CP_UNICODELITTLE	'Unicode = 1200
		CodePage(18) = CP_UNICODEBIG	'Unicode = 1201
		CodePage(19) = CP_UTF32LE		'UnicodeLE = 65005
		CodePage(20) = CP_UTF32BE		'UnicodeBE = 65006
		Set Dic = CreateObject("Scripting.Dictionary")
		For i = 0 To UBound(CodePage)
			Dic.Add(CodePage(i),i)
		Next i
		ReDim TempList(UBound(LangList) + UBound(MsgList) + 1) As String
		n = IIf(Mode = False,15,UBound(CodePage) + 1)
		Select Case GetType
		Case 3
			For i = 0 To n - 1
				TempList(i) = CStr$(CodePage(i)) & " - " & MsgList(i)
			Next i
			For i = 0 To UBound(LangList)
				With LangList(i)
					If Not Dic.Exists(.CodePage) Then
						Dic.Add(.CodePage,"")
						TempList(n) = CStr$(.CodePage) & " - " & .CPName
						n = n + 1
					End If
				End With
			Next i
		Case 4
			For i = 0 To n - 1
				TempList(i) = CStr$(CodePage(i))
			Next i
			For i = 0 To UBound(LangList)
				With LangList(i)
					If Not Dic.Exists(.CodePage) Then
						Dic.Add(.CodePage,"")
						TempList(n) = CStr$(.CodePage)
						n = n + 1
					End If
				End With
			Next i
		Case 5
			For i = 0 To n - 1
				TempList(i) = MsgList(i)
			Next i
			For i = 0 To UBound(LangList)
				With LangList(i)
					If Not Dic.Exists(.CodePage) Then
						Dic.Add(.CodePage,"")
						TempList(n) = .CPName
						n = n + 1
					End If
				End With
			Next i
		Case 6
			For i = 0 To n - 1
				TempList(i) = MsgList(i) & vbNullChar & CStr$(CodePage(i))
			Next i
			For i = 0 To UBound(LangList)
				With LangList(i)
					If Not Dic.Exists(.CodePage) Then
						Dic.Add(.CodePage,"")
						TempList(n) = .CPName & vbNullChar & CStr$(.CodePage)
						n = n + 1
					End If
				End With
			Next i
		Case 7
			For i = 1 To n - 1
				TempList(i - 1) = CStr$(CodePage(i))
			Next i
			n = n - 1
			For i = 0 To UBound(LangList)
				With LangList(i)
					If Not Dic.Exists(.CodePage) Then
						Dic.Add(.CodePage,"")
						TempList(n) = CStr$(.CodePage)
						n = n + 1
					End If
				End With
			Next i
		Case 8
			For i = 1 To n - 1
				TempList(i - 1) = MsgList(i)
			Next i
			n = n - 1
			For i = 0 To UBound(LangList)
				With LangList(i)
					If Not Dic.Exists(.CodePage) Then
						Dic.Add(.CodePage,"")
						TempList(n) = .CPName
						n = n + 1
					End If
				End With
			Next i
		End Select
	End If
	Set Dic = Nothing
	If n > 0 Then n = n - 1
	ReDim Preserve TempList(n) As String
	GetLangStrList = TempList
End Function


'打开文本文件
Public Function OpenFile(ByVal File As String,FileDataList() As String,ByVal x As Long,RunStemp As Boolean) As Boolean
	Dim i As Long,ExePath As String,ExeName As String,Argument As String,ExtName As String
	Dim TempArray() As String,MsgList() As String,WshShell As Object

	OpenFile = False
	If getMsgList(UIDataList,MsgList,"OpenFile",1) = False Then Exit Function

	If x > 0 Then
		On Error Resume Next
		Set WshShell = CreateObject("WScript.Shell")
		If WshShell Is Nothing Then
			Err.Source = "WScript.Shell"
			Call sysErrorMassage(Err,2)
			Exit Function
		End If
		On Error GoTo 0
	End If

	Select Case x
	Case 0
		If EditFile(File,FileDataList,RunStemp) = True Then OpenFile = True
	Case 1
		ExePath = Environ("SystemRoot") & "\system32\notepad.exe"
		If Dir$(ExePath) = "" Then
			ExePath = Environ("SystemRoot") & "\notepad.exe"
		End If
		If Dir(ExePath) = "" Then
			MsgBox MsgList(1),vbOkOnly+vbInformation,MsgList(0)
		Else
			If WshShell.Run("""" & ExePath & """ " & """" & File & """",1,RunStemp) <> 0 Then
				MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
			Else
				OpenFile = True
			End If
		End If
	Case 2
		i = InStrRev(File,".")
		If i > 0 Then ExtName = Mid$(File,i)
		On Error Resume Next
		ExtName = WshShell.RegRead("HKCR\" & ExtName & "\")
		If ExtName <> "" Then
			ExePath = WshShell.RegRead("HKCR\" & ExtName & "\shell\edit\command\")
			If ExePath = "" Then
				ExePath = WshShell.RegRead("HKCR\" & ExtName & "\shell\open\command\")
			End If
			If ExePath = "" Then
				ExePath = WshShell.RegRead("HKCR\" & ExtName & "\shell\preview\command\")
			End If
		End If
		On Error GoTo 0
		If ExePath <> "" Then
			i = InStr(ExePath,".")
			If i > 0 Then Argument = Trim$(Mid$(ExePath,InStr(i,ExePath," ")))
			ExePath = Left$(ExePath,Len(ExePath) - Len(Argument))
			TempArray = ReSplit(ExePath,"%")
			If UBound(TempArray) >= 2 Then
				ExePath = Replace$(ExePath,"%" & TempArray(1) & "%",Environ(TempArray(1)),,1)
			End If
			ExePath = RemoveBackslash(ExePath,"""","""",1)
			ExeName = Mid$(ExePath,InStrRev(ExePath,"\") + 1)

			If ExePath <> "" Then
				If InStr(ExePath,"\") = 0 Then
					If Dir$(Environ("SystemRoot") & "\system32\" & ExePath) <> "" Then
						ExePath = Environ("SystemRoot") & "\system32\" & ExePath
					ElseIf Dir$(Environ("SystemRoot") & "\" & ExePath) <> "" Then
						ExePath = Environ("SystemRoot") & "\" & ExePath
					End If
				End If
			End If

			If Argument <> "" Then
				If InStr(Argument,"%1") Then
					File = Replace$(Argument,"%1",File)
				ElseIf InStr(Argument,"%L") Then
					File = Replace$(Argument,"%L",File)
				Else
					File = Argument & " " & """" & File & """"
				End If
			Else
				File = """" & File & """"
			End If
		End If
		If ExePath = "" Then
			MsgBox MsgList(2),vbOkOnly+vbInformation,MsgList(0)
		ElseIf Dir$(ExePath) <> "" Then
			If WshShell.Run("""" & ExePath & """ " & File,1,RunStemp) <> 0 Then
				MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
			Else
				ExeName = Mid$(ExePath,InStrRev(ExePath,"\") + 1)
				For i = 0 To UBound(Tools)
					If InStr(LCase$(Tools(i).FilePath),LCase$(ExeName)) Then
						File = ""
						Exit For
					End If
				Next i
				If File <> "" Then Call AddTools(Tools,ExeName,ExePath,Argument)
				OpenFile = True
			End If
		Else
			MsgBox Replace$(Replace$(Replace$(MsgList(6),"%s!1!",ExeName),"%s!2!",ExePath), _
					"%s!3!",Argument) & MsgList(3),vbOkOnly+vbInformation,MsgList(0)
		End If
  	Case 3
		If CommandInput(ExePath,Argument) = True Then
			TempArray = ReSplit(ExePath,"%")
			If UBound(TempArray) = 2 Then
				ExePath = Replace$(ExePath,"%" & TempArray(1) & "%",Environ(TempArray(1)))
			End If
			ExeName = Mid$(ExePath,InStrRev(ExePath,"\") + 1)

			If Argument <> "" Then
				If InStr(Argument,"%1") Then
					File = Replace$(Argument,"%1",File)
				ElseIf InStr(Argument,"%L") Then
					File = Replace$(Argument,"%L",File)
				Else
					File = Argument & " " & """" & File & """"
				End If
			Else
				File = """" & File & """"
			End If

			If Dir$(ExePath) <> "" Then
				If WshShell.Run("""" & ExePath & """ " & File,1,RunStemp) <> 0 Then
					MsgBox MsgList(5),vbOkOnly+vbInformation,MsgList(0)
				Else
					ExeName = Mid$(ExePath,InStrRev(ExePath,"\") + 1)
					For i = 0 To UBound(Tools)
						If InStr(LCase$(Tools(i).FilePath),LCase$(ExeName)) Then
							File = ""
							Exit For
						End If
					Next i
					If File <> "" Then Call AddTools(Tools,ExeName,ExePath,Argument)
					OpenFile = True
				End If
			Else
				MsgBox ExeName & MsgList(3),vbOkOnly+vbInformation,MsgList(0)
			End If
		End If
	Case Is > 3
		ExeName = Tools(x).sName
		ExePath = Tools(x).FilePath
		Argument = Tools(x).Argument
		If Argument <> "" Then
			If InStr(Argument,"%1") Then
				File = Replace$(Argument,"%1",File)
			ElseIf InStr(Argument,"%L") Then
				File = Replace$(Argument,"%L",File)
			Else
				File = Argument & " " & """" & File & """"
			End If
		Else
			File = """" & File & """"
		End If
		If ExePath <> "" Then
			If Dir$(ExePath) <> "" Then
				If WshShell.Run("""" & ExePath & """ " & File,1,RunStemp) <> 0 Then
					MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
				Else
					OpenFile = True
				End If
			Else
				MsgBox ExeName & MsgList(3),vbOkOnly+vbInformation,MsgList(0)
			End If
		End If
	End Select
	Set WshShell = Nothing
End Function


'输入编辑程序
Private Function CommandInput(CmdPath As String,Argument As String) As Boolean
	Dim MsgList() As String,TempList() As String
	If getMsgList(UIDataList,MsgList,"CommandInput",1) = False Then Exit Function
	ToolsBak = Tools
	Begin Dialog UserDialog 540,294,MsgList(0),.CommandInputDlgFunc ' %GRID:10,7,1,1
		Text 10,7,520,140,MsgList(1),.TipText
		Text 10,154,490,14,MsgList(2),.CmdPathText
		TextBox 10,175,490,21,.CmdPath
		PushButton 500,175,30,21,MsgList(3),.BrowseButton
		Text 10,210,490,14,MsgList(4),.ArgumentText
		TextBox 10,231,490,21,.Argument
		PushButton 500,231,30,21,MsgList(5),.FileArgButton

		Text 10,7,410,14,MsgList(6),.EditerListText
		ListBox 10,28,410,119,TempList(),.EditerList
		ListBox 10,28,410,119,TempList(),.EditerListBak
		PushButton 430,28,100,21,MsgList(7),.AddButton
		PushButton 430,49,100,21,MsgList(8),.ChangeButton
		PushButton 430,70,100,21,MsgList(9),.DelButton
		PushButton 430,98,100,21,MsgList(10),.UpButton
		PushButton 430,119,100,21,MsgList(11),.DownButton

		PushButton 20,266,100,21,MsgList(12),.ClearButton
		PushButton 130,266,120,21,MsgList(13),.EditerListButton
		PushButton 130,266,100,21,MsgList(14),.ResetButton
		PushButton 310,266,100,21,MsgList(15),.SaveButton
		OKButton 310,266,100,21,.OKButton
		CancelButton 420,266,100,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.CmdPath = CmdPath
	dlg.Argument = Argument
	If Dialog(dlg) = 0 Then Exit Function
	If dlg.CmdPath <> "" Then
		CmdPath = dlg.CmdPath
		Argument = dlg.Argument
		CommandInput = True
	End If
End Function


'获取编辑程序对话框函数
Private Function CommandInputDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,x As Long,y As Long,Path As String
	Dim Temp As String,TempList() As String,TempArray() As String,MsgList() As String
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		DlgVisible "EditerListText",False
		DlgVisible "EditerList",False
		DlgVisible "EditerListBak",False
		DlgVisible "AddButton",False
		DlgVisible "ChangeButton",False
		DlgVisible "DelButton",False
		DlgVisible "UpButton",False
		DlgVisible "DownButton",False
		DlgVisible "ResetButton",False
		DlgVisible "SaveButton",False
		If UBound(Tools) < 4 Then
			DlgEnable "UpButton",False
			DlgEnable "DownButton",False
		End If
		'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			x = CreateFont(0,LFList(0))
			If x = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,x,0)
			Next i
		End If
	Case 2 ' 数值更改或者按下了按钮
		CommandInputDlgFunc = True ' 防止按下按钮关闭对话框窗口
		Select Case DlgItem$
		Case "CancelButton"
			If DlgVisible("DelButton") = False Then
				CommandInputDlgFunc = False
				Exit Function
			End If
			Tools = ToolsBak
			DlgVisible "TipText",True
			DlgVisible "OKButton",True
			DlgVisible "EditerListText",False
			DlgVisible "EditerList",False
			DlgVisible "AddButton",False
			DlgVisible "ChangeButton",False
			DlgVisible "DelButton",False
			DlgVisible "UpButton",False
			DlgVisible "DownButton",False
			DlgVisible "ResetButton",False
			DlgVisible "SaveButton",False
			DlgVisible "EditerListButton",True
			DlgEnable "ClearButton",True
			Exit Function
		Case "OKButton"
			If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
			Temp = Trim$(DlgText("CmdPath"))
			If Temp = "" Then
				MsgBox MsgList(6),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			Else
				TempList = ReSplit(Temp,"%")
				If UBound(TempList) = 2 Then
					Temp = Replace$(Temp,"%" & TempList(1) & "%",Environ(TempList(1)))
				End If
				If Dir$(Temp) = "" Then
					MsgBox MsgList(7),vbOkOnly+vbInformation,MsgList(0)
					Exit Function
				End If
			End If
			CommandInputDlgFunc = False
			Exit Function
		Case "SaveButton"
			If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
			ReDim TempList(UBound(Tools)) As String,TempArray(UBound(Tools)) As String
			For i = 4 To UBound(Tools)
				Temp = Trim$(Tools(i).FilePath)
				If Temp = "" Then
					TempList(x) = Tools(i).sName
					x = x + 1
				ElseIf Dir$(Temp) = "" Then
					TempArray(y) = Tools(i).FilePath
					y = y + 1
				End If
			Next i
			If x > 0 Then
				ReDim Preserve TempList(x - 1) As String
				MsgBox Replace$(MsgList(8),"%s",StrListJoin(TempList,vbCrLf)),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			ElseIf y > 0 Then
				ReDim Preserve TempArray(y - 1) As String
				MsgBox Replace$(MsgList(9),"%s",StrListJoin(TempArray,vbCrLf)),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
			If WriteSettings("Tools") = False Then
				MsgBox Replace(MsgList(12),"%s",RegKey),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
			MsgBox MsgList(13),vbOkOnly+vbInformation,MsgList(10)
			ToolsBak = Tools
			DlgVisible "TipText",True
			DlgVisible "OKButton",True
			DlgVisible "EditerListText",False
			DlgVisible "EditerList",False
			DlgVisible "AddButton",False
			DlgVisible "ChangeButton",False
			DlgVisible "DelButton",False
			DlgVisible "UpButton",False
			DlgVisible "DownButton",False
			DlgVisible "ResetButton",False
			DlgVisible "SaveButton",False
			DlgVisible "EditerListButton",True
			DlgEnable "ClearButton",True
			Exit Function
		Case "BrowseButton"
			If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
			If PSL.SelectFile(Path,True,MsgList(2),MsgList(1)) = False Then
				Exit Function
			End If
			DlgText "CmdPath",Path
			If DlgVisible("SaveButton") = True Then
				Temp = DlgText("EditerList")
				For i = 4 To UBound(Tools)
					If Tools(i).sName = Temp Then
						Tools(i).FilePath = Path
						Exit For
					End If
				Next i
			End If
			Exit Function
		Case "FileArgButton"
			If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
			ReDim TempList(0) As String
			TempList(0) = MsgList(3)
			If ShowPopupMenu(TempList,vbPopupUseRightButton) < 0 Then Exit Function
			DlgText "Argument",DlgText("Argument") & " " & """%1"""
			If DlgVisible("SaveButton") = True Then
				Temp = DlgText("EditerList")
				For i = 4 To UBound(Tools)
					If Tools(i).sName = Temp Then
						Tools(i).Argument = DlgText("Argument")
						Exit For
					End If
				Next i
			End If
			Exit Function
		Case "EditerListButton"
			Temp = Trim$(DlgText("CmdPath"))
			x = UBound(Tools) - 4
			If x > -1 Then
				ReDim TempList(x) As String
				For i = 0 To x
					TempList(i) = Tools(i + 4).sName
					If Tools(i + 4).FilePath = Temp Then y = i
				Next i
				DlgText "CmdPath",Tools(y + 4).FilePath
				DlgText "Argument",Tools(y + 4).Argument
			Else
				ReDim TempList(0) As String
				DlgText "CmdPath",""
				DlgText "Argument",""
			End If
			DlgListBoxArray "EditerList",TempList()
			DlgListBoxArray "EditerListBak",TempList()
			DlgValue "EditerList",y
			DlgValue "EditerListBak",y
			DlgVisible "TipText",False
			DlgVisible "OKButton",False
			DlgVisible "EditerListButton",False
			DlgVisible "EditerListText",True
			DlgVisible "EditerList",True
			DlgVisible "AddButton",True
			DlgVisible "ChangeButton",True
			DlgVisible "DelButton",True
			DlgVisible "UpButton",True
			DlgVisible "DownButton",True
			DlgVisible "ResetButton",True
			DlgVisible "SaveButton",True
		Case "AddButton"
			If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
			If PSL.SelectFile(Path,True,MsgList(2),MsgList(1)) = False Then
				Exit Function
			End If
			For i = 0 To UBound(Tools)
				If LCase$(Tools(i).FilePath) = LCase$(Path) Then
					MsgBox MsgList(11),vbOkOnly+vbInformation,MsgList(10)
					Exit Function
				End If
			Next i
			Temp = Mid$(Path,InStrRev(Path,"\") + 1)
			For i = 0 To UBound(Tools)
				If LCase$(Tools(i).sName) = LCase$(Temp) Then
					Temp = AddSet(GetToolNameList(Tools,0),Temp)
					If Temp = "" Then Exit Function
				End If
			Next i
			If AddTools(Tools,Temp,Path,"") = False Then Exit Function
			x = UBound(Tools) - 4
			TempList = GetToolNameList(Tools,4)
			DlgText "CmdPath",Path
			DlgText "Argument",""
			DlgListBoxArray "EditerList",TempList()
			DlgListBoxArray "EditerListBak",TempList()
			DlgValue "EditerList",x
			DlgValue "EditerListBak",x
		Case "ChangeButton"
			If DlgValue("EditerList") < 0 Then Exit Function
			TempList = GetToolNameList(Tools,4)
			x = DlgValue("EditerList")
			Temp = EditSet(TempList,x)
			If Temp = "" Then Exit Function
			TempList(x) = Temp
			Tools(x + 4).sName = Temp
			DlgListBoxArray "EditerList",TempList()
			DlgListBoxArray "EditerListBak",TempList()
			DlgValue "EditerList",x
			DlgValue "EditerListBak",x
			Exit Function
		Case "DelButton"
			If DlgValue("EditerList") < 0 Then Exit Function
			If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
			Temp = DlgText("EditerList")
			If MsgBox(Replace(MsgList(5),"%s",Temp),vbYesNo+vbInformation,MsgList(4)) = vbNo Then
				Exit Function
			End If
			y = DlgValue("EditerList") + 4
			Call DelToolsArray(Tools,y)
			x = UBound(Tools)
			If x > 3 Then
				If y = x + 1 Then y = x
				TempList = GetToolNameList(Tools,4)
				DlgListBoxArray "EditerList",TempList()
				DlgListBoxArray "EditerListBak",TempList()
				DlgValue "EditerList",y - 4
				DlgValue "EditerListBak",y - 4
				DlgText "CmdPath",Tools(y).FilePath
				DlgText "Argument",Tools(y).Argument
			Else
				ReDim TempList(0) As String
				DlgListBoxArray "EditerList",TempList()
				DlgListBoxArray "EditerListBak",TempList()
				DlgValue "EditerList",0
				DlgValue "EditerListBak",0
				DlgText "CmdPath",""
				DlgText "Argument",""
			End If
		Case "UpButton"
			i = DlgValue("EditerList")
			If i = 0 Then Exit Function
			ReDim TempTools(0) As TOOLS_PROPERTIE
			TempTools(0) = Tools(i + 4)
			Tools(i + 4) = Tools(i + 3)
			Tools(i + 3) = TempTools(0)
			TempList = GetToolNameList(Tools,4)
			DlgListBoxArray "EditerList",TempList()
			DlgValue "EditerList",i - 1
		Case "DownButton"
			i = DlgValue("EditerList")
			If i = UBound(Tools) Then Exit Function
			ReDim TempTools(0) As TOOLS_PROPERTIE
			TempTools(0) = Tools(i + 4)
			Tools(i + 4) = Tools(i + 5)
			Tools(i + 5) = TempTools(0)
			TempList = GetToolNameList(Tools,4)
			DlgListBoxArray "EditerList",TempList()
			DlgValue "EditerList",i + 1
		Case "ClearButton"
			If DlgVisible("EditerList") = False Then
				DlgText "CmdPath",""
 				DlgText "Argument",""
 				Exit Function
 			Else
 				If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
 				If MsgBox(MsgList(14),vbYesNo+vbInformation,MsgList(4)) = vbNo Then Exit Function
 				ReDim Preserve Tools(3) As TOOLS_PROPERTIE
 				ReDim TempList(0) As String
				DlgListBoxArray "EditerList",TempList()
				DlgListBoxArray "EditerListBak",TempList()
				DlgValue "EditerList",0
				DlgValue "EditerListBak",0
				DlgText "CmdPath",""
				DlgText "Argument",""
 			End If
 		Case "ResetButton"
			Tools = ToolsBak
			Temp = DlgText("EditerList")
			x = UBound(Tools) - 4
			If x > -1 Then
				ReDim TempList(x) As String
				For i = 0 To x
					TempList(i) = Tools(i + 4).sName
					If Tools(i + 4).FilePath = Temp Then y = i
				Next i
				DlgText "CmdPath",Tools(y + 4).FilePath
				DlgText "Argument",Tools(y + 4).Argument
			Else
				ReDim TempList(0) As String
				DlgText "CmdPath",""
				DlgText "Argument",""
			End If
			DlgListBoxArray "EditerList",TempList()
			DlgListBoxArray "EditerListBak",TempList()
			DlgValue "EditerList",y
			DlgValue "EditerListBak",y
		Case "EditerList"
			x = DlgValue("EditerList")
			If x < 0 Then Exit Function
			DlgValue "EditerList",x
			DlgValue "EditerListBak",x
			DlgText "CmdPath",Tools(x + 4).FilePath
			DlgText "Argument",Tools(x + 4).Argument
		End Select
		If DlgVisible("EditerList") = True Then
			If DlgValue("EditerList") < 0 Then
				DlgEnable "ChangeButton",False
				DlgEnable "DelButton",False
				DlgEnable "UpButton",False
				DlgEnable "DownButton",False
				DlgEnable "ResetButton",False
				DlgEnable "ClearButton",False
			Else
				DlgEnable "ChangeButton",True
				DlgEnable "DelButton",True
				DlgEnable "ResetButton",True
				DlgEnable "ClearButton",True
				If DlgListBoxArray("EditerList") < 2 Then
					DlgEnable "UpButton",False
					DlgEnable "DownButton",False
				Else
					Select Case DlgValue("EditerList")
					Case 0
						DlgEnable "UpButton",False
						DlgEnable "DownButton",True
					Case DlgListBoxArray("EditerList") - 1
						DlgEnable "UpButton",True
						DlgEnable "DownButton",False
					Case Else
						DlgEnable "UpButton",True
						DlgEnable "DownButton",True
					End Select
				End If
			End If
		End If
	Case 3 ' 文本框或者组合框文本被更改
		If DlgItem$ = "CmdPath" Or DlgItem$ = "Argument" Then
			If DlgItem$ = "CmdPath" Then
				DlgText "CmdPath",Trim$(DlgText("CmdPath"))
				Temp = DlgText("CmdPath")
				If Temp = "" Then Exit Function
				TempList = ReSplit(Temp,"%")
				If UBound(TempList) = 2 Then
					Temp = Replace$(Temp,"%" & TempList(1) & "%",Environ(TempList(1)))
				End If
				If Dir$(Temp) = "" Then
					If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
					MsgBox MsgList(7),vbOkOnly+vbInformation,MsgList(0)
				End If
			Else
				DlgText "Argument",Trim$(DlgText("Argument"))
			End If
			If DlgVisible("DelButton") = True Then
				Temp = DlgText("EditerListBak")
				If Temp = "" Then Exit Function
				For i = 4 To UBound(Tools)
					If Tools(i).sName = Temp Then
						If DlgItem$ = "CmdPath" Then
							Tools(i).FilePath = DlgText("CmdPath")
						Else
							Tools(i).Argument = DlgText("Argument")
						End If
						Exit For
					End If
				Next i
			End If
		End If
	End Select
End Function


'编辑文本文件
'Mode = True 编辑模式，如果打开文件成功返回字符编码和 True
'Mode = False 查看和确认字符编码模式，如果打开文件成功并按 [确定] 按钮返回字符编码和 True
Public Function EditFile(ByVal File As String,FileDataList() As String,ByVal Mode As Boolean) As Boolean
	Dim MsgList() As String,FileDataListBak() As String
	If getMsgList(UIDataList,MsgList,"EditFile",1) = False Then Exit Function
	'Dim objStream As Object
	'Set objStream = CreateObject("Adodb.Stream")
	'If objStream Is Nothing Then CodeList = getCodePageList(0,0)
	'If Not objStream Is Nothing Then CodeList = getCodePageList(0,49)
	'Set objStream = Nothing
	FileDataListBak = FileDataList
	Begin Dialog UserDialog 1020,595,IIf(Mode = True,MsgList(0),MsgList(1)) & " - " & File,.EditFileDlgFunc ' %GRID:10,7,1,1
		CheckBox 0,3,14,14,"",.OptionBox
		TextBox 0,0,0,21,.SuppValueBox
		Text 10,7,920,14,File,.FilePath,2
		Text 10,7,90,14,MsgList(2),.FindText
		DropListBox 110,3,270,21,MsgList(),.FindTextBox,1
		PushButton 420,3,90,21,MsgList(4),.FindButton
		PushButton 520,3,90,21,MsgList(5),.FilterButton
		PushButton 520,3,90,21,MsgList(6),.CloseFilterButton
		PushButton 380,3,30,21,MsgList(3),.RegExpTipButton
		Text 630,7,90,14,MsgList(7),.CodeText
		DropListBox 730,3,280,21,MsgList(),.CodeNameList
		TextBox 0,28,1020,532,.InTextBox,1
		PushButton 20,567,90,21,MsgList(8),.HelpButton
		PushButton 120,567,90,21,"",.ReadButton
		PushButton 220,567,90,21,MsgList(9),.PreviousButton
		PushButton 320,567,90,21,MsgList(10),.NextButton
		PushButton 790,567,100,21,MsgList(11),.SaveButton
		PushButton 900,567,90,21,MsgList(12),.ExitButton
		OKButton 790,567,100,21,.OKButton
		CancelButton 900,567,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	If Mode = False Then dlg.OptionBox = 1
	If Dialog(dlg) = 0 Then
		FileDataList = FileDataListBak
		Erase AllStrList,UseStrList
		Exit Function
	End If
	EditFile = True
	Erase AllStrList,UseStrList
End Function


'编辑对话框函数
Private Function EditFileDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,j As Long,n As Long
	Dim MsgList() As String,Temp As String,Code As String
	Dim TempArray() As String,TempList() As String

	Select Case Action%
	Case 1
		If getMsgList(UIDataList,MsgList,"EditFileDlgFunc",1) = False Then Exit Function
		DlgText "SuppValueBox",CStr$(SuppValue)
		DlgVisible "SuppValueBox",False
		DlgVisible "FilePath",False
		DlgVisible "OptionBox",False
		If DlgValue("OptionBox") = 0 Then
			DlgVisible "OKButton",False
			DlgVisible "CancelButton",False
		Else
			DlgVisible "SaveButton",False
			DlgVisible "ExitButton",False
		End If
		DlgVisible "CloseFilterButton",False
		GetHistory(TempList,"FindStrings","EditFileDlg")
		DlgListBoxArray "FindTextBox",TempList()
		DlgText "FindTextBox",TempList(0)
		Temp = DlgText("FilePath")
		For i = LBound(FileDataList) To UBound(FileDataList)
			TempList = ReSplit(FileDataList(i),JoinStr,-1)
			If TempList(0) = Temp Then
				Code = TempList(1)
				If Code = "" Then
					Code = CheckCode(Temp)
					TempList(1) = Code
					FileDataList(i) = StrListJoin(TempList,JoinStr)
				End If
				j = i
				Exit For
			End If
		Next i
		ReDim TempList(UBound(CodeList)) As String
		For i = LBound(CodeList) To UBound(CodeList)
			TempList(i) = CodeList(i).sName
			If CodeList(i).CharSet = Code Then n = i
		Next i
		DlgListBoxArray "CodeNameList",TempList()
		DlgValue "CodeNameList",n
		'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
		SetTextBoxLength GetDlgItem(SuppValue,DlgControlId("InTextBox")),FileLen(Temp),False
		DlgText "InTextBox",ReadTextFile(Temp,Code)
		If DlgText("InTextBox") <> "" Then
			DlgText "ReadButton",MsgList(9)
    	Else
    		DlgText "ReadButton",MsgList(8)
    		DlgEnable "FindButton",False
    		DlgEnable "FilterButton",False
    		DlgEnable "CloseFilterButton",False
    		DlgEnable "SaveButton",False
    	End If
    	If UBound(FileDataList) = 0 Then
			DlgEnable "PreviousButton",False
			DlgEnable "NextButton",False
		ElseIf j = 0 Then
			DlgEnable "PreviousButton",False
			DlgEnable "NextButton",True
		ElseIf j = UBound(FileDataList) Then
			DlgEnable "PreviousButton",True
			DlgEnable "NextButton",False
    	End If
    	'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	Case 2 ' 数值更改或者按下了按钮
		EditFileDlgFunc = True '防止按下按钮关闭对话框窗口
		If getMsgList(UIDataList,MsgList,"EditFileDlgFunc",1) = False Then Exit Function
		Select Case DlgItem$
		Case "HelpButton"
			Call Help("EditFileHelp")
			Exit Function
		Case "OKButton", "CancelButton"
			EditFileDlgFunc = False
			Exit Function
		Case "ExitButton"
			If DlgText("InTextBox") = ReadTextFile(DlgText("FilePath"),CodeList(DlgValue("CodeNameList")).CharSet) Then
				EditFileDlgFunc = False
				Exit Function
			End If
			Select Case MsgBox(MsgList(1),vbYesNoCancel+vbInformation,MsgList(0))
			Case vbYes
				Temp = DlgText("FilePath")
				If Dir$(Temp) <> "" Then SetAttr Temp,vbNormal
				If WriteToFile(Temp,DlgText("InTextBox"),CodeList(DlgValue("CodeNameList")).CharSet) = True Then
					MsgBox MsgList(5),vbOkOnly+vbInformation,MsgList(0)
					EditFileDlgFunc = False
					Exit Function
				Else
					MsgBox MsgList(6),vbOkOnly+vbInformation,MsgList(0)
				End If
			Case vbNo
				EditFileDlgFunc = False
				Exit Function
			End Select
		Case "SaveButton"
			If DlgText("InTextBox") = "" Then Exit Function
			Temp = DlgText("FilePath")
			If Dir$(Temp) <> "" Then SetAttr Temp,vbNormal
			If WriteToFile(Temp,DlgText("InTextBox"),CodeList(DlgValue("CodeNameList")).CharSet) = True Then
				MsgBox MsgList(5),vbOkOnly+vbInformation,MsgList(0)
			Else
				MsgBox MsgList(6),vbOkOnly+vbInformation,MsgList(0)
			End If
			DlgVisible "FilterButton",True
			DlgVisible "CloseFilterButton",False
		Case "CodeNameList"
			Code = CodeList(DlgValue("CodeNameList")).CharSet
			If Code = "_autodetect_all" Or Code = "_autodetect" Or Code = "_autodetect_kr" Then
				Code = CheckCode(DlgText("FilePath"))
				For i = LBound(CodeList) To UBound(CodeList)
					If CodeList(i).CharSet = Code Then
 						DlgValue "CodeNameList",i
 						Exit For
 					End If
				Next i
			End If
			Temp = DlgText("FilePath")
			'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
			SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox")),FileLen(Temp),True
			DlgText "InTextBox",ReadTextFile(Temp,Code)
			If DlgText("InTextBox") <> "" Then
				For i = LBound(FileDataList) To UBound(FileDataList)
					TempList = ReSplit(FileDataList(i),JoinStr)
					If TempList(0) = Temp Then
 						TempList(1) = Code
						FileDataList(i) = StrListJoin(TempList,JoinStr)
 					End If
				Next i
			End If
			Erase AllStrList,UseStrList
			DlgVisible "FilterButton",True
			DlgVisible "CloseFilterButton",False
		Case "ReadButton"
			If DlgText("ReadButton") = MsgList(8) Then
				Code = CodeList(DlgValue("CodeNameList")).CharSet
				Temp = DlgText("FilePath")
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox")),FileLen(Temp),True
				DlgText "InTextBox",ReadTextFile(Temp,Code)
				If DlgText("InTextBox") <> "" Then DlgText "ReadButton",MsgList(9)
			Else
				DlgText "InTextBox",""
				DlgText "ReadButton",MsgList(8)
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox")),0,True
			End If
			Erase AllStrList,UseStrList
			DlgVisible "FilterButton",True
			DlgVisible "CloseFilterButton",False
		Case "FindButton"
			If DlgText("FindTextBox") = "" Then
				MsgBox MsgList(11),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
			'添加查找内容
			GetHistory(TempList,"FindStrings","EditFileDlg")
			If InsertArray(TempList,DlgText("FindTextBox"),0,True) = True Then
				WriteHistory(TempList,"FindStrings","EditFileDlg")
				DlgListBoxArray "FindTextBox",TempList()
				DlgText "FindTextBox",TempList(0)
			End If
			'DlgFocus("InTextBox")  '设置焦点到文本框，2016 版会闪烁，使得光标位置移到最前面
			n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
			SendMessageLNG(n,WM_SETFOCUS,0,0)  '设置焦点到文本框
			Select Case FindCurPos(n,DlgText("FindTextBox"),False,DlgText("InTextBox"))
			Case 0
				MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
			Case -1
				MsgBox Replace$(MsgList(12),"%s",Temp),vbOkOnly+vbInformation,MsgList(0)
			Case -2
				MsgBox MsgList(13),vbOkOnly+vbInformation,MsgList(0)
			Case -3
				MsgBox MsgList(14),vbOkOnly+vbInformation,MsgList(0)
			End Select
			Exit Function
		Case "FilterButton"
			If DlgText("FindTextBox") = "" Then
				MsgBox MsgList(11),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
			'添加查找内容
			GetHistory(TempList,"FindStrings","EditFileDlg")
			If InsertArray(TempList,DlgText("FindTextBox"),0,True) = True Then
				WriteHistory(TempList,"FindStrings","EditFileDlg")
				DlgListBoxArray "FindTextBox",TempList()
				DlgText "FindTextBox",TempList(0)
			End If
			'检测查找内容的查找方式
			Temp = DlgText("FindTextBox")
			j = GetFindMode(Temp)
			If j = 1 Then Temp = "*" & Temp & "*"
			AllStrList = ReSplit(DlgText("InTextBox"),vbCrLf,-1)
			ReDim UseStrList(UBound(AllStrList)) As String
			n = 0
			For i = 0 To UBound(AllStrList)
				Select Case FilterStr(AllStrList(i),Temp,j)
				Case -1
					MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
					Exit Function
				Case -2
					MsgBox MsgList(13),vbOkOnly+vbInformation,MsgList(0)
					Exit Function
				Case -3
					MsgBox MsgList(14),vbOkOnly+vbInformation,MsgList(0)
					Exit Function
				Case Is > 0
					UseStrList(n) = "【" & CStr$(i + 1) & MsgList(10) & "】" & AllStrList(i)
					n = n + 1
				End Select
			Next i
			If n > 0 Then
				ReDim Preserve UseStrList(n - 1) As String
				DlgText "InTextBox",StrListJoin(UseStrList,vbCrLf)
				DlgVisible "FilterButton",False
				DlgVisible "CloseFilterButton",True
			Else
				Erase AllStrList,UseStrList
				MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
		Case "CloseFilterButton"
			If DlgText("InTextBox") = "" Then
				MsgBox MsgList(3),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
			If DlgText("InTextBox") <> StrListJoin(UseStrList,vbCrLf) Then
				If MsgBox(MsgList(2),vbYesNo+vbInformation,MsgList(0)) = vbYes Then
					TempArray = ReSplit(DlgText("InTextBox"),vbCrLf,-1)
					If UBound(UseStrList) = UBound(TempArray) Then
						Temp = "^【[0-9]+" & MsgList(10) & "】"
						For i = 0 To UBound(TempArray)
							If CheckStrRegExp(TempArray(i),Temp,0,2) = True Then
								TempList = ReSplit(TempArray(i),MsgList(10) & "】",2)
								AllStrList(StrToLong(Mid$(TempList(0),2)) - 1) = TempList(1)
							End If
						Next i
					Else
						MsgBox MsgList(3),vbOkOnly+vbInformation,MsgList(0)
						Exit Function
					End If
				End If
			End If
			Temp = StrListJoin(AllStrList,vbCrLf)
			'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
			SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox")),LenB(Temp),True
			DlgText "InTextBox",Temp
			Erase AllStrList,UseStrList
			DlgVisible "FilterButton",True
			DlgVisible "CloseFilterButton",False
		Case "PreviousButton","NextButton"
			Temp = DlgText("FilePath")
			For i = LBound(FileDataList) To UBound(FileDataList)
				TempList = ReSplit(FileDataList(i),JoinStr)
				If TempList(0) = Temp Then
					j = i
					Exit For
				End If
			Next i
			If DlgItem$ = "PreviousButton" Then
				If j <> 0 Then j = j - 1
			Else
				If j < UBound(FileDataList) Then j = j + 1
			End If
			If i <> j Then
				TempList = ReSplit(FileDataList(j),JoinStr)
				DlgText "FilePath",TempList(0)
				DlgText -1,Left$(DlgText(-1),InStr(DlgText(-1),"-") + 1) & TempList(0)
				Code = TempList(1)
				If Code = "" Then Code = CheckCode(TempList(0))
				For i = LBound(CodeList) To UBound(CodeList)
					If CodeList(i).CharSet = Code Then
 						DlgValue "CodeNameList",i
 						Exit For
 					End If
				Next i
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox")),FileLen(TempList(0)),True
				DlgText "InTextBox",ReadTextFile(TempList(0),Code)
				Erase AllStrList,UseStrList
			End If
			If UBound(FileDataList) = 0 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",False
			ElseIf j = 0 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",True
			ElseIf j = UBound(FileDataList) Then
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",False
			Else
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",True
			End If
			DlgVisible "FilterButton",True
			DlgVisible "CloseFilterButton",False
		Case "RegExpTipButton"
			If getMsgList(UIDataList,MsgList,"RegExpRuleTip",1) = False Then Exit Function
			i = ShowPopupMenu(MsgList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			If i = UBound(MsgList) Then
				Call Help("RegExpRuleHelp")
				Exit Function
			End If
			If DlgText("FindTextBox") = "" Then
				DlgText "FindTextBox",Mid$(MsgList(i),InStrRev(MsgList(i),vbTab) + 1)
			Else
				DlgFocus("FindTextBox")  '设置焦点到文本框
				DlgText "FindTextBox",InsertStr(GetFocus(),DlgText("FindTextBox"), _
						Mid$(MsgList(i),InStrRev(MsgList(i),vbTab) + 1))
			End If
			Exit Function
		End Select

    	If DlgText("InTextBox") <> "" Then
   			DlgText "ReadButton",MsgList(9)
   			DlgEnable "FindButton",True
   			DlgEnable "FilterButton",True
			DlgEnable "CloseFilterButton",True
			If DlgVisible("FilterButton") = True Then
				DlgEnable "SaveButton",True
				DlgEnable "ExitButton",True
				DlgEnable "CancelButton",True
			ElseIf DlgValue("OptionBox") = 0 Then
				DlgEnable "SaveButton",False
				DlgEnable "ExitButton",False
				DlgEnable "CancelButton",False
			End If
		Else
			DlgText "ReadButton",MsgList(8)
			DlgEnable "FindButton",False
			DlgEnable "FilterButton",False
			DlgEnable "CloseFilterButton",False
			DlgEnable "SaveButton",False
			DlgEnable "ExitButton",True
			DlgEnable "CancelButton",True
		End If
	Case 3 ' 文本框或者组合框文本被更改
		If getMsgList(UIDataList,MsgList,"EditFileDlgFunc",1) = False Then Exit Function
		Select Case DlgItem$
		Case "InTextBox"
			If DlgText("InTextBox") <> "" Then
				DlgText "ReadButton",MsgList(9)
				DlgEnable "FindButton",True
				DlgEnable "FilterButton",True
				DlgEnable "CloseFilterButton",True
				If DlgVisible("FilterButton") = True Then
					DlgEnable "SaveButton",True
					DlgEnable "ExitButton",True
					DlgEnable "CancelButton",True
				ElseIf DlgValue("OptionBox") = 0 Then
					DlgEnable "SaveButton",False
					DlgEnable "ExitButton",False
					DlgEnable "CancelButton",False
				End If
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox")),LenB(DlgText("InTextBox")),True
			Else
				DlgText "ReadButton",MsgList(8)
				DlgEnable "FindButton",False
				DlgEnable "FilterButton",False
				DlgEnable "CloseFilterButton",False
				DlgEnable "SaveButton",False
				DlgEnable "ExitButton",True
				DlgEnable "CancelButton",True
			End If
		End Select
	Case 4 ' 焦点被更改
		Select Case DlgItem$
		Case "InTextBox"
			i = Len(Clipboard)
			If i < Len(DlgText("InTextBox")) * 2 Then Exit Function
			'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
			SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox")),LenB(DlgText("InTextBox")) + i,True
		End Select
	Case 6 ' 函数快捷键
		If getMsgList(UIDataList,MsgList,"EditFileDlgFunc",1) = False Then Exit Function
		Select Case SuppValue
		Case 1
			Call Help("EditFileHelp")
		Case 2
			If getMsgList(UIDataList,MsgList,"RegExpRuleTip",1) = False Then Exit Function
			i = ShowPopupMenu(MsgList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			If i = UBound(MsgList) Then
				Call Help("RegExpRuleHelp")
				Exit Function
			End If
			If DlgText("FindTextBox") = "" Then
				DlgText "FindTextBox",Mid$(MsgList(i),InStrRev(MsgList(i),vbTab) + 1)
			Else
				DlgFocus("FindTextBox")  '设置焦点到文本框
				DlgText "FindTextBox",InsertStr(GetFocus(),DlgText("FindTextBox"), _
						Mid$(MsgList(i),InStrRev(MsgList(i),vbTab) + 1))
			End If
		Case 3
			If DlgText("FindTextBox") = "" Then
				MsgBox MsgList(11),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
			'DlgFocus("InTextBox")  '设置焦点到文本框，2016 版会闪烁，使得光标位置移到最前面
			n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
			SendMessageLNG(n,WM_SETFOCUS,0,0)  '设置焦点到文本框
			Select Case FindCurPos(n,DlgText("FindTextBox"),False,DlgText("InTextBox"))
			Case 0
				MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
			Case -1
				MsgBox Replace$(MsgList(12),"%s",Temp),vbOkOnly+vbInformation,MsgList(0)
			Case -2
				MsgBox MsgList(13),vbOkOnly+vbInformation,MsgList(0)
			Case -3
				MsgBox MsgList(14),vbOkOnly+vbInformation,MsgList(0)
			End Select
		Case 4
			If DlgVisible("FilterButton") = True Then
				If DlgText("FindTextBox") = "" Then
					MsgBox MsgList(11),vbOkOnly+vbInformation,MsgList(0)
					Exit Function
				End If
				'检测查找内容的查找方式
				Temp = DlgText("FindTextBox")
				j = GetFindMode(Temp)
				If j = 1 Then Temp = "*" & Temp & "*"
				AllStrList = ReSplit(DlgText("InTextBox"),vbCrLf,-1)
				ReDim UseStrList(UBound(AllStrList)) As String
				n = 0
				For i = 0 To UBound(AllStrList)
					Select Case FilterStr(AllStrList(i),Temp,j)
					Case -1
						MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
						Exit Function
					Case -2
						MsgBox MsgList(13),vbOkOnly+vbInformation,MsgList(0)
						Exit Function
					Case -3
						MsgBox MsgList(14),vbOkOnly+vbInformation,MsgList(0)
						Exit Function
					Case Is > 0
						UseStrList(n) = "【" & CStr$(i + 1) & MsgList(10) & "】" & AllStrList(i)
						n = n + 1
					End Select
				Next i
				If n > 0 Then
					ReDim Preserve UseStrList(n - 1) As String
					DlgText "InTextBox",StrListJoin(UseStrList,vbCrLf)
					DlgVisible "FilterButton",False
					DlgVisible "CloseFilterButton",True
				Else
					Erase AllStrList,UseStrList
					MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
					Exit Function
				End If
			Else
				If DlgText("InTextBox") = "" Then
					MsgBox MsgList(3),vbOkOnly+vbInformation,MsgList(0)
					Exit Function
				End If
				If DlgText("InTextBox") <> StrListJoin(UseStrList,vbCrLf) Then
					If MsgBox(MsgList(2),vbYesNo+vbInformation,MsgList(0)) = vbYes Then
						TempArray = ReSplit(DlgText("InTextBox"),vbCrLf,-1)
						If UBound(UseStrList) = UBound(TempArray) Then
							Temp = "^【[0-9]+" & MsgList(10) & "】"
							For i = 0 To UBound(TempArray)
								If CheckStrRegExp(TempArray(i),Temp,0,2) = True Then
									TempList = ReSplit(TempArray(i),MsgList(10) & "】",2)
									AllStrList(StrToLong(Mid$(TempList(0),2)) - 1) = TempList(1)
								End If
							Next i
						Else
							MsgBox MsgList(3),vbOkOnly+vbInformation,MsgList(0)
							Exit Function
						End If
					End If
				End If
				Temp = StrListJoin(AllStrList,vbCrLf)
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox")),LenB(Temp),True
				DlgText "InTextBox",Temp
				Erase AllStrList,UseStrList
				DlgVisible "FilterButton",True
				DlgVisible "CloseFilterButton",False
			End If
			If DlgVisible("FilterButton") = True Then
				DlgEnable "SaveButton",True
				DlgEnable "ExitButton",True
				DlgEnable "CancelButton",True
			ElseIf DlgValue("OptionBox") = 0 Then
				DlgEnable "SaveButton",False
				DlgEnable "ExitButton",False
				DlgEnable "CancelButton",False
			End If
		Case 5, 6
			Temp = DlgText("FilePath")
			For i = LBound(FileDataList) To UBound(FileDataList)
				TempList = ReSplit(FileDataList(i),JoinStr)
				If TempList(0) = Temp Then
					j = i
					Exit For
				End If
			Next i
			If SuppValue = 5 Then
				If j <> 0 Then j = j - 1
			Else
				If j < UBound(FileDataList) Then j = j + 1
			End If
			If i <> j Then
				TempList = ReSplit(FileDataList(j),JoinStr)
				DlgText "FilePath",TempList(0)
				DlgText -1,Left$(DlgText(-1),InStr(DlgText(-1),"-") + 1) & TempList(0)
				Code = TempList(1)
				If Code = "" Then Code = CheckCode(TempList(0))
				For i = LBound(CodeList) To UBound(CodeList)
					If CodeList(i).CharSet = Code Then
 						DlgValue "CodeNameList",i
 						Exit For
 					End If
				Next i
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox")),FileLen(TempList(0)),True
				DlgText "InTextBox",ReadTextFile(TempList(0),Code)
				Erase AllStrList,UseStrList
			End If
			If UBound(FileDataList) = 0 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",False
			ElseIf j = 0 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",True
			ElseIf j = UBound(FileDataList) Then
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",False
			Else
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",True
			End If
			DlgVisible "FilterButton",True
			DlgVisible "CloseFilterButton",False
		End Select
	End Select
End Function


'关于和帮助
Public Sub Help(ByVal HelpTip As String)
	Dim i As Long,MsgList() As String,HelpList(22) As String
	Dim HelpTipTitle As String,HelpMsg As String

	For i = 0 To UBound(UIDataList)
		With UIDataList(i)
			Select Case .Title
			Case "Windows"
				MsgList = .Value
			Case "System"
				HelpList(0) = Replace$(Replace$(StrListJoin(.Value,vbCrLf),"%s",Version),"%d",Build) & vbCrLf & vbCrLf
			Case "Description"
				HelpList(1) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "Precondition"
				HelpList(2) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "Setup"
				HelpList(3) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "CopyRight"
				HelpList(4) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "Thank"
				HelpList(5) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "Contact"
				HelpList(6) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "Logs"
				HelpList(7) = StrListJoin(.Value,vbCrLf)
			Case "MainHelp"
				HelpList(8) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "StrTypeSetHelp"
				HelpList(9) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "RefTypeSetHelp"
				HelpList(10) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "EncodeRangeSetHelp"
				HelpList(11) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "WriteSetHelp"
				HelpList(12) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "FileReadWriteAndVoiceSetHelp"
				HelpList(13) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "UpdateSetHelp"
				HelpList(14) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "UILangSetHelp"
				HelpList(15) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "GetStringHelp"
				HelpList(16) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "EditStringHelp"
				HelpList(17) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "FindSetHelp"
				HelpList(18) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "EditFileHelp"
				HelpList(19) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "FilterReserveHelp"
				HelpList(20) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "StringSearchHelp"
				HelpList(21) = StrListJoin(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "RegExpRuleHelp"
				HelpList(22) = StrListJoin(.Value,vbCrLf)
			End Select
		End With
	Next i

	Select Case HelpTip
	Case "About"
		HelpTipTitle = MsgList(2)
		HelpMsg = HelpList(0) & HelpList(1) & HelpList(2) & HelpList(3) & HelpList(4) & _
					HelpList(5) & HelpList(6) & HelpList(7)
	Case "MainHelp"
		HelpTipTitle = MsgList(3)
		HelpMsg = HelpList(8) & HelpList(7)
	Case "StrTypeSetHelp"
		HelpTipTitle = MsgList(4)
		HelpMsg = HelpList(9) & HelpList(7)
	Case "RefTypeSetHelp"
		HelpTipTitle = MsgList(5)
		HelpMsg = HelpList(10) & HelpList(7)
	Case "EncodeRangeSetHelp"
		HelpTipTitle = MsgList(6)
		HelpMsg = HelpList(11) & HelpList(7)
	Case "WriteSetHelp"
		HelpTipTitle = MsgList(7)
		HelpMsg = HelpList(12) & HelpList(7)
	Case "FileReadWriteAndVoiceSetHelp"
		HelpTipTitle = MsgList(8)
		HelpMsg = HelpList(13) & HelpList(7)
	Case "UpdateSetHelp"
		HelpTipTitle = MsgList(9)
		HelpMsg = HelpList(14) & HelpList(7)
	Case "UILangSetHelp"
		HelpTipTitle = MsgList(10)
		HelpMsg = HelpList(15) & HelpList(7)
	Case "GetStrSetHelp"
		HelpTipTitle = MsgList(11)
		HelpMsg = HelpList(16) & HelpList(7)
	Case "GetStringHelp"
		HelpTipTitle = MsgList(12)
		HelpMsg = HelpList(16) & HelpList(7)
	Case "EditStringHelp"
		HelpTipTitle = MsgList(13)
		HelpMsg = HelpList(17) & HelpList(7)
	Case "FindSetHelp"
		HelpTipTitle = MsgList(14)
		HelpMsg = Replace$(HelpList(18) & HelpList(7),"{RegExpRule}",HelpList(22))
	Case "EditFileHelp"
		HelpTipTitle = MsgList(15)
		HelpMsg = Replace$(HelpList(19) & HelpList(7),"{RegExpRule}",HelpList(22))
	Case "FilterReserveHelp"
		HelpTipTitle = MsgList(16)
		HelpMsg = HelpList(20) & HelpList(7)
	Case "StringSearchHelp"
		HelpTipTitle = MsgList(17)
		HelpMsg = Replace$(HelpList(21) & HelpList(7),"{RegExpRule}",HelpList(22))
	Case "RegExpRuleHelp"
		HelpTipTitle = MsgList(18)
		HelpMsg = HelpList(22) & vbCrLf & vbCrLf & HelpList(7)
	End Select

	Begin Dialog UserDialog 830,553,MsgList(0) & " - " & MsgList(1),.CommonDlgFunc ' %GRID:10,7,1,1
		Text 0,7,830,14,HelpTipTitle,.Text,2
		TextBox 0,28,830,490,.TextBox,1
		OKButton 370,525,100,21
	End Dialog
	Dim dlg As UserDialog
	dlg.TextBox = HelpMsg
	Dialog dlg
End Sub


'编辑设置名称和添加设置名称对话框函数
Public Function CommonDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,j As Long
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	End Select
End Function


'获取最近历史记录
'Separator 为空时，重置方式获取
'Separator 不为空时，追加方式获取，并且 Separator 后的有效值不允许重复
Public Function GetHistory(DataList() As String,ByVal ItemName As String,ByVal ValueName As String,Optional ByVal Separator As String) As Boolean
	Dim i As Long,n As Long,TempArray() As String,Dic As Object
	On Error GoTo ExitFunction
	TempArray = GetAllSettings(AppName,ItemName)
	If Separator <> "" Then
		Set Dic = CreateObject("Scripting.Dictionary")
		For i = 0 To UBound(DataList)
			If InStr(DataList(i),Separator) Then
				If Not Dic.Exists(ReSplit(DataList(i),Separator,2)(1)) Then
					Dic.Add(ReSplit(DataList(i),Separator,2)(1),"")
				End If
			End If
		Next i
		n = UBound(DataList) + 1
		ReDim Preserve DataList(n + UBound(TempArray)) As String
		For i = LBound(TempArray) To UBound(TempArray)
			If InStr(TempArray(i,0),ValueName) Then
				If InStr(TempArray(i,1),Separator) Then
					If Not Dic.Exists(ReSplit(TempArray(i,1),Separator,2)(1)) Then
						DataList(n) = TempArray(i,1)
						n = n + 1
					End If
				End If
			End If
		Next i
		Set Dic = Nothing
	Else
		ReDim DataList(UBound(TempArray)) As String
		For i = LBound(TempArray) To UBound(TempArray)
			If InStr(TempArray(i,0),ValueName) Then
				If TempArray(i,1) <> "" Then
					DataList(n) = TempArray(i,1)
					n = n + 1
				End If
			End If
		Next i
	End If
	If n > 0 Then
		n = n - 1
		GetHistory = True
	End If
	ReDim Preserve DataList(n) As String
	Exit Function
	ExitFunction:
	If Separator = "" Then
		ReDim DataList(0) As String
	Else
		ReDim Preserve DataList(UBound(DataList)) As String
	End If
End Function


'保存最近历史记录
'Mode = False 写入最多 10 个记录，否则写入所有记录，并删除 DataList 中没有的记录
Public Function WriteHistory(DataList() As String,ByVal ItemName As String,ByVal ValueName As String,Optional ByVal Mode As Boolean) As Boolean
	Dim i As Long,n As Long,TempArray() As String
	On Error Resume Next
	For i = LBound(DataList) To UBound(DataList)
		If DataList(i) <> "" Then
			SaveSetting(AppName,ItemName,ValueName & CStr$(n),DataList(i))
			n = n + 1
			If Mode = False Then
				If n = 10 Then
					WriteHistory = True
					Exit Function
				End If
			End If
		End If
	Next i
	If n > 0 Then WriteHistory = True
	TempArray = GetAllSettings(AppName,ItemName)
	For i = 0 To UBound(TempArray)
		If InStr(TempArray(i,0),ValueName & CStr$(n)) Then
			DeleteSetting(AppName,ItemName,TempArray(i,0))
			n = n + 1
		End If
	Next i
	If n < i Then Exit Function
	If WriteHistory = True Then Exit Function
	Dim WshShell As Object
	Set WshShell = CreateObject("WScript.Shell")
	WshShell.RegDelete RegKey & ItemName & "\"
	Set WshShell = Nothing
End Function


'添加设置名称
Public Function AddSet(DataArr() As String,Optional ByVal sName As String) As String
	Dim i As Long,MsgList() As String
	If getMsgList(UIDataList,MsgList,"AddSet",1) = False Then Exit Function
	Begin Dialog UserDialog 380,77,MsgList(0),.CommonDlgFunc ' %GRID:10,7,1,1
		Text 10,7,360,14,MsgList(1),.MainText
		TextBox 10,21,360,21,.TextBox
		OKButton 100,49,80,21,.OKButton
		CancelButton 200,49,80,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.TextBox = sName
	DataInPutDlg:
	If Dialog(dlg) = 0 Then
		AddSet = ""
		Exit Function
	End If
	AddSet = Trim(dlg.TextBox)
	If AddSet = "" Then
		MsgBox MsgList(3),vbOkOnly+vbInformation,MsgList(2)
		GoTo DataInPutDlg
	End If
	For i = LBound(DataArr) To UBound(DataArr)
		If LCase$(AddSet) = LCase$(DataArr(i)) Then
			MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(2)
			GoTo DataInPutDlg
		End If
	Next i
End Function


'编辑设置名称
Public Function EditSet(DataArr() As String,ByVal HeaderID As Long) As String
	Dim i As Long,MsgList() As String
	On Error GoTo ExitFunction
	If getMsgList(UIDataList,MsgList,"EditSet",1) = False Then Exit Function
	Begin Dialog UserDialog 380,126,MsgList(0),.CommonDlgFunc ' %GRID:10,7,1,1
		GroupBox 10,17,360,28,"",.GroupBox1
		Text 10,7,350,14,MsgList(1),.Text1
		Text 20,28,340,14,Replace$(DataArr(HeaderID),"&","&&"),.oldNameText
		Text 10,56,360,14,MsgList(2),.newNameText
		TextBox 10,70,360,21,.TextBox
		OKButton 100,98,80,21,.OKButton
		CancelButton 200,98,80,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	DataInPutDlg:
	dlg.TextBox = DataArr(HeaderID)
	If Dialog(dlg) = 0 Then
		EditSet = ""
		Exit Function
	End If
	EditSet = Trim$(dlg.TextBox)
	If EditSet = "" Then
		MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(3)
		GoTo DataInPutDlg
	End If
	If EditSet = DataArr(HeaderID) Then Exit Function
	For i = LBound(DataArr) To UBound(DataArr)
		If LCase$(EditSet) = LCase$(DataArr(i)) Then
			MsgBox MsgList(5),vbOkOnly+vbInformation,MsgList(3)
			GoTo DataInPutDlg
		End If
	Next i
	ExitFunction:
End Function


'删除文本数组项目
Public Sub DelArray(List() As String,ByVal IDList As Variant,Optional ByVal Separator As String)
	Dim i As Long,n As Long
	If IsArray(IDList) Then
		If UBound(List) = UBound(IDList) Then
			ReDim List(0) As String
			Exit Sub
		End If
		ReDim Stemp(UBound(List)) As Long
		For i = LBound(IDList) To UBound(IDList)
			If IDList(i) > -1 Then Stemp(IDList(i)) = 1
		Next i
		n = IDList(LBound(IDList))
		For i = IDList(LBound(IDList)) To UBound(List)
			If Stemp(i) = 0 Then
				List(n) = List(i)
				n = n + 1
			End If
		Next i
	ElseIf IsNumeric(IDList) And Separator = "" Then
		n = IDList
		For i = IDList + 1 To UBound(List)
			List(n) = List(i)
			n = n + 1
		Next i
	ElseIf IDList <> "" Then
		For i = LBound(List) To UBound(List)
			If Separator <> "" Then
				If ReSplit(List(i),Separator)(0) <> IDList Then
					List(n) = List(i)
					n = n + 1
				End If
			ElseIf List(i) <> IDList Then
				List(n) = List(i)
				n = n + 1
			End If
		Next i
	Else
		Exit Sub
	End If
	If n > 0 Then
		ReDim Preserve List(n - 1) As String
	Else
		ReDim List(0) As String
	End If
End Sub


'删除文本数组项目和索引字典
'Mode = False 不转义，否则转义
Public Sub DelArrays(Dic As Object,List() As String,ByVal IDList As Variant,Optional ByVal Mode As Boolean)
	Dim i As Long,n As Long,Temp As String
	If IsArray(IDList) Then
		If UBound(List) = UBound(IDList) Then
			Dic.RemoveAll
			ReDim List(0) As String
			Exit Sub
		End If
		ReDim Stemp(UBound(List)) As Long
		For i = LBound(IDList) To UBound(IDList)
			If IDList(i) > -1 Then Stemp(IDList(i)) = 1
		Next i
		n = IDList(LBound(IDList))
		For i = IDList(LBound(IDList)) To UBound(List)
			If List(i) <> "" Then
				If Mode = False Then
					Temp = List(i)
				Else
					Temp = Convert(List(i))
				End If
				If Stemp(i) = 0 Then
					List(n) = List(i)
					If Dic.Exists(Temp) Then
						Dic.Item(Temp) = n
					Else
						Dic.Add(Temp,n)
					End If
					n = n + 1
				ElseIf Dic.Exists(Temp) Then
					Dic.Remove(Temp)
				End If
			ElseIf Dic.Exists(List(i)) Then
				Dic.Remove(List(i))
			End If
		Next i
	ElseIf IsNumeric(IDList) Then
		n = IDList
		If Mode = False Then
			Temp = List(n)
		Else
			Temp = Convert(List(n))
		End If
		If Dic.Exists(Temp) Then Dic.Remove(Temp)
		For i = IDList + 1 To UBound(List)
			If List(i) <> "" Then
				List(n) = List(i)
				If Mode = False Then
					Temp = List(i)
				Else
					Temp = Convert(List(i))
				End If
				If Dic.Exists(Temp) Then
					Dic.Item(Temp) = n
				Else
					Dic.Add(Temp,n)
				End If
				n = n + 1
			ElseIf Dic.Exists(List(i)) Then
				Dic.Remove(List(i))
			End If
		Next i
	ElseIf IDList <> "" Then
		For i = LBound(List) To UBound(List)
			If List(i) <> "" Then
				If Mode = False Then
					Temp = List(i)
				Else
					Temp = Convert(List(i))
				End If
				If List(i) <> IDList Then
					List(n) = List(i)
					If Dic.Exists(Temp) Then
						Dic.Item(Temp) = n
					Else
						Dic.Add(Temp,n)
					End If
					n = n + 1
				ElseIf Dic.Exists(Temp) Then
					Dic.Remove(Temp)
				End If
			ElseIf Dic.Exists(List(i)) Then
				Dic.Remove(List(i))
			End If
		Next i
	Else
		Exit Sub
	End If
	If n > 0 Then
		ReDim Preserve List(n - 1) As String
	Else
		ReDim List(0) As String
	End If
End Sub


'删除数值数组项目
Public Sub DelLongArray(List() As Long,ByVal IDList As Variant)
	Dim i As Long,n As Long
	If IsArray(IDList) Then
		If UBound(List) = UBound(IDList) Then
			ReDim List(0) As Long
			Exit Sub
		End If
		ReDim Stemp(UBound(List)) As Long
		For i = LBound(IDList) To UBound(IDList)
			If IDList(i) > -1 Then Stemp(IDList(i)) = 1
		Next i
		n = IDList(LBound(IDList))
		For i = IDList(LBound(IDList)) To UBound(List)
			If Stemp(i) = 0 Then
				List(n) = List(i)
				n = n + 1
			End If
		Next i
	ElseIf IsNumeric(IDList) Then
		n = IDList
		For i = IDList + 1 To UBound(List)
			List(n) = List(i)
			n = n + 1
		Next i
	Else
		Exit Sub
	End If
	If n > 0 Then
		ReDim Preserve List(n - 1) As Long
	Else
		ReDim List(0) As Long
	End If
End Sub


'不区分大小写的字符替换 (保留未替换字符的大小写)
Public Function strReplace(ByVal s As String,ByVal Find As String,ByVal repwith As String,Optional ByVal Index As Long = 1,Optional ByVal Count As Long = -1) As String
	Dim i As Long,n As Long,fL As Long
	strReplace = s
	If s = "" Then Exit Function
	If Find = "" Then Exit Function
	If Index < 1 Then Exit Function
	If Count < -1 Then Exit Function
	s = LCase$(s)
	Find = LCase$(Find)
	i = InStr(Index,s,Find)
	If i = 0 Then Exit Function
	fL = Len(Find)
	Do While i > 0
		If Count = -1 Then
			strReplace = Replace$(strReplace,Mid$(strReplace,i,fL),repwith)
		Else
			strReplace = Replace$(strReplace,Mid$(strReplace,i,fL),repwith,,1)
			n = n + 1
			If n = Count Then Exit Do
		End If
		i = InStr(i + fL,LCase$(strReplace),Find)
	Loop
End Function


'插入数据到数组
'Mode = True 不允许重复项并插入或移位 Data 到最前面
Public Function InsertArray(List() As String,ByVal Data As String,ByVal insPos As Long,Optional ByVal Mode As Boolean) As Boolean
	Dim i As Long,j As Long
	i = LBound(List)
	j = UBound(List)
	If j = i And List(i) = "" Then
		List(i) = Data
	Else
		If Mode = True Then
			If insPos = 0 Then
				If List(0) = Data Then Exit Function
			ElseIf InStr(vbNullChar & StrListJoin(List,vbNullChar) & vbNullChar,vbNullChar & Data & vbNullChar) Then
				Exit Function
			End If
		End If
		ReDim Preserve List(j + 1) As String
		If insPos <= j Then
			For i = j + 1 To insPos + 1 Step -1
				List(i) = List(i - 1)
			Next i
			List(insPos) = Data
		Else
			List(j + 1) = Data
		End If
		If Mode = True Then List = ClearTextArray(List,True)
	End If
	InsertArray = True
End Function


'添加工具数据
Private Function AddTools(ToolsData() As TOOLS_PROPERTIE,ByVal CmdName As String,ByVal CmdPath As String,ByVal Argument As String) As Boolean
	Dim i As Long,FindName As String
	If CmdName = "" Or CmdPath = "" Then Exit Function
	FindName = LCase$(CmdName)
	For i = LBound(ToolsData) To UBound(ToolsData)
		If LCase$(ToolsData(i).sName) = FindName Then
			Exit Function
		End If
	Next i
	ReDim Preserve ToolsData(i) As TOOLS_PROPERTIE
	ToolsData(i).sName = CmdName
	ToolsData(i).FilePath = CmdPath
	ToolsData(i).Argument = Argument
	AddTools = True
End Function


'获取工具数据数组的工具名称列表
Public Function GetToolNameList(DataList() As TOOLS_PROPERTIE,ByVal StartID As Long) As String()
	Dim i As Long
	i = UBound(DataList)
	If StartID > i Or StartID < 0 Then
		ReDim TempList(0) As String
	Else
		ReDim TempList(i - StartID) As String
		For i = StartID To UBound(DataList)
			TempList(i - StartID) = DataList(i).sName
		Next i
	End If
	GetToolNameList = TempList
End Function


'删除自定义工具数组项目
Public Sub DelToolsArray(DataList() As TOOLS_PROPERTIE,ByVal IDList As Variant)
	Dim i As Long,n As Long
	If IsArray(IDList) Then
		If UBound(DataList) = UBound(IDList) Then
			ReDim DataList(0) As TOOLS_PROPERTIE
			Exit Sub
		End If
		ReDim Stemp(UBound(DataList)) As Long
		For i = LBound(IDList) To UBound(IDList)
			If IDList(i) > -1 Then Stemp(IDList(i)) = 1
		Next i
		n = IDList(LBound(IDList))
		For i = IDList(LBound(IDList)) To UBound(DataList)
			If Stemp(i) = 0 Then
				DataList(n) = DataList(i)
				n = n + 1
			End If
		Next i
	ElseIf IsNumeric(IDList) Then
		n = IDList
		For i = IDList + 1 To UBound(DataList)
			DataList(n) = DataList(i)
			n = n + 1
		Next i
	ElseIf IDList <> "" Then
		For i = LBound(DataList) To UBound(DataList)
			If DataList(i).sName <> IDList Then
				DataList(n) = DataList(i)
				n = n + 1
			End If
		Next i
	Else
		Exit Sub
	End If
	If n > 0 Then
		ReDim Preserve DataList(n - 1) As TOOLS_PROPERTIE
	Else
		ReDim DataList(0) As TOOLS_PROPERTIE
	End If
End Sub


'读取二进制文件
'BOM = False 检查并去掉 BOM，否则读入 BOM
Public Function ReadBinaryFile(ByVal FilePath As String,ByVal CodePage As Long,Optional ByVal BOM As Boolean) As String
	Dim FN As Variant
	If Dir$(FilePath) = "" Then Exit Function
	If BOM = False Then
		If FileLen(FilePath) < 3 Then Exit Function
	Else
		If FileLen(FilePath) = 0 Then Exit Function
	End If
	On Error GoTo ExitFunction
	FN = FreeFile
	Open FilePath For Binary Access Read Lock Write As #FN
	If BOM = False Then
		Select Case CodePage
		Case CP_UNICODELITTLE
			ReDim tempByte(1) As Byte
			Get #FN,,tempByte
			If tempByte = HexStr2Bytes("FFFE") Then
				ReDim tempByte(LOF(FN) - 3) As Byte
				Get #FN,3,tempByte
			Else
				ReDim tempByte(LOF(FN) - 1) As Byte
				Get #FN,,tempByte
			End If
		Case CP_UNICODEBIG
			ReDim tempByte(1) As Byte
			Get #FN,,tempByte
			If tempByte = HexStr2Bytes("FEFF") Then
				ReDim tempByte(LOF(FN) - 3) As Byte
				Get #FN,3,tempByte
			Else
				ReDim tempByte(LOF(FN) - 1) As Byte
				Get #FN,,tempByte
			End If
		Case CP_UTF8
			ReDim tempByte(2) As Byte
			Get #FN,,tempByte
			If tempByte = HexStr2Bytes("EFBBBF") Then
				ReDim tempByte(LOF(FN) - 4) As Byte
				Get #FN,4,tempByte
			Else
				ReDim tempByte(LOF(FN) - 1) As Byte
				Get #FN,,tempByte
			End If
		Case Else
			ReDim tempByte(LOF(FN) - 1) As Byte
			Get #FN,,tempByte
		End Select
	Else
		ReDim tempByte(LOF(FN) - 1) As Byte
		Get #FN,,tempByte
	End If
	ReadBinaryFile = ByteToString(tempByte,CodePage)
	ExitFunction:
	Close #FN
End Function


'写入二进制文件
'BOM = False 检查并写入 BOM，否则不写入 BOM
'Mode = False 删除文件，重新写入，仅在 File 为文件名时适用
Public Function WriteBinaryFile(ByVal File As Variant,ByVal CodePage As Long,ByVal textStr As String, _
		Optional ByVal BOM As Boolean,Optional ByVal Mode As Boolean) As Boolean
	Dim FN As Variant,tempByte() As Byte
	If IsNumeric(File) Then
		If textStr = "" Then Exit Function
		FN = File
		If BOM = False Then
			If LOF(FN) = 0 Then
				Select Case CodePage
				Case CP_UNICODELITTLE
					tempByte = HexStr2Bytes("FFFE")
					Put #FN,1,tempByte
				Case CP_UNICODEBIG
					tempByte = HexStr2Bytes("FEFF")
					Put #FN,1,tempByte
				Case CP_UTF8
					tempByte = HexStr2Bytes("EFBBBF")
					Put #FN,1,tempByte
				End Select
			End If
		End If
		tempByte = StringToByte(textStr,CodePage)
		Put #FN,LOF(FN) + 1,tempByte
		WriteBinaryFile = True
	Else
		If File = "" Then Exit Function
		On Error GoTo ExitFunction
		If Mode = False Then
			If Dir$(File) <> "" Then Kill File
		End If
		If textStr = "" Then Exit Function
		FN = FreeFile
		Open File For Binary Access Write Lock Write As #FN
		If BOM = False Then
			If LOF(FN) = 0 Then
				Select Case CodePage
				Case CP_UNICODELITTLE
					tempByte = HexStr2Bytes("FFFE")
					Put #FN,1,tempByte
				Case CP_UNICODEBIG
					tempByte = HexStr2Bytes("FEFF")
					Put #FN,1,tempByte
				Case CP_UTF8
					tempByte = HexStr2Bytes("EFBBBF")
					Put #FN,1,tempByte
				End Select
			End If
		End If
		tempByte = StringToByte(textStr,CodePage)
		Put #FN,LOF(FN) + 1,tempByte
		WriteBinaryFile = True
		ExitFunction:
		On Error Resume Next
		Close #FN
	End If
End Function


'比较二个字串数组是否相同，不相同返回 True
Public Function ArrayComp(uArray() As String,oArray() As String,Optional ByVal Index As String) As Boolean
	Dim i As Long,j As Long,SkipArray() As String,TempArray() As String
	If UBound(uArray) <> UBound(oArray) Then
		ArrayComp = True
		Exit Function
	End If
	If Index = "" Then
		For i = LBound(uArray) To UBound(uArray)
			If oArray(i) <> uArray(i) Then
				ArrayComp = True
				Exit For
			End If
		Next i
	Else
		SkipArray = ReSplit(Index,",",-1)
		For i = 0 To UBound(SkipArray)
			TempArray = ReSplit(SkipArray(i),"-")
			For j = StrToLong(TempArray(0)) To StrToLong(TempArray(UBound(TempArray)))
				If oArray(j) <> uArray(j) Then
					ArrayComp = True
					Exit Function
				End If
			Next j
		Next i
	End If
End Function


'主版本号和次版本号判断系统版本
Private Function GetWindowsVersion() As String
	Dim Ver As OSVERSIONINFO, os As String
	Ver.dwOSVersionInfoSize = Len(Ver)
	If GetVersionEx(Ver) = 0 Then Exit Function
	With Ver
		Select Case .dwMajorVersion
		Case 3
			Select Case .dwMinorVersion
			Case 10
				os = "Windows 3.10"
			Case 51
				os = "Windows 3.51"
			End Select
		Case 4
			Select Case .dwMinorVersion
			Case 0
				os = "Windows 95"
			Case 10
				os = "Windows 98"
			Case 90
				os = "Windows Me"
			End Select
		Case 5
			Select Case .dwMinorVersion
			Case 0
				os = "Windows 2000"
			Case 1
				os = "Windows XP"
			Case 2
				os = "Windows 2003"
			End Select
		Case 6
			Select Case .dwMinorVersion
			Case 0
				os = "Windows Vista"
			Case 1
				os = "Windows 7"
			Case 2
				os = "Windows 8"
			Case 3
				os = "Windows 10"
			End Select
		End Select
		os = os & Space(1) & CStr$(.szCSDVersion)
		GetWindowsVersion = CStr$(.dwMajorVersion) & CStr$(.dwMinorVersion)
	End With
End Function


'转换十进制和十六进制值为字符
'MaxVal = 0 按值计算应有的长度，> 0 按文件大小计算的位数，< 0 按指定位数
Public Function ValToStr(ByVal DecVal As Long,Optional ByVal MaxVal As Long,Optional ByVal DisPlayFormat As Boolean) As String
	On Error GoTo ExitFunction
	If DisPlayFormat = False Then
		ValToStr = CStr$(DecVal)
	Else
		ValToStr = Hex$(DecVal)
		Select Case MaxVal
		Case 0
			MaxVal = Len(ValToStr)
		Case Is > 0
			MaxVal = Len(CStr$(MaxVal))
		Case Else
			MaxVal = Abs(MaxVal)
		End Select
		If MaxVal And 1 Then MaxVal = MaxVal + 1
		ValToStr = String$(MaxVal - Len(ValToStr),"0") & ValToStr
	End If
	ExitFunction:
End Function


'转换十进制和十六进制字符为十进制值
Public Function StrToVal(ByVal textStr As String,Optional ByVal DisPlayFormat As Boolean) As Variant
	On Error GoTo ExitFunction
	If DisPlayFormat = False Then
		StrToVal = CLng(textStr)
	Else
		StrToVal = Val("&H" & textStr)
	End If
	Exit Function
	ExitFunction:
	StrToVal = -1
End Function


'批量转换十进制字符为十六进制字符
Public Function DecStrListToHexStrList(StrList() As String,Optional ByVal MaxVal As Long,Optional ByVal DisPlayFormat As Boolean) As String()
	Dim i As Long,TempList() As String
	On Error GoTo ExitFunction
	TempList = StrList
	If DisPlayFormat = True Then
		For i = 0 To UBound(TempList)
			If TempList(i) <> "" Then
				TempList(i) = ValToStr(CLng(TempList(i)),MaxVal,True)
			End If
		Next i
	End If
	DecStrListToHexStrList = TempList
	Exit Function
	ExitFunction:
	ReDim TempList(0) As String
	DecStrListToHexStrList = TempList
End Function


'批量转换十六进制字符为十进制字符
Public Function HexStrListToDecStrList(StrList() As String) As String()
	Dim i As Long,TempList() As String
	On Error GoTo ExitFunction
	TempList = StrList
	For i = 0 To UBound(TempList)
		If TempList(i) <> "" Then
			TempList(i) = CStr$(StrToVal(TempList(i),True))
		End If
	Next i
	HexStrListToDecStrList = TempList
	Exit Function
	ExitFunction:
	ReDim TempList(0) As String
	HexStrListToDecStrList = TempList
End Function


'每二个字符空格分隔
Public Function SeparatHex(ByVal textStr As String) As String
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
Public Function Str2Hex(ByVal textStr As String,ByVal CodePage As Long,ByVal FillMode As Long,Optional ByVal ByteLength As Long) As String
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


'检查输入的 Hex 字串是否符合要求
'返回 CheckHex = 0 合格, = 1 HEX 代码字符数不符, = 2 数字数不符, = 3 包含非法字符 = 4 没有要转换的代码(仅 Ignore = 1 的情况)
Public Function CheckHex(ByVal textStr As String,ByVal CodePage As Long,ByVal Mode As Integer,ByVal Ignore As Integer) As Long
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


'获取文件的类型，PE 还是 MAC 还是非 PE 文件
Public Function GetFileFormat(ByVal FilePath As String,ByVal Mode As Long,FileType As Integer) As String
	Dim i As Long,n As Long,FN As FILE_IMAGE
	On Error GoTo ExitFunction
	FileType = 0
	'打开文件
	n = FileLen(FilePath)
	If n = 0 Then Exit Function
	If n > 4096 + 16 Then n = 4096 + 16
	Mode = LoadFile(FilePath,FN,0,0,i,Mode)
	If Mode < -1 Then Exit Function
	Do While i < n - 16
		Select Case GetLong(FN,i,Mode)
		Case MH_MAGIC_32,MH_MAGIC_64,MH_MAGIC_FAT,MH_MAGIC_FAT_CIGAM
			GetFileFormat = "MAC"
			FileType = i
			Exit Do
		Case Else
			Select Case GetInteger(FN,i,Mode)
			Case IMAGE_DOS_SIGNATURE,IMAGE_OS2_SIGNATURE,IMAGE_OS2_SIGNATURE_LE,IMAGE_VXD_SIGNATURE
				GetFileFormat = "PE"
				FileType = i
				Exit Do
			End Select
		End Select
		i = getNotNullByteRegExp(FN,i + 1,4096,Mode)
	Loop
	ExitFunction:
	UnLoadFile(FN,0,Mode)
End Function


'检查文件是否已被打开或占用
Public Function IsOpen(ByVal strFilePath As String,Optional ByVal Continue As Long = 2,Optional ByVal WaitTime As Double = 0.5) As Boolean
	Dim i As Long,FN As Variant
	'尝试打开文件
	If Dir$(strFilePath) = "" Then Exit Function
	On Error Resume Next
	FN = FreeFile
	Do
		Open strFilePath For Binary Access Read Lock Write As #FN
		If Err.Number = 0 Or i > 10 Then Exit Do
		Wait WaitTime
		Err.Clear
		i = i + 1
	Loop
	Close #FN
	If Err.Number = 0 Then Exit Function
	ErrorHandle:
	IsOpen = True
	Err.Source = "NotOpenFile"
	Err.Description = Err.Description & JoinStr & strFilePath
	Call sysErrorMassage(Err,Continue)
End Function


'获取文件的创建、访问、修改日期
'Mode = 0 创建日期
'Mode = 1 访问日期
'Mode = 2 修改日期
Public Function GetFileDate(ByVal strFilePath As String,ByVal Mode As Long) As Date
	Dim fso As Object,Obj As Object
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set Obj = fso.GetFile(strFilePath)
	If Mode = 0 Then
		GetFileDate = Obj.DateCreated
	ElseIf Mode = 1 Then
		GetFileDate = Obj.DateLastAccessed
	Else
		GetFileDate = Obj.DateLastModified
	End If
	Set fso = Nothing: Set Obj = Nothing
End Function


'修改文件的创建时间、访问时间和修改时间
Public Function SetFileCreatedDate(ByVal strFilePath As String,ByVal CreatedDate As Date) As Boolean
	Dim lhFile As Long,udtFileTime As FILETIME
	Dim udtLocalTime As FILETIME,udtSystemTime As SYSTEMTIME

	'转换时间为 SYSTEMTIME
	udtSystemTime.wYear = Year(CreatedDate)
	udtSystemTime.wMonth = Month(CreatedDate)
	udtSystemTime.wDay = Day(CreatedDate)
	udtSystemTime.wDayOfWeek = Weekday(CreatedDate)
	udtSystemTime.wHour = Hour(CreatedDate)
	udtSystemTime.wMinute = Minute(CreatedDate)
	udtSystemTime.wSecond = Second(CreatedDate)
	udtSystemTime.wMilliseconds = 0

	'转换系统时间为本地时间
	If SystemTimeToFileTime(udtSystemTime, udtLocalTime) = 0 Then
		DisplayError "SystemTimeToFileTime",GetLastError()
		Exit Function
	End If

	'转换本地时间为 GMT 时间
	If LocalFileTimeToFileTime(udtLocalTime, udtFileTime) = 0 Then
		DisplayError "LocalFileTimeToFileTime",GetLastError()
		Exit Function
	End If

	'打开文件获取文件句柄
	lhFile = CreateFile(strFilePath, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
	If lhFile = 0 Then
		DisplayError "CreateFile",GetLastError()
		Exit Function
	End If

	'更改文件的日期和时间属性
	If SetFileTime(lhFile, udtFileTime, udtFileTime, udtFileTime) = 0 Then
		DisplayError "SetFileTime",GetLastError()
		If CloseHandle(lhFile) = 0 Then
			DisplayError "CloseHandle",GetLastError()
		End If
		Exit Function
	End If

	'关闭文件句柄，标记成功
	If CloseHandle(lhFile) = 0 Then
		DisplayError "CloseHandle",GetLastError()
		Exit Function
	End If
	SetFileCreatedDate = True
End Function


'加载文件
'ImageSize = 0 按文件的初始大小打开，否则按指定大小打开
'ReadOnly = 0 按只读方式打开，否则读写方式打开
'ImageByte = 0 不获取字节数组只初始化(缓存方式获取所有字节)，否则按 ImageByte 指定大小获取
'Mode < 0 直接方式，Mode = 0 缓存方式，Mode > 0 映射方式
'IsPE = 0 按一般文件映射，否则按 PE 文件映射(每个节对齐)
'LoadFile = -2 打开失败，否则实际打开方式
'LoadedImage 打开文件后获取的数据
Public Function LoadFile(ByVal strFilePath As String,LoadedImage As FILE_IMAGE,ByVal ImageSize As Long,ByVal ReadOnly As Long, _
		ByVal ImageByte As Long,ByVal Mode As Long,Optional ByVal IsPE As Long,Optional ByVal WaitTime As Double = 0.5) As Long
	Dim i As Long
	'尝试打开文件
	If IsOpen(strFilePath,2,WaitTime) = True Then
		LoadFile = -2
		Exit Function
	End If
	'加载文件
	On Error GoTo ErrorHandle
	LoadedImage.ModuleName = strFilePath
	If Mode < 0 Then
		LoadedImage.hFile = FreeFile
		If ReadOnly = 0 Then
			Open strFilePath For Binary Access Read Lock Write As #LoadedImage.hFile
		Else
			Open strFilePath For Binary Access Read Write Lock Write As #LoadedImage.hFile
		End If
		LoadedImage.SizeOfFile = LOF(LoadedImage.hFile)
		If ImageByte > 0 Then
			LoadedImage.SizeOfImage = ImageByte
			LoadedImage.ImageByte = GetBytes(LoadedImage,ImageByte,0,-1)
		Else
			LoadedImage.SizeOfImage = LoadedImage.SizeOfFile
			ReDim LoadedImage.ImageByte(0) 'As Byte
		End If
	Else
		If Mode > 0 Then
			If MapFile(strFilePath,LoadedImage,ImageSize,ReadOnly,0,IsPE) = True Then
				If ImageByte > 0 Then
					'LoadedImage.ImageByte = getBytesByMap(LoadedImage.MappedAddress,LoadedImage.SizeOfFile)
					LoadedImage.ImageByte = GetBytes(LoadedImage,ImageByte,0,Mode)
				Else
					ReDim LoadedImage.ImageByte(0) 'As Byte
				End If
			Else
				UnMapFile(LoadedImage,0)
				Mode = 0
			End If
		End If
		If Mode = 0 Then
			LoadedImage.hFile = FreeFile
			Open strFilePath For Binary Access Read Lock Write As #LoadedImage.hFile
			LoadedImage.SizeOfFile = LOF(LoadedImage.hFile)
			If ImageByte > 0 Then
				LoadedImage.SizeOfImage = ImageByte
				LoadedImage.ImageByte = GetBytes(LoadedImage,ImageByte,0,-1)
			Else
				LoadedImage.SizeOfImage = LoadedImage.SizeOfFile
				LoadedImage.ImageByte = GetBytes(LoadedImage,LoadedImage.SizeOfFile,0,-1)
			End If
			Close #LoadedImage.hFile
		End If
	End If
	LoadFile = Mode
	Exit Function
	ErrorHandle:
	UnLoadFile(LoadedImage,0,Mode)
	LoadFile = -2
	Err.Source = IIf(ReadOnly = 0,"NotOpenFile","NotWriteFile")
	Err.Description = Err.Description & JoinStr & strFilePath
	Call sysErrorMassage(Err,3)
End Function


'卸载文件
'SizeOfFile = 0 不写入，否则按指定大小写入，仅在缓存和映射方式时有效
Public Function UnLoadFile(LoadedImage As FILE_IMAGE,ByVal SizeOfFile As Long,ByVal Mode As Long) As Boolean
	On Error GoTo ErrorHandle
	If Mode < 0 Then
		Close #LoadedImage.hFile
		LoadedImage.ModuleName = ""
		LoadedImage.hFile = 0
		LoadedImage.hMap = 0
		LoadedImage.MappedAddress = 0
		LoadedImage.SizeOfFile = 0
		LoadedImage.SizeOfImage = 0
		Erase LoadedImage.ImageByte
		UnLoadFile = True
	ElseIf Mode = 0 Then
		If SizeOfFile > 0 Then
			LoadedImage.hFile = FreeFile
			Open LoadedImage.ModuleName For Binary Access Write Lock Write As #LoadedImage.hFile
			ReDim Preserve LoadedImage.ImageByte(LoadedImage.SizeOfFile - 1) 'As Byte
			Put #LoadedImage.hFile,1,LoadedImage.ImageByte
			Close #LoadedImage.hFile
		End If
		LoadedImage.ModuleName = ""
		LoadedImage.hFile = 0
		LoadedImage.hMap = 0
		LoadedImage.MappedAddress = 0
		LoadedImage.SizeOfFile = 0
		LoadedImage.SizeOfImage = 0
		Erase LoadedImage.ImageByte
		UnLoadFile = True
	ElseIf UnMapFile(LoadedImage,SizeOfFile) = True Then
		UnLoadFile = True
	End If
	ErrorHandle:
End Function


'映射文件
'MapSize = 0 按文件初始时的大小映射，否则按指定大小映射
'ReadOnly = 0 只读方式，否则读写方式
'SizeOfFile = 0 获取文件初始时的大小，否则不获取
'IsPE = 0 按一般文件映射，否则按 PE 文件映射(每个节对齐)
Public Function MapFile(ByVal strFilePath As String,LoadedImage As FILE_IMAGE,ByVal MapSize As Long, _
		ByVal ReadOnly As Long,Optional ByVal SizeOfFile As Long,Optional ByVal IsPE As Long) As Boolean
	With LoadedImage
		If UnMapFile(LoadedImage,SizeOfFile) = False Then Exit Function
		.ModuleName = strFilePath
		.hFile = CreateFile(strFilePath, _
				GENERIC_READ Or GENERIC_WRITE, _
				FILE_SHARE_READ Or FILE_SHARE_WRITE, _
				0&, _
				OPEN_EXISTING, _
				FILE_ATTRIBUTE_NORMAL Or FILE_ATTRIBUTE_ARCHIVE Or _
				FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_HIDDEN Or _
				FILE_ATTRIBUTE_SYSTEM, _
				0&)
		If .hFile = 0 Then Exit Function
		If SizeOfFile = 0 Then .SizeOfFile = GetFileSize(.hFile, 0&)
		If IsPE = 0 Then
			.hMap = CreateFileMapping(.hFile, 0&, PAGE_READWRITE, 0, MapSize, vbNullChar)
		Else
			.hMap = CreateFileMapping(.hFile, 0&, PAGE_EXECUTE_READWRITE Or SEC_IMAGE, 0, MapSize, vbNullChar)
		End If
		If .hMap = 0 Then
			DisplayError "CreateFileMapping",GetLastError()
			If CloseHandle(.hFile) = 0 Then
				DisplayError "CloseHandle",GetLastError()
			End If
			Exit Function
		End If
		If ReadOnly = 0 Then
			.MappedAddress = MapViewOfFile(.hMap, FILE_MAP_READ, 0, 0, MapSize)
		Else
			.MappedAddress = MapViewOfFile(.hMap, FILE_MAP_WRITE, 0, 0, MapSize)
		End If
		If .MappedAddress = 0 Then
			DisplayError "MapViewOfFile",GetLastError()
			If CloseHandle(.hMap) = 0 Then
				DisplayError "CloseHandle",GetLastError()
			End If
			If CloseHandle(.hFile) = 0 Then
				DisplayError "CloseHandle",GetLastError()
			End If
			Exit Function
    	End If
    	If MapSize = 0 Then
    		.SizeOfImage = GetFileSize(.hFile, 0&)
    	Else
    		.SizeOfImage = MapSize
    	End If
    	MapFile = True
    End With
End Function


'关闭映射文件
'SizeOfFile = 0 按文件实际大小保存，否则按指定大小保存
Private Function UnMapFile(LoadedImage As FILE_IMAGE,Optional ByVal SizeOfFile As Long) As Boolean
	With LoadedImage
		If .MappedAddress > 0 Then
			If UnmapViewOfFile(.MappedAddress) = 0 Then
				DisplayError "UnmapViewOfFile",GetLastError()
				Exit Function
			Else
				.MappedAddress = 0
			End If
		End If
		If .hMap > 0 Then
 			If CloseHandle(.hMap) = 0 Then
 				DisplayError "CloseHandle",GetLastError()
 				Exit Function
 			Else
 				.hMap = 0
 			End If
 		End If
 		If .hFile > 0 Then
 			If SizeOfFile > 0 Then
 				SetFilePointer(.hFile, SizeOfFile, 0&, 0)
 				SetEndOfFile(.hFile)
 				'WriteFile(.hFile,Mode, SizeOfFile, 0&, 0&)
 			End If
 			If CloseHandle(.hFile) = 0 Then
 				DisplayError "CloseHandle",GetLastError()
 				Exit Function
 			Else
 				.hFile = 0
 			End If
 		End If
 		If .MappedAddress + .hMap + .hFile = 0 Then
 			.ModuleName = ""
 			.SizeOfImage = 0
 			If SizeOfFile = 0 Then .SizeOfFile = 0
 			Erase LoadedImage.ImageByte
 			UnMapFile = True
 		End If
 	End With
End Function


'获取错误消息
Private Sub DisplayError(ByVal lpSource As String,ByVal dwError As Long)
    Dim GetLastDllErrMsg As String
    GetLastDllErrMsg = String$(256, 32)
    FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
    			FORMAT_MESSAGE_IGNORE_INSERTS, _
    			lpSource, dwError, 0&, _
    			GetLastDllErrMsg,Len(GetLastDllErrMsg),0&)
    Err.Raise(dwError,lpSource,Trim(GetLastDllErrMsg))
End Sub


'比较二个文件的文件节开始地址和大小是否一致，相同返回 0
'SecIDList 为数组时，检查制定节
Public Function CompSecHeader(srcFile As FILE_PROPERTIE,trgFile As FILE_PROPERTIE,Optional ByVal SecIDList As Variant) As Long
	Dim i As Integer,j As Integer,k As Variant
	k = trgFile.MaxSecIndex - 1
	If IsArray(SecIDList) Then
		For j = 0 To UBound(SecIDList)
			i = CInt(SecIDList(j))
			If i > k Then
				CompSecHeader = -1
				Exit For
			ElseIf srcFile.SecList(i).lPointerToRawData <> trgFile.SecList(i).lPointerToRawData Then
				CompSecHeader = 1
				Exit For
			ElseIf srcFile.SecList(i).lSizeOfRawData <> trgFile.SecList(i).lSizeOfRawData Then
				CompSecHeader = 1
				Exit For
			ElseIf srcFile.SecList(i).lVirtualAddress <> trgFile.SecList(i).lVirtualAddress Then
				CompSecHeader = 2
				Exit For
			ElseIf srcFile.SecList(i).lVirtualSize <> trgFile.SecList(i).lVirtualSize Then
				CompSecHeader = 2
				Exit For
			End If
		Next j
	Else
		j = srcFile.MaxSecIndex - 1
		j = IIf(j > k,k,j)
		Select Case srcFile.Magic
		Case "PE32","NET32","PE64","NET64"
			k = srcFile.DataDirectory(2).lPointerToRawData
		Case Else
			k = srcFile.SecList(srcFile.MaxSecID).lPointerToRawData + srcFile.SecList(srcFile.MaxSecID).lSizeOfRawData - 1
		End Select
		For i = 0 To j
			If srcFile.SecList(i).lPointerToRawData < k Then
				If srcFile.SecList(i).lPointerToRawData <> trgFile.SecList(i).lPointerToRawData Then
					CompSecHeader = 1
					Exit For
				ElseIf srcFile.SecList(i).lSizeOfRawData <> trgFile.SecList(i).lSizeOfRawData Then
					CompSecHeader = 1
					Exit For
				ElseIf srcFile.SecList(i).lVirtualAddress <> trgFile.SecList(i).lVirtualAddress Then
					CompSecHeader = 2
					Exit For
				ElseIf srcFile.SecList(i).lVirtualSize <> trgFile.SecList(i).lVirtualSize Then
					CompSecHeader = 2
					Exit For
				End If
			End If
		Next i
	End If
End Function


'获取文件节索引号
'Mode = 0 检查偏移地址(不包括隐藏节)
'Mode = 1 检查偏移地址(包括隐藏节)
'Mode = 2 检查相对虚拟地址(不包括隐藏节)
'Mode = 3 检查相对虚拟地址(包括隐藏节)
'返回文件节索引号、MinVal、MaxVal 值
Public Function SkipSection(File As FILE_PROPERTIE,ByVal Offset As Long,MinVal As Long,MaxVal As Long,Optional ByVal Mode As Long) As Long
	Dim i As Integer
	SkipSection = -1
	MinVal = 0: MaxVal = 0
	If Offset < 0 Then Exit Function
	If Mode < 2 Then
		For i = 0 To File.MaxSecIndex - 1
			With File.SecList(i)
				If Offset >= .lPointerToRawData And Offset < .lPointerToRawData + .lSizeOfRawData Then
					SkipSection = i
					MinVal = .lPointerToRawData
					MaxVal = MinVal + .lSizeOfRawData - 1
					Exit For
				End If
			End With
		Next i
		If Mode = 1 And SkipSection = -1 Then
			With File.SecList(File.MaxSecIndex)
				If .lSizeOfRawData > 0 Then
					If Offset >= .lPointerToRawData And Offset < .lPointerToRawData + .lSizeOfRawData Then
						SkipSection = File.MaxSecIndex
						MinVal = .lPointerToRawData
						MaxVal = MinVal + .lSizeOfRawData - 1
					End If
				End If
			End With
		End If
		If SkipSection = -1 Then
			If File.Magic <> "" Then
				If Offset < File.SecList(File.MinSecID).lPointerToRawData Then
					SkipSection = -2	'文件头
				ElseIf Offset > File.SecList(File.MaxSecIndex).lPointerToRawData Then
					SkipSection = -3	'子 PE 文件
				End If
			End If
		End If
	Else
		For i = 0 To File.MaxSecIndex - 1
			With File.SecList(i)
				If Offset >= .lVirtualAddress And Offset < .lVirtualAddress + .lVirtualSize Then
					SkipSection = i
					MinVal = .lVirtualAddress
					MaxVal = MinVal + .lVirtualSize - 1
					Exit For
				End If
			End With
		Next i
		If Mode = 3 And SkipSection = -1 Then
			With File.SecList(File.MaxSecIndex)
				If .lVirtualSize > 0 Then
					If Offset >= .lVirtualAddress And Offset < .lVirtualAddress + .lVirtualSize Then
						SkipSection = File.MaxSecIndex
						MinVal = .lVirtualAddress
						MaxVal = MinVal + .lVirtualSize - 1
					End If
				End If
			End With
		End If
		If SkipSection = -1 Then
			If File.Magic <> "" Then
				If Offset < File.SecList(File.MinSecID).lVirtualAddress Then
					SkipSection = -2	'文件头
				ElseIf Offset > File.SecList(File.MaxSecIndex).lVirtualAddress Then
					SkipSection = -3	'子 PE 文件
				End If
			End If
		End If
	End If
End Function


'获取子文件节索引号
'Mode = False 检查偏移地址
'Mode = True 检查相对虚拟地址
'返回文件节索引号、MinVal、MaxVal 值
Public Function SkipSubSection(Sec As SECTION_PROPERTIE,ByVal Offset As Long,MinVal As Long,MaxVal As Long,Optional ByVal Mode As Boolean) As Long
	Dim i As Integer
	SkipSubSection = -1
	MinVal = 0: MaxVal = 0
	If Sec.SubSecs = 0 Then Exit Function
	If Offset < 0 Then Exit Function
	If Mode = False Then
		For i = 0 To Sec.SubSecs - 1
			With Sec.SubSecList(i)
				If Offset >= .lPointerToRawData And Offset < .lPointerToRawData + .lSizeOfRawData Then
					SkipSubSection = i
					MinVal = .lPointerToRawData
					MaxVal = MinVal + .lSizeOfRawData - 1
					Exit For
				End If
			End With
		Next i
	Else
		For i = 0 To Sec.SubSecs - 1
			With Sec.SubSecList(i)
				If Offset >= .lVirtualAddress And Offset < .lVirtualAddress + .lVirtualSize Then
					SkipSubSection = i
					MinVal = .lVirtualAddress
					MaxVal = MinVal + .lVirtualSize - 1
					Exit For
				End If
			End With
		Next i
	End If
End Function


'获取各种文件头的索引号及地址 (这里的地址都已转换为偏移地址)
'Mode = 0 时，仅返回 RVA 所在目录的索引号
'Mode = 1 时，RVA = RVA 所在目录的最大地址 + 1，SkipVal = 比 RVA 大的目录最小地址
'Mode > 1 时，RVA = RVA 所在目录的最小地址 - 1，SkipVal = 比 RVA 小的目录最大地址
'fType = 0 时，跳过 .NET US 区段
Public Function SkipHeader(File As FILE_PROPERTIE,RVA As Long,Optional SkipVal As Long,Optional ByVal Mode As Long,Optional ByVal fType As Long) As Long
	Dim i As Integer,j As Integer,EndPos As Long
	SkipHeader = -1
	i = 15 + IIf(File.LangType = NET_FILE_SIGNATURE,7 + File.NetStreams,0)
	ReDim List(i) As IMAGE_DATA_DIRECTORY
	With File
		If .DataDirs > 0 Then
			For i = 0 To .DataDirs - 1
				List(j).lVirtualAddress = .DataDirectory(i).lPointerToRawData
				List(j).lSize = .DataDirectory(i).lSizeOfRawData
				j = j + 1
			Next i
		End If
		If .LangType = NET_FILE_SIGNATURE Then
			For i = 0 To UBound(.CLRList)
				List(j).lVirtualAddress = .CLRList(i).lPointerToRawData
				List(j).lSize = .CLRList(i).lSizeOfRawData
				j = j + 1
			Next i
		End If
		If .NetStreams > 0 Then
			For i = 0 To .NetStreams - 1
				List(j).lVirtualAddress  = .StreamList(i).lPointerToRawData
				List(j).lSize = .StreamList(i).lSizeOfRawData
				If i = .USStreamID And fType > 0 Then
					fType = 0
					EndPos = .StreamList(i).lPointerToRawData + .StreamList(i).lSizeOfRawData - 1
					If RVA >= .StreamList(i).lPointerToRawData And RVA <= EndPos Then
						fType = RVA
					ElseIf Mode = 1 And RVA < EndPos Then
						fType = .StreamList(i).lPointerToRawData
					ElseIf Mode > 1 And RVA > .StreamList(i).lPointerToRawData Then
						fType = .StreamList(i).lPointerToRawData
					End If
				End If
				j = j + 1
			Next i
		Else
			fType = 0
		End If
	End With
	If j = 0 Then Exit Function
	ReDim Preserve List(j - 1)	'As IMAGE_DATA_DIRECTORY
	For i = 0 To j - 1
		With List(i)
			If .lVirtualAddress > 0 And .lSize > 0 Then
				EndPos = .lVirtualAddress + .lSize - 1		'最大地址
				Select Case Mode
				Case 0
					If SkipHeader < 0 Then
						If RVA >= .lVirtualAddress And RVA <= EndPos Then
							SkipHeader = i
						End If
					Else
						If fType = 0 Then Exit Function
						If .lVirtualAddress >= List(SkipHeader).lVirtualAddress And _
							EndPos < List(SkipHeader).lVirtualAddress + List(SkipHeader).lSize Then
							If .lVirtualAddress <= fType And EndPos >= fType Then
								SkipHeader = -1
								Exit Function
							End If
						End If
					End If
				Case 1
					If .lVirtualAddress > RVA Then
						If RVA > SkipVal Then
							SkipVal = .lVirtualAddress
						ElseIf .lVirtualAddress < SkipVal Then
							SkipVal = .lVirtualAddress
						End If
					ElseIf RVA <= EndPos Then
						If fType > 0 Then
							If .lVirtualAddress <= fType And EndPos >= fType Then SkipHeader = i
						Else
							 SkipHeader = i
						End If
						RVA = EndPos + 1
						If SkipVal < EndPos Then SkipVal = EndPos
					ElseIf SkipHeader > -1 Then
						If .lVirtualAddress >= List(SkipHeader).lVirtualAddress And _
							EndPos < List(SkipHeader).lVirtualAddress + List(SkipHeader).lSize Then
							If .lVirtualAddress <= fType And EndPos >= fType Then
								SkipHeader = -1
								RVA = fType
								SkipVal = EndPos
								Exit Function
							End If
						End If
					End If
				Case Else
					If EndPos < RVA Then
						If RVA < SkipVal Then
							SkipVal = EndPos
						ElseIf EndPos > SkipVal Then
							SkipVal = EndPos
						End If
					ElseIf RVA >= .lVirtualAddress Then
						If fType > 0 Then
							If .lVirtualAddress <= fType And EndPos >= fType Then SkipHeader = i
						Else
							 SkipHeader = i
						End If
						RVA = .lVirtualAddress - 1
						If SkipVal > .lVirtualAddress Then SkipVal = .lVirtualAddress
					ElseIf SkipHeader > -1 Then
						If .lVirtualAddress >= List(SkipHeader).lVirtualAddress And _
							EndPos < List(SkipHeader).lVirtualAddress + List(SkipHeader).lSize Then
							If .lVirtualAddress <= fType And EndPos >= fType Then
								SkipHeader = -1
								RVA = fType - 1
								SkipVal = .lVirtualAddress
								Exit Function
							End If
						End If
					End If
				End Select
			End If
		End With
	Next i
End Function


'获取节表的最大和最小或比大或比小的地址所在节索引号
'MinID = 0 和 MaxID = 0 获取节表的最大和最小地址所在节索引号
'MinID = -1 获取比 MaxID 节小的地址所在节索引号
'MaxID = -1 获取比 MinID 节大的地址所在节索引号
'Mode = False 比较偏移地址，否则比较相对虚拟地址
Public Function GetSectionID(File As FILE_PROPERTIE,MinID As Integer,MaxID As Integer,ByVal Mode As Boolean) As Long
	Dim i As Long,MinVal As Variant,MaxVal As Variant
	With File
		If Mode = False Then
			If MinID = 0 And MaxID = 0 Then
				For i = 0 To .MaxSecIndex - 1
					If .SecList(i).lPointerToRawData > 0 Then
						If MinVal = 0 Then
							MinVal = .SecList(i).lPointerToRawData
						ElseIf .SecList(i).lPointerToRawData <= MinVal Then
							MinVal = .SecList(i).lPointerToRawData
							MinID = i
						End If
						If .SecList(i).lPointerToRawData >= MaxVal Then
							MaxVal = .SecList(i).lPointerToRawData
							MaxID = i
						End If
					End If
				Next i
			ElseIf MinID < 0 Then
				MaxVal = .SecList(MaxID).lPointerToRawData
				For i = 0 To .MaxSecIndex - 1
					If .SecList(i).lSizeOfRawData > 0 Then
						If .SecList(i).lPointerToRawData < MaxVal Then
							If MinVal = 0 Then
								MinVal = .SecList(i).lPointerToRawData
								MinID = i
							ElseIf .SecList(i).lPointerToRawData > MinVal Then
								MinVal = .SecList(i).lPointerToRawData
								MinID = i
							End If
						End If
					End If
				Next i
				GetSectionID = IIf(MinID < 0,MaxID,MinID)
			ElseIf MaxID < 0 Then
				MinVal = .SecList(MinID).lPointerToRawData
				For i = 0 To .MaxSecIndex - 1
					If .SecList(i).lSizeOfRawData > 0 Then
						If .SecList(i).lPointerToRawData > MinVal Then
							If MaxVal = 0 Then
								MaxVal = .SecList(i).lPointerToRawData
								MaxID = i
							ElseIf .SecList(i).lPointerToRawData < MaxVal Then
								MaxVal = .SecList(i).lPointerToRawData
								MaxID = i
							End If
						End If
					End If
				Next i
				GetSectionID = IIf(MaxID < 0,MinID,MaxID)
			End If
		Else
			If MinID = 0 And MaxID = 0 Then
				For i = 0 To .MaxSecIndex - 1
					If .SecList(i).lVirtualAddress > 0 Then
						If MinVal = 0 Then
							MinVal = .SecList(i).lVirtualAddress
						ElseIf .SecList(i).lVirtualAddress <= MinVal Then
							MinVal = .SecList(i).lVirtualAddress
							MinID = i
						End If
						If .SecList(i).lVirtualAddress >= MaxVal Then
							MaxVal = .SecList(i).lVirtualAddress
							MaxID = i
						End If
					End If
				Next i
			ElseIf MinID < 0 Then
				MaxVal = .SecList(MaxID).lVirtualAddress
				For i = 0 To .MaxSecIndex - 1
					If .SecList(i).lVirtualSize > 0 Then
						If .SecList(i).lVirtualAddress < MaxVal Then
							If MinVal = 0 Then
								MinVal = .SecList(i).lVirtualAddress
								MinID = i
							ElseIf .SecList(i).lVirtualAddress > MinVal Then
								MinVal = .SecList(i).lVirtualAddress
								MinID = i
							End If
						End If
					End If
				Next i
				GetSectionID = IIf(MinID < 0,MaxID,MinID)
			ElseIf MaxID < 0 Then
				MinVal = .SecList(MinID).lVirtualAddress
				For i = 0 To .MaxSecIndex - 1
					If .SecList(i).lVirtualSize > 0 Then
						If .SecList(i).lVirtualAddress > MinVal Then
							If MaxVal = 0 Then
								MaxVal = .SecList(i).lVirtualAddress
								MaxID = i
							ElseIf .SecList(i).lVirtualAddress < MaxVal Then
								MaxVal = .SecList(i).lVirtualAddress
								MaxID = i
							End If
						End If
					End If
				Next i
				GetSectionID = IIf(MaxID < 0,MinID,MaxID)
			End If
		End If
	End With
End Function


'检查和更改所有文件节名称列表
Public Sub ChangSectionNames(File As FILE_PROPERTIE,Optional ByVal HideSecName As String,Optional ByVal NoPEName As String)
	Dim i As Integer,Dic As Object
	With File
		Select Case .Magic
		Case "","NotPE32","NotPE64"
			.SecList(0).sName = NoPEName
			.SecList(1).sName = HideSecName
		Case Else
			Set Dic = CreateObject("Scripting.Dictionary")
			For i = 0 To .MaxSecIndex - 1
				If Trim$(.SecList(i).sName) = "" Then
					.SecList(i).sName = "Untitled_" & CStr$(i)
				ElseIf Dic.Exists(.SecList(i).sName) Then
					.SecList(i).sName = .SecList(i).sName & "_" & CStr$(i)
				Else
					Dic.Add(.SecList(i).sName,"")
				End If
			Next i
			.SecList(i).sName = HideSecName
			Set Dic = Nothing
		End Select
	End With
End Sub


'获取所有文件节名称或包括地址范围的列表
'Mode = 0 获取所有主区段的名称，包括隐藏区段
'Mode = 1 获取所有主区段的名称，不包括隐藏区段
'Mode = 2 获取所有主区段的名称和子区段名称，包括隐藏区段
'Mode = 3 获取所有主区段的名称和子区段名称，不包括隐藏区段
'Mode = 4 获取所有主区段的名称和地址范围，包括隐藏区段
'Mode = 5 获取所有主区段的名称和地址范围，不包括隐藏区段

'Mode = 6 获取所有主区段的名称和子区段名称和开始地址(菜单形式)，包括隐藏区段
'Mode = 7 获取所有主区段的名称和子区段名称和开始地址(菜单形式)，不包括隐藏区段
'Mode = 8 获取所有主区段的名称和子区段名称和结束地址(菜单形式)，包括隐藏区段
'Mode = 9 获取所有主区段的名称和子区段名称和结束地址(菜单形式)，不包括隐藏区段

'Mode = 10 获取所有主区段的名称和子区段ID组合，包括隐藏区段

'Mode = 11 获取所有主区段和子区段的开始地址(仅值)，包括隐藏区段
'Mode = 12 获取所有主区段和子区段的开始地址(仅值)，不包括隐藏区段
'Mode = 13 获取所有主区段和子区段的结束地址(仅值)，包括隐藏区段
'Mode = 14 获取所有主区段和子区段的结束地址(仅值)，不包括隐藏区段

'无子节输出格式 主节名称 (主节地址)
'有子节输出格式 主节名称 - 子节名称 (子节地址)
Public Function getSectionList(SecList() As SECTION_PROPERTIE,Optional ByVal Mode As Integer, _
					Optional ByVal MaxVal As Long,Optional ByVal DisPlayFormat As Boolean) As String()
	Dim i As Integer,j As Integer,k As Integer,n As Integer
	Select Case Mode
	Case 0,1	'获取所有主区段的名称
		n = UBound(SecList)
		If Mode = 0 Then
			If SecList(n).lSizeOfRawData = 0 Then n = n - 1
		Else
			n = n - 1
		End If
		ReDim TempList(n) As String
		For i = 0 To n
			TempList(i) = SecList(i).sName
		Next i
	Case 2,3	'获取所有主区段的名称和子区段名称
		n = UBound(SecList)
		If Mode = 2 Then
			If SecList(n).lSizeOfRawData = 0 Then n = n - 1
		Else
			n = n - 1
		End If
		ReDim TempList(n) As String
		For i = 0 To n
			If SecList(i).SubSecs = 0 Then
				ReDim Preserve TempList(k + 1) As String
				TempList(k) = SecList(i).sName
				k = k + 1
			Else
				ReDim Preserve TempList(k + SecList(i).SubSecs) As String
				For j = 0 To SecList(i).SubSecs - 1
					TempList(k) = SecList(i).sName & "-" & SecList(i).SubSecList(j).sName
					k = k + 1
				Next j
			End If
		Next i
		If k > 0 Then k = k - 1
		ReDim Preserve TempList(k) As String
	Case 4,5	'获取所有主区段的名称和地址范围
		n = UBound(SecList)
		If Mode = 4 Then
			If SecList(n).lSizeOfRawData = 0 Then n = n - 1
		Else
			n = n - 1
		End If
		ReDim TempList(n) As String
		For i = 0 To n
			TempList(i) = SecList(i).sName & " (" & SecList(i).lPointerToRawData & "-" & _
						SecList(i).lPointerToRawData + SecList(i).lSizeOfRawData - 1 & ")"
		Next i
	Case 6,7	'获取所有主区段的名称和子区段名称和开始地址(菜单形式)
		n = UBound(SecList)
		If Mode = 6 Then
			If SecList(n).lSizeOfRawData = 0 Then n = n - 1
		Else
			n = n - 1
		End If
		ReDim TempList(n) As String
		For i = 0 To n
			If SecList(i).SubSecs = 0 Then
				TempList(i) = SecList(i).sName & " (" & _
							ValToStr(SecList(i).lPointerToRawData,MaxVal,DisPlayFormat) & ")"
			Else
				ReDim TempArray(SecList(i).SubSecs - 1) As String
				For j = 0 To SecList(i).SubSecs - 1
					TempArray(j) = SecList(i).SubSecList(j).sName & " (" & _
								ValToStr(SecList(i).SubSecList(j).lPointerToRawData,MaxVal,DisPlayFormat) & ")"
				Next j
				TempList(i) = SecList(i).sName & vbNullChar & StrListJoin(TempArray,vbNullChar)
			End If
		Next i
	Case 8,9	'获取所有主区段的名称和子区段名称和结束地址(菜单形式)
		n = UBound(SecList)
		If Mode = 8 Then
			If SecList(n).lSizeOfRawData = 0 Then n = n - 1
		Else
			n = n - 1
		End If
		ReDim TempList(n) As String
		For i = 0 To n
			If SecList(i).SubSecs = 0 Then
				TempList(i) = SecList(i).sName & " (" & _
						ValToStr(SecList(i).lPointerToRawData + _
						SecList(i).lSizeOfRawData - 1,MaxVal,DisPlayFormat) & ")"
			Else
				ReDim TempArray(SecList(i).SubSecs - 1) As String
				For j = 0 To SecList(i).SubSecs - 1
					TempArray(j) = SecList(i).SubSecList(j).sName & " (" & _
								ValToStr(SecList(i).SubSecList(j).lPointerToRawData + _
								SecList(i).SubSecList(j).lSizeOfRawData - 1,MaxVal,DisPlayFormat) & ")"
				Next j
				TempList(i) = SecList(i).sName & vbNullChar & StrListJoin(TempArray,vbNullChar)
			End If
		Next i
	Case 10		'获取所有主区段的名称和子区段ID组合
		n = UBound(SecList)
		If SecList(n).lSizeOfRawData = 0 Then n = n - 1
		ReDim TempList(n) As String
		For i = 0 To n
			If SecList(i).SubSecs = 0 Then
				ReDim Preserve TempList(k + 1) As String
				TempList(k) = CStr$(i)
				k = k + 1
			Else
				ReDim Preserve TempList(k + SecList(i).SubSecs) As String
				For j = 0 To SecList(i).SubSecs  - 1
					TempList(k) = CStr$(i) & "-" & CStr$(j)
					k = k + 1
				Next j
			End If
		Next i
		If k > 0 Then k = k - 1
		ReDim Preserve TempList(k) As String
	Case 11,12	'获取所有主区段和子区段的开始地址(仅值)
		n = UBound(SecList)
		If Mode = 11 Then
			If SecList(n).lSizeOfRawData = 0 Then n = n - 1
		Else
			n = n - 1
		End If
		ReDim TempList(n) As String
		For i = 0 To n
			If SecList(i).SubSecs = 0 Then
				ReDim Preserve TempList(k + 1) As String
				TempList(k) = ValToStr(SecList(i).lPointerToRawData,MaxVal,DisPlayFormat)
				k = k + 1
			Else
				ReDim Preserve TempList(k + SecList(i).SubSecs) As String
				For j = 0 To SecList(i).SubSecs - 1
					TempList(k) = ValToStr(SecList(i).SubSecList(j).lPointerToRawData,MaxVal,DisPlayFormat)
					k = k + 1
				Next j
			End If
		Next i
		If k > 0 Then k = k - 1
		ReDim Preserve TempList(k) As String
	Case 13,14	'获取所有主区段和子区段的结束地址(仅值)
		n = UBound(SecList)
		If Mode = 13 Then
			If SecList(n).lSizeOfRawData = 0 Then n = n - 1
		Else
			n = n - 1
		End If
		ReDim TempList(n) As String
		For i = 0 To n
			If SecList(i).SubSecs = 0 Then
				ReDim Preserve TempList(k + 1) As String
				TempList(k) = ValToStr(SecList(i).lPointerToRawData + _
								SecList(i).lSizeOfRawData - 1,MaxVal,DisPlayFormat)
				k = k + 1
			Else
				ReDim Preserve TempList(k + SecList(i).SubSecs) As String
				For j = 0 To SecList(i).SubSecs - 1
					TempList(k) = ValToStr(SecList(i).SubSecList(j).lPointerToRawData + _
								SecList(i).SubSecList(j).lSizeOfRawData - 1,MaxVal,DisPlayFormat)
					k = k + 1
				Next j
			End If
		Next i
		If k > 0 Then k = k - 1
		ReDim Preserve TempList(k) As String
	End Select
	getSectionList = TempList
End Function


'从小到大排序鸡尾酒排序节表
'Mode = False 按偏移地址排序，否则按相对虚拟地址排序
Private Sub SortSectionByAddress(SecList() As SECTION_PROPERTIE,ByVal Mode As Boolean)
	Dim i As Long,r As Long,l As Long,TmpX As Long,TmpA As SECTION_PROPERTIE
	l = LBound(SecList): r = UBound(SecList)
	If Mode = False Then
		Do While r > l
			For i = l To r - 1
				If SecList(i).lPointerToRawData > SecList(i + 1).lPointerToRawData Then
					TmpA = SecList(i)
					SecList(i) = SecList(i + 1)
					SecList(i + 1) = TmpA
					TmpX = i
				End If
			Next i
			r = TmpX
			For i = r To l + 1 Step -1
				If SecList(i).lPointerToRawData < SecList(i - 1).lPointerToRawData Then
					TmpA = SecList(i)
					SecList(i) = SecList(i - 1)
					SecList(i - 1) = TmpA
					TmpX = i
				End If
			Next i
			l = TmpX
		Loop
	Else
		Do While r > l
			For i = l To r - 1
				If SecList(i).lVirtualAddress > SecList(i + 1).lVirtualAddress Then
					TmpA = SecList(i)
					SecList(i) = SecList(i + 1)
					SecList(i + 1) = TmpA
					TmpX = i
				End If
			Next i
			r = TmpX
			For i = r To l + 1 Step -1
				If SecList(i).lVirtualAddress < SecList(i - 1).lVirtualAddress Then
					TmpA = SecList(i)
					SecList(i) = SecList(i - 1)
					SecList(i - 1) = TmpA
					TmpX = i
				End If
			Next i
			l = TmpX
		Loop
	End If
End Sub


'按空余字节的长度快速排序空余地址数据数组
'Mode = False 从小到大排序，否则从大到小排序，l = 数组的左边界，r = 数组的右边界
Public Sub SortFreeByteByLength(MyArray() As FREE_BTYE_SPACE, ByVal l As Long,ByVal r As Long,ByVal Mode As Boolean)
	Dim i As Long, j As Long, TmpX As Long, TmpA As FREE_BTYE_SPACE
	i = l: j = r: TmpX = MyArray((l + r) \ 2).Length
	While (i <= j)
		If Mode = False Then
			While (MyArray(i).Length < TmpX And i < r)
				i = i + 1
			Wend
			While (TmpX < MyArray(j).Length And j > l)
				j = j - 1
			Wend
		Else
			While (MyArray(i).Length > TmpX And i < r)
				i = i + 1
			Wend
			While (TmpX > MyArray(j).Length And j > l)
				j = j - 1
			Wend
		End If
		If (i <= j) Then
			TmpA = MyArray(i)
			MyArray(i) = MyArray(j)
			MyArray(j) = TmpA
			i = i + 1: j = j - 1
		End If
	Wend
	If (l < j) Then Call SortFreeByteByLength(MyArray, l, j, Mode)
	If (i < r) Then Call SortFreeByteByLength(MyArray, i, r, Mode)
End Sub


'按空余字节的地址鸡尾酒排序空余地址数据数组
'Mode = False 从小到大排序，否则从大到小排序，l = 数组的左边界，r = 数组的右边界
Public Sub SortFreeByteByAddress(MyArray() As FREE_BTYE_SPACE, ByVal l As Long,ByVal r As Long,ByVal Mode As Boolean)
	Dim i As Long, TmpX As Long, TmpA As FREE_BTYE_SPACE
	If Mode = False Then
		Do While r > l
			For i = l To r - 1
				If MyArray(i).Address > MyArray(i + 1).Address Then
					TmpA = MyArray(i)
					MyArray(i) = MyArray(i + 1)
					MyArray(i + 1) = TmpA
					TmpX = i
				End If
			Next i
			r = TmpX
			For i = r To l + 1 Step -1
				If MyArray(i).Address < MyArray(i - 1).Address Then
					TmpA = MyArray(i)
					MyArray(i) = MyArray(i - 1)
					MyArray(i - 1) = TmpA
					TmpX = i
				End If
			Next i
			l = TmpX
		Loop
	Else
		Do While r > l
			For i = l To r - 1
				If MyArray(i).Address < MyArray(i + 1).Address Then
					TmpA = MyArray(i)
					MyArray(i) = MyArray(i + 1)
					MyArray(i + 1) = TmpA
					TmpX = i
				End If
			Next i
			r = TmpX
			For i = r To l + 1 Step -1
				If MyArray(i).Address > MyArray(i - 1).Address Then
					TmpA = MyArray(i)
					MyArray(i) = MyArray(i - 1)
					MyArray(i - 1) = TmpA
					TmpX = i
				End If
			Next i
			l = TmpX
		Loop
	End If
End Sub


'查看本地化文件信息
'Mode = 0 可查看新增文件节信息
'Mode = 1 不可查看新增文件节信息
'Mode > 1 查看增加文件节大小信息
Public Sub FileInfoView(File As FILE_PROPERTIE,ByteList() As FREE_BTYE_SPACE,ByVal AppID As Long,ByVal Mode As Long,ByVal DisPlayFormat As Boolean)
	Dim i As Long,j As Long,MsgList() As String,FN As Variant
	If getMsgList(UIDataList,MsgList,"FileInfoView",2) = False Then Exit Sub
	If InStr(File.Magic,"MAC") Then
		If Mode > 1 Then
			MsgList(18) = MsgList(63)
		Else
			MsgList(18) = MsgList(61)
		End If
		MsgList(19) = MsgList(62)
	ElseIf Mode > 1 Then
		MsgList(18) = MsgList(23)
	End If
	'打开文件
	On Error GoTo ErrHandle
	If Dir$(File.FilePath & ".xls") <> "" Then Kill File.FilePath & ".xls"
	FN = FreeFile
	Open File.FilePath & ".xls" For Binary Access Read Write Lock Write As #FN
	'MAC64的情况下，无法计算 64 位(8 个字节)的数值，只能用16进制显示
	If File.Magic = "MAC64" Then DisPlayFormat = True
	'写入文件属性信息
	ReDim TempList(18) As String
	With File
		TempList(0) = MsgList(0)
		TempList(1) = Replace$(MsgList(1),"%s",.FileName)
		TempList(2) = Replace$(MsgList(2),"%s",.FilePath)
		TempList(3) = Replace$(MsgList(3),"%s",.FileDescription)
		TempList(4) = Replace$(MsgList(4),"%s",.FileVersion)
		TempList(5) = Replace$(MsgList(5),"%s",.ProductName)
		TempList(6) = Replace$(MsgList(6),"%s",.ProductVersion)
		TempList(7) = Replace$(MsgList(7),"%s",.LegalCopyright)
		TempList(8) = Replace$(MsgList(8),"%s",CStr$(.FileSize))
		TempList(9) = Replace$(MsgList(9),"%s",CStr$(.DateCreated))
		TempList(10) = Replace$(MsgList(10),"%s",CStr$(.DateLastModified))
		TempList(11) = Replace$(MsgList(11),"%s",PSL.GetLangCode(Val("&H" & .LanguageID),pslCodeText))
		TempList(12) = Replace$(MsgList(12),"%s",.CompanyName)
		TempList(13) = Replace$(MsgList(13),"%s",.OrigionalFileName)
		TempList(14) = Replace$(MsgList(14),"%s",.InternalName)
		Select Case .Magic
		Case "PE32","NET32","MAC32"
			If .LangType = DELPHI_FILE_SIGNATURE Then
				TempList(15) = Replace$(MsgList(15),"%s","Delphi32")
				TempList(16) = Replace$(MsgList(16),"%s","0x" & ValToStr(.ImageBase,-8,True)) & vbCrLf
			ElseIf .LangType = NET_FILE_SIGNATURE Then
				TempList(15) = Replace$(MsgList(15),"%s",".NET32")
				TempList(16) = Replace$(MsgList(16),"%s","0x" & ValToStr(.ImageBase,-8,True)) & vbCrLf
			ElseIf InStr(.Magic,"MAC") Then
				TempList(15) = Replace$(MsgList(15),"%s","MAC32")
				TempList(16) = vbCrLf
			Else
				TempList(15) = Replace$(MsgList(15),"%s","PE32")
				TempList(16) = Replace$(MsgList(16),"%s","0x" & ValToStr(.ImageBase,-8,True)) & vbCrLf
			End If
		Case "PE64","NET64","MAC64"
			If .LangType = DELPHI_FILE_SIGNATURE Then
				TempList(15) = Replace$(MsgList(15),"%s","Delphi64")
				TempList(16) = Replace$(MsgList(16),"%s","0x" & ReverseHexCode(Byte2Hex(.ImageBase,0,-1),16)) & vbCrLf
			ElseIf .LangType = NET_FILE_SIGNATURE Then
				TempList(15) = Replace$(MsgList(15),"%s",".NET64")
				TempList(16) = Replace$(MsgList(16),"%s","0x" & ReverseHexCode(Byte2Hex(.ImageBase,0,-1),16)) & vbCrLf
			ElseIf InStr(.Magic,"MAC") Then
				TempList(15) = Replace$(MsgList(15),"%s","MAC64")
				TempList(16) = vbCrLf
			Else
				TempList(15) = Replace$(MsgList(15),"%s","PE64")
				TempList(16) = Replace$(MsgList(16),"%s","0x" & ReverseHexCode(Byte2Hex(.ImageBase,0,-1),16)) & vbCrLf
			End If
		Case Else
			TempList(15) = Replace$(MsgList(15),"%s",MsgList(59))
			TempList(16) = vbCrLf
		End Select
	End With
	TempList(17) = MsgList(17) & MsgList(20) & MsgList(20) & vbCrLf
	TempList(18) = MsgList(18) & MsgList(20) & MsgList(20) & vbCrLf
	WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,"")
	'写入每个文件节的偏移地址
	MsgList(21) = Replace$(MsgList(19),"%s!1!",MsgList(21))
	For i = 0 To File.MaxSecIndex - 1
		With File.SecList(i)
			If Mode = 0 Or (.sName <> NewPESecName And .sName <> NewMacSecName) Then
				ReDim TempList(0) As String
				TempList(0) = Replace$(MsgList(21),"%s!2!",IIf(File.Magic = "",MsgList(59),.sName))
				TempList(0) = Replace$(TempList(0),"%s!3!","")
				TempList(0) = Replace$(TempList(0),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
				TempList(0) = Replace$(TempList(0),"%s!5!",ValToStr(.lPointerToRawData + _
							IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
				TempList(0) = Replace$(TempList(0),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
				If Mode < 2 Then
					TempList(0) = Replace$(TempList(0),"%s!7!","")
				ElseIf i = File.MaxSecID Then
					TempList(0) = Replace$(TempList(0),"%s!7!",MsgList(24))
				Else
					TempList(0) = Replace$(TempList(0),"%s!7!",ValToStr(ByteList(i).Length,File.FileSize,DisPlayFormat))
				End If
				If .SubSecs > 0 Then
					ReDim Preserve TempList(.SubSecs) As String
					For j = 0 To .SubSecs - 1
						TempList(1 + j) = Replace$(MsgList(21),"%s!2!","")
						TempList(1 + j) = Replace$(TempList(1 + j),"%s!3!",.SubSecList(j).sName)
						TempList(1 + j) = Replace$(TempList(1 + j ),"%s!4!",ValToStr(.SubSecList(j).lPointerToRawData,File.FileSize,DisPlayFormat))
						TempList(1 + j) = Replace$(TempList(1 + j),"%s!5!",ValToStr(.SubSecList(j).lPointerToRawData + _
										IIf(.SubSecList(j).lSizeOfRawData = 0,0,.SubSecList(j).lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
						TempList(1 + j) = Replace$(TempList(1 + j),"%s!6!",ValToStr(.SubSecList(j).lSizeOfRawData,File.FileSize,DisPlayFormat))
						TempList(1 + j) = Replace$(TempList(1 + j),"%s!7!","")
					Next j
				End If
				WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,""),True
			End If
		End With
	Next i
	'写入隐藏节的偏移地址、子 PE 地址及数量
	With File.SecList(File.MaxSecIndex)
		If .lSizeOfRawData > 0 Then
			ReDim TempList(0) As String
			TempList(0) = Replace$(MsgList(21),"%s!2!",MsgList(25))
			TempList(0) = Replace$(TempList(0),"%s!3!","")
			TempList(0) = Replace$(TempList(0),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
			TempList(0) = Replace$(TempList(0),"%s!5!",ValToStr(.lPointerToRawData + _
						IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
			TempList(0) = Replace$(TempList(0),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
			If Mode < 2 Then
				TempList(0) = Replace$(TempList(0),"%s!7!","")
			Else
				TempList(0) = Replace$(TempList(0),"%s!7!",MsgList(27))
			End If
			WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,""),True
		End If
		If File.NumberOfSub > 0 Then
			ReDim TempList(0) As String
			TempList(0) = Replace$(MsgList(21),"%s!2!",Replace$(MsgList(60),"%s",CStr$(File.NumberOfSub)))
			TempList(0) = Replace$(TempList(0),"%s!3!","")
			TempList(0) = Replace$(TempList(0),"%s!4!",ValToStr(.lPointerToRawData + .lSizeOfRawData,File.FileSize,DisPlayFormat))
			TempList(0) = Replace$(TempList(0),"%s!5!",ValToStr(File.FileSize - 1,File.FileSize,DisPlayFormat))
			TempList(0) = Replace$(TempList(0),"%s!6!",ValToStr(File.FileSize - .lPointerToRawData - .lSizeOfRawData,File.FileSize,DisPlayFormat))
			If Mode < 2 Then
				TempList(0) = Replace$(TempList(0),"%s!7!","")
			Else
				TempList(0) = Replace$(TempList(0),"%s!7!",MsgList(27))
			End If
			WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,""),True
		End If
	End With
	WriteBinaryFile FN,CP_UNICODELITTLE,vbCrLf,True
	'写入每个文件节的相对虚拟地址
	MsgList(22) = Replace$(MsgList(19),"%s!1!",MsgList(22))
	For i = 0 To File.MaxSecIndex - 1
		With File.SecList(i)
			If Mode = 0 Or (.sName <> NewPESecName And .sName <> NewMacSecName) Then
				ReDim TempList(0) As String
				TempList(0) = Replace$(MsgList(22),"%s!2!",IIf(File.Magic = "",MsgList(59),.sName))
				TempList(0) = Replace$(TempList(0),"%s!3!","")
				If File.Magic <> "MAC64" Then
					TempList(0) = Replace$(TempList(0),"%s!4!",ValToStr(.lVirtualAddress,File.FileSize,DisPlayFormat))
					TempList(0) = Replace$(TempList(0),"%s!5!",ValToStr(.lVirtualAddress + _
								IIf(.lVirtualSize = 0,0,.lVirtualSize - 1),File.FileSize,DisPlayFormat))
				Else
					TempList(0) = Replace$(TempList(0),"%s!4!",ValToStr(.lVirtualAddress1,0,DisPlayFormat) & _
								ValToStr(.lVirtualAddress,-8,DisPlayFormat))
					TempList(0) = Replace$(TempList(0),"%s!5!",ValToStr(.lVirtualAddress1,0,DisPlayFormat) & _
								ValToStr(.lVirtualAddress + IIf(.lVirtualSize = 0,0,.lVirtualSize - 1),-8,DisPlayFormat))
				End If
				TempList(0) = Replace$(TempList(0),"%s!6!",ValToStr(.lVirtualSize,File.FileSize,DisPlayFormat))
				If Mode < 2 Then
					TempList(0) = Replace$(TempList(0),"%s!7!","")
				ElseIf i = File.MaxSecID Then
					TempList(0) = Replace$(TempList(0),"%s!7!",MsgList(24))
				Else
					TempList(0) = Replace$(TempList(0),"%s!7!",ValToStr(ByteList(i).Address,File.FileSize,DisPlayFormat))
				End If
				If .SubSecs > 0 Then
					ReDim Preserve TempList(.SubSecs) As String
					For j = 0 To .SubSecs - 1
						TempList(1 + j) = Replace$(MsgList(22),"%s!2!","")
						TempList(1 + j) = Replace$(TempList(1 + j),"%s!3!",.SubSecList(j).sName)
						If File.Magic <> "MAC64" Then
							TempList(1 + j) = Replace$(TempList(1 + j),"%s!4!",ValToStr(.SubSecList(j).lVirtualAddress,File.FileSize,DisPlayFormat))
							TempList(1 + j) = Replace$(TempList(1 + j),"%s!5!",ValToStr(.SubSecList(j).lVirtualAddress + _
										IIf(.SubSecList(j).lVirtualSize = 0,0,.lVirtualSize - 1),File.FileSize,DisPlayFormat))
						Else
							TempList(1 + j) = Replace$(TempList(1 + j),"%s!4!",ValToStr(.SubSecList(j).lVirtualAddress1,0,DisPlayFormat) & _
										ValToStr(.SubSecList(j).lVirtualAddress,-8,DisPlayFormat))
							TempList(1 + j) = Replace$(TempList(1 + j),"%s!5!",ValToStr(.SubSecList(j).lVirtualAddress1,0,DisPlayFormat) & _
										ValToStr(.SubSecList(j).lVirtualAddress + IIf(.SubSecList(j).lVirtualSize = 0,0,.lVirtualSize - 1),-8,DisPlayFormat))
						End If
						TempList(1 + j) = Replace$(TempList(1 + j),"%s!6!",ValToStr(.SubSecList(j).lVirtualSize,File.FileSize,DisPlayFormat))
						TempList(1 + j) = Replace$(TempList(1 + j),"%s!7!","")
					Next j
				End If
				WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,""),True
			End If
		End With
	Next i
	'写入隐藏节的相对虚拟地址、子 PE 地址及数量
	ReDim TempList(0) As String
	With File.SecList(File.MaxSecIndex)
		If .lVirtualSize > 0 Then
			TempList(0) = Replace$(MsgList(22),"%s!2!",MsgList(25))
			TempList(0) = Replace$(TempList(0),"%s!3!","")
			TempList(0) = Replace$(TempList(0),"%s!4!",MsgList(26))
			TempList(0) = Replace$(TempList(0),"%s!5!",MsgList(26))
			TempList(0) = Replace$(TempList(0),"%s!6!",MsgList(26))
			If Mode < 2 Then
				TempList(0) = Replace$(TempList(0),"%s!7!","")
			Else
				TempList(0) = Replace$(TempList(0),"%s!7!",MsgList(27))
			End If
			WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,""),True
		End If
		If File.NumberOfSub > 0 Then
			TempList(0) = Replace$(MsgList(22),"%s!2!",Replace$(MsgList(60),"%s",CStr$(File.NumberOfSub)))
			TempList(0) = Replace$(TempList(0),"%s!3!","")
			TempList(0) = Replace$(TempList(0),"%s!4!",MsgList(26))
			TempList(0) = Replace$(TempList(0),"%s!5!",MsgList(26))
			TempList(0) = Replace$(TempList(0),"%s!6!",MsgList(26))
			If Mode < 2 Then
				TempList(0) = Replace$(TempList(0),"%s!7!","")
			Else
				TempList(0) = Replace$(TempList(0),"%s!7!",MsgList(27))
			End If
			WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,""),True
		End If
	End With
	WriteBinaryFile FN,CP_UNICODELITTLE,MsgList(20) & MsgList(20) & vbCrLf,True
	'写入数据目录地址及所在文件节
	If File.DataDirs > 0 Then
		ReDim TempList(1) As String
		TempList(0) = vbCrLf & MsgList(28) & MsgList(20) & MsgList(20)
		TempList(1) = vbCrLf & MsgList(29) & MsgList(20) & MsgList(20) & vbCrLf
		WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,""),True
		ReDim TempList(File.DataDirs) As String
		For i = 0 To File.DataDirs - 1
			With File.DataDirectory(i)
				TempList(i) = Replace$(MsgList(19),"%s!1!",MsgList(i + 30))
				If .lPointerToRawData > 0 Then
					j = SkipSection(File,.lPointerToRawData,0,0,1)
					If j > -1 Then
						TempList(i) = Replace$(TempList(i),"%s!2!",File.SecList(j).sName)
					Else
						TempList(i) = Replace$(TempList(i),"%s!2!",MsgList(46))
					End If
				ElseIf .lSizeOfRawData > 0 Then
					TempList(i) = Replace$(TempList(i),"%s!2!",MsgList(46))
				Else
					TempList(i) = Replace$(TempList(i),"%s!2!",MsgList(47))
				End If
				TempList(i) = Replace$(TempList(i),"%s!3!","")
				TempList(i) = Replace$(TempList(i),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
				TempList(i) = Replace$(TempList(i),"%s!5!",ValToStr(.lPointerToRawData + IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
				TempList(i) = Replace$(TempList(i),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
				TempList(i) = Replace$(TempList(i),"%s!7!","")
			End With
		Next i
		WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,"") & MsgList(20) & MsgList(20) & vbCrLf,True
	End If
	'写入 .NET CLR 数据目录地址及所在文件节
	If File.LangType = NET_FILE_SIGNATURE Then
		ReDim TempList(1) As String
		TempList(0) = vbCrLf & MsgList(48) & MsgList(20) & MsgList(20)
		TempList(1) = vbCrLf & MsgList(49) & MsgList(20) & MsgList(20) & vbCrLf
		WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,""),True
		ReDim TempList(6) As String
		For i = 0 To 6
			With File.CLRList(i)
				TempList(i) = Replace$(MsgList(19),"%s!1!",MsgList(i + 50))
				If .lPointerToRawData > 0 Then
					j = SkipSection(File,.lPointerToRawData,0,0,1)
					If j > -1 Then
						TempList(i) = Replace$(TempList(i),"%s!2!",File.SecList(j).sName)
					Else
						TempList(i) = Replace$(TempList(i),"%s!2!",MsgList(46))
					End If
				ElseIf .lSizeOfRawData > 0 Then
					TempList(i) = Replace$(TempList(i),"%s!2!",MsgList(46))
				Else
					TempList(i) = Replace$(TempList(i),"%s!2!",MsgList(47))
				End If
				TempList(i) = Replace$(TempList(i),"%s!3!","")
				TempList(i) = Replace$(TempList(i),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
				TempList(i) = Replace$(TempList(i),"%s!5!",ValToStr(.lPointerToRawData + IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
				TempList(i) = Replace$(TempList(i),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
				TempList(i) = Replace$(TempList(i),"%s!7!","")
			End With
		Next i
		WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,"") & MsgList(20) & MsgList(20) & vbCrLf,True
	End If
	'写入 .NET 流地址及所在文件节
	If File.NetStreams > 0 Then
		ReDim TempList(1) As String
		TempList(0) = vbCrLf & MsgList(57) & MsgList(20) & MsgList(20)
		TempList(1) = vbCrLf & MsgList(58) & MsgList(20) & MsgList(20) & vbCrLf
		WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,"")
		ReDim TempList(File.NetStreams - 1) As String
		For i = 0 To File.NetStreams - 1
			With File.StreamList(i)
				TempList(i) = Replace$(MsgList(19),"%s!1!",.sName)
				If .lPointerToRawData > 0 Then
					j = SkipSection(File,.lPointerToRawData,0,0,1)
					If j > -1 Then
						TempList(i) = Replace$(TempList(i),"%s!2!",File.SecList(j).sName)
					Else
						TempList(i) = Replace$(TempList(i),"%s!2!",MsgList(46))
					End If
				ElseIf .lSizeOfRawData > 0 Then
					TempList(i) = Replace$(TempList(i),"%s!2!",MsgList(46))
				Else
					TempList(i) = Replace$(TempList(i),"%s!2!",MsgList(47))
				End If
				TempList(i) = Replace$(TempList(i),"%s!3!","")
				TempList(i) = Replace$(TempList(i),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
				TempList(i) = Replace$(TempList(i),"%s!5!",ValToStr(.lPointerToRawData + IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
				TempList(i) = Replace$(TempList(i),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
				TempList(i) = Replace$(TempList(i),"%s!7!","")
			End With
		Next i
		WriteBinaryFile FN,CP_UNICODELITTLE,StrListJoin(TempList,"") & MsgList(20) & MsgList(20) & vbCrLf,True
	End If
	Close #FN
	Erase MsgList,TempList
	'查看数据
	ReDim FileDataList(0) As String
	FileDataList(0) = File.FilePath & ".xls" & JoinStr & "unicodeFFFE"
	If OpenFile(File.FilePath & ".xls",FileDataList,AppID,False) = True Then
		If AppID = 3 Then WriteSettings("Tools")
	End If
	If AppID = 0 Then
		 On Error Resume Next
		 If Dir$(File.FilePath & ".xls") <> "" Then Kill File.FilePath & ".xls"
		 On Error GoTo 0
	End If
	Exit Sub
	'错误处理
	ErrHandle:
	On Error Resume Next
	Close #FN
	Err.Source = "NotWriteFile"
	Err.Description = Err.Description & JoinStr & File.FilePath & ".xls"
	Call sysErrorMassage(Err,1)
End Sub


'导出注册表关键字的值
'参数说明: KeyRoot - 根类型, KeyName - 子项名称, FileName - 导出的文件路径及文件名(原始数据库格式)
Public Function SaveKey(KeyRoot As REG_KEYROOT, KeyName As String, FileName As String) As Boolean
	Dim hKey As Long
	'Dim lpAttr As SECURITY_ATTRIBUTES	'注册表安全类型
	'lpAttr.nLength = 50				'设置安全属性为缺省值
	'lpAttr.lpSecurityDescriptor = 0
	'lpAttr.bInheritHandle = True
	If EnablePrivilege(SE_BACKUP_NAME) = False Then Exit Function
	If RegOpenKeyEx(KeyRoot, KeyName, 0&, KEY_ALL_ACCESS, hKey) <> 0 Then
		RegCloseKey(hKey)
		Exit Function
	End If
	'if RegSaveKey(hKey, FileName, lpAttr) = 0 Then SaveKey = True
	If RegSaveKey(hKey, FileName, 0&) = 0 Then SaveKey = True
	RegCloseKey(hKey)
End Function


'导入注册表关键字的值
'参数说明: KeyRoot - 根类型, KeyName - 子项名称, FileName - 导入的文件路径及文件名(原始数据库格式)
Public Function RestoreKey(KeyRoot As REG_KEYROOT, ByVal KeyName As String, ByVal FileName As String) As Boolean
	Dim hKey As Long
	If EnablePrivilege(SE_RESTORE_NAME) = False Then Exit Function
	If RegOpenKeyEx(KeyRoot, KeyName, 0&, KEY_ALL_ACCESS, hKey) <> 0 Then
		RegCloseKey(hKey)
		Exit Function
	End If
	If RegRestoreKey(hKey, FileName, REG_FORCE_RESTORE) = 0 Then RestoreKey = True
	RegCloseKey(hKey)
End Function


'使注册表允许导入导出
Private Function EnablePrivilege(ByVal seName As String) As Boolean
	Dim p_lngToken As Long
    Dim p_typLUID As REG_LUID,p_typTokenPriv As REG_TOKEN_PRIVILEGES,p_typPrevTokenPriv As REG_TOKEN_PRIVILEGES
	If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken) = 0 Then Exit Function
	If Err.LastDLLError <> 0 Then Exit Function
	If LookupPrivilegeValue(vbNullChar, seName, p_typLUID) = 0 Then Exit Function
	p_typTokenPriv.PrivilegeCount = 1
	p_typTokenPriv.Privileges.Attributes = SE_PRIVILEGE_ENABLED
	p_typTokenPriv.Privileges.pLuid = p_typLUID
	If AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, 0&) = 0 Then Exit Function
	EnablePrivilege = True
End Function


'自定义弹出菜单(最多2级)，返回菜单项的文本
'strConfig 为自定义的菜单数据,本程序以主菜单和子菜单之间以 vbNullChar 分隔
'Mode = 返回模式，False = 返回字串，True = 返回索引项
Public Function PopupMenuShow(strConfig() As String,Optional ByVal Mode As Boolean) As Variant
	Dim i As Long,j As Long,n As Long,hMenu As Long,MenuCount As Long
	Dim pt As POINTAPI,subMenuList() As String,MenuText() As String
	'获取当前光标位置
	GetCursorPos pt
	'创建一个空的菜单,获取句柄
	hMenu = CreatePopupMenu()
	'设置子菜单的句柄数量
	ReDim hPopMenu(UBound(strConfig)) As Long
	For i = 0 To UBound(strConfig)
		If InStr(strConfig(i),vbNullChar) Then
			subMenuList = Split(strConfig(i),vbNullChar)
			If subMenuList(0) <> "" Then
				'创建子级菜单项
				hPopMenu(n) = CreatePopupMenu()
				For j = 1 To UBound(subMenuList)
					If subMenuList(j) <> "" Then
						MenuCount = MenuCount + 1
						'保存菜单文本，用于菜单事件触发时识别出被选择的菜单对象
						ReDim Preserve MenuText(MenuCount) As String
						MenuText(MenuCount) = subMenuList(j)
						'添加子菜单项
						'如果是间隔线，则 wFlags = MF_SEPARATOR
						'如果要Check，则 wFlags = MF_STRING + MF_CHECKED，若令不可用，则再加 MF_GRAYED
						If LCase(subMenuList(j)) = "step" Then
							'添加间隔线,step 是间隔线的标示,可以人为定义
							AppendMenu hPopMenu(n), MF_SEPARATOR, MenuCount, MenuText(MenuCount)
						Else
							AppendMenu hPopMenu(n), MF_STRING + MF_ENABLED, MenuCount, MenuText(MenuCount)
						End If
					End If
				Next j
				AppendMenu hMenu, MF_POPUP, hPopMenu(n), subMenuList(0)	'添加父级菜单项
				n = n + 1
			End If
		ElseIf strConfig(i) <> "" Then
			MenuCount = MenuCount + 1
			If LCase(strConfig(i)) = "step" Then
				'添加间隔线,step 是间隔线的标示,可以人为定义
				AppendMenu hMenu, MF_SEPARATOR, MenuCount, strConfig(i)
			Else
				AppendMenu hMenu, MF_STRING + MF_ENABLED, MenuCount, strConfig(i)
			End If
			n = n + 1
		End If
	Next i
	If n = 0 Then Exit Function
	'显示菜单,返回 0 表示放弃
	'如果在参数 uFlags 里指定了 TPM_RETURNCMD 值，则返回值是用户选择的菜单项的标识符。
	i = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, pt.X, pt.Y, 0&, GetForegroundWindow(), 0)
	'返回项目为 0 时设置为 -1
	If i < 1 Then
    	'释放菜单资源
	    DestroyMenu hMenu
	    Exit Function
	End If
	If Mode Then
		PopupMenuShow = i - 1 '返回菜单项
		'释放菜单资源
		DestroyMenu hMenu
	Else
		'获取选中的菜单项字串
		Dim buffer As String
		buffer = Space(255)
		'MF_BYCOMMAND: 表示参数 uIDltem 给出菜单项的标识符, 缺省值
		'MF_BYPOSITION: 表示参数 uIDltem给出菜单项相对于零的位置
		i = GetMenuString(hMenu, i, buffer, Len(buffer), MF_BYCOMMAND)
		'释放菜单资源
		DestroyMenu hMenu
		If i = 0 Then Exit Function
		'返回所选菜单项的文本
		PopupMenuShow = Replace$(buffer,vbNullChar,"")
	End If
End Function


'检查自定义字串类型的数组是否为空，非空返回 True
Public Function CheckRefTypeArray(DataList() As REF_TYPE,Optional ByVal SetID As Long = -1) As Boolean
	Dim i As Long,Max As Long
	CheckRefTypeArray = True
	On Error GoTo ExitFunction
	If SetID = -1 Then
		SetID = 0: Max = UBound(DataList)
	Else
		Max = SetID
	End If
	For i = SetID To Max
		If DataList(i).Algorithm <> "" Then Exit Function
		If DataList(i).ByteLength > 0 Then Exit Function
	Next i
	ExitFunction:
	CheckRefTypeArray = False
End Function
