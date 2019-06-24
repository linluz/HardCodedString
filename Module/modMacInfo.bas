Attribute VB_Name = "modMacInfo"

'' File Information Module for PSlHardCodedString.bas
'' (c) 2015-2019 by wanfu (Last modified on 2019.03.15)

'#Uses "modCommon.bas"

Option Explicit

Private Const MAC_BIT_TYPE32 = 48 + 68
Private Const MAC_BIT_TYPE64 = 72 + 80

Private Type mac_header_32  '28个字节
	lmagic				As Long		'mach magic number identifier
	lcputype			As Long		'cpu specifier (int)
	lcpusubtype			As Long		'machine specifier (int)
	lfiletype			As Long		'type of file
	lncmds				As Long		'指定有多少个Command
	lsizeofcmds			As Long		'指定LoadCommand总的大小
	lflags				As Long		'file offset to this object file
End Type

Private Type mac_header_64  '32个字节
	lmagic				As Long		'mach magic number identifier
	lcputype			As Long		'cpu specifier (int)
	lcpusubtype			As Long		'machine specifier (int)
	lfiletype			As Long		'type of file
	lncmds				As Long		'指定有多少个Command
	lsizeofcmds			As Long		'指定LoadCommand总的大小
	lflags				As Long		'file offset to this object file
	lreserved			As Long
End Type
'magic: 可以看到文件中的内容最开始部分，是以 cafe babe开头的
'       对于一个 二进制文件 来讲，每个类型都可以在文件最初几个字节来标识出来，即“魔数”。不同类型的 二进制文件，都有自己独特的"魔数"。
'       OS X上，可执行文件的标识有这样几个魔数（不同的魔数代表不同的可执行文件类型）
'       是mach-o文件的魔数，0xfeedface代表的是32位，0xfeedfacf代表64位，cafebabe是跨处理器架构的通用格式，#!代表的是脚本文件。
'cputype 和 cupsubtype: 代表的是cpu的类型和其子类型，图上的例子是模拟器程序，cpu结构是x86_64,如果直接查看ipa，可以看到cpu是arm，subtype是armv7，arm64等
'#define CPU_TYPE_ARM((cpu_type_t) 12)
'#define CPU_SUBTYPE_ARM_V7((cpu_subtype_t) 9
'filetype: &H2 代表可执行的文件
'ncmds: 指的是加载命令(load commands)的数量，例子中一共65个，编号0-64
'sizeofcmds: 表示23个load commands的总字节大小，load commands区域是紧接着header区域的。
'flags: 例子中是0×00200085，可以按文档分析之。

'对象链接库文件
Private Type mac_header_fat_arch	'20个字节
	lcputype			As Long		'CPU specifier
	lcpusubtype			As Long		'Machine specifier
	lfileoffset			As Long		'Offset of header in file
	lsize				As Long		'size of object file
	lalign				As Long		'Alignment As a power of two
End Type

'对象链接库文件
Private Type mac_header_fat  '32个字节
	lmagic				As Long		'mach magic number identifier
	lfat_arch_size		As Long		'Number of fat_arch structs
	fat_archs() 		As mac_header_fat_arch
End Type

'Constants for the magic field of the mach_header
Private Enum mac_header_magic
	MH_MAGIC_32			= &HFEEDFACE		' 32-bit mach Object file
	MH_MAGIC_64			= &HFEEDFACF		' 64-bit mach Object file
	MH_MAGIC_FAT		= &HCAFEBABE		' Universal Object file / FAT_MAGIC
	MH_MAGIC_FAT_CIGAM	= &HBEBAFECA
End Enum

'Private Enum mac_header_cputype
'	CPU_ARCH_MASK		= &Hff000000
'	CPU_ARCH_ABI64 		= &H01000000
	' This looks ugly due To a limitation (bug?) In 010Editor template processing,
	' basically we're unable to define more constant using other constants - it doesn't
	' see them As already being processed when trying To define others (though it won't
	' Error On this Until it hits this when trying To Access that constant)
'	CPU_TYPE_X86		= &H7
'	CPU_TYPE_I386		= &H7					' CPU_TYPE_X86
'	CPU_TYPE_X86_64		= (&H7 Or &H01000000)	' (CPU_TYPE_X86 | CPU_ARCH_ABI64)
'	CPU_TYPE_POWERPC	= &H12
'	CPU_TYPE_POWERPC64	= (&H12 Or &H01000000)	' (CPU_TYPE_POWERPC | CPU_ARCH_ABI64)
'	CPU_TYPE_ARM		= &HC
'End Enum

'Constants for the filetype field of the mach_header
Private Enum mac_header_filetype
	MH_OBJECT			= &H1			'relocatable Object file
	MH_EXECUTE			= &H2			'demand paged executable file
	MH_FVMLIB			= &H3			'fixed VM Shared library file
	MH_CORE				= &H4			'core file
	MH_PRELOAD			= &H5			'preloaded executable file
	MH_DYLIB			= &H6			'dynamically bound Shared library
	MH_DYLINKER			= &H7			'dynamic link editor
	MH_BUNDLE			= &H8			'dynamically bound bundle file
	MH_DYLIB_STUB		= &H9			'Shared library stub For Static
	MH_DSYM				= &HA			'companion file With only Debug
	MH_KEXT_BUNDLE		= &HB			'x86_64 kexts
End Enum

'Constants for the mac_header_flags field
Private Enum mac_header_flags
	MH_NOUNDEFS					= &H1			'目前没有未定义的符号，不存在链接依赖
	MH_INCRLINK					= &H2
	MH_DYLDLINK					= &H4			'该文件是dyld的输入文件，无法被再次静态链接
	MH_BINDATLOAD				= &H8
	MH_PREBOUND					= &H10
	MH_SPLIT_SEGS				= &H20
	MH_LAZY_INIT				= &H40
	MH_TWOLEVEL					= &H80			'该镜像文件使用2级名称空间
	MH_FORCE_FLAT				= &H100
	MH_NOMULTIDEFS				= &H200
	MH_NOFIXPREBINDING			= &H400
	MH_PREBINDABLE				= &H800
	MH_ALLMODSBOUND				= &H1000
	MH_SUBSECTIONS_VIA_SYMBOLS	= &H2000
	MH_CANONICAL				= &H4000
	MH_WEAK_DEFINES				= &H8000
	MH_BINDS_TO_WEAK			= &H10000		'最后链接的镜像文件使用弱符号
	MH_ALLOW_STACK_EXECUTION	= &H20000
	MH_ROOT_SAFE				= &H40000
	MH_SETUID_SAFE				= &H80000
	MH_NO_REEXPORTED_DYLIBS		= &H100000
	MH_PIE						= &H200000		'加载程序在随机的地址空间，只在 MH_EXECUTE中使用
	MH_DEAD_STRIPPABLE_DYLIB	= &H400000
	MH_HAS_TLV_DESCRIPTORS		= &H800000
	MH_NO_HEAP_EXECUTION		= &H1000000
End Enum

'load command 结构
'Command有很多不同的种类，每个种类对应一个结构体但是所有的Command都会有相同的开始结构
'注意这个大小是包括了它的所有内容，包括这个结构体本身所占的大小，它后面所跟的Section结构的大小，
'和所有的Padding对齐的0.(但是不包括真正的Data,真正的Data一般在 FileOffset 中指出,根据不同Command会不同)
'所以从命令开始处加上第二个成员的大小，就可以直接定位到下一个命令的开始处。
'个人觉得这个设计相当的挫，哈哈，为撒，因为你需要先读一个Load_Command结构才能知道当前命令是个什么类型，
'然后再去读对应的结构，读完以后，还要回到命令开始处，再加上第二个成员的大小去处理下一个命令。比较挫！
'比如，如果cmd=19，它代表一个Segment_Command_64,也就是从那里开始其实是一个Segment_Command_64结构
Private Type mac_load_command	'8个字节
	lcmd				As Long		'Command 的类型
	lcmdsize			As Long		'Command 的大小
End Type
'load commmand直接跟在 header 部分的后面

'Constants for the load_command_cmd field
Private Const REQ_DYLD		= &H80000000
Private Enum load_command_cmd
	SEGMENT					= &H1
	SYM_TAB					= &H2
	SYM_SEG					= &H3
	THREAD					= &H4
	UNIX_THREAD				= &H5
	LOAD_FVM_LIB			= &H6
	ID_FVM_LIB				= &H7
	IDENT					= &H8
	FVM_FILE				= &H9
	PREPAGE					= &HA
	DY_SYM_TAB				= &HB
	LOAD_DYLIB				= &HC
	ID_DYLIB				= &HD
	LOAD_DYLINKER			= &HE
	ID_DYLINKER				= &HF
	PREBOUND_DYLIB			= &H10
	ROUTINES				= &H11
	SUB_FRAMEWORK			= &H12
	SUB_UMBRELLA			= &H13
	SUB_CLIENT				= &H14
	SUB_LIBRARY				= &H15
	TWOLEVEL_HINTS			= &H16
	PREBIND_CKSUM			= &H17
	LOAD_WEAK_DYLIB			= &H18 Or REQ_DYLD
	SEGMENT_64				= &H19
	ROUTINES_64				= &H1A
	UUID					= &H1B
	RPATH					= &H1C Or REQ_DYLD
	CODE_SIGNATURE			= &H1D
	SEGMENT_SPLIT_INFO		= &H1E
	REEXPORT_DYLIB			= &H1F Or REQ_DYLD
	LAZY_LOAD_DYLIB			= &H20
	ENCRYPTION_INFO			= &H21
	DYLD_INFO				= &H22
	DYLD_INFO_ONLY			= &H22 Or REQ_DYLD
	LOAD_UPWARD_DYLIB		= &H23 Or REQ_DYLD
	VERSION_MIN_MAC_OSX		= &H24
	VERSION_MIN_IPHONE_OS	= &H25
	FUNCTION_STARTS			= &H26
	DYLD_ENVIRONMENT		= &H27
	MAIN_CMD				= &H28
	DATA_IN_CODE			= &H29
	SOURCE_VERSION			= &H2A
	DYLIB_CODE_SIGN_DRS		= &H2B
End Enum
'一个典型的 OS X 可执行文件通常由下列五段：:
'__PAGEZERO : 定位于虚拟地址0，无任何保护权利。此段在文件中不占用空间，访问Null导致立即崩溃.
'__TEXT : 包含只读数据和可执行代码.
'__DATA : 包含可写数据. 这些 section通常由内核标志为copy-On-Write .
'__OBJC : 包含Objective C 语言运行时环境使用的数据。
'__LINKEDIT :包含动态链接器用的原始数据.
'__TEXT和 __DATA段可能包含0或更多的section. 每个section由指定类型的数据, 如, 可执行代码, 常量, C 字符串等组成

'Segment_Command 结构
'上面的结构包括了段名，和初始化的内存保护掩码，还有虚拟地址和文件偏移，和Windows上的内容差不多。
'重要的是下面，nsects和flags，这两个一个指明后面跟了多少sections，另一个代表当前的段属性。
'如果nsects>0，代表后面有节，而且节的定义紧跟的段定义。
Private Type segment_command_32	'40个字节
	segname(15) 		As Byte			'segment name  16 个字符
	lvmaddr				As Long			'memory address of this segment 段的虚拟内存地址
	lvmsize				As Long			'memory size of this segment VM Address 段的虚拟内存大小
	lfileoff			As Long			'file offset of this segment 段在文件中偏移量
	lfilesize			As Long			'amount to map from the file 段在文件中的大小
	lmaxprot			As Long			'maximum VM protection
	linitprot			As Long			'initial VM protection
	lnsects				As Long			'number of sections in segment
	lflags				As Long			'flags
End Type

Private Type segment_command_64	'64个字节
	segname(15) 		As Byte			'segment name  16 个字符
	dvmaddr1			As Long			'memory address of this segment 段的虚拟内存地址第一部分
	dvmaddr2			As Long			'memory address of this segment 段的虚拟内存地址第二部分
	dvmsize1			As Long			'memory size of this segment VM Address 段的虚拟内存大小第一部分
	dvmsize2			As Long			'memory size of this segment VM Address 段的虚拟内存大小第二部分
	dfileoff1			As Long			'file offset of this segment 段在文件中偏移量第一部分
	dfileoff2			As Long			'file offset of this segment 段在文件中偏移量第二部分
	dfilesize1			As Long			'amount to map from the file 段在文件中的大小第一部分
	dfilesize2			As Long			'amount to map from the file 段在文件中的大小第二部分
	lmaxprot			As Long			'maximum VM protection
	linitprot			As Long			'initial VM protection
	lnsects				As Long			'number of sections in segment
	lflags				As Long			'flags
End Type
'将该段对应的文件内容加载到内存中：从offset处加载 file
'size大小到虚拟内存 vmaddr 处，由于这里在内存地址空间中是_PAGEZERO段（这个段不具有访问权限，用来处理空指针）所以都是零
'还有其他段，比如_TEXT对应的就是代码段，_DATA对应的是可读／可写的数据，_LINKEDIT是支持dyld的，里面包含一些符号表等数据
'这里有个命名的问题，如下图所示，大写的__TEXT代表的是 Segment，小写的__text代表 Section

'Constants for the segment_command_flags field
Private Enum segment_command_flags
	HIGH_VM					= &H1
	FVM_LIB					= &H2
	NO_RELOC				= &H4
	PROTECTION_VERSION_1	= &H8
End Enum

'节结构
'和上面一样，有一些内存偏移和文件偏移，还有重定位节的引用，详细的需要以后慢慢理解。
'重要的也是flags，指明了当前节的属性。其中节属性可能有后面这样的
'要观察节内容，二进制数据，转到真的fileOffset处再读数据。
Private Type command_section_32	'68个字节
	sectname(15)		As Byte			'name of this section  16 个字符，节名称
	segname(15)			As Byte			'segment this section goes in  16 个字符，所在段的名称
	laddr				As Long			'memory address of this section 虚拟地址
	lsize				As Long			'size in bytes of this section 虚拟大小
	loffset				As Long			'file offset of this section 文件偏移量
	lalign				As Long			'section alignment (power of 2)节对齐值
	lreloff				As Long			'file offset of relocation entries
	lnreloc				As Long			'number of relocation entries
	lflags				As Long			'flags (section type and attributes)
	lreserved1			As Long			'reserved (for offset or index)
	lreserved2			As Long			'reserved (for count or sizeof)
End Type

Private Type command_section_64	'80字节
	sectname(15)		As Byte			'name of this section  16 个字符，节名称
	segname(15)			As Byte			'segment this section goes in  16 个字符，所在段的名称
	daddr1				As Long			'memory address of this section 虚拟地址第一部分
	daddr2				As Long			'memory address of this section 虚拟地址第二部分
	dsize1				As Long			'size in bytes of this section 虚拟大小第一部分
	dsize2				As Long			'size in bytes of this section 虚拟大小第二部分
	loffset				As Long			'file offset of this section 文件偏移量
	lalign				As Long			'section alignment (power of 2)节对齐值
	lreloff				As Long			'file offset of relocation entries
	lnreloc				As Long			'number of relocation entries
	lflags				As Long			'flags (section type and attributes)
	lreserved1			As Long			'reserved (for offset or index)
	lreserved2			As Long			'reserved (for count or sizeof)
	lreserved3			As Long
End Type
'section cmd 说明
'__text 主程序代码
'__stubs 用于动态库链接的桩
'__stub_helper 用于动态库链接的桩
'__cstring 常亮字符串符号表描述信息，通过该区信息，可以获得常亮字符串符号表地址
'__unwind_info 这里字段不是太理解啥意思，希望大家指点下

'在 __TEXT段里, 存在四个主要的 section:
'__text 主程序代码
'__const : 通用常量数据.
'__cstring : 字面量字符串常量.
'__picsymbol_stub : 动态链接器使用的位置无关码 stub 路由.
'这样保持了可执行的和不可执行的代码在段里的明显隔离.

'节属性, 默认的代码节就是这个2个属性
Private Enum command_section
	S_ATTR_PURE_INSTRUCTIONS = &H80000000
	S_ATTR_SOME_INSTRUCTIONS = &H00000400
End Enum

'LOAD COMMAND 属性
Private Type MAC_FILE_LOAD_COMMAND
	lOffset			As Long
	LoadCmd			As mac_load_command
End Type

'COMMAND_32 属性
Private Type MAC_FILE_COMMAND_32
	Index			As Integer
	CMD				As segment_command_32
	Section() 		As command_section_32
End Type

'COMMAND_64 属性
Private Type MAC_FILE_COMMAND_64
	Index			As Integer
	CMD				As segment_command_64
	Section() 		As command_section_64
End Type

Private MacHeader32 	As mac_header_32
Private MacHeader64 	As mac_header_64
Private MacLoadCmd()	As MAC_FILE_LOAD_COMMAND
Private MacCmd32()		As MAC_FILE_COMMAND_32
Private MacCmd64()		As MAC_FILE_COMMAND_64


'获取文件及子文件的数据结构信息
Public Function GetMacHeaders(ByVal strFilePath As String,File As FILE_PROPERTIE,ByVal Mode As Long) As Boolean
	Dim i As Long,FN As FILE_IMAGE,TempList() As String,Temp As String
	On Error GoTo ExitFunction
	File.FileSize = FileLen(strFilePath)
	'打开文件
	Mode = LoadFile(strFilePath,FN,0,0,0,Mode)
	If Mode < -1 Then Exit Function
	'获取主文件头
	GetMacHeaders = GetMacHeader(FN,File,Mode)
	If GetMacHeaders = False Then GoTo ExitFunction
	'获取子文件头
	With File.SecList(File.MaxSecIndex)
		If .lSizeOfRawData = 0 Then GoTo ExitFunction
		Temp = ByteToString(GetBytes(FN,.lSizeOfRawData,.lPointerToRawData,Mode),CP_ISOLATIN1)
		TempList = GetVAListRegExp(Temp,"(\xCE\xFA\xED\xFE)|(\xCF\xFA\xED\xFE)",.lPointerToRawData)
		If CheckArray(TempList) = False Then GoTo ExitFunction
		Dim SubFile As FILE_PROPERTIE
		File.NumberOfSub = UBound(TempList) + 1
		For i = 0 To File.NumberOfSub - 1
			'If GetMacHeader(FN,SubFile,Mode,CLng(TempList(i))) = True Then
				'修改主文件的隐藏节大小
				.lSizeOfRawData = CLng(TempList(i)) - .lPointerToRawData
				Exit For
			'End If
		Next i
	End With
	ExitFunction:
	'关闭文件
	On Error Resume Next
	UnLoadFile(FN,0,Mode)
End Function


'获取文件数据结构信息
Private Function GetMacHeader(FN As FILE_IMAGE,File As FILE_PROPERTIE,ByVal Mode As Long,Optional ByVal Offset As Long = -1) As Boolean
	Dim i As Long,j As Long,k As Long,n As Long
	Dim tmpMacHeader32 	As mac_header_32
	Dim tmpMacHeader64 	As mac_header_64
	Dim tmpMacHeaderFAT	As mac_header_fat
	Dim tmpMacLoadCmd()	As MAC_FILE_LOAD_COMMAND
	Dim tmpMacCmd32()	As MAC_FILE_COMMAND_32
	Dim tmpMacCmd64()	As MAC_FILE_COMMAND_64
	ReDim File.SecList(1)				'As SECTION_PROPERTIE
	ReDim File.SecList(0).SubSecList(0)	'As SUB_SECTION_PROPERTIE
	ReDim File.SecList(1).SubSecList(0)	'As SUB_SECTION_PROPERTIE
	ReDim File.DataDirectory(0)			'As SECTION_PROPERTIE
	ReDim File.CLRList(0)				'As SECTION_PROPERTIE
	ReDim File.StreamList(0)			'As SECTION_PROPERTIE
	k = Offset
	If k = -1 Then k = File.FileType
	On Error GoTo ExitFunction
	With File
		'初始化
		.Magic = ""
		.FileAlign = 0
		.SecAlign = 0
		.ImageBase = 0
		.DataDirs = 0
		.LangType = 0
		.MinSecID = 0
		.MaxSecID = 0
		.MaxSecIndex = 1
		.USStreamID = -1
		.NumberOfSub = 0
		.NetStreams = 0
		.SecList(0).SubSecs = 0
		.SecList(1).SubSecs = 0

		'获取 FAT Header
		Select Case GetLong(FN,k,Mode)
		Case MH_MAGIC_FAT,MH_MAGIC_FAT_CIGAM
			With tmpMacHeaderFAT
				.lmagic = GetLong(FN,k,Mode)
				.lfat_arch_size = Bytes2Val(GetBytes(FN,4,k + 4,Mode),4,True)
				If .lfat_arch_size = 0 Then GoTo ExitFunction
				ReDim tmpMacHeaderFAT.fat_archs(.lfat_arch_size - 1)	'As mac_header_fat_arch
				k = k + 8
				For i = 0 To .lfat_arch_size - 1
					.fat_archs(i).lcputype = Bytes2Val(GetBytes(FN,4,k,Mode),4,True)
					.fat_archs(i).lcpusubtype = Bytes2Val(GetBytes(FN,4,k + 4,Mode),4,True)
					.fat_archs(i).lfileoffset = Bytes2Val(GetBytes(FN,4,k + 8,Mode),4,True)
					.fat_archs(i).lsize	= Bytes2Val(GetBytes(FN,4,k + 12,Mode),4,True)
					.fat_archs(i).lalign = Bytes2Val(GetBytes(FN,4,k + 16,Mode),4,True)
					k = k + 20
				Next i
				k = tmpMacHeaderFAT.fat_archs(0).lfileoffset
			End With
		End Select

		'获取 Header
		Select Case GetLong(FN,k,Mode)
		Case MH_MAGIC_32
			Select Case Mode
			Case Is < 0
				Get #FN.hFile, k + 1, tmpMacHeader32
			Case 0
				CopyMemory tmpMacHeader32, FN.ImageByte(k), Len(tmpMacHeader32)
			Case Else
				MoveMemory tmpMacHeader32, FN.MappedAddress + k, Len(tmpMacHeader32)
			End Select
			If tmpMacHeader32.lncmds = 0 Then GoTo ExitFunction
			.Magic = "MAC32"
			.MaxSecIndex = tmpMacHeader32.lncmds
			k = k + Len(tmpMacHeader32)
		Case MH_MAGIC_64
			Select Case Mode
			Case Is < 0
				Get #FN.hFile, k + 1, tmpMacHeader64
			Case 0
				CopyMemory tmpMacHeader64, FN.ImageByte(k), Len(tmpMacHeader64)
			Case Else
				MoveMemory tmpMacHeader64, FN.MappedAddress + k, Len(tmpMacHeader64)
			End Select
			If tmpMacHeader64.lncmds = 0 Then GoTo ExitFunction
			.Magic = "MAC64"
			.MaxSecIndex = tmpMacHeader64.lncmds
			k = k + Len(tmpMacHeader64)
		Case Else
			GoTo ExitFunction
		End Select

		'获取 Command 段
		ReDim tmpMacLoadCmd(.MaxSecIndex - 1)	'As MAC_FILE_LOAD_COMMAND
		ReDim tmpMacCmd32(.MaxSecIndex - 1) 	'As MAC_FILE_COMMAND_32
		ReDim tmpMacCmd64(.MaxSecIndex - 1)	'As MAC_FILE_COMMAND_64
		ReDim File.SecList(.MaxSecIndex)	'As SECTION_PROPERTIE
		For i = 0 To .MaxSecIndex - 1
			tmpMacLoadCmd(i).loffset = k
			If k + tmpMacLoadCmd(i).LoadCmd.lcmdsize <= .FileSize Then
				'获取 Load Command 段
				Select Case Mode
				Case Is < 0
					Get #FN.hFile, k + 1, tmpMacLoadCmd(i).LoadCmd
				Case 0
					CopyMemory tmpMacLoadCmd(i).LoadCmd, FN.ImageByte(k), Len(tmpMacLoadCmd(i).LoadCmd)
				Case Else
					MoveMemory tmpMacLoadCmd(i).LoadCmd, FN.MappedAddress + k, Len(tmpMacLoadCmd(i).LoadCmd)
				End Select

				'获取 Command 段
				Select Case tmpMacLoadCmd(i).LoadCmd.lcmd
				Case SEGMENT	'32位标准 Command
					'获取 Command 数据
					j = k + Len(tmpMacLoadCmd(i).LoadCmd)
					tmpMacCmd32(n).Index = i
					Select Case Mode
					Case Is < 0
						Get #FN.hFile, j + 1, tmpMacCmd32(n).CMD
					Case 0
						CopyMemory tmpMacCmd32(n).CMD, FN.ImageByte(j), Len(tmpMacCmd32(n).CMD)
					Case Else
						MoveMemory tmpMacCmd32(n).CMD, FN.MappedAddress + j, Len(tmpMacCmd32(n).CMD)
					End Select

					'获取节数据
					ReDim Preserve File.SecList(n).SubSecList(0) 'SECTION_PROPERTIE
					ReDim Preserve tmpMacCmd32(n).Section(0)	'As command_section_32
					If tmpMacCmd32(n).CMD.lnsects > 0 Then
						ReDim Preserve File.SecList(n).SubSecList(tmpMacCmd32(n).CMD.lnsects - 1) 'SECTION_PROPERTIE
						ReDim Preserve tmpMacCmd32(n).Section(tmpMacCmd32(n).CMD.lnsects - 1)	'As command_section_32
						j = j + Len(tmpMacCmd32(n).CMD)
						Select Case Mode
						Case Is < 0
							Get #FN.hFile, j + 1, tmpMacCmd32(n).Section
						Case 0
							CopyMemory tmpMacCmd32(n).Section(0), FN.ImageByte(j), Len(tmpMacCmd32(n).Section(0)) * tmpMacCmd32(n).CMD.lnsects
						Case Else
							MoveMemory tmpMacCmd32(n).Section(0), FN.MappedAddress + j, Len(tmpMacCmd32(n).Section(0)) * tmpMacCmd32(n).CMD.lnsects
						End Select
						'记录节地址
						For j = 0 To tmpMacCmd32(n).CMD.lnsects - 1
							With tmpMacCmd32(n).Section(j)
								File.SecList(n).SubSecList(j).sName = Replace$(StrConv$(.sectname,vbUnicode),vbNullChar,"")
								File.SecList(n).SubSecList(j).lPointerToRawData = .loffset
								File.SecList(n).SubSecList(j).lSizeOfRawData = .lsize
								File.SecList(n).SubSecList(j).lVirtualAddress = .laddr
								File.SecList(n).SubSecList(j).lVirtualSize = .lsize
							End With
						Next j
					Else
						ReDim File.SecList(n).SubSecList(0) 'SECTION_PROPERTIE
					End If
					'记录区段地址
					.SecList(n).sName = Replace$(StrConv$(tmpMacCmd32(n).CMD.segname,vbUnicode),vbNullChar,"")
					.SecList(n).lPointerToRawData = tmpMacCmd32(n).CMD.lfileoff
					.SecList(n).lSizeOfRawData = tmpMacCmd32(n).CMD.lfilesize
					.SecList(n).lVirtualAddress = tmpMacCmd32(n).CMD.lvmaddr
					.SecList(n).lVirtualSize = tmpMacCmd32(n).CMD.lvmsize
					.SecList(n).SubSecs = tmpMacCmd32(n).CMD.lnsects
					If .SecList(n).lSizeOfRawData > 0 Then n = n + 1
				Case SEGMENT_64	'64位标准 Command
					'获取 Command 数据类型
					j = k + Len(tmpMacLoadCmd(i).LoadCmd)
					tmpMacCmd64(n).Index = i
					Select Case Mode
					Case Is < 0
						Get #FN.hFile, j + 1, tmpMacCmd64(n).CMD
					Case 0
						CopyMemory tmpMacCmd64(n).CMD, FN.ImageByte(j), Len(tmpMacCmd64(n).CMD)
					Case Else
						MoveMemory tmpMacCmd64(n).CMD, FN.MappedAddress + j, Len(tmpMacCmd64(n).CMD)
					End Select

					'获取节数据
					ReDim Preserve File.SecList(n).SubSecList(0) 'SECTION_PROPERTIE
					ReDim Preserve tmpMacCmd64(n).Section(0)	'As command_section_32
					If tmpMacCmd64(n).CMD.lnsects > 0 Then
						ReDim Preserve File.SecList(n).SubSecList(tmpMacCmd64(n).CMD.lnsects - 1) 'SECTION_PROPERTIE
						ReDim Preserve tmpMacCmd64(n).Section(tmpMacCmd64(n).CMD.lnsects - 1)	'As command_section_64
						j = j + Len(tmpMacCmd64(n).CMD)
						Select Case Mode
						Case Is < 0
							Get #FN.hFile, j + 1, tmpMacCmd64(n).Section
						Case 0
							CopyMemory tmpMacCmd64(n).Section(0), FN.ImageByte(j), Len(tmpMacCmd64(n).Section(0)) * tmpMacCmd64(n).CMD.lnsects
						Case Else
							MoveMemory tmpMacCmd64(n).Section(0), FN.MappedAddress + j, Len(tmpMacCmd64(n).Section(0)) * tmpMacCmd64(n).CMD.lnsects
						End Select
						'记录节地址
						For j = 0 To tmpMacCmd64(n).CMD.lnsects - 1
							With tmpMacCmd64(n).Section(j)
								File.SecList(n).SubSecList(j).sName = Replace$(StrConv$(.sectname,vbUnicode),vbNullChar,"")
								File.SecList(n).SubSecList(j).lPointerToRawData = .loffset
								File.SecList(n).SubSecList(j).lSizeOfRawData = .dsize1
								File.SecList(n).SubSecList(j).lVirtualAddress = .daddr1
								File.SecList(n).SubSecList(j).lVirtualAddress1 = .daddr2
								File.SecList(n).SubSecList(j).lVirtualSize = .dsize1
							End With
						Next j
					Else
						ReDim File.SecList(n).SubSecList(0) 'SECTION_PROPERTIE
					End If
					'记录区段地址
					.SecList(n).sName = Replace$(StrConv$(tmpMacCmd64(n).CMD.segname,vbUnicode),vbNullChar,"")
					.SecList(n).lPointerToRawData = tmpMacCmd64(n).CMD.dfileoff1
					.SecList(n).lSizeOfRawData = tmpMacCmd64(n).CMD.dfilesize1
					.SecList(n).lVirtualAddress = tmpMacCmd64(n).CMD.dvmaddr1
					.SecList(n).lVirtualAddress1 = tmpMacCmd64(n).CMD.dvmaddr2
					.SecList(n).lVirtualSize = tmpMacCmd64(n).CMD.dvmsize1
					.SecList(n).SubSecs = tmpMacCmd64(n).CMD.lnsects
					If .SecList(n).lSizeOfRawData > 0 Then n = n + 1
				End Select
			End If
			k = k + tmpMacLoadCmd(i).LoadCmd.lcmdsize
		Next i
		If n > 0 Then
			.MaxSecIndex = n
			ReDim Preserve tmpMacCmd32(n - 1) As MAC_FILE_COMMAND_32
			ReDim Preserve tmpMacCmd64(n - 1) As MAC_FILE_COMMAND_64
			ReDim Preserve File.SecList(n) 'SECTION_PROPERTIE
			'第一个段都是从0开始，包括了文件头，所以要调整
			For i = 0 To .MaxSecIndex - 1
				If .SecList(i).lPointerToRawData = 0 Then
					If .SecList(i).SubSecs > 0 Then
						.SecList(i).lPointerToRawData = .SecList(i).SubSecList(0).lPointerToRawData
						.SecList(i).lSizeOfRawData = .SecList(i).lSizeOfRawData - .SecList(i).lPointerToRawData
					End If
				End If
				If .SecList(i).lVirtualAddress = 0 Then
					If .SecList(i).SubSecs > 0 Then
						.SecList(i).lVirtualAddress = .SecList(i).SubSecList(0).lVirtualAddress
						.SecList(i).lVirtualSize = .SecList(i).lVirtualSize - .SecList(i).lVirtualAddress
					End If
				End If
			Next i
		Else
			.MaxSecIndex = 1
			ReDim Preserve tmpMacCmd32(n - 1) As MAC_FILE_COMMAND_32
			ReDim Preserve tmpMacCmd64(n - 1) As MAC_FILE_COMMAND_64
			ReDim File.SecList(1) 'SECTION_PROPERTIE
			GoTo ExitFunction
		End If

		'获取文件节最大索引号、最小和最大偏移地址所在节的索引号
		Call GetSectionID(File,.MinSecID,.MaxSecID,False)

		'获取隐藏节信息
		.SecList(.MaxSecIndex).lPointerToRawData = .SecList(.MaxSecID).lPointerToRawData + .SecList(.MaxSecID).lSizeOfRawData
		.SecList(.MaxSecIndex).lSizeOfRawData = GetFileLength(FN,Mode) - .SecList(.MaxSecIndex).lPointerToRawData
		.SecList(.MaxSecIndex).lVirtualAddress = .SecList(.MaxSecID).lVirtualAddress + .SecList(.MaxSecID).lVirtualSize
		.SecList(.MaxSecIndex).lVirtualSize = .SecList(.MaxSecIndex).lSizeOfRawData
	End With

	'记录主程序的各种头数据
	If Offset = -1 Then
		MacHeader32 = tmpMacHeader32
		MacHeader64 = tmpMacHeader64
		MacLoadCmd = tmpMacLoadCmd
		MacCmd32 = tmpMacCmd32
		MacCmd64 = tmpMacCmd64
	End If

	'标记成功
	GetMacHeader = True
	Exit Function

	ExitFunction:
	ReDim File.SecList(1)			'As SECTION_PROPERTIE
	ReDim File.DataDirectory(15)		'As SECTION_PROPERTIE
	ReDim File.CLRList(0)			'As SECTION_PROPERTIE
	ReDim File.StreamList(0)		'As SUB_SECTION_PROPERTIE
	With File
		.FileType = 0
		.Magic = ""
		.FileAlign = 0
		.SecAlign = 0
		.ImageBase = 0
		.DataDirs = 0
		.LangType = 0
		.MinSecID = 0
		.MaxSecID = 0
		.MaxSecIndex = 1
		.USStreamID = -1
		.NumberOfSub = 0
		.NetStreams = 0
		.SecList(0).SubSecs = 0
		.SecList(1).SubSecs = 0
		'设置整个文件为一个节
		.SecList(0).lPointerToRawData = 0
		.SecList(0).lSizeOfRawData = GetFileLength(FN,Mode)
		.SecList(0).lVirtualAddress = 0
		.SecList(0).lVirtualSize = .SecList(0).lSizeOfRawData
		'设置隐藏节信息，用于显示文件信息
		.SecList(1).lPointerToRawData = .SecList(0).lSizeOfRawData
		.SecList(1).lSizeOfRawData = 0
		.SecList(1).lVirtualAddress = .0
		.SecList(1).lVirtualSize = 0
	End With
End Function


'相对虚拟地址转偏移地址
Private Function RvaToOffset(File As FILE_PROPERTIE,ByVal dwRvaAddr As Long) As Long
	Dim i As Integer
	On Error GoTo ErrHandle
	For i = 0 To UBound(File.SecList)
		With File.SecList(i)
			If dwRvaAddr >= .lVirtualAddress Then
				If dwRvaAddr < .lVirtualAddress + .lVirtualSize Then
					RvaToOffset = dwRvaAddr + .lPointerToRawData - .lVirtualAddress
					Exit Function
				End If
			End If
		End With
	Next i
	Exit Function
	ErrHandle:
	RvaToOffset = 0
End Function


'偏移地址转相对虚拟地址
Private Function OffsetToRva(File As FILE_PROPERTIE,ByVal dwOffset As Long) As Long
	Dim i As Integer
	On Error GoTo ErrHandle
	For i = 0 To UBound(File.SecList)
		With File.SecList(i)
			If dwOffset >= .lPointerToRawData Then
				If dwOffset < .lPointerToRawData + .lSizeOfRawData Then
					OffsetToRva = dwOffset + .lVirtualAddress - .lPointerToRawData
					Exit Function
				End If
			End If
		End With
	Next i
	Exit Function
	ErrHandle:
	OffsetToRva = 0
End Function


'获取段的名称
Private Function LoadCommandTypeRead(ByVal LoadCommandType As Long) As String
	Select Case	LoadCommandType
		Case SEGMENT
			LoadCommandTypeRead = "SEGMENT"
		Case SYM_TAB
			LoadCommandTypeRead = "SYM_TAB"
		Case SYM_SEG
			LoadCommandTypeRead = "SYM_SEG"
		Case THREAD
			LoadCommandTypeRead = "THREAD"
		Case UNIX_THREAD
			LoadCommandTypeRead = "UNIX_THREAD"
		Case LOAD_FVM_LIB
			LoadCommandTypeRead = "LOAD_FVM_LIB"
		Case ID_FVM_LIB
			LoadCommandTypeRead = "ID_FVM_LIB"
		Case IDENT
			LoadCommandTypeRead = "IDENT"
		Case FVM_FILE
			LoadCommandTypeRead = "FVM_FILE"
		Case PREPAGE
			LoadCommandTypeRead = "PREPAGE"
		Case DY_SYM_TAB
			LoadCommandTypeRead = "DY_SYM_TAB"
		Case LOAD_DYLIB
			LoadCommandTypeRead = "LOAD_DYLIB"
		Case ID_DYLIB
			LoadCommandTypeRead = "ID_DYLIB"
		Case LOAD_DYLINKER
			LoadCommandTypeRead = "LOAD_DYLINKER"
		Case ID_DYLINKER
			LoadCommandTypeRead = "ID_DYLINKER"
		Case PREBOUND_DYLIB
			LoadCommandTypeRead = "PREBOUND_DYLIB"
		Case ROUTINES
			LoadCommandTypeRead = "ROUTINES"
		Case SUB_FRAMEWORK
			LoadCommandTypeRead = "SUB_FRAMEWORK"
		Case SUB_UMBRELLA
			LoadCommandTypeRead = "SUB_UMBRELLA"
		Case SUB_CLIENT
			LoadCommandTypeRead = "SUB_CLIENT"
		Case SUB_LIBRARY
			LoadCommandTypeRead = "SUB_LIBRARY"
		Case TWOLEVEL_HINTS
			LoadCommandTypeRead = "TWOLEVEL_HINTS"
		Case PREBIND_CKSUM
			LoadCommandTypeRead = "PREBIND_CKSUM"
		Case LOAD_WEAK_DYLIB
			LoadCommandTypeRead = "LOAD_WEAK_DYLIB"
		Case SEGMENT_64
			LoadCommandTypeRead = "SEGMENT_64"
		Case ROUTINES_64
			LoadCommandTypeRead = "ROUTINES_64"
		Case UUID
			LoadCommandTypeRead = "UUID"
		Case RPATH
			LoadCommandTypeRead = "RPATH"
		Case CODE_SIGNATURE
			LoadCommandTypeRead = "CODE_SIGNATURE"
		Case SEGMENT_SPLIT_INFO
			LoadCommandTypeRead = "SEGMENT_SPLIT_INFO"
		Case REEXPORT_DYLIB
			LoadCommandTypeRead = "REEXPORT_DYLIB"
		Case LAZY_LOAD_DYLIB
			LoadCommandTypeRead = "LAZY_LOAD_DYLIB"
		Case ENCRYPTION_INFO
			LoadCommandTypeRead = "ENCRYPTION_INFO"
		Case DYLD_INFO
			LoadCommandTypeRead = "DYLD_INFO"
		Case DYLD_INFO_ONLY
			LoadCommandTypeRead = "DYLD_INFO_ONLY"
		Case LOAD_UPWARD_DYLIB
			LoadCommandTypeRead = "LOAD_UPWARD_DYLIB"
		Case VERSION_MIN_MAC_OSX
			LoadCommandTypeRead = "VERSION_MIN_MAC_OSX"
		Case VERSION_MIN_IPHONE_OS
			LoadCommandTypeRead = "VERSION_MIN_IPHONE_OS"
		Case FUNCTION_STARTS
			LoadCommandTypeRead = "FUNCTION_STARTS"
		Case DYLD_ENVIRONMENT
			LoadCommandTypeRead = "DYLD_ENVIRONMENT"
		Case MAIN_CMD
			LoadCommandTypeRead = "MAIN"
		Case DATA_IN_CODE
			LoadCommandTypeRead = "DATA_IN_CODE"
		Case SOURCE_VERSION
			LoadCommandTypeRead = "SOURCE_VERSION"
		Case DYLIB_CODE_SIGN_DRS
			LoadCommandTypeRead = "DYLIB_CODE_SIGN_DRS"
		Case Else	'Default
			LoadCommandTypeRead = "Error"
	End Select
End Function


'增加文件节长度
'fType = 0 获取可增加的字节数，修改原始文件头不写入
'fType = 1 写入指定长度对齐后的字节数，不修改原始文件头仅写入
'fType = 2 写入指定长度对齐后的字节数，修改原始文件头
'fType = 3 不修改不写入，仅获取可增加值(AddSecSize(x).Length 为偏移大小，AddSecSize(x).Address 为虚拟大小)
'AddSecSize(x).Length = 0，按可增加的最大值增加，否则按 AddSecSize(x).Length 对齐值增加
Public Function AddMacSectionSize(trnFile As FILE_PROPERTIE,AddSecSize() As FREE_BTYE_SPACE,ByVal fType As Long,ByVal Mode As Long) As Long
	Dim i As Long,j As Integer,k As Long,x As Long,n As Long
	Dim SecAlign As Long,AddRAW As Long,AddRVA As Long
	Dim FN As FILE_IMAGE,TempByte() As Byte,File As FILE_PROPERTIE

	'错误处理
	On Error GoTo localError

	'获取 PE 头
	File = trnFile
	If GetMacHeaders(File.FilePath,File,Mode) = False Then Exit Function

	'修改文件节的开始地址和大小
	With File
		For i = 0 To UBound(AddSecSize)
			j = AddSecSize(i).inSectionID
			'根据是否是最后节获取可增加字节
			If j = File.MaxSecID Then
				k = .SecList(j).lSizeOfRawData + AddSecSize(i).Length
				n = 1
			Else
				k = .SecList(GetSectionID(File,j,-1,True)).lVirtualAddress - .SecList(j).lVirtualAddress
				n = 0
			End If
			'按文件对齐值对齐
			x = k - .SecList(j).lSizeOfRawData
			If x > 0 Or n > 0 Then
				'根据实际需要增加当前节的偏移大小并对齐
				If AddSecSize(i).Length > 0 Then
					x = IIf(x > AddSecSize(i).Length,AddSecSize(i).Length,x)
				End If
				AddSecSize(i).Length = x: AddRAW = AddRAW + x

				'增加当前节的虚拟大小，虚拟大小不用对齐
				If AddSecSize(i).Length > 0 Then
					x = k - .SecList(j).lVirtualSize
					If x > 0 Then
						If x > AddSecSize(i).Length Then x = AddSecSize(i).Length
						.SecList(j).lVirtualSize = .SecList(j).lVirtualSize + x
						If fType > 2 Then AddSecSize(i).Address = x
						AddRVA = AddRVA + x
					End If
				End If

				'记录可增加值
				If fType = 0 Then
					'计算并记录移位地址，用于字串移位计算操作
					AddSecSize(i).Address = .SecList(j).lPointerToRawData + .SecList(j).lSizeOfRawData
					'AddSecSize(i).inSectionID = j
					If .SecList(j).SubSecs = 0 Then
						AddSecSize(i).inSubSecID = 0
					Else
						AddSecSize(i).inSubSecID = .SecList(j).SubSecs - 1
					End If
					AddSecSize(i).MaxAddress = AddSecSize(i).Address + AddSecSize(i).Length - 1
					AddSecSize(i).lNumber = -AddSecSize(i).Address
					AddSecSize(i).MoveType = -3	'新增节尾空位，不参与空位管理
				ElseIf fType = 1 Or fType = 2 Then
					'计算原文件当前节的偏移地址和大小，用于后面的实际移位操作
					AddSecSize(i).Address = trnFile.SecList(j).lPointerToRawData
					AddSecSize(i).MaxAddress = AddSecSize(i).Address + trnFile.SecList(j).lSizeOfRawData - 1
				End If

				'修改当前节的偏移大小及后面节的偏移地址
				If AddSecSize(i).Length > 0 Then
					For k = 0 To .MaxSecIndex - 1
						If .SecList(k).lPointerToRawData > .SecList(j).lPointerToRawData Then
							.SecList(k).lPointerToRawData = .SecList(k).lPointerToRawData + AddSecSize(i).Length
							If .SecList(k).SubSecs > 0 Then
								With .SecList(k)
									For n = 0 To .SubSecs - 1
										.SubSecList(n).lPointerToRawData = .SubSecList(n).lPointerToRawData + AddSecSize(i).Length
									Next n
								End With
							End If
						End If
					Next k
					.SecList(j).lSizeOfRawData = .SecList(j).lSizeOfRawData + AddSecSize(i).Length
				End If
			Else
				AddSecSize(i).Length = 0
			End If
		Next i
	End With
	If AddRAW = 0 Then Exit Function
	If fType > 2 Then
		AddMacSectionSize = AddRAW
		Exit Function
	End If

	'修改隐藏节的偏移地址，还原隐藏节大小
	'由于重新读取的文件可能已在文件尾部写入了字串，这些字节将作为隐藏节被读取，故需要使用原始文件的隐藏节信息
	With File.SecList(File.MaxSecIndex)
		.lPointerToRawData = File.SecList(File.MaxSecID).lPointerToRawData + File.SecList(File.MaxSecID).lSizeOfRawData
		.lVirtualAddress = File.SecList(File.MaxSecID).lVirtualAddress + File.SecList(File.MaxSecID).lVirtualSize
		.lSizeOfRawData = trnFile.SecList(trnFile.MaxSecIndex).lSizeOfRawData
		.lVirtualSize = trnFile.SecList(trnFile.MaxSecIndex).lVirtualSize
	End With

	'修改目标文件的区段数据
	If fType = 0 Then
		trnFile = File
		AddMacSectionSize = AddRAW
		Exit Function
	End If

	'打开文件
	Mode = LoadFile(File.FilePath,FN,0,1,0,Mode)
	If Mode < -1 Then GoTo localError

	'修改 Header 数据
	With File
		For i = AddSecSize(0).inSectionID To .MaxSecIndex - 1
			If .Magic = "MAC32" Then
				j = MacCmd32(i).Index
			Else
				j = MacCmd64(i).Index
			End If
			'获取 Command 段类型
			Select Case MacLoadCmd(j).LoadCmd.lcmd
			Case SEGMENT	'32位标准 Command
				'修改 MacCmdHeader 数据
				If i > 0 Then MacCmd32(i).CMD.lfileoff = .SecList(i).lPointerToRawData
				MacCmd32(i).CMD.lfilesize = .SecList(i).lSizeOfRawData
				MacCmd32(i).CMD.lvmsize = .SecList(i).lVirtualSize
				'修改子节的地址和大小
				If MacCmd32(i).CMD.lnsects > 0 Then
					For k = 0 To MacCmd32(i).CMD.lnsects - 1
						.SecAlign = MacCmd32(i).Section(k).lalign
						MacCmd32(i).Section(k).loffset = .SecList(i).SubSecList(k).lPointerToRawData
						MacCmd32(i).Section(k).lsize = Alignment(.SecList(i).SubSecList(k).lSizeOfRawData,.SecAlign,1)
					Next k
					'If PutTypeValue(FN,MacBitType + MacCmd32(i).lOffset,MacCmd32(i),Mode) = False Then GoTo localError
					x = MacLoadCmd(j).lOffset + Len(MacLoadCmd(0).LoadCmd)
					k = Len(MacCmd32(i).CMD)
					Select Case Mode
					Case Is < 0
						Put #FN.hFile,x + 1,MacCmd32(i).CMD
						Put #FN.hFile,x + k + 1,MacCmd32(i).Section
					Case 0
						CopyMemory FN.ImageByte(x),MacCmd32(i).CMD,Len(MacCmd32(i).CMD)
						CopyMemory FN.ImageByte(x + k),MacCmd32(i).Section(0),Len(MacCmd32(i).Section(0)) * MacCmd32(i).CMD.lnsects
					Case Else
						WriteMemory FN.MappedAddress + x,MacCmd32(i).CMD,Len(MacCmd32(i).CMD)
						WriteMemory FN.MappedAddress + x + k,MacCmd32(i).Section(0),Len(MacCmd32(i).Section(0)) * MacCmd32(i).CMD.lnsects
					End Select
				Else
					'If PutTypeValue(FN,MacBitType + MacCmd32(i).lOffset,MacCmd32(i),Mode) = False Then GoTo localError
					x = MacLoadCmd(j).lOffset + Len(MacLoadCmd(0).LoadCmd)
					Select Case Mode
					Case Is < 0
						Put #FN.hFile,x + 1,MacCmd32(i).CMD
					Case 0
						CopyMemory FN.ImageByte(x),MacCmd32(i).CMD,Len(MacCmd32(i).CMD)
					Case Else
						WriteMemory FN.MappedAddress + x,MacCmd32(i).CMD,Len(MacCmd32(i).CMD)
					End Select
				End If
			Case SEGMENT_64	'64位标准 Command
				'修改父节的地址和大小
				If i > 0 Then MacCmd64(i).CMD.dfileoff1 = .SecList(i).lPointerToRawData
				MacCmd64(i).CMD.dfilesize1 = .SecList(i).lSizeOfRawData
				MacCmd64(i).CMD.dvmsize1 = .SecList(i).lVirtualSize
				'修改子节的地址和大小
				If MacCmd64(i).CMD.lnsects > 0 Then
					For k = 0 To MacCmd64(i).CMD.lnsects - 1
						.SecAlign = MacCmd64(i).Section(k).lalign
						MacCmd64(i).Section(k).loffset = .SecList(i).SubSecList(k).lPointerToRawData
						MacCmd64(i).Section(k).dsize1 = Alignment(.SecList(i).SubSecList(k).lSizeOfRawData,.SecAlign,1)
					Next k
					'If PutTypeValue(FN,MacBitType + MacCmd64(i).lOffset,MacCmd64(i),Mode) = False Then GoTo localError
					x = MacLoadCmd(j).lOffset + Len(MacLoadCmd(0).LoadCmd)
					k = Len(MacCmd64(i).CMD)
					Select Case Mode
					Case Is < 0
						Put #FN.hFile,x + 1,MacCmd64(i).CMD
						Put #FN.hFile,x + k + 1,MacCmd64(i).Section
					Case 0
						CopyMemory FN.ImageByte(x),MacCmd64(i).CMD,k
						CopyMemory FN.ImageByte(x + k),MacCmd64(i).Section(0),Len(MacCmd64(i).Section(0)) * MacCmd64(i).CMD.lnsects
					Case Else
						WriteMemory FN.MappedAddress + x,MacCmd64(i).CMD,k
						WriteMemory FN.MappedAddress + x + k,MacCmd64(i).Section(0),Len(MacCmd64(i).Section(0)) * MacCmd64(i).CMD.lnsects
					End Select
				Else
					'If PutTypeValue(FN,MacBitType + MacCmd64(i).lOffset,MacCmd64(i),Mode) = False Then GoTo localError
					x = MacLoadCmd(j).lOffset + Len(MacLoadCmd(0).LoadCmd)
					Select Case Mode
					Case Is < 0
						Put #FN.hFile,x + 1,MacCmd64(i).CMD
					Case 0
						CopyMemory FN.ImageByte(x),MacCmd64(i).CMD,Len(MacCmd64(i).CMD)
					Case Else
						WriteMemory FN.MappedAddress + x,MacCmd64(i).CMD,Len(MacCmd64(i).CMD)
					End Select
				End If
			End Select
		Next i
	End With

	'从小到大对偏移地址进行排序，以避免因原节表乱序而扩展错误
	Call SortFreeByteByAddress(AddSecSize,0,UBound(AddSecSize),False)

	'移位最初要增加大小所在节后面每个节，并在节尾增加可要增加的空字节
	n = UBound(AddSecSize)
	For i = n To 0 Step -1
		With AddSecSize(i)
			j = AddSecSize(i).inSectionID
			'获取当前节后的全部字节长度
			If i = n Then
				If j = File.MaxSecID Then
					'由于获取文件头时会重新获取文件大小，而扩展最后节时会在扩展前先写入字串，
					'故要使用原来的文件大小，否则这些写入字串会被复制到最后
					k = trnFile.FileSize - .MaxAddress - 1
				Else
					k = File.FileSize - .MaxAddress - 1
				End If
			Else
				k = AddSecSize(i + 1).MaxAddress - .MaxAddress + .Length
			End If
			'移位字节到当前节的最大地址后面
			If k > 0 Then
				TempByte = GetBytes(FN,k,.MaxAddress + 1,Mode)
				PutBytes(FN,File.SecList(j).lPointerToRawData + File.SecList(j).lSizeOfRawData,TempByte,k,Mode)
			End If

			'增加当前节后的空字节(置空)
			'由于重新读取的文件可能已在文件尾部写入了字串，这些字节将作为隐藏节被读取，故需要使用原始文件的隐藏节信息
			If .Length > 0 Then
				'必须使用虚拟大小，因为其大小为最大节后的所有字节(包括子 PE)
				If j = File.MaxSecID And trnFile.Seclist(trnFile.MaxSecIndex).lVirtualSize < 1 Then
					x = GetFileLength(FN,Mode)
					k = trnFile.FileSize + .Length - x
					If k > 0 Then
						ReDim TempByte(k - 1) As Byte
						PutBytes(FN,x,TempByte,k,Mode)
					End If
				Else
					ReDim TempByte(.Length - 1) As Byte
					PutBytes(FN,.MaxAddress + 1,TempByte,.Length,Mode)
				End If
			End If
		End With
	Next i
	UnLoadFile(FN,FN.SizeOfFile,Mode)
	AddMacSectionSize = AddRAW

	'修改目标文件的区段数据
	If fType = 2 Then trnFile = File
	Exit Function

	'安全退出函数
	localError:
	UnLoadFile(FN,0,Mode)
	AddMacSectionSize = -1
End Function


'在文件尾部增加一个文件节
'fType = 0 只修改文件节数据不写入
'fType = 1 不修改文件节数据但写入
'fType = 2 既修改文件节数据又写入
Public Function AddMacSection(trnFile As FILE_PROPERTIE,AddSecSize As FREE_BTYE_SPACE,ByVal SecName As String,ByVal fType As Long,ByVal Mode As Long) As Long
	Dim i As Long,k As Long,n As Long
	Dim NewHeadersOffset As Long,NewSizeOfHeader As Long
	Dim FN As FILE_IMAGE,TempByte() As Byte,File As FILE_PROPERTIE

	'错误处理
	On Error GoTo localError
	If AddSecSize.Length = 0 Then Exit Function

	'获取 PE 头
	File = trnFile
	If GetMacHeaders(File.FilePath,File,Mode) = False Then Exit Function

	'修改文件对齐值，以减少新增节的大小
	File.FileAlign = 512
	File.SecAlign = 512

	'检测后面是否有多少空余空间可以容纳新段的需求
	Select Case File.Magic
	Case "MAC32"
		NewSizeOfHeader = MAC_BIT_TYPE32
		n = MacHeader32.lncmds
	Case "MAC64"
		NewSizeOfHeader = MAC_BIT_TYPE64
		n = MacHeader64.lncmds
	End Select
	k = MacLoadCmd(n - 1).loffset + MacLoadCmd(n - 1).LoadCmd.lcmdsize
	For i = 0 To File.MaxSecIndex - 1
		If File.Seclist(i).lPointerToRawData >= k Then
			If k + NewSizeOfHeader > File.Seclist(i).lPointerToRawData Then GoTo localError
			Exit For
		End If
	Next i

	'新段数据
	ReDim Preserve File.SecList(File.MaxSecIndex + 1) 'As SECTION_PROPERTIE
	ReDim Preserve File.SecList(File.MaxSecIndex + 1).SubSecList(0) 'As SUB_SECTION_PROPERTIE
	ReDim Preserve File.SecList(File.MaxSecIndex).SubSecList(0) 'As SUB_SECTION_PROPERTIE
	With File
		.SecList(.MaxSecIndex).sName = SecName
		.SecList(.MaxSecIndex).lPointerToRawData = .SecList(.MaxSecID).lPointerToRawData + Alignment(.SecList(.MaxSecID).lSizeOfRawData,.FileAlign,1)
		.SecList(.MaxSecIndex).lSizeOfRawData = Alignment(AddSecSize.Length,.FileAlign,1)
		.SecList(.MaxSecIndex).lVirtualAddress = .SecList(.MaxSecID).lVirtualAddress + Alignment(.SecList(.MaxSecID).lVirtualSize,.SecAlign,1)
		.SecList(.MaxSecIndex).lVirtualSize = AddSecSize.Length
		.SecList(.MaxSecIndex).RWA = .SecList(.MaxSecIndex).lPointerToRawData
		.SecList(.MaxSecIndex).SubSecList(0).sName = SecName
		.SecList(.MaxSecIndex).SubSecList(0).lPointerToRawData = .SecList(.MaxSecIndex).lPointerToRawData
		.SecList(.MaxSecIndex).SubSecList(0).lSizeOfRawData = .SecList(.MaxSecIndex).lSizeOfRawData
		.SecList(.MaxSecIndex).SubSecList(0).lVirtualAddress = .SecList(.MaxSecIndex).lVirtualAddress
		.SecList(.MaxSecIndex).SubSecList(0).lVirtualSize = .SecList(.MaxSecIndex).lVirtualSize
		.SecList(.MaxSecIndex).SubSecs = 1
	End With

	'修改隐藏节的偏移地址，还原隐藏节大小
	'由于重新读取的文件可能已在文件尾部写入了字串，这些字节将作为隐藏节被读取，故需要使用原始文件的隐藏节信息
	With File.SecList(File.MaxSecIndex + 1)
		.lPointerToRawData = File.SecList(File.MaxSecIndex).lPointerToRawData + File.SecList(File.MaxSecIndex).lSizeOfRawData
		.lVirtualAddress = File.SecList(File.MaxSecIndex).lVirtualAddress + File.SecList(File.MaxSecIndex).lVirtualSize
		.lSizeOfRawData = trnFile.SecList(trnFile.MaxSecIndex).lSizeOfRawData
		.lVirtualSize = trnFile.SecList(trnFile.MaxSecIndex).lVirtualSize
	End With

	'修改目标文件的区段数据
	If fType < 1 Then
		AddSecSize.Address = File.SecList(File.MaxSecIndex).lPointerToRawData
		AddSecSize.inSectionID = File.MaxSecIndex
		AddSecSize.inSubSecID = 0
		AddSecSize.Length = File.SecList(File.MaxSecIndex).lSizeOfRawData
		AddSecSize.MaxAddress = File.SecList(File.MaxSecIndex).lPointerToRawData + File.SecList(File.MaxSecIndex).lSizeOfRawData - 1
		AddSecSize.lNumber = -File.SecList(File.MaxSecIndex).lPointerToRawData
		AddSecSize.MoveType = -4	'新增节空位，不参与空位管理
		AddMacSection = File.SecList(File.MaxSecIndex).lSizeOfRawData
		File.MaxSecID = File.MaxSecIndex: File.MaxSecIndex = File.MaxSecIndex + 1: trnFile = File
		Exit Function
	End	If

	'打开文件
	Mode = LoadFile(File.FilePath,FN,0,1,0,Mode)
	If Mode < -1 Then GoTo localError

	'获取新增节的索引号
	With File
		Dim NewLoadCmd As mac_load_command
		NewLoadCmd.lcmdsize = NewSizeOfHeader
		'增加区段数
		Select Case .Magic
		Case "MAC32"
			NewLoadCmd.lcmd = SEGMENT
			MacHeader32.lncmds = MacHeader32.lncmds + 1
			MacHeader32.lsizeofcmds = MacHeader32.lsizeofcmds + NewSizeOfHeader

			Dim NewCMD32 As segment_command_32
			Dim NewSection32 As command_section_32
			'设置区段的属性
			For i = 1 To Len(SecName)
				NewCMD32.segname(i - 1) = AscW(Mid$(UCase$(SecName),i,1))
			Next i
			NewCMD32.lnsects = 1
			NewCMD32.lfileoff = .SecList(.MaxSecIndex).lPointerToRawData
			NewCMD32.lfilesize = .SecList(.MaxSecIndex).lSizeOfRawData
			NewCMD32.lvmaddr = .SecList(.MaxSecIndex).lVirtualAddress
			NewCMD32.lvmsize = .SecList(.MaxSecIndex).lVirtualSize

			'设置子节的属性
			For i = 0 To UBound(NewCMD32.segname)
				NewSection32.segname(i) = NewCMD32.segname(i)
			Next i
			For i = 1 To Len("__cstring")
				NewSection32.sectname(i - 1) = AscW(Mid$("__cstring",i,1))
			Next i
			NewSection32.lalign = 2 	'(需2的倍数)节对齐值
			NewSection32.loffset = .SecList(.MaxSecIndex).lPointerToRawData
			NewSection32.laddr = .SecList(.MaxSecIndex).lVirtualAddress
			NewSection32.lsize = .SecList(.MaxSecIndex).lVirtualSize
			NewSection32.lflags = S_ATTR_PURE_INSTRUCTIONS Or S_ATTR_SOME_INSTRUCTIONS
			'计算插入地址，将插入到 MacLoadCmd 的中间，位于 MacCmd32 最大段的后面
			With MacCmd32(.MaxSecID)
				If .CMD.lnsects > 0 Then
					NewHeadersOffset = MacLoadCmd(.Index).lOffset + Len(MacLoadCmd(0).LoadCmd) + _
										Len(.CMD) + Len(.Section(0)) * .CMD.lnsects
				Else
					NewHeadersOffset = MacLoadCmd(.Index).lOffset + Len(MacLoadCmd(0).LoadCmd) + Len(.CMD)
				End If
			End With
		Case "MAC64"
			NewLoadCmd.lcmd = SEGMENT_64
			MacHeader64.lncmds = MacHeader64.lncmds + 1
			MacHeader64.lsizeofcmds = MacHeader64.lsizeofcmds + NewSizeOfHeader

			Dim NewCMD64 As segment_command_64
			Dim NewSection64 As command_section_64
			'设置区段的属性
			For i = 1 To Len(SecName)
				NewCMD64.segname(i - 1) = AscW(Mid$(UCase$(SecName),i,1))
			Next i
			NewCMD64.lnsects = 1
			NewCMD64.dfileoff1 = .SecList(.MaxSecIndex).lPointerToRawData
			NewCMD64.dfileoff2 = MacCmd64(.MaxSecID).CMD.dfileoff2
			NewCMD64.dfilesize1 = .SecList(.MaxSecIndex).lSizeOfRawData
			NewCMD64.dfilesize2 = 0
			NewCMD64.dvmaddr1 = .SecList(.MaxSecIndex).lVirtualAddress
			NewCMD64.dvmaddr2 = MacCmd64(.MaxSecID).CMD.dvmaddr2
			NewCMD64.dvmsize1 = .SecList(.MaxSecIndex).lVirtualSize
			NewCMD64.dvmsize2 = 0

			'设置子节的属性
			For i = 0 To UBound(NewCMD64.segname)
				NewSection64.segname(i) = NewCMD64.segname(i)
			Next i
			For i = 1 To Len("__cstring")
				NewSection64.sectname(i - 1) = AscW(Mid$("__cstring",i,1))
			Next i
			NewSection64.lalign = 2 	'(需2的倍数)节对齐值
			NewSection64.loffset = .SecList(.MaxSecIndex).lPointerToRawData
			NewSection64.daddr1 = .SecList(.MaxSecIndex).lVirtualAddress
			NewSection64.daddr2 = MacCmd64(.MaxSecID).CMD.dvmaddr2
			NewSection64.dsize1 = .SecList(.MaxSecIndex).lVirtualSize
			NewSection64.dsize2 = 0
			NewSection64.lflags = S_ATTR_PURE_INSTRUCTIONS Or S_ATTR_SOME_INSTRUCTIONS
			'计算插入地址，将插入到 MacLoadCmd 的中间，位于 MacCmd64 最大段的后面
			With MacCmd64(File.MaxSecID)
				If .CMD.lnsects > 0 Then
					NewHeadersOffset = MacLoadCmd(.Index).lOffset + Len(MacLoadCmd(0).LoadCmd) + _
										Len(.CMD) + Len(.Section(0)) * .CMD.lnsects
				Else
					NewHeadersOffset = MacLoadCmd(.Index).lOffset + Len(MacLoadCmd(0).LoadCmd) + Len(.CMD)
				End If
			End With
		End Select
	End With

	'移位原文件的隐藏节及隐藏节后的全部字节长度
	'由于重新读取的文件可能已在文件尾部写入了字串，这些字节将作为隐藏节被读取，故需要使用原始文件的隐藏节信息
	'必须使用虚拟大小，因为其大小为最大节后的所有字节(包括子 MAC)
	If trnFile.SecList(trnFile.MaxSecIndex).lVirtualSize > 0 Then
		i = trnFile.FileSize - trnFile.SecList(trnFile.MaxSecIndex).lPointerToRawData
		TempByte = GetBytes(FN,i,trnFile.SecList(trnFile.MaxSecIndex).lPointerToRawData,Mode)
		PutBytes(FN,File.SecList(File.MaxSecIndex + 1).lPointerToRawData,TempByte,i,Mode)
		'置空原最大节对齐后多出的字节和新增节为空字节
		i = File.SecList(File.MaxSecIndex).lPointerToRawData + File.SecList(File.MaxSecIndex).lSizeOfRawData - _
			trnFile.SecList(trnFile.MaxSecIndex).lPointerToRawData
		ReDim TempByte(i - 1) As Byte
		PutBytes(FN,trnFile.SecList(trnFile.MaxSecIndex).lPointerToRawData,TempByte,i,Mode)
	End If

	'移动 k 位置以后的所有文件头一个新文件头位置，只在在文件头区域内
	'NewHeadersOffset = k			'排在 MacLoadCmd 最大节后
	If NewHeadersOffset < k Then	'插入到 MacLoadCmd 的中间
		i = k - NewHeadersOffset
		TempByte = GetBytes(FN,i,NewHeadersOffset,Mode)
		PutBytes(FN,NewHeadersOffset + NewSizeOfHeader,TempByte,i,Mode)
		'清空 NewHeadersOffset 后 NewSizeOfHeaders 个数据
		ReDim TempByte(NewSizeOfHeader - 1) As Byte
		PutBytes(FN,NewHeadersOffset,TempByte,NewSizeOfHeader,Mode)
	End If

	'写入新区段数据
	'If PutTypeArray(FN,NewHeadersOffset,NewLoadCmd,Mode) = False Then GoTo localError
	Select Case Mode
	Case Is < 0
		Put #FN.hFile,NewHeadersOffset + 1,NewLoadCmd
	Case 0
		CopyMemory FN.ImageByte(NewHeadersOffset),NewLoadCmd,Len(NewLoadCmd)
	Case Else
		WriteMemory FN.MappedAddress + NewHeadersOffset,NewLoadCmd,Len(NewLoadCmd)
	End Select
	Select Case File.Magic
	Case "MAC32"
		Select Case Mode
		Case Is < 0
			Put #FN.hFile,File.FileType + 1,MacHeader32
			Put #FN.hFile,NewHeadersOffset + Len(NewLoadCmd) + 1,NewCMD32
			Put #FN.hFile,NewHeadersOffset + Len(NewLoadCmd) + Len(NewCMD32) + 1,NewSection32
		Case 0
			CopyMemory FN.ImageByte(File.FileType),MacHeader32,Len(MacHeader32)
			CopyMemory FN.ImageByte(NewHeadersOffset + Len(NewLoadCmd)),NewCMD32,Len(NewCMD32)
			CopyMemory FN.ImageByte(NewHeadersOffset + Len(NewLoadCmd) + Len(NewCMD32)),NewSection32,Len(NewSection32)
		Case Else
			WriteMemory FN.MappedAddress + File.FileType,MacHeader32,Len(MacHeader32)
			WriteMemory FN.MappedAddress + NewHeadersOffset + Len(NewLoadCmd),NewCMD32,Len(NewCMD32)
			WriteMemory FN.MappedAddress + NewHeadersOffset + Len(NewLoadCmd) + Len(NewCMD32),NewSection32,Len(NewSection32)
		End Select
	Case "MAC64"
		Select Case Mode
		Case Is < 0
			Put #FN.hFile,File.FileType + 1,MacHeader64
			Put #FN.hFile,NewHeadersOffset + Len(NewLoadCmd) + 1,NewCMD64
			Put #FN.hFile,NewHeadersOffset + Len(NewLoadCmd) + Len(NewCMD64) + 1,NewSection64
		Case 0
			CopyMemory FN.ImageByte(File.FileType),MacHeader64,Len(MacHeader64)
			CopyMemory FN.ImageByte(NewHeadersOffset + Len(NewLoadCmd)),NewCMD64,Len(NewCMD64)
			CopyMemory FN.ImageByte(NewHeadersOffset + Len(NewLoadCmd) + Len(NewCMD64)),NewSection64,Len(NewSection64)
		Case Else
			WriteMemory FN.MappedAddress + File.FileType,MacHeader64,Len(MacHeader64)
			WriteMemory FN.MappedAddress + NewHeadersOffset + Len(NewLoadCmd),NewCMD64,Len(NewCMD64)
			WriteMemory FN.MappedAddress + NewHeadersOffset + Len(NewLoadCmd) + Len(NewCMD64),NewSection64,Len(NewSection64)
		End Select
	End Select

	'在隐藏节前增加 AddSecSize.Length 空字节
	'由于重新读取的文件可能已在文件尾部写入了字串，这些字节将作为隐藏节被读取，故需要使用原始文件的隐藏节信息
	'必须使用虚拟大小，因为其大小为最大节后的所有字节(包括子 PE)
	If trnFile.SecList(trnFile.MaxSecIndex).lVirtualSize < 1 Then
		i = GetFileLength(FN,Mode)
		k = File.SecList(File.MaxSecIndex).lPointerToRawData + File.SecList(File.MaxSecIndex).lSizeOfRawData - i
		If k > 0 Then
			ReDim TempByte(k - 1) As Byte
			PutBytes(FN,i,TempByte,k,Mode)
		End If
	End If
	UnLoadFile(FN,FN.SizeOfFile,Mode)
	AddMacSection = File.SecList(File.MaxSecIndex).lSizeOfRawData

	'修改目标文件的区段数据
	If fType = 2 Then
		File.MaxSecID = File.MaxSecIndex: File.MaxSecIndex = File.MaxSecIndex + 1: trnFile = File
	End If
	Exit Function

	'安全退出函数
	localError:
	UnLoadFile(FN,0,Mode)
	AddMacSection = -1
End Function
