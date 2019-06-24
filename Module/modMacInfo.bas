Attribute VB_Name = "modMacInfo"

'' File Information Module for PSlHardCodedString.bas
'' (c) 2015-2019 by wanfu (Last modified on 2019.03.15)

'#Uses "modCommon.bas"

Option Explicit

Private Const MAC_BIT_TYPE32 = 48 + 68
Private Const MAC_BIT_TYPE64 = 72 + 80

Private Type mac_header_32  '28���ֽ�
	lmagic				As Long		'mach magic number identifier
	lcputype			As Long		'cpu specifier (int)
	lcpusubtype			As Long		'machine specifier (int)
	lfiletype			As Long		'type of file
	lncmds				As Long		'ָ���ж��ٸ�Command
	lsizeofcmds			As Long		'ָ��LoadCommand�ܵĴ�С
	lflags				As Long		'file offset to this object file
End Type

Private Type mac_header_64  '32���ֽ�
	lmagic				As Long		'mach magic number identifier
	lcputype			As Long		'cpu specifier (int)
	lcpusubtype			As Long		'machine specifier (int)
	lfiletype			As Long		'type of file
	lncmds				As Long		'ָ���ж��ٸ�Command
	lsizeofcmds			As Long		'ָ��LoadCommand�ܵĴ�С
	lflags				As Long		'file offset to this object file
	lreserved			As Long
End Type
'magic: ���Կ����ļ��е������ʼ���֣����� cafe babe��ͷ��
'       ����һ�� �������ļ� ������ÿ�����Ͷ��������ļ���������ֽ�����ʶ����������ħ��������ͬ���͵� �������ļ��������Լ����ص�"ħ��"��
'       OS X�ϣ���ִ���ļ��ı�ʶ����������ħ������ͬ��ħ������ͬ�Ŀ�ִ���ļ����ͣ�
'       ��mach-o�ļ���ħ����0xfeedface�������32λ��0xfeedfacf����64λ��cafebabe�ǿ紦�����ܹ���ͨ�ø�ʽ��#!������ǽű��ļ���
'cputype �� cupsubtype: �������cpu�����ͺ��������ͣ�ͼ�ϵ�������ģ��������cpu�ṹ��x86_64,���ֱ�Ӳ鿴ipa�����Կ���cpu��arm��subtype��armv7��arm64��
'#define CPU_TYPE_ARM((cpu_type_t) 12)
'#define CPU_SUBTYPE_ARM_V7((cpu_subtype_t) 9
'filetype: &H2 �����ִ�е��ļ�
'ncmds: ָ���Ǽ�������(load commands)��������������һ��65�������0-64
'sizeofcmds: ��ʾ23��load commands�����ֽڴ�С��load commands�����ǽ�����header����ġ�
'flags: ��������0��00200085�����԰��ĵ�����֮��

'�������ӿ��ļ�
Private Type mac_header_fat_arch	'20���ֽ�
	lcputype			As Long		'CPU specifier
	lcpusubtype			As Long		'Machine specifier
	lfileoffset			As Long		'Offset of header in file
	lsize				As Long		'size of object file
	lalign				As Long		'Alignment As a power of two
End Type

'�������ӿ��ļ�
Private Type mac_header_fat  '32���ֽ�
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
	MH_NOUNDEFS					= &H1			'Ŀǰû��δ����ķ��ţ���������������
	MH_INCRLINK					= &H2
	MH_DYLDLINK					= &H4			'���ļ���dyld�������ļ����޷����ٴξ�̬����
	MH_BINDATLOAD				= &H8
	MH_PREBOUND					= &H10
	MH_SPLIT_SEGS				= &H20
	MH_LAZY_INIT				= &H40
	MH_TWOLEVEL					= &H80			'�þ����ļ�ʹ��2�����ƿռ�
	MH_FORCE_FLAT				= &H100
	MH_NOMULTIDEFS				= &H200
	MH_NOFIXPREBINDING			= &H400
	MH_PREBINDABLE				= &H800
	MH_ALLMODSBOUND				= &H1000
	MH_SUBSECTIONS_VIA_SYMBOLS	= &H2000
	MH_CANONICAL				= &H4000
	MH_WEAK_DEFINES				= &H8000
	MH_BINDS_TO_WEAK			= &H10000		'������ӵľ����ļ�ʹ��������
	MH_ALLOW_STACK_EXECUTION	= &H20000
	MH_ROOT_SAFE				= &H40000
	MH_SETUID_SAFE				= &H80000
	MH_NO_REEXPORTED_DYLIBS		= &H100000
	MH_PIE						= &H200000		'���س���������ĵ�ַ�ռ䣬ֻ�� MH_EXECUTE��ʹ��
	MH_DEAD_STRIPPABLE_DYLIB	= &H400000
	MH_HAS_TLV_DESCRIPTORS		= &H800000
	MH_NO_HEAP_EXECUTION		= &H1000000
End Enum

'load command �ṹ
'Command�кܶ಻ͬ�����࣬ÿ�������Ӧһ���ṹ�嵫�����е�Command��������ͬ�Ŀ�ʼ�ṹ
'ע�������С�ǰ����������������ݣ���������ṹ�屾����ռ�Ĵ�С��������������Section�ṹ�Ĵ�С��
'�����е�Padding�����0.(���ǲ�����������Data,������Dataһ���� FileOffset ��ָ��,���ݲ�ͬCommand�᲻ͬ)
'���Դ����ʼ�����ϵڶ�����Ա�Ĵ�С���Ϳ���ֱ�Ӷ�λ����һ������Ŀ�ʼ����
'���˾����������൱�Ĵ죬������Ϊ������Ϊ����Ҫ�ȶ�һ��Load_Command�ṹ����֪����ǰ�����Ǹ�ʲô���ͣ�
'Ȼ����ȥ����Ӧ�Ľṹ�������Ժ󣬻�Ҫ�ص����ʼ�����ټ��ϵڶ�����Ա�Ĵ�Сȥ������һ������Ƚϴ죡
'���磬���cmd=19��������һ��Segment_Command_64,Ҳ���Ǵ����￪ʼ��ʵ��һ��Segment_Command_64�ṹ
Private Type mac_load_command	'8���ֽ�
	lcmd				As Long		'Command ������
	lcmdsize			As Long		'Command �Ĵ�С
End Type
'load commmandֱ�Ӹ��� header ���ֵĺ���

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
'һ�����͵� OS X ��ִ���ļ�ͨ����������Σ�:
'__PAGEZERO : ��λ�������ַ0�����κα���Ȩ�����˶����ļ��в�ռ�ÿռ䣬����Null������������.
'__TEXT : ����ֻ�����ݺͿ�ִ�д���.
'__DATA : ������д����. ��Щ sectionͨ�����ں˱�־Ϊcopy-On-Write .
'__OBJC : ����Objective C ��������ʱ����ʹ�õ����ݡ�
'__LINKEDIT :������̬�������õ�ԭʼ����.
'__TEXT�� __DATA�ο��ܰ���0������section. ÿ��section��ָ�����͵�����, ��, ��ִ�д���, ����, C �ַ��������

'Segment_Command �ṹ
'����Ľṹ�����˶������ͳ�ʼ�����ڴ汣�����룬���������ַ���ļ�ƫ�ƣ���Windows�ϵ����ݲ�ࡣ
'��Ҫ�������棬nsects��flags��������һ��ָ��������˶���sections����һ������ǰ�Ķ����ԡ�
'���nsects>0����������нڣ����ҽڵĶ�������Ķζ��塣
Private Type segment_command_32	'40���ֽ�
	segname(15) 		As Byte			'segment name  16 ���ַ�
	lvmaddr				As Long			'memory address of this segment �ε������ڴ��ַ
	lvmsize				As Long			'memory size of this segment VM Address �ε������ڴ��С
	lfileoff			As Long			'file offset of this segment �����ļ���ƫ����
	lfilesize			As Long			'amount to map from the file �����ļ��еĴ�С
	lmaxprot			As Long			'maximum VM protection
	linitprot			As Long			'initial VM protection
	lnsects				As Long			'number of sections in segment
	lflags				As Long			'flags
End Type

Private Type segment_command_64	'64���ֽ�
	segname(15) 		As Byte			'segment name  16 ���ַ�
	dvmaddr1			As Long			'memory address of this segment �ε������ڴ��ַ��һ����
	dvmaddr2			As Long			'memory address of this segment �ε������ڴ��ַ�ڶ�����
	dvmsize1			As Long			'memory size of this segment VM Address �ε������ڴ��С��һ����
	dvmsize2			As Long			'memory size of this segment VM Address �ε������ڴ��С�ڶ�����
	dfileoff1			As Long			'file offset of this segment �����ļ���ƫ������һ����
	dfileoff2			As Long			'file offset of this segment �����ļ���ƫ�����ڶ�����
	dfilesize1			As Long			'amount to map from the file �����ļ��еĴ�С��һ����
	dfilesize2			As Long			'amount to map from the file �����ļ��еĴ�С�ڶ�����
	lmaxprot			As Long			'maximum VM protection
	linitprot			As Long			'initial VM protection
	lnsects				As Long			'number of sections in segment
	lflags				As Long			'flags
End Type
'���öζ�Ӧ���ļ����ݼ��ص��ڴ��У���offset������ file
'size��С�������ڴ� vmaddr ���������������ڴ��ַ�ռ�����_PAGEZERO�Σ�����β����з���Ȩ�ޣ����������ָ�룩���Զ�����
'���������Σ�����_TEXT��Ӧ�ľ��Ǵ���Σ�_DATA��Ӧ���ǿɶ�����д�����ݣ�_LINKEDIT��֧��dyld�ģ��������һЩ���ű������
'�����и����������⣬����ͼ��ʾ����д��__TEXT������� Segment��Сд��__text���� Section

'Constants for the segment_command_flags field
Private Enum segment_command_flags
	HIGH_VM					= &H1
	FVM_LIB					= &H2
	NO_RELOC				= &H4
	PROTECTION_VERSION_1	= &H8
End Enum

'�ڽṹ
'������һ������һЩ�ڴ�ƫ�ƺ��ļ�ƫ�ƣ������ض�λ�ڵ����ã���ϸ����Ҫ�Ժ�������⡣
'��Ҫ��Ҳ��flags��ָ���˵�ǰ�ڵ����ԡ����н����Կ����к���������
'Ҫ�۲�����ݣ����������ݣ�ת�����fileOffset���ٶ����ݡ�
Private Type command_section_32	'68���ֽ�
	sectname(15)		As Byte			'name of this section  16 ���ַ���������
	segname(15)			As Byte			'segment this section goes in  16 ���ַ������ڶε�����
	laddr				As Long			'memory address of this section �����ַ
	lsize				As Long			'size in bytes of this section �����С
	loffset				As Long			'file offset of this section �ļ�ƫ����
	lalign				As Long			'section alignment (power of 2)�ڶ���ֵ
	lreloff				As Long			'file offset of relocation entries
	lnreloc				As Long			'number of relocation entries
	lflags				As Long			'flags (section type and attributes)
	lreserved1			As Long			'reserved (for offset or index)
	lreserved2			As Long			'reserved (for count or sizeof)
End Type

Private Type command_section_64	'80�ֽ�
	sectname(15)		As Byte			'name of this section  16 ���ַ���������
	segname(15)			As Byte			'segment this section goes in  16 ���ַ������ڶε�����
	daddr1				As Long			'memory address of this section �����ַ��һ����
	daddr2				As Long			'memory address of this section �����ַ�ڶ�����
	dsize1				As Long			'size in bytes of this section �����С��һ����
	dsize2				As Long			'size in bytes of this section �����С�ڶ�����
	loffset				As Long			'file offset of this section �ļ�ƫ����
	lalign				As Long			'section alignment (power of 2)�ڶ���ֵ
	lreloff				As Long			'file offset of relocation entries
	lnreloc				As Long			'number of relocation entries
	lflags				As Long			'flags (section type and attributes)
	lreserved1			As Long			'reserved (for offset or index)
	lreserved2			As Long			'reserved (for count or sizeof)
	lreserved3			As Long
End Type
'section cmd ˵��
'__text ���������
'__stubs ���ڶ�̬�����ӵ�׮
'__stub_helper ���ڶ�̬�����ӵ�׮
'__cstring �����ַ������ű�������Ϣ��ͨ��������Ϣ�����Ի�ó����ַ������ű��ַ
'__unwind_info �����ֶβ���̫���ɶ��˼��ϣ�����ָ����

'�� __TEXT����, �����ĸ���Ҫ�� section:
'__text ���������
'__const : ͨ�ó�������.
'__cstring : �������ַ�������.
'__picsymbol_stub : ��̬������ʹ�õ�λ���޹��� stub ·��.
'���������˿�ִ�еĺͲ���ִ�еĴ����ڶ�������Ը���.

'������, Ĭ�ϵĴ���ھ������2������
Private Enum command_section
	S_ATTR_PURE_INSTRUCTIONS = &H80000000
	S_ATTR_SOME_INSTRUCTIONS = &H00000400
End Enum

'LOAD COMMAND ����
Private Type MAC_FILE_LOAD_COMMAND
	lOffset			As Long
	LoadCmd			As mac_load_command
End Type

'COMMAND_32 ����
Private Type MAC_FILE_COMMAND_32
	Index			As Integer
	CMD				As segment_command_32
	Section() 		As command_section_32
End Type

'COMMAND_64 ����
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


'��ȡ�ļ������ļ������ݽṹ��Ϣ
Public Function GetMacHeaders(ByVal strFilePath As String,File As FILE_PROPERTIE,ByVal Mode As Long) As Boolean
	Dim i As Long,FN As FILE_IMAGE,TempList() As String,Temp As String
	On Error GoTo ExitFunction
	File.FileSize = FileLen(strFilePath)
	'���ļ�
	Mode = LoadFile(strFilePath,FN,0,0,0,Mode)
	If Mode < -1 Then Exit Function
	'��ȡ���ļ�ͷ
	GetMacHeaders = GetMacHeader(FN,File,Mode)
	If GetMacHeaders = False Then GoTo ExitFunction
	'��ȡ���ļ�ͷ
	With File.SecList(File.MaxSecIndex)
		If .lSizeOfRawData = 0 Then GoTo ExitFunction
		Temp = ByteToString(GetBytes(FN,.lSizeOfRawData,.lPointerToRawData,Mode),CP_ISOLATIN1)
		TempList = GetVAListRegExp(Temp,"(\xCE\xFA\xED\xFE)|(\xCF\xFA\xED\xFE)",.lPointerToRawData)
		If CheckArray(TempList) = False Then GoTo ExitFunction
		Dim SubFile As FILE_PROPERTIE
		File.NumberOfSub = UBound(TempList) + 1
		For i = 0 To File.NumberOfSub - 1
			'If GetMacHeader(FN,SubFile,Mode,CLng(TempList(i))) = True Then
				'�޸����ļ������ؽڴ�С
				.lSizeOfRawData = CLng(TempList(i)) - .lPointerToRawData
				Exit For
			'End If
		Next i
	End With
	ExitFunction:
	'�ر��ļ�
	On Error Resume Next
	UnLoadFile(FN,0,Mode)
End Function


'��ȡ�ļ����ݽṹ��Ϣ
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
		'��ʼ��
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

		'��ȡ FAT Header
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

		'��ȡ Header
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

		'��ȡ Command ��
		ReDim tmpMacLoadCmd(.MaxSecIndex - 1)	'As MAC_FILE_LOAD_COMMAND
		ReDim tmpMacCmd32(.MaxSecIndex - 1) 	'As MAC_FILE_COMMAND_32
		ReDim tmpMacCmd64(.MaxSecIndex - 1)	'As MAC_FILE_COMMAND_64
		ReDim File.SecList(.MaxSecIndex)	'As SECTION_PROPERTIE
		For i = 0 To .MaxSecIndex - 1
			tmpMacLoadCmd(i).loffset = k
			If k + tmpMacLoadCmd(i).LoadCmd.lcmdsize <= .FileSize Then
				'��ȡ Load Command ��
				Select Case Mode
				Case Is < 0
					Get #FN.hFile, k + 1, tmpMacLoadCmd(i).LoadCmd
				Case 0
					CopyMemory tmpMacLoadCmd(i).LoadCmd, FN.ImageByte(k), Len(tmpMacLoadCmd(i).LoadCmd)
				Case Else
					MoveMemory tmpMacLoadCmd(i).LoadCmd, FN.MappedAddress + k, Len(tmpMacLoadCmd(i).LoadCmd)
				End Select

				'��ȡ Command ��
				Select Case tmpMacLoadCmd(i).LoadCmd.lcmd
				Case SEGMENT	'32λ��׼ Command
					'��ȡ Command ����
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

					'��ȡ������
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
						'��¼�ڵ�ַ
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
					'��¼���ε�ַ
					.SecList(n).sName = Replace$(StrConv$(tmpMacCmd32(n).CMD.segname,vbUnicode),vbNullChar,"")
					.SecList(n).lPointerToRawData = tmpMacCmd32(n).CMD.lfileoff
					.SecList(n).lSizeOfRawData = tmpMacCmd32(n).CMD.lfilesize
					.SecList(n).lVirtualAddress = tmpMacCmd32(n).CMD.lvmaddr
					.SecList(n).lVirtualSize = tmpMacCmd32(n).CMD.lvmsize
					.SecList(n).SubSecs = tmpMacCmd32(n).CMD.lnsects
					If .SecList(n).lSizeOfRawData > 0 Then n = n + 1
				Case SEGMENT_64	'64λ��׼ Command
					'��ȡ Command ��������
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

					'��ȡ������
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
						'��¼�ڵ�ַ
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
					'��¼���ε�ַ
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
			'��һ���ζ��Ǵ�0��ʼ���������ļ�ͷ������Ҫ����
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

		'��ȡ�ļ�����������š���С�����ƫ�Ƶ�ַ���ڽڵ�������
		Call GetSectionID(File,.MinSecID,.MaxSecID,False)

		'��ȡ���ؽ���Ϣ
		.SecList(.MaxSecIndex).lPointerToRawData = .SecList(.MaxSecID).lPointerToRawData + .SecList(.MaxSecID).lSizeOfRawData
		.SecList(.MaxSecIndex).lSizeOfRawData = GetFileLength(FN,Mode) - .SecList(.MaxSecIndex).lPointerToRawData
		.SecList(.MaxSecIndex).lVirtualAddress = .SecList(.MaxSecID).lVirtualAddress + .SecList(.MaxSecID).lVirtualSize
		.SecList(.MaxSecIndex).lVirtualSize = .SecList(.MaxSecIndex).lSizeOfRawData
	End With

	'��¼������ĸ���ͷ����
	If Offset = -1 Then
		MacHeader32 = tmpMacHeader32
		MacHeader64 = tmpMacHeader64
		MacLoadCmd = tmpMacLoadCmd
		MacCmd32 = tmpMacCmd32
		MacCmd64 = tmpMacCmd64
	End If

	'��ǳɹ�
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
		'���������ļ�Ϊһ����
		.SecList(0).lPointerToRawData = 0
		.SecList(0).lSizeOfRawData = GetFileLength(FN,Mode)
		.SecList(0).lVirtualAddress = 0
		.SecList(0).lVirtualSize = .SecList(0).lSizeOfRawData
		'�������ؽ���Ϣ��������ʾ�ļ���Ϣ
		.SecList(1).lPointerToRawData = .SecList(0).lSizeOfRawData
		.SecList(1).lSizeOfRawData = 0
		.SecList(1).lVirtualAddress = .0
		.SecList(1).lVirtualSize = 0
	End With
End Function


'��������ַתƫ�Ƶ�ַ
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


'ƫ�Ƶ�ַת��������ַ
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


'��ȡ�ε�����
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


'�����ļ��ڳ���
'fType = 0 ��ȡ�����ӵ��ֽ������޸�ԭʼ�ļ�ͷ��д��
'fType = 1 д��ָ�����ȶ������ֽ��������޸�ԭʼ�ļ�ͷ��д��
'fType = 2 д��ָ�����ȶ������ֽ������޸�ԭʼ�ļ�ͷ
'fType = 3 ���޸Ĳ�д�룬����ȡ������ֵ(AddSecSize(x).Length Ϊƫ�ƴ�С��AddSecSize(x).Address Ϊ�����С)
'AddSecSize(x).Length = 0���������ӵ����ֵ���ӣ����� AddSecSize(x).Length ����ֵ����
Public Function AddMacSectionSize(trnFile As FILE_PROPERTIE,AddSecSize() As FREE_BTYE_SPACE,ByVal fType As Long,ByVal Mode As Long) As Long
	Dim i As Long,j As Integer,k As Long,x As Long,n As Long
	Dim SecAlign As Long,AddRAW As Long,AddRVA As Long
	Dim FN As FILE_IMAGE,TempByte() As Byte,File As FILE_PROPERTIE

	'������
	On Error GoTo localError

	'��ȡ PE ͷ
	File = trnFile
	If GetMacHeaders(File.FilePath,File,Mode) = False Then Exit Function

	'�޸��ļ��ڵĿ�ʼ��ַ�ʹ�С
	With File
		For i = 0 To UBound(AddSecSize)
			j = AddSecSize(i).inSectionID
			'�����Ƿ������ڻ�ȡ�������ֽ�
			If j = File.MaxSecID Then
				k = .SecList(j).lSizeOfRawData + AddSecSize(i).Length
				n = 1
			Else
				k = .SecList(GetSectionID(File,j,-1,True)).lVirtualAddress - .SecList(j).lVirtualAddress
				n = 0
			End If
			'���ļ�����ֵ����
			x = k - .SecList(j).lSizeOfRawData
			If x > 0 Or n > 0 Then
				'����ʵ����Ҫ���ӵ�ǰ�ڵ�ƫ�ƴ�С������
				If AddSecSize(i).Length > 0 Then
					x = IIf(x > AddSecSize(i).Length,AddSecSize(i).Length,x)
				End If
				AddSecSize(i).Length = x: AddRAW = AddRAW + x

				'���ӵ�ǰ�ڵ������С�������С���ö���
				If AddSecSize(i).Length > 0 Then
					x = k - .SecList(j).lVirtualSize
					If x > 0 Then
						If x > AddSecSize(i).Length Then x = AddSecSize(i).Length
						.SecList(j).lVirtualSize = .SecList(j).lVirtualSize + x
						If fType > 2 Then AddSecSize(i).Address = x
						AddRVA = AddRVA + x
					End If
				End If

				'��¼������ֵ
				If fType = 0 Then
					'���㲢��¼��λ��ַ�������ִ���λ�������
					AddSecSize(i).Address = .SecList(j).lPointerToRawData + .SecList(j).lSizeOfRawData
					'AddSecSize(i).inSectionID = j
					If .SecList(j).SubSecs = 0 Then
						AddSecSize(i).inSubSecID = 0
					Else
						AddSecSize(i).inSubSecID = .SecList(j).SubSecs - 1
					End If
					AddSecSize(i).MaxAddress = AddSecSize(i).Address + AddSecSize(i).Length - 1
					AddSecSize(i).lNumber = -AddSecSize(i).Address
					AddSecSize(i).MoveType = -3	'������β��λ���������λ����
				ElseIf fType = 1 Or fType = 2 Then
					'����ԭ�ļ���ǰ�ڵ�ƫ�Ƶ�ַ�ʹ�С�����ں����ʵ����λ����
					AddSecSize(i).Address = trnFile.SecList(j).lPointerToRawData
					AddSecSize(i).MaxAddress = AddSecSize(i).Address + trnFile.SecList(j).lSizeOfRawData - 1
				End If

				'�޸ĵ�ǰ�ڵ�ƫ�ƴ�С������ڵ�ƫ�Ƶ�ַ
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

	'�޸����ؽڵ�ƫ�Ƶ�ַ����ԭ���ؽڴ�С
	'�������¶�ȡ���ļ����������ļ�β��д�����ִ�����Щ�ֽڽ���Ϊ���ؽڱ���ȡ������Ҫʹ��ԭʼ�ļ������ؽ���Ϣ
	With File.SecList(File.MaxSecIndex)
		.lPointerToRawData = File.SecList(File.MaxSecID).lPointerToRawData + File.SecList(File.MaxSecID).lSizeOfRawData
		.lVirtualAddress = File.SecList(File.MaxSecID).lVirtualAddress + File.SecList(File.MaxSecID).lVirtualSize
		.lSizeOfRawData = trnFile.SecList(trnFile.MaxSecIndex).lSizeOfRawData
		.lVirtualSize = trnFile.SecList(trnFile.MaxSecIndex).lVirtualSize
	End With

	'�޸�Ŀ���ļ�����������
	If fType = 0 Then
		trnFile = File
		AddMacSectionSize = AddRAW
		Exit Function
	End If

	'���ļ�
	Mode = LoadFile(File.FilePath,FN,0,1,0,Mode)
	If Mode < -1 Then GoTo localError

	'�޸� Header ����
	With File
		For i = AddSecSize(0).inSectionID To .MaxSecIndex - 1
			If .Magic = "MAC32" Then
				j = MacCmd32(i).Index
			Else
				j = MacCmd64(i).Index
			End If
			'��ȡ Command ������
			Select Case MacLoadCmd(j).LoadCmd.lcmd
			Case SEGMENT	'32λ��׼ Command
				'�޸� MacCmdHeader ����
				If i > 0 Then MacCmd32(i).CMD.lfileoff = .SecList(i).lPointerToRawData
				MacCmd32(i).CMD.lfilesize = .SecList(i).lSizeOfRawData
				MacCmd32(i).CMD.lvmsize = .SecList(i).lVirtualSize
				'�޸��ӽڵĵ�ַ�ʹ�С
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
			Case SEGMENT_64	'64λ��׼ Command
				'�޸ĸ��ڵĵ�ַ�ʹ�С
				If i > 0 Then MacCmd64(i).CMD.dfileoff1 = .SecList(i).lPointerToRawData
				MacCmd64(i).CMD.dfilesize1 = .SecList(i).lSizeOfRawData
				MacCmd64(i).CMD.dvmsize1 = .SecList(i).lVirtualSize
				'�޸��ӽڵĵ�ַ�ʹ�С
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

	'��С�����ƫ�Ƶ�ַ���������Ա�����ԭ�ڱ��������չ����
	Call SortFreeByteByAddress(AddSecSize,0,UBound(AddSecSize),False)

	'��λ���Ҫ���Ӵ�С���ڽں���ÿ���ڣ����ڽ�β���ӿ�Ҫ���ӵĿ��ֽ�
	n = UBound(AddSecSize)
	For i = n To 0 Step -1
		With AddSecSize(i)
			j = AddSecSize(i).inSectionID
			'��ȡ��ǰ�ں��ȫ���ֽڳ���
			If i = n Then
				If j = File.MaxSecID Then
					'���ڻ�ȡ�ļ�ͷʱ�����»�ȡ�ļ���С������չ����ʱ������չǰ��д���ִ���
					'��Ҫʹ��ԭ�����ļ���С��������Щд���ִ��ᱻ���Ƶ����
					k = trnFile.FileSize - .MaxAddress - 1
				Else
					k = File.FileSize - .MaxAddress - 1
				End If
			Else
				k = AddSecSize(i + 1).MaxAddress - .MaxAddress + .Length
			End If
			'��λ�ֽڵ���ǰ�ڵ�����ַ����
			If k > 0 Then
				TempByte = GetBytes(FN,k,.MaxAddress + 1,Mode)
				PutBytes(FN,File.SecList(j).lPointerToRawData + File.SecList(j).lSizeOfRawData,TempByte,k,Mode)
			End If

			'���ӵ�ǰ�ں�Ŀ��ֽ�(�ÿ�)
			'�������¶�ȡ���ļ����������ļ�β��д�����ִ�����Щ�ֽڽ���Ϊ���ؽڱ���ȡ������Ҫʹ��ԭʼ�ļ������ؽ���Ϣ
			If .Length > 0 Then
				'����ʹ�������С����Ϊ���СΪ���ں�������ֽ�(������ PE)
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

	'�޸�Ŀ���ļ�����������
	If fType = 2 Then trnFile = File
	Exit Function

	'��ȫ�˳�����
	localError:
	UnLoadFile(FN,0,Mode)
	AddMacSectionSize = -1
End Function


'���ļ�β������һ���ļ���
'fType = 0 ֻ�޸��ļ������ݲ�д��
'fType = 1 ���޸��ļ������ݵ�д��
'fType = 2 ���޸��ļ���������д��
Public Function AddMacSection(trnFile As FILE_PROPERTIE,AddSecSize As FREE_BTYE_SPACE,ByVal SecName As String,ByVal fType As Long,ByVal Mode As Long) As Long
	Dim i As Long,k As Long,n As Long
	Dim NewHeadersOffset As Long,NewSizeOfHeader As Long
	Dim FN As FILE_IMAGE,TempByte() As Byte,File As FILE_PROPERTIE

	'������
	On Error GoTo localError
	If AddSecSize.Length = 0 Then Exit Function

	'��ȡ PE ͷ
	File = trnFile
	If GetMacHeaders(File.FilePath,File,Mode) = False Then Exit Function

	'�޸��ļ�����ֵ���Լ��������ڵĴ�С
	File.FileAlign = 512
	File.SecAlign = 512

	'�������Ƿ��ж��ٿ���ռ���������¶ε�����
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

	'�¶�����
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

	'�޸����ؽڵ�ƫ�Ƶ�ַ����ԭ���ؽڴ�С
	'�������¶�ȡ���ļ����������ļ�β��д�����ִ�����Щ�ֽڽ���Ϊ���ؽڱ���ȡ������Ҫʹ��ԭʼ�ļ������ؽ���Ϣ
	With File.SecList(File.MaxSecIndex + 1)
		.lPointerToRawData = File.SecList(File.MaxSecIndex).lPointerToRawData + File.SecList(File.MaxSecIndex).lSizeOfRawData
		.lVirtualAddress = File.SecList(File.MaxSecIndex).lVirtualAddress + File.SecList(File.MaxSecIndex).lVirtualSize
		.lSizeOfRawData = trnFile.SecList(trnFile.MaxSecIndex).lSizeOfRawData
		.lVirtualSize = trnFile.SecList(trnFile.MaxSecIndex).lVirtualSize
	End With

	'�޸�Ŀ���ļ�����������
	If fType < 1 Then
		AddSecSize.Address = File.SecList(File.MaxSecIndex).lPointerToRawData
		AddSecSize.inSectionID = File.MaxSecIndex
		AddSecSize.inSubSecID = 0
		AddSecSize.Length = File.SecList(File.MaxSecIndex).lSizeOfRawData
		AddSecSize.MaxAddress = File.SecList(File.MaxSecIndex).lPointerToRawData + File.SecList(File.MaxSecIndex).lSizeOfRawData - 1
		AddSecSize.lNumber = -File.SecList(File.MaxSecIndex).lPointerToRawData
		AddSecSize.MoveType = -4	'�����ڿ�λ���������λ����
		AddMacSection = File.SecList(File.MaxSecIndex).lSizeOfRawData
		File.MaxSecID = File.MaxSecIndex: File.MaxSecIndex = File.MaxSecIndex + 1: trnFile = File
		Exit Function
	End	If

	'���ļ�
	Mode = LoadFile(File.FilePath,FN,0,1,0,Mode)
	If Mode < -1 Then GoTo localError

	'��ȡ�����ڵ�������
	With File
		Dim NewLoadCmd As mac_load_command
		NewLoadCmd.lcmdsize = NewSizeOfHeader
		'����������
		Select Case .Magic
		Case "MAC32"
			NewLoadCmd.lcmd = SEGMENT
			MacHeader32.lncmds = MacHeader32.lncmds + 1
			MacHeader32.lsizeofcmds = MacHeader32.lsizeofcmds + NewSizeOfHeader

			Dim NewCMD32 As segment_command_32
			Dim NewSection32 As command_section_32
			'�������ε�����
			For i = 1 To Len(SecName)
				NewCMD32.segname(i - 1) = AscW(Mid$(UCase$(SecName),i,1))
			Next i
			NewCMD32.lnsects = 1
			NewCMD32.lfileoff = .SecList(.MaxSecIndex).lPointerToRawData
			NewCMD32.lfilesize = .SecList(.MaxSecIndex).lSizeOfRawData
			NewCMD32.lvmaddr = .SecList(.MaxSecIndex).lVirtualAddress
			NewCMD32.lvmsize = .SecList(.MaxSecIndex).lVirtualSize

			'�����ӽڵ�����
			For i = 0 To UBound(NewCMD32.segname)
				NewSection32.segname(i) = NewCMD32.segname(i)
			Next i
			For i = 1 To Len("__cstring")
				NewSection32.sectname(i - 1) = AscW(Mid$("__cstring",i,1))
			Next i
			NewSection32.lalign = 2 	'(��2�ı���)�ڶ���ֵ
			NewSection32.loffset = .SecList(.MaxSecIndex).lPointerToRawData
			NewSection32.laddr = .SecList(.MaxSecIndex).lVirtualAddress
			NewSection32.lsize = .SecList(.MaxSecIndex).lVirtualSize
			NewSection32.lflags = S_ATTR_PURE_INSTRUCTIONS Or S_ATTR_SOME_INSTRUCTIONS
			'��������ַ�������뵽 MacLoadCmd ���м䣬λ�� MacCmd32 ���εĺ���
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
			'�������ε�����
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

			'�����ӽڵ�����
			For i = 0 To UBound(NewCMD64.segname)
				NewSection64.segname(i) = NewCMD64.segname(i)
			Next i
			For i = 1 To Len("__cstring")
				NewSection64.sectname(i - 1) = AscW(Mid$("__cstring",i,1))
			Next i
			NewSection64.lalign = 2 	'(��2�ı���)�ڶ���ֵ
			NewSection64.loffset = .SecList(.MaxSecIndex).lPointerToRawData
			NewSection64.daddr1 = .SecList(.MaxSecIndex).lVirtualAddress
			NewSection64.daddr2 = MacCmd64(.MaxSecID).CMD.dvmaddr2
			NewSection64.dsize1 = .SecList(.MaxSecIndex).lVirtualSize
			NewSection64.dsize2 = 0
			NewSection64.lflags = S_ATTR_PURE_INSTRUCTIONS Or S_ATTR_SOME_INSTRUCTIONS
			'��������ַ�������뵽 MacLoadCmd ���м䣬λ�� MacCmd64 ���εĺ���
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

	'��λԭ�ļ������ؽڼ����ؽں��ȫ���ֽڳ���
	'�������¶�ȡ���ļ����������ļ�β��д�����ִ�����Щ�ֽڽ���Ϊ���ؽڱ���ȡ������Ҫʹ��ԭʼ�ļ������ؽ���Ϣ
	'����ʹ�������С����Ϊ���СΪ���ں�������ֽ�(������ MAC)
	If trnFile.SecList(trnFile.MaxSecIndex).lVirtualSize > 0 Then
		i = trnFile.FileSize - trnFile.SecList(trnFile.MaxSecIndex).lPointerToRawData
		TempByte = GetBytes(FN,i,trnFile.SecList(trnFile.MaxSecIndex).lPointerToRawData,Mode)
		PutBytes(FN,File.SecList(File.MaxSecIndex + 1).lPointerToRawData,TempByte,i,Mode)
		'�ÿ�ԭ���ڶ���������ֽں�������Ϊ���ֽ�
		i = File.SecList(File.MaxSecIndex).lPointerToRawData + File.SecList(File.MaxSecIndex).lSizeOfRawData - _
			trnFile.SecList(trnFile.MaxSecIndex).lPointerToRawData
		ReDim TempByte(i - 1) As Byte
		PutBytes(FN,trnFile.SecList(trnFile.MaxSecIndex).lPointerToRawData,TempByte,i,Mode)
	End If

	'�ƶ� k λ���Ժ�������ļ�ͷһ�����ļ�ͷλ�ã�ֻ�����ļ�ͷ������
	'NewHeadersOffset = k			'���� MacLoadCmd ���ں�
	If NewHeadersOffset < k Then	'���뵽 MacLoadCmd ���м�
		i = k - NewHeadersOffset
		TempByte = GetBytes(FN,i,NewHeadersOffset,Mode)
		PutBytes(FN,NewHeadersOffset + NewSizeOfHeader,TempByte,i,Mode)
		'��� NewHeadersOffset �� NewSizeOfHeaders ������
		ReDim TempByte(NewSizeOfHeader - 1) As Byte
		PutBytes(FN,NewHeadersOffset,TempByte,NewSizeOfHeader,Mode)
	End If

	'д������������
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

	'�����ؽ�ǰ���� AddSecSize.Length ���ֽ�
	'�������¶�ȡ���ļ����������ļ�β��д�����ִ�����Щ�ֽڽ���Ϊ���ؽڱ���ȡ������Ҫʹ��ԭʼ�ļ������ؽ���Ϣ
	'����ʹ�������С����Ϊ���СΪ���ں�������ֽ�(������ PE)
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

	'�޸�Ŀ���ļ�����������
	If fType = 2 Then
		File.MaxSecID = File.MaxSecIndex: File.MaxSecIndex = File.MaxSecIndex + 1: trnFile = File
	End If
	Exit Function

	'��ȫ�˳�����
	localError:
	UnLoadFile(FN,0,Mode)
	AddMacSection = -1
End Function
