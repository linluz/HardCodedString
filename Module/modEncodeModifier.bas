'' Character Encode Modifier for Passolo
'' (c) 2018 - 2019 by wanfu (Last modified on 2019.11.08)

'' Command Line Format: Command <FilePath>
'' Command: Name of this Macros file
'' FilePath: Full path of PE file.
'' Return: No
'' For example: modEncodeModifier,"d:\my folder\my file.exe"

Option Explicit

Private Const Version = "2019.11.08"
Private Const Build = "191108"
Private Const JoinStr = vbFormFeed  'vbBack
Private Const TextJoinStr = vbCrLf
Private Const LoadMode = 0&
Private Const AppName = "Character Encode Modifier"
Private Const RefFrontChar64 = "[\x0F\x4C\x48\x8B\xF2][\x05\x10\x0F\x8B\x8D\xB7][\x00-\xFF]{3}"
Private Const ImportFunRefFrontChar32 = "\xFF\x15"
Private Const ImportFunRefFrontChar64 = "\xFF\x15[\x00-\xFF]{2}"
Private Const CodePageFrontChar32 = "[\x68\x6A]"
Private Const CodePageFrontChar64 = "[\xB9\xBA\xBE]"

'程序编写语言或平台定义
Private Enum PELangType
	DELPHI_FILE_SIGNATURE = &H50
	NET_FILE_SIGNATURE = &H424A5342
End Enum

'PE文件结构(Visual Basic版)部分代码一
'签名定义
Private Enum ImageSignatureTypes
	IMAGE_DOS_SIGNATURE = &H5A4D			'// MZ
	IMAGE_OS2_SIGNATURE = &H454E			'// NE
	IMAGE_OS2_SIGNATURE_LE = &H454C			'// LE
	IMAGE_VXD_SIGNATURE = &H454C			'// LE
	IMAGE_NT_SIGNATURE = &H4550				'// PE00
End Enum

'判断是32位还是64位PE文件
Private Enum ImageOptionalHeaderMagicType
	IMAGE_NT_OPTIONAL_HDR32_MAGIC = &H10B			'32位PE文件
	IMAGE_NT_OPTIONAL_HDR64_MAGIC = &H20B			'64位PE文件
End Enum

'文件类型(标志集合)定义
'IMAGE_FILE.Characteristics
'Private Enum ImageFileCharacteristicsTypes
'	IMAGE_FILE_RELOCS_STRIPPED = &H1				'重定位信息被移除
'	IMAGE_FILE_EXECUTABLE_IMAGE = &H2				'文件可执行
'	IMAGE_FILE_LINE_NUMS_STRIPPED = &H4				'行号被移除
'	IMAGE_FILE_LOCAL_SYMS_STRIPPED = &H8			'符号被移除
'	IMAGE_FILE_AGGRESIVE_WS_TRIM = &H10				'Agressively Trim working Set
'	IMAGE_FILE_LARGE_ADDRESS_AWARE = &H20			'程序能处理大于2G的地址
'	IMAGE_FILE_BYTES_REVERSED_LO = &H80				'保留的机器类型低位
'	IMAGE_FILE_32BIT_MACHINE = &H100				'32位机器
'	IMAGE_FILE_DEBUG_STRIPPED = &H200				'.dbg文件的调试信息被移除
'	IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP = &H400		'如果在移动介质中,拷到交换文件中运行
'	IMAGE_FILE_NET_RUN_FROM_SWAP = &H800			'如果在网络中,拷到交换文件中运行
'	IMAGE_FILE_SYSTEM = &H1000						'系统文件
'	IMAGE_FILE_DLL = &H2000							'文件是一个dll
'	IMAGE_FILE_UP_SYSTEM_ONLY = &H4000				'文件只能运行在单处理器上
'	IMAGE_FILE_BYTES_REVERSED_HI = &H8000			'保留的机器类型高位.
'End Enum

'应用程序执行的环境及平台代码定义
'IMAGE_FILE_HEADER.iMachine
'======================================================================================
'Private Enum ImageFileMachineTypes
'	IMAGE_FILE_MACHINE_UNKNOWN = &H0		'未知
'	IMAGE_FILE_MACHINE_I386 = &H14C			'Intel 80386 处理器以上
'	IMAGE_FILE_MACHINE_I486 = &H14D			'Intel 80486 处理器以上
'	IMAGE_FILE_MACHINE_IPTM = &H14E			'Intel Pentium 处理器以上
'	IMAGE_FILE_MACHINE_R 	= &H160			'R3000(MIPS)处理器，big endian
'	IMAGE_FILE_MACHINE_R3000 = &H162		'R3000(MIPS)处理器，little endian
'	IMAGE_FILE_MACHINE_R4000 = &H166		'R4000(MIPS)处理器，little endian
'	IMAGE_FILE_MACHINE_R10000 = &H168		'R10000(MIPS)处理器，little endian
'	IMAGE_FILE_MACHINE_WCEMIPSV2 = &H169	'MIPS Little-endian WCE v2
'	IMAGE_FILE_MACHINE_ALPHA = &H184		'DEC Alpha AXP处理器
'	IMAGE_FILE_MACHINE_POWERPC= &H1F0		'IBM Power PC，little endian
'	IMAGE_FILE_MACHINE_SH3 = &H1A2			'SH3 little-endian
'	IMAGE_FILE_MACHINE_SH3E = &H1A4			'SH3E little-endian
'	IMAGE_FILE_MACHINE_SH4 = &H1A6			'SH4 little-endian
'	IMAGE_FILE_MACHINE_SH5 = &H1A8			'SH5
'	IMAGE_FILE_MACHINE_ARM = &H1C0			'ARM Little-Endian
'	IMAGE_FILE_MACHINE_THUMB = &H1C2
'	IMAGE_FILE_MACHINE_ARM33 = &H1D3
'	IMAGE_FILE_MACHINE_IA64 = &H200			'Intel 64
'	IMAGE_FILE_MACHINE_MIPS16 = &H266		'MIPS
'	IMAGE_FILE_MACHINE_ALPHA64 = &H284		'ALPHA64
'	IMAGE_FILE_MACHINE_MIPSFPU = &H366		'MIPS
'	IMAGE_FILE_MACHINE_MIPSFPU16 = &H466	'MIPS
'	IMAGE_FILE_MACHINE_AMD64 = &H500		'AMD K8
'	IMAGE_FILE_MACHINE_TRICORE = &H520		'Infineon
'	IMAGE_FILE_MACHINE_CEF = &HCEF
'	IMAGE_FILE_MACHINE_AMD64 = &H8664  		'AMD64 (K8)
'End Enum

'数据目录表
'======================================================================================
'Private Enum ImageDirectoryEntry
'	IMAGE_DIRECTORY_ENTRY_EXPORT = 0				'导出目录
'	IMAGE_DIRECTORY_ENTRY_IMPORT = 1				'导入目录
'	IMAGE_DIRECTORY_ENTRY_RESOURCE = 2				'资源目录
'	IMAGE_DIRECTORY_ENTRY_EXCEPTION = 3				'异常目录
'	IMAGE_DIRECTORY_ENTRY_SECURITY = 4				'安全目录
'	IMAGE_DIRECTORY_ENTRY_BASERELOC = 5				'重定位基本表
'	IMAGE_DIRECTORY_ENTRY_DEBUG = 6					'调试目录
'	IMAGE_DIRECTORY_ENTRY_COPYRIGHT = 7				'X86使用-描述文字
'	IMAGE_DIRECTORY_ENTRY_ARCHITECTURE = 7			'Architecture Specific Data
'	IMAGE_DIRECTORY_ENTRY_GLOBALPTR = 8				'机器值(MIPS GP),即 RVA of GlobalPtr
'	IMAGE_DIRECTORY_ENTRY_TLS = 9					'线程本地存储(Thread Local Storage,TLS)目录
'	IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG = 10			'载入配置目录
'	IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT = 11			'绑定输入表(Bound Import Directory in headers)
'	IMAGE_DIRECTORY_ENTRY_IAT = 12					'导入地址表
'	IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT = 13			'Delay Load Import Descriptors
'	IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR = 14		'COM 运行标志
'	IMAGE_DIRECTORY_ENTRY_RESERVED = 15				'保留
'End Enum

'节属性定义
Private Enum ImageSectionType
	IMAGE_SCN_CNT_CODE = &H20						'节中包含代码
	IMAGE_SCN_CNT_INITIALIZED_DATA = &H40			'节中包含已初始化数据
	IMAGE_SCN_CNT_UNINITIALIZED_DATA = &H80			'节中包含未初始化数据
	IMAGE_SCN_MEM_DISCARDABLE = &H2000000			'是一个可丢弃的节，即节中的数据在进程开始后将被丢弃
	IMAGE_SCN_MEM_NOT_CACHED = &H4000000			'节中数据不经过缓存
	IMAGE_SCN_MEM_NOT_PAGED = &H8000000				'节中数据不被交换出内存
	IMAGE_SCN_MEM_SHARED = &H10000000				'节中数据可共享
	IMAGE_SCN_MEM_EXECUTE = &H20000000				'可执行节
	IMAGE_SCN_MEM_READ = &H40000000					'可读节
	IMAGE_SCN_MEM_WRITE = &H80000000				'可写节
End Enum

'资源定义
'======================================================================================
'Private Enum ImageResourceEntry
'	IMAGE_RESOURCE_CURSOR = 1						'光标						'
'	IMAGE_RESOURCE_BITMAP = 2						'位图
'	IMAGE_RESOURCE_ICON = 3							'图标
'	IMAGE_RESOURCE_MENU = 4							'菜单
'	IMAGE_RESOURCE_DIALOG = 5						'对话框
'	IMAGE_RESOURCE_STRING_TABLE = 6					'字符串表
'	IMAGE_RESOURCE_FONT_DIRECTORY = 7				'字体目录
'	IMAGE_RESOURCE_FONT	= 8							'字体
'	IMAGE_RESOURCE_ACCELERATORS = 9					'加速器
'	IMAGE_RESOURCE_UNFORMATTED_RESOURCE_DATA = 10	'未格式化资源数据
'	IMAGE_RESOURCE_MESSAGE_TABEL = 11				'消息表
'	IMAGE_RESOURCE_GROUP_CURSOR = 12				'光标组
'	IMAGE_RESOURCE_GROUP_ICON = 13					'图标组
'	IMAGE_RESOURCE_VERSION_INFO = 14				'版本信息
'End Enum

'======================================================================================
'结构名: IMAGE_DOS_HEADER
'结构大小: 64字节
'结构说明: DOS映像头(或EXE头)
'======================================================================================
Private Type IMAGE_DOS_HEADER
	iSignature						As Integer		'&H0    签名("MZ",即&H5A4D)
	iLastPageBytes					As Integer		'&H2    文件最后页中的字节数
	iPages							As Integer		'&H4    文件页数
	iRelocateItems					As Integer		'&H6    重定位元素个数
	iHeaderSize						As Integer		'&H8    头部大小
	iMinAlloc						As Integer		'&HA    分配的最小附加段
	iMaxAlloc						As Integer		'&HC    分配的最大附加段
	iInitialSS						As Integer		'&HE    初始SS值
	iInitialSP						As Integer		'&H10   初始SP值
	iCheckSum						As Integer		'&H12   校验和
	iInitialIP						As Integer		'&H14   初始IP值
	iInitialCS						As Integer		'&H16   初始CS值(相对偏移量)
	iRelocateTable					As Integer		'&H18   重定向表文件地址
	iOverlay						As Integer		'&H1A   覆盖号
	iReserved(3)					As Integer		'&H22   保留字
	iOEMID							As Integer		'&H24   OEM标识符
	iOEMInformation					As Integer		'&H26   OEM信息
	iReserved2(9)					As Integer		'&H28   保留字2
	lPointerToPEHeader				As Long			'&H3C   PE头部位置
End Type

'======================================================================================
'结构名: IMAGE_FILE_HEADER
'结构大小: 24字节
'结构说明: 映像文件头
'======================================================================================
Private Type IMAGE_FILE_HEADER
	lSignature						As Long			'&H4	PE文件头标志("PE00",即&H4550),4字节
	iMachine						As Integer		'&H6    执行该程序的环境及平台
	iNumberOfSections				As Integer		'&H8    文件中节的个数
	lTimeDateStamp					As Long			'&HC    文件建立时间(时间戳)
	lPointerToSymbolTable			As Long			'&H10   COFF符号表偏移
	lNumberOfSymbols				As Long			'&H14   符号数目
	iSizeOfOptionalHeader			As Integer		'&H16   可选头部大小
	iCharacteristics				As Integer		'&H18   标志集合
End Type

'======================================================================================
'结构名: IMAGE_DATA_DIRECTORY
'结构大小: 每个目录8字节，共16个目录，共120字节
'结构说明: 数据目录表
'0 = 导出目录
'1 = 导入目录
'2 = 资源目录
'3 = 异常目录
'4 = 安全目录
'5 = 基址重定位表
'6 = 调试目录
'7 = 版权目录
'8 = 机器值(GP RVA)
'9 = 线程本地存储表
'10 = 载入配置目录
'11 = 绑定导入目录
'12 = 导入地址表
'13 = 延迟加载导入符
'14 = COM 运行库标志(.NET 程序的 RVA = CLR 地址，Size = CLR header 的大小，固定为48字节)
'15 = 保留目录
'======================================================================================
Private Type IMAGE_DATA_DIRECTORY
	lVirtualAddress			As Long			'起始RVA地址
	lSize					As Long			'lVirtualAddress所指向数据结构的字节数
End Type

'======================================================================================
'结构名: IMAGE_OPTIONAL_HEADER32
'结构大小: 224字节
'结构说明: 可选映像头
'======================================================================================
Private Type IMAGE_OPTIONAL_HEADER32
	'******************
	'标准域
	'******************
	iMagic							As Integer		'&H18   32位PE是&H10B，64位PE是&H20B
	bMajorLinkerVersion				As Byte			'&H1A   链接器主版本
	bMinorLinkerVersion				As Byte			'&H1B   链接器次版本
	lSizeOfCode						As Long			'&H1C   可执行代码长度
	lSizeOfInitializedData			As Long			'&H20   初始化数据长度(数据节)
	lSizeOfUninitializedData		As Long			'&H24   未初始化数据长度(bss节)
	lAddressOfEntryPoint			As Long			'&H28   代码入口RVA地址,程序从这开始执行
	lBaseOfCode						As Long			'&H2C   可执行代码起始位置
	lBaseOfData						As Long			'&H30   初始化数据起始位置
	'******************
	'NT 附加域
	'******************
	lImageBase						As Long			'&H34   载入程序首选的RVA地址(32位)
	lSectionAlignment				As Long			'&H38   加载后节在内存中的对齐方式
	lFileAlignment					As Long			'&H3C   节在文件中的对齐方式
	iMajorOperatingSystemVersion	As Integer		'&H40   操作系统主版本
	iMinorOperatingSystemVersion	As Integer		'&H42   操作系统次版本
	iMajorImageVersion				As Integer		'&H44   可执行文件主版本
	iMinorImageVersion				As Integer		'&H46   可执行文件次版本
	iMajorSubsystemVersion			As Integer		'&H48   子系统主版本
	iMinorSubsystemVersion			As Integer		'&H50   子系统次版本
	lWin32VersionValue				As Long			'&H52   Win32版本号,一般为0
	lSizeOfImage					As Long			'&H56   程序调入后占用内存大小(虚拟大小)
	lSizeOfHeaders					As Long			'&H5A   头部大小(偏移大小)
	lCheckSum						As Long			'&H5E   校验和
	iSubsystem						As Integer		'&H62   可执行文件的子系统
	iDllCharacteristics				As Integer		'&H64   何时DllMain被调用,一般为0
	lSizeOfStackReserve				As Long			'&H66   初始化线程时保留堆栈大小
	lSizeOfStackCommit				As Long			'&H6A   初始化线程时提交堆栈大小
	lSizeOfHeapReserve				As Long			'&H6E   进程初始化时保留堆栈大小
	lSizeOfHeapCommit				As Long			'&H72   进程初始化时提交堆栈大小
	lLoaderFlags					As Long			'&H76   装载标志,与调试相关
	lNumberOfRvaAndSizes			As Long			'&H7A   数据目录的项数,一般为16
	DataDirectory(15) 				As IMAGE_DATA_DIRECTORY			'DataDirectory(14)
End Type

'======================================================================================
'结构名: IMAGE_OPTIONAL_HEADER64
'结构大小: 240字节
'结构说明: 可选映像头
'======================================================================================
Private Type IMAGE_OPTIONAL_HEADER64
	'******************
	'标准域
	'******************
	iMagic							As Integer		'&H18   32位PE是&H10B，64位PE是&H20B
	bMajorLinkerVersion				As Byte			'&H1A   链接器主版本
	bMinorLinkerVersion				As Byte			'&H1B   链接器次版本
	lSizeOfCode						As Long			'&H1C   可执行代码长度
	lSizeOfInitializedData			As Long			'&H20   初始化数据长度(数据节)
	lSizeOfUninitializedData		As Long			'&H24   未初始化数据长度(bss节)
	lAddressOfEntryPoint			As Long			'&H28   代码入口RVA地址,程序从这开始执行
	lBaseOfCode						As Long			'&H2C   可执行代码起始位置
	'lBaseOfData					As Long			'&H30   初始化数据起始位置
	'******************
	'NT 附加域
	'******************
	dImageBase(7)					As Byte			'&H34   载入程序首选的RVA地址(64位)
	lSectionAlignment				As Long			'&H38   加载后节在内存中的对齐方式
	lFileAlignment					As Long			'&H3C   节在文件中的对齐方式
	iMajorOperatingSystemVersion	As Integer		'&H40   操作系统主版本
	iMinorOperatingSystemVersion	As Integer		'&H42   操作系统次版本
	iMajorImageVersion				As Integer		'&H44   可执行文件主版本
	iMinorImageVersion				As Integer		'&H46   可执行文件次版本
	iMajorSubsystemVersion			As Integer		'&H48   子系统主版本
	iMinorSubsystemVersion			As Integer		'&H50   子系统次版本
	lWin32VersionValue				As Long			'&H52   Win32版本号,一般为0
	lSizeOfImage					As Long			'&H56   程序调入后占用内存大小(虚拟大小)
	lSizeOfHeaders					As Long			'&H5A   头部大小(偏移大小)
	lCheckSum						As Long			'&H5E   校验和
	iSubsystem						As Integer		'&H62   可执行文件的子系统
	iDllCharacteristics				As Integer		'&H64   何时DllMain被调用,一般为0
	dSizeOfStackReserve				As Double		'&H66   初始化线程时保留堆栈大小(64位)
	dSizeOfStackCommit				As Double		'&H6A   初始化线程时提交堆栈大小(64位)
	dSizeOfHeapReserve				As Double		'&H6E   进程初始化时保留堆栈大小(64位)
	dSizeOfHeapCommit				As Double		'&H72   进程初始化时提交堆栈大小(64位)
	lLoaderFlags					As Long			'&H76   装载标志,与调试相关
	lNumberOfRvaAndSizes			As Long			'&H7A   数据目录的项数,一般为16
	DataDirectory(15) 				As IMAGE_DATA_DIRECTORY
End Type

'======================================================================================
'结构名: IMAGE_SECTION_HEADER
'结构大小: 40字节
'结构说明: 节映像头
'======================================================================================
Private Type IMAGE_SECTION_HEADER
	sName(7)						As Byte			'&H0    节名(最大8个单字节字符)
	'lPhysicalAddress				As Long			'&H8    OBJ文件中表示本节的物理地址
	lVirtualSize					As Long			'&H8    EXE文件中表示节的实际字节数
	lVirtualAddress					As Long			'&HC    本节的RVA
	lSizeOfRawData					As Long			'&H10   本节经文件对齐后的尺寸
	lPointerToRawData				As Long			'&H14   本节原始数据在文件中的位置
	lPointerToRelocations			As Long			'&H18   OBJ文件中表示本节重定位信息的偏移,EXE文件中无意义
	lPointerToLineNumbers			As Long			'&H1C   行号偏移
	iNumberOfRelocations			As Integer		'&H20   本节需重定位的数目
	iNumberOfLineNumbers			As Integer		'&H22   本节在行号表中的行号数目
	lCharacteristics				As Long			'&H24   节属性
End Type

'======================================================================================
'结构名: .NET CLR 2.0 头结构
'结构大小: 72字节
'结构说明: CLR 头
'======================================================================================
Private Type IMAGE_CLR20_HEADER
	'Header versioning
	cb						As Long					'CLR头的大小，以byte为单位
	MajorRuntimeVersion		As Integer				'能运行该程序的最小.NET版本的主版本号
	MinorRuntimeVersion		As Integer				'能运行该程序的.NET版本的副版本号

	'Symbol table And startup information
	METADATA				As IMAGE_DATA_DIRECTORY	'元数据的RVA和Size
	Flags					As Long					'属性字段，可以在IL中以.corflags进行显式设置，
													'也可以在编译时用/FLAGS选项进行设置，其中命令行设置的优先级较高
	EntryPointToken			As Long					'入口方法的元数据ID（也就是token），在EXE文件必须有，
													'而DLL文件中此项可以为0（.NET 2.0中，此项还可以是本地入口代码的RVA值）
	'Binding information
	Resources				As IMAGE_DATA_DIRECTORY	'托管资源的RVA和Size
	StrongNameSignature		As IMAGE_DATA_DIRECTORY	'强名称数据的RVA和Size（强名称的意义在后面介绍）

	'Regular fixup And binding information
	CodeManagerTable		As IMAGE_DATA_DIRECTORY	'CodeManagerTable的RVA与Size，此项暂未使用，为0
    VTableFixups			As IMAGE_DATA_DIRECTORY	'v-table项的RVA和Size，主要供使用v-table的C++语言进行重定位
    ExportAddressTableJumps	As IMAGE_DATA_DIRECTORY	'用于C++的输出跳转地址表的RVA和Size，大多数情况为0

    'Precompiled image info (internal use only - Set To zero)
    ManagedNativeHeader		As IMAGE_DATA_DIRECTORY	'仅在由ngen生成本地模块中该项不为0，其余情况均为0
End Type

'数据目录表 CLR Header Flags
'======================================================================================
'Private Enum CLR_HEADER_FLAGES
'	COMIMAGE_FLAGS_ILONLY = &H1				'此CLR程序由纯IL代码组成
'	COMIMAGE_FLAGS_32BITREQUIRED = &H2		'此CLR映像只能在32位系统上执行
'	COMIMAGE_FLAGS_IL_LIBRARY = &H4			'此CLR映像是作为IL代码库存在的
'	COMIMAGE_FLAGS_STRONGNAMESIGNED = &H8	'文件受到强名称签名的保护
'	COMIMAGE_FLAGS_NATIVE_ENTRYPOINT =&H8	'此程序入口方法为非托管
'	COMIMAGE_FLAGS_TRACKDEBUGDATA = &H10000	'Loader和JIT需要追踪调试信息，缺省置0
'End Enum

'======================================================================================
'结构名: .NET MetaData 头结构
'结构大小: 不固定
'结构说明: MetaData 头
'======================================================================================
Private Type IMAGE_METADATA_HEADER
	lSignature		As Long   		'Magic signature For physical metadata, currently 0x424A5342(BSJB 为.Net 文件的标志)
	iMajorVersion	As Integer		'Major version (1 for the first release of the common language runtime)
	iMinorVersion	As Integer		'Minor Version (1 For the first release of the common language runtime)
	lExtraData		As Long			'Reserved, always 0
	lLength			As Long			'Length of Version String In bytes
	Version()		As Byte			'版本字符串, UTF8 编码，4字节对齐
	fFlags			As Integer		'Reserved, always 0
	iStreams		As Integer		'Number of streams
	'StreamHeader()	As IMAGE_STREAM_HEADER
End Type

'======================================================================================
'结构名: .NET Stream 头结构
'结构大小: 72字节
'结构说明: Stream 头
'常用流
'#Strings: UTF8格式的字符串堆，包含各种元数据的名称（比如类名、方法名、成员名、参数名等）。
'          流的首部总有一个0作为空字符串，各字符串以0表示结尾。CLR中这些名称的最大长度是1024。
'#Blob:    二进制数据堆，存储程序中的非字符串信息，比如常量值、方法的signature、PublicKey等。
'          每个数据的长度由该数据的前1～3位决定：0表示长度１字节，10表示长度2字节，110表示长度4字节。
'#GUID:    存储所有的全局唯一标识（Global Unique Identifier）。
'#US:      以Unicode格式存放的IL代码中使用的用户字符串（User String），比如ldstr调用的字符串。
'#~:       元数据表流，最重要的流，几乎所有的元数据信息都以表的形式保存于此。每个.Net 程序都必须包含。
'#-:       #~的未压缩（或称为未优化）存储，不常见。
'======================================================================================
Private Type IMAGE_STREAM_HEADER
	lOffset			As Long			'相对于 Metadata Root 的内存偏移
	lSize			As Long			'流的字节大小，4 的倍数
	rcName()		As Byte			'以空字节终止的 ASCII 字符串数组，4字节对齐
	RWA				As Long			'数据记录所在物理偏移地址 (实际没有这个值，只是为了方便定位各个流的位置)
End Type

'#Blob流
'#Blob流是一个二进制数据堆，程序中的所有非字符串形式数据都堆放在这个流里面，
'如常数的值，Public Key的值，方法的Signature等等。
'在每个二进制数据块头，都有一个块长度数据，但为了节约存储空间，CLR使用了比较麻烦的编码方法。
'如果开始一个字节最高位为0，则此数据块长度为一个字节；
'如果开始一个字节最高位为10，则此数据块长度为两个字节；
'如果开始一个字节最高位为110，则此数据块长度为四个字节；
'在屏蔽标志位后，通过移位运算即可计算出数据块的实际长度值，并依据此获得数据。

'#US流
'一个Blob堆，包括了用户自定义的字符串。
'这个流包括了定义在用户代码中的字符串常量。这些字符串以UTF-16的编码格式保存，附带着额外的一个尾部设置为0或1的字节，
'用以指出在字符串中是否有大于0x007F的代码字符。
'这个尾部字节被添加到流线上的在由用户定义的字符串常量生成的字符串对象上的代码转换操作。
'这个流的最有趣的特征是，这个用户字符串不仅会被任意元数据表引用到，还会显示地被IL代码表明地址（使用ldstr指令）。
'此外，作为一个实际上的blob堆，US堆不仅可以存储Unicode字符串，还可以存储任意二进制对象，这使那些有趣的实现成为可能。

'======================================================================================
'结构名: .NET #~ Stream 结构
'结构大小: 24字节
'结构说明: #~ Stream，保存了与强签名相关的数据
'======================================================================================
'Private Type TClrTableStreamHeader
'	Reserved		As Long			'保留，为0
'	MajorVersion	As Byte			'元数据表的主版本号，与.Net主版本号一致
'	MinorVersion	As Byte			'元数据表的副版本号，一般为0
'	HeapSizes		As Byte			'heaps 中定位数据时的索引的大小，为0表示16位索引值
									'若堆中数据超出16位数据表示范围，则使用32位索引值。
									'01h代表 strings 堆，02h代表 GUID 堆，04h代表 blob 堆
									'在#-流中可以为20h或80h，前者代表流中包含在Edit-and-Continue的调试中修改的数据，
									'后者表示元数据中个别项被标识为已删除
'	Rid				As Byte			'所有元数据表中记录的最大索引值，在运行时由.Net计算，文件中通常为1
'	MaskValid		As Double		'8字节长度的掩码，每个位代表一个表，为1表示该表有效，为0表示无该表
'	Sorted			As Double		'8字节长度的掩码，每个位代表一个表，为1表示该表已排序，反之为0
'End Type


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

'COMMAND_64 属性
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
Private Type FILE_PROPERTIE
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
	SubFileDir			As String	'子文件所在文件夹路径
	Info 				As String	'文件所有信息，避免重复获取

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
End Type

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

'错误消息定义
Private Enum FormatMSG
	FORMAT_MESSAGE_FROM_SYSTEM = &H1000
	FORMAT_MESSAGE_IGNORE_INSERTS = &H200
End Enum

'打开文件方式的结构体
Private Type FILE_IMAGE
	ModuleName				As String	'被加载文件的文件名
	hFile					As Long		'调用 Create 文件映射或 OpenFile 的句柄
	hMap					As Long		'调用 CreateFileMapping 文件映射的句柄
	MappedAddress			As Long		'文件映射到的内存地址
	SizeOfImage				As Long		'映射的 Image 或字节数组的大小
	SizeOfFile				As Long		'实际文件大小
	ImageByte()				As Byte		'文件的字节数组
End Type

'代码页转换
Private Declare Function MultiByteToWideChar Lib "kernel32.dll" ( _
	ByVal CodePage As Long, _
	ByVal dwFlags As Long, _
	ByRef lpMultiByteStr As Any, _
	ByVal cchMultiByte As Long, _
	ByVal lpWideCharStr As Long, _
	ByVal cchWideChar As Long) As Long

'内存复制和比较函数
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
	ByRef Destination As Any, _
	ByRef Source As Any, _
	ByVal Length As Long)
Private Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
	Dest As Any, _
	ByVal Source As Long, _
	ByVal Length As Long)
Private Declare Sub WriteMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
	ByVal Dest As Long, _
	ByRef Source As Any, _
	ByVal Length As Long)
'Private Declare Function vbVarPtr Lib "msvbvm60.dll" Alias "VarPtr" (Ptr As Any) As Long

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
Private Declare Function WriteFile Lib "kernel32.dll" ( _
	ByVal hFile As Long, _
	ByVal lpBuffer As Long, _
	ByVal nNumberOfBytesToWrite As Long, _
	lpNumberOfBytesWritten As Long, _
	ByVal lpOverlapped As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" ( _
	ByVal hFile As Long, _
	ByVal lpFileSizeHigh As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32.dll" (ByVal hFile As Long) As Long

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

'窗口文本
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
	ByVal hwnd As Long, _
	ByVal wMsg As Long, _
	ByVal wParam As Long, _
	ByRef lParam As Any) As Long
Private Declare Function SendMessageLNG Lib "user32.dll" Alias "SendMessage" ( _
	ByVal hwnd As Long, _
	ByVal wMsg As Long, _
	ByVal wParam As Long, _
	ByVal lParam As Long) As Long

'用于返回控件ID的句柄
Private Declare Function GetDlgItem Lib "user32.dll" ( _
	ByVal hDlg As Long, _
	ByVal nIDDlgItem As Long) As Long

'SendMessage API 部分常数
Private Enum SendMsgValue
	EM_GETLIMITTEXT = &HD5			'0,0				获取一个编辑控件中文本的最大长度
	EM_LIMITTEXT = &HC5				'最大值,0			设置编辑控件中的最大文本长度
	WM_GETTEXT = &H0D				'字节数,字符串地址	获取窗口文本控件的文本
	WM_GETTEXTLENGTH = &H0E			'0,0				获取窗口文本控件的文本的长度(字节数)
	WM_SETTEXT = &H0C				'0,字符串地址		设置窗口文本控件的文本
	WM_VSCROLL = &H115				'控件句柄,滚动条类型,滚动条位置	设置 SB_BOTTOM 指定的垂直滚动条位置
	SB_BOTTOM = &H07				'控件句柄,滚动条类型,滚动条位置	使用 WM_VSCROLL 设置 SB_BOTTOM 指定的垂直滚动条位置

	LB_GETCOUNT = &H018B			'0,0			返回列表项的总项数，若出错则返回LB_ERR
	LB_GETSELCOUNT = &H0190			'0,0			本操作仅用于多重选择列表框，它返回选择项的数目，若出错函数返回LB_ERR
	LB_GETSELITEMS = &H0191			'数组的大小,缓冲区	本操作仅用于多重选择列表框，用来获得选中的项的数目及位置。参数lParam指向一个整型数数组缓冲区，用来存放选中的列表项的索引。wParam说明了数组缓冲区的大小。本操作返回放在缓冲区中的选择项的实际数目，若出错函数返回LB_ERR
	LB_SETSEL = &H0185				'TRUE或FALSE,索引	仅适用于多重选择列表框，它使指定的列表项选中或落选，并自动滚动到可见区域。参数lParam指定了列表项的索引，若为-1，则相当于指定了所有的项。参数wParam为TRUE时选中列表项，否则使之落选。若出错则返回LB_ERR
	LB_SETTOPINDEX = &H0197			'索引,0			用来将指定的列表项设置为列表框的第一个可见项，该函数会将列表框滚动到合适的位置。wParam指定了列表项的索引．若操作成功，返回0值，否则返回LB_ERR
End Enum

'获取和设置滚动条位置函数
'Private Declare Function GetScrollPos Lib "user32.dll" ( _
'	ByVal hwnd As Long, _
'	ByVal nBar As Long) As Long
'Private Declare Function SetScrollPos Lib "user32.dll" ( _
'	ByVal hwnd As Long, _
'	ByVal nBar As Long, _
'	ByVal nPos As Long, _
'	ByVal bRedraw As Long) As Long

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

'======================================================================================
'结构名: IMAGE_IMPORT_DESCRIPTOR
'结构大小: 20字节
'结构说明: 输入表
'======================================================================================
Private Type IMAGE_IMPORT_DESCRIPTOR
	'Characteristics				As Long			'0 For terminating Null import descriptor
	OriginalFirstThunk				As Long			'RVA To original unbound IAT (PIMAGE_THUNK_DATA)
	TimeDateStamp					As Long			'0 if not bound
	ForwarderChain					As Long			'-1 if no forwarders
	pName							As Long
	FirstThunk						As Long			'RVA to IAT (if bound this IAT has actual addresses)
End Type

Private Type IMPORT_API_LOOKUP
	ApiName							As String
	Address							As Long
End Type

Private Type IMPORT_LIST_TYPE
	strName							As String
	strGuid							As String
	strLib							As String
End Type

Private Type IMAGE_IMPORT_BY_NAME
	Hint							As Integer		'函数输出序号
	pName(259)						As Byte			'输出函数名称
End Type

Private Type IMAGE_THUNK_DATA
	AddressOfData					As Long			'指向 IMAGE_IMPORT_BY_NAME 的指针
End Type

Private Type IMAGE_THUNK_DATA32
	pForwarderString				As Long			'RVA,指向输入函数名的指针
	pFunction						As Long			'输入函数的32位内存地址
	Ordinal							As Long			'判定当前结构数据是不是以序号为输出的，如果是的话该值为0x800000000
	AddressOfData					As Long			'指向 IMAGE_IMPORT_BY_NAME 的指针
End Type

Private Type IMAGE_THUNK_DATA64
	pForwarderString1				As Long			'RVA,指向输入函数名的指针
	pForwarderString2				As Long			'RVA,指向输入函数名的指针
	pFunction						As Double		'输入函数的64位内存地址
	Ordinal							As Double		'判定当前结构数据是不是以序号为输出的，如果是的话该值为0x800000000
	AddressOfData					As Double		'指向 IMAGE_IMPORT_BY_NAME 的指针
End Type

'重定向表
Private Type IMAGE_BASE_RELOCATION
	VirtualAddress					As Long			'重定位数据开始RVA地址
	SizeOfBlock						As Long			'重定位块得长度
	TypeOffset						As Integer		'重定项位数组
End Type

Private Type IMPORT_DETAIL_INFO
	strDllName					As String
	strFuns()					As String
End Type

Private Type IMPORT_INFO
	strExePath					As String
	pDetailInfo()				As IMPORT_DETAIL_INFO
End Type

'代码页
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

'代码页属性
Private Type CODEPAGE_PROPERTIE
	CodePage 		As Long
	LEPatrn 		As String
	BEPatrn 		As String
	ByteLength		As Integer
End Type

Private Type REFERENCE_PROPERTIE
	sOrgCode		As String	'原始引用代码
	sUseCode		As String	'更改引用代码
	lAddress		As Long		'引用地址
	inSecID			As Integer	'引用所在节的索引号
	lOrgCodePage	As Long		'原始代码页
	lUseCodePage	As Long		'更改代码页
	lCPAddress		As Long		'代码页地址
	iCPByteLength	As Integer	'代码页字节长度
	iCPByteOrder	As Integer	'字节序, -1 = 大端在前, 0 = 小端在前, 1 = 未知
	sUseFunc		As String	'更改函数名
	lUseFuncAddress	As Long		'更改函数名的地址
	UseFuncInSecID	As Integer	'更改函数名所在节的索引号
End Type

Private Type STRING_SUB_PROPERTIE
	DllName			As String	'Dll 文件名
	DllLang			As String	'Dll 文件属于的编译语言，如：C++，QT
	sString			As String	'函数名称
	lStartAddress	As Long		'函数的开始地址
	inSectionID		As Integer	'函数所在节的索引号
	inSubSecID		As Integer	'函数所在节的子节索引号
	lReferenceNum	As Long		'引用次数
	GetRefState		As Integer	'获取函数引用列表的状态，0 = 未获取，1 = 已获取
	lFunNameAddress	As Long		'函数名称的开始地址
	FuncVA			As Long		'函数的虚拟地址
	Reference()		As REFERENCE_PROPERTIE
End Type

'获取Dll函数的地址
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" ( _
	ByVal hModule As Long, _
	ByVal lpProcName As String) As Long
'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
'	ByVal lpPrevWndFunc As Long, _
'	ByVal hWnd As Long, _
'	ByRef Msg As Any, _
'	ByRef wParam As Any, _
'	ByRef lParam As Any) As Long

'用于读取文件版本信息函数
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" ( _
	ByVal lptstrFilename As String, _
	ByVal dwhandle As Long, _
	ByVal dwlen As Long, _
	lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" ( _
	ByVal lptstrFilename As String, _
	lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" ( _
	pBlock As Any, _
	ByVal lpSubBlock As String, _
	lplpBuffer As Any, _
	puLen As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" ( _
	ByVal lpString1 As String, _
	ByVal lpString2 As Long) As Long

Private DosHeader 			As IMAGE_DOS_HEADER
Private FileHeader			As IMAGE_FILE_HEADER
Private OptionalHeader32	As IMAGE_OPTIONAL_HEADER32
Private OptionalHeader64	As IMAGE_OPTIONAL_HEADER64
Private SecHeader() 		As IMAGE_SECTION_HEADER

Private MacHeader32 	As mac_header_32
Private MacHeader64 	As mac_header_64
Private MacLoadCmd()	As MAC_FILE_LOAD_COMMAND
Private MacCmd32()		As MAC_FILE_COMMAND_32
Private MacCmd64()		As MAC_FILE_COMMAND_64

Private MsgList() As String,RegExp As Object
Private RefAddList() As Long,AppCPList() As String,UseCPList() As String,FuncList() As String
Private File As FILE_PROPERTIE,DataList() As STRING_SUB_PROPERTIE,CodePageList() As CODEPAGE_PROPERTIE


'主程序
Sub Main()
	Dim Obj As Object,Temp As String,TempList() As String
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
	'检测 Scripting.Dictionary 是否存在
	Set Obj = CreateObject("Scripting.Dictionary")
	If Obj Is Nothing Then
		MsgBox Err.Description & " - " & "Scripting.Dictionary",vbInformation
		Exit Sub
	End If
	Set Obj = Nothing
	'检测 VBScript.RegExp 是否存在
	On Error Resume Next
	Set RegExp = CreateObject("VBScript.RegExp")
	If RegExp Is Nothing Then
		MsgBox(Err.Description & " - " & "VBScript.RegExp",vbInformation)
		Exit Sub
	End If
	RegExp.MultiLine = True
	On Error GoTo SysErrorMsg
	If UCase(Temp) <> Right$("0" & Hex$(PSL.Option(pslOptionSystemLanguage)),4) Then
		Temp = Right$("0" & Hex$(PSL.Option(pslOptionSystemLanguage)),4)
	End If
	If GetMsgList(MsgList,Temp) = False Then GoTo SysErrorMsg
	Call GetCPList(CodePageList)
	Begin Dialog UserDialog 590,371,Replace$(Replace$(MsgList(12),"%v",Version),"%b",Build),.MainDlgFunc ' %GRID:10,7,1,1
		TextBox 0,0,0,21,.SuppValueBox
		CheckBox 0,0,0,21,"",.WriteStateBox
		TextBox 0,0,0,21,.NumberBox
		TextBox 10,7,540,21,.FilePathBox
		PushButton 550,7,30,21,MsgList(13),.FilePathButton

		Text 10,38,80,14,MsgList(14),.FunNameText
		DropListBox 90,35,260,21,TempList(),.FunNameList
		Text 360,38,80,14,MsgList(15),.FunAddText
		TextBox 450,35,130,21,.FunAddBox

		Text 10,66,190,14,MsgList(16),.RefAddText
		Text 200,66,120,14,MsgList(17),.AppCPText
		Text 320,66,120,14,MsgList(18),.UseCPText
		MultiListBox 10,84,190,259,TempList(),.RefAddList
		MultiListBox 200,84,120,259,TempList(),.AppCPList
		MultiListBox 320,84,120,259,TempList(),.UseCPList

		Text 10,66,160,14,MsgList(16),.RefAddText1
		Text 170,66,270,14,MsgList(81),.FuncText
		MultiListBox 10,84,160,259,TempList(),.RefAddList1
		MultiListBox 170,84,270,256,TempList(),.FuncList,1
		Text 10,347,560,14,Replace$(MsgList(19),"%s",""),.StatusText

		PushButton 450,84,130,21,MsgList(20),.AboutButton
		PushButton 450,105,130,21,MsgList(21),.LangButton
		PushButton 450,133,130,21,MsgList(22),.SelectAllButton
		PushButton 450,154,130,21,MsgList(23),.EditCPButton
		PushButton 450,154,130,21,MsgList(82),.EditFuncButton
		PushButton 450,175,130,21,MsgList(24),.UndoButton
		PushButton 450,196,130,21,MsgList(86),.UndoAllButton
		PushButton 450,217,130,21,MsgList(33),.InfoButton
		PushButton 450,245,130,21,MsgList(26),.WriteButton
		PushButton 450,266,130,21,MsgList(27),.AllWriteButton
		PushButton 450,294,130,21,MsgList(28),.ShowSetButton
		PushButton 450,315,130,21,MsgList(29),.ShowAllButton
		PushButton 450,343,130,21,MsgList(30),.StopButton
		CancelButton 450,63,130,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then Exit Sub
	Exit Sub
	SysErrorMsg:
	If Err.Source <> "ExitSub" Then Call sysErrorMassage(Err,0)
End Sub


'请务必查看对话框帮助主题以了解更多信息。
Private Function MainDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,n As Long,Temp As String,IntList() As Long,TempList() As String
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		DlgText "SuppValueBox",CStr$(SuppValue)
		DlgVisible "SuppValueBox",False
		DlgVisible "CancelButton",False
		DlgVisible "StopButton",False
		DlgVisible "WriteStateBox",False
		DlgVisible "NumberBox",False
		DlgEnable "FilePathBox",False
		DlgEnable "FunAddBox",False

		'转递参数值
		File.FilePath = Command
		If Dir$(File.FilePath) = "" Then File.FilePath = ""
		Temp = File.FilePath
		If Len(Temp) > 60 Then
			Temp = Left$(Temp,InStr(Temp,"\")) & "..." & Right(Temp,60 - Len(Left$(Temp,InStr(Temp,"\"))))
		End If
		DlgText "FilePathBox",Temp
		ReDim DataList(0) As STRING_SUB_PROPERTIE
		DlgText "NumberBox","0"
		If File.FilePath <> "" And Dir$(File.FilePath) <> "" Then
			Call GetFileInfo(File.FilePath,File)
			If GetHeaders(File.FilePath,File,LoadMode,File.FileType) = True Then
				DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(64))
				If GetImportTable(File,DataList,LoadMode) = True Then
					'Call GetFuncVA(DataList)
					DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(65))
					DlgText "NumberBox",CStr$(GetFuncReference(File,DataList,LoadMode,GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("StatusText"))))
					DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(66))
					Call GetCodePage(File,DataList,CodePageList,LoadMode,GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("StatusText")))
					Call SetDataList(DataList)
					'DlgVisible "StopButton",False
				Else
					DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(67))
				End If
			Else
				DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(68))
			End If
		End If
		If DataList(0).sString = "" Then
			DlgText "FunAddBox",""
			ReDim RefAddList(0) As Long,AppCPList(0) As String,UseCPList(0) As String,FuncList(0) As String
			'DlgText "StatusText",Replace$(MsgList(19),"%s",Replace$(Replace$(MsgList(69),"%s",CStr$(0)),"%d",CStr$(0)))
			DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)),"%s",CStr$(0)),"%d",CStr$(0))
			DlgText "RefAddText1",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)),"%s",CStr$(0)),"%d",CStr$(0))
			DlgVisible "RefAddText1",False
			DlgVisible "FuncText",False
			DlgVisible "RefAddList1",False
			DlgVisible "FuncList",False
			DlgVisible "EditFuncButton",False
		Else
			DlgText "FunAddBox",ValToStr(DataList(0).lStartAddress,File.FileSize,True)
			TempList = GetDataValueList(DataList,0)
			DlgListBoxArray "FunNameList",TempList()
			DlgValue "FunNameList",0
			TempList = GetDataValueList(DataList,2,0,File.FileSize,True)
			Call StrListToLongList(TempList,RefAddList,True)
			Select Case DataList(0).DllLang
			Case "MS",""
				DlgListBoxArray "RefAddList",TempList()
				DlgValue "RefAddList",Array(0)
				AppCPList = GetDataValueList(DataList,3,0)
				DlgListBoxArray "AppCPList",AppCPList()
				DlgValue "AppCPList",Array(0)
				UseCPList = GetDataValueList(DataList,4,0)
				DlgListBoxArray "UseCPList",UseCPList()
				DlgValue "UseCPList",Array(0)
				DlgVisible "RefAddText1",False
				DlgVisible "FuncText",False
				DlgVisible "RefAddList1",False
				DlgVisible "FuncList",False
				DlgVisible "EditFuncButton",False
				DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)), _
								"%s",CStr$(1)),"%d",CStr$(DataList(0).lReferenceNum))
			Case "QT"
				DlgListBoxArray "RefAddList1",TempList()
				DlgValue "RefAddList1",Array(0)
				FuncList = GetDataValueList(DataList,26,0)
				DlgListBoxArray "FuncList",FuncList()
				DlgValue "FuncList",Array(0)
				DlgVisible "RefAddText",False
				DlgVisible "AppCPText",False
				DlgVisible "UseCPText",False
				DlgVisible "RefAddList",False
				DlgVisible "AppCPList",False
				DlgVisible "UseCPList",False
				DlgVisible "EditCPButton",False
				DlgText "RefAddText1",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)), _
								"%s",CStr$(1)),"%d",CStr$(DataList(0).lReferenceNum))
			End Select
			DlgText "StatusText",Replace$(MsgList(19),"%s",Replace$(Replace$(MsgList(69),"%s",CStr$(UBound(DataList) + 1)),"%d",DlgText("NumberBox")))
		End If
		If DataList(0).lReferenceNum = 0 Then
			DlgEnable "SelectAllButton",False
			DlgEnable "EditCPButton",False
			DlgEnable "EditFuncButton",False
			DlgEnable "UndoButton",False
			DlgEnable "UndoAllButton",False
			DlgEnable "InfoButton",False
			DlgEnable "WriteButton",False
			DlgEnable "AllWriteButton",False
			DlgEnable "ShowSetButton",False
		End If
		DlgEnable "ShowAllButton",False
	Case 2 ' 数值更改或者按下按钮时
		MainDlgFunc = True ' 防止按下按钮时关闭对话框窗口
		Select Case DlgItem$
		Case "CancelButton"
			MainDlgFunc = False
		Case "FilePathButton"
			If PSL.SelectFile(Temp,True,MsgList(60),MsgList(61)) = False Then Exit Function
			If File.FilePath = Temp Then Exit Function
			If IsOpen(Temp,2,0) = True Then Exit Function
			File.FilePath = Temp
			If Len(Temp) > 60 Then
				Temp = Left$(Temp,InStr(Temp,"\")) & "..." & Right(Temp,60 - Len(Left$(Temp,InStr(Temp,"\"))))
			End If
			DlgText "FilePathBox",Temp
			ReDim DataList(0) As STRING_SUB_PROPERTIE
			DlgText "NumberBox","0"
			File.Info = ""
			Call GetFileInfo(File.FilePath,File)
			If GetHeaders(File.FilePath,File,LoadMode,File.FileType) = True Then
				DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(64))
				If GetImportTable(File,DataList,LoadMode) = True Then
					'Call GetFuncVA(DataList)
					DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(65))
					DlgText "NumberBox",CStr$(GetFuncReference(File,DataList,LoadMode,GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("StatusText"))))
					DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(66))
					Call GetCodePage(File,DataList,CodePageList,LoadMode,GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("StatusText")))
					Call SetDataList(DataList)
					'DlgVisible "StopButton",False
				Else
					DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(67))
				End If
			Else
				DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(68))
			End If
			ReDim RefAddList(0) As Long,AppCPList(0) As String,UseCPList(0) As String,FuncList(0) As String
			If DataList(0).sString = "" Then
				DlgText "FunAddBox",""
				DlgListBoxArray "FunNameList",TempList()
				DlgValue "FunNameList",0
				DlgListBoxArray "RefAddList",TempList()
				DlgValue "RefAddList",Array(0)
				DlgListBoxArray "AppCPList",AppCPList()
				DlgValue "AppCPList",Array(0)
				DlgListBoxArray "UseCPList",UseCPList()
				DlgValue "UseCPList",Array(0)
				DlgListBoxArray "RefAddList1",TempList()
				DlgValue "RefAddList1",Array(0)
				DlgListBoxArray "FuncList",FuncList()
				DlgValue "FuncList",Array(0)
				DlgVisible "RefAddText",True
				DlgVisible "AppCPText",True
				DlgVisible "UseCPText",True
				DlgVisible "RefAddList",True
				DlgVisible "AppCPList",True
				DlgVisible "UseCPList",True
				DlgVisible "EditCPButton",True
				DlgVisible "RefAddText1",False
				DlgVisible "FuncText",False
				DlgVisible "RefAddList1",False
				DlgVisible "FuncList",False
				DlgVisible "EditFuncButton",False
				'DlgText "StatusText",Replace$(MsgList(19),"%s",Replace$(Replace$(MsgList(69),"%s",CStr$(0)),"%d",CStr$(0)))
				DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)),"%s",CStr$(0)),"%d",CStr$(0))
				DlgText "RefAddText1",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)),"%s",CStr$(0)),"%d",CStr$(0))
			Else
				DlgText "FunAddBox",ValToStr(DataList(0).lStartAddress,File.FileSize,True)
				TempList = GetDataValueList(DataList,0)
				DlgListBoxArray "FunNameList",TempList()
				DlgValue "FunNameList",0
				TempList = GetDataValueList(DataList,2,0,File.FileSize,True)
				Call StrListToLongList(TempList,RefAddList,True)
				Select Case DataList(0).DllLang
				Case "MS",""
					DlgListBoxArray "RefAddList",TempList()
					DlgValue "RefAddList",Array(0)
					AppCPList = GetDataValueList(DataList,3,0)
					DlgListBoxArray "AppCPList",AppCPList()
					DlgValue "AppCPList",Array(0)
					UseCPList = GetDataValueList(DataList,4,0)
					DlgListBoxArray "UseCPList",UseCPList()
					DlgValue "UseCPList",Array(0)
					DlgVisible "RefAddText",True
					DlgVisible "AppCPText",True
					DlgVisible "UseCPText",True
					DlgVisible "RefAddList",True
					DlgVisible "AppCPList",True
					DlgVisible "UseCPList",True
					DlgVisible "EditCPButton",True
					DlgVisible "RefAddText1",False
					DlgVisible "FuncText",False
					DlgVisible "RefAddList1",False
					DlgVisible "FuncList",False
					DlgVisible "EditFuncButton",False
					DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i", _
									CStr$(0)),"%s",CStr$(1)),"%d",CStr$(DataList(0).lReferenceNum))
				Case "QT"
					DlgListBoxArray "RefAddList1",TempList()
					DlgValue "RefAddList1",Array(0)
					FuncList = GetDataValueList(DataList,26,0)
					DlgListBoxArray "FuncList",FuncList()
					DlgValue "FuncList",Array(0)
					DlgVisible "RefAddText",False
					DlgVisible "AppCPText",False
					DlgVisible "UseCPText",False
					DlgVisible "RefAddList",False
					DlgVisible "AppCPList",False
					DlgVisible "UseCPList",False
					DlgVisible "EditCPButton",False
					DlgVisible "RefAddText1",True
					DlgVisible "FuncText",True
					DlgVisible "RefAddList1",True
					DlgVisible "FuncList",True
					DlgVisible "EditFuncButton",True
					DlgText "RefAddText1",Replace$(Replace$(Replace$(MsgList(16),"%i", _
									CStr$(0)),"%s",CStr$(1)),"%d",CStr$(DataList(0).lReferenceNum))
				End Select
				DlgText "StatusText",Replace$(MsgList(19),"%s",Replace$(Replace$(MsgList(69),"%s",CStr$(UBound(DataList) + 1)),"%d",DlgText("NumberBox")))
			End If
			If DataList(0).lReferenceNum = 0 Then
				DlgEnable "SelectAllButton",False
				DlgEnable "EditCPButton",False
				DlgEnable "EditFuncButton",False
				DlgEnable "UndoButton",False
				DlgEnable "UndoAllButton",False
				DlgEnable "InfoButton",False
				DlgEnable "WriteButton",False
				DlgEnable "AllWriteButton",False
				DlgEnable "ShowSetButton",False
			Else
				DlgEnable "SelectAllButton",True
				DlgEnable "EditCPButton",True
				DlgEnable "EditFuncButton",True
				DlgEnable "UndoButton",True
				DlgEnable "UndoAllButton",True
				DlgEnable "InfoButton",True
				DlgEnable "WriteButton",True
				DlgEnable "AllWriteButton",True
				DlgEnable "ShowSetButton",True
			End If
			DlgEnable "ShowAllButton",False
			DlgValue "WriteStateBox",0
		Case "FunNameList"
			i = DlgValue("FunNameList")
			If i < 0 Then Exit Function
			DlgText "FunAddBox",ValToStr(DataList(i).lStartAddress,File.FileSize,True)
			ReDim RefAddList(0) As Long,AppCPList(0) As String,UseCPList(0) As String,FuncList(0) As String
			Select Case DataList(i).DllLang
			Case "MS",""
				DlgVisible "RefAddText",True
				DlgVisible "AppCPText",True
				DlgVisible "UseCPText",True
				DlgVisible "RefAddList",True
				DlgVisible "AppCPList",True
				DlgVisible "UseCPList",True
				DlgVisible "EditCPButton",True
				DlgVisible "RefAddText1",False
				DlgVisible "FuncText",False
				DlgVisible "RefAddList1",False
				DlgVisible "FuncList",False
				DlgVisible "EditFuncButton",False
				If DataList(i).lReferenceNum = 0 Then
					DlgListBoxArray "RefAddList",TempList()
					DlgValue "RefAddList",Array(0)
					DlgListBoxArray "AppCPList",AppCPList()
					DlgValue "AppCPList",Array(0)
					DlgListBoxArray "UseCPList",UseCPList()
					DlgValue "UseCPList",Array(0)
					DlgListBoxArray "RefAddList1",TempList()
					'DlgText "StatusText",Replace$(MsgList(19),"%s",Replace$(Replace$(MsgList(69),"%s",CStr$(0)),"%d",CStr$(0)))
					DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)),"%s",CStr$(0)),"%d",CStr$(0))
				Else
					TempList = GetDataValueList(DataList,2,i,File.FileSize,True)
					Call StrListToLongList(TempList,RefAddList,True)
					DlgListBoxArray "RefAddList",TempList()
					DlgValue "RefAddList",Array(0)
					AppCPList = GetDataValueList(DataList,3,i)
					DlgListBoxArray "AppCPList",AppCPList()
					DlgValue "AppCPList",Array(0)
					UseCPList = GetDataValueList(DataList,4,i)
					DlgListBoxArray "UseCPList",UseCPList()
					DlgValue "UseCPList",Array(0)
					DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)), _
										"%s",CStr$(1)),"%d",CStr$(DlgListBoxArray("RefAddList")))
				End If
			Case "QT"
				DlgVisible "RefAddText",False
				DlgVisible "AppCPText",False
				DlgVisible "UseCPText",False
				DlgVisible "RefAddList",False
				DlgVisible "AppCPList",False
				DlgVisible "UseCPList",False
				DlgVisible "EditCPButton",False
				DlgVisible "RefAddText1",True
				DlgVisible "FuncText",True
				DlgVisible "RefAddList1",True
				DlgVisible "FuncList",True
				DlgVisible "EditFuncButton",True
				If DataList(i).lReferenceNum = 0 Then
					DlgListBoxArray "RefAddList1",TempList()
					DlgValue "RefAddList1",Array(0)
					DlgListBoxArray "FuncList",FuncList()
					DlgValue "FuncList",Array(0)
					'DlgText "StatusText",Replace$(MsgList(19),"%s",Replace$(Replace$(MsgList(69),"%s",CStr$(0)),"%d",CStr$(0)))
					DlgText "RefAddText1",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)),"%s",CStr$(0)),"%d",CStr$(0))
				Else
					TempList = GetDataValueList(DataList,2,i,File.FileSize,True)
					Call StrListToLongList(TempList,RefAddList,True)
					DlgListBoxArray "RefAddList1",TempList()
					DlgValue "RefAddList1",Array(0)
					FuncList = GetDataValueList(DataList,26,i)
					DlgListBoxArray "FuncList",FuncList()
					DlgValue "FuncList",Array(0)
					DlgText "RefAddText1",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)), _
										"%s",CStr$(1)),"%d",CStr$(DlgListBoxArray("RefAddList1")))
				End If
			End Select
			If DataList(i).lReferenceNum = 0 Then
				DlgEnable "SelectAllButton",False
				DlgEnable "EditCPButton",False
				DlgEnable "EditFuncButton",False
				DlgEnable "UndoButton",False
				DlgEnable "UndoAllButton",False
				DlgEnable "InfoButton",False
				DlgEnable "WriteButton",False
				DlgEnable "AllWriteButton",False
				DlgEnable "ShowSetButton",False
			Else
				DlgEnable "SelectAllButton",True
				DlgEnable "EditCPButton",True
				DlgEnable "EditFuncButton",True
				DlgEnable "UndoButton",True
				DlgEnable "UndoAllButton",True
				DlgEnable "InfoButton",True
				DlgEnable "WriteButton",True
				DlgEnable "AllWriteButton",True
				DlgEnable "ShowSetButton",True
			End If
			DlgEnable "ShowAllButton",False
		Case "RefAddList"
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")))
			If CheckArrEmpty(IntList) = False Then Exit Function
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AppCPList")),IntList)
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("UseCPList")),IntList)
			DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(IntList(0))), _
								"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(DlgListBoxArray("RefAddList")))
			MsgBox Reference2Str(File,DataList(DlgValue("FunNameList")),IntList,DlgEnable("ShowAllButton"),0),vbOkOnly,MsgList(33)
		Case "RefAddList1"
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")))
			If CheckArrEmpty(IntList) = False Then Exit Function
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("FuncList")),IntList)
			DlgText "RefAddText1",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(IntList(0))), _
								"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(DlgListBoxArray("RefAddList1")))
			MsgBox Reference2Str(File,DataList(DlgValue("FunNameList")),IntList,DlgEnable("ShowAllButton"),1),vbOkOnly,MsgList(33)
		Case "AppCPList"
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AppCPList")))
			If CheckArrEmpty(IntList) = False Then Exit Function
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")),IntList)
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("UseCPList")),IntList)
			DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(IntList(0))), _
								"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(DlgListBoxArray("RefAddList")))
		Case "UseCPList"
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("UseCPList")))
			If CheckArrEmpty(IntList) = False Then Exit Function
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")),IntList)
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AppCPList")),IntList)
			DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(IntList(0))), _
								"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(DlgListBoxArray("RefAddList")))
		Case "FuncList"
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("FuncList")))
			If CheckArrEmpty(IntList) = False Then Exit Function
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")),IntList)
			DlgText "RefAddText1",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(IntList(0))), _
								"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(DlgListBoxArray("RefAddList1")))
		Case "AboutButton"
			MsgBox Replace$(Replace$(MsgList(58),"%v",Version),"%b",Build),vbOkOnly+vbInformation,MsgList(59)
		Case "LangButton"
			ReDim TempList(0) As String
			TempList = ReSplit(MsgList(62),";")
			i = ShowPopupMenu(TempList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			TempList = ReSplit(MsgList(63),";")
			If GetMsgList(MsgList,TempList(i)) = False Then Exit Function
			File.Info = ""
			'更改文本框语言
			i = DlgValue("FunNameList")
			DlgText -1,Replace$(Replace$(MsgList(12),"%v",Version),"%b",Build)
			DlgText "FilePathButton",MsgList(13)
			DlgText "FunNameText",MsgList(14)
			DlgText "FunAddText",MsgList(15)
			i = DlgValue("FunNameList")
			If i < 0 Then
				DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)), _
									"%s",CStr$(0)),"%d",CStr$(DlgListBoxArray("RefAddList")))
			Else
				Select Case DataList(i).DllLang
				Case "MS",""
					IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")))
					If CheckArrEmpty(IntList) = False Then
						DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)), _
									"%s",CStr$(0)),"%d",CStr$(DlgListBoxArray("RefAddList")))
					Else
						DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(IntList(0))), _
									"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(DlgListBoxArray("RefAddList")))
					End If
				Case "QT"
					IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")))
					If CheckArrEmpty(IntList) = False Then
						DlgText "RefAddText1",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)), _
									"%s",CStr$(0)),"%d",CStr$(DlgListBoxArray("RefAddList1")))
					Else
						DlgText "RefAddText1",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(IntList(0))), _
									"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(DlgListBoxArray("RefAddList1")))
					End If
				End Select
			End If
			DlgText "AppCPText",MsgList(17)
			DlgText "UseCPText",MsgList(18)
			DlgText "FuncText",MsgList(81)
			If File.Magic = "" Then
				DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(68))
			ElseIf DataList(0).sString = "" Then
				DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(67))
			ElseIf DlgValue("WriteStateBox") = 0 Then
				DlgText "StatusText",Replace$(MsgList(19),"%s",Replace$(Replace$(MsgList(69),"%s",CStr$(UBound(DataList) + 1)),"%d",DlgText("NumberBox")))
			Else
				DlgText "StatusText",Replace$(MsgList(19),"%s",Replace$(MsgList(72),"%s",DlgText("NumberBox")))
			End If
			DlgText "AboutButton",MsgList(20)
			DlgText "LangButton",MsgList(21)
			DlgText "SelectAllButton",MsgList(22)
			DlgText "EditCPButton",MsgList(23)
			DlgText "EditFuncButton",MsgList(82)
			DlgText "UndoButton",MsgList(24)
			DlgText "UndoAllButton",MsgList(86)
			DlgText "InfoButton",MsgList(33)
			DlgText "WriteButton",MsgList(26)
			DlgText "AllWriteButton",MsgList(27)
			DlgText "ShowSetButton",MsgList(28)
			DlgText "ShowAllButton",MsgList(29)
			DlgText "StopButton",MsgList(30)
		Case "SelectAllButton"
			i = DlgValue("FunNameList")
			If i < 0 Then Exit Function
			Select Case DataList(i).DllLang
			Case "MS",""
				n = DlgListBoxArray("RefAddList")
				If n < 1 Then Exit Function
				SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")),-1)
				SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AppCPList")),-1)
				SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("UseCPList")),-1)
				DlgItem$ = "RefAddText"
			Case "QT"
				n = DlgListBoxArray("RefAddList1")
				If n < 1 Then Exit Function
				SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")),-1)
				SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("FuncList")),-1)
				DlgItem$ = "RefAddText1"
			End Select
			DlgText DlgItem$,Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)),"%s",CStr$(n)),"%d",CStr$(n))
		Case "EditCPButton"
			i = DlgValue("FunNameList")
			If i < 0 Then Exit Function
			ReDim TempList(21) As String
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
			'TempList(23) = CStr$(CP_UTF32LE)		'UnicodeLE = 65005
			'TempList(24) = CStr$(CP_UTF32BE)		'UnicodeBE = 65006
			For n = 0 To 21
				TempList(n) = TempList(n) & " - " & MsgList(n + 34)
			Next n
			n = ShowPopupMenu(TempList,vbPopupUseRightButton)
			If n < 0 Then Exit Function
			n = StrToLong(Trim$(ReSplit(TempList(n),"-")(0)),-1)
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")))
			i = ChangeCodePage(DataList(i),IntList,n,DlgEnable("ShowAllButton"),0)
			If i = 0 Then Exit Function
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("UseCPList")))
			DlgListBoxArray "UseCPList",UseCPList()
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("UseCPList")),IntList)
		Case "EditFuncButton"
			i = DlgValue("FunNameList")
			If i < 0 Then Exit Function
			TempList = GetDataValueList(DataList,23)
			n = ShowPopupMenu(TempList,vbPopupUseRightButton)
			If n < 0 Then Exit Function
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")))
			i = ChangeFunc(File,DataList(i),IntList,n,DlgEnable("ShowAllButton"),0,LoadMode)
			If i = 0 Then Exit Function
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("FuncList")))
			DlgListBoxArray "FuncList",FuncList()
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("FuncList")),IntList)
		Case "UndoButton"
			i = DlgValue("FunNameList")
			If i < 0 Then Exit Function
			Select Case DataList(i).DllLang
			Case "MS"
				IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")))
				i = ChangeCodePage(DataList(i),IntList,i,DlgEnable("ShowAllButton"),1)
				If i = 0 Then Exit Function
				IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("UseCPList")))
				DlgListBoxArray "UseCPList",UseCPList()
				SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("UseCPList")),IntList)
			Case "QT"
				IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")))
				i = ChangeFunc(File,DataList(i),IntList,i,DlgEnable("ShowAllButton"),1,LoadMode)
				If i = 0 Then Exit Function
				IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("FuncList")))
				DlgListBoxArray "FuncList",FuncList()
				SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("FuncList")),IntList)
			End Select
		Case "UndoAllButton"
			If DataList(0).sString = "" Then Exit Function
			For i = 0 To UBound(DataList)
				Select Case DataList(i).DllLang
				Case "MS"
					If i = DlgValue("FunNameList") Then
						n = ChangeCodePage(DataList(i),IntList,i,False,2)
					Else
						ChangeCodePage(DataList(i),IntList,i,False,2)
					End If
				Case "QT"
					If i = DlgValue("FunNameList") Then
						n = ChangeFunc(File,DataList(i),IntList,i,False,2,LoadMode)
					Else
						ChangeFunc(File,DataList(i),IntList,i,False,2,LoadMode)
					End If
				End Select
			Next i
			If n = 0 Then Exit Function
			ReDim LngList(0) As Long
			LngList = RefAddList
			i = DlgValue("FunNameList")
			Select Case DataList(i).DllLang
			Case "MS"
				If DataList(i).lReferenceNum = 0 Then Exit Function
				IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")))
				TempList = GetDataValueList(DataList,2,i,File.FileSize,True)
				Call StrListToLongList(TempList,RefAddList,True)
				DlgListBoxArray "RefAddList",TempList()
				DlgValue "RefAddList",Array(0)
				AppCPList = GetDataValueList(DataList,3,i)
				DlgListBoxArray "AppCPList",AppCPList()
				DlgValue "AppCPList",Array(0)
				UseCPList = GetDataValueList(DataList,4,i)
				DlgListBoxArray "UseCPList",UseCPList()
				DlgValue "UseCPList",Array(0)
				If CheckArrEmpty(IntList) = False Then
					DlgValue "RefAddList",Array(0)
					DlgValue "AppCPList",Array(0)
					DlgValue "UseCPList",Array(0)
				Else
					Call GetLngIndexList(LngList,RefAddList,IntList)
					SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")),IntList)
					SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")),IntList)
					SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("UseCPList")),IntList)
				End If
				DlgText "RefAddText",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(IntList(0))), _
									"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(n))
			Case "QT"
				If DataList(i).lReferenceNum = 0 Then Exit Function
				IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")))
				TempList = GetDataValueList(DataList,2,i,File.FileSize,True)
				Call StrListToLongList(TempList,RefAddList,True)
				DlgListBoxArray "RefAddList1",TempList()
				DlgValue "RefAddList1",Array(0)
				FuncList = GetDataValueList(DataList,26,i)
				DlgListBoxArray "FuncList",FuncList()
				DlgValue "FuncList",Array(0)
				If CheckArrEmpty(IntList) = False Then
					DlgValue "RefAddList1",Array(0)
				Else
					Call GetLngIndexList(LngList,RefAddList,IntList)
					SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")),IntList)
					SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("FuncList")),IntList)
				End If
				DlgText "RefAddText1",Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(IntList(0))), _
									"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(n))
			End Select
			DlgEnable "ShowAllButton",False
		Case "InfoButton"
			ReDim IntList(0) As Long
			TempList = ReSplit(MsgList(87),";")
			IntList(0) = ShowPopupMenu(TempList,vbPopupUseRightButton)
			Select Case IntList(0)
			Case Is < 0
				Exit Function
			Case 0
				If File.Info = "" Then Call FileInfoView(File,True)
				ShowInfo File.FilePath,File.Info
			Case 1
				i = DlgValue("FunNameList")
				If i < 0 Then Exit Function
				Select Case DataList(i).DllLang
				Case "MS",""
					IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")))
					If CheckArrEmpty(IntList) = False Then Exit Function
					ShowInfo File.FilePath,Reference2Str(File,DataList(i),IntList,DlgEnable("ShowAllButton"),2)
				Case "QT"
					IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")))
					If CheckArrEmpty(IntList) = False Then Exit Function
					ShowInfo File.FilePath,Reference2Str(File,DataList(i),IntList,DlgEnable("ShowAllButton"),3)
				End Select
			Case Else
				ReDim TempList(UBound(DataList)) As String
				n = 0
				For i = 0 To UBound(DataList)
					Select Case DataList(i).DllLang
					Case "MS",""
						TempList(n) = Reference2Str(File,DataList(i),IntList,False,2 + IntList(0))
					Case "QT"
						TempList(n) = Reference2Str(File,DataList(i),IntList,False,7 + IntList(0))
					End Select
					If TempList(n) <> "" Then n = n + 1
				Next i
				If n > 0 Then n = n - 1
				ReDim Preserve TempList(n) As String
				ShowInfo File.FilePath,StrListJoin(TempList,TextJoinStr)
			End Select
		Case "WriteButton", "AllWriteButton"
			i = DlgValue("FunNameList")
			If i < 0 Then Exit Function
			If PSL.SelectFile(Temp,False,MsgList(60),MsgList(61)) = False Then Exit Function
			If InStr(Temp,"\") Then
				If (Mid$(Temp,InStrRev(Temp,"\")) Like "*.*") = False Then
					If (File.FilePath Like "*.*") = True Then
						Temp = Temp & Mid$(File.FilePath,InStrRev(File.FilePath,"."))
					End If
				End If
			ElseIf (Temp Like "*.*") = False Then
				If (File.FilePath Like "*.*") = True Then
					Temp = Temp & Mid$(File.FilePath,InStrRev(File.FilePath,"."))
				End If
			End If
			If Temp = File.FilePath Then
				MsgBox MsgList(70),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
			On Error Resume Next
			FileCopy File.FilePath,Temp
			If Err.Number <> 0 Then
				Err.Source = "NotWriteFile"
				Err.Description = Err.Description & JoinStr & Temp
				Call sysErrorMassage(Err,2)
				Exit Function
			End If
			On Error GoTo 0
			DlgText "StatusText",Replace$(MsgList(19),"%s",MsgList(71))
			n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("StatusText"))
			Select Case DataList(i).DllLang
			Case "MS",""
				IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")))
			Case "QT"
				IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")))
			End Select
			If DlgItem$ = "AllWriteButton" Then
				DlgText "NumberBox",CStr$(WriteDatas(Temp,DataList,IntList,DlgValue("FunNameList"),False,LoadMode,DlgEnable("ShowAllButton"),n))
			Else
				DlgText "NumberBox",CStr$(WriteDatas(Temp,DataList,IntList,DlgValue("FunNameList"),True,LoadMode,DlgEnable("ShowAllButton"),n))
			End If
			DlgText "StatusText",Replace$(MsgList(19),"%s",Replace$(MsgList(72),"%s",DlgText("NumberBox")))
			DlgValue "WriteStateBox",1
		Case "ShowSetButton","ShowAllButton"
			i = DlgValue("FunNameList")
			If i < 0 Then Exit Function
			ReDim LngList(0) As Long,IntArray(0) As Long
			LngList = RefAddList
			Select Case DataList(i).DllLang
			Case "MS"
				IntArray = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")))
				Select Case DlgItem$
				Case "ShowAllButton"
					TempList = GetDataValueList(DataList,2,i,File.FileSize,True)
					Call StrListToLongList(TempList,RefAddList,True)
					DlgListBoxArray "RefAddList",TempList()
					AppCPList = GetDataValueList(DataList,3,i)
					DlgListBoxArray "AppCPList",AppCPList()
					UseCPList = GetDataValueList(DataList,4,i)
					DlgListBoxArray "UseCPList",UseCPList()
					DlgEnable "ShowAllButton",False
				Case "ShowSetButton"
					TempList = ReSplit(MsgList(32),";")
					n = ShowPopupMenu(TempList,vbPopupUseRightButton)
					Select Case n
					Case 0 To 3
						TempList = GetDataValueList(DataList,5 + 3 * n,i,File.FileSize,True)
						Call StrListToLongList(TempList,RefAddList,True)
						DlgListBoxArray "RefAddList",TempList()
						AppCPList = GetDataValueList(DataList,6 + 3 * n,i)
						DlgListBoxArray "AppCPList",AppCPList()
						UseCPList = GetDataValueList(DataList,7 + 3 * n,i)
						DlgListBoxArray "UseCPList",UseCPList()
					Case 4 To 5
						If n = 4 Then
							IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AppCPList")))
							If CheckArrEmpty(IntList) = False Then Exit Function
							ReDim TempList(UBound(IntList)) As String
							For i = 0 To UBound(IntList)
								TempList(i) = AppCPList(IntList(i))
							Next i
						Else
							IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("UseCPList")))
							If CheckArrEmpty(IntList) = False Then Exit Function
							ReDim TempList(UBound(IntList)) As String
							For i = 0 To UBound(IntList)
								TempList(i) = UseCPList(IntList(i))
							Next i
						End If
						Temp = vbNullChar & Join$(TempList,vbNullChar) & vbNullChar
						i = DlgValue("FunNameList")
						TempList = GetDataValueList(DataList,5 + 3 * n,i,File.FileSize,True,Temp)
						Call StrListToLongList(TempList,RefAddList,True)
						DlgListBoxArray "RefAddList",TempList()
						AppCPList = GetDataValueList(DataList,6 + 3 * n,i,,,Temp)
						DlgListBoxArray "AppCPList",AppCPList()
						UseCPList = GetDataValueList(DataList,7 + 3 * n,i,,,Temp)
						DlgListBoxArray "UseCPList",UseCPList()
					Case Else
						Exit Function
					End Select
					DlgEnable "ShowAllButton",IIf(DlgListBoxArray("RefAddList") = DataList(i).lReferenceNum,False,True)
				End Select
				If CheckArrEmpty(IntArray) = False Then
					DlgValue "RefAddList",Array(0)
					DlgValue "AppCPList",Array(0)
					DlgValue "UseCPList",Array(0)
				Else
					Call GetLngIndexList(LngList,RefAddList,IntArray)
					SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")),IntArray)
					SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AppCPList")),IntArray)
					SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("UseCPList")),IntArray)
				End If
				n = DlgListBoxArray("RefAddList")
				IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList")))
				DlgItem$ = "RefAddText"
			Case "QT"
				IntArray = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")))
				Select Case DlgItem$
				Case "ShowAllButton"
					TempList = GetDataValueList(DataList,2,i,File.FileSize,True)
					Call StrListToLongList(TempList,RefAddList,True)
					DlgListBoxArray "RefAddList1",TempList()
					FuncList = GetDataValueList(DataList,26,i)
					DlgListBoxArray "FuncList",FuncList()
					DlgEnable "ShowAllButton",False
				Case "ShowSetButton"
					TempList = ReSplit(MsgList(84),";")
					n = ShowPopupMenu(TempList,vbPopupUseRightButton)
					Select Case n
					Case 0 To 1
						TempList = GetDataValueList(DataList,27 + 3 * n,i,File.FileSize,True)
						Call StrListToLongList(TempList,RefAddList,True)
						DlgListBoxArray "RefAddList1",TempList()
						FuncList = GetDataValueList(DataList,29 + 3 * n,i)
						DlgListBoxArray "FuncList",FuncList()
					Case 2
						IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("FuncList")))
						If CheckArrEmpty(IntList) = False Then Exit Function
						ReDim TempList(UBound(IntList)) As String
						For i = 0 To UBound(IntList)
							TempList(i) = FuncList(IntList(i))
						Next i
						Temp = vbNullChar & Join$(TempList,vbNullChar) & vbNullChar
						i = DlgValue("FunNameList")
						TempList = GetDataValueList(DataList,27 + 3 * n,i,File.FileSize,True,Temp)
						Call StrListToLongList(TempList,RefAddList,True)
						DlgListBoxArray "RefAddList1",TempList()
						FuncList = GetDataValueList(DataList,29 + 3 * n,i,,,Temp)
						DlgListBoxArray "FuncList",FuncList()
					Case Else
						Exit Function
					End Select
					DlgEnable "ShowAllButton",IIf(DlgListBoxArray("RefAddList1") = DataList(i).lReferenceNum,False,True)
				End Select
				If CheckArrEmpty(IntArray) = False Then
					DlgValue "RefAddList1",Array(0)
				Else
					Call GetLngIndexList(LngList,RefAddList,IntArray)
					SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")),IntArray)
					SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("FuncList")),IntArray)
				End If
				n = DlgListBoxArray("RefAddList1")
				IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("RefAddList1")))
				DlgItem$ = "RefAddText1"
			End Select
			If n < 1 Then
				DlgText DlgItem$,Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(0)),"%s",CStr$(0)),"%d",CStr$(0))
				DlgEnable "SelectAllButton",False
				DlgEnable "EditCPButton",False
				DlgEnable "EditFuncButton",False
				DlgEnable "UndoButton",False
				DlgEnable "UndoAllButton",False
				DlgEnable "InfoButton",False
				DlgEnable "WriteButton",False
				DlgEnable "AllWriteButton",False
				i = DlgValue("FunNameList")
				If i < 0 Then Exit Function
				DlgEnable "ShowSetButton",IIf(DataList(i).lReferenceNum > 0,True,False)
			Else
				DlgText DlgItem$,Replace$(Replace$(Replace$(MsgList(16),"%i",CStr$(IntList(0))),"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(n))
				DlgEnable "SelectAllButton",True
				DlgEnable "EditCPButton",True
				DlgEnable "EditFuncButton",True
				DlgEnable "UndoButton",True
				DlgEnable "UndoAllButton",True
				DlgEnable "InfoButton",True
				DlgEnable "WriteButton",True
				DlgEnable "AllWriteButton",True
				DlgEnable "ShowSetButton",True
			End If
		End Select
	'Case 3 ' 文本框或者组合框文本更改时
	Case 6 ' 功能键
		Select Case SuppValue
		Case 1
			MsgBox Replace$(Replace$(MsgList(58),"%v",Version),"%b",Build),vbOkOnly+vbInformation,MsgList(59)
		End Select
	End Select
End Function


'找出二个数值数组中的值相同或不相同的索引列表
'Mode = False 获取二个数值数组中的值相同的索引列表，二个索引列表没有对应关系
'Mode = True 获取二个数值数组中的值不相同的索引列表，二个索引列表没有对应关系
Private Function GetLngIndexList(SrcList() As Long,TrgList() As Long,IntList() As Long,Optional ByVal Mode As Boolean) As Boolean
	Dim i As Long,n As Long,Dic As Object
	Set Dic = CreateObject("Scripting.Dictionary")
	For i = 0 To UBound(IntList)
		If Not Dic.Exists(SrcList(IntList(i))) Then
			Dic.Add(SrcList(IntList(i)),i)
		End If
	Next i
	ReDim IntList(UBound(IntList)) As Long
	If Mode = False Then
		For i = 0 To UBound(TrgList)
			If Dic.Exists(TrgList(i)) Then
				IntList(n) = i
				n = n + 1
			End If
		Next i
	Else
		For i = 0 To UBound(TrgList)
			If Not Dic.Exists(TrgList(i)) Then
				IntList(n) = i
				n = n + 1
			End If
		Next i
	End If
	Set Dic = Nothing
	If n > 0 Then
		ReDim Preserve IntList(n - 1) As Long
		GetLngIndexList = True
	Else
		ReDim IntList(0) As Long
	End If
End Function


'更改代码页值
'ChangeType = 0 更改指定函数的代码页值为指定值
'ChangeType = 1 撤消指定函数的代码页更改值为原始值
'ChangeType = 2 撤消全部函数的代码页更改值为原始值
Private Function ChangeCodePage(strData As STRING_SUB_PROPERTIE,IndexList() As Long,ByVal CodePage As Long, _
				ByVal ShowState As Boolean,Optional ByVal ChangeType As Integer) As Long
	Dim i As Long,j As Long,MinLength As Integer,Dic As Object
	With strData
		Select Case ChangeType
		Case 0
			MinLength = GetEvenPos(Len(Hex$(CodePage))) \ 2
			If ShowState = False Then
				For i = 0 To UBound(IndexList)
					j = IndexList(i)
					If .Reference(j).lOrgCodePage > 0 Then
						If .Reference(j).iCPByteLength >= MinLength Then
							If .Reference(j).lUseCodePage <> CodePage Then
								UseCPList(j) = CStr$(CodePage)
								.Reference(j).lUseCodePage = CodePage
								ChangeCodePage = ChangeCodePage + 1
							End If
						End If
					End If
				Next i
			Else
				Set Dic = CreateObject("Scripting.Dictionary")
				For i = 0 To .lReferenceNum - 1
					If .Reference(i).lOrgCodePage > 0 Then
						If Not Dic.Exists(.Reference(i).lAddress) Then
							Dic.Add(.Reference(i).lAddress,i)
						End If
					End If
				Next i
				For i = 0 To UBound(IndexList)
					If Dic.Exists(RefAddList(IndexList(i))) Then
						j = Dic.Item(RefAddList(IndexList(i)))
						If .Reference(j).iCPByteLength >= MinLength Then
							If .Reference(j).lUseCodePage <> CodePage Then
								UseCPList(IndexList(i)) = CStr$(CodePage)
								.Reference(j).lUseCodePage = CodePage
								ChangeCodePage = ChangeCodePage + 1
							End If
						End If
					End If
				Next i
				Set Dic = Nothing
			End If
		Case 1
			If ShowState = False Then
				For i = 0 To UBound(IndexList)
					j = IndexList(i)
					If .Reference(j).lOrgCodePage > 0 Then
						If .Reference(j).lUseCodePage <> .Reference(j).lOrgCodePage Then
							UseCPList(j) = CStr$(.Reference(j).lOrgCodePage)
							.Reference(j).lUseCodePage = .Reference(j).lOrgCodePage
							ChangeCodePage = ChangeCodePage + 1
						End If
					End If
				Next i
			Else
				Set Dic = CreateObject("Scripting.Dictionary")
				For i = 0 To .lReferenceNum - 1
					If .Reference(i).lOrgCodePage > 0 Then
						If Not Dic.Exists(.Reference(i).lAddress) Then
							Dic.Add(.Reference(i).lAddress,i)
						End If
					End If
				Next i
				For i = 0 To UBound(IndexList)
					If Dic.Exists(RefAddList(IndexList(i))) Then
						j = Dic.Item(RefAddList(IndexList(i)))
						If .Reference(j).lUseCodePage <> .Reference(j).lOrgCodePage Then
							UseCPList(IndexList(i)) = CStr$(.Reference(j).lOrgCodePage)
							.Reference(j).lUseCodePage = .Reference(j).lOrgCodePage
							ChangeCodePage = ChangeCodePage + 1
						End If
					End If
				Next i
				Set Dic = Nothing
			End If
		Case 2
			For i = 0 To .lReferenceNum - 1
				If .Reference(i).lOrgCodePage > 0 Then
					If .Reference(i).lUseCodePage <> .Reference(i).lOrgCodePage Then
						'UseCPList(i) = CStr$(.Reference(i).lOrgCodePage)
						.Reference(i).lUseCodePage = .Reference(i).lOrgCodePage
						ChangeCodePage = ChangeCodePage + 1
					End If
				End If
			Next i
		End Select
	End With
End Function


'更改函数
'ChangeType = 0 更改指定函数为指定值
'ChangeType = 1 撤消指定函数的更改值为原始值
'ChangeType = 2 撤消全部函数的更改值为原始值
Private Function ChangeFunc(File As FILE_PROPERTIE,strData As STRING_SUB_PROPERTIE,IndexList() As Long,ByVal FuncID As Long, _
				ByVal ShowState As Boolean,Optional ByVal ChangeType As Integer,Optional ByVal Mode As Long) As Long
	Dim i As Long,j As Long,Dic As Object,FN As FILE_IMAGE
	With strData
		Select Case ChangeType
		Case 0
			If ShowState = False Then
				For i = 0 To UBound(IndexList)
					j = IndexList(i)
					If .Reference(j).lUseFuncAddress <> DataList(FuncID).lStartAddress Then
						FuncList(j) = DataList(FuncID).sString
						.Reference(j).sUseFunc = DataList(FuncID).sString
						.Reference(j).lUseFuncAddress = DataList(FuncID).lStartAddress
						.Reference(j).UseFuncInSecID = DataList(FuncID).inSectionID
						ChangeFunc = ChangeFunc + 1
					End If
				Next i
			Else
				Set Dic = CreateObject("Scripting.Dictionary")
				For i = 0 To .lReferenceNum - 1
					If Not Dic.Exists(.Reference(i).lAddress) Then
						Dic.Add(.Reference(i).lAddress,i)
					End If
				Next i
				For i = 0 To UBound(IndexList)
					If Dic.Exists(RefAddList(IndexList(i))) Then
						j = Dic.Item(RefAddList(IndexList(i)))
						If .Reference(j).lUseFuncAddress <> DataList(FuncID).lStartAddress Then
							FuncList(IndexList(i)) = DataList(FuncID).sString
							.Reference(j).sUseFunc = DataList(FuncID).sString
							.Reference(j).lUseFuncAddress = DataList(FuncID).lStartAddress
							.Reference(j).UseFuncInSecID = DataList(FuncID).inSectionID
							ChangeFunc = ChangeFunc + 1
						End If
					End If
				Next i
				Set Dic = Nothing
			End If
			If ChangeFunc > 0 Then
				'打开文件
				Mode = LoadFile(File.FilePath,FN,0,0,0,Mode)
				If Mode < -1 Then
					UnLoadFile(FN,0,Mode)
					Exit Function
				End If
				GetVARefList(File,FN,strData,"",1,Mode)
				UnLoadFile(FN,0,Mode)
			End If
		Case 1
			If ShowState = False Then
				For i = 0 To UBound(IndexList)
					j = IndexList(i)
					If .Reference(j).lUseFuncAddress <> .lStartAddress Then
						FuncList(j) = .sString
						.Reference(j).sUseFunc = .sString
						.Reference(j).sUseCode = .Reference(IndexList(i)).sOrgCode
						.Reference(j).lUseFuncAddress = .lStartAddress
						.Reference(j).UseFuncInSecID = .inSectionID
						ChangeFunc = ChangeFunc + 1
					End If
				Next i
			Else
				Set Dic = CreateObject("Scripting.Dictionary")
				For i = 0 To .lReferenceNum - 1
					If Not Dic.Exists(.Reference(i).lAddress) Then
						Dic.Add(.Reference(i).lAddress,i)
					End If
				Next i
				For i = 0 To UBound(IndexList)
					If Dic.Exists(RefAddList(IndexList(i))) Then
						j = Dic.Item(RefAddList(IndexList(i)))
						If .Reference(j).lUseFuncAddress <> .lStartAddress Then
							FuncList(IndexList(i)) = .sString
							.Reference(j).sUseFunc = .sString
							.Reference(j).sUseCode = .Reference(j).sOrgCode
							.Reference(j).lUseFuncAddress = .lStartAddress
							.Reference(j).UseFuncInSecID = .inSectionID
							ChangeFunc = ChangeFunc + 1
						End If
					End If
				Next i
				Set Dic = Nothing
			End If
		Case 2
			For i = 0 To .lReferenceNum - 1
				If .Reference(i).lUseFuncAddress <> .lStartAddress Then
					'FuncList(i) = .sString
					.Reference(i).sUseFunc = .sString
					.Reference(i).sUseCode = .Reference(i).sOrgCode
					.Reference(i).lUseFuncAddress = .lStartAddress
					.Reference(i).UseFuncInSecID = .inSectionID
					ChangeFunc = ChangeFunc + 1
				End If
			Next i
		End Select
	End With
End Function


'获取代码页属性
Private Sub GetCPList(CPList() As CODEPAGE_PROPERTIE)
	Dim i As Long
	ReDim CPList(23) As CODEPAGE_PROPERTIE
	CPList(0).CodePage = CP_OEMCP			'系统默认
	CPList(1).CodePage = CP_MACCP
	CPList(2).CodePage = CP_THREAD_ACP
	CPList(3).CodePage = CP_WESTEUROPE		'拉丁文 1 (ANSI) = 1252
	CPList(4).CodePage = CP_EASTEUROPE		'拉丁文 2 (中欧) = 1250
	CPList(5).CodePage = CP_RUSSIAN			'西里尔文 (斯拉夫) = 1251
	CPList(6).CodePage = CP_GREEK			'希腊文 = 1253
	CPList(7).CodePage = CP_TURKISH			'拉丁文 5 (土耳其) = 1254
	CPList(8).CodePage = CP_HEBREW			'希伯来文 = 1255
	CPList(9).CodePage = CP_ARABIC			'阿拉伯文 = 1256
	CPList(10).CodePage = CP_BALTIC			'波罗的海文 = 1257
	CPList(11).CodePage = CP_VIETNAMESE		'越南文 = 1258
	CPList(12).CodePage = CP_JAPAN			'日文 = 932
	CPList(13).CodePage = CP_CHINA			'简体中文 = 936
	CPList(14).CodePage = CP_GB18030		'简体中文 = 54936
	CPList(15).CodePage = CP_KOREA			'韩文 = 949
	CPList(16).CodePage = CP_TAIWAN	 		'繁体中文 = 950
	CPList(17).CodePage = CP_THAI			'泰文 = 874
	CPList(18).CodePage = CP_UTF7			'UTF-7 = 65000
	CPList(19).CodePage = CP_UTF8			'UTF-8 = 65001
	CPList(20).CodePage = CP_UNICODELITTLE	'UnicodeLE = 1200
	CPList(21).CodePage = CP_UNICODEBIG		'UnicodeBE = 1201
	CPList(22).CodePage = CP_UTF32LE		'UnicodeLE = 65005
	CPList(23).CodePage = CP_UTF32BE		'UnicodeBE = 65006
	For i = 0 To 23
		With CPList(i)
			.ByteLength = GetEvenPos(Len(Hex$(.CodePage))) \ 2
			.LEPatrn = Byte2RegExpPattern(Val2Bytes(.CodePage,.ByteLength,False),1,CP_ISOLATIN1) & "\x00*"
			.BEPatrn = "\x00*" & Byte2RegExpPattern(Val2Bytes(.CodePage,.ByteLength,True),1,CP_ISOLATIN1)
		End With
	Next i
End Sub


'获取数据的某个函数值列表，用于显示
'ItemType = 0 获取函数名称
'ItemType = 1 获取函数地址

'ItemType = 2 获取函数的全部引用地址
'ItemType = 3 获取函数的全部原始代码页项
'ItemType = 4 获取函数的全部使用代码页项

'ItemType = 5 获取函数的代码页值可更改的引用地址
'ItemType = 6 获取函数的代码页值可更改的原始代码页项
'ItemType = 7 获取函数的代码页值可更改的使用代码页项

'ItemType = 8 获取函数的代码页值未更改的引用地址
'ItemType = 9 获取函数的代码页值未更改的原始代码页项
'ItemType = 10 获取函数的代码页值未更改的使用代码页项

'ItemType = 11 获取函数的代码页值已更改的引用地址
'ItemType = 12 获取函数的代码页值已更改的原始代码页项
'ItemType = 13 获取函数的代码页值已更改的使用代码页项

'ItemType = 14 获取函数的代码页值可更改但未更改的引用地址
'ItemType = 15 获取函数的代码页值可更改但未更改的原始代码页项
'ItemType = 16 获取函数的代码页值可更改但未更改的使用代码页项

'ItemType = 17 获取函数的指定原始代码页值的引用地址
'ItemType = 18 获取函数的指定原始代码页值的原始代码页项
'ItemType = 19 获取函数的指定原始代码页值的使用代码页项

'ItemType = 20 获取函数的指定使用代码页值的引用地址
'ItemType = 21 获取函数的指定使用代码页值的原始代码页项
'ItemType = 22 获取函数的指定使用代码页值的使用代码页项

'ItemType = 23 仅获取 QT 函数名称

'ItemType = 24 获取函数的全部引用地址
'ItemType = 25 获取函数的全部原始函数名
'ItemType = 26 获取函数的全部使用函数名

'ItemType = 27 获取函数的使用函数名未更改的引用地址
'ItemType = 28 获取函数的使用函数名未更改的原始函数名项
'ItemType = 29 获取函数的使用函数名未更改的使用函数名项

'ItemType = 30 获取函数的使用函数名已更改的引用地址
'ItemType = 31 获取函数的使用函数名已更改的原始函数名项
'ItemType = 32 获取函数的使用函数名已更改的使用函数名项

'ItemType = 33 获取函数的指定使用函数名的引用地址
'ItemType = 34 获取函数的指定使用函数名的原始函数名项
'ItemType = 35 获取函数的指定使用函数名的使用函数名项

Private Function GetDataValueList(DataList() As STRING_SUB_PROPERTIE,ByVal ItemType As Long,Optional ByVal Index As Long, _
			Optional ByVal Size As Long,Optional ByVal Display As Boolean,Optional ByVal Value As String) As String()
	Dim i As Long,n As Long
	Select Case ItemType
	Case 0
		n = UBound(DataList)
		ReDim TempList(n) As String
		For i = 0 To n
			TempList(i) = DataList(i).sString
		Next i
	Case 1
		n = UBound(DataList)
		ReDim TempList(n) As String
		For i = 0 To n
			TempList(i) = CStr$(DataList(i).lStartAddress)
		Next i
	Case 2 To 4
		If DataList(Index).lReferenceNum < 1 Then
			ReDim TempList(0) As String
		Else
			ReDim TempList(DataList(Index).lReferenceNum - 1) As String
			With DataList(Index)
				For i = 0 To .lReferenceNum - 1
					Select Case	ItemType
					Case 2
						TempList(i) = ValToStr(.Reference(i).lAddress,Size,Display)
					Case 3
						TempList(i) = CStr$(.Reference(i).lOrgCodePage)
					Case 4
						TempList(i) = CStr$(.Reference(i).lUseCodePage)
					End Select
				Next i
			End With
		End If
	Case 5 To 7
		If DataList(Index).lReferenceNum < 1 Then
			ReDim TempList(0) As String
		Else
			ReDim TempList(DataList(Index).lReferenceNum - 1) As String
			With DataList(Index)
				For i = 0 To .lReferenceNum - 1
					If .Reference(i).lOrgCodePage > 0 Then
						Select Case	ItemType
						Case 5
							TempList(n) = ValToStr(.Reference(i).lAddress,Size,Display)
						Case 6
							TempList(n) = CStr$(.Reference(i).lOrgCodePage)
						Case 7
							TempList(n) = CStr$(.Reference(i).lUseCodePage)
						End Select
						n = n + 1
					End If
				Next i
			End With
			If n > 0 Then n = n - 1
			ReDim Preserve TempList(n) As String
		End If
	Case 8 To 10
		If DataList(Index).lReferenceNum < 1 Then
			ReDim TempList(0) As String
		Else
			ReDim TempList(DataList(Index).lReferenceNum - 1) As String
			With DataList(Index)
				For i = 0 To .lReferenceNum - 1
					If .Reference(i).lOrgCodePage = .Reference(i).lUseCodePage Then
						Select Case	ItemType
						Case 8
							TempList(n) = ValToStr(.Reference(i).lAddress,Size,Display)
						Case 9
							TempList(n) = CStr$(.Reference(i).lOrgCodePage)
						Case 10
							TempList(n) = CStr$(.Reference(i).lUseCodePage)
						End Select
						n = n + 1
					End If
				Next i
			End With
			If n > 0 Then n = n - 1
			ReDim Preserve TempList(n) As String
		End If
	Case 11 To 13
		If DataList(Index).lReferenceNum < 1 Then
			ReDim TempList(0) As String
		Else
			ReDim TempList(DataList(Index).lReferenceNum - 1) As String
			With DataList(Index)
				For i = 0 To .lReferenceNum - 1
					If .Reference(i).lOrgCodePage <> .Reference(i).lUseCodePage Then
						Select Case	ItemType
						Case 11
							TempList(n) = ValToStr(.Reference(i).lAddress,Size,Display)
						Case 12
							TempList(n) = CStr$(.Reference(i).lOrgCodePage)
						Case 13
							TempList(n) = CStr$(.Reference(i).lUseCodePage)
						End Select
						n = n + 1
					End If
				Next i
			End With
			If n > 0 Then n = n - 1
			ReDim Preserve TempList(n) As String
		End If
	Case 14 To 16
		If DataList(Index).lReferenceNum < 1 Then
			ReDim TempList(0) As String
		Else
			ReDim TempList(DataList(Index).lReferenceNum - 1) As String
			With DataList(Index)
				For i = 0 To .lReferenceNum - 1
					If .Reference(i).lOrgCodePage > 0 Then
						If .Reference(i).lOrgCodePage = .Reference(i).lUseCodePage Then
							Select Case	ItemType
							Case 14
								TempList(n) = ValToStr(.Reference(i).lAddress,Size,Display)
							Case 15
								TempList(n) = CStr$(.Reference(i).lOrgCodePage)
							Case 16
								TempList(n) = CStr$(.Reference(i).lUseCodePage)
							End Select
							n = n + 1
						End If
					End If
				Next i
			End With
			If n > 0 Then n = n - 1
			ReDim Preserve TempList(n) As String
		End If
	Case 17 To 19
		If DataList(Index).lReferenceNum < 1 Then
			ReDim TempList(0) As String
		Else
			ReDim TempList(DataList(Index).lReferenceNum - 1) As String
			With DataList(Index)
				For i = 0 To .lReferenceNum - 1
					If InStr(Value,vbNullChar & CStr$(.Reference(i).lOrgCodePage) & vbNullChar) Then
						Select Case	ItemType
						Case 17
							TempList(n) = ValToStr(.Reference(i).lAddress,Size,Display)
						Case 18
							TempList(n) = CStr$(.Reference(i).lOrgCodePage)
						Case 19
							TempList(n) = CStr$(.Reference(i).lUseCodePage)
						End Select
						n = n + 1
					End If
				Next i
			End With
			If n > 0 Then n = n - 1
			ReDim Preserve TempList(n) As String
		End If
	Case 20 To 22
		If DataList(Index).lReferenceNum < 1 Then
			ReDim TempList(0) As String
		Else
			ReDim TempList(DataList(Index).lReferenceNum - 1) As String
			With DataList(Index)
				For i = 0 To .lReferenceNum - 1
					If InStr(Value,vbNullChar & CStr$(.Reference(i).lUseCodePage) & vbNullChar) Then
						Select Case	ItemType
						Case 20
							TempList(n) = ValToStr(.Reference(i).lAddress,Size,Display)
						Case 21
							TempList(n) = CStr$(.Reference(i).lOrgCodePage)
						Case 22
							TempList(n) = CStr$(.Reference(i).lUseCodePage)
						End Select
						n = n + 1
					End If
				Next i
			End With
			If n > 0 Then n = n - 1
			ReDim Preserve TempList(n) As String
		End If
	Case 23
		n = UBound(DataList)
		ReDim TempList(n) As String
		For i = 0 To n
			If DataList(i).DllLang = "QT" Then
				TempList(i) = DataList(i).sString
			End If
		Next i
	Case 24 To 26
		If DataList(Index).lReferenceNum < 1 Then
			ReDim TempList(0) As String
		Else
			ReDim TempList(DataList(Index).lReferenceNum - 1) As String
			With DataList(Index)
				For i = 0 To .lReferenceNum - 1
					Select Case	ItemType
					Case 24
						TempList(i) = ValToStr(.Reference(i).lAddress,Size,Display)
					Case 25
						TempList(i) = .sString
					Case 26
						TempList(i) = .Reference(i).sUseFunc
					End Select
				Next i
			End With
		End If
	Case 27 To 29
		If DataList(Index).lReferenceNum < 1 Then
			ReDim TempList(0) As String
		Else
			ReDim TempList(DataList(Index).lReferenceNum - 1) As String
			With DataList(Index)
				For i = 0 To .lReferenceNum - 1
					If .lStartAddress = .Reference(i).lUseFuncAddress Then
						Select Case	ItemType
						Case 27
							TempList(n) = ValToStr(.Reference(i).lAddress,Size,Display)
						Case 28
							TempList(n) = .sString
						Case 29
							TempList(n) = .Reference(i).sUseFunc
						End Select
						n = n + 1
					End If
				Next i
			End With
			If n > 0 Then n = n - 1
			ReDim Preserve TempList(n) As String
		End If
	Case 30 To 32
		If DataList(Index).lReferenceNum < 1 Then
			ReDim TempList(0) As String
		Else
			ReDim TempList(DataList(Index).lReferenceNum - 1) As String
			With DataList(Index)
				For i = 0 To .lReferenceNum - 1
					If .lStartAddress <> .Reference(i).lUseFuncAddress Then
						Select Case	ItemType
						Case 30
							TempList(n) = ValToStr(.Reference(i).lAddress,Size,Display)
						Case 31
							TempList(n) = .sString
						Case 32
							TempList(n) = .Reference(i).sUseFunc
						End Select
						n = n + 1
					End If
				Next i
			End With
			If n > 0 Then n = n - 1
			ReDim Preserve TempList(n) As String
		End If
	Case 33 To 35
		If DataList(Index).lReferenceNum < 1 Then
			ReDim TempList(0) As String
		Else
			ReDim TempList(DataList(Index).lReferenceNum - 1) As String
			With DataList(Index)
				For i = 0 To .lReferenceNum - 1
					If InStr(Value,vbNullChar & CStr$(.Reference(i).sUseFunc) & vbNullChar) Then
						Select Case	ItemType
						Case 33
							TempList(n) = ValToStr(.Reference(i).lAddress,Size,Display)
						Case 34
							TempList(n) = .sString
						Case 35
							TempList(n) = .Reference(i).sUseFunc
						End Select
						n = n + 1
					End If
				Next i
			End With
			If n > 0 Then n = n - 1
			ReDim Preserve TempList(n) As String
		End If
	End Select
	GetDataValueList = TempList
End Function


'获取数据的某个函数值列表，用于显示
Private Sub StrListToLongList(strList() As String,LongList() As Long,Optional ByVal Display As Boolean)
	Dim i As Long
	ReDim LongList(UBound(strList)) As Long
	For i = 0 To UBound(strList)
		LongList(i) = StrToVal(strList(i),Display)
	Next i
End Sub


'获取函数地址
Private Sub GetFuncVA(DataList() As STRING_SUB_PROPERTIE)
	Dim i As Long,lb As Long,pa As Long
	On Error Resume Next
	For i = 0 To UBound(DataList)
		lb = LoadLibrary(DataList(i).DllName)
		DataList(i).FuncVA = GetProcAddress(lb, DataList(i).sString)
		'CallWindowProc pa, Me.hWnd, "Hello !", ByVal 0&, ByVal 0&
		FreeLibrary lb
	Next i
End Sub


'获取函数的引用地址
Private Function GetFuncReference(File As FILE_PROPERTIE,DataList() As STRING_SUB_PROPERTIE,ByVal Mode As Long,Optional ByVal ShowMsg As Long) As Long
	Dim i As Long,j As Long,FN As FILE_IMAGE,Msg As String,MsgBak As String
	'打开文件
	Mode = LoadFile(File.FilePath,FN,0,0,0,Mode)
	If Mode < -1 Then
		UnLoadFile(FN,0,Mode)
		Exit Function
	End If
	If ShowMsg > 0 Then
		Msg = GetTextBoxString(ShowMsg)
	ElseIf ShowMsg < 0 Then
		ReDim TempList(PSL.OutputWnd(0).LineCount - 1) As String
		For i = 1 To PSL.OutputWnd(0).LineCount
			TempList(i - 1) = PSL.OutputWnd(0).Text(i)
		Next i
		Msg = StrListJoin(TempList,vbCrLf)
	End If
	For i = 0 To UBound(DataList)
		If ShowMsg > 0 Then
			MsgBak = DataList(i).sString
			j = Len(MsgBak)
			If j > 20 Then MsgBak = Left$(MsgBak,16) & "..." & Mid$(MsgBak,j - 2)
			SetTextBoxString ShowMsg,Replace$(Msg,"%s",MsgBak)
		ElseIf ShowMsg < 0 Then
			MsgBak = DataList(i).sString
			j = Len(MsgBak)
			If j > 20 Then MsgBak = Left$(MsgBak,16) & "..." & Mid$(MsgBak,j - 2)
			PSL.OutputWnd(0).Clear
			PSL.Output Replace$(Msg,"%s",MsgBak)
		End If
		GetFuncReference = GetFuncReference + GetVARefList(File,FN,DataList(i),"",0,Mode,ShowMsg)
	Next i
	UnLoadFile(FN,0,Mode)
End Function


'获取函数的引用地址所在的代码页
Private Function GetCodePage(File As FILE_PROPERTIE,DataList() As STRING_SUB_PROPERTIE,CPList() As CODEPAGE_PROPERTIE,ByVal Mode As Long,Optional ByVal ShowMsg As Long) As Long
	Dim i As Long,j As Long,k As Integer,n As Integer,x As Integer,y As Integer
	Dim FN As FILE_IMAGE,Temp As String,Matches As Object,Stemp As Boolean
	Dim Msg As String,MsgBak As String,CodePageList() As CODEPAGE_PROPERTIE
	Const m = 48&
	'打开文件
	Mode = LoadFile(File.FilePath,FN,0,0,0,Mode)
	If Mode < -1 Then
		UnLoadFile(FN,0,Mode)
		Exit Function
	End If
	If ShowMsg > 0 Then
		Msg = GetTextBoxString(ShowMsg) & " "
	ElseIf ShowMsg < 0 Then
		ReDim TempList(PSL.OutputWnd(0).LineCount - 1) As String
		For i = 1 To PSL.OutputWnd(0).LineCount
			TempList(i - 1) = PSL.OutputWnd(0).Text(i)
		Next i
		Msg = StrListJoin(TempList,vbCrLf) & " "
	End If
	CodePageList = CPList
	Select Case File.Magic
	Case "PE32","NET32","MAC32"
		For i = 0 To 23
			CodePageList(i).LEPatrn = CodePageFrontChar32 & CodePageList(i).LEPatrn
			CodePageList(i).BEPatrn = CodePageFrontChar32 & CodePageList(i).BEPatrn
		Next i
	Case Else
		For i = 0 To 23
			CodePageList(i).LEPatrn = CodePageFrontChar64 & CodePageList(i).LEPatrn
			CodePageList(i).BEPatrn = CodePageFrontChar64 & CodePageList(i).BEPatrn
		Next i
	End Select
	RegExp.Global = True
	RegExp.IgnoreCase = False
	For i = 0 To UBound(DataList)
		With DataList(i)
			If ShowMsg > 0 Then
				MsgBak = .sString
				j = Len(MsgBak)
				If j > 20 Then MsgBak = Left$(MsgBak,16) & "..." & Mid$(MsgBak,j - 2)
				MsgBak = Replace$(Msg,"%s",MsgBak)
				SetTextBoxString ShowMsg,MsgBak
			ElseIf ShowMsg < 0 Then
				MsgBak = .sString
				j = Len(MsgBak)
				If j > 20 Then MsgBak = Left$(MsgBak,16) & "..." & Mid$(MsgBak,j - 2)
				MsgBak = Replace$(Msg,"%s",MsgBak)
				PSL.OutputWnd(0).Clear
				PSL.Output MsgBak
			End If
			If .DllLang = "QT" Then GoTo NextNo
			For j = 0 To .lReferenceNum - 1
				Stemp = False
				Temp = ByteToString(GetBytes(FN,m - 2,.Reference(j).lAddress - m,Mode),CP_ISOLATIN1)
				For k = 3 To 21
					For x = 0 To 1
						If x = 0 Then
							RegExp.Pattern = CodePageList(k).LEPatrn
						Else
							RegExp.Pattern = CodePageList(k).BEPatrn
						End If
						Set Matches = RegExp.Execute(Temp)
						If Matches.Count > 0 Then
							n = Matches.Count - 1
							.Reference(j).lOrgCodePage = CodePageList(k).CodePage
							.Reference(j).lUseCodePage = CodePageList(k).CodePage
							.Reference(j).iCPByteLength = GetValByteLength(Matches(n).Length - 1)
							If x = 0 Then
								.Reference(j).lCPAddress = .Reference(j).lAddress + Matches(n).FirstIndex - m + 1
								.Reference(j).iCPByteOrder = False
							ElseIf Matches(n).Length > .Reference(j).iCPByteLength Then
								y = Matches(n).Length - .Reference(j).iCPByteLength
								.Reference(j).lCPAddress = .Reference(j).lAddress + Matches(n).FirstIndex - m + 1 + y
								.Reference(j).iCPByteOrder = True
							Else
								.Reference(j).lCPAddress = .Reference(j).lAddress + Matches(n).FirstIndex - m + 1
								.Reference(j).iCPByteOrder = True
							End If
							GetCodePage = GetCodePage + 1
							Stemp = True
							Exit For
						End If
					Next x
					If Stemp = True Then Exit For
				Next k
				If Stemp = False Then
					Select Case File.Magic
					Case "PE32","NET32","MAC32"
						RegExp.Pattern = CodePageFrontChar32 & "\x00+" & ImportFunRefFrontChar32 & "$"
						Set Matches = RegExp.Execute(Temp)
						If Matches.Count > 0 Then
							.Reference(j).lOrgCodePage = CP_ACP
							.Reference(j).lUseCodePage = CP_ACP
							.Reference(j).iCPByteLength = GetValByteLength(Matches(0).Length - 3)
							.Reference(j).lCPAddress = .Reference(j).lAddress + Matches(0).FirstIndex - m + 1
							.Reference(j).iCPByteOrder = False
							GetCodePage = GetCodePage + 1
							Stemp = True
						End If
					Case Else
						RegExp.Pattern = CodePageFrontChar64 & "\x00+" & ImportFunRefFrontChar64 & "$"
						Set Matches = RegExp.Execute(Temp)
						If Matches.Count > 0 Then
							.Reference(j).lOrgCodePage = CP_ACP
							.Reference(j).lUseCodePage = CP_ACP
							.Reference(j).iCPByteLength = GetValByteLength(Matches(0).Length - 5)
							.Reference(j).lCPAddress = .Reference(j).lAddress + Matches(0).FirstIndex - m + 1
							.Reference(j).iCPByteOrder = False
							GetCodePage = GetCodePage + 1
							Stemp = True
						End If
					End Select
				End If
				If Stemp = False Then
					.Reference(j).lOrgCodePage = CP_UNKNOWN
					.Reference(j).lUseCodePage = CP_UNKNOWN
					.Reference(j).iCPByteLength = 0
					.Reference(j).iCPByteOrder = 1
				End If
				If ShowMsg > 0 Then
					SetTextBoxString ShowMsg,MsgBak & Format$(j / (.lReferenceNum + 1),"#%")
				ElseIf ShowMsg < 0 Then
					PSL.OutputWnd(0).Clear
					PSL.Output MsgBak & Format$(j / (.lReferenceNum + 1),"#%")
				End If
			Next j
			NextNo:
		End With
	Next i
	UnLoadFile(FN,0,Mode)
End Function


'获取函数的引用地址所在的函数名
Private Sub SetDataList(DataList() As STRING_SUB_PROPERTIE)
	Dim i As Long,j As Long
	For i = 0 To UBound(DataList)
		With DataList(i)
			For j = 0 To .lReferenceNum - 1
				.Reference(j).sUseCode = .Reference(j).sOrgCode
				.Reference(j).sUseFunc = .sString
				.Reference(j).lUseFuncAddress = .lStartAddress
				.Reference(j).UseFuncInSecID = .inSectionID
			Next j
		End With
	Next i
End Sub


'获取数值字节的适合长度
Private Function GetValByteLength(StrLength As Long) As Long
	Select Case StrLength
	Case 0
		GetValByteLength = 0
	Case 1
		GetValByteLength = 1
	Case 2 To 3
		GetValByteLength = 2
	Case Else
		GetValByteLength = 4
	End Select
End Function


'修改所有或选定数据
'fType = False 全部写入，否则按选定的项目写入
Private Function WriteDatas(ByVal FilePath As String,DataList() As STRING_SUB_PROPERTIE,IndexList() As Long, _
			ByVal FunID As Long,ByVal fType As Boolean,ByVal Mode As Long,ByVal ShowState As Boolean,Optional ByVal ShowMsg As Long) As Long
	Dim i As Long,j As Long,k As Long,Msg As String,FN As FILE_IMAGE
	If FunID = -1 And fType = True Then Exit Function
	'打开文件
	Mode = LoadFile(FilePath,FN,0,1,0,Mode)
	If Mode < -1 Then
		UnLoadFile(FN,0,Mode)
		Exit Function
	End If
	If ShowMsg > 0 Then
		Msg = GetTextBoxString(ShowMsg) & " "
	ElseIf ShowMsg < 0 Then
		ReDim TempList(PSL.OutputWnd(0).LineCount - 1) As String
		For i = 1 To PSL.OutputWnd(0).LineCount
			TempList(i - 1) = PSL.OutputWnd(0).Text(i)
		Next i
		Msg = StrListJoin(TempList,vbCrLf) & " "
	End If
	If fType = False Then
		Dim MsgBak As String
		For i = 0 To UBound(DataList)
			With DataList(i)
				If ShowMsg > 0 Then
					MsgBak = Replace$(Msg,"%s",.sString)
					SetTextBoxString ShowMsg,MsgBak
				ElseIf ShowMsg < 0 Then
					MsgBak = Replace$(Msg,"%s",.sString)
					PSL.OutputWnd(0).Clear
					PSL.Output MsgBak
				End If
				For j = 0 To .lReferenceNum - 1
					If WriteData(FN,DataList(i),j,Mode) = True Then
						WriteDatas = WriteDatas + 1
					End If
					If ShowMsg > 0 Then
						SetTextBoxString ShowMsg,MsgBak & Format$((j + 1) / .lReferenceNum,"#%")
					ElseIf ShowMsg < 0 Then
						PSL.OutputWnd(0).Clear
						PSL.Output MsgBak & Format$((j + 1) / .lReferenceNum,"#%")
					End If
				Next j
			End With
		Next i
	ElseIf ShowState = False Then
		Msg = Replace$(Msg,"%s",DataList(FunID).sString)
		k = UBound(IndexList) + 1
		For i = 0 To k - 1
			j = IndexList(i)
			If WriteData(FN,DataList(FunID),j,Mode) = True Then
				WriteDatas = WriteDatas + 1
			End If
			If ShowMsg > 0 Then
				SetTextBoxString ShowMsg,Msg & Format$((i + 1) / k,"#%")
			ElseIf ShowMsg < 0 Then
				PSL.OutputWnd(0).Clear
				PSL.Output Msg & Format$((i + 1) / k,"#%")
			End If
		Next i
	Else
		Msg = Replace$(Msg,"%s",DataList(FunID).sString)
		Dim Dic As Object
		Set Dic = CreateObject("Scripting.Dictionary")
		With DataList(FunID)
			For i = 0 To .lReferenceNum - 1
				If Not Dic.Exists(.Reference(i).lAddress) Then
					Dic.Add(.Reference(i).lAddress,i)
				End If
			Next i
		End With
		k = UBound(IndexList) + 1
		For i = 0 To k - 1
			If Dic.Exists(RefAddList(IndexList(i))) Then
				j = Dic.Item(RefAddList(IndexList(i)))
				If WriteData(FN,DataList(FunID),j,Mode) = True Then
					WriteDatas = WriteDatas + 1
				End If
			End If
			If ShowMsg > 0 Then
				SetTextBoxString ShowMsg,Msg & Format$((i + 1) / k,"#%")
			ElseIf ShowMsg < 0 Then
				PSL.OutputWnd(0).Clear
				PSL.Output Msg & Format$((i + 1) / k,"#%")
			End If
		Next i
		Set Dic = Nothing
	End If
	If WriteDatas > 0 Then
		UnLoadFile(FN,FN.SizeOfFile,Mode)
	Else
		UnLoadFile(FN,0,Mode)
	End If
End Function


'修改单项数据
Private Function WriteData(FN As FILE_IMAGE,strData As STRING_SUB_PROPERTIE,ByVal RefIndex As Long,ByVal Mode As Long) As Boolean
	Dim i As Long,j As Long,k As Long,TempBytes() As Byte
	Select Case strData.DllLang
	Case "MS"
		With strData.Reference(RefIndex)
			If .lOrgCodePage > 0 Then
				If .lOrgCodePage <>  .lUseCodePage Then
					TempBytes = Val2Bytes(.lOrgCodePage,.iCPByteLength,.iCPByteOrder)
					If GetBytes(FN,.iCPByteLength,.lCPAddress,Mode) = TempBytes Then
						TempBytes = Val2Bytes(.lUseCodePage,.iCPByteLength,.iCPByteOrder)
						PutBytes(FN,.lCPAddress,TempBytes,.iCPByteLength,Mode)
						WriteData = True
					End If
				End If
			End If
		End With
	Case "QT"
		With strData.Reference(RefIndex)
			If .sOrgCode <> .sUseCode Then
				TempBytes = HexStr2Bytes(.sOrgCode)
				If GetBytes(FN,4,.lAddress,Mode) = TempBytes Then
					TempBytes = HexStr2Bytes(.sUseCode)
					PutBytes(FN,.lAddress,TempBytes,4,Mode)
					WriteData = True
				End If
			End If
		End With
	End Select
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
			endPos = IIf(Offset + 512 < Max,Offset + 512,Max)
			Set Matches = .Execute(ByteToString(GetBytes(FN,endPos - Offset + 1,Offset,Mode),CP_ISOLATIN1))
			If Matches.Count > 0 Then
				Offset = Offset + Matches(0).FirstIndex
				Exit Do
			End If
			Offset = endPos
		Loop Until endPos >= Max
	End With
	getNotNullByteRegExp = Offset
End Function


'正向跳到指定数量的空字节位置，并返回空字节开始位置，Bit 为最小空字节数
Private Function getNullByte(FN As FILE_IMAGE,ByVal Offset As Long,ByVal Max As Long,ByVal Mode As Long,ByVal Bit As Integer) As Long
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


'转换字符为 Long 整数值
Private Function StrToLong(ByVal mStr As String,Optional ByVal DefaultValue As Long) As Long
	On Error GoTo errHandle
	StrToLong = CLng(mStr)
	Exit Function
	errHandle:
	StrToLong = DefaultValue
End Function


'获取4个字节值 (32 位值,4个字节, -2,147,483,648 到 2,147,483,647)
Private Function GetLong(Source As Variant,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Long
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
Private Function GetInteger(Source As Variant,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Integer
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
Private Function GetByte(Source As Variant,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Byte
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
Private Function GetBytes(Source As Variant,ByVal Length As Long,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Byte()
	On Error GoTo errHandle
	If Offset + Length > Source.SizeOfImage Then GoTo errHandle
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
Private Function PutBytes(Target As FILE_IMAGE,ByVal Offset As Long,Source() As Byte,ByVal Length As Long,Optional ByVal Mode As Long = -1) As Boolean
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


'获取变量字节长度
Private Function GetFileLength(Source As FILE_IMAGE,Optional ByVal Mode As Long = -1) As Long
	Select Case Mode
	Case Is < 0
		GetFileLength = LOF(Source.hFile)
	Case 0
		GetFileLength = Source.SizeOfFile	'UBound(Source.ImageByte) + 1
	Case Else
		GetFileLength = Source.SizeOfFile	'GetFileSize(Source.hFile, 0&)
	End Select
End Function


'检查文件是否已被打开或占用
Private Function IsOpen(ByVal strFilePath As String,Optional ByVal Continue As Long = 2,Optional ByVal WaitTime As Double = 0.5) As Boolean
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


'获取文件及子文件的数据结构信息
Private Function GetPEHeaders(ByVal strFilePath As String,File As FILE_PROPERTIE,ByVal Mode As Long) As Boolean
	Dim i As Long,FN As FILE_IMAGE,TempList() As String,Temp As String
	On Error GoTo ExitFunction
	File.FileSize = FileLen(strFilePath)
	'打开文件
	Mode = LoadFile(strFilePath,FN,0,0,0,Mode)
	If Mode < -1 Then Exit Function
	'获取主文件头
	GetPEHeaders = GetPEHeader(FN,File,Mode)
	If GetPEHeaders = False Then GoTo ExitFunction
	'获取子文件头
	With File.SecList(File.MaxSecIndex)
		If .lSizeOfRawData = 0 Then GoTo ExitFunction
		Temp = ByteToString(GetBytes(FN,.lSizeOfRawData,.lPointerToRawData,Mode),CP_ISOLATIN1)
		TempList = GetVAListRegExp(Temp,"MZ[\x00-\xFF]{64,384}?PE\x00",.lPointerToRawData)
		If CheckArray(TempList) = False Then GoTo ExitFunction
		Dim SubFile As FILE_PROPERTIE
		File.NumberOfSub = UBound(TempList) + 1
		For i = 0 To File.NumberOfSub - 1
			'If GetPEHeader(FN,SubFile,Mode,CLng(TempList(i))) = True Then
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
Private Function GetPEHeader(FN As FILE_IMAGE,File As FILE_PROPERTIE,ByVal Mode As Long,Optional ByVal Offset As Long = -1) As Boolean
	Dim i As Long,j As Long
	Dim tmpDosHeader 		As IMAGE_DOS_HEADER
	Dim tmpFileHeader		As IMAGE_FILE_HEADER
	Dim tmpOptionalHeader32	As IMAGE_OPTIONAL_HEADER32
	Dim tmpOptionalHeader64	As IMAGE_OPTIONAL_HEADER64
	Dim tmpSecHeader() 		As IMAGE_SECTION_HEADER
	ReDim File.SecList(1)				'As SECTION_PROPERTIE
	ReDim File.SecList(0).SubSecList(0)	'As SUB_SECTION_PROPERTIE
	ReDim File.SecList(1).SubSecList(0)	'As SUB_SECTION_PROPERTIE
	ReDim File.DataDirectory(0)			'As SUB_SECTION_PROPERTIE
	ReDim File.CLRList(0)				'As SUB_SECTION_PROPERTIE
	ReDim File.StreamList(0)			'As SUB_SECTION_PROPERTIE
	i = Offset
	If i = -1 Then i = File.FileType
	On Error GoTo ExitFunction
	With File
		'初始化数据
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

		'获取 IMAGE_DOS_HEADERS 结构
		'GetTypeValue(FN,i,tmpDosHeader,Mode)
		Select Case Mode
		Case Is < 0
			Get #FN.hFile, i + 1, tmpDosHeader
		Case 0
			CopyMemory tmpDosHeader, FN.ImageByte(i), Len(tmpDosHeader)
		Case Else
			MoveMemory tmpDosHeader, FN.MappedAddress + i, Len(tmpDosHeader)
		End Select
		If tmpDosHeader.iSignature <> IMAGE_DOS_SIGNATURE Then GoTo ExitFunction

		'读取 IMAGE_FILE_HEADERS 结构
		i = i + tmpDosHeader.lPointerToPEHeader
		'GetTypeValue(FN,i,tmpFileHeader,Mode)
		Select Case Mode
		Case Is < 0
			Get #FN.hFile, i + 1, tmpFileHeader
		Case 0
			CopyMemory tmpFileHeader, FN.ImageByte(i), Len(tmpFileHeader)
		Case Else
			MoveMemory tmpFileHeader, FN.MappedAddress + i, Len(tmpFileHeader)
		End Select
		If tmpFileHeader.lSignature <> IMAGE_NT_SIGNATURE Then GoTo ExitFunction
		'检查是否有文件节结构
		If tmpFileHeader.iNumberOfSections = 0 Then GoTo ExitFunction

		'按 PE 位数读取 IMAGE_OPTIONAL_HEADER 结构，32位PE是&H10B，64位PE是&H20B
		i = i + Len(tmpFileHeader)
		Select Case GetInteger(FN,i,Mode)
		Case IMAGE_NT_OPTIONAL_HDR32_MAGIC	'32位PE文件
			'GetTypeValue(FN,i.tmpOptionalHeader32,Mode)
			Select Case Mode
			Case Is < 0
				Get #FN.hFile, i + 1, tmpOptionalHeader32
			Case 0
				CopyMemory tmpOptionalHeader32, FN.ImageByte(i), Len(tmpOptionalHeader32)
			Case Else
				MoveMemory tmpOptionalHeader32, FN.MappedAddress + i, Len(tmpOptionalHeader32)
			End Select

			'获取文件节结构
			i = i + Len(tmpOptionalHeader32)
			ReDim tmpSecHeader(tmpFileHeader.iNumberOfSections - 1) 'As IMAGE_SECTION_HEADER
			'GetTypeArray(FN,i,tmpSecHeader,Mode)
			Select Case Mode
			Case Is < 0
				Get #FN.hFile, i + 1, tmpSecHeader
			Case 0
				CopyMemory tmpSecHeader(0), FN.ImageByte(i), Len(tmpSecHeader(0)) * tmpFileHeader.iNumberOfSections
			Case Else
				MoveMemory tmpSecHeader(0), FN.MappedAddress + i, Len(tmpSecHeader(0)) * tmpFileHeader.iNumberOfSections
			End Select

			'记录区段地址
			ReDim File.SecList(tmpFileHeader.iNumberOfSections) 'As SECTION_PROPERTIE
			j = 0
			For i = 0 To tmpFileHeader.iNumberOfSections - 1
				.SecList(i).sName = Replace$(StrConv$(tmpSecHeader(i).sName,vbUnicode),vbNullChar,"")
				.SecList(i).lPointerToRawData = tmpSecHeader(i).lPointerToRawData
				.SecList(i).lSizeOfRawData = tmpSecHeader(i).lSizeOfRawData
				.SecList(i).lVirtualAddress = tmpSecHeader(i).lVirtualAddress
				.SecList(i).lVirtualSize = tmpSecHeader(i).lVirtualSize
				.SecList(i).SubSecs = 0
				If .SecList(i).lSizeOfRawData = 0 Then j = j + 1
			Next i
			If j = tmpFileHeader.iNumberOfSections Then GoTo ExitFunction

			'记录 DataDirectory 地址
			ReDim File.DataDirectory(15)			'As SUB_SECTION_PROPERTIE
			For i = 0 To 15
				.DataDirectory(i).lPointerToRawData = RvaToOffset(File,tmpOptionalHeader32.DataDirectory(i).lVirtualAddress)
				.DataDirectory(i).lSizeOfRawData = tmpOptionalHeader32.DataDirectory(i).lSize
				.DataDirectory(i).lVirtualAddress = tmpOptionalHeader32.DataDirectory(i).lVirtualAddress
				.DataDirectory(i).lVirtualSize = tmpOptionalHeader32.DataDirectory(i).lSize
			Next i
			.Magic = "PE32"
			.FileAlign = tmpOptionalHeader32.lFileAlignment
			.SecAlign = tmpOptionalHeader32.lSectionAlignment
			.ImageBase = tmpOptionalHeader32.lImageBase
			'记录数据目录数
			.DataDirs = 16
		Case IMAGE_NT_OPTIONAL_HDR64_MAGIC	'64位PE文件
			'GetTypeValue(FN,i,tmpOptionalHeader64,Mode)
			Select Case Mode
			Case Is < 0
				Get #FN.hFile, i + 1, tmpOptionalHeader64
			Case 0
				CopyMemory tmpOptionalHeader64, FN.ImageByte(i), Len(tmpOptionalHeader64)
			Case Else
				MoveMemory tmpOptionalHeader64, FN.MappedAddress + i, Len(tmpOptionalHeader64)
			End Select

			'获取文件节结构
			i = i + Len(tmpOptionalHeader64)
			ReDim tmpSecHeader(tmpFileHeader.iNumberOfSections - 1) 'As IMAGE_SECTION_HEADER
			'GetTypeArray(FN,i,tmpSecHeader,Mode)
			Select Case Mode
			Case Is < 0
				Get #FN.hFile, i + 1, tmpSecHeader
			Case 0
				CopyMemory tmpSecHeader(0), FN.ImageByte(i), Len(tmpSecHeader(0)) * tmpFileHeader.iNumberOfSections
			Case Else
				MoveMemory tmpSecHeader(0), FN.MappedAddress + i, Len(tmpSecHeader(0)) * tmpFileHeader.iNumberOfSections
			End Select

			'记录区段地址
			ReDim File.SecList(tmpFileHeader.iNumberOfSections) 'As SECTION_PROPERTIE
			j = 0
			For i = 0 To tmpFileHeader.iNumberOfSections - 1
				.SecList(i).sName = Replace$(StrConv$(tmpSecHeader(i).sName,vbUnicode),vbNullChar,"")
				.SecList(i).lPointerToRawData = tmpSecHeader(i).lPointerToRawData
				.SecList(i).lSizeOfRawData = tmpSecHeader(i).lSizeOfRawData
				.SecList(i).lVirtualAddress = tmpSecHeader(i).lVirtualAddress
				.SecList(i).lVirtualSize = tmpSecHeader(i).lVirtualSize
				.SecList(i).SubSecs = 0
				If .SecList(i).lSizeOfRawData = 0 Then j = j + 1
			Next i
			If j = tmpFileHeader.iNumberOfSections Then GoTo ExitFunction

			'记录 DataDirectory 地址
			ReDim File.DataDirectory(15)			'As SUB_SECTION_PROPERTIE
			For i = 0 To 15
				.DataDirectory(i).lPointerToRawData = RvaToOffset(File,tmpOptionalHeader64.DataDirectory(i).lVirtualAddress)
				.DataDirectory(i).lSizeOfRawData = tmpOptionalHeader64.DataDirectory(i).lSize
				.DataDirectory(i).lVirtualAddress = tmpOptionalHeader64.DataDirectory(i).lVirtualAddress
				.DataDirectory(i).lVirtualSize = tmpOptionalHeader64.DataDirectory(i).lSize
			Next i
			.Magic = "PE64"
			.FileAlign = tmpOptionalHeader64.lFileAlignment
			.SecAlign = tmpOptionalHeader64.lSectionAlignment
			.ImageBase = tmpOptionalHeader64.dImageBase
			'记录数据目录数
			.DataDirs = 16
		Case Else
			GoTo ExitFunction
		End Select

		'获取文件节最大索引号、最小和最大偏移地址所在节的索引号
		.MaxSecIndex = tmpFileHeader.iNumberOfSections
		Call GetSectionID(File,.MinSecID,.MaxSecID,False)
		.LangType = tmpDosHeader.iLastPageBytes

		'获取 .NET 各种头结构
		i = Offset
		If i = -1 Then i = File.FileType
		If GetNETHeader(FN,File,Mode,i) = True Then .LangType = NET_FILE_SIGNATURE

		'获取隐藏节信息
		.SecList(.MaxSecIndex).lPointerToRawData = .SecList(.MaxSecID).lPointerToRawData + .SecList(.MaxSecID).lSizeOfRawData
		.SecList(.MaxSecIndex).lSizeOfRawData = GetFileLength(FN,Mode) - .SecList(.MaxSecIndex).lPointerToRawData
		.SecList(.MaxSecIndex).lVirtualAddress = .SecList(.MaxSecID).lVirtualAddress + .SecList(.MaxSecID).lVirtualSize
		.SecList(.MaxSecIndex).lVirtualSize = .SecList(.MaxSecIndex).lSizeOfRawData
	End With

	'记录主程序的各种头数据
	If Offset = -1 Then
		DosHeader = tmpDosHeader
		FileHeader = tmpFileHeader
		OptionalHeader32 = tmpOptionalHeader32
		OptionalHeader64 = tmpOptionalHeader64
		SecHeader = tmpSecHeader
	End If

	'标记成功
	GetPEHeader = True
	Exit Function

	ExitFunction:
	ReDim File.SecList(1)			'As SECTION_PROPERTIE
	ReDim File.DataDirectory(0)		'As SUB_SECTION_PROPERTIE
	ReDim File.CLRList(0)			'As SUB_SECTION_PROPERTIE
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


'获取 .NET 各种头结构
Private Function GetNETHeader(FN As FILE_IMAGE,File As FILE_PROPERTIE,ByVal Mode As Long,Optional ByVal Offset As Long) As Boolean
	Dim i As Long,Length As Long,dwOffset As Long
	Dim CLRHeader			As IMAGE_CLR20_HEADER
	Dim MetaDataHeader		As IMAGE_METADATA_HEADER
	On Error GoTo ExitFunction
	With File
		'转换第15个数据目录的相对虚拟地址转偏移地址
		dwOffset = RvaToOffset(File,.DataDirectory(14).lVirtualAddress)
		If dwOffset = 0 Then Exit Function
		'读取 CLR 结构
		'GetTypeValue(FN,Offset + dwOffset,CLRHeader,Mode)
		Select Case Mode
		Case Is < 0
			Get #FN.hFile, Offset + dwOffset + 1, CLRHeader
		Case 0
			CopyMemory CLRHeader, FN.ImageByte(Offset + dwOffset), Len(CLRHeader)
		Case Else
			MoveMemory CLRHeader, Offset + FN.MappedAddress + dwOffset, Len(CLRHeader)
		End Select
		'转换 CLR 中的 MetaData 相对虚拟地址转偏移地址
		dwOffset = RvaToOffset(File,CLRHeader.METADATA.lVirtualAddress)
		If dwOffset = 0 Then Exit Function
	End With

	'检查是否为 .NET 程序文件
	With MetaDataHeader
		.lSignature = GetLong(FN,Offset + dwOffset,Mode)
		If .lSignature <> NET_FILE_SIGNATURE Then Exit Function
		'获取 MetaDataHeader.Version 的字节长度
		.lLength = GetLong(FN,Offset + dwOffset + 12,Mode)
		'按 4 个字节对齐 MetaDataHeader.Version 的字节长度
		.lLength = Alignment(.lLength,4,1)
		'读取 METADATA 结构
		.iMajorVersion = GetInteger(FN,Offset + dwOffset + 4,Mode)
		.iMinorVersion = GetInteger(FN,Offset + dwOffset + 6,Mode)
		.lExtraData = GetLong(FN,Offset + dwOffset + 8,Mode)
		.Version = GetBytes(FN,.lLength,Offset + dwOffset + 16,Mode)
		.fFlags = GetInteger(FN,Offset + dwOffset + 16 + .lLength,Mode)
		.iStreams = GetInteger(FN,Offset + dwOffset + 18 + .lLength,Mode)

		'获取各个流的文件头结构
		If .iStreams > 0 Then
			ReDim StreamHeader(.iStreams - 1) As IMAGE_STREAM_HEADER
			dwOffset = dwOffset + 20 + .lLength
			For i = 0 To .iStreams - 1
				'lOffset 相对于 Metadata Root，实际 RVA = .CLRHeader.MetaData.lVirtualAddress + .StreamHeader(i).lOffset
				StreamHeader(i).RWA = dwOffset + 0
				StreamHeader(i).lOffset = GetLong(FN,Offset + dwOffset + 0,Mode)
				StreamHeader(i).lSize = GetLong(FN,Offset + dwOffset + 4,Mode)
				'获取 .StreamHeader.rcName 的字节长度
				dwOffset = dwOffset + 8
				Length = getNullByte(FN,Offset + dwOffset,Offset + dwOffset + 16,Mode,1) - dwOffset + 1
				'按 4 个字节对齐 .StreamHeader.rcName 的字节长度
				Length = Alignment(Length,4,1)
				StreamHeader(i).rcName = GetBytes(FN,Length,Offset + dwOffset,Mode)
				dwOffset = dwOffset + Length
			Next i
		End If
	End With

	'转换.NET 各种头结构的相对虚拟地址为物理地址
	ReDim File.CLRList(6)	'As SUB_SECTION_PROPERTIE
	With File
		.CLRList(0).lPointerToRawData = RvaToOffset(File,CLRHeader.METADATA.lVirtualAddress)
		.CLRList(1).lPointerToRawData = RvaToOffset(File,CLRHeader.Resources.lVirtualAddress)
		.CLRList(2).lPointerToRawData = RvaToOffset(File,CLRHeader.StrongNameSignature.lVirtualAddress)
		.CLRList(3).lPointerToRawData = RvaToOffset(File,CLRHeader.CodeManagerTable.lVirtualAddress)
		.CLRList(4).lPointerToRawData = RvaToOffset(File,CLRHeader.VTableFixups.lVirtualAddress)
		.CLRList(5).lPointerToRawData = RvaToOffset(File,CLRHeader.ExportAddressTableJumps.lVirtualAddress)
		.CLRList(6).lPointerToRawData = RvaToOffset(File,CLRHeader.ManagedNativeHeader.lVirtualAddress)
		.CLRList(0).lSizeOfRawData = CLRHeader.METADATA.lSize
		.CLRList(1).lSizeOfRawData = CLRHeader.Resources.lSize
		.CLRList(2).lSizeOfRawData = CLRHeader.StrongNameSignature.lSize
		.CLRList(3).lSizeOfRawData = CLRHeader.CodeManagerTable.lSize
		.CLRList(4).lSizeOfRawData = CLRHeader.VTableFixups.lSize
		.CLRList(5).lSizeOfRawData = CLRHeader.ExportAddressTableJumps.lSize
		.CLRList(6).lSizeOfRawData = CLRHeader.ManagedNativeHeader.lSize

		.CLRList(0).lVirtualAddress = CLRHeader.METADATA.lVirtualAddress
		.CLRList(1).lVirtualAddress = CLRHeader.Resources.lVirtualAddress
		.CLRList(2).lVirtualAddress = CLRHeader.StrongNameSignature.lVirtualAddress
		.CLRList(3).lVirtualAddress = CLRHeader.CodeManagerTable.lVirtualAddress
		.CLRList(4).lVirtualAddress = CLRHeader.VTableFixups.lVirtualAddress
		.CLRList(5).lVirtualAddress = CLRHeader.ExportAddressTableJumps.lVirtualAddress
		.CLRList(6).lVirtualAddress = CLRHeader.ManagedNativeHeader.lVirtualAddress
		.CLRList(0).lVirtualSize = CLRHeader.METADATA.lSize
		.CLRList(1).lVirtualSize = CLRHeader.Resources.lSize
		.CLRList(2).lVirtualSize = CLRHeader.StrongNameSignature.lSize
		.CLRList(3).lVirtualSize = CLRHeader.CodeManagerTable.lSize
		.CLRList(4).lVirtualSize = CLRHeader.VTableFixups.lSize
		.CLRList(5).lVirtualSize = CLRHeader.ExportAddressTableJumps.lSize
		.CLRList(6).lVirtualSize = CLRHeader.ManagedNativeHeader.lSize

		.NetStreams = MetaDataHeader.iStreams
		If .NetStreams > 0 Then
			ReDim File.StreamList(.NetStreams - 1) 'As SUB_SECTION_PROPERTIE
			For i = 0 To .NetStreams - 1
				.StreamList(i).lPointerToRawData = .CLRList(0).lPointerToRawData + StreamHeader(i).lOffset
				.StreamList(i).lSizeOfRawData = StreamHeader(i).lSize
				.StreamList(i).lVirtualAddress = .CLRList(0).lVirtualAddress + StreamHeader(i).lOffset
				.StreamList(i).lVirtualSize = StreamHeader(i).lSize
				.StreamList(i).sName = Replace$(StrConv$(StreamHeader(i).rcName,vbUnicode),vbNullChar,"")
				If UCase$(.StreamList(i).sName) = "#US" Then .USStreamID = i
			Next i
		End If
	End With
	'标记成功
	GetNETHeader = True
	ExitFunction:
End Function


'获取输入表 (读文件版)
Private Function GetImportTable(File As FILE_PROPERTIE,DataList() As STRING_SUB_PROPERTIE,ByVal Mode As Long) As Boolean
	Dim i As Integer,j As Integer,k As Integer,n As Long
	Dim RvaAddrLength As Integer,Offset As Long,dwOffset As Long,Temp As String
	Dim FN As FILE_IMAGE,TempBytes() As Byte

	Dim PEImport As IMAGE_IMPORT_DESCRIPTOR
	Dim pImportInfo As IMPORT_INFO
	'Dim pIMPTable As IMAGE_IMPORT_BY_NAME
	Dim pIMPTunk As IMAGE_THUNK_DATA
	Dim p32IMPTunk As IMAGE_THUNK_DATA32
	Dim p64IMPTunk As IMAGE_THUNK_DATA64

	ReDim QtFunNameList(0) As String
	QtFunNameList = ReSplit("toAscii;toLatin1;toUtf8;toStdString;toLocal8Bit;toUnicode;" & _
					"fromAscii;fromLatin1;fromUtf8;fromStdString;fromLocal8Bit;fromUnicode",";")

	On Error GoTo ExitFunction
	With File
		'检查是否有输入表结构
		If .DataDirs = 0 Then Exit Function
		If .DataDirectory(1).lVirtualAddress = 0 Then Exit Function

		'打开文件
		Mode = LoadFile(.FilePath,FN,0,0,0,Mode)
		If Mode < -1 Then
			UnLoadFile(FN,0,Mode)
			Exit Function
		End If

		'初始化
		RvaAddrLength = IIf(.Magic = "PE32",4,8)
		'获取输入表
		j = 0
		Do
			'获取IMAGE_THUNK_DATA32结构在文件中的真实地址
			dwOffset = .DataDirectory(1).lPointerToRawData + j * Len(PEImport)
			'GetTypeValue(FN,dwOffset.PEImport,Mode)
			Select Case Mode
			Case Is < 0
				Get #FN.hFile, dwOffset + 1, PEImport
			Case 0
				CopyMemory PEImport, FN.ImageByte(dwOffset), Len(PEImport)
			Case Else
				MoveMemory PEImport, FN.MappedAddress + dwOffset, Len(PEImport)
			End Select
			If PEImport.FirstThunk = 0 And PEImport.OriginalFirstThunk = 0 Then Exit Do

			'获取模块名称
			dwOffset = RvaToOffset(File, PEImport.pName)
			Temp = StrConv$(GetBytes(FN,130,dwOffset,Mode), vbUnicode)
			Temp = Mid$(Temp, 1, InStr(Temp & vbNullChar, vbNullChar) - 1)
			ReDim Preserve pImportInfo.pDetailInfo(j) 'As IMPORT_DETAIL_INFO
			pImportInfo.pDetailInfo(j).strDllName = Temp
			'PSL.Output "输入表 DLL 模块为：" & pImportInfo.pDetailInfo(j).strDllName

			'遍历输入表
			i = 0
			Do
				'If PEImport.OriginalFirstThunk <> 0 Then
				'	dwOffset = RvaToOffset(File, PEImport.OriginalFirstThunk + i * RvaAddrLength)
				'	If dwOffset = 0 Then Exit Do
				'	offset = dwOffset
				'	'GetTypeValue(FN,dwOffset.pIMPTunk,Mode)
				'	Select Case Mode
				'	Case Is < 0
				'		Get #FN.hFile, dwOffset + 1, pIMPTunk
				'	Case 0
				'		CopyMemory pIMPTunk, FN.ImageByte(dwOffset), Len(pIMPTunk)
				'	Case Else
				'		MoveMemory pIMPTunk, FN.MappedAddress + dwOffset, Len(pIMPTunk)
				'	End Select
				'	If pIMPTunk.AddressOfData = 0 Then Exit Do
				'	dwOffset = RvaToOffset(File, pIMPTunk.AddressOfData)
				'Else
					dwOffset = RvaToOffset(File, PEImport.FirstThunk + i * RvaAddrLength)
					If dwOffset = 0 Then Exit Do
					Offset = dwOffset
					If RvaAddrLength = 4 Then
						'GetTypeValue(FN,dwOffset.p32IMPTunk,Mode)
						Select Case Mode
						Case Is < 0
							Get #FN.hFile, dwOffset + 1, p32IMPTunk
						Case 0
							CopyMemory p32IMPTunk, FN.ImageByte(dwOffset), Len(p32IMPTunk)
						Case Else
							MoveMemory p32IMPTunk, FN.MappedAddress + dwOffset, Len(p32IMPTunk)
						End Select
						If p32IMPTunk.pForwarderString = 0 Then Exit Do
						dwOffset = RvaToOffset(File, p32IMPTunk.pForwarderString)
						'PSL.Output  dwOffset & "- 1"
					Else
						'GetTypeValue(FN,dwOffset.p64IMPTunk,Mode)
						Select Case Mode
						Case Is < 0
							Get #FN.hFile, dwOffset + 1, p64IMPTunk
						Case 0
							CopyMemory p64IMPTunk, FN.ImageByte(dwOffset), Len(p64IMPTunk)
						Case Else
							MoveMemory p64IMPTunk, FN.MappedAddress + dwOffset, Len(p64IMPTunk)
						End Select
						If p64IMPTunk.pForwarderString1 = 0 Then Exit Do
						dwOffset = RvaToOffset(File, p64IMPTunk.pForwarderString1)
						'PSL.Output  dwOffset & "- 2"
					End If
				'End If

				ReDim Preserve pImportInfo.pDetailInfo(j).strFuns(i)
				If dwOffset <= 0 Then
					pImportInfo.pDetailInfo(j).strFuns(i) = "空"
				Else
					'GetTypeValue(FN,dwOffset.pIMPTable,Mode)
					'Select Case Mode
					'Case Is < 0
					'	Get #FN.hFile, dwOffset + 1, pIMPTable
					'Case 0
					'	CopyMemory pIMPTable, FN.ImageByte(dwOffset), Len(pIMPTable)
					'Case Else
					'	MoveMemory pIMPTable, FN.MappedAddress + dwOffset, Len(pIMPTable)
					'End Select
					'Temp = StrConv$(pIMPTable.pName, vbUnicode)
					TempBytes = GetBytes(FN,300,dwOffset + 2,Mode)
					Temp = StrConv$(TempBytes, vbUnicode)
					Temp = Mid$(Temp, 1, InStr(Temp & vbNullChar, vbNullChar) - 1)
					pImportInfo.pDetailInfo(j).strFuns(i) = Temp
					Select Case Temp
					Case "MultiByteToWideChar", "WideCharToMultiByte"
						ReDim Preserve DataList(n) 'As STRING_SUB_PROPERTIE
						With DataList(n)
							.DllName = pImportInfo.pDetailInfo(j).strDllName
							.DllLang = "MS"
							.sString = Temp
							.lStartAddress = Offset
							.lFunNameAddress = dwOffset
							.inSectionID = SkipSection(File,.lStartAddress,0,0)
							.inSubSecID = SkipSubSection(File.SecList(.inSectionID),.lStartAddress,0,0)
						End With
						n = n + 1
					Case Else
						For k = 0 To UBound(QtFunNameList)
							If InStr(Temp,QtFunNameList(k)) Then
								ReDim Preserve DataList(n) 'As STRING_SUB_PROPERTIE
								With DataList(n)
									.DllName = pImportInfo.pDetailInfo(j).strDllName
									.DllLang = "QT"
									.sString = Temp
									.lStartAddress = Offset
									.lFunNameAddress = dwOffset
									.inSectionID = SkipSection(File,.lStartAddress,0,0)
									.inSubSecID = SkipSubSection(File.SecList(.inSectionID),.lStartAddress,0,0)
								End With
								n = n + 1
								Exit For
							End If
						Next k
					End Select
				End If
				'PSL.Output "       ----函数为：" & pImportInfo.pDetailInfo(j).strFuns(i)
				i = i + 1
			Loop
			j = j + 1
		Loop
	End With
	If n > 0 Then GetImportTable = True
	ExitFunction:
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


'获取文件及子文件的数据结构信息
Private Function GetMacHeaders(ByVal strFilePath As String,File As FILE_PROPERTIE,ByVal Mode As Long) As Boolean
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
	ReDim File.DataDirectory(0)		'As SECTION_PROPERTIE
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


'映射文件
'MapSize = 0 按文件初始时的大小映射，否则按指定大小映射
'ReadOnly = 0 只读方式，否则读写方式
'SizeOfFile = 0 获取文件初始时的大小，否则不获取
'IsPE = 0 按一般文件映射，否则按 PE 文件映射(每个节对齐)
Private Function MapFile(ByVal strFilePath As String,LoadedImage As FILE_IMAGE,ByVal MapSize As Long, _
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


'获取文件节索引号
'Mode = 0 检查偏移地址(不包括隐藏节)
'Mode = 1 检查偏移地址(包括隐藏节)
'Mode = 2 检查相对虚拟地址(不包括隐藏节)
'Mode = 3 检查相对虚拟地址(包括隐藏节)
'返回文件节索引号、MinVal、MaxVal 值
Private Function SkipSection(File As FILE_PROPERTIE,ByVal Offset As Long,MinVal As Long,MaxVal As Long,Optional ByVal Mode As Long) As Long
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
Private Function SkipSubSection(Sec As SECTION_PROPERTIE,ByVal Offset As Long,MinVal As Long,MaxVal As Long,Optional ByVal Mode As Boolean) As Long
	Dim i As Integer
	SkipSubSection = -1
	If Sec.SubSecs = 0 Then Exit Function
	MinVal = 0: MaxVal = 0
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
Private Function SkipHeader(File As FILE_PROPERTIE,RVA As Long,Optional SkipVal As Long,Optional ByVal Mode As Long,Optional ByVal fType As Long) As Long
	Dim i As Integer,j As Integer,endPos As Long
	SkipHeader = -1
	If File.DataDirs = 0 Then Exit Function
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
				List(j).lVirtualAddress  = .StreamList(i).lSizeOfRawData
				List(j).lSize = .StreamList(i).lSizeOfRawData
				If i = .USStreamID And fType > 0 Then
					fType = 0
					endPos = .StreamList(i).lPointerToRawData + .StreamList(i).lSizeOfRawData - 1
					If RVA >= .StreamList(i).lPointerToRawData And RVA <= endPos Then
						fType = RVA
					ElseIf Mode = 1 And RVA < endPos Then
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
				endPos = .lVirtualAddress + .lSize - 1		'最大地址
				Select Case Mode
				Case 0
					If SkipHeader < 0 Then
						If RVA >= .lVirtualAddress And RVA <= endPos Then
							SkipHeader = i
						End If
					Else
						If fType = 0 Then Exit Function
						If .lVirtualAddress >= List(SkipHeader).lVirtualAddress And _
							endPos < List(SkipHeader).lVirtualAddress + List(SkipHeader).lSize Then
							If .lVirtualAddress <= fType And endPos >= fType Then
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
					ElseIf RVA <= endPos Then
						SkipHeader = i
						RVA = endPos + 1
						If SkipVal < endPos Then SkipVal = endPos
					ElseIf SkipHeader > -1 And fType > 0 Then
						If .lVirtualAddress >= List(SkipHeader).lVirtualAddress And _
							endPos < List(SkipHeader).lVirtualAddress + List(SkipHeader).lSize Then
							If .lVirtualAddress <= fType And endPos >= fType Then
								SkipHeader = -1
								RVA = fType
								SkipVal = endPos
								Exit Function
							End If
						End If
					End If
				Case Else
					If endPos < RVA Then
						If RVA < SkipVal Then
							SkipVal = endPos
						ElseIf endPos > SkipVal Then
							SkipVal = endPos
						End If
					ElseIf RVA >= .lVirtualAddress Then
						SkipHeader = i
						RVA = .lVirtualAddress - 1
						If SkipVal > .lVirtualAddress Then SkipVal = .lVirtualAddress
					ElseIf SkipHeader > -1 And fType > 0 Then
						If .lVirtualAddress >= List(SkipHeader).lVirtualAddress And _
							endPos < List(SkipHeader).lVirtualAddress + List(SkipHeader).lSize Then
							If .lVirtualAddress <= fType And endPos >= fType Then
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
Private Function GetSectionID(File As FILE_PROPERTIE,MinID As Integer,MaxID As Integer,ByVal Mode As Boolean) As Long
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


'加载文件
'ImageSize = 0 按文件的初始大小打开，否则按指定大小打开
'ReadOnly = 0 按只读方式打开，否则读写方式打开
'ImageByte = 0 不获取字节数组只初始化(缓存方式获取所有字节)，否则按 ImageByte 指定大小获取
'Mode < 0 直接方式，Mode = 0 缓存方式，Mode > 0 映射方式
'IsPE = 0 按一般文件映射，否则按 PE 文件映射(每个节对齐)
'LoadFile = -2 打开失败，否则实际打开方式
'LoadedImage 打开文件后获取的数据
Private Function LoadFile(ByVal strFilePath As String,LoadedImage As FILE_IMAGE,ByVal ImageSize As Long,ByVal ReadOnly As Long, _
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
Private Function UnLoadFile(LoadedImage As FILE_IMAGE,ByVal SizeOfFile As Long,ByVal Mode As Long) As Boolean
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


'查找引用代码和引用列表
'fType < 0 按现有的引用地址列表，查找翻译的引用代码列表，fType 为原来的翻译开始地址
'fType = 0 查找来源的引用列表和引用代码
'fType = 1 查找翻译的引用列表和引用代码，如果 RefAdds 为空，则按照原来引用地址计算引用代码
'fType = 2 查找翻译的引用列表和引用代码，如果 RefAdds 为空，则初始化，清空引用列表
'fType > 2 初始化，清空翻译引用列表和引用代码
'虚拟地址(VA) = StartPos + ImageBase + VRK
Private Function GetVARefList(File As FILE_PROPERTIE,FN As Variant,strData As STRING_SUB_PROPERTIE, _
	ByVal RefAdds As String,ByVal fType As Long,ByVal Mode As Long,Optional ByVal ShowMsg As Long) As Long
	Dim i As Long,j As Long,k As Long,m As Long,n As Long
	Dim RVA As Long,VRK As Long,MaxPos As Long,RSize As Long,SkipVal As Long
	Dim Msg As String,TempList() As String
	On Error GoTo ExitFunction
	If File.Magic = "" Then GoTo ExitFunction
	If fType > 2 Then GoTo ExitFunction
	With strData
	If fType > -1 Then
		If .inSectionID < 0 Then .inSectionID = SkipSection(File,.lStartAddress,0,0)
		If .inSectionID < 0 Then GoTo ExitFunction
		If File.SecList(.inSectionID).SubSecs > 0 Then
			If .inSubSecID < 0 Then .inSubSecID = SkipSubSection(File.SecList(.inSectionID),.lStartAddress,0,0)
		End If
		If ShowMsg > 0 Then
			Msg = GetTextBoxString(ShowMsg) & " "
		ElseIf ShowMsg < 0 Then
			ReDim TempList(PSL.OutputWnd(0).LineCount - 1) As String
			For i = 1 To PSL.OutputWnd(0).LineCount
				TempList(i - 1) = PSL.OutputWnd(0).Text(i)
			Next i
			Msg = StrListJoin(TempList,vbCrLf) & " "
		End If
		If fType > 0 Then RefAdds = Trim$(RefAdds)
	End If
	Select Case File.Magic
	Case "PE64","NET64","MAC64"
		'按原来翻译开始地址的引用代码获取新地址的引用代码列表
		If fType < 0 Then
			'If RefAdds <> "" Then Call getRefList(strData,ReSplit(RefAdds,RefJoinStr),TempList,True)
			If .lReferenceNum = 0 Then GoTo ExitFunction
			i = .Reference(0).inSecID
			VRK = File.SecList(i).lVirtualAddress - File.SecList(i).lPointerToRawData
			RVA = Val("&H" & ReverseHexCode(.Reference(0).sUseCode,8)) + .Reference(0).lAddress + VRK + 4
			RVA = RVA + fType + .lStartAddress
			For i = 0 To .lReferenceNum - 1
				j = .Reference(i).inSecID
				VRK = File.SecList(j).lVirtualAddress - File.SecList(j).lPointerToRawData
				.Reference(i).sUseCode = ReverseHexCode(Hex$(RVA - (.Reference(i).lAddress + VRK + 4)),8)
				'.Reference(i).sUseCode = Byte2Hex(Val2Bytes(RVA - (.Reference(i).lAddress + VRK + 4),4),0,3)
			Next i
			.GetRefState = 1
			Exit Function
		End If
		If .inSectionID > File.MaxSecIndex - 1 Then GoTo ExitFunction
		If fType = 0 Then
			'获取过引用的退出程序
			'If .GetRefState > 0 Then Exit Function
			'If SkipHeader(File,strData.lStartAddress,0,0) > -1 Then GoTo ExitFunction
			'获取字串的虚拟地址
			With File.SecList(.inSectionID)
				If strData.lStartAddress >= .lPointerToRawData And strData.lStartAddress < .lPointerToRawData + .lSizeOfRawData Then
					RVA = strData.lStartAddress + .lVirtualAddress - .lPointerToRawData
				Else
					GoTo ExitFunction
				End If
			End With
			MaxPos = strData.lStartAddress - 4
			.lReferenceNum = 0
			ReDim strData.Reference(0) 'As REFERENCE_PROPERTIE
			For j = 0 To IIf(.inSectionID = 0,0,.inSectionID - 1)
				With File.SecList(j)
					i = .lPointerToRawData
					RSize = i + .lSizeOfRawData - 4
					VRK = RVA - (.lVirtualAddress - .lPointerToRawData) - 4
				End With
				If RSize > MaxPos Then RSize = MaxPos
				'SkipVal = i - 1
				SkipVal = RSize
				Do While i < RSize
					'排除某些数据目录区段和 .NET 文件数据区段
					'If i > SkipVal Then
					'	k = i: m = SkipHeader(File,k,SkipVal,1)
					'	If m = 2 Or m = 4 Or m = 5 Or m > 15 Then i = k
					'	If i > SkipVal Or SkipVal > RSize Then SkipVal = RSize + 1
					'	If i > RSize Then Exit Do
					'End If
					i = i + GetVAListPE64(FN,strData,n,j,VRK,i,SkipVal,Mode) + 1
					If ShowMsg > 0 Then
						SetTextBoxString ShowMsg,Msg & Format$(i / MaxPos,"#%")
					ElseIf ShowMsg < 0 Then
						PSL.OutputWnd(0).Clear
						PSL.Output Msg & Format$(i / MaxPos,"#%")
					End If
				Loop
			Next j
			If .lReferenceNum > 0 Then
				ReDim Preserve strData.Reference(.lReferenceNum - 1) 'As REFERENCE_PROPERTIE
				'If TagType > 1 Then Call GetCustomStrType(FN,strData,TypeList,Mode,-1)
			End If
			.GetRefState = 1
		ElseIf RefAdds <> "" Or (.lReferenceNum > 0 And fType < 2) Then
			'If RefAdds <> "" Then Call GetRefList(strData,ReSplit(RefAdds,RefJoinStr),TempList,True)
			If .lReferenceNum > 0 Then
				For i = 0 To .lReferenceNum - 1
					If strData.lStartAddress <> .Reference(i).lUseFuncAddress Then
						j = .Reference(i).UseFuncInSecID
						RVA = .Reference(i).lUseFuncAddress + File.SecList(j).lVirtualAddress - File.SecList(j).lPointerToRawData
						j = .Reference(i).inSecID
						VRK = File.SecList(j).lVirtualAddress - File.SecList(j).lPointerToRawData
						.Reference(i).sUseCode = ReverseHexCode(Hex$(RVA - (.Reference(i).lAddress + VRK + 4)),8)
						'.Reference(i).sUseCode = Byte2Hex(Val2Bytes(RVA - (.Reference(i).lAddress + VRK + 4),4),0,3)
					End If
					If ShowMsg > 0 Then
						SetTextBoxString ShowMsg,Msg & Format$(i / .lReferenceNum,"#%")
					ElseIf ShowMsg < 0 Then
						PSL.OutputWnd(0).Clear
						PSL.Output Msg & Format$(i / .lReferenceNum,"#%")
					End If
				Next i
				.GetRefState = 1
			End If
		Else
			GoTo ExitFunction
		End If
	Case Else
		'按原来翻译开始地址的引用代码获取新地址的引用代码列表
		If fType < 0 Then
			'If RefAdds <> "" Then Call getRefList(strData,ReSplit(RefAdds,RefJoinStr),TempList,True)
			If .lReferenceNum = 0 Then GoTo ExitFunction
			VRK = Val("&H" & ReverseHexCode(.Reference(0).sUseCode,8)) + fType
			.Reference(0).sUseCode = ReverseHexCode(Hex$(.lStartAddress + VRK),8)
			'.Reference(0).sUseCode = Byte2Hex(Val2Bytes(.lStartAddress + VRK,4),0,3)
			For i = 0 To .lReferenceNum - 1
				.Reference(i).sUseCode = .Reference(0).sUseCode
			Next i
			.GetRefState = 1
			Exit Function
		End If
		If .inSectionID > File.MaxSecIndex - 1 Then GoTo ExitFunction
		'获取引用地址及引用代码列表
		If fType = 0 Then
			'获取过引用的退出程序
			'If .GetRefState > 0 Then Exit Function
			'If SkipHeader(File,strData.lStartAddress,0,0) > -1 Then GoTo ExitFunction
			'获取字串的虚拟地址
			If File.SecList(.inSectionID).SubSecs > 0 Then
				With File.SecList(.inSectionID).SubSecList(.inSubSecID)
					If strData.lStartAddress >= .lPointerToRawData And strData.lStartAddress < .lPointerToRawData + .lSizeOfRawData Then
						VRK = .lVirtualAddress - .lPointerToRawData + File.ImageBase
					Else
						GoTo ExitFunction
					End If
				End With
			Else
				With File.SecList(.inSectionID)
					If strData.lStartAddress >= .lPointerToRawData And strData.lStartAddress < .lPointerToRawData + .lSizeOfRawData Then
						VRK = .lVirtualAddress - .lPointerToRawData + File.ImageBase
					Else
						GoTo ExitFunction
					End If
				End With
			End If
			.lReferenceNum = 0
			ReDim strData.Reference(0) 'As REFERENCE_PROPERTIE
			.Reference(0).sOrgCode = ReverseHexCode(Hex$(.lStartAddress + VRK),8)
			With File
				SkipVal = .SecList(.MinSecID).lPointerToRawData
				If RefAdds = "" Then
					If .DataDirs > 0 Then
						If .DataDirectory(2).lPointerToRawData > 0 Then
							If SkipSection(File,.DataDirectory(2).lPointerToRawData,0,0) > -1 Then
								RSize = .DataDirectory(2).lPointerToRawData - 1
							Else
								RSize = .SecList(.MaxSecID).lPointerToRawData + .SecList(.MaxSecID).lSizeOfRawData - 1
							End If
						Else
							RSize = .SecList(.MaxSecID).lPointerToRawData + .SecList(.MaxSecID).lSizeOfRawData - 1
						End If
					Else
						RSize = .SecList(.MaxSecID).lPointerToRawData + .SecList(.MaxSecID).lSizeOfRawData - 1
					End If
					RefAdds = ByteToString(GetBytes(FN,RSize - SkipVal + 1,SkipVal,Mode),CP_ISOLATIN1)
				End If
			End With
			TempList = GetVAListRegExp(RefAdds,ImportFunRefFrontChar32 & HexStr2RegExpPattern(.Reference(0).sOrgCode,1),SkipVal)
			If CheckArray(TempList) = True Then
				.lReferenceNum = UBound(TempList) + 1
				ReDim Preserve strData.Reference(.lReferenceNum - 1) 'As REFERENCE_PROPERTIE
				m = 0: n = 0
				For i = 0 To .lReferenceNum - 1
					.Reference(i).lAddress = CLng(TempList(i)) + 2
					.Reference(i).sOrgCode = .Reference(0).sOrgCode
					If .Reference(i).lAddress < n Or .Reference(i).lAddress > m Then
						k = SkipSection(File,.Reference(i).lAddress,n,m)
					End If
					.Reference(i).inSecID = k
					If ShowMsg > 0 Then
						SetTextBoxString ShowMsg,Msg & Format$(i / .lReferenceNum,"#%")
					ElseIf ShowMsg < 0 Then
						PSL.OutputWnd(0).Clear
						PSL.Output Msg & Format$(i / .lReferenceNum,"#%")
					End If
				Next i
				'If TagType > 1 Then Call GetCustomStrType(FN,strData,TypeList,Mode,-1)
			End If
			.GetRefState = 1
		ElseIf RefAdds <> "" Or (.lReferenceNum > 0 And fType < 2) Then
			'If RefAdds <> "" Then Call GetRefList(strData,ReSplit(RefAdds,RefJoinStr),TempList,True)
			If .lReferenceNum > 0 Then
				'.Reference(0).sUseCode = ReverseHexCode(Hex$(.lStartAddress + VRK),8)
				'.Reference(0).sUseCode = Byte2Hex(Val2Bytes(.lStartAddress + VRK,4),0,3)
				For i = 0 To .lReferenceNum - 1
					If strData.lStartAddress <> .Reference(i).lUseFuncAddress Then
						j = .Reference(i).UseFuncInSecID
						VRK = File.SecList(j).lVirtualAddress - File.SecList(j).lPointerToRawData + File.ImageBase
						.Reference(i).sUseCode = ReverseHexCode(Hex$(.Reference(i).lUseFuncAddress + VRK),8)
					End If
					If ShowMsg > 0 Then
						SetTextBoxString ShowMsg,Msg & Format$(i / .lReferenceNum,"#%")
					ElseIf ShowMsg < 0 Then
						PSL.OutputWnd(0).Clear
						PSL.Output Msg & Format$(i / .lReferenceNum,"#%")
					End If
				Next i
				.GetRefState = 1
			End If
		Else
			GoTo ExitFunction
		End If
	End Select
	GetVARefList = .lReferenceNum
	End With
	If ShowMsg > 0 Then
		SetTextBoxString ShowMsg,Msg & "100%"
	ElseIf ShowMsg < 0 Then
		PSL.OutputWnd(0).Clear
		PSL.Output Msg & "100%"
	End If
	Exit Function
	'退出函数
	ExitFunction:
	ReDim strData.Reference(0) 'As REFERENCE_PROPERTIE
	strData.lReferenceNum = 0
	strData.GetRefState = 0
	If ShowMsg > 0 Then
		SetTextBoxString ShowMsg,Msg & "100%"
	ElseIf ShowMsg < 0 Then
		PSL.OutputWnd(0).Clear
		PSL.Output Msg & "100%"
	End If
End Function


'获取 PE64 原始字串的引用地址和代码
'返回 GetVAListPE64 = 步进长度
Private Function GetVAListPE64(FN As Variant,strData As STRING_SUB_PROPERTIE,RefMaxNum As Long,ByVal SecID As Long, _
				ByVal VRK As Long,ByVal StartPos As Long,ByVal RSize As Long,ByVal Mode As Long) As Long
	Dim i As Long,Temp As String
	With strData
		'i = (VRK - StartPos) And &HFF	'后3个字节查找，速度较慢
		'i = Val("&H" & Right$("0000" & Hex$(VRK - StartPos),4))	'后2个字节查找，速度较快
		i = (VRK - StartPos) And 65535	'后2个字节查找，速度较快，这里的 65535 不能替换成 &HFFFF，因为 &HFFFF 返回为 -1
		If i > RSize - StartPos Then i = RSize - StartPos
		If i > .lStartAddress - StartPos - 4 Then i = .lStartAddress - StartPos - 4
		If i > 0 Then
			GetVAListPE64 = i
			'正则表达式查找，速度较快，后3个字节查找时，开始地址为 StartPos + 1，否则为 StartPos + 2
			ReDim TempList(0) As String
			Temp = ImportFunRefFrontChar64 & HexStr2RegExpPattern(Right$(ReverseHexCode(Hex$(VRK - StartPos),8),4),1)
			TempList = GetVAListRegExp(ByteToString(GetBytes(FN,i + 3,StartPos - 2,Mode),CP_ISOLATIN1),Temp,StartPos - 2)
			'字节数组查找，速度较慢，后3个字节查找时，开始地址为 StartPos + 1，否则为 StartPos + 2
			'TempList = GetVAList(FN.ImageByte,Val2BytesRev(VRK - StartPos,4,2),StartPos + 2,StartPos + 2 + GetVAListPE64)
			If CheckArray(TempList) = False Then Exit Function
			For i = 0 To UBound(TempList)
				StartPos = CLng(TempList(i)) + 2	'前3个为引用的特征码，所以往后3个字节
				If VRK > StartPos Then
					'获取虚拟地址(即引用代码值)，并判断其是否正确
					RSize = GetLong(FN,StartPos,Mode)
					If RSize > 0 And RSize = VRK - StartPos Then
						If .lReferenceNum > RefMaxNum Then
							RefMaxNum = .lReferenceNum + 20
							ReDim Preserve strData.Reference(RefMaxNum) 'As REFERENCE_PROPERTIE
						End If
						'保存找到的引用地址及引用代码
						.Reference(.lReferenceNum).lAddress = StartPos
						.Reference(.lReferenceNum).sOrgCode = Byte2Hex(GetBytes(FN,4,StartPos,Mode),0,3)
						.Reference(.lReferenceNum).inSecID = SecID
						.lReferenceNum = .lReferenceNum + 1
					End If
				End If
			Next i
		ElseIf VRK > StartPos Then
			'获取虚拟地址(即引用代码值)，并判断其是否正确
			TempList = GetVAListRegExp(ByteToString(GetBytes(FN,5,StartPos - 2,Mode),CP_ISOLATIN1),ImportFunRefFrontChar64,StartPos - 2)
			If CheckArray(TempList) = False Then Exit Function
			RSize = GetLong(FN,StartPos,Mode)
			If RSize > 0 And RSize = VRK - StartPos Then
				GetVAListPE64 = 3
				If .lReferenceNum > RefMaxNum Then
					RefMaxNum = .lReferenceNum + 20
					ReDim Preserve strData.Reference(RefMaxNum) 'As REFERENCE_PROPERTIE
				End If
				'保存找到的引用地址及引用代码
				.Reference(.lReferenceNum).lAddress = StartPos
				.Reference(.lReferenceNum).sOrgCode = Byte2Hex(GetBytes(FN,4,StartPos,Mode),0,3)
				.Reference(.lReferenceNum).inSecID = SecID
				.lReferenceNum = .lReferenceNum + 1
			End If
		End If
	End With
End Function


'获取在字节数组中找到的匹配数组的列表(VB方式)
'注意：StartPos、EndPos 均为以 0 开始的地址
Private Function GetVAList(Bytes() As Byte,Find() As Byte,Optional ByVal StartPos As Long,Optional ByVal EndPos As Long) As String()
	Dim i As Long,j As Long,k As Long,m As Long,n As Long,Length As Long,TempByte() As Byte
	Length = UBound(Find) + 1
	TempByte = Find
	k = -1
	For i = 0 To Length - 1
		If Find(i) <> 0 Then
			If i <> 0 Then
				ReDim TempByte(Length - i - 1) As Byte
				CopyMemory TempByte(0), Find(i), Length - i
			End If
			k = i
			Exit For
		End If
	Next i
	If k = -1 Then
		If EndPos <= StartPos Then Exit Function
		ReDim TempByte(EndPos - StartPos) As Byte
		CopyMemory TempByte(0), Bytes(StartPos), EndPos - StartPos + 1
		GetVAList = GetVAListRegExp(ByteToString(TempByte,CP_ISOLATIN1),"\x00{" & CStr$(Length) & ",}",StartPos)
		Exit Function
	End If
	m = 50
	ReDim TempList(m) As String
	i = InStrB(StartPos + k + 1,Bytes,TempByte)
	Do While i > 0
		If k > 0 Then
			For j = i - k - 1 To i - 2
				If Bytes(j) <> 0 Then GoTo NextNum
			Next j
			i = i - k
		End If
		If EndPos > 0 Then
			If i + Length - 2 > EndPos Then Exit Do
		End If
		If n > m Then
			m = m * 2
			ReDim Preserve TempList(m) As String
		End If
		TempList(n) = CStr(i - 1)    '注意 InStrB 函数找到第一个数就返回"1"
		n = n + 1
		NextNum:
		i = InStrB(i + Length,Bytes,TempByte)
	Loop
	If n > 0 Then n = n - 1
	ReDim Preserve TempList(n) As String
	GetVAList = TempList
End Function


'获取在字节数组中找到的匹配数组的列表(正则表达式方式)
'注意：StartPos、EndPos 均为以 0 开始的地址
Private Function GetVAListRegExp(ByVal StrText As String,ByVal Patrn As String,ByVal StartPos As Long) As String()
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


'反转 Hex 码
Private Function ReverseHexCode(ByVal HexStr As String,ByVal Num As Long) As String
	Dim i As Long
	i = Len(HexStr)
	If i < Num Then HexStr = String$(Num - i,"0") & HexStr
	ReverseHexCode = HexStr
	For i = 1 To Num - 1 Step 2
		Mid$(ReverseHexCode,i,2) = Mid$(HexStr,Num - i,2)
	Next i
End Function


'字节转 Hex 码
'StartPos <= EndPos 获取低位到高位的 Hex 代码，否则获取高位到低位的 Hex 代码
Private Function Byte2Hex(Bytes As Variant,ByVal StartPos As Long,ByVal EndPos As Long) As String
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


'转换数值为字节数组(短于长度的高位截断)
Private Function Val2Bytes(ByVal Value As Long,ByVal Length As Integer,Optional ByVal ByteOrder As Boolean) As Byte()
	'On Error GoTo errHandle
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


'转换十进制和十六进制值为字符
'MaxVal = 0 按值计算应有的长度，> 0 按文件大小计算的位数，< 0 按指定位数
Private Function ValToStr(ByVal DecVal As Long,Optional ByVal MaxVal As Long,Optional ByVal DisPlayFormat As Boolean) As String
	On Error GoTo ExitFunction
	If DisPlayFormat = False Then
		ValToStr = CStr$(DecVal)
	Else
		ValToStr = Hex$(DecVal)
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


'转换十进制和十六进制字符为十进制值
Private Function StrToVal(ByVal textStr As String,Optional ByVal DisPlayFormat As Boolean) As Variant
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


'检查数组是否已经初始化
'返回值:TRUE 已经初始化, FALSE 未初始化
Private Function CheckArrEmpty(ByRef MyArr As Variant) As Boolean
	On Error Resume Next
	If UBound(MyArr) >= 0 Then CheckArrEmpty = True
	Err.Clear
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


'字节数组转正则表达式使用的转义符模板
'Mode = 0 转为有 [] 形式，否则为无 [] 形式
Private Function Byte2RegExpPattern(Bytes() As Byte,Optional ByVal Mode As Long,Optional ByVal CodePage As Long) As String
	If Mode = 0 Then
		Byte2RegExpPattern = "[" & Byte2HexEsc(Bytes,0,-1,CodePage) & "]"
	Else
		Byte2RegExpPattern = Byte2HexEsc(Bytes,0,-1,CodePage)
	End If
End Function


'Hex 字符串转正则表达式使用的 Hex 转义符模板
'Mode = 0 转为有 [] 形式，否则为无 [] 形式
Private Function HexStr2RegExpPattern(ByVal HexStr As String,Optional ByVal Mode As Long) As String
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


'格式化 HEX 字串
Private Function FormatHexStr(ByVal textStr As String,ByVal Length As Integer) As String
	If textStr = "" Then Exit Function
	If (Len(textStr) Mod Length) = 0 Then
		FormatHexStr = textStr
	Else
		FormatHexStr = "0" & textStr
	End If
End Function


'引用数据转字串
'Mode = 0 转换选定引用数据为代码页方式的简单信息字串
'Mode = 1 转换选定引用数据为函数方式的简单信息字串

'Mode = 2 转换选定引用数据为代码页方式的详细信息字串
'Mode = 3 转换选定引用数据为函数方式的详细信息字串

'Mode = 4 转换所有引用数据为可用的代码页方式的详细信息字串
'Mode = 5 转换所有引用数据为未改变的代码页方式的详细信息字串
'Mode = 6 转换所有引用数据为已改变的代码页方式的详细信息字串
'Mode = 7 转换所有引用数据为可用但未改变的代码页方式的详细信息字串
'Mode = 8 转换所有引用数据为全部代码页方式的详细信息字串

'Mode = 9 转换所有引用数据为可用的函数方式的详细信息字串
'Mode = 10 转换所有引用数据为未改变的函数方式的详细信息字串
'Mode = 12 转换所有引用数据为已改变的函数方式的详细信息字串
'Mode = 13 转换所有引用数据为可用但未改变的函数方式的详细信息字串
'Mode = 14 转换所有引用数据为全部函数方式的详细信息字串
Private Function Reference2Str(File As FILE_PROPERTIE,strData As STRING_SUB_PROPERTIE,IndexList() As Long, _
				ByVal ShowState As Boolean,ByVal Mode As Integer) As String
	Dim i As Long,j As Long,k As Long,n As Long,Dic As Object
	On Error Resume Next
	If strData.lReferenceNum = 0 Then
		Reference2Str = MsgList(73)
		Exit Function
	End If
	With strData
		Select Case Mode
		Case 0,1
			If ShowState = False Then
				j = IndexList(0)
			Else
				j = -1
				For i = 0 To .lReferenceNum - 1
					If .Reference(i).lAddress = RefAddList(IndexList(0)) Then
						j = i
						Exit For
					End If
				Next i
				If j = -1 Then Exit Function
			End If
			If Mode = 0 Then
				Reference2Str = Replace$(MsgList(31),"%s!1!",ValToStr(.Reference(j).lAddress,File.FileSize,True))
				Reference2Str = Replace$(Reference2Str,"%s!2!",.Reference(j).sOrgCode)
				Reference2Str = Replace$(Reference2Str,"%s!3!",File.SecList(.Reference(j).inSecID).sName)
				Reference2Str = Replace$(Reference2Str,"%s!4!",CStr$(.Reference(j).lOrgCodePage))
				Reference2Str = Replace$(Reference2Str,"%s!5!",CStr$(.Reference(j).lUseCodePage))
				Reference2Str = Replace$(Reference2Str,"%s!6!",ValToStr(.Reference(j).lCPAddress,File.FileSize,True))
				Reference2Str = Replace$(Reference2Str,"%s!7!",CStr$(.Reference(j).iCPByteLength))
				Reference2Str = Replace$(Reference2Str,"%s!8!",IIf(.Reference(j).iCPByteOrder = -1,MsgList(78), _
														IIf(.Reference(j).iCPByteOrder = 0,MsgList(79),MsgList(80))))
			Else
				Reference2Str = Replace$(MsgList(83),"%s!1!",ValToStr(.Reference(j).lAddress,File.FileSize,True))
				Reference2Str = Replace$(Reference2Str,"%s!2!",File.SecList(.Reference(j).inSecID).sName)
				Reference2Str = Replace$(Reference2Str,"%s!3!",.sString)
				Reference2Str = Replace$(Reference2Str,"%s!4!",ValToStr(.lStartAddress,File.FileSize,True))
				Reference2Str = Replace$(Reference2Str,"%s!5!",.Reference(j).sOrgCode)
				Reference2Str = Replace$(Reference2Str,"%s!6!",.Reference(j).sUseFunc)
				Reference2Str = Replace$(Reference2Str,"%s!7!",ValToStr(.Reference(j).lUseFuncAddress,File.FileSize,True))
				Reference2Str = Replace$(Reference2Str,"%s!8!",.Reference(j).sUseCode)
			End If
		Case 2, 4 To 8
			If Mode = 2 Then k = UBound(IndexList) Else k = .lReferenceNum - 1
			ReDim List(k + 6) As String
			Select Case File.Magic
			Case "PE32","NET32","MAC32"
				If File.LangType = DELPHI_FILE_SIGNATURE Then
					List(0) = Replace$(MsgList(74),"%s!1!","Delphi32")
				ElseIf File.LangType = NET_FILE_SIGNATURE Then
					List(0) = Replace$(MsgList(74),"%s!1!",".NET32")
				ElseIf InStr(File.Magic,"MAC") Then
					List(0) = Replace$(MsgList(74),"%s!1!","MAC32")
				Else
					List(0) = Replace$(MsgList(74),"%s!1!","PE32")
				End If
			Case "PE64","NET64","MAC64"
				If File.LangType = DELPHI_FILE_SIGNATURE Then
					List(0) = Replace$(MsgList(74),"%s!1!","Delphi64")
				ElseIf File.LangType = NET_FILE_SIGNATURE Then
					List(0) = Replace$(MsgList(74),"%s!1!",".NET64")
				ElseIf InStr(File.Magic,"MAC") Then
					List(0) = Replace$(MsgList(74),"%s!1!","MAC64")
				Else
					List(0) = Replace$(MsgList(74),"%s!1!","PE64")
				End If
			Case Else
				List(0) = Replace$(MsgList(74),"%s!1!",MsgList(88))
			End Select
			List(0) = Replace$(List(0),"%s!2!",.sString)
			List(0) = Replace$(List(0),"%s!3!",ValToStr(.lStartAddress,File.FileSize,True))
			List(0) = Replace$(List(0),"%s!4!",File.SecList(.inSectionID).sName)
			List(0) = Replace$(List(0),"%s!5!",.DllName)
			List(1) = MsgList(75) & MsgList(75)
			List(2) = MsgList(76)
			List(3) = MsgList(75) & MsgList(75)
			If ShowState = False Then
				For i = 0 To k
					If Mode = 2 Then j = IndexList(i) Else j = i
					Select Case Mode
					Case 4
						ShowState = IIf(.Reference(j).lOrgCodePage > 0,True,False)
					Case 5
						ShowState = IIf(.Reference(j).lOrgCodePage = .Reference(j).lUseCodePage,True,False)
					Case 6
						ShowState = IIf(.Reference(j).lOrgCodePage <> .Reference(j).lUseCodePage,True,False)
					Case 7
						ShowState = IIf(.Reference(j).lOrgCodePage > 0 And .Reference(j).lOrgCodePage = .Reference(j).lUseCodePage,True,False)
					Case 2, 8
						ShowState = True
					End Select
					If ShowState = True Then
						List(n + 4) = Replace$(MsgList(77),"%s!1!",CStr$(n + 1))
						List(n + 4) = Replace$(List(n + 4),"%s!2!",ValToStr(.Reference(j).lAddress,File.FileSize,True))
						List(n + 4) = Replace$(List(n + 4),"%s!3!",.Reference(j).sOrgCode)
						List(n + 4) = Replace$(List(n + 4),"%s!4!",File.SecList(.Reference(j).inSecID).sName)
						List(n + 4) = Replace$(List(n + 4),"%s!5!",ValToStr(.Reference(j).lCPAddress,File.FileSize,True))
						List(n + 4) = Replace$(List(n + 4),"%s!6!",CStr$(.Reference(j).lOrgCodePage))
						List(n + 4) = Replace$(List(n + 4),"%s!7!",CStr$(.Reference(j).lUseCodePage))
						List(n + 4) = Replace$(List(n + 4),"%s!8!",CStr$(.Reference(j).iCPByteLength))
						List(n + 4) = Replace$(List(n + 4),"%s!9!",IIf(.Reference(j).iCPByteOrder = -1,MsgList(78), _
											IIf(.Reference(j).iCPByteOrder = 0,MsgList(79),MsgList(80))))
						n = n + 1
					End If
				Next i
			Else
				Set Dic = CreateObject("Scripting.Dictionary")
				For i = 0 To .lReferenceNum - 1
					If Not Dic.Exists(.Reference(i).lAddress) Then
						Dic.Add(.Reference(i).lAddress,i)
					End If
				Next i
				For i = 0 To k
					j = Dic.Ietm(RefAddList(IndexList(i)))
					List(n + 4) = Replace$(MsgList(77),"%s!1!",CStr$(n + 1))
					List(n + 4) = Replace$(List(n + 4),"%s!2!",ValToStr(.Reference(j).lAddress,File.FileSize,True))
					List(n + 4) = Replace$(List(n + 4),"%s!3!",.Reference(j).sOrgCode)
					List(n + 4) = Replace$(List(n + 4),"%s!4!",File.SecList(.Reference(j).inSecID).sName)
					List(n + 4) = Replace$(List(n + 4),"%s!5!",ValToStr(.Reference(j).lCPAddress,File.FileSize,True))
					List(n + 4) = Replace$(List(n + 4),"%s!6!",CStr$(.Reference(j).lOrgCodePage))
					List(n + 4) = Replace$(List(n + 4),"%s!7!",CStr$(.Reference(j).lUseCodePage))
					List(n + 4) = Replace$(List(n + 4),"%s!8!",CStr$(.Reference(j).iCPByteLength))
					List(n + 4) = Replace$(List(n + 4),"%s!9!",IIf(.Reference(j).iCPByteOrder = -1,MsgList(78), _
										IIf(.Reference(j).iCPByteOrder = 0,MsgList(79),MsgList(80))))
					n = n + 1
				Next i
				Set Dic = Nothing
			End If
			List(n + 4) = MsgList(75) & MsgList(75)
			If n = 0 Then Exit Function
			ReDim Preserve List(n + 5) As String
			Reference2Str = StrListJoin(List,TextJoinStr)
		Case 3, 9 To 13
			If Mode = 3 Then k = UBound(IndexList) Else k = .lReferenceNum - 1
			ReDim List(k + 6) As String
			Select Case File.Magic
			Case "PE32","NET32","MAC32"
				If File.LangType = DELPHI_FILE_SIGNATURE Then
					List(0) = Replace$(MsgList(74),"%s!1!","Delphi32")
				ElseIf File.LangType = NET_FILE_SIGNATURE Then
					List(0) = Replace$(MsgList(74),"%s!1!",".NET32")
				ElseIf InStr(File.Magic,"MAC") Then
					List(0) = Replace$(MsgList(74),"%s!1!","MAC32")
				Else
					List(0) = Replace$(MsgList(74),"%s!1!","PE32")
				End If
			Case "PE64","NET64","MAC64"
				If File.LangType = DELPHI_FILE_SIGNATURE Then
					List(0) = Replace$(MsgList(74),"%s!1!","Delphi64")
				ElseIf File.LangType = NET_FILE_SIGNATURE Then
					List(0) = Replace$(MsgList(74),"%s!1!",".NET64")
				ElseIf InStr(File.Magic,"MAC") Then
					List(0) = Replace$(MsgList(74),"%s!1!","MAC64")
				Else
					List(0) = Replace$(MsgList(74),"%s!1!","PE64")
				End If
			Case Else
				List(0) = Replace$(MsgList(74),"%s!1!",MsgList(88))
			End Select
			List(0) = Replace$(List(0),"%s!2!",.sString)
			List(0) = Replace$(List(0),"%s!3!",ValToStr(.lStartAddress,File.FileSize,True))
			List(0) = Replace$(List(0),"%s!4!",File.SecList(.inSectionID).sName)
			List(0) = Replace$(List(0),"%s!5!",.DllName)
			List(1) = MsgList(75) & MsgList(75)
			List(2) = MsgList(85)
			List(3) = MsgList(75) & MsgList(75)
			If ShowState = False Then
				For i = 0 To k
					If Mode = 3 Then j = IndexList(i) Else j = i
					Select Case Mode
					Case 10, 12
						ShowState = IIf(.Reference(j).lUseFuncAddress = .lStartAddress,True,False)
					Case 11
						ShowState = IIf(.Reference(j).lUseFuncAddress <> .lStartAddress,True,False)
					Case 3, 9, 13
						ShowState = True
					End Select
					If ShowState = True Then
						List(n + 4) = Replace$(MsgList(90),"%s!1!",CStr$(n + 1))
						List(n + 4) = Replace$(List(n + 4),"%s!2!",ValToStr(.Reference(j).lAddress,File.FileSize,True))
						List(n + 4) = Replace$(List(n + 4),"%s!3!",File.SecList(.Reference(j).inSecID).sName)
						List(n + 4) = Replace$(List(n + 4),"%s!4!",.sString)
						List(n + 4) = Replace$(List(n + 4),"%s!5!",ValToStr(.lStartAddress,File.FileSize,True))
						List(n + 4) = Replace$(List(n + 4),"%s!6!",.Reference(j).sOrgCode)
						List(n + 4) = Replace$(List(n + 4),"%s!7!",.Reference(j).sUseFunc)
						List(n + 4) = Replace$(List(n + 4),"%s!8!",ValToStr(.Reference(j).lUseFuncAddress,File.FileSize,True))
						List(n + 4) = Replace$(List(n + 4),"%s!9!",.Reference(j).sUseCode)
						n = n + 1
					End If
				Next i
			Else
				Set Dic = CreateObject("Scripting.Dictionary")
				For i = 0 To .lReferenceNum - 1
					If Not Dic.Exists(.Reference(i).lAddress) Then
						Dic.Add(.Reference(i).lAddress,i)
					End If
				Next i
				For i = 0 To k
					j = Dic.Ietm(RefAddList(IndexList(i)))
					List(n + 4) = Replace$(MsgList(90),"%s!1!",CStr$(n + 1))
					List(n + 4) = Replace$(List(n + 4),"%s!2!",ValToStr(.Reference(j).lAddress,File.FileSize,True))
					List(n + 4) = Replace$(List(n + 4),"%s!3!",File.SecList(.Reference(j).inSecID).sName)
					List(n + 4) = Replace$(List(n + 4),"%s!4!",.sString)
					List(n + 4) = Replace$(List(n + 4),"%s!5!",ValToStr(.lStartAddress,File.FileSize,True))
					List(n + 4) = Replace$(List(n + 4),"%s!6!",.Reference(j).sOrgCode)
					List(n + 4) = Replace$(List(n + 4),"%s!7!",.Reference(j).sUseFunc)
					List(n + 4) = Replace$(List(n + 4),"%s!8!",ValToStr(.Reference(j).lUseFuncAddress,File.FileSize,True))
					List(n + 4) = Replace$(List(n + 4),"%s!9!",.Reference(j).sUseCode)
					n = n + 1
				Next i
				Set Dic = Nothing
			End If
			List(n + 4) = MsgList(75) & MsgList(75)
			If n = 0 Then Exit Function
			ReDim Preserve List(n + 5) As String
			Reference2Str = StrListJoin(List,TextJoinStr)
		End Select
	End With
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


'计算 PE 文件对齐
Private Function Alignment(ByVal orgValue As Long,ByVal AlignVal As Long,ByVal RoundVal As Long) As Long
	If AlignVal < 1 Then
		Alignment = orgValue
	Else
		Alignment = IIf(orgValue Mod AlignVal = 0,orgValue,AlignVal * ((orgValue \ AlignVal) + RoundVal))
	End If
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


'转换字节数组为数值
'ByteOrder = False 按高位在后转，否则按高位在前转
Private Function Bytes2Val(Bytes() As Byte,ByVal Length As Integer,Optional ByVal ByteOrder As Boolean) As Long
	On Error GoTo errHandle
	If UBound(Bytes) + 1 < Length Then Exit Function
	If ByteOrder = False Then
		CopyMemory Bytes2Val, Bytes(0), Length
	Else
		CopyMemory Bytes2Val, ReverseValByte(Bytes,0,-1)(0), Length
	End If
	errHandle:
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
Private Function ReverseValByte(Bytes() As Byte,ByVal StartPos As Long,ByVal endPos As Long) As Byte()
	Dim i As Long,Temp() As Byte
	If StartPos < 0 Then StartPos = LBound(Bytes)
	If endPos < 0 Then endPos = UBound(Bytes)
	Temp = Bytes
	For i = StartPos To endPos
		Temp(i) = Bytes(endPos - i)
	Next i
	ReverseValByte = Temp
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
	If lRet > 0 Then
		MultiByteToUTF16 = Left$(MultiByteToUTF16, lRet)
	End If
	Exit Function
	errHandle:
	MultiByteToUTF16 = ""
End Function


'获取偶数位
'Mode = 0 奇数加 1 个字节，Mode = 1 奇数减 1 个字节
Private Function GetEvenPos(ByVal Pos As Long,Optional ByVal Mode As Long) As Long
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


'返回对话框某个控件中的字串
Private Function GetTextBoxString(ByVal hwnd As Long) As String
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
Private Function SetTextBoxString(ByVal hwnd As Long,ByVal StrText As String,Optional ByVal Mode As Boolean) As Boolean
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


'修正 PSL 2015 及以上版本宏引擎的 Split 函数拆分空字符串时返回未初始化数组的错误
Private Function ReSplit(ByVal textStr As String,Optional ByVal Sep As String = " ",Optional ByVal Max As Integer = -1) As String()
	If textStr = "" Then
		ReDim TempList(0) As String
		ReSplit = TempList
	Else
		ReSplit = Split(textStr,Sep,Max)
	End If
End Function


'获取文件的类型，PE 还是 MAC 还是非 PE 文件
Private Function GetFileFormat(ByVal FilePath As String,ByVal Mode As Long,FileType As Integer) As String
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


'获取文件及子文件的数据结构信息
Private Function GetHeaders(ByVal strFilePath As String,File As FILE_PROPERTIE,ByVal Mode As Long,FileType As Integer) As Boolean
	Select Case GetFileFormat(File.FilePath,Mode,FileType)
	Case "PE","NET",""
		GetHeaders = GetPEHeaders(File.FilePath,File,Mode)
	Case "MAC"
		GetHeaders = GetMacHeaders(File.FilePath,File,Mode)
	End Select
End Function


'显示文件信息
Private Sub ShowInfo(ByVal FilePath As String,ByVal Info As String)
	Begin Dialog UserDialog 990,448,Replace$(MsgList(89),"%s",FilePath) ' %GRID:10,7,1,1
		TextBox 0,7,990,406,.InTextBox,1
		OKButton 440,420,100,21,.OKButton
	End Dialog
	Dim dlg As UserDialog
	dlg.InTextBox = Info
	Dialog dlg
End Sub


'获取文件版本信息
Private Function GetFileInfo(ByVal strFilePath As String,File As FILE_PROPERTIE) As Boolean
	Dim i As Integer,lngBufferlen As Long,lngRc As Long,lngVerPointer As Long
	Dim bytBuffer() As Byte,strTemp As String
	Dim strBuffer As String,strLangCharset As String,strVersionInfo(7) As String
	'文件已打开时退出
	If IsOpen(strFilePath) = True Then Exit Function
	' get file size
	lngBufferlen = GetFileVersionInfoSize(strFilePath, 0&)
	If lngBufferlen > 0 Then
		ReDim bytBuffer(lngBufferlen) As Byte
		lngRc = GetFileVersionInfo(strFilePath, 0&, lngBufferlen, bytBuffer(0))
		If lngRc <> 0 Then
			lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
			If lngRc <> 0 Then
				'lngVerPointer is a pointer to four 4 bytes of Hex number,
				'first two bytes are language id, and last two bytes are code
				'page. However, strLangCharset needs a  string of
				'4 hex digits, the first two characters correspond to the
				'language id and last two the last two character correspond
				'to the code page id.
				ReDim bytBuff(lngBufferlen - 1) As Byte
				MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
				'strLangCharset = Hex(bytBuff(2) + bytBuff(3) * &H100 + bytBuff(0) * &H10000 + bytBuff(1) * &H1000000)
				strLangCharset = Byte2Hex(LowByte2HighByte(bytBuff,2),0,-1)
				'now we change the order of the language id and code page
				'and convert it into a string representation.
				'For example, it may look like 040904E4
				'Or to pull it all apart:
				'04------        = SUBLANG_ENGLISH_USA
				'--09----        = LANG_ENGLISH
				' ----04E4 = 1252 = Codepage for Windows:Multilingual

				'If Len(strLangCharset) - 2 >= 2 Then
				'    If Mid$(strLangCharset, 2, 2) = LANG_ENGLISH Then
				'    	strLangCharset2 = "English (US)"
				'    End If
				'End If

				Do While Len(strLangCharset) < 8
					strLangCharset = "0" & strLangCharset
				Loop

				' assign propertienames
				strVersionInfo(0) = "CompanyName"
				strVersionInfo(1) = "FileDescription"
				strVersionInfo(2) = "FileVersion"
				strVersionInfo(3) = "InternalName"
				strVersionInfo(4) = "LegalCopyright"
				strVersionInfo(5) = "OriginalFileName"
				strVersionInfo(6) = "ProductName"
				strVersionInfo(7) = "ProductVersion"
				' loop and get FILE_PROPERTIEs
				For i = 0 To 7
					strBuffer = String$(255, 0)
					strTemp = "\StringFileInfo\" & strLangCharset & "\" & strVersionInfo(i)
					lngRc = VerQueryValue(bytBuffer(0), strTemp, lngVerPointer, lngBufferlen)
					If lngRc <> 0 Then
						' get and format data
						lstrcpy strBuffer, lngVerPointer
						strVersionInfo(i) = Replace$(strBuffer,vbNullChar,"")
					Else
						' property not found
						strVersionInfo(i) = ""
					End If
				Next i
			End If
		End If
	End If
	' assign array to user-defined-type
	File.CompanyName = strVersionInfo(0)
	File.FileDescription = strVersionInfo(1)
	File.FileVersion = Trim$(strVersionInfo(2))
	File.InternalName = strVersionInfo(3)
	File.LegalCopyright = strVersionInfo(4)
	File.OrigionalFileName = strVersionInfo(5)
	File.ProductName = strVersionInfo(6)
	File.ProductVersion = strVersionInfo(7)
	File.LanguageID = Left$(strLangCharset,4)
	File.FileName = Mid$(File.FilePath,InStrRev(File.FilePath,"\") + 1)
	File.FileSize = FileLen(strFilePath)
	File.DateLastModified = FileDateTime(strFilePath)
	File.DateCreated = GetFileDate(strFilePath,0)
	GetFileInfo = True
End Function


'获取文件的创建、访问、修改日期
'Mode = 0 创建日期
'Mode = 1 访问日期
'Mode = 2 修改日期
Private Function GetFileDate(ByVal strFilePath As String,ByVal Mode As Long) As Date
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


'查看本地化文件信息
'DisPlayFormat = False 十进制显示数值，否则十六进制显示数值
Private Sub FileInfoView(File As FILE_PROPERTIE,ByVal DisPlayFormat As Boolean)
	Dim i As Long,j As Long,n As Long,Stemp As Boolean
	On Error GoTo ErrHandle
	If InStr(File.Magic,"MAC") Then Stemp = True
	'MAC64的情况下，无法计算 64 位(8 个字节)的数值，只能用16进制显示
	If File.Magic = "MAC64" Then
		If DisPlayFormat = False Then DisPlayFormat = True
	End If
	'写入文件属性信息
	n = 16
	ReDim List(n) As String
	With File
		List(0) = MsgList(91)
		List(1) = Replace$(MsgList(92),"%s",.FileName)
		List(2) = Replace$(MsgList(93),"%s",.FilePath)
		List(3) = Replace$(MsgList(94),"%s",.FileDescription)
		List(4) = Replace$(MsgList(95),"%s",.FileVersion)
		List(5) = Replace$(MsgList(96),"%s",.ProductName)
		List(6) = Replace$(MsgList(97),"%s",.ProductVersion)
		List(7) = Replace$(MsgList(98),"%s",.LegalCopyright)
		List(8) = Replace$(MsgList(99),"%s",CStr$(.FileSize))
		List(9) = Replace$(MsgList(100),"%s",CStr$(.DateCreated))
		List(10) = Replace$(MsgList(101),"%s",CStr$(.DateLastModified))
		List(11) = Replace$(MsgList(102),"%s",PSL.GetLangCode(Val("&H" & .LanguageID),pslCodeText))
		List(12) = Replace$(MsgList(103),"%s",.CompanyName)
		List(13) = Replace$(MsgList(104),"%s",.OrigionalFileName)
		List(14) = Replace$(MsgList(105),"%s",.InternalName)
		Select Case .Magic
		Case "PE32","NET32","MAC32"
			If .LangType = DELPHI_FILE_SIGNATURE Then
				List(15) = Replace$(MsgList(106),"%s","Delphi32")
				List(16) = Replace$(MsgList(107),"%s","0x" & ValToStr(.ImageBase,-8,True))
			ElseIf .LangType = NET_FILE_SIGNATURE Then
				List(15) = Replace$(MsgList(106),"%s",".NET32")
				List(16) = Replace$(MsgList(107),"%s","0x" & ValToStr(.ImageBase,-8,True))
			ElseIf Stemp = True Then
				List(15) = Replace$(MsgList(106),"%s","MAC32")
			Else
				List(15) = Replace$(MsgList(106),"%s","PE32")
				List(16) = Replace$(MsgList(107),"%s","0x" & ValToStr(.ImageBase,-8,True))
			End If
		Case "PE64","NET64","MAC64"
			If .LangType = DELPHI_FILE_SIGNATURE Then
				List(15) = Replace$(MsgList(106),"%s","Delphi64")
				List(16) = Replace$(MsgList(107),"%s","0x" & ReverseHexCode(Byte2Hex(.ImageBase,0,-1),16))
			ElseIf .LangType = NET_FILE_SIGNATURE Then
				List(15) = Replace$(MsgList(106),"%s",".NET64")
				List(16) = Replace$(MsgList(107),"%s","0x" & ReverseHexCode(Byte2Hex(.ImageBase,0,-1),16))
			ElseIf Stemp = True Then
				List(15) = Replace$(MsgList(106),"%s","MAC64")
			Else
				List(15) = Replace$(MsgList(106),"%s","PE64")
				List(16) = Replace$(MsgList(107),"%s","0x" & ReverseHexCode(Byte2Hex(.ImageBase,0,-1),16))
			End If
		Case Else
			List(15) = Replace$(MsgList(106),"%s",MsgList(149))
		End Select
	End With
	If List(16) = "" Then n = n - 1
	'每个文件节的偏移地址
	ReDim Preserve List(n + 6 + File.MaxSecIndex) As String
	List(n + 2) = MsgList(108)
	List(n + 3) = MsgList(111) & MsgList(111)
	List(n + 4) = IIf(Stemp = False,MsgList(109),MsgList(151))
	List(n + 5) = MsgList(111) & MsgList(111)
	n = n + 6
	For i = 0 To File.MaxSecIndex - 1
		With File.SecList(i)
			List(n) = Replace$(IIf(Stemp = False,MsgList(110),MsgList(152)),"%s!1!",MsgList(112))
			List(n) = Replace$(List(n),"%s!2!",IIf(File.Magic = "",MsgList(149),.sName))
			List(n) = Replace$(List(n),"%s!3!","")
			List(n) = Replace$(List(n),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
			List(n) = Replace$(List(n),"%s!5!",ValToStr(.lPointerToRawData + _
							IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
			List(n) = Replace$(List(n),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
			If .SubSecs > 0 Then
				n = n + 1
				ReDim Preserve List(n + .SubSecs + File.MaxSecIndex - i) As String
				For j = 0 To .SubSecs - 1
					List(n) = Replace$(MsgList(152),"%s!1!",MsgList(112))
					List(n) = Replace$(List(n),"%s!2!","")
					List(n) = Replace$(List(n),"%s!3!",.SubSecList(j).sName)
					List(n) = Replace$(List(n),"%s!4!",ValToStr(.SubSecList(j).lPointerToRawData,File.FileSize,DisPlayFormat))
					List(n) = Replace$(List(n),"%s!5!",ValToStr(.SubSecList(j).lPointerToRawData + _
									IIf(.SubSecList(j).lSizeOfRawData = 0,0,.SubSecList(j).lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
					List(n) = Replace$(List(n),"%s!6!",ValToStr(.SubSecList(j).lSizeOfRawData,File.FileSize,DisPlayFormat))
					n = n + 1
				Next j
			Else
				n = n + 1
			End If
		End With
	Next i
	'隐藏节的偏移地址、子 PE 地址及数量
	With File.SecList(File.MaxSecIndex)
		If .lSizeOfRawData > 0 Then
			ReDim Preserve List(n) As String
			List(n) = Replace$(IIf(Stemp = False,MsgList(110),MsgList(152)),"%s!1!",MsgList(112))
			List(n) = Replace$(List(n),"%s!2!",MsgList(115))
			List(n) = Replace$(List(n),"%s!3!","")
			List(n) = Replace$(List(n),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
			List(n) = Replace$(List(n),"%s!5!",ValToStr(.lPointerToRawData + _
						IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
			List(n) = Replace$(List(n),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
			n = n + 1
		End If
		If File.NumberOfSub > 0 Then
			ReDim Preserve List(n) As String
			List(n) = Replace$(IIf(Stemp = False,MsgList(110),MsgList(152)),"%s!1!",MsgList(112))
			List(n) = Replace$(List(n),"%s!2!",Replace$(MsgList(150),"%s",CStr$(File.NumberOfSub)))
			List(n) = Replace$(List(n),"%s!3!","")
			List(n) = Replace$(List(n),"%s!4!",ValToStr(.lPointerToRawData + .lSizeOfRawData,File.FileSize,DisPlayFormat))
			List(n) = Replace$(List(n),"%s!5!",ValToStr(File.FileSize - 1,File.FileSize,DisPlayFormat))
			List(n) = Replace$(List(n),"%s!6!",ValToStr(File.FileSize - .lPointerToRawData - .lSizeOfRawData,File.FileSize,DisPlayFormat))
			n = n + 1
		End If
	End With
	'每个文件节的相对虚拟地址
	n = n + 1
	ReDim Preserve List(n + File.MaxSecIndex) As String
	For i = 0 To File.MaxSecIndex - 1
		With File.SecList(i)
			List(n) = Replace$(IIf(Stemp = False,MsgList(110),MsgList(152)),"%s!1!",MsgList(113))
			List(n) = Replace$(List(n),"%s!2!",IIf(File.Magic = "",MsgList(149),.sName))
			List(n) = Replace$(List(n),"%s!3!","")
			If File.Magic <> "MAC64" Then
				List(n) = Replace$(List(n),"%s!4!",ValToStr(.lVirtualAddress,File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!5!",ValToStr(.lVirtualAddress + _
							IIf(.lVirtualSize = 0,0,.lVirtualSize - 1),File.FileSize,DisPlayFormat))
			Else
				List(n) = Replace$(List(n),"%s!4!",ValToStr(.lVirtualAddress1,0,DisPlayFormat) & _
							ValToStr(.lVirtualAddress,-8,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!5!",ValToStr(.lVirtualAddress1,0,DisPlayFormat) & _
							ValToStr(.lVirtualAddress + IIf(.lVirtualSize = 0,0,.lVirtualSize - 1),-8,DisPlayFormat))
			End If
			List(n) = Replace$(List(n),"%s!6!",ValToStr(.lVirtualSize,File.FileSize,DisPlayFormat))
			If .SubSecs > 0 Then
				n = n + 1
				ReDim Preserve List(n + .SubSecs + File.MaxSecIndex - i) As String
				For j = 0 To .SubSecs - 1
					List(n) = Replace$(MsgList(152),"%s!1!",MsgList(113))
					List(n) = Replace$(List(n),"%s!2!","")
					List(n) = Replace$(List(n),"%s!3!",.SubSecList(j).sName)
					If File.Magic <> "MAC64" Then
						List(n) = Replace$(List(n),"%s!4!",ValToStr(.SubSecList(j).lVirtualAddress,File.FileSize,DisPlayFormat))
						List(n) = Replace$(List(n),"%s!5!",ValToStr(.SubSecList(j).lVirtualAddress + _
									IIf(.SubSecList(j).lVirtualSize = 0,0,.lVirtualSize - 1),File.FileSize,DisPlayFormat))
					Else
						List(n) = Replace$(List(n),"%s!4!",ValToStr(.SubSecList(j).lVirtualAddress1,0,DisPlayFormat) & _
									ValToStr(.SubSecList(j).lVirtualAddress,-8,DisPlayFormat))
						List(n) = Replace$(List(n),"%s!5!",ValToStr(.SubSecList(j).lVirtualAddress1,0,DisPlayFormat) & _
									ValToStr(.SubSecList(j).lVirtualAddress + IIf(.SubSecList(j).lVirtualSize = 0,0,.lVirtualSize - 1),-8,DisPlayFormat))
					End If
					List(n) = Replace$(List(n),"%s!6!",ValToStr(.SubSecList(j).lVirtualSize,File.FileSize,DisPlayFormat))
					List(n) = Replace$(List(n),"%s!7!","")
					n = n + 1
				Next j
			Else
				n = n + 1
			End If
		End With
	Next i
	'隐藏节的相对虚拟地址、子 PE 地址及数量
	With File.SecList(File.MaxSecIndex)
		If .lVirtualSize > 0 Then
			ReDim Preserve List(n) As String
			List(n) = Replace$(IIf(Stemp = False,MsgList(110),MsgList(152)),"%s!1!",MsgList(113))
			List(n) = Replace$(List(n),"%s!2!",MsgList(115))
			List(n) = Replace$(List(n),"%s!3!","")
			List(n) = Replace$(List(n),"%s!4!",MsgList(116))
			List(n) = Replace$(List(n),"%s!5!",MsgList(116))
			List(n) = Replace$(List(n),"%s!6!",MsgList(116))
			n = n + 1
		End If
		If File.NumberOfSub > 0 Then
			ReDim Preserve List(n) As String
			List(n) = Replace$(IIf(Stemp = False,MsgList(110),MsgList(152)),"%s!1!",MsgList(113))
			List(n) = Replace$(List(n),"%s!2!",Replace$(MsgList(150),"%s",CStr$(File.NumberOfSub)))
			List(n) = Replace$(List(n),"%s!3!","")
			List(n) = Replace$(List(n),"%s!4!",MsgList(116))
			List(n) = Replace$(List(n),"%s!5!",MsgList(116))
			List(n) = Replace$(List(n),"%s!6!",MsgList(116))
			n = n + 1
		End If
	End With
	ReDim Preserve List(n) As String
	List(n) = MsgList(111) & MsgList(111)
	'数据目录地址及所在文件节
	If File.DataDirs > 0 Then
		ReDim Preserve List(n + 6 + File.DataDirs) As String
		List(n + 2) = MsgList(118)
		List(n + 3) = MsgList(111) & MsgList(111)
		List(n + 4) = MsgList(119)
		List(n + 5) = MsgList(111) & MsgList(111)
		n = n + 6
		For i = 0 To File.DataDirs - 1
			With File.DataDirectory(i)
				List(n) = Replace$(MsgList(110),"%s!1!",MsgList(i + 120))
				If .lPointerToRawData > 0 Then
					j = SkipSection(File,.lPointerToRawData,0,0,1)
					If j > -1 Then
						List(n) = Replace$(List(n),"%s!2!",File.SecList(j).sName)
					Else
						List(n) = Replace$(List(n),"%s!2!",MsgList(136))
					End If
				ElseIf .lSizeOfRawData > 0 Then
					List(n) = Replace$(List(n),"%s!2!",MsgList(136))
				Else
					List(n) = Replace$(List(n),"%s!2!",MsgList(137))
				End If
				List(n) = Replace$(List(n),"%s!3!","")
				List(n) = Replace$(List(n),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!5!",ValToStr(.lPointerToRawData + _
								IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
				n = n + 1
			End With
		Next i
		List(n) = MsgList(111) & MsgList(111)
	End If
	'.NET CLR 数据目录地址及所在文件节
	If File.LangType = NET_FILE_SIGNATURE Then
		ReDim Preserve List(n + 6 + 7) As String
		List(n + 2) = MsgList(138)
		List(n + 3) = MsgList(111) & MsgList(111)
		List(n + 4) = MsgList(139)
		List(n + 5) = MsgList(111) & MsgList(111)
		n = n + 6
		For i = 0 To 6
			With File.CLRList(i)
				List(n) = Replace$(MsgList(110),"%s!1!",MsgList(i + 140))
				If .lPointerToRawData > 0 Then
					j = SkipSection(File,.lPointerToRawData,0,0,1)
					If j > -1 Then
						List(n) = Replace$(List(n),"%s!2!",File.SecList(j).sName)
					Else
						List(n) = Replace$(List(n),"%s!2!",MsgList(136))
					End If
				ElseIf .lSizeOfRawData > 0 Then
					List(n) = Replace$(List(n),"%s!2!",MsgList(136))
				Else
					List(n) = Replace$(List(n),"%s!2!",MsgList(137))
				End If
				List(n) = Replace$(List(n),"%s!3!","")
				List(n) = Replace$(List(n),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!5!",ValToStr(.lPointerToRawData + _
							IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
				n = n + 1
			End With
		Next i
		List(n) = MsgList(111) & MsgList(111)
	End If
	'.NET 流地址及所在文件节
	If File.NetStreams > 0 Then
		ReDim Preserve List(n + 6 + File.NetStreams) As String
		List(n + 2) = MsgList(147)
		List(n + 3) = MsgList(111) & MsgList(111)
		List(n + 4) = MsgList(148)
		List(n + 5) = MsgList(111) & MsgList(111)
		n = n + 6
		For i = 0 To File.NetStreams - 1
			With File.StreamList(i)
				List(n) = Replace$(MsgList(110),"%s!1!",.sName)
				If .lPointerToRawData > 0 Then
					j = SkipSection(File,.lPointerToRawData,0,0,1)
					If j > -1 Then
						List(n) = Replace$(List(n),"%s!2!",File.SecList(j).sName)
					Else
						List(n) = Replace$(List(n),"%s!2!",MsgList(136))
					End If
				ElseIf .lSizeOfRawData > 0 Then
					List(n) = Replace$(List(n),"%s!2!",MsgList(136))
				Else
					List(n) = Replace$(List(n),"%s!2!",MsgList(137))
				End If
				List(n) = Replace$(List(n),"%s!3!","")
				List(n) = Replace$(List(n),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!5!",ValToStr(.lPointerToRawData + _
							IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
				n = n + 1
			End With
		Next i
		List(n) = MsgList(111) & MsgList(111)
	End If
	File.Info = StrListJoin(List,vbCrLf,True)
	Erase List
	Exit Sub
	'错误处理
	ErrHandle:
	On Error Resume Next
	Err.Source = "NotReadFile"
	Err.Description = Err.Description & JoinStr & File.FilePath & ".xls"
	Call sysErrorMassage(Err,1)
End Sub


'获取选定列表框项目的索引
Private Function GetListBoxIndexs(ByVal hwnd As Long) As Long()
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


'选定指定的列表框项目
'Indexs = -1 全选，否则选择指定项
Private Function SetListBoxItems(ByVal hwnd As Long,ByVal Indexs As Variant,Optional ByVal TopItem As Long = -1) As Boolean
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


'获取主窗口消息字符串
Private Function GetMsgList(MsgList() As String,ByVal Language As String) As Boolean
	Dim i As Integer
	ReDim MsgList(152) As String
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

		MsgList(12) = "字符编码修改器 - 版本 %v (构建 %b)"
		MsgList(13) = "..."
		MsgList(14) = "函数名:"
		MsgList(15) = "偏移地址:"
		MsgList(16) = "引用地址(%i,%s/%d)"
		MsgList(17) = "原始编码"
		MsgList(18) = "使用编码"
		MsgList(19) = "状态: %s"
		MsgList(20) = "关于"
		MsgList(21) = "语言"
		MsgList(22) = "全选"
		MsgList(23) = "更改编码"
		MsgList(24) = "撤消"
		MsgList(25) = "复制"
		MsgList(26) = "写入"
		MsgList(27) = "全部写入"
		MsgList(28) = "过滤显示"
		MsgList(29) = "全部显示"
		MsgList(30) = "取消"
		MsgList(31) = "引用地址: %s!1!\r\n引用代码: %s!2!\r\n所在区段: %s!3!\r\n原始编码: %s!4!\r\n使用编码: %s!5!\r\n编码地址: %s!6!\r\n编码字节长度: %s!7!\r\n编码字节序: %s!8!"
		MsgList(32) = "可更改;未更改;已更改;可更改但未更改;所选原始编码;所选使用编码"
		MsgList(33) = "信息"

		MsgList(34) = "OEM"
		MsgList(35) = "MAC"
		MsgList(36) = "THREAD"
		MsgList(37) = "WEST EUROPE"
		MsgList(38) = "EAST EUROPE"
		MsgList(39) = "RUSSIAN"
		MsgList(40) = "GREEK"
		MsgList(41) = "TURKISH"
		MsgList(42) = "HEBREW"
		MsgList(43) = "ARABIC"
		MsgList(44) = "BALTIC"
		MsgList(45) = "VIETNAMESE"
		MsgList(46) = "JAPAN"
		MsgList(47) = "简体中文 GBK"
		MsgList(48) = "简体中文 GB18030"
		MsgList(49) = "KOREA"
		MsgList(50) = "繁体中文 BIG5"
		MsgList(51) = "THAI"
		MsgList(52) = "UTF-7"
		MsgList(53) = "UTF-8"
		MsgList(54) = "UTF-16LE (UniCode LE)"
		MsgList(55) = "UTF-16BE (Unicode BE)"
		MsgList(56) = "UTF-32LE"
		MsgList(57) = "UTF-32BE"

		MsgList(58) = "版本 %v (构建 %b)\r\n" & _
					"OS 版本: Windows XP/2000 或以上\r\n" & _
					"Passolo 版本: Passolo 5.0 或以上\r\n" & _
					"授权: 免费软件\r\n" & _
					"网址: http://www.hanzify.org\r\n" & _
					"作者: wanfu (2018 - 2019)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(59) = "关于字符编码修改器"
		MsgList(60) = "可执行文件 (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|所有文件 (*.*)|*.*||"
		MsgList(61) = "选择文件"
		MsgList(62) = "英语;简体中文;繁体中文"
		MsgList(63) = "enu;chs;cht"
		MsgList(64) = "正在搜索转码函数，请稍候..."
		MsgList(65) = "正在搜索 %s 函数的引用，请稍候..."
		MsgList(66) = "正在搜索 %s 函数的编码，请稍候..."
		MsgList(67) = "没有找到转码函数。"
		MsgList(68) = "该文件为非 PE 文件。"
		MsgList(69) = "共找到 %s 个转码函数，%d 个调用地址。"
		MsgList(70) = "不能将原文件作为目标文件。"
		MsgList(71) = "正在修改 %s 函数的编码，请稍候..."
		MsgList(72) = "共修改了 %s 处的编码。"
		MsgList(73) = "无引用"
		MsgList(74) = "文件类型：%s!1!\r\n函数名：%s!2!\r\n函数地址：%s!3!\r\n函数所在区段：%s!4!\r\n函数所在 DLL：%s!5!"
		MsgList(75) = "========================================================="
		MsgList(76) = "序号\t引用地址\t引用代码\t所在区段\t编码地址\t原编码\t现编码\t编码字节长度\t编码字节序"
		MsgList(77) = "#%s!1!\t%s!2!\t%s!3!\t%s!4!\t%s!5!\t%s!6!\t%s!7!\t%s!8!\t%s!9!"
		MsgList(78) = "BE"
		MsgList(79) = "LE"
		MsgList(80) = "未知"
		MsgList(81) = "使用函数"
		MsgList(82) = "更改函数"
		MsgList(83) = "引用地址: %s!1!\r\n所在区段: %s!2!\r\n原函数名: %s!3!\r\n原函数地址: %s!4!\r\n原函数引用代码: %s!5!\r\n现函数名: %s!6!\r\n现函数地址: %s!7!\r\n现函数引用代码: %s!8!"
		MsgList(84) = "未更改;已更改;所选函数"
		MsgList(85) = "序号\t引用地址\t所在区段\t原函数名\t原函数地址\t原函数引用代码" & _
					"\r\n\t\t\t\t\t现函数名\t现函数地址\t现函数引用代码
		MsgList(86) = "全部撤消"
		MsgList(87) = "文件信息;选定项;全部可更改;全部未更改;全部已更改;全部可更改但未更改;全部"
		MsgList(88) = "非 PE 文件"
		MsgList(89) = "信息 - %s"
		MsgList(90) = "#%s!1!\t%s!2!\t%s!3!\t%s!4!\t%s!5!\t%s!6!\r\n\t\t\t\t%s!7!\t%s!8!\t%s!9!"

		MsgList(91) = "============ 文件信息 ============\r\n"
		MsgList(92) = "文件名称：\t%s"
		MsgList(93) = "文件路径：\t%s"
		MsgList(94) = "文件说明：\t%s"
		MsgList(95) = "文件版本：\t%s"
		MsgList(96) = "产品名称：\t%s"
		MsgList(97) = "产品版本：\t%s"
		MsgList(98) = "版权所有：\t%s"
		MsgList(99) = "文件大小：\t%s 字节"
		MsgList(100) = "创建日期：\t%s"
		MsgList(101) = "修改日期：\t%s"
		MsgList(102) = "语　　言：\t%s"
		MsgList(103) = "开 发 商：\t%s"
		MsgList(104) = "原始文件名：\t%s"
		MsgList(105) = "内部文件名：\t%s"
		MsgList(106) = "文件类型：\t%s"
		MsgList(107) = "映像基址：\t%s"
		MsgList(108) = "区段信息："
		MsgList(109) = "地址类别\t区段名\t开始地址\t结束地址\t字节大小"
		MsgList(110) = "%s!1!\t%s!2!\t%s!4!\t%s!5!\t%s!6!"
		MsgList(111) = "================================="
		MsgList(112) = "文件偏移地址"
		MsgList(113) = "相对虚拟地址"
		MsgList(114) = "任意"
		MsgList(115) = "隐藏"
		MsgList(116) = "未知"
		MsgList(117) = "不可用"
		MsgList(118) = "数据目录信息 (文件偏移地址)："
		MsgList(119) = "目录名称\t所在区段\t开始地址\t结束地址\t字节大小"
		MsgList(120) = "导出目录"
		MsgList(121) = "导入目录"
		MsgList(122) = "资源目录"
		MsgList(123) = "异常目录"
		MsgList(124) = "安全目录"
		MsgList(125) = "基址重定位表"
		MsgList(126) = "调试目录"
		MsgList(127) = "版权目录"
		MsgList(128) = "机器值(GP RVA)"
		MsgList(129) = "线程本地存储表"
		MsgList(130) = "载入配置目录"
		MsgList(131) = "绑定导入目录"
		MsgList(132) = "导入地址表"
		MsgList(133) = "延迟加载导入符"
		MsgList(134) = "COM 运行库标志"
		MsgList(135) = "保留目录"
		MsgList(136) = "异常"
		MsgList(137) = "不存在"
		MsgList(138) = ".NET CLR 数据目录信息 (文件偏移地址)："
		MsgList(139) = "目录名称\t所在区段\t开始地址\t结束地址\t字节大小"
		MsgList(140) = "元数据(MetaData)"
		MsgList(141) = "托管资源"
		MsgList(142) = "强名称签名"
		MsgList(143) = "代码管理表"
		MsgList(144) = "虚拟表(V-表)"
		MsgList(145) = "跳转导出地址表"
		MsgList(146) = "托管本机映像头"
		MsgList(147) = ".NET MetaData 流信息 (文件偏移地址)："
		MsgList(148) = "流名称\t所在区段\t开始地址\t结束地址\t字节大小"
		MsgList(149) = "非 PE 文件"
		MsgList(150) = "子PE(%s)"
		MsgList(151) = "地址类别\t段名\t节名\t\t开始地址\t结束地址\t字节大小"
		MsgList(152) = "%s!1!\t%s!2!\t%s!3!\t\t%s!4!\t%s!5!\t%s!6!"
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
		MsgList(11) = "眤╰参ぶ ""%s"" 狝叭\r\n岿粇絏: %d岿粇磞瓃: %v"

		MsgList(12) = "じ絪絏э竟 - セ %v (篶 %b)"
		MsgList(13) = "..."
		MsgList(14) = "ㄧ计:"
		MsgList(15) = "熬簿:"
		MsgList(16) = "把酚(%i,%s/%d)"
		MsgList(17) = "﹍絪絏"
		MsgList(18) = "ㄏノ絪絏"
		MsgList(19) = "篈: %s"
		MsgList(20) = "闽"
		MsgList(21) = "粂ē"
		MsgList(22) = "匡"
		MsgList(23) = "跑絪絏"
		MsgList(24) = "临"
		MsgList(25) = "狡籹"
		MsgList(26) = "糶"
		MsgList(27) = "场糶"
		MsgList(28) = "筁耾陪ボ"
		MsgList(29) = "场陪ボ"
		MsgList(30) = ""
		MsgList(31) = "把酚: %s!1!\r\n把酚絏: %s!2!\r\n┮跋琿: %s!3!\r\n﹍絪絏: %s!4!\r\nㄏノ絪絏: %s!5!\r\n絪絏: %s!6!\r\n絪絏じ舱: %s!7!\r\n絪絏じ舱: %s!8!"
		MsgList(32) = "跑;ゼ跑;跑;跑ゼ跑;┮匡﹍絪絏;┮匡ㄏノ絪絏"
		MsgList(33) = "癟"

		MsgList(34) = "OEM"
		MsgList(35) = "MAC"
		MsgList(36) = "THREAD"
		MsgList(37) = "WEST EUROPE"
		MsgList(38) = "EAST EUROPE"
		MsgList(39) = "RUSSIAN"
		MsgList(40) = "GREEK"
		MsgList(41) = "TURKISH"
		MsgList(42) = "HEBREW"
		MsgList(43) = "ARABIC"
		MsgList(44) = "BALTIC"
		MsgList(45) = "VIETNAMESE"
		MsgList(46) = "JAPAN"
		MsgList(47) = "虏砰いゅ GBK"
		MsgList(48) = "虏砰いゅ GB18030"
		MsgList(49) = "KOREA"
		MsgList(50) = "タ砰いゅ BIG5"
		MsgList(51) = "THAI"
		MsgList(52) = "UTF-7"
		MsgList(53) = "UTF-8"
		MsgList(54) = "UTF-16LE (UniCode LE)"
		MsgList(55) = "UTF-16BE (Unicode BE)"
		MsgList(56) = "UTF-32LE"
		MsgList(57) = "UTF-32BE"

		MsgList(58) = "セ %v (篶 %b)\r\n" & _
					"OS セ: Windows XP/2000 ┪\r\n" & _
					"Passolo セ: Passolo 5.0 ┪\r\n" & _
					"甭舦: 禣硁砰\r\n" & _
					"呼: http://www.hanzify.org\r\n" & _
					": wanfu (2018 - 2019)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(59) = "闽じ絪絏э竟"
		MsgList(60) = "磅︽郎 (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|┮Τ郎 (*.*)|*.*||"
		MsgList(61) = "匡郎"
		MsgList(62) = "璣粂;虏砰いゅ;タ砰いゅ"
		MsgList(63) = "enu;chs;cht"
		MsgList(64) = "タ穓锣絏ㄧ计叫祔..."
		MsgList(65) = "タ穓 %s ㄧ计把酚叫祔..."
		MsgList(66) = "タ穓 %s ㄧ计絪絏叫祔..."
		MsgList(67) = "⊿Τт锣絏ㄧ计"
		MsgList(68) = "赣郎獶 PE 郎"
		MsgList(69) = "т %s 锣絏ㄧ计%d 秸ノ"
		MsgList(70) = "ぃ盢郎ヘ夹郎"
		MsgList(71) = "タэ %s ㄧ计絪絏叫祔..."
		MsgList(72) = "э %s 矪絪絏"
		MsgList(73) = "礚把酚"
		MsgList(74) = "郎摸%s!1!\r\nㄧ计%s!2!\r\nㄧ计:%s!3!\r\nㄧ计┮跋琿%s!4!\r\nㄧ计┮ DLL%s!5!"
		MsgList(75) = "========================================================="
		MsgList(76) = "腹\t把酚\t把酚絏\t┮跋琿\t絪絏\t絪絏\t瞷絪絏\t絪絏じ舱\t絪絏じ舱"
		MsgList(77) = "#%s!1!\t%s!2!\t%s!3!\t%s!4!\t%s!5!\t%s!6!\t%s!7!\t%s!8!\t%s!9!"
		MsgList(78) = "BE"
		MsgList(79) = "LE"
		MsgList(80) = "ゼ"
		MsgList(81) = "ㄏノㄧ计"
		MsgList(82) = "跑ㄧ计"
		MsgList(83) = "把酚: %s!1!\r\n┮跋琿: %s!2!\r\nㄧ计: %s!3!\r\nㄧ计: %s!4!\r\nㄧ计把酚絏: %s!5!\r\n瞷ㄧ计: %s!6!\r\n瞷ㄧ计: %s!7!\r\n瞷ㄧ计把酚絏: %s!8!"
		MsgList(84) = "ゼ跑;跑;┮匡ㄧ计"
		MsgList(85) = "腹\t把酚\t┮跋琿\tㄧ计\tㄧ计\tㄧ计把酚絏" & _
					"\r\n\t\t\t\t\t瞷ㄧ计\t瞷ㄧ计\t瞷ㄧ计把酚絏"
		MsgList(86) = "场临"
		MsgList(87) = "郎癟;匡拒兜;场跑;场ゼ跑;场跑;场跑ゼ跑;场"
		MsgList(88) = "獶 PE 郎"
		MsgList(89) = "癟 - %s"
		MsgList(90) = "#%s!1!\t%s!2!\t%s!3!\t%s!4!\t%s!5!\t%s!6!\r\n\t\t\t\t%s!7!\t%s!8!\t%s!9!"

		MsgList(91) = "============ 郎癟 ============\r\n"
		MsgList(92) = "郎嘿\t%s"
		MsgList(93) = "郎隔畖\t%s"
		MsgList(94) = "郎弧\t%s"
		MsgList(95) = "郎セ\t%s"
		MsgList(96) = "玻珇嘿\t%s"
		MsgList(97) = "玻珇セ\t%s"
		MsgList(98) = "舦┮Τ\t%s"
		MsgList(99) = "郎\t%s じ舱"
		MsgList(100) = "ミら戳\t%s"
		MsgList(101) = "эら戳\t%s"
		MsgList(102) = "粂ē\t%s"
		MsgList(103) = "秨 祇 坝\t%s"
		MsgList(104) = "﹍郎\t%s"
		MsgList(105) = "ず场郎\t%s"
		MsgList(106) = "郎摸\t%s"
		MsgList(107) = "琈钩膀\t%s"
		MsgList(108) = "跋琿癟"
		MsgList(109) = "摸\t跋琿\t秨﹍\t挡\tじ舱"
		MsgList(110) = "%s!1!\t%s!2!\t%s!4!\t%s!5!\t%s!6!"
		MsgList(111) = "================================="
		MsgList(112) = "郎熬簿"
		MsgList(113) = "癸店览"
		MsgList(114) = "ヴ種"
		MsgList(115) = "留旅"
		MsgList(116) = "ゼ"
		MsgList(117) = "ぃノ"
		MsgList(118) = "戈ヘ魁癟 (郎熬簿)"
		MsgList(119) = "ヘ魁嘿\t┮跋琿\t秨﹍\t挡\tじ舱"
		MsgList(120) = "蹲ヘ魁"
		MsgList(121) = "蹲ヘ魁"
		MsgList(122) = "戈方ヘ魁"
		MsgList(123) = "钵盽ヘ魁"
		MsgList(124) = "ヘ魁"
		MsgList(125) = "膀﹚"
		MsgList(126) = "禘耞ヘ魁"
		MsgList(127) = "舦ヘ魁"
		MsgList(128) = "诀竟(GP RVA)"
		MsgList(129) = "磅︽狐セ诀纗"
		MsgList(130) = "更砞﹚ヘ魁"
		MsgList(131) = "竕﹚蹲ヘ魁"
		MsgList(132) = "蹲"
		MsgList(133) = "┑筐更蹲才"
		MsgList(134) = "COM 磅︽畐篨夹"
		MsgList(135) = "玂痙ヘ魁"
		MsgList(136) = "钵盽"
		MsgList(137) = "ぃ"
		MsgList(138) = ".NET CLR 戈ヘ魁癟 (郎熬簿)"
		MsgList(139) = "ヘ魁嘿\t┮跋琿\t秨﹍\t挡\tじ舱"
		MsgList(140) = "い膥戈(MetaData)"
		MsgList(141) = "Λ恨戈方"
		MsgList(142) = "眏嘿帽"
		MsgList(143) = "絏恨瞶"
		MsgList(144) = "店览(V-)"
		MsgList(145) = "铬臘蹲"
		MsgList(146) = "Λ恨セ诀琈钩繷"
		MsgList(147) = ".NET MetaData 戈瑈癟 (郎熬簿)"
		MsgList(148) = "戈瑈嘿\t┮跋琿\t秨﹍\t挡\tじ舱"
		MsgList(149) = "獶 PE 郎"
		MsgList(150) = "PE(%s)"
		MsgList(151) = "摸\t琿\t竊\t\t秨﹍\t挡\tじ舱"
		MsgList(152) = "%s!1!\t%s!2!\t%s!3!\t\t%s!4!\t%s!5!\t%s!6!"
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
					"Please restart the Passolo try and please report to the software developer."
		MsgList(6) = "The following file is missing [%s] section.\r\n%d"
		MsgList(7) = "The following file is missing [%s] Value.\r\n%d"
		MsgList(8) = "The following contents of the file is not correct.\r\n%s"
		MsgList(9) = "The following file does not exist! Please check and try again.\r\n%s"
		MsgList(10) = "The following file version is %d, requires version at least %v.\r\n%s"
		MsgList(11) = "Your system is missing %s server.\r\nError Code: %d, Content: %v"

		MsgList(12) = "Character Encoding Modifier - Version %v (Build %b)"
		MsgList(13) = "..."
		MsgList(14) = "Function:"
		MsgList(15) = "File Offset:"
		MsgList(16) = "Ref. Add(%i,%s/%d)"
		MsgList(17) = "Original Encoding"
		MsgList(18) = "Use Encoding"
		MsgList(19) = "Status: %s"
		MsgList(20) = "About"
		MsgList(21) = "Language"
		MsgList(22) = "Select All"
		MsgList(23) = "Change Encoding"
		MsgList(24) = "Undo"
		MsgList(25) = "Info"
		MsgList(26) = "Write"
		MsgList(27) = "Write All"
		MsgList(28) = "Filter Show"
		MsgList(29) = "Show All"
		MsgList(30) = "Cancel"
		MsgList(31) = "Reference Add: %s!1!\r\nReference Code: %s!2!\r\nIn Section: %s!3!\r\nOriginal Encoding: %s!4!\r\nUse Encoding: %s!5!\r\n" & _
					"Encoding Address: %s!6!\r\nEncoding Byte Length: %s!7!\r\nEncoding Byte Order: %s!8!"
		MsgList(32) = "Changeable;Unchanged;Changed;Changeable But Unchanged;Original Encoding Selected;Use Encoding Selected"
		MsgList(33) = "Information"

		MsgList(34) = "OEM"
		MsgList(35) = "MAC"
		MsgList(36) = "THREAD"
		MsgList(37) = "WEST EUROPE"
		MsgList(38) = "EAST EUROPE"
		MsgList(39) = "RUSSIAN"
		MsgList(40) = "GREEK"
		MsgList(41) = "TURKISH"
		MsgList(42) = "HEBREW"
		MsgList(43) = "ARABIC"
		MsgList(44) = "BALTIC"
		MsgList(45) = "VIETNAMESE"
		MsgList(46) = "JAPAN"
		MsgList(47) = "CHINA GBK"
		MsgList(48) = "CHINA GB18030"
		MsgList(49) = "KOREA"
		MsgList(50) = "TAIWAN"
		MsgList(51) = "THAI"
		MsgList(52) = "UTF-7"
		MsgList(53) = "UTF-8"
		MsgList(54) = "UTF-16LE (UniCode LE)"
		MsgList(55) = "UTF-16BE (Unicode BE)"
		MsgList(56) = "UTF-32LE"
		MsgList(57) = "UTF-32BE"

		MsgList(58) = "Version: %v (Build %b)\r\n" & _
					"OS Version: Windows XP/2000 or higher\r\n" & _
					"Passolo Version: Passolo 5.0 or higher\r\n" & _
					"License: Freeware\r\n" & _
					"HomePage: http://www.hanzify.org\r\n" & _
					"Author: wanfu (2018 - 2019)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(59) = "About Character Encoding Modifier"
		MsgList(60) = "Executable File (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|All File (*.*)|*.*||"
		MsgList(61) = "Select File"
		MsgList(62) = "EngLish;Chinese Simplified;Chinese Traditional"
		MsgList(63) = "enu;chs;cht"

		MsgList(64) = "Searching for transcoding function..."
		MsgList(65) = "Searching for reference of %s function..."
		MsgList(66) = "Searching for encoding of %s function..."
		MsgList(67) = "No transcoding function was found."
		MsgList(68) = "The file is not a PE file."
		MsgList(69) = "Total %s functions and %d reference addresses been found."
		MsgList(70) = "You cannot use the original file as the target file."
		MsgList(71) = "Modifing for encoding of %s function..."
		MsgList(72) = "Total %s encodings has been modified."
		MsgList(73) = "No references"
		MsgList(74) = "FileType: %s!1!\r\nFunction Name: %s!2!\r\nFunction Address: %s!3!\r\nFunction in Section: %s!4!\r\nFunction in Dll: %s!5!"
		MsgList(75) = "========================================================="
		MsgList(76) = "No.\tReference Address\tReference Code\tIn Section\tCode Address\tOld Encoding\tNew Encoding\tEncoding Byte Length\tEncoding Byte Order"
		MsgList(77) = "#%s!1!\t%s!2!\t%s!3!\t%s!4!\t%s!5!\t%s!6!\t%s!7!\t%s!8!\t%s!9!"
		MsgList(78) = "BE"
		MsgList(79) = "LE"
		MsgList(80) = "Unknown"
		MsgList(81) = "Use Function"
		MsgList(82) = "Change Function"
		MsgList(83) = "Reference Add: %s!1!\r\nIn Section: %s!2!\r\nOld Function Name: %s!3!\r\nOld Function Address: %s!4!\r\nOld Function Reference Code: %s!5!\r\n" & _
					"New Function Name: %s!6!\r\nNew Function Address: %s!7!\r\nNew Function Reference Code: %s!8!"
		MsgList(84) = "Unchanged;Changed;Function Selected"
		MsgList(85) = "No.\tReference Address\tIn Section\tOld Function Name\t\tOld Function Address\tOld Function Reference Code" & _
					"\r\n\t\t\t\t\tNew Function Name\tNew Function Address\tNew Function Reference Code"
		MsgList(86) = "Undo All"
		MsgList(87) = "File Information;Selected Items;All Changeable;All Unchanged;All Changed;All Changeable But Unchanged;All"
		MsgList(88) = "Not PE File"
		MsgList(89) = "Information - %s"
		MsgList(90) = "#%s!1!\t%s!2!\t%s!3!\t%s!4!\t%s!5!\t%s!6!\r\n\t\t\t\t%s!7!\t%s!8!\t%s!9!"

		MsgList(91) = "============ File Information ============\r\n"
		MsgList(92) = "File Name:\t%s"
		MsgList(93) = "File Path:\t\t%s"
		MsgList(94) = "Description:\t%s"
		MsgList(95) = "Version:\t\t%s"
		MsgList(96) = "Product Name:\t%s"
		MsgList(97) = "Product Version:\t%s"
		MsgList(98) = "Legal Copyright:\t%s"
		MsgList(99) = "File Size:\t\t%s bytes"
		MsgList(100) = "Date Created:\t%s"
		MsgList(101) = "Date Modified:\t%s"
		MsgList(102) = "Language:\t%s"
		MsgList(103) = "Company Name:\t%s"
		MsgList(104) = "Original File Name:\t%s"
		MsgList(105) = "Internal File Name:\t%s"
		MsgList(106) = "File Type:\t\t%s"
		MsgList(107) = "Image Base:\t%s"
		MsgList(108) = "Section Information:"
		MsgList(109) = "Address Category\tSection Name\tStart Address\tEnd Address\tByte Size"
		MsgList(110) = "%s!1!\t\t%s!2!\t%s!4!\t%s!5!\t%s!6!"
		MsgList(111) = "====================================="
		MsgList(112) = "Offset"
		MsgList(113) = "RVA"
		MsgList(114) = "Any"
		MsgList(115) = "Hide"
		MsgList(116) = "Unknown"
		MsgList(117) = "Not Available"
		MsgList(118) = "Data Directory Information (offset):"
		MsgList(119) = "Directory Name\t\t\tIn Section\tStart Address\tEnd Address\tByte Size"
		MsgList(120) = "Export Directory\t"
		MsgList(121) = "Import Directory\t"
		MsgList(122) = "Resource Directory\t"
		MsgList(123) = "Exception Directory\t"
		MsgList(124) = "Security Directory\t"
		MsgList(125) = "Base Relocation Table"
		MsgList(126) = "Debug Directory\t"
		MsgList(127) = "Copyright\t\t"
		MsgList(128) = "RVA of GP\t"
		MsgList(129) = "TLS Directory\t"
		MsgList(130) = "Load Configuration Directory"
		MsgList(131) = "Bound Import Directory"
		MsgList(132) = "Import Address Table"
		MsgList(133) = "Delay Load Import Descriptor"
		MsgList(134) = "COM Runtime Descriptor"
		MsgList(135) = "Reserved Directory\t"
		MsgList(136) = "Exception"
		MsgList(137) = "Not Exist"
		MsgList(138) = ".NET CLR Data Directory Information (offset):"
		MsgList(139) = "Directory Name\t\t\tIn Section\tStart Address\tEnd Address\tByte Size"
		MsgList(140) = "Meta Data\t"
		MsgList(141) = "Managed Resource\t"
		MsgList(142) = "Strong Name Signature"
		MsgList(143) = "Code Manager Table"
		MsgList(144) = "V-Table Fixups\t"
		MsgList(145) = "Export Address Table Jumps"
		MsgList(146) = "Managed Native Heade"
		MsgList(147) = ".NET MetaData iStreams Information (offset):"
		MsgList(148) = "Stream Name\tIn Section\tStart Address\tEnd Address\tByte Size"
		MsgList(149) = "Not PE File"
		MsgList(150) = "Sub PE(%s)"
		MsgList(151) = "Address Category\tSegment Name\tSection Name\tStart Address\tEnd Address\tByte Size"
		MsgList(152) = "%s!1!\t\t%s!2!\t\t%s!3!\t\t%s!4!\t%s!5!\t%s!6!"
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
