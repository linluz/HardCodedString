Attribute VB_Name = "modPEInfo"
'' File Information Module for PSlHardCodedString.bas
'' (c) 2015-2019 by wanfu (Last modified on 2019.11.08)

'#Uses "modCommon.bas"

Option Explicit

Private Const PE_BIT_TYPE32 = 224 + 24
Private Const PE_BIT_TYPE64 = 240 + 24

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
	DataDirectory(15) 				As IMAGE_DATA_DIRECTORY
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

'***************************************************************************************************************************************************
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


'获取文件版本信息
Public Function GetFileInfo(ByVal strFilePath As String,File As FILE_PROPERTIE) As Boolean
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


'获取文件及子文件的数据结构信息
Public Function GetPEHeaders(ByVal strFilePath As String,File As FILE_PROPERTIE,ByVal Mode As Long) As Boolean
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

		'对齐最后节
		'tmpSecHeader(.MaxSecID).lSizeOfRawData = Alignment(tmpSecHeader(.MaxSecID).lSizeOfRawData,.FileAlign,1)
		'tmpSecHeader(.MaxSecID).lVirtualSize = Alignment(tmpSecHeader(.MaxSecID).lVirtualSize,.SecAlign,1)
		'.SecList(.MaxSecID).lSizeOfRawData = tmpSecHeader(.MaxSecID).lSizeOfRawData
		'.SecList(.MaxSecID).lVirtualSize = tmpSecHeader(.MaxSecID).lVirtualSize

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


'增加文件节长度
'fType = 0 获取可增加的字节数，修改原始文件头不写入
'fType = 1 写入指定长度对齐后的字节数，不修改原始文件头仅写入
'fType = 2 写入指定长度对齐后的字节数，修改原始文件头
'fType = 3 不修改不写入，仅获取可增加值(AddSecSize(x).Length 为偏移大小，AddSecSize(x).Address 为虚拟大小)
'AddSecSize(x).Length = 0，按可增加的最大值增加，否则按 AddSecSize(x).Length 对齐值增加
Public Function AddPESectionSize(trnFile As FILE_PROPERTIE,AddSecSize() As FREE_BTYE_SPACE,ByVal fType As Long,ByVal Mode As Long) As Long
	Dim i As Long,j As Integer,k As Long,x As Long,n As Long
	Dim AddRAW As Long,AddRVA As Long,PEBitType As Long
	Dim FN As FILE_IMAGE,TempByte() As Byte,File As FILE_PROPERTIE

	'错误处理
	On Error GoTo localError

	'获取 PE 头
	File = trnFile
	If GetPEHeaders(File.FilePath,File,Mode) = False Then
		If RefTypeList(0).sName = "" Then Exit Function
		File.Magic = RefTypeList(0).FileMagic
	End If

	'修改文件节的开始地址和大小
	With File
		'修改文件对齐值，以减少新增节的大小
		If Selected(12) = "1" Then
			If .FileAlign > 512 Then
				If .FileAlign Mod 512 = 0 Then
					.FileAlign = 512
					AddRVA = 1
				End If
			End If
		End If
		For i = 0 To UBound(AddSecSize)
			j = AddSecSize(i).inSectionID
			'根据是否是最后节获取可增加字节
			If j = .MaxSecID Then
				k = .SecList(j).lSizeOfRawData + AddSecSize(i).Length
				n = 1
			Else
				k = .SecList(GetSectionID(File,j,-1,True)).lVirtualAddress - .SecList(j).lVirtualAddress
				n = 0
			End If
			'按文件对齐值对齐
			x = Alignment(k,.FileAlign,n) - .SecList(j).lSizeOfRawData
			If x > 0 Or n > 0 Then
				'根据实际需要增加当前节的偏移大小并对齐
				If AddSecSize(i).Length > 0 Then
					x = Alignment(IIf(x > AddSecSize(i).Length,AddSecSize(i).Length,x),.FileAlign,1)
				Else
					x = Alignment(x,.FileAlign,n)
				End If
				AddSecSize(i).Length = x: AddRAW = AddRAW + x

				'增加当前节的虚拟大小，虚拟大小不用对齐
				If AddSecSize(i).Length > 0 Then
					x = Alignment(k,.SecAlign,n) - .SecList(j).lVirtualSize
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
		AddPESectionSize = AddRAW
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
		AddPESectionSize = AddRAW
		Exit Function
	End If

	'打开文件
	Mode = LoadFile(File.FilePath,FN,0,1,0,Mode)
	If Mode < -1 Then GoTo localError

	'修改 OptionalHeader 数据并写入
	Select Case File.Magic
	Case "PE32","NET32"
		PEBitType = PE_BIT_TYPE32
		If AddRVA > 0 Then
			'修改文件对齐值，以减少文件大小
			OptionalHeader32.lFileAlignment = File.FileAlign
			'修改文件头的映像大小并节对齐
			OptionalHeader32.lSizeOfImage = File.SecList(File.MaxSecID).lVirtualAddress + File.SecList(File.MaxSecID).lVirtualSize
			'If PutTypeValue(FN,DosHeader.lPointerToPEHeader + Len(FileHeader),OptionalHeader32,Mode) = False Then GoTo localError
			Select Case Mode
			Case Is < 0
				Put #FN.hFile,DosHeader.lPointerToPEHeader + Len(FileHeader) + 1,OptionalHeader32
			Case 0
				CopyMemory FN.ImageByte(DosHeader.lPointerToPEHeader + Len(FileHeader)),OptionalHeader32,Len(OptionalHeader32)
			Case Else
				WriteMemory FN.MappedAddress + DosHeader.lPointerToPEHeader + Len(FileHeader),OptionalHeader32,Len(OptionalHeader32)
			End Select
		End If
	Case "PE64","NET64"
		PEBitType = PE_BIT_TYPE64
		If AddRVA > 0 Then
			'修改文件对齐值，以减少文件大小
			OptionalHeader64.lFileAlignment = File.FileAlign
			'修改文件头的映像大小并节对齐
			OptionalHeader64.lSizeOfImage = File.SecList(File.MaxSecID).lVirtualAddress + File.SecList(File.MaxSecID).lVirtualSize
			'If PutTypeValue(FN,DosHeader.lPointerToPEHeader + Len(FileHeader),OptionalHeader64,Mode) = False Then GoTo localError
			Select Case Mode
			Case Is < 0
				Put #FN.hFile,DosHeader.lPointerToPEHeader + Len(FileHeader) + 1,OptionalHeader64
			Case 0
				CopyMemory FN.ImageByte(DosHeader.lPointerToPEHeader + Len(FileHeader)),OptionalHeader64,Len(OptionalHeader64)
			Case Else
				WriteMemory FN.MappedAddress + DosHeader.lPointerToPEHeader + Len(FileHeader),OptionalHeader64,Len(OptionalHeader64)
			End Select
		End If
	End Select

	'修改 SecHeader 数据并写入
	If PEBitType <> 0 Then
		For i = 0 To File.MaxSecIndex - 1
			If i > 0 Then SecHeader(i).lPointerToRawData = File.SecList(i).lPointerToRawData
			SecHeader(i).lSizeOfRawData = File.SecList(i).lSizeOfRawData
			SecHeader(i).lVirtualSize = File.SecList(i).lVirtualSize
		Next i
		'If PutTypeArray(FN,DosHeader.lPointerToPEHeader + PEBitType,SecHeader,Mode) = False Then GoTo localError
		Select Case Mode
		Case Is < 0
			Put #FN.hFile,DosHeader.lPointerToPEHeader + PEBitType + 1,SecHeader
		Case 0
			CopyMemory FN.ImageByte(DosHeader.lPointerToPEHeader + PEBitType),SecHeader(0),Len(SecHeader(0)) * File.MaxSecIndex
		Case Else
			WriteMemory FN.MappedAddress + DosHeader.lPointerToPEHeader + PEBitType,SecHeader(0),Len(SecHeader(0)) * File.MaxSecIndex
		End Select
	End If

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
				If j = File.MaxSecID And trnFile.SecList(trnFile.MaxSecIndex).lVirtualSize < 1 Then
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
	AddPESectionSize = AddRAW

	'修改目标文件的区段数据
	If fType = 2 Then trnFile = File
	Exit Function

	'安全退出函数
	localError:
	UnLoadFile(FN,0,Mode)
	AddPESectionSize = -1
End Function


'在文件尾部增加一个文件节
'fType = 0 只修改文件节数据不写入
'fType = 1 不修改文件节数据但写入
'fType = 2 既修改文件节数据又写入
Public Function AddPESection(trnFile As FILE_PROPERTIE,AddSecSize As FREE_BTYE_SPACE,ByVal SecName As String,ByVal fType As Long,ByVal Mode As Long) As Long
	Dim i As Long,k As Long
	Dim OrgHeadersOffset As Long,NewHeadersOffset As Long,FristSecOffset As Long,PEBitType As Long
	Dim FN As FILE_IMAGE,TempByte() As Byte,File As FILE_PROPERTIE

	'错误处理
	On Error GoTo localError
	If AddSecSize.Length = 0 Then Exit Function

	'获取 PE 头
	File = trnFile
	If GetPEHeaders(File.FilePath,File,Mode) = False Then Exit Function

	'修改文件对齐值，以减少新增节的大小
	If Selected(12) = "1" Then
		If File.FileAlign > 512 Then
			If (File.FileAlign Mod 512) = 0 Then File.FileAlign = 512
		End If
	End If

	'计算 NT 头的大小和映像大小
	FristSecOffset = SecHeader(File.MinSecID).lPointerToRawData
	Select Case File.Magic
	Case "PE32", "NET32"
		'计算文件头实际占用大小
		OrgHeadersOffset = DosHeader.lPointerToPEHeader + Len(FileHeader) + Len(OptionalHeader32) + Len(SecHeader(0)) * File.MaxSecIndex
		'搜索文件头最大节后面有没有其他数据
		For i = FristSecOffset - 1 To 0 Step -1
			If GetByte(FN,i,Mode) > 0 Then
				k = i + 1
				Exit For
			End If
		Next i
		'计算新增节表后需要的文件头大小
		If k > OrgHeadersOffset Then
			NewHeadersOffset = k + Len(SecHeader(0))
			If NewHeadersOffset > File.SecAlign Then
				NewHeadersOffset = OrgHeadersOffset + Len(SecHeader(0))
				k = -1
			End If
		Else
			NewHeadersOffset = OrgHeadersOffset + Len(SecHeader(0))
		End If
		'当文件头大小超过节对齐值时取消增加新节
		If NewHeadersOffset > File.SecAlign Then GoTo localError
		'按文件对齐值对齐需要的文件头大小
		If NewHeadersOffset > FristSecOffset Then
			NewHeadersOffset = Alignment(NewHeadersOffset,File.FileAlign,1)
			OptionalHeader32.lSizeOfHeaders = NewHeadersOffset
		End If
	Case "PE64", "NET64"
		'计算文件头实际占用大小
		OrgHeadersOffset = DosHeader.lPointerToPEHeader + Len(FileHeader) + Len(OptionalHeader64) + Len(SecHeader(0)) * File.MaxSecIndex
		'搜索文件头最大节后面有没有其他数据
		For i = FristSecOffset - 1 To 0 Step -1
			If GetByte(FN,i,Mode) > 0 Then
				k = i + 1
				Exit For
			End If
		Next i
		'计算新增节表后需要的文件头大小
		If k > OrgHeadersOffset Then
			NewHeadersOffset = k + Len(SecHeader(0))
			If NewHeadersOffset > File.SecAlign Then
				NewHeadersOffset = OrgHeadersOffset + Len(SecHeader(0))
				k = -1
			End If
		Else
			NewHeadersOffset = OrgHeadersOffset + Len(SecHeader(0))
		End If
		'当文件头大小超过节对齐值时取消增加新节
		If NewHeadersOffset > File.SecAlign Then GoTo localError
		'按文件对齐值对齐需要的文件头大小
		If NewHeadersOffset > FristSecOffset Then
			NewHeadersOffset = Alignment(NewHeadersOffset,File.FileAlign,1)
			OptionalHeader64.lSizeOfHeaders = NewHeadersOffset
		End If
	End Select
	'增加节数
	FileHeader.iNumberOfSections = FileHeader.iNumberOfSections + 1

	'因插入新的区段信息，故可能会增加文件头
	If NewHeadersOffset > FristSecOffset Then
		For i = 0 To File.MaxSecIndex - 1
			If File.SecList(i).lPointerToRawData > 0 Then
				File.SecList(i).lPointerToRawData = File.SecList(i).lPointerToRawData + NewHeadersOffset - FristSecOffset
			End If
		Next i
	End If

	'计算新增节地址及大小，用于移位计算
	With File
		.SecList(.MaxSecIndex).sName = SecName
		.SecList(.MaxSecIndex).lPointerToRawData = .SecList(.MaxSecID).lPointerToRawData + Alignment(.SecList(.MaxSecID).lSizeOfRawData,.FileAlign,1)
		.SecList(.MaxSecIndex).lSizeOfRawData = Alignment(AddSecSize.Length,.FileAlign,1)
		.SecList(.MaxSecIndex).lVirtualAddress = .SecList(.MaxSecID).lVirtualAddress + Alignment(.SecList(.MaxSecID).lVirtualSize,.SecAlign,1)
		.SecList(.MaxSecIndex).lVirtualSize = AddSecSize.Length
		.SecList(.MaxSecIndex).SubSecs = 0
		.SecList(.MaxSecIndex).RWA = .SecList(.MaxSecIndex).lPointerToRawData
	End With

	'修改隐藏节的偏移地址，还原隐藏节大小
	'由于重新读取的文件可能已在文件尾部写入了字串，这些字节将作为隐藏节被读取，故需要使用原始文件的隐藏节信息
	ReDim Preserve File.SecList(File.MaxSecIndex + 1) 'As SECTION_PROPERTIE
	ReDim Preserve File.SecList(File.MaxSecIndex + 1).SubSecList(0) 'As SUB_SECTION_PROPERTIE
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
		AddPESection = File.SecList(File.MaxSecIndex).lSizeOfRawData
		If fType = 0 Then
			File.MaxSecID = File.MaxSecIndex: File.MaxSecIndex = File.MaxSecIndex + 1: trnFile = File
		End If
		Exit Function
	End	If

	'打开文件
	Mode = LoadFile(File.FilePath,FN,0,1,0,Mode)
	If Mode < -1 Then GoTo localError

	'移位原文件的隐藏节及隐藏节后的全部字节长度
	'由于重新读取的文件可能已在文件尾部写入了字串，这些字节将作为隐藏节被读取，故需要使用原始文件的隐藏节信息
	'必须使用虚拟大小，因为其大小为最大节后的所有字节(包括子 PE)
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

	'复制 FristSecOffset 后的数据到 NewSizeOfHeaders
	If NewHeadersOffset > FristSecOffset Then
		'复制 FristSecOffset 至原文件最大区段地址到 NewHeadersOffset
		i = trnFile.SecList(trnFile.MaxSecID).lPointerToRawData + _
			trnFile.SecList(trnFile.MaxSecID).lSizeOfRawData - FristSecOffset
		TempByte = GetBytes(FN,i,FristSecOffset,Mode)
		PutBytes(FN,NewHeadersOffset,TempByte,i,Mode)
		'清空 FristSecOffset 到 NewSizeOfHeaders 之间的数据
		i = NewHeadersOffset - FristSecOffset
		ReDim TempByte(i - 1) As Byte
		PutBytes(FN,FristSecOffset,TempByte,i,Mode)
	End If

	'节表结束后的多余字节处理 (一般系脱壳留下)
	If k > OrgHeadersOffset Then
		'移动 OrgHeadersOffset 到 k 之间的字节为一个节数据大小
		TempByte = GetBytes(FN,k - OrgHeadersOffset,OrgHeadersOffset,Mode)
		PutBytes(FN,OrgHeadersOffset + Len(SecHeader(0)),TempByte,k - OrgHeadersOffset,Mode)
	ElseIf k < 0 Then
		'置空节表结束到原始节尾之间的无用字节
		ReDim TempByte(FristSecOffset - OrgHeadersOffset - 1) As Byte
		PutBytes(FN,OrgHeadersOffset,TempByte,FristSecOffset - OrgHeadersOffset,Mode)
	End If

	'写入已修改的 FileHeader 数据
	'If PutTypeValue(FN,.DosHeader.lPointerToPEHeader,FileHeader,Mode) = False Then GoTo localError
	Select Case Mode
	Case Is < 0
		Put #FN.hFile,DosHeader.lPointerToPEHeader + 1,FileHeader
	Case 0
		CopyMemory FN.ImageByte(DosHeader.lPointerToPEHeader),FileHeader,Len(FileHeader)
	Case Else
		WriteMemory FN.MappedAddress + DosHeader.lPointerToPEHeader,FileHeader,Len(FileHeader)
	End Select

	'修改 OptionalHeader 数据并写入
	Select Case File.Magic
	Case "PE32","NET32"
		PEBitType = PE_BIT_TYPE32
		'修改文件对齐值，以减少文件大小
		OptionalHeader32.lFileAlignment = File.FileAlign
		'修改文件头的映像大小并节对齐
		OptionalHeader32.lSizeOfImage = File.SecList(File.MaxSecIndex).lVirtualAddress + File.SecList(File.MaxSecIndex).lVirtualSize
		'If PutTypeValue(FN,.DosHeader.lPointerToPEHeader + Len(FileHeader),OptionalHeader32,Mode) = False Then GoTo localError
		Select Case Mode
		Case Is < 0
			Put #FN.hFile,DosHeader.lPointerToPEHeader + Len(FileHeader) + 1,OptionalHeader32
		Case 0
			CopyMemory FN.ImageByte(DosHeader.lPointerToPEHeader + Len(FileHeader)),OptionalHeader32,Len(OptionalHeader32)
		Case Else
			WriteMemory FN.MappedAddress + DosHeader.lPointerToPEHeader + Len(FileHeader),OptionalHeader32,Len(OptionalHeader32)
		End Select
	Case "PE64","NET64"
		PEBitType = PE_BIT_TYPE64
		'修改文件对齐值，以减少文件大小
		OptionalHeader64.lFileAlignment = File.FileAlign
		'修改文件头的映像大小并节对齐
		OptionalHeader64.lSizeOfImage = File.SecList(File.MaxSecIndex).lVirtualAddress + File.SecList(File.MaxSecIndex).lVirtualSize
		'If PutTypeValue(FN,DosHeader.lPointerToPEHeader + Len(FileHeader),OptionalHeader64,Mode) = False Then GoTo localError
		Select Case Mode
		Case Is < 0
			Put #FN.hFile,DosHeader.lPointerToPEHeader + Len(FileHeader) + 1,OptionalHeader64
		Case 0
			CopyMemory FN.ImageByte(DosHeader.lPointerToPEHeader + Len(FileHeader)),OptionalHeader64,Len(OptionalHeader64)
		Case Else
			WriteMemory FN.MappedAddress + DosHeader.lPointerToPEHeader + Len(FileHeader),OptionalHeader64,Len(OptionalHeader64)
		End Select
	End Select

	'更新新增节地址及大小，用于写入
	If NewHeadersOffset > FristSecOffset Then
		For i = 0 To File.MaxSecIndex - 1
			SecHeader(i).lPointerToRawData = File.SecList(i).lPointerToRawData
			SecHeader(i).lSizeOfRawData = File.SecList(i).lSizeOfRawData
			SecHeader(i).lVirtualSize = File.SecList(i).lVirtualSize
		Next i
	End If
	'设置新区段的属性
	ReDim Preserve SecHeader(File.MaxSecIndex) 'As IMAGE_SECTION_HEADER
	For i = 1 To Len(SecName)
		SecHeader(File.MaxSecIndex).sName(i - 1) = AscW(Mid$(SecName,i,1))
	Next i
	SecHeader(File.MaxSecIndex).lPointerToRawData = File.SecList(File.MaxSecIndex).lPointerToRawData
	SecHeader(File.MaxSecIndex).lSizeOfRawData = File.SecList(File.MaxSecIndex).lSizeOfRawData
	SecHeader(File.MaxSecIndex).lVirtualAddress = File.SecList(File.MaxSecIndex).lVirtualAddress
	SecHeader(File.MaxSecIndex).lVirtualSize = File.SecList(File.MaxSecIndex).lVirtualSize
	SecHeader(File.MaxSecIndex).lCharacteristics = IMAGE_SCN_MEM_EXECUTE Or IMAGE_SCN_MEM_READ Or IMAGE_SCN_MEM_WRITE
	'If PutTypeArray(FN,DosHeader.lPointerToPEHeader + PEBitType,SecHeader,Mode) = False Then GoTo localError
	Select Case Mode
	Case Is < 0
		Put #FN.hFile,DosHeader.lPointerToPEHeader + PEBitType + 1,SecHeader
	Case 0
		CopyMemory FN.ImageByte(DosHeader.lPointerToPEHeader + PEBitType),SecHeader(0),Len(SecHeader(0)) * (File.MaxSecIndex + 1)
	Case Else
		WriteMemory FN.MappedAddress + DosHeader.lPointerToPEHeader + PEBitType,SecHeader(0),Len(SecHeader(0)) * (File.MaxSecIndex + 1)
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
	AddPESection = File.SecList(File.MaxSecIndex).lSizeOfRawData

	'修改目标文件的区段数据
	If fType = 2 Then
		File.MaxSecID = File.MaxSecIndex: File.MaxSecIndex = File.MaxSecIndex + 1: trnFile = File
	End If
	Exit Function

	'安全退出函数
	localError:
	UnLoadFile(FN,0,Mode)
	AddPESection = -1
End Function
