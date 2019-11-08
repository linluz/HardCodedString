'' PESubFile for Passolo
'' (c) 2018 - 2019 by wanfu (Last modified on 2019.11.08)

'' Command Line Format: Command <FilePath>
'' Command: Name of this Macros file
'' FilePath: Full path of PE file.
'' Return: No
'' For example: modPESubFile,"d:\my folder\my file.exe"

Option Explicit

Private Const Version = "2019.05.21"
Private Const Build = "190521"
Private Const JoinStr = vbFormFeed  'vbBack
Private Const TextJoinStr = vbCrLf
Private Const LoadMode = 0&
Private Const AppName = "PESubFile"
Private Const DataFileName = "SubFileList.dat"

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
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" ( _
	ByVal CodePage As Long, _
	ByVal dwFlags As Long, _
	ByVal lpWideCharStr As Long, _
	ByVal cchWideChar As Long, _
	ByRef lpMultiByteStr As Any, _
	ByVal cchMultiByte As Long, _
	ByVal lpDefaultChar As Long, _
	ByVal lpUsedDefaultChar As Long) As Long

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

'浏览文件夹
Private Type BrowseInfo
	hWndOwner		As Long		'浏览文件夹对话框的父窗体句柄
	pIDLRoot		As Long		'ITEMIDLIST结构的地址，包含浏览时的初始根目录，可以是NULL，此时桌面目录将被使用
	pszDisplayName	As Long		'用来保存用户选中的目录字符串的内存地址
	lpszTitle		As String	'显示位于对话框左上部的标题
	ulFlags			As Long		'指定对话框的外观和功能的标志
	lpfnCallback	As Long		'处理事件的回调函数
	lParam			As Long		'应用程序传给回调函数的参数
	iImage			As Long		'保存被选取的文件夹的图片索引
End Type

'浏览文件夹参数
Private Enum BrowseFolder
	BIF_RETURNONLYFSDIRS = &H1		'仅返回文件系统的目录
	BIF_DONTGOBELOWDOMAIN = &H2		'在树形视窗中，不包含域名底下的网络目录结构
	BIF_STATUSTEXT = &H4&			'包含一个状态区域。通过给对话框发送消息使回调函数设置状态文本
	BIF_EDITBOX = &H10				'包含一个编辑框，用户可以输入选中项的名字
	BIF_BROWSEINCLUDEURLS = &H80
	BIF_RETURNFSANCESTORS = &H8		'返回文件系统的一个节点
	BIF_VALIDATE = &H20				'没有BIF_EDITBOX标志位时，该标志位被忽略。如果用户输入的名字非法，将发送BFFM_VALIDATEFAILED消息给回调函数
	BIF_NEWDIALOGSTYLE = &H40
	BIF_USENEWUI = BIF_EDITBOX Or BIF_NEWDIALOGSTYLE	'对话框上有新建文件夹按钮
	BIF_UAHINT = &H100
	BIF_NONEWFOLDERBUTTON = &H200
	BIF_NOTRANSLATETARGETS = &H400
	BIF_BROWSEFORCOMPUTER = &H1000	'返回计算机名
	BIF_BROWSEFORPRINTER = &H2000	'返回打印机名
	BIF_BROWSEINCLUDEFILES = &H4000	'浏览器将显示目录，同时也显示文件
	BIF_SHAREABLE = &H8000
	BIF_BROWSEFILEJUNCTIONS = &H10000
End Enum

'浏览文件夹函数
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" ( _
	ByVal pidList As Long, _
	ByVal lpBuffer As String) As Long

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

'子文件属性
Private Type SUB_FILE
	FileName			As String	'文件名 (不含路径)
	FilePath			As String	'文件路径 (含文件名)
	FileSize			As Long		'原文件大小
	NewFileSize			As Long		'现文件大小
	FileAdd				As Long		'文件类型的开始地址在主文件中的文件偏移
	Info 				As String	'文件所有信息，避免重复获取
End Type

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

Private MsgList() As String,RegExp As Object,StrList() As String,SubFileNameList() As String
Private MainFile As FILE_PROPERTIE,SubFileList() As SUB_FILE,FindValue As String,FilterValue As String


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
	Begin Dialog UserDialog 620,392,Replace$(Replace$(MsgList(12),"%v",Version),"%b",Build),.MainDlgFunc ' %GRID:10,7,1,1
		TextBox 0,0,0,21,.SuppValueBox
		TextBox 10,7,570,21,.FilePathBox
		TextBox 10,7,570,21,.SplitStateBox
		TextBox 10,7,570,21,.SubFileNameBox
		PushButton 580,7,30,21,MsgList(13),.FilePathButton

		Text 10,38,150,14,MsgList(14),.AddText
		Text 170,38,270,14,MsgList(15),.SubFileText
		MultiListBox 10,56,160,308,TempList(),.AddList
		MultiListBox 170,56,300,308,TempList(),.SubFileList
		Text 10,367,600,14,Replace$(MsgList(16),"%s",""),.StatusText

		PushButton 480,56,130,21,MsgList(17),.AboutButton
		PushButton 480,77,130,21,MsgList(18),.LangButton
		PushButton 480,105,130,21,MsgList(19),.SplitButton
		PushButton 480,126,130,21,MsgList(20),.ImportButton
		PushButton 480,154,130,21,MsgList(124),.SelectAllButton
		PushButton 480,175,130,21,MsgList(21),.CopyButton
		PushButton 480,196,130,21,MsgList(22),.MainFileInfoButton
		PushButton 480,217,130,21,MsgList(23),.SubFileInfoButton
		PushButton 480,245,130,21,MsgList(125),.FindButton
		PushButton 480,266,130,21,MsgList(126),.FilterButton
		PushButton 480,287,130,21,MsgList(129),.ShowAllButton
		PushButton 480,315,130,21,MsgList(24),.DirectMergeButton
		PushButton 480,336,130,21,MsgList(25),.AlignMergeButton
		CancelButton 480,28,130,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then Exit Sub
	Exit Sub
	SysErrorMsg:
	If Err.Source <> "ExitSub" Then Call sysErrorMassage(Err,0)
End Sub


'请务必查看对话框帮助主题以了解更多信息。
Private Function MainDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,Temp As String,IntList() As Long,TempList() As String
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		DlgText "SuppValueBox",CStr$(SuppValue)
		DlgVisible "SuppValueBox",False
		DlgVisible "CancelButton",False
		DlgVisible "SplitStateBox",False
		DlgVisible "SubFileNameBox",False

		DlgEnable "FilePathBox",False
		DlgEnable "SelectAllButton",False
		DlgEnable "CopyButton",False
		DlgEnable "FindButton",False
		DlgEnable "FilterButton",False
		DlgEnable "ShowAllButton",False
		DlgEnable "SubFileInfoButton",False
		DlgEnable "DirectMergeButton",False
		DlgEnable "AlignMergeButton",False
		'转递参数值
		MainFile.FilePath = Command
		If Dir$(MainFile.FilePath) = "" Then MainFile.FilePath = ""
		Temp = MainFile.FilePath
		If Len(Temp) > 70 Then
			Temp = Left$(Temp,InStr(Temp,"\")) & "..." & Right(Temp,70 - Len(Left$(Temp,InStr(Temp,"\"))))
		End If
		DlgText "FilePathBox",Temp
		ReDim SubFileList(0) As SUB_FILE,SubFileNameList(0) As String
		If MainFile.FilePath = "" Then
			DlgEnable "SplitButton",False
			DlgEnable "ImportButton",False
			DlgEnable "MainFileInfoButton",False
			DlgText "SplitStateBox","0"
		Else
			Call GetFileInfo(MainFile.FilePath,MainFile)
			Call GetHeaders(MainFile.FilePath,MainFile,LoadMode,MainFile.FileType)
			If MainFile.NumberOfSub = 0 Then
				DlgEnable "SplitButton",False
				DlgEnable "ImportButton",False
			End If
			DlgEnable "MainFileInfoButton",True
			DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(112)),"%s",CStr$(MainFile.NumberOfSub))
			DlgText "SplitStateBox","1"
		End If
		DlgText "AddText",Replace$(Replace$(MsgList(14),"%s","0"),"%d","0")
	Case 2 ' 数值更改或者按下按钮时
		MainDlgFunc = True ' 防止按下按钮时关闭对话框窗口
		Select Case DlgItem$
		Case "CancelButton"
			MainDlgFunc = False
		Case "FilePathButton"
			If PSL.SelectFile(Temp,True,MsgList(28),MsgList(29)) = False Then Exit Function
			If MainFile.FilePath = Temp Then Exit Function
			If IsOpen(Temp,2,0) = True Then Exit Function
			MainFile.FilePath = Temp
			If Len(Temp) > 70 Then
				Temp = Left$(Temp,InStr(Temp,"\")) & "..." & Right(Temp,70 - Len(Left$(Temp,InStr(Temp,"\"))))
			End If
			DlgText "FilePathBox",Temp
			ReDim SubFileList(0) As SUB_FILE,SubFileNameList(0) As String
			ReDim StrList(0) As String,TempList(0) As String
			ReDim TempList(0) As String
			DlgListBoxArray "AddList",TempList()
			DlgValue "AddList",Array(0)
			DlgListBoxArray "SubFileList",TempList()
			DlgValue "SubFileList",Array(0)
			MainFile.Info = ""
			Call GetFileInfo(MainFile.FilePath,MainFile)
			Call GetHeaders(MainFile.FilePath,MainFile,LoadMode,MainFile.FileType)
			If MainFile.NumberOfSub = 0 Then
				DlgEnable "SplitButton",False
				DlgEnable "ImportButton",False
			Else
				DlgEnable "SplitButton",True
				DlgEnable "ImportButton",True
			End If
			DlgEnable "SelectAllButton",False
			DlgEnable "CopyButton",False
			DlgEnable "FindButton",False
			DlgEnable "FilterButton",False
			DlgEnable "ShowAllButton",False
			DlgEnable "MainFileInfoButton",True
			DlgEnable "SubFileInfoButton",False
			DlgEnable "DirectMergeButton",False
			DlgEnable "AlignMergeButton",False
			DlgText "StatusText",Replace$(MsgList(16),"%s",MsgList(32))
			DlgText "AddText",Replace$(Replace$(MsgList(14),"%s","0"),"%d","0")
			DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(112)),"%s",CStr$(MainFile.NumberOfSub))
			DlgText "SplitStateBox","1"
		Case "AddList"
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AddList")))
			If CheckArrEmpty(IntList) = False Then Exit Function
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("SubFileList")),IntList)
			DlgText "AddText",Replace$(Replace$(MsgList(14),"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(DlgListBoxArray("AddList")))
		Case "SubFileList"
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("SubFileList")))
			If CheckArrEmpty(IntList) = False Then Exit Function
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AddList")),IntList)
			DlgText "AddText",Replace$(Replace$(MsgList(14),"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(DlgListBoxArray("AddList")))
		Case "AboutButton"
			MsgBox Replace$(Replace$(MsgList(26),"%v",Version),"%b",Build),vbOkOnly+vbInformation,MsgList(27)
		Case "LangButton"
			ReDim TempList(0) As String
			TempList = ReSplit(MsgList(30),";")
			i = ShowPopupMenu(TempList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			TempList = ReSplit(MsgList(31),";")
			If GetMsgList(MsgList,TempList(i)) = False Then Exit Function
			ReDim StrList(0) As String
			MainFile.Info = ""
			For i = 0 To UBound(SubFileList)
				SubFileList(i).Info = ""
			Next i
			'更改文本框语言
			DlgText -1,Replace$(Replace$(MsgList(12),"%v",Version),"%b",Build)
			DlgText "FilePathButton",MsgList(13)
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AddList")))
			If CheckArrEmpty(IntList) = False Then
				DlgText "AddText",Replace$(Replace$(MsgList(14),"%s","0"),"%d",CStr$(DlgListBoxArray("AddList")))
			Else
				DlgText "AddText",Replace$(Replace$(MsgList(14),"%s",CStr$(UBound(IntList) + 1)),"%d",CStr$(DlgListBoxArray("AddList")))
			End If
			DlgText "SubFileText",MsgList(15)
			Select Case DlgText("SplitStateBox")
			Case "","0"
				DlgText "StatusText",Replace$(MsgList(16),"%s","")
			Case "1"
				DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(112)),"%s",CStr$(MainFile.NumberOfSub))
			Case "2"
				DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(33)),"%s",CStr$(DlgListBoxArray("AddList")))
			Case "3"
				DlgText "StatusText",Replace$(Replace$(Replace$(MsgList(16),"%s",MsgList(34)), _
										"%s",CStr$(DlgListBoxArray("AddList"))),"%d",MainFile.FileName & "_SubFiles")
			Case "4"
				DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(33)),"%s","0")
			Case "5"
				DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(45)),"%s",CStr$(DlgListBoxArray("AddList")))
			Case "6"
				DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(45)),"%s","0")
			Case "7"
				DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(108)),"%s",DlgText("SubFileNameBox"))
			Case "8"
				DlgText "StatusText",Replace$(MsgList(16),"%s",MsgList(38))
			Case "9"
				DlgText "StatusText",Replace$(MsgList(16),"%s",MsgList(39))
			End Select
			DlgText "AboutButton",MsgList(17)
			DlgText "LangButton",MsgList(18)
			DlgText "SplitButton",MsgList(19)
			DlgText "ImportButton",MsgList(20)
			DlgText "SelectAllButton",MsgList(124)
			DlgText "CopyButton",MsgList(21)
			DlgText "MainFileInfoButton",MsgList(22)
			DlgText "SubFileInfoButton",MsgList(23)
			DlgText "FindButton",MsgList(125)
			DlgText "FilterButton",MsgList(126)
			DlgText "ShowAllButton",MsgList(129)
			DlgText "DirectMergeButton",MsgList(24)
			DlgText "AlignMergeButton",MsgList(25)
		Case "SplitButton"
			If BrowseForFolder(Temp,MsgList(40)) = False Then Exit Function
			If Temp = "" Then Exit Function
			DlgText "StatusText",Replace$(MsgList(16),"%s",MsgList(32))
			If SplitFile(Temp,MainFile,SubFileList,LoadMode,GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("StatusText"))) = True Then
				Call WriteDataToFile(Temp,MainFile,SubFileList)
				TempList = GetDataValueList(SubFileList,MainFile.FileSize,"",0)
				DlgListBoxArray "AddList",TempList()
				DlgValue "AddList",Array(0)
				SubFileNameList = GetDataValueList(SubFileList,MainFile.FileSize,"",1)
				DlgListBoxArray "SubFileList",SubFileNameList()
				DlgValue "SubFileList",Array(0)
				DlgEnable "SelectAllButton",True
				DlgEnable "CopyButton",True
				DlgEnable "FindButton",True
				DlgEnable "FilterButton",True
				DlgEnable "ShowAllButton",False
				DlgEnable "SubFileInfoButton",True
				DlgEnable "DirectMergeButton",True
				DlgEnable "AlignMergeButton",True
				DlgText "AddText",Replace$(Replace$(MsgList(14),"%s","1"),"%d",CStr$(DlgListBoxArray("AddList")))
				If Temp = MainFile.SubFileDir Then
					DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(33)),"%s",CStr$(DlgListBoxArray("AddList")))
					DlgText "SplitStateBox","2"
				Else
					DlgText "StatusText",Replace$(Replace$(Replace$(MsgList(16),"%s",MsgList(34)), _
										"%s",CStr$(DlgListBoxArray("AddList"))),"%d",MainFile.FileName & "_SubFiles")
					DlgText "SplitStateBox","3"
				End If
			Else
				ReDim SubFileNameList(0) As String,TempList(0) As String
				DlgListBoxArray "AddList",TempList()
				DlgValue "AddList",Array(0)
				DlgListBoxArray "SubFileList",SubFileNameList()
				DlgValue "SubFileList",Array(0)
				DlgEnable "SelectAllButton",False
				DlgEnable "CopyButton",False
				DlgEnable "FindButton",False
				DlgEnable "FilterButton",False
				DlgEnable "ShowAllButton",False
				DlgEnable "SubFileInfoButton",False
				DlgEnable "DirectMergeButton",False
				DlgEnable "AlignMergeButton",False
				DlgText "AddText",Replace$(Replace$(MsgList(14),"%s","0"),"%d","0")
				DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(33)),"%s","0")
				DlgText "SplitStateBox","4"
			End If
		Case "ImportButton"
			If BrowseForFolder(Temp,MsgList(41)) = False Then Exit Function
			If Temp = "" Then Exit Function
			DlgText "StatusText",Replace$(MsgList(16),"%s",MsgList(44))
			i = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("StatusText"))
			i = ImportSubFile(Temp,MainFile,SubFileList,i)
			If i = -2 Then
				i = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("StatusText"))
				If SplitFileOld(Temp,MainFile,SubFileList,LoadMode,i) = False Then i = -1 Else i = 0
			End If
			Select Case i
			Case 0
				TempList = GetDataValueList(SubFileList,MainFile.FileSize,"",0)
				DlgListBoxArray "AddList",TempList()
				DlgValue "AddList",Array(0)
				SubFileNameList = GetDataValueList(SubFileList,MainFile.FileSize,"",1)
				DlgListBoxArray "SubFileList",SubFileNameList()
				DlgValue "SubFileList",Array(0)
				DlgEnable "SelectAllButton",True
				DlgEnable "CopyButton",True
				DlgEnable "FindButton",True
				DlgEnable "FilterButton",True
				DlgEnable "ShowAllButton",False
				DlgEnable "SubFileInfoButton",True
				DlgEnable "DirectMergeButton",True
				DlgEnable "AlignMergeButton",True
				DlgText "AddText",Replace$(Replace$(MsgList(14),"%s","1"),"%d",CStr$(DlgListBoxArray("AddList")))
				DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(45)),"%s",CStr$(DlgListBoxArray("AddList")))
				DlgText "SplitStateBox","5"
			Case 1
				MsgBox MsgList(114),vbOkOnly+vbInformation,MsgList(113)
			Case 2
				MsgBox MsgList(115),vbOkOnly+vbInformation,MsgList(113)
			Case 3
				MsgBox Replace$(MsgList(116),"%s",MainFile.FileName & "_" & DataFileName),vbOkOnly+vbInformation,MsgList(113)
			Case 4
				MsgBox MsgList(117),vbOkOnly+vbInformation,MsgList(113)
			Case 5
				MsgBox MsgList(118),vbOkOnly+vbInformation,MsgList(113)
			Case 6
				MsgBox MsgList(119),vbOkOnly+vbInformation,MsgList(113)
			Case 7
				MsgBox MsgList(120),vbOkOnly+vbInformation,MsgList(113)
			Case 8
				MsgBox Replace$(MsgList(122),"%s",MainFile.FileName & "_" & DataFileName),vbOkOnly+vbInformation,MsgList(113)
			Case 9
				MsgBox MsgList(123),vbOkOnly+vbInformation,MsgList(113)
			Case Is < 0
				ReDim SubFileNameList(0) As String,TempList(0) As String
				DlgListBoxArray "AddList",TempList()
				DlgValue "AddList",Array(0)
				DlgListBoxArray "SubFileList",SubFileNameList()
				DlgValue "SubFileList",Array(0)
				DlgEnable "SelectAllButton",False
				DlgEnable "CopyButton",False
				DlgEnable "FindButton",False
				DlgEnable "FilterButton",False
				DlgEnable "ShowAllButton",False
				DlgEnable "SubFileInfoButton",False
				DlgEnable "DirectMergeButton",False
				DlgEnable "AlignMergeButton",False
				DlgText "AddText",Replace$(Replace$(MsgList(14),"%s","0"),"%d","0")
				DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(45)),"%s","0")
				DlgText "SplitStateBox","6"
			End Select
		Case "SelectAllButton"
			i = DlgListBoxArray("AddList")
			If i < 1 Then Exit Function
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AddList")),-1)
			SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("SubFileList")),-1)
			DlgText "AddText",Replace$(Replace$(MsgList(14),"%s",CStr$(i)),"%d",CStr$(i))
		Case "CopyButton"
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AddList")))
			If CheckArrEmpty(IntList) = False Then Exit Function
			ReDim TempList(UBound(IntList)) As String
			For i = 0 To UBound(IntList)
				TempList(i) = SubFileNameList(IntList(i))
			Next i
			Clipboard StrListJoin(TempList,TextJoinStr)
		Case "MainFileInfoButton"
			If MainFile.FilePath = "" Then Exit Function
			ReDim TempList(0) As String,StrList(0) As String
			If MainFile.Info = "" Then
				Call FileInfoView(MainFile,True)
			End If
			TempList(0) = DlgText("FilePathBox")
			StrList(0) = MainFile.Info
			Call ShowFileInfo(TempList,StrList)
		Case "SubFileInfoButton"
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AddList")))
			If CheckArrEmpty(IntList) = False Then Exit Function
			i = UBound(IntList)
			ReDim TempList(i) As String,StrList(i) As String
			Dim TempFile As FILE_PROPERTIE
			For i = 0 To UBound(IntList)
				If SubFileList(IntList(i)).info = "" Then
					DlgText "StatusText",Replace$(Replace$(MsgList(16),"%s",MsgList(108)),"%s",SubFileList(IntList(i)).FileName)
					DlgText "SubFileNameBox",SubFileList(IntList(i)).FileName
					DlgText "SplitStateBox","7"
					TempFile.FilePath = SubFileList(IntList(i)).FilePath
					Call GetFileInfo(TempFile.FilePath,TempFile)
					Call GetHeaders(TempFile.FilePath,TempFile,LoadMode,TempFile.FileType)
					Call FileInfoView(TempFile,True)
					SubFileList(IntList(i)).Info = TempFile.Info
				End If
				TempList(i) = SubFileList(IntList(i)).FilePath
				If Len(TempList(i)) > 70 Then
					TempList(i) = Left$(TempList(i),InStr(TempList(i),"\")) & "..." & _
								Right(TempList(i),70 - Len(Left$(TempList(i),InStr(TempList(i),"\"))))
				End If
				StrList(i) = SubFileList(IntList(i)).Info
			Next i
			Call ShowFileInfo(TempList,StrList)
		Case "FindButton"
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AddList")))
			If CheckArrEmpty(IntList) = False Then Exit Function
			Do
				FindValue = InputBox(MsgList(130),MsgList(127),FindValue)
				If Trim$(FindValue) = "" Then Exit Function
				Select Case FilterStr("CheckRegExp",FindValue,GetFindMode(FindValue))
				Case -2
					MsgBox MsgList(134),vbOkOnly+vbInformation,MsgList(0)
				Case -3
					MsgBox MsgList(135),vbOkOnly+vbInformation,MsgList(0)
				Case Else
					Exit Do
				End Select
			Loop
			i = FindString(SubFileNameList,FindValue,IntList(0))
			If i < 0 Then i = FindString(SubFileNameList,FindValue,-1)
			Select Case i
			Case IntList(0)
				MsgBox Replace$(MsgList(132),"%s",FindValue),vbOkOnly+vbInformation,MsgList(113)
			Case Is > -1
				DlgValue "AddList",Array(i)
				DlgValue "SubFileList",Array(i)
			Case Else
				MsgBox Replace$(MsgList(133),"%s",FindValue),vbOkOnly+vbInformation,MsgList(113)
			End Select
		Case "FilterButton"
			TempList = ReSplit(MsgList(43),";")
			i = ShowPopupMenu(TempList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			If i = UBound(TempList) Then
				Do
					FilterValue = InputBox(MsgList(131),MsgList(128),FilterValue)
					If Trim$(FilterValue) = "" Then Exit Function
					Select Case FilterStr("CheckRegExp",FilterValue,GetFindMode(FilterValue))
					Case -2
						MsgBox MsgList(136),vbOkOnly+vbInformation,MsgList(0)
					Case -3
						MsgBox MsgList(137),vbOkOnly+vbInformation,MsgList(0)
					Case Else
						Exit Do
					End Select
				Loop
			End If
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AddList")))
			TempList = GetDataValueList(SubFileList,MainFile.FileSize,FilterValue,2 + 2 * i)
			DlgListBoxArray "AddList",TempList()
			TempList = SubFileNameList
			SubFileNameList = GetDataValueList(SubFileList,MainFile.FileSize,FilterValue,3 + 2 * i)
			DlgListBoxArray "SubFileList",SubFileNameList()
			If CheckArrEmpty(IntList) = False Then
				DlgValue "AddList",Array(0)
				DlgValue "SubFileList",Array(0)
			Else
				Call GetStrIndexList(TempList,SubFileNameList,IntList)
				SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AddList")),IntList)
				SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("SubFileList")),IntList)
			End If
			If DlgListBoxArray("AddList") < 1 Then
				DlgEnable "SelectAllButton",False
				DlgEnable "CopyButton",False
				DlgEnable "FindButton",False
				DlgEnable "SubFileInfoButton",False
				DlgText "AddText",Replace$(Replace$(MsgList(14),"%s","0"),"%d","0")
			Else
				DlgEnable "SelectAllButton",True
				DlgEnable "CopyButton",True
				DlgEnable "FindButton",True
				DlgEnable "SubFileInfoButton",True
				DlgText "AddText",Replace$(Replace$(MsgList(14),"%s","1"),"%d",CStr$(DlgListBoxArray("AddList")))
			End If
			DlgEnable "ShowAllButton",IIf(DlgListBoxArray("AddList") > UBound(SubFileList),False,True)
		Case "ShowAllButton"
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AddList")))
			TempList = GetDataValueList(SubFileList,MainFile.FileSize,"",0)
			DlgListBoxArray "AddList",TempList()
			TempList = SubFileNameList
			SubFileNameList = GetDataValueList(SubFileList,MainFile.FileSize,"",1)
			DlgListBoxArray "SubFileList",SubFileNameList()
			If CheckArrEmpty(IntList) = False Then
				DlgValue "AddList",Array(0)
				DlgValue "SubFileList",Array(0)
			Else
				Call GetStrIndexList(TempList,SubFileNameList,IntList)
				SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AddList")),IntList)
				SetListBoxItems(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("SubFileList")),IntList)
			End If
			If DlgListBoxArray("AddList") < 1 Then
				DlgEnable "SelectAllButton",False
				DlgEnable "CopyButton",False
				DlgEnable "FindButton",False
				DlgEnable "SubFileInfoButton",False
				DlgText "AddText",Replace$(Replace$(MsgList(14),"%s","0"),"%d","0")
			Else
				DlgEnable "SelectAllButton",True
				DlgEnable "CopyButton",True
				DlgEnable "FindButton",True
				DlgEnable "SubFileInfoButton",True
				DlgText "AddText",Replace$(Replace$(MsgList(14),"%s","1"),"%d",CStr$(DlgListBoxArray("AddList")))
			End If
			DlgEnable "ShowAllButton",False
		Case "DirectMergeButton","AlignMergeButton"
			If DlgText("FilePathButton") = "" Then Exit Function
			If PSL.SelectFile(Temp,False,MsgList(28),MsgList(29)) = False Then Exit Function
			If InStr(Temp,"\") Then
				If (Mid$(Temp,InStrRev(Temp,"\")) Like "*.*") = False Then
					If (MainFile.FilePath Like "*.*") = True Then
						Temp = Temp & Mid$(MainFile.FilePath,InStrRev(MainFile.FilePath,"."))
					End If
				End If
			ElseIf (Temp Like "*.*") = False Then
				If (MainFile.FilePath Like "*.*") = True Then
					Temp = Temp & Mid$(MainFile.FilePath,InStrRev(MainFile.FilePath,"."))
				End If
			End If
			If Temp = MainFile.FilePath Then
				MsgBox MsgList(36),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
			On Error Resume Next
			FileCopy MainFile.FilePath,Temp
			If Err.Number <> 0 Then
				Err.Source = "NotWriteFile"
				Err.Description = Err.Description & JoinStr & Temp
				Call sysErrorMassage(Err,2)
				Exit Function
			End If
			On Error GoTo 0
			DlgText "StatusText",Replace$(MsgList(16),"%s",MsgList(37))
			If MergeFile(Temp,MainFile,SubFileList,LoadMode,IIf(DlgItem$ = "DirectMergeButton",False,True)) = True Then
				DlgText "StatusText",Replace$(MsgList(16),"%s",MsgList(38))
				DlgText "SplitStateBox","8"
			Else
				DlgText "StatusText",Replace$(MsgList(16),"%s",MsgList(39))
				DlgText "SplitStateBox","9"
			End If
		End Select
	'Case 3 ' 文本框或者组合框文本更改时
	Case 6 ' 功能键
		Select Case SuppValue
		Case 1
			MsgBox Replace$(Replace$(MsgList(26),"%v",Version),"%b",Build),vbOkOnly+vbInformation,MsgList(27)
		Case 3
			IntList = GetListBoxIndexs(GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("AddList")))
			If CheckArrEmpty(IntList) = False Then Exit Function
			If FindValue = "" Then
				Do
					FindValue = InputBox(MsgList(130),MsgList(127),FindValue)
					If Trim$(FindValue) = "" Then Exit Function
					Exit Do
				Loop
			End If
			i = FindString(SubFileNameList,FindValue,IntList(0))
			If i = -1 Then i = FindString(SubFileNameList,FindValue,-1)
			Select Case i
			Case IntList(0)
				MsgBox Replace$(MsgList(132),"%s",FindValue),vbOkOnly+vbInformation,MsgList(113)
			Case Is > -1
				DlgValue "AddList",Array(i)
				DlgValue "SubFileList",Array(i)
			Case -1
				MsgBox Replace$(MsgList(133),"%s",FindValue),vbOkOnly+vbInformation,MsgList(113)
			Case -3
				MsgBox MsgList(134),vbOkOnly+vbInformation,MsgList(0)
			Case -4
				MsgBox MsgList(135),vbOkOnly+vbInformation,MsgList(0)
			End Select
		End Select
	End Select
End Function


'找出二个字符数组中的值相同或不相同的索引列表
'Mode = False 获取二个字符数组中的值相同的索引列表，二个索引列表没有对应关系
'Mode = True 获取二个字符数组中的值不相同的索引列表，二个索引列表没有对应关系
Private Function GetStrIndexList(SrcList() As String,TrgList() As String,IntList() As Long,Optional ByVal Mode As Boolean) As Boolean
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
		GetStrIndexList = True
	Else
		ReDim IntList(0) As Long
	End If
End Function


'拆分和导入子文件
Private Function SplitFileOld(ByVal trgFolder As String,File As FILE_PROPERTIE,DataList() As SUB_FILE, _
				ByVal Mode As Long,Optional ByVal ShowMsg As Long) As Boolean
	Dim i As Long,n As Long,FN As FILE_IMAGE,FN2 As Variant,Bytes() As Byte,Msg As String
	On Error GoTo ExitFunction
	If File.NumberOfSub = 0 Then
		ReDim DataList(0) As SUB_FILE
		Exit Function
	End If
	'打开文件
	Mode = LoadFile(File.FilePath,FN,0,0,0,Mode)
	If Mode < -1 Then Exit Function
	File.SubFileDir = trgFolder
	If Right$(File.SubFileDir ,1) <> "\" Then File.SubFileDir = File.SubFileDir & "\"
	If ShowMsg > 0 Then
		Msg = GetTextBoxString(ShowMsg) & " "
	ElseIf ShowMsg < 0 Then
		ReDim TempList(PSL.OutputWnd(0).LineCount - 1) As String
		For i = 1 To PSL.OutputWnd(0).LineCount
			TempList(i - 1) = PSL.OutputWnd(0).Text(i)
		Next i
		Msg = StrListJoin(TempList,vbCrLf) & " "
	End If
	'获取子文件头
	With File.SecList(File.MaxSecIndex)
		ReDim TempList(0) As String
		i = .lPointerToRawData + .lSizeOfRawData
		TempList(0) = ByteToString(GetBytes(FN,File.FileSize - i,i,Mode),CP_ISOLATIN1)
		If InStr(File.Magic,"MAC") Then
			TempList = GetVAListRegExp(TempList(0),"(\xCE\xFA\xED\xFE)|(\xCF\xFA\xED\xFE)",i)
		Else
			TempList = GetVAListRegExp(TempList(0),"MZ[\x00-\xFF]{64,384}?PE\x00",i)
		End If
		If CheckArray(TempList) = False Then GoTo ExitFunction
		ReDim DataList(File.NumberOfSub - 1) As SUB_FILE
		For i = 0 To File.NumberOfSub - 1
			If i < File.NumberOfSub - 1 Then
				DataList(n).FileSize = CLng(TempList(i + 1)) - CLng(TempList(i))
			Else
				DataList(n).FileSize = File.FileSize - CLng(TempList(i))
			End If
			If DataList(n).FileSize > 0 Then
				DataList(n).FilePath = File.SubFileDir & File.FileName & "_" & i
				If Dir$(DataList(n).FilePath) <> "" Then
					DataList(n).FileName = File.FileName & "_" & i
					DataList(n).FileAdd = CLng(TempList(i))
					DataList(n).NewFileSize = FileLen(DataList(n).FilePath)
					n = n + 1
				End If
			End If
			If ShowMsg > 0 Then
				SetTextBoxString ShowMsg,Msg & Format$(i / File.NumberOfSub,"#%")
			ElseIf ShowMsg < 0 Then
				PSL.OutputWnd(0).Clear
				PSL.Output Msg & Format$(i / File.NumberOfSub,"#%")
			End If
		Next i
	End With
	If n > 0 Then
		SplitFileOld = True
		ReDim Preserve DataList(n - 1) As SUB_FILE
	Else
		ReDim DataList(0) As SUB_FILE
	End If
	ExitFunction:
	'关闭文件
	On Error Resume Next
	UnLoadFile(FN,0,Mode)
End Function


'拆分子文件列表
Private Function SplitFile(trgFolder As String,File As FILE_PROPERTIE,DataList() As SUB_FILE,ByVal Mode As Long, _
				Optional ByVal ShowMsg As Long,Optional ByVal fType As Boolean) As Boolean
	Dim i As Long,n As Long,Dic As Object,Msg As String,Temp As String
	Dim FN As FILE_IMAGE,FN2 As Variant,Bytes() As Byte,SubFile As FILE_PROPERTIE
	On Error GoTo ExitFunction
	If File.NumberOfSub = 0 Then
		ReDim DataList(0) As SUB_FILE
		Exit Function
	End If
	'打开文件
	Mode = LoadFile(File.FilePath,FN,0,0,0,Mode)
	If Mode < -1 Then Exit Function
	Temp = trgFolder
	If Right$(Temp ,1) <> "\" Then Temp = Temp & "\"
	If Dir$(Temp & "*.*") <> "" Then
		Temp = Temp & File.FileName & "_SubFiles\"
		trgFolder = Temp
		If Dir$(Temp & "*.*") <> "" Then DelDirs(Temp)
	End If
	If MkSubDir(Temp) = False Then GoTo ExitFunction
	File.SubFileDir = Temp
	If ShowMsg > 0 Then
		Msg = GetTextBoxString(ShowMsg) & " "
	ElseIf ShowMsg < 0 Then
		ReDim TempList(PSL.OutputWnd(0).LineCount - 1) As String
		For i = 1 To PSL.OutputWnd(0).LineCount
			TempList(i - 1) = PSL.OutputWnd(0).Text(i)
		Next i
		Msg = StrListJoin(TempList,vbCrLf) & " "
	End If
	i = InStrRev(File.FileName,".")
	If i > 1 Then
		SubFile.Info = Left$(File.FileName,i - 1)
		Temp = Mid$(File.FileName,i)
	Else
		SubFile.Info = File.FileName
		Temp = ""
	End If
	'获取子文件头
	With File.SecList(File.MaxSecIndex)
		ReDim TempList(0) As String
		i = .lPointerToRawData + .lSizeOfRawData
		TempList(0) = ByteToString(GetBytes(FN,File.FileSize - i,i,Mode),CP_ISOLATIN1)
		If InStr(File.Magic,"MAC") Then
			TempList = GetVAListRegExp(TempList(0),"(\xCE\xFA\xED\xFE)|(\xCF\xFA\xED\xFE)",i)
		Else
			TempList = GetVAListRegExp(TempList(0),"MZ[\x00-\xFF]{64,384}?PE\x00",i)
		End If
		If CheckArray(TempList) = False Then GoTo ExitFunction
		Set Dic = CreateObject("Scripting.Dictionary")
		Dic.Add(File.FileName,0)
		ReDim DataList(File.NumberOfSub - 1) As SUB_FILE
		For i = 0 To File.NumberOfSub - 1
			If i < File.NumberOfSub - 1 Then
				DataList(n).FileSize = CLng(TempList(i + 1)) - CLng(TempList(i))
			Else
				DataList(n).FileSize = File.FileSize - CLng(TempList(i))
			End If
			If DataList(n).FileSize > 0 Then
				'获取子文件
				DataList(n).FilePath = File.SubFileDir & SubFile.Info & "_" & n & Temp
				DataList(n).FileName = SubFile.Info & "_" & n & Temp
				DataList(n).FileAdd = CLng(TempList(i))
				Bytes = GetBytes(FN,DataList(n).FileSize,CLng(TempList(i)),Mode)
				FN2 = FreeFile
				Open DataList(n).FilePath For Binary Access Write Lock Write As #FN2
				Put #FN2,1,Bytes
				Close #FN2
				With SubFile
					.FilePath = DataList(n).FilePath
					.OrigionalFileName = ""
					Call GetFileInfo(.FilePath,SubFile)
					If .OrigionalFileName <> "" Then
						If Not Dic.Exists(.OrigionalFileName) Then
							Dic.Add(.OrigionalFileName,0)
							DataList(n).FileName = .OrigionalFileName
						Else
							.MinSecID = Dic.Item(.OrigionalFileName) + 1
							Dic.Item(.OrigionalFileName) = .MinSecID
							.MaxSecID = InStrRev(.OrigionalFileName,".")
							If .MaxSecID > 1 Then
								DataList(n).FileName = Left$(.OrigionalFileName,.MaxSecID - 1) & _
												"_" & CStr$(.MinSecID) & Mid$(.OrigionalFileName,.MaxSecID)
							Else
								DataList(n).FileName = .OrigionalFileName & "_" & .MinSecID
							End If
						End If
						DataList(n).FilePath = File.SubFileDir & DataList(n).FileName
						Name .FilePath As DataList(n).FilePath
					End If
				End With
				DataList(n).NewFileSize = DataList(n).FileSize
				n = n + 1
			End If
			If ShowMsg > 0 Then
				SetTextBoxString ShowMsg,Msg & Format$(i / File.NumberOfSub,"#%")
			ElseIf ShowMsg < 0 Then
				PSL.OutputWnd(0).Clear
				PSL.Output Msg & Format$(i / File.NumberOfSub,"#%")
			End If
		Next i
		If n > 0 Then
			SplitFile = True
			ReDim Preserve DataList(n - 1) As SUB_FILE
			SubFile.FilePath = File.SubFileDir & File.FileName
			SubFile.FileName = File.FileName
			SubFile.FileSize = .lPointerToRawData + .lSizeOfRawData
			Bytes = GetBytes(FN,SubFile.FileSize,0,Mode)
			FN2 = FreeFile
 			Open SubFile.FilePath For Binary Access Write Lock Write As #FN2
			Put #FN2,1,Bytes
			Close #FN2
		Else
			ReDim DataList(0) As SUB_FILE
		End If
	End With
	If ShowMsg > 0 Then
		SetTextBoxString ShowMsg,Msg & "100%"
	ElseIf ShowMsg < 0 Then
		PSL.OutputWnd(0).Clear
		PSL.Output Msg & "100%"
	End If
	ExitFunction:
	'关闭文件
	Set Dic = Nothing
	On Error Resume Next
	UnLoadFile(FN,0,Mode)
	If DataList(0).FilePath = "" Then Exit Function
End Function


'导入子文件列表
Private Function WriteDataToFile(ByVal trgFolder As String,File As FILE_PROPERTIE,DataList() As SUB_FILE) As Boolean
	Dim i As Long,sb As Object
	'打开文件
	On Error GoTo ErrHandle
	If Right$(trgFolder ,1) <> "\" Then trgFolder = trgFolder & "\"
	trgFolder = trgFolder & File.FileName & "_" & DataFileName
	If Dir$(trgFolder) <> "" Then Kill trgFolder
	ReDim List(10) As String
	List(0) = "; PESubFile Data File"
	List(1) = "; " & String$(70,"-")
	List(2) = "; Application Version: " & Version
	List(3) = "; Main File Name: " & File.OrigionalFileName
	List(4) = "; Main File Version: " & File.FileVersion
	List(5) = "; Main File Size: " & CStr$(File.FileSize)
	List(6) = "; Main File DateTime: " & Format$(File.DateLastModified,"yyyy-M-d H:mm:ss")
	List(7) = "; Main File LanguageID: " & File.LanguageID
	List(8) = "; " & String$(70,"-")
	On Error Resume Next
	Set sb = CreateObject("System.Text.StringBuilder")
	On Error GoTo 0
	If sb Is Nothing Then
		ReDim List(UBound(DataList) + 10) As String
		For i = 0 To UBound(DataList)
			With DataList(i)
				List(i + 10) = CStr$(i) & "," & CStr$(.FileAdd) & "," & .FileName
			End With
		Next i
	Else
		For i = 0 To UBound(DataList)
			With DataList(i)
				sb.AppendFormat("{0}",CStr$(i) & "," & CStr$(.FileAdd) & "," & .FileName & vbNullChar)
			End With
		Next i
		List(10) = sb.ToString()
	End If
	WriteBinaryFile trgFolder,CP_UTF8,StrListJoin(List,vbNullChar,True),True
	Set sb = Nothing
	WriteDataToFile = True
	Exit Function
	ErrHandle:
End Function


'导入子文件列表
'ImportSubFile = -1 程序错误
'ImportSubFile = -2 提取旧格式
'ImportSubFile = 1 没有子文件
'ImportSubFile = 2 指定文件夹中不存在子文件
'ImportSubFile = 3 没有子文件列表数据文件
'ImportSubFile = 4 已提取子文件的原始文件名和现在的不匹配
'ImportSubFile = 5 已提取子文件的文件版本和现在的不匹配
'ImportSubFile = 6 已提取子文件的语言ID和现在的不匹配
'ImportSubFile = 7 已提取子文件的文件大小和现在的不匹配
'ImportSubFile = 8 子文件列表数据文件格式不匹配
'ImportSubFile = 9 一部分子文件不存在
Private Function ImportSubFile(ByVal trgFolder As String,File As FILE_PROPERTIE,DataList() As SUB_FILE,Optional ByVal ShowMsg As Long) As Long
	Dim i As Long,n As Long,Msg As String,VerCompVal As Integer
	On Error GoTo ExitFunction
	If Right$(trgFolder,1) <> "\" Then trgFolder = trgFolder & "\"
	If File.NumberOfSub = 0 Then
		ImportSubFile = 1
		Exit Function
	ElseIf Dir$(trgFolder & "*.*") = "" Then
		ImportSubFile = 2
		Exit Function
	Else
		For i = 0 To 10
			If Dir$(trgFolder & File.FileName & "_" & i) = "" Then Exit For
			n = n + 1
		Next i
		If n > 0 Then
			ImportSubFile = -2
			Exit Function
		ElseIf Dir$(trgFolder & File.FileName & "_" & DataFileName) = "" Then
			ImportSubFile = 3
			Exit Function
		End If
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
	ReDim List(0) As String,TempList(0) As String
	List = ReSplit(ReadBinaryFile(trgFolder & File.FileName & "_" & DataFileName,CP_UTF8,True),vbNullChar)
	'检查数据文件
	For i = 0 To UBound(List)
		List(i) = Trim$(List(i))
		If List(i) <> "" Then
			If Left$(List(i),1) = ";" Then
				If InStr(List(i),"Main File Name") Then
					List(i) = Trim$(Mid$(List(i),InStr(List(i),":") + 1))
					If List(i) <> File.OrigionalFileName Then
						ImportSubFile = 4
						Exit Function
					End If
				'If InStr(List(i),"Main File Path") Then
				'	List(i) = Trim$(Mid$(List(i),InStr(List(i),":") + 1))
				'	If List(i) <> File.FilePath Then
				'		ImportSubFile = 3
				'		Exit Function
				'	End If
				ElseIf InStr(List(i),"Main File Version") Then
					List(i) = Trim$(Mid$(List(i),InStr(List(i),":") + 1))
					If List(i) <> File.FileVersion Then
						ImportSubFile = 5
						Exit Function
					End If
				ElseIf InStr(List(i),"Main File LanguageID") Then
					List(i) = Trim$(Mid$(List(i),InStr(List(i),":") + 1))
					If List(i) <> File.LanguageID Then
						ImportSubFile = 6
						Exit Function
					End If
				ElseIf InStr(List(i),"Main File Size") Then
					List(i) = Trim$(Mid$(List(i),InStr(List(i),":") + 1))
					If StrToLong(List(i)) <> File.FileSize Then
						ImportSubFile = 7
						Exit Function
					End If
				ElseIf InStr(List(i),"Main File DateTime") Then
					List(i) = Trim$(Mid$(List(i),InStr(List(i),":") + 1))
					If IsDate(List(i)) Then
						If CStr(CDate(List(i))) <> CStr(File.DateLastModified) Then
							If MsgBox(MsgList(121),vbYesNo+vbInformation,MsgList(42)) = vbNo Then GoTo ExitFunction
						End If
					Else
						If CStr(StrToDate(List(i))) <> CStr(File.DateLastModified) Then
							If MsgBox(MsgList(121),vbYesNo+vbInformation,MsgList(42)) = vbNo Then GoTo ExitFunction
						End If
					End If
				End If
			End If
		End If
	Next i
	'解析数据文件
	File.SubFileDir = trgFolder
	ReDim DataList(File.NumberOfSub - 1) As SUB_FILE
	For i = 10 To UBound(List)
		List(i) = Trim$(List(i))
		If List(i) <> "" Then
			If Left$(List(i),1) <> ";" Then
				TempList = ReSplit(List(i),",",3)
				If StrToLong(TempList(0)) <> i - 10 Then
					ImportSubFile = 8
					GoTo ExitFunction
				End If
				DataList(n).FileName = TempList(2)
				DataList(n).FilePath = trgFolder & DataList(n).FileName
				If Dir$(DataList(n).FilePath) = "" Then
					ImportSubFile = 9
					GoTo ExitFunction
				End If
				DataList(n).FileAdd = StrToLong(TempList(1))
				DataList(n).NewFileSize = FileLen(DataList(n).FilePath)
				If n > 0 Then
					DataList(n - 1).FileSize = DataList(n).FileAdd - DataList(n - 1).FileAdd
				End If
				n = n + 1
			End If
		End If
		If ShowMsg > 0 Then
			SetTextBoxString ShowMsg,Msg & Format$(i / File.NumberOfSub,"#%")
		ElseIf ShowMsg < 0 Then
			PSL.OutputWnd(0).Clear
			PSL.Output Msg & Format$(i / File.NumberOfSub,"#%")
		End If
	Next i
	If ShowMsg > 0 Then
		SetTextBoxString ShowMsg,Msg & "100%"
	ElseIf ShowMsg < 0 Then
		PSL.OutputWnd(0).Clear
		PSL.Output Msg & "100%"
	End If
	If n > 0 Then
		n = n - 1
		DataList(n).FileSize = File.FileSize - DataList(n).FileAdd
		ReDim Preserve DataList(n) As SUB_FILE
	Else
		ReDim DataList(0) As SUB_FILE
	End If
	Exit Function
	ExitFunction:
	ReDim DataList(0) As SUB_FILE
	If ImportSubFile = 0 Then ImportSubFile = -1
End Function


'转换包含非数字日期时间的字符串为日期
Private Function StrToDate(ByVal strDate As String) As Date
	Dim i As Long,Matches As Object
	With RegExp
		.Global = True
		.IgnoreCase = True
		.Pattern = "[^\x20-\x7E]+"
		Set Matches = .Execute(strDate)
		If Matches.Count > 0 Then
			For i = 0 To Matches.Count - 1
				strDate = Replace$(strDate,Matches(i).Value," ")
			Next i
		End If
	End With
	If IsDate(strDate) Then StrToDate = CDate(strDate)
End Function


'合并文件
'fType = False 直接合并，否则文件对齐合并
Private Function MergeFile(trgFile As String,srcFile As FILE_PROPERTIE,DataList() As SUB_FILE,ByVal Mode As Long,ByVal fType As Boolean) As Boolean
	Dim i As Long,n As Long,k As Long,FN As FILE_IMAGE,FN2 As Variant,Bytes() As Byte
	If srcFile.NumberOfSub = 0 Then Exit Function
	On Error GoTo ExitFunction
	With srcFile.SecList(srcFile.MaxSecIndex)
		'打开文件
		i = .lPointerToRawData + .lSizeOfRawData
		Mode = LoadFile(trgFile,FN,i,1,i,Mode)
		If Mode < -1 Then Exit Function
		FN.SizeOfFile = i
		'合并子文件
		If fType = False Then
			For i = 0 To UBound(DataList)
				If Dir$(DataList(i).FilePath) <> "" Then
					k = FileLen(DataList(i).FilePath)
					If k > 0 Then
						ReDim Bytes(k - 1) As Byte
 						FN2 = FreeFile
 						Open DataList(i).FilePath For Binary Access Read Lock Write As #FN2
 						Get #FN2, , Bytes
 						Close #FN2
						PutBytes(FN,GetFileLength(FN,Mode),Bytes,k,Mode)
						n = n + 1
					End If
				End If
			Next i
		Else
			For i = 0 To UBound(DataList)
				If Dir$(DataList(i).FilePath) <> "" Then
					k = FileLen(DataList(i).FilePath)
					ReDim Bytes(k - 1) As Byte
 					FN2 = FreeFile
 					Open DataList(i).FilePath For Binary Access Read Lock Write As #FN2
 					Get #FN2, , Bytes
 					Close #FN2
 					If n = 0 Then
 						With srcFile.SecList(srcFile.MaxSecID)
							n = Alignment(.lSizeOfRawData,srcFile.FileAlign,1) - .lSizeOfRawData
 						End With
 						k = FileLen(DataList(i).FilePath)
						PutBytes(FN,GetFileLength(FN,Mode) + n,Bytes,k,Mode)
						n = 1
 					Else
 						PutBytes(FN,GetFileLength(FN,Mode),Bytes,k,Mode)
					End If
					n = n + 1
				End If
			Next i
		End If
	End With
	If n > 0 Then MergeFile = True
	UnLoadFile(FN,FN.SizeOfFile,Mode)
	Exit Function
	ExitFunction:
	'关闭文件
	On Error Resume Next
	UnLoadFile(FN,0,Mode)
End Function


'过滤字串
'Mode = 0 常规，= 1 通配符, = 2 正则表达式
'FilterStr = 1 已找到，= 0 未找到, = -1 程序错误, = -2 通配符语法错误 = -3 正则表达式语法错误
Private Function FilterStr(ByVal txtStr As String,ByVal FindStr As String,ByVal Mode As Long,Optional ByVal IgnoreCase As Boolean) As Long
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


'查找子文件名称符合查找内容的索引号
'FindString > -1 找到的字串列表索引号，= -1 未找到
Private Function FindString(FileList() As String,ByVal FindStr As String,ByVal Index As Long) As Long
	Dim i As Long,Max As Long
	FindString = -1
	Max = UBound(FileList)
	If Index < Max Then Index = Index + 1 Else Index = 0
	FindStr = StrToRegExpPattern(FindStr)
	For i = Index To Max
		If CheckStrRegExp(FileList(i),FindStr,0,2,True) = True Then
			FindString = i
			Exit Function
		End If
	Next i
End Function


'获取子文件列表
'Mode = 0 获取全部子文件地址
'Mode = 1 获取全部子文件名称

'Mode = 2 获取子文件小于原始文件的子文件地址
'Mode = 3 获取子文件小于原始文件的子文件名称

'Mode = 4 获取子文件等于原始文件的子文件地址
'Mode = 5 获取子文件等于原始文件的子文件名称

'Mode = 6 获取子文件大于原始文件的子文件地址
'Mode = 7 获取子文件大于原始文件的子文件名称

'Mode = 8 获取子文件名称符合查找内容的子文件地址
'Mode = 9 获取子文件名称符合查找内容的子文件名称
Private Function GetDataValueList(DataList() As SUB_FILE,ByVal MaxVal As Long,ByVal FindStr As String,ByVal Mode As Long) As String()
	Dim i As Long,n As Long
	ReDim TempList(UBound(DataList)) As String
	Select Case Mode
	Case 0 To 1
		For i = 0 To UBound(DataList)
			Select Case Mode
			Case 0
				TempList(i) = ValToStr(DataList(i).FileAdd,MaxVal,True)
			Case 1
				TempList(i) = DataList(i).FileName
			End Select
		Next i
	Case 2 To 3
		For i = 0 To UBound(DataList)
			If DataList(i).NewFileSize < DataList(i).FileSize Then
				Select Case Mode
				Case 2
					TempList(n) = ValToStr(DataList(i).FileAdd,MaxVal,True)
				Case 3
					TempList(n) = DataList(i).FileName
				End Select
				n = n + 1
			End If
		Next i
		If n > 0 Then n = n - 1
		ReDim Preserve TempList(n) As String
	Case 4 To 5
		For i = 0 To UBound(DataList)
			If DataList(i).NewFileSize = DataList(i).FileSize Then
				Select Case Mode
				Case 4
					TempList(n) = ValToStr(DataList(i).FileAdd,MaxVal,True)
				Case 5
					TempList(n) = DataList(i).FileName
				End Select
				n = n + 1
			End If
		Next i
		If n > 0 Then n = n - 1
		ReDim Preserve TempList(n) As String
	Case 6 To 7
		For i = 0 To UBound(DataList)
			If DataList(i).NewFileSize > DataList(i).FileSize Then
				Select Case Mode
				Case 6
					TempList(n) = ValToStr(DataList(i).FileAdd,MaxVal,True)
				Case 7
					TempList(n) = DataList(i).FileName
				End Select
				n = n + 1
			End If
		Next i
		If n > 0 Then n = n - 1
		ReDim Preserve TempList(n) As String
	Case 8 To 9
		FindStr = StrToRegExpPattern(FindStr)
		For i = 0 To UBound(DataList)
			If CheckStrRegExp(DataList(i).FileName,FindStr,0,2,True) = True Then
				Select Case Mode
				Case 8
					TempList(n) = ValToStr(DataList(i).FileAdd,MaxVal,True)
				Case 9
					TempList(n) = DataList(i).FileName
				End Select
				n = n + 1
			End If
		Next i
		If n > 0 Then n = n - 1
		ReDim Preserve TempList(n) As String
	End Select
	GetDataValueList = TempList
End Function


'获取查找字串的查找方式
'GetFindMode = 0 常规，= 1 通配符, = 2 正则表达式
Private Function GetFindMode(FindStr As String) As Long
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


'转换查找内容为正则表达式模板
Private Function StrToRegExpPattern(ByVal strText As String) As String
	Dim i As Long,TempList() As String
	StrToRegExpPattern = strText
	Select Case GetFindMode(StrToRegExpPattern)
	Case 0
		If (StrToRegExpPattern Like "*\[*?#[]*") = True Then
			TempList = ReSplit("*,?,#,[",",",-1)
			For i = 0 To UBound(TempList)
				StrToRegExpPattern = Replace$(StrToRegExpPattern,"\" & TempList(i),TempList(i))
			Next i
		End If
	Case 1
		'转换通配符为正则表达式模板
		TempList = ReSplit("\*,\#,\?,\[",",",-1)
		For i = 0 To UBound(TempList)
			StrToRegExpPattern = Replace$(StrToRegExpPattern,TempList(i),CStr$(i) & vbNullChar & CStr$(i) & vbNullChar & CStr$(i))
		Next i
		StrToRegExpPattern = Replace$(StrToRegExpPattern,"?",".")
		StrToRegExpPattern = Replace$(StrToRegExpPattern,"*",".*")
		StrToRegExpPattern = Replace$(StrToRegExpPattern,"#","\d")
		StrToRegExpPattern = Replace$(StrToRegExpPattern,"[!","[^")
		For i = 0 To UBound(TempList)
			StrToRegExpPattern = Replace(StrToRegExpPattern,CStr$(i) & vbNullChar & CStr$(i) & vbNullChar & CStr$(i),TempList(i))
		Next i
		StrToRegExpPattern = Replace$(StrToRegExpPattern,"\#","#")
	End Select
End Function


'检查正则表达式是否正确
Private Function CheckRegExp(ByVal RegEx As Object,ByVal Patrn As String) As Boolean
	If Patrn = "" Then Exit Function
	On Error GoTo ExitFunction
	With RegEx
		.Pattern = Patrn
		.Test("CheckRegExp")
	End With
	CheckRegExp = True
	ExitFunction:
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
Private Function Byte2Hex(Bytes As Variant,ByVal StartPos As Long,ByVal endPos As Long) As String
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


'转换十进制和十六进制值为字符
'MaxVal = 0 按值计算应有的长度，> 0 按文件大小计算的位数，< 0 按指定位数
Private Function ValToStr(ByVal DecVal As Long,Optional ByVal MaxVal As Long,Optional ByVal DisPlayFormat As Boolean) As String
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


'创建子文件夹
Private Function MkSubDir(ByVal DirPath As String) As Boolean
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


'删除文件夹，不会删除子文件夹
Private Function DelDir(ByVal DirPath As String) As Boolean
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
Private Function DelDirs(ByVal DirPath As String) As Boolean
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


'浏览文件夹
Private Function BrowseForFolder(Folder As String,sTitle As String) As Boolean
	Dim lpIDList As Long, lResult As Long, udtBI As BrowseInfo
	With udtBI
		.hWndOwner = 0&
		.lpszTitle = sTitle
		.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_USENEWUI
	End With
	lpIDList = SHBrowseForFolder(udtBI)
	If lpIDList Then
		Folder = String$(260, 0)
		SHGetPathFromIDList lpIDList, Folder
		CoTaskMemFree lpIDList
		Folder = Replace$(Folder,vbNullChar,"")
		BrowseForFolder = True
	End If
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
		List(0) = MsgList(46)
		List(1) = Replace$(MsgList(47),"%s",.FileName)
		List(2) = Replace$(MsgList(48),"%s",.FilePath)
		List(3) = Replace$(MsgList(49),"%s",.FileDescription)
		List(4) = Replace$(MsgList(50),"%s",.FileVersion)
		List(5) = Replace$(MsgList(51),"%s",.ProductName)
		List(6) = Replace$(MsgList(52),"%s",.ProductVersion)
		List(7) = Replace$(MsgList(53),"%s",.LegalCopyright)
		List(8) = Replace$(MsgList(54),"%s",CStr$(.FileSize))
		List(9) = Replace$(MsgList(55),"%s",CStr$(.DateCreated))
		List(10) = Replace$(MsgList(56),"%s",CStr$(.DateLastModified))
		List(11) = Replace$(MsgList(57),"%s",PSL.GetLangCode(Val("&H" & .LanguageID),pslCodeText))
		List(12) = Replace$(MsgList(58),"%s",.CompanyName)
		List(13) = Replace$(MsgList(59),"%s",.OrigionalFileName)
		List(14) = Replace$(MsgList(60),"%s",.InternalName)
		Select Case .Magic
		Case "PE32","NET32","MAC32"
			If .LangType = DELPHI_FILE_SIGNATURE Then
				List(15) = Replace$(MsgList(61),"%s","Delphi32")
				List(16) = Replace$(MsgList(62),"%s","0x" & ValToStr(.ImageBase,-8,True))
			ElseIf .LangType = NET_FILE_SIGNATURE Then
				List(15) = Replace$(MsgList(61),"%s",".NET32")
				List(16) = Replace$(MsgList(62),"%s","0x" & ValToStr(.ImageBase,-8,True))
			ElseIf Stemp = True Then
				List(15) = Replace$(MsgList(61),"%s","MAC32")
			Else
				List(15) = Replace$(MsgList(61),"%s","PE32")
				List(16) = Replace$(MsgList(62),"%s","0x" & ValToStr(.ImageBase,-8,True))
			End If
		Case "PE64","NET64","MAC64"
			If .LangType = DELPHI_FILE_SIGNATURE Then
				List(15) = Replace$(MsgList(61),"%s","Delphi64")
				List(16) = Replace$(MsgList(62),"%s","0x" & ReverseHexCode(Byte2Hex(.ImageBase,0,-1),16))
			ElseIf .LangType = NET_FILE_SIGNATURE Then
				List(15) = Replace$(MsgList(61),"%s",".NET64")
				List(16) = Replace$(MsgList(62),"%s","0x" & ReverseHexCode(Byte2Hex(.ImageBase,0,-1),16))
			ElseIf Stemp = True Then
				List(15) = Replace$(MsgList(61),"%s","MAC64")
			Else
				List(15) = Replace$(MsgList(61),"%s","PE64")
				List(16) = Replace$(MsgList(62),"%s","0x" & ReverseHexCode(Byte2Hex(.ImageBase,0,-1),16))
			End If
		Case Else
			List(15) = Replace$(MsgList(61),"%s",MsgList(104))
		End Select
	End With
	If List(16) = "" Then n = n - 1
	'每个文件节的偏移地址
	ReDim Preserve List(n + 6 + File.MaxSecIndex) As String
	List(n + 2) = MsgList(63)
	List(n + 3) = MsgList(66) & MsgList(66)
	List(n + 4) = IIf(Stemp = False,MsgList(64),MsgList(106))
	List(n + 5) = MsgList(66) & MsgList(66)
	n = n + 6
	For i = 0 To File.MaxSecIndex - 1
		With File.SecList(i)
			List(n) = Replace$(IIf(Stemp = False,MsgList(65),MsgList(107)),"%s!1!",MsgList(67))
			List(n) = Replace$(List(n),"%s!2!",IIf(File.Magic = "",MsgList(104),.sName))
			List(n) = Replace$(List(n),"%s!3!","")
			List(n) = Replace$(List(n),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
			List(n) = Replace$(List(n),"%s!5!",ValToStr(.lPointerToRawData + _
							IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
			List(n) = Replace$(List(n),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
			If .SubSecs > 0 Then
				n = n + 1
				ReDim Preserve List(n + .SubSecs + File.MaxSecIndex - i) As String
				For j = 0 To .SubSecs - 1
					List(n) = Replace$(MsgList(107),"%s!1!",MsgList(67))
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
			List(n) = Replace$(IIf(Stemp = False,MsgList(65),MsgList(107)),"%s!1!",MsgList(67))
			List(n) = Replace$(List(n),"%s!2!",MsgList(70))
			List(n) = Replace$(List(n),"%s!3!","")
			List(n) = Replace$(List(n),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
			List(n) = Replace$(List(n),"%s!5!",ValToStr(.lPointerToRawData + _
						IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
			List(n) = Replace$(List(n),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
			n = n + 1
		End If
		If File.NumberOfSub > 0 Then
			ReDim Preserve List(n) As String
			List(n) = Replace$(IIf(Stemp = False,MsgList(65),MsgList(107)),"%s!1!",MsgList(67))
			List(n) = Replace$(List(n),"%s!2!",Replace$(MsgList(105),"%s",CStr$(File.NumberOfSub)))
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
			List(n) = Replace$(IIf(Stemp = False,MsgList(65),MsgList(107)),"%s!1!",MsgList(68))
			List(n) = Replace$(List(n),"%s!2!",IIf(File.Magic = "",MsgList(104),.sName))
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
					List(n) = Replace$(MsgList(107),"%s!1!",MsgList(68))
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
			List(n) = Replace$(IIf(Stemp = False,MsgList(65),MsgList(107)),"%s!1!",MsgList(68))
			List(n) = Replace$(List(n),"%s!2!",MsgList(70))
			List(n) = Replace$(List(n),"%s!3!","")
			List(n) = Replace$(List(n),"%s!4!",MsgList(71))
			List(n) = Replace$(List(n),"%s!5!",MsgList(71))
			List(n) = Replace$(List(n),"%s!6!",MsgList(71))
			n = n + 1
		End If
		If File.NumberOfSub > 0 Then
			ReDim Preserve List(n) As String
			List(n) = Replace$(IIf(Stemp = False,MsgList(65),MsgList(107)),"%s!1!",MsgList(68))
			List(n) = Replace$(List(n),"%s!2!",Replace$(MsgList(105),"%s",CStr$(File.NumberOfSub)))
			List(n) = Replace$(List(n),"%s!3!","")
			List(n) = Replace$(List(n),"%s!4!",MsgList(71))
			List(n) = Replace$(List(n),"%s!5!",MsgList(71))
			List(n) = Replace$(List(n),"%s!6!",MsgList(71))
			n = n + 1
		End If
	End With
	ReDim Preserve List(n) As String
	List(n) = MsgList(66) & MsgList(66)
	'数据目录地址及所在文件节
	If File.DataDirs > 0 Then
		ReDim Preserve List(n + 6 + File.DataDirs) As String
		List(n + 2) = MsgList(73)
		List(n + 3) = MsgList(66) & MsgList(66)
		List(n + 4) = MsgList(74)
		List(n + 5) = MsgList(66) & MsgList(66)
		n = n + 6
		For i = 0 To File.DataDirs - 1
			With File.DataDirectory(i)
				List(n) = Replace$(MsgList(65),"%s!1!",MsgList(i + 75))
				If .lPointerToRawData > 0 Then
					j = SkipSection(File,.lPointerToRawData,0,0,1)
					If j > -1 Then
						List(n) = Replace$(List(n),"%s!2!",File.SecList(j).sName)
					Else
						List(n) = Replace$(List(n),"%s!2!",MsgList(91))
					End If
				ElseIf .lSizeOfRawData > 0 Then
					List(n) = Replace$(List(n),"%s!2!",MsgList(91))
				Else
					List(n) = Replace$(List(n),"%s!2!",MsgList(92))
				End If
				List(n) = Replace$(List(n),"%s!3!","")
				List(n) = Replace$(List(n),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!5!",ValToStr(.lPointerToRawData + _
								IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
				n = n + 1
			End With
		Next i
		List(n) = MsgList(66) & MsgList(66)
	End If
	'.NET CLR 数据目录地址及所在文件节
	If File.LangType = NET_FILE_SIGNATURE Then
		ReDim Preserve List(n + 6 + 7) As String
		List(n + 2) = MsgList(93)
		List(n + 3) = MsgList(66) & MsgList(66)
		List(n + 4) = MsgList(94)
		List(n + 5) = MsgList(66) & MsgList(66)
		n = n + 6
		For i = 0 To 6
			With File.CLRList(i)
				List(n) = Replace$(MsgList(65),"%s!1!",MsgList(i + 95))
				If .lPointerToRawData > 0 Then
					j = SkipSection(File,.lPointerToRawData,0,0,1)
					If j > -1 Then
						List(n) = Replace$(List(n),"%s!2!",File.SecList(j).sName)
					Else
						List(n) = Replace$(List(n),"%s!2!",MsgList(91))
					End If
				ElseIf .lSizeOfRawData > 0 Then
					List(n) = Replace$(List(n),"%s!2!",MsgList(91))
				Else
					List(n) = Replace$(List(n),"%s!2!",MsgList(92))
				End If
				List(n) = Replace$(List(n),"%s!3!","")
				List(n) = Replace$(List(n),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!5!",ValToStr(.lPointerToRawData + _
							IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
				n = n + 1
			End With
		Next i
		List(n) = MsgList(66) & MsgList(66)
	End If
	'.NET 流地址及所在文件节
	If File.NetStreams > 0 Then
		ReDim Preserve List(n + 6 + File.NetStreams) As String
		List(n + 2) = MsgList(102)
		List(n + 3) = MsgList(66) & MsgList(66)
		List(n + 4) = MsgList(103)
		List(n + 5) = MsgList(66) & MsgList(66)
		n = n + 6
		For i = 0 To File.NetStreams - 1
			With File.StreamList(i)
				List(n) = Replace$(MsgList(65),"%s!1!",.sName)
				If .lPointerToRawData > 0 Then
					j = SkipSection(File,.lPointerToRawData,0,0,1)
					If j > -1 Then
						List(n) = Replace$(List(n),"%s!2!",File.SecList(j).sName)
					Else
						List(n) = Replace$(List(n),"%s!2!",MsgList(91))
					End If
				ElseIf .lSizeOfRawData > 0 Then
					List(n) = Replace$(List(n),"%s!2!",MsgList(91))
				Else
					List(n) = Replace$(List(n),"%s!2!",MsgList(92))
				End If
				List(n) = Replace$(List(n),"%s!3!","")
				List(n) = Replace$(List(n),"%s!4!",ValToStr(.lPointerToRawData,File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!5!",ValToStr(.lPointerToRawData + _
							IIf(.lSizeOfRawData = 0,0,.lSizeOfRawData - 1),File.FileSize,DisPlayFormat))
				List(n) = Replace$(List(n),"%s!6!",ValToStr(.lSizeOfRawData,File.FileSize,DisPlayFormat))
				n = n + 1
			End With
		Next i
		List(n) = MsgList(66) & MsgList(66)
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


'显示文件信息
Private Sub ShowFileInfo(FileList() As String,StrList() As String)
	If CheckArrEmpty(FileList) = False Then Exit Sub
	If CheckArrEmpty(StrList) = False Then Exit Sub
	Begin Dialog UserDialog 890,448,Replace$(MsgList(109),"%s",FileList(0)),.ShowFileInfoDlgFunc ' %GRID:10,7,1,1
		TextBox 280,420,410,21,.IndexBox
		ListBox 280,420,420,21,FileList(),.FileList
		TextBox 0,7,890,406,.InTextBox,1
		PushButton 40,420,100,21,MsgList(110),.PreviousButton
		PushButton 160,420,100,21,MsgList(111),.NextButton
		OKButton 750,420,100,21,.OKButton
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub


'显示文件信息对话框函数
Private Function ShowFileInfoDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 1
		DlgVisible "IndexBox",False
		DlgVisible "FileList",False
		DlgText "InTextBox",StrList(0)
    	If UBound(StrList) = 0 Then
			DlgEnable "PreviousButton",False
			DlgEnable "NextButton",False
		Else
			DlgEnable "PreviousButton",False
			DlgEnable "NextButton",True
    	End If
	Case 2 ' 数值更改或者按下了按钮
		ShowFileInfoDlgFunc = True '防止按下按钮关闭对话框窗口
		Select Case DlgItem$
		Case "OKButton"
			ShowFileInfoDlgFunc = False
			Exit Function
		Case "PreviousButton"
			DlgText "IndexBox",CStr$(StrToLong(DlgText("IndexBox")) - 1)
			DlgValue "FileList",StrToLong(DlgText("IndexBox"))
			DlgText -1,Replace$(MsgList(109),"%s",DlgText("FileList"))
			DlgText "InTextBox",StrList(StrToLong(DlgText("IndexBox")))
			If StrToLong(DlgText("IndexBox")) < 1 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",True
			ElseIf StrToLong(DlgText("IndexBox")) = UBound(StrList) Then
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",False
			Else
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",True
			End If
		Case "NextButton"
			DlgText "IndexBox",CStr$(StrToLong(DlgText("IndexBox")) + 1)
			DlgValue "FileList",StrToLong(DlgText("IndexBox"))
			DlgText -1,Replace$(MsgList(109),"%s",DlgText("FileList"))
			DlgText "InTextBox",StrList(StrToLong(DlgText("IndexBox")))
			If StrToLong(DlgText("IndexBox")) < 1 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",True
			ElseIf StrToLong(DlgText("IndexBox")) = UBound(StrList) Then
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",False
			Else
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",True
			End If
		End Select
	End Select
End Function


'写入二进制文件
'BOM = False 检查并写入 BOM，否则不写入 BOM
'Mode = False 删除文件，重新写入，仅在 File 为文件名时适用
Private Function WriteBinaryFile(ByVal File As Variant,ByVal CodePage As Long,ByVal textStr As String, _
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


'读取二进制文件
'BOM = False 检查并去掉 BOM，否则读入 BOM
Private Function ReadBinaryFile(ByVal FilePath As String,ByVal CodePage As Long,Optional ByVal BOM As Boolean) As String
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
	ReDim MsgList(137) As String
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

		MsgList(12) = "PE 子文件管理器 - 版本 %v (构建 %b)"
		MsgList(13) = "..."
		MsgList(14) = "地址(%s/%d)"
		MsgList(15) = "子文件列表"
		MsgList(16) = "状态: %s"
		MsgList(17) = "关于"
		MsgList(18) = "语言"
		MsgList(19) = "提取"
		MsgList(20) = "导入"
		MsgList(21) = "复制文件名"
		MsgList(22) = "主文件信息"
		MsgList(23) = "子文件信息"
		MsgList(24) = "直接合并"
		MsgList(25) = "文件对齐合并"

		MsgList(26) = "版本 %v (构建 %b)\r\n" & _
					"OS 版本: Windows XP/2000 或以上\r\n" & _
					"Passolo 版本: Passolo 5.0 或以上\r\n" & _
					"授权: 免费软件\r\n" & _
					"网址: http://www.hanzify.org\r\n" & _
					"作者: wanfu (2018 - 2019)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(27) = "关于 PE 子文件管理器"
		MsgList(28) = "可执行文件 (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|所有文件 (*.*)|*.*||"
		MsgList(29) = "选择文件"
		MsgList(30) = "英语;简体中文;繁体中文"
		MsgList(31) = "enu;chs;cht"
		MsgList(32) = "正在提取文件..."
		MsgList(33) = "共提取 %s 个子文件。"
		MsgList(34) = "共提取 %s 个子文件，已保存到 %d 文件夹。"
		MsgList(35) = "该文件为非 PE 文件。"
		MsgList(36) = "不能将原文件作为目标文件。"
		MsgList(37) = "正在合并文件..."
		MsgList(38) = "合并成功！"
		MsgList(39) = "合并失败！缺少子文件。"
		MsgList(40) = "选择释放子文件的文件夹"
		MsgList(41) = "选择子文件所在的文件夹"
		MsgList(42) = "确认"
		MsgList(43) = "比原始小;和原始相同;比原始大;子文件名"
		MsgList(44) = "正在导入文件..."
		MsgList(45) = "共导入 %s 个子文件。"

		MsgList(46) = "============ 文件信息 ============\r\n"
		MsgList(47) = "文件名称：\t%s"
		MsgList(48) = "文件路径：\t%s"
		MsgList(49) = "文件说明：\t%s"
		MsgList(50) = "文件版本：\t%s"
		MsgList(51) = "产品名称：\t%s"
		MsgList(52) = "产品版本：\t%s"
		MsgList(53) = "版权所有：\t%s"
		MsgList(54) = "文件大小：\t%s 字节"
		MsgList(55) = "创建日期：\t%s"
		MsgList(56) = "修改日期：\t%s"
		MsgList(57) = "语　　言：\t%s"
		MsgList(58) = "开 发 商：\t%s"
		MsgList(59) = "原始文件名：\t%s"
		MsgList(60) = "内部文件名：\t%s"
		MsgList(61) = "文件类型：\t%s"
		MsgList(62) = "映像基址：\t%s"
		MsgList(63) = "区段信息："
		MsgList(64) = "地址类别\t区段名\t开始地址\t结束地址\t字节大小"
		MsgList(65) = "%s!1!\t%s!2!\t%s!4!\t%s!5!\t%s!6!"
		MsgList(66) = "================================="
		MsgList(67) = "文件偏移地址"
		MsgList(68) = "相对虚拟地址"
		MsgList(69) = "任意"
		MsgList(70) = "隐藏"
		MsgList(71) = "未知"
		MsgList(72) = "不可用"
		MsgList(73) = "数据目录信息 (文件偏移地址)："
		MsgList(74) = "目录名称\t所在区段\t开始地址\t结束地址\t字节大小"
		MsgList(75) = "导出目录"
		MsgList(76) = "导入目录"
		MsgList(77) = "资源目录"
		MsgList(78) = "异常目录"
		MsgList(79) = "安全目录"
		MsgList(80) = "基址重定位表"
		MsgList(81) = "调试目录"
		MsgList(82) = "版权目录"
		MsgList(83) = "机器值(GP RVA)"
		MsgList(84) = "线程本地存储表"
		MsgList(85) = "载入配置目录"
		MsgList(86) = "绑定导入目录"
		MsgList(87) = "导入地址表"
		MsgList(88) = "延迟加载导入符"
		MsgList(89) = "COM 运行库标志"
		MsgList(90) = "保留目录"
		MsgList(91) = "异常"
		MsgList(92) = "不存在"
		MsgList(93) = ".NET CLR 数据目录信息 (文件偏移地址)："
		MsgList(94) = "目录名称\t所在区段\t开始地址\t结束地址\t字节大小"
		MsgList(95) = "元数据(MetaData)"
		MsgList(96) = "托管资源"
		MsgList(97) = "强名称签名"
		MsgList(98) = "代码管理表"
		MsgList(99) = "虚拟表(V-表)"
		MsgList(100) = "跳转导出地址表"
		MsgList(101) = "托管本机映像头"
		MsgList(102) = ".NET MetaData 流信息 (文件偏移地址)："
		MsgList(103) = "流名称\t所在区段\t开始地址\t结束地址\t字节大小"
		MsgList(104) = "非 PE 文件"
		MsgList(105) = "子PE(%s)"
		MsgList(106) = "地址类别\t段名\t节名\t\t开始地址\t结束地址\t字节大小"
		MsgList(107) = "%s!1!\t%s!2!\t%s!3!\t\t%s!4!\t%s!5!\t%s!6!"

		MsgList(108) = "正在获取 %s 文件信息..."
		MsgList(109) = "信息 - %s"
		MsgList(110) = "上一个"
		MsgList(111) = "下一个"
		MsgList(112) = "找到 %s 个子文件。"

		MsgList(113) = "信息"
		MsgList(114) = "文件没有子文件。"
		MsgList(115) = "要导入的文件夹中没有文件。"
		MsgList(116) = "要导入的文件夹中没有找到 %s 数据文件，无法导入。"
		MsgList(117) = "子文件的原始文件名称不符，无法导入。"
		MsgList(118) = "子文件的原始文件版本不符，无法导入。"
		MsgList(119) = "子文件的原始文件语言不符，无法导入。"
		MsgList(120) = "子文件的原始文件大小不符，无法导入。"
		MsgList(121) = "子文件的原始文件日期已更改，是否继续？\r\n日期已更改，说明文件已被修改过，修改过的文件可能不适用。"
		MsgList(122) = "%s 数据文件格式不对，无法导入。"
		MsgList(123) = "一部分子文件不存在，无法导入。\r\n合并时不能少一个子文件。"

		MsgList(124) = "全选"
		MsgList(125) = "查找(&F3)"
		MsgList(126) = "过滤显示"
		MsgList(127) = "查找"
		MsgList(128) = "过滤"
		MsgList(129) = "全部显示"
		MsgList(130) = "请输入要查找的内容。\r\n- 可使用 F3 快捷键调用此功能。查找内容不为空时，不显示该对话框。\r\n- 查找内容支持常规、通配符和正则表达式并自动判断。"
		MsgList(131) = "请输入要过滤的内容。\r\n注意：过滤内容支持常规、通配符和正则表达式并自动判断。"
		MsgList(132) = "仅找到 %s 一项。"
		MsgList(133) = "未找到 %s。"
		MsgList(134) = "查找内容判断为通配符，但语法错误。"
		MsgList(135) = "查找内容判断为正则表达式，但语法错误。"
		MsgList(136) = "过滤内容判断为通配符，但语法错误。"
		MsgList(137) = "过滤内容判断为正则表达式，但语法错误。"
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

		MsgList(12) = "PE 郎恨瞶竟 - セ %v (篶 %b)"
		MsgList(13) = "..."
		MsgList(14) = "(%s/%d)"
		MsgList(15) = "郎睲虫"
		MsgList(16) = "篈: %s"
		MsgList(17) = "闽"
		MsgList(18) = "粂ē"
		MsgList(19) = "耝"
		MsgList(20) = "蹲"
		MsgList(21) = "狡籹郎"
		MsgList(22) = "郎癟"
		MsgList(23) = "郎癟"
		MsgList(24) = "钡ㄖ"
		MsgList(25) = "郎癸霍ㄖ"

		MsgList(26) = "セ %v (篶 %b)\r\n" & _
					"OS セ: Windows XP/2000 ┪\r\n" & _
					"Passolo セ: Passolo 5.0 ┪\r\n" & _
					"甭舦: 禣硁砰\r\n" & _
					"呼: http://www.hanzify.org\r\n" & _
					": wanfu (2018 - 2019)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(27) = "闽 PE 郎恨瞶竟"
		MsgList(28) = "磅︽郎 (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|┮Τ郎 (*.*)|*.*||"
		MsgList(29) = "匡郎"
		MsgList(30) = "璣粂;虏砰いゅ;タ砰いゅ"
		MsgList(31) = "enu;chs;cht"
		MsgList(32) = "タ耝郎..."
		MsgList(33) = "耝 %s 郎"
		MsgList(34) = "耝 %s 郎纗 %d 戈Ж"
		MsgList(35) = "赣郎獶 PE 郎"
		MsgList(36) = "ぃ盢郎ヘ夹郎"
		MsgList(37) = "タㄖ郎..."
		MsgList(38) = "ㄖΘ"
		MsgList(39) = "ㄖア毖ぶ郎"
		MsgList(40) = "匡睦郎戈Ж"
		MsgList(41) = "匡郎┮戈Ж"
		MsgList(42) = "絋粄"
		MsgList(43) = "ゑ﹍;㎝﹍;ゑ﹍;郎"
		MsgList(44) = "タ蹲郎..."
		MsgList(45) = "蹲 %s 郎"

		MsgList(46) = "============ 郎癟 ============\r\n"
		MsgList(47) = "郎嘿\t%s"
		MsgList(48) = "郎隔畖\t%s"
		MsgList(49) = "郎弧\t%s"
		MsgList(50) = "郎セ\t%s"
		MsgList(51) = "玻珇嘿\t%s"
		MsgList(52) = "玻珇セ\t%s"
		MsgList(53) = "舦┮Τ\t%s"
		MsgList(54) = "郎\t%s じ舱"
		MsgList(55) = "ミら戳\t%s"
		MsgList(56) = "эら戳\t%s"
		MsgList(57) = "粂ē\t%s"
		MsgList(58) = "秨 祇 坝\t%s"
		MsgList(59) = "﹍郎\t%s"
		MsgList(60) = "ず场郎\t%s"
		MsgList(61) = "郎摸\t%s"
		MsgList(62) = "琈钩膀\t%s"
		MsgList(63) = "跋琿癟"
		MsgList(64) = "摸\t跋琿\t秨﹍\t挡\tじ舱"
		MsgList(65) = "%s!1!\t%s!2!\t%s!4!\t%s!5!\t%s!6!"
		MsgList(66) = "================================="
		MsgList(67) = "郎熬簿"
		MsgList(68) = "癸店览"
		MsgList(69) = "ヴ種"
		MsgList(70) = "留旅"
		MsgList(71) = "ゼ"
		MsgList(72) = "ぃノ"
		MsgList(73) = "戈ヘ魁癟 (郎熬簿)"
		MsgList(74) = "ヘ魁嘿\t┮跋琿\t秨﹍\t挡\tじ舱"
		MsgList(75) = "蹲ヘ魁"
		MsgList(76) = "蹲ヘ魁"
		MsgList(77) = "戈方ヘ魁"
		MsgList(78) = "钵盽ヘ魁"
		MsgList(79) = "ヘ魁"
		MsgList(80) = "膀﹚"
		MsgList(81) = "禘耞ヘ魁"
		MsgList(82) = "舦ヘ魁"
		MsgList(83) = "诀竟(GP RVA)"
		MsgList(84) = "磅︽狐セ诀纗"
		MsgList(85) = "更砞﹚ヘ魁"
		MsgList(86) = "竕﹚蹲ヘ魁"
		MsgList(87) = "蹲"
		MsgList(88) = "┑筐更蹲才"
		MsgList(89) = "COM 磅︽畐篨夹"
		MsgList(90) = "玂痙ヘ魁"
		MsgList(91) = "钵盽"
		MsgList(92) = "ぃ"
		MsgList(93) = ".NET CLR 戈ヘ魁癟 (郎熬簿)"
		MsgList(94) = "ヘ魁嘿\t┮跋琿\t秨﹍\t挡\tじ舱"
		MsgList(95) = "い膥戈(MetaData)"
		MsgList(96) = "Λ恨戈方"
		MsgList(97) = "眏嘿帽"
		MsgList(98) = "絏恨瞶"
		MsgList(99) = "店览(V-)"
		MsgList(100) = "铬臘蹲"
		MsgList(101) = "Λ恨セ诀琈钩繷"
		MsgList(102) = ".NET MetaData 戈瑈癟 (郎熬簿)"
		MsgList(103) = "戈瑈嘿\t┮跋琿\t秨﹍\t挡\tじ舱"
		MsgList(104) = "獶 PE 郎"
		MsgList(105) = "PE(%s)"
		MsgList(106) = "摸\t琿\t竊\t\t秨﹍\t挡\tじ舱"
		MsgList(107) = "%s!1!\t%s!2!\t%s!3!\t\t%s!4!\t%s!5!\t%s!6!"

		MsgList(108) = "タ莉 %s 郎癟..."
		MsgList(109) = "癟 - %s"
		MsgList(110) = ""
		MsgList(111) = ""
		MsgList(112) = "т %s 郎"

		MsgList(113) = "癟"
		MsgList(114) = "郎⊿Τ郎"
		MsgList(115) = "璶蹲戈Жい⊿Τ郎"
		MsgList(116) = "璶蹲戈Жい⊿Τт %s 戈郎礚猭蹲"
		MsgList(117) = "郎﹍郎嘿ぃ才礚猭蹲"
		MsgList(118) = "郎﹍郎セぃ才礚猭蹲"
		MsgList(119) = "郎﹍郎粂ēぃ才礚猭蹲"
		MsgList(120) = "郎﹍郎ぃ才礚猭蹲"
		MsgList(121) = "郎﹍郎ら戳跑琌膥尿\r\nら戳跑弧郎砆э筁э筁 郎ぃ続ノ"
		MsgList(122) = "%s 戈郎Αぃ癸礚猭蹲"
		MsgList(123) = "场郎ぃ礚猭蹲\r\nㄖぃぶ郎"

		MsgList(124) = "匡"
		MsgList(125) = "穓碝(&F3)"
		MsgList(126) = "筁耾陪ボ"
		MsgList(127) = "穓碝"
		MsgList(128) = "筁耾"
		MsgList(129) = "场陪ボ"
		MsgList(130) = "叫块璶穓碝ず甧\r\n- ㄏノ F3 е硉龄秸ノ穓碝ず甧ぃぃ陪ボ赣癸杠よ遏\r\n- 穓碝ず甧や穿盽砏窾ノじ㎝砏玥笲衡Α笆耞"
		MsgList(131) = "叫块璶筁耾ず甧\r\n猔種筁耾ず甧や穿盽砏窾ノじ㎝砏玥笲衡Α笆耞"
		MsgList(132) = "度т %s 兜"
		MsgList(133) = "ゼт %s"
		MsgList(134) = "穓碝ず甧耞窾ノじ粂猭岿粇"
		MsgList(135) = "穓碝ず甧耞砏玥笲衡Α粂猭岿粇"
		MsgList(136) = "筁耾ず甧耞窾ノじ粂猭岿粇"
		MsgList(137) = "筁耾ず甧耞砏玥笲衡Α粂猭岿粇"
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

		MsgList(12) = "PE Subfile Manager - Version %v (Build %b)"
		MsgList(13) = "..."
		MsgList(14) = "Add(%s/%d)"
		MsgList(15) = "Subfile List"
		MsgList(16) = "Status: %s"
		MsgList(17) = "About"
		MsgList(18) = "Language"
		MsgList(19) = "Extract"
		MsgList(20) = "Import"
		MsgList(21) = "Copy File Name"
		MsgList(22) = "Main File Info"
		MsgList(23) = "Sub File Info"
		MsgList(24) = "Direct Merge"
		MsgList(25) = "File Align Merge"

		MsgList(26) = "Version: %v (Build %b)\r\n" & _
					"OS Version: Windows XP/2000 or higher\r\n" & _
					"Passolo Version: Passolo 5.0 or higher\r\n" & _
					"License: Freeware\r\n" & _
					"HomePage: http://www.hanzify.org\r\n" & _
					"Author: wanfu (2018 - 2019)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(27) = "About PE Subfile Manager"
		MsgList(28) = "Executable File (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|All File (*.*)|*.*||"
		MsgList(29) = "Select File"
		MsgList(30) = "EngLish;Chinese Simplified;Chinese Traditional"
		MsgList(31) = "enu;chs;cht"
		MsgList(32) = "Extracting files..."
		MsgList(33) = "Total %s subfiles was extracted."
		MsgList(34) = "Total %s subfiles was extracted, saved to %d folder."
		MsgList(35) = "The file is not a PE file."
		MsgList(36) = "You cannot use the original file as the target file."
		MsgList(37) = "Merging files..."
		MsgList(38) = "Merge Successful!"
		MsgList(39) = "Merge failed! The sub file is missing."
		MsgList(40) = "Select folder to release subfiles"
		MsgList(41) = "Select folder that has subfiles"
		MsgList(42) = "Confirm"
		MsgList(43) = "Smaller than Original;Same as Original;Larger than Original;Sub File Name"
		MsgList(44) = "Importing files..."
		MsgList(45) = "Total %s subfiles was imported."

		MsgList(46) = "============ File Information ============\r\n"
		MsgList(47) = "File Name:\t%s"
		MsgList(48) = "File Path:\t\t%s"
		MsgList(49) = "Description:\t%s"
		MsgList(50) = "Version:\t\t%s"
		MsgList(51) = "Product Name:\t%s"
		MsgList(52) = "Product Version:\t%s"
		MsgList(53) = "Legal Copyright:\t%s"
		MsgList(54) = "File Size:\t\t%s bytes"
		MsgList(55) = "Date Created:\t%s"
		MsgList(56) = "Date Modified:\t%s"
		MsgList(57) = "Language:\t%s"
		MsgList(58) = "Company Name:\t%s"
		MsgList(59) = "Original File Name:\t%s"
		MsgList(60) = "Internal File Name:\t%s"
		MsgList(61) = "File Type:\t\t%s"
		MsgList(62) = "Image Base:\t%s"
		MsgList(63) = "Section Information:"
		MsgList(64) = "Address Category\tSection Name\tStart Address\tEnd Address\tByte Size"
		MsgList(65) = "%s!1!\t\t%s!2!\t%s!4!\t%s!5!\t%s!6!"
		MsgList(66) = "====================================="
		MsgList(67) = "Offset"
		MsgList(68) = "RVA"
		MsgList(69) = "Any"
		MsgList(70) = "Hide"
		MsgList(71) = "Unknown"
		MsgList(72) = "Not Available"
		MsgList(73) = "Data Directory Information (offset):"
		MsgList(74) = "Directory Name\t\t\tIn Section\tStart Address\tEnd Address\tByte Size"
		MsgList(75) = "Export Directory\t"
		MsgList(76) = "Import Directory\t"
		MsgList(77) = "Resource Directory\t"
		MsgList(78) = "Exception Directory\t"
		MsgList(79) = "Security Directory\t"
		MsgList(80) = "Base Relocation Table"
		MsgList(81) = "Debug Directory\t"
		MsgList(82) = "Copyright\t\t"
		MsgList(83) = "RVA of GP\t"
		MsgList(84) = "TLS Directory\t"
		MsgList(85) = "Load Configuration Directory"
		MsgList(86) = "Bound Import Directory"
		MsgList(87) = "Import Address Table"
		MsgList(88) = "Delay Load Import Descriptor"
		MsgList(89) = "COM Runtime Descriptor"
		MsgList(90) = "Reserved Directory\t"
		MsgList(91) = "Exception"
		MsgList(92) = "Not Exist"
		MsgList(93) = ".NET CLR Data Directory Information (offset):"
		MsgList(94) = "Directory Name\t\t\tIn Section\tStart Address\tEnd Address\tByte Size"
		MsgList(95) = "Meta Data\t"
		MsgList(96) = "Managed Resource\t"
		MsgList(97) = "Strong Name Signature"
		MsgList(98) = "Code Manager Table"
		MsgList(99) = "V-Table Fixups\t"
		MsgList(100) = "Export Address Table Jumps"
		MsgList(101) = "Managed Native Heade"
		MsgList(102) = ".NET MetaData iStreams Information (offset):"
		MsgList(103) = "Stream Name\tIn Section\tStart Address\tEnd Address\tByte Size"
		MsgList(104) = "Not PE File"
		MsgList(105) = "Sub PE(%s)"
		MsgList(106) = "Address Category\tSegment Name\tSection Name\tStart Address\tEnd Address\tByte Size"
		MsgList(107) = "%s!1!\t\t%s!2!\t\t%s!3!\t\t%s!4!\t%s!5!\t%s!6!"

		MsgList(108) = "Getting %s file information..."
		MsgList(109) = "Information - %s"
		MsgList(110) = "Previou"
		MsgList(111) = "Next"
		MsgList(112) = "Total %s subfiles has been found."

		MsgList(113) = "Information"
		MsgList(114) = "The file does not have subfiles."
		MsgList(115) = "No files in the folder to import."
		MsgList(116) = "The %s data file was not found in the folder to be imported and cannot be imported."
		MsgList(117) = "Original file name of subfiles does not match and cannot be imported."
		MsgList(118) = "Original file version of subfiles does not match and cannot be imported."
		MsgList(119) = "Original file language of subfiles does not match and cannot be imported."
		MsgList(120) = "Original file size of subfiles does not match and cannot be imported."
		MsgList(121) = "Original file date of subfiles has been changed, do you want to continue?" & _
						"\r\nIf the date has changed, the file has been modified and the modified file may not be applicable."
		MsgList(122) = "%s data file is not in the correct format and cannot be imported."
		MsgList(123) = "Some subfiles do not exist and cannot be imported. You cannot merge without one child file."

		MsgList(124) = "Select All"
		MsgList(125) = "Find(&F3)"
		MsgList(126) = "Filter Show"
		MsgList(127) = "Find"
		MsgList(128) = "Filter"
		MsgList(129) = "Show All"
		MsgList(130) = "Please enter the content to find." & _
						"\r\n- Use the F3 to call this feature." & _
						"The dialog box does not appear when the find content is not empty." & _
						"\r\n- Find content supports General, Wildcard and Regular expressions and automatic determine."
		MsgList(131) = "Please enter the content to filter." & _
						"\r\nNote: Find content supports General, Wildcard and Regular expressions and automatic determine."
		MsgList(132) = "Only one item of %s was found."
		MsgList(133) = "No %s found."
		MsgList(134) = "Find content is determined as a Wildcard, but the syntax is incorrect."
		MsgList(135) = "Find content is determined as a Regular expressions, but the syntax is incorrect."
		MsgList(136) = "Filter content is determined as a Wildcard, but the syntax is incorrect."
		MsgList(137) = "Filter content is determined as a Regular expressions, but the syntax is incorrect."
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
