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

'�����д���Ի�ƽ̨����
Private Enum PELangType
	DELPHI_FILE_SIGNATURE = &H50
	NET_FILE_SIGNATURE = &H424A5342
End Enum

'PE�ļ��ṹ(Visual Basic��)���ִ���һ
'ǩ������
Private Enum ImageSignatureTypes
	IMAGE_DOS_SIGNATURE = &H5A4D			'// MZ
	IMAGE_OS2_SIGNATURE = &H454E			'// NE
	IMAGE_OS2_SIGNATURE_LE = &H454C			'// LE
	IMAGE_VXD_SIGNATURE = &H454C			'// LE
	IMAGE_NT_SIGNATURE = &H4550				'// PE00
End Enum

'�ж���32λ����64λPE�ļ�
Private Enum ImageOptionalHeaderMagicType
	IMAGE_NT_OPTIONAL_HDR32_MAGIC = &H10B			'32λPE�ļ�
	IMAGE_NT_OPTIONAL_HDR64_MAGIC = &H20B			'64λPE�ļ�
End Enum

'�ļ�����(��־����)����
'IMAGE_FILE.Characteristics
'Private Enum ImageFileCharacteristicsTypes
'	IMAGE_FILE_RELOCS_STRIPPED = &H1				'�ض�λ��Ϣ���Ƴ�
'	IMAGE_FILE_EXECUTABLE_IMAGE = &H2				'�ļ���ִ��
'	IMAGE_FILE_LINE_NUMS_STRIPPED = &H4				'�кű��Ƴ�
'	IMAGE_FILE_LOCAL_SYMS_STRIPPED = &H8			'���ű��Ƴ�
'	IMAGE_FILE_AGGRESIVE_WS_TRIM = &H10				'Agressively Trim working Set
'	IMAGE_FILE_LARGE_ADDRESS_AWARE = &H20			'�����ܴ������2G�ĵ�ַ
'	IMAGE_FILE_BYTES_REVERSED_LO = &H80				'�����Ļ������͵�λ
'	IMAGE_FILE_32BIT_MACHINE = &H100				'32λ����
'	IMAGE_FILE_DEBUG_STRIPPED = &H200				'.dbg�ļ��ĵ�����Ϣ���Ƴ�
'	IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP = &H400		'������ƶ�������,���������ļ�������
'	IMAGE_FILE_NET_RUN_FROM_SWAP = &H800			'�����������,���������ļ�������
'	IMAGE_FILE_SYSTEM = &H1000						'ϵͳ�ļ�
'	IMAGE_FILE_DLL = &H2000							'�ļ���һ��dll
'	IMAGE_FILE_UP_SYSTEM_ONLY = &H4000				'�ļ�ֻ�������ڵ���������
'	IMAGE_FILE_BYTES_REVERSED_HI = &H8000			'�����Ļ������͸�λ.
'End Enum

'Ӧ�ó���ִ�еĻ�����ƽ̨���붨��
'IMAGE_FILE_HEADER.iMachine
'======================================================================================
'Private Enum ImageFileMachineTypes
'	IMAGE_FILE_MACHINE_UNKNOWN = &H0		'δ֪
'	IMAGE_FILE_MACHINE_I386 = &H14C			'Intel 80386 ����������
'	IMAGE_FILE_MACHINE_I486 = &H14D			'Intel 80486 ����������
'	IMAGE_FILE_MACHINE_IPTM = &H14E			'Intel Pentium ����������
'	IMAGE_FILE_MACHINE_R 	= &H160			'R3000(MIPS)��������big endian
'	IMAGE_FILE_MACHINE_R3000 = &H162		'R3000(MIPS)��������little endian
'	IMAGE_FILE_MACHINE_R4000 = &H166		'R4000(MIPS)��������little endian
'	IMAGE_FILE_MACHINE_R10000 = &H168		'R10000(MIPS)��������little endian
'	IMAGE_FILE_MACHINE_WCEMIPSV2 = &H169	'MIPS Little-endian WCE v2
'	IMAGE_FILE_MACHINE_ALPHA = &H184		'DEC Alpha AXP������
'	IMAGE_FILE_MACHINE_POWERPC= &H1F0		'IBM Power PC��little endian
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

'����Ŀ¼��
'======================================================================================
'Private Enum ImageDirectoryEntry
'	IMAGE_DIRECTORY_ENTRY_EXPORT = 0				'����Ŀ¼
'	IMAGE_DIRECTORY_ENTRY_IMPORT = 1				'����Ŀ¼
'	IMAGE_DIRECTORY_ENTRY_RESOURCE = 2				'��ԴĿ¼
'	IMAGE_DIRECTORY_ENTRY_EXCEPTION = 3				'�쳣Ŀ¼
'	IMAGE_DIRECTORY_ENTRY_SECURITY = 4				'��ȫĿ¼
'	IMAGE_DIRECTORY_ENTRY_BASERELOC = 5				'�ض�λ������
'	IMAGE_DIRECTORY_ENTRY_DEBUG = 6					'����Ŀ¼
'	IMAGE_DIRECTORY_ENTRY_COPYRIGHT = 7				'X86ʹ��-��������
'	IMAGE_DIRECTORY_ENTRY_ARCHITECTURE = 7			'Architecture Specific Data
'	IMAGE_DIRECTORY_ENTRY_GLOBALPTR = 8				'����ֵ(MIPS GP),�� RVA of GlobalPtr
'	IMAGE_DIRECTORY_ENTRY_TLS = 9					'�̱߳��ش洢(Thread Local Storage,TLS)Ŀ¼
'	IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG = 10			'��������Ŀ¼
'	IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT = 11			'�������(Bound Import Directory in headers)
'	IMAGE_DIRECTORY_ENTRY_IAT = 12					'�����ַ��
'	IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT = 13			'Delay Load Import Descriptors
'	IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR = 14		'COM ���б�־
'	IMAGE_DIRECTORY_ENTRY_RESERVED = 15				'����
'End Enum

'�����Զ���
Private Enum ImageSectionType
	IMAGE_SCN_CNT_CODE = &H20						'���а�������
	IMAGE_SCN_CNT_INITIALIZED_DATA = &H40			'���а����ѳ�ʼ������
	IMAGE_SCN_CNT_UNINITIALIZED_DATA = &H80			'���а���δ��ʼ������
	IMAGE_SCN_MEM_DISCARDABLE = &H2000000			'��һ���ɶ����Ľڣ������е������ڽ��̿�ʼ�󽫱�����
	IMAGE_SCN_MEM_NOT_CACHED = &H4000000			'�������ݲ���������
	IMAGE_SCN_MEM_NOT_PAGED = &H8000000				'�������ݲ����������ڴ�
	IMAGE_SCN_MEM_SHARED = &H10000000				'�������ݿɹ���
	IMAGE_SCN_MEM_EXECUTE = &H20000000				'��ִ�н�
	IMAGE_SCN_MEM_READ = &H40000000					'�ɶ���
	IMAGE_SCN_MEM_WRITE = &H80000000				'��д��
End Enum

'��Դ����
'======================================================================================
'Private Enum ImageResourceEntry
'	IMAGE_RESOURCE_CURSOR = 1						'���						'
'	IMAGE_RESOURCE_BITMAP = 2						'λͼ
'	IMAGE_RESOURCE_ICON = 3							'ͼ��
'	IMAGE_RESOURCE_MENU = 4							'�˵�
'	IMAGE_RESOURCE_DIALOG = 5						'�Ի���
'	IMAGE_RESOURCE_STRING_TABLE = 6					'�ַ�����
'	IMAGE_RESOURCE_FONT_DIRECTORY = 7				'����Ŀ¼
'	IMAGE_RESOURCE_FONT	= 8							'����
'	IMAGE_RESOURCE_ACCELERATORS = 9					'������
'	IMAGE_RESOURCE_UNFORMATTED_RESOURCE_DATA = 10	'δ��ʽ����Դ����
'	IMAGE_RESOURCE_MESSAGE_TABEL = 11				'��Ϣ��
'	IMAGE_RESOURCE_GROUP_CURSOR = 12				'�����
'	IMAGE_RESOURCE_GROUP_ICON = 13					'ͼ����
'	IMAGE_RESOURCE_VERSION_INFO = 14				'�汾��Ϣ
'End Enum

'======================================================================================
'�ṹ��: IMAGE_DOS_HEADER
'�ṹ��С: 64�ֽ�
'�ṹ˵��: DOSӳ��ͷ(��EXEͷ)
'======================================================================================
Private Type IMAGE_DOS_HEADER
	iSignature						As Integer		'&H0    ǩ��("MZ",��&H5A4D)
	iLastPageBytes					As Integer		'&H2    �ļ����ҳ�е��ֽ���
	iPages							As Integer		'&H4    �ļ�ҳ��
	iRelocateItems					As Integer		'&H6    �ض�λԪ�ظ���
	iHeaderSize						As Integer		'&H8    ͷ����С
	iMinAlloc						As Integer		'&HA    �������С���Ӷ�
	iMaxAlloc						As Integer		'&HC    �������󸽼Ӷ�
	iInitialSS						As Integer		'&HE    ��ʼSSֵ
	iInitialSP						As Integer		'&H10   ��ʼSPֵ
	iCheckSum						As Integer		'&H12   У���
	iInitialIP						As Integer		'&H14   ��ʼIPֵ
	iInitialCS						As Integer		'&H16   ��ʼCSֵ(���ƫ����)
	iRelocateTable					As Integer		'&H18   �ض�����ļ���ַ
	iOverlay						As Integer		'&H1A   ���Ǻ�
	iReserved(3)					As Integer		'&H22   ������
	iOEMID							As Integer		'&H24   OEM��ʶ��
	iOEMInformation					As Integer		'&H26   OEM��Ϣ
	iReserved2(9)					As Integer		'&H28   ������2
	lPointerToPEHeader				As Long			'&H3C   PEͷ��λ��
End Type

'======================================================================================
'�ṹ��: IMAGE_FILE_HEADER
'�ṹ��С: 24�ֽ�
'�ṹ˵��: ӳ���ļ�ͷ
'======================================================================================
Private Type IMAGE_FILE_HEADER
	lSignature						As Long			'&H4	PE�ļ�ͷ��־("PE00",��&H4550),4�ֽ�
	iMachine						As Integer		'&H6    ִ�иó���Ļ�����ƽ̨
	iNumberOfSections				As Integer		'&H8    �ļ��нڵĸ���
	lTimeDateStamp					As Long			'&HC    �ļ�����ʱ��(ʱ���)
	lPointerToSymbolTable			As Long			'&H10   COFF���ű�ƫ��
	lNumberOfSymbols				As Long			'&H14   ������Ŀ
	iSizeOfOptionalHeader			As Integer		'&H16   ��ѡͷ����С
	iCharacteristics				As Integer		'&H18   ��־����
End Type

'======================================================================================
'�ṹ��: IMAGE_DATA_DIRECTORY
'�ṹ��С: ÿ��Ŀ¼8�ֽڣ���16��Ŀ¼����120�ֽ�
'�ṹ˵��: ����Ŀ¼��
'0 = ����Ŀ¼
'1 = ����Ŀ¼
'2 = ��ԴĿ¼
'3 = �쳣Ŀ¼
'4 = ��ȫĿ¼
'5 = ��ַ�ض�λ��
'6 = ����Ŀ¼
'7 = ��ȨĿ¼
'8 = ����ֵ(GP RVA)
'9 = �̱߳��ش洢��
'10 = ��������Ŀ¼
'11 = �󶨵���Ŀ¼
'12 = �����ַ��
'13 = �ӳټ��ص����
'14 = COM ���п��־(.NET ����� RVA = CLR ��ַ��Size = CLR header �Ĵ�С���̶�Ϊ48�ֽ�)
'15 = ����Ŀ¼
'======================================================================================
Private Type IMAGE_DATA_DIRECTORY
	lVirtualAddress			As Long			'��ʼRVA��ַ
	lSize					As Long			'lVirtualAddress��ָ�����ݽṹ���ֽ���
End Type

'======================================================================================
'�ṹ��: IMAGE_OPTIONAL_HEADER32
'�ṹ��С: 224�ֽ�
'�ṹ˵��: ��ѡӳ��ͷ
'======================================================================================
Private Type IMAGE_OPTIONAL_HEADER32
	'******************
	'��׼��
	'******************
	iMagic							As Integer		'&H18   32λPE��&H10B��64λPE��&H20B
	bMajorLinkerVersion				As Byte			'&H1A   ���������汾
	bMinorLinkerVersion				As Byte			'&H1B   �������ΰ汾
	lSizeOfCode						As Long			'&H1C   ��ִ�д��볤��
	lSizeOfInitializedData			As Long			'&H20   ��ʼ�����ݳ���(���ݽ�)
	lSizeOfUninitializedData		As Long			'&H24   δ��ʼ�����ݳ���(bss��)
	lAddressOfEntryPoint			As Long			'&H28   �������RVA��ַ,������⿪ʼִ��
	lBaseOfCode						As Long			'&H2C   ��ִ�д�����ʼλ��
	lBaseOfData						As Long			'&H30   ��ʼ��������ʼλ��
	'******************
	'NT ������
	'******************
	lImageBase						As Long			'&H34   ���������ѡ��RVA��ַ(32λ)
	lSectionAlignment				As Long			'&H38   ���غ�����ڴ��еĶ��뷽ʽ
	lFileAlignment					As Long			'&H3C   �����ļ��еĶ��뷽ʽ
	iMajorOperatingSystemVersion	As Integer		'&H40   ����ϵͳ���汾
	iMinorOperatingSystemVersion	As Integer		'&H42   ����ϵͳ�ΰ汾
	iMajorImageVersion				As Integer		'&H44   ��ִ���ļ����汾
	iMinorImageVersion				As Integer		'&H46   ��ִ���ļ��ΰ汾
	iMajorSubsystemVersion			As Integer		'&H48   ��ϵͳ���汾
	iMinorSubsystemVersion			As Integer		'&H50   ��ϵͳ�ΰ汾
	lWin32VersionValue				As Long			'&H52   Win32�汾��,һ��Ϊ0
	lSizeOfImage					As Long			'&H56   ��������ռ���ڴ��С(�����С)
	lSizeOfHeaders					As Long			'&H5A   ͷ����С(ƫ�ƴ�С)
	lCheckSum						As Long			'&H5E   У���
	iSubsystem						As Integer		'&H62   ��ִ���ļ�����ϵͳ
	iDllCharacteristics				As Integer		'&H64   ��ʱDllMain������,һ��Ϊ0
	lSizeOfStackReserve				As Long			'&H66   ��ʼ���߳�ʱ������ջ��С
	lSizeOfStackCommit				As Long			'&H6A   ��ʼ���߳�ʱ�ύ��ջ��С
	lSizeOfHeapReserve				As Long			'&H6E   ���̳�ʼ��ʱ������ջ��С
	lSizeOfHeapCommit				As Long			'&H72   ���̳�ʼ��ʱ�ύ��ջ��С
	lLoaderFlags					As Long			'&H76   װ�ر�־,��������
	lNumberOfRvaAndSizes			As Long			'&H7A   ����Ŀ¼������,һ��Ϊ16
	DataDirectory(15) 				As IMAGE_DATA_DIRECTORY			'DataDirectory(14)
End Type

'======================================================================================
'�ṹ��: IMAGE_OPTIONAL_HEADER64
'�ṹ��С: 240�ֽ�
'�ṹ˵��: ��ѡӳ��ͷ
'======================================================================================
Private Type IMAGE_OPTIONAL_HEADER64
	'******************
	'��׼��
	'******************
	iMagic							As Integer		'&H18   32λPE��&H10B��64λPE��&H20B
	bMajorLinkerVersion				As Byte			'&H1A   ���������汾
	bMinorLinkerVersion				As Byte			'&H1B   �������ΰ汾
	lSizeOfCode						As Long			'&H1C   ��ִ�д��볤��
	lSizeOfInitializedData			As Long			'&H20   ��ʼ�����ݳ���(���ݽ�)
	lSizeOfUninitializedData		As Long			'&H24   δ��ʼ�����ݳ���(bss��)
	lAddressOfEntryPoint			As Long			'&H28   �������RVA��ַ,������⿪ʼִ��
	lBaseOfCode						As Long			'&H2C   ��ִ�д�����ʼλ��
	'lBaseOfData					As Long			'&H30   ��ʼ��������ʼλ��
	'******************
	'NT ������
	'******************
	dImageBase(7)					As Byte			'&H34   ���������ѡ��RVA��ַ(64λ)
	lSectionAlignment				As Long			'&H38   ���غ�����ڴ��еĶ��뷽ʽ
	lFileAlignment					As Long			'&H3C   �����ļ��еĶ��뷽ʽ
	iMajorOperatingSystemVersion	As Integer		'&H40   ����ϵͳ���汾
	iMinorOperatingSystemVersion	As Integer		'&H42   ����ϵͳ�ΰ汾
	iMajorImageVersion				As Integer		'&H44   ��ִ���ļ����汾
	iMinorImageVersion				As Integer		'&H46   ��ִ���ļ��ΰ汾
	iMajorSubsystemVersion			As Integer		'&H48   ��ϵͳ���汾
	iMinorSubsystemVersion			As Integer		'&H50   ��ϵͳ�ΰ汾
	lWin32VersionValue				As Long			'&H52   Win32�汾��,һ��Ϊ0
	lSizeOfImage					As Long			'&H56   ��������ռ���ڴ��С(�����С)
	lSizeOfHeaders					As Long			'&H5A   ͷ����С(ƫ�ƴ�С)
	lCheckSum						As Long			'&H5E   У���
	iSubsystem						As Integer		'&H62   ��ִ���ļ�����ϵͳ
	iDllCharacteristics				As Integer		'&H64   ��ʱDllMain������,һ��Ϊ0
	dSizeOfStackReserve				As Double		'&H66   ��ʼ���߳�ʱ������ջ��С(64λ)
	dSizeOfStackCommit				As Double		'&H6A   ��ʼ���߳�ʱ�ύ��ջ��С(64λ)
	dSizeOfHeapReserve				As Double		'&H6E   ���̳�ʼ��ʱ������ջ��С(64λ)
	dSizeOfHeapCommit				As Double		'&H72   ���̳�ʼ��ʱ�ύ��ջ��С(64λ)
	lLoaderFlags					As Long			'&H76   װ�ر�־,��������
	lNumberOfRvaAndSizes			As Long			'&H7A   ����Ŀ¼������,һ��Ϊ16
	DataDirectory(15) 				As IMAGE_DATA_DIRECTORY
End Type

'======================================================================================
'�ṹ��: IMAGE_SECTION_HEADER
'�ṹ��С: 40�ֽ�
'�ṹ˵��: ��ӳ��ͷ
'======================================================================================
Private Type IMAGE_SECTION_HEADER
	sName(7)						As Byte			'&H0    ����(���8�����ֽ��ַ�)
	'lPhysicalAddress				As Long			'&H8    OBJ�ļ��б�ʾ���ڵ������ַ
	lVirtualSize					As Long			'&H8    EXE�ļ��б�ʾ�ڵ�ʵ���ֽ���
	lVirtualAddress					As Long			'&HC    ���ڵ�RVA
	lSizeOfRawData					As Long			'&H10   ���ھ��ļ������ĳߴ�
	lPointerToRawData				As Long			'&H14   ����ԭʼ�������ļ��е�λ��
	lPointerToRelocations			As Long			'&H18   OBJ�ļ��б�ʾ�����ض�λ��Ϣ��ƫ��,EXE�ļ���������
	lPointerToLineNumbers			As Long			'&H1C   �к�ƫ��
	iNumberOfRelocations			As Integer		'&H20   �������ض�λ����Ŀ
	iNumberOfLineNumbers			As Integer		'&H22   �������кű��е��к���Ŀ
	lCharacteristics				As Long			'&H24   ������
End Type

'======================================================================================
'�ṹ��: .NET CLR 2.0 ͷ�ṹ
'�ṹ��С: 72�ֽ�
'�ṹ˵��: CLR ͷ
'======================================================================================
Private Type IMAGE_CLR20_HEADER
	'Header versioning
	cb						As Long					'CLRͷ�Ĵ�С����byteΪ��λ
	MajorRuntimeVersion		As Integer				'�����иó������С.NET�汾�����汾��
	MinorRuntimeVersion		As Integer				'�����иó����.NET�汾�ĸ��汾��

	'Symbol table And startup information
	METADATA				As IMAGE_DATA_DIRECTORY	'Ԫ���ݵ�RVA��Size
	Flags					As Long					'�����ֶΣ�������IL����.corflags������ʽ���ã�
													'Ҳ�����ڱ���ʱ��/FLAGSѡ��������ã��������������õ����ȼ��ϸ�
	EntryPointToken			As Long					'��ڷ�����Ԫ����ID��Ҳ����token������EXE�ļ������У�
													'��DLL�ļ��д������Ϊ0��.NET 2.0�У���������Ǳ�����ڴ����RVAֵ��
	'Binding information
	Resources				As IMAGE_DATA_DIRECTORY	'�й���Դ��RVA��Size
	StrongNameSignature		As IMAGE_DATA_DIRECTORY	'ǿ�������ݵ�RVA��Size��ǿ���Ƶ������ں�����ܣ�

	'Regular fixup And binding information
	CodeManagerTable		As IMAGE_DATA_DIRECTORY	'CodeManagerTable��RVA��Size��������δʹ�ã�Ϊ0
    VTableFixups			As IMAGE_DATA_DIRECTORY	'v-table���RVA��Size����Ҫ��ʹ��v-table��C++���Խ����ض�λ
    ExportAddressTableJumps	As IMAGE_DATA_DIRECTORY	'����C++�������ת��ַ���RVA��Size����������Ϊ0

    'Precompiled image info (internal use only - Set To zero)
    ManagedNativeHeader		As IMAGE_DATA_DIRECTORY	'������ngen���ɱ���ģ���и��Ϊ0�����������Ϊ0
End Type

'����Ŀ¼�� CLR Header Flags
'======================================================================================
'Private Enum CLR_HEADER_FLAGES
'	COMIMAGE_FLAGS_ILONLY = &H1				'��CLR�����ɴ�IL�������
'	COMIMAGE_FLAGS_32BITREQUIRED = &H2		'��CLRӳ��ֻ����32λϵͳ��ִ��
'	COMIMAGE_FLAGS_IL_LIBRARY = &H4			'��CLRӳ������ΪIL�������ڵ�
'	COMIMAGE_FLAGS_STRONGNAMESIGNED = &H8	'�ļ��ܵ�ǿ����ǩ���ı���
'	COMIMAGE_FLAGS_NATIVE_ENTRYPOINT =&H8	'�˳�����ڷ���Ϊ���й�
'	COMIMAGE_FLAGS_TRACKDEBUGDATA = &H10000	'Loader��JIT��Ҫ׷�ٵ�����Ϣ��ȱʡ��0
'End Enum

'======================================================================================
'�ṹ��: .NET MetaData ͷ�ṹ
'�ṹ��С: ���̶�
'�ṹ˵��: MetaData ͷ
'======================================================================================
Private Type IMAGE_METADATA_HEADER
	lSignature		As Long   		'Magic signature For physical metadata, currently 0x424A5342(BSJB Ϊ.Net �ļ��ı�־)
	iMajorVersion	As Integer		'Major version (1 for the first release of the common language runtime)
	iMinorVersion	As Integer		'Minor Version (1 For the first release of the common language runtime)
	lExtraData		As Long			'Reserved, always 0
	lLength			As Long			'Length of Version String In bytes
	Version()		As Byte			'�汾�ַ���, UTF8 ���룬4�ֽڶ���
	fFlags			As Integer		'Reserved, always 0
	iStreams		As Integer		'Number of streams
	'StreamHeader()	As IMAGE_STREAM_HEADER
End Type

'======================================================================================
'�ṹ��: .NET Stream ͷ�ṹ
'�ṹ��С: 72�ֽ�
'�ṹ˵��: Stream ͷ
'������
'#Strings: UTF8��ʽ���ַ����ѣ���������Ԫ���ݵ����ƣ���������������������Ա�����������ȣ���
'          �����ײ�����һ��0��Ϊ���ַ��������ַ�����0��ʾ��β��CLR����Щ���Ƶ���󳤶���1024��
'#Blob:    ���������ݶѣ��洢�����еķ��ַ�����Ϣ�����糣��ֵ��������signature��PublicKey�ȡ�
'          ÿ�����ݵĳ����ɸ����ݵ�ǰ1��3λ������0��ʾ���ȣ��ֽڣ�10��ʾ����2�ֽڣ�110��ʾ����4�ֽڡ�
'#GUID:    �洢���е�ȫ��Ψһ��ʶ��Global Unique Identifier����
'#US:      ��Unicode��ʽ��ŵ�IL������ʹ�õ��û��ַ�����User String��������ldstr���õ��ַ�����
'#~:       Ԫ���ݱ���������Ҫ�������������е�Ԫ������Ϣ���Ա����ʽ�����ڴˡ�ÿ��.Net ���򶼱��������
'#-:       #~��δѹ�������Ϊδ�Ż����洢����������
'======================================================================================
Private Type IMAGE_STREAM_HEADER
	lOffset			As Long			'����� Metadata Root ���ڴ�ƫ��
	lSize			As Long			'�����ֽڴ�С��4 �ı���
	rcName()		As Byte			'�Կ��ֽ���ֹ�� ASCII �ַ������飬4�ֽڶ���
	RWA				As Long			'���ݼ�¼��������ƫ�Ƶ�ַ (ʵ��û�����ֵ��ֻ��Ϊ�˷��㶨λ��������λ��)
End Type

'#Blob��
'#Blob����һ�����������ݶѣ������е����з��ַ�����ʽ���ݶ��ѷ�����������棬
'�糣����ֵ��Public Key��ֵ��������Signature�ȵȡ�
'��ÿ�����������ݿ�ͷ������һ���鳤�����ݣ���Ϊ�˽�Լ�洢�ռ䣬CLRʹ���˱Ƚ��鷳�ı��뷽����
'�����ʼһ���ֽ����λΪ0��������ݿ鳤��Ϊһ���ֽڣ�
'�����ʼһ���ֽ����λΪ10��������ݿ鳤��Ϊ�����ֽڣ�
'�����ʼһ���ֽ����λΪ110��������ݿ鳤��Ϊ�ĸ��ֽڣ�
'�����α�־λ��ͨ����λ���㼴�ɼ�������ݿ��ʵ�ʳ���ֵ�������ݴ˻�����ݡ�

'#US��
'һ��Blob�ѣ��������û��Զ�����ַ�����
'����������˶������û������е��ַ�����������Щ�ַ�����UTF-16�ı����ʽ���棬�����Ŷ����һ��β������Ϊ0��1���ֽڣ�
'����ָ�����ַ������Ƿ��д���0x007F�Ĵ����ַ���
'���β���ֽڱ���ӵ������ϵ������û�������ַ����������ɵ��ַ��������ϵĴ���ת��������
'�����������Ȥ�������ǣ�����û��ַ��������ᱻ����Ԫ���ݱ����õ���������ʾ�ر�IL���������ַ��ʹ��ldstrָ���
'���⣬��Ϊһ��ʵ���ϵ�blob�ѣ�US�Ѳ������Դ洢Unicode�ַ����������Դ洢��������ƶ�����ʹ��Щ��Ȥ��ʵ�ֳ�Ϊ���ܡ�

'======================================================================================
'�ṹ��: .NET #~ Stream �ṹ
'�ṹ��С: 24�ֽ�
'�ṹ˵��: #~ Stream����������ǿǩ����ص�����
'======================================================================================
'Private Type TClrTableStreamHeader
'	Reserved		As Long			'������Ϊ0
'	MajorVersion	As Byte			'Ԫ���ݱ�����汾�ţ���.Net���汾��һ��
'	MinorVersion	As Byte			'Ԫ���ݱ�ĸ��汾�ţ�һ��Ϊ0
'	HeapSizes		As Byte			'heaps �ж�λ����ʱ�������Ĵ�С��Ϊ0��ʾ16λ����ֵ
									'���������ݳ���16λ���ݱ�ʾ��Χ����ʹ��32λ����ֵ��
									'01h���� strings �ѣ�02h���� GUID �ѣ�04h���� blob ��
									'��#-���п���Ϊ20h��80h��ǰ�ߴ������а�����Edit-and-Continue�ĵ������޸ĵ����ݣ�
									'���߱�ʾԪ�����и������ʶΪ��ɾ��
'	Rid				As Byte			'����Ԫ���ݱ��м�¼���������ֵ��������ʱ��.Net���㣬�ļ���ͨ��Ϊ1
'	MaskValid		As Double		'8�ֽڳ��ȵ����룬ÿ��λ����һ����Ϊ1��ʾ�ñ���Ч��Ϊ0��ʾ�޸ñ�
'	Sorted			As Double		'8�ֽڳ��ȵ����룬ÿ��λ����һ����Ϊ1��ʾ�ñ������򣬷�֮Ϊ0
'End Type


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

'COMMAND_64 ����
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

'�ڶ���
Private Type SUB_SECTION_PROPERTIE
	sName				As String	'������
	RWA					As Long		'���ݼ�¼��������ƫ�Ƶ�ַ (ʵ��û�����ֵ��ֻ��Ϊ�˷��㶨λ��������λ��)
	lVirtualSize		As Long		'EXE�ļ��б�ʾ�ڵ�ʵ���ֽ���
	lVirtualAddress		As Long		'���ڵĵ�λRVA
	lVirtualAddress1	As Long		'���ڵĸ�λRVA (64 λ�ļ�)
	lSizeOfRawData		As Long		'���ھ��ļ������ĳߴ�
	lPointerToRawData	As Long		'����ԭʼ�������ļ��е�λ��
End Type

'�ڶ���
Private Type SECTION_PROPERTIE
	sName				As String	'������
	RWA					As Long		'���ݼ�¼��������ƫ�Ƶ�ַ (ʵ��û�����ֵ��ֻ��Ϊ�˷��㶨λ��������λ��)
	lVirtualSize		As Long		'EXE�ļ��б�ʾ�ڵ�ʵ���ֽ���
	lVirtualAddress		As Long		'���ڵĵ�λRVA
	lVirtualAddress1	As Long		'���ڵĸ�λRVA (64 λ�ļ�)
	lSizeOfRawData		As Long		'���ھ��ļ������ĳߴ�
	lPointerToRawData	As Long		'����ԭʼ�������ļ��е�λ��
	SubSecs				As Integer	'�ӽ���
	SubSecList()		As SUB_SECTION_PROPERTIE
End Type

'�ļ�����
Private Type FILE_PROPERTIE
	CompanyName			As String	'������˾����
	FileDescription		As String	'�ļ�����
	FileVersion			As String	'�ļ��汾
	InternalName		As String	'�ڲ�����
	LegalCopyright		As String	'��Ȩ
	OrigionalFileName	As String	'ԭʼ�ļ���
	ProductName			As String	'��Ʒ����
	ProductVersion		As String	'��Ʒ�汾
	LanguageID			As String	'����ID
	FileSize			As Long		'�ļ���С
	DateCreated			As Date		'�ļ���������
	DateLastModified	As Date		'�ļ�����޸�����
	CopePage			As Long		'�ļ�����ҳ
	FileName			As String	'�ļ��� (����·��)
	FilePath			As String	'�ļ�·�� (���ļ���)
	FileType			As Integer	'�ļ����͵Ŀ�ʼ��ַ���磺PE �ļ���ʼ���� MZ
	SubFileDir			As String	'���ļ������ļ���·��
	Info 				As String	'�ļ�������Ϣ�������ظ���ȡ

	Magic				As String	'�ļ����ͣ�PE32,NET32,PE64,NET64,MAC32,MAC64������Ϊ��
	SecAlign			As Long		'�ڴ����ֵ
	FileAlign			As Long		'�ļ�����ֵ
	MinSecID 			As Integer	'��Сƫ�ƿ�ʼ��ַ���ڽ�������
	MaxSecID 			As Integer	'���ƫ�ƿ�ʼ��ַ���ڽ�������
	MaxSecIndex 		As Integer	'�ļ��ڵ����������
	USStreamID 			As Integer	'.NET �ַ���������������
	LangType			As Long		'�ļ��ı�д����
	NumberOfSub			As Integer	'��Ƕ���� PE ��
	ImageBase			As Variant	'�����ļ���RVA����ַ
	DataDirs			As Integer	'����Ŀ¼������
	NetStreams			As Integer	'.NET �ļ���������
	SecList() 			As SECTION_PROPERTIE
	DataDirectory()		As SUB_SECTION_PROPERTIE
	CLRList()			As SUB_SECTION_PROPERTIE
	StreamList()		As SUB_SECTION_PROPERTIE
End Type

'�ڴ�ӳ���ļ�����
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

'������Ϣ����
Private Enum FormatMSG
	FORMAT_MESSAGE_FROM_SYSTEM = &H1000
	FORMAT_MESSAGE_IGNORE_INSERTS = &H200
End Enum

'���ļ���ʽ�Ľṹ��
Private Type FILE_IMAGE
	ModuleName				As String	'�������ļ����ļ���
	hFile					As Long		'���� Create �ļ�ӳ��� OpenFile �ľ��
	hMap					As Long		'���� CreateFileMapping �ļ�ӳ��ľ��
	MappedAddress			As Long		'�ļ�ӳ�䵽���ڴ��ַ
	SizeOfImage				As Long		'ӳ��� Image ���ֽ�����Ĵ�С
	SizeOfFile				As Long		'ʵ���ļ���С
	ImageByte()				As Byte		'�ļ����ֽ�����
End Type

'����ҳת��
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

'�ڴ渴�ƺͱȽϺ���
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

'���ڶ�д�ļ�����
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

'�����ļ�ӳ�亯��
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

'������ʾ������Ϣ
Private Declare Function GetLastError Lib "kernel32.dll" () As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" ( _
	ByVal dwFlags As Long, _
	lpSource As Any, _
	ByVal dwMessageId As Long, _
	ByVal dwLanguageId As Long, _
	ByVal lpBuffer As String, _
	ByVal nSize As Long, _
	Arguments As Long) As Long

'���ڶ�ȡ�ļ��汾��Ϣ����
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

'����ļ���
Private Type BrowseInfo
	hWndOwner		As Long		'����ļ��жԻ���ĸ�������
	pIDLRoot		As Long		'ITEMIDLIST�ṹ�ĵ�ַ���������ʱ�ĳ�ʼ��Ŀ¼��������NULL����ʱ����Ŀ¼����ʹ��
	pszDisplayName	As Long		'���������û�ѡ�е�Ŀ¼�ַ������ڴ��ַ
	lpszTitle		As String	'��ʾλ�ڶԻ������ϲ��ı���
	ulFlags			As Long		'ָ���Ի������ۺ͹��ܵı�־
	lpfnCallback	As Long		'�����¼��Ļص�����
	lParam			As Long		'Ӧ�ó��򴫸��ص������Ĳ���
	iImage			As Long		'���汻ѡȡ���ļ��е�ͼƬ����
End Type

'����ļ��в���
Private Enum BrowseFolder
	BIF_RETURNONLYFSDIRS = &H1		'�������ļ�ϵͳ��Ŀ¼
	BIF_DONTGOBELOWDOMAIN = &H2		'�������Ӵ��У��������������µ�����Ŀ¼�ṹ
	BIF_STATUSTEXT = &H4&			'����һ��״̬����ͨ�����Ի�������Ϣʹ�ص���������״̬�ı�
	BIF_EDITBOX = &H10				'����һ���༭���û���������ѡ���������
	BIF_BROWSEINCLUDEURLS = &H80
	BIF_RETURNFSANCESTORS = &H8		'�����ļ�ϵͳ��һ���ڵ�
	BIF_VALIDATE = &H20				'û��BIF_EDITBOX��־λʱ���ñ�־λ�����ԡ�����û���������ַǷ���������BFFM_VALIDATEFAILED��Ϣ���ص�����
	BIF_NEWDIALOGSTYLE = &H40
	BIF_USENEWUI = BIF_EDITBOX Or BIF_NEWDIALOGSTYLE	'�Ի��������½��ļ��а�ť
	BIF_UAHINT = &H100
	BIF_NONEWFOLDERBUTTON = &H200
	BIF_NOTRANSLATETARGETS = &H400
	BIF_BROWSEFORCOMPUTER = &H1000	'���ؼ������
	BIF_BROWSEFORPRINTER = &H2000	'���ش�ӡ����
	BIF_BROWSEINCLUDEFILES = &H4000	'���������ʾĿ¼��ͬʱҲ��ʾ�ļ�
	BIF_SHAREABLE = &H8000
	BIF_BROWSEFILEJUNCTIONS = &H10000
End Enum

'����ļ��к���
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" ( _
	ByVal pidList As Long, _
	ByVal lpBuffer As String) As Long

'�����ı�
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

'���ڷ��ؿؼ�ID�ľ��
Private Declare Function GetDlgItem Lib "user32.dll" ( _
	ByVal hDlg As Long, _
	ByVal nIDDlgItem As Long) As Long

'SendMessage API ���ֳ���
Private Enum SendMsgValue
	EM_GETLIMITTEXT = &HD5			'0,0				��ȡһ���༭�ؼ����ı�����󳤶�
	EM_LIMITTEXT = &HC5				'���ֵ,0			���ñ༭�ؼ��е�����ı�����
	WM_GETTEXT = &H0D				'�ֽ���,�ַ�����ַ	��ȡ�����ı��ؼ����ı�
	WM_GETTEXTLENGTH = &H0E			'0,0				��ȡ�����ı��ؼ����ı��ĳ���(�ֽ���)
	WM_SETTEXT = &H0C				'0,�ַ�����ַ		���ô����ı��ؼ����ı�
	WM_VSCROLL = &H115				'�ؼ����,����������,������λ��	���� SB_BOTTOM ָ���Ĵ�ֱ������λ��
	SB_BOTTOM = &H07				'�ؼ����,����������,������λ��	ʹ�� WM_VSCROLL ���� SB_BOTTOM ָ���Ĵ�ֱ������λ��

	LB_GETCOUNT = &H018B			'0,0			�����б�������������������򷵻�LB_ERR
	LB_GETSELCOUNT = &H0190			'0,0			�����������ڶ���ѡ���б��������ѡ�������Ŀ��������������LB_ERR
	LB_GETSELITEMS = &H0191			'����Ĵ�С,������	�����������ڶ���ѡ���б���������ѡ�е������Ŀ��λ�á�����lParamָ��һ�����������黺�������������ѡ�е��б����������wParam˵�������黺�����Ĵ�С�����������ط��ڻ������е�ѡ�����ʵ����Ŀ��������������LB_ERR
	LB_SETSEL = &H0185				'TRUE��FALSE,����	�������ڶ���ѡ���б����ʹָ�����б���ѡ�л���ѡ�����Զ��������ɼ����򡣲���lParamָ�����б������������Ϊ-1�����൱��ָ�������е������wParamΪTRUEʱѡ���б������ʹ֮��ѡ���������򷵻�LB_ERR
	LB_SETTOPINDEX = &H0197			'����,0			������ָ�����б�������Ϊ�б��ĵ�һ���ɼ���ú����Ὣ�б����������ʵ�λ�á�wParamָ�����б�����������������ɹ�������0ֵ�����򷵻�LB_ERR
End Enum

'��ȡ���ڱ����ؼ��е��ı��ַ���
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLength" (ByVal hwnd As Long) As Long
'��ȡ���ڱ����ؼ��е��ı�
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowText" (ByVal hwnd As Long, _
	ByVal lpString As String, _
	ByVal cch As Long) As Long
'���ô��ڱ����ؼ��е��ı�
'���ܸı�������Ӧ�ó����еĿؼ����ı����ݣ������Ҫ������ SengMessage ��������һ�� WM_SETTEX ��Ϣ��
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowText" (ByVal hwnd As Long, _
	ByVal lpString As String) As Long

'����ҳ
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
	CP_ISOLATIN1 = 28591	'ISO 8859-1 Latin 1; Western European (ISO)  ��ŷ����
	CP_ISOEASTEUROPE = 28592	'ISO 8859-2 Central European; Central European (ISO)  ��ŷ����
	CP_ISOTURKISH = 28593	'ISO 8859-3 Latin 3  ��ŷ���ԡ�������Ҳ���ô��ַ�����ʾ��
	CP_ISOBALTIC = 28594	'ISO 8859-4 Baltic	 ��ŷ����
	CP_ISORUSSIAN = 28595	'ISO 8859-5 Cyrillic   ˹��������
	CP_ISOARABIC = 28596	'ISO 8859-6 Arabic  ��������
	CP_ISOGREEK = 28597		'ISO 8859-7 Greek ϣ����
	CP_ISOHEBREW = 28598	'ISO 8859-8 Hebrew; Hebrew (ISO-Visual) ϣ������Ӿ�˳�򣩣�ISO 8859-8-I�� ϣ������߼�˳��
	CP_ISOTURKISH2 = 28599	'ISO 8859-9 Turkish ����Latin-1�ı�������ĸ���ߣ���������������ĸ
	CP_ISOESTONIAN = 28603	'ISO 8859-13 Estonian   ���޵�����
	CP_ISOLATIN9 = 28605	'ISO 8859-15 Latin 9 ��ŷ���ԣ�����Latin-1Ƿȱ�ķ�������ĸ�ʹ�д����������ĸ���Լ�ŷԪ��������š�
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

'���ļ�����
Private Type SUB_FILE
	FileName			As String	'�ļ��� (����·��)
	FilePath			As String	'�ļ�·�� (���ļ���)
	FileSize			As Long		'ԭ�ļ���С
	NewFileSize			As Long		'���ļ���С
	FileAdd				As Long		'�ļ����͵Ŀ�ʼ��ַ�����ļ��е��ļ�ƫ��
	Info 				As String	'�ļ�������Ϣ�������ظ���ȡ
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


'������
Sub Main()
	Dim Obj As Object,Temp As String,TempList() As String
	'���ϵͳ����
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
	'��� Scripting.Dictionary �Ƿ����
	Set Obj = CreateObject("Scripting.Dictionary")
	If Obj Is Nothing Then
		MsgBox Err.Description & " - " & "Scripting.Dictionary",vbInformation
		Exit Sub
	End If
	Set Obj = Nothing
	'��� VBScript.RegExp �Ƿ����
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


'����ز鿴�Ի�������������˽������Ϣ��
Private Function MainDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,Temp As String,IntList() As Long,TempList() As String
	Select Case Action%
	Case 1 ' �Ի��򴰿ڳ�ʼ��
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
		'ת�ݲ���ֵ
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
	Case 2 ' ��ֵ���Ļ��߰��°�ťʱ
		MainDlgFunc = True ' ��ֹ���°�ťʱ�رնԻ��򴰿�
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
			'�����ı�������
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
	'Case 3 ' �ı��������Ͽ��ı�����ʱ
	Case 6 ' ���ܼ�
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


'�ҳ������ַ������е�ֵ��ͬ����ͬ�������б�
'Mode = False ��ȡ�����ַ������е�ֵ��ͬ�������б����������б�û�ж�Ӧ��ϵ
'Mode = True ��ȡ�����ַ������е�ֵ����ͬ�������б����������б�û�ж�Ӧ��ϵ
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


'��ֺ͵������ļ�
Private Function SplitFileOld(ByVal trgFolder As String,File As FILE_PROPERTIE,DataList() As SUB_FILE, _
				ByVal Mode As Long,Optional ByVal ShowMsg As Long) As Boolean
	Dim i As Long,n As Long,FN As FILE_IMAGE,FN2 As Variant,Bytes() As Byte,Msg As String
	On Error GoTo ExitFunction
	If File.NumberOfSub = 0 Then
		ReDim DataList(0) As SUB_FILE
		Exit Function
	End If
	'���ļ�
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
	'��ȡ���ļ�ͷ
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
	'�ر��ļ�
	On Error Resume Next
	UnLoadFile(FN,0,Mode)
End Function


'������ļ��б�
Private Function SplitFile(trgFolder As String,File As FILE_PROPERTIE,DataList() As SUB_FILE,ByVal Mode As Long, _
				Optional ByVal ShowMsg As Long,Optional ByVal fType As Boolean) As Boolean
	Dim i As Long,n As Long,Dic As Object,Msg As String,Temp As String
	Dim FN As FILE_IMAGE,FN2 As Variant,Bytes() As Byte,SubFile As FILE_PROPERTIE
	On Error GoTo ExitFunction
	If File.NumberOfSub = 0 Then
		ReDim DataList(0) As SUB_FILE
		Exit Function
	End If
	'���ļ�
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
	'��ȡ���ļ�ͷ
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
				'��ȡ���ļ�
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
	'�ر��ļ�
	Set Dic = Nothing
	On Error Resume Next
	UnLoadFile(FN,0,Mode)
	If DataList(0).FilePath = "" Then Exit Function
End Function


'�������ļ��б�
Private Function WriteDataToFile(ByVal trgFolder As String,File As FILE_PROPERTIE,DataList() As SUB_FILE) As Boolean
	Dim i As Long,sb As Object
	'���ļ�
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


'�������ļ��б�
'ImportSubFile = -1 �������
'ImportSubFile = -2 ��ȡ�ɸ�ʽ
'ImportSubFile = 1 û�����ļ�
'ImportSubFile = 2 ָ���ļ����в��������ļ�
'ImportSubFile = 3 û�����ļ��б������ļ�
'ImportSubFile = 4 ����ȡ���ļ���ԭʼ�ļ��������ڵĲ�ƥ��
'ImportSubFile = 5 ����ȡ���ļ����ļ��汾�����ڵĲ�ƥ��
'ImportSubFile = 6 ����ȡ���ļ�������ID�����ڵĲ�ƥ��
'ImportSubFile = 7 ����ȡ���ļ����ļ���С�����ڵĲ�ƥ��
'ImportSubFile = 8 ���ļ��б������ļ���ʽ��ƥ��
'ImportSubFile = 9 һ�������ļ�������
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
	'��������ļ�
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
	'���������ļ�
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


'ת����������������ʱ����ַ���Ϊ����
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


'�ϲ��ļ�
'fType = False ֱ�Ӻϲ��������ļ�����ϲ�
Private Function MergeFile(trgFile As String,srcFile As FILE_PROPERTIE,DataList() As SUB_FILE,ByVal Mode As Long,ByVal fType As Boolean) As Boolean
	Dim i As Long,n As Long,k As Long,FN As FILE_IMAGE,FN2 As Variant,Bytes() As Byte
	If srcFile.NumberOfSub = 0 Then Exit Function
	On Error GoTo ExitFunction
	With srcFile.SecList(srcFile.MaxSecIndex)
		'���ļ�
		i = .lPointerToRawData + .lSizeOfRawData
		Mode = LoadFile(trgFile,FN,i,1,i,Mode)
		If Mode < -1 Then Exit Function
		FN.SizeOfFile = i
		'�ϲ����ļ�
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
	'�ر��ļ�
	On Error Resume Next
	UnLoadFile(FN,0,Mode)
End Function


'�����ִ�
'Mode = 0 ���棬= 1 ͨ���, = 2 ������ʽ
'FilterStr = 1 ���ҵ���= 0 δ�ҵ�, = -1 �������, = -2 ͨ����﷨���� = -3 ������ʽ�﷨����
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


'�������ļ����Ʒ��ϲ������ݵ�������
'FindString > -1 �ҵ����ִ��б������ţ�= -1 δ�ҵ�
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


'��ȡ���ļ��б�
'Mode = 0 ��ȡȫ�����ļ���ַ
'Mode = 1 ��ȡȫ�����ļ�����

'Mode = 2 ��ȡ���ļ�С��ԭʼ�ļ������ļ���ַ
'Mode = 3 ��ȡ���ļ�С��ԭʼ�ļ������ļ�����

'Mode = 4 ��ȡ���ļ�����ԭʼ�ļ������ļ���ַ
'Mode = 5 ��ȡ���ļ�����ԭʼ�ļ������ļ�����

'Mode = 6 ��ȡ���ļ�����ԭʼ�ļ������ļ���ַ
'Mode = 7 ��ȡ���ļ�����ԭʼ�ļ������ļ�����

'Mode = 8 ��ȡ���ļ����Ʒ��ϲ������ݵ����ļ���ַ
'Mode = 9 ��ȡ���ļ����Ʒ��ϲ������ݵ����ļ�����
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


'��ȡ�����ִ��Ĳ��ҷ�ʽ
'GetFindMode = 0 ���棬= 1 ͨ���, = 2 ������ʽ
Private Function GetFindMode(FindStr As String) As Long
	'����ͨ�����������ʽר���ַ�ʱ
	If (FindStr Like "*[$()+.^{|*?#[\]*") = False Then
		Exit Function
	End If
	'����������ʽר���ַ�ʱ
	If (FindStr Like "*[$()+.^{|\]*") = False Then
		If (FindStr Like "*\[*?#[]*") = False Then
			GetFindMode = 1
		End If
		Exit Function
	End If
	GetFindMode = 2
End Function


'ת����������Ϊ������ʽģ��
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
		'ת��ͨ���Ϊ������ʽģ��
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


'���������ʽ�Ƿ���ȷ
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


'����ִ��Ƿ����ָ���ַ�(������ʽ�Ƚ�)
'Mode = 0 ����ִ��Ƿ����ָ���ַ������ҳ�ָ���ַ���λ��
'Mode = 1 ����ִ��Ƿ�ֻ����ָ���ַ�
'Mode = 2 ����ִ��Ƿ����ָ���ַ�
'Mode = 3 ����ִ��Ƿ�ֻ������С��д��ָ���ַ�����ʱ IgnoreCase ������Ч
'Mode = 4 ����ִ��Ƿ���������ͬ���ַ���StrNum Ϊ�����ظ��ַ�����
'Mode = 5 ����ִ��Ƿ����ָ���ִ���������ƥ����ִ��ܳ��� (�ʺ��ַ���ϲ�ѯ)
'Patrn  Ϊ������ʽģ��
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


'���������ǵ����ֽ�λ�ã������طǵ����ֽڿ�ʼλ��
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


'��������ָ�������Ŀ��ֽ�λ�ã������ؿ��ֽڿ�ʼλ�ã�Bit Ϊ��С���ֽ���
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


'ת���ַ�Ϊ Long ����ֵ
Private Function StrToLong(ByVal mStr As String,Optional ByVal DefaultValue As Long) As Long
	On Error GoTo errHandle
	StrToLong = CLng(mStr)
	Exit Function
	errHandle:
	StrToLong = DefaultValue
End Function


'��ȡ4���ֽ�ֵ (32 λֵ,4���ֽ�, -2,147,483,648 �� 2,147,483,647)
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


'��ȡ2���ֽ�ֵ (16 λֵ, 2���ֽ�, -32,768 �� 32,767)
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


'��ָ����ַ��ȡһ���ֽ�(8 λֵ, 1���ֽ�)
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


'��ȡ�����ڵ��ֽ�����
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


'д���ֽ�����
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


'��ȡ�����ֽڳ���
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


'����ļ��Ƿ��ѱ��򿪻�ռ��
Private Function IsOpen(ByVal strFilePath As String,Optional ByVal Continue As Long = 2,Optional ByVal WaitTime As Double = 0.5) As Boolean
	Dim i As Long,FN As Variant
	'���Դ��ļ�
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


'��ȡ�ļ������ļ������ݽṹ��Ϣ
Private Function GetPEHeaders(ByVal strFilePath As String,File As FILE_PROPERTIE,ByVal Mode As Long) As Boolean
	Dim i As Long,FN As FILE_IMAGE,TempList() As String,Temp As String
	On Error GoTo ExitFunction
	File.FileSize = FileLen(strFilePath)
	'���ļ�
	Mode = LoadFile(strFilePath,FN,0,0,0,Mode)
	If Mode < -1 Then Exit Function
	'��ȡ���ļ�ͷ
	GetPEHeaders = GetPEHeader(FN,File,Mode)
	If GetPEHeaders = False Then GoTo ExitFunction
	'��ȡ���ļ�ͷ
	With File.SecList(File.MaxSecIndex)
		If .lSizeOfRawData = 0 Then GoTo ExitFunction
		Temp = ByteToString(GetBytes(FN,.lSizeOfRawData,.lPointerToRawData,Mode),CP_ISOLATIN1)
		TempList = GetVAListRegExp(Temp,"MZ[\x00-\xFF]{64,384}?PE\x00",.lPointerToRawData)
		If CheckArray(TempList) = False Then GoTo ExitFunction
		Dim SubFile As FILE_PROPERTIE
		File.NumberOfSub = UBound(TempList) + 1
		For i = 0 To File.NumberOfSub - 1
			'If GetPEHeader(FN,SubFile,Mode,CLng(TempList(i))) = True Then
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
		'��ʼ������
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

		'��ȡ IMAGE_DOS_HEADERS �ṹ
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

		'��ȡ IMAGE_FILE_HEADERS �ṹ
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
		'����Ƿ����ļ��ڽṹ
		If tmpFileHeader.iNumberOfSections = 0 Then GoTo ExitFunction

		'�� PE λ����ȡ IMAGE_OPTIONAL_HEADER �ṹ��32λPE��&H10B��64λPE��&H20B
		i = i + Len(tmpFileHeader)
		Select Case GetInteger(FN,i,Mode)
		Case IMAGE_NT_OPTIONAL_HDR32_MAGIC	'32λPE�ļ�
			'GetTypeValue(FN,i.tmpOptionalHeader32,Mode)
			Select Case Mode
			Case Is < 0
				Get #FN.hFile, i + 1, tmpOptionalHeader32
			Case 0
				CopyMemory tmpOptionalHeader32, FN.ImageByte(i), Len(tmpOptionalHeader32)
			Case Else
				MoveMemory tmpOptionalHeader32, FN.MappedAddress + i, Len(tmpOptionalHeader32)
			End Select

			'��ȡ�ļ��ڽṹ
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

			'��¼���ε�ַ
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

			'��¼ DataDirectory ��ַ
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
			'��¼����Ŀ¼��
			.DataDirs = 16
		Case IMAGE_NT_OPTIONAL_HDR64_MAGIC	'64λPE�ļ�
			'GetTypeValue(FN,i,tmpOptionalHeader64,Mode)
			Select Case Mode
			Case Is < 0
				Get #FN.hFile, i + 1, tmpOptionalHeader64
			Case 0
				CopyMemory tmpOptionalHeader64, FN.ImageByte(i), Len(tmpOptionalHeader64)
			Case Else
				MoveMemory tmpOptionalHeader64, FN.MappedAddress + i, Len(tmpOptionalHeader64)
			End Select

			'��ȡ�ļ��ڽṹ
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

			'��¼���ε�ַ
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

			'��¼ DataDirectory ��ַ
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
			'��¼����Ŀ¼��
			.DataDirs = 16
		Case Else
			GoTo ExitFunction
		End Select

		'��ȡ�ļ�����������š���С�����ƫ�Ƶ�ַ���ڽڵ�������
		.MaxSecIndex = tmpFileHeader.iNumberOfSections
		Call GetSectionID(File,.MinSecID,.MaxSecID,False)
		.LangType = tmpDosHeader.iLastPageBytes

		'��ȡ .NET ����ͷ�ṹ
		i = Offset
		If i = -1 Then i = File.FileType
		If GetNETHeader(FN,File,Mode,i) = True Then .LangType = NET_FILE_SIGNATURE

		'��ȡ���ؽ���Ϣ
		.SecList(.MaxSecIndex).lPointerToRawData = .SecList(.MaxSecID).lPointerToRawData + .SecList(.MaxSecID).lSizeOfRawData
		.SecList(.MaxSecIndex).lSizeOfRawData = GetFileLength(FN,Mode) - .SecList(.MaxSecIndex).lPointerToRawData
		.SecList(.MaxSecIndex).lVirtualAddress = .SecList(.MaxSecID).lVirtualAddress + .SecList(.MaxSecID).lVirtualSize
		.SecList(.MaxSecIndex).lVirtualSize = .SecList(.MaxSecIndex).lSizeOfRawData
	End With

	'��¼������ĸ���ͷ����
	If Offset = -1 Then
		DosHeader = tmpDosHeader
		FileHeader = tmpFileHeader
		OptionalHeader32 = tmpOptionalHeader32
		OptionalHeader64 = tmpOptionalHeader64
		SecHeader = tmpSecHeader
	End If

	'��ǳɹ�
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


'��ȡ .NET ����ͷ�ṹ
Private Function GetNETHeader(FN As FILE_IMAGE,File As FILE_PROPERTIE,ByVal Mode As Long,Optional ByVal Offset As Long) As Boolean
	Dim i As Long,Length As Long,dwOffset As Long
	Dim CLRHeader			As IMAGE_CLR20_HEADER
	Dim MetaDataHeader		As IMAGE_METADATA_HEADER
	On Error GoTo ExitFunction
	With File
		'ת����15������Ŀ¼����������ַתƫ�Ƶ�ַ
		dwOffset = RvaToOffset(File,.DataDirectory(14).lVirtualAddress)
		If dwOffset = 0 Then Exit Function
		'��ȡ CLR �ṹ
		'GetTypeValue(FN,Offset + dwOffset,CLRHeader,Mode)
		Select Case Mode
		Case Is < 0
			Get #FN.hFile, Offset + dwOffset + 1, CLRHeader
		Case 0
			CopyMemory CLRHeader, FN.ImageByte(Offset + dwOffset), Len(CLRHeader)
		Case Else
			MoveMemory CLRHeader, Offset + FN.MappedAddress + dwOffset, Len(CLRHeader)
		End Select
		'ת�� CLR �е� MetaData ��������ַתƫ�Ƶ�ַ
		dwOffset = RvaToOffset(File,CLRHeader.METADATA.lVirtualAddress)
		If dwOffset = 0 Then Exit Function
	End With

	'����Ƿ�Ϊ .NET �����ļ�
	With MetaDataHeader
		.lSignature = GetLong(FN,Offset + dwOffset,Mode)
		If .lSignature <> NET_FILE_SIGNATURE Then Exit Function
		'��ȡ MetaDataHeader.Version ���ֽڳ���
		.lLength = GetLong(FN,Offset + dwOffset + 12,Mode)
		'�� 4 ���ֽڶ��� MetaDataHeader.Version ���ֽڳ���
		.lLength = Alignment(.lLength,4,1)
		'��ȡ METADATA �ṹ
		.iMajorVersion = GetInteger(FN,Offset + dwOffset + 4,Mode)
		.iMinorVersion = GetInteger(FN,Offset + dwOffset + 6,Mode)
		.lExtraData = GetLong(FN,Offset + dwOffset + 8,Mode)
		.Version = GetBytes(FN,.lLength,Offset + dwOffset + 16,Mode)
		.fFlags = GetInteger(FN,Offset + dwOffset + 16 + .lLength,Mode)
		.iStreams = GetInteger(FN,Offset + dwOffset + 18 + .lLength,Mode)

		'��ȡ���������ļ�ͷ�ṹ
		If .iStreams > 0 Then
			ReDim StreamHeader(.iStreams - 1) As IMAGE_STREAM_HEADER
			dwOffset = dwOffset + 20 + .lLength
			For i = 0 To .iStreams - 1
				'lOffset ����� Metadata Root��ʵ�� RVA = .CLRHeader.MetaData.lVirtualAddress + .StreamHeader(i).lOffset
				StreamHeader(i).RWA = dwOffset + 0
				StreamHeader(i).lOffset = GetLong(FN,Offset + dwOffset + 0,Mode)
				StreamHeader(i).lSize = GetLong(FN,Offset + dwOffset + 4,Mode)
				'��ȡ .StreamHeader.rcName ���ֽڳ���
				dwOffset = dwOffset + 8
				Length = getNullByte(FN,Offset + dwOffset,Offset + dwOffset + 16,Mode,1) - dwOffset + 1
				'�� 4 ���ֽڶ��� .StreamHeader.rcName ���ֽڳ���
				Length = Alignment(Length,4,1)
				StreamHeader(i).rcName = GetBytes(FN,Length,Offset + dwOffset,Mode)
				dwOffset = dwOffset + Length
			Next i
		End If
	End With

	'ת��.NET ����ͷ�ṹ����������ַΪ�����ַ
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
	'��ǳɹ�
	GetNETHeader = True
	ExitFunction:
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


'��ȡ�ļ������ļ������ݽṹ��Ϣ
Private Function GetMacHeaders(ByVal strFilePath As String,File As FILE_PROPERTIE,ByVal Mode As Long) As Boolean
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


'ӳ���ļ�
'MapSize = 0 ���ļ���ʼʱ�Ĵ�Сӳ�䣬����ָ����Сӳ��
'ReadOnly = 0 ֻ����ʽ�������д��ʽ
'SizeOfFile = 0 ��ȡ�ļ���ʼʱ�Ĵ�С�����򲻻�ȡ
'IsPE = 0 ��һ���ļ�ӳ�䣬���� PE �ļ�ӳ��(ÿ���ڶ���)
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


'�ر�ӳ���ļ�
'SizeOfFile = 0 ���ļ�ʵ�ʴ�С���棬����ָ����С����
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


'��ȡ������Ϣ
Private Sub DisplayError(ByVal lpSource As String,ByVal dwError As Long)
    Dim GetLastDllErrMsg As String
    GetLastDllErrMsg = String$(256, 32)
    FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
    			FORMAT_MESSAGE_IGNORE_INSERTS, _
    			lpSource, dwError, 0&, _
    			GetLastDllErrMsg,Len(GetLastDllErrMsg),0&)
    Err.Raise(dwError,lpSource,Trim(GetLastDllErrMsg))
End Sub

'��ȡ�ļ���������
'Mode = 0 ���ƫ�Ƶ�ַ(���������ؽ�)
'Mode = 1 ���ƫ�Ƶ�ַ(�������ؽ�)
'Mode = 2 �����������ַ(���������ؽ�)
'Mode = 3 �����������ַ(�������ؽ�)
'�����ļ��������š�MinVal��MaxVal ֵ
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
					SkipSection = -2	'�ļ�ͷ
				ElseIf Offset > File.SecList(File.MaxSecIndex).lPointerToRawData Then
					SkipSection = -3	'�� PE �ļ�
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
					SkipSection = -2	'�ļ�ͷ
				ElseIf Offset > File.SecList(File.MaxSecIndex).lVirtualAddress Then
					SkipSection = -3	'�� PE �ļ�
				End If
			End If
		End If
	End If
End Function


'��ȡ�ڱ��������С��ȴ���С�ĵ�ַ���ڽ�������
'MinID = 0 �� MaxID = 0 ��ȡ�ڱ��������С��ַ���ڽ�������
'MinID = -1 ��ȡ�� MaxID ��С�ĵ�ַ���ڽ�������
'MaxID = -1 ��ȡ�� MinID �ڴ�ĵ�ַ���ڽ�������
'Mode = False �Ƚ�ƫ�Ƶ�ַ������Ƚ���������ַ
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


'�����ļ�
'ImageSize = 0 ���ļ��ĳ�ʼ��С�򿪣�����ָ����С��
'ReadOnly = 0 ��ֻ����ʽ�򿪣������д��ʽ��
'ImageByte = 0 ����ȡ�ֽ�����ֻ��ʼ��(���淽ʽ��ȡ�����ֽ�)������ ImageByte ָ����С��ȡ
'Mode < 0 ֱ�ӷ�ʽ��Mode = 0 ���淽ʽ��Mode > 0 ӳ�䷽ʽ
'IsPE = 0 ��һ���ļ�ӳ�䣬���� PE �ļ�ӳ��(ÿ���ڶ���)
'LoadFile = -2 ��ʧ�ܣ�����ʵ�ʴ򿪷�ʽ
'LoadedImage ���ļ����ȡ������
Private Function LoadFile(ByVal strFilePath As String,LoadedImage As FILE_IMAGE,ByVal ImageSize As Long,ByVal ReadOnly As Long, _
		ByVal ImageByte As Long,ByVal Mode As Long,Optional ByVal IsPE As Long,Optional ByVal WaitTime As Double = 0.5) As Long
	Dim i As Long
	'���Դ��ļ�
	If IsOpen(strFilePath,2,WaitTime) = True Then
		LoadFile = -2
		Exit Function
	End If
	'�����ļ�
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


'ж���ļ�
'SizeOfFile = 0 ��д�룬����ָ����Сд�룬���ڻ����ӳ�䷽ʽʱ��Ч
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


'��ȡ���ֽ��������ҵ���ƥ��������б�(������ʽ��ʽ)
'ע�⣺StartPos��EndPos ��Ϊ�� 0 ��ʼ�ĵ�ַ
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


'��ת Hex ��
Private Function ReverseHexCode(ByVal HexStr As String,ByVal Num As Long) As String
	Dim i As Long
	i = Len(HexStr)
	If i < Num Then HexStr = String$(Num - i,"0") & HexStr
	ReverseHexCode = HexStr
	For i = 1 To Num - 1 Step 2
		Mid$(ReverseHexCode,i,2) = Mid$(HexStr,Num - i,2)
	Next i
End Function


'�ֽ�ת Hex ��
'StartPos <= EndPos ��ȡ��λ����λ�� Hex ���룬�����ȡ��λ����λ�� Hex ����
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


'ת��ʮ���ƺ�ʮ������ֵΪ�ַ�
'MaxVal = 0 ��ֵ����Ӧ�еĳ��ȣ�> 0 ���ļ���С�����λ����< 0 ��ָ��λ��
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


'��������Ƿ��Ѿ���ʼ��
'����ֵ:TRUE �Ѿ���ʼ��, FALSE δ��ʼ��
Private Function CheckArrEmpty(ByRef MyArr As Variant) As Boolean
	On Error Resume Next
	If UBound(MyArr) >= 0 Then CheckArrEmpty = True
	Err.Clear
End Function


'����ִ������Ƿ�Ϊ�գ��ǿշ��� True
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


'�����ִ�����Ϊ�ִ�����Ϊ Join ����Ч��̫��
'Mode = False �� Join ������ʽ���ӣ����治�����ӷ��������������ӷ�
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


'���� PE �ļ�����
Private Function Alignment(ByVal orgValue As Long,ByVal AlignVal As Long,ByVal RoundVal As Long) As Long
	If AlignVal < 1 Then
		Alignment = orgValue
	Else
		Alignment = IIf(orgValue Mod AlignVal = 0,orgValue,AlignVal * ((orgValue \ AlignVal) + RoundVal))
	End If
End Function


'������������Ϣ
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


'�ֽ�����ת�ַ���
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
			'����Ƿ���� CP_UTF32LE, CP_UTF_32LE �����׼
			If Bytes(i + 3) = 0 Then
				CopyMemory CodePage, Bytes(i), 4
				'ת������&H7FFF��ֵΪ��ֵ
				'ChrW��������ֵ��ΧΪ Integer ��ȡֵ��Χ��-32768 �� 32767������ 32767 ʱ��Ҫ��ȥ 65536
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
			'����Ƿ���� CP_UTF32BE, CP_UTF_32BE �����׼
			If Bytes(i) = 0 Then
				CopyMemory CodePage, tmpBytes(i), 4
				'ת������&H7FFF��ֵΪ��ֵ
				'ChrW��������ֵ��ΧΪ Integer ��ȡֵ��Χ��-32768 �� 32767������ 32767 ʱ��Ҫ��ȥ 65536
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


'�ַ���ת�ֽ�����
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
			'ת����ֵΪ��ֵ
			k = AscW(Mid$(textStr,i,1)) And 65535	'����4���ֽڵ��ڴ�ռ�
			CopyMemory Bytes(n), k, 4
			n = n + 4
		Next i
		StringToByte = Bytes
	Case CP_UTF32BE, CP_UTF_32BE
		ReDim Bytes(Len(textStr) * 4 - 1) As Byte
		For i = 1 To Len(textStr)
			'ת����ֵΪ��ֵ
			k = AscW(Mid$(textStr,i,1)) And 65535	'����4���ֽڵ��ڴ�ռ�
			CopyMemory Bytes(n), k, 4
			n = n + 4
		Next i
		StringToByte = LowByte2HighByte(Bytes,4)
	Case Else
		StringToByte = UTF16ToMultiByte(textStr,CodePage)	'��ָ������ҳתUnicode�ַ�ΪANSI����
		'StringToByte = StrConv(textStr,vbFromUnicode)		'����������ҳתUnicode�ַ�ΪANSI����
	End Select
End Function


'�ַ����ֽ�����ĸ��ֽں͵��ֽڻ���
'������ UNICODE LITTLE �� UNICODE BIG �ֽ�������໥ת��
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


'ת���ֽ�����Ϊ��ֵ
'ByteOrder = False ����λ�ں�ת�����򰴸�λ��ǰת
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


'��ת�ֽ����飬��������ֵ���ֽ�����ĸ��ֽں͵��ֽڻ���
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


'���õĴ���ҳ��UTF-8:65001��GB2312:936��GB18030��54936��UTF-7��65000
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


'���ضԻ���ĳ���ؼ��е��ִ�
Private Function GetTextBoxString(ByVal hwnd As Long) As String
	Dim i As Long
	On Error GoTo errHandle
	'���Է�����������Ĵ��ںͿؼ�
	'i = SendMessageLNG(hwnd,WM_GETTEXTLENGTH,0&,0&)
	'ֻ�ܷ�������Ĵ��ںͿؼ�
	i = GetWindowTextLength(hwnd)
	If i > 0 Then
		GetTextBoxString = String$(i + 1,0)
		'���Է�����������Ĵ��ںͿؼ������ٶȺ���
		'SendMessageLNG hwnd, WM_GETTEXT, i + 1, StrPtr(GetTextBoxString)
		'GetTextBoxString = Replace$(StrConv$(GetTextBoxString,vbUnicode),vbNullChar,"")
		'ֻ�ܷ�������Ĵ��ںͿؼ������ٶȿ�ü���
		GetWindowText hwnd, GetTextBoxString, i + 1
		GetTextBoxString = Replace$(GetTextBoxString,vbNullChar,"")
	End If
	Exit Function
	errHandle:
	GetTextBoxString = ""
End Function


'���öԻ���ĳ���ؼ��е��ִ�
'Mode = False �滻��ʾ������׷����ʾ
Private Function SetTextBoxString(ByVal hwnd As Long,ByVal StrText As String,Optional ByVal Mode As Boolean) As Boolean
	On Error GoTo errHandle
	If Mode = False Then
		'���Է�����������Ĵ��ںͿؼ�
		'SetTextBoxString = SendMessageLNG(hwnd,WM_SETTEXT,0&,StrPtr(StrConv$(StrText,vbFromUnicode)))
		'ֻ�ܷ�������Ĵ��ںͿؼ������ᱣ������ȷ��ʾ��ǰ���һ�����ַ�
		SetTextBoxString = SetWindowText(hwnd,StrText)
	Else
		StrText = GetTextBoxString(hwnd) & vbCrLf & StrText
		'�����ı�������������
		SetTextBoxLength hwnd, Len(StrText), True
		'���Է�����������Ĵ��ںͿؼ������ᶪʧ��������ǰ���һ�����ַ�
		'SetTextBoxString = SendMessageLNG(hwnd,WM_SETTEXT,0&,StrPtr(StrConv$(StrText,vbFromUnicode)))
		'ֻ�ܷ�������Ĵ��ںͿؼ������ᱣ������ȷ��ʾ��ǰ���һ�����ַ�
		SetTextBoxString = SetWindowText(hwnd,StrText)
		'���ù��������ı���ײ�
		SendMessageLNG hwnd, WM_VSCROLL, SB_BOTTOM, 0&
	End If
	errHandle:
End Function


'���ñ༭�ؼ��е�����ı����ȣ�ԭ��󳤶�Ϊ30000���ַ���˫�ֽ��ַ���1����
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


'���� PSL 2015 �����ϰ汾������� Split ������ֿ��ַ���ʱ����δ��ʼ������Ĵ���
Private Function ReSplit(ByVal textStr As String,Optional ByVal Sep As String = " ",Optional ByVal Max As Integer = -1) As String()
	If textStr = "" Then
		ReDim TempList(0) As String
		ReSplit = TempList
	Else
		ReSplit = Split(textStr,Sep,Max)
	End If
End Function


'��ȡ�ļ������ͣ�PE ���� MAC ���Ƿ� PE �ļ�
Private Function GetFileFormat(ByVal FilePath As String,ByVal Mode As Long,FileType As Integer) As String
	Dim i As Long,n As Long,FN As FILE_IMAGE
	On Error GoTo ExitFunction
	FileType = 0
	'���ļ�
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


'��ȡ�ļ������ļ������ݽṹ��Ϣ
Private Function GetHeaders(ByVal strFilePath As String,File As FILE_PROPERTIE,ByVal Mode As Long,FileType As Integer) As Boolean
	Select Case GetFileFormat(File.FilePath,Mode,FileType)
	Case "PE","NET",""
		GetHeaders = GetPEHeaders(File.FilePath,File,Mode)
	Case "MAC"
		GetHeaders = GetMacHeaders(File.FilePath,File,Mode)
	End Select
End Function


'��ȡ�ļ��汾��Ϣ
Private Function GetFileInfo(ByVal strFilePath As String,File As FILE_PROPERTIE) As Boolean
	Dim i As Integer,lngBufferlen As Long,lngRc As Long,lngVerPointer As Long
	Dim bytBuffer() As Byte,strTemp As String
	Dim strBuffer As String,strLangCharset As String,strVersionInfo(7) As String
	'�ļ��Ѵ�ʱ�˳�
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


'��ȡ�ļ��Ĵ��������ʡ��޸�����
'Mode = 0 ��������
'Mode = 1 ��������
'Mode = 2 �޸�����
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


'�������ļ���
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


'ɾ���ļ��У�����ɾ�����ļ���
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


'ɾ���ļ��У������������ļ���
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


'����ļ���
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


'�鿴���ػ��ļ���Ϣ
'DisPlayFormat = False ʮ������ʾ��ֵ������ʮ��������ʾ��ֵ
Private Sub FileInfoView(File As FILE_PROPERTIE,ByVal DisPlayFormat As Boolean)
	Dim i As Long,j As Long,n As Long,Stemp As Boolean
	On Error GoTo ErrHandle
	If InStr(File.Magic,"MAC") Then Stemp = True
	'MAC64������£��޷����� 64 λ(8 ���ֽ�)����ֵ��ֻ����16������ʾ
	If File.Magic = "MAC64" Then
		If DisPlayFormat = False Then DisPlayFormat = True
	End If
	'д���ļ�������Ϣ
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
	'ÿ���ļ��ڵ�ƫ�Ƶ�ַ
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
	'���ؽڵ�ƫ�Ƶ�ַ���� PE ��ַ������
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
	'ÿ���ļ��ڵ���������ַ
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
	'���ؽڵ���������ַ���� PE ��ַ������
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
	'����Ŀ¼��ַ�������ļ���
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
	'.NET CLR ����Ŀ¼��ַ�������ļ���
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
	'.NET ����ַ�������ļ���
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
	'������
	ErrHandle:
	On Error Resume Next
	Err.Source = "NotReadFile"
	Err.Description = Err.Description & JoinStr & File.FilePath & ".xls"
	Call sysErrorMassage(Err,1)
End Sub


'��ʾ�ļ���Ϣ
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


'��ʾ�ļ���Ϣ�Ի�����
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
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
		ShowFileInfoDlgFunc = True '��ֹ���°�ť�رնԻ��򴰿�
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


'д��������ļ�
'BOM = False ��鲢д�� BOM������д�� BOM
'Mode = False ɾ���ļ�������д�룬���� File Ϊ�ļ���ʱ����
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


'��ȡ�������ļ�
'BOM = False ��鲢ȥ�� BOM��������� BOM
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


'ת�� HEX ����Ϊ�ֽ�����
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


'��ȡѡ���б����Ŀ������
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


'ѡ��ָ�����б����Ŀ
'Indexs = -1 ȫѡ������ѡ��ָ����
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


'��ȡ��������Ϣ�ַ���
Private Function GetMsgList(MsgList() As String,ByVal Language As String) As Boolean
	Dim i As Integer
	ReDim MsgList(137) As String
	On Error GoTo errHandle
	Language = LCase$(Language)
	Select Case Language
	Case "chs","0804","1004"
		MsgList(0) = "����"
		MsgList(1) = "\r\n\r\n�����޷��������У����˳���"
		MsgList(2) = "\r\n\r\nҪ�������г�����"
		MsgList(3) = "\r\n\r\n���򽫼������С�"
		MsgList(4) = "�޷����ļ�����ȷ���ļ�·�����ļ������Ƿ�����������Ե��ַ���\r\n" & _
					"ע�⣺Passolo 2015 �汾�ĺ������޷�ʶ��������������ַ����ļ�·�����ļ�����"
		MsgList(5) = "������������ϵĴ���\r\n�������: %d����������: %v\r\n" & _
					"���������� Passolo ���ԣ��򱨸����������ߡ�"
		MsgList(6) = "�����ļ���ȱ�� [%s] �ڡ�\r\n%d"
		MsgList(7) = "�����ļ���ȱ�� [%s] ֵ��\r\n%d"
		MsgList(8) = "�����ļ������ݲ���ȷ��\r\n%s"
		MsgList(9) = "�����ļ������ڣ���������ԡ�\r\n%s"
		MsgList(10) = "�����ļ��汾Ϊ %d����Ҫ�İ汾����Ϊ %v��\r\n%s"
		MsgList(11) = "����ϵͳȱ�� ""%s"" ����\r\n�������: %d����������: %v"

		MsgList(12) = "PE ���ļ������� - �汾 %v (���� %b)"
		MsgList(13) = "..."
		MsgList(14) = "��ַ(%s/%d)"
		MsgList(15) = "���ļ��б�"
		MsgList(16) = "״̬: %s"
		MsgList(17) = "����"
		MsgList(18) = "����"
		MsgList(19) = "��ȡ"
		MsgList(20) = "����"
		MsgList(21) = "�����ļ���"
		MsgList(22) = "���ļ���Ϣ"
		MsgList(23) = "���ļ���Ϣ"
		MsgList(24) = "ֱ�Ӻϲ�"
		MsgList(25) = "�ļ�����ϲ�"

		MsgList(26) = "�汾 %v (���� %b)\r\n" & _
					"OS �汾: Windows XP/2000 ������\r\n" & _
					"Passolo �汾: Passolo 5.0 ������\r\n" & _
					"��Ȩ: ������\r\n" & _
					"��ַ: http://www.hanzify.org\r\n" & _
					"����: wanfu (2018 - 2019)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(27) = "���� PE ���ļ�������"
		MsgList(28) = "��ִ���ļ� (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|�����ļ� (*.*)|*.*||"
		MsgList(29) = "ѡ���ļ�"
		MsgList(30) = "Ӣ��;��������;��������"
		MsgList(31) = "enu;chs;cht"
		MsgList(32) = "������ȡ�ļ�..."
		MsgList(33) = "����ȡ %s �����ļ���"
		MsgList(34) = "����ȡ %s �����ļ����ѱ��浽 %d �ļ��С�"
		MsgList(35) = "���ļ�Ϊ�� PE �ļ���"
		MsgList(36) = "���ܽ�ԭ�ļ���ΪĿ���ļ���"
		MsgList(37) = "���ںϲ��ļ�..."
		MsgList(38) = "�ϲ��ɹ���"
		MsgList(39) = "�ϲ�ʧ�ܣ�ȱ�����ļ���"
		MsgList(40) = "ѡ���ͷ����ļ����ļ���"
		MsgList(41) = "ѡ�����ļ����ڵ��ļ���"
		MsgList(42) = "ȷ��"
		MsgList(43) = "��ԭʼС;��ԭʼ��ͬ;��ԭʼ��;���ļ���"
		MsgList(44) = "���ڵ����ļ�..."
		MsgList(45) = "������ %s �����ļ���"

		MsgList(46) = "============ �ļ���Ϣ ============\r\n"
		MsgList(47) = "�ļ����ƣ�\t%s"
		MsgList(48) = "�ļ�·����\t%s"
		MsgList(49) = "�ļ�˵����\t%s"
		MsgList(50) = "�ļ��汾��\t%s"
		MsgList(51) = "��Ʒ���ƣ�\t%s"
		MsgList(52) = "��Ʒ�汾��\t%s"
		MsgList(53) = "��Ȩ���У�\t%s"
		MsgList(54) = "�ļ���С��\t%s �ֽ�"
		MsgList(55) = "�������ڣ�\t%s"
		MsgList(56) = "�޸����ڣ�\t%s"
		MsgList(57) = "����ԣ�\t%s"
		MsgList(58) = "�� �� �̣�\t%s"
		MsgList(59) = "ԭʼ�ļ�����\t%s"
		MsgList(60) = "�ڲ��ļ�����\t%s"
		MsgList(61) = "�ļ����ͣ�\t%s"
		MsgList(62) = "ӳ���ַ��\t%s"
		MsgList(63) = "������Ϣ��"
		MsgList(64) = "��ַ���\t������\t��ʼ��ַ\t������ַ\t�ֽڴ�С"
		MsgList(65) = "%s!1!\t%s!2!\t%s!4!\t%s!5!\t%s!6!"
		MsgList(66) = "================================="
		MsgList(67) = "�ļ�ƫ�Ƶ�ַ"
		MsgList(68) = "��������ַ"
		MsgList(69) = "����"
		MsgList(70) = "����"
		MsgList(71) = "δ֪"
		MsgList(72) = "������"
		MsgList(73) = "����Ŀ¼��Ϣ (�ļ�ƫ�Ƶ�ַ)��"
		MsgList(74) = "Ŀ¼����\t��������\t��ʼ��ַ\t������ַ\t�ֽڴ�С"
		MsgList(75) = "����Ŀ¼"
		MsgList(76) = "����Ŀ¼"
		MsgList(77) = "��ԴĿ¼"
		MsgList(78) = "�쳣Ŀ¼"
		MsgList(79) = "��ȫĿ¼"
		MsgList(80) = "��ַ�ض�λ��"
		MsgList(81) = "����Ŀ¼"
		MsgList(82) = "��ȨĿ¼"
		MsgList(83) = "����ֵ(GP RVA)"
		MsgList(84) = "�̱߳��ش洢��"
		MsgList(85) = "��������Ŀ¼"
		MsgList(86) = "�󶨵���Ŀ¼"
		MsgList(87) = "�����ַ��"
		MsgList(88) = "�ӳټ��ص����"
		MsgList(89) = "COM ���п��־"
		MsgList(90) = "����Ŀ¼"
		MsgList(91) = "�쳣"
		MsgList(92) = "������"
		MsgList(93) = ".NET CLR ����Ŀ¼��Ϣ (�ļ�ƫ�Ƶ�ַ)��"
		MsgList(94) = "Ŀ¼����\t��������\t��ʼ��ַ\t������ַ\t�ֽڴ�С"
		MsgList(95) = "Ԫ����(MetaData)"
		MsgList(96) = "�й���Դ"
		MsgList(97) = "ǿ����ǩ��"
		MsgList(98) = "��������"
		MsgList(99) = "�����(V-��)"
		MsgList(100) = "��ת������ַ��"
		MsgList(101) = "�йܱ���ӳ��ͷ"
		MsgList(102) = ".NET MetaData ����Ϣ (�ļ�ƫ�Ƶ�ַ)��"
		MsgList(103) = "������\t��������\t��ʼ��ַ\t������ַ\t�ֽڴ�С"
		MsgList(104) = "�� PE �ļ�"
		MsgList(105) = "��PE(%s)"
		MsgList(106) = "��ַ���\t����\t����\t\t��ʼ��ַ\t������ַ\t�ֽڴ�С"
		MsgList(107) = "%s!1!\t%s!2!\t%s!3!\t\t%s!4!\t%s!5!\t%s!6!"

		MsgList(108) = "���ڻ�ȡ %s �ļ���Ϣ..."
		MsgList(109) = "��Ϣ - %s"
		MsgList(110) = "��һ��"
		MsgList(111) = "��һ��"
		MsgList(112) = "�ҵ� %s �����ļ���"

		MsgList(113) = "��Ϣ"
		MsgList(114) = "�ļ�û�����ļ���"
		MsgList(115) = "Ҫ������ļ�����û���ļ���"
		MsgList(116) = "Ҫ������ļ�����û���ҵ� %s �����ļ����޷����롣"
		MsgList(117) = "���ļ���ԭʼ�ļ����Ʋ������޷����롣"
		MsgList(118) = "���ļ���ԭʼ�ļ��汾�������޷����롣"
		MsgList(119) = "���ļ���ԭʼ�ļ����Բ������޷����롣"
		MsgList(120) = "���ļ���ԭʼ�ļ���С�������޷����롣"
		MsgList(121) = "���ļ���ԭʼ�ļ������Ѹ��ģ��Ƿ������\r\n�����Ѹ��ģ�˵���ļ��ѱ��޸Ĺ����޸Ĺ����ļ����ܲ����á�"
		MsgList(122) = "%s �����ļ���ʽ���ԣ��޷����롣"
		MsgList(123) = "һ�������ļ������ڣ��޷����롣\r\n�ϲ�ʱ������һ�����ļ���"

		MsgList(124) = "ȫѡ"
		MsgList(125) = "����(&F3)"
		MsgList(126) = "������ʾ"
		MsgList(127) = "����"
		MsgList(128) = "����"
		MsgList(129) = "ȫ����ʾ"
		MsgList(130) = "������Ҫ���ҵ����ݡ�\r\n- ��ʹ�� F3 ��ݼ����ô˹��ܡ��������ݲ�Ϊ��ʱ������ʾ�öԻ���\r\n- ��������֧�ֳ��桢ͨ�����������ʽ���Զ��жϡ�"
		MsgList(131) = "������Ҫ���˵����ݡ�\r\nע�⣺��������֧�ֳ��桢ͨ�����������ʽ���Զ��жϡ�"
		MsgList(132) = "���ҵ� %s һ�"
		MsgList(133) = "δ�ҵ� %s��"
		MsgList(134) = "���������ж�Ϊͨ��������﷨����"
		MsgList(135) = "���������ж�Ϊ������ʽ�����﷨����"
		MsgList(136) = "���������ж�Ϊͨ��������﷨����"
		MsgList(137) = "���������ж�Ϊ������ʽ�����﷨����"
	Case "cht","0404","0c04","1404"
		MsgList(0) = "���~"
		MsgList(1) = "\r\n\r\n�{���L�k�~�����A�N�����C"
		MsgList(2) = "\r\n\r\n�n�~�����{���ܡH"
		MsgList(3) = "\r\n\r\n�{���N�~�����C"
		MsgList(4) = "�L�k�}���ɮסA�нT�{�ɮ׸��|�M�ɮצW���O�_�]�t�Ȭw�y�����r���C\r\n" & _
					"�`�N�GPassolo 2015 ���������������L�k���ѥ]�t�Ȭw�y���r�����ɮ׸��|�M�ɮצW�C"
		MsgList(5) = "�o�͵{���]�p�W�����~�C\r\n���~�N�X: %d�A���~�y�z: %v\r\n" & _
					"�Э��s�Ұ� Passolo �A�աA�γ��i���n��}�o�̡C"
		MsgList(6) = "�U�C�ɮפ��ʤ� [%s] �`�C\r\n%d"
		MsgList(7) = "�U�C�ɮפ��ʤ� [%s] �ȡC\r\n%d"
		MsgList(8) = "�U�C�ɮת����e�����T�C\r\n%s"
		MsgList(9) = "�U�C�ɮפ��s�b�I���ˬd��A�աC\r\n%s"
		MsgList(10) = "�U�C�ɮת����� %d�A�ݭn�������ܤ֬� %v�C\r\n%s"
		MsgList(11) = "�z���t�ίʤ� ""%s"" �A�ȡC\r\n���~�N�X: %d�A���~�y�z: %v"

		MsgList(12) = "PE �l�ɮ׺޲z�� - ���� %v (�c�� %b)"
		MsgList(13) = "..."
		MsgList(14) = "��}(%s/%d)"
		MsgList(15) = "�l�ɮײM��"
		MsgList(16) = "���A: %s"
		MsgList(17) = "����"
		MsgList(18) = "�y��"
		MsgList(19) = "�^��"
		MsgList(20) = "�פJ"
		MsgList(21) = "�ƻs�ɮצW"
		MsgList(22) = "�D�ɮװT��"
		MsgList(23) = "�l�ɮװT��"
		MsgList(24) = "�����X��"
		MsgList(25) = "�ɮ׹���X��"

		MsgList(26) = "���� %v (�c�� %b)\r\n" & _
					"OS ����: Windows XP/2000 �ΥH�W\r\n" & _
					"Passolo ����: Passolo 5.0 �ΥH�W\r\n" & _
					"���v: �K�O�n��\r\n" & _
					"���}: http://www.hanzify.org\r\n" & _
					"�@��: wanfu (2018 - 2019)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(27) = "���� PE �l�ɮ׺޲z��"
		MsgList(28) = "�i�����ɮ� (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|�Ҧ��ɮ� (*.*)|*.*||"
		MsgList(29) = "����ɮ�"
		MsgList(30) = "�^�y;²�餤��;���餤��"
		MsgList(31) = "enu;chs;cht"
		MsgList(32) = "���b�^���ɮ�..."
		MsgList(33) = "�@�^�� %s �Ӥl�ɮסC"
		MsgList(34) = "�@�^�� %s �Ӥl�ɮסA�w�x�s�� %d ��Ƨ��C"
		MsgList(35) = "���ɮ׬��D PE �ɮסC"
		MsgList(36) = "����N���ɮק@���ؼ��ɮסC"
		MsgList(37) = "���b�X���ɮ�..."
		MsgList(38) = "�X�֦��\�I"
		MsgList(39) = "�X�֥��ѡI�ʤ֤l�ɮסC"
		MsgList(40) = "�������l�ɮת���Ƨ�"
		MsgList(41) = "����l�ɮשҦb����Ƨ�"
		MsgList(42) = "�T�{"
		MsgList(43) = "���l�p;�M��l�ۦP;���l�j;�l�ɮצW"
		MsgList(44) = "���b�פJ�ɮ�..."
		MsgList(45) = "�@�פJ %s �Ӥl�ɮסC"

		MsgList(46) = "============ �ɮװT�� ============\r\n"
		MsgList(47) = "�ɮצW�١G\t%s"
		MsgList(48) = "�ɮ׸��|�G\t%s"
		MsgList(49) = "�ɮ׻����G\t%s"
		MsgList(50) = "�ɮת����G\t%s"
		MsgList(51) = "���~�W�١G\t%s"
		MsgList(52) = "���~�����G\t%s"
		MsgList(53) = "���v�Ҧ��G\t%s"
		MsgList(54) = "�ɮפj�p�G\t%s �줸��"
		MsgList(55) = "�إߤ���G\t%s"
		MsgList(56) = "�ק����G\t%s"
		MsgList(57) = "�y�@�@���G\t%s"
		MsgList(58) = "�} �o �ӡG\t%s"
		MsgList(59) = "��l�ɮצW�G\t%s"
		MsgList(60) = "�����ɮצW�G\t%s"
		MsgList(61) = "�ɮ������G\t%s"
		MsgList(62) = "�M����}�G\t%s"
		MsgList(63) = "�Ϭq�T���G"
		MsgList(64) = "��}���O\t�Ϭq�W\t�}�l��}\t������}\t�줸�դj�p"
		MsgList(65) = "%s!1!\t%s!2!\t%s!4!\t%s!5!\t%s!6!"
		MsgList(66) = "================================="
		MsgList(67) = "�ɮװ�����}"
		MsgList(68) = "�۹������}"
		MsgList(69) = "���N"
		MsgList(70) = "����"
		MsgList(71) = "����"
		MsgList(72) = "���i��"
		MsgList(73) = "��ƥؿ��T�� (�ɮװ�����})�G"
		MsgList(74) = "�ؿ��W��\t�Ҧb�Ϭq\t�}�l��}\t������}\t�줸�դj�p"
		MsgList(75) = "�ץX�ؿ�"
		MsgList(76) = "�פJ�ؿ�"
		MsgList(77) = "�귽�ؿ�"
		MsgList(78) = "���`�ؿ�"
		MsgList(79) = "�w���ؿ�"
		MsgList(80) = "��}���w���"
		MsgList(81) = "�E�_�ؿ�"
		MsgList(82) = "���v�ؿ�"
		MsgList(83) = "������(GP RVA)"
		MsgList(84) = "����������s�x��"
		MsgList(85) = "���J�]�w�ؿ�"
		MsgList(86) = "�j�w�פJ�ؿ�"
		MsgList(87) = "�פJ��}��"
		MsgList(88) = "������J�פJ��"
		MsgList(89) = "COM ����w�X��"
		MsgList(90) = "�O�d�ؿ�"
		MsgList(91) = "���`"
		MsgList(92) = "���s�b"
		MsgList(93) = ".NET CLR ��ƥؿ��T�� (�ɮװ�����})�G"
		MsgList(94) = "�ؿ��W��\t�Ҧb�Ϭq\t�}�l��}\t������}\t�줸�դj�p"
		MsgList(95) = "���~���(MetaData)"
		MsgList(96) = "���޸귽"
		MsgList(97) = "�j�W��ñ�W"
		MsgList(98) = "�N�X�޲z��"
		MsgList(99) = "������(V-��)"
		MsgList(100) = "���D�ץX��}��"
		MsgList(101) = "���ޥ����M���Y"
		MsgList(102) = ".NET MetaData ��Ƭy�T�� (�ɮװ�����})�G"
		MsgList(103) = "��Ƭy�W��\t�Ҧb�Ϭq\t�}�l��}\t������}\t�줸�դj�p"
		MsgList(104) = "�D PE �ɮ�"
		MsgList(105) = "�lPE(%s)"
		MsgList(106) = "��}���O\t�q�W\t�`�W\t\t�}�l��}\t������}\t�줸�դj�p"
		MsgList(107) = "%s!1!\t%s!2!\t%s!3!\t\t%s!4!\t%s!5!\t%s!6!"

		MsgList(108) = "���b��� %s �ɮװT��..."
		MsgList(109) = "�T�� - %s"
		MsgList(110) = "�W�@��"
		MsgList(111) = "�U�@��"
		MsgList(112) = "��� %s �Ӥl�ɮסC"

		MsgList(113) = "�T��"
		MsgList(114) = "�ɮרS���l�ɮסC"
		MsgList(115) = "�n�פJ����Ƨ����S���ɮסC"
		MsgList(116) = "�n�פJ����Ƨ����S����� %s ����ɮסA�L�k�פJ�C"
		MsgList(117) = "�l�ɮת���l�ɮצW�٤��šA�L�k�פJ�C"
		MsgList(118) = "�l�ɮת���l�ɮת������šA�L�k�פJ�C"
		MsgList(119) = "�l�ɮת���l�ɮ׻y�����šA�L�k�פJ�C"
		MsgList(120) = "�l�ɮת���l�ɮפj�p���šA�L�k�פJ�C"
		MsgList(121) = "�l�ɮת���l�ɮפ���w�ܧ�A�O�_�~��H\r\n����w�ܧ�A�����ɮפw�Q�ק�L�A�ק�L�� �ɥi�ण�A�ΡC"
		MsgList(122) = "%s ����ɮ׮榡����A�L�k�פJ�C"
		MsgList(123) = "�@�����l�ɮפ��s�b�A�L�k�פJ�C\r\n�X�֮ɤ���֤@�Ӥl�ɮסC"

		MsgList(124) = "����"
		MsgList(125) = "�j�M(&F3)"
		MsgList(126) = "�L�o���"
		MsgList(127) = "�j�M"
		MsgList(128) = "�L�o"
		MsgList(129) = "�������"
		MsgList(130) = "�п�J�n�j�M�����e�C\r\n- �i�ϥ� F3 �ֳt��եΦ��\��C�j�M���e�����ŮɡA����ܸӹ�ܤ���C\r\n- �j�M���e�䴩�`�W�B�U�Φr���M�W�h�B�⦡�æ۰ʧP�_�C"
		MsgList(131) = "�п�J�n�L�o�����e�C\r\n�`�N�G�L�o���e�䴩�`�W�B�U�Φr���M�W�h�B�⦡�æ۰ʧP�_�C"
		MsgList(132) = "�ȧ�� %s �@���C"
		MsgList(133) = "����� %s�C"
		MsgList(134) = "�j�M���e�P�_���U�Φr���A���y�k���~�C"
		MsgList(135) = "�j�M���e�P�_���W�h�B�⦡�A���y�k���~�C"
		MsgList(136) = "�L�o���e�P�_���U�Φr���A���y�k���~�C"
		MsgList(137) = "�L�o���e�P�_���W�h�B�⦡�A���y�k���~�C"
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
