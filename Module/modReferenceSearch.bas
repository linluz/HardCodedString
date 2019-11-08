'' Reference Search for Passolo
'' (c) 2015 - 2019 by wanfu (Last modified on 2019.11.08)

'' Command Line Format: Command <FilePath> <-add:Address> <-lng:Hex Language Code>
'' Command: Name of this Macros file
'' FilePath: Full path of PE file
'' Address: Offset of Dec or Hex, if Hex, use like "0xFFFF" format
'' Hex Language Code: Display UI Language. Supports EngLish, Chinese Simplified and Chinese Traditional only. For sample: 0804,1004,0404,0C04,1404.
'' Return: No
'' For example: modRefernceSearch,"d:\my folder\my file.exe" -add:35148 -lng:0804

Option Explicit

Private Const Version = "2019.11.08"
Private Const Build = "191108"
Private Const JoinStr = vbFormFeed  'vbBack
Private Const TextJoinStr = vbCrLf
Private Const RefJoinStr = "|"
Private Const LoadMode = 0&
Private Const AppName = "Reference Search"
Private Const RefFrontChar = "[\x0F\x4C\x48\x8B\xF2][\x05\x10\x0F\x8B\x8D\xB7][\x00-\xFF]{3}"

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

'�����ı�
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

'����ҳ
Private Enum KnownCodePage
	CP_CHINA = 936			'ANSI/OEM Simplified Chinese (PRC, Singapore); Chinese Simplified (GB2312)
	CP_TAIWAN = 950			'ANSI/OEM Traditional Chinese (Taiwan; Hong Kong SAR, PRC); Chinese Traditional (Big5)
	CP_UNICODELITTLE = 1200	'Unicode UTF-16, little endian byte order (BMP of ISO 10646); available only to managed applications
	CP_UNICODEBIG = 1201	'Unicode UTF-16, big endian byte order; available only to managed applications
	CP_WESTEUROPE = 1252	'ANSI Latin 1; Western European (Windows)
	CP_ISOLATIN1 = 28591	'ISO 8859-1 Latin 1; Western European (ISO)  ��ŷ����
	CP_UTF_32LE = 12000  	'Unicode UTF-32, little endian byte order; available only to managed applications
	CP_UTF_32BE = 12001		'Unicode UTF-32, big endian byte order; available only to managed applications
	CP_UTF32LE = 65005  	'Unicode (UTF-32 LE)
	CP_UTF32BE = 65006		'Unicode (UTF-32 Big-Endian)
End Enum

Private Type REFERENCE_PROPERTIE
	sCode			As String	'���ô���
	lAddress		As Long		'���õ�ַ�б�
	inSecID			As Integer	'�ִ����ڽڵ�������
End Type

Private Type STRING_SUB_PROPERTIE
	lStartAddress	As Long		'�ִ��Ŀ�ʼ��ַ
	inSectionID		As Integer	'�ִ����ڽڵ�������
	inSubSecID			As Integer	'�ִ����ڽڵ��ӽ�������
	lReferenceNum	As Long		'���ô���
	GetRefState		As Integer	'��ȡ�ִ������б��״̬��0 = δ��ȡ��1 = �ѻ�ȡ
	Reference()		As REFERENCE_PROPERTIE
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

Private MsgList() As String,SetList() As String,RegExp As Object
Private File As FILE_PROPERTIE,Data As STRING_SUB_PROPERTIE


'������
Sub Main()
	Dim Obj As Object,Temp As String
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
	'��� VBScript.RegExp �Ƿ����
	On Error Resume Next
	Set RegExp = CreateObject("VBScript.RegExp")
	If RegExp Is Nothing Then
		MsgBox(Err.Description & " - " & "VBScript.RegExp",vbInformation)
		Exit Sub
	End If
	RegExp.MultiLine = True
	On Error GoTo SysErrorMsg
	SetList = SplitArgument(Command$,3)
	If SetList(2) <> "" Then
		If StrToLong(SetList(2),3) = 3 Then Temp = ReSplit(SetList(2),";")(0)
	End If
	If UCase(Temp) <> Right$("0" & Hex$(PSL.Option(pslOptionSystemLanguage)),4) Then
		Temp = Right$("0" & Hex$(PSL.Option(pslOptionSystemLanguage)),4)
	End If
	If GetMsgList(MsgList,Temp) = False Then GoTo SysErrorMsg
	'ֹͣ��ť����ʱ����ʾ�Ի����ϵ���Ϣ
	Begin Dialog UserDialog 600,294,Replace$(Replace$(MsgList(12),"%v",Version),"%b",Build),.MainDlgFunc ' %GRID:10,7,1,1
		TextBox 0,0,0,21,.SuppValueBox
		TextBox 10,7,440,21,.FilePathBox
		PushButton 450,7,30,21,MsgList(13),.FilePathButton
		OptionGroup .ValTypeGroup
			OptionButton 10,35,60,21,MsgList(14),.DecButton
			OptionButton 80,35,60,21,MsgList(15),.HexButton
		CheckBox 80,35,60,21,"",.ValTypeCheckBox
		DropListBox 150,35,150,21,SetList(),.SecNameList
		Text 310,38,50,14,MsgList(16),.FileBitText
		DropListBox 370,35,110,21,SetList(),.FileBitList
		PushButton 490,7,100,21,MsgList(17),.AboutButton
		PushButton 490,28,100,21,MsgList(63),.LangButton
		Text 10,66,50,14,MsgList(18),.RVAText
		TextBox 60,63,100,21,.RVABox
		Text 170,66,50,14,MsgList(19),.OffsetText
		TextBox 230,63,100,21,.OffsetBox
		PushButton 350,63,130,21,MsgList(20),.VACodeButton
		PushButton 490,49,100,21,MsgList(21),.SearchButton
		Text 10,98,470,14,Replace$(MsgList(22),"%s","0"),.ShowText
		TextBox 10,119,580,168,.ShowTextBox,1
		PushButton 490,70,100,21,MsgList(23),.CopyButton
		PushButton 490,91,100,21,MsgList(66),.FileInfoButton
		CancelButton 330,91,70,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then Exit Sub
	Exit Sub
	SysErrorMsg:
	If Err.Source <> "ExitSub" Then Call sysErrorMassage(Err,0)
End Sub


'����ز鿴�Ի�������������˽������Ϣ��
Private Function MainDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,Mode As Long,FN As FILE_IMAGE,Temp As String,TempList() As String
	Select Case Action%
	Case 1 ' �Ի��򴰿ڳ�ʼ��
		DlgText "SuppValueBox",CStr$(SuppValue)
		DlgVisible "SuppValueBox",False
		DlgVisible "CancelButton",False
		DlgVisible "ValTypeCheckBox",False
		DlgEnable "FilePathBox",False
		DlgEnable "CopyButton",False
		DlgEnable "VACodeButton",False
		DlgValue "ValTypeCheckBox",DlgValue("ValTypeGroup")
		TempList = ReSplit(MsgList(24),";")
		DlgListBoxArray "FileBitList",TempList()
		'ת�ݲ���ֵ
		File.FilePath = SetList(0)
		If InStr(SetList(1),"0x") Then
			Data.lStartAddress = Val("&H" & SetList(1))
		Else
			Data.lStartAddress = StrToLong(SetList(1))
		End If
		If Dir$(File.FilePath) = "" Then File.FilePath = ""
		Temp = File.FilePath
		If Len(Temp) > 50 Then
			Temp = Left$(Temp,InStr(Temp,"\")) & "..." & Right(Temp,50 - Len(Left$(Temp,InStr(Temp,"\"))))
		End If
		DlgText "FilePathBox",Temp
		If File.FilePath = "" Then
			Data.lStartAddress = -1
			DlgValue "FileBitList",0
			ReDim TempList(0) As String
			DlgListBoxArray "SecNameList",TempList()
			DlgValue "SecNameList",0
			DlgEnable "SearchButton",False
			DlgEnable "FileInfoButton",False
		Else
			Call GetFileInfo(File.FilePath,File)
			If GetHeaders(File.FilePath,File,LoadMode,File.FileType) = False Then
				DlgValue "FileBitList",1
				'��ʾ��ַ
				If Data.lStartAddress > -1 Then
					DlgText "OffsetBox",CStr$(Data.lStartAddress)
					DlgText "RVABox",CStr$(Data.lStartAddress)
				End If
				'�����ظ���������
				ChangSectionNames File,MsgList(26),MsgList(25)
				'��ȡ�ļ��������б�
				TempList = getSectionNameList(File.SecList)
				i = UBound(TempList)
				ReDim Preserve TempList(i + 1) As String
				TempList(i + 1) = MsgList(27)
			Else
				Select Case File.Magic
				Case ""
					DlgValue "FileBitList",1
				Case "PE32"
					DlgValue "FileBitList",2
				Case "PE64"
					DlgValue "FileBitList",3
				Case "MAC32"
					DlgValue "FileBitList",4
				Case "MAC64"
					DlgValue "FileBitList",5
				End Select
				'��ʾ��ַ
				If Data.lStartAddress > -1 Then
					DlgText "OffsetBox",CStr$(Data.lStartAddress)
					DlgText "RVABox",CStr$(OffsetToRva(File,Data.lStartAddress))
				End If
				'�����ظ���������
				ChangSectionNames File,MsgList(26),MsgList(25)
				'��ȡ�ļ��������б�
				TempList = getSectionNameList(File.SecList)
				i = UBound(TempList)
				ReDim Preserve TempList(i + 3) As String
				TempList(i + 1) = MsgList(27)
				TempList(i + 2) = MsgList(28)
				TempList(i + 3) = MsgList(29)
			End If
			DlgListBoxArray "SecNameList",TempList()
			Data.inSectionID = SkipSection(File,Data.lStartAddress,0,0,1)
			If Data.inSectionID > -1 Then
				Data.inSubSecID = SkipSubSection(File.SecList(Data.inSectionID),Data.lStartAddress,0,0)
				DlgValue "SecNameList",Data.inSectionID
				DlgEnable "SearchButton",IIf(File.Magic = "",False,True)
			ElseIf File.SecList(File.MaxSecIndex).lSizeOfRawData > 0 Then
				Data.inSubSecID = -1
				DlgValue "SecNameList",File.MaxSecIndex - Data.inSectionID
				DlgEnable "SearchButton",False
			Else
				Data.inSubSecID = -1
				DlgValue "SecNameList",File.MaxSecIndex - Data.inSectionID - 1
				DlgEnable "SearchButton",False
			End If
			DlgEnable "FileInfoButton",True
		End If
	Case 2 ' ��ֵ���Ļ��߰��°�ťʱ
		MainDlgFunc = True ' ��ֹ���°�ťʱ�رնԻ��򴰿�
		Select Case DlgItem$
		Case "CancelButton"
			MainDlgFunc = False
		Case "AboutButton"
			MsgBox Replace$(Replace$(MsgList(31),"%v",Version),"%b",Build),vbOkOnly+vbInformation,MsgList(32)
		Case "FilePathButton"
			If PSL.SelectFile(Temp,True,MsgList(33),MsgList(34)) = False Then Exit Function
			If File.FilePath = Temp Then Exit Function
			If IsOpen(Temp,2,0) = True Then Exit Function
			File.FilePath = Temp
			If Len(Temp) > 50 Then
				Temp = Left$(Temp,InStr(Temp,"\")) & "..." & Right(Temp,50 - Len(Left$(Temp,InStr(Temp,"\"))))
			End If
			DlgText "FilePathBox",Temp
			Data.lStartAddress = 0
			DlgText "OffsetBox",""
			DlgText "RVABox",""
			DlgText "ShowTextBox",""
			DlgText "ShowText",Replace$(MsgList(22),"%s","0")
			DlgEnable "CopyButton",False
			DlgEnable "VACodeButton",False
			File.Info = ""
			Call GetFileInfo(File.FilePath,File)
			If GetHeaders(File.FilePath,File,LoadMode,File.FileType) = False Then
				DlgValue "FileBitList",1
				'�����ظ���������
				ChangSectionNames File,MsgList(26),MsgList(25)
				'��ȡ�ļ��������б�
				TempList = getSectionNameList(File.SecList)
				i = UBound(TempList)
				ReDim Preserve TempList(i + 1) As String
				TempList(i + 1) = MsgList(27)
			Else
				'��ȡ�ļ�����
				Select Case File.Magic
				Case ""
					DlgValue "FileBitList",1
				Case "PE32"
					DlgValue "FileBitList",2
				Case "PE64"
					DlgValue "FileBitList",3
				Case "MAC32"
					DlgValue "FileBitList",4
				Case "MAC64"
					DlgValue "FileBitList",5
				End Select
				'�����ظ���������
				ChangSectionNames File,MsgList(26),MsgList(25)
				'��ȡ�ļ��������б�
				TempList = getSectionNameList(File.SecList)
				i = UBound(TempList)
				ReDim Preserve TempList(i + 3) As String
				TempList(i + 1) = MsgList(27)
				TempList(i + 2) = MsgList(28)
				TempList(i + 3) = MsgList(29)
			End If
			DlgListBoxArray "SecNameList",TempList()
			Data.inSectionID = SkipSection(File,Data.lStartAddress,0,0,1)
			If Data.inSectionID > -1 Then
				Data.inSubSecID = SkipSubSection(File.SecList(Data.inSectionID),Data.lStartAddress,0,0)
				DlgValue "SecNameList",Data.inSectionID
				DlgEnable "SearchButton",IIf(File.Magic = "",False,True)
			ElseIf File.SecList(File.MaxSecIndex).lSizeOfRawData > 0 Then
				Data.inSubSecID = -1
				DlgValue "SecNameList",File.MaxSecIndex - Data.inSectionID
				DlgEnable "SearchButton",False
			Else
				Data.inSubSecID = -1
				DlgValue "SecNameList",File.MaxSecIndex - Data.inSectionID - 1
				DlgEnable "SearchButton",False
			End If
			DlgEnable "FileInfoButton",True
		Case "ValTypeGroup"
			If DlgValue("ValTypeGroup") = DlgValue("ValTypeCheckBox") Then Exit Function
			DlgValue "ValTypeCheckBox",DlgValue("ValTypeGroup")
			If DlgValue("ValTypeGroup") = 0 Then
				If DlgText("OffsetBox") <> "" Then
					DlgText "OffsetBox",CStr$(Val("&H" & DlgText("OffsetBox")))
				End If
				If DlgText("RVABox") <> "" Then
					DlgText "RVABox",CStr$(Val("&H" & DlgText("RVABox")))
				End If
			Else
				If DlgText("OffsetBox") <> "" Then
					DlgText "OffsetBox",FormatHexStr(Hex$(StrToLong(DlgText("OffsetBox"))),2)
				End If
				If DlgText("RVABox") <> "" Then
					DlgText "RVABox",FormatHexStr(Hex$(StrToLong(DlgText("RVABox"))),2)
				End If
			End If
		Case "SecNameList"
			'�����ظ���������
			ChangSectionNames File,MsgList(26),MsgList(25)
			'��ȡ�ļ��������б�
			TempList = getSectionNameList(File.SecList)
			i = UBound(TempList)
			If File.Magic = "" Then
				ReDim Preserve TempList(i + 1) As String
				TempList(i + 1) = MsgList(27)
			Else
				ReDim Preserve TempList(i + 3) As String
				TempList(i + 1) = MsgList(27)
				TempList(i + 2) = MsgList(28)
				TempList(i + 3) = MsgList(29)
			End If
			DlgListBoxArray "SecNameList",TempList()
			Data.inSectionID = SkipSection(File,Data.lStartAddress,0,0,1)
			If Data.inSectionID > -1 Then
				DlgValue "SecNameList",Data.inSectionID
				DlgEnable "SearchButton",IIf(File.Magic = "",False,True)
			ElseIf File.SecList(File.MaxSecIndex).lSizeOfRawData > 0 Then
				DlgValue "SecNameList",File.MaxSecIndex - Data.inSectionID
				DlgEnable "SearchButton",False
			Else
				DlgValue "SecNameList",File.MaxSecIndex - Data.inSectionID - 1
				DlgEnable "SearchButton",False
			End If
		Case "FileBitList"
			'��ȡ�ļ�����
			Select Case File.Magic
			Case ""
				DlgValue "FileBitList",1
			Case "PE32"
				DlgValue "FileBitList",2
			Case "PE64"
				DlgValue "FileBitList",3
			Case "MAC32"
				DlgValue "FileBitList",4
			Case "MAC64"
				DlgValue "FileBitList",5
			End Select
		Case "VACodeButton"
			If DlgText("FilePathBox") = "" Then Exit Function
			If DlgText("OffsetBox") = "" Or File.Magic = "" Then Exit Function
			If Data.inSectionID < 0 Then Exit Function
			If Data.GetRefState = 0 Then
				MsgBox(MsgList(60),vbOkOnly+vbInformation,MsgList(59))
				Exit Function
			ElseIf Data.lReferenceNum = 0 Then
				MsgBox(MsgList(61),vbOkOnly+vbInformation,MsgList(59))
				Exit Function
			End If
			DlgEnable "CopyButton",False
			DlgEnable "SearchButton",False
			Call GetVARefList(File,FN,Data,"",1,Mode,0)
			DlgText "ShowTextBox",Reference2Str(File,Data)
			DlgText "ShowText",Replace$(MsgList(22),"%s",CStr$(Data.lReferenceNum))
			DlgEnable "SearchButton",IIf(File.Magic = "",False,True)
			DlgEnable "CopyButton",IIf(DlgText("ShowTextBox") = "",False,True)
		Case "SearchButton"
			If DlgText("FilePathBox") = "" Then Exit Function
			DlgText "ShowTextBox",""
			DlgText "ShowText",Replace$(MsgList(22),"%s","0")
			DlgEnable "CopyButton",False
			DlgEnable "VACodeButton",False
			If DlgText("OffsetBox") = "" Or File.Magic = "" Then Exit Function
			If Data.inSectionID < 0 Then Exit Function
			Mode = LoadFile(File.FilePath,FN,0,0,0,LoadMode)
			If Mode < -1 Then
				UnLoadFile(FN,0,Mode)
				Exit Function
			End If
			DlgEnable "SearchButton",False
			'PSL.OutputWnd(0).Clear
			'PSL.Output MsgList(35)
			DlgText "ShowText",MsgList(35)
			Call GetVARefList(File,FN,Data,"",0,Mode,GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("ShowText")))
			UnLoadFile(FN,0,Mode)
			DlgText "ShowTextBox",Reference2Str(File,Data)
			DlgText "ShowText",Replace$(MsgList(22),"%s",CStr$(Data.lReferenceNum))
			DlgEnable "SearchButton",True
			DlgEnable "CopyButton",IIf(DlgText("ShowTextBox") = "",False,True)
			If Data.lReferenceNum > 0 Then DlgEnable "VACodeButton",True
		Case "CopyButton"
			Clipboard Replace$(Replace$(Replace$(Replace$(MsgList(36),"%p",File.FilePath), _
					"%o",DlgText("OffsetBox")),"%r",DlgText("RVABox")),"%s",DlgText("ShowTextBox"))
		Case "FileInfoButton"
			If File.Info = "" Then Call FileInfoView(File,True)
			ShowInfo File.FilePath,File.Info
		Case "LangButton"
			ReDim TempList(0) As String
			TempList = ReSplit(MsgList(64),";")
			i = ShowPopupMenu(TempList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			TempList = ReSplit(MsgList(65),";")
			If GetMsgList(MsgList,TempList(i)) = False Then Exit Function
			File.Info = ""

			'�����ı�������
			DlgText -1,Replace$(Replace$(MsgList(12),"%v",Version),"%b",Build)
			DlgText "FilePathButton",MsgList(13)
			DlgText "DecButton",MsgList(14)
			DlgText "HexButton",MsgList(15)
			DlgText "FileBitText",MsgList(16)
			DlgText "AboutButton",MsgList(17)
			DlgText "LangButton",MsgList(63)
			DlgText "RVAText",MsgList(18)
			DlgText "OffsetText",MsgList(19)
			DlgText "VACodeButton",MsgList(20)
			DlgText "SearchButton",MsgList(21)
			DlgText "ShowText",Replace$(MsgList(22),"%s","0")
			DlgText "CopyButton",MsgList(23)
			DlgText "FileInfoButton",MsgList(66)

			'�����ļ������б�����
			i = DlgValue("FileBitList")
			TempList = ReSplit(MsgList(24),";")
			DlgListBoxArray "FileBitList",TempList()
			DlgValue "FileBitList",i

			'�����ļ�������������
			If DlgText("FilePathBox") = "" Then Exit Function
			'�����ظ���������
			ChangSectionNames File,MsgList(26),MsgList(25)
			'��ȡ�ļ��������б�
			TempList = getSectionNameList(File.SecList)
			If File.Magic <> "" Then
				i = UBound(TempList)
				ReDim Preserve TempList(i + 3) As String
				TempList(i + 1) = MsgList(27)
				TempList(i + 2) = MsgList(28)
				TempList(i + 3) = MsgList(29)
			End If
			i = DlgValue("SecNameList")
			DlgListBoxArray "SecNameList",TempList()
			DlgValue "SecNameList",i

			'���������������
			DlgText "ShowTextBox",Reference2Str(File,Data)
			DlgText "ShowText",Replace$(MsgList(22),"%s",CStr$(Data.lReferenceNum))
		End Select
	Case 3 ' �ı��������Ͽ��ı�����ʱ
		Select Case DlgItem$
		Case "OffsetBox"
			If DlgText("FilePathBox") = "" Then Exit Function
			If DlgText("OffsetBox") = "" Then
				Data.lStartAddress = -1
				Data.inSectionID = -1
				DlgText "RVABox",""
				DlgValue "SecNameList",File.MaxSecIndex - Data.inSectionID
				DlgEnable "SearchButton",False
				DlgText "ShowTextBox",""
				DlgText "ShowText",Replace$(MsgList(22),"%s","0")
				DlgEnable "CopyButton",False
				DlgEnable "SearchButton",False
				Exit Function
			End If
			DlgText "OffsetBox",Replace$(DlgText("OffsetBox")," ","")
			If DlgValue("ValTypeGroup") = 0 Then
				If CheckDecStr(DlgText("OffsetBox"),0) = False Then
					MsgBox(MsgList(40),vbOkOnly+vbInformation,MsgList(0))
					DlgText "OffsetBox",CStr$(Data.lStartAddress)
					Exit Function
				ElseIf StrToLong(DlgText("OffsetBox")) >= File.FileSize Then
					MsgBox(MsgList(41),vbOkOnly+vbInformation,MsgList(0))
					DlgText "OffsetBox",CStr$(Data.lStartAddress)
					Exit Function
				Else
					i = Data.lStartAddress
					Data.lStartAddress = StrToLong(DlgText("OffsetBox"))
					DlgText "RVABox",CStr$(OffsetToRva(File,Data.lStartAddress))
					If i = Data.lStartAddress Then Exit Function
				End If
			Else
				If CheckHexStr(DlgText("OffsetBox"),0) = False Then
					MsgBox(MsgList(40),vbOkOnly+vbInformation,MsgList(0))
					DlgText "OffsetBox",FormatHexStr(Hex$(Data.lStartAddress),2)
					Exit Function
				ElseIf Val("&H" & DlgText("OffsetBox")) >= File.FileSize Then
					MsgBox(MsgList(41),vbOkOnly+vbInformation,MsgList(0))
					DlgText "OffsetBox",FormatHexStr(Hex$(Data.lStartAddress),2)
					Exit Function
				Else
					i = Data.lStartAddress
					Data.lStartAddress = Val("&H" & DlgText("OffsetBox"))
					DlgText "RVABox",FormatHexStr(Hex$(OffsetToRva(File,Data.lStartAddress)),2)
					If i = Data.lStartAddress Then Exit Function
				End If
			End If
			Data.inSectionID = SkipSection(File,Data.lStartAddress,0,0,1)
			If Data.inSectionID > -1 Then
				DlgValue "SecNameList",Data.inSectionID
				DlgEnable "SearchButton",IIf(File.Magic = "",False,True)
			ElseIf File.SecList(File.MaxSecIndex).lSizeOfRawData > 0 Then
				DlgValue "SecNameList",File.MaxSecIndex - Data.inSectionID
				DlgEnable "SearchButton",False
			Else
				DlgValue "SecNameList",File.MaxSecIndex - Data.inSectionID - 1
				DlgEnable "SearchButton",False
			End If
			DlgText "ShowTextBox",""
			DlgText "ShowText",Replace$(MsgList(22),"%s","0")
			DlgEnable "CopyButton",False
		Case "RVABox"
			If DlgText("FilePathBox") = "" Then Exit Function
			If DlgText("RVABox") = "" Then
				Data.lStartAddress = -1
				Data.inSectionID = -1
				DlgText "OffsetBox",""
				DlgValue "SecNameList",File.MaxSecIndex - Data.inSectionID
				DlgEnable "SearchButton",False
				DlgText "ShowTextBox",""
				DlgText "ShowText",Replace$(MsgList(22),"%s","0")
				DlgEnable "CopyButton",False
				DlgEnable "SearchButton",False
				Exit Function
			End If
			DlgText "RVABox",Replace$(DlgText("RVABox")," ","")
			If DlgValue("ValTypeGroup") = 0 Then
				If CheckDecStr(DlgText("RVABox"),0) = False Then
					MsgBox(MsgList(40),vbOkOnly+vbInformation,MsgList(0))
					DlgText "RVABox",CStr$(OffsetToRva(File,Data.lStartAddress))
					Exit Function
				End If
				If SkipSection(File,StrToLong(DlgText("RVABox")),0,0,3) = -3 Then
					MsgBox(MsgList(41),vbOkOnly+vbInformation,MsgList(0))
					DlgText "RVABox",CStr$(OffsetToRva(File,Data.lStartAddress))
					Exit Function
				Else
					i = Data.lStartAddress
					Data.lStartAddress = RvaToOffset(File,StrToLong(DlgText("RVABox")))
					DlgText "OffsetBox",CStr$(Data.lStartAddress)
					If i = Data.lStartAddress Then Exit Function
				End If
			Else
				If CheckHexStr(DlgText("RVABox"),0) = False Then
					MsgBox(MsgList(40),vbOkOnly+vbInformation,MsgList(0))
					DlgText "RVABox",FormatHexStr(Hex$(OffsetToRva(File,Data.lStartAddress)),2)
					Exit Function
				End If
				If SkipSection(File,Val("&H" & DlgText("RVABox")),0,0,3) = -3 Then
					MsgBox(MsgList(41),vbOkOnly+vbInformation,MsgList(0))
					DlgText "RVABox",FormatHexStr(Hex$(OffsetToRva(File,Data.lStartAddress)),2)
					Exit Function
				Else
					i = Data.lStartAddress
					Data.lStartAddress = RvaToOffset(File,Val("&H" & DlgText("RVABox")))
					DlgText "OffsetBox",FormatHexStr(Hex$(Data.lStartAddress),2)
					If i = Data.lStartAddress Then Exit Function
				End If
			End If
			Data.inSectionID = SkipSection(File,Data.lStartAddress,0,0,1)
			If Data.inSectionID > -1 Then
				DlgValue "SecNameList",Data.inSectionID
				DlgEnable "SearchButton",IIf(File.Magic = "",False,True)
			ElseIf File.SecList(File.MaxSecIndex).lSizeOfRawData > 0 Then
				DlgValue "SecNameList",File.MaxSecIndex - Data.inSectionID
				DlgEnable "SearchButton",False
			Else
				DlgValue "SecNameList",File.MaxSecIndex - Data.inSectionID - 1
				DlgEnable "SearchButton",False
			End If
			DlgText "ShowTextBox",""
			DlgText "ShowText",Replace$(MsgList(22),"%s","0")
			DlgEnable "CopyButton",False
		End Select
	Case 6 ' ���ܼ�
		Select Case SuppValue
		Case 1
			MsgBox Replace$(Replace$(MsgList(31),"%v",Version),"%b",Build),vbOkOnly+vbInformation,MsgList(32)
		Case 2
			Clipboard Replace$(Replace$(Replace$(Replace$(MsgList(36),"%p",File.FilePath), _
					"%o",DlgText("OffsetBox")),"%r",DlgText("RVABox")),"%s",DlgText("ShowTextBox"))
		End Select
	End Select
End Function


'��������ʮ�����ִ��Ƿ����Ҫ��
Private Function CheckDecStr(ByVal textStr As String,ByVal Length As Integer) As Boolean
	If textStr = "" Then Exit Function
	If Length <> 0 Then
		If (Len(textStr) Mod Length) <> 0 Then Exit Function
	End If
	If CheckStrRegExp(textStr,"[0-9]",0,1) = False Then Exit Function
	CheckDecStr = True
End Function


'��������ʮ�������ִ��Ƿ����Ҫ��
Private Function CheckHexStr(ByVal textStr As String,ByVal Length As Integer) As Boolean
	If textStr = "" Then Exit Function
	If Length <> 0 Then
		If (Len(textStr) Mod Length) <> 0 Then Exit Function
	End If
	If CheckStrRegExp(textStr,"[0-9A-Fa-f]",0,1) = False Then Exit Function
	CheckHexStr = True
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
			Set Matches = .Execute(TextStr)
			If Matches.Count = 0 Then Exit Function
			For n = 0 To Matches.Count - 1
				CheckStrRegExp = CheckStrRegExp + Matches(n).Length
			Next n
			CheckStrRegExp = IIf(CheckStrRegExp = Len(textStr),True,False)
		End Select
	End With
End Function


'���ݷ������õ�ַ�б��ȡ��Դ�����б�(�����ִ����)
'Mode = False '�޸����õ�ַ�б����򣬻�ȡ���õ�ַ�ʹ����б�
Private Sub GetRefList(strData As STRING_SUB_PROPERTIE,RefAddList() As String,RefVAList() As String,ByVal Mode As Boolean)
	Dim i As Long,n As Long,Dic As Object
	With strData
		If .lReferenceNum = 0 Then
			ReDim RefList(0) As REFERENCE_PROPERTIE
			ReDim AddList(0) As String,VAList(0) As String
		Else
			n = UBound(RefAddList)
			ReDim RefList(n) As REFERENCE_PROPERTIE
			ReDim AddList(n) As String,VAList(n) As String
			Set Dic = CreateObject("Scripting.Dictionary")
			For i = 0 To .lReferenceNum - 1
				If Not Dic.Exists(.Reference(i).lAddress) Then
					Dic.Add(CStr(.Reference(i).lAddress),i)
				End If
			Next i
			n = 0
			For i = 0 To UBound(RefAddList)
				If Dic.Exists(RefAddList(i)) Then
					RefList(n) = .Reference(Dic.Item(RefAddList(i)))
					AddList(n) = CStr(RefList(n).lAddress)
					VAList(n) = RefList(n).sCode
					n = n + 1
				End If
			Next i
			Set Dic = Nothing
			If n > 0 Then
				ReDim Preserve RefList(n - 1) As REFERENCE_PROPERTIE
				ReDim Preserve AddList(n - 1) As String,VAList(n - 1) As String
			Else
				ReDim RefList(0) As REFERENCE_PROPERTIE
				ReDim AddList(0) As String,VAList(0) As String
			End If
		End If
		If Mode = False Then
			RefAddList = AddList
			RefVAList = VAList
		Else
			.Reference = RefList
			.lReferenceNum = n
			.GetRefState = IIf(n = 0,0,1)
		End If
	End With
End Sub


'���������ǵ����ֽ�λ�ã������طǵ����ֽڿ�ʼλ��
Private Function getNotNullByteRegExp(FN As FILE_IMAGE,ByVal Offset As Long,ByVal Max As Long,ByVal Mode As Long) As Long
	If Offset + 1 > Max Then
		getNotNullByteRegExp = Offset
		Exit Function
	End If
	Dim Matches As Object,EndPos As Long
	With RegExp
		.Global = False
		.IgnoreCase = False
		.Pattern = "[^\x00]+?"
		Do
			EndPos = IIf(Offset + 512 < Max,Offset + 512,Max)
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


'��ȡ�Զ�������ֵ
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


'��ȡ�Զ�����������
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
		'GetTypeValue(FN,i,tmpFileHeader,Mode��)
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


'��ȡ���ļ���������
'Mode = False ���ƫ�Ƶ�ַ
'Mode = True �����������ַ
'�����ļ��������š�MinVal��MaxVal ֵ
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


'��ȡ�����ļ�ͷ�������ż���ַ (����ĵ�ַ����ת��Ϊƫ�Ƶ�ַ)
'Mode = 0 ʱ�������� RVA ����Ŀ¼��������
'Mode = 1 ʱ��RVA = RVA ����Ŀ¼������ַ + 1��SkipVal = �� RVA ���Ŀ¼��С��ַ
'Mode > 1 ʱ��RVA = RVA ����Ŀ¼����С��ַ - 1��SkipVal = �� RVA С��Ŀ¼����ַ
'fType = 0 ʱ������ .NET US ����
Private Function SkipHeader(File As FILE_PROPERTIE,RVA As Long,Optional SkipVal As Long,Optional ByVal Mode As Long,Optional ByVal fType As Long) As Long
	Dim i As Integer,j As Integer,EndPos As Long
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
				EndPos = .lVirtualAddress + .lSize - 1		'����ַ
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
						SkipHeader = i
						RVA = EndPos + 1
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
					If EndPos < RVA Then
						If RVA < SkipVal Then
							SkipVal = EndPos
						ElseIf EndPos > SkipVal Then
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


'���͸��������ļ��������б�
Private Sub ChangSectionNames(File As FILE_PROPERTIE,Optional ByVal HideSecName As String,Optional ByVal NoPEName As String)
	Dim i As Integer,Dic As Object
	With File
		If .Magic = "" Then
			.SecList(0).sName = NoPEName
			.SecList(1).sName = HideSecName
		Else
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
		End If
	End With
End Sub


'��ȡ�����ļ��������б�
'Mode = 0 ��ȡ�������Σ������������Σ����򲻰�����������
Private Function getSectionNameList(SecList() As SECTION_PROPERTIE,Optional ByVal Mode As Boolean) As String()
	Dim i As Integer
	i = UBound(SecList)
	If Mode = False Then
		If SecList(i).lSizeOfRawData > 0 Then
			ReDim TempList(i) As String
		Else
			ReDim TempList(i - 1) As String
		End If
	Else
		ReDim TempList(i - 1) As String
	End If
	For i = 0 To UBound(TempList)
		TempList(i) = SecList(i).sName
	Next i
	getSectionNameList = TempList
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


'���� .NET �ִ����õ�ַ�б�
'fType = 0 ������Դ�������б�����ô���
'fType = 1 ���ҷ���������б�����ô��룬��� RefAdds Ϊ�գ�����ԭ�����õ�ַ�������ô���
'fType = 2 ���ҷ���������б�����ô��룬��� RefAdds Ϊ�գ����ʼ������������б�
'fType > 2 ��ʼ������շ��������б�����ô���
Private Function GetNETVARefList(File As FILE_PROPERTIE,FN As Variant,strData As STRING_SUB_PROPERTIE,ByVal StrTypeLength As Integer, _
				ByVal RefAdds As String,ByVal fType As Long,ByVal Mode As Long,Optional ByVal ShowMsg As Long) As Long
	Dim i As Long,j As Long,m As Long,n As Long
	Dim Msg As String,TempList() As String
	On Error GoTo ExitFunction
	If File.Magic = "" Then GoTo ExitFunction
	If fType > 2 Then GoTo ExitFunction
	If ShowMsg > 0 Then
		Msg = GetTextBoxString(ShowMsg) & " "
	ElseIf ShowMsg < 0 Then
		ReDim TempList(PSL.OutputWnd(0).LineCount - 1) As String
		For i = 1 To PSL.OutputWnd(0).LineCount
			TempList(i - 1) = PSL.OutputWnd(0).Text(i)
		Next i
		Msg = StrListJoin(TempList,vbCrLf) & " "
	End If
	With strData
		'��ԭ�����뿪ʼ��ַ�����ô����ȡ�µ�ַ�����ô����б�
		If fType < 0 Then
			'If RefAdds <> "" Then Call getRefList(strData,ReSplit(RefAdds,RefJoinStr),TempList,True)
			If .lReferenceNum = 0 Then GoTo ExitFunction
			j = File.StreamList(File.USStreamID).lPointerToRawData
			.Reference(0).sCode = ReverseHexCode(Hex$(.lStartAddress - j - StrTypeLength),6) & "70"
			For i = 0 To .lReferenceNum - 1
				.Reference(i).sCode = .Reference(0).sCode
			Next i
			.GetRefState = 1
			Exit Function
		End If
		If fType = 0 Then
			'��ȡ�����õ��˳�����
			If .GetRefState > 0 Then Exit Function
			.lReferenceNum = 0
			ReDim strData.Reference(0) 'As REFERENCE_PROPERTIE
			j = File.StreamList(File.USStreamID).lPointerToRawData
			.Reference(0).sCode = ReverseHexCode(Hex$(.lStartAddress - j - StrTypeLength),6) & "70"
			i = File.SecList(File.MinSecID).lPointerToRawData
			If RefAdds = "" Then
				j = File.StreamList(File.USStreamID).lPointerToRawData
				RefAdds = ByteToString(GetBytes(FN,j - i + 1,i,Mode),CP_ISOLATIN1)
			End If
			TempList = GetVAListRegExp(RefAdds,"\x72" & HexStr2RegExpPattern(.Reference(0).sCode,1),i)
			If CheckArray(TempList) = True Then
				.lReferenceNum = UBound(TempList) + 1
				ReDim Preserve strData.Reference(.lReferenceNum - 1) 'As REFERENCE_PROPERTIE
				For i = 0 To .lReferenceNum - 1
					.Reference(i).lAddress = CLng(TempList(i)) + 1
					.Reference(i).sCode = .Reference(0).sCode
					If .Reference(i).lAddress < n Or .Reference(i).lAddress > m Then
						j = SkipSection(File,.Reference(i).lAddress,n,m)
					End If
					.Reference(i).inSecID = j
					If ShowMsg > 0 Then
						SetTextBoxString ShowMsg,Msg & Format$(i / .lReferenceNum,"#%")
					ElseIf ShowMsg < 0 Then
						PSL.OutputWnd(0).Clear
						PSL.Output Msg & Format$(i / .lReferenceNum,"#%")
					End If
				Next i
			End If
			.GetRefState = 1
		ElseIf RefAdds <> "" Or (.lReferenceNum > 0 And fType < 2) Then
			If RefAdds <> "" Then Call GetRefList(strData,ReSplit(RefAdds,RefJoinStr),TempList,True)
			If .lReferenceNum > 0 Then
				j = File.StreamList(File.USStreamID).lPointerToRawData
				.Reference(0).sCode = ReverseHexCode(Hex$(.lStartAddress - j - StrTypeLength),6) & "70"
				For i = 0 To .lReferenceNum - 1
					.Reference(i).sCode = .Reference(0).sCode
					If ShowMsg > 0 Then
						SetTextBoxString ShowMsg,Msg & Format$(i / .lReferenceNum,"#%")
					ElseIf ShowMsg < 0 Then
						PSL.OutputWnd(0).Clear
						PSL.Output Msg & Format$(i / .lReferenceNum,"#%")
					End If
				Next i
			End If
			.GetRefState = 1
		Else
			GoTo ExitFunction
		End If
	End With
	GetNETVARefList = strData.lReferenceNum
	If ShowMsg > 0 Then
		SetTextBoxString ShowMsg,Msg & "100%"
	ElseIf ShowMsg < 0 Then
		PSL.OutputWnd(0).Clear
		PSL.Output Msg & "100%"
	End If
	Exit Function
	'�˳�����
	ExitFunction:
	ReDim Preserve strData.Reference(0) 'As REFERENCE_PROPERTIE
	strData.lReferenceNum = 0
	strData.GetRefState = 0
	If ShowMsg > 0 Then
		SetTextBoxString ShowMsg,Msg & "100%"
	ElseIf ShowMsg < 0 Then
		PSL.OutputWnd(0).Clear
		PSL.Output Msg & "100%"
	End If
End Function


'�������ô���������б�
'fType < 0 �����е����õ�ַ�б����ҷ�������ô����б�fType Ϊԭ���ķ��뿪ʼ��ַ
'fType = 0 ������Դ�������б�����ô���
'fType = 1 ���ҷ���������б�����ô��룬��� RefAdds Ϊ�գ�����ԭ�����õ�ַ�������ô���
'fType = 2 ���ҷ���������б�����ô��룬��� RefAdds Ϊ�գ����ʼ������������б�
'fType > 2 ��ʼ������շ��������б�����ô���
'�����ַ(VA) = StartPos + ImageBase + VRK
Private Function GetVARefList(File As FILE_PROPERTIE,FN As Variant,strData As STRING_SUB_PROPERTIE, _
	ByVal RefAdds As String,ByVal fType As Long,ByVal Mode As Long,Optional ByVal ShowMsg As Long) As Long
	Dim i As Long,j As Long,k As Long,m As Long,n As Long
	Dim RVA As Long,VRK As Long,MaxPos As Long,RSize As Long,SkipVal As Long
	Dim Msg As String,TempList() As String
	On Error GoTo ExitFunction
	If File.Magic = "" Then GoTo ExitFunction
	If fType > 2 Then GoTo ExitFunction
	With strData
	If File.LangType = NET_FILE_SIGNATURE And File.USStreamID > -1 Then
		i = File.StreamList(File.USStreamID).lPointerToRawData
		j = i + File.StreamList(File.USStreamID).lSizeOfRawData - 1
		If .lStartAddress >= i And .lStartAddress <= j Then
			i = .lStartAddress
			Do
				i = i - 1
				If i < i Then Exit Do
				j = GetByte(FN,i,Mode)
				n = n + 1
			Loop Until j = 0 Or j = 1
			i = i + 1
			i = CorSigUncompressData(FN,i,j,Mode)
			If j > 0 And i + 1 = n Then
				GetVARefList = GetNETVARefList(File,FN,strData,i,RefAdds,fType,Mode,ShowMsg)
			End If
			Exit Function
		End If
	End If
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
		'��ԭ�����뿪ʼ��ַ�����ô����ȡ�µ�ַ�����ô����б�
		If fType < 0 Then
			'If RefAdds <> "" Then Call getRefList(strData,ReSplit(RefAdds,RefJoinStr),TempList,True)
			If .lReferenceNum = 0 Then GoTo ExitFunction
			i = .Reference(0).inSecID
			VRK = File.SecList(i).lVirtualAddress - File.SecList(i).lPointerToRawData
			RVA = Val("&H" & ReverseHexCode(.Reference(0).sCode,8)) + .Reference(0).lAddress + VRK + 4
			RVA = RVA + fType + .lStartAddress
			For i = 0 To .lReferenceNum - 1
				j = .Reference(i).inSecID
				VRK = File.SecList(j).lVirtualAddress - File.SecList(j).lPointerToRawData
				.Reference(i).sCode = ReverseHexCode(Hex$(RVA - (.Reference(i).lAddress + VRK + 4)),8)
				'.Reference(i).sCode = Byte2Hex(Val2Bytes(RVA - (.Reference(i).lAddress + VRK + 4),4),0,3)
			Next i
			.GetRefState = 1
			Exit Function
		End If
		If .inSectionID > File.MaxSecIndex - 1 Then GoTo ExitFunction
		'��ȡ�ִ��������ַ
		With File.SecList(.inSectionID)
			If strData.lStartAddress >= .lPointerToRawData And strData.lStartAddress < .lPointerToRawData + .lSizeOfRawData Then
				RVA = strData.lStartAddress + .lVirtualAddress - .lPointerToRawData
			Else
				GoTo ExitFunction
			End If
		End With
		If fType = 0 Then
			'��ȡ�����õ��˳�����
			'If .GetRefState > 0 Then Exit Function
			If SkipHeader(File,strData.lStartAddress,0,0) > -1 Then GoTo ExitFunction
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
				SkipVal = i - 1
				Do While i < RSize
					'�ų�ĳЩ����Ŀ¼���κ� .NET �ļ���������
					If i > SkipVal Then
						k = i: m = SkipHeader(File,k,SkipVal,1)
						If m = 2 Or m = 4 Or m = 5 Or m > 15 Then i = k
						If i > SkipVal Or SkipVal > RSize Then SkipVal = RSize + 1
						If i > RSize Then Exit Do
					End If
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
			If RefAdds <> "" Then Call GetRefList(strData,ReSplit(RefAdds,RefJoinStr),TempList,True)
			If .lReferenceNum > 0 Then
				For i = 0 To .lReferenceNum - 1
					j = .Reference(i).inSecID
					VRK = File.SecList(j).lVirtualAddress - File.SecList(j).lPointerToRawData
					.Reference(i).sCode = ReverseHexCode(Hex$(RVA - (.Reference(i).lAddress + VRK + 4)),8)
					'.Reference(i).sCode = Byte2Hex(Val2Bytes(RVA - (.Reference(i).lAddress + VRK + 4),4),0,3)
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
		'��ԭ�����뿪ʼ��ַ�����ô����ȡ�µ�ַ�����ô����б�
		If fType < 0 Then
			'If RefAdds <> "" Then Call getRefList(strData,ReSplit(RefAdds,RefJoinStr),TempList,True)
			If .lReferenceNum = 0 Then GoTo ExitFunction
			VRK = Val("&H" & ReverseHexCode(.Reference(0).sCode,8)) + fType
			.Reference(0).sCode = ReverseHexCode(Hex$(.lStartAddress + VRK),8)
			'.Reference(0).sCode = Byte2Hex(Val2Bytes(.lStartAddress + VRK,4),0,3)
			For i = 0 To .lReferenceNum - 1
				.Reference(i).sCode = .Reference(0).sCode
			Next i
			.GetRefState = 1
			Exit Function
		End If
		If .inSectionID > File.MaxSecIndex - 1 Then GoTo ExitFunction
		'��ȡ�ִ��������ַ
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
		'��ȡ���õ�ַ�����ô����б�
		If fType = 0 Then
			'��ȡ�����õ��˳�����
			'If .GetRefState > 0 Then Exit Function
			If SkipHeader(File,strData.lStartAddress,0,0) > -1 Then GoTo ExitFunction
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
			.lReferenceNum = 0
			ReDim strData.Reference(0) 'As REFERENCE_PROPERTIE
			.Reference(0).sCode = ReverseHexCode(Hex$(.lStartAddress + VRK),8)
			TempList = GetVAListRegExp(RefAdds,HexStr2RegExpPattern(.Reference(0).sCode,1),SkipVal)
			If CheckArray(TempList) = True Then
				.lReferenceNum = UBound(TempList) + 1
				ReDim Preserve strData.Reference(.lReferenceNum - 1) 'As REFERENCE_PROPERTIE
				m = 0: n = 0
				For i = 0 To .lReferenceNum - 1
					.Reference(i).lAddress = CLng(TempList(i))
					.Reference(i).sCode = .Reference(0).sCode
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
			If RefAdds <> "" Then Call GetRefList(strData,ReSplit(RefAdds,RefJoinStr),TempList,True)
			If .lReferenceNum > 0 Then
				.Reference(0).sCode = ReverseHexCode(Hex$(.lStartAddress + VRK),8)
				'.Reference(0).sCode = Byte2Hex(Val2Bytes(.lStartAddress + VRK,4),0,3)
				For i = 0 To .lReferenceNum - 1
					.Reference(i).sCode = .Reference(0).sCode
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
	'�˳�����
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


'��ȡ PE64 ԭʼ�ִ������õ�ַ�ʹ���
'���� GetVAListPE64 = ��������
Private Function GetVAListPE64(FN As Variant,strData As STRING_SUB_PROPERTIE,RefMaxNum As Long,ByVal SecID As Long, _
				ByVal VRK As Long,ByVal StartPos As Long,ByVal RSize As Long,ByVal Mode As Long) As Long
	Dim i As Long,Temp As String
	With strData
		'i = (VRK - StartPos) And &HFF	'��3���ֽڲ��ң��ٶȽ���
		'i = Val("&H" & Right$("0000" & Hex$(VRK - StartPos),4))	'��2���ֽڲ��ң��ٶȽϿ�
		i = (VRK - StartPos) And 65535	'��2���ֽڲ��ң��ٶȽϿ죬����� 65535 �����滻�� &HFFFF����Ϊ &HFFFF ����Ϊ -1
		If i > RSize - StartPos Then i = RSize - StartPos
		If i > .lStartAddress - StartPos - 4 Then i = .lStartAddress - StartPos - 4
		If i > 0 Then
			GetVAListPE64 = i
			'������ʽ���ң��ٶȽϿ죬��3���ֽڲ���ʱ����ʼ��ַΪ StartPos + 1������Ϊ StartPos + 2
			ReDim TempList(0) As String
			Temp = RefFrontChar & HexStr2RegExpPattern(Right$(ReverseHexCode(Hex$(VRK - StartPos),8),4),1)
			TempList = GetVAListRegExp(ByteToString(GetBytes(FN,i + 4,StartPos - 3,Mode),CP_ISOLATIN1),Temp,StartPos - 3)
			'�ֽ�������ң��ٶȽ�������3���ֽڲ���ʱ����ʼ��ַΪ StartPos + 1������Ϊ StartPos + 2
			'TempList = GetVAList(FN.ImageByte,Val2BytesRev(VRK - StartPos,4,2),StartPos + 2,StartPos + 2 + GetVAListPE64)
			If CheckArray(TempList) = False Then Exit Function
			For i = 0 To UBound(TempList)
				StartPos = CLng(TempList(i)) + 3	'ǰ3��Ϊ���õ������룬��������3���ֽ�
				If VRK > StartPos Then
					'��ȡ�����ַ(�����ô���ֵ)�����ж����Ƿ���ȷ
					RSize = GetLong(FN,StartPos,Mode)
					If RSize > 0 And RSize = VRK - StartPos Then
						If .lReferenceNum > RefMaxNum Then
							RefMaxNum = .lReferenceNum + 20
							ReDim Preserve strData.Reference(RefMaxNum) 'As REFERENCE_PROPERTIE
						End If
						'�����ҵ������õ�ַ�����ô���
						.Reference(.lReferenceNum).lAddress = StartPos
						.Reference(.lReferenceNum).sCode = Byte2Hex(GetBytes(FN,4,StartPos,Mode),0,3)
						.Reference(.lReferenceNum).inSecID = SecID
						.lReferenceNum = .lReferenceNum + 1
					End If
				End If
			Next i
		ElseIf VRK > StartPos Then
			'��ȡ�����ַ(�����ô���ֵ)�����ж����Ƿ���ȷ
			TempList = GetVAListRegExp(ByteToString(GetBytes(FN,5,StartPos - 3,Mode),CP_ISOLATIN1),RefFrontChar,StartPos - 3)
			If CheckArray(TempList) = False Then Exit Function
			RSize = GetLong(FN,StartPos,Mode)
			If RSize > 0 And RSize = VRK - StartPos Then
				GetVAListPE64 = 3
				If .lReferenceNum > RefMaxNum Then
					RefMaxNum = .lReferenceNum + 20
					ReDim Preserve strData.Reference(RefMaxNum) 'As REFERENCE_PROPERTIE
				End If
				'�����ҵ������õ�ַ�����ô���
				.Reference(.lReferenceNum).lAddress = StartPos
				.Reference(.lReferenceNum).sCode = Byte2Hex(GetBytes(FN,4,StartPos,Mode),0,3)
				.Reference(.lReferenceNum).inSecID = SecID
				.lReferenceNum = .lReferenceNum + 1
			End If
		End If
	End With
End Function


'��ȡ���ֽ��������ҵ���ƥ��������б�(VB��ʽ)
'ע�⣺StartPos��EndPos ��Ϊ�� 0 ��ʼ�ĵ�ַ
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
		TempList(n) = CStr(i - 1)    'ע�� InStrB �����ҵ���һ�����ͷ���"1"
		n = n + 1
		NextNum:
		i = InStrB(i + Length,Bytes,TempByte)
	Loop
	If n > 0 Then n = n - 1
	ReDim Preserve TempList(n) As String
	GetVAList = TempList
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


'ת����ֵΪ�ֽ�����(���ڳ��ȵĸ�λ�ض�)
Private Function Val2Bytes(ByVal Value As Long,ByVal Length As Integer,Optional ByVal ByteOrder As Boolean) As Byte()
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


'ת����ֵΪ�ֽ�����(���ڳ��ȵĵ�λ�ض�)
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


'Hex �ַ���ת������ʽʹ�õ� Hex ת���ģ��
'Mode = 0 תΪ�� [] ��ʽ������Ϊ�� [] ��ʽ
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


'��ʽ�� HEX �ִ�
Private Function FormatHexStr(ByVal textStr As String,ByVal Length As Integer) As String
	If textStr = "" Then Exit Function
	If (Len(textStr) Mod Length) = 0 Then
		FormatHexStr = textStr
	Else
		FormatHexStr = "0" & textStr
	End If
End Function


'��������ת�ִ�
Private Function Reference2Str(File As FILE_PROPERTIE,strData As STRING_SUB_PROPERTIE) As String
	Dim i As Long,k As Long,n As Long,Temp As String,SecList() As String
	On Error Resume Next
	If strData.lReferenceNum = 0 Then
		Reference2Str = MsgList(30)
		Exit Function
	End If
	Temp = String$(Len(CStr$(File.FileSize)),"0")
	n = Len(FormatHexStr(Hex$(File.FileSize),2))
	SecList = getSectionNameList(File.SecList)
	With strData
		ReDim TempList(.lReferenceNum + 4) As String
		TempList(0) = MsgList(37)
		TempList(1) = MsgList(38)
		TempList(2) = MsgList(37)
		For i = 0 To .lReferenceNum - 1
			TempList(i + 3) = Replace$(MsgList(39),"%no",CStr$(i + 1))
			TempList(i + 3) = Replace$(TempList(i + 3),"%da",Format$(.Reference(i).lAddress,Temp))
			TempList(i + 3) = Replace$(TempList(i + 3),"%ha",Right$(Temp & Hex$(.Reference(i).lAddress),n))
			TempList(i + 3) = Replace$(TempList(i + 3),"%sc",SecList(SkipSection(File,.Reference(i).lAddress,0,0,1)))
			k = SkipHeader(File,.Reference(i).lAddress)
			If k < 16 Then
				TempList(i + 3) = Replace$(TempList(i + 3),"%dc",MsgList(42 + k + 1))
			Else
				TempList(i + 3) = Replace$(TempList(i + 3),"%dc",MsgList(140 + k - 22))
			End If
			TempList(i + 3) = Replace$(TempList(i + 3),"%rc",strData.Reference(i).sCode)
		Next i
	End With
	TempList(i + 3) = MsgList(37)
	Reference2Str = StrListJoin(TempList,TextJoinStr)
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


'��ȡż��λ
'Mode = 0 ������ 1 ���ֽڣ�Mode = 1 ������ 1 ���ֽ�
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


'���������в���
'' Command Line Format: Command <FilePath> <-add:Address> <-lng:Hex Language Code>
'' Command: Name of this Macros file
'' FilePath: Full path of PE file
'' Address: Offset of Dec or Hex, if Hex, use like "0xFFFF" format
'' Hex Language Code: Display UI Language. Supports EngLish, Chinese Simplified and Chinese Traditional only. For sample: 0804,1004,0404,0C04,1404.
'' Return: No
'' For example: modRefernceSearch,"d:\my folder\my file.exe" -add:35148 -lng:0804

Private Function SplitArgument(ByVal Argument As String,ByVal MaxNum As Long) As String()
	Dim i As Long,j As Long,k As Long,TempList() As String
	ReDim ArgArray(MaxNum - 1) As String
	If Argument = "" Then
		SplitArgument = ArgArray
		Exit Function
	End If
	'�Ӻ���ǰ������ֵ���������С����
	TempList = ReSplit(Argument," ")
	k = UBound(TempList)
	MaxNum = IIf(k - MaxNum > 0,k - MaxNum,1)
	For i = k To MaxNum Step -1
		Argument = Trim$(TempList(i))
		If CheckStrRegExp(Argument,"-add:[0-9]+",0,5,True) = True Then
			Argument = "-add:"
		ElseIf CheckStrRegExp(Argument,"-lng:[0-9a-f;]+",0,5,True) = True Then
			Argument = "-lng:"
		End If
		Select Case Argument
		Case "-add:"
			If ArgArray(1) = "" Then
				ArgArray(1) = Mid$(Trim$(TempList(i)),6)
				j = j + 1
			End If
		Case "-lng:"
			If ArgArray(2) = "" Then
				ArgArray(2) = Mid$(Trim$(TempList(i)),6)
				j = j + 1
			End If
		End Select
	Next i
	ReDim Preserve TempList(k - j) As String
	ArgArray(0) = Join$(TempList," ")
	SplitArgument = ArgArray
End Function


'��ȥ�ִ�ǰ��ָ���� PreStr �� AppStr
'fType = -1 ��ȥ���ִ�ǰ��Ŀո������ָ���� PreStr �� AppStr������ȥ���ִ���ǰ��ո�
'fType = 0 ȥ���ִ�ǰ��Ŀո������ָ���� PreStr �� AppStr������ȥ���ִ���ǰ��ո�
'fType = 1 ȥ���ִ�ǰ��Ŀո������ָ���� PreStr �� AppStr����ȥ���ִ���ǰ��ո�
'fType = 2 ȥ���ִ�ǰ��Ŀո��ָ���� PreStr �� AppStr 1 �Σ�����ȥ���ִ���ǰ��ո�
'fType > 2 ȥ���ִ�ǰ��Ŀո��ָ���� PreStr �� AppStr 1 �Σ���ȥ���ִ���ǰ��ո�
Private Function RemoveBackslash(ByVal Path As String,ByVal PreStr As String,ByVal AppStr As String,ByVal fType As Long) As String
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
Private Function ReSplit(ByVal TextStr As String,Optional ByVal Sep As String = " ",Optional ByVal Max As Integer = -1) As String()
	If TextStr = "" Then
		ReDim TempList(0) As String
		ReSplit = TempList
	Else
		ReSplit = Split(TextStr,Sep,Max)
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


'��ʾ�ļ���Ϣ
Private Sub ShowInfo(ByVal FilePath As String,ByVal Info As String)
	Begin Dialog UserDialog 890,448,Replace$(MsgList(89),"%s",FilePath) ' %GRID:10,7,1,1
		TextBox 0,7,890,406,.InTextBox,1
		OKButton 390,420,100,21,.OKButton
	End Dialog
	Dim dlg As UserDialog
	dlg.InTextBox = Info
	Dialog dlg
End Sub


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
	'ÿ���ļ��ڵ�ƫ�Ƶ�ַ
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
	'���ؽڵ�ƫ�Ƶ�ַ���� PE ��ַ������
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
	'ÿ���ļ��ڵ���������ַ
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
	'���ؽڵ���������ַ���� PE ��ַ������
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
	'����Ŀ¼��ַ�������ļ���
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
	'.NET CLR ����Ŀ¼��ַ�������ļ���
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
	'.NET ����ַ�������ļ���
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
	'������
	ErrHandle:
	On Error Resume Next
	Err.Source = "NotReadFile"
	Err.Description = Err.Description & JoinStr & File.FilePath & ".xls"
	Call sysErrorMassage(Err,1)
End Sub


'ת��ʮ���ƺ�ʮ������ֵΪ�ַ�
'MaxVal = 0 ��ֵ����Ӧ�еĳ��ȣ�> 0 ���ļ���С�����λ����< 0 ��ָ��λ��
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


'λ����
Private Function SHL(nSource As Long, n As Byte) As Double
	On Error GoTo ExitFunction:
	SHL = nSource * 2 ^ n
	ExitFunction:
End Function


'Blob �����Ƚ�ѹ��
'���� CorSigUncompressData = ѹ�����ȣ�Length = �����ȱ�ʶ������ֽڳ��ȣ������Ƿ���� > &H7F �ַ���ʶ����
'ÿ�����������ݿ�ͷ������1���������ݿ飬ͨ����λ���㣬������������ݿ��ʵ�ʳ���
'�����һ���ֽ����λΪ0��������ݿ鳤��Ϊ1���ֽ�
'�����һ���ֽ����λΪ10��������ݿ鳤��Ϊ2���ֽ�
'�����һ���ֽ����λΪ110��������ݿ鳤��Ϊ4���ֽ�
Private Function CorSigUncompressData(FN As Variant,ByVal Index As Long,Length As Long,ByVal Mode As Long) As Integer
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


'��Ϣ�ַ���
Private Function GetMsgList(MsgList() As String,ByVal Language As String) As Boolean
	Dim i As Integer
	ReDim MsgList(152) As String
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

		MsgList(12) = "�������� - �汾 %v (���� %b)"
		MsgList(13) = "..."
		MsgList(14) = "Dec"
		MsgList(15) = "Hex"
		MsgList(16) = "����"
		MsgList(17) = "����"
		MsgList(18) = "RVA"
		MsgList(19) = "ʵ��ַ"
		MsgList(20) = "�������"
		MsgList(21) = "����"
		MsgList(22) = "���� (%s):"
		MsgList(23) = "����"

		MsgList(24) = "δ֪;Not PE;PE32;PE64;MAC32;MAC64"
		MsgList(25) = "������"
		MsgList(26) = "��������"
		MsgList(27) = "�����ļ�"
		MsgList(28) = "�ļ�ͷ"
		MsgList(29) = "�� PE �ļ�"
		MsgList(30) = "������"
		MsgList(31) = "�汾 %v (���� %b)\r\n" & _
					"OS �汾: Windows XP/2000 ������\r\n" & _
					"Passolo �汾: Passolo 5.0 ������\r\n" & _
					"��Ȩ: ������\r\n" & _
					"��ַ: http://www.hanzify.org\r\n" & _
					"����: wanfu (2015 - 2019)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(32) = "������������"
		MsgList(33) = "��ִ���ļ� (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|�����ļ� (*.*)|*.*||"
		MsgList(34) = "ѡ���ļ�"
		MsgList(35) = "������������, ���Ժ�..."
		MsgList(36) = "�ļ�: %p\r\nOffset: %o, RVA: %r\r\n%s"

		MsgList(37) = "================================================"
		MsgList(38) = "���\tDec ��ַ\tHex ��ַ\t����\tĿ¼����\t���ô���"
		MsgList(39) = "#%no\t%da\t%ha\t%sc\t%dc\t%rc"
		MsgList(40) = "����ĵ�ַ�Ƿ���"
		MsgList(41) = "����ĵ�ַ�����ļ��ĳ��ȡ�"
		MsgList(42) = "��Χ��"
		MsgList(43) = "����Ŀ¼"
		MsgList(44) = "����Ŀ¼"
		MsgList(45) = "��Դ"
		MsgList(46) = "�쳣"
		MsgList(47) = "��ȫ"
		MsgList(48) = "�ض�λ������"
		MsgList(49) = "����"
		MsgList(50) = "��Ȩ"
		MsgList(51) = "����ֵ"
		MsgList(52) = "�̱߳��ش洢"
		MsgList(53) = "��������Ŀ¼"
		MsgList(54) = "�������"
		MsgList(55) = "�����ַ��"
		MsgList(56) = "�ӳٵ���"
		MsgList(57) = "COM ������"
		MsgList(58) = "����"
		MsgList(59) = "��Ϣ"
		MsgList(60) = "�밴�ִ���ԭʼ��ַ�������ã�Ȼ�������µ�ַ�Լ����µ����ô��롣"
		MsgList(61) = "�����õ��ִ����޷������µ�ַ�����ô��롣"
		MsgList(62) = "ȡ��"
		MsgList(63) = "����"
		MsgList(64) = "Ӣ��;��������;��������"
		MsgList(65) = "enu;chs;cht"
		MsgList(66) = "�ļ���Ϣ"
		MsgList(89) = "��Ϣ - %s"

		MsgList(91) = "============ �ļ���Ϣ ============\r\n"
		MsgList(92) = "�ļ����ƣ�\t%s"
		MsgList(93) = "�ļ�·����\t%s"
		MsgList(94) = "�ļ�˵����\t%s"
		MsgList(95) = "�ļ��汾��\t%s"
		MsgList(96) = "��Ʒ���ƣ�\t%s"
		MsgList(97) = "��Ʒ�汾��\t%s"
		MsgList(98) = "��Ȩ���У�\t%s"
		MsgList(99) = "�ļ���С��\t%s �ֽ�"
		MsgList(100) = "�������ڣ�\t%s"
		MsgList(101) = "�޸����ڣ�\t%s"
		MsgList(102) = "����ԣ�\t%s"
		MsgList(103) = "�� �� �̣�\t%s"
		MsgList(104) = "ԭʼ�ļ�����\t%s"
		MsgList(105) = "�ڲ��ļ�����\t%s"
		MsgList(106) = "�ļ����ͣ�\t%s"
		MsgList(107) = "ӳ���ַ��\t%s"
		MsgList(108) = "������Ϣ��"
		MsgList(109) = "��ַ���\t������\t��ʼ��ַ\t������ַ\t�ֽڴ�С"
		MsgList(110) = "%s!1!\t%s!2!\t%s!4!\t%s!5!\t%s!6!"
		MsgList(111) = "================================="
		MsgList(112) = "�ļ�ƫ�Ƶ�ַ"
		MsgList(113) = "��������ַ"
		MsgList(114) = "����"
		MsgList(115) = "����"
		MsgList(116) = "δ֪"
		MsgList(117) = "������"
		MsgList(118) = "����Ŀ¼��Ϣ (�ļ�ƫ�Ƶ�ַ)��"
		MsgList(119) = "Ŀ¼����\t��������\t��ʼ��ַ\t������ַ\t�ֽڴ�С"
		MsgList(120) = "����Ŀ¼"
		MsgList(121) = "����Ŀ¼"
		MsgList(122) = "��ԴĿ¼"
		MsgList(123) = "�쳣Ŀ¼"
		MsgList(124) = "��ȫĿ¼"
		MsgList(125) = "��ַ�ض�λ��"
		MsgList(126) = "����Ŀ¼"
		MsgList(127) = "��ȨĿ¼"
		MsgList(128) = "����ֵ(GP RVA)"
		MsgList(129) = "�̱߳��ش洢��"
		MsgList(130) = "��������Ŀ¼"
		MsgList(131) = "�󶨵���Ŀ¼"
		MsgList(132) = "�����ַ��"
		MsgList(133) = "�ӳټ��ص����"
		MsgList(134) = "COM ���п��־"
		MsgList(135) = "����Ŀ¼"
		MsgList(136) = "�쳣"
		MsgList(137) = "������"
		MsgList(138) = ".NET CLR ����Ŀ¼��Ϣ (�ļ�ƫ�Ƶ�ַ)��"
		MsgList(139) = "Ŀ¼����\t��������\t��ʼ��ַ\t������ַ\t�ֽڴ�С"
		MsgList(140) = "Ԫ����(MetaData)"
		MsgList(141) = "�й���Դ"
		MsgList(142) = "ǿ����ǩ��"
		MsgList(143) = "��������"
		MsgList(144) = "�����(V-��)"
		MsgList(145) = "��ת������ַ��"
		MsgList(146) = "�йܱ���ӳ��ͷ"
		MsgList(147) = ".NET MetaData ����Ϣ (�ļ�ƫ�Ƶ�ַ)��"
		MsgList(148) = "������\t��������\t��ʼ��ַ\t������ַ\t�ֽڴ�С"
		MsgList(149) = "�� PE �ļ�"
		MsgList(150) = "��PE(%s)"
		MsgList(151) = "��ַ���\t����\t����\t\t��ʼ��ַ\t������ַ\t�ֽڴ�С"
		MsgList(152) = "%s!1!\t%s!2!\t%s!3!\t\t%s!4!\t%s!5!\t%s!6!"
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

		MsgList(12) = "�ѷӷj�� - ���� %v (�c�� %b)"
		MsgList(13) = "..."
		MsgList(14) = "Dec"
		MsgList(15) = "Hex"
		MsgList(16) = "����"
		MsgList(17) = "����"
		MsgList(18) = "RVA"
		MsgList(19) = "���}"
		MsgList(20) = "�N�X�p��"
		MsgList(21) = "�j��"
		MsgList(22) = "�ѷ� (%s):"
		MsgList(23) = "�ƻs"

		MsgList(24) = "����;Not PE;PE32;PE64;MAC32;MAC64"
		MsgList(25) = "�L�Ϭq"
		MsgList(26) = "���ðϬq"
		MsgList(27) = "�W�X�ɮ�"
		MsgList(28) = "�ɮ��Y"
		MsgList(29) = "�l PE �ɮ�"
		MsgList(30) = "�L�ѷ�"
		MsgList(31) = "���� %v (�c�� %b)\r\n" & _
					"OS ����: Windows XP/2000 �ΥH�W\r\n" & _
					"Passolo ����: Passolo 5.0 �ΥH�W\r\n" & _
					"���v: �K�O�n��\r\n" & _
					"���}: http://www.hanzify.org\r\n" & _
					"�@��: wanfu (2015 - 2019)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(32) = "����ѷӷj��"
		MsgList(33) = "�i�����ɮ� (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|�Ҧ��ɮ� (*.*)|*.*||"
		MsgList(34) = "����ɮ�"
		MsgList(35) = "���b�j���ѷ�, �еy��..."
		MsgList(36) = "�ɮ�: %p\r\nOffset: %o, RVA: %r\r\n%s"

		MsgList(37) = "================================================"
		MsgList(38) = "�Ǹ�\tDec ��}\tHex ��}\t�Ϭq\t�ؿ����\t�ѷӥN�X"
		MsgList(39) = "#%no\t%da\t%ha\t%sc\t%dc\t%rc"
		MsgList(40) = "��J����}�D�k�C"
		MsgList(41) = "��J����}�W�L�ɮת����סC"
		MsgList(42) = "�d��~"
		MsgList(43) = "�ץX�ؿ�"
		MsgList(44) = "�פJ�ؿ�"
		MsgList(45) = "�귽"
		MsgList(46) = "���`"
		MsgList(47) = "�w��"
		MsgList(48) = "���w��򥻪�"
		MsgList(49) = "�E�_"
		MsgList(50) = "���v"
		MsgList(51) = "������"
		MsgList(52) = "����������s�x"
		MsgList(53) = "���J�]�w�ؿ�"
		MsgList(54) = "�j�w��J��"
		MsgList(55) = "�פJ��}��"
		MsgList(56) = "����פJ"
		MsgList(57) = "COM �y�z��"
		MsgList(58) = "�O�d"
		MsgList(59) = "�T��"
		MsgList(60) = "�Ы��r�ꪺ��l��}�j���ѷ�, �M���J�s��}�H�p��s���ѷӥN�X�C"
		MsgList(61) = "�L�ѷӪ��r��, �L�k�p��s��}���ѷӥN�X�C"
		MsgList(62) = "����"
		MsgList(63) = "�y��"
		MsgList(64) = "�^�y;²�餤��;���餤��"
		MsgList(65) = "enu;chs;cht"
		MsgList(66) = "�ɮװT��"
		MsgList(89) = "�T�� - %s"

		MsgList(91) = "============ �ɮװT�� ============\r\n"
		MsgList(92) = "�ɮצW�١G\t%s"
		MsgList(93) = "�ɮ׸��|�G\t%s"
		MsgList(94) = "�ɮ׻����G\t%s"
		MsgList(95) = "�ɮת����G\t%s"
		MsgList(96) = "���~�W�١G\t%s"
		MsgList(97) = "���~�����G\t%s"
		MsgList(98) = "���v�Ҧ��G\t%s"
		MsgList(99) = "�ɮפj�p�G\t%s �줸��"
		MsgList(100) = "�إߤ���G\t%s"
		MsgList(101) = "�ק����G\t%s"
		MsgList(102) = "�y�@�@���G\t%s"
		MsgList(103) = "�} �o �ӡG\t%s"
		MsgList(104) = "��l�ɮצW�G\t%s"
		MsgList(105) = "�����ɮצW�G\t%s"
		MsgList(106) = "�ɮ������G\t%s"
		MsgList(107) = "�M����}�G\t%s"
		MsgList(108) = "�Ϭq�T���G"
		MsgList(109) = "��}���O\t�Ϭq�W\t�}�l��}\t������}\t�줸�դj�p"
		MsgList(110) = "%s!1!\t%s!2!\t%s!4!\t%s!5!\t%s!6!"
		MsgList(111) = "================================="
		MsgList(112) = "�ɮװ�����}"
		MsgList(113) = "�۹������}"
		MsgList(114) = "���N"
		MsgList(115) = "����"
		MsgList(116) = "����"
		MsgList(117) = "���i��"
		MsgList(118) = "��ƥؿ��T�� (�ɮװ�����})�G"
		MsgList(119) = "�ؿ��W��\t�Ҧb�Ϭq\t�}�l��}\t������}\t�줸�դj�p"
		MsgList(120) = "�ץX�ؿ�"
		MsgList(121) = "�פJ�ؿ�"
		MsgList(122) = "�귽�ؿ�"
		MsgList(123) = "���`�ؿ�"
		MsgList(124) = "�w���ؿ�"
		MsgList(125) = "��}���w���"
		MsgList(126) = "�E�_�ؿ�"
		MsgList(127) = "���v�ؿ�"
		MsgList(128) = "������(GP RVA)"
		MsgList(129) = "����������s�x��"
		MsgList(130) = "���J�]�w�ؿ�"
		MsgList(131) = "�j�w�פJ�ؿ�"
		MsgList(132) = "�פJ��}��"
		MsgList(133) = "������J�פJ��"
		MsgList(134) = "COM ����w�X��"
		MsgList(135) = "�O�d�ؿ�"
		MsgList(136) = "���`"
		MsgList(137) = "���s�b"
		MsgList(138) = ".NET CLR ��ƥؿ��T�� (�ɮװ�����})�G"
		MsgList(139) = "�ؿ��W��\t�Ҧb�Ϭq\t�}�l��}\t������}\t�줸�դj�p"
		MsgList(140) = "���~���(MetaData)"
		MsgList(141) = "���޸귽"
		MsgList(142) = "�j�W��ñ�W"
		MsgList(143) = "�N�X�޲z��"
		MsgList(144) = "������(V-��)"
		MsgList(145) = "���D�ץX��}��"
		MsgList(146) = "���ޥ����M���Y"
		MsgList(147) = ".NET MetaData ��Ƭy�T�� (�ɮװ�����})�G"
		MsgList(148) = "��Ƭy�W��\t�Ҧb�Ϭq\t�}�l��}\t������}\t�줸�դj�p"
		MsgList(149) = "�D PE �ɮ�"
		MsgList(150) = "�lPE(%s)"
		MsgList(151) = "��}���O\t�q�W\t�`�W\t\t�}�l��}\t������}\t�줸�դj�p"
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

		MsgList(12) = "Reference Search - Version %v (Build %b)"
		MsgList(13) = "..."
		MsgList(14) = "Dec"
		MsgList(15) = "Hex"
		MsgList(16) = "Type"
		MsgList(17) = "About"
		MsgList(18) = "RVA"
		MsgList(19) = "Offset"
		MsgList(20) = "Code Calculate"
		MsgList(21) = "Search"
		MsgList(22) = "References (%s):"
		MsgList(23) = "Copy"

		MsgList(24) = "Unknown;Not PE;PE32;PE64;MAC32;MAC64"
		MsgList(25) = "No Section"
		MsgList(26) = "Hide Section"
		MsgList(27) = "Outside File"
		MsgList(28) = "File Header"
		MsgList(29) = "Sub PE File"
		MsgList(30) = "No References"
		MsgList(31) = "Version: %v (Build %b)\r\n" & _
					"OS Version: Windows XP/2000 or higher\r\n" & _
					"Passolo Version: Passolo 5.0 or higher\r\n" & _
					"License: Freeware\r\n" & _
					"HomePage: http://www.hanzify.org\r\n" & _
					"Author: wanfu (2015 - 2019)\r\n" & _
					"E-mail: z_shangyi@163.com"
		MsgList(32) = "About Reference Search"
		MsgList(33) = "Executable file (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|All file (*.*)|*.*||"
		MsgList(34) = "Select file"
		MsgList(35) = "Searching references, please wait..."
		MsgList(36) = "File: %p\r\nOffset: %o, RVA: %r\r\n%s"

		MsgList(37) = "================================================"
		MsgList(38) = "No.\tDec Offset\tHex Offset\tSection\tDirectory\tRef Code"
		MsgList(39) = "#%no\t%da\t%ha\t%sc\t%dc\t%rc"
		MsgList(40) = "Address entered is illegal."
		MsgList(41) = "Address entered exceeds the length of this file."
		MsgList(42) = "Outside"
		MsgList(43) = "Export"			'����Ŀ¼
		MsgList(44) = "Import"			'����Ŀ¼
		MsgList(45) = "Resource"		'��ԴĿ¼
		MsgList(46) = "Exception"		'�쳣Ŀ¼
		MsgList(47) = "Security"		'��ȫĿ¼
		MsgList(48) = "Basereloc"		'�ض�λ������
		MsgList(49) = "Debug"			'����Ŀ¼
		MsgList(50) = "Copyright"		'X86ʹ��-��������
		MsgList(51) = "Globalptr"		'����ֵ(Mips GP),�� RVA Of Globalptr
		MsgList(52) = "TLS"				'�̱߳��ش洢(Thread Local Storage,Tls)Ŀ¼
		MsgList(53) = "Load Config"		'��������Ŀ¼
		MsgList(54) = "Bound Import"	'�������(Bound Import Directory in Headers)
		MsgList(55) = "IAT"				'�����ַ��
		MsgList(56) = "Delay Import"	'Delay Load Import Descriptors
		MsgList(57) = "COM Descriptor"	'COM ���б�־
		MsgList(58) = "Reserved"		'����
		MsgList(59) = "Message"
		MsgList(60) = "Please search for reference by the original address of the string, and then enter a new address to calculate the new reference code."
		MsgList(61) = "String is not referenced, reference code for the new address cannot be calculated."
		MsgList(62) = "Cancel"
		MsgList(63) = "Language"
		MsgList(64) = "EngLish;Chinese Simplified;Chinese Traditional"
		MsgList(65) = "enu;chs;cht"
		MsgList(66) = "File Info"
		MsgList(89) = "Information - %s"

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
		Case "cht","0404","0C04","1404"
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
