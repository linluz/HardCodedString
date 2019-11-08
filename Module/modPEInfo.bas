Attribute VB_Name = "modPEInfo"
'' File Information Module for PSlHardCodedString.bas
'' (c) 2015-2019 by wanfu (Last modified on 2019.11.08)

'#Uses "modCommon.bas"

Option Explicit

Private Const PE_BIT_TYPE32 = 224 + 24
Private Const PE_BIT_TYPE64 = 240 + 24

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
	DataDirectory(15) 				As IMAGE_DATA_DIRECTORY
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

'***************************************************************************************************************************************************
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

Private DosHeader 			As IMAGE_DOS_HEADER
Private FileHeader			As IMAGE_FILE_HEADER
Private OptionalHeader32	As IMAGE_OPTIONAL_HEADER32
Private OptionalHeader64	As IMAGE_OPTIONAL_HEADER64
Private SecHeader() 		As IMAGE_SECTION_HEADER


'��ȡ�ļ��汾��Ϣ
Public Function GetFileInfo(ByVal strFilePath As String,File As FILE_PROPERTIE) As Boolean
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


'��ȡ�ļ������ļ������ݽṹ��Ϣ
Public Function GetPEHeaders(ByVal strFilePath As String,File As FILE_PROPERTIE,ByVal Mode As Long) As Boolean
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

		'��������
		'tmpSecHeader(.MaxSecID).lSizeOfRawData = Alignment(tmpSecHeader(.MaxSecID).lSizeOfRawData,.FileAlign,1)
		'tmpSecHeader(.MaxSecID).lVirtualSize = Alignment(tmpSecHeader(.MaxSecID).lVirtualSize,.SecAlign,1)
		'.SecList(.MaxSecID).lSizeOfRawData = tmpSecHeader(.MaxSecID).lSizeOfRawData
		'.SecList(.MaxSecID).lVirtualSize = tmpSecHeader(.MaxSecID).lVirtualSize

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


'�����ļ��ڳ���
'fType = 0 ��ȡ�����ӵ��ֽ������޸�ԭʼ�ļ�ͷ��д��
'fType = 1 д��ָ�����ȶ������ֽ��������޸�ԭʼ�ļ�ͷ��д��
'fType = 2 д��ָ�����ȶ������ֽ������޸�ԭʼ�ļ�ͷ
'fType = 3 ���޸Ĳ�д�룬����ȡ������ֵ(AddSecSize(x).Length Ϊƫ�ƴ�С��AddSecSize(x).Address Ϊ�����С)
'AddSecSize(x).Length = 0���������ӵ����ֵ���ӣ����� AddSecSize(x).Length ����ֵ����
Public Function AddPESectionSize(trnFile As FILE_PROPERTIE,AddSecSize() As FREE_BTYE_SPACE,ByVal fType As Long,ByVal Mode As Long) As Long
	Dim i As Long,j As Integer,k As Long,x As Long,n As Long
	Dim AddRAW As Long,AddRVA As Long,PEBitType As Long
	Dim FN As FILE_IMAGE,TempByte() As Byte,File As FILE_PROPERTIE

	'������
	On Error GoTo localError

	'��ȡ PE ͷ
	File = trnFile
	If GetPEHeaders(File.FilePath,File,Mode) = False Then
		If RefTypeList(0).sName = "" Then Exit Function
		File.Magic = RefTypeList(0).FileMagic
	End If

	'�޸��ļ��ڵĿ�ʼ��ַ�ʹ�С
	With File
		'�޸��ļ�����ֵ���Լ��������ڵĴ�С
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
			'�����Ƿ������ڻ�ȡ�������ֽ�
			If j = .MaxSecID Then
				k = .SecList(j).lSizeOfRawData + AddSecSize(i).Length
				n = 1
			Else
				k = .SecList(GetSectionID(File,j,-1,True)).lVirtualAddress - .SecList(j).lVirtualAddress
				n = 0
			End If
			'���ļ�����ֵ����
			x = Alignment(k,.FileAlign,n) - .SecList(j).lSizeOfRawData
			If x > 0 Or n > 0 Then
				'����ʵ����Ҫ���ӵ�ǰ�ڵ�ƫ�ƴ�С������
				If AddSecSize(i).Length > 0 Then
					x = Alignment(IIf(x > AddSecSize(i).Length,AddSecSize(i).Length,x),.FileAlign,1)
				Else
					x = Alignment(x,.FileAlign,n)
				End If
				AddSecSize(i).Length = x: AddRAW = AddRAW + x

				'���ӵ�ǰ�ڵ������С�������С���ö���
				If AddSecSize(i).Length > 0 Then
					x = Alignment(k,.SecAlign,n) - .SecList(j).lVirtualSize
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
		AddPESectionSize = AddRAW
		Exit Function
	End If

	'���ļ�
	Mode = LoadFile(File.FilePath,FN,0,1,0,Mode)
	If Mode < -1 Then GoTo localError

	'�޸� OptionalHeader ���ݲ�д��
	Select Case File.Magic
	Case "PE32","NET32"
		PEBitType = PE_BIT_TYPE32
		If AddRVA > 0 Then
			'�޸��ļ�����ֵ���Լ����ļ���С
			OptionalHeader32.lFileAlignment = File.FileAlign
			'�޸��ļ�ͷ��ӳ���С���ڶ���
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
			'�޸��ļ�����ֵ���Լ����ļ���С
			OptionalHeader64.lFileAlignment = File.FileAlign
			'�޸��ļ�ͷ��ӳ���С���ڶ���
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

	'�޸� SecHeader ���ݲ�д��
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

	'�޸�Ŀ���ļ�����������
	If fType = 2 Then trnFile = File
	Exit Function

	'��ȫ�˳�����
	localError:
	UnLoadFile(FN,0,Mode)
	AddPESectionSize = -1
End Function


'���ļ�β������һ���ļ���
'fType = 0 ֻ�޸��ļ������ݲ�д��
'fType = 1 ���޸��ļ������ݵ�д��
'fType = 2 ���޸��ļ���������д��
Public Function AddPESection(trnFile As FILE_PROPERTIE,AddSecSize As FREE_BTYE_SPACE,ByVal SecName As String,ByVal fType As Long,ByVal Mode As Long) As Long
	Dim i As Long,k As Long
	Dim OrgHeadersOffset As Long,NewHeadersOffset As Long,FristSecOffset As Long,PEBitType As Long
	Dim FN As FILE_IMAGE,TempByte() As Byte,File As FILE_PROPERTIE

	'������
	On Error GoTo localError
	If AddSecSize.Length = 0 Then Exit Function

	'��ȡ PE ͷ
	File = trnFile
	If GetPEHeaders(File.FilePath,File,Mode) = False Then Exit Function

	'�޸��ļ�����ֵ���Լ��������ڵĴ�С
	If Selected(12) = "1" Then
		If File.FileAlign > 512 Then
			If (File.FileAlign Mod 512) = 0 Then File.FileAlign = 512
		End If
	End If

	'���� NT ͷ�Ĵ�С��ӳ���С
	FristSecOffset = SecHeader(File.MinSecID).lPointerToRawData
	Select Case File.Magic
	Case "PE32", "NET32"
		'�����ļ�ͷʵ��ռ�ô�С
		OrgHeadersOffset = DosHeader.lPointerToPEHeader + Len(FileHeader) + Len(OptionalHeader32) + Len(SecHeader(0)) * File.MaxSecIndex
		'�����ļ�ͷ���ں�����û����������
		For i = FristSecOffset - 1 To 0 Step -1
			If GetByte(FN,i,Mode) > 0 Then
				k = i + 1
				Exit For
			End If
		Next i
		'���������ڱ����Ҫ���ļ�ͷ��С
		If k > OrgHeadersOffset Then
			NewHeadersOffset = k + Len(SecHeader(0))
			If NewHeadersOffset > File.SecAlign Then
				NewHeadersOffset = OrgHeadersOffset + Len(SecHeader(0))
				k = -1
			End If
		Else
			NewHeadersOffset = OrgHeadersOffset + Len(SecHeader(0))
		End If
		'���ļ�ͷ��С�����ڶ���ֵʱȡ�������½�
		If NewHeadersOffset > File.SecAlign Then GoTo localError
		'���ļ�����ֵ������Ҫ���ļ�ͷ��С
		If NewHeadersOffset > FristSecOffset Then
			NewHeadersOffset = Alignment(NewHeadersOffset,File.FileAlign,1)
			OptionalHeader32.lSizeOfHeaders = NewHeadersOffset
		End If
	Case "PE64", "NET64"
		'�����ļ�ͷʵ��ռ�ô�С
		OrgHeadersOffset = DosHeader.lPointerToPEHeader + Len(FileHeader) + Len(OptionalHeader64) + Len(SecHeader(0)) * File.MaxSecIndex
		'�����ļ�ͷ���ں�����û����������
		For i = FristSecOffset - 1 To 0 Step -1
			If GetByte(FN,i,Mode) > 0 Then
				k = i + 1
				Exit For
			End If
		Next i
		'���������ڱ����Ҫ���ļ�ͷ��С
		If k > OrgHeadersOffset Then
			NewHeadersOffset = k + Len(SecHeader(0))
			If NewHeadersOffset > File.SecAlign Then
				NewHeadersOffset = OrgHeadersOffset + Len(SecHeader(0))
				k = -1
			End If
		Else
			NewHeadersOffset = OrgHeadersOffset + Len(SecHeader(0))
		End If
		'���ļ�ͷ��С�����ڶ���ֵʱȡ�������½�
		If NewHeadersOffset > File.SecAlign Then GoTo localError
		'���ļ�����ֵ������Ҫ���ļ�ͷ��С
		If NewHeadersOffset > FristSecOffset Then
			NewHeadersOffset = Alignment(NewHeadersOffset,File.FileAlign,1)
			OptionalHeader64.lSizeOfHeaders = NewHeadersOffset
		End If
	End Select
	'���ӽ���
	FileHeader.iNumberOfSections = FileHeader.iNumberOfSections + 1

	'������µ�������Ϣ���ʿ��ܻ������ļ�ͷ
	If NewHeadersOffset > FristSecOffset Then
		For i = 0 To File.MaxSecIndex - 1
			If File.SecList(i).lPointerToRawData > 0 Then
				File.SecList(i).lPointerToRawData = File.SecList(i).lPointerToRawData + NewHeadersOffset - FristSecOffset
			End If
		Next i
	End If

	'���������ڵ�ַ����С��������λ����
	With File
		.SecList(.MaxSecIndex).sName = SecName
		.SecList(.MaxSecIndex).lPointerToRawData = .SecList(.MaxSecID).lPointerToRawData + Alignment(.SecList(.MaxSecID).lSizeOfRawData,.FileAlign,1)
		.SecList(.MaxSecIndex).lSizeOfRawData = Alignment(AddSecSize.Length,.FileAlign,1)
		.SecList(.MaxSecIndex).lVirtualAddress = .SecList(.MaxSecID).lVirtualAddress + Alignment(.SecList(.MaxSecID).lVirtualSize,.SecAlign,1)
		.SecList(.MaxSecIndex).lVirtualSize = AddSecSize.Length
		.SecList(.MaxSecIndex).SubSecs = 0
		.SecList(.MaxSecIndex).RWA = .SecList(.MaxSecIndex).lPointerToRawData
	End With

	'�޸����ؽڵ�ƫ�Ƶ�ַ����ԭ���ؽڴ�С
	'�������¶�ȡ���ļ����������ļ�β��д�����ִ�����Щ�ֽڽ���Ϊ���ؽڱ���ȡ������Ҫʹ��ԭʼ�ļ������ؽ���Ϣ
	ReDim Preserve File.SecList(File.MaxSecIndex + 1) 'As SECTION_PROPERTIE
	ReDim Preserve File.SecList(File.MaxSecIndex + 1).SubSecList(0) 'As SUB_SECTION_PROPERTIE
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
		AddPESection = File.SecList(File.MaxSecIndex).lSizeOfRawData
		If fType = 0 Then
			File.MaxSecID = File.MaxSecIndex: File.MaxSecIndex = File.MaxSecIndex + 1: trnFile = File
		End If
		Exit Function
	End	If

	'���ļ�
	Mode = LoadFile(File.FilePath,FN,0,1,0,Mode)
	If Mode < -1 Then GoTo localError

	'��λԭ�ļ������ؽڼ����ؽں��ȫ���ֽڳ���
	'�������¶�ȡ���ļ����������ļ�β��д�����ִ�����Щ�ֽڽ���Ϊ���ؽڱ���ȡ������Ҫʹ��ԭʼ�ļ������ؽ���Ϣ
	'����ʹ�������С����Ϊ���СΪ���ں�������ֽ�(������ PE)
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

	'���� FristSecOffset ������ݵ� NewSizeOfHeaders
	If NewHeadersOffset > FristSecOffset Then
		'���� FristSecOffset ��ԭ�ļ�������ε�ַ�� NewHeadersOffset
		i = trnFile.SecList(trnFile.MaxSecID).lPointerToRawData + _
			trnFile.SecList(trnFile.MaxSecID).lSizeOfRawData - FristSecOffset
		TempByte = GetBytes(FN,i,FristSecOffset,Mode)
		PutBytes(FN,NewHeadersOffset,TempByte,i,Mode)
		'��� FristSecOffset �� NewSizeOfHeaders ֮�������
		i = NewHeadersOffset - FristSecOffset
		ReDim TempByte(i - 1) As Byte
		PutBytes(FN,FristSecOffset,TempByte,i,Mode)
	End If

	'�ڱ������Ķ����ֽڴ��� (һ��ϵ�ѿ�����)
	If k > OrgHeadersOffset Then
		'�ƶ� OrgHeadersOffset �� k ֮����ֽ�Ϊһ�������ݴ�С
		TempByte = GetBytes(FN,k - OrgHeadersOffset,OrgHeadersOffset,Mode)
		PutBytes(FN,OrgHeadersOffset + Len(SecHeader(0)),TempByte,k - OrgHeadersOffset,Mode)
	ElseIf k < 0 Then
		'�ÿսڱ������ԭʼ��β֮��������ֽ�
		ReDim TempByte(FristSecOffset - OrgHeadersOffset - 1) As Byte
		PutBytes(FN,OrgHeadersOffset,TempByte,FristSecOffset - OrgHeadersOffset,Mode)
	End If

	'д�����޸ĵ� FileHeader ����
	'If PutTypeValue(FN,.DosHeader.lPointerToPEHeader,FileHeader,Mode) = False Then GoTo localError
	Select Case Mode
	Case Is < 0
		Put #FN.hFile,DosHeader.lPointerToPEHeader + 1,FileHeader
	Case 0
		CopyMemory FN.ImageByte(DosHeader.lPointerToPEHeader),FileHeader,Len(FileHeader)
	Case Else
		WriteMemory FN.MappedAddress + DosHeader.lPointerToPEHeader,FileHeader,Len(FileHeader)
	End Select

	'�޸� OptionalHeader ���ݲ�д��
	Select Case File.Magic
	Case "PE32","NET32"
		PEBitType = PE_BIT_TYPE32
		'�޸��ļ�����ֵ���Լ����ļ���С
		OptionalHeader32.lFileAlignment = File.FileAlign
		'�޸��ļ�ͷ��ӳ���С���ڶ���
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
		'�޸��ļ�����ֵ���Լ����ļ���С
		OptionalHeader64.lFileAlignment = File.FileAlign
		'�޸��ļ�ͷ��ӳ���С���ڶ���
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

	'���������ڵ�ַ����С������д��
	If NewHeadersOffset > FristSecOffset Then
		For i = 0 To File.MaxSecIndex - 1
			SecHeader(i).lPointerToRawData = File.SecList(i).lPointerToRawData
			SecHeader(i).lSizeOfRawData = File.SecList(i).lSizeOfRawData
			SecHeader(i).lVirtualSize = File.SecList(i).lVirtualSize
		Next i
	End If
	'���������ε�����
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
	AddPESection = File.SecList(File.MaxSecIndex).lSizeOfRawData

	'�޸�Ŀ���ļ�����������
	If fType = 2 Then
		File.MaxSecID = File.MaxSecIndex: File.MaxSecIndex = File.MaxSecIndex + 1: trnFile = File
	End If
	Exit Function

	'��ȫ�˳�����
	localError:
	UnLoadFile(FN,0,Mode)
	AddPESection = -1
End Function
