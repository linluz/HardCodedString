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
			'0=����ʮ������ת���,1=���� Unicode ת���,2=���ڰ˽���ת���
Public Const CheckStrHexRange = "\x00-\x07,\x0E-\x1F;" & _
								"\x00-\x40,\x5B-\x60,\x7B-\xBF;" & _
								"\x00-\x60,\x7B-\xBF;" & _
								"\x00-\x40,\x5B-\xBF;" & _
								"\x41-\x5A,\x61-\x7A;" & _
								"\x30-\x39,\x41-\x5A,\x61-\x7A;" & _
								"\x30-\x39,\x41-\x46"
			'0=�����ַ�,1=ȫΪ���ֺͷ���,2=ȫΪ��дӢ��,3=ȫΪСдӢ��,4=��Сд���Ӣ��,5=��ݼ��ַ���Χ,6=ʮ�������ַ�
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

'�ļ�����
Private Type HCS_FILE
	SourceFileName		As String	'�����ļ�����Դ�ļ���
	SourceFileVersion	As String	'�����ļ�����Դ�ļ��汾
	SourceFileSize		As Long		'�����ļ�����Դ�ļ���С
	SourceFileDateTime	As Date		'�����ļ�����Դ�ļ��޸�����
	SourceFileLangID	As String	'�����ļ�����Դ�ļ�����ID
	AppVersion			As String	'���������ļ����õĺ�汾
	DateLastModified	As Date		'�����ļ����޸�����
	FilePath			As String	'�����ļ���·�� (���ļ���)
	FileFormat			As Boolean	'�����ļ��ĸ�ʽ��Ture Ϊ����Ҫ��
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
Public Type FILE_PROPERTIE
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
	hcsFile				As HCS_FILE
End Type

'�Զ�������
Public Type REF_TYPE
	sName				As String	'�㷨����
	Algorithm 			As String	'���ô�����㷨��ʽ
	ByteLength			As Integer	'���õ��ֽڳ���
	ByteOrder			As Integer	'�ֽ���0 = С�ˣ�1 = ���
	FileMagic			As String	'�������ͣ�NotPE32,NotPE64
	PrefixByte			As String	'ǰ׺�ֽڣ���ʮ����������
	PrefixLength		As Integer	'ǰ׺�ֽڳ��ȣ�������ֵ�Զ�����
	StrAddAlgorithm 	As String	'�����ô�������ִ���ַ���㷨��ʽ
	Template			As String	'���浽�ִ������ļ�������
End Type

'PE�ļ��ṹ(Visual Basic��)���ִ���һ
'ǩ������
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

'�����д���Ի�ƽ̨����
Public Enum PELangType
	DELPHI_FILE_SIGNATURE = &H50
	NET_FILE_SIGNATURE = &H424A5342
End Enum

Private Type IMAGE_DATA_DIRECTORY
	lVirtualAddress		As Long		'��ʼRVA��ַ
	lSize				As Long		'lVirtualAddress��ָ�����ݽṹ���ֽ���
End Type

'���ֽ�����
Public Type FREE_BTYE_SPACE
	Address				As Long		'���ÿ�ʼ��ַ
	Length				As Long		'������󳤶�
	MaxAddress			As Long		'�ַ��������ַ
	inSectionID			As Integer	'�ִ����ڽڵ�������
	inSubSecID			As Integer	'�ִ������ӽڵ�������
	lNumber				As Long		'���õ�ַ�����ţ�<-1 = ���ִ��ռ�������>-1 = �ִ��ռ�ID��
	MoveType			As Long		'����״̬, -3 = ����ռ�ã�-2 = ��β�ռ䣬-1 = δ��ռ�ã�>-1 = �ѱ�ռ��(�ִ�ID��)
	IndexList()			As Long		'��ͬID�������б�
End Type

'������Ϣ����
Public Type PROGRESS_MSG
	Total				As Long		'�ܴ�����
	Passed				As Long		'�Ѵ�����
	Massage				As String	'������Ϣ
	hwnd				As Long		'��ʾ��Ϣ�ĶԻ���ؼ����
End Type

'�༭��������
Public Type TOOLS_PROPERTIE
	sName				As String	'��������
	FilePath			As String	'�����ļ�·��(���ļ���)
	Argument			As String	'���в���
End Type

'����ҳ����
Public Type CODEPAGE_DATA
	sName				As String	'��������
	CharSet				As String	'�ַ�����
End Type

'����������
Public Type FILTER_PROPERTIE
	Item				As String	'������Ŀ
	Value				As String	'�ж�ֵ
	Mode				As Integer	'�ж�����
	Other				As String	'����ֵ
End Type

'�����ļ�
Public Type UI_FILE
	FilePath			As String	'�����ļ���ȫ·��
	AppName				As String	'��������
	Version				As String	'����汾
	LangName			As String	'��������
	LangID				As String	'��������ID
	Encoding			As String	'�ַ�����
End Type

'INI �ļ�
Public Type INIFILE_DATA
	Title				As String	'����
	Item()				As String	'��Ŀ
	Value()				As String	'�ִ�ֵ
End Type

'����
Public Type REFERENCE_PROPERTIE
	sCode				As String	'���ô���
	lAddress			As Long		'���õ�ַ�б�
	inSecID				As Integer	'�ִ����ڽڵ�������
	StrType				As Integer	'�������ô�����ִ�����
	Index               As Long		'���õ������ţ��� 1 Ϊ��ʼ�����������ִ��ķ����и��������ţ��������ò��
End Type

'�ִ�������
Public Type STRING_SUB_PROPERTIE
	lStartAddress		As Long		'�ִ��Ŀ�ʼ��ַ
	lEndAddress			As Long		'�ִ��Ľ�����ַ
	lMaxAddress			As Long		'�ַ��������ַ
	lCharLength			As Long		'�ִ����ַ�����
	lHexLength			As Long		'�ִ���ʮ�����Ƴ���
	lMaxHexLength		As Long		'�ִ����������ʮ�����Ƴ���
	CodePage			As Long		'�ִ��Ĵ���ҳ
	sString				As String	'�ַ���
	inSectionID			As Integer	'�ִ����ڽڵ�������
	inSubSecID			As Integer	'�ִ����ڽڵ��ӽ�������
	StrTypeLength		As Integer	'�ִ����͵�ǰ�ñ�ʶ���ֽ���
	MoveLength			As Integer	'�ɱ����ִ����ȱ�ʶ�����ִ����ͳ��ȸ���ֵ, 0=�޸���, >0=����ֵ, <0=����ֵ
									'ԭʼ�ִ��ĸ�ֵ = �����ִ� - ԭʼ�ִ��������ִ��ĸ�ֵ = �ضϺ� - ԭʼ����
	iNullByteLength		As Integer	'�ִ���ҪԤ���Ŀ��ֽڳ���(Unicode = 2����������ҳ = 1��Pascal Short = 0)
	lReferenceNum		As Long		'���ô���
	GetRefState			As Integer	'��ȡ�ִ������б��״̬��0 = δ��ȡ��1 = �ѻ�ȡ��0=����
	NewStemp			As Integer	'�ִ���������ǣ����ڸ��ӵ���ʱ���Թ�����ʾ�������ִ���ԭʼ�ִ���0=��ȡ��1=���룬2=��ӣ������ִ���1=����
	Reference()			As REFERENCE_PROPERTIE
End Type

'�ִ�����
Public Type STRING_PROPERTIE
	ID						As Long		'�ִ����(ID��), <0 ���õ�ַ(���ִ�)��>-1 �ִ���ʼ��ַ(���ִ�)
	StrType					As Integer	'�ִ�����, <-4 Android��<0=.NET��0=Ĭ�ϣ�1=Pascal Unicode, 2=Pascal Wide, 3=Pascal Ansi, 4=Pascal Short, >4=�Զ���
	WriteType				As Integer	'д������
										'PE �ļ�: -2=��ʧ(�ѷ���)��-1=��ʧ(δ����)��0=δ���룬1=ԭַ����д�룬2=ԭַ�ض�д�룬3=ȫ����λд�룬4=������λд�룬5=ԭַ����д��
										'�� PE �ļ�: -2=��ʧ(�ѷ���)��-1=��ʧ(δ����)��0=δ���룬1=ԭ������д�룬2=�����д�룬3=��������д�룬4=ԭַ�ض�д��
	WriteState				As Integer	'д��״̬, 0=δд�룬1=����д�룬2=�ض�д�룬3=�ִ�д��ʧ�ܣ�4=���ô����޸�ʧ�ܣ�5=��ʶ���޸�ʧ��
	LockState				As Integer	'�����ִ�������״̬��0=δ������1=������
	ScapeIDBeMoved			As Long		'��ռ�õ�������λ�ִ�ID��: -1=û�б�ռ�ã�>-1=ռ�ø��ִ���ַ��������λ�ִ����ִ�ID��
	ScapeIDForMove			As Long		'��λд����ʹ�õĵ�ַID��: -1=û���ƶ���>-1=������λ�ִ����ִ������ţ�<-1=�ִ��б�����Ŀ����ַID��
	SourceStringClearState	As Integer	'�ִ�ԭַ�����״̬��0=δ��գ�1=�����
	EndByteLength			As Integer	'�ִ����͵ĺ����ʶ���ֽ���
	GetLengthState			As Integer	'�ִ����ͳ��ȵĻ�ȡ״̬��0 = δ��ȡ��1 = �ѻ�ȡ
	MoveType				As Integer	'�ƶ�����, 0=����λ, 1=�ִ���λ(�ִ�ID), 2=���ִ���λ(������), 3=�ں�ԭ�п�λ(-2)
												  '4=�ִ����ڽں���չ��λ(-3), 5=������չ��λ(-��ʼ��ַ), 6=������(-��ʼ��ַ)
	MoveMode				As Integer	'��λģʽ��0=������λ(δ��λʱ���õ�ַΪ��), 1=ǿ����λ(δ��λʱ���õ�ַ�ǿ�),
												  '2=ԭַ��λ(�ִ������ֽڳ��ȸı�), 3=�ֶ���λ��4=����λ����չ��5=ԭַ��չ
	SplitState				As Long		'���״̬, 0=δ���, >0=�Ѳ�ֵĸ���(���ֱ�ʾ�Ӵ���), <0=�Ӵ�(���ֱ�ʾ�丸��ID)
	TagType					As Integer	'��ʶ������, 0=�ޱ�ʶ��, 1=�����ִ�, 2=�������ô���
	LengthModeID			As Integer	'��⵽�ĸ����ִ����Զ����ִ����͵��ַ������ȱ�ʶ�����ִ����ȼ�������
	FillLength				As Integer	'����ԭʼ���ȵķ����ִ�����ո�Ŀո��ֽڳ���, <0=ǰ�˿ո�����ֵ, 0=�޲���, >0=��˿ո�����ֵ
	iError					As Integer	'�����ִ��б�ʱ�Ĵ���������PSL �ı���������֧�ִ����� 16k ���ϵ��ִ���0=����, 1=����, 2=ɾ��
	Moveable				As Integer  '����λ����(�����������޹�), 0=����λ, 1=�� PE �ִ�, 2=.NET ����λ�ִ�, 3=���ؽ��ִ�, 4=��ʧ�ִ�
	OverLengthWrite			As Integer  '����д������ (�ɷ�����ԭ�ִ���Ŀ�λ), 0=����, 1=ȫ������Ϊ��ֹ, 2=ѡ���ִ�����Ϊ��ֹ
	Missing					As Integer  'ԭʼ��Ŀ���ļ��ж�ʧ���ִ���0=δ��ʧ��1=��ʧ
	Source					As STRING_SUB_PROPERTIE
	Trans					As STRING_SUB_PROPERTIE
End Type

'��������
Public Type LANG_PROPERTIE
	LangName					As String	'��������
	LangID						As Long		'����ID
	CPName						As String	'����ҳ����
	CodePage					As Long		'����ҳ����ֵ
	UniCodeRange				As String	'Unidoce�ı��뷶Χ
	UniCodeRegExpPattern		As String	'Unicode�ı��뷶Χ(RegExpģ��)
	UniCodeByteRange			As String	'Unicode���ַ���Χ
	FeatureCode					As String	'�����뷶Χ
	FeatureCodeRegExpPattern	As String	'�����뷶Χ(RegExpģ��)
	FeatureCodeByteRange		As String	'��������ַ���Χ
	FeatureCodeEnable			As Integer  '������������,True Ϊ����,Ĭ�ϲ�����
	dwFlags						As Boolean	'���Ա�־, True = �û���ӵ�����
End Type

'�ֽ�
'Private Type BYTE_PROPERTIE
'	NullValNum		As Long		'&H00�ֽ���
'	CBLValNum		As Long		'&H00-&H7F�ֽ���
'	NullValPos		As Long		'&H00�ֽ�����λ��
'	ByteType		As Long		'�ֽڿ�����
'End Type

Public Type CHECK_STRING_VALUE
	AscRange		As String
	Range			As String
End Type

'�ڴ�ӳ�䷽ʽ�����ִ�
'Private Type SAFEARRAYBOUND
'	cElements		As Long		'һά�ж��ٸ�Ԫ��
'	lLbound			As Long		'������ʼֵ
'End Type

'�ڴ�ӳ�䷽ʽ�����ִ�
'Private Type SAFEARRAYID
'	cDims			As Integer 	'�����ά��
'	fFeatures		As Integer	'���������
'	cbElements		As Long		'�����ÿ��Ԫ�ش�С
'	clocks			As Long		'���鱻�����Ĵ���
'	pvData			As Long		'����������ݴ��λ��
'	rgsabound(0)	As SAFEARRAYBOUND
'End Type

'���ļ���ʽ�Ľṹ��
Public Type FILE_IMAGE
	ModuleName		As String	'�������ļ����ļ���
	hFile			As Long		'���� Create �ļ�ӳ��� OpenFile �ľ��
	hMap			As Long		'���� CreateFileMapping �ļ�ӳ��ľ��
	MappedAddress	As Long		'�ļ�ӳ�䵽���ڴ��ַ
	SizeOfImage		As Long		'ӳ��� Image ���ֽ�����Ĵ�С
	SizeOfFile		As Long		'ʵ���ļ���С
	ImageByte()		As Byte		'�ļ����ֽ�����
End Type

'Delphi �ַ������Ͷ���
'-----------------------------------------------------------------
'1 ShortString		��������255���ַ�,��ҪΪ���ϰ汾����
'2 AnsiString		��������2��31�η����ַ�,D2009ǰĬ�ϵ�String����
'3 UnicodeString	��������2��30�η����ַ�,D2009���Ժ��Ĭ��String����
'4 WideString		��������2��30�η����ַ�,��Ҫ��COM���õıȽ϶�
'-----------------------------------------------------------------
'ShortString
Public Type DELPHI_SHORT_STRING
	Length					As Byte		'�ַ����ֽڳ��ȣ�1���ֽ�
	'Strings()				As Byte		'�ַ������ֽڳ���Ϊ Length * 1
	'EndChar(0)				As Byte		'��һ����&H00����
End Type

'AnsiString
Public Type DELPHI_ANSI_STRING
	RefCount				As Long		'�ַ������ô�����4���ֽ�
	Length					As Long		'�ַ����ֽڳ��ȣ�4���ֽ�
	'Strings()				As Byte		'�ַ������ֽڳ���Ϊ Length * 1
	'EndChar(0)				As Byte		'��&H00����
End Type

'WideString
Public Type DELPHI_WIDE_STRING
	Length					As Long		'�ַ����ֽڳ��ȣ�4���ֽ�
	'Strings()				As Byte		'�ַ������ֽڳ���Ϊ Length * 1
	'EndChar(1)				As Byte		'��&H0000����
End Type

'UnicodeString
Public Type DELPHI_UNICODE_STRING
	CodePage				As Integer	'�ַ�������ҳ��2���ֽ�, ֧�� Unicode, UTF-8, ANSI
	elemSize				As Integer	'ÿ���ַ����ֽ�����2���ֽڣ�Unicode = 2, UTF-8 = 1 or 3, GB2312 = 2
	RefCount				As Long		'�ַ������ô�����4���ֽ�
	Length					As Long		'�ַ������ַ�����4���ֽ�
	'Strings()				As Byte		'�ַ������ֽڳ���Ϊ Length * elemSize
	'EndChar(1)				As Byte		'��&H0000����
End Type

'�Զ����ַ������Ͷ���
Public Type STRING_TYPE
	sName					As String	'�ַ������͵�����
	CodeLoc					As Integer	'�����ַ�����ʶ������λ��, 0 = �ִ�ǰ, 1 = ���õ�ַǰ
	FristCodePos			As Integer	'��һ���ַ�����ʶ��λ���ִ�ǰ��λ��
	CPCodePos				As Integer	'�ַ�������ҳ��ʶ��λ���ִ�ǰ��λ��
	CPCodeSize				As Integer	'�ַ�������ҳ��ʶ�����ֽ���
	CPCodeStartString		As String	'�ַ�������ҳ��ʶ����ʼ���(Hex�ı�)
	CPCodeStartLength		As Integer	'�ַ�������ҳ��ʶ����ʼ���(Hex�ı�)���ֽڳ���
	CPCodeStartByte()		As Byte		'�ַ�������ҳ��ʶ����ʼ���(Hex�ı�)���ֽ�����
	LengthCodePos			As Integer	'�ַ������ȱ�ʶ��λ���ִ�ǰ��λ��
	LengthCodeSize			As Integer	'�ַ������ȱ�ʶ�����ֽ���
	LengthMode				As Integer	'�ַ������ȱ�ʶ�����ִ����ȼ�������
	LengthReviseVal			As Integer	'�ַ������ȱ�ʶ�����ִ����ȵ���ֵ
	ByteLengthReviseVal		As Integer	'�ַ������ȱ�ʶ�����ִ����ȵ���ֵ
	CharLengthReviseVal		As Integer	'�ַ������ȱ�ʶ�����ִ����ȵ���ֵ
	LengthCodeStartString	As String	'�ַ������ȱ�ʶ����ʼ���(Hex�ı�)
	LengthCodeStartLength	As Integer	'�ַ������ȱ�ʶ����ʼ���(Hex�ı�)���ֽ���
	LengthCodeStartByte()	As Byte		'�ַ������ȱ�ʶ����ʼ���(Hex�ı�)���ֽ�����
	StartCodePos			As Integer	'�ַ�����ʼ��ʶ��λ���ִ�ǰ��λ��
	StartCodeString			As String	'�ַ�����ʼ��ʶ��(Hex�ı�)
	StartCodeLength			As Integer	'�ַ�����ʼ��ʶ�����ֽ���
	StartCodeByte()			As Byte		'�ַ�����ʼ��ʶ�����ֽ�����
	EndCodeString			As String	'�ַ���������ʶ��(Hex�ı�)
	EndCodeLength			As Integer	'�ַ���������ʶ�����ֽ���
	EndCodeByte()			As Byte		'�ַ���������ʶ�����ֽ�����
	RefCodeStartPos			As Integer	'���ô��뿪ʼλ�����ô���ǰ��λ��
	RefCodeStartString		As String	'���ô��뿪ʼ���(Hex�ı�)
	RefCodeStartLength		As Integer	'���ô��뿪ʼ���(Hex�ı�)���ֽ���
	RefCodeStartByte()		As Byte		'���ô��뿪ʼ���(Hex�ı�)���ֽ�����
	RegExpPattern()			As String	'������ʽģ��
End Type

'�ִ������ж�ֵ����
Public Type STRING_TYPE_LENGTH
	Pattern					As String	'������ʽģ��
	Length1					As Long		'�׸������ж�ֵ
	Length2					As Long		'�θ������ж�ֵ
	Size					As Long		'���ȴ�С�ж�ֵ
	Bytes()					As Byte		'���ȵ��ֽ�����
	ByteOrder				As Integer	'�ֽ���, -1 = �����ǰ, 0 = С����ǰ, 1 = δ֪
End Type

'�ļ����Ƽ������ļ��ж���
Public Type FILE_LIST
	sName					As String	'�ı�����������
	FilePath				As String	'�����������ļ�·��(���ļ���)
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
	CP_UTF32LE = 65005  	'Unicode (UTF-32 LE)
	CP_UTF32BE = 65006		'Unicode (UTF-32 Big-Endian)
End Enum

'SetLocaleInfo()�� LCTYPE values �ľ�������
Public Enum LOCALE
	LOCALE_ILANGUAGE = &H1			'����ID
	LOCALE_SLANGUAGE = &H2			'�����������ƣ���: "English (United States)"
	LOCALE_SENGLANGUAGE = &H1001	'����Ӣ������
	LOCALE_SABBREVLANGNAME = &H3	'����������д����: "ENU"
	LOCALE_SNATIVELANGNAME = &H4	'�����������ƣ���: "English"
	LOCALE_ICOUNTRY = &H5			'���Ҵ���
	LOCALE_SCOUNTRY = &H6			'���ұ�������
	LOCALE_SENGCOUNTRY = 4098		'����Ӣ������
	LOCALE_SABBREVCTRYNAME = &H7	'����������д
	LOCALE_SNATIVECTRYNAME = &H8	'�������Թ�������
	LOCALE_IDEFAULTLANGUAGE = &H9	'ȱʡ����ID
	LOCALE_IDEFAULTCOUNTRY = &HA	'ȱʡ���Ҵ���
	LOCALE_IDEFAULTCODEPAGE = &HB	'ȱʡ��OEM����
	LOCALE_IDEFAULTANSICODEPAGE = &H1004	'ȱʡ��ASCII����
	LOCALE_IDEFAULTMACCODEPAGE = &H1011		'ȱʡ��MACINTOH����
End Enum

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
Private Declare Function GetACP Lib "kernel32.dll" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" ( _
	ByVal LOCALE As Long, _
	ByVal LCType As Long, _
	ByVal lpLCData As String, _
	ByVal cchData As Long) As Long
	'cchData Ϊ lpLCData �������ĳ��ȣ�����Ϊ�㣬��ʾ��ȡ��Ҫ�Ļ���������

'�ڴ渴�ƺͱȽϺ���
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

'������궨��
Public Type POINTAPI
	x As Long
	y As Long
End Type

'SendMessage API ���ֳ���
Public Enum SendMsgValue
	'��������б����
	CB_GETEDITSEL = &H140			'���,�յ�		����ȡ����Ͽ��������༭���ӿؼ��е�ǰ��ѡ�е��ַ�������ֹλ��
	CB_SETEDITSEL = &h142			'0,0 or -1		����ѡ����Ͽ��������༭���ӿؼ��еĲ����ַ���,��Ӧ����

	'�����ı�����Ҷ�λ����
	EM_GETSEL = &HB0				'0,����			��ȡ���λ�ã��Ա���Ĭ�ϱ�����ַ�����ʾ��
	EM_SETSEL = &HB1				'���,�յ�		���ñ༭�ؼ����ı�ѡ�����ݷ�Χ�������ù��λ�ã������յ��Ϊ�ַ�ֵ
									'				��ָ����������0���յ����-1ʱ���ı�ȫ����ѡ�У��˷���������ձ༭�ؼ�
									'				��ָ����������-2���յ����-1ʱ��ȫ�ľ���ѡ����������ı�δ��
	EM_GETLINECOUNT = &HBA			'0,0			��ȡ�༭�ؼ���������
	EM_LINEINDEX = &HBB				'�к�,0			��ȡָ����(��:-1,0 ��ʾ���������)���ַ����ı��е�λ�ã����ַ�����ʾ��
	EM_LINELENGTH = &HC1			'ƫ��ֵ,0		��ȡָ��λ��������(��:-1,0 ��ʾ��������У����ı����ȣ����ַ�����ʾ��
	EM_LINEFROMCHAR = &HC9			'ƫ��ֵ,0		��ȡָ��λ��(��:-1,0 ��ʾ���λ��)���ڵ��к�
	EM_GETLINE = &HC4				'�к�,ByVal����	��ȡ�༭�ؼ�ĳһ�е����ݣ�������Ԥ�ȸ��ո�
	EM_SCROLLCARET = &HB7			'0,0 			�ѿɼ���Χ������괦
	EM_UNDO = &HC7					'0,0 			����ǰһ�α༭���������ظ����ͱ���Ϣ���ؼ����ڳ����ͻָ��������л�
	EM_REPLACESEL = &HC2			'1(0),�ַ���	��ָ���ַ����滻�༭�ؼ��еĵ�ǰѡ������
									'				�������������wParamΪ1���򱾴β�����������0��ֹ�������ַ������ô�ֵ��ʽ��Ҳ���ô�ַ��ʽ
									'				������SendMessage Text1.hwnd, EM_REPLACESEL, 0, Text2.Text '���Ǵ�ֵ��ʽ��
	EM_GETMODIFY = &HB8				'0,0			�жϱ༭�ؼ��������Ƿ��ѷ����仯������TRUE(1)��ؼ��ı��ѱ��޸ģ�����FALSE(0)��δ��
	EN_CHANGE = &H300 				'				�༭�ؼ������ݷ����ı䡣��EN_UPDATE��ͬ������Ϣ���ڱ༭����ʾ�����ı�ˢ�º�ŷ�����
	EN_UPDATE = &H400				'				�ؼ�׼����ʾ�ı��˵�����ʱ���͸���Ϣ������EN_CHANGE֪ͨ��Ϣ���ƣ�ֻ���������ڸ����ı���ʾ����֮ǰ
	EM_GETLIMITTEXT = &HD5			'0,0			��ȡһ���༭�ؼ����ı�����󳤶�
	EM_LIMITTEXT = &HC5				'���ֵ,0		���ñ༭�ؼ��е�����ı�����
	EM_GETFIRSTVISIBLEINE = &HCE	'0,0			����ı��ؼ��д��ڿɼ�λ�õ�������ı����ڵ��к�
	EM_GETHANDLE = &HBD				'0,0			ȡ���ı�������
	'EM_SETCHARFORMAT = &H444		'��ɫֵ,0		�ı�ѡ���ı�����ɫ

	'�����б����
	'LB_ADDFILE = &H0196			'0,�ļ�����ַ	�����ļ���
	'LB_ADDSTRING = &H0180			'0,�ַ�����ַ	׷��һ���б�������������ָ����LBS_SORT��񣬱���������򣬷��򽫱�׷�����б������һ��
	LB_DELETESTRING = &H0182		'�б������,0	ɾ��ָ�����б�����б��ʣ�N���
	'LB_DIR = &H018D				'DDL_ARCHIVE,ָ��ͨ�����ַ	����ļ����б���������һ����ӵ��ļ���������
	'LB_FINDSTRING = &H018F			'��ʼ�������,�ַ�����ַ	����ƥ���ַ��������Դ�Сд����ָ����ʼ������ſ�ʼ���ң����鵽ĳ������ı��ַ�����ǰ�����ָ�����ַ�����������Ҳ�����ת���б���һ��������ң�ֱ���������б�����wParamΪ-1����б���һ�ʼ���ң�����ҵ��򷵻ر�����ţ����򷵻�LB_ERR���磺�����ַ���Ϊ"abc123"��ָ���ִ�"ABC"����ƥ��
	'LB_FINDSTRINGEXACT = &H01A2	'��ʼ�������,�ַ�����ַ	�����ַ��������Դ�Сд����LB_FINDSTRING��ͬ�����������������ַ�����ͬ������ҵ��򷵻ر�����ţ����򷵻�LB_ERR
	'LB_GETANCHORINDEX = &H019D		'0,0			����������ѡ�е��������
	'LB_GETCARETINDEX = &H019F		'0,0			���ؾ��о��ν�����������
	LB_GETCOUNT = &H018B			'0,0			�����б�������������������򷵻�LB_ERR
	'LB_GETCURSEL = &H0188			'0,0			�������������ڵ�ѡ���б���������ص�ǰ��ѡ��������������û���б��ѡ����д��������򷵻�LB_ERR
	LB_GETHORIZONTALEXTENT = &H0193	'0,0		�����б��Ŀɹ����Ŀ�ȣ����أ�
	'LB_GETITEMDATA = &H0199		'����,0			ÿ���б����һ��32λ�ĸ������ݣ��ú�������ָ���б���ĸ������ݡ���������������LB_ERR
	'LB_GETITEMHEIGHT = &H01A1		'����,0			�����б����ĳһ��ĸ߶ȣ����أ�
	'LB_GETITEMRECT = &H0198		'����,RECT�ṹ��ַ	����б���Ŀͻ�����RECT
	'LB_GETLOCALE = &H01A6			'0,0			ȡ�б��ǰ������������Դ��룬���û�ʹ��LB_ADDSTRING����Ͽ��е��б������Ӽ�¼��ʹ��LBS_SORT��������������ʱ������ʹ�ø����Դ��롣����ֵ�и�16λΪ���Ҵ���
	LB_GETSEL = &H0187				'����,0			����ָ���б����״̬�������ѯ���б��ѡ���ˣ���������һ����ֵ�����򷵻�0���������򷵻�LB_ERR
	LB_GETSELCOUNT = &H0190			'0,0			�����������ڶ���ѡ���б��������ѡ�������Ŀ��������������LB_ERR
	LB_GETSELITEMS = &H0191			'����Ĵ�С,������	�����������ڶ���ѡ���б���������ѡ�е������Ŀ��λ�á�����lParamָ��һ�����������黺�������������ѡ�е��б����������wParam˵�������黺�����Ĵ�С�����������ط��ڻ������е�ѡ�����ʵ����Ŀ��������������LB_ERR
	'LB_GETTEXT = &H0189			'����,������ 	���ڻ�ȡָ���б�����ַ���������lParamָ��һ�������ַ����Ļ�������wParam��ָ���˽����ַ������б������������ػ�õ��ַ����ĳ��ȣ��������򷵻�LB_ERR
	'LB_GETTEXTLEN = &H018A			'����,0 		����ָ���б�����ַ������ֽڳ��ȡ�wParamָ�����б�����������������򷵻�LB_ERR���غͽo������P���ַ����L�ȣ���λ���ַ���
	LB_GETTOPINDEX = &H018E			'0,0			�����б���е�һ���ɼ�����������������򷵻�LB_ERR
	'LB_INITSTORAGE = &H01A8		'������,�ڴ��ֽ���	������ֻ������Windows95�汾�����㽫Ҫ���б���м���ܶ������кܴ�ı���ʱ����������Ԥ�ȷ���һ���ڴ棬�����ڽ��Ĳ�����һ��һ�εط����ڴ棬�Ӷ��ӿ���������ٶ�
	LB_INSERTSTRING = &H0181		'����,�ַ�����ַ	���б���е�ָ��λ�ò����ַ�����wParamָ�����б�������������Ϊ-1�����ַ���������ӵ��б��ĩβ��lParamָ��Ҫ������ַ���������������ʵ�ʵĲ���λ�ã����������󣬻᷵��LB_ERR��LB_ERRSPACE����LB_ADDSTRING��ͬ�����������ᵼ��LBS_SORT�����б���������򡣽��鲻Ҫ�ھ���LBS_SORT�����б����ʹ�ñ������������ƻ��б���Ĵ���
	'LB_ITEMFROMPOINT = &H01A9		'0,λ��			�����ָ������������������lParamָ�����б��ͻ�������16λΪX���꣬��16λΪY����
	LB_RESETCONTENT = &H0184		'0,0			��������б���
	'LB_SELECTSTRING = &H018C		'��ʼ�������,�ַ�����ַ	�������������ڵ�ѡ���б���趨��ָ���ַ�����ƥ����б���Ϊѡ���������������б����ʹѡ����ɼ������������弰�����ķ�����LB_FINDSTRING���ơ�����ҵ���ƥ�������ظ�������������û��ƥ��������LB_ERR���ҵ�ǰ��ѡ������ı�
	'LB_SELITEMRANGE = &H019B		'TRUE��FALSE,��Χ	�����������ڶ���ѡ���б������ʹָ����Χ�ڵ��б���ѡ�л���ѡ������lParamָ�����б��������ķ�Χ����16λΪ��ʼ���16λΪ������������wParamΪTRUE����ô��ѡ����Щ�б�������ʹ������ѡ��������������LB_ERR
	'LB_SELITEMRANGEEX = &H0183		'���,�յ�		�����ڶ���ѡ���б����ָ���յ����������趨�÷�ΧΪѡ�У���ָ���������յ����趨�÷�ΧΪ��ѡ
	'LB_SETANCHORINDEX = &H019C		'����,0			����������ѡ�еı����ָ������
	'LB_SETCARETINDEX = &H019E		'����,TRUE��FALSE	���ü������뽹�㵽ָ�������lParamΪTRUE�������ָ����ݿɼ�����lParamΪFALSE�������ָ����ȫ���ɼ�
	'LB_SETCOLUMNWIDTH = &H0195		'���(��),0		�����еĿ�ȣ���λ�����أ�
	'LB_SETCOUNT = &H01A7			'����,0			���ñ�����Ŀ
	'LB_SETCURSEL = &H0186			'����,0			�������ڵ�ѡ���б������ָ�����б���Ϊ��ǰѡ������Զ��������ɼ����򡣲���wParamָ�����б������������Ϊ-1����ô������б���е�ѡ��������������LB_ERR
	LB_SETHORIZONTALEXTENT = &H0194	'���(��),0 �����б��Ĺ�����ȣ���λ�����أ�
	'LB_SETITEMDATA = &H019A		'����,����ֵ	����ָ���б����32λ�������ݡ�
	'LB_SETITEMHIEGHT = &H01A0		'����,�߶�(��)	ָ���б�����ʾ�߶ȣ�����LBS_OWNERDRAWVARIABLE(�Ի��б���)���Ŀؼ���ֻ������wParamָ����ĸ߶ȣ�������񽫸������е��б���ĸ߶ȣ���λ�����أ�
	'LB_SETLOCALE = &H01A5			'���Դ���,0		ȡ�б��ǰ������������Դ��룬���û�ʹ��LB_ADDSTRING����Ͽ��е��б������Ӽ�¼��ʹ��LBS_SORT��������������ʱ������ʹ�ø����Դ��롣����ֵ�и�16λΪ���Ҵ���
	LB_SETSEL = &H0185				'TRUE��FALSE,����	�������ڶ���ѡ���б����ʹָ�����б���ѡ�л���ѡ�����Զ��������ɼ����򡣲���lParamָ�����б������������Ϊ-1�����൱��ָ�������е������wParamΪTRUEʱѡ���б������ʹ֮��ѡ���������򷵻�LB_ERR
	'LB_SETTABSTOPS = &H0192		'վ��,����˳���	�����б��Ĺ��(���뽹��)վ��������˳���
	LB_SETTOPINDEX = &H0197			'����,0			������ָ�����б�������Ϊ�б��ĵ�һ���ɼ���ú����Ὣ�б����������ʵ�λ�á�wParamָ�����б�����������������ɹ�������0ֵ�����򷵻�LB_ERR
	'LB_MULTIPLEADDSTRING = &H01B1
	'LB_GETLISTBOXINFO = &H01B2
	'LB_MSGMAX_501 = &H01B3
	'LB_MSGMAX_WCE4 = &H01B1
	'LB_MSGMAX_4 = &H01B0
	'LB_MSGMAX_PRE4 = &H01A8

	'�����ı��ؼ�����
	WM_GETTEXT = &H0D				'�ֽ���,�ַ�����ַ	��ȡ�����ı��ؼ����ı�
	WM_GETTEXTLENGTH = &H0E			'0,0				��ȡ�����ı��ؼ����ı��ĳ��ȣ����������ַ���(�ַ���)
	WM_SETTEXT = &H0C				'0,�ַ�����ַ		���ô����ı��ؼ����ı�

	'���ڶԻ�������
	WM_SETFONT = &H30				'������,True		�����ı�ʱ�����ʹ���Ϣ��ȡ�ؼ�Ҫ�õ�����
	WM_GETFONT = &H31				'0,0 				��ȡ��ǰ�ؼ������ı���������
	WM_FONTCHANGE = &H1D 			'0,0				��ϵͳ��������Դ��仯ʱ���ʹ���Ϣ�����ж�������
	WM_SETREDRAW = &H0B 			'Boolean,0			���ô����Ƿ����ػ���False ��ֹ�ػ���True �����ػ�
	'WM_CTLCOLORMSGBOX = &H132		'�豸���,�ؼ����	������Ϣ����ɫ
	'WM_CTLCOLOREDIT = &H133		'�豸���,�ؼ����	���ñ༭����ɫ
	'WM_CTLCOLORLISTBOX = &H134		'�豸���,�ؼ����	�����б����ɫ
	'WM_CTLCOLORBTN = &H135			'�豸���,�ؼ����	���ð�ť��ɫ
	'WM_CTLCOLORDLG = &H136			'�豸���,�ؼ����	���öԻ�����ɫ
	'WM_CTLCOLORSCROLLBAR = &H137	'�豸���,�ؼ����	���ù�������ɫ
	'WM_CTLCOLORSTATIC = &H138		'�豸���,�ؼ����	����״̬����ɫ
	WM_SETFOCUS = &H7				'�ؼ����,0,0		���ý���

	'WM_KEYDOWN = &H100				'�ؼ����,�����,0	ģ�ⰴ�°�ť
	'WM_KEYUP = &H101				'�ؼ����,�����,0	ģ��̧��ť
	'WM_LBUTTONDOWN = &H201			'�ƶ����
	'WM_LBUTTONUP = &H202			'����������
	'WM_LBUTTONDBLCLK = &H203		'�ͷ�������
	'WM_RBUTTONDOWN = &H204			'˫��������
	'WM_RBUTTONUP = &H205			'��������Ҽ�
	'WM_RBUTTONDBLCLK = &H206		'�ͷ�����Ҽ�
	'WM_MBUTTONDOWN = &H207			'˫������Ҽ�
	WM_MBUTTONUP = &H208			'��������м�
	'WM_MBUTTONDBLCLK = &H209		'�ͷ�����м�
	'WM_MOUSEWHEEL = &H20A			'˫������м�

	'WM_HSCROLL= &H114				'�ؼ����,����������,������λ��	���� SB_BOTTOM ָ����ˮƽ������λ��
	WM_VSCROLL = &H115				'�ؼ����,����������,������λ��	���� SB_BOTTOM ָ���Ĵ�ֱ������λ��
	'SB_TOP = &H06					'������λ��, ���ô�ֱ������������
	'SB_LEFT = &H06					'������λ��, ����ˮƽ���������ұ�
	SB_BOTTOM = &H07				'������λ��, ���ô�ֱ���������ײ�
	'SB_RIGHT = &H07				'������λ��, ����ˮƽ���������ұ�

	'��ť�¼�
	'BM_GETCHECK = &HF0				'�ؼ����,0,0		��ȡ��ѡ��ť��ѡ���ѡ��״̬
	'BM_SETCHECK = &HF1				'�ؼ����,0,0		���õ�ѡ��ť��ѡ���ѡ��״̬
	BM_GETSTATE = &HF2				'�ؼ����,0,0		��ȡ��ť�Ƿ񱻰��¹�
	BM_SETSTATE = &HF3				'�ؼ����,��ť״̬,0	���ð�ť��״̬��True δ����״̬��False ����״̬
	'BM_SETSTYLE = &HF4				'�ؼ����,��ť��ʽ,0	���ð�ť����ʽ
	BM_CLICK = &HF5					'�ؼ����,0,0		ģ������ť
	'BM_GETIMAGE = &HF6
	'BM_SETIMAGE = &HF7

	'BN_CLICKED = &H0				'�ؼ����,0,0		�û������˰�ť

	'BST_UNCHECKED = &H0      		'���õ�ѡ��͸�ѡ��ѡ��Ϊδѡ��״̬
	'BST_CHECKED = &H1				'���õ�ѡ��͸�ѡ��ѡ��Ϊ��ѡ��״̬
	'BST_INDETERMINATE = &H2
	BST_PUSHED = &H4      			'���ð�ťΪ����״̬
	'BST_FOCUS = &H8      			'���ý���
End Enum

'�����ı�����Ҷ�λ����
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

'���ڿؼ�����ʾ������
'ShowWindow ���ֳ���
Public Enum ShowWindowValue
	SW_SHOW = 5
	SW_HIDE = 0
End Enum
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function EnableWindow Lib "user32" Alias "EnableWindow" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
'���أ������ʾ�ɹ������ʾʧ�ܡ�������GetLastError
'hwnd�����ڻ�ؼ����
'fEnable Long�������������ֹ

'�������ý��㵽�ؼ���������޷�ʹ��
'Private Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long)
'���ڷ��ؽ���ؼ��ľ��
Public Declare Function GetFocus Lib "user32.dll" () As Long
'���ڷ��ؿؼ�ID�ľ��
Public Declare Function GetDlgItem Lib "user32.dll" ( _
	ByVal hDlg As Long, _
	ByVal nIDDlgItem As Long) As Long
'��ȡָ���㴰�ڵľ��
Public Declare Function WindowFromPoint Lib "user32.dll" ( _
	ByVal xPoint As Long, _
	ByVal yPoint As Long) As Long
'��ȡָ�����Ӵ��ڵľ��
'Private Declare Function ChildWindowFromPoint Lib "user32.dll" ( _
'	ByVal hwnd As Long, _
'	ByVal xPoint As Long, _
'	ByVal yPoint As Long) As Long

'��ȡ�����ù�����λ�ú���
'Private Declare Function GetScrollPos Lib "user32.dll" ( _
'	ByVal hwnd As Long, _
'	ByVal nBar As Long) As Long
'Private Declare Function SetScrollPos Lib "user32.dll" ( _
'	ByVal hwnd As Long, _
'	ByVal nBar As Long, _
'	ByVal nPos As Long, _
'	ByVal bRedraw As Long) As Long

'���ؾ������궨��
Private Type RECT
	Left As Long
	Top As Long
	Right As Long
	Bottom As Long
End Type

'DrawText �� wFormat ��������
Private Enum DrawTextConstants
	DT_BOTTOM = &H8				'�����ĵ��������εײ�����ֵ����� DT_SINGLELINE ��ϡ�
	DT_CALCRECT = &H400			'�������εĿ�͸ߡ���������ж��У�DrawTextʹ��lpRect����ľ��εĿ�ȣ�����չ���εĵ�ѵ���������ĵ����һ�У�
								'�������ֻ��һ�У���DrawText�ı���ε��ұ߽磬�������������е����һ���ַ��������κ�һ�������DrawText���ظ�ʽ�����ĵĸ߶ȶ�����д���ġ�
	DT_CENTER = &H1				'ʹ�����ھ�����ˮƽ���С�
	DT_EXPANDTABS = &H40		'��չ�Ʊ����ÿ���Ʊ����ȱʡ�ַ�����8
	DT_EXTERNALLEADING = &H200	'���еĸ߶������������ⲿ��ͷ��ͨ�����ⲿ��ͷ���������������еĸ߶��
	DT_INTERNAL = &H1000		'��ϵͳ�������������Ķ�����
	DT_LEFT = &H0				'��������롣
	DT_NOCLIP = &H100			'�޲ü����Ƶ�DT_NOCLIPʹ��ʱDrawText��ʹ�û������ӿ졣
	DT_NOPREFIX = &H800			'�ر�ǰ׺�ַ��Ĵ���ͨ��DrawText��������ǰ׺�ַ���&Ϊ�������ַ����»��ߣ�����&&Ϊ��ʾ����&��ָ��DT_NOPREFIX�����ִ����رա�
	DT_RIGHT = &H2				'�����Ҷ��롣
	DT_SINGLELINE = &H20		'��ʾ���ĵ�ͬһ�У��س��ͻ��з����������С�
	DT_TABSTOP = &H80			'�����Ʊ�����uFormat��15"C8λ����λ���еĸ�λ�ֽڣ�ָ��ÿ���Ʊ�����ַ�����ÿ���Ʊ����ȱʡ�ַ�����8��
	DT_TOP = &H0				'���Ķ��˶��루���Ե��У���
	DT_VCENTER = &H4			'����ˮƽ���У����Ե��У���
	DT_WORDBREAK = &H10			'�Ͽ��֡���һ���е��ַ��������쵽��lpRectָ���ľ��εı߿�ʱ�������Զ�������֮��Ͽ���һ���س�һ����Ҳ��ʹ���۶ϡ�
	DT_EDITCONTROL = &H2000&	'���ƶ��б༭���Ƶ�������ʾ���ԣ�����أ�Ϊ�༭���Ƶ�ƽ���ַ��������ͬ���ķ�������ģ��˺�������ʾֻ�ǲ��ֿɼ������һ�С�
	DT_END_ELLIPSIS = &H8000&	'����ָ��DT_END_ELLIPSIS���滻���ַ���ĩβ���ַ�����ָ��DT_PATH_ELLIPSIS���滻�ַ����м���ַ���
								'����ַ����ﺬ�з�б����DT_PATH_ELLIPSIS�����ܵر������һ����б��֮������ġ�
	DT_MODIFYSTRING = &H10000	'�޸ĸ������ַ�����ƥ����ʾ�����ģ��˱�־�����DT_END_ELLIPSIS��DT_PATH_ELLIPSISͬʱʹ�á�
	DT_PATH_ELLIPSIS = &H4000&
	DT_RTLREADING = &H20000		'��ѡ����豸������������Hebrew��Arabicfʱ��Ϊ˫�����İ��Ŵ��ҵ�����Ķ�˳���Ǵ����ҵġ�
	DT_WORD_ELLIPSIS = &H40000	'�ض̲����Ͼ��ε����ģ���������Բ��
	'ע�⣺DT_CALCRECT, DT_EXTERNALLEADING, DT_INTERNAL, DT_NOCLIP, DT_NOPREFIXֵ���ܺ�DT_TABSTOPֵһ��ʹ�á�
End Enum

'��ȡ�ַ��������ش�Сʹ�õĺ���
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

'�����ı���ɫ
'Private Declare Function SetTextColor Lib "gdi32.dll" ( _
'	ByVal hDC As Long, _
'	ByVal crColor As Long) As Long

'��� VK ��ֵ���壬���� GetAsyncKeyState ����
Public Enum VK
	VK_LBUTTON = &H01	'������
	VK_RBUTTON = &H02	'����Ҽ�
	VK_MBUTTON = &H04	'����м�
	'VK_END = &H23		'���� End ��
	'VK_HOME = &H24		'���� Home ��
	'VK_LEFT = &H25		'���������
	'VK_UP = &H26		'�������ϼ�
	'VK_RIGHT = &H27	'�������Ҽ�
	'VK_DOWN = &H28		'�������¼�
	VK_ESCAPE = &H1B	'Esc ��
End Enum

'��ȡ��갴��״̬
'Private Declare Function GetKeyState Lib "user32" Alias "GetKeyState" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
'GetAsyncKeyState �������ص���ָ�������˲ʱ��Ӳ���ж�״ֵ̬���������ַ���ֵ��
'0 ����ǰδ���ڰ���״̬���������ϴε���GetAsyncKeyState��ļ�Ҳδ������
'1 ����ǰδ���ڰ���״̬�����ڴ�֮ǰ�����ϴε���GetAsyncKeyState�󣩼�����������
'-32768����16������&H8000������ǰ���ڰ���״̬�����ڴ�֮ǰ�����ϴε���GetAsyncKeyState�󣩼�δ������
'-32767����16������&H8001������ǰ���ڰ���״̬�������ڴ�֮ǰ�����ϴε���GetAsyncKeyState�󣩼�Ҳ����������

'��ȡ������Ļλ��
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
'�ƶ���굽��Ļ��ָ��λ��
'Private Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
'ת����Ļ�������Ϊ�ͻ�������
'Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'ת���ͻ����������Ϊ��Ļ����
'Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

' Logical Font
'Private Const LF_FACESIZE = 32
'Private Const LF_FULLFACESIZE = 64

'�������ͣ����� ChooseFont ����
'Private Enum FONTTYPE
'	BOLD_FONTTYPE				'����Ϊ���塣����Ϣ�� LOGFONT �ṹ��Ա lfWeight ���ƣ�����FW_BOLD ��Ч��
'	ITALIC_FONTTYPE				'����Ϊб�塣����Ϣ�� LOGFONT �ṹ��Ա lfItalic ���ơ�
'	PRINTER_FONTTYPE			'����Ϊ��ӡ�����塣
'	REGULAR_FONTTYPE = &H400	'����Ϊ��׼������Ϣ�Ǵ��� LOGFONT �ṹ��Ա lfWeight ���ƣ����� FW_REGULAR ��Ч��
'	SCREEN_FONTTYPE				'����Ϊ��Ļ���塣
'	SIMULATED_FONTTYPE			'���屻ͼ���豸�ӿ� (GDI) ģ��
'End Enum

'�ַ��������� LOG_FONT ���͵� lfCharSet
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

'ChooseFont ���͵� flags ��������
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
	CF_SCREENFONTS = &H1		'��ʾ��Ļ����
	CF_PRINTERFONTS = &H2		'��ʾ��ӡ������
	CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)	'���߶���ʾ
	CF_EFFECTS = &H100&			'�������Ч��
	CF_LIMITSIZE = &H2000&		'���������С����
End Enum

'��������
Public Type LOG_FONT
	lfHeight As Long			'�����С
	lfWidth As Long				'������
	lfEscapement As Long		'������ʾ�Ƕ�
	lfOrientation As Long		'����Ƕ�
	lfWeight As Long			'�Ƿ����
	lfItalic As Byte			'�Ƿ�б��
	lfUnderline As Byte			'�Ƿ��»���
	lfStrikeOut As Byte			'�Ƿ�ɾ����
	lfCharSet As Byte			'�ַ���
	lfOutPrecision As Byte		'�������
	lfClipPrecision As Byte		'�ü�����
	lfQuality As Byte			'�߼�����������豸ʵ������֮��ľ���
	lfPitchAndFamily As Byte	'����������弯
	'lfFaceName As String * LF_FACESIZE	'��������(�����������壬��������ʱ�����)
	lfFaceName(31) As Byte		'��������
	lfColor As Long				'������ɫ
End Type

'����Ի�������
Private Type CHOOSE_FONT
	lStructSize As Long			' size of CHOOSEFONT structure in byte
	hwndOwner As Long			' caller's window handle
	hDC As Long					' printer DC/IC or NULL
	lpLogFont As Long			' LogFont �ṹ��ַ
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

'RedrawWindow ������ fuRedraw ��������
Private Enum RDW
	RDW_INVALIDATE = &H1		'���ã����Σ��ػ�����
	RDW_INTERNALPAINT = &H2		'��ʹ���ڲ�����Ч��Ҳ����Ͷ��һ��WM_PAINT��Ϣ
	RDW_ERASE = &H4				'�ػ�ǰ��������ػ�����ı�����Ҳ����ָ��RDW_INVALIDATE
	RDW_VALIDATE = &H8			'�����ػ�����
	RDW_NOINTERNALPAINT = &H10	'��ֹ�ڲ����ɻ�������������ɵ��κδ���WM_PAINT��Ϣ�������Ч�����Ի�����WM_PAINT��Ϣ
	RDW_NOERASE = &H20			'��ֹɾ���ػ�����ı���
	RDW_NOCHILDREN = &H40		'�ػ������ų��Ӵ��ڣ�ǰ�������Ǵ������ػ�����
	RDW_ALLCHILDREN = &H80		'�ػ����������Ӵ��ڣ�ǰ�������Ǵ������ػ�����
	RDW_UPDATENOW = &H100		'��������ָ�����ػ�����
	RDW_ERASENOW = &H200		'����ɾ��ָ�����ػ�����
	RDW_FRAME = &H400			'��ǿͻ����������ػ������У���Էǿͻ������и��¡�Ҳ����ָ��RDW_INVALIDATE
	RDW_NOFRAME = &H800			'��ֹ�ǿͻ������ػ�����������ػ������һ���֣���Ҳ����ָ��RDW_VALIDATE
End Enum

'�ػ��Ի�����
Private Declare Function RedrawWindow Lib "user32.dll" ( _
	ByVal hwnd As Long, _
	ByVal lprcUpdate As Long, _
	ByVal hrgnUpdate As Long, _
	ByVal fuRedraw As Long) As Long

'��ȡע���ֵ������
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

'��ȡ Windows �汾��
Private Type OSVERSIONINFO
	dwOSVersionInfoSize As Long		'��ʹ��GetVersionEx֮ǰҪ���˳�ʼ��Ϊ�ṹ�Ĵ�С
	dwMajorVersion As Long			'ϵͳ���汾��
	dwMinorVersion As Long			'ϵͳ�ΰ汾��
	dwBuildNumber As Long			'ϵͳ������
	dwPlatformId As Long			'ϵͳ֧�ֵ�ƽ̨(�����1)
	szCSDVersion As String * 128	'ϵͳ������������
	wServicePackMajor As Integer	'ϵͳ�����������汾
	wServicePackMinor As Integer	'ϵͳ�������Ĵΰ汾
	wSuiteMask As Integer			'��ʶϵͳ�ϵĳ�����(�����2)
	wProductType As Byte			'��ʶϵͳ����(�����3)
	wReserved As Byte				'����,δʹ��
End Type
'��ȡ�汾��Ϣ
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
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
'Private Declare Function MapAndLoad Lib "imagehlp.dll" ( _
'	ByVal ImageName As String, _
'	ByVal DllPath As String, _
'	LoadedImage As LOADED_IMAGE, _
'	ByVal DotDll As Boolean, _
'	ByVal ReadOnly As Boolean) As Boolean
	'ImageName	����� PE �ļ����ļ���
	'DllPath ��λ�ļ���·���������� Null�������� PATH ���������е�·����
	'LoadedImage �ṹ�� LOADED_IMAGE ������ IMAGEHLP.H file
	'DotDll ����Ҫ���Ҹ��ļ�����û��ָ����չ������ʹ�� .exe �� .dll ����չ��
	'�� DotDll ��־��Ϊ True����ʹ�� .dll ��չ��; ������ .exe ��չ��
	'ReadOnly ����Ϊ True�����ļ���ӳ��Ϊ Read-only��
	'����ֵ ���ɹ������� True; ���򷵻� False
'Private Declare Function UnMapAndLoad Lib "imagehlp.dll" (LoadedImage As LOADED_IMAGE) As Boolean
	'ʹ����ӳ���ļ���Ӧ�õ��� UnMapAndLoad() ������ �˺������ PE �ļ���ӳ�䲢������ MapAndLoad() �������Դ��
	'LoadedImage ָ�� LOADED_IMAGE �ṹ���ָ�룬��ָ�����ǰ����� MapAndLoad() �������ص�ָ�롣
	'����ֵ ���ɹ����򷵻� True; ���򷵻� False

'������Ϣ����
Private Enum FormatMSG
	FORMAT_MESSAGE_FROM_SYSTEM = &H1000
	FORMAT_MESSAGE_IGNORE_INSERTS = &H200
End Enum

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

'�����޸��ļ���ʱ������
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

'�����޸��ļ���ʱ�����Ժ���
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

'�й�ע����뵼������
Private Enum REG_KEY_IMPORT_EXPORT
	REG_FORCE_RESTORE = 8&
	TOKEN_QUERY= &H8&
	TOKEN_ADJUST_PRIVILEGES = &H20&
	SE_PRIVILEGE_ENABLED = &H2
End Enum
Private Const SE_RESTORE_NAME As String = "SeRestorePrivilege"
Private Const SE_BACKUP_NAME As String = "SeBackupPrivilege"

'ע���ؼ��ַ���ѡ��
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

'ע���ؼ��ָ�����...
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

'�Զ��嵯���˵�
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

'��ȡ��Ļ�ֱ��ʺ����ô��ڴ�С
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


'����ı��Ĵ���ҳ
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


'�Զ�����ֽ�����Ĵ���ҳ
'fType = 0 ��ǰ��λ������.CodePage < 1000��UniCode��UniCode BE ǰ��λ
Public Function getCodePageRegExp(FN As FILE_IMAGE,strData As STRING_SUB_PROPERTIE,CPList() As LANG_PROPERTIE, _
			EncodeList() As Integer,ByVal MinLength As Integer,StrTypeEndChar() As String,EndChar() As String, _
			ByVal Mode As Long,ByVal fType As Integer,ByVal LengthFilterSet As Long,ByVal MaxLength As Long) As Long
	Dim i As Long,j As Long,k As Long,l As Long,m As Long,n As Long,sp As Long,ep As Long
	Dim Matches As Object,Data As STRING_SUB_PROPERTIE
	On Error GoTo ExitFunciton
	'��ʼ��
	getCodePageRegExp = strData.lEndAddress
	strData.CodePage = CP_UNKNOWN
	With Data
	.lStartAddress = strData.lStartAddress
	.lEndAddress = strData.lEndAddress
	.lMaxHexLength = strData.lHexLength
	'������ҳ����ȡ�ִ���ַ
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


'��ȡ������ʽ�ִ��ĵ�ַ
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


'��ת Hex ��
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


'�ֽ�ת Hex ��
'StartPos <= EndPos ��ȡ��λ����λ�� Hex ���룬�����ȡ��λ����λ�� Hex ����
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


'�ֽ�ת Hex ת����
'StartPos <= EndPos ��ȡ��λ����λ�� Hex ���룬�����ȡ��λ����λ�� Hex ����
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


'�����ֵ���ֽ���
'CheckByteOrder ���ؼ������-1 = ��ˣ�0 = С�ˣ�1 = δ֪
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


'ת���ֽ�����Ϊ��ֵ
'ByteOrder = False ����λ�ں�ת�����򰴸�λ��ǰת
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


'ת����ֵΪ�ֽ�����(���ڳ��ȵĸ�λ�ض�)
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


'ת�� HEX ����Ϊ�ֽ�����
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


'����ִ��Ƿ����ָ���ַ�(�ı���ͨ����Ƚ�)
'Mode = 0 ����ִ��Ƿ����ָ���ַ������ҳ�ָ���ַ���λ��
'Mode = 1 ����ִ��Ƿ�ֻ����ָ���ַ�
'Mode = 2 ����ִ��Ƿ����ָ���ַ�
'Mode = 3 ����ִ��Ƿ�ֻ������С��д��ָ���ַ�����ʱ IgnoreCase ������Ч
'Mode = 4 ����ִ����Ƿ���������ͬ���ַ���StrNum Ϊ�����ַ�����
'StrRange  �����ִ���鷶Χ (���� [Min - Max|Min - Max] ��ʾ��Χ)
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


'����ִ��Ƿ����ָ���ַ�(������ʽ�Ƚ�)
'Mode = 0 ����ִ��Ƿ����ָ���ַ������ҳ�ָ���ַ���λ��
'Mode = 1 ����ִ��Ƿ�ֻ����ָ���ַ�
'Mode = 2 ����ִ��Ƿ����ָ���ַ�
'Mode = 3 ����ִ��Ƿ�ֻ������С��д��ָ���ַ�����ʱ IgnoreCase ������Ч
'Mode = 4 ����ִ��Ƿ���������ͬ���ַ���StrNum Ϊ�����ظ��ַ�����
'Mode = 5 ����ִ��Ƿ����ָ���ִ���������ƥ����ִ��ܳ��� (�ʺ��ַ���ϲ�ѯ)
'Patrn  Ϊ������ʽģ��
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


'�ִ���������ת��
'Mode = 0 �� PSL �汾�ĺ����治ͬ�ֱ�ת������ַ���ȫ���򲿷���������չ�ַ�
'Mode = 1 ת������ַ���������������չ�ַ�
'Mode = 2 ת������ַ��Ͳ���ϵͳ������ʾ����������չ�ַ�
'Mode = 3 ��ת������ַ�
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


'ת����������չ�ַ�Ϊʮ������ת���
'Mode = 0 �� PSL �汾�ĺ����治ͬ�ֱ�ת������ַ���ȫ���򲿷���������չ�ַ�
'Mode = 1 ת������ַ���������������չ�ַ�
'Mode = 2 ת������ַ��Ͳ���ϵͳ������ʾ����������չ�ַ�
'Mode = 3 ��ת������ַ�
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


'ת����������չ�ַ�Ϊʮ������ת���
'Mode = 0 �� PSL �汾�ĺ����治ͬ�ֱ�ת������ַ���ȫ���򲿷���������չ�ַ�
'Mode = 1 ת������ַ���������������չ�ַ�
'Mode = 2 ת������ַ��Ͳ���ϵͳ������ʾ����������չ�ַ�
'Mode = 3 ��ת������ַ�
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


'�ִ���������ת��
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


'ת���˽��ƻ�ʮ������ת���
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


'ת���˽��ƻ�ʮ������ת���
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


'ת�� Athena-A ת���Ϊ��׼��ʽ
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


'ת�� Athena-A �˽��ƻ�ʮ������ת���Ϊ��׼��ʽ
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


'ת���ַ�Ϊ Long ����ֵ
Public Function StrToLong(ByVal mStr As String,Optional ByVal DefaultValue As Long) As Long
	On Error GoTo errHandle
	StrToLong = CLng(mStr)
	Exit Function
	errHandle:
	StrToLong = DefaultValue
End Function


'��ȥ�ִ�ǰ��ָ���� PreStr �� AppStr
'fType = -1 ��ȥ���ִ�ǰ��Ŀո������ָ���� PreStr �� AppStr������ȥ���ִ���ǰ��ո�
'fType = 0 ȥ���ִ�ǰ��Ŀո������ָ���� PreStr �� AppStr������ȥ���ִ���ǰ��ո�
'fType = 1 ȥ���ִ�ǰ��Ŀո������ָ���� PreStr �� AppStr����ȥ���ִ���ǰ��ո�
'fType = 2 ȥ���ִ�ǰ��Ŀո��ָ���� PreStr �� AppStr 1 �Σ�����ȥ���ִ���ǰ��ո�
'fType > 2 ȥ���ִ�ǰ��Ŀո��ָ���� PreStr �� AppStr 1 �Σ���ȥ���ִ���ǰ��ո�
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


'�ִ�ǰ�󸽼�ָ���� PreStr �� AppStr
'fType = 0 ��ȥ���ִ�ǰ��ո񣬵����ִ�ǰ�󸽼�ָ���� PreStr �� AppStr
'fType = 1 ȥ���ִ�ǰ��ո񣬲����ִ�ǰ�󸽼�ָ���� PreStr �� AppStr
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


'��ȡ4���ֽ�ֵ (32 ֵ, 4���ֽ�)
Private Function GetDWord(Source As Variant,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Double
	GetDWord# = GetWord(Source, Offset, Mode)
	GetDWord# = GetDWord# + 65536# * GetWord(Source, Offset + 2, Mode)
End Function


'��ȡ2���ֽ�ֵ (16 ֵ, 2���ֽ�)
Private Function GetWord(Source As Variant,Optional ByVal Offset As Long = -1,Optional ByVal Mode As Long = -1) As Long
	GetWord& = GetByte(Source, Offset, Mode)
	GetWord& = GetWord& + 256& * GetByte(Source, Offset + 1,Mode)
End Function


'��ȡ8���ֽ�ֵ (64 λֵ,8���ֽ�)
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


'��ȡ4���ֽ�ֵ (32 λֵ,4���ֽ�, -2,147,483,648 �� 2,147,483,647)
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


'��ȡ2���ֽ�ֵ (16 λֵ, 2���ֽ�, -32,768 �� 32,767)
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


'��ָ����ַ��ȡһ���ֽ�(8 λֵ, 1���ֽ�)
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


'��ȡ�����ڵ��ֽ�����
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


'��ȡ�ֽ�����(�ڴ�ӳ�䷽ʽ)
'Private Function getBytesByMap(ByVal Source As Long,ByVal Length As Long) As Byte()
'	Dim ppSA As Long, pSA As Long
'	Dim tagNewSA As SAFEARRAYID, tagOldSA As SAFEARRAYID
'	ReDim Bytes(0) As Byte							'��ʼ������
'	ppSA = VarPtr(Bytes(0))							'���ָ��SAFEARRAY��ָ���ָ��
'	MoveMemory pSA, ppSA, 4							'���ָ��SAFEARRAY��ָ��
'	MoveMemory tagOldSA, pSA, Len(tagOldSA)			'����ԭ����SAFEARRAY��Ա��Ϣ
'	CopyMemory tagNewSA, tagOldSA, Len(tagNewSA)	'����SAFEARRAY��Ա��Ϣ
'	tagNewSA.rgsabound(0).cElements = Length		'�޸�����Ԫ�ظ���
'	tagNewSA.pvData = Source						'�޸��������ݵ�ַ
'	WriteMemory pSA, tagNewSA, Len(tagNewSA)		'��ӳ�������ݵ�ַ��������
'	getBytesByMap = Bytes
'	WriteMemory pSA, tagOldSA, Len(tagOldSA)		'�ָ������SAFEARRAY�ṹ��Ա��Ϣ
'End Function


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


'д���ֽ�����
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


'д���Զ�������ֵ
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


'д���Զ�����������
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


'��ȡ�����ֽڳ���
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


'�������ļ���
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


'�����ı��ļ�
Public Function CreateTXTFile(ByVal FilePath As String,ByVal Text As String,ByVal Code As String) As Boolean
	Dim i As Long
	i = InStrRev(FilePath,"\")
	If i > 0 Then
		If MkSubDir(Left$(FilePath,i)) = False Then Exit Function
		CreateTXTFile = WriteToFile(FilePath,Text,Code)
	End If
End Function


'��ȡ��ǰ�ļ����е��ļ��б�
'Mode = True �����Ƿ���� DefaultFile ������Ϊ�������Ϊ�׸��ļ�
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


'��ȡ��ǰ�ļ��е����ļ����е��ļ��б�
'FindFile = "" ʱ .sName = �ļ�������Ŀ¼�������� .sName = ��Ŀ¼�е��ļ�
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


'ɾ���ļ��У�����ɾ�����ļ���
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


'ɾ���ļ��У������������ļ���
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


'�ֽ�����ת������ʽʹ�õ�ת���ģ��
'Mode = 0 תΪ�� [] ��ʽ������Ϊ�� [] ��ʽ
Public Function Byte2RegExpPattern(Bytes() As Byte,Optional ByVal Mode As Long,Optional ByVal CodePage As Long) As String
	If Mode = 0 Then
		Byte2RegExpPattern = "[" & Byte2HexEsc(Bytes,0,-1,CodePage) & "]"
	Else
		Byte2RegExpPattern = Byte2HexEsc(Bytes,0,-1,CodePage)
	End If
End Function


'�ַ���ת������ʽʹ�õ� Unicode ת���ģ��
'Mode = 0 תΪ�� [] ��ʽ������Ϊ�� [] ��ʽ
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


'Hex �ַ���ת������ʽʹ�õ� Hex ת���ģ��
'Mode = 0 תΪ�� [] ��ʽ������Ϊ�� [] ��ʽ
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


'��ȡż��λ
'Mode = 0 ������ 1 ���ֽڣ�Mode = 1 ������ 1 ���ֽ�
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


'�ַ���ת�ֽ�����
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


'�ֽ�����ת�ַ���
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


'��ת�ֽ����飬��������ֵ���ֽ�����ĸ��ֽں͵��ֽڻ���
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


'���ַ���������ַ���ʮ�������ֽڳ���
'Mode = 1 ת��, ����ת��
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


'��ȡ�ַ����Ľ�ȡλ�ã������ȡ���˫�ֽ��ַ�
'���� IsDBCSLeadPos = ��ȡλ��(��ȡ����ֽڳ���)
'Mode = False �����ؽ�ȡ����ִ�, ���򷵻ؽ�ȡ����ִ�
'fType = False ���ȡ������ǰ��ȡ
'ByteLength Ϊ��Ҫ���ֽڳ���
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


'��ȡ�ַ����Ŀո����ֽ���
'���� FillStrWithSpape = ��������Ŀո��ֽ���(fType = False Ϊ��ֵ������Ϊ��ֵ)
'Mode = False �����ز����ַ�������ִ�, ���򷵻ز����ַ�������ִ�
'fType = False ��˿ո���, ����ǰ�˿ո���
'ByteLength Ϊ��Ҫ���ֽڳ���
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


'˳�����ֽ������в���ƥ������
'Mode = 0������ StartPos �� EndPos ֮���λ��(���� StartPos �� EndPos)
'Mode <> 0�����ҳ� StartPos �� EndPos ֮�������λ��(������ StartPos �� EndPos)
'���� InByte ֵ = 0��û���ҵ���> 0 �ҵ� (�� 1 ��ʼ�ĵ�ַ)
'ע�⣺StartPos��EndPos �� MaxPos ��Ϊ�� 0 ��ʼ�ĵ�ַ
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
			InByte = i	'ע�� InStrB �����ҵ���һ�����ͷ���"1"
			Exit Do
		ElseIf (i + Length - 2 < StartPos) Or (i + Length - 2 > EndPos) Then
			If MaxPos > 0 Then
				If i + Length - 2 > MaxPos Then Exit Do
			End If
			InByte = i	'ע�� InStrB �����ҵ���һ�����ͷ���"1"
			Exit Do
		End If
		NextNum:
		i = InStrB(i + Length,Bytes,tempByte)
	Loop
End Function


'˳�����ֽ������в���ƥ������
'Mode = 0������ StartPos �� EndPos ֮���λ��(���� StartPos �� EndPos)
'Mode <> 0�����ҳ� StartPos �� EndPos ֮�������λ��(������ StartPos �� EndPos)
'���� InByte ֵ = 0��û���ҵ���> 0 �ҵ� (�� 1 ��ʼ�ĵ�ַ)
'ע�⣺StartPos��EndPos �� MaxPos ��Ϊ�� 0 ��ʼ�ĵ�ַ
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
			InByteRegExp = i	'ע���ҵ���һ�����ͷ���"1"
			Exit For
		ElseIf (i + Length - 2 < StartPos) Or (i + Length - 2 > endPos) Then
			InByteRegExp = i	'ע���ҵ���һ�����ͷ���"1"
			Exit For
		End If
	Next j
	ExitFunction:
End Function


'�������ֽ������в���ƥ������
'Mode = 0������ StartPos �� EndPos ֮���λ��(���� StartPos �� EndPos)
'Mode <> 0�����ҳ� StartPos �� EndPos ֮�������λ��(������ StartPos �� EndPos)
'���� InByteRev ֵ = 0��û���ҵ���> 0 �ҵ� (�� 1 ��ʼ�ĵ�ַ)
'ע�⣺StartPos��EndPos �� MaxPos ��Ϊ�� 0 ��ʼ�ĵ�ַ
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
			InByteRev = i	'ע�� InStrB �����ҵ���һ�����ͷ���"1"
		ElseIf (i + Length - 2 < StartPos) Or (i + Length - 2 > EndPos) Then
			If MaxPos > 0 Then
				If i + Length - 2 > MaxPos Then Exit Do
			End If
			InByteRev = i	'ע�� InStrB �����ҵ���һ�����ͷ���"1"
		End If
		NextNum:
		i = InStrB(i + Length,Bytes,tempByte)
	Loop
End Function


'�������ֽ������в���ƥ������
'Mode = 0������ StartPos �� EndPos ֮���λ��(���� StartPos �� EndPos)
'Mode <> 0�����ҳ� StartPos �� EndPos ֮�������λ��(������ StartPos �� EndPos)
'���� InByteRev ֵ = 0��û���ҵ���> 0 �ҵ� (�� 1 ��ʼ�ĵ�ַ)
'ע�⣺StartPos��EndPos �� MaxPos ��Ϊ�� 0 ��ʼ�ĵ�ַ
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
			InByteRevRegExp = i	'ע���ҵ���һ�����ͷ���"1"
			Exit For
		ElseIf (i + Length - 2 < StartPos) Or (i + Length - 2 > EndPos) Then
			InByteRevRegExp = i	'ע���ҵ���һ�����ͷ���"1"
			Exit For
		End If
	Next j
	ExitFunction:
End Function


'����ָ�������Ŀ��ֽ�λ�ã������ؿ��ֽڿ�ʼλ�� (�� 1 ��ʼ�ĵ�ַ)
'Mode = 0 �����׸���ʼλ�ã�= 1 �������һ����ʼλ�ã�= 2 �����׸�����λ�ã�>= 3 �������һ������λ��
'���� NullInByteRegExp ֵ = 0��û���ҵ���> 0 �ҵ� (�� 1 ��ʼ�ĵ�ַ)
'ע�⣺StartPos��EndPos ��Ϊ�� 0 ��ʼ�ĵ�ַ
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


'�ϲ��ֽ�����
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


'���ù��λ�ã���ʼ�к���ʼ��Ϊ0��
'LineNo = �к�(�ı���ͷ��ʼ��)����ʼ�� = 0
'ColNo = �к�(��ǰ���׿�ʼ��)����ʼ�� = 0
'Length = ����ѡ������
'strText = ""�� ColNo �� Length Ϊ�ַ��������� Win7 �����°汾
'strText <> ""��ColNo �� Length Ϊ�ֽ��������� Win10 �����ϰ汾
'strText Ϊ�ı����е��ı��������� GetTextBoxString(hwnd) ����ȡ
Private Sub SetCurPos(ByVal hwnd As Long,ByVal LineNo As Long,ByVal ColNo As Long,ByVal Length As Long,Optional ByVal strText As String)
	'��ȡָ���е����ַ����ı��е��ַ���ƫ��
	LineNo = SendMessageLNG(hwnd, EM_LINEINDEX, LineNo, 0&)
	'Win10 �� EM_SETSEL ��Ҫ�ֽ����������ص� LineNo Ϊ�ַ���������ҪתΪ�ֽ���ƫ��
	If strText <> "" Then
		'LineNo = StrHexLength(Left$(strText,LineNo),GetACP,0) '�������׵��ֽ���
		ReDim tmpByte(0) As Byte
		tmpByte = StrConv(Left$(strText,LineNo),vbFromUnicode)
		LineNo = UBound(tmpByte) + 1
	End If
	'ѡ��ָ���ı���������Χ (Win10 �µĲ���Ϊ�ֽ���ƫ��)
	SendMessageLNG hwnd, EM_SETSEL, LineNo + ColNo, LineNo + ColNo + Length
	'��ѡ�����ݷŵ����ӷ�Χ֮��
	SendMessageLNG hwnd, EM_SCROLLCARET, 0&, 0&
End Sub


'��ȡ���λ�ã��кź��кţ������к���ʼ�о�Ϊ0��
'Mode = False��GetCurPos.y Ϊ�кţ�GetCurPos.x Ϊ����յ���к�(���׿�ʼ����ַ���)
'Mode = True�� GetCurPos.y Ϊ�кţ�GetCurPos.x Ϊ���ʼ����к�(���׿�ʼ����ַ���)
'strText = ""����ת�� GetCurPos.x Ϊ�ַ��������� Win7 �����°汾
'strText <> ""�� ת�� GetCurPos.x Ϊ�ַ��������� Win10 �����ϰ汾
'strText Ϊ�ı����е��ı��������� GetTextBoxString(hwnd) ����ȡ
Private Function GetCurPos(ByVal hwnd As Long,ByVal Mode As Boolean,Optional ByVal strText As String) As POINTAPI
	If Mode = False Then
		'��ȡ����������λ�����ı��е��ֽ���ƫ��
		SendMessage hwnd, EM_GETSEL, 0&, GetCurPos
	Else
		'��ȡ�������λ�����ı��е��ֽ���ƫ��(��16λ = ���ʼ��,��16λ = ����յ�)���������������
		GetCurPos.x = SendMessageLNG(hwnd, EM_GETSEL, 0&, 0&)	'��λ����λ����󷵻�ֵΪ65535�����򷵻� -1
		If GetCurPos.x > -1 Then
			'Int hi = DWORD / 0x10000; Int low = DWORD And 0xffff;
			If Mode = False Then
				GetCurPos.x = GetCurPos.x / &H10000	'��λ������յ�
			Else
				GetCurPos.x = GetCurPos.x And 65535	'��λ�����ʼ��
			End If
		Else
			SendMessage hwnd, EM_GETSEL, 0&, GetCurPos
		End If
	End If
	If strText <> "" Then
		'ת�� GetCurPos.x Ϊ�ַ�������ΪWin10 �� EM_LINEFROMCHAR ���ַ���ƫ�ƻ�ȡ�к�
		ReDim tmpByte(0) As Byte
		'tmpByte = StringToByte(strText,GetACP)
		tmpByte = StrConv(strText,vbFromUnicode)
		ReDim Preserve tmpByte(GetCurPos.x + 1) As Byte
		'GetCurPos.x = Len(ByteToString(tmpByte,,GetACP))
		GetCurPos.x = Len(StrConv$(tmpByte,vbUnicode))
	End If
	'��ù�������е��к�
	GetCurPos.y = SendMessageLNG(hwnd, EM_LINEFROMCHAR, GetCurPos.x, 0&)
	'���ع�������е��ַ���
	GetCurPos.x = GetCurPos.x - SendMessageLNG(hwnd, EM_LINEINDEX, GetCurPos.y, 0&)
End Function


'��ȡ��������е������ı�
Private Function GetCurPosLine(ByVal hwnd As Long,ByVal LineNo As Long) As String
	Dim Length As Long
	'��ȡ��������е����ַ����ı��е��ַ���ƫ��
	Length = SendMessageLNG(hwnd, EM_LINEINDEX, LineNo, 0&)
	'��ȡ��������е��ı�����(�ַ���)
	Length = SendMessageLNG(hwnd, EM_LINELENGTH, Length, 0&)
	If Length < 1 Then Exit Function
	'Ԥ��ɽ����ı����ݵ��ֽ�������Ԥ�ȸ��ո�
	ReDim byteBuffer(Length * 2 + 1) As Byte
	'��������� 1024 ���ַ�
	byteBuffer(1) = 4
	'��ȡ��������е��ı��ֽ�����
	SendMessage hwnd, EM_GETLINE, LineNo, byteBuffer(0)
	'ת��Ϊ�ı�����������ַ�
	GetCurPosLine = Replace$(StrConv$(byteBuffer,vbUnicode),vbNullChar,"")
End Function


'��ȡ�����ִ��Ĳ��ҷ�ʽ
'GetFindMode = 0 ���棬= 1 ͨ���, = 2 ������ʽ
Public Function GetFindMode(FindStr As String) As Long
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


'�����ִ����ƶ����λ��
'Mode = False ��ͷ��β�������β��ͷ
'FindCurPos > 0 ���ҵ���= 0 δ�ҵ�, = -1 ��ͬλ��(ֻ�ҵ�һ��), = -2 ͨ����﷨����, = -3 ������ʽ�﷨����
Private Function FindCurPos(ByVal hwnd As Long,ByVal FindStr As String,ByVal Mode As Boolean,Optional strText As String) As Long
	Dim i As Long,n As Long,Lines As Long,Stemp As Integer,sLength As Long,bLength As Long
	Dim ptPos As POINTAPI,bkPos As POINTAPI,Matches As Object
	On Error GoTo errHandle
	'��ȡ���λ�ü��ı������ִ�������
	Lines = SendMessageLNG(hwnd, EM_GETLINECOUNT, 0&, 0&) - 1 'Lines �� 1 Ϊ���
	If Lines = -1 Then Exit Function
	'���������ݵĲ��ҷ�ʽ
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
		'ת��ͨ���Ϊ������ʽģ��
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
	'��ʼ��������ʽ
	With RegExp
		.Global = Mode
		.IgnoreCase = False
		.Pattern = FindStr
	End With
	'Win10 �£���Ҫ�õ��༭���е������ִ�
	'If StrToLong(GetWindowsVersion()) < 62 Then
		strText = ""	'Win10 ���°汾����Ҫ
	'ElseIf strText = "" Then
	'	strText = GetTextBoxString(hwnd)
	'End If
	'��ȡ���λ��  .y Ϊ�к� (��� = 0)��.x Ϊ�����е��ַ���ƫ��(��� = 0)
	ptPos = GetCurPos(hwnd,Mode,strText)
	'������ʼ�㣬�����ж��ҵ���λ���Ƿ���ǹ�����λ��
	bkPos = ptPos
	'�����ִ�
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


'�����ִ�
'Mode = 0 ���棬= 1 ͨ���, = 2 ������ʽ
'FilterStr = 1 ���ҵ���= 0 δ�ҵ�, = -1 �������, = -2 ͨ����﷨���� = -3 ������ʽ�﷨����
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


'�����ı����ı��������ڿ�ʼ�������� lpPoint �������
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


'���������ʽ�Ƿ���ȷ
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


'����Ƿ���Ҫ���������
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


'ת���ַ��������ȡ˳��
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


'�ж��Ƿ���Ҫ���������õ�ַ���Զ����ַ�����
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


'��ȡ�ַ�������������
'Index = 1 �ϸ��飬= 2 ȥ�����������ƣ�= 3 ��������ǰ�� "?"�������ִ�������= 0 �� > 3 ת��Ϊǰ�÷��ͽ�����������ʽ
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


'�ϲ��ַ�������������
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


'����ַ�������������
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


'ת���ַ����������������Ա�������ֽ����������ַ�
Public Function ConvertStrEndCharSet(ByVal UseEndChar As String,ByVal Mode As Boolean) As String
	If Trim$(UseEndChar) = "" Then Exit Function
	If Mode = False Then
		ConvertStrEndCharSet = Replace$(UseEndChar,")" & ValJoinStr,")" & JoinStr)
	Else
		ConvertStrEndCharSet = Replace$(UseEndChar,")" & JoinStr,")" & ValJoinStr)
	End If
End Function


'ת���ִ��б�Ϊ���б�ֵΪ Key ���ֵ�
'DelAccKey <> 0 ɾ����ݼ������ת����ִ����ֵ䣬����ɾ�������ص��ִ��б���ת��
'Mode = 0 ת�岻�����ִ��б�Mode = 1 ��ת�岻�����ִ��б�Mode = 2 ��ת�岢���ر�ת��Ϊ�ֵ�� List �ִ��б�
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


'ɾ����ݼ� (���������ŵ��������ԵĿ�ݼ�)
Public Function DelAccKey(ByVal strText As String) As String
	Dim i As Long,j As Long,TempList() As String,Matches As Object
	On Error GoTo errHandle
	DelAccKey = strText
	i = InStr(strText,"&")
	If i < 1 Then Exit Function
	'��ʼ��������ʽ
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


'��;����ʮ����ת��Ϊ������
'���룺Dec(ʮ������)
'�����DECtoBIN(��������)
'����������Ϊ2147483647,��������Ϊ1111111111111111111111111111111(31��1)
Private Function DECtoBIN(Dec As Long) As String
	Do While Dec > 0
		DECtoBIN = Dec Mod 2 & DECtoBIN
		If Dec < 2 Then Exit Do
		Dec = Dec \ 2
	Loop
End Function


'λ����
Private Function SHL(nSource As Long, n As Byte) As Double
	On Error GoTo ExitFunction:
	SHL = nSource * 2 ^ n
	ExitFunction:
End Function


'λ����
Private Function SHR(nSource As Long, n As Byte) As Long
	On Error GoTo ExitFunction:
	SHR = nSource / 2 ^ n
	ExitFunction:
End Function


'���ָ����λ
Private Function GetBits(nSource As Long, n As Byte) As Long
	On Error GoTo ExitFunction:
	GetBits = nSource And 2 ^ n
	ExitFunction:
End Function


'����ָ����λ
Private Function SetBits(nSource As Long, n As Byte) As Long
	On Error GoTo ExitFunction:
	SetBits = nSource Or 2 ^ n
	ExitFunction:
End Function


'���ָ����λ
Private Function ResetBits(nSource As Long, n As Byte) As Long
	On Error GoTo ExitFunction:
	ResetBits = nSource And Not 2 ^ n
	ExitFunction:
End Function


'��ת�ֽ�˳�� (16-bit)
Private Function UInt16ReverseBytes(ByVal Value As Integer) As Integer
	UInt16ReverseBytes = SHL((Value And &HFF),8) Or SHR((Value And &HFF00),8)
End Function


'��ת�ֽ�˳�� (32-bit)
Private Function UInt32ReverseBytes(ByVal Value As Long) As Long
	UInt32ReverseBytes = SHL((Value And &HFF),24) Or SHL((Value And &HFF00),8) Or _
         SHR((Value And &HFF0000),8) Or SHR((Value And &HFF000000),24)
End Function


'��ȡ N ���ַ�����������Ӵ�
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


'���㲢������ Long ��ֵ���ֵ�Ƿ����
'����ֵ: =0 ���, >0 δ���
Public Function CheckLongPlus(ByVal Long1 As Long,ByVal Long2 As Long) As Variant
	On Error GoTo errHandle
	CheckLongPlus = Long1 + Long2
	Exit Function
	errHandle:
	Err.Clear
	CheckLongPlus = 0
End Function


'��������Ƿ��Ѿ���ʼ��
'����ֵ:TRUE �Ѿ���ʼ��, FALSE δ��ʼ��
Public Function CheckArrEmpty(ByRef MyArr As Variant) As Boolean
	On Error Resume Next
	If UBound(MyArr) >= 0 Then CheckArrEmpty = True
	Err.Clear
End Function


'��������ĳ������ID�Ƿ����
'����ֵ:TRUE �Ѿ���ʼ��, FALSE δ��ʼ��
Private Function CheckArrID(ByRef MyArr As Variant,ID As Long) As Boolean
	Dim i As Variant
	On Error Resume Next
	i = MyArr(ID)
	CheckArrID = (Err.Number = 0)
	Err.Clear
End Function


'����ִ������Ƿ�Ϊ�գ��ǿշ��� True
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


'��� INI ���������Ƿ�Ϊ�գ��ǿշ��� True
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


'��� HCS �ִ����������Ƿ�Ϊ�գ��ǿշ��� True
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


'����Զ����ִ����͵������Ƿ�Ϊ�գ��ǿշ��� True
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


'�����б��������Ŀ��
'Mode = 0  �����= 1 ѡ���= 2 δѡ����
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


'��ȡѡ���б����Ŀ������
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


'�����б���е�һ���ɼ��������
Public Function GetListBoxTopIndex(ByVal hwnd As Long) As Long
	On Error GoTo errHandle
	GetListBoxTopIndex = SendMessageLNG(hwnd, LB_GETTOPINDEX, 0&, 0&)
	Exit Function
	errHandle:
	GetListBoxTopIndex = 0
End Function


'�����б���е�һ���ɼ��������
Private Function SetListBoxTopIndex(ByVal hwnd As Long,ByVal TopItem As Long) As Boolean
	On Error GoTo errHandle
	SendMessageLNG(hwnd, LB_SETTOPINDEX, TopItem, 0&)
	SetListBoxTopIndex = True
	errHandle:
End Function


'ѡ��ָ�����б����Ŀ
'Indexs = -1 ȫѡ������ѡ��ָ����
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


'����򸽼��б����Ŀ
'InsPos = -1 ���ӵ���󣬷�����뵽ָ������λ��
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


'ɾ���б����Ŀ
'DelPos = -1 ɾ������< -1 ȫ����գ�����ɾ��ָ�������ŵ���Ŀ
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


'�����б����Ŀ
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


'���ضԻ���ĳ���ؼ��е��ִ�
Public Function GetTextBoxString(ByVal hwnd As Long) As String
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
Public Function SetTextBoxString(ByVal hwnd As Long,ByVal StrText As String,Optional ByVal Mode As Boolean) As Boolean
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


'��ȡ�ִ������ش�С
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


'��������Ƿ�Ϊ�գ��ǿշ��� True
Public Function CheckFont(LF As LOG_FONT) As Boolean
	If ReSplit(StrConv$(LF.lfFaceName,vbUnicode),vbNullChar,2)(0) <> "" Then CheckFont = True
End Function


'��ȡ�������ƺ��ֺ�
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


'�Ƚ϶������������Ƿ���ͬ������ͬ���� True
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


'�Ƚ϶��������Ƿ���ͬ������ͬ���� True
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


'����ϵͳ����Ի���ѡ�����壬ȷ��ʱ���ط���
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
		.lpLogFont = VarPtr(LF2)		'LogFont�ṹ��ַ
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


'�������壬����������
Public Function CreateFont(ByVal hwnd As Long,LF As LOG_FONT) As Long
	Dim LF2 As LOG_FONT
	LF2 = LF
	If CheckFont(LF2) = False Then
		If hwnd = 0 Then Exit Function
		GetObjectAPI(SendMessageLNG(hwnd,WM_GETFONT,0,0),Len(LF2),VarPtr(LF2))
	End If
	CreateFont = CreateFontIndirect(LF2)
End Function


'�ػ������Ի���
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


'�ػ��Ի����ı�
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


'�ִ����ͽ�����ʶ��ת������ʽģ��
'Setting = "0" ʱ��StrTypeEndChar2Pattern ���� 2 ά������
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


'��ȡ Blob ������ѹ�������賤��
'���� CorSigCompressLength = ǩ������
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


'Blob ������ѹ��
'���� CorSigCompressByte = ѹ������ֽ����飬Length = �ֽڳ���
Public Function CorSigCompressByte(ByVal Length As Long) As Byte()
	If Length > &H3FFF Then
		CorSigCompressByte = Val2Bytes((&HC0000000 Or Length),4,True)
	ElseIf Length > &H7F Then
		CorSigCompressByte = Val2Bytes((&H8000 Or Length),2,True)
	Else
		CorSigCompressByte = Val2Bytes(Length,1,True)
	End If
End Function


'Blob ������ѹ��
'���� CorSigCompressData = ǩ�����ȣ�Length = �ֽڳ���
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


'Blob �����Ƚ�ѹ��
'���� CorSigUncompressData = ѹ�����ȣ�Length = �����ȱ�ʶ������ֽڳ��ȣ������Ƿ���� > &H7F �ַ���ʶ����
'ÿ�����������ݿ�ͷ������1���������ݿ飬ͨ����λ���㣬������������ݿ��ʵ�ʳ���
'�����һ���ֽ����λΪ0��������ݿ鳤��Ϊ1���ֽ�
'�����һ���ֽ����λΪ10��������ݿ鳤��Ϊ2���ֽ�
'�����һ���ֽ����λΪ110��������ݿ鳤��Ϊ4���ֽ�
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


'���� PE �ļ�����
Public Function Alignment(ByVal orgValue As Long,ByVal AlignVal As Long,ByVal RoundVal As Long) As Long
	If AlignVal < 1 Then
		Alignment = orgValue
	Else
		Alignment = IIf(orgValue Mod AlignVal = 0,orgValue,AlignVal * ((orgValue \ AlignVal) + RoundVal))
	End If
End Function


'���� PSL 2015 �����ϰ汾������� Split ������ֿ��ַ���ʱ����δ��ʼ������Ĵ���
Public Function ReSplit(ByVal textStr As String,Optional ByVal Sep As String = " ",Optional ByVal Max As Integer = -1) As String()
	If textStr = "" Then
		ReDim TempList(0) As String
		ReSplit = TempList
	Else
		ReSplit = Split(textStr,Sep,Max)
	End If
End Function


'�����ִ�����Ϊ�ִ�����Ϊ Join ����Ч��̫��
'Mode = False �� Join ������ʽ���ӣ����治�����ӷ��������������ӷ�
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


'�����ֵ�Ϊ����
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


'��ȡ�ֵ�ָ������ֵ
Public Function GetDicVal(ByVal Dic As Object,ByVal Key As Variant,Optional ByVal DefaultValue As Variant) As Variant
	If Dic.Exists(Key) Then
		GetDicVal = Dic.Item(Key)
	Else
		GetDicVal = DefaultValue
	End If
End Function


'��ȡ���ػ����Լ�����ҳ��Ϣ
Public Function GetLCIDInfo(LCID As Long,LCType As Long) As String
	Dim iRet As Long
	iRet = GetLocaleInfo(LCID, LCType, "", 0)
	If iRet = 0 Then Exit Function
	GetLCIDInfo = String$(iRet, 0)
    If GetLocaleInfo(LCID, LCType, GetLCIDInfo, iRet) = 0 Then Exit Function
    GetLCIDInfo = Replace$(GetLCIDInfo,vbNullChar,"")
End Function


'��ȡ�ִ��е����ַ���(������ʽ)
'Patrn Ϊ������ʽģ��
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


'�ϲ����������ϵ���޷�ƥ��� \x ת���������ʽΪ [\x-\x] ��ʽ�ı��ʽ
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


'������������Ϣ
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


'��ȫ�˳�����
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


'��ȡ���ֽ��������ҵ���ƥ��������б�(������ʽ��ʽ)
'ע�⣺StartPos��EndPos ��Ϊ�� 0 ��ʼ�ĵ�ַ
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


'���������ǵ����ֽ�λ�ã������طǵ����ֽڿ�ʼλ��
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


'���������ǵ����ֽ�λ�ã����������һ���ǵ����ֽڽ���λ��
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


'���������ǵ����ֽ�λ�ã����������һ���ǵ����ֽڽ���λ��
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


'��������ָ�������Ŀ��ֽ�λ�ã������ؿ��ֽڿ�ʼλ�ã�Bit Ϊ��С���ֽ���
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


'��������ָ�������Ŀ��ֽ�λ�ã����������һ�������ֽڽ���λ�ã�Bit Ϊ��С���ֽ���
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


'��������ָ�������Ľ�����λ�ã�CodePage Ϊ����ҳ
'getEndByteRegExp�����ؽ�������ʼλ�� (fType = False) �����λ�� + 1 (fType = True)
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


'��������ָ�������Ŀ��ֽ�λ�ã�CodePage Ϊ����ҳ
'getEndByteRevRegExp�����ؽ���������λ��
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


' ����ļ�����
' ----------------------------------------------------
' ANSI      �޸�ʽ����
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
' �����е� XX ��ʾ����ʮ�������ַ�

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


' ��ȡ�ı��ļ�
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


' д���ı��ļ�
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
				'ȥ������ BOM ��ʽ�� BOM
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


'���� Adodb.Stream ʹ�õĴ���ҳ����
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


'����Դ�б��л�ȡָ����Ŀ��Ŀ���б�
'Mode = 0 ��������ֱ���˳�����Mode = 1 ������������Ƿ��˳�������ʾ
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


'д������
Public Function WriteSettings(ByVal WriteType As String) As Boolean
	Dim i As Long,n As Long,Temp As String,TempArray() As String
	On Error GoTo ExitFunction
	SaveSetting(AppName,"Option","Version",Version)
	'���������������
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
	'�����ִ���ȡѡ��
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
	'�����Զ���������
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
	'�����Զ����ִ�����
	If WriteType = "Sets" Or WriteType = "All" Then
		'ɾ��ԭ������
		Temp = GetSetting(AppName,"Option","StringTypeCount","")
		If Temp <> "" Then
			On Error Resume Next
			For i = 0 To StrToLong(Temp)
				DeleteSetting(AppName,"StringType_" & CStr$(i))
			Next i
			On Error GoTo 0
		End If
		'д����������
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
	'�����Զ��������㷨
	If WriteType = "Sets" Or WriteType = "All" Then
		'ɾ��ԭ������
		Temp = GetSetting(AppName,"Option","RefAlgorithmCount","")
		If Temp <> "" Then
			On Error Resume Next
			For i = 0 To StrToLong(Temp)
				DeleteSetting(AppName,"RefAlgorithm_" & CStr$(i))
			Next i
			On Error GoTo 0
		End If
		'д����������
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
	'�����Զ��幤��
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
	'�����Զ������ҳ
	If WriteType = "Sets" Or WriteType = "All" Then
		'ɾ��ԭ������
		On Error Resume Next
		DeleteSetting(AppName,"Languages")
		On Error GoTo 0
		'д����������
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
	'��������ִ��б�����
	If WriteType = "FilterStrDic" Or WriteType = "All" Then
		SaveSetting(AppName,"GetString","UseMatchFilterListFile",ExtractSet(34))
		SaveSetting(AppName,"GetString","FilterListFileSelect",ExtractSet(39))
		SaveSetting(AppName,"GetString","FilterListFile",Mid$(ExtractSet(35),InStrRev(ExtractSet(35),"\") + 1))
	End If
	'���汣���ִ��б�����
	If WriteType = "ReserveStrDic" Or WriteType = "All" Then
		SaveSetting(AppName,"GetString","UseMatchReserveListFile",ExtractSet(36))
		SaveSetting(AppName,"GetString","ReserveListFileSelect",ExtractSet(40))
		SaveSetting(AppName,"GetString","ReserveListFile",Mid$(ExtractSet(37),InStrRev(ExtractSet(37),"\") + 1))
	End If
	'����Ի�����������
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


'�����ִ��������ظ�����
'Mode = False �����������������������
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


'������ֵ�������ظ�����
'Mode = False �����Ϊ����������Ϊ����
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


'��ȡ INI �ļ�
'Mode = 0 ɾ����Ŀֵǰ��ո�˫���ţ���ת����Ŀֵ
'Mode = 1 ɾ����Ŀֵǰ�ո񣬲�ת��
'Mode = 2 ɾ����Ŀֵǰ��ո񣬲�ת��
'Mode > 2 ��ɾ����Ŀֵǰ��ո񣬲�ת��
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


'�ָ�������������ƺ�ֵ
'GetType = 0 ��ȡ���������б�
'GetType = 1 ��ȡ���� ID �б�
'GetType = 2 ��ȡ����ҳֵ�б� (��ͬ�Ĵ���ҳ������)
'GetType = 3 ��ȡ����ҳֵ - ����ҳ���Ƹ�ʽ���б� (��ͬ�Ĵ���ҳ�����ˣ�Ĭ�ϴ���ҳ��û�еĴ���ҳ����ӣ�������ʾ)
'GetType = 4 ��ȡ����ҳֵ�б� (��ͬ�Ĵ���ҳ�����ˣ�Ĭ�ϴ���ҳ��û�еĴ���ҳ����ӣ�������ʾ)
'GetType = 5 ��ȡ����ҳ�����б� (��ͬ�Ĵ���ҳ�����ˣ�Ĭ�ϴ���ҳ��û�еĴ���ҳ����ӣ�������ʾ)
'GetType = 6 ��ȡ����ҳ���� + vbNullChar + ����ҳֵ���б� (��ͬ�Ĵ���ҳ�����ˣ�Ĭ�ϴ���ҳ��û�еĴ���ҳ����ӣ�������ʾ)
'GetType = 7 ��ȡ����ҳֵ�б� (��ͬ�Ĵ���ҳ�����ˣ�Ĭ�ϴ���ҳ��û�еĴ���ҳ����ӣ��Զ������ȥ����������ʾ)
'GetType = 8 ��ȡ����ҳ�����б� (��ͬ�Ĵ���ҳ�����ˣ�Ĭ�ϴ���ҳ��û�еĴ���ҳ����ӣ��Զ������ȥ����������ʾ)
'Mode = False ���˵� Unicode ������Ŀ�����򲻹���
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
		CodePage(0) = CP_UNKNOWN		'δ֪ (�Զ����) = -1
		CodePage(1) = CP_WESTEUROPE		'������ 1 (ANSI) = 1252
		CodePage(2) = CP_EASTEUROPE	    '������ 2 (��ŷ) = 1250
		CodePage(3) = CP_RUSSIAN		'������� (˹����) = 1251
		CodePage(4) = CP_GREEK			'ϣ���� = 1253
		CodePage(5) = CP_TURKISH		'������ 5 (������) = 1254
		CodePage(6) = CP_HEBREW			'ϣ������ = 1255
		CodePage(7) = CP_ARABIC			'�������� = 1256
		CodePage(8) = CP_BALTIC			'���޵ĺ��� = 1257
		CodePage(9) = CP_VIETNAMESE		'Խ���� = 1258
		CodePage(10) = CP_JAPAN			'���� = 932
		CodePage(11) = CP_CHINA			'�������� = 936
		CodePage(12) = CP_KOREA			'���� = 949
		CodePage(13) = CP_TAIWAN 		'�������� = 950
		CodePage(14) = CP_THAI 			'̩�� = 874
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


'���ı��ļ�
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


'����༭����
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


'��ȡ�༭����Ի�����
Private Function CommandInputDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,x As Long,y As Long,Path As String
	Dim Temp As String,TempList() As String,TempArray() As String,MsgList() As String
	Select Case Action%
	Case 1 ' �Ի��򴰿ڳ�ʼ��
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
		'���õ�ǰ�Ի�������
		If CheckFont(LFList(0)) = True Then
			x = CreateFont(0,LFList(0))
			If x = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,x,0)
			Next i
		End If
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
		CommandInputDlgFunc = True ' ��ֹ���°�ť�رնԻ��򴰿�
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
	Case 3 ' �ı��������Ͽ��ı�������
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


'�༭�ı��ļ�
'Mode = True �༭ģʽ��������ļ��ɹ������ַ������ True
'Mode = False �鿴��ȷ���ַ�����ģʽ��������ļ��ɹ����� [ȷ��] ��ť�����ַ������ True
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


'�༭�Ի�����
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
		'���ñ༭�ؼ��е�����ı����ȣ�ԭ��󳤶�Ϊ30000���ַ���˫�ֽ��ַ���1����
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
    	'���õ�ǰ�Ի�������
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
		EditFileDlgFunc = True '��ֹ���°�ť�رնԻ��򴰿�
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
			'���ñ༭�ؼ��е�����ı����ȣ�ԭ��󳤶�Ϊ30000���ַ���˫�ֽ��ַ���1����
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
				'���ñ༭�ؼ��е�����ı����ȣ�ԭ��󳤶�Ϊ30000���ַ���˫�ֽ��ַ���1����
				SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox")),FileLen(Temp),True
				DlgText "InTextBox",ReadTextFile(Temp,Code)
				If DlgText("InTextBox") <> "" Then DlgText "ReadButton",MsgList(9)
			Else
				DlgText "InTextBox",""
				DlgText "ReadButton",MsgList(8)
				'���ñ༭�ؼ��е�����ı����ȣ�ԭ��󳤶�Ϊ30000���ַ���˫�ֽ��ַ���1����
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
			'��Ӳ�������
			GetHistory(TempList,"FindStrings","EditFileDlg")
			If InsertArray(TempList,DlgText("FindTextBox"),0,True) = True Then
				WriteHistory(TempList,"FindStrings","EditFileDlg")
				DlgListBoxArray "FindTextBox",TempList()
				DlgText "FindTextBox",TempList(0)
			End If
			'DlgFocus("InTextBox")  '���ý��㵽�ı���2016 �����˸��ʹ�ù��λ���Ƶ���ǰ��
			n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
			SendMessageLNG(n,WM_SETFOCUS,0,0)  '���ý��㵽�ı���
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
			'��Ӳ�������
			GetHistory(TempList,"FindStrings","EditFileDlg")
			If InsertArray(TempList,DlgText("FindTextBox"),0,True) = True Then
				WriteHistory(TempList,"FindStrings","EditFileDlg")
				DlgListBoxArray "FindTextBox",TempList()
				DlgText "FindTextBox",TempList(0)
			End If
			'���������ݵĲ��ҷ�ʽ
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
					UseStrList(n) = "��" & CStr$(i + 1) & MsgList(10) & "��" & AllStrList(i)
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
						Temp = "^��[0-9]+" & MsgList(10) & "��"
						For i = 0 To UBound(TempArray)
							If CheckStrRegExp(TempArray(i),Temp,0,2) = True Then
								TempList = ReSplit(TempArray(i),MsgList(10) & "��",2)
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
			'���ñ༭�ؼ��е�����ı����ȣ�ԭ��󳤶�Ϊ30000���ַ���˫�ֽ��ַ���1����
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
				'���ñ༭�ؼ��е�����ı����ȣ�ԭ��󳤶�Ϊ30000���ַ���˫�ֽ��ַ���1����
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
				DlgFocus("FindTextBox")  '���ý��㵽�ı���
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
	Case 3 ' �ı��������Ͽ��ı�������
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
				'���ñ༭�ؼ��е�����ı����ȣ�ԭ��󳤶�Ϊ30000���ַ���˫�ֽ��ַ���1����
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
	Case 4 ' ���㱻����
		Select Case DlgItem$
		Case "InTextBox"
			i = Len(Clipboard)
			If i < Len(DlgText("InTextBox")) * 2 Then Exit Function
			'���ñ༭�ؼ��е�����ı����ȣ�ԭ��󳤶�Ϊ30000���ַ���˫�ֽ��ַ���1����
			SetTextBoxLength GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox")),LenB(DlgText("InTextBox")) + i,True
		End Select
	Case 6 ' ������ݼ�
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
				DlgFocus("FindTextBox")  '���ý��㵽�ı���
				DlgText "FindTextBox",InsertStr(GetFocus(),DlgText("FindTextBox"), _
						Mid$(MsgList(i),InStrRev(MsgList(i),vbTab) + 1))
			End If
		Case 3
			If DlgText("FindTextBox") = "" Then
				MsgBox MsgList(11),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
			'DlgFocus("InTextBox")  '���ý��㵽�ı���2016 �����˸��ʹ�ù��λ���Ƶ���ǰ��
			n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
			SendMessageLNG(n,WM_SETFOCUS,0,0)  '���ý��㵽�ı���
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
				'���������ݵĲ��ҷ�ʽ
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
						UseStrList(n) = "��" & CStr$(i + 1) & MsgList(10) & "��" & AllStrList(i)
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
							Temp = "^��[0-9]+" & MsgList(10) & "��"
							For i = 0 To UBound(TempArray)
								If CheckStrRegExp(TempArray(i),Temp,0,2) = True Then
									TempList = ReSplit(TempArray(i),MsgList(10) & "��",2)
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
				'���ñ༭�ؼ��е�����ı����ȣ�ԭ��󳤶�Ϊ30000���ַ���˫�ֽ��ַ���1����
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
				'���ñ༭�ؼ��е�����ı����ȣ�ԭ��󳤶�Ϊ30000���ַ���˫�ֽ��ַ���1����
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


'���ںͰ���
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


'�༭�������ƺ�����������ƶԻ�����
Public Function CommonDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,j As Long
	Select Case Action%
	Case 1 ' �Ի��򴰿ڳ�ʼ��
		'���õ�ǰ�Ի�������
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	End Select
End Function


'��ȡ�����ʷ��¼
'Separator Ϊ��ʱ�����÷�ʽ��ȡ
'Separator ��Ϊ��ʱ��׷�ӷ�ʽ��ȡ������ Separator �����Чֵ�������ظ�
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


'���������ʷ��¼
'Mode = False д����� 10 ����¼������д�����м�¼����ɾ�� DataList ��û�еļ�¼
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


'�����������
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


'�༭��������
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


'ɾ���ı�������Ŀ
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


'ɾ���ı�������Ŀ�������ֵ�
'Mode = False ��ת�壬����ת��
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


'ɾ����ֵ������Ŀ
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


'�����ִ�Сд���ַ��滻 (����δ�滻�ַ��Ĵ�Сд)
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


'�������ݵ�����
'Mode = True �������ظ���������λ Data ����ǰ��
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


'��ӹ�������
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


'��ȡ������������Ĺ��������б�
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


'ɾ���Զ��幤��������Ŀ
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


'��ȡ�������ļ�
'BOM = False ��鲢ȥ�� BOM��������� BOM
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


'д��������ļ�
'BOM = False ��鲢д�� BOM������д�� BOM
'Mode = False ɾ���ļ�������д�룬���� File Ϊ�ļ���ʱ����
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


'�Ƚ϶����ִ������Ƿ���ͬ������ͬ���� True
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


'���汾�źʹΰ汾���ж�ϵͳ�汾
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


'ת��ʮ���ƺ�ʮ������ֵΪ�ַ�
'MaxVal = 0 ��ֵ����Ӧ�еĳ��ȣ�> 0 ���ļ���С�����λ����< 0 ��ָ��λ��
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


'ת��ʮ���ƺ�ʮ�������ַ�Ϊʮ����ֵ
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


'����ת��ʮ�����ַ�Ϊʮ�������ַ�
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


'����ת��ʮ�������ַ�Ϊʮ�����ַ�
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


'ÿ�����ַ��ո�ָ�
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


'�ַ���ת Hex
'FillMode: 0 = ��ʵ�������ֽ�ת��
'          1 = ����ʱ����ضϣ�������������ֽ�����
'          2 = ����ʱǰ��ضϣ�����ǰ�������ֽ�����
'          3 = ����ʱ����ضϣ���������ÿ��ֽ�����
'          4 = ����ʱǰ��ضϣ�����ǰ���ÿ��ֽ�����
'ByteLength Ϊ��Ҫ���ֽڳ��ȣ����ڽض�ʱ��Ҫ
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


'�������� Hex �ִ��Ƿ����Ҫ��
'���� CheckHex = 0 �ϸ�, = 1 HEX �����ַ�������, = 2 ����������, = 3 �����Ƿ��ַ� = 4 û��Ҫת���Ĵ���(�� Ignore = 1 �����)
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


'��ȡ�ļ������ͣ�PE ���� MAC ���Ƿ� PE �ļ�
Public Function GetFileFormat(ByVal FilePath As String,ByVal Mode As Long,FileType As Integer) As String
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


'����ļ��Ƿ��ѱ��򿪻�ռ��
Public Function IsOpen(ByVal strFilePath As String,Optional ByVal Continue As Long = 2,Optional ByVal WaitTime As Double = 0.5) As Boolean
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


'��ȡ�ļ��Ĵ��������ʡ��޸�����
'Mode = 0 ��������
'Mode = 1 ��������
'Mode = 2 �޸�����
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


'�޸��ļ��Ĵ���ʱ�䡢����ʱ����޸�ʱ��
Public Function SetFileCreatedDate(ByVal strFilePath As String,ByVal CreatedDate As Date) As Boolean
	Dim lhFile As Long,udtFileTime As FILETIME
	Dim udtLocalTime As FILETIME,udtSystemTime As SYSTEMTIME

	'ת��ʱ��Ϊ SYSTEMTIME
	udtSystemTime.wYear = Year(CreatedDate)
	udtSystemTime.wMonth = Month(CreatedDate)
	udtSystemTime.wDay = Day(CreatedDate)
	udtSystemTime.wDayOfWeek = Weekday(CreatedDate)
	udtSystemTime.wHour = Hour(CreatedDate)
	udtSystemTime.wMinute = Minute(CreatedDate)
	udtSystemTime.wSecond = Second(CreatedDate)
	udtSystemTime.wMilliseconds = 0

	'ת��ϵͳʱ��Ϊ����ʱ��
	If SystemTimeToFileTime(udtSystemTime, udtLocalTime) = 0 Then
		DisplayError "SystemTimeToFileTime",GetLastError()
		Exit Function
	End If

	'ת������ʱ��Ϊ GMT ʱ��
	If LocalFileTimeToFileTime(udtLocalTime, udtFileTime) = 0 Then
		DisplayError "LocalFileTimeToFileTime",GetLastError()
		Exit Function
	End If

	'���ļ���ȡ�ļ����
	lhFile = CreateFile(strFilePath, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
	If lhFile = 0 Then
		DisplayError "CreateFile",GetLastError()
		Exit Function
	End If

	'�����ļ������ں�ʱ������
	If SetFileTime(lhFile, udtFileTime, udtFileTime, udtFileTime) = 0 Then
		DisplayError "SetFileTime",GetLastError()
		If CloseHandle(lhFile) = 0 Then
			DisplayError "CloseHandle",GetLastError()
		End If
		Exit Function
	End If

	'�ر��ļ��������ǳɹ�
	If CloseHandle(lhFile) = 0 Then
		DisplayError "CloseHandle",GetLastError()
		Exit Function
	End If
	SetFileCreatedDate = True
End Function


'�����ļ�
'ImageSize = 0 ���ļ��ĳ�ʼ��С�򿪣�����ָ����С��
'ReadOnly = 0 ��ֻ����ʽ�򿪣������д��ʽ��
'ImageByte = 0 ����ȡ�ֽ�����ֻ��ʼ��(���淽ʽ��ȡ�����ֽ�)������ ImageByte ָ����С��ȡ
'Mode < 0 ֱ�ӷ�ʽ��Mode = 0 ���淽ʽ��Mode > 0 ӳ�䷽ʽ
'IsPE = 0 ��һ���ļ�ӳ�䣬���� PE �ļ�ӳ��(ÿ���ڶ���)
'LoadFile = -2 ��ʧ�ܣ�����ʵ�ʴ򿪷�ʽ
'LoadedImage ���ļ����ȡ������
Public Function LoadFile(ByVal strFilePath As String,LoadedImage As FILE_IMAGE,ByVal ImageSize As Long,ByVal ReadOnly As Long, _
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


'ӳ���ļ�
'MapSize = 0 ���ļ���ʼʱ�Ĵ�Сӳ�䣬����ָ����Сӳ��
'ReadOnly = 0 ֻ����ʽ�������д��ʽ
'SizeOfFile = 0 ��ȡ�ļ���ʼʱ�Ĵ�С�����򲻻�ȡ
'IsPE = 0 ��һ���ļ�ӳ�䣬���� PE �ļ�ӳ��(ÿ���ڶ���)
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


'�Ƚ϶����ļ����ļ��ڿ�ʼ��ַ�ʹ�С�Ƿ�һ�£���ͬ���� 0
'SecIDList Ϊ����ʱ������ƶ���
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


'��ȡ�ļ���������
'Mode = 0 ���ƫ�Ƶ�ַ(���������ؽ�)
'Mode = 1 ���ƫ�Ƶ�ַ(�������ؽ�)
'Mode = 2 �����������ַ(���������ؽ�)
'Mode = 3 �����������ַ(�������ؽ�)
'�����ļ��������š�MinVal��MaxVal ֵ
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


'��ȡ�����ļ�ͷ�������ż���ַ (����ĵ�ַ����ת��Ϊƫ�Ƶ�ַ)
'Mode = 0 ʱ�������� RVA ����Ŀ¼��������
'Mode = 1 ʱ��RVA = RVA ����Ŀ¼������ַ + 1��SkipVal = �� RVA ���Ŀ¼��С��ַ
'Mode > 1 ʱ��RVA = RVA ����Ŀ¼����С��ַ - 1��SkipVal = �� RVA С��Ŀ¼����ַ
'fType = 0 ʱ������ .NET US ����
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


'��ȡ�ڱ��������С��ȴ���С�ĵ�ַ���ڽ�������
'MinID = 0 �� MaxID = 0 ��ȡ�ڱ��������С��ַ���ڽ�������
'MinID = -1 ��ȡ�� MaxID ��С�ĵ�ַ���ڽ�������
'MaxID = -1 ��ȡ�� MinID �ڴ�ĵ�ַ���ڽ�������
'Mode = False �Ƚ�ƫ�Ƶ�ַ������Ƚ���������ַ
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


'���͸��������ļ��������б�
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


'��ȡ�����ļ������ƻ������ַ��Χ���б�
'Mode = 0 ��ȡ���������ε����ƣ�������������
'Mode = 1 ��ȡ���������ε����ƣ���������������
'Mode = 2 ��ȡ���������ε����ƺ����������ƣ�������������
'Mode = 3 ��ȡ���������ε����ƺ����������ƣ���������������
'Mode = 4 ��ȡ���������ε����ƺ͵�ַ��Χ��������������
'Mode = 5 ��ȡ���������ε����ƺ͵�ַ��Χ����������������

'Mode = 6 ��ȡ���������ε����ƺ����������ƺͿ�ʼ��ַ(�˵���ʽ)��������������
'Mode = 7 ��ȡ���������ε����ƺ����������ƺͿ�ʼ��ַ(�˵���ʽ)����������������
'Mode = 8 ��ȡ���������ε����ƺ����������ƺͽ�����ַ(�˵���ʽ)��������������
'Mode = 9 ��ȡ���������ε����ƺ����������ƺͽ�����ַ(�˵���ʽ)����������������

'Mode = 10 ��ȡ���������ε����ƺ�������ID��ϣ�������������

'Mode = 11 ��ȡ���������κ������εĿ�ʼ��ַ(��ֵ)��������������
'Mode = 12 ��ȡ���������κ������εĿ�ʼ��ַ(��ֵ)����������������
'Mode = 13 ��ȡ���������κ������εĽ�����ַ(��ֵ)��������������
'Mode = 14 ��ȡ���������κ������εĽ�����ַ(��ֵ)����������������

'���ӽ������ʽ �������� (���ڵ�ַ)
'���ӽ������ʽ �������� - �ӽ����� (�ӽڵ�ַ)
Public Function getSectionList(SecList() As SECTION_PROPERTIE,Optional ByVal Mode As Integer, _
					Optional ByVal MaxVal As Long,Optional ByVal DisPlayFormat As Boolean) As String()
	Dim i As Integer,j As Integer,k As Integer,n As Integer
	Select Case Mode
	Case 0,1	'��ȡ���������ε�����
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
	Case 2,3	'��ȡ���������ε����ƺ�����������
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
	Case 4,5	'��ȡ���������ε����ƺ͵�ַ��Χ
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
	Case 6,7	'��ȡ���������ε����ƺ����������ƺͿ�ʼ��ַ(�˵���ʽ)
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
	Case 8,9	'��ȡ���������ε����ƺ����������ƺͽ�����ַ(�˵���ʽ)
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
	Case 10		'��ȡ���������ε����ƺ�������ID���
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
	Case 11,12	'��ȡ���������κ������εĿ�ʼ��ַ(��ֵ)
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
	Case 13,14	'��ȡ���������κ������εĽ�����ַ(��ֵ)
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


'��С��������β������ڱ�
'Mode = False ��ƫ�Ƶ�ַ���򣬷�����������ַ����
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


'�������ֽڵĳ��ȿ�����������ַ��������
'Mode = False ��С�������򣬷���Ӵ�С����l = �������߽磬r = ������ұ߽�
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


'�������ֽڵĵ�ַ��β����������ַ��������
'Mode = False ��С�������򣬷���Ӵ�С����l = �������߽磬r = ������ұ߽�
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


'�鿴���ػ��ļ���Ϣ
'Mode = 0 �ɲ鿴�����ļ�����Ϣ
'Mode = 1 ���ɲ鿴�����ļ�����Ϣ
'Mode > 1 �鿴�����ļ��ڴ�С��Ϣ
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
	'���ļ�
	On Error GoTo ErrHandle
	If Dir$(File.FilePath & ".xls") <> "" Then Kill File.FilePath & ".xls"
	FN = FreeFile
	Open File.FilePath & ".xls" For Binary Access Read Write Lock Write As #FN
	'MAC64������£��޷����� 64 λ(8 ���ֽ�)����ֵ��ֻ����16������ʾ
	If File.Magic = "MAC64" Then DisPlayFormat = True
	'д���ļ�������Ϣ
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
	'д��ÿ���ļ��ڵ�ƫ�Ƶ�ַ
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
	'д�����ؽڵ�ƫ�Ƶ�ַ���� PE ��ַ������
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
	'д��ÿ���ļ��ڵ���������ַ
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
	'д�����ؽڵ���������ַ���� PE ��ַ������
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
	'д������Ŀ¼��ַ�������ļ���
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
	'д�� .NET CLR ����Ŀ¼��ַ�������ļ���
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
	'д�� .NET ����ַ�������ļ���
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
	'�鿴����
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
	'������
	ErrHandle:
	On Error Resume Next
	Close #FN
	Err.Source = "NotWriteFile"
	Err.Description = Err.Description & JoinStr & File.FilePath & ".xls"
	Call sysErrorMassage(Err,1)
End Sub


'����ע���ؼ��ֵ�ֵ
'����˵��: KeyRoot - ������, KeyName - ��������, FileName - �������ļ�·�����ļ���(ԭʼ���ݿ��ʽ)
Public Function SaveKey(KeyRoot As REG_KEYROOT, KeyName As String, FileName As String) As Boolean
	Dim hKey As Long
	'Dim lpAttr As SECURITY_ATTRIBUTES	'ע���ȫ����
	'lpAttr.nLength = 50				'���ð�ȫ����Ϊȱʡֵ
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


'����ע���ؼ��ֵ�ֵ
'����˵��: KeyRoot - ������, KeyName - ��������, FileName - ������ļ�·�����ļ���(ԭʼ���ݿ��ʽ)
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


'ʹע��������뵼��
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


'�Զ��嵯���˵�(���2��)�����ز˵�����ı�
'strConfig Ϊ�Զ���Ĳ˵�����,�����������˵����Ӳ˵�֮���� vbNullChar �ָ�
'Mode = ����ģʽ��False = �����ִ���True = ����������
Public Function PopupMenuShow(strConfig() As String,Optional ByVal Mode As Boolean) As Variant
	Dim i As Long,j As Long,n As Long,hMenu As Long,MenuCount As Long
	Dim pt As POINTAPI,subMenuList() As String,MenuText() As String
	'��ȡ��ǰ���λ��
	GetCursorPos pt
	'����һ���յĲ˵�,��ȡ���
	hMenu = CreatePopupMenu()
	'�����Ӳ˵��ľ������
	ReDim hPopMenu(UBound(strConfig)) As Long
	For i = 0 To UBound(strConfig)
		If InStr(strConfig(i),vbNullChar) Then
			subMenuList = Split(strConfig(i),vbNullChar)
			If subMenuList(0) <> "" Then
				'�����Ӽ��˵���
				hPopMenu(n) = CreatePopupMenu()
				For j = 1 To UBound(subMenuList)
					If subMenuList(j) <> "" Then
						MenuCount = MenuCount + 1
						'����˵��ı������ڲ˵��¼�����ʱʶ�����ѡ��Ĳ˵�����
						ReDim Preserve MenuText(MenuCount) As String
						MenuText(MenuCount) = subMenuList(j)
						'����Ӳ˵���
						'����Ǽ���ߣ��� wFlags = MF_SEPARATOR
						'���ҪCheck���� wFlags = MF_STRING + MF_CHECKED��������ã����ټ� MF_GRAYED
						If LCase(subMenuList(j)) = "step" Then
							'��Ӽ����,step �Ǽ���ߵı�ʾ,������Ϊ����
							AppendMenu hPopMenu(n), MF_SEPARATOR, MenuCount, MenuText(MenuCount)
						Else
							AppendMenu hPopMenu(n), MF_STRING + MF_ENABLED, MenuCount, MenuText(MenuCount)
						End If
					End If
				Next j
				AppendMenu hMenu, MF_POPUP, hPopMenu(n), subMenuList(0)	'��Ӹ����˵���
				n = n + 1
			End If
		ElseIf strConfig(i) <> "" Then
			MenuCount = MenuCount + 1
			If LCase(strConfig(i)) = "step" Then
				'��Ӽ����,step �Ǽ���ߵı�ʾ,������Ϊ����
				AppendMenu hMenu, MF_SEPARATOR, MenuCount, strConfig(i)
			Else
				AppendMenu hMenu, MF_STRING + MF_ENABLED, MenuCount, strConfig(i)
			End If
			n = n + 1
		End If
	Next i
	If n = 0 Then Exit Function
	'��ʾ�˵�,���� 0 ��ʾ����
	'����ڲ��� uFlags ��ָ���� TPM_RETURNCMD ֵ���򷵻�ֵ���û�ѡ��Ĳ˵���ı�ʶ����
	i = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, pt.X, pt.Y, 0&, GetForegroundWindow(), 0)
	'������ĿΪ 0 ʱ����Ϊ -1
	If i < 1 Then
    	'�ͷŲ˵���Դ
	    DestroyMenu hMenu
	    Exit Function
	End If
	If Mode Then
		PopupMenuShow = i - 1 '���ز˵���
		'�ͷŲ˵���Դ
		DestroyMenu hMenu
	Else
		'��ȡѡ�еĲ˵����ִ�
		Dim buffer As String
		buffer = Space(255)
		'MF_BYCOMMAND: ��ʾ���� uIDltem �����˵���ı�ʶ��, ȱʡֵ
		'MF_BYPOSITION: ��ʾ���� uIDltem�����˵�����������λ��
		i = GetMenuString(hMenu, i, buffer, Len(buffer), MF_BYCOMMAND)
		'�ͷŲ˵���Դ
		DestroyMenu hMenu
		If i = 0 Then Exit Function
		'������ѡ�˵�����ı�
		PopupMenuShow = Replace$(buffer,vbNullChar,"")
	End If
End Function


'����Զ����ִ����͵������Ƿ�Ϊ�գ��ǿշ��� True
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
