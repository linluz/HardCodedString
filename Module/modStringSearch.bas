Attribute VB_Name = "modStringSearch"
'' String Search Module for PSlHardCodedString.bas
'' (c) 2015-2019 by wanfu (Last modified on 2019.10.28)

'#Uses "modCommon.bas"
'#Uses "modPEInfo.bas"
'#Uses "modMacInfo.bas"

Option Explicit

'�����ִ�����
Private Type FindStrList
	StartPos		As Long
	EndPos			As Long
	inSecID			As Integer
	Matches 		As Object
End Type

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

Public SearchSet() As String,SearchResult() As String
Private StopHwnd As Long,EscHwnd As Long,StopMsg As String,StopTitle As String,FreeByteList() As FREE_BTYE_SPACE


'����Ŀ¼�������ļ��е��ַ���
Public Function StringSearch(ByVal Mode As Long) As String
	Dim i As Integer,TempList() As String,MsgList() As String
	If getMsgList(UIDataList,MsgList,"StringSearch",1) = False Then Exit Function
	Begin Dialog UserDialog 780,476,MsgList(0),.StringSearchDlgFunc ' %GRID:10,7,1,1
		TextBox 0,0,0,21,.SuppValueBox
		PushButton 10,28,70,14,"",.HideButton
		CheckBox 10,49,70,14,"",.ModeBox
		Text 10,9,90,14,MsgList(1),.FindText
		TextBox 100,7,520,42,.FindTextBox,1
		PushButton 620,7,30,21,MsgList(2),.FindStrButton
		PushButton 620,28,30,21,MsgList(2),.RegExpTipButton
		Text 100,56,80,14,MsgList(3),.LangText
		DropListBox 180,53,230,21,TempList(),.LangNameList
		DropListBox 180,53,230,21,TempList(),.LangValueList
		Text 430,56,80,14,MsgList(4),.EncodeText
		DropListBox 510,53,110,21,TempList(),.EncodeList
		Text 100,77,80,14,MsgList(5),.EndCharText
		DropListBox 180,74,410,21,TempList(),.EndCharList
		PushButton 590,74,30,21,MsgList(2),.EndCharButton
		Text 10,101,90,14,MsgList(6),.DirectoryText
		TextBox 100,98,520,21,.DirectoryPathBox
		PushButton 620,98,30,21,MsgList(7),.BrowseButton
		GroupBox 10,126,640,42,MsgList(8),.FindModeGroup
		CheckBox 30,140,140,21,MsgList(9),.MatchCaseBox
		CheckBox 180,140,140,21,MsgList(10),.MatchFullWordBox
		CheckBox 330,140,140,21,MsgList(11),.MatchFullTextBox
		CheckBox 480,140,150,21,MsgList(12),.IgnoreAcckeyBox
		GroupBox 10,175,640,84,MsgList(13),.FindRangeGroup
		Text 30,192,90,14,MsgList(14),.FileTypeText
		TextBox 130,189,190,21,.FileTypeBox
		Text 340,192,90,14,MsgList(15),.FileNameText
		TextBox 440,189,190,21,.FileNameBox
		Text 30,215,90,14,MsgList(16),.IgnoreSubFolderText
		DropListBox 130,213,470,21,TempList(),.IgnoreSubFolderList,1
		PushButton 600,213,30,21,MsgList(7),.IgnoreSubFolderButton
		CheckBox 30,234,140,21,MsgList(17),.SubFolderBox
		CheckBox 180,234,140,21,MsgList(18),.IgnoreHideFolderBox
		CheckBox 330,234,140,21,MsgList(19),.IgnoreHideFileBox
		CheckBox 480,234,150,21,MsgList(20),.SkipHeadersCheckBox
		Text 10,266,640,14,MsgList(21),.FileResultText
		ListBox 10,287,760,182,TempList(),.FileResultList,1
		PushButton 660,7,110,28,MsgList(22),.FindButton
		PushButton 660,49,110,21,MsgList(23),.CopyResultButton
		PushButton 660,70,110,21,MsgList(24),.CleanButton
		PushButton 660,105,110,21,MsgList(25),.ViewFileInfoButton
		PushButton 660,126,110,21,MsgList(26),.ViewStrInfoButton
		PushButton 660,147,110,21,MsgList(27),.OpenFileButton
		PushButton 660,168,110,21,MsgList(28),.GetStringButton
		PushButton 660,224,110,21,MsgList(29),.HelpButton
		PushButton 660,203,110,21,MsgList(73),.ExitButton
		CancelButton 660,203,110,21,.CancelButton
		PushButton 660,259,110,21,MsgList(74),.StopButton
	End Dialog
	Dim dlg As UserDialog
	dlg.ModeBox = Mode
	If Dialog(dlg) = 0 Then Exit Function
	StringSearch = SearchSet(17)
End Function


'������Ի�����
Private Function StringSearchDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,j As Long,k As Long,n As Long,m As Long,Max As Long
	Dim TempList() As String,TempArray() As String,MsgList() As String
	Dim Temp As String,File As FILE_PROPERTIE,FN As FILE_IMAGE
	Dim StartTime As Date,RegEx As Object
	Select Case Action%
	Case 1 ' �Ի��򴰿ڳ�ʼ��
		If getMsgList(UIDataList,MsgList,"StringSearch",1) = False Then Exit Function
		DlgText "SuppValueBox",CStr$(SuppValue)
		DlgVisible "SuppValueBox",False
		DlgVisible "HideButton",False
		DlgVisible "CancelButton",False
		DlgVisible "StopButton",False
		DlgVisible "ModeBox",False
		DlgVisible "LangValueList",False
		If SearchSet(16) = "" Then
			'��ʼ��ÿ���ļ��Ĳ�������
			ReDim SearchResult(0) As String
			DlgEnable "CopyResultButton",False
			DlgEnable "ViewFileInfoButton",False
			DlgEnable "ViewStrInfoButton",False
			DlgEnable "OpenFileButton",False
			DlgEnable "GetStringButton",False
			DlgEnable "CleanButton",False
		End If
		If DlgValue("ModeBox") = 1 Then
			DlgEnable "GetStringButton",False
		End If
		EscHwnd = GetDlgItem(SuppValue,DlgControlId("CancelButton"))
		StopHwnd = GetDlgItem(SuppValue,DlgControlId("StopButton"))
		StopMsg = MsgList(75)
		StopTitle = MsgList(33)

		'�ָ����ҳ��������ƺ�ֵ��������ʾ
		TempList = GetLangStrList(UniLangList,0)
		DlgListBoxArray "LangNameList",TempList()
		TempList = GetLangStrList(UniLangList,1)
		DlgListBoxArray "LangValueList",TempList()
		TempList = ReSplit(MsgList(30),ItemJoinStr)
		DlgListBoxArray "EncodeList",TempList()

		'��ʼ������
		If SearchSet(1) = "" Then SearchSet(1) = CStr$(UseLangList(0).LangID)
		If SearchSet(2) = "" Then SearchSet(2) = "0"
		If SearchSet(3) = "" Then SearchSet(3) = ExtractSet(31)
		If SearchSet(4) = "" Then SearchSet(4) = PSL.ActiveProject.Location
		If SearchSet(5) = "" Then SearchSet(5) = "0"
		If SearchSet(6) = "" Then SearchSet(6) = "0"
		If SearchSet(7) = "" Then SearchSet(7) = "0"
		If SearchSet(8) = "" Then SearchSet(8) = "0"
		If SearchSet(12) = "" Then SearchSet(12) = "0"
		If SearchSet(13) = "" Then SearchSet(13) = "0"
		If SearchSet(14) = "" Then SearchSet(14) = "0"
		If SearchSet(15) = "" Then SearchSet(15) = "1"

		DlgText "FindTextBox",SearchSet(0)
		DlgText "LangValueList",SearchSet(1)
		DlgValue "EncodeList",StrToLong(SearchSet(2))
		TempList = ReSplit(ReSplit(SearchSet(3),ItemJoinStr,2)(1),JoinStr)
		DlgListBoxArray "EndCharList",TempList()
		DlgValue "EndCharList",StrToLong(ReSplit(SearchSet(3),ItemJoinStr,2)(0))
		DlgText "DirectoryPathBox",SearchSet(4)
		DlgValue "MatchCaseBox",StrToLong(SearchSet(5))
		DlgValue "MatchFullWordBox",StrToLong(SearchSet(6))
		DlgValue "MatchFullTextBox",StrToLong(SearchSet(7))
		DlgValue "IgnoreAcckeyBox",StrToLong(SearchSet(8))
		DlgText "FileTypeBox",SearchSet(9)
		DlgText "FileNameBox",SearchSet(10)
		TempList = ReSplit(SearchSet(11),TextJoinStr)
		DlgListBoxArray "IgnoreSubFolderList",TempList()
		DlgValue "IgnoreSubFolderList",0
		If DlgListBoxArray("IgnoreSubFolderList") > 0 Then
			DlgText "IgnoreSubFolderButton",MsgList(2)
		End If
		DlgValue "SubFolderBox",StrToLong(SearchSet(12))
		DlgValue "IgnoreHideFolderBox",StrToLong(SearchSet(13))
		DlgValue "IgnoreHideFileBox",StrToLong(SearchSet(14))
		DlgValue "SkipHeadersCheckBox",StrToLong(SearchSet(15))
		TempList = ReSplit(SearchSet(16),TextJoinStr)
		DlgListBoxArray "FileResultList",TempList()
		DlgValue "FileResultList",0
		Temp = DlgText("FileResultText")
		If CheckArray(TempList) = False Then
			DlgText "FileResultText",Replace$(Replace$(Temp,"%s","0"),"%d","0")
		Else
			DlgText "FileResultText",Replace$(Replace$(Temp,"%s",CStr$(UBound(TempList) + 1)),"%d","1")
		End If

		If DlgValue("SubFolderBox") = 0 Then
			DlgEnable "IgnoreSubFolderList",False
			DlgEnable "IgnoreSubFolderButton",False
		End If

		If DlgValue("LangValueList") < 0 Then DlgText "LangValueList","1033"
		DlgValue "LangNameList",DlgValue("LangValueList")
		SearchSet(1) = DlgText("LangValueList")

		'���õ�ǰ�Ի�������
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
			DrawWindow(SuppValue,j)
		End If
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
		StringSearchDlgFunc = True ' ��ֹ���°�ť�رնԻ��򴰿�
		Select Case DlgItem$
		Case "StopButton"
			Exit Function
		Case "CancelButton", "ExitButton"
			'ɾ����ʱ�ļ�
			If SearchSet(16) <> "" Then
				On Error Resume Next
				TempList = ReSplit(SearchSet(16),TextJoinStr)
				For i = 0 To UBound(TempList)
					Temp = ReSplit(TempList(i),vbTab)(0)
					If Dir$(Temp & ".xls") <> "" Then Kill Temp & ".xls"
				Next i
				On Error GoTo 0
			End If
			DlgFocus("CancelButton")  '���ý��㵽ȡ����ť���Ա�Ի��򷵻�ֵΪ��
			SendMessageLNG GetFocus(), BM_CLICK, 0, 0
			StringSearchDlgFunc = False
			Exit Function
		Case "HelpButton"
			If StrToLong(Selected(30)) = 1 Then
				If OpenCHM(CLng(DlgText("SuppValueBox")),1025,Selected(0),OSLanguage,UIFileList) = True Then Exit Function
			End If
			Call Help("StringSearchHelp")
		Case "FindStrButton"
			GetHistory(TempList,"FindStrings","StringSearchDlg")
			If CheckArray(TempList) = False Then Exit Function
			i = ShowPopupMenu(TempList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			DlgText "FindTextBox",TempList(i)
			SearchSet(0) = DlgText("FindTextBox")
		Case "RegExpTipButton"
			If getMsgList(UIDataList,MsgList,"RegExpRuleTip",1) = False Then Exit Function
			i = ShowPopupMenu(MsgList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			If i = UBound(MsgList) Then
				If StrToLong(Selected(30)) = 1 Then
					If OpenCHM(CLng(DlgText("SuppValueBox")),1022,Selected(0),OSLanguage,UIFileList) = True Then Exit Function
				End If
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
			SearchSet(0) = DlgText("FindTextBox")
		Case "LangNameList"
			DlgValue "LangValueList",DlgValue("LangNameList")
			SearchSet(1) = DlgText("LangValueList")
		Case "EncodeList"
			SearchSet(2) = DlgValue("EncodeList")
		Case "EndCharList"
			SearchSet(3) = DlgValue("EndCharList") & ItemJoinStr & ReSplit(SearchSet(3),ItemJoinStr,2)(1)
		Case "MatchCaseBox"
			SearchSet(5) = CStr$(DlgValue("MatchCaseBox"))
		Case "MatchFullWordBox"
			If DlgValue("MatchFullWordBox") = 1 Then DlgValue "MatchFullTextBox",0
			SearchSet(6) = CStr$(DlgValue("MatchFullWordBox"))
			SearchSet(7) = CStr$(DlgValue("MatchFullTextBox"))
		Case "MatchFullTextBox"
			If DlgValue("MatchFullTextBox") = 1 Then DlgValue "MatchFullWordBox",0
			SearchSet(6) = CStr$(DlgValue("MatchFullWordBox"))
			SearchSet(7) = CStr$(DlgValue("MatchFullTextBox"))
		Case "IgnoreAcckeyBox"
			SearchSet(8) = CStr$(DlgValue("IgnoreAcckeyBox"))
		Case "SubFolderBox"
			SearchSet(12) = CStr$(DlgValue("SubFolderBox"))
			If DlgValue("SubFolderBox") = 0 Then
				DlgEnable "IgnoreSubFolderList",False
				DlgEnable "IgnoreSubFolderButton",False
			Else
				DlgEnable "IgnoreSubFolderList",True
				DlgEnable "IgnoreSubFolderButton",True
			End If
		Case "IgnoreHideFolderBox"
			SearchSet(13) = CStr$(DlgValue("IgnoreHideFolderBox"))
		Case "IgnoreHideFileBox"
			SearchSet(14) = CStr$(DlgValue("IgnoreHideFileBox"))
		Case "SkipHeadersCheckBox"
			SearchSet(15) = CStr$(DlgValue("SkipHeadersCheckBox"))
		Case "BrowseButton"
			If getMsgList(UIDataList,MsgList,"StringSearch",1) = False Then Exit Function
			'If PSL.SelectFolder(Temp,MsgList(36)) = False Then Exit Function
			If BrowseForFolder(Temp,MsgList(36)) = False Then Exit Function
			If Temp = "" Then Exit Function
			If LCase$(Temp) = LCase$(SearchSet(4)) Then Exit Function
			DlgText "DirectoryPathBox",Temp
			ReDim TempList(0) As String
			DlgListBoxArray "IgnoreSubFolderList",TempList()
			DlgValue "IgnoreSubFolderList",0
			DlgText "IgnoreSubFolderButton",MsgList(7)
			SearchSet(4) = DlgText("DirectoryPathBox")
			SearchSet(11) = ""
		Case "EndCharButton"
			If getMsgList(UIDataList,MsgList,"StringSearch",1) = False Then Exit Function
			TempArray = ReSplit(MsgList(62),ItemJoinStr)
			ReDim Preserve TempArray(UBound(TempArray) - 1)
			i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			TempList = ReSplit(ReSplit(SearchSet(3),ItemJoinStr,2)(1),JoinStr)
			If i = 0 Then
				Do
					Temp = InputBox(MsgList(64),MsgList(63),Temp)
					If Temp = "" Then Exit Function
					If InStr(Temp,"(") = 0 Or InStr(Temp,")") = 0 Then
						MsgBox(MsgList(65),vbOkOnly+vbInformation,MsgList(34))
					ElseIf StrEndChar2Pattern(Temp,1)(0) = "" Then
						MsgBox(MsgList(65),vbOkOnly+vbInformation,MsgList(34))
					Else
						TempArray(0) = Mid$(Temp,InStr(Temp,"("))
						TempArray(0) = Left$(TempArray(0),InStrRev(TempArray(0),")"))
						If InStr(SearchSet(3),TempArray(0)) Then
							MsgBox(MsgList(66),vbOkOnly+vbInformation,MsgList(34))
						ElseIf CheckRegExp(RegExp,TempArray(0)) = False Then
							MsgBox(MsgList(67),vbOkOnly+vbInformation,MsgList(35))
						Else
							Exit Do
						End If
					End If
				Loop
				n = UBound(TempList) + 1
				ReDim Preserve TempList(n) As String
				TempList(n) = Temp
			ElseIf i = 1 Then
				n = DlgValue("EndCharList")
				If n < 0 Then Exit Function
				If n <= UBound(ReSplit(EndCharOfString,ValJoinStr)) Then
					MsgBox(MsgList(68),vbOkOnly+vbInformation,MsgList(34))
					Exit Function
				End If
				Do
					Temp = EditSet(TempList,n)
					If Temp = "" Then Exit Function
					If InStr(Temp,"(") = 0 Or InStr(Temp,")") = 0 Then
						MsgBox(MsgList(65),vbOkOnly+vbInformation,MsgList(34))
					ElseIf StrEndChar2Pattern(Temp,1)(0) = "" Then
						MsgBox(MsgList(65),vbOkOnly+vbInformation,MsgList(34))
					Else
						TempArray(0) = Mid$(Temp,InStr(Temp,"("))
						TempArray(0) = Left$(TempArray(0),InStrRev(TempArray(0),")"))
						If InStr(Replace$(SearchSet(3),TempList(n),""),TempArray(0)) Then
							MsgBox(MsgList(66),vbOkOnly+vbInformation,MsgList(34))
						ElseIf CheckRegExp(RegExp,TempArray(0)) = False Then
							MsgBox(MsgList(67),vbOkOnly+vbInformation,MsgList(35))
						Else
							Exit Do
						End If
					End If
				Loop
				TempList(n) = Temp
			ElseIf i = 2 Then
				n = DlgValue("EndCharList")
				If n < 0 Then Exit Function
				If n <= UBound(ReSplit(EndCharOfString,ValJoinStr)) Then
					MsgBox(MsgList(68),vbOkOnly+vbInformation,MsgList(34))
					Exit Function
				ElseIf MsgBox(MsgList(69),vbYesNo+vbInformation,MsgList(33)) = vbNo Then
					Exit Function
				End If
				Call DelArray(TempList,n)
				n = n - 1
			End If
			SearchSet(3) = CStr$(n) & ItemJoinStr & StrListJoin(TempList,JoinStr)
			SaveSetting(AppName,"GetString","EndCharOfString",ConvertStrEndCharSet(MergeStrEndCharSet(SearchSet(3)),True))
			DlgListBoxArray "EndCharList",TempList()
			DlgValue "EndCharList",n
		Case "IgnoreSubFolderButton"
			If getMsgList(UIDataList,MsgList,"StringSearch",1) = False Then Exit Function
			If DlgListBoxArray("IgnoreSubFolderList") > 0 Then
				TempList = ReSplit(MsgList(62),ItemJoinStr)
				i = ShowPopupMenu(TempList,vbPopupUseRightButton)
				If i < 0 Then Exit Function
			End If
			If i < 2 Then
				If i = 1 Then
					n = DlgValue("IgnoreSubFolderList")
					If n < 0 Then Exit Function
				End If
				If SearchSet(4) = "" Then
					MsgBox(MsgList(37),vbOkOnly+vbInformation,MsgList(35))
					Exit Function
				End If
				'If PSL.SelectFolder(Temp,MsgList(36)) = False Then Exit Function
				If BrowseForFolder(Temp,MsgList(36)) = False Then Exit Function
				If Temp = "" Then Exit Function
				If LCase$(Temp) = LCase$(SearchSet(4)) Then
					MsgBox(Replace$(MsgList(38),"%s",DlgText("DirectoryPathBox")),vbOkOnly+vbInformation,MsgList(35))
					Exit Function
				End If
				If InStr(LCase$(Temp),LCase$(SearchSet(4))) = 0 Then
					MsgBox(Replace$(MsgList(38),"%s",DlgText("DirectoryPathBox")),vbOkOnly+vbInformation,MsgList(35))
					Exit Function
				End If
				Temp = strReplace(Temp,SearchSet(4),"...")
				If InStr(TextJoinStr & LCase$(SearchSet(11)) & TextJoinStr,TextJoinStr & LCase$(Temp) & TextJoinStr) Then
					MsgBox(MsgList(39),vbOkOnly+vbInformation,MsgList(35))
					DlgText "IgnoreSubFolderList",Temp
					Exit Function
				End If
				TempList = ReSplit(SearchSet(11),TextJoinStr)
				If i = 0 Then
					n = DlgListBoxArray("IgnoreSubFolderList")
					ReDim Preserve TempList(n) As String
				End If
				TempList(n) = Temp
				DlgListBoxArray "IgnoreSubFolderList",TempList()
				DlgValue "IgnoreSubFolderList",n
				DlgText "IgnoreSubFolderButton",MsgList(2)
				SearchSet(11) = StrListJoin(TempList,TextJoinStr)
			ElseIf i = 2 Then
				n = DlgValue("IgnoreSubFolderList")
				If n < 0 Then Exit Function
				If MsgBox(MsgList(40),vbYesNo+vbInformation,MsgList(33)) = vbNo Then Exit Function
				TempList = ReSplit(SearchSet(11),TextJoinStr)
				Call DelArray(TempList,n)
				DlgListBoxArray "IgnoreSubFolderList",TempList()
				i = UBound(TempList)
				DlgValue "IgnoreSubFolderList",IIf(n > i,i,n)
				If DlgListBoxArray("IgnoreSubFolderList") > 0 Then
					DlgText "IgnoreSubFolderButton",MsgList(2)
				Else
					DlgText "IgnoreSubFolderButton",MsgList(7)
				End If
				SearchSet(11) = StrListJoin(TempList,TextJoinStr)
			Else
				If DlgValue("IgnoreSubFolderList") < 0 Then Exit Function
				If MsgBox(MsgList(41),vbYesNo+vbInformation,MsgList(33)) = vbNo Then Exit Function
				ReDim TempList(0) As String
				DlgListBoxArray "IgnoreSubFolderList",TempList()
				DlgValue "IgnoreSubFolderList",0
				DlgText "IgnoreSubFolderButton",MsgList(7)
				SearchSet(11) = ""
			End If
		Case "FindButton"
			If getMsgList(UIDataList,MsgList,"StringSearch",1) = False Then Exit Function
			'ɾ����ʱ�ļ�
			If SearchSet(16) <> "" Then
				On Error Resume Next
				TempList = ReSplit(SearchSet(16),TextJoinStr)
				For i = 0 To UBound(TempList)
					Temp = ReSplit(TempList(i),vbTab)(0)
					If Dir$(Temp & ".xls") <> "" Then Kill Temp & ".xls"
				Next i
				On Error GoTo 0
			End If
			'������Ŀ¼�Ƿ�Ϊ��
			If SearchSet(4) = "" Then
				MsgBox(MsgList(37),vbOkOnly+vbInformation,MsgList(35))
				Exit Function
			End If
			'�����������Ƿ����
			If SearchSet(0) <> "" Then
				Select Case FilterStr("CheckRegExp",SearchSet(0),GetFindMode(SearchSet(0)))
				Case -2
					MsgBox MsgList(77),vbOkOnly+vbInformation,MsgList(35)
					Exit Function
				Case -3
					MsgBox MsgList(52),vbOkOnly+vbInformation,MsgList(35)
					Exit Function
				End Select
			End If
			'��ʼ���ؼ���ʾ
			SearchSet(16) = ""
			ReDim SearchResult(0) As String
			DlgText "FileResultText",Replace$(Replace$(MsgList(21),"%s","0"),"%d","0")
			ReDim TempList(0) As String
			DlgListBoxArray "FileResultList",TempList()
			DlgValue "FileResultList",0
			DlgEnable "CopyResultButton",False
			DlgEnable "ViewFileInfoButton",False
			DlgEnable "ViewStrInfoButton",False
			DlgEnable "OpenFileButton",False
			DlgEnable "GetStringButton",False
			DlgEnable "CleanButton",False
			'��ȡ�ļ��б�
			TempList = ReSplit(SearchSet(11),TextJoinStr)
			For i = 0 To UBound(TempList)
				TempList(i) = Replace$(TempList(i),"...",SearchSet(4),,1)
			Next i
			Temp = StrListJoin(TempList,";")
			ReDim TempList(5) As String
			TempList(0) = SearchSet(9)	'�����ļ�����
			TempList(1) = SearchSet(10)	'Ҫ���Ե��ļ���
			TempList(2) = SearchSet(12)	'�������ļ���
			TempList(3) = SearchSet(13)	'�����������ļ���
			TempList(4) = SearchSet(14)	'���������ļ�
			TempList(5) = Temp			'�������ļ����б�
			j = 0
			If StrToLong(Selected(17)) > 0 Then
				j = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("FileResultText"))
			ElseIf StrToLong(Selected(17)) < 0 Then
				j = -2
			End If
			If GetFiles(AppendBackslash(SearchSet(4),"","\",1),TempList,TempArray,MsgList(44),j) < 0 Then Exit Function
			If CheckArray(TempArray) = False Then
				DlgText "FileResultText",Replace$(MsgList(44),"%s","0")
				If SearchSet(12) = "0" Then
					MsgBox(MsgList(42),vbOkOnly+vbInformation,MsgList(34))
					Exit Function
				Else
					MsgBox(MsgList(43),vbOkOnly+vbInformation,MsgList(34))
					Exit Function
				End If
			End If
			SearchSet(0) = DlgText("FindTextBox")
			'��������Ϊ��ʱ��ʾ�ļ��б��˳�
			If SearchSet(0) = "" Then
				DlgListBoxArray "FileResultList",TempArray()
				DlgValue "FileResultList",UBound(TempArray)
				DlgText "FileResultText",Replace$(MsgList(44),"%s",CStr$(UBound(TempArray) + 1))
				DlgEnable "CopyResultButton",True
				DlgEnable "ViewFileInfoButton",True
				DlgEnable "ViewStrInfoButton",True
				DlgEnable "OpenFileButton",True
				If DlgValue("ModeBox") = 0 Then DlgEnable "GetStringButton",True
				DlgEnable "CleanButton",True
				SearchSet(16) = StrListJoin(TempArray,TextJoinStr)
				'��ʼ��ÿ���ļ��Ĳ�������
				ReDim SearchResult((UBound(TempArray) + 1) * 2 - 1) As String
				Exit Function
			Else
				'���������ݵĲ��ҷ�ʽ��ת��Ϊ������ʽģ��
				Temp = SearchSet(0)
				If SearchSet(8) = "1" Then Temp = DelAccKey(Temp)
				Temp = StrToRegExpPattern(Temp)
				'�����������Ƿ����������ʽҪ��
				If CheckRegExp(RegExp,Temp) = False Then
					MsgBox MsgList(52),vbOkOnly+vbInformation,MsgList(35)
					Exit Function
				End If
			End If
			'��Ӳ�������
			GetHistory(TempList,"FindStrings","StringSearchDlg")
			If InsertArray(TempList,DlgText("FindTextBox"),0,True) = True Then
				WriteHistory(TempList,"FindStrings","StringSearchDlg")
			End If
			'ת�����Ҳ���
			ReDim lngList(4) As Long
			Select Case SearchSet(2)
			Case "0"
				lngList(0) = UniLangList(GetDicVal(LangIDIndexDic,StrToLong(SearchSet(1)),0)).CodePage	'����ҳ
				If lngList(0) = CP_UNKNOWN Then
					MsgBox(MsgList(71),vbOkOnly+vbInformation,MsgList(35))
					Exit Function
				End If
				lngList(3) = 1							'����ҳ�ֽڳ���
			Case "1"
				lngList(0) = CP_UNICODELITTLE
				lngList(3) = 2
			Case "2"
				lngList(0) = CP_UTF8
				lngList(3) = 1
			Case "3"
				lngList(0) = CP_UNICODEBIG
				lngList(3) = 2
			Case "4"
				lngList(0) = CP_UTF7
				lngList(3) = 1
			Case "5"
				lngList(0) = CP_UTF32LE
				lngList(3) = 4
			Case "6"
				lngList(0) = CP_UTF32BE
				lngList(3) = 4
			End Select
			lngList(1) = StrToLong(SearchSet(8))		'���Կ�ݼ�
			lngList(2) = StrToLong(SearchSet(15))		'���� PE ����Ŀ¼
			lngList(4) = 1	'Len(Convert(SearchSet(0)))	'���������ַ�����
			'����������ʽģ��
			Set RegEx = CreateObject("VBScript.RegExp")
			If SearchSet(6) = "1" Then
				'���ȫ��ƥ��ʱ�ǲ���ȫӢ�Ļ����
				If CheckStrRegExp(Temp,"[\x01-\xFE]",0,1) = True Then
					RegEx.Pattern = AppendBackslash(Temp,"\b","\b",1)
				Else
					RegEx.Pattern = Temp
				End If
			ElseIf SearchSet(7) = "1" Then
				TempList = StrEndChar2Pattern(SearchSet(3),2)
				RegEx.Pattern = TempList(0) & Temp & TempList(0)
			Else
				RegEx.Pattern = Temp
			End If
			RegEx.Global = True
			RegEx.IgnoreCase = IIf(SearchSet(5) = "0",True,False)
			'��ʼ����
			ReDim MatcheList(0) As FindStrList
			Max = UBound(TempArray) + 1
			ReDim TempList(Max) As String
			DlgEnable "FindButton",False
			StartTime = Timer
			'��ʾ�����ڵ�ȡ��������ť�ͽ�ֹ�����ڵ� Esc ����Ӧ�˳�������
			Call ShowButton(StopHwnd,VK_ESCAPE,True)
			For i = 0 To Max - 1
				Temp = Replace$(MsgList(45),"%s",TempArray(i))
				If Len(Temp) > 60 Then
					Temp = Left$(Temp,InStr(Temp,"\")) & "..." & Right(Temp,60 - Len(Left$(Temp,InStr(Temp,"\"))))
				End If
				DlgText "FileResultText",Temp & Format$((i + 1) / Max,"#%")
			   	If Dir$(TempArray(i)) <> "" Then
					File.FilePath = TempArray(i)
					j = FindStringCount(MatcheList,RegEx,File,FN,lngList,0)
					UnLoadFile(FN,0,0)
					If j > 0 Then
						TempList(n) = TempArray(i) & vbTab & Replace$(MsgList(46),"%s",CStr$(j))
						DlgListBoxArray "FileResultList",TempList()
						DlgValue "FileResultList",n
						n = n + 1: m = m + j
					ElseIf j < 0 Then
						Exit For
					End If
				End If
			Next i
			Set RegEx = Nothing
			'�����������е�ȡ��������ť������ Esc �����˳���Ӧ
			Call ShowButton(StopHwnd,VK_ESCAPE,False)
			'��ʾ���
			DlgText "FileResultText",Replace$(MsgList(45),"%s",TempArray(i - 1)) & "100%"
			If n = 0 Then
				ReDim TempList(0) As String
				TempList(0) = MsgList(47)
				DlgListBoxArray "FileResultList",TempList()
				DlgValue "FileResultList",0
				DlgEnable "CopyResultButton",False
				DlgEnable "ViewFileInfoButton",False
				DlgEnable "ViewStrInfoButton",False
				DlgEnable "OpenFileButton",False
				DlgEnable "GetStringButton",False
				SearchSet(16) = ""
				'��ʼ��ÿ���ļ��Ĳ�������
				ReDim SearchResult(0) As String
			Else
				ReDim Preserve TempList(n - 1) As String
				DlgListBoxArray "FileResultList",TempList()
				DlgValue "FileResultList",n - 1
				DlgEnable "CopyResultButton",True
				DlgEnable "ViewFileInfoButton",True
				DlgEnable "ViewStrInfoButton",True
				DlgEnable "OpenFileButton",True
				If DlgValue("ModeBox") = 0 Then DlgEnable "GetStringButton",True
				SearchSet(16) = StrListJoin(TempList,TextJoinStr)
				'��ʼ��ÿ���ļ��Ĳ�������
				ReDim SearchResult((UBound(TempList) + 1) * 2 - 1) As String
			End If
			DlgEnable "FindButton",True
			DlgEnable "CleanButton",True
			DlgText "FileResultText",Replace$(Replace$(Replace$(Replace$(MsgList(48),"%s",CStr$(Max)),"%d",CStr$(n)), _
								"%n",CStr$(m)),"%t",Format$(DateAdd("s",Timer - StartTime,0),MsgList(49)))
		Case "CopyResultButton"
			Clipboard SearchSet(16)
		Case "ViewFileInfoButton", "OpenFileButton", "GetStringButton"
			i = DlgValue("FileResultList")
			If i < 0 Then Exit Function
			If getMsgList(UIDataList,MsgList,"StringSearch",1) = False Then Exit Function
			'ѡ���б������ļ�·��Ϊ��ʱ�˳�
			SearchSet(17) = ReSplit(ReSplit(SearchSet(16),TextJoinStr)(i),vbTab)(0)
			If SearchSet(17) = "" Then
				MsgBox(MsgList(50),vbOkOnly+vbInformation,MsgList(34))
				Exit Function
			End If
			Select Case DlgItem$
			Case "OpenFileButton"
				ReDim TempList(UBound(Tools)) As String
				For i = 3 To UBound(Tools)
					TempList(i - 3) = Tools(i).sName
				Next i
				i = ShowPopupMenu(TempList,vbPopupUseRightButton)
				If i < 0 Then Exit Function
				ReDim TempList(0) As String
				TempList(0) = SearchSet(17) & JoinStr
				If OpenFile(SearchSet(17),TempList,i + 3,False) = True Then
					If i = 0 Then WriteSettings("Tools")
				End If
			Case "GetStringButton"
				MsgBox(MsgList(51),vbOkOnly+vbInformation,MsgList(34))
				StringSearchDlgFunc = False
			Case "ViewFileInfoButton"
				ReDim TempList(UBound(Tools)) As String
				For i = 0 To UBound(Tools)
					TempList(i) = Tools(i).sName
				Next i
				i = ShowPopupMenu(TempList,vbPopupUseRightButton)
				If i < 0 Then Exit Function
				File.FilePath = SearchSet(17)
				If GetFileInfo(File.FilePath,File) = False Then Exit Function
				GetHeaders(File.FilePath,File,StrToLong(Selected(1)),File.FileType)
				Call FileInfoView(File,FreeByteList,i,0,StrToLong(Selected(16)))
			End Select
		Case "ViewStrInfoButton"
			k = DlgValue("FileResultList")
			If k < 0 Then Exit Function
			If getMsgList(UIDataList,MsgList,"StringSearch",1) = False Then Exit Function
			'��������Ϊ��ʱ�˳�
			If SearchSet(0) = "" Then
				MsgBox(MsgList(32),vbOkOnly+vbInformation,MsgList(35))
				Exit Function
			'�����������Ƿ����
			Else
				Select Case FilterStr("CheckRegExp",SearchSet(0),GetFindMode(SearchSet(0)))
				Case -2
					MsgBox MsgList(77),vbOkOnly+vbInformation,MsgList(35)
					Exit Function
				Case -3
					MsgBox MsgList(52),vbOkOnly+vbInformation,MsgList(35)
					Exit Function
				End Select
			End If
			'ѡ���б������ļ�·��Ϊ��ʱ�˳�
			SearchSet(17) = ReSplit(ReSplit(SearchSet(16),TextJoinStr)(k),vbTab)(0)
			If SearchSet(17) = "" Then
				MsgBox(MsgList(50),vbOkOnly+vbInformation,MsgList(34))
				Exit Function
			End If
			'ѡ����ִ���Ϣ�Ĺ���
			ReDim TempList(UBound(Tools)) As String
			For i = 0 To UBound(Tools)
				TempList(i) = Tools(i).sName
			Next i
			i = ShowPopupMenu(TempList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			'��������Ƿ񱻸��ģ�δ����ʱ����ȡ�����ִ�����
			If SearchResult(k) <> "" Then
				TempList = ReSplit(SearchResult(k),vbNullChar)
				If ArrayComp(TempList,SearchSet,"0-8,15") = False Then
					j = DlgListBoxArray("FileResultList")
					If SearchResult(j + k) <> "" Then
						TempList = ReSplit(SearchResult(j + k),TextJoinStr)
						n = UBound(TempList) - 9
						DlgText "FileResultText",Replace$(Replace$(MsgList(54),"%n",CStr$(n)),"%t",Format$(DateAdd("s",0,0),MsgList(49)))
						'������Ϊ��ʱ�ļ�
						If WriteToFile(SearchSet(17) & ".xls",SearchResult(j + k),"unicodeFFFE") = False Then
							Exit Function
						End If
						'����ʱ�ļ��鿴��Ϣ
						ReDim FileDataList(0) As String
						FileDataList(0) = SearchSet(17) & ".xls" & JoinStr & "unicodeFFFE"
						If OpenFile(SearchSet(17) & ".xls",FileDataList,i,False) = True Then
							If i = 3 Then WriteSettings("Tools")
						End If
					Else
						MsgBox MsgList(47),vbOkOnly+vbInformation,MsgList(34)
					End If
					Exit Function
				End If
			End If
			SearchResult(k) = StrListJoin(SearchSet,vbNullChar)
			'���������ݵĲ��ҷ�ʽ��ת��Ϊ������ʽģ��
			Temp = SearchSet(0)
			If SearchSet(8) = "1" Then Temp = DelAccKey(Temp)
			Temp = StrToRegExpPattern(Temp)
			'�����������Ƿ����������ʽҪ��
			If CheckRegExp(RegExp,Temp) = False Then
				MsgBox MsgList(52),vbOkOnly+vbInformation,MsgList(35)
				Exit Function
			End If
			'��Ӳ�������
			GetHistory(TempList,"FindStrings","StringSearchDlg")
			If InsertArray(TempList,DlgText("FindTextBox"),0,True) = True Then
				WriteHistory(TempList,"FindStrings","StringSearchDlg")
			End If
			'ת�����Ҳ���
			ReDim lngList(4) As Long
			Select Case SearchSet(2)
			Case "0"
				lngList(0) = UniLangList(GetDicVal(LangIDIndexDic,StrToLong(SearchSet(1)),0)).CodePage	'����ҳ
				If lngList(0) = CP_UNKNOWN Then
					MsgBox(MsgList(71),vbOkOnly+vbInformation,MsgList(35))
					Exit Function
				End If
				lngList(3) = 1							'����ҳ�ֽڳ���
			Case "1"
				lngList(0) = CP_UNICODELITTLE
				lngList(3) = 2
			Case "2"
				lngList(0) = CP_UTF8
				lngList(3) = 1
			Case "3"
				lngList(0) = CP_UNICODEBIG
				lngList(3) = 2
			Case "4"
				lngList(0) = CP_UTF7
				lngList(3) = 1
			Case "5"
				lngList(0) = CP_UTF32LE
				lngList(3) = 4
			Case "6"
				lngList(0) = CP_UTF32BE
				lngList(3) = 4
			End Select
			lngList(1) = StrToLong(SearchSet(8))		'���Կ�ݼ�
			lngList(2) = StrToLong(SearchSet(15))		'���� PE ����Ŀ¼
			lngList(4) = 1 'Len(Convert(SearchSet(0))) 	'���������ַ�����
			'����������ʽģ��
			Set RegEx = CreateObject("VBScript.RegExp")
			With UniLangList(GetDicVal(LangIDIndexDic,StrToLong(SearchSet(1)),0))
				If SearchSet(7) = "1" Then
					TempList = StrEndChar2Pattern(SearchSet(3),2)
					RegEx.Pattern = TempList(0) & "(" & Temp & ")()" & TempList(0)
				ElseIf SearchSet(6) = "1" Then
					If .UniCodeRange = "" Then
						MsgBox(MsgList(72),vbOkOnly+vbInformation,MsgList(35))
						Exit Function
					End If
					TempList = StrEndChar2Pattern(SearchSet(3),3)
					'���ȫ��ƥ��ʱ�ǲ���ȫӢ�Ļ����
					If CheckStrRegExp(Temp,"[\x01-\xFE]",0,1) = True Then
						RegEx.Pattern = TempList(0) & .UniCodeRegExpPattern & "*?" & _
									AppendBackslash(Temp,"\b","\b",1) & .UniCodeRegExpPattern & "*?)(" & TempList(1)
					Else
						RegEx.Pattern = TempList(0) & .UniCodeRegExpPattern & "*?" & Temp & _
									.UniCodeRegExpPattern & "*?)(" & TempList(1)
					End If
				Else
					If .UniCodeRange = "" Then
						MsgBox(MsgList(72),vbOkOnly+vbInformation,MsgList(35))
						Exit Function
					End If
					TempList = StrEndChar2Pattern(SearchSet(3),3)
					RegEx.Pattern = TempList(0) & .UniCodeRegExpPattern & "*?" & Temp & _
								.UniCodeRegExpPattern & "*?)(" & TempList(1)
				End If
			End With
			RegEx.Global = True
			RegEx.IgnoreCase = IIf(SearchSet(5) = "0",True,False)
			'��ʼ����
			Dim strData As STRING_SUB_PROPERTIE
			ReDim MatcheList(0) As FindStrList
			Temp = Replace$(MsgList(53),"%s",SearchSet(17))
			If Len(Temp) > 60 Then
				Temp = Left$(Temp,InStr(Temp,"\")) & "..." & Right(Temp,60 - Len(Left$(Temp,InStr(Temp,"\"))))
			End If
			DlgText "FileResultText",Temp
			n = 7
			StartTime = Timer
			File.FilePath = SearchSet(17)
			j = 0
			If StrToLong(Selected(17)) > 0 Then
				j = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("FileResultText"))
			ElseIf StrToLong(Selected(17)) < 0 Then
				j = -2
			End If
			'��ʾ�����ڵ�ȡ��������ť�ͽ�ֹ�����ڵ� Esc ����Ӧ�˳�������
			Call ShowButton(StopHwnd,VK_ESCAPE,True)
			Max = FindStringCount(MatcheList,RegEx,File,FN,lngList,0,j)
			If Max < 1 Then
				'�����������е�ȡ��������ť������ Esc �����˳���Ӧ
				Call ShowButton(StopHwnd,VK_ESCAPE,False)
				MsgBox MsgList(47),vbOkOnly+vbInformation,MsgList(34)
				Exit Function
			End If
			Set RegEx = Nothing
			'�����ظ���������
			ChangSectionNames File,MsgList(76),MsgList(31)
			'��ȡ��������
			TempArray = getSectionList(File.SecList,2)
			ReDim TempList(Max + n + 1) As String
			TempList(0) = MsgList(55)
			TempList(1) = Replace$(MsgList(56),"%s",SearchSet(17))
			TempList(2) = Replace$(MsgList(57),"%s",SearchSet(0))
			TempList(3) = Replace$(Replace$(MsgList(58),"%s",UniLangList(LangIDIndexDic.Item(StrToLong(SearchSet(1)))).LangName),"%d",CStr$(lngList(0)))
			TempList(4) = MsgList(60) & MsgList(60)
			TempList(5) = MsgList(59)
			TempList(6) = MsgList(60) & MsgList(60)
			With strData
				.CodePage = lngList(0)
				For j = 0 To UBound(MatcheList)
					If MatcheList(j).Matches.Count > 0 Then
						.lMaxAddress = MatcheList(j).StartPos
						.inSectionID = MatcheList(j).inSecID
						For k = 0 To MatcheList(j).Matches.Count - 1
							.lHexLength = -1
							If MatcheList(j).Matches(k).Length >= lngList(4) Then
								.lStartAddress = MatcheList(j).StartPos
								.lEndAddress = MatcheList(j).EndPos
								Call GetStrAddress(FN,strData,MatcheList(j).Matches(k),lngList(3))
								.sString = MatcheList(j).Matches(k).SubMatches(1)
							End If
							If .lHexLength >= lngList(3) * lngList(4) Then
								TempList(n) = "#" & CStr$(n - 6) & vbTab & _
										ValToStr(.lStartAddress,File.FileSize,StrToLong(Selected(16))) & vbTab & _
										TempArray(.inSectionID) & vbTab & ReConvert(.sString)
								n = n + 1
							End If
							m = m + 1
							DlgText "FileResultText",Temp & Format$(m / Max,"#%")
						Next k
					End If
				Next j
				DlgText "FileResultText",Temp & "100%"
			End With
			'�����������е�ȡ��������ť������ Esc �����˳���Ӧ
			Call ShowButton(StopHwnd,VK_ESCAPE,False)
			UnLoadFile(FN,0,0)
			TempList(n) = MsgList(60) & MsgList(60)
			TempList(n + 1) = Replace$(Replace$(MsgList(61),"%n",CStr$(n - 7)),"%t",Format$(DateAdd("s",Timer - StartTime,0),MsgList(49)))
			ReDim Preserve TempList(n + 1) As String
			DlgText "FileResultText",Replace$(Replace$(MsgList(54),"%n",CStr$(n - 7)),"%t",Format$(DateAdd("s",Timer - StartTime,0),MsgList(49)))
			j = DlgListBoxArray("FileResultList")
			k = DlgValue("FileResultList")
			SearchResult(j + k) = StrListJoin(TempList,TextJoinStr)
			If n = 7 Then
				MsgBox MsgList(47),vbOkOnly+vbInformation,MsgList(34)
				Exit Function
			End If
			'������Ϊ��ʱ�ļ�
			If WriteToFile(SearchSet(17) & ".xls",SearchResult(j + k),"unicodeFFFE") = False Then
				Exit Function
			End If
			'����ʱ�ļ��鿴��Ϣ
			ReDim FileDataList(0) As String
			FileDataList(0) = SearchSet(17) & ".xls" & JoinStr & "unicodeFFFE"
			If OpenFile(SearchSet(17) & ".xls",FileDataList,i,False) = True Then
				If i = 3 Then WriteSettings("Tools")
			End If
		Case "CleanButton"
			If getMsgList(UIDataList,MsgList,"StringSearch",1) = False Then Exit Function
			'ɾ����ʱ�ļ�
			If SearchSet(16) <> "" Then
				On Error Resume Next
				TempList = ReSplit(SearchSet(16),TextJoinStr)
				For i = 0 To UBound(TempList)
					Temp = ReSplit(TempList(i),vbTab)(0)
					If Dir$(Temp & ".xls") <> "" Then Kill Temp & ".xls"
				Next i
				On Error GoTo 0
			End If
			DlgText "FileResultText",Replace$(Replace$(MsgList(21),"%s","0"),"%d","0")
			ReDim TempList(0) As String
			DlgListBoxArray "FileResultList",TempList()
			DlgValue "FileResultList",0
			DlgEnable "CopyResultButton",False
			DlgEnable "ViewFileInfoButton",False
			DlgEnable "ViewStrInfoButton",False
			DlgEnable "OpenFileButton",False
			DlgEnable "GetStringButton",False
			DlgEnable "CleanButton",False
			SearchSet(16) = ""
			'��ʼ��ÿ���ļ��Ĳ�������
			ReDim SearchResult(0) As String
		Case "FileResultList"
			If getMsgList(UIDataList,MsgList,"StringSearch",1) = False Then Exit Function
			i = DlgValue("FileResultList")
			If i < 0 Then
				DlgText "FileResultText",Replace$(Replace$(MsgList(21),"%s","0"),"%d","0")
			Else
				n = DlgListBoxArray("FileResultList")
				DlgText "FileResultText",Replace$(Replace$(MsgList(21),"%s",CStr$(n)),"%d",CStr$(i + 1))
			End If
		End Select
	Case 3 ' �ı��������Ͽ��ı�������
		Select Case DlgItem$
		Case "FindTextBox"
			SearchSet(0) = DlgText("FindTextBox")
		Case "DirectoryPathBox"
			DlgText "DirectoryPathBox",Trim$(DlgText("DirectoryPathBox"))
			ReDim TempList(0) As String
			DlgListBoxArray "IgnoreSubFolderList",TempList()
			DlgValue "IgnoreSubFolderList",0
			SearchSet(4) = DlgText("DirectoryPathBox")
			SearchSet(11) = ""
		Case "FileTypeBox"
			DlgText "FileTypeBox",Trim$(DlgText("FileTypeBox"))
			SearchSet(9) = DlgText("FileTypeBox")
		Case "FileNameBox"
			DlgText "FileNameBox",Trim$(DlgText("FileNameBox"))
			SearchSet(10) = DlgText("FileNameBox")
		Case "IgnoreSubFolderList"
			If getMsgList(UIDataList,MsgList,"StringSearch",1) = False Then Exit Function
			If SearchSet(4) = "" Then
				MsgBox(MsgList(37),vbOkOnly+vbInformation,MsgList(35))
				Exit Function
			End If
			Temp = Trim$(DlgText("IgnoreSubFolderList"))
			If Temp = "" Then Exit Function
			If InStr(SearchSet(11),Temp) Then Exit Function
			Temp = RemoveBackslash(Temp,"","\",1)
			If InStr(Temp,"...\") = 1 Then
				Temp = RemoveBackslash(SearchSet(4),"","\",1) & Mid$(Temp,4)
			End If
			If LCase$(Temp) = LCase$(SearchSet(4)) Then
				MsgBox(Replace$(MsgList(38),"%s",DlgText("DirectoryPathBox")),vbOkOnly+vbInformation,MsgList(35))
				Exit Function
			End If
			If InStr(LCase$(Temp),LCase$(SearchSet(4))) = 0 Then
				MsgBox(Replace$(MsgList(38),"%s",DlgText("DirectoryPathBox")),vbOkOnly+vbInformation,MsgList(35))
				Exit Function
			End If
			If Dir$(Temp & "\*.*",vbDirectory Or vbHidden) = "" Then
				MsgBox(MsgList(70),vbOkOnly+vbInformation,MsgList(35))
				Exit Function
			End If
			Temp = strReplace(Temp,SearchSet(4),"...")
			If InStr(TextJoinStr & LCase$(SearchSet(11)) & TextJoinStr,TextJoinStr & LCase$(Temp) & TextJoinStr) Then
				MsgBox(MsgList(39),vbOkOnly+vbInformation,MsgList(35))
				DlgText "IgnoreSubFolderList",Temp
				Exit Function
			End If
			TempList = ReSplit(SearchSet(11),TextJoinStr)
			If InsertArray(TempList,Temp,0,True) = True Then
				DlgListBoxArray "IgnoreSubFolderList",TempList()
				DlgValue "IgnoreSubFolderList",0
				SearchSet(11) = StrListJoin(TempList,TextJoinStr)
			End If
		End Select
	Case 6 ' ������ݼ�
		Select Case SuppValue
		Case 1
			If StrToLong(Selected(30)) = 1 Then
				If OpenCHM(CLng(DlgText("SuppValueBox")),1025,Selected(0),OSLanguage,UIFileList) = True Then Exit Function
			End If
			Call Help("StringSearchHelp")
		Case 2
			If getMsgList(UIDataList,MsgList,"RegExpRuleTip",1) = False Then Exit Function
			i = ShowPopupMenu(MsgList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			If i = UBound(MsgList) Then
				If StrToLong(Selected(30)) = 1 Then
					If OpenCHM(CLng(DlgText("SuppValueBox")),1022,Selected(0),OSLanguage,UIFileList) = True Then Exit Function
				End If
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
			SearchSet(0) = DlgText("FindTextBox")
		Case 9, 11
			If getMsgList(UIDataList,MsgList,"StringSearch",1) = False Then Exit Function
			i = DlgValue("FileResultList")
			If i < 0 Then Exit Function
			SearchSet(17) = ReSplit(ReSplit(SearchSet(16),TextJoinStr)(i),vbTab)(0)
			If SearchSet(17) = "" Then
				MsgBox(MsgList(50),vbOkOnly+vbInformation,MsgList(34))
				Exit Function
			End If
			Select Case SuppValue
			Case 9
				ReDim TempList(UBound(Tools)) As String
				For i = 3 To UBound(Tools)
					TempList(i - 3) = Tools(i).sName
				Next i
				i = ShowPopupMenu(TempList,vbPopupUseRightButton)
				If i < 0 Then Exit Function
				ReDim TempList(0) As String
				TempList(0) = SearchSet(17) & JoinStr
				If OpenFile(SearchSet(17),TempList,i + 3,False) = True Then
					If i = 0 Then WriteSettings("Tools")
				End If
			Case 11
				ReDim TempList(UBound(Tools)) As String
				For i = 0 To UBound(Tools)
					TempList(i) = Tools(i).sName
				Next i
				i = ShowPopupMenu(TempList,vbPopupUseRightButton)
				If i < 0 Then Exit Function
				File.FilePath = SearchSet(17)
				If GetFileInfo(File.FilePath,File) = False Then Exit Function
				GetHeaders(File.FilePath,File,StrToLong(Selected(1)),File.FileType)
				Call FileInfoView(File,FreeByteList,i,0,StrToLong(Selected(16)))
			End Select
		End Select
	End Select
End Function


'��ȡ��ǰ�ļ����е�ÿ���ļ�
'ConditionList(0) Ҫ�����ļ���չ����";" �ָ������չ��
'ConditionList(1) Ҫ���Ե��ļ����ƣ�";" �ָ�����ļ���
'ConditionList(2) �Ƿ�������ļ��У�0 = �����ң�1 = ����
'ConditionList(3) �Ƿ���������ļ��У�0 = �����ԣ�1 = ����
'ConditionList(4) �Ƿ���������ļ���0 = �����ԣ�1 = ����
'ConditionList(5) Ҫ���Ե����ļ��У�";" �ָ�������ļ���
Private Function GetFiles(ByVal Folder As String,ConditionList() As String,gFiles() As String,ByVal Msg As String,ByVal ShowMsg As Long) As Long
	Dim i As Long,n As Long,m As Long
	Dim File As String,FindList() As String,SkipList() As String
	m = 20
	ReDim gFiles(m) As String
	FindList = ReSplit(UCase$(ConditionList(0)),";",-1)
	SkipList = ReSplit(UCase$(ConditionList(1)),";",-1)
	On Error Resume Next
	File = Dir$(Folder & "*.*",vbHidden)
	Do While File <> ""
		If GetAttr(Folder & File) And vbDirectory Then GoTo NextNo
		If (UCase$(File) Like "*.LPU") Then GoTo NextNo
		If ConditionList(4) = "1" Then
			If GetAttr(Folder & File) And vbHidden Then GoTo NextNo
		End If
		If ConditionList(1) <> "" Then
			For i = 0 To UBound(SkipList)
				If UCase$(File) Like Trim$(SkipList(i)) Then GoTo NextNo
			Next i
		End If
		If n > m Then
			m = m * 2
			ReDim Preserve gFiles(m) As String
		End If
		If ConditionList(0) <> "" Then
			For i = 0 To UBound(FindList)
				If UCase$(File) Like Trim$(FindList(i)) Then
					gFiles(n) = Folder & File
					n = n + 1
					If ShowMsg > 0 Then
						SetTextBoxString ShowMsg,Replace$(Msg,"%s",CStr$(n))
					ElseIf ShowMsg < 0 Then
						PSL.OutputWnd(0).Clear
						PSL.Output Replace$(Msg,"%s",CStr$(n))
					End If
					Exit For
				End If
			Next i
		Else
			gFiles(n) = Folder & File
			n = n + 1
			If ShowMsg > 0 Then
				SetTextBoxString ShowMsg,Replace$(Msg,"%s",CStr$(n))
			ElseIf ShowMsg < 0 Then
				PSL.OutputWnd(0).Clear
				PSL.Output Replace$(Msg,"%s",CStr$(n))
			End If
		End If
		NextNo:
		File = Dir$()
	Loop
	If ConditionList(2) = "1" Then
		GetFiles = FindSubFiles(Folder,ConditionList,gFiles,n,Msg,ShowMsg)
	End If
	If n > 0 Then n = n - 1
	ReDim Preserve gFiles(n) As String
End Function


'��ȡ���ļ����е�ÿ���ļ�
'ConditionList(0) Ҫ�����ļ���չ��
'ConditionList(1) Ҫ���Ե��ļ�����
'ConditionList(2) �Ƿ�������ļ��У�0 = �����ң�1 = ����
'ConditionList(3) �Ƿ���������ļ��У�0 = �����ԣ�1 = ����
'ConditionList(4) �Ƿ���������ļ���0 = �����ԣ�1 = ����
'ConditionList(5) Ҫ���Ե����ļ��У�";" �ָ�������ļ���
Private Function FindSubFiles(ByVal Folder As String,ConditionList() As String,gFiles() As String,Index As Long,ByVal Msg As String,ByVal ShowMsg As Long) As Long
	Dim i As Long,j As Long,k As Long,m As Long,File As String
	Dim FindList() As String,SkipList() As String,FolderList() As String
	m = Index + 20
	ReDim Preserve gFiles(m) As String
	ReDim subFolders(0) As String
	FindList = ReSplit(UCase$(ConditionList(0)),";",-1)
	SkipList = ReSplit(UCase$(ConditionList(1)),";",-1)
	FolderList = ReSplit(UCase$(ConditionList(5)),";",-1)
	subFolders(0) = Folder
	On Error Resume Next
	Do
		Folder = subFolders(j)
		File = Dir$(Folder & "*.*",vbDirectory Or vbHidden)
		While File <> ""
			If File <> "." And File <> ".." Then
				If GetAttr(Folder & File) And vbDirectory Then
             		If ConditionList(3) = "1" Then
						If GetAttr(Folder & File) And vbHidden Then GoTo NextNo
					End If
					If ConditionList(5) <> "" Then
						For i = 0 To UBound(FolderList)
							If InStr(UCase$(Folder & File),FolderList(i)) = 1 Then GoTo NextNo
						Next i
					End If
					k = k + 1
            	 	ReDim Preserve subFolders(k) As String
             		subFolders(k) = Folder & File & "\"
             	ElseIf Folder <> subFolders(0) Then
					If (UCase$(File) Like "*.LPU") Then GoTo NextNo
					If ConditionList(4) = "1" Then
						If GetAttr(Folder & File) And vbHidden Then GoTo NextNo
					End If
					If ConditionList(1) <> "" Then
						For i = 0 To UBound(SkipList)
							If UCase$(File) Like Trim$(SkipList(i)) Then GoTo NextNo
						Next i
					End If
					If Index > m Then
						m = m * 2
						ReDim Preserve gFiles(m) As String
					End If
					If ConditionList(0) <> "" Then
						For i = 0 To UBound(FindList)
							If UCase$(File) Like Trim$(FindList(i)) Then
								gFiles(Index) = Folder & File
								Index = Index + 1
								If ShowMsg > 0 Then
									SetTextBoxString ShowMsg,Replace$(Msg,"%s",CStr$(Index))
								ElseIf ShowMsg < 0 Then
									PSL.OutputWnd(0).Clear
									PSL.Output Replace$(Msg,"%s",CStr$(Index))
								End If
								Exit For
							End If
						Next i
					Else
						gFiles(Index) = Folder & File
						Index = Index + 1
						If ShowMsg > 0 Then
							SetTextBoxString ShowMsg,Replace$(Msg,"%s",CStr$(Index))
						ElseIf ShowMsg < 0 Then
							PSL.OutputWnd(0).Clear
							PSL.Output Replace$(Msg,"%s",CStr$(Index))
						End If
					End If
				End If
				NextNo:
			End If
			File = Dir$()
		Wend
		j = j + 1
	Loop Until j = k + 1
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


'����ָ���ı����ִ�
Private Function FindStringCount(DataList() As FindStrList,ByVal RegEx As Object,File As FILE_PROPERTIE,FN As FILE_IMAGE, _
				SetList() As Long,Mode As Long,Optional ByVal ShowMsg As Long) As Long
	Dim i As Long,j As Long,m As Long,n As Long,SkipVal As Long,RSize As Long,Max As Long,Temp As String
	On Error GoTo ExitFunction
	ReDim DataList(0) As FindStrList
	With File
		.FileSize = FileLen(.FilePath)
		If .FileSize = 0 Then Exit Function
		Mode = LoadFile(.FilePath,FN,0,0,0,Mode,0,0)
		If Mode < -1 Then Exit Function
		If ShowMsg > 0 Then Temp = GetTextBoxString(ShowMsg) & " "
		GetHeaders(File.FilePath,File,StrToLong(Selected(1)),File.FileType)
		Max = 15
		ReDim DataList(Max) As FindStrList
		For j = 0 To .MaxSecIndex
			If j < .MaxSecIndex Then
				i = .SecList(j).lPointerToRawData
				RSize = .SecList(j).lPointerToRawData + .SecList(j).lSizeOfRawData - 1
				SkipVal = i - 1
			ElseIf .SecList(.MaxSecIndex).lSizeOfRawData > 0 Then
				i = .SecList(j).lPointerToRawData
				RSize = .SecList(j).lPointerToRawData + .SecList(j).lSizeOfRawData - 1
				SkipVal = i - 1
			Else
				Exit For
			End If
			Do While i < RSize
				If i <= SkipVal Then i = SkipVal + 1
				If i > SkipVal Then
					'�ų������ļ�ͷ�������Σ�.NET �û��ִ��������
					m = SkipHeader(File,i,SkipVal,1,1)
					If SetList(2) > 0 Then
						If m = 0 Or m = 1 Or m = 5 Or m = 10 Then Exit Do
					End If
					If i > SkipVal Or SkipVal > RSize Then SkipVal = RSize + 1
					If i > RSize Then Exit Do
				End If
				If n >= Max Then
					Max = Max * 2
					ReDim Preserve DataList(Max) As FindStrList
				End If
				DataList(n).StartPos = i
				DataList(n).EndPos = SkipVal - 1
				DataList(n).inSecID = j
				If  SetList(1) = 0 Then
					Set DataList(n).Matches = RegEx.Execute(ByteToString(GetBytes(FN,SkipVal - i,i,Mode),SetList(0)) & vbNullChar)
				Else
					Set DataList(n).Matches = RegEx.Execute(DelAccKey(ByteToString(GetBytes(FN,SkipVal - i,i,Mode),SetList(0))) & vbNullChar)
				End If
				FindStringCount = FindStringCount + DataList(n).Matches.Count
				n = n + 1
				If ShowMsg > 0 Then
					SetTextBoxString ShowMsg,Temp & Format$(SkipVal / .FileSize,"#%")
				ElseIf ShowMsg < 0 Then
					PSL.OutputWnd(0).Clear
					PSL.Output Temp & Format$(SkipVal / .FileSize,"#%")
				End If
				'DoEvents 'ת�ÿ���Ȩ���������ϵͳ���������¼�
				If StopProcess(StopHwnd,VK_ESCAPE) = True Then
					FindStringCount = -1
					Exit For
				End If
			Loop
		Next j
	End With
	ExitFunction:
	If n = 0 Then Set DataList(0).Matches = RegEx.Execute("")
	If n > 0 Then n = n - 1
	ReDim Preserve DataList(n) As FindStrList
	If Mode <> 0 Then UnLoadFile(FN,0,Mode)
End Function


'��ʾ��ȡ��ȡ�����̰�ť
Private Sub ShowButton(ByVal ButtonHwnd As Long,ByVal KeyHwnd As Long,ByVal Mode As Boolean)
	If Mode = True Then
		'��ʾȡ��������ť
		ShowWindow ButtonHwnd, SW_SHOW
		'���� Esc ���������Ի������Ӧ���˳�
		EnableWindow KeyHwnd, False
		'���������¼����Ϊ GetAsyncKeyState ���¼���һ��
		GetAsyncKeyState KeyHwnd
	Else
		'����ȡ��������ť
		ShowWindow ButtonHwnd, SW_HIDE
		'���� Esc ���������Ի������Ӧ��������
		EnableWindow KeyHwnd, True
	End If
End Sub


'ȡ������ȷ��
Private Function StopProcess(ByVal ButtonHwnd As Long,ByVal KeyHwnd As Long) As Boolean
	If GetAsyncKeyState(KeyHwnd) < 0 Then
		SendMessageLNG(ButtonHwnd,BM_SETSTATE,True,0)
		If MsgBox(StopMsg,vbYesNo+vbInformation,StopTitle) = vbNo Then Exit Function
	ElseIf SendMessageLNG(ButtonHwnd,BM_GETSTATE,0,0) = WM_MBUTTONUP Then
		SendMessageLNG(ButtonHwnd,BM_SETSTATE,True,0)
		If MsgBox(StopMsg,vbYesNo+vbInformation,StopTitle) = vbNo Then Exit Function
	Else
		Exit Function
	End If
	'����ȡ��������ť
	ShowWindow ButtonHwnd, SW_HIDE
	'���� Esc ���������Ի������Ӧ���˳�
	EnableWindow KeyHwnd, True
	StopProcess = True
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
