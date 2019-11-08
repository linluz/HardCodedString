Attribute VB_Name = "modUpdate"
'' Update Module for PSlHardCodedString.bas
'' (c) 2010-2019 by wanfu (Last modified on 2019.09.16

'#Uses "modCommon.bas"

Option Explicit

Private Const MacroLoc = Left$(MacroDir,InStrRev(MacroDir,"\") - 1)
'Public UpdateSet() As String,UpdateSetBak() As String
Public Const DefaultObject = "Microsoft.XMLHTTP;Msxml2.XMLHTTP"
'Private Const JoinStr = vbFormFeed  'vbBack
Private Const DefaultWaitTimes = 2    '2��
Private Const updateMainFile = "PSLHardCodedString.bas"
Private Const updateINIFile = "PSLMacrosUpdates.rar"
Private Const updateINIMainUrl = "http://jp.wanfutrade.com/download/PSLMacrosUpdates.rar"
Private Const updateINIMinorUrl = "http://www.wanfutrade.com/software/hanhua/PSLMacrosUpdates.rar"
Public Const updateMainUrl = "http://jp.wanfutrade.com/download/PSLHardCodedString.rar"
Public Const updateMinorUrl = "http://www.wanfutrade.com/software/hanhua/PSLHardCodedString.rar"


'��Ⲣ�����°汾
Public Function CheckUpdate(UpdateSet() As String,ByVal ShowMsg As Long) As Boolean
	Dim i As Long,j As Long,n As Long
	'��ȡ�������ݲ�����°汾
	If CheckArray(UpdateSet) = True Then
		If UpdateSet(0) = "" Then UpdateSet(0) = "1"
		If UpdateSet(1) = "" Then UpdateSet(1) = updateMainUrl & vbCrLf & updateMinorUrl
		If UpdateSet(2) = "" Or (UpdateSet(2) <> "" And Dir$(UpdateSet(2)) = "") Then
			getCMDPath(".rar",UpdateSet(2),UpdateSet(3))
		End If
	Else
		UpdateSet = ReSplit("1" & JoinStr & updateMainUrl & vbCrLf & updateMinorUrl & JoinStr & _
					getCMDPath(".rar","","") & JoinStr & "7" & JoinStr,JoinStr)
	End If
	If UpdateSet(0) <> "" And UpdateSet(0) <> "2" Then
		If UpdateSet(5) <> "" Then
			i = CLng(DateDiff("d",CDate(UpdateSet(5)),Date))
			j = StrComp(Format(Date,"yyyy-MM-dd"),UpdateSet(5))
			If UpdateSet(4) <> "" Then n = i - CLng(UpdateSet(4))
		End If
		If UpdateSet(5) = "" Or (j = 1 And n >= 0) Then
			i = Download(UpdateSet,UpdateSet(1),StrToLong(UpdateSet(0)),ShowMsg)
			If i > 0 Then
				If UpdateSet(5) < Format(Date,"yyyy-MM-dd") Then
					UpdateSet(5) = Format(Date,"yyyy-MM-dd")
					WriteSettings("Update")
				End If
				If i = 3 Then CheckUpdate = True
			End If
		End If
	End If
End Function


'��Ⲣ�����°汾
'Mode: 0 = �Զ����ز���װ, 1 = �û�����, 2 = �ر�, 3 = �ֶ���� 4 = ����, 5 = ����
'����ֵ = 0 ʧ��, 1 = ��������Ϣ�ɹ�, 2 = �����ļ��ɹ�, 3 = ���³ɹ�
Public Function Download(UpdateSet() As String,ByVal Url As String,ByVal Mode As Long,ByVal ShowMsg As Long) As Long
	Dim i As Long,j As Long,m As Long,n As Long
	Dim xmlHttp As Object,Body() As Byte,BodyBak() As Byte
	Dim UrlList() As String,TempList() As String,TempArray() As String,MsgList() As String
	Dim ExePath As String,Argument As String,Temp As String,UpdateData As INIFILE_DATA
	Dim WebVersion As String,WebBuild As String,IniBuild As String,TempPath As String

	If Mode = 2 Then Exit Function
	If getMsgList(UIDataList,MsgList,"Download",1) = False Then Exit Function

	'��ȡ��ѹ����Ͳ���
	If CheckArray(UpdateSet) = True Then
		If Url = "" Then Url = UpdateSet(1)
		ExePath = Trim$(UpdateSet(2))
		Argument = UpdateSet(3)
	End If

	'����ѹ����Ͳ���
	If Mode < 5 Then
		If ShowMsg > 0 Then
			SetTextBoxString ShowMsg,IIf(Mode = 4,MsgList(22),MsgList(23))
		ElseIf ShowMsg < 0 Then
			PSL.OutputWnd.Clear
			PSL.Output IIf(Mode = 4,MsgList(22),MsgList(23))
		End If
	End If
	If ExePath = "" Then
		If Mode < 5 Then
			MsgBox(IIf(Mode <> 4,MsgList(1),MsgList(2)) & vbCrLf & MsgList(3),vbOkOnly+vbInformation,MsgList(0))
		End If
		Exit Function
	End If
	If Url = "" Or Argument = "" Then
		i = 1
	ElseIf InStr(Argument,"%1") = 0 Then
		i = 1
	'ElseIf InStr(Argument,"%2") = 0 Then
	'	i = 1
	ElseIf InStr(Argument,"%3") = 0 Then
		i = 1
	End If
	If i = 1 Then
		If Mode < 5 Then
			MsgBox(IIf(Mode <> 4,MsgList(1),MsgList(2)) & vbCrLf & MsgList(4),vbOkOnly+vbInformation,MsgList(0))
		End If
		Exit Function
	End If

	'������ط����Ƿ����
	On Error Resume Next
	TempList = ReSplit(DefaultObject,";")
	For i = 0 To UBound(TempList)
		Set xmlHttp = CreateObject(TempList(i))
		If Not xmlHttp Is Nothing Then Exit For
	Next i
	If xmlHttp Is Nothing Then
		If Mode < 5 Then
			Err.Source = StrListJoin(TempList,"; ")
			Call sysErrorMassage(Err,2)
		End If
		Exit Function
	End If
	On Error GoTo 0

	'��ȡ����������Ϣ
	j = 0
	If Mode <> 4 Then
		'�ϲ����������ļ���Ĭ�Ϻ��Զ���������ַ
		UrlList = ReSplit(Url,vbCrLf)
		For i = 0 To UBound(UrlList)
			Temp = Trim$(UrlList(i))
			n = InStrRev(Temp,"/")
			If n > 0 Then
				UrlList(i) = Left$(Temp,n) & updateINIFile
			End If
		Next i
		UrlList = ReSplit(updateINIMainUrl & vbCrLf & updateINIMinorUrl & vbCrLf & StrListJoin(UrlList,vbCrLf),vbCrLf)

		'���ز������������ļ�
		For i = 0 To UBound(UrlList)
			'����ֵ��1 = �ɹ���0 = ʧ�ܣ�-1 = �ļ������ڣ�-2 = ����
			Select Case DownloadFile(Body,xmlHttp,UrlList(i))
			Case 1
				Temp = BytesToBstr(Body,"utf-8")
				If Temp <> "" Then
					If InStr(LCase$(Temp),LCase$(AppName)) Then
						'�����������ļ�
						IniBuild = CheckUpdateINIFile(UpdateData,Temp)
						If IniBuild <> "" Then Exit For
					End If
				End If
			Case -2
				Set xmlHttp = Nothing
				If Mode < 5 Then MsgBox(MsgList(35),vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End Select
		Next i

		'��ʾ������Ϣ
		If UpdateData.Title <> "" Then
			Download = 1
			Select Case StrComp(UpdateData.Title,Version) + StrComp(IniBuild,Build)
			Case Is > 0
				If Mode = 1 Or Mode = 3 Then
					Temp = Replace$(Replace$(MsgList(15),"%s",UpdateData.Title),"%d",IniBuild) & vbCrLf & vbCrLf & MsgList(20)
					If MsgBox(Temp & vbCrLf & StrListJoin(UpdateData.Value,vbCrLf),vbYesNo+vbInformation,MsgList(17)) = vbNo Then
						Set xmlHttp = Nothing
						Exit Function
					End If
				End If
			Case 0
				If Mode = 3 Then
					Temp = Replace$(Replace$(MsgList(14),"%s",UpdateData.Title),"%d",IniBuild) & vbCrLf & MsgList(21)
					If MsgBox(Temp,vbYesNo+vbInformation,MsgList(17)) = vbNo Then
						Set xmlHttp = Nothing
						Exit Function
					End If
				Else
					Set xmlHttp = Nothing
					Exit Function
				End If
			Case Is < 0
				If Mode = 3 Then
					MsgBox Replace$(Replace$(MsgList(14),"%s",UpdateData.Title),"%d",IniBuild),vbOkOnly+vbInformation,MsgList(16)
				End If
				Set xmlHttp = Nothing
				Exit Function
			End Select
		End If
	End If

	'���س����ļ�
	If Mode < 4 Then
		If ShowMsg > 0 Then
			SetTextBoxString ShowMsg,MsgList(24),True
		ElseIf ShowMsg < 0 Then
			PSL.Output MsgList(24)
		End If
	End If
	m = 0: n = 0: j = 0
	If UpdateData.Title <> "" Then
		UrlList = ClearTextArray(ReSplit(Url & vbCrLf & StrListJoin(UpdateData.Item,vbCrLf),vbCrLf),True)
	Else
		UrlList = ReSplit(Url,vbCrLf)
	End If
	For i = 0 To UBound(UrlList)
		If UrlList(i) <> "" Then
			'����ֵ��1 = �ɹ���0 = ʧ�ܣ�-1 = �ļ������ڣ�-2 = ����
			Select Case DownloadFile(Body,xmlHttp,UrlList(i))
			Case 1
				If Mode <> 4 Then Exit For
				If LenB(BodyBak) = 0 Then BodyBak = Body
			Case 0
				If Mode = 4 Then
					ReDim Preserve TempList(m) As String
					TempList(m) = UrlList(i)
				End If
				m = m + 1
			Case -1
				If Mode = 4 Then
					ReDim Preserve TempArray(n) As String
					TempArray(n) = UrlList(i)
				End If
				n = n + 1
			Case -2
				If Mode < 5 Then MsgBox(MsgList(35),vbOkOnly+vbInformation,MsgList(0))
				Set xmlHttp = Nothing
				Exit Function
			End Select
			j = j + 1
		End If
	Next i
	If m + n <> 0 Then
		If Mode <> 4 Then
			If m = j Or n = j Then
				If Mode < 5 Then MsgBox(MsgList(1) & vbCrLf & IIf(n = j,MsgList(5),MsgList(6)),vbOkOnly+vbInformation,MsgList(0))
				Set xmlHttp = Nothing
				Exit Function
			End If
		Else
			If m <> 0 And n <> 0 Then
				Temp = MsgList(2) & vbCrLf & MsgList(33) & vbCrLf & StrListJoin(TempArray,vbCrLf) & _
						vbCrLf & vbCrLf & MsgList(34) & vbCrLf & StrListJoin(TempList,vbCrLf)
			ElseIf m <> 0 Then
				Temp = MsgList(2) & vbCrLf & MsgList(34) & vbCrLf & StrListJoin(TempList,vbCrLf)
			ElseIf n <> 0 Then
				Temp = MsgList(2) & vbCrLf & MsgList(33) & vbCrLf & StrListJoin(TempArray,vbCrLf)
			End If
			MsgBox(Temp,vbOkOnly+vbInformation,MsgList(12))
			Set xmlHttp = Nothing
			Exit Function
		End If
	End If
	Set xmlHttp = Nothing

	'�������صĳ����ļ�
	If Mode = 4 Then Body = BodyBak
	TempPath = MacroLoc & "\temp\"
	Temp = TempPath & "temp.rar"
	On Error Resume Next
	If Dir$(TempPath & "*.*") = "" Then MkDir TempPath
	If BytesToFile(Body,Temp) = False Then
		i = FreeFile
		Open Temp For Binary Access Write As #i
		Put #i,,Body
		Close #i
	End If
	On Error GoTo 0

	'��ѹ�ļ�
	i = 0
	If Dir$(Temp) <> "" Then
		i = ExtractFile(Temp,TempPath,ExePath,Argument)
		If i = 1 Then
			Temp = TempPath & updateMainFile
		ElseIf Mode < 5 Then
			If i = -2 Then
				Temp = Mid$(Left$(ExePath,InStrRev(ExePath,".") - 1),InStrRev(ExePath,"\") + 1)
				Temp = Replace$(MsgList(8),"%s",Temp) & vbCrLf & _
						Replace$(MsgList(9),"%s",ExePath) & vbCrLf & _
						Replace$(MsgList(10),"%s",Argument) & vbCrLf & vbCrLf & MsgList(11)
				MsgBox(Temp,vbOkOnly+vbInformation,MsgList(0))
			ElseIf i = -3 Then
				MsgBox(IIf(Mode <> 4,MsgList(1),MsgList(2)) & vbCrLf & MsgList(7),vbOkOnly+vbInformation,MsgList(0))
			End If
		End If
	End If
	If i <> 1 Then
		DelDirs(TempPath)
		Exit Function
	End If

	'��ȡ���صĳ���汾��
	If Mode < 4 Then
		If ShowMsg > 0 Then
			SetTextBoxString ShowMsg,MsgList(26),True
		ElseIf ShowMsg < 0 Then
			PSL.Output MsgList(26)
		End If
	End If
	WebVersion = GetWebVersion(Temp,"Const Version = ")
	WebBuild = GetWebVersion(Temp,"Const Build = ")
	If WebVersion = "" Or WebBuild = "" Then
		If Mode < 5 Then MsgBox(MsgList(19),vbOkOnly+vbInformation,MsgList(0))
		DelDirs(TempPath)
		Exit Function
	End If

	'�Ƚϰ汾����ʾ������Ϣ
	If Url <> StrListJoin(UrlList,vbCrLf) Then UpdateSet(1) = StrListJoin(UrlList,vbCrLf)
	If Mode = 4 Then
		MsgBox(MsgList(13) & vbCrLf & Replace$(Replace$(MsgList(14),"%s",WebVersion),"%d",WebBuild),vbOkOnly+vbInformation,MsgList(12))
		Download = 2
		DelDirs(TempPath)
		Exit Function
	End If
	n = StrComp(WebVersion,Version) + StrComp(WebBuild,Build)
	If Mode < 5 Then
		If n > 0 Or (n = 0 And Mode = 3) Then
			If UpdateData.Title = "" Then
				If Mode = 1 Then
					Temp = Replace$(Replace$(MsgList(15),"%s",WebVersion),"%d",WebBuild)
				ElseIf Mode = 3 Then
					If n = 0 Then
						Temp = Replace$(Replace$(MsgList(14),"%s",WebVersion),"%d",WebBuild) & vbCrLf & MsgList(21)
					Else
						Temp = Replace$(Replace$(MsgList(15),"%s",WebVersion),"%d",WebBuild)
					End If
				End If
				If MsgBox(Temp,vbYesNo+vbInformation,MsgList(17)) = vbNo Then
					DelDirs(TempPath)
					Exit Function
				End If
			ElseIf UpdateData.Title <> WebVersion Or IniBuild <> WebBuild Then
				Temp = MsgList(1) & vbCrLf & Replace$(Replace$(MsgList(29),"%s",WebVersion),"%d",WebBuild) & vbCrLf & MsgList(32)
				MsgBox(Temp,vbOkOnly+vbInformation,MsgList(0))
				DelDirs(TempPath)
				Exit Function
			End If
		Else
			If Mode < 2 Then
				If ShowMsg > 0 Then
					SetTextBoxString ShowMsg,Replace$(Replace$(MsgList(30),"%s",WebVersion),"%d",WebBuild),True
				ElseIf ShowMsg < 0 Then
					PSL.Output Replace$(Replace$(MsgList(30),"%s",WebVersion),"%d",WebBuild)
				End If
			Else
				MsgBox Replace$(Replace$(MsgList(31),"%s",WebVersion),"%d",WebBuild),vbOkOnly+vbInformation,MsgList(16)
			End If
			DelDirs(TempPath)
			Exit Function
		End If
	ElseIf n < 1 Then
		DelDirs(TempPath)
		Exit Function
	End If

	'��װ�°汾
	If Mode < 5 Then
		If ShowMsg > 0 Then
			SetTextBoxString ShowMsg,MsgList(27),True
		ElseIf ShowMsg < 0 Then
			PSL.Output MsgList(27)
		End If
		If SetupNewVersion(TempPath,MacroLoc) = True Then
			If ShowMsg > 0 Then
				SetTextBoxString ShowMsg,MsgList(28),True
			ElseIf ShowMsg < 0 Then
				PSL.Output MsgList(28)
			End If
			MsgBox(MsgList(18),vbOkOnly+vbInformation,MsgList(17))
			Download = 3
		End If
	ElseIf SetupNewVersion(TempPath,MacroLoc) = True Then
		Download = 3
	End If
	DelDirs(TempPath)
End Function


'��ע����л�ȡ RAR ��չ����Ĭ�ϳ���
Public Function getCMDPath(ByVal ExtName As String,CmdPath As String,Argument As String) As String
	Dim i As Long,WshShell As Object,TempArray() As String
	On Error Resume Next
	Set WshShell = CreateObject("WScript.Shell")
	If WshShell Is Nothing Then
		getCMDPath = CmdPath & JoinStr & Argument
		Err.Source = "WScript.Shell"
		Call sysErrorMassage(Err,2)
		Exit Function
	End If
	ExtName = WshShell.RegRead("HKCR\" & ExtName & "\")
	If ExtName <> "" Then
		CmdPath = WshShell.RegRead("HKCR\" & ExtName & "\shell\open\command\")
	End If
	On Error GoTo 0
	Set WshShell = Nothing
	If CmdPath <> "" Then
		i = InStr(CmdPath,".")
		If i > 0 Then Argument = Trim$(Mid$(CmdPath,InStr(i,CmdPath," ")))
		CmdPath = Left$(CmdPath,Len(CmdPath) - Len(Argument))
		TempArray = ReSplit(CmdPath,"%")
		If UBound(TempArray) = 2 Then
			CmdPath = Replace$(CmdPath,"%" & TempArray(1) & "%",Environ(TempArray(1)),,1)
		End If
		CmdPath = RemoveBackslash(CmdPath,"""","""",1)

		If InStr(CmdPath,"\") = 0 Then
			If Dir$(Environ("SystemRoot") & "\system32\" & CmdPath) <> "" Then
				CmdPath = Environ("SystemRoot") & "\system32\" & CmdPath
			ElseIf Dir$(Environ("SystemRoot") & "\" & CmdPath) <> "" Then
				CmdPath = Environ("SystemRoot") & "\" & CmdPath
			End If
		End If

		If InStr(LCase$(CmdPath),"winrar.exe") Then
			If Argument <> "" Then
				If InStr(Argument,"""%1""") Then
					Argument = "e -ibck " & Replace$(Argument,"""%1""","""%1"" %2 ""%3""")
				ElseIf InStr(Argument,"%1") Then
					Argument = "e -ibck " & Replace$(Argument,"%1","""%1"" %2 ""%3""")
				Else
					Argument = "e -ibck ""%1"" %2 ""%3"" " & Argument
				End If
			Else
				Argument = "e -ibck ""%1"" %2 ""%3"""
			End If
		ElseIf InStr(LCase$(CmdPath),"winzip.exe") Then
			CmdPath = strReplace(CmdPath,"WinZip.exe","WzunZip.exe")
			If Argument <> "" Then
				If InStr(Argument,"""%1""") Then
					Argument = Replace$(Argument,"""%1""","""%1"" %2 ""%3""")
				ElseIf InStr(Argument,"%1") Then
					Argument = Replace$(Argument,"%1","""%1"" %2 ""%3""")
				Else
					Argument = """%1"" %2 ""%3"" " & Argument
				End If
			Else
				Argument = " ""%1"" %2 ""%3"""
			End If
		ElseIf InStr(LCase$(CmdPath),"wzunzip.exe") Then
			If Argument <> "" Then
				If InStr(Argument,"""%1""") Then
					Argument = Replace$(Argument,"""%1""","""%1"" %2 ""%3""")
				ElseIf InStr(Argument,"%1") Then
					Argument = Replace$(Argument,"%1","""%1"" %2 ""%3""")
				Else
					Argument = """%1"" %2 ""%3"" " & Argument
				End If
			Else
				Argument = " ""%1"" %2 ""%3"""
			End If
		ElseIf InStr(LCase$(CmdPath),"7z.exe") Then
			If Argument <> "" Then
				If InStr(Argument,"""%1""") Then
					Argument = "e -r " & Replace$(Argument,"""%1""","""%1"" -o""%3"" %2")
				ElseIf InStr(Argument,"%1") Then
					Argument = "e -r " & Replace$(Argument,"%1","""%1"" -o""%3"" %2")
				Else
					Argument = "e -r ""%1"" -o""%3"" %2 " & Argument
				End If
			Else
				Argument = "e -r ""%1"" -o""%3"" %2"
			End If
		ElseIf InStr(LCase$(CmdPath),"7zfm.exe") Then
			CmdPath = strReplace(CmdPath,"7zFM.exe","7z.exe")
			If Argument <> "" Then
				If InStr(Argument,"""%1""") Then
					Argument = "e -r " & Replace$(Argument,"""%1""","""%1"" -o""%3"" %2")
				ElseIf InStr(Argument,"%1") Then
					Argument = "e -r " & Replace$(Argument,"%1","""%1"" -o""%3"" %2")
				Else
					Argument = "e -r ""%1"" -o""%3"" %2 " & Argument
				End If
			Else
				Argument = "e -r ""%1"" -o""%3"" %2"
			End If
		ElseIf InStr(LCase$(CmdPath),"haozip.exe") Then
			CmdPath = strReplace(CmdPath,"HaoZip.exe","HaoZipC.exe")
			If Argument <> "" Then
				If InStr(Argument,"""%1""") Then
					Argument = "e " & Replace$(Argument,"""%1""","""%1"" -r -o""%3"" %2")
				ElseIf InStr(Argument,"%1") Then
					Argument = "e " & Replace$(Argument,"%1","""%1"" -r -o""%3"" %2")
				Else
					Argument = "e ""%1"" -r -o""%3"" %2 " & Argument
				End If
			Else
				Argument = "e ""%1"" -r -o""%3"" %2"
			End If
		ElseIf InStr(LCase$(CmdPath),"haozipc.exe") Then
			If Argument <> "" Then
				If InStr(Argument,"""%1""") Then
					Argument = "e " & Replace$(Argument,"""%1""","""%1"" -r -o""%3"" %2")
				ElseIf InStr(Argument,"%1") Then
					Argument = "e " & Replace$(Argument,"%1","""%1"" -r -o""%3"" %2")
				Else
					Argument = "e ""%1"" -r -o""%3"" %2 " & Argument
				End If
			Else
				Argument = "e ""%1"" -r -o""%3"" %2"
			End If
		End If
	End If
	getCMDPath = CmdPath & JoinStr & Argument
End Function


'�� wTimes �ȴ�ʱ������ѯ��������״̬
'tValue ΪĿ��ֵ���� wTimes = 0 ʱΪĬ�ϵȴ�ʱ��
Private Function OnReadyStateChange(xmlHttp As Object,ByVal tValue As Long,wTimes As Long) As Long
	Dim StartTime As Long
	StartTime = Timer
	If wTimes = 0 Then wTimes = DefaultWaitTimes
	OnReadyStateChange = xmlHttp.readyState
	Do While OnReadyStateChange < tValue
		OnReadyStateChange = xmlHttp.readyState
		If (Timer - StartTime) > wTimes Then Exit Do
	Loop
End Function


'ת������������Ϊָ�������ʽ���ַ�
Public Function BytesToBstr(strBody As Variant,ByVal outCode As String) As String
	Dim objStream As Object
	If LenB(strBody) = 0 Or outCode = "" Then Exit Function
	On Error GoTo ErrorMsg
	Set objStream = CreateObject("Adodb.Stream")
	If Not objStream Is Nothing Then
		With objStream
			.Type = 1
			.Mode = 3
			.Open
			.Write strBody
			.Position = 0
			.Type = 2
			.Charset = outCode
			BytesToBstr = .ReadText
			.Close
		End With
		Set objStream = Nothing
	End If
	Exit Function
	ErrorMsg:
	Err.Source = "Adodb.Stream"
	Call sysErrorMassage(Err,1)
End Function


'д����������ݵ��ļ�
Public Function BytesToFile(strBody As Variant,ByVal File As String) As Boolean
	Dim objStream As Object
	BytesToFile = False
	If LenB(strBody) = 0 Or File = "" Then Exit Function
	On Error GoTo ErrorMsg
	Set objStream = CreateObject("Adodb.Stream")
	If Not objStream Is Nothing Then
		With objStream
			.Type = 1
			.Mode = 3
			.Open
			.Write(strBody)
			.Position = 0
			.SaveToFile File,2
			.Flush
			.Close
		End With
		Set objStream = Nothing
		BytesToFile = True
	End If
	Exit Function
	ErrorMsg:
	Err.Source = "Adodb.Stream"
	Call sysErrorMassage(Err,1)
End Function


'�����ļ�
'����ֵ��1 = �ɹ���0 = ʧ�ܣ�-1 = �ļ������ڣ�-2 = ����
Private Function DownloadFile(Body() As Byte,xmlHttp As Object,ByVal Url As String) As Long
	Dim FileSize As Long
	ReDim Body(0) As Byte
	If Trim$(Url) = "" Then Exit Function
	On Error GoTo ExitFunction
	xmlHttp.Open "HEAD",Url,False,"",""
	xmlHttp.send()
	If OnReadyStateChange(xmlHttp,4,DefaultWaitTimes) = 4 Then
		'FileSize = CLng(ReSplit(xmlHttp.getResponseHeader("Content-Range"),"/")(1))
		FileSize = CLng(xmlHttp.getResponseHeader("Content-Length"))
	End If
	xmlHttp.Abort
	If FileSize > 0 Then
		xmlHttp.Open "GET",Url,False,"",""
		xmlHttp.setRequestHeader "Referer", Left$(Url, InStr(InStr(Url, "//") + 2, Url, "/") - 1)
		xmlHttp.setRequestHeader "Accept", "*/*"
		'xmlHttp.setRequestHeader "Range", "bytes = " & FileSize
		xmlHttp.setRequestHeader "Content-Type", "application/octet-stream"
		xmlHttp.setRequestHeader "If-Modified-Since", "0"
		xmlHttp.setRequestHeader "Pragma", "no-cache"
		xmlHttp.setRequestHeader "Cache-Control", "no-cache"
		xmlHttp.send()
		If OnReadyStateChange(xmlHttp,4,DefaultWaitTimes) = 4 Then
			If xmlHttp.Status = 200 Then
				Body = xmlHttp.responseBody
				If LenB(Body) = FileSize Then DownloadFile = 1
			End If
		End If
	Else
		DownloadFile = -1
	End If
	On Error GoTo 0
	ExitFunction:
	xmlHttp.Abort
	If Err.Number <> 0 Then DownloadFile = -2
End Function


'�����������ļ�
Private Function CheckUpdateINIFile(Data As INIFILE_DATA,ByVal UpdateINIText As String) As String
	Dim i As Long,j As Long,m As Long,n As Long
	Dim DefaultLng As String,LangName As String,DataList() As INIFILE_DATA

	If Trim$(UpdateINIText) = "" Then Exit Function
	If getINIFile(DataList,"",UpdateINIText,2) = False Then Exit Function
	For i = 0 To UBound(DataList)
		With DataList(i)
			Select Case .Title
			Case "Option"
				For j = 0 To UBound(.Item)
					If .Item(j) = "DefaultLanguage" Then
						DefaultLng = LCase$(Trim$(.Value(j)))
						Exit For
					End If
				Next j
			Case "Language"
				UpdateINIText = LCase$(OSLanguage)
				For j = 0 To UBound(.Item)
					If InStr(LCase$(.Value(j)),UpdateINIText) Then
						LangName = LCase$(.Item(j))
						Exit For
					End If
				Next j
				If LangName = "" Then LangName = DefaultLng
			Case AppName
				For j = 0 To UBound(.Item)
					UpdateINIText = LCase$(.Item(j))
					If UpdateINIText = "version" Then
						Data.Title = Trim$(.Value(j))
					ElseIf UpdateINIText = "build" Then
						CheckUpdateINIFile = Trim$(.Value(j))
					ElseIf InStr(UpdateINIText,"url_") Then
						ReDim Preserve Data.Item(m) 'As String
						Data.Item(m) = Trim$(.Value(j))
						m = m + 1
					ElseIf InStr(UpdateINIText,"des_" & LangName) Then
						ReDim Preserve Data.Value(n) 'As String
						Data.Value(n) = Trim$(.Value(j))
						n = n + 1
					End If
				Next j
				Exit For
			End Select
		End With
	Next i
	If m = 0 Or n = 0 Then
		CheckUpdateINIFile = ""
		Data.Title = ""
		ReDim Data.Item(0) 'As String,
		ReDim Data.Value(0) 'As String
	Else
		If CheckUpdateINIFile = "" Then CheckUpdateINIFile = Build
		ReDim Preserve Data.Item(m - 1) 'As String
		ReDim Preserve Data.Value(n - 1) 'As String
	End If
End Function


'��ѹ�ļ�
'����ֵ 1 = �ɹ���0 = Ҫ��ѹ���ļ������ڻ��СΪ�㣬-1 = ���������Ҳ�����-2 = ��ѹ�����Ҳ�����-3 = ��ѹ����
Private Function ExtractFile(ByVal File As String,ByVal Path As String,ByVal ExePath As String,ByVal Argument As String) As Long
	Dim WshShell As Object,TempList() As String

	On Error Resume Next
	Set WshShell = CreateObject("WScript.Shell")
	If WshShell Is Nothing Then
		Err.Source = "WScript.Shell"
		Call sysErrorMassage(Err,2)
		Exit Function
	End If
	On Error GoTo 0

	If Dir$(File) = "" Then Exit Function
	If FileLen(File) = 0 Then Exit Function
	If ExePath <> "" Then
		TempList = ReSplit(ExePath,"%",-1)
		If UBound(TempList) >= 2 Then
			ExePath = Replace$(ExePath,"%" & TempList(1) & "%",Environ(TempList(1)),,1)
		End If
		ExePath = RemoveBackslash(ExePath,"""","""",1)
	End If
	If Argument <> "" Then
		If InStr(Argument,"""%1""") Then Argument = Replace$(Argument,"%1",File)
		If InStr(Argument,"""%2""") Then Argument = Replace$(Argument,"%2","*")
		If InStr(Argument,"""%3""") Then Argument = Replace$(Argument,"%3",Path)
		If InStr(Argument,"%1") Then Argument = Replace$(Argument,"%1","""" & File & """")
		If InStr(Argument,"%2") Then Argument = Replace$(Argument,"%2","*")
		If InStr(Argument,"%3") Then Argument = Replace$(Argument,"%3","""" & Path & """")
	End If
	If ExePath <> "" Then
		If Dir$(ExePath) <> "" Then
			If WshShell.Run("""" & ExePath & """ " & Argument,0,True) = 0 Then
				ExtractFile = IIf(Dir$(Path & updateMainFile) <> "",1,-1)
			Else
				ExtractFile = -3
			End If
		Else
			ExtractFile = -2
		End If
	Else
		ExtractFile = -2
	End If
	Set WshShell = Nothing
End Function


'��ȡ���صĳ���汾��
Private Function GetWebVersion(ByVal File As String,ByVal CheckStr As String) As String
	Dim n As Long,Temp As String,FN As Long
	On Error GoTo ExitFunction
	FN = FreeFile
	Open File For Input As #FN
	Do While Not EOF(FN)
		Line Input #FN,Temp
		n = InStr(Temp,CheckStr)
		If n > 0 Then
			GetWebVersion = Mid$(Temp,n + Len(CheckStr) + 1)
			GetWebVersion = Left$(GetWebVersion,Len(GetWebVersion) - 1)
			Exit Do
		End If
	Loop
	ExitFunction:
	On Error Resume Next
	Close #FN
End Function


'��װ�°汾
Private Function SetupNewVersion(ByVal FromPath As String,ByVal TargetDir As String) As Boolean
	Dim i As Long
	On Error GoTo ExitFunction
	If Right$(FromPath,1) <> "\" Then FromPath = FromPath & "\"

	'����Ƿ������Ӧ��Ŀ¼
	If Dir$(FromPath & "*.lng") <> "" Or Dir$(FromPath & "*.ini") <> "" Or Dir$(FromPath & "*.dat") <> "" Then
		If Dir$(TargetDir & "\Data\" & "*.*") = "" Then MkDir TargetDir & "\Data\"
	End If
	If Dir$(FromPath & "*.txt") <> "" Or Dir$(FromPath & "*.doc") <> "" Or Dir$(FromPath & "*.pdf") <> "" Then
		If Dir$(TargetDir & "\Doc\" & "*.*") = "" Then MkDir TargetDir & "\Doc\"
	End If
	If Dir$(FromPath & "mod*.bas") <> "" Or Dir$(FromPath & "*.cls") <> "" Then
		If Dir$(TargetDir & "\Module\" & "*.*") = "" Then MkDir TargetDir & "\Module\"
	End If
	If Dir$(FromPath & "*.chm") <> "" Or Dir$(FromPath & "*.hlp") <> "" Then
		If Dir$(TargetDir & "\Help\" & "*.*") = "" Then MkDir TargetDir & "\Help\"
	End If

	'��ȡ��ǰ�ļ��У����������ļ��У��е��°汾�ļ��б�
	ReDim FileList(0) As FILE_LIST
	If CheckArray(GetFiles(FileList,FromPath,"","*.*")) = False Then
		If CheckArray(getSubFiles(FileList,FromPath,"","*.*")) = False Then Exit Function
		SetupNewVersion = True
	End If
	Do
		'���Ƶ�ǰ�ļ���Ŀ���ļ���
		For i = 0 To UBound(FileList)
			Select Case LCase$(Mid$(FileList(i).sName,InStrRev(FileList(i).sName,".") + 1))
			Case "bas"
				If FileList(i).sName Like "mod*.bas" = False Then
					FileCopy FileList(i).FilePath,TargetDir & "\" & FileList(i).sName
				Else
					FileCopy FileList(i).FilePath,TargetDir & "\Module\" & FileList(i).sName
				End If
			Case "lng", "dat"
				FileCopy FileList(i).FilePath,TargetDir & "\Data\" & FileList(i).sName
			Case "obm", "cls"
				FileCopy FileList(i).FilePath,TargetDir & "\Module\" & FileList(i).sName
			Case "txt","doc","pdf"
				FileCopy FileList(i).FilePath,TargetDir & "\Doc\" & FileList(i).sName
				If Dir$(TargetDir & "\Data\" & FileList(i).sName) <> "" Then
					Kill TargetDir & "\Data\" & FileList(i).sName
				End If
			Case "ini"
				If Dir$(TargetDir & "\Data\" & FileList(i).sName) <> "" Then
					If FileDateTime(FileList(i).FilePath) > FileDateTime(TargetDir & "\Data\" & FileList(i).sName) Then
						FileCopy FileList(i).FilePath,TargetDir & "\Data\" & FileList(i).sName
					End If
				Else
					FileCopy FileList(i).FilePath,TargetDir & "\Data\" & FileList(i).sName
				End If
			Case "chm", "hlp"
				FileCopy FileList(i).FilePath,TargetDir & "\Help\" & FileList(i).sName
				If Dir$(TargetDir & "\Data\" & FileList(i).sName) <> "" Then
					Kill TargetDir & "\Data\" & FileList(i).sName
				End If
			Case Is <> "rar"
				FileCopy FileList(i).FilePath,TargetDir & "\Module\" & FileList(i).sName
			End Select
		Next i
		If SetupNewVersion = True Then Exit Do
		SetupNewVersion = True
		'��ȡ���ļ����е��°汾�ļ��б�
		ReDim FileList(0) As FILE_LIST
		If CheckArray(getSubFiles(FileList,FromPath,"","*.*")) = False Then Exit Do
	Loop
	SetupNewVersion = True
	ExitFunction:
End Function
