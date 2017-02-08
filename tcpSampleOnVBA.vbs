'===========================================================
'TCP/IP�ŒʐM����T���v��(VBS/VBA)
'�ʐM�I�u�W�F�N�g����
'===========================================================
'�y���ӎ����z
'�@[regsvr32.exe NONCOMSCK.OCX]���K�v
'�@����VBS�T���v����64bit(x64)��VBS�ł�CreateObject�G���[�ɂȂ�܂��B
'�@32bit(x86)�ł�WSH(C:\Windows\SysWow64\cscript.exe)���g�p���Ă��������B
'===========================================================

'===========================================================
'���C������
'===========================================================
'----------
' ���������^�ݒ�
'----------
Dim ipAddess, portNo
Dim i
Dim commandStr(10), k
Dim commandFileName

ipAddess        = "127.0.0.1"
portNo          = 4000
commandFileName = "testUtf8.txt"

'----------
' ����
'----------
Set Winsock1 = CreateObject("NonComSck.Winsock")
i = 0

'�R�}���h�p�������z��ɃZ�b�g
Call readCommandFile(commandFileName)

'�R�}���h�p������𑗐M���郋�[�v
Do While True
	Call startConnection
	Call transData
Loop

WSCript.Quit

'===========================================================
'���C������
'===========================================================
Sub transData()
	
	Dim wText
	Dim wSendStr
	
	WScript.Echo "---transData-----"
	'----------
	' �f�[�^���M(�������Byte�z��ɕϊ����đ��M)�^End�̏ꍇ�͋����I��
	'----------
	
	'�R�}���h�p��������P�s���������o
	wText = speechText(i)
	WScript.Echo i & ":" & wText
	
	i = i + 1

	'�R�}���h�p������i���s�R�[�h<LF>���j��UTF-8�ɕϊ�����
	wSendStr = encodeStr(wText & vbLf, "UTF-8")
	
	'�T�[�o���փR�}���h�p������𑗐M
	Winsock1.SEndData wSendStr

	'----------
	' �f�[�^��M�i�T�[�o����̎�M�������m�F�j
	'----------
	Winsock1.Start_EventForScript()
	Do
		WScript.Sleep(500)
		Evt = Winsock1.GetEventParameters()
		If Ubound(Evt) >= 0 Then
		
			' Evt(0) : �C�x���g��
			If Evt(0) = "DataArrival" Then
				' Evt(9) : ��M�f�[�^��Byte�z��
				' Byte�z��𕶎���ɕϊ�
				WScript.Echo Winsock1.ByteArrayToStr(Evt(9))
				Exit Do
				
			End If
			
		End If
	Loop
	Winsock1.End_EventForScript()
	
	'�P�`���̑���M���m�F������ؒf�iTCP/IP�ʐM�̐���j
	Call disConnection()
	
	'�I���R�}���h���ݒ肳��Ă�����A�v���O�����I��
	IF wText = "End" THEN
		Call disConnection()
		WSCript.Quit
	End IF
	
End Sub

'===========================================================
' TCP�ʐM�J�n
'===========================================================
Sub startConnection()
	WScript.Echo "---startConnection-----"
	'----------
	' TCP/IP�ڑ�
	'----------
	Winsock1.Connect ipAddess, portNo

	'----------
	' TCP/IP�ڑ��҂�
	'----------
	Do While Winsock1.State = 6
	    WScript.Sleep(500)
	Loop
End Sub

'===========================================================
' TCP�ʐM�ؒf
'===========================================================
Sub disConnection()
	WScript.Echo "---disconnection-----"
	
	Winsock1.Close2
	
End Sub

'===========================================================
' �R�}���h�p������̒��o�i�P�����j
'===========================================================
Function speechText(Byval pSpeechNo) 
	Dim wRetText
	
	wRetText = commandStr(pSpeechNo)
	speechText = wRetText
	
End Function

'===========================================================
' �R�}���h�p�t�@�C���iUTF-8�̃e�L�X�g�t�@�C���j��Ǎ���
'===========================================================
Sub readCommandFile(Byval pFileName)
	Dim objStream

	'----------
	' �t�@�C����Ǎ���
	'----------
	Set objStream = CreateObject("ADODB.Stream")
	
	objStream.Type = 2							' 1�F�o�C�i��, 2�F�e�L�X�g
	objStream.Charset = "UTF-8"					' �����R�[�h�w��
	objStream.Open
	
	objStream.LoadFromFile pFileName
	
	'----------
	' �Ǎ��݃t�@�C������1�s���R�}���h�p������i�z��j�ɏ�����
	'----------
	k = 0
	Do Until objStream.EOS
		commandStr(k) = objStream.ReadText(-2)	' -1�F�S�s�Ǎ���, -2�F��s�Ǎ���
		'WScript.Echo commandStr(k)
		
		k = k + 1
	Loop
	
	'----------
	'�I������
	'----------
	objStream.Close
	Set objStream = Nothing
	
End Sub

'===========================================================
' �����R�[�h�ϊ�
'===========================================================
Function encodeStr(Byval pStrUni, Byval pCharSet) 

	Set objStream = CreateObject("ADODB.Stream")
	
	'----------
	'�w�肳�ꂽ�������Stream�ɏ�����
	'----------
	objStream.Open
	objStream.Type = 2					' 1�F�o�C�i��, 2�F�e�L�X�g
	objStream.Charset = pCharSet
	objStream.WriteText pStrUni 
	objStream.Position = 0

	'----------
	'�����R�[�h�ϊ�����Strem����ǂݏo��
	'----------
	'BOM�����镶���R�[�h�̏ꍇ�́A�ŏ���BOM�����X�L�b�v
	objStream.Type = 1					' 1�F�o�C�i��, 2�F�e�L�X�g
	Select Case UCase(pCharSet)
		Case "UNICODE", "UTF-16"
			objStream.Position = 2
			
		Case "UTF-8"
			objStream.Position = 3
			
	End Select
	
	encodeStr = objStream.Read()
	
	'----------
	'�I������
	'----------
	objStream.Close
	Set objStream = Nothing
	
End Function
