'===========================================================
'TCP/IP�ŒʐM����T���v��(VBS/VBA)
'�ʐM�I�u�W�F�N�g����
'===========================================================
'�y���ӎ����z
'�@[regsvr32.exe NONCOMSCK.OCX]���K�v
'�@����VBS�T���v����64bit(x64)��VBS�ł�CreateObject�G���[�ɂȂ�܂��B
'�@32bit(x86)�ł�WSH(C:\Windows\SysWow64\cscript.exe)���g�p���Ă��������B
'===========================================================

'----------
' ��������
'----------
Dim i

Set Winsock1 = CreateObject("NonComSck.Winsock")
i = 0

do while true
	call startConnection
	call main
loop

'===========================================================
'���C������
'===========================================================
SUB main()
	WScript.Echo "---main-----"
	'----------
	' �f�[�^���M(�������Byte�z��ɕϊ����đ��M)�^END�̏ꍇ�͋����I��
	'----------
	''''''wText = InputBox("���M�e�L�X�g�����","����","red")
	wText = speechText(i)
	WScript.Echo i & ":" & wText
	
	''''''''''''''''''wText = encodeUTF8(wText)
	i = i + 1
	
	'Winsock1.SendData Winsock1.StrToByteArray(wText & vbLf)
	Winsock1.SendData wText & vbLf

	'----------
	' �f�[�^��M
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
	            'MsgBox Winsock1.ByteArrayToStr(Evt(9))
	            WScript.Echo Winsock1.ByteArrayToStr(Evt(9))
	            Exit Do
	        End If
	    End If
	Loop
	Winsock1.End_EventForScript()
	
	call disConnection()


	IF wText = "end" THEN
		call disConnection()
		WSCript.Quit
	END IF
	
END SUB

'===========================================================
' TCP�ʐM�J�n
'===========================================================
SUB startConnection()
	WScript.Echo "---startConnection-----"
	'----------
	' TCP/IP�ڑ�
	'----------
	Winsock1.Connect "172.16.168.46", 5001

	'----------
	' TCP/IP�ڑ��҂�
	'----------
	Do While Winsock1.State = 6
	    WScript.Sleep(500)
	Loop
End SUB

'===========================================================
' TCP�ʐM�ؒf
'===========================================================
SUB disConnection()
	WScript.Echo "---disConnection-----"
	'----------
	' TCP/IP�ؒf
	'----------
	Winsock1.Close2

	WScript.Echo "�I��"
end SUB

'===========================================================
' �ǂݎ��e�L�X�g�̒��o
'===========================================================
FUNCTION speechText(byval pSpeechNo) 
	Dim wRetText
	
	select case pSpeechNo
	case 0
		'wRetText = "ko re ka ra ka i sha se tu me i ka i wo ha ji me ma su"
		wRetText = "���ꂩ���А�������͂��߂܂�"
	case 1
		'wRetText = "so no 1"
		wRetText = "���̂���"
	case 2
		wRetText = "are ya kore ya"
	case 3
		wRetText = "so no 2"
	case 4
		wRetText = "dou tara kou tara"
	case else
		wRetText = "owari"
	end select
	
	speechText = wRetText
	
END FUNCTION

'===========================================================
' �������UTF-8�ŃG���R�[�h����
'===========================================================
FUNCTION encodeUTF8(byval mytext) 
    Dim mystream
    Dim mybinary, mynumber
    
    Set mystream = CreateObject("ADODB.Stream")
    
    With mystream
        .Open
        .Type = 2				'adTypeText
        .Charset = "UTF-8"
        .LineSeparator = 10		'���s�R�[�h�FLF
        .WriteText mytext
        .Position = 0
        .Type = 1				'adTypeBinary
        .Position = 3
        mybinary = .Read
        .Close
    End With
    
    WScript.Echo mybinary
    
    
    For Each mynumber In mybinary
        encodeUTF8 = encodeUTF8 & "%" & Hex(mynumber)
    Next
END FUNCTION
