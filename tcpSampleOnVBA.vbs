'===========================================================
'TCP/IPで通信するサンプル(VBS/VBA)
'通信オブジェクト生成
'===========================================================
'【注意事項】
'　[regsvr32.exe NONCOMSCK.OCX]が必要
'　このVBSサンプルは64bit(x64)版VBSではCreateObjectエラーになります。
'　32bit(x86)版のWSH(C:\Windows\SysWow64\cscript.exe)を使用してください。
'===========================================================

'----------
' 初期処理
'----------
Dim i

Set Winsock1 = CreateObject("NonComSck.Winsock")
i = 0

do while true
	call startConnection
	call main
loop

'===========================================================
'メイン処理
'===========================================================
SUB main()
	WScript.Echo "---main-----"
	'----------
	' データ送信(文字列をByte配列に変換して送信)／ENDの場合は強制終了
	'----------
	''''''wText = InputBox("送信テキストを入力","入力","red")
	wText = speechText(i)
	WScript.Echo i & ":" & wText
	
	''''''''''''''''''wText = encodeUTF8(wText)
	i = i + 1
	
	'Winsock1.SendData Winsock1.StrToByteArray(wText & vbLf)
	Winsock1.SendData wText & vbLf

	'----------
	' データ受信
	'----------
	Winsock1.Start_EventForScript()
	Do
	    WScript.Sleep(500)
	    Evt = Winsock1.GetEventParameters()
	    If Ubound(Evt) >= 0 Then
	        ' Evt(0) : イベント名
	        If Evt(0) = "DataArrival" Then
	            ' Evt(9) : 受信データのByte配列
	            ' Byte配列を文字列に変換
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
' TCP通信開始
'===========================================================
SUB startConnection()
	WScript.Echo "---startConnection-----"
	'----------
	' TCP/IP接続
	'----------
	Winsock1.Connect "172.16.168.46", 5001

	'----------
	' TCP/IP接続待ち
	'----------
	Do While Winsock1.State = 6
	    WScript.Sleep(500)
	Loop
End SUB

'===========================================================
' TCP通信切断
'===========================================================
SUB disConnection()
	WScript.Echo "---disConnection-----"
	'----------
	' TCP/IP切断
	'----------
	Winsock1.Close2

	WScript.Echo "終了"
end SUB

'===========================================================
' 読み取りテキストの抽出
'===========================================================
FUNCTION speechText(byval pSpeechNo) 
	Dim wRetText
	
	select case pSpeechNo
	case 0
		'wRetText = "ko re ka ra ka i sha se tu me i ka i wo ha ji me ma su"
		wRetText = "これから会社説明会をはじめます"
	case 1
		'wRetText = "so no 1"
		wRetText = "そのいち"
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
' 文字列をUTF-8でエンコードする
'===========================================================
FUNCTION encodeUTF8(byval mytext) 
    Dim mystream
    Dim mybinary, mynumber
    
    Set mystream = CreateObject("ADODB.Stream")
    
    With mystream
        .Open
        .Type = 2				'adTypeText
        .Charset = "UTF-8"
        .LineSeparator = 10		'改行コード：LF
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
