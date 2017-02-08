# TcpSampleOnVBA
TCP通信を行うためのVBAサンプル<BR>
Sota君がTCPサーバとして起動している場合に、TCPクライアントとしてメッセージを送付するツールです。<BR>
<BR>
《セットアップ方法》
* RequiredModuleForTcpConnection.zip を解凍
* VBAでTCP通信を行うサンプル.pdf に記載の手順で、NonComSck.ocx、MSWINSCK.OCX をレジストリに登録
<BR>

《実行方法》
* コマンドテキスト（testUtf8.txt）を用意する。
* tcpSampleOnVBA.vbs を起動する。

※このVBSサンプルは64bit(x64)版VBSではCreateObjectエラーになります。32bit(x86)版のWSH(C:\Windows\SysWow64\cscript.exe)を使用してください。
