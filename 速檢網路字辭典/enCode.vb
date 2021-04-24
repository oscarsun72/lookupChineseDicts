Module enCode


    Private browserAppValue As String
    Public ReadOnly Property BrowserApp() As String
        Get
            browserAppValue = DefaultBrowser()
            'If Not browserAppValue.IndexOf("iexplore") > -1 Then
            If browserAppValue.IndexOf("iexplore") = -1 Then
                Return browserAppValue
            Else
                Dim bApp As New BrowserChrome
                Return bApp.ChromeAppFileName
            End If
        End Get
    End Property


    Function 查詢字串轉換_小學堂碼(ByVal w As String) '即十六位元Unicode碼
        Dim u8 As System.Text.Encoding = System.Text.Encoding.Unicode
        Dim bytes As Byte() = u8.GetBytes(w)
        Dim uni As String = PrintHexBytes_Unicode(bytes)
        查詢字串轉換_小學堂碼 = "0." & uni
    End Function
    Function 查詢字串轉換_國語會碼(ByVal w As String) 'Big5碼
        Dim u8 As System.Text.Encoding = System.Text.Encoding.GetEncoding("big5") 'https://msdn.microsoft.com/zh-tw/library/system.text.encoding(v=vs.110).aspx
        Dim bytes As Byte() = u8.GetBytes(w)
        查詢字串轉換_國語會碼 = PrintHexBytes(bytes)
    End Function
    Function 查詢字串轉換_網路碼(ByVal w As String)
        Dim u8 As System.Text.Encoding = System.Text.Encoding.UTF8
        Dim bytes As Byte() = u8.GetBytes(w)
        查詢字串轉換_網路碼 = PrintHexBytes(bytes)
    End Function
    Public Function PrintHexBytes(ByVal bytes() As Byte) As String ''https://msdn.microsoft.com/zh-tw/library/system.text.encoding.utf8(v=vs.110).aspx
        PrintHexBytes = ""
        If bytes Is Nothing OrElse bytes.Length = 0 Then
            'Console.WriteLine("<none>")
            MsgBox("<none>")
        Else
            Dim i As Integer
            For i = 0 To bytes.Length - 1
                PrintHexBytes &= "%" & Hex(bytes(i))
                'PrintHexBytes &= Hex(bytes(i))
                'Console.Write("{0:X2} ", bytes(i))
            Next i
            'Console.WriteLine()
        End If
    End Function 'PrintHexBytes 
    Public Function PrintHexBytes_漢典(ByVal bytes() As Byte) As String ''https://msdn.microsoft.com/zh-tw/library/system.text.encoding.utf8(v=vs.110).aspx
        PrintHexBytes_漢典 = ""
        If bytes Is Nothing OrElse bytes.Length = 0 Then
            'Console.WriteLine("<none>")
            MsgBox("<none>")
        Else
            Dim i As Integer
            For i = 0 To bytes.Length - 1
                PrintHexBytes_漢典 &= "Zdic" & Hex(bytes(i))
            Next i
        End If
    End Function
    Public Function PrintHexBytes_Unicode(ByVal bytes() As Byte) As String ''https://msdn.microsoft.com/zh-tw/library/system.text.encoding.utf8(v=vs.110).aspx
        PrintHexBytes_Unicode = ""
        If bytes Is Nothing OrElse bytes.Length = 0 Then
            'Console.WriteLine("<none>")
            MsgBox("<none>")
        Else
            Dim i As Integer
            For i = bytes.Length - 1 To 0 Step -1
                If bytes(i) < 16 Then
                    PrintHexBytes_Unicode &= "0" & Hex(bytes(i))
                Else
                    PrintHexBytes_Unicode &= Hex(bytes(i))
                End If
            Next i
        End If
    End Function

    Function GetDefaultBrowserEXE() '2010/10/18由http://chijanzen.net/wp/?p=156#comment-1303(取得預設瀏覽器(default web browser)的名稱? chijanzen 雜貨舖)而來.
        Dim objShell
        objShell = CreateObject("WScript.Shell")
        'HKEY_CLASSES_ROOT\HTTP\shell\open\ddeexec\Application
        '取得註冊表中的值
        GetDefaultBrowserEXE = objShell.RegRead _
                ("HKCR\http\shell\open\command\")
    End Function
End Module
