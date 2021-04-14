Module Module1
    Dim uni As String = "", UTF8str As String = ""

    Sub 查詢中國哲學書電子化計劃字典(ByVal x As String)
        'Clipboard.SetText(x)
        Shell(Replace(GetDefaultBrowserEXE, """%1", "http://ctext.org/dictionary.pl?if=gb&char=" & UTF8str))
    End Sub
    Sub 查詢古今文字集成(ByVal x As String)
        'Clipboard.SetText(x)
        Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.ccamc.co/cjkv.php?cjkv=" & UTF8str))
    End Sub
    Sub 查詢漢語多功能字庫(ByVal x As String)
        'Clipboard.SetText(x)
        Shell(Replace(GetDefaultBrowserEXE, """%1", "http://humanum.arts.cuhk.edu.hk/Lexis/lexi-mf/search.php?word=" & UTF8str))
    End Sub

    Sub 查詢國語辭典(ByVal x As String)
        'If Len(x) > 1 Then
        '    Shell(Replace(GetDefaultBrowserEXE, """%1", "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?idx=dict.idx&cond=" & 查詢字串轉換_國語會碼(x) & "&pieceLen=50&fld=1&cat=&imgFont=1"))
        'Else
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?idx=dict.idx&cond=^" & 查詢字串轉換_國語會碼(x) & "$&pieceLen=50&fld=1&cat=&imgFont=1"))
        Shell(Replace(GetDefaultBrowserEXE, """%1", "http://dict.revised.moe.edu.tw/cbdic/search.htm"))
        Shell(Replace(GetDefaultBrowserEXE, """%1", "https://dict.variants.moe.edu.tw/variants/rbt/query_by_standard_tiles.rbt?command=clear"))
        'End If
    End Sub
    Sub 查詢漢典(ByVal x As String) 'http://www.zdic.net/zd/zi/ZdicE4ZdicB8Zdic8B.htm
        Dim u8 As System.Text.Encoding = System.Text.Encoding.UTF8
        Dim bytes As Byte() = u8.GetBytes(x)
        Const HDurl As String = "http://www.zdic.net/zd/zi/"
        Shell(Replace(GetDefaultBrowserEXE, """%1", HDurl & PrintHexBytes_漢典(bytes) & ".htm"))
    End Sub
    Sub 查詢小學堂(ByVal x As String)
        UTF8str = 查詢字串轉換_網路碼(x)
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://xiaoxue.iis.sinica.edu.tw/char?fontcode=" & 查詢字串轉換_小學堂碼(x)))
        查詢字串轉換_小學堂碼(x) '國學大師要參照，要用其中的uni變數
        Shell(Replace(GetDefaultBrowserEXE, """%1", "http://xiaoxue.iis.sinica.edu.tw/?char=" & UTF8str))

    End Sub
    Sub 查詢國學大師汉语字典(ByVal x As String) 'http://www.guoxuedashi.net/zidian/93F5.html
        On Error GoTo EH
        'Dim u8 As System.Text.Encoding = System.Text.Encoding.Unicode
        'Dim bytes As Byte() = u8.GetBytes(x)
        'Const HDurl As String = "http://www.guoxuedashi.net/zidian/"
        'Shell(Replace(GetDefaultBrowserEXE, """%1", HDurl & uni & ".html"))
        ''Shell(Replace(GetDefaultBrowserEXE, """%1", HDurl & PrintHexBytes_Unicode(bytes) & ".html"))

        Const HDurl As String = "http://www.guoxuedashi.net/zidian/"
        'Dim u8 As System.Text.Encoding = System.Text.Encoding.Unicode
        Dim pu As String = "", Tpu As String = ""
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.guoxuedashi.net/so.php?sokeytm=" & x & "&ka=100"))
        Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.guoxuedashi.net/zidian/so.php?sokeyzi=" & x & "&kz=1")) 'http://www.guoxuedashi.net/help/setie.php
        Exit Sub


        If Len(x) > 1 Then
            'u8 = System.Text.Encoding.UTF32
            Dim dbe As New DAO.DBEngine, db As DAO.Database = dbe.OpenDatabase("DB1.mdb")
            Dim qf As DAO.QueryDef = db.QueryDefs("qr"), p As DAO.Parameter = qf.Parameters("q")
            p.Value = x
            Dim rst As DAO.Recordset = qf.OpenRecordset
            If rst.RecordCount > 0 Then
                pu = rst.Fields("pu").Value
            Else
                MsgBox("Check!!", vbExclamation)
                'Dim bytes As Byte() = u8.GetBytes(x)
                'pu = PrintHexBytes_Unicode(bytes)
                pu = uni
            End If

        Else
1:          'Dim bytes As Byte() = u8.GetBytes(x)
            pu = uni 'PrintHexBytes_Unicode(bytes)
        End If
        Tpu = HDurl & pu
2:      Shell(Replace(GetDefaultBrowserEXE, """%1", Tpu & ".html"))
        Exit Sub
EH:
        Select Case Err.Number
            Case 3024 '找不到檔案
                If Dir("w:\!!for hpr\速檢網路字辭典\速檢網路字辭典\bin\Debug\DB1.mdb") <> "" Then
                    Resume
                ElseIf Dir("C:\Program Files\孫守真\速檢網路字辭典\DB1.mdb") <> "" Then
                    Resume
                ElseIf Dir("C:\Program Files (x86)\孫守真\速檢網路字辭典\DB1.mdb") <> "" Then
                    Resume
                Else
                    GoTo ET
                End If

            Case Else
                Debug.Print(Err.Number & Err.Description)
                MsgBox(Err.Number & Err.Description)
                'Stop : Resume

        End Select
ET:
        If Len(x) > 1 Then
            Tpu = HDurl & uni

            'u8 = System.Text.Encoding.UTF32
            'Dim bytes As Byte() = u8.GetBytes(x)
            'pu = PrintHexBytes_Unicode(bytes)
            'If Left(pu, 2) = "02" Then
            '    If Len(pu) = 6 Then
            '        Tpu = HDurl & Mid(pu, 2)
            '    Else
            '        Tpu = HDurl & Replace(pu, "02", "20", 1, 1)
            '    End If
            '    'ElseIf Left(pu, 1) = 2 And Len(pu) = 4 Then

            'Else
            '    Tpu = HDurl & pu
            'End If
            GoTo 2
        Else
            GoTo 1
        End If


    End Sub
    Function 查詢字串轉換_小學堂碼(ByVal w As String) '即十六位元Unicode碼
        Dim u8 As System.Text.Encoding = System.Text.Encoding.Unicode
        Dim bytes As Byte() = u8.GetBytes(w)
        uni = PrintHexBytes_Unicode(bytes)
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
