Module LookUP
    Dim uni As String = "", UTF8str As String = ""
    Dim X As String = Clipboard.GetText
    Sub LookUPs()
        If X <> "" Then
            'If Len(X) = 1 Then
            '    查詢國語辭典(X)
            'Else
            '    Shell(Replace(GetDefaultBrowserEXE, """%1", "https://dict.variants.moe.edu.tw/variants/rbt/query_by_standard_tiles.rbt?command=clear"))
            'End If
            查詢小學堂(X)
            '查詢中國哲學書電子化計劃字典(X)
            '查詢漢典(X)
            查詢國學大師汉语字典(X)
            查詢古今文字集成(X)
            查詢漢語多功能字庫(X)
            'Shell(Replace(GetDefaultBrowserEXE, """%1",
            '    "https://dict.variants.moe.edu.tw/variants/rbt/query_by_standard_tiles.rbt?command=clear"))
            '異體字字典
            Process.Start(BrowserApp,
              "https://dict.variants.moe.edu.tw/variants/rbt/query_by_standard_tiles.rbt?command=clear")
            End
        End If
    End Sub
    Sub 查詢中國哲學書電子化計劃字典(ByVal x As String)
        'Clipboard.SetText(x)
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://ctext.org/dictionary.pl?if=gb&char=" & UTF8str))
        BrowserOps.openUrl(BrowserApp,
            "http://ctext.org/dictionary.pl?if=gb&char=" & UTF8str)
    End Sub
    Sub 查詢古今文字集成(ByVal x As String)
        'Clipboard.SetText(x)
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.ccamc.co/cjkv.php?cjkv=" & UTF8str))
        BrowserOps.openUrl(BrowserApp,
            "http://www.ccamc.co/cjkv.php?cjkv=" & UTF8str)
    End Sub
    Sub 查詢漢語多功能字庫(ByVal x As String)
        'Clipboard.SetText(x)
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://humanum.arts.cuhk.edu.hk/Lexis/lexi-mf/search.php?word=" & UTF8str))
        BrowserOps.openUrl(BrowserApp,
          "http://humanum.arts.cuhk.edu.hk/Lexis/lexi-mf/search.php?word=" & UTF8str)
    End Sub

    Sub 查詢國語辭典(ByVal x As String)
        'If Len(x) > 1 Then
        '    Shell(Replace(GetDefaultBrowserEXE, """%1", "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?idx=dict.idx&cond=" & 查詢字串轉換_國語會碼(x) & "&pieceLen=50&fld=1&cat=&imgFont=1"))
        'Else
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?idx=dict.idx&cond=^" & 查詢字串轉換_國語會碼(x) & "$&pieceLen=50&fld=1&cat=&imgFont=1"))
        'Shell(Replace(GetDefaultBrowserEXE, """%1",
        '    "http://dict.revised.moe.edu.tw/cbdic/search.htm"))
        'Shell(Replace(GetDefaultBrowserEXE, """%1",
        '    "https://dict.variants.moe.edu.tw/variants/rbt/query_by_standard_tiles.rbt?command=clear"))
        BrowserOps.openUrl(BrowserApp,
                "http://dict.revised.moe.edu.tw/cbdic/search.htm")
        BrowserOps.openUrl(BrowserApp,
           "https://dict.variants.moe.edu.tw/variants/rbt/query_by_standard_tiles.rbt?command=clear")
        'End If
    End Sub
    Sub 查詢漢典(ByVal x As String) 'http://www.zdic.net/zd/zi/ZdicE4ZdicB8Zdic8B.htm
        Dim u8 As System.Text.Encoding = System.Text.Encoding.UTF8
        Dim bytes As Byte() = u8.GetBytes(x)
        Const HDurl As String = "http://www.zdic.net/zd/zi/"
        'Shell(Replace(GetDefaultBrowserEXE, """%1",
        'HDurl & PrintHexBytes_漢典(bytes) & ".htm"))
        BrowserOps.openUrl(BrowserApp,
            HDurl & PrintHexBytes_漢典(bytes) & ".htm")
    End Sub
    Sub 查詢小學堂(ByVal x As String)
        UTF8str = 查詢字串轉換_網路碼(x)
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://xiaoxue.iis.sinica.edu.tw/char?fontcode=" & 查詢字串轉換_小學堂碼(x)))
        查詢字串轉換_小學堂碼(x) '國學大師要參照，要用其中的uni變數

        BrowserOps.openUrl(BrowserApp,
            "http://xiaoxue.iis.sinica.edu.tw/?char=" & UTF8str)
        'Shell(Replace(GetDefaultBrowserEXE, """%1",
        '   "http://xiaoxue.iis.sinica.edu.tw/?char=" & UTF8str))

    End Sub
    Sub 查詢國學大師汉语字典(ByVal x As String) 'http://www.guoxuedashi.net/zidian/93F5.html
        'On Error GoTo EH
        'Dim u8 As System.Text.Encoding = System.Text.Encoding.Unicode
        'Dim bytes As Byte() = u8.GetBytes(x)
        'Const HDurl As String = "http://www.guoxuedashi.net/zidian/"
        'Shell(Replace(GetDefaultBrowserEXE, """%1", HDurl & uni & ".html"))
        ''Shell(Replace(GetDefaultBrowserEXE, """%1", HDurl & PrintHexBytes_Unicode(bytes) & ".html"))

        Const HDurl As String = "http://www.guoxuedashi.net/zidian/"
        'Dim u8 As System.Text.Encoding = System.Text.Encoding.Unicode
        Dim pu As String = "", Tpu As String = ""
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.guoxuedashi.net/so.php?sokeytm=" & x & "&ka=100"))
        'Shell(Replace(GetDefaultBrowserEXE, """%1",
        '"http://www.guoxuedashi.net/zidian/so.php?sokeyzi=" & x & "&kz=1")) 'http://www.guoxuedashi.net/help/setie.php
        BrowserOps.openUrl(BrowserApp,
            HDurl + "so.php?sokeyzi=" & x & "&kz=1")
        Exit Sub


        '        If Len(x) > 1 Then
        '            'u8 = System.Text.Encoding.UTF32
        '            Dim dbe As New DAO.DBEngine, db As DAO.Database = dbe.OpenDatabase("DB1.mdb")
        '            Dim qf As DAO.QueryDef = db.QueryDefs("qr"), p As DAO.Parameter = qf.Parameters("q")
        '            p.Value = x
        '            Dim rst As DAO.Recordset = qf.OpenRecordset
        '            If rst.RecordCount > 0 Then
        '                pu = rst.Fields("pu").Value
        '            Else
        '                MsgBox("Check!!", vbExclamation)
        '                'Dim bytes As Byte() = u8.GetBytes(x)
        '                'pu = PrintHexBytes_Unicode(bytes)
        '                pu = uni
        '            End If

        '        Else
        '1:          'Dim bytes As Byte() = u8.GetBytes(x)
        '            pu = uni 'PrintHexBytes_Unicode(bytes)
        '        End If
        '        Tpu = HDurl & pu
        '2:      Shell(Replace(GetDefaultBrowserEXE, """%1", Tpu & ".html"))
        '        Exit Sub
        'EH:
        '        Select Case Err.Number
        '            Case 3024 '找不到檔案
        '                If Dir("w:\!!for hpr\速檢網路字辭典\速檢網路字辭典\bin\Debug\DB1.mdb") <> "" Then
        '                    Resume
        '                ElseIf Dir("C:\Program Files\孫守真\速檢網路字辭典\DB1.mdb") <> "" Then
        '                    Resume
        '                ElseIf Dir("C:\Program Files (x86)\孫守真\速檢網路字辭典\DB1.mdb") <> "" Then
        '                    Resume
        '                Else
        '                    GoTo ET
        '                End If

        '            Case Else
        '                Debug.Print(Err.Number & Err.Description)
        '                MsgBox(Err.Number & Err.Description)
        '                'Stop : Resume

        '        End Select
        'ET:
        '        If Len(x) > 1 Then
        '            Tpu = HDurl & uni

        '            'u8 = System.Text.Encoding.UTF32
        '            'Dim bytes As Byte() = u8.GetBytes(x)
        '            'pu = PrintHexBytes_Unicode(bytes)
        '            'If Left(pu, 2) = "02" Then
        '            '    If Len(pu) = 6 Then
        '            '        Tpu = HDurl & Mid(pu, 2)
        '            '    Else
        '            '        Tpu = HDurl & Replace(pu, "02", "20", 1, 1)
        '            '    End If
        '            '    'ElseIf Left(pu, 1) = 2 And Len(pu) = 4 Then

        '            'Else
        '            '    Tpu = HDurl & pu
        '            'End If
        '            GoTo 2
        '        Else
        '            GoTo 1
        '        End If


    End Sub
End Module
