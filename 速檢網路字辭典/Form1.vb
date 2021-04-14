Public Class Form1
    Dim X As String = Clipboard.GetText
    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If X <> "" Then
            If Len(X) = 1 Then
                查詢國語辭典(X)
            Else
                Shell(Replace(GetDefaultBrowserEXE, """%1", "http://dict2.variants.moe.edu.tw/variants/rbt/query_by_standard_tiles.rbt?command=clear"))
            End If
            '查詢小學堂(X)
            Shell(Replace(GetDefaultBrowserEXE, """%1", "http://xiaoxue.iis.sinica.edu.tw/?char=" & X))
            '查詢中國哲學書電子化計劃字典(X)
            '查詢漢典(X)
            '查詢國學大師汉语字典(X)
            Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.guoxuedashi.com/zidian/so.php?sokeyzi=" & X & "&kz=1"))
            '查詢古今文字集成(X)
            Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.ccamc.co/cjkv.php?cjkv=" & X))
            End
        End If
    End Sub
End Class
