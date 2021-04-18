Public Class Form1
    Dim X As String = Clipboard.GetText
    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
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
End Class
