Attribute VB_Name = "Integermultiple"
Sub Integermultiple()

'How-To-Use
MsgBox ("Ⅰ.整数倍を掛ける列を整数で入力(列勉号は左から1から始まる)" & vbCrLf & _
vbTab & "例 A列→1 , D列→4 " & vbCrLf & "Ⅱ.何倍に掛けるのかを入力" & vbCrLf & _
"Ⅲ.コードは終了する " & vbCrLf & "※尚、行は2から100(任意)の行だけで実行される。")

Dim a As Integer

Dim c As Integer

a = InputBox("倍数を掛ける列を入力(半角数字で)", "列番号の入力", 1, 100, 100)
'MsgBox a

c = InputBox("何倍にする?", "指定倍の数字", 1, 100, 100)
'MsgBox c

Dim d, b As Integer
b = 0
d = 0

For d = 2 To 100 '2にしたのは1行目に並び替えのタイトルが来ることを想定している為

    'Cells(d, a).Select ' 行が何処を指しているのか分からない時に、これを使うといい。

    Cells(d, a).Value = Cells(d, a).Value * c

    If (Cells(d, a).Value = 0) Then '0が表示されるのは鬱陶しい(必要に応じてコメント化)

        Cells(d, a).Value = ""

    End If

Next d

End Sub
