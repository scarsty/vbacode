rc = ActiveSheet.UsedRange.Rows.Count
cc = ActiveSheet.UsedRange.Columns.Count
'MsgBox cc
'清除无用
Dim r, r1, k As Integer
r = 0
r1 = 999999
For i = 2 To rc
    If Not (Cells(i, 1).Value <= 1 And Cells(i, 1).Value >= 0.5 And Cells(i, 1).Value < Cells(i + 1, 1).Value) Then
        For j = 1 To cc
            Cells(i, j).Value = ""
        Next j
    Else
        r1 = WorksheetFunction.Min(i, r1)
        r = r + 1
    End If
Next i

'放到上面
k = 5
For i = k To r + k - 1
    For j = 1 To cc
        Cells(i, j).Value = Cells(i - k + r1, j).Value
        Cells(i - k + r1, j).Value = ""
    Next j
Next i

'将公式写入表格
'y=kx+b, k=y(1)-y(0), b=y(0)
r = r + k - 1
Cells(2, 1).Value = "k"
Cells(3, 1).Value = "b"
For j = 2 To cc
cd = Chr(j + 64)
'MsgBox cd
Cells(3, j).Value = "=TREND(" + cd + "5:" + cd + CStr(r) + ",A5:A" + CStr(r) + ",0)"
Cells(2, j).Value = "=TREND(" + cd + "5:" + cd + CStr(r) + ",A5:A" + CStr(r) + ",1)-" + cd + "3"
Next j


