Attribute VB_Name = "Module1"
Option Explicit

Sub 產能標記()
Dim i, rowcnt As Integer
Dim target As Integer
target = CInt(InputBox("請輸入標記上限值(0-1000)"))
Dim rangestr As String
rowcnt = Cells(Rows.Count, 1).End(xlUp).Row
rangestr = "b3:b" & rowcnt
MsgBox "目前運算範圍" & rangestr
Range(rangestr).Interior.Color = xlNone
For i = 3 To rowcnt
    If Cells(i, "B") > target Then
        Cells(i, "B").Interior.Color = vbYellow
    End If
Next
Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

End Sub
