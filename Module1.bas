Attribute VB_Name = "Module1"
Option Explicit

Sub ����аO()
Dim i, rowcnt As Integer
Dim target As Integer
target = CInt(InputBox("�п�J�аO�W����(0-1000)"))
Dim rangestr As String
rowcnt = Cells(Rows.Count, 1).End(xlUp).Row
rangestr = "b3:b" & rowcnt
MsgBox "�ثe�B��d��" & rangestr
Range(rangestr).Interior.Color = xlNone
For i = 3 To rowcnt
    If Cells(i, "B") > target Then
        Cells(i, "B").Interior.Color = vbYellow
    End If
Next
Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

End Sub
