Attribute VB_Name = "Display"
Sub Display_day()
Dim i, j, jj, q, d, m, y As Integer
Dim dd As Date
Dim lastrow As Long

Set sht = Worksheets("data")
Set sht2 = Worksheets("Summary")
lastrow = sht.Cells(4, 1).End(xlDown).Row

q = 0
 
d = Cells(38, 4).Value
m = Cells(38, 3).Value
y = Cells(38, 2).Value
dd = DateSerial(y, m, d)

j = 40
jj = 61

Call Clear.Clear_day

sht.Unprotect "1234"
sht2.Unprotect "1234"

'pause calculation
Application.Calculation = xlManual

For i = 2 To lastrow
'Trade date
If sht.Cells(i, 6) = dd And j < 59 Then
    sht2.Cells(j, 1).Value = sht.Cells(i, 1).Value
    sht2.Cells(j, 2).Value = sht.Cells(i, 2).Value
    sht2.Cells(j, 4).Value = sht.Cells(i, 3).Value
    sht2.Cells(j, 6).Value = sht.Cells(i, 4).Value
    sht2.Cells(j, 8).Value = sht.Cells(i, 5).Value
    sht2.Cells(j, 10).Value = sht.Cells(i, 6).Value
    sht2.Cells(j, 12).Value = sht.Cells(i, 7).Value
    sht2.Cells(j, 14).Value = sht.Cells(i, 8).Value
    sht2.Cells(j, 15).Value = sht.Cells(i, 9).Value
    sht2.Range("Q" & j).Value = i
    j = j + 1
    q = 1
'Value Date
ElseIf sht.Cells(i, 8) = dd And jj < 80 Then
    sht2.Cells(jj, 1).Value = sht.Cells(i, 1).Value
    sht2.Cells(jj, 2).Value = sht.Cells(i, 2).Value
    sht2.Cells(jj, 4).Value = sht.Cells(i, 3).Value
    sht2.Cells(jj, 6).Value = sht.Cells(i, 4).Value
    sht2.Cells(jj, 8).Value = sht.Cells(i, 5).Value
    sht2.Cells(jj, 10).Value = sht.Cells(i, 6).Value
    sht2.Cells(jj, 12).Value = sht.Cells(i, 7).Value
    sht2.Cells(jj, 14).Value = sht.Cells(i, 8).Value
    sht2.Cells(jj, 15).Value = sht.Cells(i, 9).Value
    sht2.Range("Q" & jj).Value = i
    jj = jj + 1
    q = 1
End If
Next i

If j = 59 Or jj = 80 Then
result = MsgBox("Cannot display all entries", vbCritical, "Error")
End If

If q = 0 Then
MsgBox ("Noting to disply")
End If

sht2.Range("A40").Select
ActiveWindow.ScrollRow = ActiveCell.Row - 5

Application.Calculation = xlAutomatic
sht.Protect "1234"
sht2.Protect "1234"

End Sub

Sub Display_mem()
Dim i, j, q As Integer
Dim m As String
Dim dd As Date
Dim lastrow As Long

Set sht = Worksheets("data")
Set sht2 = Worksheets("Member Summary")
lastrow = sht.Cells(4, 1).End(xlDown).Row

q = 0
m = sel_mem.ComboBox1.Value
dd = Now()

If m = "" Then
MsgBox ("No member selected")
End
End If

j = 4

Call Clear.Clear_mem

sht.Unprotect "1234"
sht2.Unprotect "1234"

'pause calculation
Application.Calculation = xlManual

For i = 2 To lastrow
    'Trade date
    If sht.Cells(i, 6) >= dd Then
        'Member name
        If sht.Cells(i, 2) = m Then
        sht2.Cells(j, 1).Value = sht.Cells(i, 1).Value
        sht2.Cells(j, 2).Value = sht.Cells(i, 2).Value
        sht2.Cells(j, 4).Value = sht.Cells(i, 3).Value
        sht2.Cells(j, 6).Value = sht.Cells(i, 4).Value
        sht2.Cells(j, 8).Value = sht.Cells(i, 5).Value
        sht2.Cells(j, 10).Value = sht.Cells(i, 6).Value
        sht2.Cells(j, 12).Value = sht.Cells(i, 7).Value
        sht2.Cells(j, 14).Value = sht.Cells(i, 8).Value
        sht2.Cells(j, 15).Value = sht.Cells(i, 9).Value
        sht2.Range("Q" & j).Value = i
        j = j + 1
        q = 1
        End If
    End If
Next i

If q = 0 Then
MsgBox ("Noting to disply")
sht.Protect "1234"
sht2.Protect "1234"
End
End If

sht2.Visible = -1
sht2.Activate
sht2.Cells(4, 1).Select

Application.Calculation = xlAutomatic
sht.Protect "1234"
sht2.Protect "1234"
End

End Sub

