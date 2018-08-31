Attribute VB_Name = "Edit"
Sub Edit_start()
Dim line As Integer
Set sht = Sheets("Data")

sht.Visible = -1
sht.Activate
sht.Cells(3, 1).Select
sht.Unprotect "1234"

line = InputBox("Line number to edit")

sht.Range("A1").Value = line

For i = 1 To 9
    sht.Cells(3, i).Value = sht.Cells(line, i).Value2
Next i

sht.Shapes("Button 1").Visible = False
sht.Shapes("Button 5").Visible = False
sht.Shapes("Button 3").Visible = True
sht.Shapes("Button 7").Visible = True

sht.Protect "1234"

End Sub

Sub Edit_finish()
Dim line As Integer
Set sht = Sheets("Data")

sht.Unprotect "1234"

line = sht.Range("A1")

sht.Range("A1").Value = line

For i = 1 To 9
    sht.Cells(line, i).Value2 = sht.Cells(3, i).Value
    sht.Cells(3, i).Value = ""
Next i

sht.Shapes("Button 1").Visible = True
sht.Shapes("Button 5").Visible = True
sht.Shapes("Button 3").Visible = False
sht.Shapes("Button 7").Visible = False

sht.Protect "1234"


End Sub

Sub edit_delete()
Dim line As Integer
Set sht = Sheets("Data")

sht.Unprotect "1234"

line = sht.Range("A1")
sht.Range("A1").Value = line

answer = MsgBox("Are you sure you want to delete line " & line, vbYesNo + vbQuestion, "Delete?")
    If answer = vbNo Then
        sht.Protect "1234"
        End
    ElseIf answer = vbYes Then
        Rows(line).EntireRow.Delete
        For i = 1 To 9
            sht.Cells(3, i).Value = ""
        Next i
        
        sht.Shapes("Button 1").Visible = True
        sht.Shapes("Button 5").Visible = True
        sht.Shapes("Button 3").Visible = False
        sht.Shapes("Button 7").Visible = False
        sht.Protect "1234"
    End If

End Sub
