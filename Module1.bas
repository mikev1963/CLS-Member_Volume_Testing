Attribute VB_Name = "Module1"
Sub Add_testing()
Dim i As Integer
Set sht = Sheets("No testing dates")

i = sht.Cells(3, 1).End(xlDown).Row

sht.Activate
sht.Cells(i + 1, 1).Select

End Sub
