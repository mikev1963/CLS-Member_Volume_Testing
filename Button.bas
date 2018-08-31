Attribute VB_Name = "Button"
Sub Add_testing()
Attribute Add_testing.VB_Description = "New member testing"
Attribute Add_testing.VB_ProcData.VB_Invoke_Func = "n\n14"
Add_new.Show
End Sub

Sub view_mem()
sel_mem.Show
End Sub

Sub Add_no_test()
Dim i As Integer
Set sht = ActiveWorkbook.Sheets("No testing dates")

sht.Visible = -1
i = sht.Cells(3, 1).End(xlDown).Row

sht.Activate
sht.Cells(i + 1, 1).Select

End Sub

Sub Make_edit()
Set sht = Sheets("Data")
sht.Activate
sht.Cells(3, 1).Select
Call Edit.Edit_start
End Sub

Sub return_no_test()
Set sht = ActiveWorkbook.Sheets("No testing dates")
Set sht2 = ActiveWorkbook.Sheets("Summary")

sht2.Activate
sht2.Cells(2, 1).Select

sht.Visible = 2

End Sub

Sub see_data()
Set sht = ActiveWorkbook.Sheets("data")
Set sht2 = ActiveWorkbook.Sheets("Member summary")

sht2.Visible = 2
sht.Visible = -1
sht.Activate
sht.Cells(3, 1).Select

End Sub
Sub return_data()
Set sht = ActiveWorkbook.Sheets("data")
Set sht2 = ActiveWorkbook.Sheets("Summary")

sht2.Activate
sht2.Cells(2, 1).Select
sht.Visible = 2

End Sub
Sub see_members()
Set sht = ActiveWorkbook.Sheets("Members")

sht.Visible = -1
sht.Activate
sht.Cells(3, 1).Select

End Sub
Sub return_members()
Set sht = ActiveWorkbook.Sheets("Members")
Set sht2 = ActiveWorkbook.Sheets("Summary")

sht2.Activate
sht2.Cells(2, 1).Select
sht.Visible = 2

End Sub
Sub return_BH()
Set sht = ActiveWorkbook.Sheets("Bank Holidays")
Set sht2 = ActiveWorkbook.Sheets("Summary")

sht2.Activate
sht2.Cells(2, 1).Select

sht.Visible = 2

End Sub
Sub Add_BH()
Dim i As Integer
Set sht = ActiveWorkbook.Sheets("Bank Holidays")

sht.Visible = -1
i = sht.Cells(3, 1).End(xlDown).Row

sht.Activate
sht.Cells(i + 1, 1).Select

End Sub
Sub return_mem_sum()
Set sht = ActiveWorkbook.Sheets("Member Summary")
Set sht2 = ActiveWorkbook.Sheets("Summary")

sht2.Activate
sht2.Cells(2, 1).Select

sht.Visible = 2

End Sub
