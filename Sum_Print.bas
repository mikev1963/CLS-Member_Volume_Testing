Attribute VB_Name = "Sum_Print"
Sub Print_day()
Set sht = Worksheets("summary")

With sht.PageSetup
    .Orientation = xlLandscape
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = 1
End With

sht.Range("A37:T70").PrintPreview

End Sub
Sub Print_mem()
Set sht = Worksheets("Member Summary")

With sht.PageSetup
    .Orientation = xlPortrait
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = 1
End With

sht.Range("A1:Q100").PrintPreview


End Sub

Sub Print_mon()
Set sht = Worksheets("summary")

With sht.PageSetup
    .Orientation = xlLandscape
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = 1
End With

sht.Range("A1:L35").PrintPreview


End Sub
