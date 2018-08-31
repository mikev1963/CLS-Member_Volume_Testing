Attribute VB_Name = "Clear"
Sub Clear_day()
' clear the day summary
 Sheets("Summary").Unprotect "1234"
 Range("A40:Q59").Value = ""
 Range("A61:Q80").Value = ""
 Sheets("Summary").Protect "1234"
End Sub

Sub Clear_mem()
' clear the day summary
 Sheets("Member Summary").Unprotect "1234"
  Sheets("Member Summary").Range("A4:Q134").Value = ""
 Sheets("Member Summary").Protect "1234"
End Sub

