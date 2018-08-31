Attribute VB_Name = "ok_button"
Sub ok_button()
Set sht = Worksheets("data")
Dim lastrow, emptyrow As Long
lastrow = sht.Cells(4, 1).End(xlDown).Row
emptyrow = lastrow + 1

'TBR
If Add_new.OptionButton2.Value = True Then
    Call Add_values((emptyrow), 1)
End If

'annual
If Add_new.OptionButton3.Value = True Then
    Call weekend_holiday
End If

'Volume
If Add_new.OptionButton1.Value = True Then
    Call weekend_holiday
End If


End Sub
Sub weekend_holiday()

Dim i, j, k, answer As Integer
Dim member As String
Dim lastrow, lastrow2, emptyrow As Long

Set sht = Worksheets("data")
Set sht2 = Worksheets("summary")
Set sht3 = Worksheets("No Testing Dates")
Set sht4 = Worksheets("Bank Holidays")
lastrow = sht.Cells(4, 1).End(xlDown).Row
emptyrow = lastrow + 1
lastrow3 = sht3.Cells(4, 1).End(xlDown).Row
lastrow4 = sht4.Cells(3, 1).End(xlDown).Row

sht.Unprotect "1234"
sht2.Unprotect "1234"

'bank holiday check
For k = 4 To lastrow4
If Add_new.value_DTPicker.Value = sht4.Cells(k, 1) Then
    answer = MsgBox("Settlement cannot happen on CLS Bank Holiday" & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "CLS Bank Holiday Settlement")
    If answer = vbNo Then
        sht.Protect "1234"
        sht2.Protect "1234"
        End
    ElseIf answer = vbYes Then
        Call check_testing((lastrow), (emptyrow))
    End If
ElseIf Add_new.Trade_DTPicker.Value = sht3.Cells(k, 1) Then
    answer = MsgBox("Input cannot happen on CLS Bank Holiday" & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "CLS Bank Holiday Settlement")
    If answer = vbNo Then
        sht.Protect "1234"
        sht2.Protect "1234"
        End
    ElseIf answer = vbYes Then
        Call check_testing((lastrow), (emptyrow))
    End If
End If
Next k

'weekend check
If Weekday(Add_new.Trade_DTPicker.Value) = 1 Or Weekday(Add_new.Trade_DTPicker.Value) = 7 Then
    answer = MsgBox("Input cannot happen at weekend" & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "Weekend input")
    If answer = vbNo Then
        sht.Protect "1234"
        sht2.Protect "1234"
        End
    ElseIf answer = vbYes Then
        Call check_testing((lastrow), (emptyrow))
    End If
ElseIf Weekday(Add_new.value_DTPicker.Value) = 1 Or Weekday(Add_new.value_DTPicker.Value) = 7 Then
    answer = MsgBox("Settlement cannot happen at weekend" & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "Weekend Settlement")
    If answer = vbNo Then
        sht.Protect "1234"
        sht2.Protect "1234"
        End
    ElseIf answer = vbYes Then
        Call check_testing((lastrow), (emptyrow))
    End If
End If


'no memeber testing
For j = 4 To lastrow3
If Add_new.value_DTPicker.Value = sht3.Cells(j, 1) And sht3.Cells(j, 3).Value = False Then
    MsgBox ("No member Input allowed on " & Add_new.value_DTPicker.Value)
    sht.Protect "1234"
    sht2.Protect "1234"
    End
ElseIf Add_new.Trade_DTPicker.Value = sht3.Cells(j, 1).Value And sht3.Cells(j, 2).Value = False Then
    MsgBox ("No member settlement allowed on " & Add_new.Trade_DTPicker.Value)
    sht.Protect "1234"
    sht2.Protect "1234"
    End
End If
Next j

If j >= lastrow3 And k >= lastrow4 Then
Call check_testing((lastrow), (emptyrow))
End If

End Sub

Sub check_testing(lastrow As Integer, emptyrow As Integer)
Dim answer, i As Integer

Set sht = Worksheets("data")
Set sht2 = Worksheets("summary")

'Other member testing check
For i = 2 To lastrow
If Add_new.Trade_DTPicker.Value = sht.Cells(i, 6) Or Add_new.Trade_DTPicker.Value = sht.Cells(i, 8) Then
    sht2.Cells(38, 2).Value = Year(Add_new.Trade_DTPicker.Value)
    sht2.Cells(38, 3).Value = Month(Add_new.Trade_DTPicker.Value)
    sht2.Cells(38, 4).Value = Day(Add_new.Trade_DTPicker.Value)
    Call Display.Display_day
    sht.Unprotect "1234"
    sht2.Unprotect "1234"
    answer = MsgBox("Member testing already on trade date." & vbNewLine & "Check day summary for current details." & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "Continue?")
    If answer = vbYes Then
        Call Add_values((emptyrow), 0)
    Else
        sht.Protect "1234"
        sht2.Protect "1234"
        End
    End If
    Exit For
ElseIf Add_new.value_DTPicker.Value = sht.Cells(i, 6) Or Add_new.value_DTPicker.Value = sht.Cells(i, 8) Then
    sht2.Cells(38, 2).Value = Year(Add_new.value_DTPicker.Value)
    sht2.Cells(38, 3).Value = Month(Add_new.value_DTPicker.Value)
    sht2.Cells(38, 4).Value = Day(Add_new.value_DTPicker.Value)
    Call Display.Display_day
    sht.Unprotect "1234"
    sht2.Unprotect "1234"
    answer = MsgBox("Member testing already on value date." & vbNewLine & "Check day summary for current details." & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "Continue?")
    If answer = vbYes Then
        Call Add_values((emptyrow), 0)
    Else
        sht.Protect "1234"
        sht2.Protect "1234"
        End
    End If
    Exit For
End If
Next i

Call Add_values((emptyrow), 0)

End Sub
Sub Add_values(emptyrow As Integer, v As Integer)

Set sht = Worksheets("data")
Set sht2 = Worksheets("summary")

sht.Unprotect "1234"
sht2.Unprotect "1234"

If Add_new.OptionButton1.Value = True Then
    sht.Cells(emptyrow, 1).Value = "Volume Test"
ElseIf Add_new.OptionButton2.Value = True Then
    sht.Cells(emptyrow, 1).Value = "TBR"
ElseIf Add_new.OptionButton3.Value = True Then
    sht.Cells(emptyrow, 1).Value = "Annual Volume Test"
End If


If Add_new.Type_ComboBox.Value = "Own BIC" Then
    sht.Cells(emptyrow, 5).Value = Add_new.BIC_TextBox.Value
Else
Select Case Add_new.Reflex_ComboBox.Value
    Case "Alpha pay"
        sht.Cells(emptyrow, 5).Value = "ZYGPGB40XXX"
    Case "Beta pay"
        sht.Cells(emptyrow, 5).Value = "ZYGRGB40XXX"
    Case "Gamma pay"
        sht.Cells(emptyrow, 5).Value = "ZYGSGB40XXX"
    Case "Late pay"
        sht.Cells(emptyrow, 5).Value = "ZYGMGB40XXX"
    Case "Part pay"
        sht.Cells(emptyrow, 5).Value = "ZYGOGB40XXX"
    Case "Never pay"
        sht.Cells(emptyrow, 5).Value = "ZYGNGB40XXX"
    Case "Lambda pay"
        sht.Cells(emptyrow, 5).Value = "ZYGQGB40XXX"
    Case "Kappa pay"
        sht.Cells(emptyrow, 5).Value = "ZYGTGB40XXX"
End Select
End If

sht.Cells(emptyrow, 2).Value = Add_new.member_ComboBox.Value
sht.Cells(emptyrow, 3).Value = Add_new.Type_ComboBox.Value
sht.Cells(emptyrow, 4).Value = Add_new.Reflex_ComboBox.Value
sht.Cells(emptyrow, 6).Value = Add_new.Trade_DTPicker.Value
sht.Cells(emptyrow, 7).Value = Add_new.Sides_TextBox.Value
sht.Cells(emptyrow, 8).Value = Add_new.value_DTPicker.Value
sht.Cells(emptyrow, 9).Value = Add_new.ref_TextBox.Value

If v = 1 Then
'display value day
sht2.Cells(38, 2).Value = Year(Add_new.value_DTPicker.Value)
sht2.Cells(38, 3).Value = Month(Add_new.value_DTPicker.Value)
sht2.Cells(38, 4).Value = Day(Add_new.value_DTPicker.Value)
Call Display.Display_day
Else
'display trade day
sht2.Cells(38, 2).Value = Year(Add_new.Trade_DTPicker.Value)
sht2.Cells(38, 3).Value = Month(Add_new.Trade_DTPicker.Value)
sht2.Cells(38, 4).Value = Day(Add_new.Trade_DTPicker.Value)
Call Display.Display_day
End If

sht.Protect "1234"
sht2.Protect "1234"

End

End Sub

Sub trade_volume(emptyrow As Integer)
Dim v As Long
Set sht = Worksheets("data")
Set sht2 = Worksheets("summary")

'display trade day to check volumes
sht2.Cells(38, 2).Value = Year(Add_new.Trade_DTPicker.Value)
sht2.Cells(38, 3).Value = Month(Add_new.Trade_DTPicker.Value)
sht2.Cells(38, 4).Value = Day(Add_new.Trade_DTPicker.Value)
Call Display.Display_day
sht.Unprotect "1234"
sht2.Unprotect "1234"

'check volumes are within limited
Select Case Add_new.Reflex_ComboBox.Value
    Case "Alpha pay"
        v = Add_new.Sides_TextBox.Value + sht2.Range("S41").Value
        If v >= sht2.Range("T41").Value Then
            answer = MsgBox("Trade volumes will go above recommended levels." & vbNewLine & "Check day summary for current details." & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "Continue?")
            If answer = vbYes Then
                Call Add_values((emptyrow), 0)
                Call Display.Display_day
            Else
                sht.Protect "1234"
                sht2.Protect "1234"
                End
            End If
        End If
    Case "Beta pay"
        v = Add_new.Sides_TextBox.Value + sht2.Range("S42").Value
        If v >= sht2.Range("T42") Then
            answer = MsgBox("Trade volumes will go above recommended levels." & vbNewLine & "Check day summary for current details." & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "Continue?")
            If answer = vbYes Then
                Call Add_values((emptyrow), 0)
                Call Display.Display_day
            Else
                sht.Protect "1234"
                sht2.Protect "1234"
                End
            End If
        End If
    Case "Gamma pay"
        v = Add_new.Sides_TextBox.Value + sht2.Range("S43").Value
        If v >= sht2.Range("T43") Then
            answer = MsgBox("Trade volumes will go above recommended levels." & vbNewLine & "Check day summary for current details." & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "Continue?")
            If answer = vbYes Then
                Call Add_values((emptyrow), 0)
                Call Display.Display_day
            Else
                sht.Protect "1234"
                sht2.Protect "1234"
                End
            End If
        End If
    Case "Late pay"
        v = Add_new.Sides_TextBox.Value + sht2.Range("S44").Value
        If v >= sht2.Range("T44") Then
            answer = MsgBox("Trade volumes will go above recommended levels." & vbNewLine & "Check day summary for current details." & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "Continue?")
            If answer = vbYes Then
                Call Add_values((emptyrow), 0)
                Call Display.Display_day
            Else
                sht.Protect "1234"
                sht2.Protect "1234"
                End
            End If
        End If
    Case "Part pay"
        v = Add_new.Sides_TextBox.Value + sht2.Range("S45").Value
        If v >= sht2.Range("T45") Then
            answer = MsgBox("Trade volumes will go above recommended levels." & vbNewLine & "Check day summary for current details." & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "Continue?")
            If answer = vbYes Then
                Call Add_values((emptyrow), 0)
                Call Display.Display_day
            Else
                sht.Protect "1234"
                sht2.Protect "1234"
                End
            End If
        End If
    Case "Never pay"
        v = Add_new.Sides_TextBox.Value + sht2.Range("S46").Value
        If v >= sht2.Range("T46") Then
            answer = MsgBox("Trade volumes will go above recommended levels." & vbNewLine & "Check day summary for current details." & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "Continue?")
            If answer = vbYes Then
                Call Add_values((emptyrow), 0)
                Call Display.Display_day
            Else
                sht.Protect "1234"
                sht2.Protect "1234"
                End
            End If
        End If
    Case "Lambda pay"
        v = Add_new.Sides_TextBox.Value + sht2.Range("S47").Value
        If v >= sht2.Range("T47") Then
            answer = MsgBox("Trade volumes will go above recommended levels." & vbNewLine & "Check day summary for current details." & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "Continue?")
            If answer = vbYes Then
                Call Add_values((emptyrow), 0)
                Call Display.Display_day
            Else
                sht.Protect "1234"
                sht2.Protect "1234"
                End
            End If
        End If
    Case "Kappa pay"
        v = Add_new.Sides_TextBox.Value + sht2.Range("S48").Value
        If v >= sht2.Range("T48") Then
            answer = MsgBox("Trade volumes will go above recommended levels." & vbNewLine & "Check day summary for current details." & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbQuestion, "Continue?")
            If answer = vbYes Then
                Call Add_values((emptyrow), 0)
                Call Display.Display_day
            Else
                sht.Protect "1234"
                sht2.Protect "1234"
                End
            End If
        End If
End Select

End Sub
