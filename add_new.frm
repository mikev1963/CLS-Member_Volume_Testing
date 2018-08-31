VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} add_new 
   Caption         =   "Add New"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   OleObjectBlob   =   "add_new.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "add_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()

Call ok_button.ok_button

End Sub



Private Sub Label2_Click()

End Sub

Private Sub OptionButton3_Click()

End Sub

Private Sub OptionButton2_Change()

If OptionButton2.Value = True Then
    Add_new.BIC_TextBox.Visible = True
    Add_new.Label5.Visible = True
    Add_new.Label4.Visible = False
    Add_new.Reflex_ComboBox.Visible = False
    Add_new.Type_ComboBox.Visible = False
    Add_new.Label10.Visible = False
    Add_new.Trade_DTPicker.Visible = False
    Add_new.Label6.Visible = False
Else
    Add_new.BIC_TextBox.Visible = False
    Add_new.Label5.Visible = False
    Add_new.Label4.Visible = False
    Add_new.Reflex_ComboBox.Visible = False
    Add_new.Type_ComboBox.Visible = True
    Add_new.Label10.Visible = True
    Add_new.Trade_DTPicker.Visible = True
    Add_new.Label6.Visible = True
End If

End Sub

Private Sub OptionButton3_Change()

If OptionButton3.Value = True Then
    Add_new.BIC_TextBox.Visible = True
    Add_new.Label5.Visible = True
    Add_new.Label4.Visible = False
    Add_new.Reflex_ComboBox.Visible = False
    Add_new.Type_ComboBox.Visible = False
    Add_new.Label10.Visible = False
    Add_new.Type_ComboBox.Value = "Own BIC"
Else
    Add_new.BIC_TextBox.Visible = False
    Add_new.Label5.Visible = False
    Add_new.Label4.Visible = False
    Add_new.Reflex_ComboBox.Visible = False
    Add_new.Type_ComboBox.Visible = True
    Add_new.Label10.Visible = True
    Add_new.Type_ComboBox.Value = ""
End If

End Sub



Private Sub type_comboBox_Change()

' Hide/Show options
If Type_ComboBox.Value = "Own BIC" Then
    Add_new.BIC_TextBox.Visible = True
    Add_new.Label5.Visible = True
    Add_new.Label4.Visible = False
    Add_new.Reflex_ComboBox.Visible = False
ElseIf Type_ComboBox.Value = "Reflex" Then
    Add_new.BIC_TextBox.Visible = False
    Add_new.Label5.Visible = False
    Add_new.Label4.Visible = True
    Add_new.Reflex_ComboBox.Visible = True
Else
    Add_new.BIC_TextBox.Visible = False
    Add_new.Label5.Visible = False
    Add_new.Label4.Visible = False
    Add_new.Reflex_ComboBox.Visible = False
End If

End Sub

Sub UserForm_Initialize()
Dim lastrow As Integer


'Add options to Member textbox
lastrow = Sheets("members").Cells(1, 1).End(xlDown).Row
member_ComboBox.List = Sheets("members").Range("A2:A" & lastrow).Value

'Default Request Type
OptionButton1.Value = True

'Counterpary type options
With Type_ComboBox
    .AddItem "Own BIC"
    .AddItem "Reflex"
End With

'Clear Own BIC textbox
BIC_TextBox.Value = ""

'Reflex options
With Reflex_ComboBox
    .AddItem "Alpha pay"
    .AddItem "Beta pay"
    .AddItem "Gamma pay"
    .AddItem "Late pay"
    .AddItem "Part pay"
    .AddItem "Never pay"
    .AddItem "Lambda pay"
    .AddItem "Kappa pay"
End With
    
'clear Sides textbox
Sides_TextBox.Value = ""
 
'Default CLS Ref# textbox
ref_TextBox.Value = "RITM00"

' set dates to today
Trade_DTPicker.Value = Date
value_DTPicker.Value = Date + 1

'Set focus on Member textbox
member_ComboBox.SetFocus

'Hide conditional options
Add_new.BIC_TextBox.Visible = False
Add_new.Label5.Visible = False
Add_new.Label4.Visible = False
Add_new.Reflex_ComboBox.Visible = False

End Sub


