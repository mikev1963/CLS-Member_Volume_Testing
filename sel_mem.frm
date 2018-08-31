VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sel_mem 
   Caption         =   "Select Member"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   OleObjectBlob   =   "sel_mem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sel_mem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
 
 Call Display.Display_mem
 
End Sub

Sub UserForm_Initialize()
Dim lastrow As Integer

'Add options to Member textbox
lastrow = Sheets("members").Cells(1, 1).End(xlDown).Row
ComboBox1.List = Sheets("members").Range("A2:A" & lastrow).Value
ComboBox1.Value = ""
End Sub
