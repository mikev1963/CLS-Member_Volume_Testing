Attribute VB_Name = "Module2"
Sub savebat()
Const Module = 1
Const ClassModule = 2
Const Form = 3
Const Document = 100
Const Padding = 24

Dim VBComponent As Object
Dim extension, n, directory, path As String

n = Application.ThisWorkbook.Name
directory = "\\OFFNS001GB.prod.local\CLSUserHome$\broche\Documents\VBA code"
    
    
For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
    Select Case VBComponent.Type
        Case ClassModule, Document
            'do nothing
        Case Form
            extension = ".frm"
            path = directory & "\" & VBComponent.Name & extension
            Call VBComponent.Export(path)
        Case Module
            extension = ".bas"
            path = directory & "\" & VBComponent.Name & extension
            Call VBComponent.Export(path)
        Case Else
            extension = ".txt"
            path = directory & "\" & VBComponent.Name & extension
            Call VBComponent.Export(path)
    End Select
    

Next

End Sub

