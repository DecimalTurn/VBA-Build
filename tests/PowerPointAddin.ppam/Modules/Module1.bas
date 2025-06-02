Attribute VB_Name = "Module1"

'@Lang VBA
Option Explicit

Sub Demo()
    MsgBox "Hello, World!"
End Sub

'This code will be called via COM to test if the VBA import was successful
Sub WriteToFile()
    Dim filePath As String
    Dim fileNum As Integer
    
    ' Specify the path to the text file
    filePath = ThisAddin.Path & "\PowerPointAddin.txt"
    
    ' Get a free file number
    fileNum = FreeFile
    
    ' Open the file for output
    Open filePath For Output As #fileNum
    
    ' Write some text to the file
    Print #fileNum, "Hello, World!"
    
    ' Close the file
    Close #fileNum
End Sub

Private Function ThisAddin() As PowerPoint.Addin
    ' This function returns the current Addin object
    ' It is used to access the Addin's properties and methods
    Dim addin As PowerPoint.Addin

    For Each addin In Application.AddIns
        If addin.Name = "PowerPointAddin" Then
            Set ThisAddin = addin
            Exit Function
        End If
    Next addin
    ' If the addin is not found, return Nothing
    Set ThisAddin = Nothing
    Exit Function

End Function