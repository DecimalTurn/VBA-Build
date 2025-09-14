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
    filePath = ThisPresentation.Path & "\PowerPointPresentation.txt"
    
    ' Get a free file number
    fileNum = FreeFile
    
    ' Open the file for output
    Open filePath For Output As #fileNum
    
    ' Write some text to the file
    Print #fileNum, "Hello, World!"
    
    ' Close the file
    Close #fileNum
End Sub

Private Function ThisPresentation() As PowerPoint.Presentation
    Dim Pres As Presentation
    On Error Resume Next
        Set Pres = Presentations("PowerPointPresentation.pptm")
    On Error GoTo 0
    If Pres Is Nothing Then
        'Try to fallback using ActiveVBProject
        Dim PresName As String
        PresName = Split(Application.VBE.ActiveVBProject.FileName, "\")(UBound(Split(Application.VBE.ActiveVBProject.FileName, "\")))
        Set Pres = Presentations(PresName)
    End If
    Set ThisPresentation = Pres
    Exit Function
End Function