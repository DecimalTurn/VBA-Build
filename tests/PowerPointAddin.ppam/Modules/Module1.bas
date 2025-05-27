Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    MsgBox "This is a demo subroutine in the PowerPoint addin!"
End Sub

Public Function HelloFromPowerPointAddin() As String
    HelloFromPowerPointAddin = "Hello from PowerPoint Addin!"
End Function

' Example of a utility function that might be included in a PowerPoint addin
Public Sub InsertSlideWithTitle(titleText As String)
    Dim newSlide As Slide
    Set newSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutTitle)
    newSlide.Shapes.Title.TextFrame.TextRange.Text = titleText
End Sub

' Standard test procedure used by the build process
Public Sub WriteToFile()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim filePath As String
    filePath = ActivePresentation.Path & "\PowerPointAddin.txt"
    
    Dim fileObject As Object
    Set fileObject = fso.CreateTextFile(filePath, True)
    
    fileObject.WriteLine "Hello, World!"
    fileObject.Close
    
    Set fileObject = Nothing
    Set fso = Nothing
End Sub
