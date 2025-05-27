Attribute VB_Name = "Module1"
Option Explicit

Public Function HelloFromPowerPointAddin() As String
    HelloFromPowerPointAddin = "Hello from PowerPoint Addin!"
End Function

' Example of a utility function that might be included in a PowerPoint addin
Public Sub InsertSlideWithTitle(titleText As String)
    Dim newSlide As Slide
    Set newSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutTitle)
    newSlide.Shapes.Title.TextFrame.TextRange.Text = titleText
End Sub
