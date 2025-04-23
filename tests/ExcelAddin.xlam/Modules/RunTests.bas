Attribute VB_Name = "RunTests"

Sub RunAllTests()
    Dim Output As String
    With Application.VBE.AddIns("Rubberduck.Extension").Object
        Dim LogPath As String
        LogPath = ThisWorkbook.Path & "\RD_Log_" & ThisWorkbook.Name & ".txt"
        Output = .RunAllTestsAndGetResults(LogPath)
    End With
    ' Output the result to the Immediate Window 
    Debug.Print Output
End Sub