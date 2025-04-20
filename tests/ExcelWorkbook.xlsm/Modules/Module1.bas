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
    filePath = ThisWorkbook.Path & "\ExcelWorkbook.txt"
    
    ' Get a free file number
    fileNum = FreeFile
    
    ' Open the file for output
    Open filePath For Output As #fileNum
    
    ' Write some text to the file
    Print #fileNum, "Hello, World!"
    
    ' Close the file
    Close #fileNum
End Sub