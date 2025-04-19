VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Lang VBA
Option Explicit


Public Sub ExecuteWrite()
    Dim filePath As String
    Dim fileNum As Integer
    
    ' Specify the path to the text file
    filePath = ThisWorkbook.Path & "\ExcelAddin.txt"
    
    ' Get a free file number
    fileNum = FreeFile
    
    ' Open the file for output
    Open filePath For Output As #fileNum
    
    ' Write some text to the file
    Print #fileNum, "Hello, World!"
    
    ' Close the file
    Close #fileNum
End Sub