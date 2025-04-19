Attribute VB_Name = "Module1"

'@Lang VBA
Option Explicit

Sub Demo()
    MsgBox "Hello, World!"
End Sub

'This code will be called via COM to test if the VBA import was successful
Sub WriteToFile()
    Class1.ExecuteWrite
End Sub