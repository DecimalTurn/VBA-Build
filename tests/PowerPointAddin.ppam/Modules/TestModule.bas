Attribute VB_Name = "TestModule"
Option Explicit

' Test for the HelloFromPowerPointAddin function
Public Sub TestHelloFromPowerPointAddin()
    ' This function tests that the HelloFromPowerPointAddin function works correctly
    
    Dim expected As String
    expected = "Hello from PowerPoint Addin!"
    
    Dim actual As String
    actual = HelloFromPowerPointAddin()
    
    ' Assert
    If actual = expected Then
        Debug.Print "TestHelloFromPowerPointAddin: PASSED"
    Else
        Debug.Print "TestHelloFromPowerPointAddin: FAILED - Expected '" & expected & "' but got '" & actual & "'"
    End If
End Sub

' Test runner function
Public Sub RunAllTests()
    Debug.Print "Running all PowerPoint Addin tests..."
    TestHelloFromPowerPointAddin
    Debug.Print "All tests completed"
End Sub
