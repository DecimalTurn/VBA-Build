Attribute VB_Name = "TestsForIsCapitalLetterFunction"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize >>BeforeAll
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup >>AfterAll
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize >> BeforeEach
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup >>AfterEach
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("IsCapitalLetter Function Test >>Capital Letter Case")
Private Sub IsCapitalLetterForCapitalLetterCase()
    On Error GoTo TestFail
        Dim Actual As Boolean
        Actual = FunctionCollection.IsCapitalLetter("A")
        Assert.IsTrue Actual, "This Test should return True"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsCapitalLetter Function Test>> Small Letter Case")
Private Sub IsCapitalLetterForSmallerLetterCase()
    On Error GoTo TestFail
        Dim Actual As Boolean
        Actual = FunctionCollection.IsCapitalLetter("a")
        Assert.IsFalse Actual, "This Test should return false"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsCapitalLetter Function Test>> Digit  Case")
Private Sub IsCapitalLetterForDigitCase()
    On Error GoTo TestFail
        Dim Actual As Boolean
        Actual = FunctionCollection.IsCapitalLetter("8")
        Assert.IsFalse Actual, "This Test should return false"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsCapitalLetter Function Test>> Error Raising  Case")
Private Sub IsCapitalLetterForErrorCase()
    Const ExpectedError As Long = 13
    On Error GoTo TestFail
    Dim Actual As Boolean
    Actual = FunctionCollection.IsCapitalLetter("ISMAIL")
    Assert.IsFalse Actual, "This Test should return false"
        
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("IsCapitalLetter Function Test>> NullString  Case")
Private Sub IsCapitalLetterForNullStringCase()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    Dim Actual As Boolean
    Actual = FunctionCollection.IsCapitalLetter(vbNullString)
    Assert.IsFalse Actual, "This Test should return false"
        
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub