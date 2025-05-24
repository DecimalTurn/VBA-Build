Attribute VB_Name = "DebugAssertDemo"
Option Explicit
Option Private Module

'@TestModule

'--- Rubberduck Boilerplate ---

'Idea: Currently there is a lot of code that doesn't use the Rubberduck framework for testing.
'Usually, this means that people are using Debug.Assert to test their code.
'To automate testing we could replace Debug.Assert with a call to a function named DebugRD that uses the Rubberduck framework.
'This would allow us to run the tests in the CI/CD pipeline.
'The only thing is that we need to mark the module with '@TestModule so that Rubberduck can find it and 
' the tests with '@TestMethod so that Rubberduck can run them after inserting the boilerplate code and making
' the necessary replacements of Debug.Assert with DebugRD.

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

Sub DebugAssert(ByRef Assert As Rubberduck.AssertClass, ByVal Statement As Boolean)
        On Error GoTo TestFail
        Assert.IsTrue Statement
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'--- End of Rubberduck Boilerplate ---

'@TestMethod
Private Sub SimplestTest()
    Debug.Assert True
End Sub

Private Sub SimplestTest2()
    DebugAssert Assert, True
End Sub

