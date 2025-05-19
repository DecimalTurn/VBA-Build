Attribute VB_Name = "FunctionCollection"
Option Explicit

'Source: https://www.youtube.com/watch?v=sR8ZM0XUXdk

'@Description("This will make a valid constant name according to convention")
'@Dependency("IsCapitalLetter function")
'@ExampleCall : MakeValidConstName("SettingsTable") >> SETTINGS_TABLE
'@Date : 14 October 2021 08:05:41 PM
Public Function MakeValidConstName(ByVal GivenName As String) As String
    
    If Trim(GivenName) = vbNullString Then
        Err.Raise 13, "MakeValidConstName Function", "Constant Name can't be Nullstring"
    ElseIf Not (Left(GivenName, 1) Like "[A-Za-z]") Then
        Err.Raise 13, "MakeValidConstName Function", "Constant Name Should be start with A-Z or a-z"
    End If
    
    Dim Result As String
    If UCase(GivenName) = GivenName Then
        MakeValidConstName = GivenName
        Exit Function
    End If
    Dim Counter As Long
    Dim CurrentCharacter As String
    Const WORD_SEPARATOR As String = "_"
    Result = Left(GivenName, 1)
    For Counter = 2 To Len(GivenName)
        CurrentCharacter = Mid(GivenName, Counter, 1)
        If IsCapitalLetter(CurrentCharacter) Then
            Result = Result & WORD_SEPARATOR
        End If
        Result = Result & CurrentCharacter
    Next Counter
    Result = UCase(Result)
    MakeValidConstName = Result
    
End Function


'@Description("This will check if a given character is Capital letter > A-Z..It will throw error if length of the letter is more than 1")
'@Dependency("No Dependency")
'@ExampleCall : IsCapitalLetter(CurrentCharacter)
'@Date : 14 October 2021 10:23:19 PM
Public Function IsCapitalLetter(ByVal GivenLetter As String) As Boolean
    If Len(GivenLetter) > 1 Then
        Err.Raise 13, "IsCapitalLetter Function", "Given Letter need to be one character String"
    End If
    If GivenLetter = vbNullString Then
        Err.Raise 5, "IsCapitalLetter Function", "Given Letter can't be nullstring"
    End If
    
    Const ASCII_CODE_FOR_A As Integer = 65
    Const ASCII_CODE_FOR_Z As Integer = 90
    Dim ASCIICodeForGivenLetter As Integer
    ASCIICodeForGivenLetter = Asc(GivenLetter)
    IsCapitalLetter = (ASCIICodeForGivenLetter >= ASCII_CODE_FOR_A And ASCIICodeForGivenLetter <= ASCII_CODE_FOR_Z)

End Function

