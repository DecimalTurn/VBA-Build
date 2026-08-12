Attribute VB_Name = "modTestProcedures"
Option Compare Database
Option Explicit

Public Sub TestProcedure()
    SaveTestInfo "TestProcedure"
End Sub

Public Sub TestProcedure2(ByVal T As String, ByVal N As Long)
    SaveTestInfo "TestProcedure2", T, N
End Sub

Private Sub SaveTestInfo(ByVal ProcName As String, Optional ByVal T As Variant, Optional ByVal N As Variant)

    With CurrentDb.OpenRecordset("tabTest", dbOpenTable, dbAppendOnly)
        .AddNew
        .Fields("ProcName").Value = ProcName
        If Not IsMissing(T) Then
            .Fields("T").Value = T
        End If
        If Not IsMissing(N) Then
            .Fields("N").Value = N
        End If
        .Update
        .Close
    End With

End Sub
