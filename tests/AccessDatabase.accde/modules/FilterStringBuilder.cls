VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: data.sql.FilterStringBuilder
'---------------------------------------------------------------------------------------
'
' Create SQL criteria (filter) expression
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/FilterStringBuilder.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>text/StringCollection.cls</use>
'  <use>data/SqlTools.cls</use>
'  <test>_test/data/FilterStringBuilderTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private m_Items As StringCollection
Private m_SqlTool As SqlTools
Private m_DefaultConcatOperator As SqlLogicalOperator

Private m_ConditionGroups As Collection

Private m_SqlPrefix As String
Private m_SqlSuffix As String
Private m_ConditionPrefix As String
Private m_IgnoreSqlPreSuffixIfEmptyFilter As Boolean

Private Const WhereReplacementText As String = "[WhereStatement]"

Private Sub Class_Initialize()
   Set m_Items = New StringCollection
   m_DefaultConcatOperator = SqlLogicalOperator.SQL_And
End Sub

Private Sub Class_Terminate()
   Set m_Items = Nothing
End Sub

'##################################
' Group: Class support

'---------------------------------------------------------------------------------------
' Property: Self
'---------------------------------------------------------------------------------------
'
' Reference to Me (current instance of FilterStringBuilder class)
'
'---------------------------------------------------------------------------------------
Public Property Get Self() As FilterStringBuilder
   Set Self = Me
End Property

'##################################
' Group: Config

'---------------------------------------------------------------------------------------
' Property: DefaultConcatOperator
'---------------------------------------------------------------------------------------
'
' The default concat operator for the instance
'
'---------------------------------------------------------------------------------------
Public Property Get DefaultConcatOperator() As SqlLogicalOperator
   DefaultConcatOperator = m_DefaultConcatOperator
End Property

Public Property Let DefaultConcatOperator(ByVal NewValue As SqlLogicalOperator)
   m_DefaultConcatOperator = NewValue
End Property

'---------------------------------------------------------------------------------------
' Property: SqlTool
'---------------------------------------------------------------------------------------
'
' SqlTools instance to be used (if not set, clone of global SqlTools instance is used)
'
'---------------------------------------------------------------------------------------
Friend Property Get SqlTool() As SqlTools
   If m_SqlTool Is Nothing Then
      Set m_SqlTool = SqlTools.Clone
   End If
   Set SqlTool = m_SqlTool
End Property

Friend Property Set SqlTool(ByVal NewRef As SqlTools)
   Set m_SqlTool = NewRef
End Property

'---------------------------------------------------------------------------------------
' Function: SelectSqlDialect
'---------------------------------------------------------------------------------------
'
' Config sql text output format for specific sql dialect
'
' Parameters:
'
'     UseDialect - sql dialect .. use setting of dialect (as base)
'     SqlDateFormat - output string format for date values
'     SqlBooleanTrueString - output string format for boolean values
'     SqlWildCardString - wildcard string (e.g. * .. dao, % .. T-SQL)
'
'---------------------------------------------------------------------------------------
Friend Sub SelectSqlDialect(ByVal UseDialect As SqlDialect, _
                   Optional ByVal SqlDateFormat As String, _
                   Optional ByVal SqlBooleanTrueString As String, _
                   Optional ByVal SqlWildCardString As String)

   Set m_SqlTool = SqlTools.FromDialect(UseDialect, SqlDateFormat, SqlBooleanTrueString, SqlWildCardString)

End Sub

'---------------------------------------------------------------------------------------
' Function: ConfigSqlFormat
'---------------------------------------------------------------------------------------
'
' Config sql text output format for specific sql dialect
'
' Parameters:
'
'     SqlDateFormat - output string format for date values
'     SqlBooleanTrueString - output string format for boolean values
'     SqlWildCardString - wildcard string (e.g. * .. dao, % .. T-SQL)
'
'---------------------------------------------------------------------------------------
Friend Sub ConfigSqlFormat(ByVal SqlDateFormat As String, _
                           ByVal SqlBooleanTrueString As String, _
                           ByVal SqlWildCardString As String)

   If m_SqlTool Is Nothing Then
      Set m_SqlTool = SqlTools.NewInstance(SqlDateFormat, SqlBooleanTrueString, SqlWildCardString)
      Exit Sub
   End If

   With m_SqlTool
      .SqlDateFormat = SqlDateFormat
      .SqlBooleanTrueString = SqlBooleanTrueString
      .SqlWildCardString = SqlWildCardString
   End With

End Sub

'---------------------------------------------------------------------------------------
' Function: ConfigSqlStatement
'---------------------------------------------------------------------------------------
'
' Auxiliary function for setting SQL texts if this class is used in other filter builder classes.
' (e.g. in form.filter.FilterControlCollection)
'
' Parameters:
'
'     SqlPrefix - Is placed before the FilterString for ToString output (ToString = m_SqlPrefix & FilterString & m_SqlSuffix)
'     SqlSuffix - Is placed after the FilterString for ToString output (ToString = m_SqlPrefix & FilterString & m_SqlSuffix)
'     ConditionPrefix - Is placed before the filter condition (but only if the filter condition is > "").
'     IgnoreSqlPreSuffixIfEmptyFilter - True: SqlPrefix + SqlSuffix are not set for empty FilterString
'
'---------------------------------------------------------------------------------------
Friend Sub ConfigSqlStatement(ByVal SqlPrefix As String, ByVal SqlSuffix As String, _
                              ByVal ConditionPrefix As String, _
                     Optional ByVal IgnoreSqlPreSuffixIfEmptyFilter As Boolean = False)
   m_SqlPrefix = SqlPrefix
   m_SqlSuffix = SqlSuffix
   m_ConditionPrefix = ConditionPrefix
   m_IgnoreSqlPreSuffixIfEmptyFilter = IgnoreSqlPreSuffixIfEmptyFilter

End Sub

'##################################
' Group: Build Criteria

'---------------------------------------------------------------------------------------
' Function: Add
'---------------------------------------------------------------------------------------
'
' Add filter criteria definition
' (uses: data.SqlTools.BuildCriteria)
'
' Parameters:
'
'     FieldName            - Field name in the data source to be filtered
'     RelationalOperator   - Relational operator (=, <=, etc.)
'     Value                - Filter value (can be a single value or an array of values)
'     Value2               - Optional 2nd filter value (for Between)
'     IgnoreValue          - The value for which no filter condition is to be created. (Array transfer of values possible), Default: Null
'
'---------------------------------------------------------------------------------------
Public Sub Add(ByVal FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
               ByVal RelationalOperator As SqlRelationalOperators, _
               ByVal Value As Variant, _
      Optional ByVal Value2 As Variant = Null, _
      Optional ByVal IgnoreValue As Variant = Null)

   AddCriteria SqlTool.BuildCriteria(FieldName, FieldDataType, RelationalOperator, Value, Value2, IgnoreValue)

End Sub

'---------------------------------------------------------------------------------------
' Function: AddCriteria
'---------------------------------------------------------------------------------------
'
' Add filter criteria (String)
'
' Parameters:
'
'     Criteria - Criteria string to append
'
'---------------------------------------------------------------------------------------
Public Sub AddCriteria(ByVal Criteria As String)
   If Len(Criteria) = 0 Then Exit Sub
   m_Items.Add Criteria
End Sub

'---------------------------------------------------------------------------------------
' Function: NewConditionGroup
'---------------------------------------------------------------------------------------
'
' New filter condition group - e.g. for Or group: ( a = 1 and (x = 2 or x >= 10) )
'
' Parameters:
'
'     ConcatOperator - [SQL_And (default), SQL_Or, SQL_CommaSeparator] .. valid within the new group
'
' Returns:
'
'     FilterStringBuilder instance for new group
'
'---------------------------------------------------------------------------------------
Public Function NewConditionGroup(ByVal ConcatOperator As SqlLogicalOperator) As FilterStringBuilder

   Dim NewBuilder As FilterStringBuilder

   Set NewBuilder = New FilterStringBuilder
   Set NewBuilder.SqlTool = m_SqlTool
   NewBuilder.DefaultConcatOperator = ConcatOperator

   ConditionGroups.Add NewBuilder

   Set NewConditionGroup = NewBuilder

End Function

Private Function AppendFilterGroupsString(ByVal BaseFilterString As String, ByVal ConcatOperator As SqlLogicalOperator, _
                                 Optional ByVal IgnoreDuplicateFilters As Boolean = False) As String

   Dim CondPrefix As String
   Dim CondSuffix As String

   If m_ConditionGroups Is Nothing Then
      AppendFilterGroupsString = BaseFilterString
      Exit Function
   End If

   If m_ConditionGroups.Count = 0 Then
      AppendFilterGroupsString = BaseFilterString
      Exit Function
   End If

   If ConcatOperator <> SQL_CommaSeparator Then
      CondPrefix = "("
      CondSuffix = ")"
   End If

   Dim FSB As FilterStringBuilder

   With New StringCollection

      .Add BaseFilterString

      For Each FSB In m_ConditionGroups
         .Add FSB.ToString()
      Next

      AppendFilterGroupsString = .ToString(GetConcatOperatorString(ConcatOperator), CondPrefix, CondSuffix, True, IgnoreDuplicateFilters)

   End With

End Function

'---------------------------------------------------------------------------------------
' Function: AddSubSelectCriteria
'---------------------------------------------------------------------------------------
'
' New filter condition group for a sub select - e.g. ( a = 1 and x In (select n  from tab123) )
'
' Parameters:
'
'     FieldName - Data field name to compare
'     RelationalOperator - Relational operator for sub select: usually In(...)
'     SelectFromText - Sql text for sub select (without where, except UseWhereReplacementTextInFromText is used)
'     IgnoreIfSubSelectHasNoCriteria - True = Ignore subselect criterion if no filter values were transferred
'     UseWhereReplacementTextInFromText - Replace [Where] with Where condition string ... is required for Group By-SQL, for example.
'     SubSelectConcatOperator - [SQL_And (default), SQL_Or, SQL_CommaSeparator] .. valid within the new group
'
' Returns:
'
'     FilterStringBuilder instance for new sub select group
'
'---------------------------------------------------------------------------------------
Public Function AddSubSelectCriteria( _
               ByVal FieldName As String, _
               ByVal RelationalOperator As SqlRelationalOperators, _
               ByVal SelectFromText As String, _
               Optional ByVal IgnoreIfSubSelectHasNoCriteria As Boolean = False, _
               Optional ByVal UseWhereReplacementTextInFromText As Boolean = False, _
               Optional ByVal SubSelectConcatOperator As SqlLogicalOperator = SqlLogicalOperator.SQL_And _
      ) As FilterStringBuilder

   Dim NewBuilder As FilterStringBuilder
   Dim SqlPrefix As String
   Dim SqlSuffix As String
   Dim WhereReplacementPos As Long

   SqlPrefix = FieldName & " " & SqlTools.GetRelationalOperatorString(RelationalOperator) & " ("
   SqlSuffix = ")"
   If UseWhereReplacementTextInFromText Then
      WhereReplacementPos = InStr(1, SelectFromText, WhereReplacementText, vbTextCompare)
   End If

   If WhereReplacementPos > 0 Then
      SqlPrefix = SqlPrefix & Trim$(Left(SelectFromText, WhereReplacementPos - 1))
      SqlSuffix = " " & Trim$(Mid$(SelectFromText, WhereReplacementPos + Len(WhereReplacementText))) & SqlSuffix
   Else
      SqlPrefix = SqlPrefix & SelectFromText
   End If

   Set NewBuilder = New FilterStringBuilder
   NewBuilder.DefaultConcatOperator = SubSelectConcatOperator
   Set NewBuilder.SqlTool = m_SqlTool

   NewBuilder.ConfigSqlStatement SqlPrefix:=SqlPrefix, SqlSuffix:=SqlSuffix, _
                                 ConditionPrefix:=" Where ", IgnoreSqlPreSuffixIfEmptyFilter:=IgnoreIfSubSelectHasNoCriteria
   ConditionGroups.Add NewBuilder

   Set AddSubSelectCriteria = NewBuilder

End Function

'---------------------------------------------------------------------------------------
' Function: AddExistsCriteria
'---------------------------------------------------------------------------------------
'
' New filter condition group for a exits sub select - e.g. ( a = 1 and exists (select * from tab123 where t = a and y = 123) )
'
' Parameters:
'
'     SelectFromText - Sql text for sub select (without where, except UseWhereReplacementTextInFromText is used)
'     IgnoreIfExistsStatementHasNoCriteria - True = Ignore subselect criterion if no filter values were transferred
'     SubSelectConcatOperator - [SQL_And (default), SQL_Or, SQL_CommaSeparator] .. valid within the new group
'     UseNotExists - use not exists (..): default: false = exists (...)
'
' Returns:
'
'     FilterStringBuilder instance for new exists group
'
'---------------------------------------------------------------------------------------
Public Function AddExistsCriteria( _
                        ByVal SelectFromText As String, _
               Optional ByVal IgnoreIfExistsStatementHasNoCriteria As Boolean = False, _
               Optional ByVal SubSelectConcatOperator As SqlLogicalOperator = SqlLogicalOperator.SQL_And, _
               Optional ByVal UseNotExists As Boolean = False, _
               Optional ByVal UseWhereReplacementTextInFromText As Boolean = False _
      ) As FilterStringBuilder

   Dim NewBuilder As FilterStringBuilder
   Dim ExistsSqlPrefix As String
   Dim WhereReplacementPos As Long
   Dim ExistsSqlSuffix As String

   ExistsSqlPrefix = "Exists ("
   If UseNotExists Then ExistsSqlPrefix = "Not " & ExistsSqlPrefix

   If UseWhereReplacementTextInFromText Then
      WhereReplacementPos = InStr(1, SelectFromText, WhereReplacementText, vbTextCompare)
   End If

   If WhereReplacementPos > 0 Then
      ExistsSqlSuffix = " " & Trim$(Mid$(SelectFromText, WhereReplacementPos + Len(WhereReplacementText))) & ")"
      SelectFromText = Trim$(Left(SelectFromText, WhereReplacementPos - 1))
   Else
      ExistsSqlSuffix = ")"
   End If

   Set NewBuilder = New FilterStringBuilder
   NewBuilder.DefaultConcatOperator = SubSelectConcatOperator
   Set NewBuilder.SqlTool = m_SqlTool

   NewBuilder.ConfigSqlStatement SqlPrefix:=ExistsSqlPrefix & SelectFromText, _
                                 SqlSuffix:=ExistsSqlSuffix, ConditionPrefix:=" Where ", IgnoreSqlPreSuffixIfEmptyFilter:=IgnoreIfExistsStatementHasNoCriteria

   ConditionGroups.Add NewBuilder

   Set AddExistsCriteria = NewBuilder

End Function

'##################################
' Group: Output

'---------------------------------------------------------------------------------------
' Function: ToString
'---------------------------------------------------------------------------------------
'
' Output criteria to String
'
' Parameters:
'
'     ConcatOperator - [SQL_And (default), SQL_Or, SQL_CommaSeparator]
'     IgnoreDuplicateFilters - Do not output duplicate filter criteria
'
' Returns:
'
'     Criteria/filter string
'
'---------------------------------------------------------------------------------------
Public Function ToString(Optional ByVal ConcatOperator As SqlLogicalOperator = SqlLogicalOperator.[_SQL_Default], _
                         Optional ByVal IgnoreDuplicateFilters As Boolean = False) As String

   Dim FilterString As String
   Dim ItemPrefix As String
   Dim ItemSuffix As String

   If ConcatOperator = SqlLogicalOperator.[_SQL_Default] Then
      ConcatOperator = DefaultConcatOperator
   End If

   If ConcatOperator <> SQL_CommaSeparator Then
      ItemPrefix = "("
      ItemSuffix = ")"
   End If

   FilterString = m_Items.ToString(GetConcatOperatorString(ConcatOperator), ItemPrefix, ItemSuffix, , IgnoreDuplicateFilters)
   FilterString = AppendFilterGroupsString(FilterString, ConcatOperator, IgnoreDuplicateFilters)
   If Len(FilterString) > 0 Then
      FilterString = m_ConditionPrefix & FilterString
   End If

   If m_IgnoreSqlPreSuffixIfEmptyFilter Then
   If Len(FilterString) = 0 Then
      ToString = vbNullString
      Exit Function
   End If
   End If

   ToString = m_SqlPrefix & FilterString & m_SqlSuffix

End Function

Private Function GetConcatOperatorString(ByVal ConcatOperator As SqlLogicalOperator) As String

   Select Case ConcatOperator
      Case SqlLogicalOperator.SQL_And
         GetConcatOperatorString = " And "
      Case SqlLogicalOperator.SQL_Or
         GetConcatOperatorString = " Or "
      Case SqlLogicalOperator.SQL_CommaSeparator
         GetConcatOperatorString = ", "
      Case Else

   End Select

End Function

Private Property Get ConditionGroups() As Collection
   If m_ConditionGroups Is Nothing Then
      Set m_ConditionGroups = New Collection
   End If
   Set ConditionGroups = m_ConditionGroups
End Property
