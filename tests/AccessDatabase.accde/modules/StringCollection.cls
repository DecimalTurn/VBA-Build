VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: text.StringCollection
'---------------------------------------------------------------------------------------
'
' Collection for strings
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text/StringCollection.cls</file>
'  <license>_codelib/license.bas</license>
'  <test>_test/text/StringCollectionTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private m_Items As Collection

Private Sub Class_Initialize()
   Set m_Items = New Collection
End Sub

Private Sub Class_Terminate()
   Set m_Items = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Property: Self
'---------------------------------------------------------------------------------------
'
' Reference to self (Me)
'
' Remarks:
'     Useful for with-block
'
' Returns:
'     Database.StringCollection
'
'---------------------------------------------------------------------------------------
Public Property Get Self() As StringCollection
   Set Self = Me
End Property

'---------------------------------------------------------------------------------------
' Property: Items
'---------------------------------------------------------------------------------------
'
' Collection with items
'
' Returns:
'     VBA.Collection
'
'---------------------------------------------------------------------------------------
Public Property Get Items() As Collection
   Set Items = m_Items
End Property

'---------------------------------------------------------------------------------------
' Property: Item
'---------------------------------------------------------------------------------------
'
' Item of Collection
'
' Parameters:
'     Index - (Variant)
'
' Returns:
'     Item string - (String)
'
'---------------------------------------------------------------------------------------
Public Property Get Item(ByVal Index As Variant) As String
   Item = m_Items.Item(Index)
End Property

Public Property Let Item(ByVal Index As Variant, ByVal NewValue As String)
Attribute Item.VB_UserMemId = 0
   m_Items.Add NewValue, , , Index
   m_Items.Remove Index
End Property

'---------------------------------------------------------------------------------------
' Sub: Add
'---------------------------------------------------------------------------------------
'
' Add string to collection
'
' Parameters:
'     Item to add - (String)
'
'---------------------------------------------------------------------------------------
Public Sub Add(ByVal Item As String)
   m_Items.Add Item
End Sub

'---------------------------------------------------------------------------------------
' Sub: AddFromArray
'---------------------------------------------------------------------------------------
'
' Add items form an array to collection
'
' Parameters:
'     ArrayToAdd - (Variant)
'     ItemStringFormat - (String) Format each item of Array with ItemStringFormat before add to collection
'
'---------------------------------------------------------------------------------------
Public Sub AddFromArray(ByRef ArrayToAdd As Variant, Optional ByVal ItemStringFormat As String = vbNullString)

   Dim i As Long

   For i = LBound(ArrayToAdd) To UBound(ArrayToAdd)
      m_Items.Add Format(ArrayToAdd(i), ItemStringFormat)
   Next

End Sub

'---------------------------------------------------------------------------------------
' Sub: AddFromCollection
'---------------------------------------------------------------------------------------
'
' Add items form a collection to string collection
'
' Parameters:
'     CollectionToAppend - (Object) .. so that all collections with Enumarable and Item(index) interface can be run through
'     ItemStringFormat   - (String) Format each item of collection with ItemStringFormat before add to collection
'
'---------------------------------------------------------------------------------------
Public Sub AddFromCollection(ByVal CollectionToAppend As Object, Optional ByVal ItemStringFormat As String = vbNullString)

   Dim itm As Variant

   For Each itm In CollectionToAppend
      m_Items.Add Format(itm, ItemStringFormat)
   Next

End Sub

'---------------------------------------------------------------------------------------
' Function: ToString
'---------------------------------------------------------------------------------------
'
' Return Collection items as joined String
'
' Parameters:
'     Delimiter             - (String) Example: ", " => "Item1, Item2, Item3"
'     ItemPrefix            - (String) Prefix for each item
'     ItemSuffix            - (String) Suffix for each item
'     IgnoreEmptyValue      - (Boolean) don't output an empty item
'     IgnoreDuplicateValues - (Boolean) True = don't output duplicate items
'
' Returns:
'     String
'
'---------------------------------------------------------------------------------------
Public Function ToString(Optional ByVal Delimiter As String = ", ", _
                         Optional ByVal ItemPrefix As String = vbNullString, _
                         Optional ByVal ItemSuffix As String = vbNullString, _
                         Optional ByVal IgnoreEmptyValue As Boolean = False, _
                         Optional ByVal IgnoreDuplicateValues As Boolean = False) As String

   Dim s As String

   s = VBA.Join(ToStringArray(IgnoreEmptyValue, IgnoreDuplicateValues), ItemSuffix & Delimiter & ItemPrefix)
   If Len(s) > 0 Then s = ItemPrefix & s & ItemSuffix

   ToString = s

End Function

'---------------------------------------------------------------------------------------
' Function: ToStringArray
'---------------------------------------------------------------------------------------
'
' Return Collection items as String array
'
' Parameters:
'     IgnoreEmptyValue      - (Boolean) don't output an empty item
'     IgnoreDuplicateValues - (Boolean) True = don't output duplicate items
'
' Returns:
'     String array
'
'---------------------------------------------------------------------------------------
Public Function ToStringArray(Optional ByVal IgnoreEmptyValue As Boolean = False, _
                              Optional ByVal IgnoreDuplicateValues As Boolean = False) As String()

   Dim ItemArray() As String
   Dim MaxArrayIndex As Long
   Dim i As Long

   MaxArrayIndex = m_Items.Count - 1

   If MaxArrayIndex < 0 Then
      ToStringArray = ItemArray
      Exit Function
   End If

   If IgnoreEmptyValue Then
      If IgnoreDuplicateValues Then
         ToStringArray = RemoveDuplicateValues(GetArrayWithoutEmptyValues())
      Else
         ToStringArray = GetArrayWithoutEmptyValues()
      End If
      Exit Function
   End If

   ReDim ItemArray(0 To MaxArrayIndex)
   For i = 0 To MaxArrayIndex
      ItemArray(i) = m_Items.Item(i + 1)
   Next

   If IgnoreDuplicateValues Then
      ToStringArray = RemoveDuplicateValues(ItemArray)
   Else
      ToStringArray = ItemArray
   End If

End Function

Private Function GetArrayWithoutEmptyValues() As String()

   Dim ItemArray() As String
   Dim MaxArrayIndex As Long
   Dim ItemIndex As Long
   Dim itm As Variant

   MaxArrayIndex = m_Items.Count - 1

   If MaxArrayIndex < 0 Then
      GetArrayWithoutEmptyValues = ItemArray
      Exit Function
   End If

   ReDim ItemArray(0 To MaxArrayIndex)
   ItemIndex = -1
   For Each itm In m_Items
      If Len(itm) > 0 Then
         ItemIndex = ItemIndex + 1
         ItemArray(ItemIndex) = itm
      End If
   Next

   If ItemIndex = -1 Then
      Erase ItemArray
      GetArrayWithoutEmptyValues = ItemArray
      Exit Function
   End If

   If ItemIndex < (m_Items.Count - 1) Then
      ReDim Preserve ItemArray(0 To ItemIndex)
   End If

   GetArrayWithoutEmptyValues = ItemArray

End Function

Private Function RemoveDuplicateValues(ByRef ArrayToCheck() As String) As String()

   Dim ItemArray() As String
   Dim MaxArrayIndex As Long
   Dim ItemIndex As Long
   Dim ArrayItem As Variant

   MaxArrayIndex = UBound(ArrayToCheck)

   If MaxArrayIndex = 0 Then
      RemoveDuplicateValues = ArrayToCheck
      Exit Function
   End If

   ReDim ItemArray(MaxArrayIndex)

   ItemIndex = -1
   For Each ArrayItem In ArrayToCheck
      If Not ValueExistsInArray(ItemArray, ArrayItem, ItemIndex) Then
         ItemIndex = ItemIndex + 1
         ItemArray(ItemIndex) = ArrayItem
      End If
   Next

   If ItemIndex < (m_Items.Count - 1) Then
      ReDim Preserve ItemArray(0 To ItemIndex)
   End If

   RemoveDuplicateValues = ItemArray

End Function

Private Function ValueExistsInArray(ByRef ArrayToCheck() As String, ByVal ValueToCheck As String, ByVal CheckUntilArrayIndex As Long) As Boolean

   Dim i As Long

   If CheckUntilArrayIndex < 0 Then
      Exit Function
   End If

   For i = LBound(ArrayToCheck) To CheckUntilArrayIndex
      If StrComp(ArrayToCheck(i), ValueToCheck, vbBinaryCompare) = 0 Then
         ValueExistsInArray = True
         Exit Function
      End If
   Next

End Function
