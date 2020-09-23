Attribute VB_Name = "Sort_String"
Option Explicit
Public Function ShellSortLong(SortArray() As String, Optional ByVal IgnoreCase As Boolean = True) As Long()
Dim sVal1 As String, sVal2 As String
Dim IndexArray() As Long
Dim idx As Long, Row As Long, MaxRow As Long, MinRow As Long
Dim Swtch As Long, Limit As Long, Offset As Long

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
ReDim IndexArray(MinRow To MaxRow)
For idx = MinRow To MaxRow
    IndexArray(idx) = idx
Next

Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         'Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
                sVal1 = SortArray(IndexArray(Row))
                sVal2 = SortArray(IndexArray(Row + Offset))
                If IgnoreCase Then
                    sVal1 = LCase(sVal1)
                    sVal2 = LCase(sVal2)
                End If
                If sVal1 > sVal2 Then
                   SwapLongs IndexArray(Row), IndexArray(Row + Offset)
                   Swtch = Row
                End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop

ShellSortLong = IndexArray

End Function
Private Sub SwapLongs(ByRef var1 As Long, ByRef var2 As Long)

    Dim X As Long
    X = var1
    var1 = var2
    var2 = X

End Sub
