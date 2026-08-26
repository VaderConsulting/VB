VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Sort the items in array values() with bounds min and max.
Sub Quicksort(values() As String, ByVal min As Long, ByVal max As Long)

  Dim med_value As String
  Dim hi As Long
  Dim lo As Long
  Dim i As Long

  ' If the list has only 1 item, it's sorted.
  If min >= max Then Exit Sub

  ' Pick a dividing item randomly.
  i = min + Int(Rnd(max - min + 1))
  med_value = values(i)

  ' Swap the dividing item to the front of the list.
  values(i) = values(min)

  ' Separate the list into sublists.
  lo = min
  hi = max
  Do
    ' Look down from hi for a value < med_value.
    Do While values(hi) >= med_value
      hi = hi - 1
      If hi <= lo Then Exit Do
    Loop

    If hi <= lo Then
      ' The list is separated.
      values(lo) = med_value
      Exit Do
    End If

    ' Swap the lo and hi values.
    values(lo) = values(hi)

    ' Look up from lo for a value >= med_value.
    lo = lo + 1
    Do While values(lo) < med_value
      lo = lo + 1
      If lo >= hi Then Exit Do
    Loop

    If lo >= hi Then
      ' The list is separated.
      lo = hi
      values(hi) = med_value
      Exit Do
    End If

    ' Swap the lo and hi values.
    values(hi) = values(lo)
  Loop ' Loop until the list is separated.

  ' Recursively sort the sublists.
  Quicksort values, min, lo - 1
  Quicksort values, lo + 1, max

End Sub

