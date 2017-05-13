VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSortArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Used for sorting arrays to make them match the way GTA2 sorts maps

Option Explicit

Private IsSorted As Boolean
Private ItemArray() As Variant
Private idxArray() As Long
Public CaseSensitive As Boolean
Private CurrentIndex As Long
Private lCount As Long
Public Sub AddItem(ByVal Item As Variant)
ReDim Preserve ItemArray(0 To lCount)
ItemArray(lCount) = Item
CurrentIndex = lCount
lCount = lCount + 1
End Sub
Public Sub RemoveItem(Optional ByVal Index As Long)
Dim iCnt As Long
If IsMissing(Index) Then
Else
    CurrentIndex = Index
End If
For iCnt = CurrentIndex To lCount - 2
    ItemArray(iCnt) = ItemArray(iCnt + 1)
Next iCnt
lCount = lCount - 1
ReDim Preserve ItemArray(0 To lCount - 1)
End Sub
'------------------------------------------
'Sorting and Searching
'-------------------------------------------
Public Property Get count() As Long
count = lCount
End Property
Public Property Get SortedItem(ByVal idx As Long) As Variant
If idx > lCount - 1 Then
    displaychat strDestTab, vbRed, "Error sorting maps: " & idx & " > " & lCount
    Exit Sub
End If
If IsSorted Then
Else
    Sort
End If
SortedItem = ItemArray(idxArray(idx))
CurrentIndex = idxArray(idx)
End Property
Public Property Let SortedItem(ByVal idx As Long, ByVal x As Variant)
If idx > lCount - 1 Then
    Err.Raise 9
End If
If IsSorted Then
Else
    Sort
End If
ItemArray(idxArray(idx)) = x
Sort
CurrentIndex = SearchArray(x)
End Property

Public Property Get Item(ByVal idx As Long) As Variant
If idx > lCount - 1 Then
    Err.Raise 9
End If
Item = ItemArray(idx)
CurrentIndex = idx

End Property
Public Property Let Item(ByVal idx As Long, vValue As Variant)
If idx > lCount - 1 Then
    ReDim Preserve ItemArray(0 To idx)
    lCount = idx + 1
End If
ItemArray(idx) = vValue
CurrentIndex = idx
End Property

Private Sub Sort()

Dim NewArray() As String
Dim iCnt As Integer

If CaseSensitive Or Not (VarType(ItemArray(0)) = vbString) Then
    QuickSort ItemArray, idxArray
Else
    ReDim NewArray(0 To lCount - 1)
    For iCnt = 0 To lCount - 1
        NewArray(iCnt) = UCase(ItemArray(iCnt))
    Next iCnt
    QuickSort NewArray, idxArray
End If
IsSorted = True
End Sub
Private Sub QuickSort _
    (ByVal iArray As Variant, _
    ByRef Index As Variant, _
    Optional StartPoint As Long, _
    Optional EndPoint As Long)
If UBound(iArray) = LBound(iArray) Then
    ReDim Index(0 To lCount - 1)
    Index(0) = 0
    Exit Sub
End If
    
    Dim tempStart As Long
    Dim tempEnd As Long
    Dim MidPoint As Variant
    Dim y As Long
    Dim lCnt As Long
    
    'Initialize start and end points, if not already.
    If StartPoint = 0 And EndPoint = 0 Then
        StartPoint = 0
        EndPoint = lCount - 1
        ReDim Index(StartPoint To EndPoint)
    End If
    
    'Initialize Index, if not already initialized
    If Index(StartPoint) = 0 And Index(EndPoint) = 0 And Not (StartPoint = EndPoint) Then
    'If Index(StartPoint) = 0 And Index(EndPoint) = 0 Then
        For lCnt = StartPoint To EndPoint
            Index(lCnt) = lCnt
        Next lCnt
    End If
    
    tempStart = StartPoint
    tempEnd = EndPoint
    MidPoint = iArray(Index((StartPoint + EndPoint) / 2))

    While (tempStart <= tempEnd)

        While (iArray(Index(tempStart)) < MidPoint And tempStart < EndPoint)
            tempStart = tempStart + 1
        Wend

        While (MidPoint < iArray(Index(tempEnd)) And tempEnd > StartPoint)
            tempEnd = tempEnd - 1
        Wend

        If (tempStart <= tempEnd) Then
            y = Index(tempStart)
            Index(tempStart) = Index(tempEnd)
            Index(tempEnd) = y
            tempStart = tempStart + 1
            tempEnd = tempEnd - 1
            
        End If
        
    Wend
    
    If (StartPoint < tempEnd) Then
        QuickSort iArray, Index, StartPoint, tempEnd
    End If
    If (tempStart < EndPoint) Then
        QuickSort iArray, Index, tempStart, EndPoint
    End If
End Sub

Public Function SearchArray _
    (ByVal SearchValue As Variant, _
    Optional ByVal boolCaseSensitive As Boolean = True, _
    Optional ByVal UseWildcards As Boolean = False) As Long
Dim min As Integer
Dim max As Integer
Dim pnt As Integer
Dim Bomb As Integer
Dim WildCard As Boolean
Dim found As Boolean
Dim p As Integer
Dim srch As String
If IsSorted Then
Else
    Sort
End If
If UseWildcards And (InStr(SearchValue, "*") > 0 Or InStr(SearchValue, "?") > 0) Then
    WildCard = True
End If
min = 0
max = lCount - 1

If WildCard Then
    p = InStr(SearchValue, "*")
    For pnt = min To max
        If p > 1 Then
            srch = ResolveQuestion(Left(SearchValue, p - 1), ItemArray(idxArray(pnt))) & "*" & ResolveQuestion(Mid(SearchValue, p + 1), ItemArray(idxArray(pnt)))
        Else
            If p > 0 Then
                srch = "*"
            Else
                srch = vbNullString
            End If
            srch = srch & ResolveQuestion(Mid(SearchValue, p + 1), ItemArray(idxArray(pnt)))
        End If
        found = CheckMatch(srch, ItemArray(idxArray(pnt)))
        If found Then
            SearchArray = pnt
            Exit Function
        End If
    Next pnt
End If

Bomb = min - 1
pnt = Int((max - min) / 2 + min)
'If SearchValue = "htmlfile" Then Stop
Do Until max = min _
    Or SearchValue = ItemArray(idxArray(pnt)) _
    Or found _
    Or (UCase(SearchValue) = UCase(ItemArray(idxArray(pnt))) And _
        Not boolCaseSensitive)
    If WildCard Then
        found = CheckMatch(SearchValue, ItemArray(idxArray(pnt)))
        If found Then
            Exit Do
        End If
    End If
        
    If SearchValue > ItemArray(idxArray(pnt)) _
        Or (UCase(SearchValue) > UCase(ItemArray(idxArray(pnt))) And _
            Not boolCaseSensitive) Then
        min = pnt
    
    Else
        max = pnt
    End If
    If min = max - 1 Then
        'Max = Max - 1
        If pnt = min Then
            min = min + 1
        Else
            max = max - 1
        End If
        pnt = min
    Else
        pnt = Int((max - min) / 2 + min)
    End If
Loop

If SearchValue = ItemArray(idxArray(pnt)) Or (WildCard And found) _
    Or (UCase(SearchValue) = UCase(ItemArray(idxArray(pnt))) And _
        Not boolCaseSensitive) Then
    SearchArray = pnt
Else
    SearchArray = Bomb
End If

End Function
Private Function ResolveQuestion(ByVal txt As String, ByVal base As String) As String
Dim p As Integer
p = InStr(txt, "?")
Do While p > 0
    If p = 1 Then
        txt = Left(base, 1) & txt
    Else
        txt = Left(txt, p - 1) & Mid(base, p, 1) & Mid(txt, p + 1)
    End If
    p = InStr(txt, "?")
Loop
ResolveQuestion = txt
End Function
Private Function CheckMatch(ByVal SearchValue As String, ByVal BaseValue As String) As Boolean
Dim p As Integer
p = InStr(SearchValue, "*")
If p > 1 Then
    SearchValue = ResolveQuestion(Left(SearchValue, p - 1), Left(BaseValue, p - 1)) & Mid(SearchValue, p)
    If Left(SearchValue, p - 1) = Left(BaseValue, p - 1) Then
        If Mid(SearchValue, p + 1) = Right(BaseValue, Len(Mid(SearchValue, p + 1))) Then
            CheckMatch = True
            Exit Function
        Else
            CheckMatch = False
            Exit Function
        End If
    Else
        CheckMatch = False
        Exit Function
    End If
Else
    SearchValue = ResolveQuestion(SearchValue, BaseValue)
    If Right(SearchValue, Len(SearchValue) - p) = Right(BaseValue, Len(SearchValue) - p) Then
        CheckMatch = True
        Exit Function
    Else
        CheckMatch = False
        Exit Function
    End If
End If
End Function

Public Property Get SortedToItemIndex(ByVal SortedIndex As Long) As Long
If IsSorted Then
Else
    Sort
End If
SortedToItemIndex = idxArray(SortedIndex)
End Property
Public Property Get ItemToSortedIndex(ByVal Index As Long) As Long
ItemToSortedIndex = SearchArray(ItemArray(Index), CaseSensitive)
End Property

Private Sub Class_Initialize()
CaseSensitive = True
End Sub
