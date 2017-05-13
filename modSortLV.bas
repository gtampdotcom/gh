Attribute VB_Name = "modSortLV"
Option Explicit

Public Function SortLV(ListView As Object, ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo oops
    ' Record the starting CPU time (milliseconds since boot-up)
    Dim lngStart As Long
    lngStart = GetTickCount
   
    ' Commence sorting
    With ListView
    
        ' Display the hourglass cursor whilst sorting
        
        Dim lngCursor As Long
        lngCursor = .MousePointer
        .MousePointer = vbHourglass
        
        ' Prevent the lvGH control from updating on screen -
        ' this is to hide the changes being made to the listitems
        ' and also to speed up the sort
        
        'LockWindowUpdate .hwnd
        
        ' Check the data type of the column being sorted,
        ' and act accordingly
        
        'Dim l As Long
        'Dim strFormat As String
        'Dim strData() As String
        
        Dim lngIndex As Long
        lngIndex = ColumnHeader.Index - 1
        
        ' Sort alphabetically. This is the only sort provided
            ' by the MS lvGH control (at this time), and as
            ' such we don't really need to do much here
        
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
        'End Select
    
        ' Unlock the list window so that the OCX can update it
        
        'LockWindowUpdate 0&
        
        ' Restore the previous cursor
        
        .MousePointer = lngCursor
    
    End With
    
    '------- Show Icon in Heading --------------------------------
    ShowListViewColumnHeaderSortIcon ListView   'show the icon
    '-------------------------------------------------------------
    Exit Function
oops:
    displaychat strChannel, strGHColor, "Error sorting ListView"
End Function

Public Sub SortColumn(ListView As Object, ByVal intCurrentSortColumn As Integer)
    Dim col As ColumnHeader
    Set col = ListView.ColumnHeaders(intCurrentSortColumn) 'sort on this column
    If ListView.SortOrder = lvwAscending Then
        ListView.SortOrder = lvwDescending
    Else
        ListView.SortOrder = lvwAscending
    End If
    Call SortLV(ListView, col)
End Sub
