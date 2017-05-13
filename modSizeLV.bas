Attribute VB_Name = "modSizeLV"
Option Explicit

'Resizes the columns in any ListView.
'Returns the total width of the columns at their final size.
Public Function AutoSizeLV(LV As ListView) As Single
On Error GoTo oops
'Passing ListView control arrays wouldn't work. Got fed up and hacked around it.

'Variables:
Dim lngCol As Long, lngRow As Long
Dim arrWidest() As Single '(Col): Width of widest item in each column.
Dim arrCurrent() As Single '(Col): Width of items before resizing.
Dim sngThis As Single 'Width of current item
Dim sngWidth As Single 'total width of columns

With LV
  'Initialise arrays:
  ReDim arrCurrent(1 To .ColumnHeaders.count)
  ReDim arrWidest(1 To .ColumnHeaders.count)
  '1-based for consistency.

    'Start with column header Widths:
    For lngCol = LBound(arrWidest) To UBound(arrWidest)
        arrCurrent(lngCol) = .ColumnHeaders(lngCol).Width 'measure column
        arrWidest(lngCol) = frmGH.TextWidth(.ColumnHeaders(lngCol)) 'measure text
    Next
    'ColumnHeaders aren't counted as ListItems.
    
    'Check item contents, if there are any:
    If .ListItems.count > 0 Then
        'For each row:
        For lngRow = 1 To .ListItems.count
            '1st column:
            sngThis = frmGH.TextWidth(.ListItems(lngRow)) 'measure text
            If arrWidest(1) < sngThis Then arrWidest(1) = sngThis
    
            'Subsequent columns:
            For lngCol = 2 To .ColumnHeaders.count
                'Is column visible?
                If .ColumnHeaders(lngCol).Width > 0 Then
                    sngThis = frmGH.TextWidth(.ListItems(lngRow).SubItems(lngCol - 1)) 'measure text
                    If arrWidest(lngCol) < sngThis Then arrWidest(lngCol) = sngThis
                End If
            Next
        Next
    End If
    
    'Apply column Widths:
    For lngCol = LBound(arrWidest) To UBound(arrWidest)
        'Is column visible?
        If .ColumnHeaders(lngCol).Width > 0 Then
            'Is first column?
            If lngCol = 1 Then
                'Has items?
                If .ListItems.count > 0 Then
                    arrWidest(lngCol) = arrWidest(lngCol) + 180 'flag Width
                End If
            End If
    
            'Has sorting arrow?
            If lngCol = .SortKey + 1 Then
                'Is header+arrow wider than widest item?
                If arrWidest(lngCol) < (frmGH.TextWidth(.ColumnHeaders(lngCol)) + 285) Then
                    arrWidest(lngCol) = frmGH.TextWidth(.ColumnHeaders(lngCol)) + 285
                End If
            End If
    
            'Apply Width only if it will be different:
            arrWidest(lngCol) = arrWidest(lngCol) + 180 'normal padding (420 text needs 600 Width.)
            If CLng(.ColumnHeaders(lngCol).Width) <> CLng(arrWidest(lngCol)) Then
                .ColumnHeaders(lngCol).Width = arrWidest(lngCol)
            End If
        End If
    Next
    
    'Add up width of all columns:
    For lngCol = 1 To .ColumnHeaders.count
        sngWidth = sngWidth + .ColumnHeaders(lngCol).Width
    Next
End With
  
'Return result:
AutoSizeLV = sngWidth

   Exit Function
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, strTextColor, "AutoSizeLV error: Line " & strErrLine & " " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :AutoSizeLV error: Line " & strErrLine & " " & strErrdesc
End Function
