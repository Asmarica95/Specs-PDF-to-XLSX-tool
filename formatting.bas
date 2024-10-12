Attribute VB_Name = "formatting"
Sub formatting()
    
    Dim lastRow As Long
    
    'stop updating screen and alerts to enhance performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    With Sheets("Specs")

        lastRow = .Range("A" & Rows.Count).End(xlUp).Row

        '~~> Remove any filters
        .AutoFilterMode = False

        '~~> Filter, offset(to exclude headers) and delete visible rows
        With .Range("A1:A" & lastRow)
            .AutoFilter Field:=1, Criteria1:="=", Operator:=xlOr, Criteria2:="0"
            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End With

        '~~> Remove any filters
        .AutoFilterMode = False
    End With
    
    'reactivate updating screen and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
