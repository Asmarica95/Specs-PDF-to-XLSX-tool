Attribute VB_Name = "formatting"
'Module for formatting the processed table and exporting it to a new workbook
Sub formatting()

'error handling to jump to the the code block "error_handle"
On Error GoTo error_handle

    Dim lastRow As Long
    
    'stop updating screen and alerts to enhance performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    With Sheets("Specs")

        lastRow = .Range("A" & Rows.Count).End(xlUp).Row

        '~~> Remove any filters
        .AutoFilterMode = False

        '~~> Filter, offset(to exclude headers) and delete visible rows with blank or zero values
        With .Range("A1:A" & lastRow)
            .AutoFilter Field:=1, Criteria1:="=", Operator:=xlOr, Criteria2:="0"
            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End With

        '~~> Remove any filters
        .AutoFilterMode = False
    End With
    
    'formatting and export
    With Sheets("Specs")

        lastRow = .Range("A" & Rows.Count).End(xlUp).Row
        .Range("A1:B" & lastRow).Font.Name = "Aptos Narrow"
        .Range("A1:B" & lastRow).Font.Size = 10
        .Columns("A").ColumnWidth = 100
        .Columns("A").WrapText = True
        .Columns("B").ColumnWidth = 50
        .Range("B1").Value = "alfanar Reply"
        .Range("A1:B" & lastRow).Borders.LineStyle = xlContinuous
        .Range("A1:B1").Font.Bold = True
        .Range("B2:B" & lastRow).Font.Bold = True
        .Range("B1:B" & lastRow).Font.Color = RGB(14, 92, 220)
        .Range("A1:B" & lastRow).VerticalAlignment = xlCenter
        .Move
        
    End With
    
    'reactivate updating screen and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    'it is important to exit the sub before the error handling block or it will be executed
    Exit Sub

'error handling code block
error_handle:

    'massege to the user that there has been an error in Table 1 processing
    MsgBox ("an error has occured in formatting the excel file converted from pdf.")
    
    'restore the sheet to normal by deleting the added sheets. error catch is to avoid deleting nonexisting sheet
    For Each ws In Worksheets
        ElseIf ws.Name = "Specs" Then
            Application.DisplayAlerts = False
            Sheets("Specs").Delete
            Application.DisplayAlerts = True
        End If
    Next

    'reactivate updating screen and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    

End Sub
