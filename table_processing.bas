Attribute VB_Name = "table_processing"
'Module for processing Table 1 sheet and perpare it for the last step which is filtering and row deletion
Sub table_processing()
'error handling to jump to the the code block "error_handle"
On Error GoTo error_handle

    Dim LastCell As Range
    Dim LastCellRowIndex As Long
    Dim i, j As Integer
    
    'stop updating screen and alerts to enhance performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'unmerge all the cells in the sheet "Table 1"
    Sheets("Table 1").Cells.UnMerge
    
    'find the last cell in the column A and then save its row index
    Set LastCell = Sheets("Table 1").Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    LastCellRowIndex = LastCell.Row
    
    'join the cells in range from A to DC per row to handle the data separated in cloumns and keep everything in one column DD
    For i = 1 To LastCellRowIndex
    
        'This executes a function from the worksheet called TEXTJOIN
        Sheets("Table 1").Range("DD" & i).Value = WorksheetFunction.TextJoin(" ", True, Sheets("Table 1").Range("A" & i & ":DC" & i))
    
    Next
    
    'copy column DD to column A and delete all the rest to prepare for next step
    Sheets("Table 1").Range("DD:DD").Copy
    Sheets("Table 1").Range("A:A").PasteSpecial Paste:=xlPasteValues
    Sheets("Table 1").Range("B:DD").Delete
    
    'Apply text to column with a new line character delimiter to separate each clause in a separate column
    Sheets("Table 1").Range("A1:A" & LastCellRowIndex).TextToColumns DataType:=xlDelimited, Other:=True, OtherChar:=vbLf
    
    'addiing a new sheet to paste the separated values in it and it will be the output sheet
    Sheets.Add.Name = "Specs"
    
    'iterate through the columns of each row and copy the cells in transpose position to rearrange the clauses
    i = 2
    For j = 1 To LastCellRowIndex
    
        Sheets("Table 1").Range("A" & j & ":DD" & j).Copy
        Sheets("Specs").Range("A" & i).PasteSpecial Transpose:=True
        'the number of columns from A to DD is 108
        i = i + 108
    
    Next
    
    'add a header to enable filtering in the next step
    Sheets("Specs").Range("A1").Value = "Clause"
    
    'delete the Table 1 sheet as it is not needed anymore
    Sheets("Table 1").Delete
    
    'reactivate updating screen and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    'go to the next module
    Call formatting.formatting
    
    'it is important to exit the sub before the error handling block or it will be executed
    Exit Sub

'error handling code block
error_handle:

    'massege to the user that there has been an error in Table 1 processing
    MsgBox ("an error has occured in processing the excel file converted from pdf.")
    
    'restore the sheet to normal by deleting the added sheets. error catch is to avoid deleting nonexisting sheet
    For Each ws In Worksheets
        If ws.Name = "Table 1" Then
            Application.DisplayAlerts = False
            Sheets("Table 1").Delete
            Application.DisplayAlerts = True
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
