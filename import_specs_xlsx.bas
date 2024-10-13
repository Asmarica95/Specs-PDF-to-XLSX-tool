Attribute VB_Name = "import_specs_xlsx"
'module for importing the sheet "Table 1" from the excel file that was converted from pdf using Adobe Acrobat
'by selecting the designated file and copy the sheet with the name "Table 1" to be the first sheet of this workbook
Sub import_specs_xlsx()

'error handling to jump to the the code block "error_handle"
On Error GoTo error_handle

    'variables definition
    Dim fd As Office.FileDialog
    Dim Filepath As String
    Dim closedBook As Workbook
        
    'define the dialig fd as file picker type to allow user to select a file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'specify the fd properties
    With fd
          .AllowMultiSelect = False 'we need to allow the selection of one file only
          .Title = "Please select the excel file that was converted from pdf using Adobe Acrobat" 'dialog title
        
          'checks for successful dialog show
          If .Show = True Then
            Filepath = .SelectedItems(1) 'save the path of the first selected item in the dialog to Filepath
          End If
    End With
    
    'stop updating screen and alerts to enhance performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'select the designated workbook and copy the sheet "Table 1" from it then close it
    Set closedBook = Workbooks.Open(Filepath)
    closedBook.Sheets("Table 1").Copy Before:=ThisWorkbook.Sheets(1)
    closedBook.Close SaveChanges:=False
    
    'reactivate updating screen and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    'go to the next module
    Call table_processing.table_processing
    
    'it is important to exit the sub before the error handling block or it will be executed
    Exit Sub

'error handling code block
error_handle:

    'massege to the user that there has been an error in the process of copying the sheet from the disgnated file
    MsgBox ("an error has occured in importing specs from excel sheet.")
    
    'restore the sheet to normal by closing the opened workbookonly if it was instatiated
    If Not closedBook Is Nothing Then
        closedBook.Close SaveChanges:=False
    End If
    
    'reactivate updating screen and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
 
