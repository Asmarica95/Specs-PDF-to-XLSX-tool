# Specs-PDF-to-XLSX-tool
1- pdf with a compliance using type text field is a hopeless case. Acrobat does not handle the additional text well and every line with an added type text field is basically ruined by either adding the type text content with the same line of the clause or separating the content to multiple columns ahead. commented word files converted to pdf are also hectic but nothing compared to the previously mentioned disaster as their main issue is the production of shapes (can be handled by selection pane or code maybe?).

2- Acrobat mainly converts each page in the pdf to one combined big cell in excel. the extent of this big cell includes combined columns all the way to provide a good view point of the big cell contents. These big cells unmerge easily keeping all the text in the first column. some anomalies show in the content including deleted spaces causing words to glue to each other (but nothing that is intolerable or will never happen when the conversion is done manually by a human).

3- Some exceptions happen where there are uniformly separated clauses (e.g. references or standards section where there is IEC then space then xxxxx then space then the standard label), acrobat plays smart and convert each part of the clause into a separate column making a table (which is smart btw and very useful in other applications but a big headache in ours). actual tables will be considered as a table of course and behave in the same way.

4- General observation1: the only rock solid way to check that a column is empty when controlling excel UI is to go the first row of that column and press ctrl+Down_arrow, if it goes to row 1048576 (bikini bottom xD) then it is empty. This is to be done before unmerge or it becomes completely irrelevant. Pro tip: it is the first column from the left with the default width.

4.1- Correction: the only rock solid way was not very solid after all, it fails to catch empty cells that are combined, we are looking for all signs of manipulation done by acrobat not text only. the only possible way is the visual way in the pro tip or just to play it safe and unmerge the whole damn sheet.

5- After unmerging the combined big cells, the problem raised in (3) shows clearly.

6- A uniform solution to this data preprocessing dilemma is as follows:

	a- extend the unmerge selection area to a very large column range (DD maybe? ;p ).
	b- in the column after the column in (6.a) end of range, we can make a string combine function to concatenate all the trailing cells content separated by a space.
	c- in UI, concatenate using TEXTJOIN() to enable specifying a delimiter.

7- General observation2: headers and footers are usually deleted by acrobat after conversion (this can be disrupted by stamps or signatures on the pdf). However, some brilliant designers tend to write headers under the headers or footers above the footers and those will stay. They will act like a normal clause or as mentioned in (3) so they have to be wranglered manually.

8- I will list below the step by step explanation of my way in converting a spec file to an excel sheet where each clause is separated in an individual cell (row):

"import_specs_xlsx" start
	a- Open the excel sheet that contains the converted pdf using code and copy the sheet called "Table 1" to start the manipulations. This should be done using a button with the below defined function:

'''
 Dim fd As Office.FileDialog

  Set fd = Application.FileDialog(msoFileDialogFilePicker)

  With fd
    .AllowMultiSelect = False
    .Title = "Please select the file."
    .Filters.Clear
    .Filters.Add "Excel 2003", "*.xls?"

    If .Show = True Then
      fileName = Dir(.SelectedItems(1))

    End If
  End With

  Application.ScreenUpdating = False
  Application.DisplayAlerts = False

Set closedBook = Workbooks.Open(fileName)
    closedBook.Sheets("Table 1").Copy Before:=ThisWorkbook.Sheets(1)
    closedBook.Close SaveChanges:=False
'''
"import_specs_xlsx" end

"table_processing" start

	b- Unmerge the full sheet "Table 1".
	c- Identify the last row index by searching for any value starting from bottom in column A. This variable is essential and should be global.
	d- TEXTJOIN(with space delimiter and ignore empty) The range A:DC sequentially for each row in the sheet starting from row 1 all the way to the index in (8.c) and save the result in column DC. This will produce the required single column data. (Range("DD"&i).Value = WorksheetFunction.TextJoin(" ", True, Range("A" & i & ":DC" & i))
	e- Copy the column DD as text to column A and preferably delete everything else.
	f- Apply Range(A1:A"index from (8.c)").TextToColumns(DataType:=xlDelimited, Other:=True, OtherChar:=vbLf) to separate each clause in a separate column.
	g- Below is the copy iteration pseudocode:

'''
i	2		
			
j	1	index	
copy	Aj	DDj	copy
	paste as transpose in a new sheet		
paste_transpose	Ai		
	i=	i+108	
'''

	h- type anything in the new sheet A1 cell to enable filtering and delete the "Table 1" sheet.
"table_processing" end

"formatting" start

	i- below code is for filtering and deleting all empty rows or have 0:

'''
Sub Sample()
    Dim lastRow As Long

    With Sheets("Sheet1")

        lastRow = .Range("A" & Rows.Count).End(xlUp).Row

        '~~> Remove any filters
        .AutoFilterMode = False

        '~~> Filter, offset(to exclude headers) and delete visible rows
        With .Range("A1:A" & lastRow)
            .AutoFilter Field:=1, Criteria1:="=", Operator:=xlAnd, Criteria2:="0"
            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End With

        '~~> Remove any filters
        .AutoFilterMode = False
    End With
End Sub
'''
	j- handle the required formatting. column width 100, wrap text, All boarders, add reply column with bold blue color with the first row label "alfanar reply"
	k- export the new sheet to a new workbook (move not copy).

"formatting" end


	l- think of UX to enable the user to select the required file and the code should export the converted and formatted file as a new book and add a column for the reply. write clear steps for users and include first time setup for VBA certificate and developer tab enabling (preferably add pictures).
	













