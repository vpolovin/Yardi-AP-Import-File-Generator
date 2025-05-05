Attribute VB_Name = "Module1"
Sub Button1_Click()
'Turn off screen updating
Application.ScreenUpdating = False

'Define column variables and tables as array
Dim r As Integer, c As Integer  'for numbering of rows and columns
Dim i As Integer, j As Integer  'for looping
Dim k As Integer, m As Integer  'for looping
Dim z As Integer 'for looping (reusable)
Dim sum As Double, diff As Double, length As Integer
Dim export As Workbook 'for the new csv file
Dim doc As Worksheet, tbl As Worksheet, sht As Worksheet

'Identify workbook and create/set array to copy data
Set doc = ActiveSheet
Set tbl = ThisWorkbook.Worksheets("Tables")
c = WorksheetFunction.CountA(Range("5:5")) 'count existing columns
r = WorksheetFunction.CountA(Range("A6:A300")) 'count existing columns

Dim data() As Variant 'create array named 'data'
ReDim data(r, c) 'set dimensions
For i = 1 To c
    For j = 1 To r
        data(j, i) = doc.Cells(j + 5, i).Value 'row 6 is first in doc
    Next j
Next i

'Create new .csv and copy data over
Set export = Workbooks.Add 'create/name workbook
Set sht = export.Sheets(1) 'create/name worksheet
    'sht.Name = "Test CSV"

'Save As
'dt = Format(Now, "yymmddhhnn") & "_Exported File"
'dt = ThisWorkbook.Path & "\" & Format(Now, "yymmddhhnn") & "_Exported File"
dt = doc.Name

fname = Application.GetSaveAsFilename( _
    InitialFileName:=dt, _
    filefilter:="Comma delimited file (*.csv),*.csv")
    If fname = False Then
        export.Close SaveChanges:=False
        End
    End If

'Dim transferred As Boolean
'transferred = False
'j - array rows
k = 1 'csv rows
'i - columns
't - table rows
'z - array columns

Dim distdata() As Variant  'create distribution array
ReDim distdata(1, c) 'set as 1 dimensional

'Transfer one entered row (j) at a time
For j = 1 To r

  'If zero, skip entry
  If 0 = data(j, 6) Then
    'Do nothing. This goes to 'Next j'
  
  'if ztable, then distribute
  ElseIf 0 < InStr(UCase(doc.Cells(j + 5, 5).Value), UCase("ztable")) Then
    
    'Find relevant ztable row
    t = 1
    While tbl.Cells(t, 1).Value <> doc.Cells(j + 5, 5).Value
        t = t + 1
    Wend
    m = t 'assign table start row
    
    'add data row to distdata array
    For z = 1 To c
        distdata(1, z) = data(j, z)
    Next z
    
    'loop through each distributed amount
    While tbl.Cells(t, 1).Value = data(j, 5)
        'get dist. amount
        distdata(1, 6) = Round(data(j, 6) * tbl.Cells(t, 3).Value, 2)
        'get property code
        distdata(1, 5) = tbl.Cells(t, 2).Value
        'RowTransfer(csv,csv rcount, array, doc rcount, doc, distribute?)
        Call RowTransfer(sht, k, distdata, j, doc, 1)
        t = t + 1 'increment table row
        k = k + 1 'increment csv row
    Wend
    
    'Correct any rounding errors
    'Sum the distributed amount
    length = t - m ' (t - m) is the table's start & end rows
    sum = 0
    For z = 1 To length
        'sum uses Banker's rounding
        sum = Round(sum + sht.Cells(k - 1 - length + z, 10), 2)
    Next z
    diff = Round(data(j, 6) - sum, 2) 'find excess
    'add excess to last distributed amount
    sht.Cells(k - 1, 10).Value = sht.Cells(k - 1, 10).Value + diff
    
    
    
'    Verify math
'    MsgBox "sum + diff: " & sum + diff & vbNewLine & _
'        "entered: " & data(j, 6) & vbNewLine & _
'        "diff: " & diff
        
        
        
  Else
    'Transfer row normally
    Call RowTransfer(sht, k, data, j, doc, 0)
    k = k + 1  'increment row on .csv file
  End If
Next j


sht.SaveAs Filename:=fname, FileFormat:=6 'save as .csv
export.Close SaveChanges:=False 'close .csv

'Turn on screen updating
Application.ScreenUpdating = True


End Sub
