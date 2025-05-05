Attribute VB_Name = "Module2"
Sub RowTransfer(sht As Worksheet, k As Integer, d() As Variant, _
    j As Integer, doc As Worksheet, z As Integer)

'determine if you are using data or distdata array
    If z = 0 Then
        z = j 'z stays 1 to distribute 1 dim. array:  distdata()
    End If

'Transfers data from input array to CSV
    'NOTE: hard-coded values will be the same for all
            'records (rows) in CSV file.
    sht.Cells(k, 1).Value = "I"      'Invoice
    sht.Cells(k, 2).Value = d(z, 1)  'trannum
    sht.Cells(k, 3).Value = d(z, 2)  'person
    sht.Cells(k, 5).Value = d(z, 3)  'date
    sht.Cells(k, 6).Value = _
        Format(doc.Cells(3, 2).Value, _
        "mm/yyyy")                   'postmonth
    sht.Cells(k, 7).Value = d(z, 4)  'ref
    sht.Cells(k, 8).Value = d(z, 9)  'notes
    sht.Cells(k, 9).Value = d(z, 5)  'property/table
    sht.Cells(k, 10).Value = d(z, 6) 'amount
    sht.Cells(k, 11).Value = d(z, 7) 'account
'    sht.Cells(k, 12).Value = "2110000" 'accrual
    sht.Cells(k, 15).Value = d(z, 8) 'description
    sht.Cells(k, 79).Value = "Standard Payable Display Type" 'displaytype
    sht.Cells(k, 80).Value = "Expense"  'expensetype
    'isconsolidated
'    If UCase(doc.Cells(1, 11).Value) = UCase("true") Then
'        sht.Cells(k, 118).Value = -1      'true
'    Else
'        sht.Cells(k, 118).Value = 0       'false
'    End If

End Sub
