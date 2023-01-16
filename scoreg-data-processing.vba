
' HELPER FUNCTIONS

Function getRowCount(row As Long, col As Long) As Long
    getRowCount = Range(Cells(row, col).Address).End(xlDown).row
End Function

Function getColCount(row As Long, col As Long) As Long
    Dim col_count As Long
    
    col_count = Range(Cells(row, col).Address).End(xlToRight).Column
    If col_count = 16384 Then
        col_count = 1
    End If
    
    getColCount = col_count
End Function

Function getFilterColIndex(name As String, colCount As Long) As Long
    Dim i As Long, filter_col As Long
    
    filter_col = 0
    For i = 1 To colCount
        If InStr(1, name, Range(Cells(1, i).Address).Value, 1) > 0 Then
            filter_col = i
        End If
    Next

    If filter_col = 0 Then
        MsgBox ("ERROR: Die Spalte " & name & " wurde nicht gefunden. ")
    End If

    getFilterColIndex = filter_col
End Function

Function removeLineWithBlanks(filter_col As Long, row_count As Long)
    Dim i As Long
    
    For i = 1 To row_count
        If Range(Cells(i, filter_col).Address).Value = "" Then
            Range(Cells(i, filter_col).Address).Select
            Selection.EntireRow.Delete
        End If
    Next
End Function

'SUB remove duplicates general
Sub Remove_duplicates_general(Optional colHeader As String)
'
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    wb.Worksheets("Paste data here").Activate
    
    Dim row_count As Long
    Dim col_count As Long
    Dim filter_col  As Long
    Dim i As Long
    
    If Len(Trim(colHeader)) = 0 Then
        colHeader = InputBox("In welcher Spalte sollen duplikate entfernt werden")
    End If
    
    row_count = getRowCount(1, 1)
    col_count = getColCount(1, 1)
    
    filter_col = getFilterColIndex(colHeader, col_count)
    
    If filter_col = 0 Then
        GoTo Skip
    End If
    
    Range("$A$1:" & Cells(row_count, col_count).Address).RemoveDuplicates Columns:=Array(filter_col), _
        Header:=xlYes
Skip:
    wb.Worksheets("Control Panel").Activate
    
End Sub

' SUB remove email duplicates
Sub Remove_email_duplicates()
'
    Remove_duplicates_general ("E-Mail")
    
End Sub

' SUB remove scout-id duplicates
Sub Remove_scoutid_duplicates()
'
    Remove_duplicates_general ("Scout-ID")
    
End Sub

'SUB apply blacklist

Sub Apply_blacklist_to_data()
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
        
    wb.Worksheets("Blacklist").Activate
    
    ' Filter dims
    Dim colHeader As String
    Dim col_count As Long
    Dim filter_col As Long
    Dim row_count_filter As Long
    Dim i As Long
    Dim col_count_data As Long
    Dim row_count_data As Long
    
    colHeader = "Infomail"
    
    col_count = getColCount(1, 1)
    filter_col = getFilterColIndex(colHeader, col_count)
    
    If filter_col = 0 Then
        GoTo Skip
    End If
    
    row_count_filter = getRowCount(1, filter_col) - 1
    
    'Apply Filter
    wb.Worksheets("Paste data here").Activate
    
    Dim matchedRange As Range
    Dim firstAddress As String
    Dim data_col As Long
    
    For i = 1 To row_count_filter
        
        col_count_data = getColCount(1, 1)
        row_count_data = getRowCount(1, 1)
        
        data_col = getFilterColIndex("E-Mail", col_count_data)
        
        With Range(Cells(2, data_col).Address & ":" & Cells(row_count_data, data_col).Address)
            Set matchedRange = .Find(Worksheets("Blacklist").Range(Cells(i + 1, filter_col).Address).Value _
            , LookIn:=xlValues)
            If Not matchedRange Is Nothing Then
                firstAddress = matchedRange.Address
                Do
                    matchedRange.Select
                    matchedRange.Value = Replace(matchedRange.Value, matchedRange.Value, "")
                    Set matchedRange = .FindNext(matchedRange)
                Loop While Not matchedRange Is Nothing
            End If
        End With
    Next
    
    Call removeLineWithBlanks(data_col, row_count_data)
    
Skip:
    wb.Worksheets("Control Panel").Activate
End Sub

Sub Apply_whitelist_to_data()

    Dim wb As Workbook
    Set wb = ActiveWorkbook
        
    wb.Worksheets("Whitelist").Activate
    ' Filter dims
    Dim filterColHeader As String
    Dim dataColHeader As String
    
    Dim col_count As Long
    Dim filter_col As Long
    Dim row_count_filter As Long
    Dim row_count_data As Long
    Dim data_col As Long
    
    filterColHeader = "Infomail"
    dataColHeader = "E-Mail"
    
    col_count = getColCount(1, 1)
    filter_col = getFilterColIndex(filterColHeader, col_count)
    row_count_filter = getRowCount(1, filter_col)
    
    'Apply Filter
    wb.Worksheets("Paste data here").Activate
     
    col_count = getColCount(1, 1)
    row_count_data = getRowCount(1, 1)
    data_col = getFilterColIndex(dataColHeader, col_count)
    
    If data_col = 0 Then
        GoTo Skip
    End If
    
    Range(Cells(row_count_data + 1, 1).Address & ":" & _
        Cells(row_count_data - 1 + row_count_filter, col_count).Address).Value = "-"
        
    Worksheets("Whitelist").Range((Cells(2, filter_col).Address) & ":" & Cells(row_count_filter, filter_col).Address).Copy _
       Destination:=Range(Cells(row_count_data + 1, data_col).Address)
        
Skip:
    wb.Worksheets("Control Panel").Activate

End Sub


Sub Create_email_dist_list()

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    wb.Worksheets("Paste data here").Activate
    
    Dim row_count As Long
    Dim col_count As Long
    Dim mail_col As Long
    Dim dataColHeader As String
    
    dataColHeader = "E-Mail"
    
    col_count = getColCount(1, 1)
    row_count = getRowCount(1, 1)
    
    mail_col = getFilterColIndex(dataColHeader, col_count)
    If mail_col = 0 Then
        GoTo Skip
    End If
     
    Range((Cells(2, mail_col).Address) & ":" & Cells(row_count, mail_col).Address).Copy _
    Destination:=Worksheets("Output").Range("A1")
    
    wb.Worksheets("Output").Activate
    
    Call removeLineWithBlanks(1, col_count)
    
    'Format to blocks with certain length
    Dim block_length  As Long
    Dim num_block As Long
    Dim i As Long
    
    block_length = 249
    num_block = (row_count / block_length)

    For i = 1 To num_block
        Range((Cells((i * block_length + 1), 1).Address) & ":" & Cells((i + 1) * block_length, 1).Address).Cut _
    Destination:=Worksheets("Output").Range(Cells(1, i + 1).Address)
     
    Next
    Range("A1").Select
    
Skip:
    wb.Worksheets("Control Panel").Activate
    
End Sub

