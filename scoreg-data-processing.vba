Sub Remove_email_duplicates()
'
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    wb.Worksheets("Paste data here").Activate
    
    Dim row_count, col_count, i, mail_col  As Integer
    
    row_count = Range("A1").End(xlDown).Row
    col_count = Range("A1").End(xlToRight).Column
    
    If col_count = 16384 Then
        col_count = 1
    End If
    
    For i = 1 To col_count
        If InStr(1, "E-Mail", Range(Cells(1, i).Address).Value, 1) > 0 Then
            mail_col = i
        End If
    Next
    
     Range("$A$1:" & Cells(row_count, col_count).Address).RemoveDuplicates Columns:=Array(mail_col), _
        Header:=xlYes
        wb.Worksheets("Control Panel").Activate
    
End Sub

Sub Remove_duplicates_general()
'
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    wb.Worksheets("Paste data here").Activate
    
    Dim row_count, col_count, i, filter_col  As Integer
    Dim col_header As String
    
    mail_col = 0
    col_header = InputBox("In welcher Spalte sollen duplikate entfernt werden")
    
    row_count = Range("A1").End(xlDown).Row
    col_count = Range("A1").End(xlToRight).Column
    
    If col_count = 16384 Then
        col_count = 1
    End If
    
    For i = 1 To col_count
        If InStr(1, col_header, Range(Cells(1, i).Address).Value, 1) > 0 Then
            filter_col = i
        End If
    Next
    
    If filter_col = 0 Then
        MsgBox ("ERROR: Die Spalte " & col_header & " wurde nicht gefunden. ")
        GoTo Skip
    End If
    
     Range("$A$1:" & Cells(row_count, col_count).Address).RemoveDuplicates Columns:=Array(filter_col), _
        Header:=xlYes
Skip:
wb.Worksheets("Control Panel").Activate
    
End Sub


Sub Apply_blacklist_to_data()

Dim wb As Workbook
Set wb = ActiveWorkbook
    
wb.Worksheets("Blacklist").Activate
' Filter dims
Dim colHeader As String
Dim row_count_filter, mail_col, col_count, i, n, filter_col, numEntries, col_count_data, row_count_data   As Long

col_count = Range("A1").End(xlToRight).Column
col_header = "Infomail"

If col_count = 16384 Then
    col_count = 1
End If

For i = 1 To col_count
    If InStr(1, col_header, Range(Cells(1, i).Address).Value, 1) > 0 Then
        filter_col = i
    End If
Next
row_count_filter = Range(Cells(1, filter_col).Address).End(xlDown).Row - 1


'Apply Filter
wb.Worksheets("Paste data here").Activate

Dim c As Range
Dim firstAddress As String

For n = 1 To row_count_filter
    
    col_count_data = Range("A1").End(xlToRight).Column
    row_count_data = Range("A1").End(xlDown).Row
    
    For i = 1 To col_count_data
        If InStr(1, "E-Mail", Range(Cells(1, i).Address).Value, 1) > 0 Then
            mail_col = i
        End If
    Next

    With Range(Cells(2, mail_col).Address & ":" & Cells(row_count_data, mail_col).Address)
        Set c = .Find(Worksheets("Blacklist").Range(Cells(n + 1, filter_col).Address).Value _
        , LookIn:=xlValues)
        If Not c Is Nothing Then
            firstAddress = c.Address
            Do
                c.Select
                c.Value = Replace(c.Value, c.Value, "")
                'Selection.EntireRow.Delete
                Set c = .FindNext(c)
            Loop While Not c Is Nothing
        End If
    End With
Next

For i = 1 To row_count_data
    If Range(Cells(i, mail_col).Address).Value = "" Then
        Range(Cells(i, mail_col).Address).Select
        Selection.EntireRow.Delete
    End If
Next

wb.Worksheets("Control Panel").Activate

End Sub

Sub Apply_whitelist_to_data()

Dim wb As Workbook
Set wb = ActiveWorkbook
    
wb.Worksheets("Whitelist").Activate
' Filter dims
Dim colHeader As String
Dim row_count_filter, col_count, i, filter_col, row_count_data, mail_col  As Long

col_count = Range("A1").End(xlToRight).Column
col_header = "Infomail"

If col_count = 16384 Then
    col_count = 1
End If

For i = 1 To col_count
    If InStr(1, col_header, Range(Cells(1, i).Address).Value, 1) > 0 Then
        filter_col = i
    End If
Next
row_count_filter = Range(Cells(1, filter_col).Address).End(xlDown).Row


'Apply Filter
 
    col_count = Worksheets("Paste data here").Range("A1").End(xlToRight).Column
    row_count_data = Worksheets("Paste data here").Range("A1").End(xlDown).Row
    
    
    For i = 1 To col_count
        If InStr(1, "E-Mail", Worksheets("Paste data here").Range(Cells(1, i).Address).Value, 1) > 0 Then
            mail_col = i
        End If
    Next

Worksheets("Paste data here").Range(Cells(row_count_data + 1, 1).Address & ":" & _
    Cells(row_count_data - 1 + row_count_filter, col_count).Address).Value = "-"
Range((Cells(2, 1).Address) & ":" & Cells(row_count_filter, 1).Address).Copy _
   Destination:=Worksheets("Paste data here").Range(Cells(row_count_data + 1, mail_col).Address)
    
wb.Worksheets("Control Panel").Activate

End Sub


Sub Create_email_dist_list()

Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    wb.Worksheets("Paste data here").Activate
    
    Dim row_count, col_count, mail_col, i As Integer
    
    col_count = Range("A1").End(xlToRight).Column
    row_count = Range("A1").End(xlDown).Row
    
    For i = 1 To col_count
        If InStr(1, "E-Mail", Range(Cells(1, i).Address).Value, 1) > 0 Then
            mail_col = i
        End If
    Next
     
    Range((Cells(2, mail_col).Address) & ":" & Cells(row_count, mail_col).Address).Copy _
    Destination:=Worksheets("Output").Range("A1")
    
    wb.Worksheets("Output").Activate
    For i = 1 To row_count
        If Range("A" & i).Value = "" Then
            Range("A" & i).Select
            Selection.EntireRow.Delete
        End If
    Next
    
    Dim block_length  As Integer
    Dim num_block As Long
    
    block_length = 249
    num_block = (row_count / block_length)

    For i = 1 To num_block
        Range((Cells((i * block_length + 1), 1).Address) & ":" & Cells((i + 1) * block_length, 1).Address).Cut _
    Destination:=Worksheets("Output").Range(Cells(1, i + 1).Address)
     
    Next
    Range("A1").Select
    
    wb.Worksheets("Control Panel").Activate
    
End Sub

