Sub Dev()

Dim wb As Workbook
Set wb = ActiveWorkbook

Dim wsOverview As Worksheet
Set wsOverview = wb.Sheets("Gesamtliste")

'sets up the framework for using Word
Dim wordApp As Object
Dim wordDoc As Object

Dim name, surname, awarding_date, awarding_number, award_code, award_text As String
Dim strFileDA, strFileEZ, strFileL, strExport, done As String

Dim n, j As Integer

Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False

strFileDA = ThisWorkbook.Path & "\Template_Dankabzeichen_2023.docx"
strFileEZ = ThisWorkbook.Path & "\Template_Ehrenzeichen_2023.docx"
strFileL = ThisWorkbook.Path & "\Template_Lilien_2023.docx"
                
'now we begin the loop for the mailing sheet that is being used

n = wsOverview.Range("A:A").Find(what:="*", searchdirection:=xlPrevious).Row

With CreateObject("Scripting.FileSystemObject")
    If Not .FolderExists(ThisWorkbook.Path & "\Export") Then .CreateFolder ThisWorkbook.Path & "\Export"
End With

For j = 2 To n
        
        'collects the  strings needed for the document, skips already done awards
        done = wsOverview.Range("J" & j).Value
        name = wsOverview.Range("C" & j).Value
        
        If done <> "Nein" Or name = "" Then
            GoTo NextIteration
        End If
        
        awarding_date = wsOverview.Range("A" & j).Value
        award_code = wsOverview.Range("D" & j).Value
        award_text = wsOverview.Range("E" & j).Value
        awarding_number = wsOverview.Range("F" & j).Value
        
        'generate String for path later
        
        surname = Right(name, Len(name) - (InStrRev(name, " ")))
        strExport = ThisWorkbook.Path & "\Export\" & surname & "_" & award_code & "_" & Year(CStr(awarding_date)) & ".pdf"
        
        'opens the word doc that has the template  for sending out
        If award_code = "DA" Then
            Set wordDoc = wordApp.Documents.Open(strFileDA)
        ElseIf InStr(1, award_code, "EZ", 1) > 0 Then
            Set wordDoc = wordApp.Documents.Open(strFileEZ)
        Else
            Set wordDoc = wordApp.Documents.Open(strFileL)
        End If
        
        'fills in the word doc with the missing fields
        With wordDoc.Content.Find
            .Execute FindText:="<<name>>", ReplaceWith:=name, Replace:=wdReplaceAll
            .Execute FindText:="<<type>>", ReplaceWith:=award_text, Replace:=wdReplaceAll
            .Execute FindText:="<<number>>", ReplaceWith:=awarding_number, Replace:=wdReplaceAll
            .Execute FindText:="<<date>>", ReplaceWith:=Format(CStr(awarding_date), "dd. MMMM yyyy"), Replace:=wdReplaceAll
        End With

        ' this section saves the word doc in the folder as a pdf
        wordDoc.ExportAsFixedFormat strExport, wdExportFormatPDF, False

    'need to close word now that it has been opened before the next loop
    wordDoc.Close (wdDoNotSaveChanges)

NextIteration:
Next
Set wordDoc = Nothing
Set wordApp = Nothing
End Sub

