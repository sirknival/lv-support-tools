Sub AusbildungSammelrechnungen()
    Dim wdDoc As Document
    Dim wdMailMerge As MailMerge
    Dim i As Integer
    Dim pdfName, pdfPath, emailAddress, group, emailBody As String
    Dim saveAndMail As Boolean
    Dim olApp As Object
    Dim olMail As Object

     
    
    ' Aktuelles Dokument und MailMerge abrufen
    Set wdDoc = ActiveDocument
    Set wdMailMerge = wdDoc.MailMerge

    ' Serienbrief starten, falls nicht aktiv
    If wdMailMerge.State <> wdMainAndDataSource Then
        MsgBox "Bitte verbinden Sie das Dokument mit einer Datenquelle.", vbExclamation
        Exit Sub
    End If
    
    ' Verzeichnispfad zum Speichern der PDFs
    pdfPath = OrdnerAuswählen()
    
    If Dir(pdfPath, vbDirectory) = "" Then MkDir pdfPath
    
    saveAndMail = saveAndMailFun()
    If saveAndMail = True Then
        emailBodyText = MailTextFun()
    End If
    

    ' Outlook-Objekt erstellen
    Set olApp = CreateObject("Outlook.Application")

    ' Serienbrief starten
    wdMailMerge.Destination = wdSendToNewDocument
    wdMailMerge.Execute Pause:=False

    ' Zugriff auf Datenquelle sicherstellen
    Dim dataSource As MailMergeDataSource
    Set dataSource = wdDoc.MailMerge.dataSource
    
    ' Alle Feldnamen durchlaufen

    For i = 2 To dataSource.RecordCount
        
        dataSource.ActiveRecord = i ' Zum aktuellen Datensatz wechseln
        wdDoc.MailMerge.dataSource.FirstRecord = i
        wdDoc.MailMerge.dataSource.LastRecord = i
          
        ' E-Mail-Adresse aus der Datenquelle abrufen
        emailAddress = dataSource.DataFields(9).Value
        group = dataSource.DataFields(1).Value
        
        If group = "0" Then
            Exit For
        End If
        
        group = Replace(group, "/", "_")
        group = Replace(group, "&", "_")
        

        ' Seriendruck für den aktuellen Datensatz ausführen
        wdDoc.MailMerge.Destination = wdSendToNewDocument
        wdDoc.MailMerge.Execute Pause:=False

        ' PDF speichern
        Dim singleDoc As Document
        Set singleDoc = ActiveDocument
        pdfName = pdfPath & "\Sammelrechnung_Ausbildung_W" & group & ".pdf"
        singleDoc.ExportAsFixedFormat OutputFileName:=pdfName, ExportFormat:=wdExportFormatPDF
        
        If saveAndMail = True Then
        
            emailBody = "Hallo," & vbCrLf & vbCrLf & emailBodyText & vbCrLf & _
            "Beste Grüße & Gut Pfad" & vbCrLf & _
            "Flo" & vbCrLf & vbCrLf & _
            "Ausbildung Wiener Pfadfinder und Pfadfinderinnen" & vbCrLf & _
            "Hasnerstraße 41" & vbCrLf & _
            "1160 Wien" & vbCrLf & _
            "M: ausbildung@wpp.at" & vbCrLf & _
            "T: +43 1 495 23 15" & vbCrLf & _
            "W: wpp.at"
    
            
            ' E-Mail erstellen und senden
            Set olMail = olApp.CreateItem(0)
            With olMail
                .To = emailAddress
                .Subject = "Ausbildung Sammelrechnung W" & group
                .Body = emailBody
                .Attachments.Add pdfName
                .Send
            End With

        End If
        

        ' Dokument schließen
        singleDoc.Close False
    Next i

    ' Aufräumen
    MsgBox "Ausführung beendet", vbInformation
    Set wdDoc = Nothing
    Set wdMailMerge = Nothing
    Set olApp = Nothing
End Sub


Function OrdnerAuswählen() As String
    Dim objShell As Object
    Dim objFolder As Object
    Dim ordnerPfad As String
    
    ' Shell-Objekt erstellen
    Set objShell = CreateObject("Shell.Application")
    
    ' Dialog zum Ordnerauswählen anzeigen
    Set objFolder = objShell.BrowseForFolder(0, "Bitte wählen Sie einen Ordner aus:", 1)
    
    ' Wenn der Benutzer einen Ordner ausgewählt hat
    If Not objFolder Is Nothing Then
        ordnerPfad = objFolder.Self.Path
    Else
        ordnerPfad = ""
    End If
    
    ' Rückgabewert setzen
    OrdnerAuswählen = ordnerPfad
End Function

Function saveAndMailFun() As Boolean
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim chosenOption As Boolean
    Msg = "Sollen die Dokumente auch per Mail versendet werden?"    ' Define message.
    Style = vbYesNo Or vbInformation Or vbDefaultButton2 Or vbApplicationModal    ' Define buttons.
    Title = "Optionsausahl"    ' Define title.
    Help = "DEMO.HLP"    ' Define Help file.
    Ctxt = 1000    ' Define topic context.
            ' Display message.
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    If Response = vbYes Then    ' User chose Yes.
        chosenOption = True    ' Perform some action.
    Else    ' User chose No.
        chosenOption = False    ' Perform some action.
    End If
    
    saveAndMailFun = chosenOption
End Function


Function MailTextFun() As String
    Dim Message, Title, Default, mailBody
    Message = "Mailtext eingeben. Begrüßung & Signatur wird automtisch hinzugefügt"    ' Set prompt.
    Title = "Mailbody"    ' Set title.
    Default = "1"    ' Set default.
    ' Display message, title, and default value.
    mailBody = InputBox(Message, Title, Default)
    
    MailTextFun = mailBody

End Function
