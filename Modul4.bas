Attribute VB_Name = "Modul4"


Option Explicit                                  'Always put this at the top of your Module!

Sub sendReport()


    Dim project As String
    Dim fname As String
    Dim templatePath As String

    project = "APT"                              ' hier uss generisch gändert werden

    fname = Application.ActivePresentation.Name

    Dim tempmail As MailItem
    Dim mail As MailItem
    Dim pdfPath As String                        'Pfad und Name von PDF wo agbelegt

    Dim pdfattach As Attachment
    Dim attach As Attachments
    Dim objOL As Outlook.Application
    Set objOL = New Outlook.Application


    'Aufruf der Methode savetoPDF
    pdfPath = savetoPDF()


    'Aufruf der Methode getProject
    project = getProject()




    'templatePath = "\\Mac\Home\Downloads\forPPTMacro\templates\" & project & ".oft" 'Application.ActivePresentation.path

    templatePath = Application.ActivePresentation.path & getProject() & ".oft"

    templatePath = Replace(templatePath, "ppt", "templates\")
    'Set tempmail = objOL.CreateItemFromTemplate("\\Mac\Home\Downloads\forPPTMacro\templates\APT.oft")
    Set tempmail = objOL.CreateItemFromTemplate(templatePath)

    'Neues Mail Objekt
    Set mail = objOL.CreateItem(olMailItem)

    'aktuelle zu reportendes Slide als Body erstellen

    Dim slide As slide
    'slide = ActivePresentation.Slides.FindBySlideID(2)

    Dim strMessage As String



    '''''''''''''''''''''''''Test

    mail.Display
    Dim wordDoc As Word.Document
    Set wordDoc = mail.GetInspector.WordEditor
    'wordDoc.Range.PasteAndFormat wdChartPicture


    ''''''''''''''''''''''''Test
    'ersetze KK xx mit aktueller Kalenderwoche und speichere in newSubject
    Dim newSubject As String
    Dim tmpString
    tmpString = Replace(tempmail.Subject, "KW xx", "KW XX")
    newSubject = Replace(tmpString, "KW XX", "KW " + getReportCW)

    mail.To = tempmail.To
    mail.BodyFormat = olFormatHTML

    'pdfattach.PathName = pdfPath
    mail.Attachments.Add (pdfPath)


    'mail.Attachments = attach



    'wordDoc.Content.InsertParagraphAfter
    ActivePresentation.Slides(getLastSlide).Shapes("Draft").TextFrame.TextRange.Copy
    wordDoc.Paragraphs(1).Range.PasteSpecial , , , , wdPasteRTF
    wordDoc.Content.InsertParagraphAfter


    'TODO: muss noch für mehrere angepasst werden
    'ActivePresentation.Slides(getLastSlide).Shapes("APT").TextFrame.TextRange.Copy

    'wordDoc.Paragraphs(1).Range.PasteSpecial , , , , wdPasteRTF



    'ActivePresentation.Slides().Shapes("ADP").TextFrame.TextRange.Copy
    'wordDoc.Paragraphs(1).Range.PasteSpecial , , , , wdPasteRTF


    mail.Subject = newSubject

    'Dim str As String
    'str = mail.Body
    'Debug.Print str
    'mail.Display

    'mail.To = tempmail.Sender
    'mail.CC = tempmail.CC
    'mail.Subject = tempmail.Subject
    'replyEmail.HTMLBody = replyEmail .HTMLBody & origEmail.Reply. HTMLBody
    'mail.Display

End Sub

Public Function getLastSlide() As Integer
    Dim cnt As Integer
    cnt = ActivePresentation.Slides.Count
    getLastSlide = cnt
End Function

Public Function getReportCW() As String
    Dim str As String
    getReportCW = Format(Date - 3, "ww", vbMonday, vbFirstFullWeek)
End Function

Public Function savetoPDF() As String
    Dim path As String
    Dim fname As String
    Dim cmplt As String
    Dim fso As New FileSystemObject
    path = Application.ActivePresentation.path & "\pdf\"
    If Not fso.FolderExists(path) Then

        ' doesn't exist, so create the folder
        fso.CreateFolder path
    End If

    fname = Application.ActivePresentation.Name
    cmplt = path & fname & ".pdf"
    cmplt = Replace(cmplt, ".pptm", "")
    'ActivePresentation.SaveAs path & fname & ".pdf", 32
    ActivePresentation.SaveAs cmplt, 32
    savetoPDF = cmplt
End Function

Public Function getProject() As String
    Dim fname As String
    Dim firstOc As Integer
    Dim secondOc As Integer


    Dim prj As String

    fname = Application.ActivePresentation.Name
    Debug.Print (fname)
    'prj = Replace(fname, "KW xx", "KW " + getReportCW)

    ' newSubject = Replace(tempmail.Subject, "KW xx", "KW " + getReportCW)

    firstOc = InStr(fname, "_")
    secondOc = InStr(firstOc + 1, fname, "_")
    prj = Mid(fname, firstOc + 1, secondOc - firstOc - 1)

    getProject = prj
End Function

