Attribute VB_Name = "Modul3"


Option Explicit 'Always put this at the top of your Module!
Sub sendReport()



'Datum für KW
Dim aktuelleWoche As String

Dim tempmail As MailItem
Dim mail As MailItem

Dim objOL As Outlook.Application
Set objOL = New Outlook.Application

'''''''''''''''''''get acive shape''''''''''




''''''''''''''''''''''''''''''''''''''''


'Template Mail Laden
Set tempmail = objOL.CreateItemFromTemplate("\\Mac\Home\Downloads\forPPTMacro\templates\APT.oft")

'Neues Mail Objekt
Set mail = objOL.CreateItem(olMailItem)

'aktuelle zu reportendes Slide als Body erstellen

Dim slide As slide
'slide = ActivePresentation.Slides.FindBySlideID(2)

Dim strMessage As String
ActivePresentation.Slides(8).Shapes(4).TextFrame.TextRange.Copy
'ermitteln der aktuellen Kalendarwoche
aktuelleWoche = Format(Date - 3, "ww", vbMonday, vbFirstFullWeek)


'''''''''''''''''''''''''Test

mail.Display
Dim wordDoc As Word.Document
Set wordDoc = mail.GetInspector.WordEditor
'wordDoc.Range.PasteAndFormat wdChartPicture


''''''''''''''''''''''''Test
'ersetze KK xx mit aktueller Kalenderwoche und speichere in newSubject
Dim newSubject As String
newSubject = Replace(tempmail.Subject, "KW xx", "KW " + aktuelleWoche)
mail.To = tempmail.To
mail.BodyFormat = olFormatHTML
'mail.Body = "efkwmfkmfklf"
Dim r1 As Range
'r1 = ActiveDocument.Range
'ActivePresentation.Slides(8).Shapes(4).TextFrame.TextRange.Copy
'wordDoc.Range.PasteAndFormat wdChartPicture

'mail.Body = "egregerhe"

wordDoc.Content.InsertParagraphAfter
wordDoc.Content.InsertParagraphAfter
wordDoc.Content.InsertParagraphBefore

ActivePresentation.Slides(8).Shapes(5).TextFrame.TextRange.Copy
wordDoc.Paragraphs(2).Range.PasteAndFormat wdChartPicture

'wordDoc.Content.InsertParagraphAfter

ActivePresentation.Slides(8).Shapes(6).TextFrame.TextRange.Copy

    'wordDoc.Paragraphs(2).Range.PasteSpecial , , , , wdPasteBitmap
wordDoc.Paragraphs(1).Range.PasteAndFormat wdFormatPlainText
    
    
ActivePresentation.Slides(8).Shapes(4).TextFrame.TextRange.Copy
wordDoc.Paragraphs(3).Range.PasteAndFormat wdChartPicture

    

'wordDoc.Range.Paragraphs.Alignment = wdAlignParagraphJustifyHi
mail.Subject = newSubject
wordDoc.Content.InsertParagraphAfter
Dim str As String
str = mail.Body
Debug.Print str
'mail.Body = "egregerhe"
mail.Display

'mail.To = tempmail.Sender
'mail.CC = tempmail.CC
'mail.Subject = tempmail.Subject
'replyEmail.HTMLBody = replyEmail .HTMLBody & origEmail.Reply. HTMLBody
'mail.Display
End Sub















Sub DetermineActiveShape()
'PURPOSE: Determine the currently selected shape in PowerPoint
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim ActiveShape As Shape
Dim shp As Shape

'Determine Which Shape is Active
  If ActiveWindow.Selection.Type = ppSelectionShapes Then
    'Loop in case multiples shapes selected
       For Each shp In ActiveWindow.Selection.ShapeRange
         'ActiveShape is first shape selected
            Set ActiveShape = shp
            Exit For
       Next shp
  Else
    MsgBox "There is no shape currently selected!", vbExclamation, "No Shape Found"
  End If

'Do Something with the ActiveShape
  ActiveShape.TextFrame2.TextRange.Text = "Hello!"

End Sub

