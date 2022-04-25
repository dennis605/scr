Attribute VB_Name = "Modul2"
Option Explicit

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
