Attribute VB_Name = "ReplaceChartsWithImages"
Sub chartToimage()
'https://stackoverflow.com/questions/31537325/convert-all-shape-to-image-in-ms-word-with-macro
'
' Welcome to the wonderful world of manipulating the very collection you are looping through. The moment you cut, you are effectively removing the shape from the collection, altering your loop.
'
' If you want to loop through shapes (or table rows or whatever) and delete something from that collection, simply go backwards:
'

Dim i As Integer, oShp As InlineShape

Application.ScreenUpdating = False

For i = ActiveDocument.InlineShapes.Count To 1 Step -1
    Set oShp = ActiveDocument.InlineShapes(i)
    oShp.Select
    'remove the border - can look messy
    oShp.Chart.ChartArea.Border.LineStyle = None
  
    'oShp.Chart.ChartArea.Width = 100
    Selection.Cut
    Selection.PasteSpecial Link:=False, DataType:=wdPasteEnhancedMetafile, _
        Placement:=wdInLine, DisplayAsIcon:=False
Next i

Application.ScreenUpdating = True

End Sub

