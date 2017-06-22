Attribute VB_Name = "ExtractFigureData"
Sub ExtractFigureData()
' https://msdn.microsoft.com/en-us/library/office/ff821389.aspx
' Know errors:
' Run-time error '-2147467259 (80004005)'
' select debug then continue.

'On Error GoTo -1

'Dim oChartStyle As Word.Style'
'
'If Not oDocStyles.styleExists("Table Header") Then
'   Set oChartStyle = pDoc.Styles.Add(Name:="Table Header", Type:=wdStyleTypeParagraph)
'   With oChartStyle
'      .BaseStyle = "Normal"
'      If .Font.Size <> 10 Then .Font.Size = 10
'      'If Not .Font.Bold Then .Font.Bold = True
'      If .Font.Scaling <> 100 Then .Font.Scaling = 100
'      If .Font.Spacing <> 0 Then .Font.Spacing = 0
'      'With .ParagraphFormat
'      '   .Space1
'      '   If Not .KeepWithNext Then .KeepWithNext = True
'      '   If .Alignment <> wdAlignParagraphCenter Then .Alignment = wdAlignParagraphCenter
'      'End With
'   End With
'End If

 Dim objShape As InlineShape
 'Dim myRange As Excel.Range
     
    ' Iterates each inline shape in the active document.
    ' If the inline shape contains a chart, then display the
    ' data associated with that chart and minimize the application
    ' used to display the data.
    For Each objShape In ActiveDocument.InlineShapes
        If objShape.HasChart Then
            ' Activate the topmost window of the application used to
            ' display the data for the chart.
            objShape.Select
            'Selection.MoveUp Unit:=wdLine, Count:=1
            'Selection.MoveStart Unit:=wdLine, Count:=-1
            'Selection.MoveEnd Unit:=wdLine, Count:=1
            Selection.MoveUp Unit:=wdParagraph, Count:=2
            Selection.MoveEnd Unit:=wdParagraph, Count:=1
            Selection.Copy
            Selection.HomeKey
            With Selection.Find
              .ClearFormatting
              .Replacement.ClearFormatting
              .Text = "^13^13"
              .Replacement.Text = ""
              .Forward = True
              .Wrap = wdFindStop
              .Format = False
              .MatchWildcards = True
              .Execute
              If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
              End If

            End With
            
            Selection.MoveEnd Unit:=wdCharacter, Count:=-1
            Selection.EndKey Unit:=wdLine
            Selection.MoveDown Unit:=wdParagraph, Count:=1
            'Selection.TypeParagraph
            Selection.TypeText Text:="FIGURE DATA - "
            Selection.Paste
            
            Selection.EndKey Unit:=wdLine
            'Selection.TypeParagraph

            'objShape.Chart.ChartData.Activate
            'ActiveSheet.UsedRange.Copy
            
            'objShape.Chart.ChartData.Workbook.ActiveSheet.ListObjects("Table1").Range.Copy
            objShape.Chart.ChartData.Activate
            ' usually table1
            If objShape.Chart.ChartData.Workbook.ActiveSheet.ListObjects.Count > 0 Then
            'Set myRange = objShape.Chart.ChartData.Workbook.ActiveSheet.ListObjects("Table1").Range
            'remove emptry columns and rows - cells can merge in Word if they are next to an
            'empty cell and you don't want that.
            '  For iCounter = myRange.Columns.Count To 1 Step -1
            '    If Excel.Application.CountA(Columns(iCounter).EntireColumn) = 0 Then
            '      Columns(iCounter).Delete
            '    End If
            '  Next iCounter
            '  For iCounter = myRange.Rows.Count To 1 Step -1
            '    If Excel.Application.CountA(Rows(iCounter).EntireColumn) = 0 Then
            '      Rows(iCounter).Delete
            '    End If
            '  Next iCounter
              objShape.Chart.ChartData.Workbook.ActiveSheet.ListObjects("Table1").Range.Copy
            Else
           ' ' but sometimes we just have to use the used range
           ' Set myRange = objShape.Chart.ChartData.Workbook.ActiveSheet.UsedRange
           ' 'remove emptry columns and rows - cells can merge in Word if they are next to an
           ' 'empty cell and you don't want that.
           '   For iCounter = myRange.Columns.Count To 1 Step -1
           '     If Excel.Application.CountA(Columns(iCounter).EntireColumn) = 0 Then
           '       Columns(iCounter).Delete
           '     End If
           '   Next iCounter
           '   For iCounter = myRange.Rows.Count To 1 Step -1
           '     If Excel.Application.CountA(Rows(iCounter).EntireColumn) = 0 Then
           '       Rows(iCounter).Delete
           '     End If
           '   Next iCounter
              objShape.Chart.ChartData.Workbook.ActiveSheet.UsedRange.Copy
            End If
           ' objShape.Chart.ChartData.Workbook.ActiveSheet.UsedRange.Copy
            objShape.Chart.ChartData.Workbook.Close
            'objShape.Chart.ChartData.Workbook.Application.Quit
            Selection.Paste
            Selection.Previous(Unit:=wdTable, Count:=1).Select
            If Selection.Tables(1).Uniform Then
              Selection.Tables(1).Rows(1).Select
              Selection.Style = ActiveDocument.Styles("Table Header")
              Selection.Style = ActiveDocument.Styles("Table Header")
              Selection.Tables(1).Select
              Selection.Tables(1).Columns(1).Select
              Selection.Style = ActiveDocument.Styles("Table Header")
              Selection.Style = ActiveDocument.Styles("Table Header")
            End If
        End If
    Next
End Sub



