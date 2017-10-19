Attribute VB_Name = "ExternalExcelTables"
Sub ExernalExcelTables()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim ExcelWasNotRunning As Boolean
Dim WorkbookToWorkOn As String

Dim Starttime As Double
Dim SecondsElapsed As Double
Dim MinutesElapsed As String
Dim TblCounter As Integer

' Check the doc contains the Exhibit Title style.
On Error GoTo Style_Err_Handler
titleStyle = ActiveDocument.Styles("Exhibit Title")

' timer and counter for the complete message
Starttime = Timer
TblCounter = 0

' Stop Word updating
Application.ScreenUpdating = False

    'Always start at the top of the document
    Selection.HomeKey Unit:=wdStory

    'find an external include.  Expect the format:
    'External Excel Table: spreadsheetname {sheetname}
    With Selection.Find
        .ClearFormatting
        .Text = "External Excel Table:"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute
    End With
    
    Selection.HomeKey Unit:=wdStory
    
    Do While Selection.Find.Found = True
    
    Selection.Find.Execute
        'loop though all the found items
        If Selection.Find.Found Then
             Selection.Expand wdLine
             'eat the newline
             Selection.MoveEnd wdCharacter, -1
             sText = Application.Selection.Text
             sText = Replace(sText, "}", "")
             sText = Replace(sText, "External Excel Table: ", "")
             sText = Replace(sText, " {", ":")
             ' Split string into values
             ' sValues(0) = xls name
             ' sValues(1) = sheet name
             sValues = Split(sText, ":")


             'specify the workbook to work on
             WorkbookToWorkOn = ActiveDocument.Path & "\" & sValues(0)

             'If Excel is running, get a handle on it; otherwise start a new instance of Excel
             On Error Resume Next
             Set oXL = GetObject(, "Excel.Application")
             
             If Err Then
                ExcelWasNotRunning = True
                Set oXL = New Excel.Application
             End If

             On Error GoTo Err_Handler

             'If you want Excel to be visible, you could add the line: oXL.Visible = True here;
             'but your code will run faster if you don't make it visible
             oXL.Visible = False

             'Open the workbook only if it wasn't already open
             On Error Resume Next
             Set oWB = oXL.Workbooks(sValues(0))
             If Err Then
                 Set oWB = oXL.Workbooks.Open(FileName:=WorkbookToWorkOn)
             End If
             
             ' activate the specified sheet
             oXL.ActiveWorkbook.Sheets(sValues(1)).Activate
             oXL.FindFormat.Clear
             
             ' UsedRange and .Range("A1").CurrentRegion are inadequate
             ' select the entire used range of cells
             'Set oRng = oXL.ActiveWorkbook.Sheets(sValues(1)).Range("A1").CurrentRegion
             Set oRng = oXL.ActiveWorkbook.Sheets(sValues(1)).Range("A1")
             oRng.Copy
             Selection.PasteAndFormat (wdFormatPlainText)
             With Selection
               .Style = "Exhibit Title"
             End With
                         
             Set oRng = oXL.ActiveWorkbook.Sheets(sValues(1)).Range("A2")
             Set oRng = oRng.Resize(Cells.Find("*", , xlValues, , xlRows, xlPrevious).Row - 1, Cells.Find("*", , xlValues, , xlColumns, xlPrevious).Column)
             'oRng.Select
             oRng.Copy
             Selection.PasteAndFormat (wdFormatOriginalFormatting)
             With Selection
               .Style = "Normal"
             End With
             
        TblCounter = TblCounter + 1
        End If
    Loop
    
    'oXL.Visible = True
    oXL.ActiveWorkbook.Close
    Application.ScreenUpdating = True
   
    
    If ExcelWasNotRunning Then
       oWB.Close (False)
       oXL.Quit
    End If


    'Make sure you release object references.
    Set oRng = Nothing
    Set oSheet = Nothing
    Set oWB = Nothing
    Set oXL = Nothing
             
    SecondsElapsed = Round(Timer - Starttime, 2)
    MinutesElapsed = Format(SecondsElapsed / 86400, "hh:mm:ss")
    'MsgBox "Complete " & TblCounter & " tables inserted. " & SecondsElapsed & " seconds elapsed", vbOKOnly, "Message"
    MsgBox "Complete " & TblCounter & " tables inserted. " & MinutesElapsed & " minutes elapsed", vbOKOnly, "Message"
    Exit Sub

Style_Err_Handler:
   MsgBox "Please create the Exhibit Title style before running this macro"
   Exit Sub
   
Err_Handler:
   MsgBox WorkbookToWorkOn & " caused a problem. " & Err.Description, vbCritical, _
           "Error: " & Err.Number
   If ExcelWasNotRunning Then
       oXL.Quit
   End If
End Sub
