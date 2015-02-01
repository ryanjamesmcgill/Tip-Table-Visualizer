Dim objSession As Object 'Used by OO4O to Set up a session for use on an Oracle DB
Dim objDatabase As Object 'Used by OO4O to create a connection to the Oracle DB for the Session

Sub Main()
Dim oraDynaset_tool As Object   'correcsponds to query 1 in SQL tab
Dim SQL_tool As String  ' will hold the SQL query
Dim oraDynaset_status As Object 'correcsponds to query 2 in SQL tab
Dim SQL_status As String  ' will hold the SQL query
Dim oraDynaset_comment As Object    'correcsponds to query 3 in SQL tab
Dim SQL_comment As String  ' will hold the SQL query
Dim oraDynaset_code As Object   'correcsponds to query 4 in SQL tab
Dim SQL_code As String  ' will hold the SQL query

Dim tool_list As String 'hold list of tools to go in pivot section of sql queries

dummy = updateStatus("Starting Macro")

PreApplicationSetups
ConnectToDatabase
dummy = formatUserEntry(ActiveWorkbook.Sheets("TipTableVisualizer").Range("C4"))
updateUserEntry

'clear spread sheet
dummy = clearFormatting(8) 'pass the row to start clearing
dummy = clearComments(ActiveWorkbook.Sheets("TipTableVisualizer"))

'generate SQL strings
SQL_tool = generateSQL(ActiveWorkbook.Sheets("SQL").Range("D4"), 15)
SQL_status = generateSQL(ActiveWorkbook.Sheets("SQL").Range("D22"), 21)
SQL_comment = generateSQL(ActiveWorkbook.Sheets("SQL").Range("D46"), 21)
SQL_code = generateSQL(ActiveWorkbook.Sheets("SQL").Range("D70"), 21)
'create unique tool list for pivot columns
dummy = updateStatus("Loading SQL 1 of 4 (this one takes the longest)")
Set oraDynaset_tool = objDatabase.DBCreateDynaset(SQL_tool, 0) 'Run the SQL Statement
tool_list = populateTool_list(oraDynaset_tool)

dummy = updateStatus("Loading SQL 2 of 4")

'insert tool list into SQL strings
SQL_status = Replace(SQL_status, "tool_list", tool_list)
SQL_comment = Replace(SQL_comment, "tool_list", tool_list)
SQL_code = Replace(SQL_code, "tool_list", tool_list)

'Query SQL
Set oraDynaset_status = objDatabase.DBCreateDynaset(SQL_status, 0) 'Run the SQL Statement
dummy = updateStatus("Loading SQL 3 of 4")
Set oraDynaset_comment = objDatabase.DBCreateDynaset(SQL_comment, 0) 'Run the SQL Statement
dummy = updateStatus("Loading SQL 4 of 4")
Set oraDynaset_code = objDatabase.DBCreateDynaset(SQL_code, 0) 'Run the SQL Statement
dummy = updateStatus("Almost done...")

'Print data to worksheet
dummy = print_xaxis(oraDynaset_tool, ActiveWorkbook.Sheets("TipTableVisualizer").Range("E8"))
dummy = print_yaxis(oraDynaset_status, ActiveWorkbook.Sheets("TipTableVisualizer").Range("A10"))
dummy = print_status(oraDynaset_status, ActiveWorkbook.Sheets("TipTableVisualizer").Range("E10"), oraDynaset_tool.RecordCount)
dummy = print_comment(oraDynaset_comment, ActiveWorkbook.Sheets("TipTableVisualizer").Range("E10"), oraDynaset_tool.RecordCount) 'must be called before print_code() or it will ignore Q-Tipped lines
dummy = print_code(oraDynaset_code, ActiveWorkbook.Sheets("TipTableVisualizer").Range("E10"), oraDynaset_tool.RecordCount)

'formatting
dummy = mergeEqpIDs(ActiveWorkbook.Sheets("TipTableVisualizer").Range("E9"))  'provide range where CHAMBER IDs Start
dummy = drawBottomLine(ActiveWorkbook.Sheets("TipTableVisualizer").rows(9))
dummy = formatColumns(ActiveWorkbook.Sheets("TipTableVisualizer").Range("A8"), oraDynaset_status.RecordCount)
dummy = drawStepTypeDivders(ActiveWorkbook.Sheets("TipTableVisualizer").Range("B10"), oraDynaset_status.RecordCount)

Application.StatusBar = "DONE"

dummy = checkForEmptyQuery(oraDynaset_tool)
PostApplicationSetups
End Sub


Sub ConnectToDatabase()
    Set objSession = CreateObject("OracleInProcServer.XOraSession") 'Set up the Oracle DB Session
    Set objDatabase = objSession.OpenDatabase("MFGINFO.World", "u_msas2/sa1sfby", 0) 'Assign the Oracle DB for the Session
End Sub

Sub PreApplicationSetups()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
End Sub

Sub PostApplicationSetups()
    ActiveWorkbook.Sheets("TipTableVisualizer").Range("A9").Select 'just to clear any current selections
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    ActiveWindow.ScrollRow = 1
End Sub

Sub updateUserEntry()
    ActiveWorkbook.Sheets("SQL").Activate
    ActiveWorkbook.Sheets("SQL").Range(Cells(1, 1), Cells(70, 4)).Calculate
    ActiveWorkbook.Sheets("TipTableVisualizer").Activate
End Sub

Function populateTool_list(oraDynaset As Object) As String
    oraDynaset.movefirst
    Dim tool_list As String
    For i = 1 To oraDynaset.RecordCount
        If (i = oraDynaset.RecordCount) Then
            tool_list = tool_list & "'" & oraDynaset.Fields("TOOL").Value & "'"
        Else
            tool_list = tool_list & "'" & oraDynaset.Fields("TOOL").Value & "',"
        End If
    oraDynaset.movenext
    Next i
populateTool_list = tool_list
End Function

Function generateSQL(rng As Range, rows As Integer) As String
    Dim SQL As String
    For i = 0 To rows
        SQL = SQL & " " & rng.Offset(i, 0).Value
    Next i
generateSQL = SQL
End Function

Function print_xaxis(oraDynaset As Object, rng As Range)
    oraDynaset.movefirst
    For i = 0 To (oraDynaset.RecordCount - 1) 'because array starts at 0
        rng.Offset(0, i).Value = oraDynaset.Fields("EQP_ID").Value
        rng.Offset(1, i).Value = oraDynaset.Fields("TKIN_PREVENT_CHAMBER_ID").Value
            If (oraDynaset.Fields("EQP_TRANSN_COMMENT").Value <> "") Then
                rng.Offset(0, i).AddComment (oraDynaset.Fields("EQP_TRANSN_COMMENT").Value)
                rng.Offset(1, i).AddComment (oraDynaset.Fields("EQP_TRANSN_COMMENT").Value)
            End If
        Select Case oraDynaset.Fields("EQP_STATUS").Value
        Case "LOCAL"
            rng.Offset(0, i).Interior.Color = 65535
            rng.Offset(1, i).Interior.Color = 65535
        Case "DOWN"
            rng.Offset(0, i).Interior.ThemeColor = xlThemeColorAccent2
            rng.Offset(0, i).Interior.TintAndShade = 0.399975585192419
            rng.Offset(1, i).Interior.ThemeColor = xlThemeColorAccent2
            rng.Offset(1, i).Interior.TintAndShade = 0.399975585192419
        Case Else
            rng.Offset(0, i).Interior.Color = 5287936
            rng.Offset(1, i).Interior.Color = 5287936
        End Select
        oraDynaset.movenext
    Next i
    print_xaxis = "done"
End Function

Function print_yaxis(oraDynaset As Object, rng As Range)
    oraDynaset.movefirst
        rng.Offset(-1, 0).Value = "Process ID"
        rng.Offset(-1, 1).Value = "Step Seq"
        rng.Offset(-1, 2).Value = "PRC Group"
        rng.Offset(-1, 3).Value = "PPID"
        rng.Offset(-2, 0).Value = "~"
        rng.Offset(-2, 1).Value = "~"
        rng.Offset(-2, 2).Value = "~"
        rng.Offset(-2, 3).Value = "~"
    For i = 0 To (oraDynaset.RecordCount - 1) 'because array starts at 0
        rng.Offset(i, 0).Value = oraDynaset.Fields("PROCESS_ID").Value
        rng.Offset(i, 1).Value = oraDynaset.Fields("STEP_SEQ").Value
        rng.Offset(i, 2).Value = oraDynaset.Fields("PROCESS_GROUP_NAME").Value
        rng.Offset(i, 3).Value = oraDynaset.Fields("PPID").Value
        oraDynaset.movenext
    Next i
    print_yaxis = "done"
End Function





Function print_status(oraDynaset As Object, rng As Range, numberOfTools As Integer)
    ReDim arr(oraDynaset.RecordCount, numberOfTools) As String   'since array starts at 0 index, we subtract 1
    oraDynaset.movefirst
    For x = 0 To (oraDynaset.RecordCount - 1)
        For y = 0 To (numberOfTools - 1)
            Status = oraDynaset.Fields(4 + y).Value 'Fields start at 0, and I'm using a 4 column offset to start at the tool columns in pivot table.
            If (IsNull(Status)) Then
                arr(x, y) = "-"
            Else
                arr(x, y) = Mid(Status, 1, 1)
            End If
        Next y
    oraDynaset.movenext
    Next x
Range(rng, rng.Offset(oraDynaset.RecordCount - 1, numberOfTools - 1)) = arr
End Function

Function print_comment(oraDynaset As Object, rng As Range, numberOfTools As Integer)
    oraDynaset.movefirst
    For x = 0 To (oraDynaset.RecordCount - 1)
        For y = 0 To (numberOfTools - 1)
            If (rng.Offset(x, y).Value = "P") Then 'since called before Q-tip printer, we can just look for prevents
            Comment = oraDynaset.Fields(4 + y).Value 'Fields start at 0, and I'm using a 4 column offset to start at the tool columns in pivot table.
            rng.Offset(x, y).AddComment (Comment)
            End If
        Next y
    oraDynaset.movenext
    Next x
End Function


Function print_code(oraDynaset As Object, rng As Range, numberOfTools As Integer)
    oraDynaset.movefirst
    For x = 0 To (oraDynaset.RecordCount - 1)
        For y = 0 To (numberOfTools - 1)
            Code = oraDynaset.Fields(4 + y).Value 'Fields start at 0, and I'm using a 4 column offset to start at the tool columns in pivot table.
            If (Code = "Q") Then
                rng.Offset(x, y).Value = "Q"
            End If
        Next y
    oraDynaset.movenext
    Next x
End Function

Function mergeEqpIDs(rng As Range)
Dim rngMerge As Range
Dim rngBox As Range
Dim i As Integer
Dim EqpID As String
Set rngMerge = rng.Offset(-1, 0)
EqpID = rng.Offset(-1, 0).Value 'initialize to first tool
    Do While (rng.Offset(-1, i) <> "")
        Set rngMerge = Union(rngMerge, rng.Offset(-1, i))
        rngMerge.Select
        If (EqpID <> rng.Offset(-1, i + 1)) Then
            Set rngBox = Range(rngMerge, rngMerge.End(xlDown))
            dummy = drawOutline(rngBox)
            rngMerge.Merge
            Set rngMerge = rng.Offset(-1, i + 1)
            EqpID = rngMerge.Value
        End If
    i = i + 1
    Loop
End Function

Function drawOutline(rng As Range)
    rng.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
End Function

Function drawBottomLine(rng As Range)
    rng.Borders(xlEdgeBottom).LineStyle = xlContinuous
    rng.Borders(xlEdgeBottom).Weight = xlMedium
End Function

Function clearFormatting(row As Integer)
    Range(rows(row), rows(row).End(xlDown)).Select
    Selection.UnMerge
    Selection.ClearContents
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Function

Function clearComments(ws As Worksheet)
    Set cmt = ws.Comments
    For Each c In cmt
        c.Delete
    Next
End Function

Function formatColumns(rng As Range, numberOfSteps As Integer)
Dim tmpRange As Range
    Set tmpRange = Range(rng, rng.Offset(1, 3))
    tmpRange.Interior.Color = 13434879 'color columns yellow
    Set tmpRange = Range(tmpRange, rng.Offset(numberOfSteps + 1, 0))
    dummy = drawOutline(tmpRange)
End Function

Function updateStatus(progress As String)
    Application.ScreenUpdating = True
    Application.StatusBar = progress
    Application.ScreenUpdating = False
End Function

Function formatUserEntry(rng As Range)
    For i = 0 To 3
        If (rng.Offset(i, 0) = "") Then
            rng.Offset(i, 0) = "%"
        End If
        rng.Offset(i, 0) = UCase(Trim(rng.Offset(i, 0)))
    Next i
    If (rng.Offset(0, 0) = "%" And rng.Offset(1, 0) = "%" And rng.Offset(2, 0) = "%" And rng.Offset(3, 0) = "%") Then
        MsgBox ("Please specify user search fields, this query will overload the system")
        End
    End If
End Function


Function checkForEmptyQuery(oraDynaset As Object)
    If (oraDynaset.RecordCount = 0) Then
        MsgBox ("No Tools/Chambers found for this query, please revise search parameters")
        End
    End If
End Function

Function drawStepTypeDivders(rng As Range, TotalSteps As Integer) As String
CurrStepType = Mid(rng.Value, 1, 2)
    For i = 0 To (TotalSteps - 1)
        'This IF statement is checking for the Letters in STEP_SEQ, and has a special case for Y step sequences
        If ((CurrStepType <> Mid(rng.Offset(i, 0).Value, 1, 2) And Mid(CurrStepType, 1, 1) <> "Y") Or (Mid(CurrStepType, 1, 1) <> Mid(rng.Offset(i, 0).Value, 1, 1) And Mid(CurrStepType, 1, 1) = "Y")) Then
             Range(rng.Offset(i - 1, -1), rng.Offset(i - 1, -1).End(xlToRight)).Borders(xlEdgeBottom).Weight = xlThin
        End If
        CurrStepType = Mid(rng.Offset(i, 0).Value, 1, 2)
    Next i
End Function

