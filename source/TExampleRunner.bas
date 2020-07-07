Attribute VB_Name = "TExampleRunner"
'------------------------------------------------------------------------
' Description  : execute test steps for Gherkin scenarios / examples
'------------------------------------------------------------------------

Option Explicit

Dim mLogger As logger
Dim mTestStopped As Boolean

Public Sub run_example(pExampleLinesArray As Variant, pTestDefinitionObject As Variant)

    Dim intLineIndex As Integer
    Dim colLine As Collection
    Dim strLastStepType As String

    On Error GoTo error_handler
    TExampleRunner.TestStopped = False
    intLineIndex = 0
    Set colLine = getExampleLine(pExampleLinesArray, intLineIndex)
    If LCase(colLine("line_head")) = "rule:" Then
        print_rule colLine.Item("line")
        Exit Sub
    End If
    pTestDefinitionObject.before
    print_scenario_title colLine.Item("line")
    intLineIndex = intLineIndex + 1
    Do While TExampleRunner.TestStopped = False And intLineIndex <= UBound(pExampleLinesArray)
        Set colLine = getExampleLine(pExampleLinesArray, intLineIndex)
        If colLine.Item("line_head") <> "And" Then
            strLastStepType = colLine.Item("line_head")
        End If
        colLine.Remove "step_type"
        colLine.Add strLastStepType, "step_type"
        run_step_line colLine, pTestDefinitionObject
        intLineIndex = intLineIndex + 1
    Loop
    If Not TExampleRunner.TestStopped Then
        pTestDefinitionObject.after
    End If
    
    Debug.Print
    Exit Sub
    
error_handler:
    If Err.Number = ERR_ID_SCENARIO_SYNTAX_ERROR Then
        Log.error_log "syntax error: " & Err.description & vbCr & vbLf & "in line >" & colLine.Item("line") & "<"
    Else
        Log.log_function_error "TExampleRunner.runExample", Join(pExampleLinesArray, vbTab & vbCr & vbLf)
    End If
End Sub

Private Sub print_rule(pRule As String)
    
    Debug.Print vbTab & pRule
    Debug.Print ""
End Sub

Private Sub print_scenario_title(pExampleTitle As String)
    
    If LCase(Left(pExampleTitle, Len("Scenario:"))) <> "scenario:" And LCase(Left(pExampleTitle, Len("Example:"))) <> "example:" Then
        Err.raise ERR_ID_SCENARIO_SYNTAX_ERROR, description:="can't find scenario start in >" & pExampleTitle & "<"
    Else
        Debug.Print vbTab & pExampleTitle
    End If
End Sub

Private Sub run_step_line(pStepLine As Collection, pobjTestDefinition As Variant)

    Dim step_result As Variant

    Select Case pStepLine.Item("step_type")
    Case "Given", "When", "Then"
        step_result = pobjTestDefinition.run_step(pStepLine)
        Debug.Print vbTab & step_result(0), vbTab & pStepLine.Item("line")
        If step_result(0) = "FAILED" Or step_result(0) = "PENDING" Then
            Debug.Print "", step_result(1)
        End If
    Case Else
        Err.raise ERR_ID_SCENARIO_SYNTAX_ERROR, description:="unexpected step type " & pStepLine.Item("step_type")
    End Select
End Sub

Public Function getExampleLine(pvarExample As Variant, pintLineIndex As Integer) As Collection

    Dim colLineProps As Collection
    Dim strLine As String     'the whole line
    Dim strStepType As String 'e.g. Given, When, Then, And
    Dim strStepDef As String  ' everything behind step type
    Dim varWords As Variant   'all words of the line as array
    
    On Error GoTo error_handler
    strLine = Trim(pvarExample(pintLineIndex))
    varWords = Split(strLine, " ")
    strStepType = varWords(0)
    strStepDef = Right(strLine, Len(strLine) - Len(strStepType))
    Set colLineProps = New Collection
    With colLineProps
        .Add strLine, "line"
        .Add strStepType, "line_head"
        .Add strStepDef, "line_body"
        .Add vbNullString, "step_type"      'step type depends on context, e.g. previous steps
    End With
    Set getExampleLine = colLineProps
    Exit Function

error_handler:
    Log.log_function_error "TExampleRunner.getExampleLine"
End Function

Public Property Get TestStopped() As Boolean
    TestStopped = mTestStopped
End Property

Private Property Let TestStopped(ByVal pTestStopped As Boolean)
    mTestStopped = pTestStopped
End Property

Public Sub stop_test()
    TExampleRunner.TestStopped = True
End Sub

Public Sub pending(pPendingMsg)
    
    Debug.Print "PENDING: " & pPendingMsg
    stop_test
End Sub

Public Property Get Log() As logger
    
    If TypeName(mLogger) = "Nothing" Then
        Set mLogger = New logger
    End If
    Set Log = mLogger
End Property

