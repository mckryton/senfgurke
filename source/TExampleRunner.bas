Attribute VB_Name = "TExampleRunner"
Option Explicit

Public Const ATTR_STEP_HEAD = "step_head"
Public Const ATTR_STEP_BODY = "step_body"
Public Const ATTR_WHOLE_STEP = "whole_step"
Public Const ATTR_ERROR = "error"

Dim m_logger As Logger
Dim m_is_test_stopped As Boolean

Public Sub run_example(example_steps As Variant, feature_object As Variant)

    Dim step_index As Integer
    Dim step_attributes As Collection

    On Error GoTo error_handler
    TExampleRunner.IsTestStopped = False
    step_index = 0
    Set step_attributes = read_example_step(example_steps, step_index)
    If LCase(step_attributes(ATTR_STEP_HEAD)) = "rule:" Then
        print_rule step_attributes.Item(ATTR_WHOLE_STEP)
        Exit Sub
    End If
    feature_object.before
    print_scenario_title step_attributes.Item(ATTR_WHOLE_STEP)
    step_index = step_index + 1
    Do While TExampleRunner.IsTestStopped = False And step_index <= UBound(example_steps)
        Set step_attributes = read_example_step(example_steps, step_index)
        run_step_line step_attributes, feature_object
        step_index = step_index + 1
    Loop
    If Not TExampleRunner.IsTestStopped Then
        feature_object.after
    End If
    Debug.Print
    Exit Sub
    
error_handler:
    If Err.Number = ERR_ID_SCENARIO_SYNTAX_ERROR Then
        Log.error_log "syntax error: " & Err.description & vbCr & vbLf & "in line >" & step_attributes.Item(ATTR_WHOLE_STEP) & "<"
    Else
        Log.log_function_error "Runtime errror in TExampleRunner.runExample", Join(example_steps, vbTab & vbCr & vbLf)
    End If
End Sub

Private Sub print_rule(pRule As String)
    
    Debug.Print vbTab & pRule
    Debug.Print ""
End Sub

Private Sub print_scenario_title(example_title As String)
    
    If LCase(Left(example_title, Len("Scenario:"))) <> "scenario:" And LCase(Left(example_title, Len("Example:"))) <> "example:" Then
        Err.raise ERR_ID_SCENARIO_SYNTAX_ERROR, description:="can't find scenario start in >" & example_title & "<"
    Else
        Debug.Print vbTab & example_title
    End If
End Sub

Private Sub run_step_line(step_attributes As Collection, feature_object As Variant)

    Dim step_result As Variant

    Select Case step_attributes.Item(ATTR_STEP_HEAD)
    Case "Given", "When", "Then"
        step_result = feature_object.run_step(step_attributes)
        Debug.Print vbTab & step_result(0), vbTab & step_attributes.Item(ATTR_WHOLE_STEP)
        If step_result(0) = "FAILED" Or step_result(0) = "PENDING" Then
            Debug.Print "", step_result(1)
        End If
    Case Else
        Err.raise ERR_ID_SCENARIO_SYNTAX_ERROR, description:="unexpected step type " & step_attributes.Item(ATTR_STEP_HEAD)
    End Select
End Sub

Public Function read_example_step(example As Variant, step_index As Integer) As Collection

    Dim step_attributes As Collection
    Dim prev_step_attributes As Collection
    Dim whole_step As String
    Dim step_head As String     'Given, When, Then, And or But (synonym for and)
    Dim step_body As String
    Dim step_words As Variant
    
    whole_step = Trim(example(step_index))
    step_words = Split(whole_step, " ")
    step_head = step_words(0)
    step_body = Trim(Right(whole_step, Len(whole_step) - Len(step_head)))
    Set step_attributes = New Collection
    With step_attributes
        .Add whole_step, ATTR_WHOLE_STEP
        .Add step_head, ATTR_STEP_HEAD
        .Add step_body, ATTR_STEP_BODY
    End With
    If step_head = "And" Or step_head = "But" Then
        If step_index = 1 Then
            step_attributes.Add "Gherkin syntax error: no previous step to match """ & step_head & """ keyword", ATTR_ERROR
            Set read_example_step = step_attributes
            Exit Function
        Else
            Set prev_step_attributes = read_example_step(example, step_index - 1)
            step_attributes(ATTR_STEP_HEAD) = prev_step_attributes(ATTR_STEP_HEAD)
        End If
    End If
    step_attributes.Add validate_step_head(step_attributes(ATTR_STEP_HEAD)), ATTR_ERROR
    Set read_example_step = step_attributes
End Function
Private Function validate_step_head(step_head) As String

    Select Case step_head
        Case "Given", "When", "Then", "Rule:", "Example:", "Scenario:"
            validate_step_head = ""
        Case Else
            validate_step_head = "Gherkin syntax error: >" & step_head & "< is not a valid step type"
    End Select
End Function

Public Property Get IsTestStopped() As Boolean
    IsTestStopped = m_is_test_stopped
End Property

Private Property Let IsTestStopped(ByVal is_test_stopped As Boolean)
    m_is_test_stopped = is_test_stopped
End Property

Public Sub stop_test()
    TExampleRunner.IsTestStopped = True
End Sub

Public Property Get Log() As Logger
    
    If TypeName(m_logger) = "Nothing" Then
        Set m_logger = New Logger
    End If
    Set Log = m_logger
End Property

