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
    Dim step_result As Variant
    Dim err_msg As String

    On Error GoTo error_handler
    TExampleRunner.IsTestStopped = False
    step_index = 0
    Set step_attributes = read_step(example_steps, step_index)
    If LCase(step_attributes(ATTR_STEP_HEAD)) = "rule:" Then
        TReport.report TReport.TYPE_RULE, step_attributes.Item(ATTR_WHOLE_STEP)
        Exit Sub
    End If
    feature_object.before_example
    validate_example_title_syntax step_attributes.Item(ATTR_WHOLE_STEP)
    TReport.report TReport.TYPE_EXAMPLE_TITLE, step_attributes.Item(ATTR_WHOLE_STEP)
    step_index = step_index + 1
    Do While TExampleRunner.IsTestStopped = False And step_index <= UBound(example_steps)
        Set step_attributes = read_step(example_steps, step_index)
        step_result = execute_step(step_attributes, feature_object)
        err_msg = vbNullString
        If Not step_result(0) = "OK" Then
            err_msg = step_result(1)
            TExampleRunner.stop_example
        End If
        TReport.report TReport.TYPE_STEP, step_attributes.Item(ATTR_WHOLE_STEP), step_result(0), err_msg
        step_index = step_index + 1
    Loop
    If Not TExampleRunner.IsTestStopped Then
        feature_object.after_example
    End If
    Exit Sub
    
error_handler:
    If Err.Number = ERR_ID_SCENARIO_SYNTAX_ERROR Then
        log.error_log "syntax error: " & Err.description & vbCr & vbLf & "in line >" & step_attributes.Item(ATTR_WHOLE_STEP) & "<"
    Else
        log.log_function_error "Runtime errror in TExampleRunner.runExample", Join(example_steps, vbTab & vbCr & vbLf)
    End If
End Sub

Private Sub validate_example_title_syntax(example_title As String)
    If LCase(Left(example_title, Len("Scenario:"))) <> "scenario:" And LCase(Left(example_title, Len("Example:"))) <> "example:" Then
        Err.Raise ERR_ID_SCENARIO_SYNTAX_ERROR, description:="can't find scenario start in >" & example_title & "<"
    End If
End Sub

Public Function execute_step(step_attributes As Collection, feature_object As Variant) As Variant

    Dim step_result As Variant

    Select Case step_attributes.Item(ATTR_STEP_HEAD)
    Case "Given", "When", "Then"
        step_result = feature_object.run_step(step_attributes)
        execute_step = step_result
    Case Else
        Err.Raise ERR_ID_SCENARIO_SYNTAX_ERROR, description:="unexpected step type " & step_attributes.Item(ATTR_STEP_HEAD)
    End Select
End Function

Public Function read_step(example As Variant, step_index As Integer) As Collection

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
            Set read_step = step_attributes
            Exit Function
        Else
            Set prev_step_attributes = read_step(example, step_index - 1)
            'collection objects do not support item updates
            step_attributes.Remove ATTR_STEP_HEAD
            step_attributes.Add prev_step_attributes(ATTR_STEP_HEAD), ATTR_STEP_HEAD
        End If
    End If
    step_attributes.Add validate_step_head(step_attributes(ATTR_STEP_HEAD)), ATTR_ERROR
    Set read_step = step_attributes
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

Public Sub stop_example()
    TExampleRunner.IsTestStopped = True
End Sub

Public Property Get log() As Logger
    
    If TypeName(m_logger) = "Nothing" Then
        Set m_logger = New Logger
    End If
    Set log = m_logger
End Property

