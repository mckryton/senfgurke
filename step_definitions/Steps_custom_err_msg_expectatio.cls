VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_custom_err_msg_expectatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_an_expected_value_was_defined_as_INT_AA730CC5F5BC(example_context As TContext, step_expressions As Collection)
    'Given an expected value was defined as {integer}
    example_context.set_value step_expressions(1), "expected_val"
End Sub

Public Sub Given_the_actual_value_was_evaluated_as_INT_C1726D898F77(example_context As TContext, step_expressions As Collection)
    'And the actual value was evaluated as {integer}
    example_context.set_value step_expressions(1), "actual_val"
End Sub

Public Sub Given_a_custom_error_message_was_defined_as_S_05173C9D82CE(example_context As TContext, step_expressions As Collection)
    'And a custom error message was defined as {string}
    example_context.set_value step_expressions(1), "custom_err_msg"
End Sub

Public Sub Given_a_collection_was_used_to_save_cell_back_74F395E30263(example_context As TContext, step_expressions As Collection)
    'Given a collection was used to save cell background color names {string} , {string} and {string}
    Dim sample_collection As Collection
    
    Set sample_collection = New Collection
    With sample_collection
        .Add step_expressions(1)
        .Add step_expressions(2)
        .Add step_expressions(3)
    End With
    example_context.set_value sample_collection, "sample_collection"
End Sub

Public Sub Given_a_array_was_used_to_save_cell_backgroun_1EFBF20EF10A(example_context As TContext, step_expressions As Collection)
    'Given a array was used to save cell background color names {string} , {string} and {string}
    Dim sample_array(3) As String
    
    sample_array(0) = step_expressions(1)
    sample_array(1) = step_expressions(2)
    sample_array(2) = step_expressions(3)
    example_context.set_value sample_array, "sample_array"
End Sub

Public Sub When_expected_and_actual_value_are_being_comp_04DEE404A552(example_context As TContext)
    'When expected and actual value are being compared using to_be
    Dim actual_err_msg As String
    Dim given_err_msg As String
    Dim expected_val As Variant
    Dim actual_val As Variant
    
    actual_err_msg = vbNullString
    given_err_msg = example_context.get_value("custom_err_msg")
    expected_val = example_context.get_value("expected_val")
    actual_val = example_context.get_value("actual_val")
    On Error GoTo comparison_failed
    TSpec.expect(actual_val).to_be expected_val, given_err_msg
    On Error GoTo 0
    example_context.set_value actual_err_msg, "actual_err_msg"
    Exit Sub
    
comparison_failed:
    'sometimes VBA will overwrite any custom set Err.description property, therefore the err msg is copied to TSpec.LastFailMsg
    actual_err_msg = TSpec.LastFailMsg
    'escape linebreaks and tabs to fit the actual err msg into one line
    actual_err_msg = escape_whitespace_formatting(actual_err_msg)
    Err.Clear
    Resume Next
End Sub

Public Sub When_expected_and_actual_value_are_being_comp_7F8492096565(example_context As TContext)
    'When expected and actual value are being compared using not_to_be
    Dim actual_err_msg As String
    Dim given_err_msg As String
    Dim expected_val As Variant
    Dim actual_val As Variant
    
    actual_err_msg = vbNullString
    given_err_msg = example_context.get_value("custom_err_msg")
    expected_val = example_context.get_value("expected_val")
    actual_val = example_context.get_value("actual_val")
    On Error GoTo comparison_failed
    TSpec.expect(actual_val).not_to_be expected_val, given_err_msg
    On Error GoTo 0
    example_context.set_value actual_err_msg, "actual_err_msg"
    Exit Sub
    
comparison_failed:
    'sometimes VBA will overwrite any custom set Err.description property, therefore the err msg is copied to TSpec.LastFailMsg
    actual_err_msg = TSpec.LastFailMsg
    'escape linebreaks and tabs to fit the actual err msg into one line
    actual_err_msg = escape_whitespace_formatting(actual_err_msg)
    Err.Clear
    Resume Next
End Sub

Public Sub When_the_expectation_that_the_collection_cont_2D8FB05B11AD(example_context As TContext, step_expressions As Collection)
    'When the expectation that the collection contains {string} as member is validated
    Dim sample_collection As Collection
    Dim search_item As String
    Dim actual_err_msg As String
    Dim given_err_msg As String
     
    Set sample_collection = example_context.get_value("sample_collection")
    search_item = step_expressions(1)
    actual_err_msg = vbNullString
    given_err_msg = example_context.get_value("custom_err_msg")
    On Error GoTo comparison_failed
    TSpec.expect(sample_collection).contains_member search_item, given_err_msg
    On Error GoTo 0
    example_context.set_value actual_err_msg, "actual_err_msg"
    Exit Sub
    
comparison_failed:
    'sometimes VBA will overwrite any custom set Err.description property, therefore the err msg is copied to TSpec.LastFailMsg
    actual_err_msg = TSpec.LastFailMsg
    'escape linebreaks and tabs to fit the actual err msg into one line
    actual_err_msg = escape_whitespace_formatting(actual_err_msg)
    Err.Clear
    Resume Next
End Sub

Public Sub When_the_expectation_that_the_array_contains__998B0E969573(example_context As TContext, step_expressions As Collection)
    'When the expectation that the array contains {string} as member is validated
    Dim sample_array() As String 'to be able to retrieve the array from example context this variable must not declare a fixed size array
    Dim search_item As String
    Dim actual_err_msg As String
    Dim given_err_msg As String
     
    sample_array = example_context.get_value("sample_array")
    search_item = step_expressions(1)
    actual_err_msg = vbNullString
    given_err_msg = example_context.get_value("custom_err_msg")
    On Error GoTo comparison_failed
    TSpec.expect(sample_array).contains_member search_item, given_err_msg
    On Error GoTo 0
    example_context.set_value actual_err_msg, "actual_err_msg"
    Exit Sub
    
comparison_failed:
    'sometimes VBA will overwrite any custom set Err.description property, therefore the err msg is copied to TSpec.LastFailMsg
    actual_err_msg = TSpec.LastFailMsg
    'escape linebreaks and tabs to fit the actual err msg into one line
    actual_err_msg = escape_whitespace_formatting(actual_err_msg)
    Err.Clear
    Resume Next
End Sub

Public Sub Then__STR_is_added_to_the_expectation_error_m_4ED53D681EDD(example_context As TContext, step_expressions As Collection)
    'Then  {string} is added to the expectation error message
    Dim expected_err_msg As String
    Dim actual_err_msg As String
    
   expected_err_msg = step_expressions(1)
   actual_err_msg = example_context.get_value("actual_err_msg")
   TSpec.expect(actual_err_msg).ends_with_text expected_err_msg
End Sub

Private Function escape_whitespace_formatting(err_msg As String) As String
    Dim escaped_msg As String
    
    escaped_msg = Replace(err_msg, vbLf, "\n")
    escaped_msg = Replace(escaped_msg, vbTab, "\t")
    escape_whitespace_formatting = escaped_msg
End Function
