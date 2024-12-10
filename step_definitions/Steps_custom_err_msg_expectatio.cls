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

Public Sub Given_an_text_was_expected_to_include_STR_84A9E7DB025F(example_context As TContext, step_expressions As Collection)
    'Given an text was expected to include {string}
    Dim expected_text As String
    
    expected_text = step_expressions(1)
    example_context.set_value expected_text, "expected_text"
End Sub

Public Sub Given_the_actual_text_was_STR_5F91A842132F(example_context As TContext, step_expressions As Collection)
    'But the actual text was {string}
    Dim actual_text As String
    
    actual_text = step_expressions(1)
    example_context.set_value actual_text, "actual_text"
End Sub

Public Sub When_expected_and_actual_value_are_being_comp_04DEE404A552(example_context As TContext)
    'When expected and actual value are being compared using to_be
    Dim actual_err_msg As String
    Dim custom_err_msg As String
    Dim expected_val As Variant
    Dim actual_val As Variant
    
    expected_val = example_context.get_value("expected_val")
    actual_val = example_context.get_value("actual_val")
    custom_err_msg = example_context.get_value("custom_err_msg")
    actual_err_msg = run_comparison_expectation(actual_val, expected_val, custom_err_msg, "to_be")
    example_context.set_value actual_err_msg, "actual_err_msg"
End Sub

Public Sub When_expected_and_actual_value_are_being_comp_7F8492096565(example_context As TContext)
    'When expected and actual value are being compared using not_to_be
    Dim actual_err_msg As String
    Dim custom_err_msg As String
    Dim expected_val As Variant
    Dim actual_val As Variant
    
    expected_val = example_context.get_value("expected_val")
    actual_val = example_context.get_value("actual_val")
    custom_err_msg = example_context.get_value("custom_err_msg")
    actual_err_msg = run_comparison_expectation(actual_val, expected_val, custom_err_msg, "not_to_be")
    example_context.set_value actual_err_msg, "actual_err_msg"
End Sub

Public Sub When_the_text_is_compared_using_starts_with_A93679E7575B(example_context As TContext)
    'When the text is compared using starts_with
    Dim expected_text As String
    Dim actual_text As String
    Dim actual_err_msg As String
    Dim custom_err_msg As String
    
    expected_text = example_context.get_value("expected_text")
    actual_text = example_context.get_value("actual_text")
    custom_err_msg = example_context.get_value("custom_err_msg")
    actual_err_msg = run_comparison_expectation(actual_text, expected_text, custom_err_msg, "starts_with")
    example_context.set_value actual_err_msg, "actual_err_msg"
End Sub

Public Sub When_the_text_is_compared_using_ends_with_F7C2E74C69E2(example_context As TContext)
    'When the text is compared using ends_with
    Dim expected_text As String
    Dim actual_text As String
    Dim actual_err_msg As String
    Dim custom_err_msg As String
    
    expected_text = example_context.get_value("expected_text")
    actual_text = example_context.get_value("actual_text")
    custom_err_msg = example_context.get_value("custom_err_msg")
    actual_err_msg = run_comparison_expectation(actual_text, expected_text, custom_err_msg, "ends_with")
    example_context.set_value actual_err_msg, "actual_err_msg"
End Sub

Public Sub When_the_text_is_compared_using_includes_text_38E99F35B921(example_context As TContext)
    'When the text is compared using includes_text
    Dim expected_text As String
    Dim actual_text As String
    Dim actual_err_msg As String
    Dim custom_err_msg As String
    
    expected_text = example_context.get_value("expected_text")
    actual_text = example_context.get_value("actual_text")
    custom_err_msg = example_context.get_value("custom_err_msg")
    actual_err_msg = run_comparison_expectation(actual_text, expected_text, custom_err_msg, "includes_text")
    example_context.set_value actual_err_msg, "actual_err_msg"
End Sub

Public Sub When_the_expectation_that_the_collection_cont_2D8FB05B11AD(example_context As TContext, step_expressions As Collection)
    'When the expectation that the collection contains {string} as member is validated
    Dim sample_collection As Collection
    Dim search_item As String
    Dim actual_err_msg As String
    Dim custom_err_msg As String
     
    Set sample_collection = example_context.get_value("sample_collection")
    search_item = step_expressions(1)
    custom_err_msg = example_context.get_value("custom_err_msg")
    actual_err_msg = run_search_expectation(sample_collection, search_item, custom_err_msg, "contains_member")
    example_context.set_value actual_err_msg, "actual_err_msg"
End Sub

Public Sub When_the_expectation_that_the_array_contains__998B0E969573(example_context As TContext, step_expressions As Collection)
    'When the expectation that the array contains {string} as member is validated
    Dim sample_array() As String 'to be able to retrieve the array from example context this variable must not declare a fixed size array
    Dim search_item As String
    Dim actual_err_msg As String
    Dim custom_err_msg As String
    
    sample_array = example_context.get_value("sample_array")
    search_item = step_expressions(1)
    custom_err_msg = example_context.get_value("custom_err_msg")
    actual_err_msg = run_search_expectation(sample_array, search_item, custom_err_msg, "contains_member")
    example_context.set_value actual_err_msg, "actual_err_msg"
End Sub

Public Sub Then__STR_is_added_to_the_expectation_error_m_4ED53D681EDD(example_context As TContext, step_expressions As Collection)
    'Then  {string} is added to the expectation error message
    Dim expected_err_msg As String
    Dim actual_err_msg As String
    
   expected_err_msg = step_expressions(1)
   actual_err_msg = example_context.get_value("actual_err_msg")
   TSpec.expect(actual_err_msg).ends_with_text expected_err_msg
End Sub

Private Function run_search_expectation(search_list As Variant, search_item As Variant, custom_err_msg As String, search_type As String) As String
    Dim actual_err_msg As String
    
    'set the actual err msg to a default value to recognice if the expectation under test caused no error at all
    actual_err_msg = "Expected an error being raised by a missed expectation but none found so far..."
    On Error GoTo search_failed
    Select Case search_type
        Case "contains_member"
            TSpec.expect(search_list).contains_member search_item, custom_err_msg
    End Select
    On Error GoTo 0
    run_search_expectation = actual_err_msg
    Exit Function
    
search_failed:
    actual_err_msg = get_failure_msg
    Err.Clear
    Resume Next
End Function

Private Function run_comparison_expectation(actual_value As Variant, expected_value As Variant, custom_err_msg As String, comparison_type As String) As String
    Dim actual_err_msg As String
    
    'set the actual err msg to a default value to recognice if the expectation under test caused no error at all
    actual_err_msg = "Expected an error being raised by a missed expectation but none found so far..."
    On Error GoTo comparison_failed
    Select Case comparison_type
        Case "starts_with"
            TSpec.expect(actual_value).starts_with_text expected_value, custom_err_msg
        Case "ends_with"
            TSpec.expect(actual_value).ends_with_text expected_value, custom_err_msg
        Case "includes_text"
            TSpec.expect(actual_value).includes_text expected_value, custom_err_msg
        Case "to_be"
            TSpec.expect(actual_value).to_be expected_value, custom_err_msg
        Case "not_to_be"
            TSpec.expect(actual_value).not_to_be expected_value, custom_err_msg
    End Select
    On Error GoTo 0
    run_comparison_expectation = actual_err_msg
    Exit Function
    
comparison_failed:
    actual_err_msg = get_failure_msg
    Err.Clear
    Resume Next
End Function

Private Function get_failure_msg()
    Dim err_msg As String
    Dim escaped_msg As String
    
    'sometimes VBA will overwrite any custom set Err.description property, therefore the err msg is copied to TSpec.LastFailMsg
    err_msg = TSpec.LastFailMsg
    'escape linebreaks and tabs to fit the actual err msg into one line
    escaped_msg = Replace(err_msg, vbLf, "\n")
    escaped_msg = Replace(escaped_msg, vbTab, "\t")
    get_failure_msg = escaped_msg
End Function
