VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_report_verbose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub When_the_following_events_are_reported_as_a_result_0232CC26FE97(example_context As TContext, data_table As TDataTable)
    'When the following events are reported as a result of running a feature
    Dim session As TSession
    Dim raised_event As Variant
    
    Set session = example_context.get_value("session")
    For Each raised_event In data_table.table_rows
        Select Case LCase(raised_event("section_type"))
            'reporting is always about STARTED events except for steps where the status is only available for the FINISHED event
            Case "feature"
                session.log_event EVENT_RUN_FEATURE_STARTED, vbNullString, raised_event("section_name"), SECTION_TYPE_FEATURE
            Case "rule"
                session.log_event EVENT_RUN_RULE_STARTED, vbNullString, raised_event("section_type") & ": " & raised_event("section_name"), SECTION_TYPE_RULE
            Case "example"
                session.log_event EVENT_RUN_EXAMPLE_STARTED, vbNullString, raised_event("section_type") & ": " & raised_event("section_name"), SECTION_TYPE_EXAMPLE
            Case "step"
                session.log_event EVENT_RUN_STEP_FINISHED, vbNullString, raised_event("section_name"), SECTION_TYPE_STEP, raised_event("section_status")
        End Select
    Next
End Sub

Public Sub When_a_step_STR_followed_by_a_docstring_STR_is_rep_8C80B341FF25(example_context As TContext, step_expressions As Collection)
    'When a step {string} followed by a docstring {string} is reported with status {string}
    Dim step_name As String
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    step_name = step_expressions(1) & vbLf & """""""" & vbLf & step_expressions(2) & vbLf & """"""""
    session.log_event EVENT_RUN_STEP_FINISHED, vbNullString, step_name, SECTION_TYPE_STEP, step_expressions(3)
End Sub

Public Sub When_the_execution_of_the_feature_is_reported_7A6F8A00CA5C(example_context As TContext)
    'When the execution of the feature is reported
    Dim parsed_feature As TFeature
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    Set parsed_feature = TFeatureParser.parse_feature(example_context.get_value("gherkin_text"))
    example_context.set_value parsed_feature, "parsed_feature"
    session.log_event EVENT_RUN_FEATURE_STARTED, parsed_feature.description, parsed_feature.Name, SECTION_TYPE_FEATURE
End Sub

Public Sub When_the_example_is_reported_with_the_step_result__F52F4459A058(example_context As TContext, step_expressions As Collection)
    'And the example is reported with the step result {string}
    Dim Example As TExample
    Dim step As TStep
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    Set Example = example_context.get_value("parsed_feature").sections(1)
    Set step = Example.steps(1)
    session.log_event EVENT_RUN_EXAMPLE_STARTED, Example.description, Example.OriginalHeadline, SECTION_TYPE_EXAMPLE
    session.log_event EVENT_RUN_STEP_FINISHED, vbNullString, step.OriginalStepDefinition, SECTION_TYPE_STEP, step_expressions(1)
End Sub

Public Sub When_the_following_step_is_reported_with_the_statu_2B48A62AC066(example_context As TContext, step_expressions As Collection)
    'When the following step is reported with the status {string}
    Dim step As TStep
    Dim step_list As Collection
    Dim parent_feature As TFeature
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    Set parent_feature = THelper.new_TFeature
    parent_feature.parsed_lines = -1
    Set step_list = TStepParser.parse_step_list(step_expressions(2), parent_feature)
    Set step = step_list(1)
    session.log_event EVENT_RUN_STEP_FINISHED, vbNullString, step.OriginalStepDefinition, SECTION_TYPE_STEP, step_expressions(1)
End Sub

Public Sub Then_the_report_output_contains_STR_5A8ED3F2F3C6(example_context As TContext, step_expressions As Collection)
    'Then the report output contains {string}
    TSpec.expect(example_context.get_value("session").reporter.ConsoleOutput).includes_text step_expressions(1)
End Sub

Public Sub Then_the_report_output_doesnt_contain_STR_77804E212494(example_context As TContext, step_expressions As Collection)
    'Then the report output doesn't contain {string}
    TSpec.expect(example_context.get_value("session").reporter.ConsoleOutput).does_not_include step_expressions(1)
End Sub

Public Sub Then_the_report_output_is_STR_7C9B2337B419(example_context As TContext, step_expressions As Collection)
    'Then the report output is {string}
    TSpec.expect(ExtraVBA.align_textblock(example_context.get_value("session").reporter.ConsoleOutput)).to_be step_expressions(1) & vbLf
End Sub

Public Sub Then_line_INT_of_the_resulting_output_is_STR_A434511445AD(example_context As TContext, step_expressions As Collection)
    'Then line {integer} of the resulting output is {string}
    Dim lines As Variant
    
    lines = Split(example_context.get_value("session").reporter.ConsoleOutput, vbLf)
    TSpec.expect(lines(step_expressions(1) - 1)).to_be step_expressions(2)
End Sub
