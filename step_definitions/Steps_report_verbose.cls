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

Public Sub Given_a_feature_STR_was_reported_B5EB56508778(example_context As TContext, step_expressions As Collection)
    'Given a feature {string} was reported
    example_context.get_value("session").reporter.report REPORT_MSG_TYPE_FEATURE_NAME, "Feature: " & step_expressions(1)
End Sub

Public Sub Given_a_rule_STR_was_reported_A5D68D23D5AE(example_context As TContext, step_expressions As Collection)
    'Given a rule {string} was reported
    example_context.get_value("session").reporter.report REPORT_MSG_TYPE_RULE, "Rule: " & step_expressions(1)
End Sub

Public Sub Given_an_example_STR_was_reported_9C2B7078D5AE(example_context As TContext, step_expressions As Collection)
    'Given an example {string} was reported
    example_context.get_value("session").reporter.report REPORT_MSG_TYPE_EXAMPLE_TITLE, "Example: " & step_expressions(1)
End Sub

Public Sub When_a_step_STR_followed_by_a_docstring_STR_is_rep_8C80B341FF25(example_context As TContext, step_expressions As Collection)
    'When a step {string} followed by a docstring {string} is reported with status {string}
    Dim step_name As String
    
    step_name = step_expressions(1) & vbLf & """""""" & vbLf & step_expressions(2) & vbLf & """"""""
    example_context.get_value("session").reporter.report REPORT_MSG_TYPE_STEP, step_name, step_expressions(3)
End Sub

Public Sub When_the_execution_of_the_first_example_step_is_re_7D6F60A0AA78(example_context As TContext)
    'When the execution of the first example step is reported
    example_context.get_value("session").reporter.report REPORT_MSG_TYPE_STEP, "Given sample step", STATUS_OK
End Sub

Public Sub When_the_execution_of_the_feature_is_reported_7A6F8A00CA5C(example_context As TContext)
    'When the execution of the feature is reported
    Dim parsed_feature As TFeature
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    Set parsed_feature = TFeatureParser.parse_feature(example_context.get_value("gherkin_text"))
    example_context.set_value parsed_feature, "parsed_feature"
    session.reporter.report REPORT_MSG_TYPE_FEATURE_NAME, parsed_feature.name
    session.reporter.report REPORT_MSG_TYPE_DESC, parsed_feature.description
End Sub

Public Sub When_the_example_is_reported_with_the_step_result__F52F4459A058(example_context As TContext, step_expressions As Collection)
    'And the example is reported with the step result {string}
    Dim example As TExample
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    Set example = example_context.get_value("parsed_feature").sections(1)
    session.reporter.report REPORT_MSG_TYPE_EXAMPLE_TITLE, example.OriginalHeadline
    session.reporter.report REPORT_MSG_TYPE_STEP, example.steps(1).OriginalStepDefinition, step_expressions(1), vbNullString
End Sub

Public Sub Then_the_report_output_contains_STR_5A8ED3F2F3C6(example_context As TContext, step_expressions As Collection)
    'Then the report output contains {string}
    TSpec.expect(example_context.get_value("session").reporter.ConsoleOutput).includes_text step_expressions(1)
End Sub

Public Sub Then_the_report_output_doesnt_contain_STR_77804E212494(example_context As TContext, step_expressions As Collection)
    'Then the report output doesn't contain {string}
    TSpec.expect(example_context.get_value("session").reporter.ConsoleOutput).does_not_include step_expressions(1)
End Sub

Public Sub Then_line_INT_of_the_resulting_output_is_STR_A434511445AD(example_context As TContext, step_expressions As Collection)
    'Then line {integer} of the resulting output is {string}
    Dim lines As Variant
    
    lines = Split(example_context.get_value("session").reporter.ConsoleOutput, vbLf)
    TSpec.expect(lines(step_expressions(1) - 1)).to_be step_expressions(2)
End Sub
