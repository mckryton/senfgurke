VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_parse_tags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Then_the_parsed_features_contains_the_tags_STR_013B39CE37F4(example_context As TContext, step_expressions As Collection)
    'Then the parsed features contains the tags {string}
    
    Dim tags As Variant
    Dim tag As Variant
    
    tags = Split(step_expressions(1), ",")
    For Each tag In tags
        TSpec.expect(example_context.get_value("parsed_feature").tags).contains_member Trim(tag)
    Next
End Sub

Public Sub Then_the_parsed_example_contains_the_tags_STR_013B99DC06E4(example_context As TContext, step_expressions As Collection)
    'And the parsed example contains the tags {string}
    
    Dim tags As Variant
    Dim tag As Variant
    
    tags = Split(step_expressions(1), ",")
    For Each tag In tags
        TSpec.expect(example_context.get_value("example").tags).contains_member Trim(tag)
    Next
End Sub

Public Sub Then_the_parsed_features_contains_INT_examples_013B39CEE3B0(example_context As TContext, step_expressions As Collection)
    'Then the parsed features contains {integer} examples
    TSpec.expect(example_context.get_value("parsed_feature").sections.Count).to_be 2
    TSpec.expect(TypeName(example_context.get_value("parsed_feature").sections(2))).to_be "TExample"
    
    example_context.set_value example_context.get_value("parsed_feature").sections(2), "example"
End Sub

Public Sub Then_the_parsed_second_example_is_tagged_with_STR_EB00E4656DAE(example_context As TContext, step_expressions As Collection)
    'And the parsed second example is tagged with {string}
        
    Dim tags As Variant
    Dim tag As Variant
    
    tags = Split(step_expressions(1), ",")
    For Each tag In tags
        TSpec.expect(example_context.get_value("example").tags).contains_member Trim(tag)
    Next
End Sub

Public Sub Then_the_parsed_rule_contains_the_tags_STR_7B67E25E94E1(example_context As TContext, step_expressions As Collection)
    'Then the parsed rule contains the tags {string}
    
    Dim rule As TRule
    Dim tags As Variant
    Dim tag As Variant
    
    Set rule = example_context.get_value("parsed_feature").sections(1)
    tags = Split(step_expressions(1), ",")
    For Each tag In tags
        TSpec.expect(rule.tags).contains_member Trim(tag)
    Next
End Sub

Public Sub Then_the_included_example_contains_the_tags_STR_33546FCE3943(example_context As TContext, step_expressions As Collection)
    'And the included example contains the tags {string}
    
    Dim included_example As TExample
    Dim tags As Variant
    Dim tag As Variant
    
    If TypeName(example_context.get_value("parsed_feature").sections(1)) = "TExample" Then
        'example is 1st section of the feature
        Set included_example = example_context.get_value("parsed_feature").sections(1)
    Else
        '1st section of the feature is a rule with one example
        Set included_example = example_context.get_value("parsed_feature").sections(1).sections(1)
    End If
    tags = Split(step_expressions(1), ",")
    For Each tag In tags
        TSpec.expect(included_example.tags).contains_member Trim(tag)
    Next
End Sub

