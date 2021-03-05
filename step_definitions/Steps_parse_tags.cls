VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_parse_tags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Then_the_parsed_features_contains_the_tags_STR_013B39CE37F4(step_expressions As Collection)
    'Then the parsed features contains the tags {string}
    
    Dim tags As Variant
    Dim tag As Variant
    
    tags = Split(step_expressions(1), ",")
    For Each tag In tags
        TSpec.expect(m_parsed_feature.tags(Trim(tag))).to_be Trim(tag)
    Next
End Sub

Public Sub Then_the_parsed_example_contains_the_tags_STR_013B99DC06E4(step_expressions As Collection)
    'And the parsed example contains the tags {string}
    
    Dim tags As Variant
    Dim tag As Variant
    
    tags = Split(step_expressions(1), ",")
    For Each tag In tags
        TSpec.expect(TStepVars.example.tags(Trim(tag))).to_be Trim(tag)
    Next
End Sub

Public Sub Then_the_parsed_features_contains_INT_examples_013B39CEE3B0(step_expressions As Collection)
    'Then the parsed features contains {integer} examples
    TSpec.expect(m_parsed_feature.Clauses.Count).to_be 2
    TSpec.expect(TypeName(m_parsed_feature.Clauses(2))).to_be "TExample"
    Set TStepVars.example = m_parsed_feature.Clauses(2)
End Sub

Public Sub Then_the_parsed_second_example_contains_the_tags_S_E5AA0E4C4719(step_expressions As Collection)
    'And the parsed second example contains the tags {string}
        
    Dim tags As Variant
    Dim tag As Variant
    
    tags = Split(step_expressions(1), ",")
    For Each tag In tags
        TSpec.expect(TStepVars.example.tags(Trim(tag))).to_be Trim(tag)
    Next
End Sub