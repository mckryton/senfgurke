VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_parse_rules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Then_the_feature_contains_INT_rules_5A810804CB2F(step_expressions As Collection)
    'Then the feature contains {integer} rule(s)
    
    Dim rule_count As Integer
    Dim feature_section As Variant
    
    rule_count = 0
    For Each feature_section In m_parsed_feature.sections
        If TypeName(feature_section) = "TRule" Then
            rule_count = rule_count + 1
        End If
    Next
    TSpec.expect(rule_count).to_be step_expressions(1)
End Sub

Public Sub Then_each_rule_has_INT_example_3D7448C173EA(step_expressions As Collection)
    'And each rule has {integer} example
        
    Dim feature_section As Variant
    
    For Each feature_section In m_parsed_feature.sections
        If TypeName(feature_section) = "TRule" Then
            TSpec.expect(feature_section.sections.Count).to_be step_expressions(1)
        End If
    Next
End Sub
