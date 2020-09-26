VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Run_Examples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = ""

Dim m_example As TExample
Dim m_parsed_feature As TFeature
Dim m_example_statistics As Collection

Public Property Get Description() As String

    Description = "Senfgurke will locate the matching step implementation for each step" & vbLf _
                   & "of an example and execute it."
End Property

' in the future this property shall be replaced with a function that parses text
Public Property Get Examples() As Collection

    Dim examples_list As Collection
    Dim example As Variant

    Set examples_list = New Collection
    
    example = Array("Rule: examples are executed until a step returns a status other than ""OK""")
    examples_list.Add example
    example = Array("Example: example has 3 steps and 2nd fails", _
                        "Given an feature with one example with 3 steps where the 2nd step fails", _
                        "When the example is executed", _
                        "Then only two steps of the example were executed")
    examples_list.Add example

    Set Examples = examples_list
End Property

Public Function run_step(step_definition As String) As Variant

    Dim step_attributes As Collection
    Dim feature_clause As Collection
    
    On Error GoTo error_handler
    Select Case step_definition
                  
        Case "Given an feature with one example with 3 steps where the 2nd step fails"
            Set m_parsed_feature = TFeatureParser.parse_feature( _
                "Feature: sample feature" & vbLf & _
                vbLf & _
                "  Example: sample example" & vbLf & _
                "    Given a valid step" & vbLf & _
                "    And an invalid step" & vbLf & _
                "    And a valid step")
            
        Case "When the example is executed"
            Set m_example_statistics = TExampleRunner.run_example(m_parsed_feature.Clauses(1), silent:=True)
        
        Case "Then only two steps of the example were executed"
            TSpec.expect(m_example_statistics.Count).to_be 2
        
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array(STEP_RESULT_OK)
    Exit Function

error_handler:
    run_step = TStepRunner.fail_step(Err.Number, Err.Description)
End Function

Public Sub before_example()
    
End Sub

Public Sub after_example()

End Sub

Public Property Get tags() As Variant
    tags = cTags
End Property

Private Sub pending(pending_msg)
    TStepRunner.pending pending_msg
End Sub

Public Sub Given_to_long_name()
'Public Sub Given_an_feature_with_one_example_with_3_steps_where_the_2nd_step_fails()
    Set m_parsed_feature = TFeatureParser.parse_feature( _
                    "Feature: sample feature" & vbLf & _
                    vbLf & _
                    "  Example: sample example" & vbLf & _
                    "    Given a valid step" & vbLf & _
                    "    And an invalid step" & vbLf & _
                    "    And a valid step")
End Sub
