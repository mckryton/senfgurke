VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Parse_Features"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = "wip"
Dim m_description As String

'module variables used for steps
Dim m_feature_text As String
Dim m_example As Variant
Dim m_step_attributes As Collection
Dim m_parsed_feature As Collection

Public Property Get Description() As String
    
    If m_description = vbNullString Then
        m_description = "Senfgurke will read feature specs from text" & vbLf _
                            & "and identify it's elements like descriptions, rules and examples"
    End If
    Description = m_description
End Property

' in the future this property shall be replaced with a function that parses text
Public Property Get Examples() As Collection

    Dim examples_list As Collection
    Dim example As Variant

    Set examples_list = New Collection
    example = Array("Rule: parse features spec only if it's starting with ""Feature:"" keyword or synonyms or has preceding tags")
    examples_list.Add example
    example = Array("Example: feature spec without feature keyword", _
                        "Given a feature ""this is just some text <br> without a feature keyword""", _
                        "When the feature is parsed", _
                        "Then the parsing results in the error message ""Feature lacks feature keyword at the beginning""")
    examples_list.Add example
    
    example = Array("Rule: elements of the features are identified by lines starting with <optional whitespace><keyword>:")
    examples_list.Add example
    example = Array("Example: feature containing one rule and one example", _
                        "Given a feature named ""sample feature""", _
                        "And the feature includes a rule ""this is a sample rule""", _
                        "And the feature includes an example ""this is an example""", _
                        "When the feature is parsed", _
                        "Then the parsed result contains a separate item for each of the given elements")
    examples_list.Add example
    
    example = Array("Rule: every line after a keyword up to the next keyword line or until the first example step in a example clause " _
                        & " is considered to be description except comment lines")
    examples_list.Add example
    ' examples:
    '   feature description includes comment
    '   feature starts with comment
    '   rule has description
    '   example has description and steps
    '   feature has example step as description
    '
    
    
'    example = Array("Rule: translate synonyms ""Ability,Business Needs"" to ""Feature""")
'    examples_list.Add example
'    example = Array("Example: feature starts with Ability:", _
'                        "Given a precondition", _
'                        "And another precondition", _
'                        "When any action happens", _
'                        "Then some result is expected")
'    examples_list.Add example
    
'    example = Array("Rule: ignore trailing white space")
'    examples_list.Add example
'
'    example = Array("Rule: Steps not starting with Given, When, Then or And will produce a Gherkin syntax error")
'    examples_list.Add example
'    example = Array("Example: missing step type", _
'                        "Given an example with a step ""bla bla bla""", _
'                        "When the step is read", _
'                        "Then the step result is an Gherkin Syntax Error claiming ""bla"" is not a valid step type")
'    examples_list.Add example
'
'    example = Array("Rule: Translate ""And"" and ""But"" Given, When or Then using the same keyword as the previous step")
'    examples_list.Add example
'    example = Array("Example: And as synonym for Given", _
'                        "Given an example with two steps ""Given x is 1"" and ""And y is 2""", _
'                        "When the second step is read", _
'                        "Then the step type is set to ""Given""")
'    examples_list.Add example
'    example = Array("Example: But as synonym for When", _
'                        "Given an example with two steps ""When some action happens"" and ""But it doesn't matter""", _
'                        "When the second step is read", _
'                        "Then the step type is set to ""When""")
'    examples_list.Add example
'
'    example = Array("Rule: Translate ""Ability"" and ""Business Need"" to feature")
'    examples_list.Add example
    
    Set Examples = examples_list
End Property

Public Function run_step(pcolStepProps As Collection) As Variant
    
    On Error GoTo error_handler
    
    Select Case pcolStepProps.Item(TExampleRunner.ATTR_STEP_HEAD) & " " & pcolStepProps.Item(TExampleRunner.ATTR_STEP_BODY)
        
        Case "Given a feature ""this is just some text <br> without a feature keyword"""
            m_feature_text = Replace("this is just some text <br> without a feature keyword", " <br> ", vbLf)
        
        Case "Given an example with a step ""bla bla bla"""
            m_example = Array("Example: Gherkin syntax error", _
                                "bla bla bla")
           
        Case "Given an example with two steps ""Given x is 1"" and ""And y is 2"""
            m_example = Array("Example: Using And as Given", _
                                "Given x is 1", _
                                "And y is 2")
        
        Case "Given an example with two steps ""When some action happens"" and ""But it doesn't matter"""
            m_example = Array("Example: Using But as When", _
                                "When some action happens", _
                                "But it doesn't matter")
            
        Case "Given a feature named ""sample feature"""
            m_feature_text = "Feature: sample feature"
        
        Case "Given the feature includes a rule ""this is a sample rule"""
            m_feature_text = m_feature_text & vbLf & "Rule: this is a sample rule"
            
        Case "Given the feature includes an example ""this is an example"""
            m_feature_text = m_feature_text & vbLf & "Example: this is an example"
            
        Case "When the feature is parsed"
            Set m_parsed_feature = TFeatureParse.parse_feature(m_feature_text)
        
        Case "When the step is read"
            Set m_step_attributes = TExampleRunner.read_step(m_example, 1)
        
        Case "When the second step is read"
            Set m_step_attributes = TExampleRunner.read_step(m_example, 2)

        Case "Then the parsing results in the error message ""Feature lacks feature keyword at the beginning"""
            TSpec.expect(m_parsed_feature("error")).to_be "Feature lacks feature keyword at the beginning"
        
        Case "Then the step result is an Gherkin Syntax Error claiming ""bla"" is not a valid step type"
            TSpec.expect(m_step_attributes(TExampleRunner.ATTR_ERROR)).to_be "Gherkin syntax error: >bla< is not a valid step type"
        
        Case "Then the step type is set to ""Given"""
            TSpec.expect(m_step_attributes(TExampleRunner.ATTR_STEP_HEAD)).to_be "Given"
            
        Case "Then the step type is set to ""When"""
            TSpec.expect(m_step_attributes(TExampleRunner.ATTR_STEP_HEAD)).to_be "When"
            
        Case "Then the parsed result contains a separate item for each of the given elements"
            TSpec.expect(m_parsed_feature.Count).to_be 3
            TSpec.expect(m_parsed_feature(1)("type")).to_be TFeature.CLAUSE_TYPE_FEATURE
            TSpec.expect(m_parsed_feature(1)("name")).to_be "sample feature"
            TSpec.expect(m_parsed_feature(2)("type")).to_be TFeature.CLAUSE_TYPE_RULE
            TSpec.expect(m_parsed_feature(2)("name")).to_be "this is a sample rule"
            TSpec.expect(m_parsed_feature(3)("type")).to_be TFeature.CLAUSE_TYPE_EXAMPLE
            TSpec.expect(m_parsed_feature(3)("name")).to_be "this is an example"
            
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = TFeature.fail_step(Err.Number, Err.Description)
End Function

Public Sub before_example()
    
End Sub

Public Sub after_example()
    Set m_step_attributes = Nothing
End Sub

Public Property Get tags() As Variant
    tags = cTags
End Property

Private Sub pending(pending_msg)
    TFeature.pending pending_msg
End Sub

Public Property Let Description(ByVal new_description As String)
    m_description = new_description
End Property

