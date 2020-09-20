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
Const cTags = ""
Dim m_description As String

'module variables used for steps
Dim m_feature_text As String
Dim m_example As Variant
Dim m_step_attributes As Collection
Dim m_parsed_feature As TFeature

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
    
    example = Array("Rule: feature clauses ar limited by lines matching this format <optional whitespace><keyword>:<optional whitespace><name><optional whitespace>" _
                            & "any optional whitspace is ignored")
    examples_list.Add example
    example = Array("Example: feature containing one rule and one example", _
                        "Given a feature named ""sample feature""", _
                        "And the feature includes a rule ""this is a sample rule""", _
                        "And the feature includes an example ""this is an example""", _
                        "When the feature is parsed", _
                        "Then the parsed result contains a separate item for each of the given elements")
    examples_list.Add example
    example = Array("Example: keywords with whitespace", _
                        "Given a feature starting with ""  Feature: sample feature """, _
                        "And the feature continues with the line ""    Rule: this is a rule""", _
                        "When the feature is parsed", _
                        "Then the parsed result contains a feature with the name ""sample feature""", _
                        "And the  parsed result contains a rule with the name ""this is a rule""")
    examples_list.Add example

    ' any trailing white space is removed from description
    example = Array("Rule: every line after a keyword up to the next keyword line or until the first example step in a example clause " _
                        & " is considered to be description except comment lines") 'PENDING description is not allowed after example steps
    examples_list.Add example
    example = Array("Example: feature with 2 lines of description", _
                        "Given a feature named ""sample feature""", _
                        "And the line with the feature keyword is followed by two lines ""  this is<br>  the description""", _
                        "When the feature is parsed", _
                        "Then the feature description is set to those two lines")
    examples_list.Add example
    example = Array("Example: rule with description", _
                        "Given a feature named ""sample feature""", _
                        "And the feature includes a rule ""this is a sample rule""", _
                        "And the rule is followed by a line ""  this is a description""", _
                        "When the feature is parsed", _
                        "Then the parsed feature contains a rule", _
                        "And the rules description is set to ""this is a description""")
    examples_list.Add example
'    ' TODO add examples:
'    '   feature description includes comment
'    '   feature starts with comment
'    '   example has description and steps
'    '   feature has example step as description

    example = Array("Rule: every line in a example clause starting with ""<optional whitspace><step keyword><space>"" is considered to be a step" _
                        & " where <tep keyword> is one of those: Given, When, Then, And, But")
    examples_list.Add example
    example = Array("Example: simple example with Given, When, Then steps", _
                        "Given a feature", _
                        "And the feature includes a line ""  Example: sample example""", _
                        "And this line is followed by ""    Given a precondition""", _
                        "And this line is followed by ""    When action happens""", _
                        "And this line is followed by ""    Then some result is expected""", _
                        "When the feature is parsed", _
                        "Then the parsed result contains an example", _
                        "And the example clause from the parsed result contains all the steps")
    examples_list.Add example

''    example = Array("Rule: translate synonyms ""Ability,Business Needs"" to ""Feature""")
''    examples_list.Add example
''    example = Array("Example: feature starts with Ability:", _
''                        "Given a precondition", _
''                        "And another precondition", _
''                        "When any action happens", _
''                        "Then some result is expected")
''    examples_list.Add example
'
''    example = Array("Rule: Steps not starting with Given, When, Then or And will produce a Gherkin syntax error")
''    examples_list.Add example
''    example = Array("Example: missing step type", _
''                        "Given an example with a step ""bla bla bla""", _
''                        "When the step is read", _
''                        "Then the step result is an Gherkin Syntax Error claiming ""bla"" is not a valid step type")
''    examples_list.Add example
''
''    example = Array("Rule: Translate ""And"" and ""But"" Given, When or Then using the same keyword as the previous step")
''    examples_list.Add example
''    example = Array("Example: And as synonym for Given", _
''                        "Given an example with two steps ""Given x is 1"" and ""And y is 2""", _
''                        "When the second step is read", _
''                        "Then the step type is set to ""Given""")
''    examples_list.Add example
''    example = Array("Example: But as synonym for When", _
''                        "Given an example with two steps ""When some action happens"" and ""But it doesn't matter""", _
''                        "When the second step is read", _
''                        "Then the step type is set to ""When""")
''    examples_list.Add example

'TODO: warn when steps are longer than 255 chars
    
    Set Examples = examples_list
End Property

Public Function run_step(step_definition As Collection) As Variant
    
    Dim example_steps As Collection
    
    On Error GoTo error_handler
    Select Case step_definition.Item(TExampleRunner_Old.ATTR_STEP_HEAD) & " " & step_definition.Item(TExampleRunner_Old.ATTR_STEP_BODY)
        
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
            
        Case "Given the line with the feature keyword is followed by two lines ""  this is<br>  the description"""
            m_feature_text = m_feature_text & vbLf & "  this is" & vbLf & "  the description"
        
        Case "Given the feature includes a rule ""this is a sample rule"""
            m_feature_text = m_feature_text & vbLf & "Rule: this is a sample rule"
            
        Case "Given the rule is followed by a line ""  this is a description"""
            m_feature_text = m_feature_text & vbLf & "  this is a description"
            
        Case "Given a feature starting with ""  Feature: sample feature """
            m_feature_text = "  Feature: sample feature "
            
        Case "Given the feature continues with the line ""    Rule: this is a rule"""
            m_feature_text = m_feature_text & vbLf & "    Rule: this is a rule"
        
        Case "Given a feature"
            m_feature_text = "Feature: sample feature"
            
        Case "Given the feature includes a line ""  Example: sample example"""
            m_feature_text = m_feature_text & vbLf & "  Example: sample example"
        
        Case "Given this line is followed by ""    Given a precondition"""
            m_feature_text = m_feature_text & vbLf & "    Given a precondition"
            
        Case "Given this line is followed by ""    When action happens"""
            m_feature_text = m_feature_text & vbLf & "    When action happens"
            
        Case "Given this line is followed by ""    Then some result is expected"""
            m_feature_text = m_feature_text & vbLf & "    Then some result is expected"
                                
        Case "When the feature is parsed"
            Set m_parsed_feature = TFeatureParser.parse_feature(m_feature_text)
        
        Case "When the step is read"
            Set m_step_attributes = TExampleRunner_Old.read_step(m_example, 1)
        
        Case "When the second step is read"
            Set m_step_attributes = TExampleRunner_Old.read_step(m_example, 2)

        Case "Then the parsing results in the error message ""Feature lacks feature keyword at the beginning"""
            TSpec.expect(m_parsed_feature.ErrorStatus).to_be "Feature lacks feature keyword at the beginning"
        
        Case "Then the step result is an Gherkin Syntax Error claiming ""bla"" is not a valid step type"
            TSpec.expect(m_step_attributes(TExampleRunner_Old.ATTR_ERROR)).to_be "Gherkin syntax error: >bla< is not a valid step type"
        
        Case "Then the step type is set to ""Given"""
            TSpec.expect(m_step_attributes(TExampleRunner_Old.ATTR_STEP_HEAD)).to_be "Given"
            
        Case "Then the step type is set to ""When"""
            TSpec.expect(m_step_attributes(TExampleRunner_Old.ATTR_STEP_HEAD)).to_be "When"
            
        Case "Then the parsed result contains a separate item for each of the given elements"
            TSpec.expect(m_parsed_feature.Name).to_be "sample feature"
            TSpec.expect(m_parsed_feature.Clauses.Count).to_be 2
            TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TRule"
            TSpec.expect(m_parsed_feature.Clauses(1).Name).to_be "this is a sample rule"
            TSpec.expect(TypeName(m_parsed_feature.Clauses(2))).to_be "TExample"
            TSpec.expect(m_parsed_feature.Clauses(2).Name).to_be "this is an example"
            
        Case "Then the feature description is set to those two lines"
            TSpec.expect(m_parsed_feature.Description).to_be "this is" & vbLf & "the description"
            
        Case "Then the parsed feature contains a rule"
            TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TRule"
            
        Case "Then the rules description is set to ""this is a description"""
            TSpec.expect(m_parsed_feature.Clauses(1).Description).to_be "this is a description"
            
        Case "Then the parsed result contains a feature with the name ""sample feature"""
            TSpec.expect(m_parsed_feature.Name).to_be "sample feature"
        
        Case "Then the  parsed result contains a rule with the name ""this is a rule"""
            TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TRule"
            TSpec.expect(m_parsed_feature.Clauses(1).Name).to_be "this is a rule"
        
        Case "Then the parsed result contains an example"
            TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TExample"
        
        Case "Then the example clause from the parsed result contains all the steps"
            Set example_steps = m_parsed_feature.Clauses(1).Steps
            TSpec.expect(example_steps.Count).to_be 3
            TSpec.expect(example_steps(1).SType).to_be STEP_TYPE_GIVEN
            TSpec.expect(example_steps(1).Name).to_be "a precondition"
            TSpec.expect(example_steps(1).Head).to_be "Given"
            TSpec.expect(example_steps(2).SType).to_be STEP_TYPE_WHEN
            TSpec.expect(example_steps(2).Name).to_be "action happens"
            TSpec.expect(example_steps(2).Head).to_be "When"
            TSpec.expect(example_steps(3).SType).to_be STEP_TYPE_THEN
            TSpec.expect(example_steps(3).Name).to_be "some result is expected"
            TSpec.expect(example_steps(3).Head).to_be "Then"
            
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
            
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = TStepRunner.fail_step(Err.Number, Err.Description)
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
    TFeatureRunner.pending pending_msg
End Sub

Public Property Let Description(ByVal new_description As String)
    m_description = new_description
End Property

