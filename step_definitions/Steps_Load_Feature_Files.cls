VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Load_Feature_Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = ""
Dim m_description As String

'module variables used for steps
Dim m_doc_path As String
Dim m_feature_path As String
Dim m_features As Collection

Public Property Get description() As String
    
    If m_description = vbNullString Then
        m_description = "Senfgurke will import examples from feature files" & vbLf _
                            & "containing feature name and description, rules and examples as text."
    End If
    description = m_description
End Property

Public Property Get Examples() As Collection

    Dim examples_list As Collection
    Dim example As Variant

    Set examples_list = New Collection
    example = Array("Rule: search for feature files in the features dir beneath the current dir of the document containing the step implementation")
    examples_list.Add example
    example = Array("Example: derive feature from document dir", _
                        "Given the current application is stored under ""/Users/cuke/source/senfgurke""", _
                        "When the location for features is requested", _
                        "Then the feature dir is set to ""/Users/cuke/source/senfgurke/features""")
    examples_list.Add example
    
    example = Array("Rule: the content of all files with the suffix .feature should be loaded as text, other files should be ignored")
    examples_list.Add example
    example = Array("Example: two feature files on top level", _
                        "Given two feature files in the feature dir", _
                        "When the feature dir is read", _
                        "Then two features are loaded")
    examples_list.Add example
    example = Array("Example: one feature file and one plain text file on top level", _
                        "Given one feature file and one text file in the feature dir", _
                        "When the feature dir is read", _
                        "Then one feature is loaded")
    examples_list.Add example
    example = Array("Example: one feature file on a sub-level", _
                        "Given one feature file located at a sub directory under the feature dir", _
                        "When the feature dir is read", _
                        "Then one feature is loaded")
    examples_list.Add example

' TODO
    example = Array("Rule: return error message if feature files are not accessible")
    examples_list.Add example
'    example = Array("Example: feature dir is unavailable", _
'                        "Given feature dir is set to ""/this/path/does/not/exist/features""", _
'                        "When senfgurke reads features from feature dir", _
'                        "Then the import returns an error message ""can't access feature dir ><""")
'    examples_list.Add example
    
    Set Examples = examples_list
End Property

Public Function run_step(step_definition As String) As Variant
    
    On Error GoTo error_handler
    
    Select Case step_definition
        
        Case "Given the current application is stored under ""/Users/cuke/source/senfgurke"""
            m_doc_path = Join(Split("/Users/cuke/source/senfgurke", "/"), ExtraVBA.get_path_separator)
            
        Case "Given feature dir is set to ""/this/path/does/not/exist/features"""
            m_doc_path = Join(Split("/this/path/does/not/exist/features", "/"), ExtraVBA.get_path_separator)
        
        Case "Given two feature files in the feature dir"
            m_feature_path = ThisWorkbook.Path & ExtraVBA.get_path_separator & "testdata/features_toplevel_pure"
            'testfiles are already part of this repo, so that no write access is required for running tests
            TSpec.expect(Dir(m_feature_path)).to_be "dummy_1.feature"
            TSpec.expect(Dir()).to_be "dummy_2.feature"
            TSpec.expect(Dir()).to_be vbNullString
            
        Case "Given one feature file and one text file in the feature dir"
            m_feature_path = ThisWorkbook.Path & ExtraVBA.get_path_separator & "testdata/features_toplevel_mixed"
            'testfiles are already part of this repo, so that no write access is required for running tests
            TSpec.expect(Dir(m_feature_path)).to_be "dummy_1.txt"
            TSpec.expect(Dir()).to_be "dummy_2.feature"
            TSpec.expect(Dir()).to_be vbNullString
            
        Case "Given one feature file located at a sub directory under the feature dir"
            m_feature_path = ThisWorkbook.Path & ExtraVBA.get_path_separator & "testdata/features_subdir"
            'testfiles are already part of this repo, so that no write access is required for running tests
            TSpec.expect(Dir(m_feature_path)).to_be vbNullString
            TSpec.expect(Dir(m_feature_path & ExtraVBA.get_path_separator & "subdir")).to_be "dummy_1.feature"
            TSpec.expect(Dir()).to_be vbNullString
        
        Case "When the location for features is requested"
            m_feature_path = TFeatureLoader.get_feature_dir(m_doc_path)
            
        Case "When the feature dir is read"
            Set m_features = TFeatureLoader.load_features(m_feature_path)
            
        Case "Then the feature dir is set to ""/Users/cuke/source/senfgurke/features"""
            TSpec.expect(m_feature_path).to_be Join(Split("/Users/cuke/source/senfgurke/features", "/"), ExtraVBA.get_path_separator)
            
        Case "Then two features are loaded"
            TSpec.expect(m_features.Count).to_be 2
            
        Case "Then one feature is loaded"
            TSpec.expect(m_features.Count).to_be 1
            
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = TStepRunner.fail_step(Err.Number, Err.description)
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

Public Property Let description(ByVal new_description As String)
    m_description = new_description
End Property


