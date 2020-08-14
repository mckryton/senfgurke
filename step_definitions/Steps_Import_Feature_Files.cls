VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Import_Feature_Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = ""
Dim m_description As String

Public Property Get Description() As String
    
    If m_description = vbNullString Then
        m_description = "Senfgurke will import examples from feature files" & vbLf _
                            & "containing feature name and descripion, rules and examples as text."
    End If
    Description = m_description
End Property

' in the future this property shall be replaced with a function that parses text
Public Property Get Examples() As Collection

    Dim examples_list As Collection
    Dim example As Variant

    Set examples_list = New Collection
    example = Array("Rule: use the features dir beneath the current dir of the document containing the step implementation")
    examples_list.Add example
    example = Array("Example: derive feature from working dir", _
                        "Given the current working dir is ""/Users/cuke/source/senfgurke""", _
                        "When the location for features is set", _
                        "Then the feature dir is set to ""/Users/cuke/source/senfgurke/features""")
    examples_list.Add example
    
    example = Array("Rule: return error message if feature files are not accessible")
    example = Array("Example: feature dir is unavailable", _
                        "Given feature dir is set to ""/this/path/does/not/exist/features""", _
                        "When senfgurke imports features", _
                        "Then the import returns an error message ""can't access feature dir ><""")
    examples_list.Add example
    
    Set Examples = examples_list
End Property

Public Function run_step(pcolStepProps As Collection) As Variant
    
    On Error GoTo error_handler
    
    Select Case pcolStepProps.Item(TExampleRunner.ATTR_STEP_HEAD) & " " & pcolStepProps.Item(TExampleRunner.ATTR_STEP_BODY)
        
        Case "Given a precondition"
           
        Case "Given another precondition"
            
        Case "When any action happens"

        Case "Then some result is expected"
        
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


