VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Run_Tests"
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
        m_description = "The main purpose of senfgurke is to run tests containing" & vbLf _
                            & "numerous examples used to describe distinct features."
    End If
    Description = m_description
End Property

' in the future this property shall be replaced with a function that parses text
Public Property Get Examples() As Collection

    Dim examples_list As Collection
    Dim example As Variant

    Set examples_list = New Collection
    example = Array("Rule: stop execution when the import of feature files fails")
    examples_list.Add example
'    example = Array("Example: import results with error message", _
'                        "Given ", _
'                        "When  ", _
'                        "Then  ")
'    examples_list.Add example

    Set Examples = examples_list
End Property

Public Function run_step(step_definition As String) As Variant
    
    On Error GoTo error_handler
    
    Select Case step_definition
        
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

Public Property Let Description(ByVal new_description As String)
    m_description = new_description
End Property




