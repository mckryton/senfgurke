VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_example_head As String
Dim m_example_name As String
Dim m_example_description As String
Dim m_steps As Collection
Dim m_example_tags As Collection
Dim m_original_headline As String

Public Property Get Head() As String
    Head = m_example_head
End Property

Public Property Let Head(ByVal example_head As String)
    m_example_head = example_head
End Property

Public Property Get Name() As String
    Name = m_example_name
End Property

Public Property Let Name(ByVal example_name As String)
    m_example_name = example_name
End Property

Public Property Get description() As String
    description = m_example_description
End Property

Public Property Let description(ByVal example_description As String)
    m_example_description = example_description
End Property

Public Property Get OriginalHeadline() As String
    OriginalHeadline = m_original_headline
End Property

Public Property Let OriginalHeadline(original_headline As String)
    m_original_headline = original_headline
End Property

Public Property Get Steps() As Collection
    If m_steps Is Nothing Then
        Set m_steps = New Collection
    End If
    Set Steps = m_steps
End Property

Public Property Get Tags() As Collection
    If m_example_tags Is Nothing Then
        Set m_example_tags = New Collection
    End If
    Set Tags = m_example_tags
End Property

Public Property Let Tags(ByVal example_tags As Collection)

    Set m_example_tags = example_tags
End Property


Public Sub add_step(step_head As String, step_name As String)

    Dim new_step As TStep
    Dim previous_step_type As String

    If Me.Steps.Count = 0 Then
        previous_step_type = STEP_TYPE_GIVEN
    Else
        previous_step_type = Me.Steps(Me.Steps.Count).SType
    End If
    Set new_step = New TStep
    If step_head = "And" Or step_head = "But" Then
        new_step.SType = previous_step_type
    Else
        new_step.SType = get_step_type(step_head)
    End If
    new_step.Head = step_head
    new_step.Name = step_name
    Me.Steps.Add new_step
End Sub

Private Function get_step_type(step_head As String) As String

    Select Case step_head
        Case "Given"
            get_step_type = STEP_TYPE_GIVEN
        Case "When"
            get_step_type = STEP_TYPE_WHEN
        Case "Then"
            get_step_type = STEP_TYPE_THEN
        Case Else
            get_step_type = "unknown step type"
    End Select
End Function
