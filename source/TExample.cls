VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_example_head As String
Dim m_example_name As String
Dim m_example_description As String
Dim m_steps As Collection
Dim m_example_tags As Collection
Dim m_original_headline As String
Dim m_outlines As Collection

Private Sub Class_Terminate()
    Set m_outlines = Nothing
End Sub

Public Property Get head() As String
    head = m_example_head
End Property

Public Property Let head(ByVal example_head As String)
    m_example_head = example_head
End Property

Public Property Get name() As String
    name = m_example_name
End Property

Public Property Let name(ByVal example_name As String)
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

Public Property Get Outlines() As Collection
    If m_outlines Is Nothing Then Set m_outlines = New Collection
    Set Outlines = m_outlines
End Property

Public Property Get steps() As Collection
    If m_steps Is Nothing Then
        Set m_steps = New Collection
    End If
    Set steps = m_steps
End Property

Public Property Set steps(new_step_list As Collection)
    Set m_steps = new_step_list
End Property

Public Property Get tags() As Collection
    If m_example_tags Is Nothing Then
        Set m_example_tags = New Collection
    End If
    Set tags = m_example_tags
End Property

Public Property Let tags(ByVal example_tags As Collection)
    Set m_example_tags = example_tags
End Property

Public Sub insert_background_steps(background_steps As Collection)
    Dim background_step As TStep
    Dim step_index As Long
    
    For step_index = background_steps.Count To 1
        Set background_step = background_steps(step_index)
        If Me.steps.Count > 0 Then
            Me.steps.Add background_step, before:=1
        Else
            Me.steps.Add background_step
        End If
    Next
End Sub
