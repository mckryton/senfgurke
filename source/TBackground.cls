VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_background_description As String
Dim m_steps As Collection

Public Property Get steps() As Collection
    If m_steps Is Nothing Then
        Set m_steps = New Collection
    End If
    Set steps = m_steps
End Property

Public Property Set steps(new_step_list As Collection)
    Set m_steps = new_step_list
End Property


Public Property Get description() As String
    description = m_background_description
End Property

Public Property Let description(ByVal background_description As String)
    m_background_description = background_description
End Property


