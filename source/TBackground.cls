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
Dim m_docstring As String

Public Property Get Steps() As Collection
    If m_steps Is Nothing Then
        Set m_steps = New Collection
    End If
    Set Steps = m_steps
End Property

Public Property Get Description() As String
    Description = m_background_description
End Property

Public Property Let Description(ByVal background_description As String)
    m_background_description = background_description
End Property

Public Property Get Docstring() As String
    Docstring = m_docstring
End Property

Public Property Let Docstring(ByVal new_docstring As String)
    m_docstring = new_docstring
End Property

