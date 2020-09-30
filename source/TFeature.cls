VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_feature_head As String
Dim m_feature_name As String
Dim m_feature_description As String
Dim m_clauses As Collection
Dim m_error_status As String

Public Property Get Head() As String
    Head = m_feature_head
End Property

Public Property Let Head(ByVal feature_head As String)
    m_feature_head = feature_head
End Property

Public Property Get Name() As String
    Name = m_feature_name
End Property

Public Property Let Name(ByVal feature_name As String)
    m_feature_name = feature_name
End Property

Public Property Get description() As String
    description = m_feature_description
End Property

Public Property Let description(ByVal feature_description As String)
    m_feature_description = feature_description
End Property

Public Property Get Clauses() As Collection
    If m_clauses Is Nothing Then
        Set m_clauses = New Collection
    End If
    Set Clauses = m_clauses
End Property

Public Property Get ErrorStatus() As String
    ErrorStatus = m_error_status
End Property

Public Property Let ErrorStatus(ByVal error_status As String)
    m_error_status = error_status
End Property

Private Sub Class_Initialize()
    m_error_status = vbNullString
End Sub
