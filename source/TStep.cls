VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_step_type As String
Dim m_step_head As String
Dim m_step_name As String

Public Property Get SType() As String
    SType = m_step_type
End Property

Public Property Let SType(ByVal step_type As String)
    m_step_type = step_type
End Property

Public Property Get Head() As String
    Head = m_step_head
End Property

Public Property Let Head(ByVal step_head As String)
    m_step_head = step_head
End Property

Public Property Get Name() As String
    Name = m_step_name
End Property

Public Property Let Name(ByVal step_name As String)
    m_step_name = step_name
End Property

Public Property Get WholeStep() As String
    WholeStep = Me.Head & " " & Me.Name
End Property

Public Property Get StepImplementation() As String
    StepImplementation = Me.SType & " " & Me.Name
End Property
