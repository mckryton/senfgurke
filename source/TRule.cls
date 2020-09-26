VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_rule_name As String
Dim m_rule_description As String

Public Property Get Name() As String
    Name = m_rule_name
End Property

Public Property Let Name(ByVal rule_name As String)
    m_rule_name = rule_name
End Property

Public Property Get Description() As String
    Description = m_rule_description
End Property

Public Property Let Description(ByVal rule_description As String)
    m_rule_description = rule_description
End Property

Public Property Get WholeRule() As String
    WholeRule = "Rule: " & Me.Name
End Property
