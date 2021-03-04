VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TStepExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_value As Variant
Dim m_type_name As String
Dim m_index_start As Long
Dim m_index_end As Long

Public Property Get value() As Variant
    value = m_value
End Property

Public Property Let value(ByVal new_value As Variant)
    m_value = new_value
End Property

Public Property Get TypeName() As String
    TypeName = m_type_name
End Property

Public Property Let TypeName(ByVal type_name As String)
    m_type_name = type_name
End Property

Public Property Get IndexStart() As Long
    IndexStart = m_index_start
End Property

Public Property Let IndexStart(ByVal index_start As Long)
    m_index_start = index_start
End Property

Public Property Get IndexEnd() As Long
    IndexEnd = m_index_end
End Property

Public Property Let IndexEnd(ByVal index_end As Long)
    m_index_end = index_end
End Property
