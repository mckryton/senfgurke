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
Dim m_sections As Collection
Dim m_feature_tags As Collection
Dim m_parsed_lines_index As Long
Dim m_background As TBackground

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

Public Property Get Description() As String
    Description = m_feature_description
End Property

Public Property Let Description(ByVal feature_description As String)
    m_feature_description = feature_description
End Property

Public Property Get sections() As Collection
    If m_sections Is Nothing Then
        Set m_sections = New Collection
    End If
    Set sections = m_sections
End Property

Public Property Get tags() As Collection
    If m_feature_tags Is Nothing Then
        Set m_feature_tags = New Collection
    End If
    Set tags = m_feature_tags
End Property

Public Property Let tags(ByVal feature_tags As Collection)
    Set m_feature_tags = feature_tags
End Property

Public Property Get ParsedLinesIndex() As Long
    ParsedLinesIndex = m_parsed_lines_index
End Property

Public Property Let ParsedLinesIndex(ByVal parsed_lines_index As Long)
    m_parsed_lines_index = parsed_lines_index
End Property

Public Property Get Background() As TBackground
    If m_background Is Nothing Then
        Set m_background = New TBackground
    End If
    Set Background = m_background
End Property

Public Property Let Background(new_background As TBackground)
    Set m_background = new_background
End Property

Public Sub add_tags(feature_line As String)

    Dim tag_list As Variant
    Dim tag As Variant
    
    tag_list = Split(Trim(feature_line), " ")
    For Each tag In tag_list
        If Len(tag) > 1 Then
            If Left(tag, 1) = "@" Then
                If Not ExtraVBA.collection_has_key(tag, Me.tags) Then
                    Me.tags.Add tag, tag
                End If
            End If
        End If
    Next
End Sub

Private Sub Class_Initialize()
    m_parsed_lines_index = 0
End Sub

Private Sub Class_Terminate()
    Set m_sections = Nothing
    Set m_feature_tags = Nothing
    Set m_background = Nothing
End Sub
