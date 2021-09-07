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

Implements TExampleContainer

Const CONTAINER_TYPE_FEATURE = "feature"

Private m_feature_head As String
Private m_feature_name As String
Private m_feature_description As String
Private m_sections As Collection
Private m_feature_tags As Collection
Private m_background As TBackground
Private m_container_type As String
Private m_parsed_lines As Long

Private Sub Class_Initialize()
    m_parsed_lines = 0
End Sub

Private Sub Class_Terminate()
    Set m_sections = Nothing
    Set m_feature_tags = Nothing
    Set m_background = Nothing
End Sub

Public Property Get head() As String
    head = TExampleContainer_head
End Property

Public Property Get parsed_lines() As Long
    parsed_lines = m_parsed_lines
End Property

Public Property Let parsed_lines(ByVal line_index As Long)
    m_parsed_lines = line_index
End Property

Public Property Let head(ByVal new_head As String)
    TExampleContainer_head = new_head
End Property

Public Property Get name() As String
    name = TExampleContainer_name
End Property

Public Property Let name(ByVal new_name As String)
    TExampleContainer_name = new_name
End Property

Public Property Get description() As String
    description = TExampleContainer_description
End Property

Public Property Let description(ByVal new_description As String)
    TExampleContainer_description = new_description
End Property

Public Property Get sections() As Collection
    Set sections = TExampleContainer_sections
End Property

Public Property Get tags() As Collection
    Set tags = TExampleContainer_tags
End Property

Public Property Set tags(ByVal new_tags As Collection)
    Set TExampleContainer_tags = new_tags
End Property

Public Property Get background() As TBackground
    Set background = TExampleContainer_background
End Property

Public Property Set background(new_background As TBackground)
    Set TExampleContainer_background = new_background
End Property

Public Property Get wholeName() As String
    wholeName = TExampleContainer_wholeName
End Property

Public Property Get containerType() As String
    containerType = TExampleContainer_containerType
End Property

'--- start of the implementation of the TExampleContainer interface ---
' see https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/implements-statement
' for more information about interfaces in VBA

Private Property Get TExampleContainer_head() As String
    TExampleContainer_head = m_feature_head
End Property

Private Property Let TExampleContainer_head(ByVal feature_head As String)
    m_feature_head = feature_head
End Property

Private Property Get TExampleContainer_name() As String
    TExampleContainer_name = m_feature_name
End Property

Private Property Let TExampleContainer_name(ByVal feature_name As String)
    m_feature_name = feature_name
End Property

Private Property Get TExampleContainer_description() As String
    TExampleContainer_description = m_feature_description
End Property

Private Property Let TExampleContainer_description(ByVal feature_description As String)
    m_feature_description = feature_description
End Property

Private Property Get TExampleContainer_parent() As TFeature
    ' a feature doesn't have a parent feature, only rules can have a parent feature
    Set TExampleContainer_parent = Nothing
End Property

Private Property Set TExampleContainer_parent(ByVal new_parent As TFeature)
    'VBA interface requires full property declaration even for read-only properties
    Debug.Print "DEBUG: can't assign new value > TFeature.parent is read only"
End Property

Private Property Get TExampleContainer_sections() As Collection
    If m_sections Is Nothing Then
        Set m_sections = New Collection
    End If
    Set TExampleContainer_sections = m_sections
End Property

Private Property Set TExampleContainer_sections(ByVal new_sections As Collection)
    'VBA interface requires full property declaration even for read-only properties
    Debug.Print "DEBUG: can't assign new value > TFeature.sections is read only"
End Property

Private Property Get TExampleContainer_tags() As Collection
    If m_feature_tags Is Nothing Then
        Set m_feature_tags = New Collection
    End If
    Set TExampleContainer_tags = m_feature_tags
End Property

Private Property Set TExampleContainer_tags(ByVal feature_tags As Collection)
    Set m_feature_tags = feature_tags
End Property

Private Property Get TExampleContainer_background() As TBackground
    If m_background Is Nothing Then
        Set m_background = New TBackground
    End If
    Set TExampleContainer_background = m_background
End Property

Private Property Set TExampleContainer_background(ByVal new_background As TBackground)
    Set m_background = new_background
End Property

Private Property Get TExampleContainer_wholeName() As String
    TExampleContainer_wholeName = TExampleContainer_head & ": " & TExampleContainer_name
End Property

Private Property Let TExampleContainer_wholeName(ByVal new_name As String)
    'VBA interface requires full property declaration even for read-only properties
    Debug.Print "DEBUG: can't assign new value > Tfeature.wholeName is read only"
End Property

Private Property Get TExampleContainer_containerType() As String
    TExampleContainer_containerType = CONTAINER_TYPE_FEATURE
End Property

Private Property Let TExampleContainer_containerType(ByVal new_container_type As String)
    'VBA interface requires full property declaration even for read-only properties
    Debug.Print "DEBUG: can't assign new value > TFeature.containerType is read only"
End Property
