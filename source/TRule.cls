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

Implements TExampleContainer

Const CONTAINER_TYPE_RULE = "rule"

Private m_rule_name As String
Private m_rule_description As String
Private m_rule_head As String
Private m_rule_tags As Collection
Private m_rule_sections As Collection
Private m_rule_background As TBackground
Private m_parsed_lines_index As Long
Private m_feature_parent As TFeature

Private Sub Class_Initialize()
    m_rule_head = "Rule"
End Sub

Private Sub Class_Terminate()
    Set m_rule_tags = Nothing
    Set m_rule_sections = Nothing
    Set m_rule_background = Nothing
    Set m_feature_parent = Nothing
End Sub

Public Property Get head() As String
    head = TExampleContainer_head
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

Public Property Get parent() As TFeature
    Set parent = TExampleContainer_parent
End Property

Public Property Set parent(ByVal new_parent As Variant)
    Set TExampleContainer_parent = new_parent
End Property

Public Property Get containerType() As String
    containerType = TExampleContainer_containerType
End Property

'--- start of the implementation of the TExampleContainer interface ---
' see https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/implements-statement
' for more information about interfaces in VBA

Private Property Get TExampleContainer_background() As TBackground
    If m_rule_background Is Nothing Then
        Set m_rule_background = New TBackground
    End If
    Set TExampleContainer_background = m_rule_background
End Property

Private Property Set TExampleContainer_background(ByVal new_rule_background As TBackground)
    Set m_rule_background = new_rule_background
End Property

Private Property Get TExampleContainer_head() As String
    TExampleContainer_head = m_rule_head
End Property

Private Property Let TExampleContainer_head(ByVal new_rule_head As String)
    m_rule_head = new_rule_head
End Property

Private Property Get TExampleContainer_name() As String
    TExampleContainer_name = m_rule_name
End Property

Private Property Let TExampleContainer_name(ByVal new_rule_name As String)
    m_rule_name = new_rule_name
End Property

Private Property Get TExampleContainer_description() As String
    TExampleContainer_description = m_rule_description
End Property

Private Property Let TExampleContainer_description(ByVal new_rule_description As String)
    m_rule_description = new_rule_description
End Property

Private Property Get TExampleContainer_tags() As Collection
    If m_rule_tags Is Nothing Then
        Set m_rule_tags = New Collection
    End If
    Set TExampleContainer_tags = m_rule_tags
End Property

Private Property Set TExampleContainer_tags(ByVal new_tags As Collection)
    Set m_rule_tags = new_tags
End Property

Private Property Get TExampleContainer_sections() As Collection
    If m_rule_sections Is Nothing Then
        Set m_rule_sections = New Collection
    End If
    Set TExampleContainer_sections = m_rule_sections
End Property

Private Property Set TExampleContainer_sections(ByVal new_sections As Collection)
    'VBA interface requires full property declaration even for read-only properties
    Debug.Print "DEBUG: can't assign new value > TRule.sections is read only"
End Property

Private Property Get TExampleContainer_wholeName() As String
    TExampleContainer_wholeName = TExampleContainer_head & ": " & TExampleContainer_name
End Property

Private Property Let TExampleContainer_wholeName(ByVal new_name As String)
    'VBA interface requires full property declaration even for read-only properties
    Debug.Print "DEBUG: can't assign new value > TRule.wholeName is read only"
End Property

Private Property Get TExampleContainer_parent() As TFeature
    Set TExampleContainer_parent = m_feature_parent
End Property

Private Property Set TExampleContainer_parent(ByVal new_parent As TFeature)
    Set m_feature_parent = new_parent
End Property

Private Property Get TExampleContainer_containerType() As String
    TExampleContainer_containerType = CONTAINER_TYPE_RULE
End Property

Private Property Let TExampleContainer_containerType(ByVal new_container_type As String)
    'VBA interface requires full property declaration even for read-only properties
    Debug.Print "DEBUG: can't assign new value > TFeature.containerType is read only"
End Property

