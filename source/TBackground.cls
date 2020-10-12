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

Dim m_parent_feature As TFeature

Public Property Get ParentFeature() As TFeature

    Set ParentFeature = m_parent_feature
End Property

Public Property Let ParentFeature(ByVal parent_feature As TFeature)
    Set m_parent_feature = parent_feature
End Property

Public Sub add_step(step_head As String, step_name As String)

    Dim new_step As TStep
    
    Set new_step = New TStep
    new_step.SType = STEP_TYPE_GIVEN
    new_step.Head = step_head
    new_step.Name = step_name
    Me.ParentFeature.BackgroundSteps.Add new_step
End Sub
