VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TExampleContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is the interface definition for classes holding example objects: TFeature and TRule.
' It helps TFeatureParser to add assign examples to differenttypes of containers (features or rules).

Option Explicit

Public head As String
Public name As String
Public description As String
Public sections As Collection
Public tags As Collection
Public background As TBackground
Public wholeName As String
Public parent As TFeature
