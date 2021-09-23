VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This is a container for data exchange between steps in an example. E.g. you may assign a value
' to a variable in one step, add it to the context and retrieve the value in another step.

Option Explicit

Private context_value_collection As Collection

Private Sub Class_Initialize()
    Set context_value_collection = New Collection
End Sub

Private Sub Class_Terminate()
    Set context_value_collection = Nothing
End Sub

Public Sub set_value(new_value As Variant, value_name As String)
    If ExtraVBA.collection_has_key(value_name, context_value_collection) Then
        context_value_collection.Remove value_name
    End If
    context_value_collection.Add new_value, value_name
End Sub

Public Function value_exists(value_name As String) As Boolean
    If ExtraVBA.collection_has_key(value_name, context_value_collection) Then
        value_exists = True
    Else
        value_exists = False
    End If
End Function

Public Function get_value(value_name As String) As Variant
    
    If ExtraVBA.collection_has_key(value_name, context_value_collection) Then
        If IsObject(context_value_collection(value_name)) Then
            Set get_value = context_value_collection(value_name)
        Else
            get_value = context_value_collection(value_name)
        End If
    End If
End Function
