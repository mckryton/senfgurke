Attribute VB_Name = "TSupport"
'this module is for sharing helper functions between your step implementation classes

Option Explicit

'mocking loading features from disk
Public Sub load_feature_from_text(feature_text As String, loaded_features As Collection, Optional feature_origin As String)
    
    Dim loaded_feature As Collection
    
    If IsMissing(feature_origin) Then feature_origin = vbNullString
    Set loaded_feature = New Collection
    loaded_feature.Add feature_text, "feature_text"
    loaded_feature.Add feature_origin, "origin"
    loaded_features.Add loaded_feature
End Sub
