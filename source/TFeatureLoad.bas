Attribute VB_Name = "TFeatureLoad"
Option Explicit

Public Function get_feature_dir(this_doc_path As String) As String

    Dim path_separator As String

    path_separator = ExtraVBA.get_path_separator
    If Right(this_doc_path, 1) = path_separator Then
        get_feature_dir = this_doc_path & "features"
    Else
        get_feature_dir = this_doc_path & path_separator & "features"
    End If
End Function

Public Function load_features(feature_dir As String) As Collection

    Dim dir_entry As String
    Dim attributes As Integer
    Dim features As Collection
    Dim subdir_features As Collection
    Dim subdir As Variant
    Dim subdirs As Collection
    
    Set features = New Collection
    dir_entry = Dir(feature_dir)
    Do While dir_entry <> vbNullString
        attributes = GetAttr(feature_dir & ExtraVBA.get_path_separator & dir_entry)
        If attributes <> vbHidden And attributes <> vbSystem And Right(dir_entry, 8) = ".feature" Then
            features.Add read_feature(feature_dir & ExtraVBA.get_path_separator & dir_entry)
        End If
        dir_entry = Dir()
    Loop
    Set subdirs = get_subdirs(feature_dir)
    For Each subdir In subdirs
        Set subdir_features = TFeatureLoad.load_features(feature_dir & ExtraVBA.get_path_separator & subdir)
        merge_features features, subdir_features
        Set subdir_features = Nothing
    Next
    Set load_features = features
End Function

Private Function read_feature(feature_file As String) As String
    
    Dim feature As String
    Dim text_line As String
    Dim file_id As Integer
    
    file_id = FreeFile
    Open feature_file For Input As #file_id
    Do Until EOF(1)
        Line Input #1, text_line
        feature = feature & vbLf & text_line
    Loop
    Close #file_id
    read_feature = feature
End Function

Private Sub merge_features(target_features As Collection, source_features As Collection)

    Dim index As Integer
    
    For index = 1 To source_features.Count
        target_features.Add source_features(index)
    Next
End Sub

Private Function get_subdirs(feature_dir As String) As Collection

    Dim subdirs As Collection
    Dim subdir As String
    Dim attributes As Integer

    Set subdirs = New Collection
    subdir = Dir(feature_dir, vbDirectory)
    Do While subdir <> vbNullString
        attributes = GetAttr(feature_dir & ExtraVBA.get_path_separator & subdir)
        If attributes = vbDirectory Then
            If attributes <> vbHidden & attributes <> vbSystem And Left(subdir, 1) <> "." Then
                subdirs.Add subdir
            End If
        End If
        subdir = Dir()
    Loop
    Set get_subdirs = subdirs
End Function