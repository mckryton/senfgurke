Attribute VB_Name = "TFeatureLoader"
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

Public Function load_features(Optional feature_dir, Optional file_filter, Optional session) As Collection
    Dim dir_entry As String
    Dim attributes As Integer
    Dim features As Collection
    Dim subdir_features As Collection
    Dim subdir As Variant
    Dim subdirs As Collection
    
    If IsMissing(feature_dir) Then
        #If app_name = "Microsoft PowerPoint" Then
            TError.raise ERR_ID_UNKNOWN_FEATURE_PATH
        #ElseIf app_name = "Microsoft Excel" Then
            feature_dir = get_feature_dir(senfgurke_workbook.Path)
        #End If
    End If
    If Right(feature_dir, 1) <> ExtraVBA.get_path_separator Then
        feature_dir = feature_dir & ExtraVBA.get_path_separator
    End If
    Set features = New Collection
    dir_entry = Dir(feature_dir)
    Do While dir_entry <> vbNullString
        attributes = GetAttr(feature_dir & dir_entry)
        If attributes <> vbHidden And attributes <> vbSystem And Right(dir_entry, 8) = ".feature" Then
            If Not IsMissing(file_filter) Then
                If Left(dir_entry, Len(file_filter)) = CStr(file_filter) Then
                    features.Add read_feature(CStr(feature_dir), dir_entry, session)
                End If
            Else
                features.Add read_feature(CStr(feature_dir), dir_entry)
            End If
        End If
        dir_entry = Dir()
    Loop
    Set subdirs = get_subdirs(CStr(feature_dir))
    For Each subdir In subdirs
        Set subdir_features = TFeatureLoader.load_features(feature_dir & subdir, file_filter, session)
        merge_features features, subdir_features
        Set subdir_features = Nothing
    Next
    Set load_features = features
End Function

Private Function read_feature(feature_dir As String, feature_file As String, Optional session) As Collection
    Dim feature_data As Collection
    Dim feature As String
    Dim text_line As String
    Dim file_id As Integer
    
    Set feature_data = New Collection
    file_id = FreeFile
    Open feature_dir & feature_file For Input As #file_id
    Do Until EOF(1)
        Line Input #1, text_line
        feature = feature & text_line & vbLf
    Loop
    Close #file_id
    feature_data.Add feature, "feature_text"
    feature_data.Add feature_file, "origin"
    If Not IsMissing(session) Then session.log_event EVENT_FEATURE_LOADED, feature_file
    Set read_feature = feature_data
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
        attributes = GetAttr(feature_dir & subdir)
        If attributes = vbDirectory Then
            If attributes <> vbHidden & attributes <> vbSystem And Left(subdir, 1) <> "." Then
                subdirs.Add subdir
            End If
        End If
        subdir = Dir()
    Loop
    Set get_subdirs = subdirs
End Function

