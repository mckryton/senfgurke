VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Load_Feature_Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_the_current_application_is_stored_under_STR_E4B6AFDB477F(example_context As TContext, step_expressions As Collection)
    'Given the current application is stored under {string}
    example_context.set_value Join(Split(step_expressions(1), "/"), ExtraVBA.get_path_separator), "doc_path"
End Sub

Public Sub Given_two_feature_files_in_the_feature_dir_8FB24F121954(example_context As TContext)
    'Given two feature files in the feature dir
    example_context.set_value ThisWorkbook.path & ExtraVBA.get_path_separator & Join(Split("testdata/features_toplevel_pure", "/"), ExtraVBA.get_path_separator), "feature_path"
End Sub

Public Sub Given_one_feature_file_and_one_text_file_in_the_fe_98F5E94C9F0E(example_context As TContext)
    'Given one feature file and one text file in the feature dir
    example_context.set_value ThisWorkbook.path & ExtraVBA.get_path_separator & Join(Split("testdata/features_toplevel_mixed", "/"), ExtraVBA.get_path_separator), "feature_path"
End Sub

Public Sub Given_one_feature_file_located_at_a_sub_directory__CA04B38A1CAE(example_context As TContext)
    'Given one feature file located at a sub directory under the feature dir
    example_context.set_value ThisWorkbook.path & ExtraVBA.get_path_separator & Join(Split("testdata/features_subdir", "/"), ExtraVBA.get_path_separator), "feature_path"
End Sub

Public Sub Given_a_feature_file_STR_D7FE0A9A5327(example_context As TContext, step_expressions As Collection)
    'Given a feature file {string}
    Dim dir_entry As String
    Dim file_names As String
    Dim application_dir As String
    Dim feature_dir As String
    
    application_dir = senfgurke_steps_workbook.path & ExtraVBA.get_path_separator _
        & "testdata" & ExtraVBA.get_path_separator _
        & "filter_for_loading_features" '& ExtraVBA.get_path_separator
    example_context.set_value application_dir, "application_dir"
    feature_dir = application_dir & ExtraVBA.get_path_separator & "features" & ExtraVBA.get_path_separator
    example_context.set_value feature_dir, "feature_dir"
    file_names = read_dir_entries(feature_dir)
    'creating files for each test run might be slow and difficult, therefore just ensure that test files are part of the repository
    TSpec.expect(file_names).includes_text (step_expressions(1))
End Sub

Public Sub Given_a_feature_directory_contains_the_files_51DDA68DE9B3(example_context As TContext, data_table As TDataTable)
    'Given a feature directory contains the files
    Dim dir_entry As String
    Dim file_names As String
    Dim application_dir As String
    Dim feature_dir As String
    Dim table_row As Variant
    Dim file_name As String
    
    application_dir = senfgurke_steps_workbook.path & ExtraVBA.get_path_separator _
        & "testdata" & ExtraVBA.get_path_separator _
        & "filter_for_loading_features" & ExtraVBA.get_path_separator
    example_context.set_value application_dir, "application_dir"
    feature_dir = application_dir & "features" & ExtraVBA.get_path_separator
    example_context.set_value feature_dir, "feature_dir"
    file_names = read_dir_entries(feature_dir)
    'creating files for each test run might be slow and difficult, therefore just ensure that test files are part of the repository
    For Each table_row In data_table.table_rows
        file_name = Replace(table_row("feature_file"), "/", ExtraVBA.get_path_separator)
        TSpec.expect(file_names).includes_text file_name
    Next
End Sub

Private Function read_dir_entries(base_path As String, Optional sub_dir) As String
    
    Dim dir_entry As String
    Dim file_names As String
    Dim next_sub_dir As Variant
    Dim sub_dir_list As Collection
    
    If Not Right(base_path, 1) = ExtraVBA.get_path_separator Then
        base_path = base_path & ExtraVBA.get_path_separator
    End If
    If IsMissing(sub_dir) Then
        sub_dir = vbNullString
    End If
    dir_entry = Dir(base_path & sub_dir)
    Do While dir_entry <> vbNullString
        If Not file_names = vbNullString Then file_names = file_names & "|"
        If sub_dir <> vbNullString Then dir_entry = sub_dir & dir_entry
        file_names = file_names & dir_entry
        dir_entry = Dir()
    Loop
    Set sub_dir_list = New Collection
    dir_entry = Dir(base_path & sub_dir, vbDirectory)
    Do While dir_entry <> vbNullString
        If Left(dir_entry, 1) <> "." And GetAttr(base_path & sub_dir & dir_entry) = vbDirectory Then
            If sub_dir = vbNullString Then
                next_sub_dir = dir_entry & ExtraVBA.get_path_separator
            Else
                next_sub_dir = sub_dir & dir_entry & ExtraVBA.get_path_separator
            End If
            sub_dir_list.Add next_sub_dir
        End If
        dir_entry = Dir()
    Loop
    For Each next_sub_dir In sub_dir_list
        If Not file_names = vbNullString Then file_names = file_names & "|"
        file_names = file_names & read_dir_entries(base_path, next_sub_dir)
    Next
    read_dir_entries = file_names
End Function

Public Sub When_the_location_for_features_is_requested_FB2BB390EAA5(example_context As TContext)
    'When the location for features is requested
    example_context.set_value TFeatureLoader.get_feature_dir(example_context.get_value("doc_path")), "feature_path"
End Sub

Public Sub When_the_feature_dir_is_read_8C8F808B5CDF(example_context As TContext)
    'When the feature dir is read
    example_context.set_value TFeatureLoader.load_features(example_context.get_value("feature_path")), "features"
End Sub

Public Sub When_a_test_is_started_with_STR_as_filter_for_feat_E1CA0BEA4B3F(example_context As TContext, step_expressions As Collection)
    'When a test is started with {string} as filter for feature names
    Dim session As TSession
    Dim event_listener As EventListener
    Dim app_dir As String
    
    app_dir = example_context.get_value("application_dir")
    Set event_listener = New EventListener
    Set session = THelper.new_TSession
    Set event_listener.SenfgurkeEventLogger = session.EventLogger
    With session
        .reporter.enable_silent_mode
        .run_test New Collection, feature_filter:=step_expressions(1), application_dir:=app_dir
    End With
    example_context.set_value session, "session"
    example_context.set_value event_listener, "event_listener"
End Sub

Public Sub Then_the_feature_dir_is_set_to_STR_5A81B06F4D5F(example_context As TContext, step_expressions As Collection)
    'Then the feature dir is set to {string}
    TSpec.expect(example_context.get_value("feature_path")).to_be Join(Split(step_expressions(1), "/"), ExtraVBA.get_path_separator)
End Sub

Public Sub Then__INT_features_are_loaded_A59DDF6CC7D9(example_context As TContext, step_expressions As Collection)
    'Then  {integer} feature(s) are loaded
    TSpec.expect(example_context.get_value("features").Count).to_be step_expressions(1)
End Sub

Public Sub Then_only_STR_is_loaded_C3CA6E056B89(example_context As TContext, step_expressions As Collection)
    'Then only {string} is loaded
    TSpec.expect(example_context.get_value("loaded_files").Count).to_be 1
    TSpec.expect(example_context.get_value("loaded_files").Item(1)("origin")).includes_text step_expressions(1)
End Sub

Public Sub Then_both_features_are_loaded_7363DF6CC7D9(example_context As TContext)
    'Then both features are loaded
    TSpec.expect(example_context.get_value("loaded_files").Count).to_be 2
End Sub

Public Sub Then_only_STR_is_loaded_and_executed_1C2089F670AF(example_context As TContext, step_expressions As Collection)
    'Then only {string} is loaded and executed
    Dim event_listener As EventListener
    Dim logged_events As Collection

    Set event_listener = example_context.get_value("event_listener")
    Set logged_events = event_listener.filter_events(EVENT_FEATURE_LOADED)
    TSpec.expect(logged_events.Count).to_be 1
    TSpec.expect(logged_events(1).EventMessage).includes_text "play.feature"
    Set logged_events = event_listener.filter_events(EVENT_RUN_FEATURE_STARTED)
    TSpec.expect(logged_events.Count).to_be 1
    TSpec.expect(logged_events(1).SectionName).to_be "play"
End Sub

Public Sub Then_both_features_are_loaded_and_executed_A16D505188D1(example_context As TContext)
    'Then both features are loaded and executed
    Dim event_listener As EventListener
    Dim logged_events As Collection

    Set event_listener = example_context.get_value("event_listener")
    Set logged_events = event_listener.filter_events(EVENT_FEATURE_LOADED)
    TSpec.expect(logged_events.Count).to_be 2
    Set logged_events = event_listener.filter_events(EVENT_RUN_FEATURE_STARTED)
    TSpec.expect(logged_events.Count).to_be 2
End Sub

Public Sub Then_only_INT_features_having_names_starting_with__101BAF114ADB(example_context As TContext, step_expressions As Collection)
    'Then only {integer} features having names starting with {string} are loaded
    Dim event_listener As EventListener
    Dim logged_events As Collection
    Dim path_items As Variant
    Dim feature_name As String

    Set event_listener = example_context.get_value("event_listener")
    Set logged_events = event_listener.filter_events(EVENT_FEATURE_LOADED)
    TSpec.expect(logged_events.Count).to_be 2
    path_items = Split(logged_events(1).EventMessage, ExtraVBA.get_path_separator)
    feature_name = path_items(UBound(path_items))
    TSpec.expect(Left(feature_name, Len(step_expressions(2)))).to_be step_expressions(2)
    path_items = Split(logged_events(2).EventMessage, ExtraVBA.get_path_separator)
    feature_name = path_items(UBound(path_items))
    TSpec.expect(Left(feature_name, Len(step_expressions(2)))).to_be step_expressions(2)
End Sub
