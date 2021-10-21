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
    example_context.set_value ThisWorkbook.Path & ExtraVBA.get_path_separator & Join(Split("testdata/features_toplevel_pure", "/"), ExtraVBA.get_path_separator), "feature_path"
End Sub

Public Sub Given_one_feature_file_and_one_text_file_in_the_fe_98F5E94C9F0E(example_context As TContext)
    'Given one feature file and one text file in the feature dir
    example_context.set_value ThisWorkbook.Path & ExtraVBA.get_path_separator & Join(Split("testdata/features_toplevel_mixed", "/"), ExtraVBA.get_path_separator), "feature_path"
End Sub

Public Sub Given_one_feature_file_located_at_a_sub_directory__CA04B38A1CAE(example_context As TContext)
    'Given one feature file located at a sub directory under the feature dir
    example_context.set_value ThisWorkbook.Path & ExtraVBA.get_path_separator & Join(Split("testdata/features_subdir", "/"), ExtraVBA.get_path_separator), "feature_path"
End Sub

Public Sub Given_a_feature_file_STR_D7FE0A9A5327(example_context As TContext, step_expressions As Collection)
    'Given a feature file {string}
    Dim dir_entry As String
    Dim file_names As String
    
    If Not example_context.value_exists("feature_dir") Then
        example_context.set_value senfgurke_workbook.Path & ExtraVBA.get_path_separator _
            & "testdata" & ExtraVBA.get_path_separator _
            & "filter_for_loading_features" & ExtraVBA.get_path_separator _
            , "feature_dir"
    End If
    dir_entry = Dir(example_context.get_value("feature_dir"))
    Do While dir_entry <> vbNullString
        If Not file_names = vbNullString Then file_names = file_names & "|"
        file_names = file_names & dir_entry
        dir_entry = Dir()
    Loop
    'creating files for each test run might be slow and difficult, therefore just ensure that test files are part of the repository
    TSpec.expect(file_names).includes_text (step_expressions(1))
End Sub

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
    example_context.set_value New TSession, "session"
    With example_context.get_value("session")
        .reporter.enable_silent_mode
        .statistics.enable_event_log
        .run_test feature_filter:=step_expressions(1), feature_dir:=example_context.get_value("feature_dir")
    End With
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
    Dim session As TSession
    Dim logged_events As Collection

    Set session = example_context.get_value("session")
    Set logged_events = session.statistics.filter_events(LOG_EVENT_LOAD_FEATURE)
    TSpec.expect(logged_events.Count).to_be 1
    TSpec.expect(logged_events(1)("msg")).includes_text "play.feature"
    Set logged_events = session.statistics.filter_events(LOG_EVENT_RUN_FEATURE)
    TSpec.expect(logged_events.Count).to_be 1
    TSpec.expect(logged_events(1)("msg")).to_be "Feature: play"
End Sub

Public Sub Then_both_features_are_loaded_and_executed_A16D505188D1(example_context As TContext)
    'Then both features are loaded and executed
    Dim session As TSession
    Dim logged_events As Collection

    Set session = example_context.get_value("session")
    Set logged_events = session.statistics.filter_events(LOG_EVENT_LOAD_FEATURE)
    TSpec.expect(logged_events.Count).to_be 2
    Set logged_events = session.statistics.filter_events(LOG_EVENT_RUN_FEATURE)
    TSpec.expect(logged_events.Count).to_be 2
End Sub
