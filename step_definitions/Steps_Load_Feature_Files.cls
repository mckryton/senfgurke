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

Private m_doc_path As String
Private m_feature_path As String
Private m_features As Collection

Public Sub Given_the_current_application_is_stored_under_STR_E4B6AFDB477F(example_context As TContext, step_expressions As Collection)
    'Given the current application is stored under {string}
    m_doc_path = Join(Split(step_expressions(1), "/"), ExtraVBA.get_path_separator)
End Sub

Public Sub Given_two_feature_files_in_the_feature_dir_8FB24F121954(example_context As TContext)
    'Given two feature files in the feature dir
    m_feature_path = ThisWorkbook.Path & ExtraVBA.get_path_separator & Join(Split("testdata/features_toplevel_pure", "/"), ExtraVBA.get_path_separator)
End Sub

Public Sub Given_one_feature_file_and_one_text_file_in_the_fe_98F5E94C9F0E(example_context As TContext)
    'Given one feature file and one text file in the feature dir
    m_feature_path = ThisWorkbook.Path & ExtraVBA.get_path_separator & Join(Split("testdata/features_toplevel_mixed", "/"), ExtraVBA.get_path_separator)
End Sub

Public Sub Given_one_feature_file_located_at_a_sub_directory__CA04B38A1CAE(example_context As TContext)
    'Given one feature file located at a sub directory under the feature dir
    m_feature_path = ThisWorkbook.Path & ExtraVBA.get_path_separator & Join(Split("testdata/features_subdir", "/"), ExtraVBA.get_path_separator)
End Sub

Public Sub When_the_location_for_features_is_requested_FB2BB390EAA5(example_context As TContext)
    'When the location for features is requested
    m_feature_path = TFeatureLoader.get_feature_dir(m_doc_path)
End Sub

Public Sub When_the_feature_dir_is_read_8C8F808B5CDF(example_context As TContext)
    'When the feature dir is read
    Set m_features = TFeatureLoader.load_features(m_feature_path)
End Sub

Public Sub Then_the_feature_dir_is_set_to_STR_5A81B06F4D5F(example_context As TContext, step_expressions As Collection)
    'Then the feature dir is set to {string}
    TSpec.expect(m_feature_path).to_be Join(Split(step_expressions(1), "/"), ExtraVBA.get_path_separator)
End Sub

Public Sub Then__INT_features_are_loaded_A59DDF6CC7D9(example_context As TContext, step_expressions As Collection)
    'Then  {integer} feature(s) are loaded
    TSpec.expect(m_features.Count).to_be step_expressions(1)
End Sub
