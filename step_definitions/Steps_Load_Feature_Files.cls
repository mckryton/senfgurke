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

Dim m_doc_path As String
Dim m_feature_path As String
Dim m_features As Collection

Public Sub Given_the_current_application_is_stored_under_User_7A25EDA8AE2B()
    'Given the current application is stored under "/Users/cuke/source/senfgurke"
    m_doc_path = Join(Split("/Users/cuke/source/senfgurke", "/"), ExtraVBA.get_path_separator)
End Sub

Public Sub When_the_location_for_features_is_requested_2F12A184CE3C()
    'When the location for features is requested
    m_feature_path = TFeatureLoader.get_feature_dir(m_doc_path)
End Sub

Public Sub Then_the_feature_dir_is_set_to_Userscukesourcesenf_4E87ADCB5888()
    'Then the feature dir is set to "/Users/cuke/source/senfgurke/features"
    TSpec.expect(m_feature_path).to_be Join(Split("/Users/cuke/source/senfgurke/features", "/"), ExtraVBA.get_path_separator)
End Sub
