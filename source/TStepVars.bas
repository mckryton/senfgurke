Attribute VB_Name = "TStepVars"
Option Explicit

Public gherkin_text As String
Public example As TExample
Public example_step As TStep
Public feature_statistics As Collection
Public m_parsed_feature As TFeature
Public m_step_function_name As String
Public step As TStep
Public first_step As TStep
Public Background As TBackground
Public ParsedFeatures As Collection
Public ReportOutput As String
Public ReportFormatter As Variant
Public ReportMessages As Collection
Public CodeTemplate As String
Public Session As TSession
Public Duration As Long
Public StatisticResults As String
Public parse_error As Collection
Public loaded_features As Collection
Public feature_origin As String
