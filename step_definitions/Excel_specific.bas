Attribute VB_Name = "word_specific"
Option Explicit

Public Function get_app_root_path() As String
    'assume that the document containing the step definitions is saved in
    ' the root directory of the Senfgurke repository while the Senfgurke is
    ' the application (app) under test
    get_app_root_path = ThisDocument.Path
End Function
