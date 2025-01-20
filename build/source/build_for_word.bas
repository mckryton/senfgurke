Attribute VB_Name = "build_for_word"
Option Explicit

Public Function create_new_document(vba_project_name As String) As Document
    Dim new_document As Document
   
    Set new_document = Application.Documents.Add
    new_document.VBProject.Name = vba_project_name
    Set create_new_document = new_document
End Function

Public Function get_build_path() As String
    get_build_path = ThisDocument.Path
End Function
