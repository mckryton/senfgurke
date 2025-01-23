Attribute VB_Name = "build_for_Excel"
Option Explicit

Public Function create_new_document(doc_type As String, vba_project_name As String) As Workbook
    Dim new_document As Workbook
   
    Set new_document = Application.Workbooks.Add
    new_document.VBProject.Name = vba_project_name
    
    If doc_type = DOC_TYPE_ADDIN Then
        new_document.Worksheets(1).Range("A1").Value = _
            "This document contains the Senfgurke AddIn. Save it as AddIn first. Add the AddIn to Excels AddIns to use it."
        new_document.Worksheets(1).Name = "description"
    ElseIf doc_type = DOC_TYPE_STEP_DEF Then
        new_document.Worksheets(1).Range("A1").Value = _
            "This document contains the step definitions for the Senfgurke AddIn."
        new_document.Worksheets(1).Range("A2").Value = _
            "To be able to run all examples set a reference to the Senfgurke AddIn first. "
        new_document.Worksheets(1).Range("A3").Value = _
            "Call the TRun.test macro to execute all examples."
        new_document.Worksheets(1).Range("A4").Value = _
            "Open the command pane from the VBA IDE to see the test results."
        new_document.Worksheets(1).Name = "description"
    End If
    
    Set create_new_document = new_document
End Function

Public Function get_build_path() As String
    get_build_path = ThisWorkbook.Path
End Function

Public Function is_exclusive_for_other_office_apps(file_name As String) As Boolean
    Dim file_is_exclusive_for_other_office_apps As Boolean

    file_is_exclusive_for_other_office_apps = False
    If Left(file_name, Len("Word_")) = "Word_" Then file_is_exclusive_for_other_office_apps = True
    If Left(file_name, Len("Powerpoint_")) = "Powerpoint_" Then file_is_exclusive_for_other_office_apps = True
    is_exclusive_for_other_office_apps = file_is_exclusive_for_other_office_apps
End Function
