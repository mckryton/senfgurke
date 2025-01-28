Attribute VB_Name = "build_for_Powerpoint"
Option Explicit

Public Function create_new_document(doc_type As String, vba_project_name As String) As Presentation
    Dim new_document As Presentation
   
    Set new_document = Application.Presentations.Add
    new_document.VBProject.Name = vba_project_name
    
    If doc_type = DOC_TYPE_ADDIN Then
'        new_document.Windows(1).Selection.TypeText _
'            Text:="This document contains the Senfgurke AddIn. Save it as template with macros under Words AddIn directory to use it."
    ElseIf doc_type = DOC_TYPE_STEP_DEF Then
'        new_document.Windows(1).Selection.TypeText _
'            Text:="This document contains the step definitions for the Senfgurke AddIn." & vbLf & _
'                "To be able to run all examples set a reference to the Senfgurke AddIn first. " & _
'                "Call the TRun.test macro to execute all examples. Open the command pane from the VBA IDE to see the test results."
    End If
    
    Set create_new_document = new_document
End Function

Public Function get_build_path() As String
    get_build_path = Presentations("build_for_Powerpoint.pptm").Path
End Function

Public Function is_exclusive_for_other_office_apps(file_name As String) As Boolean
    Dim file_is_exclusive_for_other_office_apps As Boolean

    file_is_exclusive_for_other_office_apps = False
    If Left(file_name, Len("Excel_")) = "Excel_" Then file_is_exclusive_for_other_office_apps = True
    If Left(file_name, Len("Word_")) = "Word_" Then file_is_exclusive_for_other_office_apps = True
    is_exclusive_for_other_office_apps = file_is_exclusive_for_other_office_apps
End Function
