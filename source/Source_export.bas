Attribute VB_Name = "Source_export"
Option Explicit

'This module is for exporting all modules and classes as text. This is for using
' git built-in functions to find changes in text files. You may also use the scripts
' from the build directory to rebuild the AddIn and the step definitions.

'Before using this function a reference to Microsoft Visual Basic for applications Extensibility must be set.

'Private Sub exportCode()
'    Dim vbe_source_object As VBComponent
'    Dim base_path As String
'    Dim sub_path As String
'    Dim file_path As String
'    Dim path_separator As String
'    Dim file_suffix As String
'
'    path_separator = ExtraVBA.get_path_separator
'    base_path = get_app_root_path()
'    For Each vbe_source_object In Application.VBE.VBProjects("Senfgurke_Steps").VBComponents
'        Select Case vbe_source_object.Type
'            Case vbext_ct_StdModule
'                file_suffix = "bas"
'            Case vbext_ct_ClassModule
'                file_suffix = "cls"
'            Case vbext_ct_Document
'                file_suffix = "cls"
'            Case vbext_ct_MSForm
'                file_suffix = "frm"
'            Case Else
'                file_suffix = "txt"
'        End Select
'        'save the code of the step definition classes in the step_definitions directory instead of the source directory
'        sub_path = "step_definitions"
'        file_path = base_path & path_separator & sub_path & path_separator & vbe_source_object.Name & "." & file_suffix
'        Debug.Print "export code to " & file_path
'        #If Mac Then
'            'try not to change forms unless Microsoft offers full support for forms on the Mac!
'            If Not file_suffix = "frm" Then
'                vbe_source_object.Export file_path
'            End If
'        #Else
'            vbe_source_object.Export file_path
'        #End If
'    Next
'End Sub


