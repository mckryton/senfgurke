VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "senfgurke_steps_workbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub exportCode()

    Dim vbe_source_object As VBComponent
    Dim base_path As String
    Dim sub_path As String
    Dim file_path As String
    Dim path_separator As String
    Dim file_suffix As String

    path_separator = ExtraVBA.get_path_separator
    base_path = ThisWorkbook.Path
    For Each vbe_source_object In Application.VBE.VBProjects("Senfgurke_Steps").VBComponents
        Select Case vbe_source_object.Type
            Case vbext_ct_StdModule
                file_suffix = "bas"
            Case vbext_ct_ClassModule
                file_suffix = "cls"
            Case vbext_ct_Document
                file_suffix = "cls"
            Case vbext_ct_MSForm
                file_suffix = "frm"
            Case Else
                file_suffix = "txt"
        End Select
        'save the code of the step definition classes in the step_definitions directory instead of the source directory
        sub_path = "step_definitions"
        file_path = base_path & path_separator & sub_path & path_separator & vbe_source_object.name & "." & file_suffix
        Debug.Print "export code to " & file_path
        #If Mac Then
            'try not to change forms unless Microsoft offers full support for forms on the Mac!
            If Not file_suffix = "frm" Then
                vbe_source_object.Export file_path
            End If
        #Else
            vbe_source_object.Export file_path
        #End If
    Next
End Sub
