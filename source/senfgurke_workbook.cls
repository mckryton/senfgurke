VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "senfgurke_workbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Public Sub dummy()

    Dim x As New TStep
    
    x.parse_step_definition "Given a first step ""dummy"""
    Debug.Print x.get_step_function_name
End Sub
