Attribute VB_Name = "TStepRegister"
Option Explicit

'This function tells Senfgurke where to look if it tries to find the matching code for a single step in an example.

'REGISTER all classes with STEP IMPLEMENTATIONS HERE >>>
Public Function get_step_definition_classes() As Variant
    get_step_definition_classes = Array(New Steps_template)
End Function
