Attribute VB_Name = "THelper"
' This class provides helper functions for instantiating Senfgurke classes,
' so that step functions to test Senfgurke could be implemented in a separate file.
' See https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class
' for background information.

Option Explicit


Public Function new_TContext() As TContext
    Set new_TContext = New TContext
End Function

Public Function new_TFeatureRunner() As TFeatureRunner
    Set new_TFeatureRunner = New TFeatureRunner
End Function

Public Function new_TStepRunner() As TStepRunner
    Set new_TStepRunner = New TStepRunner
End Function

Public Function new_TSession() As TSession
    Set new_TSession = New TSession
End Function

Public Function new_TFeature() As TFeature
    Set new_TFeature = New TFeature
End Function
