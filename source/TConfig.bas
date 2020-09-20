Attribute VB_Name = "TConfig"
Option Explicit

Public Property Get StepImplementations() As Variant
    
    'add all classes with step implmentations here:
    StepImplementations = Array(New Steps_Load_Feature_Files, New Steps_Run_Steps)
End Property
