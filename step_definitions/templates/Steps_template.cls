VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' When you execute your examples defined inside .feature files Senfgurke will tell the name of the
' functions that will be executed for every single step from those examples. Senfgurke will look
' for this function in all step definition classes, that are registered in the TStepRegister module.
'
' So if you change the name of this class, don't forget to apply this change to the StepRegister
' module as well.
