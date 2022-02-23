VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public EventName As String
Public EventMessage As String
Public SectionName As String
Public SectionType As String
Public SectionStatus As String
Public LogTime As Single


Private Sub Class_Initialize()
    SectionStatus = STATUS_UNKNOWN
    LogTime = ExtraVBA.get_unix_timestamp_now
    EventName = vbNullString
    EventMessage = vbNullString
    SectionName = vbNullString
    SectionType = vbNullString
End Sub


