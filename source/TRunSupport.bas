Attribute VB_Name = "TRunSupport"
Option Explicit

Public Sub pending(Optional pending_msg)

    If IsMissing(pending_msg) Then pending_msg = vbNullString
    TSpec.LastFailMsg = pending_msg
    Err.Raise ERR_ID_STEP_IS_STATUS_PENDING, Description:=pending_msg
End Sub
