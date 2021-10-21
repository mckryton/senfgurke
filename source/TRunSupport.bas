Attribute VB_Name = "TRunSupport"
Option Explicit

Public Sub pending(Optional pending_msg)

    If IsMissing(pending_msg) Then pending_msg = vbNullString
    TSpec.LastFailMsg = pending_msg
    pending_msg = IIf(pending_msg = "", "x", pending_msg)
    Err.raise ERR_ID_STEP_IS_STATUS_PENDING, description:=pending_msg
End Sub
