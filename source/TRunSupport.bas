Attribute VB_Name = "TRunSupport"
Option Explicit

Public Sub pending(Optional pending_msg)

    If IsMissing(pending_msg) Then pending_msg = vbNullString
    pending_msg = IIf(pending_msg = "", "x", pending_msg)
    TError.raise ERR_ID_STEP_IS_STATUS_PENDING, "TRunSupport.pending", pending_msg
End Sub
