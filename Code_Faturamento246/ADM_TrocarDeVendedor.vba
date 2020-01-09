Option Compare Database

Sub cmdSortAsc_Click()
    On Error GoTo ErrHandler
    Screen.PreviousControl.SetFocus
    DoCmd.RunCommand acCmdSortAscending
    Exit Sub

ErrHandler:
  Select Case Err
    Case 2046
      'Command not available
      MsgBox "Sorting is not available at this time.", vbCritical, "Not Available"
    Case Else
      MsgBox Err.Number & ":-" & vbCrLf & Err.Description
  End Select
End Sub



Sub cmdSortDesc_Click()
    On Error GoTo ErrHandler
    Screen.PreviousControl.SetFocus
    DoCmd.RunCommand acCmdSortDescending
    Exit Sub

ErrHandler:
  Select Case Err
    Case 2046
      'Command not available
      MsgBox "Sorting is not available at this time.", vbCritical, "Not Available"
    Case Else
      MsgBox Err.Number & ":-" & vbCrLf & Err.Description
  End Select
End Sub

