Attribute VB_Name = "modEndProgram"
Option Explicit
Public Sub EndProgram()
  Dim Frm As Form
  Dim Ctl As Control
  
    On Local Error Resume Next '/* Some controls have no visible prpty
    For Each Frm In Forms
        For Each Ctl In Frm.Controls
            Ctl.Visible = True
            Set Ctl = Nothing
        Next Ctl
        Unload Frm
        Set Frm = Nothing
    Next Frm

End Sub

