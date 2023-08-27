Private Sub CommandButton1_Click()

End Sub

Private Sub CmdClose_Click()
Dim Anser2 As VbMsgBoxResult
Anser2 = MsgBox("åá ÊÑíÏ ÇáÎÑæÌ¿", vbCritical + vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, "ÊäÈíå")
If Anser2 = vbYes Then
ThisWorkbook.Save
Application.Quit
Else
Cancel = True
End If
End Sub

Private Sub CmdData_Click()
DataForm.Show

End Sub

Private Sub CmdSheet_Click()
Application.Visible = True

Unload Me


End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Activate()
Me.BackColor = RGB(242, 242, 242)

Me.Label1.Caption = Date
Me.Label2.Caption = Format(Date, "ddd")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)


End Sub
