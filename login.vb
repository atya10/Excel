Private Sub CheckBox1_Click()
If Me.CheckBox1.Value = True Then
Me.TextBox1.PasswordChar = ""
Me.CheckBox1.Caption = "ÅÎÝÇÁ ßáãÉ ÇáãÑæÑ"
Else
Me.TextBox1.PasswordChar = "*"
Me.CheckBox1.Caption = "ÅÙåÇÑ ßáãÉ ÇáãÑæÑ"

End If

  
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

Private Sub CmdLogin_Click()
Dim UserName, Password As String
UserName = Me.TextBox2.Text
Password = Me.TextBox1.Text
If UserName = Help.Range("A2").Text And Password = Help.Range("B2").Text Then
Application.Visible = False
Unload Me
MainForm.Show
Else

MsgBox "ãÚáæãÇÊ ÇáÏÎæá ÛíÑ ÕÍíÍÉ", vbCritical, "ÊäÈíå"


End If
End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub UserForm_Activate()
Me.BackColor = RGB(242, 242, 242)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

End Sub
