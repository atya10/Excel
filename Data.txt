Private Sub CmdDelete_Click()
DataSheet.Range("Q1").Value = Me.TextBox1.Value
Dim Answer As VbMsgBoxResult
Dim Ldelte As Integer
Answer = MsgBox("åá ÊÑíÏ ÍÐÝ ÇáÈíÇäÇÊ ÈÇáÝÚá¿", vbCritical + vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, "ÊäÈíå")

Ldelte = DataSheet.Range("S1").Value
If Answer = vbYes Then
DataSheet.Rows(Ldelte).Delete
MsgBox "Êã ÍÐÝ ÇáÈíÇäÇÊ ÈäÌÇÍ", vbInformaion + vbMsgBoxRight + vbMsgBoxRtlReading, "ÊÃßíÏ"
Else
Cancel = True
End If
Me.TextBox2.Value = ""
Me.TextBox3.Value = ""
Me.TextBox4.Value = ""
Me.TextBox5.Value = ""
Me.TextBox6.Value = ""
Me.TextBox7.Value = ""
Me.TextBox8.Value = ""
End Sub

Private Sub CmdSave_Click()
Dim Lr As Integer
Lr = DataSheet.Range("A10000").End(xlUp).Row + 1

DataSheet.Cells(Lr, "A").Value = Me.TextBox2.Value
DataSheet.Cells(Lr, "B").Value = CDate(Me.TextBox4.Value)
DataSheet.Cells(Lr, "C").Value = Me.TextBox3.Value
DataSheet.Cells(Lr, "D").Value = Me.TextBox5.Value
DataSheet.Cells(Lr, "E").Value = Me.TextBox7.Value
DataSheet.Cells(Lr, "F").Value = Me.TextBox6.Value
DataSheet.Cells(Lr, "G").Value = Me.TextBox8.Value


Me.TextBox2.Value = ""
Me.TextBox3.Value = ""
Me.TextBox4.Value = ""
Me.TextBox5.Value = ""
Me.TextBox6.Value = ""
Me.TextBox7.Value = ""
Me.TextBox8.Value = ""
MsgBox "Êã ÅÖÇÝÉ ÇáÈíÇäÇÊ ÈäÌÇÍ", vbInformaion + vbMsgBoxRight + vbMsgBoxRtlReading, "ÊÃßíÏ"

End Sub



Private Sub CmdSearch_Click()
DataSheet.Range("Q1").Value = Me.TextBox1.Value
Dim LSrch As Integer
LSrch = DataSheet.Range("S1").Value
Me.TextBox2.Value = DataSheet.Cells(LSrch, "A").Value
Me.TextBox4.Value = DataSheet.Cells(LSrch, "B").Value
Me.TextBox3.Value = DataSheet.Cells(LSrch, "C").Value
Me.TextBox5.Value = DataSheet.Cells(LSrch, "D").Value
Me.TextBox7.Value = DataSheet.Cells(LSrch, "E").Value
Me.TextBox6.Value = DataSheet.Cells(LSrch, "F").Value
Me.TextBox8.Value = DataSheet.Cells(LSrch, "G").Value
End Sub

Private Sub Cmdupdate_Click()
DataSheet.Range("Q1").Value = Me.TextBox1.Value
Dim Lupdat As Integer
Lupdat = DataSheet.Range("S1").Value
DataSheet.Cells(Lupdat, "A").Value = Me.TextBox2.Value
DataSheet.Cells(Lupdat, "B").Value = CDate(Me.TextBox4.Value)
DataSheet.Cells(Lupdat, "C").Value = Me.TextBox3.Value
DataSheet.Cells(Lupdat, "D").Value = Me.TextBox5.Value
DataSheet.Cells(Lupdat, "E").Value = Me.TextBox7.Value
DataSheet.Cells(Lupdat, "F").Value = Me.TextBox6.Value
DataSheet.Cells(Lupdat, "G").Value = Me.TextBox8.Value
Me.TextBox2.Value = ""
Me.TextBox3.Value = ""
Me.TextBox4.Value = ""
Me.TextBox5.Value = ""
Me.TextBox6.Value = ""
Me.TextBox7.Value = ""
Me.TextBox8.Value = ""
MsgBox "Êã ÊÚÏíá ÇáÈíÇäÇÊ ÈäÌÇÍ", vbInformaion + vbMsgBoxRight + vbMsgBoxRtlReading, "ÊÃßíÏ"


End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub ListBox1_Click()

Me.TextBox1.Value = Me.ListBox1.Column(0)
Me.ListBox1.Visible = False

End Sub

Private Sub TextBox1_Change()
 On Error Resume Next
If Me.TextBox1.Value <> "" Then
Me.ListBox1.Visible = True
Else
Me.ListBox1.Visible = False
Exit Sub
End If
DataSheet.Range("Q1").Value = TextBox1.Text
Me.ListBox1.Clear
Me.ListBox1.Height = 55
Dim C, Lr As Integer
Lr = DataSheet.Range("A10000").End(xlUp).Row
With DataSheet
For C = 2 To Lr
A = Len(Me.TextBox1.Text)
 If Left(DataSheet.Cells(C, "A").Value, A) = Left(Me.TextBox1.Text, A) Then
 Me.ListBox1.AddItem DataSheet.Cells(C, "A").Value
  End If
  Next C
  End With
End Sub

Private Sub UserForm_Activate()
Me.BackColor = RGB(242, 242, 242)
Me.Frame1.BackColor = RGB(242, 242, 242)
Me.Frame2.BackColor = RGB(242, 242, 242)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

End Sub
