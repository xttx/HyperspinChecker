Public Class ISOFixAsk_binWoCue_extPlus
    Private Sub ISOFixAsk_cueWoBin_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        ListBox1.Items.Clear()
        Label2.Text = Class1.askVar1
        For Each s In Class1.askList
            ListBox1.Items.Add(s)
        Next
        ListBox1.SelectedIndex = 0
        If RadioButton1.Checked Then CheckBox1.Enabled = False Else CheckBox1.Enabled = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If RadioButton1.Checked Then
            Class1.askResponse = ListBox1.SelectedIndex
        ElseIf RadioButton2.Checked Then
            Class1.askResponse = -1
        ElseIf RadioButton3.Checked Then
            Class1.askResponse = -2
        End If
        If CheckBox1.Checked And Class1.askResponse < 0 Then Class1.askResponse = Class1.askResponse + 100000
        Me.Close()
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, RadioButton3.CheckedChanged
        If RadioButton1.Checked Then CheckBox1.Enabled = False Else CheckBox1.Enabled = True
    End Sub
End Class