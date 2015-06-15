Public Class ISOFixAsk_binWoCue_ext

    Private Sub ISOFixAsk_cueWoBin_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Label2.Text = Class1.askVar1
        RadioButton1.Text = "Change '" + Class1.askVar2 + "' to match this .bin"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If RadioButton1.Checked Then
            Class1.askResponse = -3
        ElseIf RadioButton2.Checked Then
            Class1.askResponse = -1
        ElseIf RadioButton3.Checked Then
            Class1.askResponse = -2
        End If
        If CheckBox1.Checked Then Class1.askResponse = Class1.askResponse + 100000
        Me.Close()
    End Sub
End Class