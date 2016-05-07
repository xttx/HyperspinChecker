Public Class ISOFixAsk_binWoCue

    Private Sub ISOFixAsk_cueWoBin_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Label2.Text = Class1.askVar1
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Class1.askResponse = -1
        If CheckBox1.Checked Then Class1.askResponse = 99999
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Class1.askResponse = -2
        If CheckBox1.Checked Then Class1.askResponse = 99998
        Me.Close()
    End Sub

    Private Sub ISOFixAsk_cueWoBin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class