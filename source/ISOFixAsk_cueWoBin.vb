Public Class ISOFixAsk_cueWoBin
    Dim f1WOext, f2WOext, f1Ext, f2Ext As String

    Private Sub ISOFixAsk_MdfMds_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        f1Ext = Class1.askVar2.Substring(Class1.askVar2.LastIndexOf(".") + 1)
        f2Ext = Class1.askVar3.Substring(Class1.askVar3.LastIndexOf(".") + 1)
        f1WOext = Class1.askVar2.Substring(0, Class1.askVar2.LastIndexOf("."))
        f2WOext = Class1.askVar3.Substring(0, Class1.askVar3.LastIndexOf("."))
        Label2.Text = Class1.askVar1
        RadioButton1.Text = "Rename .bin filename (""" + Class1.askVar3 + """) to match file inside this CUE (new name: """ + Class1.askVar2 + """)"
        RadioButton2.Text = "Rename .bin filename inside .cue (""" + Class1.askVar2 + """) to match this existing .bin filename (""" + Class1.askVar3 + """)"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If RadioButton1.Checked Then
            Class1.askResponse = 1
        ElseIf RadioButton2.Checked Then
            Class1.askResponse = 2
        Else
            Class1.askResponse = 3
        End If
        If CheckBox1.Checked Then Class1.askResponse = Class1.askResponse + 100
        Me.Close()
    End Sub
End Class