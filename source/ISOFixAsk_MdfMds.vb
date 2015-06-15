Public Class ISOFixAsk_MdfMds
    Dim f1WOext, f2WOext, f1Ext, f2Ext As String

    Private Sub ISOFixAsk_MdfMds_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        f1Ext = Class1.askVar2.Substring(Class1.askVar2.LastIndexOf(".") + 1)
        f2Ext = Class1.askVar3.Substring(Class1.askVar3.LastIndexOf(".") + 1)
        f1WOext = Class1.askVar2.Substring(0, Class1.askVar2.LastIndexOf("."))
        f2WOext = Class1.askVar3.Substring(0, Class1.askVar3.LastIndexOf("."))
        Label2.Text = Class1.askVar1
        RadioButton1.Text = "Rename """ + Class1.askVar2 + """ to """ + f2WOext + "." + f2Ext + """"
        RadioButton2.Text = "Rename """ + Class1.askVar3 + """ to """ + f1WOext + "." + f1Ext + """"
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
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