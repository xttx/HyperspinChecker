Public Class ISOFixAsk_cueWoBin_ext
    'Dim f1WOext, f2WOext, f1Ext, f2Ext As String
    Dim f2WOext As String

    Private Sub ISOFixAsk_MdfMds_ext_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        ListBox1.Items.Clear()
        Label2.Text = Class1.askVar1
        RadioButton1.Text = "Rename selected .bin to match filename inside .cue {" + Class1.askVar2 + "}:"
        For Each s As String In Class1.askList
            'f2WOext = s.Substring(0, s.LastIndexOf("."))
            ListBox1.Items.Add(s)
        Next
        ListBox1.SelectedIndex = 0
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If RadioButton1.Checked Then
            Class1.askResponse = ListBox1.SelectedIndex + 1
        ElseIf RadioButton2.Checked Then
            Class1.askResponse = ListBox1.SelectedIndex + 10001
        Else
            Class1.askResponse = 0
        End If
        If CheckBox1.Checked Then Class1.askResponse = 100000 + Class1.askResponse
        Me.Close()
    End Sub
End Class