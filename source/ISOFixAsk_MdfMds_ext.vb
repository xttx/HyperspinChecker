Public Class ISOFixAsk_MdfMds_ext
    'Dim f1WOext, f2WOext, f1Ext, f2Ext As String
    Dim f2WOext As String

    Private Sub ISOFixAsk_MdfMds_ext_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        'f1Ext = Class1.askVar2.Substring(Class1.askVar2.LastIndexOf(".") + 1)
        'f1WOext = Class1.askVar2.Substring(0, Class1.askVar2.LastIndexOf("."))
        ListBox1.Items.Clear()
        Label2.Text = Class1.askVar1
        RadioButton1.Text = "Rename """ + Class1.askVar2 + """ to (please, select): "
        For Each s As String In Class1.askList
            f2WOext = s.Substring(0, s.LastIndexOf("."))
            ListBox1.Items.Add(f2WOext)
        Next
        ListBox1.SelectedIndex = 0
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If RadioButton1.Checked Then
            'Dim s As String = ListBox1.SelectedItem
            'f2Ext = s.Substring(s.LastIndexOf(".") + 1)
            'f2WOext = s.Substring(0, s.LastIndexOf("."))
            'Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(Class1.askVar1 + "\" + Class1.askVar2, f2WOext + "." + f1Ext)
            Class1.askResponse = ListBox1.SelectedIndex + 1
        Else
            Class1.askResponse = 0
        End If
        If CheckBox1.Checked Then Class1.askResponse = 100000 + Class1.askResponse
        Me.Close()
    End Sub
End Class