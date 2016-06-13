Imports Microsoft.VisualBasic.FileIO.FileSystem
Public Class FormF_createNewHL_system
    Public response As String = ""

    Public Sub init(msg As String, hl_path As String)
        Label1.Text = msg
        Dim s As String = ""
        Dim d = GetDirectories(hl_path + "\Settings")
        For Each dir As String In d
            s = dir.Substring(dir.LastIndexOf("\") + 1)
            ComboBox1.Items.Add(s)
        Next
        If ComboBox1.Items.Count > 0 Then ComboBox1.SelectedIndex = 0
    End Sub

    'OK
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.SelectedIndex >= 0 Then response = ComboBox1.SelectedItem.ToString
        Me.Close()
    End Sub

    'Cancel
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class