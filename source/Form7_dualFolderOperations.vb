Public Class Form7_dualFolderOperations





    '... buttons
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim fd As New FolderBrowserDialog : fd.Description = "Select folder" : fd.ShowDialog()
        TextBox1.Text = fd.SelectedPath
    End Sub
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim fd As New FolderBrowserDialog : fd.Description = "Select folder" : fd.ShowDialog()
        TextBox2.Text = fd.SelectedPath
    End Sub

    Private Sub Form7_dualFolderOperations_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Language.localize(Me)
    End Sub
End Class