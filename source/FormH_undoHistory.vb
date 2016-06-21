Public Class FormH_undoHistory
    Private Sub FormH_undoHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each list In Form1.undo_humanReadable
            For Each s In list
                ListBox1.Items.Add(s.Replace("\\", "\"))
            Next
        Next
    End Sub
End Class