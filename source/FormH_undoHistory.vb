Public Class FormH_undoHistory
    Dim frm As Form1 = DirectCast(Application.OpenForms("Form1"), Form1)
    Private Sub FormH_undoHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each list In frm.undo_humanReadable
            For Each s In list
                ListBox1.Items.Add(s.Replace("\\", "\"))
            Next
        Next
    End Sub
End Class