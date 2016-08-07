Public Class Form4_genres_favorites_preview

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            Me.TopMost = True
        Else
            Me.TopMost = False
        End If
    End Sub

    Private Sub Form4_genres_favorites_preview_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Language.localize(Me)
    End Sub
End Class