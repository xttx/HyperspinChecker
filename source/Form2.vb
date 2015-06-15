Public Class Form2

    Private Sub Form2_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ProgressBar1.Value = 0
        ComboBox1.SelectedIndex = Class1.i
        Select Case ComboBox1.SelectedIndex
            Case 0
                TextBox1.Text = Class1.romPath
            Case 1
                TextBox1.Text = Class1.videoPath
            Case 2
                TextBox1.Text = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Themes\"
            Case 3
                TextBox1.Text = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Wheel\"
            Case 4
                TextBox1.Text = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork1\"
            Case 5
                TextBox1.Text = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork2\"
            Case 6
                TextBox1.Text = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork3\"
            Case 7
                TextBox1.Text = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork4\"
        End Select
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        fillList()
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        fillList()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton1.CheckedChanged
        fillList()
    End Sub

    Private Sub fillList()
        Dim index As Integer
        Dim fWoExt As String = ""
        Dim allowedExt As String = ""
        ListBox1.Items.Clear()
        If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(TextBox1.Text) Then
            ProgressBar1.Value = 0 : ProgressBar1.Refresh()
            Dim files As Collections.ObjectModel.ReadOnlyCollection(Of String)
            files = Microsoft.VisualBasic.FileIO.FileSystem.GetFiles(TextBox1.Text)
            ProgressBar1.Maximum = files.Count
            For Each f As String In files
                f = f.Substring(f.LastIndexOf("\") + 1)
                fWoExt = f.Substring(0, f.LastIndexOf("."))
                index = Class1.romlist.IndexOf(fWoExt.ToLower)
                If index >= 0 Then
                    If RadioButton1.Checked Then
                        If Class1.data(index).ToString.Substring(ComboBox1.SelectedIndex, 1) = "N" Then ListBox1.Items.Add(f)
                    Else
                        ListBox1.Items.Add(f)
                    End If
                End If
                ProgressBar1.Value += 1
            Next
        End If
        Label3.Text = "Total found: " + ListBox1.Items.Count.ToString
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim dstPath As String = ""
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = ListBox1.Items.Count
        Select Case ComboBox1.SelectedIndex
            Case 0
                dstPath = Class1.romPath
            Case 1
                dstPath = Class1.videoPath
            Case 2
                dstPath = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Themes\"
            Case 3
                dstPath = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Wheel\"
            Case 4
                dstPath = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork1\"
            Case 5
                dstPath = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork2\"
            Case 6
                dstPath = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork3\"
            Case 7
                dstPath = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork4\"
        End Select

        For Each filename As String In ListBox1.Items
            Microsoft.VisualBasic.FileIO.FileSystem.CopyFile(TextBox1.Text + "\" + filename, dstPath + "\" + filename)
            ProgressBar1.Value += 1
            If ProgressBar1.Value Mod 3 = 0 Then ProgressBar1.Refresh()
        Next
        Me.Close()
    End Sub
End Class