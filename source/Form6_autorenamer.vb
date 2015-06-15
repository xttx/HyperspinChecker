Public Class Form6_autorenamer
    Dim clean_game_names, clean_rom_names As List(Of String)

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If Button2.Text = "GO" Then
            If Not IsNumeric(TextBox3.Text) Then MsgBox("Please, check your min ratio % param.") : Exit Sub
            Dim min_ratio As Integer = CInt(TextBox3.Text)

            'Preprocess
            Dim curGameName, curRomName As String
            clean_game_names = New List(Of String)
            clean_rom_names = New List(Of String)
            For Each item As String In Form1.ListBox1.Items
                curGameName = item.ToUpper.Trim
                If CheckBox1.Checked Then curGameName = Remove_paranteses(curGameName)
                If CheckBox2.Checked Then curGameName = Remove_brackets(curGameName)
                If CheckBox3.Checked Then curGameName = Remove_Special(curGameName)
                If CheckBox4.Checked Then curGameName = Remove_words(curGameName)
                If CheckBox5.Checked Then curGameName = Convert_roman(curGameName)
                clean_game_names.Add(Remove_spaces(curGameName))
            Next
            For Each item As String In Form1.ListBox2.Items
                curRomName = item.ToUpper.Trim
                If CheckBox1.Checked Then curRomName = Remove_paranteses(curRomName)
                If CheckBox2.Checked Then curRomName = Remove_brackets(curRomName)
                If CheckBox3.Checked Then curRomName = Remove_Special(curRomName)
                If CheckBox4.Checked Then curRomName = Remove_words(curRomName)
                If CheckBox5.Checked Then curRomName = Convert_roman(curRomName)
                clean_rom_names.Add(Remove_spaces(curRomName))
            Next

            'Comparing
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = clean_rom_names.Count
            ProgressBar1.Value = 0
            Dim counter1 As Integer = 0, counter2 As Integer = 0
            Dim stringSplit1(), stringSplit2() As String
            Dim commonList As Collections.Generic.IEnumerable(Of String)
            Dim similarity As Double
            Dim t1, t2 As String
            For Each curRomName In clean_rom_names
                counter1 = 0
                stringSplit1 = curRomName.Split(" "c)
                For Each curGameName In clean_game_names
                    stringSplit2 = curGameName.Split(" "c)
                    commonList = stringSplit1.Intersect(stringSplit2)
                    similarity = 100 * (commonList.Count * 2) / (stringSplit1.Length + stringSplit2.Length)
                    If similarity > min_ratio Then
                        t1 = Form1.ListBox1.Items(counter1).ToString
                        t2 = Form1.ListBox2.Items(counter2).ToString
                        DataGridView1.Rows.Add({t1, t2, Math.Round(similarity, 2, MidpointRounding.AwayFromZero).ToString, "X"})
                    End If
                    counter1 += 1
                Next
                counter2 += 1
                ProgressBar1.Value += 1
                If ProgressBar1.Value Mod 10 = 0 Then ProgressBar1.Refresh()
            Next
            ProgressBar1.Value = 0
            If DataGridView1.Rows.Count > 0 Then Button2.Text = "RENAME"
        Else
            Button2.Enabled = False
            Dim path As String = Form1.TextBox4.Text
            If Not path.EndsWith("\") Then path = path + "\"
            For Each r As DataGridViewRow In DataGridView1.Rows
                If r.Cells(3).Value.ToString = "X" Then
                    Try
                        Dim f = Microsoft.VisualBasic.FileIO.FileSystem.GetFiles(path, FileIO.SearchOption.SearchTopLevelOnly, {r.Cells(1).Value.ToString + ".*"})
                        If f.Count > 0 Then
                            Dim fname As String = f(0)
                            Dim ext As String = fname.Substring(fname.LastIndexOf("."))
                            Microsoft.VisualBasic.FileSystem.Rename(fname, path + r.Cells(0).Value.ToString + ext)
                        End If
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If
            Next
            MsgBox("Done!")
            Me.Close()
        End If
    End Sub

    Private Function Remove_paranteses(s As String) As String
        If s.IndexOf("(") >= 0 Then s = s.Substring(0, s.IndexOf("("))
        Return s.Trim
    End Function
    Private Function Remove_brackets(s As String) As String
        If s.IndexOf("[") >= 0 Then s = s.Substring(0, s.IndexOf("["))
        Return s.Trim
    End Function
    Private Function Remove_Special(s As String) As String
        If TextBox1.Text.Trim = "" Then Return s
        For Each c As Char In TextBox1.Text.Trim
            If c <> "" Then s = s.Replace(c, " ")
        Next
        Return s
    End Function
    Private Function Remove_words(s As String) As String
        If TextBox2.Text.Trim = "" Then Return s
        For Each w As String In TextBox2.Text.Trim.Split({","}, StringSplitOptions.RemoveEmptyEntries)
            w = w.Trim.ToUpper
            If w <> "" Then s = s.Replace(" " + w, " ").Replace(w + " ", " ")
        Next
        Return s
    End Function
    Private Function Convert_roman(s As String) As String
        Dim t As String = s
        For Each word As String In t.Split({" "}, StringSplitOptions.RemoveEmptyEntries)
            If word.Replace("I", "").Replace("V", "").Replace("X", "").Trim = "" Then s = s.Replace(word, Convert_roman2(word))
        Next
        Return s
    End Function
    Private Function Convert_roman2(s As String) As String
        Dim ch As String
        Dim new_value, old_value As Integer
        Dim result As Integer = 0

        old_value = 1000
        For i As Integer = 0 To s.Length - 1
            ch = s.Substring(i, 1)
            Select Case ch
                Case "I"
                    new_value = 1
                Case "V"
                    new_value = 5
                Case "X"
                    new_value = 10
                Case "L"
                    new_value = 50
                Case "C"
                    new_value = 100
                Case "D"
                    new_value = 500
                Case "M"
                    new_value = 1000
            End Select
            If new_value > old_value Then
                result = result + new_value - 2 * old_value
            Else
                result = result + new_value
            End If
            old_value = new_value
        Next
        Return result.ToString
    End Function
    Private Function Remove_spaces(s As String) As String
        Return s.Replace("      ", " ").Replace("     ", " ").Replace("    ", " ").Replace("   ", " ").Replace("  ", " ")
    End Function

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If DataGridView1.Item(3, e.RowIndex).Value.ToString = "X" Then DataGridView1.Item(3, e.RowIndex).Value = "" Else DataGridView1.Item(3, e.RowIndex).Value = "X"
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Space Then
            Dim r As Integer = DataGridView1.SelectedCells(0).RowIndex
            If DataGridView1.Item(3, r).Value.ToString = "X" Then DataGridView1.Item(3, r).Value = "" Else DataGridView1.Item(3, r).Value = "X"
        End If
    End Sub



    Private Sub Form6_autorenamer_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        Dim enable As Boolean = False
        Form1.Button1.Enabled = enable
        Form1.Button2_moveUnneeded.Enabled = enable
        Form1.Button5_Associate.Enabled = enable
        Form1.GroupBox1.Enabled = enable
        Form1.GroupBox3.Enabled = enable

        Button2.Text = "GO"
        Button2.Enabled = True
        DataGridView1.Rows.Clear()
    End Sub
    Private Sub Form6_autorenamer_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim enable As Boolean = True
        Form1.Button1.Enabled = enable
        Form1.Button2_moveUnneeded.Enabled = enable
        Form1.Button5_Associate.Enabled = enable
        Form1.GroupBox1.Enabled = enable
        Form1.GroupBox3.Enabled = enable
    End Sub
End Class