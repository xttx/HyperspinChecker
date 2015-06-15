Public Class Form3
    Public filter As New List(Of String)
    Dim rom_list As New Dictionary(Of String, List(Of String))

    Private Sub Form3_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        filter.Clear()
        rom_list.Clear()
        CheckedListBox1.Items.Clear()
        Dim f As String = Class1.askVar1
        Dim cur_cat As String = ""
        Dim cur_line As String = ""
        FileOpen(1, f, OpenMode.Input, OpenAccess.Read)
        Do While Not EOF(1)
            cur_line = LineInput(1)
            If cur_line.Trim.StartsWith("[") Then
                cur_cat = cur_line.Substring(1, cur_line.Length - 2)
                If cur_cat = "" Then cur_cat = "EMPTY"
                If cur_cat.ToUpper <> "FOLDER_SETTINGS" Then rom_list.Add(cur_cat, New List(Of String))
            Else
                If cur_cat <> "" And cur_cat.ToUpper <> "FOLDER_SETTINGS" And cur_line.Trim <> "" Then
                    rom_list.Item(cur_cat).Add(cur_line)
                End If
            End If
        Loop
        FileClose(1)

        For Each s As String In rom_list.Keys
            CheckedListBox1.Items.Add(s + " (" + rom_list(s).Count.ToString + " roms)")
        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        For n As Integer = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(n, True)
        Next
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        For n As Integer = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(n, False)
        Next
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        filter.Clear()
        For Each s As String In CheckedListBox1.CheckedItems
            s = s.Substring(0, s.LastIndexOf("(") - 1)
            For Each v As String In rom_list(s)
                filter.Add(v.ToLower)
            Next
        Next
        Me.Close()
    End Sub
End Class