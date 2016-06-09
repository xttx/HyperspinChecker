Public Class FormF_systemManager_exclusions
    Dim ini As New IniFileApi()

    Private Sub FormF_systemManager_exclusions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ini.path = ".\Config.conf"

        Dim tmp As String = ""
        tmp = ini.IniReadValue("SystemManager", "required_media_number")
        If tmp <> "" AndAlso IsNumeric(tmp) Then NumericUpDown1.Value = CInt(tmp)
        tmp = ini.IniReadValue("SystemManager", "required_media_list")
        If tmp <> "" Then
            For Each s As String In tmp.Split({","c})
                If IsNumeric(s.Trim) Then
                    ListView1.Items(CInt(s.Trim)).Checked = True
                End If
            Next
        End If
        tmp = ini.IniReadValue("SystemManager", "dont_show_completed")
        If tmp <> "" AndAlso tmp <> "0" Then CheckBox1.Checked = True
    End Sub

    'Save
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ini.IniWriteValue("SystemManager", "required_media_number", NumericUpDown1.Value.ToString)
        If CheckBox1.Checked Then
            ini.IniWriteValue("SystemManager", "dont_show_completed", "1")
        Else
            ini.IniWriteValue("SystemManager", "dont_show_completed", "0")
        End If

        Dim list As String = ""
        For i As Integer = 0 To ListView1.Items.Count - 1
            If ListView1.Items(i).Checked Then
                list = list + i.ToString + ","
            End If
        Next
        If Not list = "" Then list = list.Substring(0, list.Length - 1)
        ini.IniWriteValue("SystemManager", "required_media_list", list)
        Me.Close()
    End Sub
End Class