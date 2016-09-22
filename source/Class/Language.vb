Public Class Language
    Public Shared Preset_Checker As String = "Preset - Checker"
    Public Shared Preset_Editor As String = "Preset - Editor"
    Public Shared Preset_HL As String = "Preset - HL / RL Media"
    Public Shared Preset_SysMngr_Default As String = "Preset - Default"
    Public Shared Preset_SysMngr_Checker As String = "Preset - Checker"
    Public Shared Preset_SysMngr_Manager As String = "Preset - Manager"
    Public Shared Preset_SysMngr_HLMedia As String = "Preset - HL / RL Media"
    Public Shared Preset_SysMngr_Full As String = "Preset - Full"
    Public Shared Save_current_cols_conf_as_startup As String = "Save current cols config as startup"

    Public Shared strings As New Dictionary(Of String, String)

    Public Shared Sub localize(form As Form)
        For Each k In strings.Keys
            If Not k.ToUpper.StartsWith(form.Name.ToUpper) Then Continue For

            Dim s() = k.Split({"."c})
            Dim val As String = strings(k)

            Dim tmp As Control = Nothing
            Dim tmp_menu_item As ToolStripItem = Nothing
            For Each c In s
                If c.ToUpper = form.Name.ToUpper Then
                    tmp = form
                Else
                    If tmp_menu_item IsNot Nothing Then
                        tmp_menu_item = DirectCast(tmp_menu_item, ToolStripMenuItem).DropDownItems(c)
                    Else
                        If TypeOf tmp Is MenuStrip Then
                            tmp_menu_item = DirectCast(tmp, MenuStrip).Items(c)
                        Else
                            tmp = tmp.Controls(c)
                        End If
                    End If

                    If tmp.Name.ToUpper = "SPLITCONTAINER1" Then
                        Dim tmp2 = DirectCast(tmp, SplitContainer)
                        tmp2.Panel1.Name = "Panel1"
                        tmp2.Panel2.Name = "Panel2"
                    End If
                End If
            Next

            If tmp_menu_item IsNot Nothing Then
                tmp_menu_item.Text = val
            Else
                tmp.Text = val
            End If
        Next
    End Sub
End Class
