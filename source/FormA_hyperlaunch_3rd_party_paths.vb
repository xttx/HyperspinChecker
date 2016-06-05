Imports Microsoft.Win32
Imports Microsoft.VisualBasic.FileIO.FileSystem
Public Class FormA_hyperlaunch_3rd_party_paths
    Dim HLPath As String = ""
    Dim HLPathNoExt As String = ""
    Dim HLHQPath As String = ""
    Dim ini As New IniFileApi
    Dim doNotUpdate As Boolean = False

    Private Sub FormA_hyperlaunch_3rd_party_paths_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        HLPath = Form1.TextBox18.Text.Replace("\\", "\").Replace("\\", "\").Trim
        If HLPath.ToUpper.EndsWith("EXE") Then HLPathNoExt = HLPath.Substring(0, HLPath.LastIndexOf("\")).Trim Else HLPathNoExt = HLPath.Trim
        If Not HLPathNoExt.EndsWith("\") Then HLPathNoExt = HLPathNoExt + "\"

        If FileExists(HLPathNoExt + "HyperLaunchHQ.exe") Then
            HLHQPath = HLPathNoExt
        Else
            If FileExists(HLPathNoExt + "HyperLaunchHQ\HyperLaunchHQ.exe") Then
                HLHQPath = HLPathNoExt + "HyperLaunchHQ\"
            Else
                If FileExists(HLPathNoExt + "RocketLauncherUI.exe") Then
                    HLHQPath = HLPathNoExt
                Else
                    If FileExists(HLPathNoExt + "RocketLauncherUI\RocketLauncherUI.exe") Then
                        HLHQPath = HLPathNoExt + "RocketLauncherUI\"
                    End If
                End If
            End If
        End If
        If HLHQPath.Trim = "" Then TextBox_TextChanged(TextBox0, New System.EventArgs) : Exit Sub

        TextBox0.Text = HLHQPath

        If FileExists((HLPathNoExt + "\Settings\HyperLaunch.ini").Replace("\\", "\")) Then
            ini.path = (HLPathNoExt + "\Settings\HyperLaunch.ini").Replace("\\", "\")
        ElseIf FileExists((HLPathNoExt + "\Settings\RocketLauncher.ini").Replace("\\", "\")) Then
            ini.path = (HLPathNoExt + "\Settings\RocketLauncher.ini").Replace("\\", "\")
        Else
            Exit Sub
        End If

        TextBox1.Text = ini.IniReadValue("Settings", "Modules_Path")
        TextBox2.Text = ini.IniReadValue("Settings", "HyperLaunch_Media_Path")
        TextBox3.Text = ini.IniReadValue("Settings", "Frontend_Path")
        TextBox4.Text = ini.IniReadValue("Settings", "Profiles_Path")
        TextBox5.Text = ini.IniReadValue("7z", "7z_Path")
        TextBox6.Text = ini.IniReadValue("HyperPause", "HyperPause_HiToText_Path")
        TextBox7.Text = ini.IniReadValue("DAEMON Tools", "DAEMON_Tools_Path")
        TextBox8.Text = ini.IniReadValue("Keymapper", "Xpadder_Path")
        TextBox9.Text = ini.IniReadValue("Keymapper", "JoyToKey_Path")
        TextBox10.Text = ini.IniReadValue("Vjoy", "VJoy_Path")

        If TextBox2.Text = "" Then TextBox2.Text = ini.IniReadValue("Settings", "RocketLauncher_Media_Path")
        If TextBox3.Text = "" Then TextBox3.Text = ini.IniReadValue("Settings", "Default_Front_End_Path")
        If TextBox6.Text = "" Then TextBox6.Text = ini.IniReadValue("Pause", "Pause_HiToText_Path")
        If TextBox7.Text = "" Then TextBox7.Text = ini.IniReadValue("Virtual Drive", "Virtual_Drive_Path")
    End Sub

    Private Sub read_HLHQ_ini()
        If FileExists((HLHQPath + "\Settings\HyperLaunchHQ.ini").Replace("\\", "\")) Then
            ini.path = (HLHQPath + "\Settings\HyperLaunchHQ.ini").Replace("\\", "\")
        ElseIf FileExists((HLHQPath + "\Settings\RocketLauncherUI.ini").Replace("\\", "\")) Then
            ini.path = (HLHQPath + "\Settings\RocketLauncherUI.ini").Replace("\\", "\")
        End If

        TextBox11.Text = ini.IniReadValue("Paths", "HL_Folder")
        If TextBox11.Text = "" Then TextBox11.Text = ini.IniReadValue("Settings", "HL_Folder")
        If TextBox11.Text = "" Then TextBox11.Text = ini.IniReadValue("Paths", "RL_Folder")
        TextBox12.Text = ini.IniReadValue("Paths", "Media_Folder_Path")
        If TextBox12.Text = "" Then TextBox12.Text = ini.IniReadValue("Settings", "Media_Folder_Path")

        ini.path = (HLHQPath + "\Settings\Frontends.ini").Replace("\\", "\")
        TextBox13.Text = ini.IniReadValue("Hyperspin", "Path")
        If TextBox13.Text = "" Then TextBox13.Text = ini.IniReadValue("HS", "Path")

        TextBox14.Text = ini.IniReadValue("HyperLaunchHQ", "Path")
        If TextBox14.Text = "" Then TextBox14.Text = ini.IniReadValue("RocketLauncherUI", "Path")

        Dim f As String = ini.IniReadValue("Settings", "Default_Frontend").ToUpper
        If f = "HS" Or f = "HYPERSPIN" Then CheckBox2.Checked = True
        If f = "HYPERLAUNCHHQ" Or f = "ROCKETLAUNCHERUI" Then CheckBox3.Checked = True
    End Sub

    Private Sub TextBox_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox0.TextChanged, TextBox1.TextChanged, TextBox2.TextChanged, _
        TextBox3.TextChanged, TextBox4.TextChanged, TextBox5.TextChanged, TextBox6.TextChanged, TextBox7.TextChanged, TextBox8.TextChanged, _
        TextBox9.TextChanged, TextBox10.TextChanged, TextBox11.TextChanged, TextBox12.TextChanged, TextBox13.TextChanged, TextBox14.TextChanged

        If doNotUpdate Then Exit Sub
        Dim ID As Integer = CInt(DirectCast(sender, TextBox).Name.Substring(7))

        If ID = 0 Then 'HLHQ Path
            If FileExists(TextBox0.Text + "\HyperLaunchHQ.exe") OrElse FileExists(TextBox0.Text + "\RocketLauncherUI.exe") Then
                TextBox0.BackColor = Form1.colorYES : read_HLHQ_ini()
            Else
                TextBox0.BackColor = Form1.colorNO
                TextBox11.Text = "HLHQ Path not set" : TextBox12.Text = "HLHQ Path not set"
                TextBox13.Text = "HLHQ Path not set" : TextBox14.Text = "HLHQ Path not set"
            End If
            HLHQPath = (TextBox0.Text + "\").Replace("\\", "\").Replace("\\", "\").Trim
        End If

        If ID = 1 Then 'Modules Path
            If DirectoryExists(Relative_Path_to_Absolute(HLPathNoExt, TextBox1.Text)) Then TextBox1.BackColor = Form1.colorYES Else TextBox1.BackColor = Form1.colorNO
        End If

        If ID = 2 Then 'Media Path
            If DirectoryExists(Relative_Path_to_Absolute(HLPathNoExt, TextBox2.Text)) Then TextBox2.BackColor = Form1.colorYES Else TextBox2.BackColor = Form1.colorNO
        End If

        If ID = 3 Then 'Frontend Path
            If FileExists(Relative_Path_to_Absolute(HLPathNoExt, TextBox3.Text)) Then TextBox3.BackColor = Form1.colorYES Else TextBox3.BackColor = Form1.colorNO
        End If

        If ID = 4 Then 'Profiles Path
            If DirectoryExists(Relative_Path_to_Absolute(HLPathNoExt, TextBox4.Text)) Then TextBox4.BackColor = Form1.colorYES Else TextBox4.BackColor = Form1.colorNO
        End If

        If ID = 5 Then '7z Path
            If FileExists(Relative_Path_to_Absolute(HLPathNoExt, TextBox5.Text)) Then TextBox5.BackColor = Form1.colorYES Else TextBox5.BackColor = Form1.colorNO
        End If

        If ID = 6 Then 'HiToText Path
            If FileExists(Relative_Path_to_Absolute(HLPathNoExt, TextBox6.Text)) Then TextBox6.BackColor = Form1.colorYES Else TextBox6.BackColor = Form1.colorNO
        End If

        If ID = 7 Then 'Daemon Tools Path
            If FileExists(Relative_Path_to_Absolute(HLPathNoExt, TextBox7.Text)) Then TextBox7.BackColor = Form1.colorYES Else TextBox7.BackColor = Form1.colorNO
        End If

        If ID = 8 Then 'Daemon Tools Path
            If FileExists(Relative_Path_to_Absolute(HLPathNoExt, TextBox8.Text)) Then TextBox8.BackColor = Form1.colorYES Else TextBox8.BackColor = Form1.colorNO
        End If

        If ID = 9 Then 'Joy To Key Path
            If FileExists(Relative_Path_to_Absolute(HLPathNoExt, TextBox9.Text)) Then TextBox9.BackColor = Form1.colorYES Else TextBox9.BackColor = Form1.colorNO
        End If

        If ID = 10 Then 'Vjoy Path
            If FileExists(Relative_Path_to_Absolute(HLPathNoExt, TextBox10.Text)) Then TextBox10.BackColor = Form1.colorYES Else TextBox10.BackColor = Form1.colorNO
        End If

        If ID = 11 Then 'Hyperlaunch Path for HLHQ
            If DirectoryExists(Relative_Path_to_Absolute(HLHQPath, TextBox11.Text)) Then TextBox11.BackColor = Form1.colorYES Else TextBox11.BackColor = Form1.colorNO
        End If

        If ID = 12 Then 'Media Path for HLHQ
            If DirectoryExists(Relative_Path_to_Absolute(HLHQPath, TextBox12.Text)) Then TextBox12.BackColor = Form1.colorYES Else TextBox12.BackColor = Form1.colorNO
        End If

        If ID = 13 Then 'HS Plugin
            If FileExists(Relative_Path_to_Absolute(HLHQPath, TextBox13.Text)) Then TextBox13.BackColor = Form1.colorYES Else TextBox13.BackColor = Form1.colorNO
        End If

        If ID = 14 Then 'HLHQ Plugin
            If FileExists(Relative_Path_to_Absolute(HLHQPath, TextBox14.Text)) Then TextBox14.BackColor = Form1.colorYES Else TextBox14.BackColor = Form1.colorNO
        End If
    End Sub

    Private Function Relative_Path_to_Absolute(path_from As String, relative As String) As String
        If Not relative.StartsWith(".") Then Return relative
        Return (path_from + "\" + relative).Replace("\\", "\").Replace("\\", "\")
    End Function

    Private Function Absolute_Path_to_Relative(path_from As String, path_to As String) As String
        If path_to.StartsWith(".") Then Return path_to
        If String.IsNullOrEmpty(path_from) Then Return ""
        If String.IsNullOrEmpty(path_to) Then Return ""

        If path_to.Substring(0, 1).ToUpper <> path_from.Substring(0, 1).ToUpper Then Return path_to
        Dim UriFrom = New Uri(path_from)
        Dim UriTo = New Uri(path_to)

        If UriFrom.Scheme <> UriTo.Scheme Then Return path_to

        Dim UriRelative As Uri = UriFrom.MakeRelativeUri(UriTo)
        Dim relativePath As String = Uri.UnescapeDataString(UriRelative.ToString)

        If UriTo.Scheme.ToUpperInvariant = "FILE" Then
            relativePath = relativePath.Replace(System.IO.Path.AltDirectorySeparatorChar, System.IO.Path.DirectorySeparatorChar)
        End If
        If Not relativePath.StartsWith(".") Then relativePath = ".\" + relativePath
        Return relativePath
    End Function

    'Try autodetect paths
    Private Sub Button15_Click(sender As System.Object, e As System.EventArgs) Handles Button15.Click
        'HLHQ Path
        Class1.Log("Detecting HLHQ Path")
        If Not CheckBox1.Checked Or TextBox0.BackColor = Form1.colorNO Then
            If FileExists(HLPathNoExt + "\HyperLaunchHQ.exe") Then
                TextBox0.Text = (HLPathNoExt + "\").Replace("\\", "\").Replace("\\", "\")
            ElseIf FileExists(HLPathNoExt + "\HyperLaunchHQ\HyperLaunchHQ.exe") Then
                TextBox0.Text = (HLPathNoExt + "\HyperLaunchHQ\").Replace("\\", "\").Replace("\\", "\")
            ElseIf FileExists(HLPathNoExt + "\RocketLauncherUI.exe") Then
                TextBox0.Text = (HLPathNoExt + "\").Replace("\\", "\").Replace("\\", "\")
            ElseIf FileExists(HLPathNoExt + "\RocketLauncherUI\RocketLauncherUI.exe") Then
                TextBox0.Text = (HLPathNoExt + "\RocketLauncherUI\").Replace("\\", "\").Replace("\\", "\")
            End If
        End If

        'Modules
        Class1.Log("Detecting Modules Path")
        If Not CheckBox1.Checked Or TextBox1.BackColor = Form1.colorNO Then
            If DirectoryExists(HLPathNoExt + "\Modules") Then
                TextBox1.Text = Absolute_Path_to_Relative(HLPathNoExt, (HLPathNoExt + "\Modules\").Replace("\\", "\").Replace("\\", "\"))
            End If
        End If

        'Media
        Class1.Log("Detecting Media Path")
        If Not CheckBox1.Checked Or TextBox2.BackColor = Form1.colorNO Then
            If DirectoryExists(HLPathNoExt + "\Media") Then
                TextBox2.Text = Absolute_Path_to_Relative(HLPathNoExt, (HLPathNoExt + "\Media\").Replace("\\", "\").Replace("\\", "\"))
            End If
        End If

        'Frontend
        Class1.Log("Detecting Frontend Path")
        If Not CheckBox1.Checked Or TextBox3.BackColor = Form1.colorNO Then
            If FileExists(Class1.HyperspinPath + "\HyperSpin.exe") Then
                TextBox3.Text = Absolute_Path_to_Relative(HLPath, (Class1.HyperspinPath + "\HyperSpin.exe").Replace("\\", "\").Replace("\\", "\"))
            End If
        End If

        'Profiles
        Class1.Log("Detecting Profiles Path")
        If Not CheckBox1.Checked Or TextBox4.BackColor = Form1.colorNO Then
            If DirectoryExists(HLPathNoExt + "\Profiles") Then
                TextBox4.Text = Absolute_Path_to_Relative(HLPathNoExt, (HLPathNoExt + "\Profiles\").Replace("\\", "\").Replace("\\", "\"))
            End If
        End If

        '7z
        Class1.Log("Detecting 7z Path")
        If Not CheckBox1.Checked Or TextBox5.BackColor = Form1.colorNO Then
            If FileExists(HLPathNoExt + "\Module Extensions\7z.exe") Then
                TextBox5.Text = Absolute_Path_to_Relative(HLPathNoExt, (HLPathNoExt + "\Module Extensions\7z.exe").Replace("\\", "\").Replace("\\", "\"))
            End If
        End If

        'HiToText
        Class1.Log("Detecting HiToText Path")
        If Not CheckBox1.Checked Or TextBox6.BackColor = Form1.colorNO Then
            If FileExists(HLPathNoExt + "\Module Extensions\HiToText.exe") Then
                TextBox6.Text = Absolute_Path_to_Relative(HLPathNoExt, (HLPathNoExt + "\Module Extensions\HiToText.exe").Replace("\\", "\").Replace("\\", "\"))
            End If
        End If

        'Daemon Tools
        Class1.Log("Detecting Daemon Tools Path")
        If Not CheckBox1.Checked Or TextBox7.BackColor = Form1.colorNO Then
            Dim regKey As RegistryKey = Registry.ClassesRoot.OpenSubKey("DAEMON.Tools.Lite\\shell\\open\\command", False)
            If (Not regKey Is Nothing) Then
                Dim tmp As String = regKey.GetValue("", 0).ToString
                If tmp IsNot Nothing AndAlso tmp.IndexOf("""") = 0 Then
                    tmp = tmp.Substring(1)
                    If tmp.IndexOf("""") > 0 Then
                        TextBox7.Text = tmp.Substring(0, tmp.IndexOf(""""))
                    End If
                End If
            End If
            If regKey IsNot Nothing Then regKey.Close()
        End If

        'Xpadder

        'JoyToKey

        'Vjoy

        'HL For HLHQ
        Class1.Log("Detecting Daemon HLHQ Path")
        If Not CheckBox1.Checked Or TextBox11.BackColor = Form1.colorNO Then
            If FileExists(HLPathNoExt + "\HyperLaunch.exe") OrElse FileExists(HLPathNoExt + "\RocketLauncher.exe") Then
                TextBox11.Text = Absolute_Path_to_Relative(HLHQPath, (HLPathNoExt + "\").Replace("\\", "\").Replace("\\", "\"))
            End If
        End If

        'HLHQ Media
        Class1.Log("Detecting Daemon HLHQ Media Path")
        If Not CheckBox1.Checked Or TextBox12.BackColor = Form1.colorNO Then
            If DirectoryExists(HLHQPath + "\Media") Then
                TextBox12.Text = Absolute_Path_to_Relative(HLHQPath, (HLHQPath + "\Media\").Replace("\\", "\").Replace("\\", "\"))
            End If
        End If

        'HS Plugin
        'TODO Maybe this should be relative to HL path and not to HLHQ path
        Class1.Log("Detecting HS Plugin Path")
        If Not CheckBox1.Checked Or TextBox13.BackColor = Form1.colorNO Then
            If FileExists(Class1.HyperspinPath + "\hyperspin.exe") Then
                TextBox13.Text = Absolute_Path_to_Relative(HLHQPath, (Class1.HyperspinPath + "\HyperSpin.exe").Replace("\\", "\").Replace("\\", "\"))
            End If
        End If

        'HLHQ Plugin
        'TODO Maybe this should be relative to HL path and not to HLHQ path
        Class1.Log("Detecting HLHQ Plugin Path")
        If Not CheckBox1.Checked Or TextBox14.BackColor = Form1.colorNO Then
            If FileExists(HLHQPath + "\HyperLaunchHQ.exe") Then
                TextBox14.Text = Absolute_Path_to_Relative(HLHQPath, (HLHQPath + "\HyperLaunchHQ.exe").Replace("\\", "\").Replace("\\", "\"))
            ElseIf FileExists(HLHQPath + "\RocketLauncherUI.exe") Then
                TextBox14.Text = Absolute_Path_to_Relative(HLHQPath, (HLHQPath + "\RocketLauncherUI.exe").Replace("\\", "\").Replace("\\", "\"))
            End If
        End If
    End Sub

    'Make paths relative
    Private Sub Button17_Click(sender As System.Object, e As System.EventArgs) Handles Button17.Click
        TextBox1.Text = Absolute_Path_to_Relative(HLPath, (TextBox1.Text + "\").Replace("\\", "\").Replace("\\", "\"))
        TextBox2.Text = Absolute_Path_to_Relative(HLPath, (TextBox2.Text + "\").Replace("\\", "\").Replace("\\", "\"))
        TextBox3.Text = Absolute_Path_to_Relative(HLPath, TextBox3.Text.Replace("\\", "\").Replace("\\", "\"))
        TextBox4.Text = Absolute_Path_to_Relative(HLPath, (TextBox4.Text + "\").Replace("\\", "\").Replace("\\", "\"))
        TextBox5.Text = Absolute_Path_to_Relative(HLPath, TextBox5.Text.Replace("\\", "\").Replace("\\", "\"))
        TextBox6.Text = Absolute_Path_to_Relative(HLPath, TextBox6.Text.Replace("\\", "\").Replace("\\", "\"))
        TextBox11.Text = Absolute_Path_to_Relative(HLHQPath, (TextBox11.Text + "\").Replace("\\", "\").Replace("\\", "\"))
        TextBox12.Text = Absolute_Path_to_Relative(HLHQPath, (TextBox12.Text + "\").Replace("\\", "\").Replace("\\", "\"))
        'TODO Maybe this should be relative to HL path and not to HLHQ path
        TextBox13.Text = Absolute_Path_to_Relative(HLHQPath, TextBox13.Text.Replace("\\", "\").Replace("\\", "\"))
        'TODO Maybe this should be relative to HL path and not to HLHQ path
        TextBox14.Text = Absolute_Path_to_Relative(HLHQPath, TextBox14.Text.Replace("\\", "\").Replace("\\", "\"))
    End Sub

    'Update INI
    Private Sub Button16_Click(sender As System.Object, e As System.EventArgs) Handles Button16.Click
        doNotUpdate = True
        For Each c As Control In Me.Controls
            If c.GetType Is GetType(TextBox) Then
                Dim t As TextBox = DirectCast(c, TextBox)
                t.Text = t.Text.Replace("\\", "\").Replace("\\", "\").Replace("\\", "\").Trim
                If t.Text.EndsWith("\") Then t.Text = t.Text.Substring(0, t.Text.Length - 1)
            End If
        Next
        doNotUpdate = False

        'HL/RL ini
        If FileExists(HLPathNoExt + "Settings\HyperLaunch.ini") Then
            ini.path = (HLPathNoExt + "\Settings\HyperLaunch.ini").Replace("\\", "\")
            ini.IniWriteValue("Settings", "HyperLaunch_Media_Path", TextBox2.Text)
            If TextBox3.Text.Trim <> "" Then ini.IniWriteValue("Settings", "Frontend_Path", TextBox3.Text)
            ini.IniWriteValue("HyperPause", "HyperPause_HiToText_Path", TextBox6.Text)
            ini.IniWriteValue("DAEMON Tools", "DAEMON_Tools_Path", TextBox7.Text)
        ElseIf FileExists(HLPathNoExt + "Settings\RocketLauncher.ini") Then
            ini.path = (HLPathNoExt + "Settings\RocketLauncher.ini").Replace("\\", "\")
            ini.IniWriteValue("Settings", "RocketLauncher_Media_Path", TextBox2.Text)
            If TextBox3.Text.Trim <> "" Then ini.IniWriteValue("Settings", "Default_Front_End_Path", TextBox3.Text)
            ini.IniWriteValue("Pause", "Pause_HiToText_Path", TextBox6.Text)
            ini.IniWriteValue("Virtual Drive", "Virtual_Drive_Path", TextBox7.Text)
        Else
            MsgBox("Error updating HL/RL ini")
            Exit Sub
        End If

        ini.IniWriteValue("Settings", "Modules_Path", TextBox1.Text)
        ini.IniWriteValue("Settings", "Profiles_Path", TextBox4.Text)
        ini.IniWriteValue("7z", "7z_Path", TextBox5.Text)
        ini.IniWriteValue("Keymapper", "Xpadder_Path", TextBox8.Text)
        ini.IniWriteValue("Keymapper", "JoyToKey_Path", TextBox9.Text)
        ini.IniWriteValue("Vjoy", "VJoy_Path", TextBox10.Text)

        'HLHQ ini
        If FileExists(HLHQPath + "\HyperLaunchHQ.exe") Then
            ini.path = (HLHQPath + "\Settings\HyperLaunchHQ.ini").Replace("\\", "\")
            If ini.IniReadValue("Paths", "HL_Folder").Trim <> "" Then
                ini.IniWriteValue("Paths", "HL_Folder", TextBox11.Text)
            ElseIf ini.IniReadValue("Settings", "HL_Folder").Trim <> "" Then
                ini.IniWriteValue("Settings", "HL_Folder", TextBox11.Text)
            Else
                ini.IniWriteValue("Paths", "HL_Folder", TextBox11.Text)
                ini.IniWriteValue("Settings", "HL_Folder", TextBox11.Text)
            End If

            If ini.IniReadValue("Paths", "Media_Folder_Path").Trim <> "" Then
                ini.IniWriteValue("Paths", "Media_Folder_Path", TextBox12.Text)
            ElseIf ini.IniReadValue("Settings", "Media_Folder_Path").Trim <> "" Then
                ini.IniWriteValue("Settings", "Media_Folder_Path", TextBox12.Text)
            Else
                ini.IniWriteValue("Paths", "Media_Folder_Path", TextBox12.Text)
                ini.IniWriteValue("Settings", "Media_Folder_Path", TextBox12.Text)
            End If

            Dim foundSection As String = "HS"
            ini.path = (HLHQPath + "\Settings\Frontends.ini").Replace("\\", "\")
            If ini.IniReadValue("Hyperspin", "Path").Trim <> "" Then
                foundSection = "Hyperspin"
                ini.IniWriteValue("Hyperspin", "Path", TextBox13.Text)
            ElseIf ini.IniReadValue("HS", "Path").Trim <> "" Then
                foundSection = "HS"
                ini.IniWriteValue("HS", "Path", TextBox13.Text)
            Else
                ini.IniWriteValue("Hyperspin", "Path", TextBox13.Text)
                ini.IniWriteValue("HS", "Path", TextBox13.Text)
            End If
            ini.IniWriteValue("HyperLaunchHQ", "Path", TextBox14.Text)

            If CheckBox2.Checked Then
                ini.IniWriteValue("Settings", "Default_Frontend", foundSection)
            ElseIf CheckBox3.Checked Then
                ini.IniWriteValue("Settings", "Default_Frontend", "HyperLaunchHQ")
            End If
        ElseIf FileExists(HLHQPath + "\RocketLauncherUI.exe") Then
            ini.path = (HLHQPath + "\Settings\RocketLauncherUI.ini").Replace("\\", "\")

            ini.IniWriteValue("Paths", "RL_Folder", TextBox11.Text)
            ini.IniWriteValue("Paths", "Media_Folder_Path", TextBox12.Text)

            ini.path = (HLHQPath + "\Settings\Frontends.ini").Replace("\\", "\")
            ini.IniWriteValue("Hyperspin", "Path", TextBox13.Text)
            ini.IniWriteValue("RocketLauncherUI", "Path", TextBox14.Text)

            If CheckBox2.Checked Then
                ini.IniWriteValue("Settings", "Default_Frontend", "HyperSpin")
            ElseIf CheckBox3.Checked Then
                ini.IniWriteValue("Settings", "Default_Frontend", "RocketLauncherUI")
            End If
        Else
            MsgBox("Error updating HLHQ/RL UI ini")
            Exit Sub
        End If
    End Sub

    'Default plugin change
    Private Sub CheckBox2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox2.CheckedChanged, CheckBox3.CheckedChanged
        Dim name As String = DirectCast(sender, CheckBox).Name.ToUpper
        If name = "CHECKBOX2" Then
            If CheckBox2.Checked Then CheckBox3.Checked = False Else CheckBox3.Checked = True
        End If
        If name = "CHECKBOX3" Then
            If CheckBox3.Checked Then CheckBox2.Checked = False Else CheckBox2.Checked = True
        End If
    End Sub
End Class