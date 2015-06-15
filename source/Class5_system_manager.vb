Public Class Class5_system_manager
    Dim ini As New IniFileApi
    Dim modules As New Dictionary(Of String, List(Of String))
    Private frm As New Form8_systemProperties
    Private Systems As New Dictionary(Of String, List(Of String))
    Private WithEvents Btn_scan As Button = Form1.Button33
    Private WithEvents Btn_prop As Button = Form1.Button34
    Private WithEvents grid As DataGridView = Form1.DataGridView2

    Public Sub New()
        AddHandler Form8_systemProperties.paths_updated, AddressOf SystemPathUpdated
    End Sub

    'Main Scan Systems
    Private Sub scan() Handles Btn_scan.Click
        If Form1.Label23.BackColor <> Color.LightGreen Then MsgBox("HyperSpin path is incorrect or not set. Set HS path in 'Program Settings' tab.") : Exit Sub
        Systems.Clear()
        modules.Clear()
        Form1.DataGridView2.Rows.Clear()
        Dim iniClass As New IniFile, iniclass2 As New IniFile
        Dim activeCount As Integer = 0

        iniClass.Load(Class1.HyperspinPath + "\Settings\Settings.ini")
        Dim HL_Path As String = iniClass.GetKeyValue("Main", "Hyperlaunch_Path").Trim
        If Not HL_Path = "" AndAlso Not FileIO.FileSystem.FileExists(HL_Path + "\HyperLaunch.exe") Then HL_Path = ""
        If Not HL_Path = "" Then iniclass2.Load(HL_Path + "\Settings\Global Emulators.ini")

        'SCANING MODULES
        If Not HL_Path = "" Then
            Dim s As String = ""
            Dim modules_files As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = FileIO.FileSystem.GetFiles(HL_Path + "\Modules\", FileIO.SearchOption.SearchAllSubDirectories, {"*.ahk"})
            For Each module_file As String In modules_files
                s = ""
                FileOpen(1, module_file, OpenMode.Input, OpenAccess.Read)
                Do While Not EOF(1)
                    s = LineInput(1)
                    If s.Trim.ToUpper.StartsWith("MSYSTEM") Then
                        s = s.Substring(s.IndexOf("=") + 1).Trim
                        Dim l As New List(Of String)
                        l.Add(module_file)
                        For Each supportedSys In s.Split({","}, StringSplitOptions.RemoveEmptyEntries)
                            supportedSys = supportedSys.Replace("""", "").Trim
                            l.Add(supportedSys.ToUpper)
                        Next
                        modules.Add(module_file.Substring(module_file.LastIndexOf("\") + 1), l)
                    End If
                Loop
                FileClose(1)
            Next
        End If


        Dim mainMenuXml As New Xml.XmlDocument
        mainMenuXml.Load(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml")
        activeCount = mainMenuXml.SelectNodes("/menu/game").Count
        For Each node As Xml.XmlNode In mainMenuXml.SelectNodes("/menu/game")
            Systems.Add(node.Attributes("name").Value.ToUpper, New List(Of String))
            Systems(node.Attributes("name").Value.ToUpper).Add(node.Attributes("name").Value)
            Systems(node.Attributes("name").Value.ToUpper).Add("%1%")
        Next

        Dim sys As String = ""
        For Each Dir As String In FileIO.FileSystem.GetDirectories(Class1.HyperspinPath + "\Databases\")
            sys = Dir.Substring(Dir.LastIndexOf("\") + 1)
            If FileIO.FileSystem.FileExists(Dir + "\" + sys + ".xml") And Not sys.ToUpper = "MAIN MENU" Then
                If Not Systems.Keys.Contains(sys.ToUpper) Then
                    Systems.Add(sys.ToUpper, New List(Of String))
                    Systems(sys.ToUpper).Add(sys)
                End If
                Systems(sys.ToUpper).Add("%2%")
            End If
        Next

        Dim files As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = FileIO.FileSystem.GetFiles(Class1.HyperspinPath + "\Media\Main Menu\Themes\", FileIO.SearchOption.SearchTopLevelOnly, {"*.zip"})
        For Each file As String In files
            sys = file.Substring(file.LastIndexOf("\") + 1)
            sys = sys.Substring(0, sys.LastIndexOf("."))
            If Not Systems.Keys.Contains(sys.ToUpper) Then
                Systems.Add(sys.ToUpper, New List(Of String))
                Systems(sys.ToUpper).Add(sys)
            End If
            Systems(sys.ToUpper).Add("%3%")
        Next

        For Each Dir As String In FileIO.FileSystem.GetDirectories(Class1.HyperspinPath + "\Media\")
            sys = Dir.Substring(Dir.LastIndexOf("\") + 1)
            If FileIO.FileSystem.FileExists(Dir + "\Themes\default.zip") And Not sys.ToUpper = "MAIN MENU" Then
                If Not Systems.Keys.Contains(sys.ToUpper) Then
                    Systems.Add(sys.ToUpper, New List(Of String))
                    Systems(sys.ToUpper).Add(sys)
                End If
                Systems(sys.ToUpper).Add("%4%")
            End If
        Next

        Dim c As Integer = 0
        Dim emuPath As String = ""
        Dim romPath As String = ""
        For Each file As String In FileIO.FileSystem.GetFiles(Class1.HyperspinPath + "\Settings\", FileIO.SearchOption.SearchTopLevelOnly, {"*.ini"})
            c += 1 : emuPath = ""
            sys = file.Substring(file.LastIndexOf("\") + 1)
            sys = sys.Substring(0, sys.LastIndexOf("."))
            If Not sys.ToUpper = "SETTINGS" And Not sys.ToUpper = "BETABRITE" And Not sys.ToUpper = "MAIN MENU" And Not sys.ToUpper = "GLOBAL SETTINGS" Then
                If Not Systems.Keys.Contains(sys.ToUpper) Then
                    Systems.Add(sys.ToUpper, New List(Of String))
                    Systems(sys.ToUpper).Add(sys)
                End If
                Systems(sys.ToUpper).Add("%5%")

                iniClass.Load(file)
                If iniClass.GetKeyValue("EXE INFO", "hyperlaunch").Trim.ToUpper = "TRUE" Then
                    'HYPERLAUNCH INI
                    Systems(sys.ToUpper).Add("%5HL%")
                    If HL_Path = "" Then
                        Systems(sys.ToUpper).Add("%5HL_PATH_NOT_SET%")
                    Else
                        Dim emu As String = ""
                        If FileIO.FileSystem.FileExists(HL_Path + "\Settings\" + sys + "\Emulators.ini") Then
                            iniClass.Load(HL_Path + "\Settings\" + sys + "\Emulators.ini")
                            emu = iniClass.GetKeyValue("ROMS", "Default_Emulator").Trim
                            romPath = iniClass.GetKeyValue("ROMS", "Rom_Path").Trim
                            If romPath.StartsWith(".") Then romPath = HL_Path + "\" + romPath

                            If emu <> "" Then
                                'iniClass.Load(HL_Path + "\Settings\Global Emulators.ini")
                                emuPath = iniclass2.GetKeyValue(emu, "Emu_Path").Trim
                                If emuPath.StartsWith(".") Then emuPath = HL_Path + "\" + emuPath
                            End If
                        End If
                    End If
                Else
                    'HYPERSPIN INI
                    emuPath = iniClass.GetKeyValue("EXE INFO", "path").Trim
                    If emuPath <> "" AndAlso Not emuPath.EndsWith("\") Then emuPath = emuPath + "\"
                    If emuPath <> "" Then emuPath = emuPath + iniClass.GetKeyValue("EXE INFO", "exe").Trim
                    If emuPath <> "" And emuPath.StartsWith(".") Then emuPath = Class1.HyperspinPath + "\" + emuPath

                    romPath = iniClass.GetKeyValue("EXE INFO", "rompath").Trim
                    If romPath <> "" And romPath.StartsWith(".") Then romPath = Class1.HyperspinPath + "\" + romPath
                End If
                If FileIO.FileSystem.FileExists(emuPath) Then Systems(sys.ToUpper).Add("%6%")
                If FileIO.FileSystem.DirectoryExists(romPath) Then Systems(sys.ToUpper).Add("%7%")
            End If

            Form1.Label41.Text = "Scanning... " + c.ToString + " of " + files.Count.ToString : Form1.Label41.Refresh()
        Next

        'FILL GRID
        Dim x(7) As String
        For Each l As List(Of String) In Systems.Values
            x(0) = l(0)
            If l.Contains("%1%") Then x(1) = "X" Else x(1) = ""
            If l.Contains("%2%") Then x(2) = "X" Else x(2) = ""
            If l.Contains("%3%") Then x(3) = "X" Else x(3) = ""
            If l.Contains("%4%") Then x(4) = "X" Else x(4) = ""
            If l.Contains("%5%") Then x(5) = "X" Else x(5) = ""
            If l.Contains("%5HL%") Then x(5) = "HL"
            If l.Contains("%5HL_PATH_NOT_SET%") Then x(5) = "INVALID HL PATH"
            If l.Contains("%6%") Then x(6) = "X" Else x(6) = ""
            If l.Contains("%7%") Then x(7) = "X" Else x(7) = ""
            Dim r As New DataGridViewRow()
            r.CreateCells(grid, x)
            For i As Integer = 1 To 7
                If x(i) = "X" Or x(i) = "HL" Then
                    r.Cells(i).Style.BackColor = Form1.colorYES
                Else
                    r.Cells(i).Style.BackColor = Form1.colorNO

                    If x(i) = "INVALID HL PATH" Then
                        Dim f As Font = New Font(Control.DefaultFont.FontFamily, 6, FontStyle.Regular)
                        r.Cells(i).Style.Font = f
                    End If
                End If
            Next
            Form1.DataGridView2.Rows.Add(r)
        Next
        Form1.Label41.Text = "System count: " + Form1.DataGridView2.Rows.Count.ToString + ", Active:" + activeCount.ToString
    End Sub

#Region "Drag'n'drop"
    Private Sub drag_enter(sender As Object, e As Windows.Forms.DragEventArgs) Handles grid.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Move
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect
            grid.DefaultCellStyle.SelectionBackColor = Color.LightGoldenrodYellow
        End If
    End Sub

    Private Sub drag_over(sender As Object, e As Windows.Forms.DragEventArgs) Handles grid.DragOver
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            Dim clientPoint As Point = grid.PointToClient(New Point(e.X, e.Y))
            Dim hitTest As DataGridView.HitTestInfo = grid.HitTest(clientPoint.X, clientPoint.Y)
            If hitTest.RowIndex >= 0 And hitTest.ColumnIndex >= 0 Then
                grid.Rows(hitTest.RowIndex).Cells(hitTest.ColumnIndex).Selected = True
            End If
        End If
    End Sub

    Private Sub drag_drop(sender As Object, e As System.EventArgs) Handles grid.DragLeave
        Dim r As Integer = -1
        grid.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight
        If grid.SelectedCells.Count > 0 Then r = grid.SelectedCells(0).RowIndex
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        If r >= 0 Then grid.Rows(r).Selected = True
    End Sub

    Private Sub drag_drop(sender As Object, e As Windows.Forms.DragEventArgs) Handles grid.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            grid.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight
            Dim clientPoint As Point = grid.PointToClient(New Point(e.X, e.Y))
            Dim hitTest As DataGridView.HitTestInfo = grid.HitTest(clientPoint.X, clientPoint.Y)

            Dim failed As New List(Of String)
            Dim files As String() = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
            If files.Count > 0 Then
                Select Case hitTest.ColumnIndex
                    Case 2 'Database
                        For Each file As String In files
                            If file.ToUpper.EndsWith(".XML") Then
                                Dim sys As String = file.ToUpper.Replace(".XML", "")
                                If Systems.Keys.Contains(sys) Then
                                    FileCopy(file, Class1.HyperspinPath + "\Databases\" + Systems(sys)(0) + "\" + Systems(sys)(0) + ".xml")
                                Else
                                    failed.Add(file)
                                End If
                            End If

                            If failed.Count = 1 Then
                                Dim sys As String = grid.Rows(hitTest.RowIndex).Cells(0).Value.ToString
                                FileCopy(failed(0), Class1.HyperspinPath + "\Databases\" + sys + "\" + sys + ".xml")
                                grid.Rows(hitTest.RowIndex).Cells(4).Style.BackColor = Form1.colorYES
                                If grid.Rows(hitTest.RowIndex).Cells(4).Value.ToString = "" Then
                                    grid.Rows(hitTest.RowIndex).Cells(4).Value = "X"
                                Else
                                    grid.Rows(hitTest.RowIndex).Cells(4).Value = "NEW"
                                End If
                            End If
                        Next
                    Case 3 'Main menu theme
                        For Each file As String In files
                            If file.ToUpper.EndsWith(".ZIP") Then
                                Dim sys As String = file.ToUpper.Replace(".ZIP", "")
                                If Systems.Keys.Contains(sys) Then
                                    FileCopy(file, Class1.HyperspinPath + "\Media\Main Menu\Themes\" + Systems(sys)(0) + ".zip")
                                Else
                                    failed.Add(file)
                                End If
                            End If
                        Next

                        If failed.Count = 1 Then
                            Dim sys As String = grid.Rows(hitTest.RowIndex).Cells(0).Value.ToString
                            Dim res As Microsoft.VisualBasic.MsgBoxResult = MsgBox("Following file were not recognized as system: " + failed(0).Substring(failed(0).LastIndexOf("\")) + vbCrLf + "Do you want to set it as " + sys + "?", MsgBoxStyle.YesNo)
                            If res = MsgBoxResult.Yes Then
                                FileCopy(failed(0), Class1.HyperspinPath + "\Media\Main Menu\Themes\" + sys + ".zip")
                                grid.Rows(hitTest.RowIndex).Cells(3).Style.BackColor = Form1.colorYES
                                If grid.Rows(hitTest.RowIndex).Cells(3).Value.ToString = "" Then
                                    grid.Rows(hitTest.RowIndex).Cells(3).Value = "X"
                                Else
                                    grid.Rows(hitTest.RowIndex).Cells(3).Value = "NEW"
                                End If
                            End If
                        Else
                            MsgBox("Following files were not recognized as system: " + vbCrLf + String.Join(vbCrLf, failed))
                        End If
                    Case 4 'System Theme
                        For Each file As String In files
                            If file.ToUpper.EndsWith(".ZIP") Then
                                Dim sys As String = file.ToUpper.Replace(".ZIP", "")
                                If Systems.Keys.Contains(sys) Then
                                    FileCopy(file, Class1.HyperspinPath + "\Media\" + Systems(sys)(0) + "\Themes\default.zip")
                                Else
                                    failed.Add(file)
                                End If
                            End If

                            If failed.Count = 1 Then
                                Dim sys As String = grid.Rows(hitTest.RowIndex).Cells(0).Value.ToString.ToUpper
                                FileCopy(failed(0), Class1.HyperspinPath + "\Media\" + Systems(sys)(0) + "\Themes\default.zip")
                                grid.Rows(hitTest.RowIndex).Cells(4).Style.BackColor = Form1.colorYES
                                If grid.Rows(hitTest.RowIndex).Cells(4).Value.ToString = "" Then
                                    grid.Rows(hitTest.RowIndex).Cells(4).Value = "X"
                                Else
                                    grid.Rows(hitTest.RowIndex).Cells(4).Value = "NEW"
                                End If
                            End If
                        Next
                End Select
            End If

            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            If hitTest.RowIndex >= 0 And hitTest.ColumnIndex >= 0 Then grid.Rows(hitTest.RowIndex).Selected = True
        End If
    End Sub
#End Region

    'Show properties
    Private Sub show_properties() Handles Btn_prop.Click, grid.CellDoubleClick
        frm = New Form8_systemProperties
        setPropertiesFormData() : frm.Show()
    End Sub

    'Grid selection changed
    Private Sub grid_SelectionChanged(sender As Object, e As EventArgs) Handles grid.SelectionChanged
        If grid.SelectionMode <> DataGridViewSelectionMode.FullRowSelect Then Exit Sub
        If frm.Visible = False Then Exit Sub
        setPropertiesFormData()
    End Sub

    'setPropertiesFormData()
    Private Sub setPropertiesFormData()
        Dim data = New List(Of String)
        Dim modules_for_cur_system As New List(Of String)

        If grid.SelectedRows.Count > 0 Then
            Dim sys = grid.SelectedRows(0).Cells(0).Value.ToString.ToUpper
            If Systems(sys).Contains("%1%") Then data.Add("1") Else data.Add("") 'Main menu active
            If Systems(sys).Contains("%2%") Then data.Add("1") Else data.Add("") 'Database exist
            If Systems(sys).Contains("%3%") Then data.Add("1") Else data.Add("") 'Main Menu Theme Exist
            If Systems(sys).Contains("%4%") Then data.Add("1") Else data.Add("") 'System Theme Exist
            If Systems(sys).Contains("%5%") Then data.Add("1") Else data.Add("") 'HS Settings Exist
            If Systems(sys).Contains("%5HL%") Then data.Add("1") Else data.Add("") 'HS Set to use HyperLaunch for this system
            If Systems(sys).Contains("%5HL_PATH_NOT_SET%") Then data.Add("1") Else data.Add("") 'HL path is invalid
            If Systems(sys).Contains("%6%") Then data.Add("1") Else data.Add("") 'Emulator exists
            If Systems(sys).Contains("%7%") Then data.Add("1") Else data.Add("") 'Rom Path exists
            data.Add(sys)

            'SET modules
            For Each m As KeyValuePair(Of String, List(Of String)) In modules
                If m.Value.Contains(sys) Then
                    modules_for_cur_system.Add(m.Key)
                End If
            Next
        End If
        frm.data = data
        frm.modules = modules_for_cur_system
    End Sub

    'System path Updated Event from form_properties
    Private Sub SystemPathUpdated(sys As String)
        Dim system_grid_index As Integer = -1
        For i As Integer = 0 To grid.Rows.Count - 1
            If grid.Rows(i).Cells(0).Value.ToString = sys Then system_grid_index = i : Exit For
        Next
        If system_grid_index = -1 Then MsgBox("Can't find system '" + sys + "' in the grid.") : Exit Sub
        Dim r As DataGridViewRow = grid.Rows(system_grid_index)

        Dim iniClass As New IniFile, iniclass2 As New IniFile
        iniClass.Load(Class1.HyperspinPath + "\Settings\Settings.ini")
        Dim HL_Path As String = iniClass.GetKeyValue("Main", "Hyperlaunch_Path").Trim
        If Not HL_Path = "" AndAlso Not FileIO.FileSystem.FileExists(HL_Path + "\HyperLaunch.exe") Then HL_Path = ""
        If Not HL_Path = "" Then iniclass2.Load(HL_Path + "\Settings\Global Emulators.ini")

        If Not FileIO.FileSystem.FileExists(Class1.HyperspinPath + "\Settings\" + sys + ".ini") Then
            Dim f As Font = r.Cells(4).Style.Font
            r.Cells(5).Value = "INVALID HL PATH"
            r.Cells(6).Value = ""
            r.Cells(7).Value = ""
            r.Cells(5).Style.Font = f
            r.Cells(5).Style.BackColor = Form1.colorNO
            r.Cells(6).Style.BackColor = Form1.colorNO
            r.Cells(7).Style.BackColor = Form1.colorNO
            Exit Sub
        End If
        iniClass.Load(Class1.HyperspinPath + "\Settings\" + sys + ".ini")

        If iniClass.GetKeyValue("EXE INFO", "hyperlaunch").Trim.ToUpper = "TRUE" Then
            'Check Hyperlaunch paths
            If HL_Path = "" Then
                Dim f As Font = New Font(Control.DefaultFont.FontFamily, 6, FontStyle.Regular)
                r.Cells(5).Value = "INVALID HL PATH"
                r.Cells(6).Value = ""
                r.Cells(7).Value = ""
                r.Cells(5).Style.Font = f
                r.Cells(5).Style.BackColor = Form1.colorNO
                r.Cells(6).Style.BackColor = Form1.colorNO
                r.Cells(7).Style.BackColor = Form1.colorNO
            Else
                Dim f As Font = r.Cells(4).Style.Font
                r.Cells(5).Value = "HL"
                r.Cells(5).Style.Font = f

                Dim emu As String = ""
                Dim emuPath As String = ""
                Dim romPath As String = ""
                If FileIO.FileSystem.FileExists(HL_Path + "\Settings\" + sys + "\Emulators.ini") Then
                    iniClass.Load(HL_Path + "\Settings\" + sys + "\Emulators.ini")
                    emu = iniClass.GetKeyValue("ROMS", "Default_Emulator").Trim
                    romPath = iniClass.GetKeyValue("ROMS", "Rom_Path").Trim
                    If romPath.StartsWith(".") Then romPath = HL_Path + "\" + romPath

                    If emu <> "" Then
                        emuPath = iniclass2.GetKeyValue(emu, "Emu_Path").Trim
                        If emuPath.StartsWith(".") Then emuPath = HL_Path + "\" + emuPath
                    End If

                    If FileIO.FileSystem.FileExists(emuPath) Then r.Cells(6).Value = "X" : r.Cells(6).Style.BackColor = Form1.colorYES Else r.Cells(6).Value = "" : r.Cells(6).Style.BackColor = Form1.colorNO
                    If FileIO.FileSystem.DirectoryExists(romPath) Then r.Cells(7).Value = "X" : r.Cells(7).Style.BackColor = Form1.colorYES Else r.Cells(7).Value = "" : r.Cells(7).Style.BackColor = Form1.colorNO
                Else
                    r.Cells(6).Value = ""
                    r.Cells(7).Value = ""
                    r.Cells(6).Style.BackColor = Form1.colorNO
                    r.Cells(7).Style.BackColor = Form1.colorNO
                End If
            End If
        Else
            'Check Hyperspin paths
            Dim f As Font = r.Cells(4).Style.Font
            r.Cells(5).Value = "X"
            r.Cells(5).Style.Font = f

            Dim emuPath As String = iniClass.GetKeyValue("EXE INFO", "path").Trim
            If emuPath <> "" AndAlso Not emuPath.EndsWith("\") Then emuPath = emuPath + "\"
            If emuPath <> "" Then emuPath = emuPath + iniClass.GetKeyValue("EXE INFO", "exe").Trim
            If emuPath <> "" And emuPath.StartsWith(".") Then emuPath = Class1.HyperspinPath + "\" + emuPath

            Dim romPath As String = iniClass.GetKeyValue("EXE INFO", "rompath").Trim
            If romPath <> "" And romPath.StartsWith(".") Then romPath = Class1.HyperspinPath + "\" + romPath

            If FileIO.FileSystem.FileExists(emuPath) Then r.Cells(6).Value = "X" : r.Cells(6).Style.BackColor = Form1.colorYES Else r.Cells(6).Value = "" : r.Cells(6).Style.BackColor = Form1.colorNO
            If FileIO.FileSystem.DirectoryExists(romPath) Then r.Cells(7).Value = "X" : r.Cells(7).Style.BackColor = Form1.colorYES Else r.Cells(7).Value = "" : r.Cells(7).Style.BackColor = Form1.colorNO
        End If
    End Sub
End Class
