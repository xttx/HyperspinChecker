Imports System.Runtime.CompilerServices
Public Module Module1_Extensions
    Interface INI_File_Class_Interface
        Function IniReadValue(Section As String, Key As String) As String
        Sub IniWriteValue(Section As String, Key As String, Value As String)
        Function IniListKey(Optional section As String = Nothing) As String()
        Sub save()
    End Interface

    Public Class AnchorStylesExt
        Public Shared TopLeft As AnchorStyles = AnchorStyles.Top Or AnchorStyles.Left
        Public Shared TopRight As AnchorStyles = AnchorStyles.Top Or AnchorStyles.Right
        Public Shared BottomLeft As AnchorStyles = AnchorStyles.Bottom Or AnchorStyles.Left
        Public Shared BottomRight As AnchorStyles = AnchorStyles.Bottom Or AnchorStyles.Right
        Public Shared LeftRight As AnchorStyles = AnchorStyles.Left Or AnchorStyles.Right
        Public Shared TopBottom As AnchorStyles = AnchorStyles.Top Or AnchorStyles.Bottom
        Public Shared All As AnchorStyles = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
    End Class

    <Extension()>
    Public Function IndexOfCaseInsensitive(ByVal IEnum As IEnumerable(Of String), search As String) As Integer
        Dim ind = -1
        For i As Integer = 0 To IEnum.Count - 1
            If IEnum(i).ToUpper = search.ToUpper Then ind = i : Exit For
        Next
        Return ind
    End Function
    <Extension()>
    Public Function AllIndexesOf(ByVal IEnum As IEnumerable(Of String), search As String) As IEnumerable(Of Integer)

        Dim index = 0
        Dim l = IEnum.ToList
        Dim indexes As New List(Of Integer)
        Do
            index = l.IndexOf(search, index)
            If index = -1 Then Exit Do
            indexes.Add(index)
            index += 1
            If index > l.Count - 1 Then Exit Do
        Loop

        Return indexes
    End Function
    <Extension()>
    Public Function DeserializeToDictionary(ByVal str As String, separator_entry As String, separator_keyValue As String) As Dictionary(Of String, String)
        Dim dict As New Dictionary(Of String, String)
        If str Is Nothing Then Return dict
        For Each entry In str.Split({separator_entry}, StringSplitOptions.RemoveEmptyEntries)
            Dim kv_arr = entry.Split({separator_keyValue}, StringSplitOptions.None)
            If kv_arr.Count = 2 AndAlso kv_arr(0) <> "" Then dict(kv_arr(0)) = kv_arr(1)
        Next
        Return dict
    End Function

    'Validate a path
    Public Function IsPathValid(path As String) As Boolean
        Dim fi As IO.FileInfo = Nothing
        Try
            fi = New IO.FileInfo(path)
        Catch ex As Exception
            Return False
        End Try

        If fi IsNot Nothing Then Return True Else Return False
    End Function
    'Absolute path to relative
    Public Function Absolute_Path_to_Relative(path_from As String, path_to As String) As String
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
    'Relative path to absolute
    Public Function Relative_Path_to_Absolute(path_from As String, relative As String) As String
        If Not relative.StartsWith(".") Then Return relative
        Return (path_from + "\" + relative).Replace("\\", "\").Replace("\\", "\")
    End Function

    'Backup file or folder
    Public Sub Backup_FileOrFolder(path As String, Optional copy As Boolean = True, Optional backupFileToSubFolder As Boolean = False)
        Dim newPath = Backup_FileOrFolder_FindUnusedBackupName(path)
        If newPath.Trim = "" Then Exit Sub

        If IO.Directory.Exists(path) Then
            If copy Then
                My.Computer.FileSystem.CopyDirectory(path, newPath)
            Else
                IO.Directory.Move(path, newPath)
            End If
        ElseIf IO.File.Exists(path) Then
            If copy Then
                IO.File.Copy(path, newPath)
            Else
                IO.File.Move(path, newPath)
            End If
        End If
    End Sub
    Public Function Backup_FileOrFolder_FindUnusedBackupName(basePath As String) As String
        Dim c As Integer = 1
        If Not IO.Directory.Exists(basePath) AndAlso Not IO.File.Exists(basePath) Then Return ""

        Dim fName = IO.Path.GetFileName(basePath)
        Dim rx = New System.Text.RegularExpressions.Regex("\.backup\s\(\d+\)")
        rx.Replace(fName, "")
        basePath = IO.Path.GetDirectoryName(basePath) + "\" + fName

        Dim newPath = basePath + ".backup (" + c.ToString + ")"
        Do While IO.Directory.Exists(newPath) Or IO.File.Exists(newPath)
            c += 1
            newPath = basePath + ".backup (" + c.ToString + ")"
        Loop

        Return newPath
    End Function
End Module

Public Class CustomToolTip
    Inherits Form
    Const LINE_LIMIT = 100
    Const WS_EX_NOACTIVATE = &H8000000
    Dim binded As New Dictionary(Of Control, String)
    Dim binded_area As New Dictionary(Of Control, Rectangle)
    Dim l As New Label With {.AutoSize = False, .Dock = DockStyle.Fill, .Padding = New Padding(5, 5, 25, 5)}

    Public Sub New()
        Me.ShowInTaskbar = False
        Me.FormBorderStyle = FormBorderStyle.None
        Me.Width = 500 : Me.Height = 80
        Me.BackColor = SystemColors.Info

        Me.Controls.Add(l)
        l.ForeColor = SystemColors.InfoText
        'l.Text = "Tafgasdfasdf," + vbCrLf + "Test."
    End Sub

    Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
        Get
            Return True
            'Return MyBase.ShowWithoutActivation
        End Get
    End Property

    Sub Bind(c As Control, str As String, Optional area As Rectangle = Nothing, Optional noWarp As Boolean = False)
        Dim warpedStr = ""
        If Not noWarp Then
            'WordWarp for long strings
            Dim line = ""
            For Each word In str.Split(" "c)
                If (line + word).Length > LINE_LIMIT Then
                    warpedStr += line + vbCrLf : line = ""
                End If
                line += word + " "
            Next
            If line.Length > 0 Then warpedStr += line
        Else
            warpedStr = str
        End If


        If Not binded.ContainsKey(c) Then
            binded.Add(c, warpedStr)
            AddHandler c.MouseEnter, AddressOf showToolTip
            AddHandler c.MouseMove, AddressOf showToolTipMove
            AddHandler c.MouseLeave, AddressOf showToolTipHide
        Else
            binded(c) = warpedStr
        End If

        If area <> Nothing Then binded_area(c) = area
    End Sub
    Sub Bind(c As Control, str As String, noWarp As Boolean)
        Bind(c, str, Nothing, noWarp)
    End Sub

    Sub unBind(c As Control)
        If binded.ContainsKey(c) Then
            binded.Remove(c)
            RemoveHandler c.MouseEnter, AddressOf showToolTip
            RemoveHandler c.MouseMove, AddressOf showToolTipMove
            RemoveHandler c.MouseLeave, AddressOf showToolTipHide
        End If
        If binded_area.ContainsKey(c) Then binded_area.Remove(c)
    End Sub

    Sub unBindAll()
        For Each c In binded.Keys
            RemoveHandler c.MouseEnter, AddressOf showToolTip
            RemoveHandler c.MouseMove, AddressOf showToolTipMove
            RemoveHandler c.MouseLeave, AddressOf showToolTipHide
        Next
        binded.Clear()
        binded_area.Clear()
    End Sub

    Sub showToolTip(o As Object, e As EventArgs)
        Dim c = DirectCast(o, Control)
        l.Text = binded(c)

        Dim g = Graphics.FromImage(New Bitmap(1, 1))
        Me.Width = CInt(g.MeasureString(l.Text, l.Font).Width) + Me.Padding.Left + Me.Padding.Right + l.Padding.Left + l.Padding.Right
        Me.Height = CInt(g.MeasureString(l.Text, l.Font).Height) + Me.Padding.Top + Me.Padding.Bottom + l.Padding.Top + l.Padding.Bottom
        If l.Text.Contains(vbCrLf) Then Me.Height += 3
        If checkIfInArea(c) Then Me.Show()
    End Sub
    Sub showToolTipMove(o As Object, e As EventArgs)
        If checkIfInArea(DirectCast(o, Control)) Then
            Dim x = MousePosition.X
            Dim w = Screen.FromPoint(MousePosition).Bounds.Width
            If x + Me.Width <= w Then
                Me.Location = New Point(MousePosition.X, MousePosition.Y + 10)
            Else
                Me.Location = New Point(w - Me.Width, MousePosition.Y + 10)
            End If

            If Not Me.Visible Then Me.Show()
        Else
            If Me.Visible Then Me.Hide()
        End If
    End Sub
    Sub showToolTipHide(o As Object, e As EventArgs)
        Me.Hide()
    End Sub

    Function checkIfInArea(c As Control) As Boolean
        If binded_area.ContainsKey(c) Then
            Dim area = binded_area(c)
            Dim p = c.PointToClient(MousePosition)
            If p.X >= area.X AndAlso p.X <= c.Width - area.Width AndAlso p.Y >= area.Y AndAlso p.Y <= c.Height - area.Height Then
                Return True
            Else
                Return False
            End If
        Else
            Return True
        End If
    End Function

    'OBSOLETE
    'Protected Overrides ReadOnly Property CreateParams As CreateParams
    '    Get
    '        Dim ret = MyBase.CreateParams
    '        ret.ExStyle = ret.ExStyle Or WS_EX_NOACTIVATE
    '        Return ret
    '    End Get
    'End Property
End Class
Public Class AddSystemToHSMainMenu
    Dim panel As New Panel With {.BorderStyle = BorderStyle.Fixed3D, .Visible = False, .Anchor = AnchorStylesExt.All}
    Dim WithEvents listView As New ListView With {.Anchor = AnchorStylesExt.All, .FullRowSelect = True, .HideSelection = False, .MultiSelect = False, .View = View.Details}
    Dim WithEvents but_toTop As New Button With {.Text = "Set to top", .Anchor = AnchorStylesExt.BottomLeft}
    Dim WithEvents but_toBot As New Button With {.Text = "Set to bottom", .Anchor = AnchorStylesExt.BottomLeft}
    Dim WithEvents but_OK As New Button With {.Text = "OK", .Anchor = AnchorStylesExt.BottomRight}
    Dim WithEvents but_Cancel As New Button With {.Text = "Cancel", .Anchor = AnchorStylesExt.BottomRight}

    Dim sys As String = ""
    Public Event SystemAdded(index As Integer)
    Public Event SystemAddCanceled()

    Public Sub New()
        listView.Columns.Add("Where put new system on main wheel?")
        listView.Columns(0).Width = 300
        listView.Font = New Font(listView.Font.FontFamily, 12)
        but_OK.Font = New Font(but_OK.Font, FontStyle.Bold)
        but_Cancel.Font = New Font(but_Cancel.Font, FontStyle.Bold)

        panel.Size = New Size(420, 200)
        panel.Controls.Add(listView) : listView.Location = New Point(3, 4) : listView.Size = New Size(410, 152)
        panel.Controls.Add(but_toTop) : but_toTop.Location = New Point(3, 162) : but_toTop.Size = New Size(79, 30)
        panel.Controls.Add(but_toBot) : but_toBot.Location = New Point(88, 162) : but_toBot.Size = New Size(79, 30)
        panel.Controls.Add(but_OK) : but_OK.Location = New Point(279, 162) : but_OK.Size = New Size(64, 30)
        panel.Controls.Add(but_Cancel) : but_Cancel.Location = New Point(349, 162) : but_Cancel.Size = New Size(64, 30)
    End Sub

    Public Sub show(parent As Control, newSystem As String)
        sys = newSystem
        listView.Items.Clear()

        For Each r As DataRow In Form1.system_list.Rows
            listView.Items.Add(New ListViewItem(r.Item(0).ToString))
        Next
        listView.Items.Insert(0, New ListViewItem(sys))
        listView.TopItem = listView.Items(0)
        listView.Items(0).Selected = True

        parent.Controls.Add(panel)
        panel.Top = 10
        panel.Left = CInt(parent.Width / 2) - CInt(panel.Width / 2)
        panel.Height = parent.Height - 20
        panel.Visible = True : panel.BringToFront() : listView.Focus()
    End Sub

    'Main menu systems list selection change
    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles listView.SelectedIndexChanged
        If listView.SelectedItems.Count <> 1 Then Exit Sub
        Dim c As Integer = 0
        Dim f As Font = listView.Font
        Dim f_bold As Font = New Font(f.FontFamily, f.SizeInPoints + 4, FontStyle.Bold)
        For Each l As ListViewItem In listView.Items
            If l.Selected Then
                l.Text = sys
                l.Font = f_bold
            Else
                l.Text = Form1.system_list.Rows(c).Item(0).ToString
                l.Font = f
                c += 1
            End If
        Next
    End Sub
    'Set to top
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles but_toTop.Click
        If listView.SelectedItems.Count <> 1 Then Exit Sub

        Dim c As Integer = 0
        Dim f As Font = listView.Font
        Dim f_bold As Font = New Font(f.FontFamily, f.SizeInPoints + 4, FontStyle.Bold)
        For i As Integer = 0 To listView.Items.Count - 1
            Dim l As ListViewItem = listView.Items(i)
            If i = 0 Then
                l.Text = sys
                l.Font = f_bold
            Else
                l.Text = Form1.system_list.Rows(c).Item(0).ToString
                l.Font = f
                c += 1
            End If
        Next
        listView.TopItem = listView.Items(0)
        listView.Items(0).Selected = True
        listView.Focus()
    End Sub
    'Set to bottom
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles but_toBot.Click
        If listView.SelectedItems.Count <> 1 Then Exit Sub

        Dim c As Integer = 0
        Dim f As Font = listView.Font
        Dim f_bold As Font = New Font(f.FontFamily, f.SizeInPoints + 4, FontStyle.Bold)
        For i As Integer = 0 To listView.Items.Count - 1
            Dim l As ListViewItem = listView.Items(i)
            If i = listView.Items.Count - 1 Then
                l.Text = sys
                l.Font = f_bold
            Else
                l.Text = Form1.system_list.Rows(c).Item(0).ToString
                l.Font = f
                c += 1
            End If
        Next
        listView.TopItem = listView.Items(listView.Items.Count - 1)
        listView.Items(listView.Items.Count - 1).Selected = True
        listView.Focus()
    End Sub
    'OK
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles but_OK.Click
        'Add to main menu
        If listView.SelectedIndices.Count < 1 Then Exit Sub
        Dim x As New Xml.XmlDocument
        x.Load(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml")

        Dim xm = x.SelectSingleNode("/menu")
        'If xm Is Nothing Then xm = x.CreateElement("menu") : x.AppendChild(xm) 'Am I paranoid?
        Dim xg = x.SelectNodes("/menu/game")

        Dim xs = x.CreateElement("game")
        xs.SetAttribute("name", sys)
        xm.InsertBefore(xs, xg(listView.SelectedIndices(0)))

        Dim row = Form1.system_list.NewRow() : row.Item(0) = sys
        Form1.system_list.Rows.InsertAt(row, listView.SelectedIndices(0))
        Form1.ComboBox1.Items.Insert(listView.SelectedIndices(0), sys)

        x.Save(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml")

        panel.Parent.Controls.Remove(panel)
        RaiseEvent SystemAdded(listView.SelectedIndices(0))
    End Sub
    'Cancel
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles but_Cancel.Click
        panel.Parent.Controls.Remove(panel)
        RaiseEvent SystemAddCanceled()
    End Sub
End Class
