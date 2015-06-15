Imports Microsoft.VisualBasic.FileIO.FileSystem
Public Class Form9_database_statistic
    Private aliases As New List(Of String)
    Private excludeList As New List(Of String)
    Private frm As Form9_database_statistic_comparer

    Private Sub Form9_database_statistic_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = Class1.HyperspinPath + "\Databases"
        TextBox2.Text = "MainDB"
        LoadExcludeList() 'Reload exclude list
    End Sub 'Form Load

    Private Sub LoadExcludeList()
        excludeList.Clear()
        If FileExists(".\DBExcludeList.txt") Then
            FileOpen(1, ".\DBExcludeList.txt", OpenMode.Input)
            Dim s As String = ""
            Do While Not EOF(1)
                s = LineInput(1).Trim.ToUpper
                If Not excludeList.Contains(s) Then excludeList.Add(s)
            Loop
            FileClose(1)
        End If
    End Sub 'Load Exclude List

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        aliases.Clear()
        ListBox1.Items.Clear()
        Label3.Text = "Path: "
        Label4.Text = "Total Games: "
    End Sub 'Clear all

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Not DirectoryExists(TextBox1.Text) Then MsgBox("Path not Exist") : Exit Sub
        Dim al As String = TextBox2.Text.Trim.ToUpper
        If aliases.Contains(al) Then MsgBox("This alias already exist") : Exit Sub
        aliases.Add(al)
        Dim c As Integer = 0
        Dim xml As New Xml.XmlDocument
        For Each f As String In GetFiles(TextBox1.Text, FileIO.SearchOption.SearchAllSubDirectories, {"*.xml"})
            Dim i As String = f.Substring(f.LastIndexOf("\") + 1) : i = i.Substring(0, i.LastIndexOf("."))
            If excludeList.Contains(i.Trim.ToUpper) Then Continue For

            Try
                xml.Load(f)
                c = xml.SelectNodes("/menu/game").Count
            Catch ex As Exception
                c = -1
            End Try
            ListBox1.Items.Add(New myItem(i, f, TextBox2.Text, c))
        Next
    End Sub 'SCAN

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedIndex < 0 Then Label3.Text = "Path: " : Label4.Text = "Total Games: " : Exit Sub
        Dim l As myItem = DirectCast(ListBox1.SelectedItem, myItem)
        Label3.Text = "Path: " + l.Tag1
        Label4.Text = "Total Games: " + l.count.ToString
        If l.count = -1 Then Label4.Text = "Total Games: INVALID XML"
    End Sub 'List selection change

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ListBox1.SelectedIndex < 0 Then Exit Sub
        Dim s As String = DirectCast(ListBox1.SelectedItem, myItem).Text.Trim.ToUpper
        If Not excludeList.Contains(s) Then
            excludeList.Add(s)
            FileOpen(1, ".\DBExcludeList.txt", OpenMode.Output)
            For Each s In excludeList
                PrintLine(1, s)
            Next
            FileClose(1)
            Button4_Click_RefreshList()
        End If
    End Sub 'Put to exclude

    Private Sub Button4_Click_RefreshList()
        Dim itm As String = ""
        Dim indexes_to_remove As New List(Of Integer)
        For c As Integer = ListBox1.Items.Count - 1 To 0 Step -1
            itm = DirectCast(ListBox1.Items(c), myItem).Text
            If excludeList.Contains(itm.Trim.ToUpper) Then indexes_to_remove.Add(c)
        Next
        For Each c As Integer In indexes_to_remove
            ListBox1.Items.RemoveAt(c)
        Next
    End Sub

    Dim WithEvents p As Process
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Not FileExists(".\DBExcludeList.txt") Then MsgBox("List is empty.") : Exit Sub
        p = System.Diagnostics.Process.Start("notepad.exe", ".\DBExcludeList.txt")
        AddHandler p.Exited, AddressOf Button5_Click_NotepadExited
    End Sub 'Edit Exclude List

    Private Sub Button5_Click_NotepadExited(sender As Object, e As System.EventArgs) Handles p.Exited
        LoadExcludeList()
        Button4_Click_RefreshList()
        MsgBox("asd")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If ListBox1.SelectedItems.Count < 2 Then MsgBox("You have to select at least 2 databases to compare.") : Exit Sub
        If ListBox1.SelectedItems.Count > 3 Then MsgBox("You can compare up to 3 databases.") : Exit Sub
        If DirectCast(ListBox1.SelectedItems(0), myItem).count = -1 Then MsgBox("One of selected XMLs is invalid.") : Exit Sub
        If DirectCast(ListBox1.SelectedItems(1), myItem).count = -1 Then MsgBox("One of selected XMLs is invalid.") : Exit Sub
        If ListBox1.SelectedItems.Count = 3 AndAlso DirectCast(ListBox1.SelectedItems(2), myItem).count = -1 Then MsgBox("One of selected XMLs is invalid.") : Exit Sub

        frm = New Form9_database_statistic_comparer
        frm.Width = 850
        frm.file1 = DirectCast(ListBox1.SelectedItems(0), myItem).Tag1
        frm.file2 = DirectCast(ListBox1.SelectedItems(1), myItem).Tag1
        If ListBox1.SelectedItems.Count = 3 Then frm.Width = 1250 : frm.file3 = DirectCast(ListBox1.SelectedItems(2), myItem).Tag1
        frm.Show()
    End Sub 'Show comparator
End Class

Public Class myItem
    Public Tag1 As String
    Public Tag2 As String
    Public Text As String
    Public count As Integer
    Public Shared switch As Integer = 0

    Public Sub New(txt As String, tg As String, tg2 As String, c As Integer)
        Text = txt : Tag1 = tg : Tag2 = tg2 : count = c
    End Sub

    Public Overrides Function ToString() As String
        Return Me.Text + " (" + Tag2 + ")"
    End Function
End Class
