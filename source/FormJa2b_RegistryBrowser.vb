Imports Microsoft.Win32
Public Class FormJa2b_RegistryBrowser
    Dim HKCR As RegistryKey = Registry.ClassesRoot
    Dim HKCU As RegistryKey = Registry.CurrentUser
    Dim HKLM As RegistryKey = Registry.LocalMachine
    Dim HKCC As RegistryKey = Registry.CurrentConfig

    Public return_result As String = ""

    'OnLoad is called every time we show a form. The constructor should only be called once
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim t As TreeNode
        t = TreeView1.Nodes.Add("HKCR", "HKEY_CLASSES_ROOT")
        t.Nodes.Add("dummy", "dummy")
        t = TreeView1.Nodes.Add("HKCU", "HKEY_CURRENT_USER")
        t.Nodes.Add("dummy", "dummy")
        t = TreeView1.Nodes.Add("HKLM", "HKEY_LOCAL_MACHINE")
        t.Nodes.Add("dummy", "dummy")
        t = TreeView1.Nodes.Add("HKCC", "HKEY_CURRENT_CONFIG")
        t.Nodes.Add("dummy", "dummy")
    End Sub

    Private Sub FormJa2b_RegistryBrowser_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        TextBox1.Text = ""
        return_result = ""
    End Sub

    Private Sub TreeView1_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles TreeView1.BeforeExpand
        If e.Node.Nodes.Count = 0 Then Exit Sub
        If e.Node.Nodes(0).Text <> "dummy" Then Exit Sub

        e.Node.Nodes.RemoveAt(0)
        Try
            Dim k As RegistryKey = getKeyFromNode(e.Node)

            For Each sub_k In k.GetSubKeyNames
                Dim tmp_k = k.OpenSubKey(sub_k)
                Dim sub_node = e.Node.Nodes.Add(sub_k, sub_k)
                If tmp_k.SubKeyCount > 0 Then sub_node.Nodes.Add("dummy", "dummy")
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        ListView1.Items.Clear()
        If e.Node.Parent Is Nothing Then Exit Sub

        Try
            Dim k As RegistryKey = getKeyFromNode(e.Node)
            For Each v In k.GetValueNames
                Dim l As New ListViewItem(v)
                l.SubItems.Add(k.GetValue(v).ToString)

                If l.Text.Trim = "" Then l.Text = "@default"
                ListView1.Items.Add(l)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Function getKeyFromNode(n As TreeNode) As RegistryKey
        Dim path = n.FullPath

        Dim k As RegistryKey = Nothing
        If path.StartsWith("HKEY_CLASSES_ROOT") Then
            k = HKCR
        ElseIf path.StartsWith("HKEY_CURRENT_USER") Then
            k = HKCU
        ElseIf path.StartsWith("HKEY_LOCAL_MACHINE") Then
            k = HKLM
        ElseIf path.StartsWith("HKEY_CURRENT_CONFIG") Then
            k = HKCC
        End If

        If n.Parent IsNot Nothing Then
            k = k.OpenSubKey(path.Substring(path.IndexOf("\") + 1))
        End If

        Return k
    End Function

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        TextBox1.Text = ""
        If TreeView1.SelectedNode Is Nothing Then Exit Sub
        If ListView1.SelectedItems.Count <> 1 Then Exit Sub
        TextBox1.Text = TreeView1.SelectedNode.FullPath + "\" + ListView1.SelectedItems(0).Text
    End Sub

    'OK Button
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        return_result = TextBox1.Text.Trim
        Me.Close()
    End Sub
End Class