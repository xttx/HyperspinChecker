Imports Microsoft.VisualBasic.FileIO
Public Class Form4_genres_favorites
    Dim p As String = ""
    Dim moveValue As Boolean = False

    Private Sub Form4_genres_favorites_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.Text = "Genre / favorites manager: " + Form1.ComboBox1.SelectedItem.ToString
        Button4.Enabled = True
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        ComboBox1.Items.Add("All Games")
        p = Form1.xmlPath.Substring(0, Form1.xmlPath.LastIndexOf("\") + 1)
        If FileSystem.FileExists(p + "favorites.txt") Then ComboBox1.Items.Add("Favorites") : ComboBox2.Items.Add("Favorites") : Button4.Enabled = False
        If FileSystem.FileExists(p + "genre.xml") Then
            Dim x As New Xml.XmlDocument
            x.Load(p + "genre.xml")
            For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
                Dim t As String = node.Attributes.GetNamedItem("name").Value
                ComboBox1.Items.Add(t) : ComboBox2.Items.Add(t)
            Next
        End If
        If ComboBox1.Items.Count > 0 Then ComboBox1.SelectedIndex = 0
        If ComboBox2.Items.Count > 0 Then ComboBox2.SelectedIndex = 0
        CheckBox1.Checked = False
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        ListBox1.Items.Clear()
        ListBox3.Items.Clear()
        Dim x As New Xml.XmlDocument
        If ComboBox1.SelectedIndex = 0 Then
            x.Load(Form1.xmlPath)
            moveValue = CheckBox1.Checked
            CheckBox1.Checked = False
            CheckBox1.Enabled = False
        Else
            If CheckBox1.Enabled = False Then CheckBox1.Checked = moveValue : CheckBox1.Enabled = True
        End If

        If ComboBox1.SelectedIndex > 0 And ComboBox1.SelectedItem.ToString <> "Favorites" Then
            If Not FileSystem.FileExists(p + ComboBox1.SelectedItem.ToString + ".xml") Then MsgBox(ComboBox1.SelectedItem.ToString + ".xml does not exist.") : Exit Sub
            x.Load(p + ComboBox1.SelectedItem.ToString + ".xml")
        End If
        If ComboBox1.SelectedItem.ToString <> "Favorites" Then
            For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
                Dim t As String = node.SelectSingleNode("description").InnerText
                ListBox1.Items.Add(t)
                ListBox3.Items.Add(node.Attributes.GetNamedItem("name").Value)
            Next
        Else
            FileOpen(1, p + "favorites.txt", OpenMode.Input)
            Do While Not EOF(1)
                Dim l As String = LineInput(1)
                ListBox1.Items.Add(l)
                ListBox3.Items.Add(l)
            Loop
            FileClose(1)
        End If
        Label3.Text = "Total:" + vbCrLf + ListBox1.Items.Count.ToString
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        ListBox2.Items.Clear()
        ListBox4.Items.Clear()
        Dim x As New Xml.XmlDocument
        If ComboBox2.SelectedIndex > 0 And ComboBox2.SelectedItem.ToString <> "Favorites" Then
            If Not FileSystem.FileExists(p + ComboBox2.SelectedItem.ToString + ".xml") Then MsgBox(ComboBox2.SelectedItem.ToString + ".xml does not exist.") : Exit Sub
            x.Load(p + ComboBox2.SelectedItem.ToString + ".xml")
        End If
        If ComboBox2.SelectedItem.ToString <> "Favorites" Then
            For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
                Dim t As String = node.SelectSingleNode("description").InnerText
                ListBox2.Items.Add(t)
                ListBox4.Items.Add(node.Attributes.GetNamedItem("name").Value)
            Next
        Else
            FileOpen(1, p + "favorites.txt", OpenMode.Input)
            Do While Not EOF(1)
                Dim l As String = LineInput(1)
                ListBox2.Items.Add(l)
                ListBox4.Items.Add(l)
            Loop
            FileClose(1)
        End If
        Label4.Text = "Total:" + vbCrLf + ListBox2.Items.Count.ToString
    End Sub

    'Create new genre
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim g As String = InputBox("Enter genre name:", "Genre name", "new genre")
        If g = "" Then Exit Sub

        Dim x As New Xml.XmlDocument
        Dim nodeMenu As Xml.XmlNode
        If FileSystem.FileExists(p + "genre.xml") Then
            x.Load(p + "genre.xml")
            nodeMenu = x.SelectSingleNode("/menu")
        Else
            nodeMenu = x.CreateElement("menu")
            x.AppendChild(nodeMenu)
        End If
        Dim x2 As New Xml.XmlDocument
        Dim nodeMenu_x2 As Xml.XmlElement
        nodeMenu_x2 = x2.CreateElement("menu")
        x2.AppendChild(nodeMenu_x2)
        x2.Save(p + g + ".xml")

        Dim newGenre As Xml.XmlElement = x.CreateElement("game")
        newGenre.SetAttribute("name", g)
        nodeMenu.AppendChild(newGenre)
        x.Save(p + "genre.xml")

        ComboBox1.Items.Add(g)
        ComboBox2.Items.Add(g)
        ComboBox2.SelectedIndex = ComboBox2.Items.Count - 1
    End Sub

    'Create favorites
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        FileOpen(1, p + "favorites.txt", OpenMode.Output) : FileClose(1)
        Button4.Enabled = False
        ComboBox1.Items.Insert(1, "Favorites")
        ComboBox2.Items.Insert(0, "Favorites")
        ComboBox2.SelectedIndex = 0
    End Sub

    'ADD
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ListBox1.SelectedIndex < 0 Then MsgBox("Please, select a game to add.") : Exit Sub
        Dim x As New Xml.XmlDocument
        x.Load(Form1.xmlPath)
        Dim gameNode As Xml.XmlNode = x.SelectSingleNode("/menu/game[@name=""" + ListBox3.Items(ListBox1.SelectedIndex).ToString + """]")

        If ComboBox2.SelectedItem.ToString <> "Favorites" Then
            If Not FileSystem.FileExists(p + ComboBox2.SelectedItem.ToString + ".xml") Then MsgBox(ComboBox1.SelectedItem.ToString + ".xml does not exist.") : Exit Sub

            Dim xg As New Xml.XmlDocument
            Dim xgMenuNode As Xml.XmlNode
            xg.Load(p + ComboBox2.SelectedItem.ToString + ".xml")
            xgMenuNode = xg.SelectSingleNode("/menu")
            Dim xgNewNode As Xml.XmlElement = xg.CreateElement("game")
            xgNewNode.SetAttribute("name", gameNode.Attributes.GetNamedItem("name").Value)
            xgNewNode.InnerXml = gameNode.InnerXml


            If (RadioButton1.Checked Or RadioButton2.Checked) And ListBox2.SelectedIndex >= 0 Then
                Dim xg_node As Xml.XmlNode = xg.SelectSingleNode("/menu/game[@name=""" + ListBox4.Items(ListBox2.SelectedIndex).ToString + """]")
                If RadioButton1.Checked Then
                    xgMenuNode.InsertBefore(xgNewNode, xg_node)
                    ListBox2.Items.Insert(ListBox2.SelectedIndex, xgNewNode.SelectSingleNode("description").InnerText)
                    ListBox4.Items.Insert(ListBox2.SelectedIndex, xgNewNode.Attributes.GetNamedItem("name").Value)
                Else
                    xgMenuNode.InsertAfter(xgNewNode, xg_node)
                    ListBox2.Items.Insert(ListBox2.SelectedIndex + 1, xgNewNode.SelectSingleNode("description").InnerText)
                    ListBox4.Items.Insert(ListBox2.SelectedIndex + 1, xgNewNode.Attributes.GetNamedItem("name").Value)
                End If
            Else
                xgMenuNode.AppendChild(xgNewNode)
                ListBox2.Items.Add(xgNewNode.SelectSingleNode("description").InnerText)
                ListBox4.Items.Add(xgNewNode.Attributes.GetNamedItem("name").Value)
            End If
            xg.Save(p + ComboBox2.SelectedItem.ToString + ".xml")
        Else
            If Not FileSystem.FileExists(p + "favorites.txt") Then MsgBox("favorites.txt does not exist.") : Exit Sub
            FileOpen(1, p + "favorites.txt", OpenMode.Append) : PrintLine(1, gameNode.Attributes.GetNamedItem("name").Value) : FileClose(1)
            ListBox2.Items.Add(gameNode.SelectSingleNode("description").InnerText)
            ListBox4.Items.Add(gameNode.Attributes.GetNamedItem("name").Value)
        End If
        If ListBox1.Items.Count > ListBox1.SelectedIndex + 1 Then ListBox1.SelectedIndex = ListBox1.SelectedIndex + 1
    End Sub

    'Delete
    Private Function Delete(ByVal romNameToRemove As String, ByVal removeFrom As String) As Integer
        If removeFrom <> "Favorites" Then
            If Not FileSystem.FileExists(p + removeFrom + ".xml") Then MsgBox(removeFrom + ".xml does not exist.") : Return -1

            Dim x As New Xml.XmlDocument
            x.Load(p + ComboBox2.SelectedItem.ToString + ".xml")
            Dim x_node As Xml.XmlNode = x.SelectSingleNode("/menu/game[@name=""" + romNameToRemove + """]")
            x.SelectSingleNode("/menu").RemoveChild(x_node)
            x.Save(p + removeFrom + ".xml")
        Else
            If Not FileSystem.FileExists(p + "favorites.txt") Then MsgBox("favorites.txt does not exist.") : Return -1
            FileOpen(1, p + "favorites.txt", OpenMode.Input)
            FileOpen(2, p + "favorites_.txt", OpenMode.Output)
            Dim s As String
            While Not EOF(1)
                s = LineInput(1)
                If s.ToUpper <> romNameToRemove.ToUpper Then PrintLine(2, s)
            End While
            FileClose(1)
            FileClose(2)
            FileSystem.MoveFile(p + "favorites_.txt", p + "favorites.txt", True)
        End If
        Return 0
    End Function

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        If Form4_genres_favorites_preview.Visible = False Then Exit Sub
        Dim t As String = "file:///" + Application.StartupPath.Replace("\", "/") + "/player.swf"
        Dim tv As String
        If FileSystem.FileExists(Class1.videoPath + ListBox3.Items(ListBox1.SelectedIndex).ToString + ".flv") Then
            tv = "file:///" + Class1.videoPath.Replace("\", "/") + ListBox3.Items(ListBox1.SelectedIndex).ToString + ".flv"
        ElseIf FileSystem.FileExists(Class1.videoPath + ListBox3.Items(ListBox1.SelectedIndex).ToString + ".mp4") Then
            tv = "file:///" + Class1.videoPath.Replace("\", "/") + ListBox3.Items(ListBox1.SelectedIndex).ToString + ".mp4"
        Else
            tv = ""
        End If
        tv = tv.Replace("&", "%26").Replace("+", "%2B")
        Dim swfparam As String = t + "?file=" + tv + "&autostart=true&logo=logo.PNG&smoothing=true"
        Form4_genres_favorites_preview.AxShockwaveFlash1.LoadMovie(0, swfparam)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Form4_genres_favorites_preview.Show()
    End Sub 'Show Preview

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ListBox2.SelectedIndex < 0 Then Exit Sub
        Dim i As Integer = ListBox2.SelectedIndex
        If Delete(ListBox4.Items(ListBox2.SelectedIndex).ToString, ComboBox2.SelectedItem.ToString) < 0 Then Exit Sub
        ListBox2.Items.RemoveAt(i)
        ListBox4.Items.RemoveAt(i)
        If ListBox2.Items.Count > i Then
            ListBox2.SelectedIndex = i
        Else
            If ListBox2.Items.Count > 0 Then ListBox2.SelectedIndex = ListBox2.Items.Count - 1
        End If
    End Sub 'Delete from listbox2

    Private Sub Form4_genres_favorites_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Language.localize(Me)
    End Sub
End Class