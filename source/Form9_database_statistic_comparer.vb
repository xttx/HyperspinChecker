Public Class Form9_database_statistic_comparer
    Public file1 As String = ""
    Public file2 As String = ""
    Public file3 As String = ""

    Private Sub Form9_database_statistic_comparer_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim xml1 As New Xml.XmlDocument
        Dim xml2 As New Xml.XmlDocument
        Dim xml3 As New Xml.XmlDocument
        Dim gamelist1 As New List(Of String)
        Dim gamelist2 As New List(Of String)
        Dim gamelist3 As New List(Of String)
        Dim intersect As New List(Of String)
        xml1.Load(file1) : gamelist1 = xmlToGamelist(xml1)
        xml2.Load(file2) : gamelist2 = xmlToGamelist(xml2)
        If file3 <> "" Then xml3.Load(file3) : gamelist3 = xmlToGamelist(xml3)

        Dim intersect_tmp As System.Collections.Generic.IEnumerable(Of String)
        If file3 = "" Then
            intersect_tmp = gamelist1.Intersect(gamelist2)
        Else
            intersect_tmp = gamelist1.Intersect(gamelist2).Intersect(gamelist3)
        End If

        intersect.AddRange(intersect_tmp)
        For Each s As String In intersect
            If gamelist1.Contains(s) Then gamelist1.Remove(s)
            If gamelist2.Contains(s) Then gamelist2.Remove(s)
            If gamelist3.Contains(s) Then gamelist3.Remove(s)
        Next


        TextBox1.Text = "Unique for:" + vbCrLf
        TextBox1.Text = TextBox1.Text + file1 + vbCrLf
        TextBox1.Text = TextBox1.Text + gamelist1.Count.ToString + " entries" + vbCrLf
        TextBox1.Text = TextBox1.Text + "------------------------------------------------------------------------" + vbCrLf
        For Each s In gamelist1
            TextBox1.Text = TextBox1.Text + s + vbCrLf
        Next
        TextBox1.Text = TextBox1.Text + "------------------------------------------------------------------------" + vbCrLf

        TextBox2.Text = "Unique for:" + vbCrLf
        TextBox2.Text = TextBox2.Text + file2 + vbCrLf
        TextBox2.Text = TextBox2.Text + gamelist2.Count.ToString + " entries" + vbCrLf
        TextBox2.Text = TextBox2.Text + "------------------------------------------------------------------------" + vbCrLf
        For Each s In gamelist2
            TextBox2.Text = TextBox2.Text + s + vbCrLf
        Next
        TextBox2.Text = TextBox2.Text + "------------------------------------------------------------------------" + vbCrLf

        If file3 <> "" Then
            TextBox3.Text = "Unique for:" + vbCrLf
            TextBox3.Text = TextBox3.Text + file3 + vbCrLf
            TextBox3.Text = TextBox3.Text + gamelist3.Count.ToString + " entries" + vbCrLf
            TextBox3.Text = TextBox3.Text + "------------------------------------------------------------------------" + vbCrLf
            For Each s In gamelist3
                TextBox3.Text = TextBox3.Text + s + vbCrLf
            Next
            TextBox3.Text = TextBox3.Text + "------------------------------------------------------------------------" + vbCrLf
        End If
    End Sub

    Private Function xmlToGamelist(xml As Xml.XmlDocument) As List(Of String)
        Dim tmpList As New List(Of String)
        For Each x As Xml.XmlNode In xml.SelectNodes("/menu/game")
            tmpList.Add(Class2_xml.streap_brackets(x.Attributes("name").Value))
        Next
        Return tmpList
    End Function
End Class