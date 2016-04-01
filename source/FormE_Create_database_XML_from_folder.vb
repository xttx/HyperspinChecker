Public Class FormE_Create_database_XML_from_folder

    Private Sub FormE_Create_database_XML_from_folder_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Height = 150
        ComboBox12.SelectedIndex = 0
        ComboBox13.SelectedIndex = 0
        ComboBox14.SelectedIndex = 0
        ComboBox15.SelectedIndex = 0
        ComboBox16.SelectedIndex = 0
        ComboBox17.SelectedIndex = 0
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        If Button4.Text.EndsWith(">>>") Then
            Button4.Text = "Advanced Options <<<"
            Me.Height = 357
        Else
            Button4.Text = "Advanced Options >>>"
            Me.Height = 150
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim fb As New FolderBrowserDialog
        If TextBox1.Text.Trim <> "" Then fb.SelectedPath = TextBox1.Text
        fb.ShowDialog()
        TextBox1.Text = fb.SelectedPath
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim fb As New SaveFileDialog
        fb.RestoreDirectory = True
        fb.Filter = "XML Files|*.xml"
        If TextBox2.Text.Trim <> "" Then fb.InitialDirectory = TextBox2.Text
        fb.ShowDialog()
        TextBox2.Text = fb.FileName
    End Sub

    'GO
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Class2_xml.xml_from_folder_options_src = TextBox1.Text
        Class2_xml.xml_from_folder_options_dst = TextBox2.Text
        Class2_xml.xml_from_folder_options_fillCRC = CheckBox21.Checked
        Class2_xml.xml_from_folder_options_removeParanthesis = CheckBox17.Checked
        Class2_xml.xml_from_folder_options_adv = Format(ComboBox12.SelectedIndex, "00") + Format(ComboBox13.SelectedIndex, "00")
        Class2_xml.xml_from_folder_options_adv += Format(ComboBox14.SelectedIndex, "00") + Format(ComboBox15.SelectedIndex, "00")
        Class2_xml.xml_from_folder_options_adv += Format(ComboBox16.SelectedIndex, "00") + Format(ComboBox17.SelectedIndex, "00")
        Class2_xml.xml_from_folder_options_removeParanthesis_exceptions = TextBox3.Text
        Me.Close()
    End Sub
End Class