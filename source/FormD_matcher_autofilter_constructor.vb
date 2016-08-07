Imports System.Text.RegularExpressions

Public Class FormD_matcher_autofilter_constructor
    Dim regexList As New List(Of String)

    Private Sub FormD_matcher_autofilter_constructor_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Language.localize(Me)
        If Class3_matcher.autofilter_regex_options(0) Then CheckBox1.Checked = True Else CheckBox1.Checked = False
        If Class3_matcher.autofilter_regex_options(1) Then CheckBox2.Checked = True Else CheckBox2.Checked = False

        If Class3_matcher.autofilter_regex.Trim = "" Then test_refresh() : Exit Sub
        Dim r As String = Class3_matcher.autofilter_regex
        If r.StartsWith("%") Then
            RadioButton2.Checked = True
            r = r.Substring(1)
        Else
            RadioButton1.Checked = True
        End If
        TextBox1.Text = r
    End Sub

    Private Sub create_expression()
        Dim expr As String = ""
        If regexList.Contains("---------") Then expr = "("
        For Each s As String In regexList
            If s = "---------" Then
                expr = expr + ")("
            Else
                expr = expr + s
            End If
        Next
        If regexList.Contains("---------") Then expr = expr + ")"
        TextBox1.Text = expr
    End Sub

    Private Sub test_refresh()
        Try
            Dim regexp As New Regex(TextBox1.Text)
            Dim m As MatchCollection = regexp.Matches(TextBox2.Text)
            If m.Count = 0 Then
                Label4.Text = "Become: //empty//"
                NumericUpDown1.Maximum = 0
            Else
                'Label4.Text = "Become: " + m.Item(0).Value
                NumericUpDown1.Maximum = m.Item(0).Groups.Count
                Label4.Text = "Become: " + m.Item(0).Groups(CInt(NumericUpDown1.Value)).Value
                'If m.Item(0).Groups.Count > 1 Then
                'Label4.Text = "Become: " + m.Item(0).Groups(1).Value
                'Else
                'Label4.Text = "Become: " + m.Item(0).Groups(0).Value
                'End If
            End If
        Catch ex As Exception
            Label4.Text = "Become: //invalid expression//"
        End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged
        test_refresh()
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        test_refresh()
        If RadioButton1.Checked Then Class3_matcher.autofilter_regex = TextBox1.Text Else Class3_matcher.autofilter_regex = "%" + TextBox1.Text
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        test_refresh()
    End Sub
#Region "Add op buttons"
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If RadioButton3.Checked Then
            ListBox1.Items.Add("Number (Single)")
            regexList.Add("[0-9]")
        ElseIf RadioButton4.Checked Then
            ListBox1.Items.Add("Number (Single or Multiple)")
            regexList.Add("[0-9]+")
        ElseIf RadioButton5.Checked Then
            ListBox1.Items.Add("Number (Zerro or more)")
            regexList.Add("[0-9]*")
        End If
        create_expression()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If RadioButton3.Checked Then
            ListBox1.Items.Add("Letter (Single)")
            regexList.Add("[A-Za-z]")
        ElseIf RadioButton4.Checked Then
            ListBox1.Items.Add("Letter (Single or Multiple)")
            regexList.Add("[A-Za-z]+")
        ElseIf RadioButton5.Checked Then
            ListBox1.Items.Add("Letter (Zerro or more)")
            regexList.Add("[A-Za-z]*")
        End If
        create_expression()
    End Sub


    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        If RadioButton3.Checked Then
            ListBox1.Items.Add("AlphaNumeric (Single)")
            regexList.Add("\w")
        ElseIf RadioButton4.Checked Then
            ListBox1.Items.Add("AlphaNumeric (Single or Multiple)")
            regexList.Add("\w+")
        ElseIf RadioButton5.Checked Then
            ListBox1.Items.Add("AlphaNumeric (Zerro or more)")
            regexList.Add("\w*")
        End If
        create_expression()
    End Sub

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Button7.Click
        If RadioButton3.Checked Then
            ListBox1.Items.Add("Not AlphaNumeric (Single)")
            regexList.Add("\W")
        ElseIf RadioButton4.Checked Then
            ListBox1.Items.Add("Not AlphaNumeric (Single or Multiple)")
            regexList.Add("\W+")
        ElseIf RadioButton5.Checked Then
            ListBox1.Items.Add("Not AlphaNumeric (Zerro or more)")
            regexList.Add("\W*")
        End If
        create_expression()
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        If RadioButton3.Checked Then
            ListBox1.Items.Add("Space (Single)")
            regexList.Add("\s")
        ElseIf RadioButton4.Checked Then
            ListBox1.Items.Add("Space (Single or Multiple)")
            regexList.Add("\s+")
        ElseIf RadioButton5.Checked Then
            ListBox1.Items.Add("Space (Zerro or more)")
            regexList.Add("\s*")
        End If
        create_expression()
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        If RadioButton3.Checked Then
            ListBox1.Items.Add("Not Space (Single)")
            regexList.Add("\S")
        ElseIf RadioButton4.Checked Then
            ListBox1.Items.Add("Not Space (Single or Multiple)")
            regexList.Add("\S+")
        ElseIf RadioButton5.Checked Then
            ListBox1.Items.Add("Not Space (Zerro or more)")
            regexList.Add("\S*")
        End If
        create_expression()
    End Sub

    Private Sub Button10_Click(sender As System.Object, e As System.EventArgs) Handles Button10.Click
        If RadioButton3.Checked Then
            ListBox1.Items.Add("Any Char (Single)")
            regexList.Add(".")
        ElseIf RadioButton4.Checked Then
            ListBox1.Items.Add("Any Char (Single or Multiple)")
            regexList.Add(".+")
        ElseIf RadioButton5.Checked Then
            ListBox1.Items.Add("Any Char (Zerro or more)")
            regexList.Add(".*")
        End If
        create_expression()
    End Sub

    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        ListBox1.Items.Add("---------")
        regexList.Add("---------")
    End Sub
#End Region

    'Delete operator button
    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Dim curSel As Integer = ListBox1.SelectedIndex
        If curSel >= 0 Then
            regexList.RemoveAt(curSel)
            ListBox1.Items.RemoveAt(curSel)
        End If
        create_expression()

        If curSel >= 0 And ListBox1.Items.Count > 0 Then
            If curSel <= ListBox1.Items.Count - 1 Then ListBox1.SelectedIndex = curSel : Exit Sub
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        End If
    End Sub

    'reset to default
    Private Sub Button9_Click(sender As System.Object, e As System.EventArgs) Handles Button9.Click
        regexList.Clear()
        ListBox1.Items.Clear()
        RadioButton2.Checked = True
        TextBox1.Text = "[A-Za-z]{4}[A-Za-z]*"
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton1.CheckedChanged
        TextBox1_TextChanged(TextBox1, New System.EventArgs)
    End Sub

    Private Sub CheckBox_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged, CheckBox2.CheckedChanged
        Class3_matcher.autofilter_regex_options(0) = CheckBox1.Checked
        Class3_matcher.autofilter_regex_options(1) = CheckBox2.Checked
    End Sub


End Class