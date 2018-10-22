Imports System.ComponentModel
Imports System.Text.RegularExpressions

Public Class FormD_matcher_autofilter_constructor
    Dim regexList As New List(Of String)
    Dim textbox_refresh As Boolean = False

    Private Sub FormD_matcher_autofilter_constructor_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Language.localize(Me)
        If Class3_matcher.autofilter_regex_options(0) Then CheckBox1.Checked = True Else CheckBox1.Checked = False
        If Class3_matcher.autofilter_regex_options(1) Then CheckBox2.Checked = True Else CheckBox2.Checked = False
        NumericUpDown1.Value = Class3_matcher.autofilter_regex_opt_outGroup

        If Class3_matcher.autofilter_regex.Trim = "" Then test_refresh() : Exit Sub
        Dim r As String = Class3_matcher.autofilter_regex
        If r.StartsWith("%") Then
            RadioButton2.Checked = True
            r = r.Substring(1)
        Else
            RadioButton1.Checked = False
        End If
        TextBox1.Text = r

        For Each kv In Class3_matcher.autofilter_regex_presets
            ComboBox1.Items.Add(kv.Key)
        Next
    End Sub

    Private Sub parse_regex()
        regexList.Clear()
        ListBox1.Items.Clear()
        Try
            'test regex: https://regex101.com/r/WRqiTt/1
            '(?<!\\) - negative look behind: don't match pattern if it starts with "\"
            '\( - match open paranthesis
            '.*? - match any symbol zero or unlimited times, but as few times as possible
            '\) - match close paranthesis
            Dim rx_groups = New Regex("(?<!\\)\(.*?\)").Matches(TextBox1.Text)

            Dim groups_arr() As String
            If rx_groups.Count = 0 Then groups_arr = {TextBox1.Text} Else groups_arr = rx_groups.Cast(Of Match).Select(Of String)(Function(m, b) m.Value).ToArray
            Dim need_add_grp_separator As Boolean = False
            For Each grp In groups_arr
                grp = grp.Trim
                need_add_grp_separator = False
                If grp.StartsWith("(") And grp.EndsWith(")") Then grp = grp.Substring(1, grp.Length - 2) : need_add_grp_separator = True

                'Dim rx_split = New Regex("((?<!\\)\[.*?\][\+\*]*)|(\\[A-Za-z][\+\*]*)|(.[\+\*]*)").Matches(grp)
                Dim rx_split = New Regex("((?<!\\)\[.*?\](\*|\+|\{[0-9]+\})*)|(\\[A-Za-z](\*|\+|\{[0-9]+\})*)|(.(\*|\+|\{[0-9]+\})*)").Matches(grp)
                For Each expr As Match In rx_split
                    Dim v = expr.Value

                    Dim expr_str As String = ""
                    If v.StartsWith("[") Then
                        If v.StartsWith("[0-9]") Then expr_str = "Number"
                        If v.StartsWith("[A-Za-z]") Then expr_str = "Letter"
                        If v.StartsWith("[^0-9]") Then expr_str = "Not Number"
                        If v.StartsWith("[^A-Za-z]") Then expr_str = "Not Letter"
                    ElseIf v.StartsWith("\") Then
                        If v.StartsWith("\w") Then expr_str = "AlphaNumeric"
                        If v.StartsWith("\W") Then expr_str = "Not AlphaNumeric"
                        If v.StartsWith("\s") Then expr_str = "Space"
                        If v.StartsWith("\S") Then expr_str = "Not Space"
                    Else
                        If v.StartsWith(".") Then expr_str = "Any Char"
                    End If

                    Dim qualifier As String = ""
                    If v.Contains("{") And v.EndsWith("}") Then
                        Dim num = v.Substring(v.IndexOf("{") + 1)
                        num = num.Substring(0, num.Length - 1)
                        qualifier = "(" + num + " times)"
                    ElseIf v.EndsWith("*") Then
                        qualifier = "(Zerro or more)"
                    ElseIf v.EndsWith("+") Then
                        qualifier = "(Single or Multiple)"
                    ElseIf v.EndsWith("?") Then
                        qualifier = "(Zerro or one)"
                    Else
                        qualifier = "(Single)"
                    End If

                    regexList.Add(v)
                    ListBox1.Items.Add(expr_str + " " + qualifier)
                    If need_add_grp_separator Then ListBox1.Items.Add("---------")
                Next
            Next
        Catch ex As Exception
            ListBox1.Items.Add("There was an error while parsing regular expression.")
        End Try
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

        textbox_refresh = True
        TextBox1.Text = expr
        textbox_refresh = False
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

    'Demo string changed
    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged
        test_refresh()
    End Sub

    'Regex changed
    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        test_refresh()
        If RadioButton1.Checked Then Class3_matcher.autofilter_regex = TextBox1.Text Else Class3_matcher.autofilter_regex = "%" + TextBox1.Text
        If Not textbox_refresh Then parse_regex()
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        test_refresh()
        Class3_matcher.autofilter_regex_opt_outGroup = CInt(NumericUpDown1.Value)
    End Sub
#Region "Add op buttons"
    'Add number
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
    'Add NOT number
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If RadioButton3.Checked Then
            ListBox1.Items.Add("Not Number (Single)")
            'regexList.Add("\D")
            regexList.Add("[^0-9]")
        ElseIf RadioButton4.Checked Then
            ListBox1.Items.Add("Not Number (Single or Multiple)")
            'regexList.Add("\D+")
            regexList.Add("[^0-9]+")
        ElseIf RadioButton5.Checked Then
            ListBox1.Items.Add("Not Number (Zerro or more)")
            'regexList.Add("\D*")
            regexList.Add("[^0-9]*")
        End If
        create_expression()
    End Sub
    'Add Letter
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
    'Add NOT Letter
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If RadioButton3.Checked Then
            ListBox1.Items.Add("Not Letter (Single)")
            regexList.Add("[^A-Za-z]")
        ElseIf RadioButton4.Checked Then
            ListBox1.Items.Add("Not Letter (Single or Multiple)")
            regexList.Add("[^A-Za-z]+")
        ElseIf RadioButton5.Checked Then
            ListBox1.Items.Add("Not Letter (Zerro or more)")
            regexList.Add("[^A-Za-z]*")
        End If
        create_expression()
    End Sub
    'Add AlphaNumeric
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
    'Add NOT AlphaNumeric
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
    'Add Space
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
    'Add NOT Space
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
    'Add Any char
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
    'Add group separator
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

    'Preset - Load
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If ComboBox1.Text.Trim = "" Then MsgBox("Please, choose a preset to load.") : Exit Sub
        If Not Class3_matcher.autofilter_regex_presets.ContainsKey(ComboBox1.Text.Trim) Then MsgBox("There is no such preset.") : Exit Sub

        Dim p = Class3_matcher.autofilter_regex_presets(ComboBox1.Text.Trim).Trim.Split({"^^^"}, StringSplitOptions.RemoveEmptyEntries)
        If p.Count <> 4 Then MsgBox("Invalid preset data.") : Exit Sub

        If p(0).StartsWith("%") Then RadioButton1.Checked = True : p(0) = p(0).Substring(1) Else RadioButton1.Checked = False
        TextBox1.Text = p(0)
        If p(1).ToUpper.Trim = "TRUE" Then CheckBox1.Checked = True Else CheckBox1.Checked = False
        If p(2).ToUpper.Trim = "TRUE" Then CheckBox2.Checked = True Else CheckBox2.Checked = False
        If IsNumeric(p(3)) Then NumericUpDown1.Value = CInt(p(3))
    End Sub
    'Preset - Save
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If ComboBox1.Text.Trim = "" Then MsgBox("Please, enter preset name.") : Exit Sub

        Dim preset_string = TextBox1.Text.Trim + "^^^"
        If CheckBox1.Checked Then preset_string += "TRUE" Else preset_string += "FALSE"
        preset_string += "^^^"
        If CheckBox2.Checked Then preset_string += "TRUE" Else preset_string += "FALSE"
        preset_string += "^^^"
        preset_string += NumericUpDown1.Value.ToString

        If Class3_matcher.autofilter_regex_presets.ContainsKey(ComboBox1.Text.Trim) Then
            Class3_matcher.autofilter_regex_presets(ComboBox1.Text.Trim) = preset_string
        Else
            ComboBox1.Items.Add(ComboBox1.Text.Trim)
            Class3_matcher.autofilter_regex_presets.Add(ComboBox1.Text.Trim, preset_string)
        End If
    End Sub


End Class