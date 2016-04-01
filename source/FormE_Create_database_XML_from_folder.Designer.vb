<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormE_Create_database_XML_from_folder
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBox17 = New System.Windows.Forms.CheckBox()
        Me.CheckBox21 = New System.Windows.Forms.CheckBox()
        Me.ComboBox12 = New System.Windows.Forms.ComboBox()
        Me.ComboBox13 = New System.Windows.Forms.ComboBox()
        Me.ComboBox14 = New System.Windows.Forms.ComboBox()
        Me.ComboBox15 = New System.Windows.Forms.ComboBox()
        Me.ComboBox16 = New System.Windows.Forms.ComboBox()
        Me.ComboBox17 = New System.Windows.Forms.ComboBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(12, 23)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(266, 20)
        Me.TextBox1.TabIndex = 22
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(284, 23)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(28, 20)
        Me.Button1.TabIndex = 23
        Me.Button1.Text = "..."
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Source folder:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 46)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 13)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "Destination XML:"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(284, 62)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(28, 20)
        Me.Button2.TabIndex = 26
        Me.Button2.Text = "..."
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(12, 62)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(266, 20)
        Me.TextBox2.TabIndex = 25
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(241, 88)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(71, 29)
        Me.Button3.TabIndex = 28
        Me.Button3.Text = "Go"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TextBox3)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.CheckBox17)
        Me.GroupBox1.Controls.Add(Me.CheckBox21)
        Me.GroupBox1.Controls.Add(Me.ComboBox12)
        Me.GroupBox1.Controls.Add(Me.ComboBox13)
        Me.GroupBox1.Controls.Add(Me.ComboBox14)
        Me.GroupBox1.Controls.Add(Me.ComboBox15)
        Me.GroupBox1.Controls.Add(Me.ComboBox16)
        Me.GroupBox1.Controls.Add(Me.ComboBox17)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 123)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(315, 190)
        Me.GroupBox1.TabIndex = 29
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Advanced Options"
        '
        'CheckBox17
        '
        Me.CheckBox17.AutoSize = True
        Me.CheckBox17.Location = New System.Drawing.Point(16, 136)
        Me.CheckBox17.Name = "CheckBox17"
        Me.CheckBox17.Size = New System.Drawing.Size(186, 17)
        Me.CheckBox17.TabIndex = 29
        Me.CheckBox17.Text = "remove all ( ) and [ ] tags from files"
        Me.CheckBox17.UseVisualStyleBackColor = True
        '
        'CheckBox21
        '
        Me.CheckBox21.AutoSize = True
        Me.CheckBox21.Location = New System.Drawing.Point(16, 113)
        Me.CheckBox21.Name = "CheckBox21"
        Me.CheckBox21.Size = New System.Drawing.Size(113, 17)
        Me.CheckBox21.TabIndex = 28
        Me.CheckBox21.Text = "fill CRC (very slow)"
        Me.CheckBox21.UseVisualStyleBackColor = True
        '
        'ComboBox12
        '
        Me.ComboBox12.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox12.FormattingEnabled = True
        Me.ComboBox12.Items.AddRange(New Object() {"DON'T USE", "First parenthesis", "First brackets"})
        Me.ComboBox12.Location = New System.Drawing.Point(16, 29)
        Me.ComboBox12.Name = "ComboBox12"
        Me.ComboBox12.Size = New System.Drawing.Size(97, 21)
        Me.ComboBox12.TabIndex = 22
        '
        'ComboBox13
        '
        Me.ComboBox13.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox13.FormattingEnabled = True
        Me.ComboBox13.Items.AddRange(New Object() {"Year", "Country", "Manufacturer"})
        Me.ComboBox13.Location = New System.Drawing.Point(119, 29)
        Me.ComboBox13.Name = "ComboBox13"
        Me.ComboBox13.Size = New System.Drawing.Size(116, 21)
        Me.ComboBox13.TabIndex = 23
        '
        'ComboBox14
        '
        Me.ComboBox14.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox14.FormattingEnabled = True
        Me.ComboBox14.Items.AddRange(New Object() {"DON'T USE", "Second parenthesis", "Second brackets"})
        Me.ComboBox14.Location = New System.Drawing.Point(16, 59)
        Me.ComboBox14.Name = "ComboBox14"
        Me.ComboBox14.Size = New System.Drawing.Size(97, 21)
        Me.ComboBox14.TabIndex = 24
        '
        'ComboBox15
        '
        Me.ComboBox15.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox15.FormattingEnabled = True
        Me.ComboBox15.Items.AddRange(New Object() {"Year", "Country", "Manufacturer"})
        Me.ComboBox15.Location = New System.Drawing.Point(119, 59)
        Me.ComboBox15.Name = "ComboBox15"
        Me.ComboBox15.Size = New System.Drawing.Size(116, 21)
        Me.ComboBox15.TabIndex = 25
        '
        'ComboBox16
        '
        Me.ComboBox16.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox16.FormattingEnabled = True
        Me.ComboBox16.Items.AddRange(New Object() {"DON'T USE", "Third parenthesis", "Third brackets"})
        Me.ComboBox16.Location = New System.Drawing.Point(16, 86)
        Me.ComboBox16.Name = "ComboBox16"
        Me.ComboBox16.Size = New System.Drawing.Size(97, 21)
        Me.ComboBox16.TabIndex = 26
        '
        'ComboBox17
        '
        Me.ComboBox17.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox17.FormattingEnabled = True
        Me.ComboBox17.Items.AddRange(New Object() {"Year", "Country", "Manufacturer"})
        Me.ComboBox17.Location = New System.Drawing.Point(119, 86)
        Me.ComboBox17.Name = "ComboBox17"
        Me.ComboBox17.Size = New System.Drawing.Size(116, 21)
        Me.ComboBox17.TabIndex = 27
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(12, 88)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(139, 22)
        Me.Button4.TabIndex = 30
        Me.Button4.Text = "Advanced Options >>>"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(45, 162)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(43, 13)
        Me.Label3.TabIndex = 30
        Me.Label3.Text = "Except:"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(98, 159)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(168, 20)
        Me.TextBox3.TabIndex = 31
        Me.TextBox3.Text = "disk|disc|tape"
        '
        'FormE_Create_database_XML_from_folder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(342, 329)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "FormE_Create_database_XML_from_folder"
        Me.Text = "Create database XML from rom folder"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox17 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox21 As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBox12 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox13 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox14 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox15 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox16 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox17 As System.Windows.Forms.ComboBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
