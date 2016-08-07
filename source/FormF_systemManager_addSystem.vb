Public Class FormF_systemManager_addSystem
    Public _systems As List(Of String)
    Public value As String = ""

    Public Property systems As List(Of String)
        Get
            Return _systems
        End Get
        Set(value As List(Of String))
            _systems = value
            _systems.Sort()
            ComboBox1.Items.Clear()
            For Each sys As String In systems
                ComboBox1.Items.Add(sys)
            Next
        End Set
    End Property

    'OK
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        value = ComboBox1.Text
        Me.Close()
    End Sub

    'Cancel button
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        value = ""
        Me.Close()
    End Sub

    Private Sub FormF_systemManager_addSystem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Language.localize(Me)
    End Sub
End Class