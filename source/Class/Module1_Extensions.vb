Imports System.Runtime.CompilerServices
Module Module1_Extensions
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
End Module
