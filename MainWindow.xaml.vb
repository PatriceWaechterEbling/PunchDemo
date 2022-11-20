Class MainWindow
    Dim sb As String
    Dim PD1 As New LiaisonExcel()
    Private Sub Arrivee_Click(sender As Object, e As RoutedEventArgs) Handles Arrivee.Click

    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub Grid_Loaded(sender As Object, e As RoutedEventArgs)
        For x = 1 To 16
            UsrLst.Items.Add(x)

            Debug.Print(PD1.LireCellule(PD1.Fichier, "Users", x, 1, vbNull))
        Next
    End Sub

    Private Sub Main_Loaded(sender As Object, e As RoutedEventArgs) Handles Main.Loaded
        PD1.Init()

    End Sub
End Class
