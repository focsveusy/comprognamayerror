Public Class Dashboard
    Private Sub btnDBooks_Click(sender As Object, e As EventArgs) Handles btnDBooks.Click
        Dim nextForm As New Books()
        nextForm.Show()
        Me.Close()
    End Sub

    Private Sub btnDUsers_Click(sender As Object, e As EventArgs) Handles btnDUsers.Click
        Dim usermain As New Users()
        Users.Show()
        Me.Close()
    End Sub
End Class