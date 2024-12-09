Public Class MainBookshop
    Private Sub btnDBooks_Click(sender As Object, e As EventArgs) Handles btnDBooks.Click
        ' Perform necessary actions, such as saving data

        ' Show the next form and close the current form
        Dim nextForm As New Books()
        nextForm.Show()
        Me.Close() ' Close the current form
    End Sub

    Private Sub btnDUsers_Click(sender As Object, e As EventArgs) Handles btnDUsers.Click
        Dim usermain As New Users()
        Users.Show()
        Me.Close()
    End Sub

    Private Sub btnDDashboard_Click(sender As Object, e As EventArgs) Handles btnDDashboard.Click

        Dim dsboard As New Dashboard()
        Dashboard.Show()
        Me.Close()
    End Sub
End Class