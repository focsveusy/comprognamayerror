Imports System.Data.OleDb
Public Class Books
    Dim con As New OleDbConnection
    Dim cmd As OleDbCommand
    Dim dt As New DataTable
    Dim adap As New OleDbDataAdapter(cmd)

    Private bitmap As Bitmap
    Private Sub btnBUsers_Click(sender As Object, e As EventArgs) Handles btnBUsers.Click

        Dim usermain As New Users()
        Users.Show()
        Me.Close()
    End Sub

    Private Sub btnBDashboard_Click(sender As Object, e As EventArgs) Handles btnBDashboard.Click
        Dim dsboard As New Dashboard()
        Dashboard.Show()
        Me.Close()
    End Sub

    Private Sub Books_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & Application.StartupPath & "\Database\dbase.mdb"
    End Sub

    Private Sub btnBSave_Click(sender As Object, e As EventArgs) Handles btnBSave.Click
        Try

            If txtBTitle.Text = " " Or txtAuthor.Text = " " Or cboCategories.Text = " " Or txtQuantity.Text = " " Or txtPrice.Text = " " Then
                MsgBox("Please fill in all the fields", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim quantity As Integer
            If Not Integer.TryParse(txtQuantity.Text, quantity) Then
                MsgBox("Invalid quantity, please enter a number", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim price As Decimal
            If Not Decimal.TryParse(txtPrice.Text, price) Then
                MsgBox("Invalid price, please enter a valid decimal number", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            con.Open()
            cmd = New OleDbCommand("insert into books([Book Title], [Author], [Category], [Quantity], [Price]) values (?, ?, ?, ?, ?)", con)
            cmd.Parameters.AddWithValue("?", txtBTitle.Text)
            cmd.Parameters.AddWithValue("?", txtAuthor.Text)
            cmd.Parameters.AddWithValue("?", cboCategories.Text)
            cmd.Parameters.AddWithValue("?", quantity)
            cmd.Parameters.AddWithValue("?", price)
            cmd.ExecuteNonQuery()

            MsgBox("Record saved successfully!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            con.Close()

            con.Open()
            adap = New OleDbDataAdapter("select * from books", con)
            adap.Fill(dt)
            dgBooksList.DataSource = dt
            con.Close()

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub btnBLogout_Click(sender As Object, e As EventArgs) Handles btnBLogout.Click
        Dim loginForm As New Form1()
        Form1.Show()
        MsgBox("User Succefully Log out!")
        Me.Close()


    End Sub
    Dim key = 0
    Private Sub dgBooksList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgBooksList.CellContentClick
        If dgBooksList.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = dgBooksList.SelectedRows(0)

            Try
                txtID.Text = selectedRow.Cells(1).Value.ToString()
                txtBTitle.Text = selectedRow.Cells(2).Value.ToString()
                txtAuthor.Text = selectedRow.Cells(3).Value.ToString()
                cboCategories.Text = selectedRow.Cells(4).Value.ToString()
                txtQuantity.Text = selectedRow.Cells(5).Value.ToString()
                txtPrice.Text = selectedRow.Cells(6).Value.ToString()
                If txtBTitle.Text = "" Then
                    key = 0
                End If
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MsgBox("Please select a book to edit.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub btnBDelete_Click(sender As Object, e As EventArgs) Handles btnBDelete.Click
        If dgBooksList.SelectedRows.Count = 0 Then
            MsgBox("Please select a book to delete", MessageBoxButtons.OK + MessageBoxIcon.Warning, Title:=Nothing)
            Return
        End If

        Dim selectedRow As DataGridViewRow = dgBooksList.SelectedRows(0)
        Dim userName As String = selectedRow.Cells("Book ID").Value.ToString()
        Dim userPhone As String = selectedRow.Cells("Book Title").Value.ToString()
        Dim result As DialogResult = MsgBox("Are you sure you want to delete this user?", MessageBoxButtons.YesNo + MessageBoxIcon.Question, Title:=Nothing)
        If result = DialogResult.No Then
            Return
        End If
        con.Open()

        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
        If rowsAffected > 0 Then
            MsgBox("Book deleted successfully!", MessageBoxButtons.OK + MessageBoxIcon.Information, Title:=Nothing)
        Else
            MsgBox("No book found with the provided details.", MessageBoxButtons.OK + MessageBoxIcon.Warning, Title:=Nothing)
        End If


        con.Close()
        dgBooksList.Rows.Remove(selectedRow)

        If con.State = ConnectionState.Open Then
            con.Close()
        End If
    End Sub
End Class