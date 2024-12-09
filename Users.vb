Imports System.Data.OleDb
Public Class Users
    Public Shared usermain
    Dim con As New OleDbConnection
    Dim cmd As OleDbCommand
    Dim dt As New DataTable
    Dim adap As New OleDbDataAdapter(cmd)

    Private bitmap As Bitmap


    Private Sub Users_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & Application.StartupPath & "\Database\dbase.mdb"
    End Sub

    Private Sub btnUSave_Click(sender As Object, e As EventArgs) Handles btnUSave.Click
        Try
            If txtName.Text = " " Or txtPhone.Text = " " Or txtAddress.Text = " " Or txtUsername.Text = " " Or txtPassword.Text = " " Then
                MsgBox("Please fill in all the fields", MessageBoxButtons.OK + MessageBoxIcon.Exclamation, "Incomplete")
                Return
            End If

            con.Open()
            cmd = New OleDbCommand("insert into Login([Name], [Phone], [Address], [Username], [Password]) values (?, ?, ?, ?, ?)", con)

            cmd.Parameters.AddWithValue("?", txtName.Text)
            cmd.Parameters.AddWithValue("?", txtPhone.Text)
            cmd.Parameters.AddWithValue("?", txtAddress.Text)
            cmd.Parameters.AddWithValue("?", txtUsername.Text)
            cmd.Parameters.AddWithValue("?", txtPassword.Text)
            cmd.ExecuteNonQuery()

            MsgBox("Record saved successfully!", MessageBoxButtons.OK + MessageBoxIcon.Information, Title:=Nothing)

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
        con.Open()
        cmd = con.CreateCommand()
        adap = New OleDbDataAdapter("select * from Login", con)
        adap.Fill(dt)
        dgUsersList.DataSource = dt
        con.Close()
    End Sub

    Private Sub btnUBooks_Click(sender As Object, e As EventArgs) Handles btnUBooks.Click
        Dim nextForm As New Books()
        nextForm.Show()
        Me.Close()
    End Sub

    Private Sub btnUDashboard_Click(sender As Object, e As EventArgs) Handles btnUDashboard.Click
        Dim dsboard As New Dashboard()
        Dashboard.Show()
        Me.Close()
    End Sub

    Private Sub btnUDelete_Click(sender As Object, e As EventArgs) Handles btnUDelete.Click
        Try
            If dgUsersList.SelectedRows.Count = 0 Then
                MsgBox("Please select a user to delete", MessageBoxButtons.OK + MessageBoxIcon.Warning, Title:=Nothing)
                Return
            End If

            Dim selectedRow As DataGridViewRow = dgUsersList.SelectedRows(0)
            Dim userName As String = selectedRow.Cells("Name").Value.ToString()
            Dim userPhone As String = selectedRow.Cells("Phone").Value.ToString()
            Dim result As DialogResult = MsgBox("Are you sure you want to delete this user?", MessageBoxButtons.YesNo + MessageBoxIcon.Question, Title:=Nothing)
            If result = DialogResult.No Then
                Return
            End If
            con.Open()

            cmd = New OleDbCommand("DELETE FROM Login WHERE [Name] = ? AND [Phone] = ?", con)
            cmd.Parameters.AddWithValue("?", userName)
            cmd.Parameters.AddWithValue("?", userPhone)
            Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
            If rowsAffected > 0 Then
                MsgBox("User deleted successfully!", MessageBoxButtons.OK + MessageBoxIcon.Information, Title:=Nothing)
            Else
                MsgBox("No user found with the provided details.", MessageBoxButtons.OK + MessageBoxIcon.Warning, Title:=Nothing)
            End If


            con.Close()
            dgUsersList.Rows.Remove(selectedRow)

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub btnUEdit_Click(sender As Object, e As EventArgs) Handles btnUEdit.Click
        If dgUsersList.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = dgUsersList.SelectedRows(0)

            Try
                txtName.Text = selectedRow.Cells("Name").Value.ToString()
                txtPhone.Text = selectedRow.Cells("Phone").Value.ToString()
                txtAddress.Text = selectedRow.Cells("Address").Value.ToString()
                txtUsername.Text = selectedRow.Cells("Username").Value.ToString()
                txtPassword.Text = selectedRow.Cells("Password").Value.ToString()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MsgBox("Please select a user to edit.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub

    Private Sub btnclear_Click(sender As Object, e As EventArgs) Handles btnclear.Click
        txtName.Text = ""
        txtPhone.Text = ""
        txtAddress.Text = ""
        txtUsername.Text = ""
        txtPassword.Text = ""
    End Sub

    Private Sub dgUsersList_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgUsersList.CellFormatting
        If e.ColumnIndex = 5 Then
            If e.Value IsNot Nothing Then
                e.Value = New String("*", e.Value.ToString().Length)
            End If
        End If
    End Sub
End Class