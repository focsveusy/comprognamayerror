Imports System.Data.OleDb
Imports System.Data

Public Class Form1
    Dim dbCON As New OleDb.OleDbConnection
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dbCON = New OleDb.OleDbConnection
        dbCON.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & Application.StartupPath & "\Database\dbase.mdb"
        dbCON.Open()

        'MsgBox("Connected")
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click

        If txtUname.Text = Nothing Or txtPassword.Text = Nothing Then
            MsgBox("Enter credentials", MsgBoxStyle.Exclamation)
        Else
            If dbCON.State = ConnectionState.Closed Then
                dbCON.Open()
            End If
            Dim cmd As New OleDbCommand("select count(*) from login where username=? and password=?", dbCON)
            cmd.Parameters.AddWithValue("@1", OleDbType.VarChar).Value = txtUname.Text
            cmd.Parameters.AddWithValue("@2", OleDbType.VarChar).Value = txtPassword.Text
            Dim count = Convert.ToInt32(cmd.ExecuteScalar())



            If (count > 0) Then
                MsgBox("Login Successfully!", MsgBoxStyle.Information)
            Else
                MsgBox("Acount not found, check the credential!", MsgBoxStyle.Critical)

            End If
            Dim userbills As New Bills
            userbills.Show()

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim admin As New admin
        admin.Show()
        Me.Hide()
    End Sub

    Private Sub txtPassword_TextChanged(sender As Object, e As EventArgs) Handles txtPassword.TextChanged
        If txtPassword.UseSystemPasswordChar = True Then
            txtPassword.UseSystemPasswordChar = False

        End If
    End Sub
End Class
