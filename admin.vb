Imports finalproject_bookshop_.dbaseDataSet1TableAdapters

Public Class admin
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim login = New Form1
        login.Show()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If txtpass.Text = "mulatoandarlantico" Then
            Dim Obj = New Books()
            Obj.Show()
            Me.Hide()
        Else
            MsgBox("Wrong password!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, " ")
        End If
    End Sub

    Private Sub admin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtpass.UseSystemPasswordChar = True

    End Sub

    Private Sub btnshowhide_Click(sender As Object, e As EventArgs) Handles btnshowhide.Click
        If txtpass.UseSystemPasswordChar = True Then

            txtpass.UseSystemPasswordChar = False
            btnshowhide.Text = "Hide"

        Else
            txtpass.UseSystemPasswordChar = True

            btnshowhide.Text = "Show"
        End If
    End Sub
End Class