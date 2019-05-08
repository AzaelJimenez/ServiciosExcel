Imports System.Data.SqlClient
Public Class FrmMenu
    Private Sub FrmMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' MetodoMetasInf2019()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Frmdoce.ShowDialog()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        FrmModificacion.ShowDialog()
    End Sub


End Class