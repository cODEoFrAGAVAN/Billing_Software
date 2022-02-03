Imports System.Data.OleDb

Public Class Form1
    Dim con As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Public UT As String
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Form2.Show()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=E:\Billing_software\project\database1.mdb")
        con.Open()
        con.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim user_name As String
        Dim Upassword As String
        user_name = TextBox1.Text
        Upassword = TextBox2.Text
        Dim dr As OleDbDataReader

        If (user_name = "ADMIN" And Upassword = "ADMIN") Then
            UT = "ADMIN"
            Form12.Show()
        End If
        con.Open()
        cmd = New OleDbCommand("select * from table1 where user_name='" & user_name & "' and Upassword='" & Upassword & "'", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            UT = "OTHERS"
            Form12.Show()

        Else
            MsgBox("login denied")
        End If
        con.Close()
    End Sub
End Class
