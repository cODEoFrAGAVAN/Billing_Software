Imports System.Data.OleDb
Public Class Form15
    Dim con As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Form15_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=E:\Billing_software\project\database1.mdb")
        con.Open()

        con.Close()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim department As String
        department = ComboBox1.Text

        con.Open()
        Dim da As OleDbDataAdapter
        Dim ds As New Data.DataSet
        da = New OleDbDataAdapter("select * from table5 where department='" & department & "'", con)
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.Refresh()
        con.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim acfrom As String
        acfrom = ComboBox2.Text

        con.Open()
        Dim da As OleDbDataAdapter
        Dim ds As New Data.DataSet
        da = New OleDbDataAdapter("select * from table5 where year_start='" & acfrom & "'", con)
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.Refresh()
        con.Close()
    End Sub
End Class