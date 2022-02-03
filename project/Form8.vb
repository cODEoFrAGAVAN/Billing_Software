Imports System.Data.OleDb

Public Class Form8
    Dim con As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        con.Open()
        Dim da As OleDbDataAdapter
        Dim ds As New Data.DataSet
        da = New OleDbDataAdapter("select * from table3 where product_code>0", con)
        da.Fill(ds)

        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.Refresh()
        con.Close()
    End Sub

    Private Sub Form8_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=E:\Billing_software\project\database1.mdb")
        con.Open()

        con.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim product_code As Integer
        product_code = TextBox1.Text

        con.Open()
        Dim da As OleDbDataAdapter
        Dim ds As New Data.DataSet
        da = New OleDbDataAdapter("select * from table3 where product_code=" & product_code, con)
        da.Fill(ds)

        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.Refresh()
        con.Close()
    End Sub
End Class