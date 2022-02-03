Imports System.Data.OleDb
Public Class Form10
    Dim con As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim register_no As Integer
        Dim full_name, department, shift As String
        Dim year_start, year_end As String


        register_no = TextBox1.Text
        full_name = TextBox2.Text
        department = ComboBox1.Text
        shift = ComboBox2.Text
        year_start = ComboBox3.Text
        year_end = ComboBox4.Text

        con.Open()
        cmd = New OleDbCommand("insert into table5 values(" & register_no & ",'" & full_name & "','" & department & "','" & shift & "','" & year_start & "','" & year_end & "')", con)
        cmd.ExecuteNonQuery()

        MsgBox("Registered")
        con.Close()

    End Sub

    Private Sub Form10_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=E:\Billing_software\project\database1.mdb")
        con.Open()
        con.Close()
        ComboBox3.Text = Now.Year
        ComboBox4.Text = ""


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""

        ComboBox2.Text = ""
        ComboBox3.Text = Now.Year
        ComboBox4.Text = ""



    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim register_no As Integer

        con.Open()
        register_no = TextBox1.Text
        cmd = New OleDbCommand("select * from table5 where  register_no= " & register_no, con)
        dr = cmd.ExecuteReader

        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0)
            TextBox2.Text = dr(1)
            ComboBox1.Text = dr(2)
            ComboBox2.Text = dr(3)
            ComboBox3.Text = dr(4)
            ComboBox4.Text = dr(5)

        End If


        con.Close()

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim register_no As Integer

        con.Open()
        register_no = TextBox1.Text
        cmd = New OleDbCommand("delete from table5 where register_no=" & register_no, con)
        cmd.ExecuteNonQuery()

        MsgBox("deleted")
        con.Close()

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim register_no As Long
        Dim full_name, department, shift As String
        Dim year_start, year_end As String


        register_no = TextBox1.Text
        full_name = TextBox2.Text
        department = ComboBox1.SelectedItem
        shift = ComboBox2.SelectedItem
        year_start = ComboBox3.SelectedItem
        year_end = ComboBox4.SelectedItem
        con.Open()
        cmd = New OleDbCommand("update table5 set full_name='" & full_name & "', department='" & department & "', shift='" & shift & "', year_start='" & year_start & "', year_end='" & year_end & "'where register_no=" & register_no, con)
        cmd.ExecuteNonQuery()
        MsgBox("updated")

        con.Close()


    End Sub
End Class