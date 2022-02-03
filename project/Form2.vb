Imports System.Data.OleDb
Public Class Form2
    Dim con As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader


    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=E:\Billing_software\project\database1.mdb")
        con.Open()
        cmd = New OleDbCommand("select MAX(user_account_id) from table1", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0) + 1

        End If
        DateTimePicker1.Value = Now.Date

        con.Close()
        con.Open()
        cmd = New OleDbCommand("select MAX(employeecode) from table1", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox3.Text = dr(0) + 1

        End If
        con.Close()


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim id As Integer
        Dim name, employeecode, designation, department, email, password, contact_no As String
        Dim date_of_joining As Date
        con.Open()
        id = TextBox1.Text
        name = TextBox2.Text
        employeecode = TextBox3.Text
        designation = ComboBox1.SelectedItem
        date_of_joining = DateTimePicker1.Value
        department = ComboBox2.SelectedItem
        email = TextBox4.Text
        contact_no = TextBox5.Text
        password = TextBox6.Text

        cmd = New OleDbCommand("insert into table1 values(" & id & ",' " & name & " ',' " & employeecode & " ' ,' " & designation & " ',' " & date_of_joining & " ',' " & department & " ',' " & email & " ',' " & contact_no & " ',' " & password & " ')", con)
        cmd.ExecuteNonQuery()
        MsgBox("New User Created")

        con.Close()



    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim id As Integer

        Dim cmd As OleDbCommand
        con.Open()
        id = TextBox1.Text
        cmd = New OleDbCommand("delete from table1 where user_account_id=" & id, con)
        cmd.ExecuteNonQuery()
        MsgBox("user removed")
        con.Close()

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim id As Integer
        Dim dr As OleDbDataReader

        con.Open()
        id = TextBox1.Text
        cmd = New OleDbCommand("select * from table1 where user_account_id= " & id, con)
        dr = cmd.ExecuteReader

        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0)
            TextBox2.Text = dr(1)
            TextBox3.Text = dr(2)
            ComboBox1.Text = dr(3)
            DateTimePicker1.Value = dr(4)
            ComboBox2.Text = dr(5)
            TextBox4.Text = dr(6)
            TextBox5.Text = dr(7)
            TextBox6.Text = dr(8)

        End If
        MsgBox("record selected")
        con.Close()

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click


        Dim id As Integer
        Dim name, employeecode, designation, department, email, password, contact_no As String
        Dim date_of_joining As Date
        id = TextBox1.Text
        name = TextBox2.Text
        employeecode = TextBox3.Text
        designation = ComboBox1.Text
        date_of_joining = DateTimePicker1.Value.Date
        department = ComboBox2.Text
        email = TextBox4.Text
        contact_no = TextBox5.Text
        password = TextBox6.Text
        con.Open()
        cmd = New OleDbCommand("update table1 set user_name= '" & name & "', employeecode='" & employeecode & "',designation='" & designation & "',date_of_joining='" & date_of_joining & "',department='" & department & "',email='" & email & "',contact_no='" & contact_no & "',Upassword='" & password & "' where user_account_id=" & id, con)
        cmd.ExecuteNonQuery()
        MsgBox("Updated")
        con.Close()

    End Sub

   
    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        con.Open()
        cmd = New OleDbCommand("select MAX(user_account_id) from table1", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0) + 1

        End If
        con.Close()
        con.Open()
        cmd = New OleDbCommand("select MAX(employeecode) from table1", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox3.Text = dr(0) + 1

        End If
        con.Close()

        DateTimePicker1.Value = Now.Date
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""

        ComboBox2.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""

    End Sub
End Class