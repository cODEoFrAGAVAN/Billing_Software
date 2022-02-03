Imports System.Data.OleDb

Public Class Form3
    Dim dr As OleDbDataReader
    Dim cmd As OleDbCommand
    Dim con As OleDbConnection
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim id As Integer
        Dim name, address, city, con_number, con_person As String

        Dim count As String
        Dim x As Integer
        Dim goods As String
        id = TextBox1.Text
        name = TextBox2.Text
        address = RichTextBox1.Text
        city = ComboBox1.SelectedItem
        con_number = TextBox3.Text
        con_person = TextBox4.Text
        goods = ""
        count = ListBox1.Items.Count
        For x = 0 To count - 1
            If (ListBox1.GetSelected(x)) Then

                goods = goods & vbCrLf & ListBox1.Items.Item(x)


            End If
        Next
        con.Open()
        cmd = New OleDbCommand("insert into Table2 values (" & id & ",'" & name & "','" & address & "','" & city & "','" & con_number & "','" & con_person & "','" & goods & "')", con)
        cmd.ExecuteNonQuery()
        MsgBox("profile created")
        con.Close()

    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=E:\Billing_software\project\database1.mdb")
        con.Open()
        cmd = New OleDbCommand("select MAX(supplier_id) from table2", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0) + 1

        End If
        con.Close()


    End Sub

    
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim id As Integer
        con.Open()

        id = TextBox1.Text
        cmd = New OleDbCommand("delete from table2 where supplier_id=" & id, con)
        cmd.ExecuteNonQuery()
        MsgBox("profile deleted")
        con.Close()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim id As Integer


        con.Open()
        id = TextBox1.Text
        cmd = New OleDbCommand("select * from table2 where supplier_id=" & id, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0)
            TextBox2.Text = dr(1)
            RichTextBox1.Text = dr(2)
            ComboBox1.Text = dr(3)
            TextBox3.Text = dr(4)
            TextBox4.Text = dr(5)
            RichTextBox2.Text = dr(6)
        End If

        con.Close()

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        TextBox2.Text = ""
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""

        ComboBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ListBox1.Text = ""

        con.Open()
        cmd = New OleDbCommand("select MAX(supplier_id) from table2", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0) + 1

        End If
        con.Close()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim id As Integer
        Dim name, address, city, con_number, con_person As String

        Dim count As String
        Dim x As Integer
        Dim goods As String
        id = TextBox1.Text
        name = TextBox2.Text
        address = RichTextBox1.Text
        city = ComboBox1.SelectedItem
        con_number = TextBox3.Text
        con_person = TextBox4.Text
        goods = ""
        count = ListBox1.Items.Count
        For x = 0 To count - 1
            If (ListBox1.GetSelected(x)) Then

                goods = goods & ListBox1.Items.Item(x)


            End If
        Next
        con.Open()
        cmd = New OleDbCommand("update  Table2 set supplier_name='" & name & "',contact_address='" & address & "',city='" & city & "',contact_number='" & con_number & "',contact_person_name='" & con_person & "',supplying_goods='" & goods & "'where supplier_id=" & id, con)

        cmd.ExecuteNonQuery()
        MsgBox("profile updated")
        con.Close()

    End Sub
End Class