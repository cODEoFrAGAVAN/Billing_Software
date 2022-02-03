Imports System.Data.OleDb

Public Class Form5
    Dim dr As OleDbDataReader
    Dim cmd As OleDbCommand
    Dim con As OleDbConnection


    Private Sub Form5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
        con = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=E:\Billing_software\project\database1.mdb")
        con.Open()
        cmd = New OleDbCommand("select MAX(transaction_id) from table4", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0) + 1

        End If

        TextBox2.Text = Now.Date
        
        con.Close()
        TextBox4.Text = 0
        TextBox5.Text = 0
        TextBox6.Text = 0
        TextBox7.Text = 0
        TextBox8.Text = 0


    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim transaction_id, supplier_id, product_code, purchase_qty, purchase_cost, gst, total, total_cost As Integer
        Dim name_of_the_product As String
        Dim tdate As Date
        transaction_id = TextBox1.Text
        supplier_id = ComboBox1.SelectedItem
        product_code = ComboBox2.SelectedItem
        name_of_the_product = TextBox3.Text
        purchase_qty = TextBox4.Text
        purchase_cost = TextBox5.Text
        gst = TextBox6.Text
        total = TextBox7.Text
        total_cost = TextBox8.Text
        tdate = TextBox2.Text
        con.Open()

        cmd = New OleDbCommand("insert into table4 values(" & transaction_id & ",'" & tdate & "'," & supplier_id & "," & product_code & ",'" & name_of_the_product & "'," & purchase_qty & "," & purchase_cost & "," & gst & "," & total & "," & total_cost & ")", con)
        cmd.ExecuteNonQuery()
        cmd = New OleDbCommand("update table3 set stock=stock+" & purchase_qty & " where product_code=" & product_code, con)
        cmd.ExecuteNonQuery()

        MsgBox("Stock Transaction Successfully Completed")
        con.Close()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        TextBox2.Text = ""

        ComboBox1.Text = ""
        ComboBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = 0
        TextBox5.Text = 0
        TextBox6.Text = 0
        TextBox7.Text = 0
        TextBox8.Text = 0
        con.Open()
        cmd = New OleDbCommand("select MAX(transaction_id) from table4", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0) + 1

        End If

        TextBox2.Text = Now.Date
        con.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim transaction_id As Integer
        con.Open()

        transaction_id = TextBox1.Text
        cmd = New OleDbCommand("delete from table4 where transaction_id=" & transaction_id, con)
        cmd.ExecuteNonQuery()
        MsgBox("transaction deleted")
        con.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim transaction_id As Integer

        con.Open()
        transaction_id = TextBox1.Text
        cmd = New OleDbCommand("select * from table4 where transaction_id=" & transaction_id, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0)
            TextBox2.Text = dr(1)
            ComboBox1.Text = dr(2)
            ComboBox2.Text = dr(3)
            TextBox3.Text = dr(4)
            TextBox4.Text = dr(5)
            TextBox5.Text = dr(6)
            TextBox6.Text = dr(7)
            TextBox7.Text = dr(8)
            TextBox8.Text = dr(9)
        End If

        con.Close()

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim transaction_id, supplier_id, product_code, purchase_qty, purchase_cost, gst, total, total_cost As Integer
        Dim name_of_the_product As String
        Dim tdate As Date
        transaction_id = TextBox1.Text
        supplier_id = ComboBox1.SelectedItem
        product_code = ComboBox2.SelectedItem
        name_of_the_product = TextBox3.Text
        purchase_qty = TextBox4.Text
        purchase_cost = TextBox5.Text
        gst = TextBox6.Text
        total = TextBox7.Text
        total_cost = TextBox8.Text
        tdate = TextBox2.Text
        con.Open()

        cmd = New OleDbCommand("update table4 set tdate='" & tdate & "', supplier_id=" & supplier_id & ", product_code=" & product_code & ", name_of_the_product='" & name_of_the_product & "', purchase_qty=" & purchase_qty & ", purchase_cost=" & purchase_cost & ", gst=" & gst & ", total=" & total & ", total_cost=" & total_cost & " where  transaction_id= " & transaction_id, con)

        cmd.ExecuteNonQuery()
        MsgBox("transaction updated")
        con.Close()

    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim purchase_cost, gst, total_price, qty As Integer
        qty = TextBox4.Text
        purchase_cost = TextBox5.Text
        gst = TextBox6.Text
        total_price = purchase_cost + (purchase_cost * (gst / 100))
        TextBox7.Text = total_price
        TextBox8.Text = total_price * qty
        

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim product_code As Integer
        product_code = ComboBox2.Text
        con.Open()
        cmd = New OleDbCommand("select * from table3 where product_code=" & product_code, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox3.Text = dr(1)

        End If

        con.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class