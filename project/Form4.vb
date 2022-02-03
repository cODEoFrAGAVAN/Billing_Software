Imports System.Data.OleDb

Public Class Form4
    Dim dr As OleDbDataReader
    Dim cmd As OleDbCommand
    Dim con As OleDbConnection
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim product_code, supplier_id, purchase_cost, gst, total_price, student_price As Integer

        Dim name_of_product, name_of_supplier As String
        Dim stock As Integer

        product_code = TextBox1.Text
        name_of_product = TextBox2.Text
        supplier_id = ComboBox1.Text
        name_of_supplier = TextBox3.Text
        purchase_cost = TextBox4.Text
        gst = TextBox5.Text
        total_price = TextBox6.Text
        student_price = TextBox7.Text
        stock = TextBox8.Text

        con.Open()

        cmd = New OleDbCommand("insert into Table3 values (" & product_code & ",'" & name_of_product & "'," & supplier_id & ",'" & name_of_supplier & "'," & purchase_cost & "," & gst & "," & total_price & "," & student_price & "," & stock & ")", con)
        cmd.ExecuteNonQuery()

        MsgBox("product registered")

        con.Close()

    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=E:\Billing_software\project\database1.mdb")
        con.Open()
        cmd = New OleDbCommand("select MAX(product_code) from table3", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0) + 1

        End If

        con.Close()
        TextBox4.Text = 0
        TextBox5.Text = 0
        TextBox6.Text = 0
        TextBox7.Text = 0
        TextBox8.Text = 0


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        TextBox2.Text = ""
        ComboBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = 0
        TextBox5.Text = 0
        TextBox6.Text = 0
        TextBox7.Text = 0
        TextBox8.Text = 0
        con.Open()
        cmd = New OleDbCommand("select MAX(product_code) from table3", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0) + 1

        End If

        con.Close()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim product_code As Integer
        con.Open()

        product_code = TextBox1.Text
        cmd = New OleDbCommand("delete from table3 where product_code=" & product_code, con)
        cmd.ExecuteNonQuery()
        MsgBox("product deleted")
        con.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim product_code As Integer

        con.Open()
        product_code = TextBox1.Text
        cmd = New OleDbCommand("select * from table3 where product_code=" & product_code, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0)
            TextBox2.Text = dr(1)
            ComboBox1.Text = dr(2)
            TextBox3.Text = dr(3)
            TextBox4.Text = dr(4)
            TextBox5.Text = dr(5)
            TextBox6.Text = dr(6)
            TextBox7.Text = dr(7)
            TextBox8.Text = dr(8)

        End If

        con.Close()

    End Sub

    
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim product_code, supplier_id, purchase_cost, gst, total_price, student_price As Integer

        Dim name_of_product, name_of_supplier As String

        Dim stock As Integer

        product_code = TextBox1.Text
        name_of_product = TextBox2.Text
        supplier_id = ComboBox1.Text
        name_of_supplier = TextBox3.Text
        purchase_cost = TextBox4.Text
        gst = TextBox5.Text
        total_price = TextBox6.Text
        student_price = TextBox7.Text
        stock = TextBox8.Text

        con.Open()
        cmd = New OleDbCommand("update table3 set name_of_product='" & name_of_product & "', supplier_id=" & supplier_id & ", name_of_supplier='" & name_of_supplier & "', purchase_cost=" & purchase_cost & ", gst=" & gst & ", total_price=" & total_price & ", student_price=" & student_price & ", stock=" & stock & " where product_code= " & product_code, con)
        cmd.ExecuteNonQuery()
        MsgBox("product deatils updated")
        con.Close()

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim purchase_cost, gst, total_price As Integer
        purchase_cost = TextBox4.Text
        gst = TextBox5.Text
        total_price = purchase_cost + (purchase_cost * gst / 100)
        TextBox6.Text = total_price
        
    End Sub

   
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim supplier_id As Integer
        supplier_id = ComboBox1.Text
        con.Open()
        cmd = New OleDbCommand("select * from table2 where supplier_id=" & supplier_id, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox3.Text = dr(1)

        End If

        con.Close()

    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub
End Class