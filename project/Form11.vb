Imports System.Data.OleDb

Public Class Form11
    Dim dr As OleDbDataReader
    Dim cmd As OleDbCommand
    Dim con As OleDbConnection
    Private Sub Form11_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=E:\Billing_software\project\database1.mdb")
        con.Open()
        cmd = New OleDbCommand("select MAX(bill_no) from table6", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0) + 1

        End If
        con.Close()
        TextBox6.Text = 0
        TextBox10.Text = 0
        TextBox14.Text = 0
        TextBox18.Text = 0
        TextBox22.Text = 0
        TextBox26.Text = 0

        TextBox8.Text = 0
        TextBox12.Text = 0
        TextBox16.Text = 0
        TextBox20.Text = 0
        TextBox24.Text = 0
        TextBox28.Text = 0

        ComboBox1.Text = 0
        ComboBox2.Text = 0
        ComboBox3.Text = 0
        ComboBox4.Text = 0
        ComboBox5.Text = 0
        ComboBox6.Text = 0

        TextBox9.Text = 0
        TextBox13.Text = 0
        TextBox17.Text = 0
        TextBox21.Text = 0
        TextBox25.Text = 0
        TextBox29.Text = 0

        TextBox30.Text = 0
        TextBox31.Text = 0
        TextBox32.Text = 0

        TextBox2.Text = Now.Date


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim bill_no As Integer
        Dim stu_reg_no As Integer

        Dim cur_date, stu_name, department As String
        Dim name_1, name_2, name_3, name_4, name_5, name_6 As String
        Dim product_code_1, product_code_2, product_code_3, product_code_4, product_code_5, product_code_6 As Integer
        Dim price_1, price_2, price_3, price_4, price_5, price_6 As Integer
        Dim quantity_1, quantity_2, quantity_3, quantity_4, quantity_5, quantity_6 As Integer
        Dim total_1, total_2, total_3, total_4, total_5, total_6 As Integer
        Dim grand_total, recived_amount, balance_amount As Integer
        bill_no = TextBox1.Text
        cur_date = TextBox2.Text
        stu_reg_no = TextBox3.Text
        department = TextBox4.Text
        stu_name = TextBox5.Text

        product_code_1 = TextBox6.Text
        product_code_2 = TextBox10.Text
        product_code_3 = TextBox14.Text
        product_code_4 = TextBox18.Text
        product_code_5 = TextBox22.Text
        product_code_6 = TextBox26.Text

        name_1 = TextBox7.Text
        name_2 = TextBox11.Text
        name_3 = TextBox15.Text
        name_4 = TextBox19.Text
        name_5 = TextBox23.Text
        name_6 = TextBox27.Text

        price_1 = TextBox8.Text
        price_2 = TextBox12.Text
        price_3 = TextBox16.Text
        price_4 = TextBox20.Text
        price_5 = TextBox24.Text
        price_6 = TextBox28.Text

        quantity_1 = ComboBox1.Text
        quantity_2 = ComboBox2.Text
        quantity_3 = ComboBox3.Text
        quantity_4 = ComboBox4.Text
        quantity_5 = ComboBox5.Text
        quantity_6 = ComboBox6.Text

        total_1 = TextBox9.Text
        total_2 = TextBox13.Text
        total_3 = TextBox17.Text
        total_4 = TextBox21.Text
        total_5 = TextBox25.Text
        total_6 = TextBox29.Text

        grand_total = TextBox30.Text
        recived_amount = TextBox31.Text
        balance_amount = TextBox32.Text
        con.Open()
        cmd = New OleDbCommand("insert into table6 values(" & bill_no & ",'" & cur_date & "'," & stu_reg_no & ",'" & department & "','" & stu_name & "'," & product_code_1 & ",'" & name_1 & "'," & price_1 & "," & quantity_1 & "," & total_1 & "," & product_code_2 & ",'" & name_2 & "'," & price_2 & "," & quantity_2 & "," & total_2 & "," & product_code_3 & ",'" & name_3 & "'," & price_3 & "," & quantity_3 & "," & total_3 & "," & product_code_4 & ",'" & name_4 & "'," & price_4 & "," & quantity_4 & "," & total_4 & "," & product_code_5 & ",'" & name_5 & "'," & price_5 & "," & quantity_5 & "," & total_5 & "," & product_code_6 & ",'" & name_6 & "'," & price_6 & "," & quantity_6 & "," & total_6 & "," & grand_total & "," & recived_amount & "," & balance_amount & ")", con)
        cmd.ExecuteNonQuery()
        con.Close()

        con.Open()

        cmd = New OleDbCommand("update table3 set stock=stock-" & quantity_1 & " where product_code=" & product_code_1, con)
        cmd.ExecuteNonQuery()
        con.Close()

        con.Open()
        cmd = New OleDbCommand("update table3 set stock=stock-" & quantity_2 & " where product_code=" & product_code_2, con)
        cmd.ExecuteNonQuery()
        con.Close()

        con.Open()
        cmd = New OleDbCommand("update table3 set stock=stock-" & quantity_3 & " where product_code=" & product_code_3, con)
        cmd.ExecuteNonQuery()
        con.Close()

        con.Open()
        cmd = New OleDbCommand("update table3 set stock=stock-" & quantity_4 & " where product_code=" & product_code_4, con)
        cmd.ExecuteNonQuery()
        con.Close()

        con.Open()
        cmd = New OleDbCommand("update table3 set stock=stock-" & quantity_5 & " where product_code=" & product_code_5, con)
        cmd.ExecuteNonQuery()
        con.Close()

        con.Open()
        cmd = New OleDbCommand("update table3 set stock=stock-" & quantity_6 & " where product_code=" & product_code_6, con)
        cmd.ExecuteNonQuery()
        con.Close()

        MsgBox("amount added")
        con.Close()

        

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim price_1, price_2, price_3, price_4, price_5, price_6 As Integer
        Dim quantity_1, quantity_2, quantity_3, quantity_4, quantity_5, quantity_6 As Integer
        Dim total_1, total_2, total_3, total_4, total_5, total_6 As Integer
        Dim grand_total, recived_amount, balance_amount As Integer

        price_1 = Val(TextBox8.Text)

        price_2 = Val(TextBox12.Text)
        price_3 = Val(TextBox16.Text)
        price_4 = Val(TextBox20.Text)
        price_5 = Val(TextBox24.Text)
        price_6 = Val(TextBox28.Text)

        quantity_1 = ComboBox1.Text
        quantity_2 = ComboBox2.Text
        quantity_3 = ComboBox3.Text

        quantity_4 = ComboBox4.Text
        quantity_5 = ComboBox5.Text
        quantity_6 = ComboBox6.Text


        grand_total = Val(TextBox30.Text)
        recived_amount = Val(TextBox31.Text)
     
        total_1 = price_1 * quantity_1
        total_2 = price_2 * quantity_2
        total_3 = price_3 * quantity_3
        total_4 = price_4 * quantity_4
        total_5 = price_5 * quantity_5
        total_6 = price_6 * quantity_6

        grand_total = total_1 + total_2 + total_3 + total_4 + total_5 + total_6
        balance_amount = grand_total - recived_amount
        TextBox30.Text = grand_total

        TextBox9.Text = total_1
        TextBox13.Text = total_2
        TextBox17.Text = total_3
        TextBox21.Text = total_4
        TextBox25.Text = total_5
        TextBox29.Text = total_6

        TextBox32.Text = balance_amount
        TextBox31.Text = recived_amount


    End Sub

    Private Sub TextBox32_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox32.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        con.Open()
        cmd = New OleDbCommand("select MAX(bill_no) from table6", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0) + 1

        End If
        con.Close()
        TextBox6.Text = 0
        TextBox10.Text = 0
        TextBox14.Text = 0
        TextBox18.Text = 0
        TextBox22.Text = 0
        TextBox26.Text = 0

        TextBox8.Text = 0
        TextBox12.Text = 0
        TextBox16.Text = 0
        TextBox20.Text = 0
        TextBox24.Text = 0
        TextBox28.Text = 0

        ComboBox1.Text = 0
        ComboBox2.Text = 0
        ComboBox3.Text = 0
        ComboBox4.Text = 0
        ComboBox5.Text = 0
        ComboBox6.Text = 0

        TextBox9.Text = 0
        TextBox13.Text = 0
        TextBox17.Text = 0
        TextBox21.Text = 0
        TextBox25.Text = 0
        TextBox29.Text = 0

        TextBox30.Text = 0
        TextBox31.Text = 0
        TextBox32.Text = 0

        TextBox2.Text = Now.Date
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox7.Text = ""
        TextBox11.Text = ""
        TextBox15.Text = ""
        TextBox19.Text = ""
        TextBox23.Text = ""
        TextBox27.Text = ""

    End Sub

    

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim bill_no As Integer


        con.Open()
        bill_no = TextBox1.Text
        cmd = New OleDbCommand("select * from table6 where bill_no=" & bill_no, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0)
            TextBox2.Text = dr(1)
            TextBox3.Text = dr(2)
            TextBox4.Text = dr(3)
            TextBox5.Text = dr(4)
            TextBox6.Text = dr(5)
            TextBox7.Text = dr(6)
            TextBox8.Text = dr(7)
            ComboBox1.Text = dr(8)
            TextBox9.Text = dr(9)
            TextBox10.Text = dr(10)
            TextBox11.Text = dr(11)
            TextBox12.Text = dr(12)
            ComboBox2.Text = dr(13)
            TextBox13.Text = dr(14)
            TextBox14.Text = dr(15)
            TextBox15.Text = dr(16)
            TextBox16.Text = dr(17)
            ComboBox3.Text = dr(18)
            TextBox17.Text = dr(19)
            TextBox18.Text = dr(20)
            TextBox19.Text = dr(21)
            TextBox20.Text = dr(22)
            ComboBox4.Text = dr(23)
            TextBox21.Text = dr(24)
            TextBox22.Text = dr(25)
            TextBox23.Text = dr(26)
            TextBox24.Text = dr(27)
            ComboBox5.Text = dr(28)
            TextBox25.Text = dr(29)
            TextBox26.Text = dr(30)
            TextBox27.Text = dr(31)
            TextBox28.Text = dr(32)
            ComboBox6.Text = dr(33)
            TextBox29.Text = dr(34)
            TextBox30.Text = dr(35)
            TextBox31.Text = dr(36)
            TextBox32.Text = dr(37)


        End If


        con.Close()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim product_code As Integer


        con.Open()
        product_code = TextBox6.Text
        cmd = New OleDbCommand("select * from table3 where product_code=" & product_code, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox6.Text = dr(0)
            TextBox7.Text = dr(1)
            TextBox8.Text = dr(7)

        End If

        con.Close()

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim product_code As Integer


        con.Open()
        product_code = TextBox10.Text
        cmd = New OleDbCommand("select * from table3 where product_code=" & product_code, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox10.Text = dr(0)
            TextBox11.Text = dr(1)
            TextBox12.Text = dr(7)
        End If
        con.Close()

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim product_code As Integer


        con.Open()
        product_code = TextBox14.Text
        cmd = New OleDbCommand("select * from table3 where product_code=" & product_code, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox14.Text = dr(0)
            TextBox15.Text = dr(1)
            TextBox16.Text = dr(7)
        End If
        con.Close()

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim product_code As Integer


        con.Open()
        product_code = TextBox18.Text
        cmd = New OleDbCommand("select * from table3 where product_code=" & product_code, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox18.Text = dr(0)
            TextBox19.Text = dr(1)
            TextBox20.Text = dr(7)
        End If
        con.Close()

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim product_code As Integer


        con.Open()
        product_code = TextBox22.Text
        cmd = New OleDbCommand("select * from table3 where product_code=" & product_code, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox22.Text = dr(0)
            TextBox23.Text = dr(1)
            TextBox24.Text = dr(7)
        End If
        con.Close()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim product_code As Integer


        con.Open()
        product_code = TextBox26.Text
        cmd = New OleDbCommand("select * from table3 where product_code=" & product_code, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox26.Text = dr(0)
            TextBox27.Text = dr(1)
            TextBox28.Text = dr(7)
        End If
        con.Close()

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim register_no As Integer
        register_no = TextBox3.Text
        con.Open()
        cmd = New OleDbCommand("select * from table5 where register_no=" & register_no, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox3.Text = dr(0)
            TextBox4.Text = dr(2)
            TextBox5.Text = dr(1)

        End If
        con.Close()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim bill_no As Integer
        Dim stu_reg_no As Integer

        Dim cur_date, stu_name, department As String
        Dim name_1, name_2, name_3, name_4, name_5, name_6 As String
        Dim product_code_1, product_code_2, product_code_3, product_code_4, product_code_5, product_code_6 As Integer
        Dim price_1, price_2, price_3, price_4, price_5, price_6 As Integer
        Dim quantity_1, quantity_2, quantity_3, quantity_4, quantity_5, quantity_6 As Integer
        Dim total_1, total_2, total_3, total_4, total_5, total_6 As Integer
        Dim grand_total, recived_amount, balance_amount As Integer
        bill_no = TextBox1.Text
        cur_date = TextBox2.Text
        stu_reg_no = TextBox3.Text
        department = TextBox4.Text
        stu_name = TextBox5.Text

        product_code_1 = TextBox6.Text
        product_code_2 = TextBox10.Text
        product_code_3 = TextBox14.Text
        product_code_4 = TextBox18.Text
        product_code_5 = TextBox22.Text
        product_code_6 = TextBox26.Text

        name_1 = TextBox7.Text
        name_2 = TextBox11.Text
        name_3 = TextBox15.Text
        name_4 = TextBox19.Text
        name_5 = TextBox23.Text
        name_6 = TextBox27.Text

        price_1 = TextBox8.Text
        price_2 = TextBox12.Text
        price_3 = TextBox16.Text
        price_4 = TextBox20.Text
        price_5 = TextBox24.Text
        price_6 = TextBox28.Text

        quantity_1 = ComboBox1.Text
        quantity_2 = ComboBox2.Text
        quantity_3 = ComboBox3.Text
        quantity_4 = ComboBox4.Text
        quantity_5 = ComboBox5.Text
        quantity_6 = ComboBox6.Text

        total_1 = TextBox9.Text
        total_2 = TextBox13.Text
        total_3 = TextBox17.Text
        total_4 = TextBox21.Text
        total_5 = TextBox25.Text
        total_6 = TextBox29.Text

        grand_total = TextBox30.Text
        recived_amount = TextBox31.Text
        balance_amount = TextBox32.Text
        con.Open()
        cmd = New OleDbCommand("update table6 set stu_reg_no=" & stu_reg_no & ", department='" & department & "', stu_name='" & stu_name & "', product_code_1=" & product_code_1 & ", name_1='" & name_1 & "', price_1=" & price_1 & ", quantity_1=" & quantity_1 & ", total_1=" & total_1 & ", product_code_2=" & product_code_2 & ", name_2='" & name_2 & "', price_2=" & price_2 & ", quantity_2=" & quantity_2 & ", total_2=" & total_2 & ", product_code_3=" & product_code_3 & ", name_3='" & name_3 & "', price_3=" & price_3 & ", quantity_3=" & quantity_3 & ", total_3=" & total_3 & ", product_code_4=" & product_code_4 & ", name_4='" & name_4 & "', price_4=" & price_4 & ", quantity_4=" & quantity_4 & ", total_4=" & total_4 & ", product_code_5=" & product_code_5 & ", name_5='" & name_5 & "', price_5=" & price_5 & ", quantity_5=" & quantity_5 & ", total_5=" & total_5 & ", product_code_6=" & product_code_6 & ", name_6='" & name_6 & "', price_6=" & price_6 & ", quantity_6=" & quantity_6 & ", total_6=" & total_6 & ", grand_total=" & grand_total & ", recived_amount=" & recived_amount & ", balance_amount=" & balance_amount & " where bill_no=" & bill_no, con)
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
End Class