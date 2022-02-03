Imports System.Data.OleDb
Public Class Form14
    Dim dr As OleDbDataReader
    Dim cmd As OleDbCommand
    Dim con As OleDbConnection
    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim product_code As Integer
        con.Open()
        product_code = TextBox12.Text
        cmd = New OleDbCommand("select * from table3 where product_code=" & product_code, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox12.Text = dr(0)
            TextBox13.Text = dr(1)
            TextBox14.Text = dr(7)

        End If

        con.Close()
    End Sub



    Private Sub Form14_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=E:\Billing_software\project\database1.mdb")
        con.Open()
        cmd = New OleDbCommand("select MAX(bill_no) from table7", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0) + 1

        End If
        con.Close()
        TextBox2.Text = Now.Date
        TextBox4.Text = 0
        TextBox5.Text = 0
        TextBox6.Text = 0
        TextBox8.Text = 0
        TextBox9.Text = 0
        TextBox10.Text = 0
        TextBox11.Text = 0
        TextBox14.Text = 0
        TextBox15.Text = 0
        TextBox18.Text = 0
        TextBox19.Text = 0
        TextBox22.Text = 0
        TextBox23.Text = 0
        ComboBox2.Text = 0
        ComboBox3.Text = 0
        ComboBox4.Text = 0
        TextBox24.Text = 0
        TextBox25.Text = 0
        TextBox26.Text = 0
        TextBox27.Text = 0
        TextBox16.Text = 0
        TextBox20.Text = 0
        TextBox12.Text = 0
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim price_1, price_2 As Integer
        Dim qty1, qty2 As Integer
        Dim total_1, total_2, loc As Integer
        price_1 = TextBox4.Text
        qty1 = TextBox5.Text
        price_2 = TextBox8.Text
        qty2 = TextBox9.Text
        total_1 = price_1 * qty1
        total_2 = price_2 * qty2
        loc = total_1 + total_2
        TextBox11.Text = loc
        TextBox6.Text = total_1
        TextBox10.Text = total_2

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim ptotal1, ptotal2, ptotal3, loc As Integer
        Dim pqty1, pqty2, pqty3 As Integer
        Dim pprice1, pprice2, pprice3 As Integer
        pprice1 = TextBox14.Text
        pprice2 = TextBox18.Text
        pprice3 = TextBox22.Text
        pqty1 = ComboBox2.Text
        pqty2 = ComboBox3.Text
        pqty3 = ComboBox4.Text
        ptotal1 = pprice1 * pqty1
        ptotal2 = pprice2 * pqty2
        ptotal3 = pprice3 * pqty3

        loc = ptotal1 + ptotal2 + ptotal3
        TextBox24.Text = loc
        TextBox15.Text = ptotal1
        TextBox19.Text = ptotal2
        TextBox23.Text = ptotal3

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim grand_total1, grand_total2, total_amount As Integer
        grand_total1 = TextBox11.Text
        grand_total2 = TextBox24.Text
        total_amount = grand_total1 + grand_total2
        TextBox25.Text = total_amount

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim total_amount, recived_amount, balance_amount As Integer
        recived_amount = TextBox26.Text
        total_amount = TextBox25.Text
        balance_amount = total_amount - recived_amount
        TextBox27.Text = balance_amount

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim product_code As Integer
        con.Open()
        product_code = TextBox16.Text
        cmd = New OleDbCommand("select * from table3 where product_code=" & product_code, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox16.Text = dr(0)
            TextBox17.Text = dr(1)
            TextBox18.Text = dr(7)

        End If

        con.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim product_code As Integer
        con.Open()
        product_code = TextBox20.Text
        cmd = New OleDbCommand("select * from table3 where product_code=" & product_code, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox20.Text = dr(0)
            TextBox21.Text = dr(1)
            TextBox22.Text = dr(7)

        End If

        con.Close()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        con.Open()
        cmd = New OleDbCommand("select MAX(bill_no) from table7", con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0) + 1

        End If
        con.Close()
        TextBox2.Text = Now.Date
        TextBox4.Text = 0
        TextBox5.Text = 0
        TextBox6.Text = 0
        TextBox8.Text = 0
        TextBox9.Text = 0
        TextBox10.Text = 0
        TextBox11.Text = 0
        TextBox14.Text = 0
        TextBox15.Text = 0
        TextBox18.Text = 0
        TextBox19.Text = 0
        TextBox22.Text = 0
        TextBox23.Text = 0
        ComboBox2.Text = 0
        ComboBox3.Text = 0
        ComboBox4.Text = 0
        TextBox24.Text = 0
        TextBox25.Text = 0
        TextBox26.Text = 0
        TextBox27.Text = 0

        TextBox3.Text = ""
        TextBox7.Text = ""
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        ComboBox1.Text = ""
        TextBox16.Text = 0
        TextBox12.Text = 0
        TextBox20.Text = 0
        TextBox13.Text = ""
        TextBox17.Text = ""
        TextBox21.Text = ""

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim billno As Integer
        con.Open()
        billno = TextBox1.Text
        cmd = New OleDbCommand("select * from table7 where bill_no=" & billno, con)
        dr = cmd.ExecuteReader
        If (dr.HasRows) Then
            dr.Read()
            TextBox1.Text = dr(0)
            TextBox2.Text = dr(1)
            ComboBox1.Text = dr(2)
            If (dr(3) = "Single-Side") Then
                CheckBox1.Checked = True
            End If
            TextBox3.Text = dr(4)
            TextBox4.Text = dr(5)
            TextBox5.Text = dr(6)
            TextBox6.Text = dr(7)
            If (dr(8) = "Double-Side") Then
                CheckBox2.Checked = True
            End If
            TextBox7.Text = dr(9)
            TextBox8.Text = dr(10)
            TextBox9.Text = dr(11)
            TextBox10.Text = dr(12)
            TextBox11.Text = dr(13)
            TextBox12.Text = dr(14)
            TextBox13.Text = dr(15)
            TextBox14.Text = dr(16)
            ComboBox2.Text = dr(17)
            TextBox15.Text = dr(18)
            TextBox16.Text = dr(19)
            TextBox17.Text = dr(20)
            TextBox18.Text = dr(21)
            ComboBox3.Text = dr(22)
            TextBox19.Text = dr(23)
            TextBox20.Text = dr(24)
            TextBox21.Text = dr(25)
            TextBox22.Text = dr(26)
            ComboBox4.Text = dr(27)
            TextBox23.Text = dr(28)
            TextBox24.Text = dr(29)
            TextBox25.Text = dr(30)
            TextBox26.Text = dr(31)
            TextBox27.Text = dr(32)

        End If

        con.Close()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim billno As Integer
        Dim department, udate, type_1, type_2 As String
        type_1 = ""
        type_2 = ""
        Dim desc1, desc2 As String
        Dim price_1, price_2 As Integer
        Dim qty1, qty2 As Integer
        Dim total_1, total_2 As Integer
        Dim grand_total1 As Integer

        Dim ptotal1, ptotal2, ptotal3, loc As Integer
        Dim pqty1, pqty2, pqty3 As Integer
        Dim pprice1, pprice2, pprice3 As Integer
        Dim product_code1, product_code2, product_code3 As Integer
        Dim pname1, pname2, pname3 As String
        Dim grand_total2 As Integer

        Dim total_amount, recived_amount, balance_amount As Integer


        billno = TextBox1.Text
        department = ComboBox1.Text
        udate = TextBox2.Text

        If (CheckBox1.Checked = True) Then
            type_1 = "Single-Side"
        End If
        If (CheckBox2.Checked = True) Then
            type_2 = "Double-Side"
        End If
        desc1 = TextBox3.Text
        desc2 = TextBox7.Text
        price_1 = TextBox4.Text
        qty1 = TextBox5.Text
        price_2 = TextBox8.Text
        qty2 = TextBox9.Text
        total_1 = TextBox6.Text
        total_2 = TextBox10.Text
        grand_total1 = TextBox11.Text

        product_code1 = TextBox12.Text
        product_code2 = TextBox16.Text
        product_code3 = TextBox20.Text
        pname1 = TextBox13.Text
        pname2 = TextBox17.Text
        pname3 = TextBox21.Text
        pprice1 = TextBox14.Text
        pprice2 = TextBox18.Text
        pprice3 = TextBox22.Text
        pqty1 = ComboBox2.Text
        pqty2 = ComboBox3.Text
        pqty3 = ComboBox4.Text
        ptotal1 = TextBox15.Text
        ptotal2 = TextBox19.Text
        ptotal3 = TextBox23.Text
        grand_total2 = TextBox24.Text

        total_amount = TextBox25.Text
        recived_amount = TextBox26.Text
        balance_amount = TextBox27.Text
        con.Open()
        cmd = New OleDbCommand("insert into table7 values(" & billno & ",'" & udate & "','" & department & "','" & type_1 & "','" & desc1 & "'," & price_1 & "," & qty1 & "," & total_1 & ",'" & type_2 & "','" & desc2 & "'," & price_2 & "," & qty2 & "," & total_2 & "," & grand_total1 & "," & product_code1 & ",'" & pname1 & "'," & pprice1 & "," & pqty1 & "," & ptotal1 & "," & product_code2 & ",'" & pname2 & "'," & pprice2 & "," & pqty2 & "," & ptotal2 & "," & product_code3 & ",'" & pname3 & "'," & pprice3 & "," & pqty3 & "," & ptotal3 & "," & grand_total2 & "," & total_amount & "," & recived_amount & "," & balance_amount & ")", con)
        cmd.ExecuteNonQuery()
        MsgBox("AmountAdded")
        con.Close()

        con.Open()
        cmd = New OleDbCommand("update table3 set stock=stock-" & pqty1 & " where product_code=" & product_code1, con)
        cmd.ExecuteNonQuery()
        con.Close()

        con.Open()
        cmd = New OleDbCommand("update table3 set stock=stock-" & pqty2 & " where product_code=" & product_code2, con)
        cmd.ExecuteNonQuery()
        con.Close()

        con.Open()
        cmd = New OleDbCommand("update table3 set stock=stock-" & pqty3 & " where product_code=" & product_code3, con)
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim billno As Integer
        Dim department, udate, type_1, type_2 As String
        type_1 = ""
        type_2 = ""
        Dim desc1, desc2 As String
        Dim price_1, price_2 As Integer
        Dim qty1, qty2 As Integer
        Dim total_1, total_2 As Integer
        Dim grand_total1 As Integer

        Dim ptotal1, ptotal2, ptotal3, loc As Integer
        Dim pqty1, pqty2, pqty3 As Integer
        Dim pprice1, pprice2, pprice3 As Integer
        Dim product_code1, product_code2, product_code3 As Integer
        Dim pname1, pname2, pname3 As String
        Dim grand_total2 As Integer

        Dim total_amount, recived_amount, balance_amount As Integer


        billno = TextBox1.Text
        department = ComboBox1.Text
        udate = TextBox2.Text

        If (CheckBox1.Checked = True) Then
            type_1 = "Single-Side"
        End If
        If (CheckBox2.Checked = True) Then
            type_2 = "Double-Side"
        End If
        desc1 = TextBox3.Text
        desc2 = TextBox7.Text
        price_1 = TextBox4.Text
        qty1 = TextBox5.Text
        price_2 = TextBox8.Text
        qty2 = TextBox9.Text
        total_1 = TextBox6.Text
        total_2 = TextBox10.Text
        grand_total1 = TextBox11.Text

        product_code1 = TextBox12.Text
        product_code2 = TextBox16.Text
        product_code3 = TextBox20.Text
        pname1 = TextBox13.Text
        pname2 = TextBox17.Text
        pname3 = TextBox21.Text
        pprice1 = TextBox14.Text
        pprice2 = TextBox18.Text
        pprice3 = TextBox22.Text
        pqty1 = ComboBox2.Text
        pqty2 = ComboBox3.Text
        pqty3 = ComboBox4.Text
        ptotal1 = TextBox15.Text
        ptotal2 = TextBox19.Text
        ptotal3 = TextBox23.Text
        grand_total2 = TextBox24.Text

        total_amount = TextBox25.Text
        recived_amount = TextBox26.Text
        balance_amount = TextBox27.Text
        con.Open()
        cmd = New OleDbCommand("update table7 set department='" & department & "', type_1='" & type_1 & "', description__1='" & desc1 & "', price_1=" & price_1 & ", quantity_1=" & qty1 & ", total_1=" & total_1 & ", type_2='" & type_2 & "', description_2='" & desc2 & "', price_2=" & price_2 & ", quantity_2=" & qty2 & ", total_2=" & total_2 & ", grand_total1=" & grand_total1 & ", pcode1=" & product_code1 & ", pname1='" & pname1 & "', pprice1=" & pprice1 & ", pquantity1=" & pqty1 & ", ptotal1=" & ptotal1 & ", pcode2=" & product_code2 & ", pname2='" & pname2 & "', pprice2=" & pprice2 & ", pquantity2=" & pqty2 & ", ptotal2=" & ptotal2 & ", pcode3=" & product_code3 & ", pname3='" & pname3 & "', pprice3=" & pprice3 & ", pquantity3=" & pqty3 & ", ptotal3=" & ptotal3 & ", grand_total2=" & grand_total2 & ", total_amount=" & total_amount & ", recived_amount=" & recived_amount & ", balance_amount=" & balance_amount & " where bill_no=" & billno, con)
        cmd.ExecuteNonQuery()
        MsgBox("Amount Updated")
        con.Close()

    End Sub
End Class