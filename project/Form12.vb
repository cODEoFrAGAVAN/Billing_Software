Public Class Form12

    Private Sub UserManagementToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserManagementToolStripMenuItem.Click
        Form2.Show()

    End Sub

    Private Sub SupplierManagementToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupplierManagementToolStripMenuItem.Click
        Form3.Show()

    End Sub

    Private Sub ProductRegistrationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductRegistrationToolStripMenuItem.Click
        Form4.Show()

    End Sub

    Private Sub StockPurchaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StockPurchaseToolStripMenuItem.Click
        Form5.Show()

    End Sub

    Private Sub StockReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub UsersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsersToolStripMenuItem.Click
        Form6.Show()

    End Sub

    Private Sub SuppliersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuppliersToolStripMenuItem.Click
        Form7.Show()

    End Sub

    Private Sub ProductsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductsToolStripMenuItem.Click
        Form8.Show()

    End Sub

    Private Sub PurchaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PurchaseToolStripMenuItem.Click
        Form9.Show()

    End Sub

    Private Sub StudentInformationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StudentInformationToolStripMenuItem.Click


    End Sub

    Private Sub ReportsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReportsToolStripMenuItem.Click

    End Sub

    Private Sub SalesToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesToolStripMenuItem1.Click


    End Sub

    Private Sub SalesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesToolStripMenuItem.Click


    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub StudentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StudentToolStripMenuItem.Click
        Form11.Show()

    End Sub

    Private Sub DepartmentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DepartmentToolStripMenuItem.Click
        Form14.Show()
    End Sub

    Private Sub StudentRegistrationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StudentRegistrationToolStripMenuItem.Click
        Form10.Show()

    End Sub

    Private Sub Form12_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If (Form1.UT = "OTHERS") Then
            UserManagementToolStripMenuItem.Visible = False
            SupplierManagementToolStripMenuItem.Visible = False
            ReportsToolStripMenuItem.Visible = False
            ReportsToolStripMenuItem.Visible = False
        End If
    End Sub

    Private Sub StudentToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StudentToolStripMenuItem1.Click
        Form13.Show()

    End Sub

    Private Sub StudentDeatilsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StudentDeatilsToolStripMenuItem.Click
        Form15.Show()

    End Sub

    Private Sub StudentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StudentsToolStripMenuItem.Click
        Form16.Show()

    End Sub

    Private Sub DepartmentToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DepartmentToolStripMenuItem2.Click
        Form17.Show()

    End Sub

    Private Sub DepartmentToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DepartmentToolStripMenuItem1.Click
        Form18.Show()

    End Sub

    Private Sub LogOutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogOutToolStripMenuItem.Click
        End
    End Sub
End Class