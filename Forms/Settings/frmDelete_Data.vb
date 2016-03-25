Public Class frmDelete_Data
    Dim Myconn As New Connection
    Private Sub btn_kind_Click(sender As Object, e As EventArgs) Handles btn_kind.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Items]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_Customers_Click(sender As Object, e As EventArgs) Handles btn_Customers.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Customers]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_Suppliers_Click(sender As Object, e As EventArgs) Handles btn_Suppliers.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Supplier]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_Purchasas_Click(sender As Object, e As EventArgs) Handles btn_Purchasas.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Purchases]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_Sales_Click(sender As Object, e As EventArgs) Handles btn_Sales.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Sales]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_recive_Click(sender As Object, e As EventArgs) Handles btn_recive.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Safe_Recive_per]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_Payment_Click(sender As Object, e As EventArgs) Handles btn_Payment.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Safe_payment_per]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_Employees_Click(sender As Object, e As EventArgs) Handles btn_Employees.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Employees]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_Users_Click(sender As Object, e As EventArgs) Handles btn_Users.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Users_ID]")
        Myconn.ExecQuery("delete from [Users_Permission]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_Stocks_Click(sender As Object, e As EventArgs) Handles btn_Stocks.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Stocks]")
        Myconn.ExecQuery("delete from [Items_move]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_Group_Click(sender As Object, e As EventArgs) Handles btn_Group.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [group]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_Check_Click(sender As Object, e As EventArgs) Handles btn_Check.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Bank_checks]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_Bank_Click(sender As Object, e As EventArgs) Handles btn_Bank.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Bank_Operations]")
        MessageBox.Show("تم حذف البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btn_All_data_Click(sender As Object, e As EventArgs) Handles btn_All_data.Click
        If MsgBox("هل أنت متأكد من عملية الحذف فإن ذلك سيؤدي لفقد جمبيع البيانات ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.ExecQuery("delete from [Items]")
        Myconn.ExecQuery("delete from [Customers]")
        Myconn.ExecQuery("delete from [Supplier]")
        Myconn.ExecQuery("delete from [Purchases]")
        Myconn.ExecQuery("delete from [Sales]")
        Myconn.ExecQuery("delete from [Safe_Recive_per]")
        Myconn.ExecQuery("delete from [Safe_payment_per]")
        Myconn.ExecQuery("delete from [Employees]")
        Myconn.ExecQuery("delete from [Users_ID]")
        Myconn.ExecQuery("delete from [Users_Permission]")
        Myconn.ExecQuery("delete from [Stocks]")
        Myconn.ExecQuery("delete from [Items_move]")
        Myconn.ExecQuery("delete from [group]")
        Myconn.ExecQuery("delete from [Bank_checks]")
        Myconn.ExecQuery("delete from [Bank_Operations]")
        MessageBox.Show("تم حذف جميع البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub
End Class