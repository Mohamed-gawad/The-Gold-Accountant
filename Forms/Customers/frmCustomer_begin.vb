Public Class frmCustomer_begin

    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st As String

    Private Sub New_record()
        Myconn.ClearAllControls(GroupBox1, True)

        cbo_Customer.SelectedIndex = -1
    End Sub
    Private Sub Filldrg()
        drg.Rows.Clear()
        Myconn.ExecQuery("Select S.ID,S.Sales_Date,S.Customer_ID,S.Final_Total_Price,C.Customer_Name,U.Employee_Name From (Sales S left join Customers C on S.customer_ID = C.Customer_ID)
                          Left join Users_ID U on S.Users_ID = U.Employee_ID where S.Sales_Bill_ID = 0")
        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = Format(CDate(r("Sales_Date")), "yyyy/MM/dd")
            drg.Rows(i).Cells(2).Value = r("Customer_Name")
            drg.Rows(i).Cells(3).Value = r("Customer_ID")
            drg.Rows(i).Cells(4).Value = r("Final_Total_Price")
            drg.Rows(i).Cells(5).Value = r("Employee_Name")
            drg.Rows(i).Cells(6).Value = r("ID")
        Next

    End Sub
    Private Sub Binding()
        Myconn.ExecQuery("Select S.ID,S.Sales_Date,S.Customer_ID,S.Final_Total_Price,C.Customer_Name,U.Employee_Name From (Sales S left join Customers C on S.customer_ID = C.Customer_ID)
                          Left join Users_ID U on S.Users_ID = U.Employee_ID where S.ID = " & CInt(drg.CurrentRow.Cells(6).Value))
        Dim r As DataRow = Myconn.dt.Rows(0)
        D_date.Text = r("Sales_Date")
        cbo_Customer.SelectedValue = r("Customer_ID")
        txtAmount.Text = r("Final_Total_Price")


    End Sub

    Private Sub Save_recod()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@Sales_Date", Format(CDate(D_date.Text), "yyyy/MM/dd")) ' التاريخ
                .Addparam("@Sales_Bill_ID", 0) ' التاريخ
                .Addparam("@Final_Total_Price", txtAmount.Text) '
                .Addparam("@Customer_ID", cbo_Customer.SelectedValue) '
                .Addparam("@Users_ID", My.Settings.user_ID) '
                .Addparam("@Status", 1) '
                .ExecQuery("insert into [Sales] (Sales_Date,Sales_Bill_ID,Final_Total_Price,Customer_ID,Users_ID,Status)
                                          values(@Sales_Date,@Sales_Bill_ID,@Final_Total_Price,@Customer_ID,@Users_ID,@Status)")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Update_record()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@Sales_Date", Format(CDate(D_date.Text), "yyyy/MM/dd")) ' التاريخ
                .Addparam("@Sales_Bill_ID", 0) ' التاريخ
                .Addparam("@Final_Total_Price", txtAmount.Text) '
                .Addparam("@Customer_ID", cbo_Customer.SelectedValue) '
                .Addparam("@Users_ID", My.Settings.user_ID) '
                .Addparam("@ID", drg.CurrentRow.Cells(6).Value) '
                .ExecQuery("Update [Sales] set Sales_Date=@Sales_Date,Sales_Bill_ID=@Sales_Bill_ID,Final_Total_Price=@Final_Total_Price,Customer_ID=@Customer_ID,Users_ID=@Users_ID where ID = @ID ")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub frmCustomer_begin_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width

            If  F <> 1Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Exit Sub
                Dim r As DataRow = Myconn.dt.Rows(0)
                If r("U_full").ToString = False Then
                    btnSave.Enabled = r("U_add").ToString
                    btnUpdat.Enabled = r("U_updat").ToString
                    btnNew.Enabled = r("U_new").ToString
                    btnDel.Enabled = r("U_delete").ToString
                    btnPrint.Enabled = r("U_print").ToString
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        fin = False
        Myconn.Fillcombo("select * from Customers order by Customer_Name", "Customers", "Customer_ID", "Customer_Name", Me, cbo_Customer)
        fin = True
        Filldrg()

    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        New_record()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is TextBox Then
                If txt.Text = "" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            ElseIf TypeOf txt Is ComboBox Then
                If txt.Text = "" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            End If
        Next
        Save_recod()
        Filldrg()
        MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        With Myconn
            .Addparam("@ID", drg.CurrentRow.Cells(6).Value)
            .ExecQuery("delete from [Sales] where ID = @ID ")
        End With
        If Myconn.NoErrors(True) = False Then Exit Sub
        drg.Rows.Remove(drg.SelectedRows(0))
        Myconn.ClearAllControls(GroupBox2, True)
        Filldrg()
        Binding()
    End Sub

    Private Sub btnUpdat_Click(sender As Object, e As EventArgs) Handles btnUpdat.Click
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is TextBox Then
                If txt.Text = "" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            ElseIf TypeOf txt Is ComboBox Then
                If txt.Text = "" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            End If
        Next
        Update_record()
        Filldrg()
        MessageBox.Show("تمت عملية التعديل بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
    End Sub

    Private Sub cbo_Customer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_Customer.SelectedIndexChanged
        If Not fin Then Return
        Myconn.ExecQuery("Select Customer_ID from Customers where Customer_ID =" & CInt(cbo_Customer.SelectedValue) & "")
        If Myconn.Recodcount = 0 Then Return
        Dim r As DataRow = Myconn.dt.Rows(0)
        txtCustomer_ID.Text = r("Customer_ID")
    End Sub

    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Binding()

    End Sub

    Private Sub cbo_Customer_Enter(sender As Object, e As EventArgs) Handles cbo_Customer.Enter
        Myconn.langAR()

    End Sub
End Class