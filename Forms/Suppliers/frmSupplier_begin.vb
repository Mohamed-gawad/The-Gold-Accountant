Public Class frmSupplier_begin
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st As String

    Private Sub New_record()
        Myconn.ClearAllControls(GroupBox1, True)

        cbo_Customer.SelectedIndex = -1
    End Sub
    Private Sub Filldrg()
        drg.Rows.Clear()
        Myconn.ExecQuery("Select S.ID,S.Pur_Date,S.Supplier_ID,S.Total_Price,C.Supplier_Name,U.Employee_Name From (Purchases S left join Supplier C on S.Supplier_ID = C.Supplier_ID)
                          Left join Users_ID U on S.Employee_ID = U.Employee_ID where S.Pur_Bill_num = 0")
        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = Format(CDate(r("Pur_Date")), "yyyy/MM/dd")
            drg.Rows(i).Cells(2).Value = r("Supplier_Name")
            drg.Rows(i).Cells(3).Value = r("Supplier_ID")
            drg.Rows(i).Cells(4).Value = r("Total_Price")
            drg.Rows(i).Cells(5).Value = r("Employee_Name")
            drg.Rows(i).Cells(6).Value = r("ID")
        Next

    End Sub
    Private Sub Binding()
        Myconn.ExecQuery("Select S.ID,S.Pur_Date,S.Supplier_ID,S.Total_Price,C.Supplier_Name,U.Employee_Name From (Purchases S left join Supplier C on S.Supplier_ID = C.Supplier_ID)
                          Left join Users_ID U on S.Employee_ID = U.Employee_ID where S.ID = " & CInt(drg.CurrentRow.Cells(6).Value))
        Dim r As DataRow = Myconn.dt.Rows(0)
        D_date.Text = r("Pur_Date")
        cbo_Customer.SelectedValue = r("Supplier_ID")
        txtAmount.Text = r("Total_Price")


    End Sub

    Private Sub Save_recod()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@Pur_Date", Format(CDate(D_date.Text), "yyyy/MM/dd")) ' التاريخ
                .Addparam("@Pur_Bill_num", 0) ' الفاتورة
                .Addparam("@Total_Price", txtAmount.Text) 'المبلغ
                .Addparam("@Supplier_ID", cbo_Customer.SelectedValue) ' المورد
                .Addparam("@Employee_ID", My.Settings.user_ID) ' المستخدم
                .Addparam("@Status", 1) ' الحالة
                .ExecQuery("insert into [Purchases] (Pur_Date,Pur_Bill_num,Total_Price,Supplier_ID,Employee_ID,Status)
                                          values(@Pur_Date,@Pur_Bill_num,@Total_Price,@Supplier_ID,@Employee_ID,@Status)")

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
                .Addparam("@Pur_Date", Format(CDate(D_date.Text), "yyyy/MM/dd")) ' التاريخ
                .Addparam("@Pur_Bill_num", 0) ' الفاتورة
                .Addparam("@Total_Price", txtAmount.Text) 'المبلغ
                .Addparam("@Supplier_ID", cbo_Customer.SelectedValue) ' المورد
                .Addparam("@Employee_ID", My.Settings.user_ID) ' المستخدم
                .Addparam("@ID", drg.CurrentRow.Cells(6).Value) ' رقم السجل
                .ExecQuery("Update [Purchases] set Pur_Date=@Pur_Date,Pur_Bill_num=@Pur_Bill_num,Total_Price=@Total_Price,Supplier_ID=@Supplier_ID,Employee_ID=@Employee_ID where ID = @ID ")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub frmSupplier_begin_Load(sender As Object, e As EventArgs) Handles Me.Load
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
        Myconn.Fillcombo("select * from Supplier order by Supplier_Name", "Supplier", "Supplier_ID", "Supplier_Name", Me, cbo_Customer)
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
            .ExecQuery("delete from [Purchases] where ID = @ID ")
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
        Myconn.ExecQuery("Select Supplier_ID from Supplier where Supplier_ID =" & CInt(cbo_Customer.SelectedValue) & "")
        If Myconn.Recodcount = 0 Then Return
        Dim r As DataRow = Myconn.dt.Rows(0)
        txtCustomer_ID.Text = r("Supplier_ID")
    End Sub

    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Binding()

    End Sub

    Private Sub cbo_Customer_Enter(sender As Object, e As EventArgs) Handles cbo_Customer.Enter
        Myconn.langAR()

    End Sub
End Class