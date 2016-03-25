Public Class frmAdd_Customer
    Dim Myconn As New Connection

    Private Sub New_record()
        Myconn.ClearAllControls(GroupBox1, True)
        Myconn.Autonumber("Customer_ID", "Customers", txtID, Me)
    End Sub
    Private Sub Filldrg()
        drg.Rows.Clear()
        Myconn.ExecQuery("SELECT * from [Customers] order by Customer_ID")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub

        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("Customer_Name")
            drg.Rows(i).Cells(2).Value = r("Customer_ID")
            drg.Rows(i).Cells(3).Value = r("Customer_Address")
            drg.Rows(i).Cells(4).Value = r("Customer_tel")
            drg.Rows(i).Cells(5).Value = r("Customer_mobil1")
            drg.Rows(i).Cells(6).Value = r("Customer_mobil2")
            drg.Rows(i).Cells(7).Value = r("ID")
        Next
        Myconn.DataGridview_MoveLast(drg, 2)
    End Sub
    Private Sub Binding()
        Myconn.ExecQuery("SELECT * from [Customers]  where ID =" & CInt(drg.CurrentRow.Cells(7).Value))
        Dim r As DataRow = Myconn.dt.Rows(0)
        txtID.Text = r("Customer_ID").ToString
        txtName.Text = r("Customer_name").ToString
        txtAddress.Text = r("Customer_Address").ToString
        txt_tel.Text = r("Customer_tel").ToString
        txtMobil1.Text = r("Customer_mobil1").ToString
        txtMobil2.Text = r("Customer_mobil2").ToString
    End Sub
    Private Sub Save_recod()
        With Myconn
            .Parames.Clear()
            .Addparam("@Customer_ID", CInt(txtID.Text))
            .Addparam("@Customer_Name", txtName.Text)
            .Addparam("@Customer_tel", txt_tel.Text)
            .Addparam("@Customer_mobil1", txtMobil1.Text)
            .Addparam("@Customer_mobil2", txtMobil2.Text)
            .Addparam("@Customer_Address", txtAddress.Text)

            .ExecQuery("insert into  [Customers] (Customer_ID,Customer_Name,Customer_tel,Customer_mobil1,Customer_mobil2,Customer_Address) 
                                           values(@Customer_ID,@Customer_Name,@Customer_tel,@Customer_mobil1,@Customer_mobil2,@Customer_Address)")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
    End Sub
    Private Sub Update_record()
        With Myconn
            .Parames.Clear()
            .Addparam("@Customer_Name", txtName.Text)
            .Addparam("@Customer_tel", txt_tel.Text)
            .Addparam("@Customer_mobil1", txtMobil1.Text)
            .Addparam("@Customer_mobil2", txtMobil2.Text)
            .Addparam("@Customer_Address", txtAddress.Text)
            .Addparam("@ID", drg.CurrentRow.Cells(7).Value)
            .ExecQuery("Update  [Customers] set  Customer_Name=@Customer_Name,Customer_tel=@Customer_tel,Customer_mobil1=@Customer_mobil1,Customer_mobil2=@Customer_mobil2,Customer_Address=@Customer_Address where ID =@ID")

            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
    End Sub

    Private Sub frmAdd_Customer_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width
            If F <> 1 Then
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
            New_record()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Filldrg()
    End Sub
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        New_record()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If txtID.Text = "" Or txtName.Text = "" Then
            ErrorProvider1.SetError(txtName, "أكمل البيانات")
            MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
            Return

        End If
        Save_recod()
        Filldrg()

        MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
        New_record()
    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        With Myconn
            .Addparam("@ID", drg.CurrentRow.Cells(7).Value)
            .ExecQuery("delete from [Customers] where ID = @ID")
        End With
        If Myconn.NoErrors(True) = False Then Exit Sub
        drg.Rows.Remove(drg.SelectedRows(0))
        Myconn.ClearAllControls(GroupBox1, True)
    End Sub

    Private Sub btnUpdat_Click(sender As Object, e As EventArgs) Handles btnUpdat.Click
        If txtID.Text = "" Or txtName.Text = "" Then
            ErrorProvider1.SetError(txtName, "أكمل البيانات")
            MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
            Return

        End If
        Update_record()
        Filldrg()
        MessageBox.Show("تمت عملية التعديل بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Binding()
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Print_drg()
    End Sub

    Sub Print_drg()
        Try
            Dim rpt As New rpt_Customers
            Dim table As New DataTable
            For i As Integer = 1 To 6
                Dim x As String
                x = Format(i, "00")
                table.Columns.Add(x)
            Next

            For Each dr As DataGridViewRow In drg.Rows
                table.Rows.Add()
                table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count ' المسلسل
                table.Rows(table.Rows.Count - 1)(1) = dr.Cells(1).Value ' العميل
                table.Rows(table.Rows.Count - 1)(2) = dr.Cells(2).Value ' الكود
                table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value ' العنوان
                table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value ' التليفون
                table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value ' المحمول

            Next
            rpt.SetDataSource(table)
            rpt.SetParameterValue("Co_name", My.Settings.Co_name)
            rpt.SetParameterValue("Address", My.Settings.Co_address & " ت : " & My.Settings.Co_tel)

            rpt.PrintOptions.PrinterName = My.Settings.Printer_Sales
            If My.Settings.Print = True Then
                frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                frmReportViewer.Show()
            Else
                rpt.PrintOptions.PrinterName = My.Settings.Printer_report
                rpt.PrintToPrinter(1, False, 0, 0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
End Class