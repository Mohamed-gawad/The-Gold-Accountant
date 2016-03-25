Imports System.Windows.Forms
Imports System.Drawing.Drawing2D
Imports System.ComponentModel

Public Class FrmMain
    Dim myconn As New Connection
    Private Sub bgColor()
        Dim child As Control
        For Each child In Me.Controls
            If TypeOf child Is MdiClient Then
                child.BackColor = Color.LavenderBlush
                'child.BackgroundImage = Me.BackgroundImage

                Exit For
            End If
        Next
        child = Nothing
    End Sub
    Private Sub myMdiControlPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)

        'e.Graphics.DrawImage(Me.PictureBox1.BackgroundImage, 0, 0, Me.Width, Me.Height)
    End Sub
    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        MenuStrip.Renderer = New clsMenuRenderer
        bgColor()
        For Each ctl As Control In Me.Controls
            If TypeOf ctl Is MdiClient Then
                ctl.BackgroundImage = Me.BackgroundImage
            End If
        Next ctl
        Me.KeyPreview = True
        myconn.ExecQuery("Select * from Bank_checks where  Bank_checks.Check_date >= #" & Format(Today.Date, "yyyy/MM/dd") & "#")
        If myconn.Recodcount > 0 Then
            Dim frm As New frmRecodr_sheek("x")
            frm.MdiParent = Me
            frm.Show()

        End If
    End Sub

    Private Sub Out_Click(sender As Object, e As EventArgs) Handles Out.Click
        Application.Exit()
    End Sub

    Private Sub mini_Click(sender As Object, e As EventArgs) Handles mini.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub K_list_kind_Click(sender As Object, e As EventArgs) Handles K_list_kind.Click

        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'K_list_kind'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmList_kinds.MdiParent = Me
        frmList_kinds.Show()

    End Sub

    Private Sub S_Bill_Sales_Click(sender As Object, e As EventArgs) Handles S_Bill_Sales.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_Bill_Sales'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        Dim frm As New frmBill_Sale
        frm.MdiParent = Me
        frm.Show()

    End Sub

    Private Sub K_new_kind_Click(sender As Object, e As EventArgs) Handles K_new_kind.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'K_new_kind'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmAdd_Kind.MdiParent = Me
        frmAdd_Kind.Show()

    End Sub

    Private Sub K_new_group_Click(sender As Object, e As EventArgs) Handles K_new_group.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'K_new_group'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmAdd_Group.MdiParent = Me
        frmAdd_Group.Show()

    End Sub

    Private Sub p_Bill_Purc_Click(sender As Object, e As EventArgs) Handles p_Bill_Purc.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'p_Bill_Purc'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmPurchases_Bill.MdiParent = Me
        frmPurchases_Bill.Show()

    End Sub

    Private Sub اS_new_stock_Click(sender As Object, e As EventArgs) Handles S_new_stock.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_new_stock'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmAdd_Stock.MdiParent = Me
        frmAdd_Stock.Show()

    End Sub

    Private Sub P_Bills_Purc_Click(sender As Object, e As EventArgs) Handles P_Bills_Purc.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'P_Bills_Purc'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmAll_Purchases_Bills.MdiParent = Me
        frmAll_Purchases_Bills.Show()

    End Sub

    Private Sub P_Back_Purc_Click(sender As Object, e As EventArgs) Handles P_Back_Purc.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'P_Back_Purc'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmBack_Purchases.MdiParent = Me
        frmBack_Purchases.Show()

    End Sub

    Private Sub E_new_employee_Click(sender As Object, e As EventArgs) Handles E_new_employee.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'E_new_employee'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmAdd_Employee.MdiParent = Me
        frmAdd_Employee.Show()

    End Sub

    Private Sub U_permisstion_Click(sender As Object, e As EventArgs) Handles U_permisstion.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'U_permisstion'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmPermission.MdiParent = Me
        frmPermission.Show()
    End Sub

    Private Sub U_new_user_Click(sender As Object, e As EventArgs) Handles U_new_user.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'U_new_user'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmAdd_User.MdiParent = Me
        frmAdd_User.Show()
    End Sub

    Private Sub FrmMain_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        frmLogin.Hide()
    End Sub

    Private Sub S_Bills_Sales_Click(sender As Object, e As EventArgs) Handles S_Bills_Sales.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_Bills_Sales'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmAll_Sales_Bills.MdiParent = Me
        frmAll_Sales_Bills.Show()
    End Sub

    Private Sub S_Back_Sales_Click(sender As Object, e As EventArgs) Handles S_Back_Sales.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_Back_Sales'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmBack_Sales.MdiParent = Me
        frmBack_Sales.Show()
    End Sub

    Private Sub C_new_customer_Click(sender As Object, e As EventArgs) Handles C_new_customer.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'C_new_customer'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmAdd_Customer.MdiParent = Me
        frmAdd_Customer.Show()
    End Sub

    Private Sub S_recive_Click(sender As Object, e As EventArgs) Handles S_recive.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_recive'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmRecive.MdiParent = Me
        frmRecive.Show()
    End Sub

    Private Sub S_pay_Click(sender As Object, e As EventArgs) Handles S_pay.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_pay'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmPayment.MdiParent = Me
        frmPayment.Show()
    End Sub

    Private Sub S_safe_move_Click(sender As Object, e As EventArgs) Handles S_safe_move.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_safe_move'")
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmSafe_moving.MdiParent = Me
        frmSafe_moving.Show()
    End Sub

    Private Sub s_ezn_back_Click(sender As Object, e As EventArgs) Handles s_ezn_back.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 's_ezn_back'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmBack_ezn.MdiParent = Me
        frmBack_ezn.Show()
    End Sub

    Private Sub K_kind_move_Click(sender As Object, e As EventArgs) Handles K_kind_move.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'K_kind_move'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmKind_move.MdiParent = Me
        frmKind_move.Show()
    End Sub

    Private Sub K_Kinds_move_Click(sender As Object, e As EventArgs) Handles K_Kinds_move.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'K_kinds_move'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmKinds_move.MdiParent = Me
        frmKinds_move.Show()
    End Sub

    Private Sub C_customer_account_Click(sender As Object, e As EventArgs) Handles C_customer_account.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'C_customer_account'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmCustomer_account.MdiParent = Me
        frmCustomer_account.Show()
    End Sub

    Private Sub C_customers_account_Click(sender As Object, e As EventArgs) Handles C_customers_account.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'C_customers_account'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmCustomers_account.MdiParent = Me
        frmCustomers_account.Show()
    End Sub

    Private Sub C_first_amount_Click(sender As Object, e As EventArgs) Handles C_first_amount.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'C_first_amount'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmCustomer_begin.MdiParent = Me
        frmCustomer_begin.Show()
    End Sub

    Private Sub S_new_supplier_Click(sender As Object, e As EventArgs) Handles S_new_supplier.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_new_supplier'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmAdd_Supplier.MdiParent = Me
        frmAdd_Supplier.Show()
    End Sub

    Private Sub S_supplier_account_Click(sender As Object, e As EventArgs) Handles S_supplier_account.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_supplier_account'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmSupplier_account.MdiParent = Me
        frmSupplier_account.Show()
    End Sub

    Private Sub S_first_amount_Click(sender As Object, e As EventArgs) Handles S_first_amount.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_first_amount'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmSupplier_begin.MdiParent = Me
        frmSupplier_begin.Show()
    End Sub

    Private Sub S_suppliers_acount_Click(sender As Object, e As EventArgs) Handles S_suppliers_acount.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_suppliers_acount'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmSuppliers_account.MdiParent = Me
        frmSuppliers_account.Show()
    End Sub

    Private Sub S_goods_supplier_Click(sender As Object, e As EventArgs) Handles S_goods_supplier.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_goods_supplier'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmSupplier_kinds_move.MdiParent = Me
        frmSupplier_kinds_move.Show()
    End Sub

    Private Sub S_amount_stock_Click(sender As Object, e As EventArgs) Handles S_amount_stock.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_amount_stock'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmStocks_account.MdiParent = Me
        frmStocks_account.Show()
    End Sub

    Private Sub S_transport_stocks_Click(sender As Object, e As EventArgs) Handles S_transport_stocks.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_transport_stocks'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmTransforme_kinds.MdiParent = Me
        frmTransforme_kinds.Show()
    End Sub

    Private Sub S_programe_Click(sender As Object, e As EventArgs) Handles S_programe.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_programe'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmSetting.MdiParent = Me
        frmSetting.Show()
    End Sub

    Private Sub S_data_Click(sender As Object, e As EventArgs) Handles S_data.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_data'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmCo_data.MdiParent = Me
        frmCo_data.Show()
    End Sub

    Private Sub S_Barcode_Click(sender As Object, e As EventArgs) Handles S_Barcode.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_Barcode'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmBarcode_Setting.MdiParent = Me
        frmBarcode_Setting.Show()
    End Sub

    Private Sub S_Delete_Data_Click(sender As Object, e As EventArgs) Handles S_Delete_Data.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_Delete_Data'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmDelete_Data.MdiParent = Me
        frmDelete_Data.Show()
    End Sub

    Private Sub S_backup_Click(sender As Object, e As EventArgs) Handles S_backup.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_backup'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmBackup_Data.MdiParent = Me
        frmBackup_Data.Show()
    End Sub

    Private Sub S_Repair_Data_Click(sender As Object, e As EventArgs) Handles S_Repair_Data.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_Repair_Data'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmCompact_DataBase.MdiParent = Me
        frmCompact_DataBase.Show()
    End Sub

    Private Sub S_restor_Click(sender As Object, e As EventArgs) Handles S_restor.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_restor'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmRestor_Data.MdiParent = Me
        frmRestor_Data.Show()
    End Sub

    Private Sub E_Jobs_Click(sender As Object, e As EventArgs) Handles E_Jobs.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'E_Jobs'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmJob.MdiParent = Me
        frmJob.Show()
    End Sub

    Private Sub R_Sales_Click(sender As Object, e As EventArgs) Handles R_Sales.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'R_Sales'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmReport_Sales.MdiParent = Me
        frmReport_Sales.Show()
    End Sub

    Private Sub R_Purchases_Click(sender As Object, e As EventArgs) Handles R_Purchases.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'R_Purchases'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmReport_Purchases.MdiParent = Me
        frmReport_Purchases.Show()
    End Sub

    Private Sub R_Earning_Click(sender As Object, e As EventArgs) Handles R_Earning.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'R_Earning'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmReport_Erning.MdiParent = Me
        frmReport_Erning.Show()
    End Sub

    Private Sub R_payments_Click(sender As Object, e As EventArgs) Handles R_payments.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'R_payments'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmreport_Payment.MdiParent = Me
        frmreport_Payment.Show()
    End Sub

    Private Sub B_Add_bank_Click(sender As Object, e As EventArgs) Handles B_Add_bank.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'B_Add_bank'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frm_Add_bank.MdiParent = Me
        frm_Add_bank.Show()
    End Sub

    Private Sub B_Record_ezn_Click(sender As Object, e As EventArgs) Handles B_Record_ezn.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'B_Record_ezn'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmRecord_ezn.MdiParent = Me
        frmRecord_ezn.Show()
    End Sub

    Private Sub B_Record_sheek_Click(sender As Object, e As EventArgs) Handles B_Record_sheek.Click
        myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'B_Record_sheek'")
        If myconn.Recodcount = 0 Then Return
        Dim r As DataRow = myconn.dt.Rows(0)
        Per = r("Sub_menu_ID")
        frmRecodr_sheek.MdiParent = Me
        frmRecodr_sheek.Show()
    End Sub

    Private Sub FrmMain_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If e.KeyCode = Keys.F12 Then ' لفتح فاتورة مبيعات
            myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'S_Bill_Sales'")
            If myconn.Recodcount = 0 Then Return
            Dim r As DataRow = myconn.dt.Rows(0)
            Per = r("Sub_menu_ID")

            If F <> 1 Then
                myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If myconn.dt.Rows.Count = 0 Then Return
            End If

            Dim frm As New frmBill_Sale
            frm.MdiParent = Me
            frm.Show()
        End If

        If e.KeyCode = Keys.F1 Then ' لفتح فاتورة مشتريات
            myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'p_Bill_Purc'")
            If myconn.Recodcount = 0 Then Return
            Dim r As DataRow = myconn.dt.Rows(0)
            Per = r("Sub_menu_ID")

            If F <> 1 Then
                myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If myconn.dt.Rows.Count = 0 Then Return
            End If

            frmPurchases_Bill.MdiParent = Me
            frmPurchases_Bill.Show()
        End If

        If e.KeyCode = Keys.F5 Then '
            myconn.ExecQuery("Select * from Sub_menu where Sub_menu_name  like 'K_kind_move'")
            If myconn.Recodcount = 0 Then Return
            Dim r As DataRow = myconn.dt.Rows(0)
            Per = r("Sub_menu_ID")

            If F <> 1 Then
                myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If myconn.dt.Rows.Count = 0 Then Return
            End If

            frmKind_move.MdiParent = Me
            frmKind_move.Show()
        End If
    End Sub
End Class