Imports System.Globalization
Public Class frmRecive
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st As String
    Dim x, y As Integer

    Private Sub New_record()
        Myconn.ClearAllControls(GroupBox1, True)
        Myconn.Autonumber("per_ID", "Safe_Recive_per", txt_ID, Me)
    End Sub
    Private Sub Filldrg()
        Try
            drg.Rows.Clear()

            Select Case y
                Case 0
                    st = " where Safe_Recive_per.per_date = #" & Format(CDate(Now.Date), "yyyy/MM/dd") & "#"
                Case 1
                    st = " where Safe_Recive_per.per_ID =" & CInt(txtSearch.Text)
                Case 2
                    st = " where Safe_Recive_per.per_date = #" & Format(CDate(txtSearch.Text), "yyyy/MM/dd") & "#"

            End Select

            Myconn.ExecQuery("SELECT Safe_Recive_per.ID,Bank.Bank_name,Safe_Recive_per.Bank_ID,Safe_Recive_per.per_ID, Safe_Recive_per.per_date, Safe_Recive_per.per_time, Safe_Recive_per.users_ID, Safe_Recive_per.Recive_Item_ID, Safe_Recive_per.perm_ID, Safe_Recive_per.Amount, Safe_Recive_per.Note_per, Safe_Recive_per.Status, Customers.Customer_Name, Recive_Items.Recive_Item_name, Users_ID.Employee_Name
                            FROM (((Safe_Recive_per LEFT JOIN Recive_Items ON Safe_Recive_per.Recive_Item_ID = Recive_Items.Recive_Item_ID)
                            LEFT JOIN Customers ON Safe_Recive_per.Customer_ID = Customers.Customer_ID )
                            LEFT JOIN Users_ID ON Safe_Recive_per.users_ID = Users_ID.Employee_ID )
                            LEFT JOIN Bank ON Safe_Recive_per.bank_ID = Bank.Bank_ID                             
                            " & st & " order by Safe_Recive_per.ID ")

            If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub

            If Myconn.Recodcount = 0 Then
                Select Case y
                    Case 0
                        MsgBox("لا توجد أذونات هذا اليوم", MsgBoxStyle.Information, "رسالة")
                    Case 1
                        MsgBox("لا يوجد إذن بهذا الرقم", MsgBoxStyle.Information, "رسالة")
                    Case 2
                        MsgBox("لا توجد أذونات بهذا التاريخ", MsgBoxStyle.Information, "رسالة")
                End Select
                Return
            End If

            Dim V1 As Double = 0
            Dim V2 As Double = 0
            Dim B As Double = 0
            For i As Integer = 0 To Myconn.dt.Rows.Count - 1
                Dim r As DataRow = Myconn.dt.Rows(i)
                drg.Rows.Add()
                drg.Rows(i).Cells(0).Value = i + 1
                drg.Rows(i).Cells(1).Value = r("per_time").ToString
                drg.Rows(i).Cells(2).Value = Format(CDate(r("per_date").ToString), "yyyy/MM/dd")
                drg.Rows(i).Cells(3).Value = r("per_ID")
                drg.Rows(i).Cells(4).Value = If(IsDBNull(r("Bank_ID")), r("Customer_Name"), r("Bank_Name"))
                drg.Rows(i).Cells(5).Value = r("Amount")
                drg.Rows(i).Cells(6).Value = r("Recive_Item_name")
                drg.Rows(i).Cells(7).Value = r("note_per")
                drg.Rows(i).Cells(8).Value = r("Employee_Name")
                drg.Rows(i).Cells(9).Value = r("Status")
                drg.Rows(i).Cells(10).Value = r("ID")

                If drg.Rows(i).Cells(9).Value = True Then
                    drg.Rows(i).DefaultCellStyle.BackColor = Color.LemonChiffon
                Else
                    drg.Rows(i).DefaultCellStyle.BackColor = Color.Red
                End If
            Next
            Myconn.DataGridview_MoveLast(drg, 2)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Binding()
        Try
            Myconn.ExecQuery("Select * from Safe_Recive_per where ID =" & CInt(drg.CurrentRow.Cells(10).Value))
            Dim r As DataRow = Myconn.dt.Rows(0)
            D_date.Value = r("per_date")
            txt_ID.Text = r("per_ID")
            txtAmount.Text = If(IsDBNull(r("Amount")), "", r("Amount"))
            txtNotes.Text = If(IsDBNull(r("Note_per")), "", r("Note_per"))
            cbo_Band.SelectedValue = If(IsDBNull(r("Recive_Item_ID")), 0, r("Recive_Item_ID"))
            cbo_Customer.SelectedValue = If(IsDBNull(r("Customer_ID")), 0, r("Customer_ID"))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Save_recod()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@per_ID", txt_ID.Text)
                .Addparam("@per_date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
                .Addparam("@per_time", Label12.Text)
                .Addparam("@users_ID", My.Settings.user_ID)
                .Addparam("@Recive_Item_ID", cbo_Band.SelectedValue)
                .Addparam("@perm_ID", 1)
                .Addparam("@Amount", txtAmount.Text)
                .Addparam("@Customer_ID", If(RB1.Checked = True And cbo_Customer.SelectedIndex >= 0, cbo_Customer.SelectedValue, DBNull.Value))
                .Addparam("@Note_per", If(txtNotes.Text = "", DBNull.Value, txtNotes.Text))
                .Addparam("@Status", 1)
                .Addparam("@Bank_ID", If(RB2.Checked = True And cbo_Customer.SelectedIndex >= 0, cbo_Customer.SelectedValue, DBNull.Value))

                .ExecQuery("insert into [Safe_Recive_per] (per_ID,per_date,per_time,users_ID,Recive_Item_ID,perm_ID,Amount,Customer_ID,Note_per,Status,Bank_ID) Values(@per_ID,@per_date,@per_time,@users_ID,@Recive_Item_ID,@perm_ID,@Amount,@Customer_ID,@Note_per,@Status,@Bank_ID)")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub
    Private Sub Update_record()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@per_date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
                .Addparam("@per_time", Label12.Text)
                .Addparam("@users_ID", My.Settings.user_ID)
                .Addparam("@Recive_Item_ID", cbo_Band.SelectedValue)
                .Addparam("@perm_ID", 1)
                .Addparam("@Amount", txtAmount.Text)
                .Addparam("@Customer_ID", If(RB1.Checked = True And cbo_Customer.SelectedIndex >= 0, cbo_Customer.SelectedValue, DBNull.Value))
                .Addparam("@Note_per", If(txtNotes.Text = "", DBNull.Value, txtNotes.Text))
                .Addparam("@Bank_ID", If(RB2.Checked = True And cbo_Customer.SelectedIndex >= 0, cbo_Customer.SelectedValue, DBNull.Value))

                .Addparam("@ID", drg.CurrentRow.Cells(10).Value)

                .ExecQuery("Update [Safe_Recive_per]  Set per_date=@per_date,per_time=@per_time,users_ID=@users_ID,Recive_Item_ID=@Recive_Item_ID,perm_ID=@perm_ID,Amount=@Amount,Customer_ID=@Customer_ID,Note_per=@Note_per,Bank_ID=@Bank_ID where ID =@ID")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Myconn.ExecQuery("SELECT Safe_Recive_per.ID, Safe_Recive_per.per_ID, Safe_Recive_per.per_date, Safe_Recive_per.per_time, Safe_Recive_per.users_ID, Safe_Recive_per.Recive_Item_ID, Safe_Recive_per.perm_ID, Safe_Recive_per.Amount, Safe_Recive_per.Note_per, Safe_Recive_per.Status, Customers.Customer_Name, Recive_Items.Recive_Item_name, Users_ID.Employee_Name
                            FROM ((Safe_Recive_per LEFT JOIN Recive_Items ON Safe_Recive_per.Recive_Item_ID = Recive_Items.Recive_Item_ID) LEFT JOIN Customers ON Safe_Recive_per.Customer_ID = Customers.Customer_ID) LEFT JOIN Users_ID ON Safe_Recive_per.users_ID = Users_ID.Employee_ID where Safe_Recive_per.ID = " & CInt(drg.CurrentRow.Cells(10).Value) & "")
        Dim r As DataRow = Myconn.dt.Rows(0)
        drg.CurrentRow.Cells(1).Value = r("per_time")
        drg.CurrentRow.Cells(2).Value = Format(CDate(r("per_date").ToString), "yyyy/MM/dd")
        drg.CurrentRow.Cells(3).Value = r("per_ID")
        drg.CurrentRow.Cells(4).Value = r("Customer_Name")
        drg.CurrentRow.Cells(5).Value = r("Amount")
        drg.CurrentRow.Cells(6).Value = r("Recive_Item_name")
        drg.CurrentRow.Cells(7).Value = r("note_per")
        drg.CurrentRow.Cells(8).Value = r("Employee_Name")

        MessageBox.Show("تمت عملية التعديل بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub frmRecive_Load(sender As Object, e As EventArgs) Handles Me.Load
        Label5.Left = 0
        Label5.Width = Me.Width
        Try
            If  F <> 1Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " And Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Exit Sub
                Dim r As DataRow = Myconn.dt.Rows(0)
                If r("U_full").ToString = False Then
                    btnBack.Enabled = r("U_back").ToString
                    btnSave.Enabled = r("U_add").ToString
                    btnSearch.Enabled = r("U_search").ToString
                    btnUpdat.Enabled = r("U_updat").ToString
                    btnNew.Enabled = r("U_new").ToString
                    btnDel.Enabled = r("U_delete").ToString
                    btnPrint.Enabled = r("U_print").ToString
                End If
            End If
        Catch ex As Exception

        End Try

        fin = False
        Myconn.Fillcombo("Select * from [Recive_Items] order by [Recive_Item_name]", "[Recive_Items]", "Recive_Item_ID", "Recive_Item_name", Me, cbo_Band)
        Myconn.Fillcombo("select * from Customers order by Customer_Name", "Customers", "Customer_ID", "Customer_Name", Me, cbo_Customer)
        fin = True
        Timer1.Start()
        x = 0
        Filldrg()

        New_record()
        '-------------------------------------------------------------------------------------------------- النسخة التجريبية
        'Myconn.ExecQuery("select * from Safe_Recive_per")
        'If Myconn.Recodcount > 100 Then
        '    MsgBox("هذه النسخة تجريبية")
        '    btnSave.Enabled = False
        '    btnNew.Enabled = False
        '    Return
        'End If
    End Sub
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        New_record()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Myconn.ExecQuery("Select * from Safe_Recive_per where per_ID =" & CInt(txt_ID.Text))
        If Myconn.dt.Rows.Count > 0 Then
            MsgBox("هذا الرقم مكرر", MsgBoxStyle.Critical, "تحذير")
            Return
        End If
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is TextBox Then
                If txt.Text = "" And txt.Name <> "txtNotes" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            ElseIf TypeOf txt Is ComboBox Then
                If txt.Text = "" And txt.Name <> "cbo_Customer" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            End If
        Next
        If btnSave.Enabled = False Then MsgBox("هذه النسخة تجريبية", MsgBoxStyle.Critical, "رسالة") : Return
        Save_recod()
        y = 0
        Filldrg()
        New_record()

    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        With Myconn
            .Addparam("@ID", drg.CurrentRow.Cells(10).Value)
            .ExecQuery("delete from [Safe_Recive_per] where ID = @ID ")
        End With
        If Myconn.NoErrors(True) = False Then Exit Sub
        drg.Rows.Remove(drg.SelectedRows(0))
        Myconn.ClearAllControls(GroupBox2, True)
    End Sub

    Private Sub btnUpdat_Click(sender As Object, e As EventArgs) Handles btnUpdat.Click
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is TextBox Then
                If txt.Text = "" And txt.Name <> "txtNotes" Then
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

    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        If txtSearch.Text = "" Then Return
        If txtSearch.Text.IndexOf("/") > -1 Then
            y = 2
            Filldrg()
        Else
            y = 1
            Filldrg()
        End If
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        If drg.CurrentRow.Cells(9).Value = True Then
            With Myconn
                .Parames.Clear()
                .Addparam("@Status", False)
                .Addparam("@ID", drg.CurrentRow.Cells(10).Value)

            End With
            drg.CurrentRow.Cells(9).Value = False
            drg.CurrentRow.DefaultCellStyle.BackColor = Color.Red
        Else
            With Myconn
                .Parames.Clear()
                .Addparam("@Status", True)
                .Addparam("@ID", drg.CurrentRow.Cells(10).Value)
            End With
            drg.CurrentRow.Cells(9).Value = True
            drg.CurrentRow.DefaultCellStyle.BackColor = Color.LemonChiffon
        End If
        Myconn.ExecQuery(" Update  Safe_Recive_per set Status = @Status where ID = @ID")

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label12.Text = TimeOfDay.ToString("hh:mm:ss tt", CultureInfo.CreateSpecificCulture("ar-eg"))
    End Sub
    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Binding()

    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Dim table As New DataTable
        For i As Integer = 1 To 9
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(1).Value
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(2).Value
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(8).Value
            table.Rows(table.Rows.Count - 1)(8) = dr.Cells(9).Value
        Next
        Dim rpt As New rpt_recive
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub
    Private Sub drg_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellDoubleClick
        drg.CurrentRow.Selected = False
    End Sub
    Private Sub RB2_CheckedChanged(sender As Object, e As EventArgs) Handles RB2.CheckedChanged
        cbo_Customer.DataSource = Nothing
        If RB2.Checked = True Then
            Myconn.Fillcombo("select * from Bank order by Bank_Name", "Bank", "Bank_ID", "Bank_Name", Me, cbo_Customer)
        Else
            Myconn.Fillcombo("select * from Customers order by Customer_Name", "Customers", "Customer_ID", "Customer_Name", Me, cbo_Customer)
        End If
    End Sub
End Class