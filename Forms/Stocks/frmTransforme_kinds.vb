Imports System.Globalization
Public Class frmTransforme_kinds
    Dim Myconn As New Connection
    Dim fin As Boolean
    Dim x As Integer
    Dim st As String
    Sub New_record()
        fin = False
        Myconn.ClearAllControls(GroupBox1, True)
        Myconn.ClearAllControls(GroupBox2, True)
        fin = True
        Myconn.Autonumber("move_Id", "[Items_move]", txtID, Me)
    End Sub
    Sub Filldrg()
        drg.Rows.Clear()
        Myconn.ExecQuery("SELECT Items_move.ID,Items_move.move_Id, Items_move.move_date, Items_move.move_time, Items_move.items_Cod, Items.Items_Name,Items.Parcode, Items_move.items_amount,Items_move.Stock_from,(Stocks_1.Stock_Name) as Stock_f, Items_move.Stock_to, (Stocks.Stock_Name) as Stock_t, Users_ID.Employee_Name, Items_move.Notes
                            FROM (((Items RIGHT JOIN Items_move ON Items.items_Cod = Items_move.items_Cod) LEFT JOIN Stocks ON Items_move.Stock_to = Stocks.Stock_ID) LEFT JOIN Users_ID ON Items_move.Employee_ID = Users_ID.Employee_ID) LEFT JOIN Stocks AS Stocks_1 ON Items_move.Stock_from = Stocks_1.Stock_ID
                           " & st & " order by Items_move.move_Id")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub

        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("move_Id")
            drg.Rows(i).Cells(2).Value = r("move_time")
            drg.Rows(i).Cells(3).Value = Format(CDate(r("move_date")), "yyyy/MM/dd")
            drg.Rows(i).Cells(4).Value = r("Items_Name")
            drg.Rows(i).Cells(5).Value = r("Parcode")
            drg.Rows(i).Cells(6).Value = r("items_amount")
            drg.Rows(i).Cells(7).Value = r("Stock_f")
            drg.Rows(i).Cells(8).Value = r("Stock_t")
            drg.Rows(i).Cells(9).Value = r("Employee_Name")
            drg.Rows(i).Cells(10).Value = r("ID")
        Next
        Myconn.DataGridview_MoveLast(drg, 2)
    End Sub
    Private Sub Binding()

        Select Case x
            Case 0 ' المخزن الذي به البضاعة
                Myconn.ExecQuery("Select i.Parcode, i.Customer_Price, i.cost_Price, i.Total_Price, i.Supplier_ID, IIf(ISNULL((c.Pur_num + T.amount_t) - (s.Sales_num + F.amount_F)), 0,((c.Pur_num + T.amount_t) - (s.Sales_num + F.amount_F))) as rest 
                                    from ((( Items i Left join  (select iif(ISNULL(sum(Iteme_Number)),0,sum(Iteme_Number)) as Pur_num,items_cod from Purchases group by items_Cod,Status,Stock_ID having Status = True and Stock_ID = " & CInt(cboStock1.SelectedValue) & " ) c
                                    on i.items_Cod = c.items_Cod  )
                                    left join (Select iif(ISNULL(sum(Items_num)),0,sum(Items_num)) as Sales_num ,items_Cod From Sales group by items_cod,Status,Stock_ID having Status = True and Stock_ID = " & CInt(cboStock1.SelectedValue) & "   ) S
                                    on i.items_Cod = S.items_Cod)
                                    left join  (Select iif(Isnull(Sum(items_amount)),0,Sum(items_amount)) as amount_F, items_cod From Items_move  group by items_cod,Stock_From having Stock_From = " & CInt(cboStock1.SelectedValue) & " ) F
                                    on  i.items_Cod = F.items_cod)
                                    left join  (Select iif(Isnull(Sum(items_amount)),0,Sum(items_amount)) as amount_t, items_cod From Items_move  group by items_cod,Stock_to having Stock_to = " & CInt(cboStock1.SelectedValue) & " ) T
                                    on  i.items_Cod = T.items_cod
                                    group by i.items_Cod, i.Parcode,i.Customer_Price,i.cost_Price,i.Total_Price,i.Supplier_ID, IIf(ISNULL((c.Pur_num + T.amount_t) - (s.Sales_num + F.amount_F)), 0,((c.Pur_num + T.amount_t) - (s.Sales_num + F.amount_F)))
                                    having i.items_Cod =" & CInt(cbokind.SelectedValue))
                If Myconn.Recodcount = 0 Then Return
                Dim r As DataRow = Myconn.dt.Rows(0)
                txtAmount1.Text = r("rest")

            Case 1 ' المخزن الذي ستنقل اليه البضاعة

                Myconn.ExecQuery("Select i.Parcode, i.Customer_Price, i.cost_Price, i.Total_Price, i.Supplier_ID, IIf(ISNULL((c.Pur_num + T.amount_t) - (s.Sales_num + F.amount_F)), 0,((c.Pur_num + T.amount_t) - (s.Sales_num + F.amount_F))) as rest 
                                    from ((( Items i Left join  (select iif(ISNULL(sum(Iteme_Number)),0,sum(Iteme_Number)) as Pur_num,items_cod from Purchases group by items_Cod,Status,Stock_ID having  Status = True and Stock_ID = " & CInt(cboStock2.SelectedValue) & "  ) c
                                    on i.items_Cod = c.items_Cod  )
                                    left join (Select iif(ISNULL(sum(Items_num)),0,sum(Items_num)) as Sales_num ,items_Cod From Sales group by items_cod,Status,Stock_ID having  Status = True and Stock_ID = " & CInt(cboStock2.SelectedValue) & "  ) S
                                    on i.items_Cod = S.items_Cod)
                                    left join  (Select iif(Isnull(Sum(items_amount)),0,Sum(items_amount)) as amount_F, items_cod From Items_move  group by items_cod,Stock_From having Stock_From = " & CInt(cboStock2.SelectedValue) & " ) F
                                    on  i.items_Cod = F.items_cod)
                                    left join  (Select iif(Isnull(Sum(items_amount)),0,Sum(items_amount)) as amount_t, items_cod From Items_move  group by items_cod,Stock_to having Stock_to = " & CInt(cboStock2.SelectedValue) & " ) T
                                    on  i.items_Cod = T.items_cod
                                    group by i.items_Cod, i.Parcode,i.Customer_Price,i.cost_Price,i.Total_Price,i.Supplier_ID, IIf(ISNULL((c.Pur_num + T.amount_t) - (s.Sales_num + F.amount_F)), 0,((c.Pur_num + T.amount_t) - (s.Sales_num + F.amount_F)))
                                    having i.items_Cod =" & CInt(cbokind.SelectedValue))
                If Myconn.Recodcount = 0 Then Return
                Dim r As DataRow = Myconn.dt.Rows(0)
                txtAmount2.Text = r("rest")
        End Select



    End Sub
    Private Sub Save_record()
        With Myconn
            .Parames.Clear()
            .Addparam("@move_Id", txtID.Text)
            .Addparam("@move_date", Format(Now.Date, "yyyy/MM/dd"))
            .Addparam("@move_time", Label10.Text)
            .Addparam("@items_Cod", cbokind.SelectedValue)
            .Addparam("@items_amount", txtAmount_trans.Text)
            .Addparam("@Stock_from", cboStock1.SelectedValue)
            .Addparam("@Stock_to", cboStock2.SelectedValue)
            .Addparam("@Employee_ID", My.Settings.user_ID)
            .Addparam("@Notes", txtNotes.Text)

            .ExecQuery("insert into [Items_move] (move_Id,move_date,move_time,items_Cod,items_amount,Stock_from,Stock_to,Employee_ID,Notes)
                                  values(@move_Id,@move_date,@move_time,@items_Cod,@items_amount,@Stock_from,@Stock_to,@Employee_ID,@Notes)")

            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
    End Sub
    Private Sub Update_record()
        With Myconn
            .Parames.Clear()
            .Addparam("@move_date", Format(Now.Date, "yyyy/MM/dd"))
            .Addparam("@move_time", Label10.Text)
            .Addparam("@items_Cod", cbokind.SelectedValue)
            .Addparam("@items_amount", txtAmount_trans.Text)
            .Addparam("@Stock_from", cboStock1.SelectedValue)
            .Addparam("@Stock_to", cboStock2.SelectedValue)
            .Addparam("@Notes", txtNotes.Text)
            .Addparam("@Employee_ID", My.Settings.user_ID)
            .Addparam("@ID", drg.CurrentRow.Cells(10).Value)

            .ExecQuery("update  [Items_move] set move_date=@move_date,move_time=@move_time,items_Cod=@items_Cod,items_amount=@items_amount,Stock_from=@Stock_from,Stock_to=@Stock_to,Notes=@Notes,Employee_ID=@Employee_ID where ID = @ID")

            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
    End Sub
    Private Sub frmTransforme_kinds_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label6.Left = 0
            Label6.Width = Me.Width

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
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        fin = False
        Myconn.Fillcombo("Select Items_Name, items_Cod from [items] order by [Items_Name]", "[items]", "items_Cod", "Items_Name", Me, cbokind)
        Myconn.Fillcombo("select * from [Stocks] order by Stock_ID", "[Stocks]", "Stock_ID", "Stock_Name", Me, cboStock1)
        Myconn.Fillcombo("select * from [Stocks] order by Stock_ID", "[Stocks]", "Stock_ID", "Stock_Name", Me, cboStock2)
        fin = True
        Timer1.Start()
        New_record()
        Filldrg()
    End Sub
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        New_record()
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
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
        If cboStock1.SelectedValue = cboStock2.SelectedValue Then MsgBox("لا يمكن نقل البضاعة من المخزن إلى نفسه", MsgBoxStyle.Critical, "تنبيه") : Return
        Myconn.ExecQuery("Select * from Items_move where move_Id =" & CInt(txtID.Text))
        If Myconn.Recodcount > 0 Then MsgBox(" رقم هذا الإذن مسجل من قبل", MsgBoxStyle.Critical & MsgBoxStyle.MsgBoxRtlReading, "رسالة") : Return

        Save_record()

        Filldrg()
        cbokind_SelectedIndexChanged(Nothing, Nothing)
        MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

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
        If cboStock1.SelectedValue = cboStock2.SelectedValue Then MsgBox("لا يمكن نقل البضاعة من المخزن إلى نفسه", MsgBoxStyle.Critical, "تنبيه") : Return

        Update_record()
        Filldrg()
        cbokind_SelectedIndexChanged(Nothing, Nothing)
        MessageBox.Show("تمت عملية التعديل بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)


    End Sub
    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        With Myconn
            .Addparam("@ID", drg.CurrentRow.Cells(10).Value)
            .ExecQuery("delete from [Items_move] where ID = @ID ")
        End With
        If Myconn.NoErrors(True) = False Then Exit Sub
        drg.Rows.Remove(drg.SelectedRows(0))
        cbokind_SelectedIndexChanged(Nothing, Nothing)
    End Sub
    Private Sub cbokind_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbokind.SelectedIndexChanged
        If Not fin Then Return
        If cboStock1.SelectedIndex = -1 Then MsgBox("من فضلك قم باختيار المخزن الذي به البضاعة", MsgBoxStyle.Critical & MsgBoxStyle.MsgBoxRtlReading, "رسالة") : Return
        x = 0
        Binding()
        x = 1
        Binding()
    End Sub

    Private Sub cboStock2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStock2.SelectedIndexChanged
        If Not fin Then Return
        If cboStock2.SelectedIndex = -1 Then MsgBox("من فضلك قم باختيار المخزن الذي ستنقل إليه البضاعة", MsgBoxStyle.Critical & MsgBoxStyle.MsgBoxRtlReading, "رسالة") : Return
        If cbokind.SelectedIndex = -1 Then Return
        x = 1
        Binding()

    End Sub

    Private Sub cboStock1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStock1.SelectedIndexChanged
        If Not fin Then Return
        If cbokind.SelectedIndex = -1 Then Return
        x = 0
        Binding()

    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label10.Text = TimeOfDay.ToString("hh:mm:ss tt", CultureInfo.CreateSpecificCulture("ar-eg"))
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        st = " where Items_move.move_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
        Filldrg()

    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Dim table As New DataTable
        For i As Integer = 1 To 11
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
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(7).Value
            table.Rows(table.Rows.Count - 1)(8) = dr.Cells(8).Value
            table.Rows(table.Rows.Count - 1)(9) = dr.Cells(9).Value
        Next
        Dim rpt As New rpt_Transforme_kinds
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Bill_num", txt1.Text)
        rpt.SetParameterValue("F_date", txt2.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub

    Private Sub txtBarcode_KeyUp(sender As Object, e As KeyEventArgs) Handles txtBarcode.KeyUp
        If e.KeyCode = Keys.Enter Then
            Myconn.ExecQuery("Select * from Items where Parcode Like '" & txtBarcode.Text & "'")
            If Myconn.Recodcount = 0 Then MsgBox(" الصنف غير موجود", MsgBoxStyle.Critical & MsgBoxStyle.MsgBoxRtlReading, "رسالة") : Return
            Dim r As DataRow = Myconn.dt.Rows(0)
            cbokind.SelectedValue = r("items_Cod")
            txtAmount_trans.Focus()

        End If
    End Sub

    Private Sub txtAmount_trans_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAmount_trans.KeyPress
        Myconn.NumberOnly(txtAmount_trans, e)
    End Sub

    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Myconn.ExecQuery("Select * from Items_move where ID =" & CInt(drg.CurrentRow.Cells(10).Value))
        If Myconn.Recodcount = 0 Then Return
        Dim r As DataRow = Myconn.dt.Rows(0)

        cboStock1.SelectedValue = r("Stock_from")
        cboStock2.SelectedValue = r("Stock_to")
        cbokind.SelectedValue = r("items_Cod")
        txtAmount_trans.Text = r("items_amount")
        txtBarcode.Text = drg.CurrentRow.Cells(5).Value
        txtID.Text = r("move_Id")
        txtNotes.Text = r("Notes")
    End Sub
End Class