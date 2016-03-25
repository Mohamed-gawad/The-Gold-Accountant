Public Class frmList_kinds
    Private Myconn As New Connection
    Dim fin As Boolean
    Dim St As String
    Sub Filldrg()
        drg.Rows.Clear()
        Myconn.ExecQuery("SELECT Items.Items_Name, Items.Parcode, Items.Customer_Price, Items.Total_Price, Items.cost_Price, group.group_Name, Supplier.Supplier_Name, Items.items_Cod
                            FROM (Items LEFT JOIN [group] ON Items.Group_cod = group.Group_cod) LEFT JOIN Supplier ON Items.Supplier_ID = Supplier.Supplier_ID" & St & "
                            ORDER BY Items.Items_Name;")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub

        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("Parcode")
            drg.Rows(i).Cells(2).Value = r("Items_Name")
            drg.Rows(i).Cells(3).Value = r("cost_Price")
            drg.Rows(i).Cells(4).Value = r("Customer_Price")
            drg.Rows(i).Cells(5).Value = r("Total_Price")
            drg.Rows(i).Cells(6).Value = r("group_Name")
            drg.Rows(i).Cells(7).Value = r("Supplier_Name")
            drg.Rows(i).Cells(8).Value = r("items_Cod")
        Next
        Myconn.DataGridview_MoveLast(drg, 2)
    End Sub
    Sub Update_record()
        Try

            Myconn.ExecQuery("SELECT Items.Items_Name, Items.Parcode, Items.Customer_Price, Items.Total_Price, Items.cost_Price, group.group_Name, Supplier.Supplier_Name, Items.items_Cod
                            FROM (Items LEFT JOIN [group] ON Items.Group_cod = group.Group_cod) LEFT JOIN Supplier ON Items.Supplier_ID = Supplier.Supplier_ID where Items.items_Cod = " & drg.CurrentRow.Cells(8).Value & "
                            ORDER BY Items.Items_Name;")

            If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
            Dim r As DataRow = Myconn.dt.Rows(0)
            drg.CurrentRow.Cells(1).Value = r("Parcode")
            drg.CurrentRow.Cells(2).Value = r("Items_Name")
            drg.CurrentRow.Cells(3).Value = r("cost_Price")
            drg.CurrentRow.Cells(4).Value = r("Customer_Price")
            drg.CurrentRow.Cells(5).Value = r("Total_Price")
            drg.CurrentRow.Cells(6).Value = r("group_Name")
            drg.CurrentRow.Cells(7).Value = r("Supplier_Name")
            drg.CurrentRow.Cells(8).Value = r("items_Cod")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical & MsgBoxStyle.MsgBoxRtlReading, "رسالة")
        End Try
    End Sub
    Sub Binding()
        Try
            Myconn.ExecQuery("SELECT Items.Items_Name, Items.Parcode, Items.Customer_Price, Items.Total_Price, Items.cost_Price, group.group_Name, Supplier.Supplier_Name, Items.items_Cod
                            FROM (Items LEFT JOIN [group] ON Items.Group_cod = group.Group_cod) LEFT JOIN Supplier ON Items.Supplier_ID = Supplier.Supplier_ID 
                            where  Items.items_Cod = " & CInt(drg.CurrentRow.Cells(8).Value) & " ORDER BY Items.Items_Name ")
            Dim r As DataRow = Myconn.dt.Rows(0)
            txtID.Text = r("items_Cod").ToString
            txtName.Text = r("Items_Name").ToString
            txtParcode.Text = r("Parcode").ToString
            txtCustomer.Text = If(IsDBNull(r("Customer_Price")), "", r("Customer_Price"))
            txtCost.Text = If(IsDBNull(r("cost_Price")), "", r("cost_Price"))
            txtTotal.Text = If(IsDBNull(r("Total_Price")), "", r("Total_Price"))
            cbo_Group.Text = If(IsDBNull(r("group_Name")), "", r("group_Name"))
            cbo_Supplier.Text = If(IsDBNull(r("Supplier_Name")), Nothing, r("Supplier_Name"))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical & MsgBoxStyle.MsgBoxRtlReading, "رسالة")
        End Try
    End Sub
    Sub New_record()
        Myconn.ClearAllControls(GroupBox1, True)
        Myconn.Autonumber("items_Cod", "Items", txtID, Me)
        txtParcode.Text = Format(CInt(txtID.Text), "000000")
    End Sub
    Sub Save_record()
        With Myconn
            .Parames.Clear()
            .Addparam("@items_Cod", txtID.Text)
            .Addparam("@Items_Name", txtName.Text)
            .Addparam("@Group_cod", cbo_Group.SelectedValue)
            .Addparam("@Parcode", txtParcode.Text)
            .Addparam("@Customer_Price", If(txtCustomer.Text = "", DBNull.Value, txtCustomer.Text))
            .Addparam("@Total_Price", If(txtTotal.Text = "", DBNull.Value, txtTotal.Text))
            .Addparam("@cost_Price", If(txtCost.Text = "", DBNull.Value, txtCost.Text))
            .Addparam("@Supplier_ID", cbo_Supplier.SelectedValue)
            .ExecQuery("insert into  Items (items_Cod,Items_Name,Group_cod,Parcode,Customer_Price,Total_Price,cost_Price,Supplier_ID) 
                                     values(@items_Cod,@Items_Name,@Group_cod,@Parcode,@Customer_Price,@Total_Price,@cost_Price,@Supplier_ID)")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
        MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
    End Sub
    Private Sub Check_Parcode()
        Try
            Myconn.ExecQuery("Select * from Items where Parcode ='" & txtParcode.Text & "'")
            If Myconn.Recodcount = 0 Then Return
            Dim r As DataRow = Myconn.dt.Rows(0)
            If Myconn.Recodcount > 0 Then
                MsgBox("  الباركود " & r("Parcode").ToString & " مسجل باسم الصنف" & vbNewLine & r("Items_name").ToString, MsgBoxStyle.MsgBoxRtlReading & MsgBoxStyle.Critical, "رسالة")
                txtParcode.Text = ""
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical & MsgBoxStyle.MsgBoxRtlReading, "رسالة")
        End Try
    End Sub
    Private Sub frmNew_kind_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label5.Left = 0
        Label5.Width = Me.Width
        fin = False
        Myconn.Fillcombo("select * from [group] order by [group_Name]", "[group]", "Group_cod", "group_Name", Me, cbo_Group)
        Myconn.Fillcombo("select * from Supplier order by Supplier_Name", "Supplier", "Supplier_ID", "Supplier_Name", Me, cbo_Supplier)
        fin = True

    End Sub
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        New_record()
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Check_Parcode()
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is TextBox Then
                If txt.Text = "" AndAlso txt.Name <> "txtCost" AndAlso txt.Name <> "txtCustomer" AndAlso txt.Name <> "txtTotal" Then
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

        Save_record()
        St = " where Items.items_Cod =" & CInt(txtID.Text)
        Filldrg()
        New_record()
    End Sub
    Private Sub btnUpdat_Click(sender As Object, e As EventArgs) Handles btnUpdat.Click
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is TextBox Then
                If txt.Text = "" AndAlso txt.Name <> "txtCost" AndAlso txt.Name <> "txtCustomer" AndAlso txt.Name <> "txtTotal" Then
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
        If Myconn.NoErrors(True) = False Then Exit Sub
        With Myconn
            .Parames.Clear()
            .Addparam("@Items_Name", txtName.Text)
            .Addparam("@Group_cod", cbo_Group.SelectedValue)
            .Addparam("@Parcode", txtParcode.Text)
            .Addparam("@Customer_Price", If(txtCustomer.Text = "", DBNull.Value, txtCustomer.Text))
            .Addparam("@Total_Price", If(txtTotal.Text = "", DBNull.Value, txtTotal.Text))
            .Addparam("@cost_Price", If(txtCost.Text = "", DBNull.Value, txtCost.Text))
            .Addparam("@Supplier_ID", If(cbo_Supplier.SelectedIndex = -1, DBNull.Value, cbo_Supplier.SelectedValue))
            .Addparam("@items_Cod", txtID.Text)
            .ExecQuery("Update Items Set Items_Name =@Items_Name ,Group_cod =@Group_cod ,Parcode =@Parcode ,Customer_Price =@Customer_Price ,Total_Price =@Total_Price ,cost_Price =@cost_Price ,Supplier_ID =@Supplier_ID where items_Cod =@items_Cod ")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
        Update_record()
        MessageBox.Show("تمت عملية التعديل بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
    End Sub
    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        Myconn.Addparam("@Items_Code", txtID.Text)
        Myconn.ExecQuery("Delete from Items where items_Cod = @Items_Code ")
        If Myconn.NoErrors(True) = False Then Exit Sub
        drg.Rows.Remove(drg.SelectedRows(0))
        Myconn.ClearAllControls(GroupBox1, True)
    End Sub
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        If cboArgment.SelectedIndex = -1 Then Return
        Select Case cboArgment.SelectedIndex
            Case 0
                St = " where Items.Group_cod =" & cboSearch.ComboBox.SelectedValue
            Case 1
                St = " where Items.Items_Name like '%" & cboSearch.Text & "%'"
            Case 2
                St = ""
            Case 3
                St = " where Items.Parcode like '" & cboSearch.Text & "'"

        End Select
        Filldrg()
        Binding()
    End Sub
    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Binding()
    End Sub
    Private Sub cboArgment_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboArgment.SelectedIndexChanged
        Select Case cboArgment.SelectedIndex
            Case 0
                Myconn.Fillcombo("Select * from [group] order by [group_Name]", "[group]", "Group_cod", "group_Name", Me, cboSearch.ComboBox)
            Case 1
                cboSearch.ComboBox.DataSource = Nothing
                cboSearch.Items.Clear()
            Case 2
                cboSearch.ComboBox.DataSource = Nothing
                cboSearch.Items.Clear()
            Case 3
                cboSearch.ComboBox.DataSource = Nothing
                cboSearch.Items.Clear()
        End Select
    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Print_drg()
    End Sub
    Sub Print_drg()
        Try
            Dim rpt As New rpt_Kinds
            Dim table As New DataTable
            For i As Integer = 1 To 7
                Dim x As String
                x = Format(i, "00")
                table.Columns.Add(x)
            Next

            For Each dr As DataGridViewRow In drg.Rows
                table.Rows.Add()
                table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count ' المسلسل
                table.Rows(table.Rows.Count - 1)(1) = dr.Cells(2).Value ' الصنف
                table.Rows(table.Rows.Count - 1)(2) = dr.Cells(1).Value ' الباركود
                table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value 'التكلفة
                table.Rows(table.Rows.Count - 1)(4) = dr.Cells(5).Value ' الجملة
                table.Rows(table.Rows.Count - 1)(5) = dr.Cells(4).Value ' المستهلك
                table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value ' المجموعة
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

    Private Sub cboSearch_KeyUp(sender As Object, e As KeyEventArgs) Handles cboSearch.KeyUp
        If cboArgment.SelectedIndex = -1 Then Return
        If e.KeyCode = Keys.Enter Then
            Select Case cboArgment.SelectedIndex
                Case 0
                    St = " where Items.Group_cod =" & cboSearch.ComboBox.SelectedValue
                Case 1
                    St = " where Items.Items_Name like '%" & cboSearch.Text & "%'"
                Case 2
                    St = ""
                Case 3
                    St = " where Items.Parcode like '" & cboSearch.Text & "'"
            End Select
            Filldrg()
            Binding()
        End If
    End Sub

End Class