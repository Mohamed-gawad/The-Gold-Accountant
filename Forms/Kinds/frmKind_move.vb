Public Class frmKind_move
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st, st1 As String
    Dim y, x As Integer

    Private Sub Filldrg()
        drg.Rows.Clear()
        x = 0
        Select Case y
            Case 0
                If String.IsNullOrWhiteSpace(txt1.Text) OrElse String.IsNullOrWhiteSpace(txt2.Text) Then
                    st = " where Items.items_Cod =" & CInt(cbo_kind.ComboBox.SelectedValue) & " and Sales.Status = True"
                Else
                    st = " where Items.items_Cod =" & CInt(cbo_kind.ComboBox.SelectedValue) & " and  Sales.Sales_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "# and Sales.Status = True"
                End If

            Case 1
                If String.IsNullOrWhiteSpace(txt1.Text) OrElse String.IsNullOrWhiteSpace(txt2.Text) Then
                    st = " where Items.Parcode like '" & txtBarcode.Text & "' and Sales.Status = True"
                Else
                    st = " where Items.Parcode like '" & txtBarcode.Text & "' and  Sales.Sales_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "# and Sales.Status = True"
                End If

        End Select



        Myconn.ExecQuery("Select Items.items_Cod,Items.Parcode, Items.Items_Name, Items.cost_Price, Sales.Sales_Date, Sales.Sales_Bill_ID, Sales.Items_Price, Sales.Items_num, Sales.Total_Price, Sales.Reduce, Sales.Final_Total_Price, Sales.Final_Total_Price, Sales.Users_ID,Sales.Status,sales.ID ,Sales.Earning, Customers.Customer_Name, Users_ID.Employee_Name
                            FROM ((Items LEFT JOIN Sales ON Items.items_Cod = Sales.items_Cod) LEFT JOIN Customers ON Sales.Customer_ID = Customers.Customer_ID) LEFT JOIN Users_ID ON Sales.Users_ID = Users_ID.Employee_ID
                             " & st & " order by Sales.ID ")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        Dim V1 As Double = 0
        Dim V2 As Double = 0

        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = Format(CDate(r("Sales_Date")), "yyyy/MM/dd")
            drg.Rows(i).Cells(2).Value = r("Sales_Bill_ID")
            drg.Rows(i).Cells(3).Value = r("Customer_Name")
            drg.Rows(i).Cells(4).Value = r("Items_Name")
            drg.Rows(i).Cells(5).Value = r("Parcode")
            drg.Rows(i).Cells(6).Value = r("Items_num")
            drg.Rows(i).Cells(7).Value = r("Items_Price")
            drg.Rows(i).Cells(8).Value = r("Reduce")
            drg.Rows(i).Cells(9).Value = r("Final_Total_Price")
            drg.Rows(i).Cells(10).Value = r("Employee_Name")
            drg.Rows(i).Cells(11).Value = r("Status")
            drg.Rows(i).Cells(12).Value = r("ID")
            V1 += r("Items_num")
            x = r("items_Cod")
        Next
        Label6.Text = V1

        '---------------------------------------------------------------------------------- المشتريات
        drg1.Rows.Clear()
        Select Case y
            Case 0
                If String.IsNullOrWhiteSpace(txt1.Text) OrElse String.IsNullOrWhiteSpace(txt2.Text) Then
                    st1 = " where Items.items_Cod =" & CInt(cbo_kind.ComboBox.SelectedValue) & " and Purchases.Status = True"
                Else
                    st1 = " where Items.items_Cod =" & CInt(cbo_kind.ComboBox.SelectedValue) & " and  Purchases.Pur_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "# and Purchases.Status = True"
                End If
            Case 1

                If String.IsNullOrWhiteSpace(txt1.Text) OrElse String.IsNullOrWhiteSpace(txt2.Text) Then
                    st1 = " where Items.Parcode like '" & txtBarcode.Text & "'  and Purchases.Status = True"
                Else
                    st1 = " where Items.Parcode like '" & txtBarcode.Text & "'  and  Purchases.Pur_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "# and Purchases.Status = True"
                End If

        End Select




        Myconn.ExecQuery("SELECT Purchases.ID, Purchases.Pur_Date,Items.Items_Name,Items.Parcode, Purchases.Pur_Bill_num, Purchases.Supplier_ID, Purchases.items_Cod, Purchases.Customer_Price, Purchases.Reduce, Purchases.Pur_Price, Purchases.Iteme_Number, Purchases.Total_Price, Purchases.Status, Supplier.Supplier_Name, Users_ID.Employee_Name
                            FROM Users_ID RIGHT JOIN ((Items LEFT JOIN Purchases ON Items.items_Cod = Purchases.items_Cod) LEFT JOIN Supplier ON Purchases.Supplier_ID = Supplier.Supplier_ID) ON Users_ID.Employee_ID = Purchases.Employee_ID
                            " & st1 & " order by Purchases.ID ")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        If Myconn.Recodcount = 0 Then MsgBox("الصنف غير موجود", MsgBoxStyle.MsgBoxRtlReading, "رسالة") : cbo_kind.SelectedIndex = -1

        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg1.Rows.Add()
            drg1.Rows(i).Cells(0).Value = i + 1
            drg1.Rows(i).Cells(1).Value = Format(CDate(r("Pur_Date")), "yyyy/MM/dd")
            drg1.Rows(i).Cells(2).Value = r("Pur_Bill_num")
            drg1.Rows(i).Cells(3).Value = r("Supplier_Name")
            drg1.Rows(i).Cells(4).Value = r("Items_Name")
            drg1.Rows(i).Cells(5).Value = r("Parcode")
            drg1.Rows(i).Cells(6).Value = r("Iteme_Number")
            drg1.Rows(i).Cells(7).Value = r("Customer_Price")
            drg1.Rows(i).Cells(8).Value = r("Reduce")
            drg1.Rows(i).Cells(9).Value = r("Pur_Price")
            drg1.Rows(i).Cells(10).Value = r("Total_Price")
            drg1.Rows(i).Cells(11).Value = r("Employee_Name")
            drg1.Rows(i).Cells(12).Value = r("Status")
            drg1.Rows(i).Cells(13).Value = r("ID")
            V2 += r("Iteme_Number")
        Next
        Label4.Text = V2
        Label7.Text = V2 - V1
    End Sub

    Private Sub frmKind_move_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width

            If  F <> 1Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Exit Sub
                Dim r As DataRow = Myconn.dt.Rows(0)
                If r("U_full").ToString = False Then
                    btnSearch.Enabled = r("U_search").ToString
                    btnPrint.Enabled = r("U_print").ToString
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        fin = False
        Myconn.Fillcombo("Select Items_Name, items_Cod from [items] order by [Items_Name]", "[items]", "items_Cod", "Items_Name", Me, cbo_kind.ComboBox)
        fin = True
    End Sub
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        y = 0
        Filldrg()
    End Sub
    Private Sub cbo_kind_Enter(sender As Object, e As EventArgs) Handles cbo_kind.Enter
        Myconn.langAR()
    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Select Case cboPrint.SelectedIndex
            Case 0 ' مشتريات الصنف
                Print_Pur_Kind()
            Case 1 ' مبيعات الصنف
                Print_Sales_Kind()
        End Select
    End Sub
    Private Sub Print_Pur_Kind() 'مشتريات الصنف
        Dim table As New DataTable
        For i As Integer = 1 To 11
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg1.Rows
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(2).Value ' رقم الفاتورة
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(1).Value ' التاريخ
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value ' المورد
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value ' الصنف
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value ' الباركود
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(7).Value ' سعر المصنع 
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(6).Value ' العدد
            table.Rows(table.Rows.Count - 1)(8) = Math.Round((dr.Cells(7).Value * dr.Cells(6).Value), 2) ' الاجمالي
            table.Rows(table.Rows.Count - 1)(9) = dr.Cells(8).Value ' الخصم
            table.Rows(table.Rows.Count - 1)(10) = dr.Cells(10).Value ' سعر التكلفة بعد الخصم
        Next
        Dim rpt As New rpt_Pur_Kind
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)

        rpt.SetParameterValue("Price", Label4.Text)
        rpt.SetParameterValue("Bill_num", cbo_kind.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub
    Private Sub Print_Sales_Kind() 'مبيعات الصنف
        Dim table As New DataTable
        For i As Integer = 1 To 11
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(2).Value ' رقم الفاتورة
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(1).Value ' التاريخ
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value ' المورد
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value ' الصنف
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value ' الباركود
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(7).Value ' سعر  
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(6).Value ' العدد
            table.Rows(table.Rows.Count - 1)(8) = Math.Round((dr.Cells(7).Value * dr.Cells(6).Value), 2) ' الاجمالي
            table.Rows(table.Rows.Count - 1)(9) = dr.Cells(8).Value ' الخصم
            table.Rows(table.Rows.Count - 1)(10) = dr.Cells(9).Value ' سعر التكلفة بعد الخصم
        Next
        Dim rpt As New rpt_Pur_Kind
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)

        rpt.SetParameterValue("Price", Label6.Text)
        rpt.SetParameterValue("Bill_num", cbo_kind.Text)

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
            y = 1
            Filldrg()
        End If
        cbo_kind.ComboBox.SelectedValue = x
    End Sub

    Private Sub txtBarcode_Enter(sender As Object, e As EventArgs) Handles txtBarcode.Enter
        txtBarcode.Text = Nothing
    End Sub

    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Dim frm As New frmBill_Sale
        frm.MdiParent = FrmMain
        frm.txtSearch.Text = drg.CurrentRow.Cells(2).Value

        frm.Show()
        frm.btnSearch_Click(Nothing, Nothing)
    End Sub

    Private Sub drg1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg1.CellClick
        Dim frm As New frmPurchases_Bill
        Dim T As String = drg1.CurrentRow.Cells(5).Value
        frm.MdiParent = FrmMain
        frm.txtSearch.Text = drg1.CurrentRow.Cells(2).Value

        frm.drg.ClearSelection()
        frm.Show()
        frm.btnSearch_Click(Nothing, Nothing)
        For W As Integer = 0 To frm.drg.Rows.Count - 1

            If frm.drg.Rows(W).Cells(3).Value.ToString.Equals(T, StringComparison.CurrentCultureIgnoreCase) Then
                frm.drg.Rows(W).Cells(2).Selected = True
                frm.drg.CurrentCell = frm.drg.SelectedCells(1)
                Exit For
            End If
        Next
    End Sub
End Class