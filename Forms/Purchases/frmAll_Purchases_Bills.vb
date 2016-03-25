Public Class frmAll_Purchases_Bills
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st As String
    Private Sub fillgrd()
        drg.Rows.Clear()

        Select Case cboSearch.SelectedIndex
            Case 0 '  كل الفواتير
                st = Nothing
            Case 1 ' فواتير خلال فترة
                st = " and p.Pur_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
            Case 2 ' فواتير مورد
                st = " and p.Supplier_ID = " & CInt(cboSupplier.ComboBox.SelectedValue)
            Case 3 ' فواتير مورد خلال فترة
                st = " and p.Pur_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#" & " and p.Supplier_ID = " & CInt(cboSupplier.ComboBox.SelectedValue)
        End Select
        Myconn.ExecQuery("Select  p.Pur_Bill_num,p.Pur_Date,s.supplier_name, iif( isnull(sum(Customer_Price * Iteme_Number)),0,sum(Customer_Price*Iteme_Number)) As Customer_Price2,count(items_Cod) As Kind_num,
                                    iif(isnull(sum(Pur_Price * Iteme_Number)),0,sum(Pur_Price * Iteme_Number)) as Cost,iif(isnull(b.back),0,b.back) as back2,p.supplier_ID,iif(isnull(b.back_cost),0,(b.back_cost)) as back_cost2
                                    From (Purchases p left join (select count(Iteme_Number) as back,sum(Pur_Price * Iteme_Number) as back_cost,Pur_Bill_num from Purchases group by Pur_Bill_num,Status having Status = false   ) b
                                       on  p.Pur_Bill_num = b.Pur_Bill_num  )
                                    left join Supplier  s on p.supplier_ID = s.Supplier_ID
                                    group by b.back,p.Pur_Bill_num,p.Pur_Date,s.supplier_name,P.supplier_ID,b.back_cost,p.Status having p.Status = true " & st & " and P.Pur_Bill_num > 0 order by p.Pur_Bill_num")


        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        Dim V1 As Double = 0
        Dim V2 As Double = 0
        Dim B As Double = 0
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("Pur_Bill_num")
            drg.Rows(i).Cells(2).Value = Format(CDate(r("Pur_Date")), "yyyy/MM/dd")
            drg.Rows(i).Cells(3).Value = r("supplier_name")
            drg.Rows(i).Cells(4).Value = r("supplier_ID")
            drg.Rows(i).Cells(5).Value = r("Kind_num")
            drg.Rows(i).Cells(6).Value = r("back2")
            drg.Rows(i).Cells(7).Value = r("back_cost2")
            drg.Rows(i).Cells(8).Value = r("Customer_Price2")
            drg.Rows(i).Cells(9).Value = Math.Round(((((r("Customer_Price2") - r("Cost")) / r("Customer_Price2"))) * 100), 2)
            drg.Rows(i).Cells(10).Value = r("Cost")
            V1 += r("Customer_Price2")
            V2 += r("Cost")
            B += r("back_cost2")
        Next
        Myconn.DataGridview_MoveLast(drg, 2)
        Label19.Text = V2
        Label18.Text = V1
        Label22.Text = B
        Label20.Text = Math.Round(((Val(Val(Label18.Text) - Val(Label19.Text)) / Val(Label18.Text)) * 100), 2)
    End Sub

    Private Sub cboSearch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSearch.SelectedIndexChanged
        Select Case cboSearch.SelectedIndex
            Case 0 '  كل الفواتير
                L1.Visible = False
                L2.Visible = False
                txt1.Visible = False
                txt2.Visible = False
                cboSupplier.Visible = False

            Case 1 ' فواتير خلال فترة
                L1.Visible = True
                L2.Visible = True
                txt1.Visible = True
                txt2.Visible = True
                cboSupplier.Visible = False
            Case 2 ' فواتير مورد
                L1.Visible = False
                L2.Visible = False
                txt1.Visible = False
                txt2.Visible = False
                cboSupplier.Visible = True
                cboSupplier.ComboBox.DataSource = Nothing
                Myconn.Fillcombo("select * from Supplier order by Supplier_Name", "Supplier", "Supplier_ID", "Supplier_Name", Me, cboSupplier.ComboBox)
            Case 3 ' فواتير مورد خلال فترة
                L1.Visible = True
                L2.Visible = True
                txt1.Visible = True
                txt2.Visible = True
                cboSupplier.Visible = True
                cboSupplier.ComboBox.DataSource = Nothing
                Myconn.Fillcombo("select * from Supplier order by Supplier_Name", "Supplier", "Supplier_ID", "Supplier_Name", Me, cboSupplier.ComboBox)
        End Select


    End Sub

    Private Sub frmAll_Purchases_Bills_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width

            If F <> 1 Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Exit Sub
                Dim r As DataRow = Myconn.dt.Rows(0)
                If r("U_full").ToString = False Then
                    btnSearch.Enabled = r("U_search").ToString
                    btnPrint.Enabled = r("U_print").ToString
                End If
            End If
            'If My.Settings.Factory_Price = True Then
            '    drg.Columns(8).HeaderText = " سعر المصنع"
            'Else
            '    drg.Columns(8).HeaderText = " سعر الشراء"
            'End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        fillgrd()
    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Print_bills()
    End Sub
    Private Sub Print_bills()
        Dim table As New DataTable
        For i As Integer = 1 To 11
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(1).Value ' رقم الفاتورة
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(2).Value ' التاريخ
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value ' المورد
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value ' الكود
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value ' عدد الأصناف
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value ' عدد المرتجعات 
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(7).Value ' قيمة المرتجعات
            table.Rows(table.Rows.Count - 1)(8) = dr.Cells(8).Value ' سعر المصنع
            table.Rows(table.Rows.Count - 1)(9) = " % " & dr.Cells(9).Value ' الخصم
            table.Rows(table.Rows.Count - 1)(10) = dr.Cells(10).Value ' سعر التكلفة بعد الخصم
        Next
        Dim rpt As New rpt_Pur_Bills
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        'rpt.SetParameterValue("Bill_num", Format(CDate(txt1.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("Price", Label18.Text)
        rpt.SetParameterValue("reduce", Math.Round((Label18.Text - Label19.Text), 2))
        rpt.SetParameterValue("Nisba", Label20.Text)
        rpt.SetParameterValue("Total", Label19.Text)
        rpt.SetParameterValue("Supplier", Label22.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If

    End Sub
End Class